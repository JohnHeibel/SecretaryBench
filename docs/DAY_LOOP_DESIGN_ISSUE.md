# Design issue: the runner delivers email-at-a-time, not day-at-a-time

**Status:** open · **Filed:** 2026-06-01 · **Area:** `sb/live/runner.py`, `sb/live/mcp_app.py`

## Summary

The benchmark is *supposed* to prompt the model **once per simulated day** and make
it **pull its own inbox through tools** — list the day's new emails, read each one,
decide what (if anything) to do. The current runner instead prompts **once per
email** and **injects the email body directly into the prompt**. Both differences
make the task more artificial than intended.

## What we want (the intended design)

One turn = one day. The model is woken up and told, in effect:

> It's Thursday, March 12. You have new mail. Go handle your inbox.

It is **not** handed any email text. To do the day it must:

1. Call a tool to **list the new/unread emails for today** (sees "there are 5").
2. **Read** email 1 (tool call) → decide: nothing to do.
3. **Read** email 2 → "schedule the signing" → `create_event`.
4. **Read** email 3 → look back at an earlier email via `search_inbox` → act.
5. ...continue until it has worked through the day's batch, then stop.

This mirrors a real executive assistant: mail arrives in a pile, you open the
client, you triage. The model owns the loop of *discovering* what's in front of it,
not just reacting to a single pre-opened message. That discovery + triage step —
"which of these 5 actually need action, and in what order?" — is a core thing we
want to measure.

## What the code does today

### 1. Loop is per-email, not per-day

`sb/live/runner.py:224` flattens the day structure away before the loop even starts:

```python
order = [e for day in plan.per_day for e in day]   # day batches collapsed to a flat list
...
for i, eid in enumerate(order, 1):                  # one iteration per EMAIL
    ...
    cmd = ["claude", "-p", ...]                     # one claude turn per EMAIL
```

So a 3-email day becomes 3 separate `claude -p` turns (resumed into one session via
`--resume`), each stamped with the same `Today:` date. The model never sees "today's
batch" as a unit and never has to order or triage within a day.

### 2. The email is spoon-fed into the prompt

`sb/live/runner.py:250-254` builds the user message *from the email body itself*:

```python
user_msg = (
    f"Email-Id: {eid}\n"
    f"Today: {sd.strftime('%A, %B %d, %Y')}\n"
    f"From: {email.sender}\nTo: {', '.join(email.recipients)}\n"
    f"Subject: {email.subject}\n\n{rendered}")     # full body handed over directly
```

The model doesn't *discover* the email — it's already open on its desk. The
`search_inbox` / `get_email` tools exist (`sb/live/mcp_app.py:103-125`) but today
they're only useful for looking *backward* at older facts, not for finding out
what arrived today.

### 3. There is no "list today's inbox" tool

The MCP surface (`sb/live/mcp_app.py`) has:

- `create/update/delete_event`, `create/update/delete_todo`, `list_events`, `list_todos`
- `search_inbox(query, sender, ...)` — backward fact retrieval
- `get_email(email_id)` — fetch one past email by id

There is **no** tool like `list_new_emails()` / `list_unread()` that returns the
emails delivered *today* (without their ids being known in advance). Until that
exists, a per-day prompt has no honest way for the model to find the day's mail.

## Why the current behavior understates the task

| Dimension | Per-email + injected (current) | Per-day + tool-pulled (intended) |
|---|---|---|
| Triage | none — one email, pre-opened | must list, skim, and prioritize a batch |
| Intra-day ordering | n/a (serialized for the model) | model decides read/act order itself |
| Discovery | n/a — email is handed over | model must call a tool to even see the mail |
| Realism | a bot reacting to a single push | an assistant working an inbox |
| Same-day vs next-day email | indistinguishable except date stamp | a real "5 things landed at once" pile |

The upstream scheduler already models days correctly: `plan.per_day` is a
list-of-batches, and the readiness rule forbids same-day dependencies
(`served on a strictly earlier day`, `< d`) — which only makes sense if a day is one
atomic unit the model handles together. The runner is the only layer that collapses
that structure.

## What a fix touches (sketch, not prescriptive)

1. **Add an inbox-listing tool** in `sb/live/mcp_app.py`, e.g.
   `list_new_emails()` → the emails whose `served_date == today`, returning
   id + sender + subject (and maybe a snippet), *not* the full body. Back it with a
   store endpoint that knows "today."
2. **Rewrite the runner loop** in `sb/live/runner.py` to iterate
   `for day, batch in enumerate(plan.per_day)` instead of over the flattened
   `order`. Per day:
   - POST all of the day's emails to `/inbox` (so the list tool can see them).
   - Send **one** `claude -p` turn: a short "it's <date>, you have new mail, handle
     your inbox" prompt with **no bodies**.
   - Let the model list → read (`get_email`) → act across the whole batch in that
     single turn.
3. **Grade per email, as today**: after the day's turn, diff the store state and run
   `grade_email` for each email in the batch against its own `answer.expect` /
   `forbid`. Grading is already state-based, so it doesn't care whether the actions
   came from one turn or several — only the loop shape changes.
4. **Decide the store's "today" semantics** so `list_new_emails` returns only the
   current day's arrivals while `search_inbox` / `get_email` still reach everything
   delivered so far.

## Open questions

- Should `list_new_emails` mark emails read once fetched, or stay idempotent within
  the day? (Affects whether "did it open all 5?" is observable.)
- Per-turn timeout: a day-turn does more work than an email-turn — `per_turn_timeout`
  likely needs raising / making per-day.
- Should the day prompt hint at the count ("you have new mail") or stay fully blind
  so listing is mandatory?
