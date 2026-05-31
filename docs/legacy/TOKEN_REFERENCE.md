# Token Reference

Every date/time token the simulation understands, so hand-edits to
`Emails.xlsx` use forms the code can actually read. If a token isn't on this
list, the grader can't check a date for it — it falls back to "did the model
create *something*" (existence-only).

**Source of truth:** `engine._resolve_one_token` (resolution) and `grader.py`
(what gets checked). If this doc and the code disagree, the code wins — update
this doc.

---

## TL;DR for editing criteria

1. **`{...}` means "check a date/time." Plain text means "check the words."**
   - `CC-{date+2}` → an event must exist **on** that date.
   - `TC-greenlight` → a todo whose title/description **contains** "greenlight".
   - A braced token that *isn't* a date (`{C}`, `{deadline}`) checks nothing —
     it just means "an item exists."

2. **Spacing and capitalization inside `{...}` don't matter.** Before matching,
   the token is lowercased and **all whitespace is removed**. So
   `{date+1- 3PM}`, `{date+1-3pm}`, and `{DATE +1 - 3 PM}` are identical.

3. **Bare `{date}` = the day the email is delivered** (the "served" day), and it
   is strict-matched. The served day is arbitrary relative to the email's story.
   If the email's real target is a *different* day ("next Tuesday", "the 3rd
   Friday of next month"), **don't use bare `{date}`** — use one of the relative
   tokens below, or the date will grade against the wrong day.

4. **Prefixes:** `TC-` todo, `CC-` calendar event, `RS-` reschedule (event
   exists), `No action` (nothing should be created). Anything with no prefix is
   *free-text* and is reported as ungraded (not scored).

5. **Commas separate required items; `or` separates alternatives.**
   - `TC-{date}, TC-{date}` → **two** todos required (count = number of subs).
   - `CC-{date+1- 3PM} or CC-{date+2- 11AM}` → passes if **either** matches.
   - `(single to-do only, not three)` forces the count to exactly 1.

6. **Date checks now apply to todos too (G2).** `TC-{date+2}` checks the todo's
   `due_date`, not just that a todo exists. Bare/unresolved tokens stay
   existence-only.

---

## Supported tokens

Examples below assume the **served day = Wednesday, March 15, 2000**.

### Anchored to the served day

| Token | Resolves to | Example |
|---|---|---|
| `{date}` | the served day | `March 15, 2000` |
| `{date+N}` | N days after | `{date+2}` → `March 17, 2000` |
| `{date+Nweeks}` | N weeks after | `{date+3weeks}` → `April 05, 2000` |
| `{date-tomorrow}` / `{date-nextday}` | served + 1 day | `March 16, 2000` |
| `{date-thisweek}` / `{date-this week}` | the served day (treated as "today") | `March 15, 2000` |
| `{date-beginningmonth}` | the 1st of the served month | `March 01, 2000` |

### Times of day (produce a date **and** time)

Time must be `Nam`/`Npm`, optionally with minutes: `3PM`, `10am`, `11pm`,
`1:14PM`, `12:30pm`. (24-hour times like `14:00` are **not** parsed.)

| Token | Resolves to |
|---|---|
| `{date-3PM}` | `March 15, 2000 at 03:00 PM` |
| `{date-1:14PM}` | `March 15, 2000 at 01:14 PM` |
| `{date+1- 3PM}` | `March 16, 2000 at 03:00 PM` |
| `{date+2- 11AM}` | `March 17, 2000 at 11:00 AM` |
| `{date- this week 11pm}` | `March 15, 2000 at 11:00 PM` |

When a time is present the grader checks **date + hour + minute**; the event
(or todo) must start at exactly that time.

### Day-of-month and weekday ("next future occurrence")

These mean "the next time that day comes up, on or after the served day." If
this month's instance is already past, they roll to next month.

| Token | Resolves to | Note |
|---|---|---|
| `{date-14th}` | `April 14, 2000` | the 14th already passed → next month |
| `{date-25th}` | `March 25, 2000` | still upcoming this month |
| `{date-Tuesday}` | `March 21, 2000` | next Tuesday on/after served |
| `{date-Wednesday}` | `March 15, 2000` | served day *is* Wednesday → today |

### Next week

"Next week" = the calendar week **after** the served day's week
(weeks start Monday).

| Token | Resolves to | Note |
|---|---|---|
| `{date-nextweek}` | `March 22, 2000` | served + 7 days |
| `{nextweek-date +3}` | `March 25, 2000` | served + 7, then +3 |
| `{nextweek-date -2}` | `March 20, 2000` | served + 7, then −2 |
| `{nextweek-wednesday}` | `March 22, 2000` | next week's Wednesday |
| `{nextweek-friday}` | `March 24, 2000` | next week's Friday |
| `{nextweek-Thursday, 3pm GMT}` | `March 23, 2000 at 03:00 PM` | weekday + time; `GMT`/`UTC` is dropped (the whole sim runs in UTC) |

`{date-nextweek}`, `{date-next week}`, `{date-next-week}`, `{nextweek-date}`,
`{nextweek - date}` are all the same token (whitespace/hyphens ignored).

### Nth weekday of a month

Ordinal + weekday. Ordinal is `first`–`fifth` (or `1st`–`5th`) or `last`.
Default month is the **next future occurrence**; add `nextmonth` or
`of <month>` to pin the month. An optional `±N` shifts by days; any trailing
words (a label like "dinner") are ignored.

| Token | Resolves to | Note |
|---|---|---|
| `{third Friday}` | `March 17, 2000` | this month's 3rd Friday (still upcoming) |
| `{first Monday}` | `April 03, 2000` | March's 1st Monday (the 6th) already passed |
| `{last Friday}` | `March 31, 2000` | last Friday of the served month |
| `{third Friday nextmonth}` | `April 21, 2000` | pinned to next month |
| `{third Friday of November}` | `November 17, 2000` | pinned to a named month |
| `{third Friday -1 dinner}` | `March 16, 2000` | 3rd Friday minus 1 day; "dinner" ignored |
| `{fifth Friday}` | rolls forward to a month that *has* a 5th Friday | |

> **Heads-up:** bare `{third Friday}` is "this month if upcoming, else next."
> If an email says "the third Friday of **next** month," write
> `{third Friday nextmonth}` so it doesn't grade against this month.

### Misc

| Token | Resolves to |
|---|---|
| `{meeting-link}` / `{link}` | `[meeting link]` (placeholder, not a date) |

---

## NOT supported (these grade existence-only)

These either aren't dates or aren't a form the resolver reads. They don't
error — the criterion just can't check a date and falls back to "an item
exists."

| Token | Why | What to do |
|---|---|---|
| `{C}`, `{A}`, `{B}` | abstract "date A/B/C", no real date | leave as-is, or rewrite the email to use a concrete token |
| `{deadline}` | a judgment ("flag ambiguity"), not a date | leave as-is (structural grader can't verify it) |
| `{greenlight product A}` | a *content* check wearing braces | drop the braces: `TC-greenlight` |
| `{date-12:30-2:00PM}` | a time **range** | pick one time: `{date-12:30PM}` |
| 24-hour times (`{date-14:00}`) | parser needs `am`/`pm` | write `{date-2PM}` |
| `{Tuesday- this week at 3:00 PM}` | bare-weekday prefix with "at" | use `{date-Tuesday}` (+ a time form if needed) |

---

## How to fix a wrong criterion (cookbook)

- **The email targets a specific date that isn't the arrival day, and a token
  exists for it** → use that token. E.g. "by next Wednesday" → `TC-{nextweek-wednesday}`.
- **The criterion is checking *words*, not a date** → drop the braces.
  E.g. S21: `TC-{greenlight product A}` → **`TC-greenlight`**.
- **"Third Friday of next month"** → `{third Friday nextmonth}` (bare
  `{third Friday}` resolves to *this* month when it's still upcoming).
- **Two events where one is "the evening before"** → `CC-{third Friday nextmonth}, CC-{third Friday nextmonth -1 dinner}`.
- **The real target can't be expressed as any token** (rare prose date) → leave
  bare `{date}` (it grades existence-only-ish against the served day) and note
  it; this is a known dataset limitation, not a code bug.

See `HANDOFF.md` §4 and `docs/REMAINING_WORK.md` (G1/G2/G5) for the full
division of labor: the code grades faithfully whatever the criterion says;
the answer key is edited by hand.
