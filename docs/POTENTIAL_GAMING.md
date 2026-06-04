# Potential Gaming Surfaces

Status: monitor-first, created 2026-06-02.

This note records two benchmark hardening risks we are not blocking yet. The current
plan is to log them during live runs, inspect whether models actually exploit them,
and only then tighten the tool contract. This keeps the pilot realistic without
quietly changing what the grader measures.

## 1. Broad inbox search

`search_inbox` is intentionally available because long-horizon tasks should let the
assistant recover old facts. The possible gaming pattern is using it as an archive
dump instead of a targeted retrieval tool, for example calling it with no query or
asking for a very large result set.

Why this matters: if a model refreshes the whole inbox every day, compaction matters
less. The task drifts from "notice that this email needs an old fact and retrieve it"
toward "absorb a huge tool dump."

Current stance: monitor, do not block. Dumping many emails still forces the model to
read and reason over them, and we do not yet know whether models will do this in
real runs.

Current logging:

- `broad_search`: `search_inbox` was called without `query` or `sender`.
- `large_search_result`: `search_inbox` requested `limit=0` or `limit > 25`.

Potential future fix:

- require `query` or `sender`;
- cap `limit`;
- return snippets from search and require `get_email(email_id)` for full bodies;
- report search-query quality separately in `sb.analyze`.

## 2. Suspicious email_id attribution

Created events and todos carry an `email_id` so the day-loop grader can split a
day's new objects back to the email they answer. The possible gaming pattern is
creating objects with a fake, stale, or wrong email id so the grader ignores them
or routes them away from a no-action email.

Why this matters: unlike broad search, attribution abuse can directly distort the
score. It can hide over-action or misfile work outside the node being graded.

Current stance: monitor first in the pilot log, then harden if it appears. The
prompt already requires the model to copy the current email id exactly; logging will
tell us whether this is a real failure mode or just theoretical.

Current logging:

- `invalid_email_id`: created object used an id that was never delivered.
- `stale_email_id`: created object used a delivered id, but not one from the current
  day's batch.

Potential future fix:

- reject create calls unless `email_id` is in today's delivered batch;
- count any new object with an invalid or stale id as a day-level grading failure;
- add raw attribution diagnostics to the run summary.

## PR note

The DAG/day-loop PR should mention that these are known monitor-first risks, not
silent unknowns. If either warning appears regularly in pilot logs, hardening it is
the next benchmark-integrity task.
