# Running the benchmark

Two steps: build a corpus, then run a model on it. Everything is reproducible.
Same seed in, same result out.

## Glossary (every email is one of these, and all three are graded)

- **needle** — a test whose answer needs a fact from an **earlier email** (an
  `@anchor` reference), so the model has to retrieve it. Example: setup says
  `{!signing = +5d}`, a later email's answer wants `@signing+2w`. This is the
  retrieval kind, and the **only** kind whose difficulty grows as you add filler
  (the gap between setup and payoff is the "span").
- **test (self-contained)** — a test whose answer is computable from **that email
  alone** (e.g. "meet next Thursday" → `next:THU`). Graded, but no retrieval, so
  filler doesn't make it harder.
- **filler / distractor** — an email with **no expected action**: pure haystack
  noise. Still graded — the model must correctly do nothing; over-acting on it
  fails that email.

"Not a needle" does **not** mean "not valuable": self-contained tests probe
reasoning and distractors punish over-acting. The needle is just the retrieval axis.

## How it fits together

```
  AUTHOR
  ------
  webapp  (drag date-token blocks, lint + oracle check, save to DB)
     |
     |  export nodes
     v
  corpus/nodes/*.json     the real tests: emails with {date tokens} + answer keys
     |
     |  sb.scale          (keep the real tests, bury them in junk filler)
     v
  build/scaled/           your tests + a big haystack of no-action junk
     |
     |  sb.scheduler      (fake calendar from Jun 1 2026, + seed + days)
     v
  serve plan              which email lands on which fake day
     |
     |  sb.live.runner    (hand the model one email at a time)
     v
  +------------------------------------------+
  |  email  ->  MODEL  ->  calendar / todos  |   the model acts with real tools (MCP)
  +-------------------+----------------------+
                      |
                      |  grader: do the calendar/todos match the answer key?
                      v
  run.log             PASS / FAIL per email
     |
     |  sb.analyze   (+ span = how far back the needed fact was)
     v
  tier x span grid    where it passes, and whether distance hurts it
```

Two things hold this together underneath:

```
  THE GRAMMAR keeps the email and the answer in sync (one source, can't drift):

     email body:  the signing is {!signing = +5d}   ->   "Sunday, June 14, 2026"
     answer key:  start = @signing+2w               ->   June 28, 2026
                                                          ^ same token, cannot disagree

  THE ORACLE solves the whole corpus from the answer keys BEFORE any model runs:

     100% = every test is winnable.  less than 100% = a test is impossible, fix it.
```

## The commands

```
# 1. build a scaled corpus (your handwritten nodes + a junk haystack)
.venv/bin/python -m sb.scale --filler 300 --needles 0 --seed 42 --days 300

# 2. run a model against it (this step costs API calls)
NO_COLOR=1 ./run.sh --model claude-haiku-4-5 --corpus build/scaled --seed 42 --days 300 > build/run.log 2>&1

# 3. score it
.venv/bin/python -m sb.analyze build/run.log --corpus build/scaled --seed 42 --days 300
```

Use the SAME `--seed` and `--days` in all three or the numbers won't line up.

## The parameters

| flag | what it does | typical |
|------|--------------|---------|
| `--filler`  | how many junk emails to bury the facts under | 200-300 |
| `--needles` | machine-made tests per tier. 0 once you write your own | 0 |
| `--seed`    | the dice. same seed = the exact same run, forever | 42 |
| `--days`    | how long the inbox keeps delivering mail | 200-300 |
| `--corpus`  | which corpus folder to read | `build/scaled` |
| `--model`   | which model to test (run step only) | `claude-haiku-4-5` |

`NO_COLOR=1` just keeps the log file plain text so `sb.analyze` can read it. Nothing more.

## Start and end dates

The clock is fake. It always starts on **June 1, 2026** (day 0). It never looks at
today's real date.

`--days` is the **ceiling**, not the finish line. The sim stops the moment every
email has been delivered, which is usually well before the ceiling.

```
  start                                                --days ceiling
  Jun 1                                                Jun 1 + days
   |                                                        |
   v                                                        v
   |== emails delivered, 1 to 5 per day ==|................|
   day 0                               day N
                                        ^
                                        last email out = the REAL end.
                                        everything is delivered, so it stops here.
```

So `--days` only has to be **big enough**:

```
  too small  ->  error: "InfeasibleSchedule" (a deadline can't be met in time)
  just right ->  everything fits, sim stops early, you're good
  too big    ->  no harm, it just stops early anyway
```

Rule of thumb: keep `--days` comfortably above what the corpus needs. More emails
or longer date chains need more days.

## What a "needle" looks like in the stream

A needle is two emails with a pile of junk between them. The gap is the whole point.

```
  SETUP                       ...junk...junk...junk...        PAYOFF
  "migration is Nov 27"       (the haystack, ~60 emails)      "book the review one
                                                               week after the
                                                               migration"
  day 8                                                        day 70
    |<-------------------- span = 62 emails --------------------->|
    states the fact           by now the fact has probably         needs the fact,
                              scrolled out of the model's          but it's gone, so
                              memory                               it has to search
```

Small span = the fact is still in the model's memory, easy. Large span = it got
pushed out, so the model has to go dig for it in the inbox. That gap is the thing
we are actually measuring.
