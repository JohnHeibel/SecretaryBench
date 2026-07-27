# claude-haiku-4-5 — SecretaryBench full run (raw CLI evidence)

- **Driver:** claude (`claude -p`)
- **Corpus:** corpus/ (167 emails, sha d737d44e14dc7d20) · seed 42 · days 30 · daily_max 21
- **Score:** `SCORE 98/167 (59%)`
- **Tally:** PASS 98 · FAIL 69 · ERROR 0 (clean retry — supersedes the earlier rate-limited 84/176 run on the old 176-email corpus)
- **Run log:** build/run_claude-haiku-4-5_v2.log (completed 2026-07-04)

## Exact stdout

```text

╔═══ SecretaryBench · live run ═══╗
║ model claude-haiku-4-5 via claude
║ seed 42 · 16 days · 167 emails · start 2026-06-01
╚═════════════════════════════════╝

── day 1 · Mon Jun 01 · 21 new email(s) ──
   tools  ToolSearch, ToolSearch, list_new_emails, get_email, get_email, get_email, get_email, get_email, create_event, create_todo
   model said  “Done. Processed 21 emails today:  **Events created (11):** - Anthropic AI meeting (Tue, Jun 2) - Google Gemini demo (Tue, Jun 2)   - OpenAI ChatGPT demo (Wed, Jun 3) - FBS delegation planning meeting …”

 ✓ PASS  [1] Rebrand-goes-company-wide.hard-no-on-teasers  · served Mon Jun 01
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [2] Innovation-comp.sponsor-mixer-before-the-final  · served Mon Jun 01
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [3] press-tour.press-plan-for-launch  · served Mon Jun 01
   ✓ expected  event ~"interview" @ Thu Aug 06
     actual    "Cover story interview for launch press tour" Thu Aug 06 9 AM (60m)
     why       matched

 ✓ PASS  [4] Innovation-comp.come-check-out-the-semifinal-showcase  · served Mon Jun 01
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [5] Innovation-comp.need-a-budget-meeting-for-the-pitch-comp  · served Mon Jun 01
   ✓ expected  event ~"budget" @ Sun Jul 05
     actual    "Pitch comp budget meeting with Finance & VP Product" Sun Jul 05 9 AM (60m)
     why       matched

 ✓ PASS  [6] Innovation-comp.pitch-comp-budget-review-optional-sit-in  · served Mon Jun 01
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [7] project_atlas.freeze  · served Mon Jun 01
   ✓ expected  event ~"freeze" @ Mon Jun 08
     actual    "Project Atlas code freeze - Go/No-go decision" Mon Jun 08 9 AM (60m)
     why       matched

 ✓ PASS  [8] shoe-product-launch-delays.product-launch-ready  · served Mon Jun 01
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [9] Planning.rebrand-pitch  · served Mon Jun 01
   ✓ expected  event ~"vision" @ Mon Jun 15
     actual    "Rebrand vision sync with VP Marketing" Mon Jun 15 9 AM (60m)
     why       matched

 ✗ FAIL  [10] Innovation-comp.one-of-our-designers-got-a-job-offer  · served Mon Jun 01
   ✗ expected  event ~"conversation" @ Wed Jun 03
     actual    (nothing matching created)
     why       no event titled like "conversation" was created

 ✓ PASS  [11] Enterprise_Ai_Selection.openai-review  · served Mon Jun 01
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [12] World_Cup_Cleat_Launch.project-design-kickoff  · served Mon Jun 01
   ✓ expected  to-do ~"decision" @ Fri Jun 05
     actual    "Decide on World Cup cleat launch window (early vs. late)" Fri Jun 05 5 PM
     why       matched

 ✓ PASS  [13] Planning.acme-trademark-clear  · served Mon Jun 01
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [14] Pre-Launch.launch-day-locked  · served Mon Jun 01
   ✗ expected  event ~"site" @ Mon Aug 10
     actual    (nothing matching created)
     why       no event titled like "site" was created
   ✓ expected  to-do ~"sign" @ by Mon Aug 03
     actual    "Review and sign-off on COO items" Mon Aug 03 5 PM
     why       matched

 ✗ FAIL  [15] Partnership-with-deeptech-companies.whoop-collaboration-meeting  · served Mon Jun 01
   ✗ expected  event ~"WHOOP Meeting" @ Wed Jun 10
     actual    (nothing matching created)
     why       no event titled like "WHOOP Meeting" was created

 ✗ FAIL  [16] Sponsoring-Marathon.approval-of-event  · served Mon Jun 01
   ✗ expected  to-do ~"event" @ Mon Aug 10
     actual    (nothing matching created)
     why       no to-do titled like "event" was created

 ✓ PASS  [17] Day-of-execution_and_Aftermath.we-are-live  · served Mon Jun 01
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [18] Innovation-comp.pitch-comp-is-on-locking-in-the-final-da  · served Mon Jun 01
   ✗ expected  event ~"pitch" @ Wed Jul 15
     actual    "Pitch comp budget meeting with Finance & VP Product" Sun Jul 05 9 AM (60m); "Innovation Pitch Competition Final Event" Wed Jul 15 9 AM (120m)
     why       found 2 matching, expected exactly 1 (duplicate / double-booked)

 ✓ PASS  [19] press-tour.keynote-at-the-footwear  · served Mon Jun 01
   ✓ expected  event ~"keynote" @ Wed Aug 12
     actual    "Keynote at Footwear Expo" Wed Aug 12 9 AM (60m)
     why       matched

 ✗ FAIL  [20] Enterprise_Ai_Selection.ai-meeting-schedule  · served Mon Jun 01
   ✓ expected  event ~"anthropic" @ Tue Jun 02
     actual    "Anthropic AI meeting (Zoom)" Tue Jun 02 9 AM (60m)
     why       matched
   ✓ expected  event ~"google" @ Tue Jun 02
     actual    "Google Gemini demo (Zoom)" Tue Jun 02 10:30 AM (60m)
     why       matched
   ✗ expected  event ~"through" @ Wed Jun 03
     actual    (nothing matching created)
     why       no event titled like "through" was created

 ✗ FAIL  [21] Partnership-with-deeptech-companies.caltech-conference-invitation  · served Mon Jun 01
   ✗ expected  to-do ~"team" @ Fri Jun 05
     actual    (nothing matching created)
     why       no to-do titled like "team" was created
   ✓ expected  event ~"fbs" @ Fri Jun 05
     actual    "FBS delegation planning meeting" Fri Jun 05 9 AM (60m)
     why       matched
   ✗ expected  event ~"conference" @ Tue Aug 25
     actual    (nothing matching created)
     why       no event titled like "conference" was created

── day 2 · Tue Jun 02 · 19 new email(s) ──
   tools  list_new_emails, get_email, get_email, get_email, get_email, create_event, create_todo
   model said  “Done. Processed 19 emails today:  **Events created/updated (9):** - Company Retreat starts (Mon, Jun 22) - Private design team meeting (Wed, Jun 3) - Board greenlight call - WC cleat (Mon, Jun 8) - WH…”

 ✓ PASS  [22] Company-Retreat.company-retreat-date-conformed  · served Tue Jun 02
   ✓ expected  event ~"Company Retreat" @ Mon Jun 22
     actual    "Company Retreat - Start" Mon Jun 22 9 AM (480m)
     why       matched

 ✓ PASS  [23] Innovation-comp.let-s-set-up-a-recap  · served Tue Jun 02
   ✓ expected  event ~"recap" @ Tue Aug 25
     actual    "Innovation comp recap meeting" Tue Aug 25 9 AM (60m)
     why       matched

 ✓ PASS  [24] press-tour.keynote-slot-swapped  · served Tue Jun 02
   ✓ expected  event ~"keynote" @ Thu Aug 13
     actual    "Keynote at Footwear Expo" Thu Aug 13 11 AM (60m)
     why       matched

 ✗ FAIL  [25] Partnership-with-deeptech-companies.whoop-meeting-reschedule  · served Tue Jun 02
   ✗ expected  event ~"WHOOP Meeting" @ Wed Jun 10
     actual    (nothing matching created)
     why       no event titled like "WHOOP Meeting" was moved

 ✓ PASS  [26] Innovation-comp.storm-warning-for-finals-week  · served Tue Jun 02
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [27] Innovation-comp.final-lineup-confirmed  · served Tue Jun 02
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [28] Innovation-comp.this-pitch-comp-thing-is-going-around  · served Tue Jun 02
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [29] Pre-Launch.dry-run-of-the-flip  · served Tue Jun 02
   ✗ expected  event ~"launch" @ Thu Aug 06
     actual    (nothing matching created)
     why       no event titled like "launch" was created

 ✓ PASS  [30] Innovation-comp.vp-product  · served Tue Jun 02
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [31] shoe-product-launch-delays.design-delays-internal-conflicts  · served Tue Jun 02
   ✗ expected  to-do ~"advertising" @ Mon Jun 15
     actual    (nothing matching created)
     why       no to-do titled like "advertising" was created
   ✗ expected  event ~"melissa" @ Wed Jun 03
     actual    (nothing matching created)
     why       no event titled like "melissa" was created

 ✓ PASS  [32] Day-of-execution_and_Aftermath.congrats-on-the-relaunch  · served Tue Jun 02
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [33] World_Cup_Cleat_Launch.wc-cleat-launch-window-options  · served Tue Jun 02
   ✗ expected  to-do ~"decision" @ Mon Jun 08
     actual    "Decide on World Cup cleat launch window (early vs. late)" Fri Jun 05 5 PM; "Confirm WC cleat launch window decision before June 8th" Mon Jun 08 5 PM
     why       found 2 matching, expected exactly 1 (duplicate / double-booked)

 ✓ PASS  [34] Innovation-comp.kickoff-went-well  · served Tue Jun 02
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [35] Innovation-comp.found-a-typo-on-the-trophy  · served Tue Jun 02
   ✗ expected  event ~"correction" @ Wed Jun 03
     actual    (nothing matching created)
     why       no event titled like "correction" was created

 ✗ FAIL  [36] World_Cup_Cleat_Launch.prototype-status-ahead-of-greenlight  · served Tue Jun 02
   ✗ expected  to-do ~"delivery" @ Fri Jun 05
     actual    (nothing matching created)
     why       no to-do titled like "delivery" was created

 ✓ PASS  [37] World_Cup_Cleat_Launch.greenlight-call-for-the-wc-cleat  · served Tue Jun 02
   ✓ expected  event ~"greenlight" @ Mon Jun 08
     actual    "Board greenlight call - WC cleat" Mon Jun 08 9 AM (60m)
     why       matched

 ✓ PASS  [38] press-tour.morning-show-launch-day  · served Tue Jun 02
   ✓ expected  event ~"morning" @ Mon Aug 10
     actual    "National morning show appearance - LAUNCH DAY" Mon Aug 10 6 AM (60m)
     why       matched

 ✗ FAIL  [39] Innovation-comp.trophy-design-quick-look  · served Tue Jun 02
   ✗ expected  to-do ~"trophy" @ Fri Jun 05
     actual    "Approve corrected trophy engraving text" Mon Jun 08 5 PM; "Review trophy design render - approve or provide feedback" Wed Jun 10 5 PM
     why       found 2 matching, expected exactly 1 (duplicate / double-booked)

 ✓ PASS  [40] Innovation-comp.travel-heads-up-for-finals-week  · served Tue Jun 02
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

── day 3 · Wed Jun 03 · 2 new email(s) ──
   tools  list_new_emails, get_email
   model said  “Done. Processed 2 emails today:  **No new calendar events or to-dos needed.**  - **Add to recap agenda:** VP Marketing wants to include discussion about promoting next year's competition earlier—add t…”

 ✗ FAIL  [41] Innovation-comp.add-this-to-the-recap-agenda  · served Wed Jun 03
   ✗ expected  to-do ~"agenda" @ Tue Aug 25
     actual    (nothing matching created)
     why       no to-do titled like "agenda" was created

 ✗ FAIL  [42] World_Cup_Cleat_Launch.prototype-status-ahead-of-greenlight-2  · served Wed Jun 03
   ✗ expected  to-do ~"outsole" @ Wed Jun 03
     actual    "Push vendor - outsole sample due by Friday June 5th" Fri Jun 05 5 PM
     why       on the wrong day

── day 4 · Thu Jun 04 · 19 new email(s) ──
   tools  list_new_emails, get_email, get_email, get_email, get_email, create_event, create_todo
   model said  “Done. Processed 19 emails today:  **Events created (9):** - End-of-year pizza party (Mon, Jun 8) - Final Meeting Review - AI vendors (Fri, Jun 5, time TBD) - Design team meeting - staffing decisions (…”

 ✗ FAIL  [43] pizza-party.end-of-year-pizza-party  · served Thu Jun 04
   ✗ expected  to-do ~"Team_pizza_party" @ Mon Jun 08
     actual    (nothing matching created)
     why       no to-do titled like "Team_pizza_party" was created

 ✗ FAIL  [44] Enterprise_Ai_Selection.final-review  · served Thu Jun 04
   ✗ expected  to-do ~"ai" @ Fri Jun 05
     actual    (nothing matching created)
     why       no to-do titled like "ai" was created

 ✓ PASS  [45] Innovation-comp.mid-comp-check-in-with-the-teams  · served Thu Jun 04
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [46] Rebrand-goes-company-wide.we-need-to-talk-about-morale  · served Thu Jun 04
   ✓ expected  event ~"morale" @ Fri Jul 10
     actual    "Morale discussion with HR - rebrand concerns" Fri Jul 10 9 AM (45m)
     why       matched

 ✓ PASS  [47] shoe-product-launch-delays.product-launch-ready-soon  · served Thu Jun 04
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [48] Sponsoring-Marathon.launching-sponsoring-eugene-marathon  · served Thu Jun 04
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [49] project_atlas.board-demo  · served Thu Jun 04
   ✗ expected  event ~"atlas" @ Mon Jun 22
     actual    "Project Atlas code freeze - Go/No-go decision" Mon Jun 08 9 AM (60m); "Atlas live board demo" Mon Jun 22 9 AM (90m)
     why       found 2 matching, expected exactly 1 (duplicate / double-booked)

 ✗ FAIL  [50] shoe-product-launch-delays.design-team-meeting  · served Thu Jun 04
   ✗ expected  event ~"design" @ Sun Jun 07
     actual    "Private meeting with Design Lead - team structure discussion" Wed Jun 03 9 AM (60m); "Design team meeting - staffing decisions" Sun Jun 07 10 AM (60m)
     why       found 2 matching, expected exactly 1 (duplicate / double-booked)
   ✗ expected  to-do ~"who" @ by Sun Jun 07
     actual    (nothing matching created)
     why       no to-do titled like "who" was created

 ✓ PASS  [51] project_atlas.legal-fyi  · served Thu Jun 04
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [52] World_Cup_Cleat_Launch.reveal-event-date-and-venue  · served Thu Jun 04
   ✓ expected  event ~"reveal" @ Mon Jul 13
     actual    "WC cleat reveal event" Mon Jul 13 10 AM (120m)
     why       matched

 ✓ PASS  [53] Day-of-execution_and_Aftermath.billboard-photo-op  · served Thu Jun 04
   ✓ expected  event ~"billboard" @ Mon Aug 10
     actual    "Billboard unveiling photo op - LAUNCH DAY" Mon Aug 10 11 AM (60m)
     why       matched

 ✓ PASS  [54] Innovation-comp.team-morale-lunch-during-the-comp  · served Thu Jun 04
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [55] Innovation-comp.two-weeks-out-from-the-final  · served Thu Jun 04
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [56] Innovation-comp.recap-agenda-finalized  · served Thu Jun 04
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [57] Innovation-comp.new-sponsor-wants-to-chat  · served Thu Jun 04
   ✗ expected  event ~"prospect" @ Tue Aug 25
     actual    (nothing matching created)
     why       no event titled like "prospect" was created

 ✓ PASS  [58] project_atlas.ceo-note  · served Thu Jun 04
   ✓ expected  event ~"interview" @ Mon Aug 24
     actual    "First press interview - post-launch" Mon Aug 24 10 AM (60m)
     why       matched

 ✓ PASS  [59] Innovation-comp.quick-win-to-share  · served Thu Jun 04
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [60] Innovation-comp.press-preview-ahead-of-the-final  · served Thu Jun 04
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [61] Innovation-comp.mixer-was-a-hit  · served Thu Jun 04
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

── day 5 · Fri Jun 05 · 13 new email(s) ──
   tools  list_new_emails, get_email, get_email, get_email, create_todo
   model said  “Done. Processed 13 emails today:  **Events created/updated (4):** - WC cleat reveal - rehearsal with athlete (Fri, Jun 12) - Atlas live board demo moved (Fri, Jun 19 - rescheduled from Jun 22) - All-h…”

 ✓ PASS  [62] Innovation-comp.confirmed-for-the-new-sponsor-call  · served Fri Jun 05
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [63] World_Cup_Cleat_Launch.reveal-event-date-and-venue-2  · served Fri Jun 05
   ✗ expected  event ~"rehearsal" @ Sun Jul 12
     actual    "WC cleat reveal - rehearsal with athlete" Fri Jun 12 10 AM (90m)
     why       on the wrong day
   ✗ expected  event ~"Design Lead Stage Slot" @ Sat Jun 20
     actual    (nothing matching created)
     why       no event titled like "Design Lead Stage Slot" was created

 ✓ PASS  [64] Sponsoring-Marathon.pitch-deck  · served Fri Jun 05
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [65] Innovation-comp.post-recap-follow-ups  · served Fri Jun 05
   ✗ expected  to-do ~"judges" @ Thu Aug 27
     actual    "Send thank you note to innovation comp judges" Wed Aug 26 5 PM
     why       on the wrong day

 ✓ PASS  [66] Innovation-comp.great-year-overall  · served Fri Jun 05
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [67] Innovation-comp.order-placed-for-trophies  · served Fri Jun 05
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [68] Innovation-comp.brief-for-the-new-sponsor-call  · served Fri Jun 05
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [69] World_Cup_Cleat_Launch.reveal-event-budget-is-over  · served Fri Jun 05
   ✗ expected  to-do ~"Approve revised event budget" @ Fri Jun 05
     actual    (nothing matching created)
     why       no to-do titled like "Approve revised event budget" was created

 ✗ FAIL  [70] Innovation-comp.need-your-sign-off-on-prize-amounts  · served Fri Jun 05
   ✗ expected  to-do ~"approve" @ Sun Jun 07
     actual    "Approve corrected trophy engraving text" Mon Jun 08 5 PM; "Review trophy design render - approve or provide feedback" Wed Jun 10 5 PM; "Approve prize amounts for innovation comp" Fri Jun 05 5 PM
     why       found 3 matching, expected exactly 1 (duplicate / double-booked)

 ✓ PASS  [71] pizza-party.pizza-party-budget  · served Fri Jun 05
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [72] project_atlas.demo-moved  · served Fri Jun 05
   ✗ expected  event ~"atlas" @ Fri Jun 19
     actual    "Project Atlas code freeze - Go/No-go decision" Mon Jun 08 9 AM (60m); "Atlas live board demo" Fri Jun 19 9 AM (90m)
     why       found 2 matching, expected exactly 1 (duplicate / double-booked)

 ✓ PASS  [73] Rebrand-goes-company-wide.all-hands-to-get-everyone-aligned  · served Fri Jun 05
   ✓ expected  event ~"webinar" @ by Fri Aug 07
     actual    "All-hands webinar - rebrand explanation" Fri Aug 07 10 AM (60m)
     why       matched

 ✓ PASS  [74] Innovation-comp.pitch-comp-is-on-locking-in-the-final-da-2  · served Fri Jun 05
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

── day 6 · Sat Jun 06 · 11 new email(s) ──
   tools  list_new_emails, get_email, get_email, get_email, create_todo
   model said  “Done. Processed 11 emails today:  **Events created (5):** - Press interview - trade outlet follow-up (Tue, Aug 18) - Teaser campaign walkthrough with COMMS (Tue, Jun 30) - Shoe product launch - budget…”

 ✓ PASS  [75] Innovation-comp.press-wants-a-short-interview  · served Sat Jun 06
   ✓ expected  event ~"interview" @ Tue Aug 18
     actual    "Press interview - trade outlet follow-up" Tue Aug 18 10 AM (20m)
     why       matched

 ✓ PASS  [76] Innovation-comp.judges-prep-session-want-in  · served Sat Jun 06
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [77] Rebrand-goes-company-wide.teaser-campaign  · served Sat Jun 06
   ✓ expected  event ~"teaser" @ Tue Jun 30
     actual    "Teaser campaign walkthrough with COMMS" Tue Jun 30 9 AM (60m)
     why       matched

 ✗ FAIL  [78] Partnership-with-deeptech-companies.boston-partnership-trip  · served Sat Jun 06
   ✗ expected  to-do ~"brainstorm" @ by Mon Jun 08
     actual    (nothing matching created)
     why       no to-do titled like "brainstorm" was created
   ✗ expected  event ~"strategy" @ Mon Jun 08
     actual    (nothing matching created)
     why       no event titled like "strategy" was created

 ✓ PASS  [79] Enterprise_Ai_Selection.anthropic-google-review  · served Sat Jun 06
   ✓ expected  to-do ~"meetings" @ Sun Jun 07
     actual    "Send AI vendor review notes and fill out rubric" Sun Jun 07 5 PM
     why       matched

 ✓ PASS  [80] Sponsoring-Marathon.approval-needed-race-sponsorship-budget  · served Sat Jun 06
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [81] Innovation-comp.heads-up-one-of-the-pitches-might-step-o  · served Sat Jun 06
   ✗ expected  event ~"overlap" @ Thu Jul 16
     actual    (nothing matching created)
     why       no event titled like "overlap" was created

 ✗ FAIL  [82] World_Cup_Cleat_Launch.reveal-event-budget-is-over-2  · served Sat Jun 06
   ✗ expected  to-do ~"Approve revised event budget" @ Wed Jun 03
     actual    (nothing matching created)
     why       no to-do titled like "Approve revised event budget" was moved

 ✓ PASS  [83] Company-Retreat.tasks-before-the-retreat  · served Sat Jun 06
   ✓ expected  to-do ~"location" @ Fri Jun 19
     actual    "Contact retreat location to approve plans" Fri Jun 19 5 PM
     why       matched
   ✓ expected  to-do ~"list" @ by Mon Jun 22
     actual    "Provide VIP access employee list for retreat" Mon Jun 22 5 PM
     why       matched

 ✗ FAIL  [84] shoe-product-launch-delays.budget-exceeded  · served Sat Jun 06
   ✗ expected  to-do ~"advertising" @ Mon Jun 22
     actual    (nothing matching created)
     why       no to-do titled like "advertising" was moved
   ✗ expected  event ~"first" @ Mon Jun 15
     actual    (nothing matching created)
     why       no event titled like "first" was created
   ✗ expected  event ~"second" @ Mon Jun 29
     actual    (nothing matching created)
     why       no event titled like "second" was created

 ✓ PASS  [85] press-tour.podcast-taping-pick-a-day  · served Sat Jun 06
   ✓ expected  event ~"podcast" @ Mon Aug 03 or Tue Aug 04 or Wed Aug 05
     actual    "Podcast taping - "Mind Your Own Business"" Tue Aug 04 3 PM (60m)
     why       matched

── day 7 · Sun Jun 07 · 4 new email(s) ──
   tools  list_new_emails, get_email, create_event
   model said  “Done. Processed 4 emails today:  **Events deleted (2):** - Teaser campaign walkthrough (Jun 30 - cancelled per decision to go cold on teasers) - Billboard unveiling photo op (Aug 10 - cancelled due to…”

 ✓ PASS  [86] Rebrand-goes-company-wide.standing-down-on-teasers  · served Sun Jun 07
   ✓ expected  event ~"teaser" cancelled
     actual    (nothing — cancelled)
     why       cancelled

 ✓ PASS  [87] Day-of-execution_and_Aftermath.billboard-s-off  · served Sun Jun 07
   ✓ expected  event ~"billboard" cancelled
     actual    (nothing — cancelled)
     why       cancelled

 ✓ PASS  [88] Sponsoring-Marathon.race-sponsorship-budget-goals  · served Sun Jun 07
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [89] Innovation-comp.sponsor-wants-a-follow-up-call  · served Sun Jun 07
   ✗ expected  event ~"sponsor" @ Sun Jul 12
     actual    "New sponsor intro call" Tue Aug 25 2 PM (60m); "Retail sponsor partnership extension call" Sun Jul 12 10 AM (60m)
     why       found 2 matching, expected exactly 1 (duplicate / double-booked)

── day 8 · Mon Jun 08 · 18 new email(s) ──
   tools  ToolSearch, list_new_emails, ToolSearch, get_email, get_email, get_email, get_email, create_event, create_todo
   model said  “Done. Processed 18 emails today:  **Events created/updated/deleted (10):** - Deleted: Retail sponsor call (no longer needed) - Updated: Company Retreat moved from Jun 22 to Jun 29 - Created: Major ath…”

 ✓ PASS  [90] project_atlas.beta  · served Mon Jun 08
   ✓ expected  to-do ~"feedback" @ by Mon Jun 22
     actual    "Compile and summarize Atlas beta feedback" Mon Jun 22 5 PM
     why       matched

 ✓ PASS  [91] Enterprise_Ai_Selection.enterprise-ai-selection-kickoff  · served Mon Jun 08
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [92] Innovation-comp.today-s-the-day  · served Mon Jun 08
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [93] project_atlas.launch-dinner  · served Mon Jun 08
   ✗ expected  event ~"launch" @ Mon Aug 17
     actual    "First press interview - post-launch" Mon Aug 24 10 AM (60m); "Atlas launch team dinner" Mon Aug 17 6 PM (120m)
     why       found 2 matching, expected exactly 1 (duplicate / double-booked)

 ✓ PASS  [94] Innovation-comp.a-kid-drew-the-new-logo-concept  · served Mon Jun 08
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [95] Sponsoring-Marathon.sponsorship-tiers  · served Mon Jun 08
   ✗ expected  to-do ~"approval" @ Fri Jun 19
     actual    (nothing matching created)
     why       no to-do titled like "approval" was created

 ✓ PASS  [96] Innovation-comp.she-s-staying  · served Mon Jun 08
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [97] World_Cup_Cleat_Launch.press-briefing-and-embargo  · served Mon Jun 08
   ✓ expected  event ~"briefing" @ Mon Jul 06
     actual    "WC cleat press briefing - under embargo" Mon Jul 06 10 AM (90m)
     why       matched
   ✗ expected  to-do ~"Approve Press Embargo" @ Tue Jun 09
     actual    (nothing matching created)
     why       no to-do titled like "Approve Press Embargo" was created

 ✓ PASS  [98] press-tour.local-evening-news-hit  · served Mon Jun 08
   ✓ expected  event ~"segment" @ Tue Aug 11
     actual    "Local evening news segment" Tue Aug 11 5 PM (60m)
     why       matched

 ✓ PASS  [99] Innovation-comp.great-session-today  · served Mon Jun 08
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [100] Pre-Launch.embargo-lifts-at-six  · served Mon Jun 08
   ✓ expected  event ~"embargo" @ Mon Aug 10
     actual    "Press embargo lifts - coverage releases" Mon Aug 10 6 AM (30m)
     why       matched

 ✗ FAIL  [101] Enterprise_Ai_Selection.fridays-meeting-time  · served Mon Jun 08
   ✗ expected  event ~"friday" @ Fri Jun 12
     actual    (nothing matching created)
     why       no event titled like "friday" was created

 ✗ FAIL  [102] Sponsoring-Marathon.pitch-deck-2  · served Mon Jun 08
   ✗ expected  event ~"breifing" @ Thu Jul 02
     actual    (nothing matching created)
     why       no event titled like "breifing" was created

 ✗ FAIL  [103] World_Cup_Cleat_Launch.final-colorway-for-sign-off  · served Mon Jun 08
   ✗ expected  to-do ~"design" @ Fri Jun 19
     actual    "Push vendor - outsole sample due by Friday June 5th" Fri Jun 05 5 PM; "Sign off on final WC cleat colorway - design freeze" Fri Jun 19 5 PM
     why       found 2 matching, expected exactly 1 (duplicate / double-booked)

 ✗ FAIL  [104] Innovation-comp.sponsor-call-no-longer-needed  · served Mon Jun 08
   ✗ expected  event ~"sponsor" cancelled
     actual    "New sponsor intro call" Tue Aug 25 2 PM (60m)
     why       should be cancelled, but 1 still on the calendar

 ✓ PASS  [105] pizza-party.conference-room-reservation  · served Mon Jun 08
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [106] Company-Retreat.in-town-and-would-love-to-connect  · served Mon Jun 08
   ✗ expected  event ~"Company Retreat" @ Mon Jun 15
     actual    "Company Retreat - Start" Mon Jun 29 9 AM (480m)
     why       on the wrong day
   ✗ expected  to-do ~"location" @ Mon Jun 15
     actual    "Contact retreat location to approve plans" Fri Jun 19 5 PM
     why       on the wrong day
   ✓ expected  event ~"athlete" @ Mon Jun 22
     actual    "Meeting with major athlete" Mon Jun 22 10 AM (60m)
     why       matched

 ✓ PASS  [107] World_Cup_Cleat_Launch.greenlight-call-for-the-wc-cleat-2  · served Mon Jun 08
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

── day 9 · Tue Jun 09 · 14 new email(s) ──
   tools  list_new_emails, get_email, get_email, get_email, create_todo
   model said  “Done. Processed 14 emails today:  **Events created/updated (5):** - Boston partnership trip - WHOOP & Boston Dynamics (Mon-Thu, Jul 6-9) - Retreat planning call with retreat company (Fri, Jun 21) - Re…”

 ✗ FAIL  [108] Partnership-with-deeptech-companies.boston-trip-details  · served Tue Jun 09
   ✗ expected  event ~"start" @ Mon Jul 06–Jul 12
     actual    (nothing matching created)
     why       no event titled like "start" was created
   ✗ expected  event ~"end" @ Thu Jul 09
     actual    (nothing matching created)
     why       no event titled like "end" was created
   ✗ expected  event ~"WHOOP HQ Visit" @ Mon Jul 06–Jul 12
     actual    (nothing matching created)
     why       no event titled like "WHOOP HQ Visit" was created
   ✗ expected  event ~"dynamics" @ Wed Jul 08
     actual    "Boston partnership trip - WHOOP & Boston Dynamics" Mon Jul 06 9 AM (4800m)
     why       on the wrong day

 ✗ FAIL  [109] Company-Retreat.planning-call-and-forms-for-your-company  · served Tue Jun 09
   ✗ expected  event ~"Retreat Company Meeting Call" @ Sun Jun 21
     actual    (nothing matching created)
     why       no event titled like "Retreat Company Meeting Call" was created
   ✓ expected  to-do ~"forms" @ Sun Jun 21
     actual    "Complete and submit retreat planning forms" Sun Jun 21 10 AM
     why       matched

 ✗ FAIL  [110] Company-Retreat.athelete-visit  · served Tue Jun 09
   ✓ expected  to-do ~"athlete" @ by Mon Jun 22
     actual    "Create guest list and notify selected people for athlete visit" Sun Jun 21 5 PM
     why       matched
   ✗ expected  event ~"people" @ Mon Jun 22
     actual    (nothing matching created)
     why       no event titled like "people" was created

 ✓ PASS  [111] Innovation-comp.last-year-s-recap-notes-for-reference  · served Tue Jun 09
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [112] Pre-Launch.store-floors-reset-overnight  · served Tue Jun 09
   ✓ expected  event ~"floor" @ Sun Aug 09
     actual    "Retail store floor reset - overnight rebrand" Sun Aug 09 10 PM (480m)
     why       matched

 ✗ FAIL  [113] pizza-party.pizza-place-selection  · served Tue Jun 09
   ✗ expected  to-do ~"decision" @ by Mon Jun 15
     actual    (nothing matching created)
     why       no to-do titled like "decision" was created

 ✗ FAIL  [114] World_Cup_Cleat_Launch.final-colorway-for-sign-off-2  · served Tue Jun 09
   ✗ expected  to-do ~"design" @ Wed Jun 10
     actual    "Push vendor - outsole sample due by Friday June 5th" Fri Jun 05 5 PM; "Sign off on final WC cleat colorway - design freeze" Fri Jun 19 5 PM
     why       found 2 matching, expected exactly 1 (duplicate / double-booked)

 ✓ PASS  [115] Innovation-comp.final-headcount-for-the-recap  · served Tue Jun 09
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [116] Innovation-comp.press-interview-confirmed  · served Tue Jun 09
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [117] Innovation-comp.board-slot-the-week-after-finals  · served Tue Jun 09
   ✓ expected  event ~"presentation" @ Wed Jul 22
     actual    "Board presentation - innovation comp wins" Wed Jul 22 9 AM (120m)
     why       matched

 ✓ PASS  [118] Rebrand-goes-company-wide.please-reconsider-this-rebrand  · served Tue Jun 09
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [119] World_Cup_Cleat_Launch.tooling-po-needs-approval  · served Tue Jun 09
   ✗ expected  to-do ~"Approve tooling PO" @ Fri Jun 12
     actual    (nothing matching created)
     why       no to-do titled like "Approve tooling PO" was created

 ✓ PASS  [120] Pre-Launch.design-locked  · served Tue Jun 09
   ✓ expected  to-do ~"production" @ by Mon Aug 03
     actual    "Ensure campaign design final and to production by Aug 3" Mon Aug 03 5 PM
     why       matched

 ✗ FAIL  [121] Marketing-campaign-new-product-delay.need-to-delay-the-release-of-the-new-pro  · served Tue Jun 09
   ✗ expected  event ~"delayed" @ Tue Aug 04
     actual    (nothing matching created)
     why       no event titled like "delayed" was created

── day 10 · Wed Jun 10 · 8 new email(s) ──
   tools  list_new_emails, get_email, get_email, create_todo
   model said  “Done. Processed 8 emails today:  **Events updated/deleted/created (7):** - Moved: WC cleat reveal rehearsal to Friday, Jun 19 (from Jun 12) - Moved: WC cleat reveal main event to Saturday, Jun 20 (fro…”

 ✗ FAIL  [122] World_Cup_Cleat_Launch.reveal-event-date-and-venue-3  · served Wed Jun 10
   ✗ expected  event ~"reveal" @ Sat Jun 20
     actual    "WC cleat reveal event" Sat Jun 20 10 AM (120m); "WC cleat reveal - rehearsal with athlete" Fri Jun 19 10 AM (90m)
     why       found 2 matching, expected exactly 1 (duplicate / double-booked)
   ✓ expected  event ~"rehearsal" @ Fri Jun 19
     actual    "WC cleat reveal - rehearsal with athlete" Fri Jun 19 10 AM (90m)
     why       matched

 ✗ FAIL  [123] Partnership-with-deeptech-companies.new-meeting-during-boston-trip  · served Wed Jun 10
   ✗ expected  event ~"actnano" @ Thu Jul 09
     actual    "Boston partnership trip - WHOOP & Boston Dynamics" Mon Jul 06 9 AM (6240m)
     why       on the wrong day

 ✗ FAIL  [124] project_atlas.dinner-cancel  · served Wed Jun 10
   ✗ expected  event ~"launch" cancelled
     actual    "First press interview - post-launch" Mon Aug 24 10 AM (60m)
     why       should be cancelled, but 1 still on the calendar

 ✗ FAIL  [125] pizza-party.pizza-order-deadline  · served Wed Jun 10
   ✗ expected  to-do ~"order-the-pizzas" @ Fri Jun 12
     actual    (nothing matching created)
     why       no to-do titled like "order-the-pizzas" was created

 ✓ PASS  [126] Innovation-comp.prep-doc-for-the-board-presentation  · served Wed Jun 10
   ✓ expected  to-do ~"slides" @ Mon Jul 20
     actual    "Review board presentation slides - gut check" Mon Jul 20 5 PM
     why       matched

 ✗ FAIL  [127] Planning.might-be-too-much  · served Wed Jun 10
   ✗ expected  event ~"signoff" @ Thu Jun 11
     actual    (nothing matching created)
     why       no event titled like "signoff" was created

 ✓ PASS  [128] press-tour.news-hit-fell-through  · served Wed Jun 10
   ✓ expected  event ~"segment" cancelled
     actual    (nothing — cancelled)
     why       cancelled

 ✗ FAIL  [129] pizza-party.client-demo-conflict  · served Wed Jun 10
   ✗ expected  to-do ~"Team_pizza_party" @ Tue Jun 09
     actual    (nothing matching created)
     why       no to-do titled like "Team_pizza_party" was moved

── day 11 · Thu Jun 11 · 9 new email(s) ──
   tools  list_new_emails, get_email, get_email, create_todo
   model said  “Done. Processed 9 emails today:  **Events updated/deleted/created (4):** - Deleted: Design team staffing meeting (no longer needed - HR fixed the issue) - Moved: WC cleat press briefing to Saturday, J…”

 ✗ FAIL  [130] Sponsoring-Marathon.launching-sponsoring-eugene-marathon-2  · served Thu Jun 11
   ✗ expected  event ~"launchmeeting" @ Thu Jun 11
     actual    (nothing matching created)
     why       no event titled like "launchmeeting" was created

 ✗ FAIL  [131] World_Cup_Cleat_Launch.manufacturing-kickoff  · served Thu Jun 11
   ✗ expected  event ~"manufacturing" @ Wed Jun 10
     actual    (nothing matching created)
     why       no event titled like "manufacturing" was created

 ✗ FAIL  [132] shoe-product-launch-delays.re-design-team-meeting  · served Thu Jun 11
   ✓ expected  to-do ~"who" cancelled
     actual    (nothing — cancelled)
     why       cancelled
   ✗ expected  event ~"design" cancelled
     actual    "Private meeting with Design Lead - team structure discussion" Wed Jun 03 9 AM (60m)
     why       should be cancelled, but 1 still on the calendar

 ✓ PASS  [133] Enterprise_Ai_Selection.ai-selection-finalization  · served Thu Jun 11
   ✓ expected  to-do ~"sign" @ Sun Jun 14
     actual    "Sign Anthropic deal documents" Sun Jun 14 5 PM
     why       matched

 ✓ PASS  [134] World_Cup_Cleat_Launch.press-briefing-and-embargo-2  · served Thu Jun 11
   ✓ expected  event ~"briefing" @ Sat Jun 13
     actual    "WC cleat press briefing - under embargo" Sat Jun 13 10 AM (90m)
     why       matched

 ✗ FAIL  [135] Planning.can-t-do-thursday  · served Thu Jun 11
   ✗ expected  event ~"signoff" @ Mon Jun 15
     actual    (nothing matching created)
     why       no event titled like "signoff" was moved

 ✗ FAIL  [136] Partnership-with-deeptech-companies.spadxtech-meeting-before-fbs  · served Thu Jun 11
   ✗ expected  event ~"spad" @ Wed Jul 08
     actual    (nothing matching created)
     why       no event titled like "spad" was created

 ✓ PASS  [137] World_Cup_Cleat_Launch.endorsement-terms-for-the-reveal  · served Thu Jun 11
   ✓ expected  to-do ~"endorsement" @ Fri Jun 12
     actual    "Review & approve athlete endorsement terms for reveal" Fri Jun 12 5 PM
     why       matched

 ✓ PASS  [138] pizza-party.updated-pizza-party-schedule  · served Thu Jun 11
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

── day 12 · Fri Jun 12 · 9 new email(s) ──
   tools  list_new_emails, get_email, get_email, create_event
   model said  “Done. Processed 9 emails today:  **Events updated/created (2):** - Moved: Marathon sponsorship launch meeting to Friday, Jun 12 at 2:00pm (from Jun 11 - CFO availability) - Updated: Boston trip - Bost…”

 ✓ PASS  [139] Planning.locked-in-for  · served Fri Jun 12
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [140] World_Cup_Cleat_Launch.endorsement-terms-for-the-reveal-2  · served Fri Jun 12
   ✓ expected  to-do ~"endorsement" @ Wed Jun 17
     actual    "Review & approve athlete endorsement terms for reveal" Wed Jun 17 5 PM
     why       matched

 ✓ PASS  [141] Sponsoring-Marathon.launch-meeting-conflict  · served Fri Jun 12
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [142] Sponsoring-Marathon.approval-of-budget-tier  · served Fri Jun 12
   ✗ expected  event ~"sponsorship & budget approval meeting" @ Wed Jul 01
     actual    (nothing matching created)
     why       no event titled like "sponsorship & budget approval meeting" was created

 ✓ PASS  [143] Innovation-comp.thanks-for-the-sign-off  · served Fri Jun 12
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [144] pizza-party.room-change  · served Fri Jun 12
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [145] Partnership-with-deeptech-companies.confirmation-from-spadxtech  · served Fri Jun 12
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [146] Partnership-with-deeptech-companies.boston-dynamics  · served Fri Jun 12
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✗ FAIL  [147] Partnership-with-deeptech-companies.boston-dynamics-cancel  · served Fri Jun 12
   ✗ expected  event ~"dynamics" cancelled
     actual    "Boston partnership trip - WHOOP & Boston Dynamics" Mon Jul 06 9 AM (6240m)
     why       should be cancelled, but 1 still on the calendar

── day 13 · Sat Jun 13 · 7 new email(s) ──
   tools  list_new_emails, get_email, get_email, create_todo
   model said  “Done. Processed 7 emails today:  **Events created (4):** - Live launch stream - walk through new ACME (Mon, Aug 10 at 7pm) - Design credit discussion - WC cleat reveal (Tue, Jun 16 - urgent before rev…”

 ✓ PASS  [148] Day-of-execution_and_Aftermath.launch-livestream  · served Sat Jun 13
   ✓ expected  event ~"livestream" @ Mon Aug 10
     actual    "Live launch stream - walk through new ACME" Mon Aug 10 7 PM (60m)
     why       matched

 ✓ PASS  [149] Planning.memo-before-the-board-sits  · served Sat Jun 13
   ✓ expected  to-do ~"budget" @ Sun Jun 14
     actual    "Finalize rebrand budget memo for board" Sun Jun 14 5 PM
     why       matched

 ✗ FAIL  [150] Innovation-comp.quick-favor-before-the-final  · served Sat Jun 13
   ✗ expected  to-do ~"board" @ Wed Jul 15
     actual    "Review board presentation slides - gut check" Mon Jul 20 5 PM; "Create one-pager on innovation comp for board deck" Mon Jul 20 5 PM
     why       found 2 matching, expected exactly 1 (duplicate / double-booked)

 ✗ FAIL  [151] World_Cup_Cleat_Launch.manufacturing-kickoff-2  · served Sat Jun 13
   ✗ expected  event ~"Design Lead 1:1" @ Fri Jun 12
     actual    (nothing matching created)
     why       no event titled like "Design Lead 1:1" was created

 ✓ PASS  [152] Marketing-campaign-new-product-delay.need-to-reschedule-athletes-for-marketin  · served Sat Jun 13
   ✓ expected  event ~"william" @ Tue Sep 08
     actual    "Serena Williams marketing campaign shoot" Tue Sep 08 10 AM (120m)
     why       matched
   ✓ expected  event ~"michael" @ Tue Sep 15
     actual    "Michael Phelps marketing campaign shoot" Tue Sep 15 10 AM (120m)
     why       matched

 ✗ FAIL  [153] Sponsoring-Marathon.launch-meeting-conflict-2  · served Sat Jun 13
   ✗ expected  to-do ~"launch" @ by Sun Jun 14
     actual    (nothing matching created)
     why       no to-do titled like "launch" was created

 ✗ FAIL  [154] Company-Retreat.inquiry-on-vip-list  · served Sat Jun 13
   ✗ expected  to-do ~"list" @ by Sun Jun 14
     actual    "Provide VIP access employee list for retreat" Mon Jun 22 5 PM; "Create guest list and notify selected people for athlete visit" Sun Jun 21 5 PM; "Create and finalize VIP access list for company retreat" Sun Jun 14 5 PM
     why       found 3 matching, expected exactly 1 (duplicate / double-booked)

── day 14 · Sun Jun 14 · 9 new email(s) ──
   tools  list_new_emails, get_email, get_email, create_todo
   model said  “Done. Processed 9 emails today:  **Events created/updated (6):** - Green room - mic check & talking points (Mon, Aug 10 at 6pm - before livestream) - Launch team thank-you lunch (Tue, Aug 11) - Ronald…”

 ✗ FAIL  [155] Sponsoring-Marathon.launch-meeting-recap  · served Sun Jun 14
   ✗ expected  event ~"sponsorshippitch" @ Thu Jul 09
     actual    (nothing matching created)
     why       no event titled like "sponsorshippitch" was created

 ✓ PASS  [156] Day-of-execution_and_Aftermath.green-room-before-we-go-live  · served Sun Jun 14
   ✓ expected  event ~"green" @ Mon Aug 10
     actual    "Green room - mic check & talking points" Mon Aug 10 6 PM (60m)
     why       matched

 ✓ PASS  [157] Innovation-comp.one-pager-looks-good  · served Sun Jun 14
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [158] World_Cup_Cleat_Launch.manufacturing-kickoff-3  · served Sun Jun 14
   ✓ expected  event ~"Design Lead Stage Slot" cancelled
     actual    (nothing — cancelled)
     why       cancelled

 ✓ PASS  [159] Marketing-campaign-new-product-delay.lebron-james-cancelled-marketing-campaig  · served Sun Jun 14
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

 ✓ PASS  [160] Day-of-execution_and_Aftermath.thank-the-team  · served Sun Jun 14
   ✓ expected  event ~"thank" @ Tue Aug 11
     actual    "Launch team thank-you lunch" Tue Aug 11 12 PM (90m)
     why       matched

 ✗ FAIL  [161] Marketing-campaign-new-product-delay.ronaldo-s-availability-for-marketing-cam  · served Sun Jun 14
   ✗ expected  event ~"Giano Ronaldo marketing campaign " @ Wed Aug 05
     actual    (nothing matching created)
     why       no event titled like "Giano Ronaldo marketing campaign " was created

 ✓ PASS  [162] Day-of-execution_and_Aftermath.metrics-readout  · served Sun Jun 14
   ✓ expected  to-do ~"metrics" @ by Thu Aug 13
     actual    "Compile launch metrics readout for board/investors" Thu Aug 13 5 PM
     why       matched

 ✗ FAIL  [163] Marketing-campaign-new-product-delay.serena-williams-reschedule  · served Sun Jun 14
   ✗ expected  event ~"marketing" @ Sat Sep 19
     actual    "Serena Williams marketing campaign shoot" Sat Sep 19 10 AM (120m); "Michael Phelps marketing campaign shoot" Tue Sep 15 10 AM (120m); "Ronaldo marketing campaign shoot" Wed Aug 05 10 AM (120m)
     why       found 3 matching, expected exactly 1 (duplicate / double-booked)

── day 15 · Mon Jun 15 · 3 new email(s) ──
   tools  list_new_emails, get_email, create_event
   model said  “Done. Processed 3 emails today:  **Events created (3):** - LeBron James marketing campaign shoot (Fri, Aug 21 - rescheduled) - PR discussion - LeBron James cancellation (Tue, Jun 16 with COMMS) - Boar…”

 ✗ FAIL  [164] Marketing-campaign-new-product-delay.able-to-get-lebron-james-to-reschedule  · served Mon Jun 15
   ✗ expected  event ~"LeBron James marketing campaign scheduled" @ Fri Aug 21
     actual    (nothing matching created)
     why       no event titled like "LeBron James marketing campaign scheduled" was created

 ✗ FAIL  [165] Marketing-campaign-new-product-delay.lebron-james-canceled  · served Mon Jun 15
   ✗ expected  event ~"cancellation" @ Tue Jun 16
     actual    "LeBron James marketing campaign shoot" Fri Aug 21 10 AM (120m); "PR discussion - LeBron James cancellation" Tue Jun 16 9 AM (60m)
     why       found 2 matching, expected exactly 1 (duplicate / double-booked)

 ✗ FAIL  [166] World_Cup_Cleat_Launch.board-sync-on-the-credit-issue  · served Mon Jun 15
   ✗ expected  event ~"board" @ Thu Jun 18
     actual    "Board greenlight call - WC cleat" Mon Jun 08 9 AM (60m); "Board sync - design credit issue" Thu Jun 18 9 AM (60m)
     why       found 2 matching, expected exactly 1 (duplicate / double-booked)

── day 16 · Tue Jun 16 · 1 new email(s) ──
   tools  list_new_emails, get_email
   model said  “Done. Processed 1 email today:  **Email requiring no action:** VP Marketing confirms all athletes have confirmed meeting availability with no further scheduling conflicts. All marketing campaign shoot…”

 ✓ PASS  [167] Marketing-campaign-new-product-delay.confirmations-from-all-athletes-and-vp-o  · served Tue Jun 16
   ✓ expected  take no action (FYI / no scheduling needed)
     actual    (nothing)
     why       correctly took no action

══════════════════════════════════
  ●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●●
  SCORE 98/167 (59%)
══════════════════════════════════

```
