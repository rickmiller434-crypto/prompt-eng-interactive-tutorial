---
description: Run a retrospective and capture actions
argument-hint: <project> <sprint-or-period-id>
---

Run a retrospective for `pm-suite/projects/$1/`, sprint/period `$2`.

Steps:

1. Read the prior retro's actions (most recent file in `sprint-records/sprint-N/retro.md` or equivalent).
2. Invoke the `pm-team` subagent via the `Agent` tool to facilitate the retro structure.
3. Ask the user (or whichever participants are present in the session) for:
   - **Review of prior actions** — done / in progress / dropped (with reason)
   - **What went well** (specific, cited)
   - **What did not go well** (specific, cited; above-the-line)
   - **What we'll change** — at least one action with owner and due date
4. Capture process / quality signals (rework %, defect escape, anything triggering variance analysis).
5. Write to `pm-suite/projects/$1/sprint-records/sprint-$2/retro.md` (or appropriate path) using `pm-suite/templates/monitoring-controlling/retro.md` as the template.

If the prior retro had unfinished actions with no progress, flag them — that pattern is itself a retro topic.

Cite Principle 2 (Team), 6 (Leadership), 8 (Quality), 11 (Adaptability).
