---
description: Preflight-check MISSION.md and start the recurring mission loop
argument-hint: [optional interval, e.g. 15m — defaults to dynamic pacing]
---

Start the mission loop. This command only launches; the per-round logic lives
in /mission:round.

1. **Preflight.** Read MISSION.md in full and refuse to launch unless all of
   these hold (tell the user exactly what's missing otherwise):
   - Status is ACTIVE (run /mission:define first if DRAFT).
   - The Objective is filled in and the Definition of Done has at least one
     unchecked, concretely verifiable item.
   - You are on the designated feature branch with a clean working tree.

2. **Launch.** Start the recurring loop with the built-in /loop skill,
   passing /mission:round as the command to repeat. If the user gave an
   interval in $ARGUMENTS, use it (e.g. `/loop 15m /mission:round`);
   otherwise let /loop pace itself dynamically. If the /loop skill is not
   available in this environment, fall back to running /mission:round
   directly and, at the end of each round, scheduling the next one with
   whatever self-scheduling tool exists (ScheduleWakeup or send_later);
   if none exists, run rounds back-to-back in this session until the stop
   condition in /mission:round step 4 is met.

3. Tell the user the loop is running, how to watch it (the Round log in
   MISSION.md is the audit trail), and that /mission:stop ends it at any time.
