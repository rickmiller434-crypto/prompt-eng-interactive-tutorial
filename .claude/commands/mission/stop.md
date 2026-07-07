---
description: Stop the mission loop cleanly and record why
argument-hint: [optional reason]
---

Stop the mission loop.

1. If a /loop or scheduled wakeup is driving rounds, cancel it (ScheduleWakeup
   with stop: true, or delete the pending trigger).
2. In MISSION.md, set Status to STOPPED and append a final Round log entry:
   `### Stopped — <date>` with the reason ($ARGUMENTS if given, otherwise
   "stopped by user").
3. Commit and push MISSION.md so the stop is durable, then summarize for the
   user: how many rounds ran, which Definition of Done items are checked, and
   what remains if they relaunch later.

Do not delete the Round log or reset the file — a stopped mission can be
resumed by setting Status back to ACTIVE and running /mission:launch.
