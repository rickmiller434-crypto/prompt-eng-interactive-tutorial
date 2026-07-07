---
description: Interview the user and write MISSION.md for the mission loop
argument-hint: [optional one-line mission statement to start from]
---

You are defining a mission for the long-running mission loop. The output of
this command is a completed MISSION.md — nothing else. Do not start doing the
mission's work.

1. Read MISSION.md. If its Status is ACTIVE, stop and tell the user to run
   /mission:stop first — never overwrite a live mission.

2. Gather what you need from the user, using $ARGUMENTS as the starting point
   if provided. You need four things, and you should push back until each is
   concrete:
   - **Objective**: one paragraph a stranger could judge success by.
   - **Definition of Done**: 3–8 checklist items, each verifiable by a command
     or by inspecting a file — reject vague items like "code is clean" and ask
     the user to restate them as something checkable.
   - **Constraints**: hard rules (branch, compatibility, budget, style).
   - **Non-goals**: adjacent work that is explicitly out of scope. If the user
     can't name any, propose two or three likely drift directions yourself and
     confirm them.

3. Rewrite MISSION.md with those answers, keeping the file's structure and
   comments intact. Set Status to ACTIVE. Leave the Round log empty.

4. Read the finished file back to the user in summary form and tell them to
   start the loop with /mission:launch.
