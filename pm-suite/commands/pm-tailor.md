---
description: Invoke pm-tailor to produce or revalidate the Tailoring Record from a filled intake
argument-hint: <project>
---

Invoke the `pm-tailor` subagent on `pm-suite/projects/$ARGUMENTS/tailoring.md`.

If the file does not exist, tell the user to run `/pm-init $ARGUMENTS` first.

If Sections A–H are not at least minimally filled, list the missing answers and stop.

Otherwise:

1. Use the `Agent` tool with `subagent_type: pm-tailor`. Pass the file path and the goal: produce (or revalidate) the Tailoring Recommendation section per the rules in the agent definition.

2. After the agent returns, write the updated `tailoring.md`, then summarize the lifecycle choice, the engaged-agents table, and any principles-in-tension entries.

3. If this is a revalidation (the file already has a Tailoring Recommendation), also append a decision-log entry capturing what changed and why.

Cite principles 7 (Tailoring), 5 (Systems thinking), 9 (Complexity).
