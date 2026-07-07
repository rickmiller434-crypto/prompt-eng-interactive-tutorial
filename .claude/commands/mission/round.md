---
description: Run one round of the mission loop - reload mission, work, isolated advisor review, log
---

Run exactly one round of the mission loop. Follow every step in order. Do not
skip the advisor step, even if you are certain the round went well — the whole
point of the advisor is that your certainty is not evidence.

## 1. Reload the mission (compaction-proofing)

Read MISSION.md in full, even if you believe you remember it — after an
auto-compaction your memory of it is a summary, and this step is what keeps
the goal from decaying.

- If Status is COMPLETE or STOPPED: report that and end the turn. If this
  round was triggered by /loop, stop the loop (ScheduleWakeup with stop: true).
- If Status is DRAFT: tell the user to run /mission:define first, and stop.

## 2. Work

From the Definition of Done and the latest Round log entry's carry-over, pick
the **single highest-value next step**. Prefer the advisor's top finding from
last round over your own new ideas. Then do it.

- Stay inside the Constraints; touch nothing on the Non-goals list.
- Verify what you did with a real command (tests, build, running the thing) —
  record the command, because the advisor will re-run it.
- Commit the round's work on the designated branch with a clear message. One
  round, one commit (or a few coherent ones) — the advisor reads the diff.

## 3. Advisor review (isolated context)

Spawn the `mission-advisor` agent (Agent tool, subagent_type: mission-advisor,
run_in_background: false) and wait for its verdict. Its prompt must contain:

- The full current text of MISSION.md.
- A plain factual summary of this round: what you changed, which files, what
  command you ran to verify and its result.
- The instruction: "Audit this round against the mission. Do not trust the
  summary — verify against the repository itself."

Do NOT include your opinion of how well the round went, and do NOT argue with
the verdict afterward. If a finding seems wrong, the next round can disprove
it with evidence.

## 4. Log and act

Append a Round log entry to MISSION.md in the format described there,
including the advisor's verdict. Never rewrite old entries.

- Findings that are quick, unambiguous fixes: fix them now, note it in the log.
- Larger findings: record the top one as the carry-over for the next round.
- If the advisor returned NOTHING_TO_ADD **and** the previous round's entry
  was also NOTHING_TO_ADD **and** every Definition of Done box is checked:
  set Status to COMPLETE, commit and push, and stop the loop. Two consecutive
  clean audits are required — one can be luck.

## 5. Hand off

End the turn with a two-or-three sentence summary: what this round did, the
advisor's verdict, and what the next round will do. If running under /loop,
let the next tick fire — do not start another round in the same turn.
