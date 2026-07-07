# Mission loop

A long-running autonomous work loop for Claude Code built entirely from
built-in features — file-based state, project slash commands, a custom
subagent, and the built-in /loop skill. No third-party plugins, no external
code execution, no headless `claude -p`.

It addresses the two ways long autonomous runs go wrong:

1. **The goal decays across auto-compaction.** The mission lives in
   `MISSION.md` on disk, not in the transcript. Every round starts by
   re-reading it in full, so compaction can summarize the conversation all it
   wants — the goal is reloaded verbatim each round.
2. **The agent drifts and rationalizes.** After every round, an isolated
   `mission-advisor` subagent (`.claude/agents/mission-advisor.md`) audits the
   work against the mission in its own context. It is instructed to distrust
   the working agent's summary, verify claims against the actual repo state,
   and attack any claim of completion. Its findings drive the next round.

## Usage

```
/mission:define     # interview → writes MISSION.md, sets Status: ACTIVE
/mission:launch     # preflight, then repeats /mission:round via /loop
/mission:stop       # cancel the loop, mark STOPPED, keep the audit trail
```

Each `/mission:round` is: reload mission → do the single highest-value step →
commit → advisor audit → log the verdict in MISSION.md's Round log. The loop
ends on its own only after **two consecutive** `NOTHING_TO_ADD` verdicts with
every Definition of Done box checked.

## Portability

The harness is just these files. To use it in another repository, copy:

- `.claude/agents/mission-advisor.md`
- `.claude/commands/mission/` (define, launch, round, stop)
- `MISSION.md` (reset to the template state, Status: DRAFT)
