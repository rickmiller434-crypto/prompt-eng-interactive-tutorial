# Agents

Claude subagent definitions. Each agent is a self-contained markdown file with YAML frontmatter (`name`, `description`, `tools`) and a system prompt body, in the format expected by Claude Code subagents.

## Roster

| Agent | Status | Role |
|---|---|---|
| `pm-lead` | ✅ Phase 2 | Orchestrator; tailoring-aware routing; principle enforcement |
| `pm-tailor` | ✅ Phase 2 | Owns the Tailoring Record |
| `pm-principles-auditor` | ✅ Phase 2 | Audits artifacts against the 12 principles |
| `pm-stakeholders` | ✅ Phase 3 | Stakeholder register, engagement matrix, comms plan |
| `pm-team` | ✅ Phase 3 | Team charter, RACI, leadership behaviors, retros |
| `pm-approach` | ✅ Phase 3 | Lifecycle selection, complexity assessment, dependency map |
| `pm-planning` | ✅ Phase 3 | Scope/WBS or backlog, schedule, cost, resources |
| `pm-work` | ✅ Phase 4 | Execution, procurement, knowledge management |
| `pm-delivery` | ✅ Phase 4 | Acceptance, value realization, quality verification |
| `pm-measurement` | ✅ Phase 4 | KPIs, EVM, velocity, forecasts, status reports |
| `pm-uncertainty` | ✅ Phase 4 | Risks, issues, change control, resilience |

## Conventions

- Every agent reads `pm-suite/projects/<project>/tailoring.md` before acting.
- Every agent declares which of the 12 principles its action serves.
- Every artifact gets a `Principles Honored` footer.
- Agents refuse to produce artifacts excluded by the Tailoring Record without an explicit override (logged in the decision log).

## Using these as live Claude Code subagents

These files are written in Claude Code's subagent format. To make them callable as subagents from a Claude Code session:

```bash
mkdir -p .claude/agents
cp pm-suite/agents/*.md .claude/agents/
```

After that, invoke them via the `Agent` tool with `subagent_type: pm-lead` (or any other name from the roster).

The suite keeps the source of truth in `pm-suite/agents/` so it ships with the project; `.claude/agents/` is the local runtime location.
