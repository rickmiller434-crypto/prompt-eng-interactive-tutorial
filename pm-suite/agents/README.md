# Agents

Claude subagent definitions land here in Phase 2+.

## Planned roster

| Agent | Phase | Role |
|---|---|---|
| `pm-lead` | 2 | Orchestrator; tailoring-aware routing; principle enforcement |
| `pm-tailor` | 2 | Owns the Tailoring Record; produced from the intake questionnaire |
| `pm-principles-auditor` | 2 | Reads artifacts, flags drift from the 12 principles |
| `pm-stakeholders` | 3 | Stakeholder register, engagement matrix, comms plan |
| `pm-team` | 3 | Team charter, RACI, leadership behaviors, retros |
| `pm-approach` | 3 | Lifecycle selection, complexity assessment, dependency map |
| `pm-planning` | 3 | Scope/WBS or backlog, schedule, cost, resources |
| `pm-work` | 4 | Execution, procurement, knowledge management |
| `pm-delivery` | 4 | Acceptance, value realization, quality verification |
| `pm-measurement` | 4 | KPIs, EVM, velocity, forecasts, status reports |
| `pm-uncertainty` | 4 | Risks, issues, change control, resilience |

## Conventions (will apply once agents land)

- Every agent reads `pm-suite/projects/<project>/tailoring.md` before acting.
- Every agent declares which of the 12 principles its action serves.
- Every artifact gets a `Principles Honored` footer.
- Agents refuse to produce artifacts excluded by the Tailoring Record without an explicit override (logged in the decision log).
