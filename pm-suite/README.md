# PM Suite — AI-Assisted Project Delivery (PMBOK 8 Aligned)

A Claude-powered suite for running a project end-to-end with discipline. It pairs **specialized subagents** (one per PMBOK Performance Domain, plus orchestration and audit) with **storage-agnostic artifact templates** that work for predictive, adaptive, and hybrid lifecycles.

## Why this exists

Most "AI for PM" tooling generates one-shot artifacts (a charter here, a status report there) without enforcing the discipline that makes PM work: tailoring, principle adherence, traceability from value to delivery, and continuous measurement. This suite is opinionated about that discipline.

## Foundation

- **The 12 Principles** (`principles/`) — the why. Every agent is required to cite the principles it is honoring when it acts.
- **The 8 Performance Domains** — the what. One specialist subagent per domain (see `agents/` once Phase 2 lands).
- **Tailoring** (`intake/tailoring-questionnaire.md`) — the how. Filled out at project initiation; drives which lifecycle, which artifacts, and which agents engage.
- **Models, Methods, Artifacts** (`models-methods/`, `templates/`) — the toolbox. Used in service of the domains and principles.

## Layout

```
pm-suite/
  agents/                  # Claude subagent definitions (Phase 2+)
  intake/                  # Project intake & tailoring questionnaire
  principles/              # The 12 PM principles + auditor checks
  models-methods/          # Tailoring, estimation, EVM, risk models (Phase 3+)
  templates/
    initiating/            # Charter, business case, stakeholder register, ...
    planning/              # WBS, schedule, cost, risk, quality, comms, RAID, RACI, backlog, ...
    executing/             # Change requests, status reports, sprint plans, ...
    monitoring-controlling/# EVM, variance, risk burndown, velocity, CFD, ...
    closing/               # Lessons learned, benefits realization, closure
  projects/                # Per-project working artifacts (gitignored in real use)
  examples/                # End-to-end worked example (Phase 5)
```

## Lifecycle support

Hybrid by default. Templates ship in two flavors where it matters — **predictive** (baselines, EVM, change control) and **adaptive** (backlog, increments, velocity) — and the tailoring intake determines which the project draws from, or how to blend them.

## Storage

Artifacts are plain markdown under `projects/<project-name>/`. Later phases will add thin adapters to sync to Notion, GitHub Projects, or other systems of record without changing the agents.

## Build phases

1. **Phase 1 (this PR)** — Scaffold, 12 principles, tailoring intake.
2. **Phase 2** — `pm-lead` orchestrator, `pm-tailor`, `pm-principles-auditor`.
3. **Phase 3** — Initiating + Planning domain agents and their templates.
4. **Phase 4** — Executing + Monitoring domain agents and their templates.
5. **Phase 5** — Closing templates + end-to-end worked example.
6. **Phase 6** — Storage adapters (Notion, GitHub Projects), `/pm-*` slash commands.

## Source standard

Aligned with the **PMI PMBOK Guide, 8th Edition** — principle-based, performance-domain organized, tailoring-first. Where the suite makes opinionated choices beyond the standard, those are flagged in the relevant file.
