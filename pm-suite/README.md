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
  agents/                  # 11 Claude subagent definitions (pm-lead, pm-tailor, …)
  commands/                # /pm-* slash commands for Claude Code
  adapters/                # Storage adapters (local-markdown, notion, github-projects)
  intake/                  # Project intake & tailoring questionnaire
  principles/              # The 12 PM principles + auditor checks
  models-methods/          # Models, methods, heuristics (extensible)
  templates/
    initiating/            # Charter, business case, stakeholder register, ...
    planning/              # WBS, schedule, cost, risk, quality, comms, RAID, RACI, backlog, ...
    executing/             # Change requests, status reports, sprint plans, ...
    monitoring-controlling/# EVM, variance, risk burndown, velocity, CFD, ...
    closing/               # Lessons learned, benefits realization, handover, closure
  projects/                # Per-project working artifacts (includes worked example)
  examples/                # Pointer to projects/acme-onboarding-revamp/
```

## Lifecycle support

Hybrid by default. Templates ship in two flavors where it matters — **predictive** (baselines, EVM, change control) and **adaptive** (backlog, increments, velocity) — and the tailoring intake determines which the project draws from, or how to blend them.

## Storage

Artifacts are plain markdown under `projects/<project-name>/`. Later phases will add thin adapters to sync to Notion, GitHub Projects, or other systems of record without changing the agents.

## Build phases (all complete)

1. ✅ **Phase 1** — Scaffold, 12 principles, tailoring intake.
2. ✅ **Phase 2** — `pm-lead`, `pm-tailor`, `pm-principles-auditor`.
3. ✅ **Phase 3** — Initiating + Planning agents (`pm-stakeholders`, `pm-team`, `pm-approach`, `pm-planning`) and their 16 templates.
4. ✅ **Phase 4** — Executing + Monitoring agents (`pm-work`, `pm-delivery`, `pm-measurement`, `pm-uncertainty`) and their 14 templates.
5. ✅ **Phase 5** — Closing templates and the end-to-end worked example (`projects/acme-onboarding-revamp/`).
6. ✅ **Phase 6** — Storage adapters (`adapters/`) and `/pm-*` slash commands (`commands/`).

## Getting started

```bash
# Activate the subagents and slash commands
mkdir -p .claude/agents .claude/commands
cp pm-suite/agents/*.md .claude/agents/
cp pm-suite/commands/*.md .claude/commands/

# In a Claude Code session in this repo, start a new project:
# /pm-init my-new-project
# (then fill out Sections A–H of the tailoring intake)
# /pm-tailor my-new-project
# /pm-status my-new-project   (later, once execution is underway)
# /pm-audit my-new-project    (at phase gates)
# /pm-close my-new-project    (at closure)
```

See `projects/acme-onboarding-revamp/README.md` for a full worked example.

## Source standard

Aligned with the **PMI PMBOK Guide, 8th Edition** — principle-based, performance-domain organized, tailoring-first. Where the suite makes opinionated choices beyond the standard, those are flagged in the relevant file.
