---
description: Generate a status report for the current reporting period
argument-hint: <project>
---

Generate a status report for `pm-suite/projects/$ARGUMENTS/`.

Steps:

1. Read `tailoring.md` to determine lifecycle (predictive / adaptive / hybrid) and reporting cadence.
2. Read the latest artifacts: backlog or WBS, risk register, decision log, prior status report (if any), and the latest retro.
3. Invoke the `pm-measurement` subagent with `Agent` tool.
4. Have it produce the appropriate template:
   - Predictive / hybrid-governance → `pm-suite/templates/executing/status-report.md`
   - Adaptive / hybrid-delivery → `pm-suite/templates/executing/status-report-adaptive.md`
5. Save to `pm-suite/projects/$ARGUMENTS/status-reports/<YYYY-MM-DD>-status.md`.

The report must include:
- Vision restated and top 3 priorities (Principle 6)
- Outputs delivered AND outcomes (benefits realization) (Principle 4)
- Honest forecast with confidence interval (no point estimates) (Principle 10)
- Variance analysis link if any threshold breached (Principle 8)
- Top 5 risks with delta (Principle 10)
- Decisions needed (Principle 6)

If `pm-measurement` reports a threshold breach (SPI/CPI/velocity drop), tell the user they should also run a variance analysis (point them at `pm-suite/templates/monitoring-controlling/variance-analysis.md`) — do not produce it automatically.

Cite principles 4 (Value), 6 (Leadership), 8 (Quality), 10 (Risk).
