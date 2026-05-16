# Adapter — Local Markdown (default)

The default. Artifacts live as markdown under `pm-suite/projects/<project>/`. No external system involved.

## Layout

```
pm-suite/projects/<project>/
  tailoring.md
  charter.md
  business-case.md
  stakeholder-register.md
  engagement-matrix.md
  risk-register.md
  product-backlog.md            # or wbs.md / schedule-gantt.md for predictive
  quality-plan.md
  communications-plan.md
  ocm-plan.md
  dependency-map.md
  raci-matrix.md
  team-charter.md
  resource-plan.md
  procurement-plan.md
  raid-log.md
  decisions/
    decision-log.md
  status-reports/
    <YYYY-MM-DD>-status.md
  sprint-records/               # adaptive
    sprint-<n>/plan.md, retro.md, review.md
  audits/
    audit-<YYYY-MM-DD>.md
  closure/
    lessons-learned.md
    benefits-realization.md
    handover-to-operations.md
    closure-report.md
```

## Operations

| Op | Implementation |
|----|----------------|
| push | n/a (already local) |
| pull | n/a |
| link | n/a |
| diff | `git diff` between commits |

## Versioning

Use the project's own git repo. Each artifact change is a commit; significant changes (e.g., re-baselines) are tagged.

## Sensitive content

If artifacts contain customer or regulatory-sensitive content:

- Add `pm-suite/projects/<sensitive-project>/` to `.gitignore` and keep the repo private; **or**
- Replace sensitive values with references to a secret store; **or**
- Push to a private mirror only (and use the Notion or GitHub Projects adapter for sharing visible-but-redacted views).

Coordinate with `pm-stewardship` requirements in `principles/01-stewardship.md`.
