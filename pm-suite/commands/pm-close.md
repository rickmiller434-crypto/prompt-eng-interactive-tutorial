---
description: Run the closure workflow — lessons learned, benefits realization, handover, closure report
argument-hint: <project>
---

Run the project closure workflow for `pm-suite/projects/$ARGUMENTS/`.

This is a multi-step workflow. Confirm with the user before each step.

## Step 1 — Final principle audit

Run `/pm-audit $ARGUMENTS`. Surface any remaining severity-3 or severity-4 findings. The user decides whether to remediate before closure or to record persistent items as handover obligations.

## Step 2 — Lessons learned

Invoke `pm-delivery` to produce `pm-suite/projects/$ARGUMENTS/closure/lessons-learned.md` using the template at `pm-suite/templates/closing/lessons-learned.md`. Inputs:
- All retros under `sprint-records/`
- All variance analyses
- All audit reports
- The risk register's realized-risk log
- Stakeholder feedback summaries

## Step 3 — Benefits realization plan

Invoke `pm-delivery` + `pm-measurement` to produce `pm-suite/projects/$ARGUMENTS/closure/benefits-realization.md`. Must include:
- Benefits register at handover with current values
- Measurement plan beyond closure (cadence, owner, audience per benefit)
- Trigger conditions for re-engagement
- Named sustainment owners

## Step 4 — Handover to operations

Invoke `pm-delivery` + `pm-work` to produce `pm-suite/projects/$ARGUMENTS/closure/handover-to-operations.md`. Must include:
- Inventory of systems and documents
- Runbooks
- Support model and known issues
- Named acceptance from operations owner

## Step 5 — Closure report

Invoke `pm-lead` (with `pm-measurement` and `pm-delivery`) to produce `pm-suite/projects/$ARGUMENTS/closure/closure-report.md`. Must reference the prior closure artifacts and the final audit.

## Step 6 — Sign-off

Collect named sign-offs (sponsor, PM, operations owner, benefits owner) into the closure report.

Cite Principle 1 (Stewardship), 4 (Value), 8 (Quality), 12 (Change).
