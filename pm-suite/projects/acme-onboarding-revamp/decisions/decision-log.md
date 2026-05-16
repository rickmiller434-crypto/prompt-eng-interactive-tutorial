# Decision Log — Acme Onboarding Revamp

| Date | ID | Decision | Rationale | Alternatives | Level | Steward | Affected | Principles | Linked artifacts |
|------|----|---------|-----------|--------------|-------|---------|----------|------------|-------------------|
| 2026-01-29 | D-001 | Adopt hybrid lifecycle: adaptive product delivery + predictive compliance workstream | Section D answers split between volatile UX and stable compliance; hard 1 Sep regulatory date | All-adaptive (rejected — compliance gating needed); all-predictive (rejected — UX iteration needed) | Sponsor | Marcus | Whole team | 7, 9, 11 | tailoring.md |
| 2026-02-12 | D-002 | Jordan Park co-owns the OCM workstream | Engages a resistant-leaning stakeholder; reduces R1 exposure; aligns with Principle 12 | Marcus owns alone (rejected — Jordan's authority and credibility with CSMs matters) | PM | Marcus | CS team, Jordan | 3, 12 | engagement-matrix.md; risk-register.md (R1) |
| 2026-02-19 | D-003 | In-product videos limited to EN + DE at GA; ES and FR deferred to fast-follow | Localization capacity constraint; mid-market segment 78% EN/DE; revisit at 60d review | Localize all four at GA (rejected — slips schedule); EN-only (rejected — DE customer base 23%) | PM | Marcus | Customer Advisory; Localization | 7, 11 | product-backlog.md (OB-4) |
| 2026-02-26 | D-004 | Approve CR-001: add segmentation guardrails (OB-10) before internal beta | Mitigates R8; cheap to add; protects out-of-scope customers | Defer (rejected — risk realized would be high-visibility) | PM | Marcus | Eng | 5, 10 | risk-register.md (R8); CRs/CR-001 |
| 2026-03-05 | D-005 | Pair Lin Wei with engineer Kemi for 4 weeks; create instrumentation runbook | Reduces R7 (single-point-of-failure); aligns with Principle 11 | Hire faster (rejected — no candidate ready); accept SPoF (rejected — material impact) | PM | Marcus | Data team; Eng | 11 | risk-register.md (R7); resource-plan.md |
| 2026-03-12 | D-006 | Push external-beta opt-in (OB-12) to internal beta + 2 weeks rather than at internal beta | Internal beta workload heavier than estimated; protects quality | Hold opt-in mechanism at internal beta (rejected — risked DoD slip) | PM | Marcus | Sales; Customer Advisory | 8, 11 | product-backlog.md (OB-12); status-reports/ |

## Conventions

- Append new decisions at the bottom.
- Decisions that overturn earlier ones reference the prior ID and explain the change.
- Tailoring updates, baseline re-baselines, and approved CRs each produce an entry.
