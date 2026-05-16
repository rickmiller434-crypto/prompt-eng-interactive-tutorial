# Product Backlog — Acme Onboarding Revamp

- **Owner:** pm-planning (with Sam Reyes, Product)
- **Last refined:** 2026-03-12

## Backlog (representative slice)

| ID | Title | As a … I want … so that … | Benefit | Stakeholder | Priority (WSJF) | Size | Status | Acceptance criteria summary |
|----|-------|---------------------------|---------|-------------|------------------|------|--------|------------------------------|
| OB-1 | Guided sign-up wizard | As a new admin, I want a 3-step setup so that I can connect my first data source within 10 minutes | B1 (TTFV) | Customer Advisory | 32 | M | Done | Wizard completes in ≤ 3 steps; ≥ 80% finish in usability test |
| OB-2 | In-product checklist with first-value milestones | As an admin, I want a visible "next 5 minutes / next 30 minutes" checklist so that I know what to do | B1, B2 | Customer Advisory | 28 | M | Done | Checklist persists; clicks instrumented |
| OB-3 | Sample data set | As an admin, I want a one-click sample data set so that I can see value before integrating | B1 | Customer Advisory | 24 | S | Done | One-click load; tagged "sample"; removable in 1 click |
| OB-4 | Embedded video walkthroughs (3 core flows) | As an admin, I want short videos at decision points so that I don't need a kickoff call | B1, B3 | Customer Advisory; CS | 20 | L | In sprint | Accessibility ✓; ≤ 90s each; localized for EN + DE |
| OB-5 | Compliance: data-residency selection at sign-up | As an admin, I want to select my data region so that my data is stored compliantly | Regulatory | Aisha; GDPR | 30 | M | In sprint | Region locked after first write; audit-logged |
| OB-6 | Instrumentation: TTFV, drop-off, CS minutes per onboarding | As a stakeholder, I want these measured so that we can verify benefit realization | All | Data team | 25 | M | Done | Events firing; dashboard live; data validated by Lin Wei |
| OB-7 | CSM dashboard: stalled onboardings + suggested action | As a CSM, I want to see who needs help so that I act on high-value cases | B3 | Jordan; CSMs | 18 | M | Ready | Dashboard updates daily; ≤ 2 click drill-down |
| OB-8 | Sales→CS handoff redesign | As a CSM, I want a structured handoff so that I'm not chasing context | B3 | Sales; CS | 12 | S | Ready | Handoff form ≤ 5 fields; auto-populated from CRM |
| OB-9 | Empty-state coaching for inactive admins (day 3, 7, 14) | As an admin who has lapsed, I want gentle nudges so that I come back | B2 | CS | 14 | S | Ready | 3 messages; opt-out works; instrumented |
| OB-10 | Mid-market segmentation guardrails | As the team, I want the new flow scoped to mid-market only so that we don't break out-of-scope segments | (Risk R8 mitigation) | Eng | 22 | S | Done | Segment check at entry; out-of-scope routed to legacy flow |
| OB-11 | OCM: CS sandbox week + 3 workshops | As a CSM, I want to practice in the new flow so that I feel ready | B3 (people) | Jordan; CS Team | 26 | M | In sprint | Sandbox env live; 3 workshops scheduled; attendance ≥ 90% |
| OB-12 | External-beta opt-in mechanism | As a customer, I want to opt into the new onboarding so that I can try it on next signup | (Release approach) | Customer Advisory | 16 | S | Ready | Toggle in admin; feature-flagged; reversible |

(Many more items exist below the cut-line in the real backlog.)

## Definition of Ready

An item is **Ready** when it has: clear user/outcome, acceptance criteria, size estimate, linked benefit, no unresolved blocking dependencies.

## Definition of Done

Per the quality plan: acceptance criteria verified, accessibility ✓, instrumentation ✓, docs updated, deployable, accepted by named acceptor.

## Prioritization

**WSJF** = (Business Value + Time Criticality + Risk Reduction) / Job Size. Inputs from Priya (BV) and Sam/Marcus (TC, RR). Refreshed each refinement.

## Hygiene

Refinement weekly. "New" items aged > 4 weeks are reviewed and dropped or sized.

---

**Principles Honored:** 4 (value), 8 (quality), 11 (adaptability & resiliency).
