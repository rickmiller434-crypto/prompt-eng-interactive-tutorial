# Tailoring Intake & Record — Acme Onboarding Revamp

## Section A — Project identity

1. **Project name:** Acme Onboarding Revamp
2. **Sponsor:** Priya Shah, VP Customer Success
3. **PM / delivery lead:** Marcus Lin, Senior PM
4. **One-line description:** Redesign the new-customer onboarding flow to cut time-to-first-value and reduce 30-day drop-off.
5. **Desired future state:** A new customer signs up, completes guided onboarding in under a week, and is actively using the product daily by day 7. The Customer Success team spends its time on expansion conversations, not chasing stuck onboardings.
6. **Hard deadline or target window:** Soft target end of May 2026 (before summer freeze). Hard deadline: regulator compliance changes land 1 Sep 2026 — anything we ship needs to comply.

## Section B — Value & success

7. **Primary value:** Revenue (retention + expansion) and cost reduction (CS time per onboarding).
8. **Top 3 measurable benefits:**
   - Median time-to-first-value: baseline 14 days → target ≤ 5 days, measured 60 days post-launch.
   - 30-day drop-off rate: baseline 28% → target < 15%, measured at 90 days post-launch.
   - CS hours per onboarding: baseline 6.4 hr → target ≤ 2.0 hr, measured at 60 days post-launch.
9. **Success at 30 / 90 / 365 days:** 30d — new flow live for all new customers. 90d — drop-off measurably down. 365d — expansion revenue from cohorts onboarded post-launch ≥ 15% higher than pre-launch cohorts.
10. **Clear failure:** Drop-off rate unchanged or worse, and CS hours not reduced.

## Section C — Stakeholders

11. **Sponsor engagement:** Sets vision; biweekly steering; approves baseline changes.
12. **Primary users:** Mid-market customers (5–50 seats). Admins are primary persona.
13. **Regulators / external:** SOC 2, GDPR (EU customers), upcoming 1 Sep 2026 data-residency rule.
14. **Affected teams:** Sales (handoff process changes), CS (their day-to-day changes), Engineering (shared infra), Data (instrumentation).
15. **Controversy / opposition:** Moderate — CS team initially anxious about role change.

## Section D — Scope & uncertainty

16. **Requirements stability:** Mostly stable for compliance and instrumentation; volatile for UX flows.
17. **Technology novelty:** Mostly proven (we have the components); one novelty: in-product video walkthroughs.
18. **Solution clarity:** We know the problem; we are still framing parts of the solution (which UX patterns).
19. **Anticipated increments:** ~8 two-week increments.
20. **Anticipated change rate:** Medium (UX work will iterate; compliance work will not).

## Section E — Size & resources

21. **Team size:** 8 core + 3 extended (Legal, Data, Marketing).
22. **Duration:** 4 months.
23. **Budget tier:** M ($300k–$600k for this org).
24. **Geographic distribution:** Hybrid; two timezones (US ET + UK).
25. **Domain experience:** High for the engineering team, mixed for design (newer hires).

## Section F — Constraints & obligations

26. **Regulatory:** SOC 2, GDPR; 1 Sep 2026 data-residency requirement.
27. **Contractual:** None project-driving; existing customer contracts referenced for SLAs.
28. **Security / data classification:** Customer PII; standard internal controls apply.
29. **Sustainability / ESG:** N/A material.
30. **Health & safety:** N/A.

## Section G — Delivery cadence

31. **Delivery mode:** Phased — internal beta at week 8, limited external at week 12, GA at week 16.
32. **Acceptable batch size:** Per increment to internal, gated rollout to external.
33. **Customer ability to absorb change:** Medium — needs comms + in-product guidance.

## Section H — Organizational context

34. **PM maturity:** Defined.
35. **Mandated methodology:** None; team uses Scrum-flavored for product work.
36. **Mandatory tools:** Jira (work tracking), Notion (docs), GitHub (code).
37. **Reporting cadence:** Biweekly to steering, weekly to Product/Eng leadership.

---

## Tailoring Recommendation (pm-tailor)

### Recommended lifecycle

- [ ] Predictive
- [ ] Adaptive
- [x] **Hybrid**

**Rationale:** Section D Q16 splits cleanly — compliance/instrumentation work is stable, UX work is volatile. Section F Q26 imposes a hard 1 Sep regulatory date that argues for predictive governance over the compliance workstream. Section G Q31 supports phased delivery. Section E Q25 (mixed design experience) argues for short feedback loops on UX, which suits adaptive delivery. Overall: **adaptive delivery in 2-week increments inside a predictive governance overlay**, with the compliance workstream run as a separate predictive plan with hard milestones.

### Engaged agents

| Agent | Engaged? | Notes |
|---|---|---|
| pm-lead | ✅ | Orchestrator |
| pm-tailor | ✅ | This record; re-validate at week 8 (internal beta) and week 12 (external beta) |
| pm-principles-auditor | ✅ | At week 4 and at each beta gate |
| pm-stakeholders | ✅ | High — CS-team change management is critical |
| pm-team | ✅ | New cross-functional team; charter needed at kickoff |
| pm-approach | ✅ | Confirm hybrid; map dependencies on shared infra |
| pm-planning | ✅ | Backlog for UX/product; predictive plan for compliance + instrumentation |
| pm-work | ✅ | Sprint execution + procurement (video tooling vendor) |
| pm-delivery | ✅ | DoD, acceptance, value trace |
| pm-measurement | ✅ | Velocity + burn-up for adaptive; SPI/CPI for compliance workstream; benefits dashboard |
| pm-uncertainty | ✅ | Risk register; CS-resistance and 1 Sep regulatory date are headline risks |

### Required artifacts

Initiating: charter, business-case, stakeholder-register, assumption-constraint-log.

Planning (adaptive parts): product-backlog, release-plan, communications-plan, engagement-matrix, raid-log, raci-matrix, resource-plan, team-charter, ocm-plan, dependency-map, risk-register, quality-plan.

Planning (predictive — compliance workstream only): wbs, schedule-gantt, cost-baseline, procurement-plan (for video tooling).

Executing / Monitoring: change-request, decision-log, status-report-adaptive (project), status-report (compliance workstream — with EVM-lite), velocity-tracker, burn-up, risk-burndown, retro (every increment).

Closing: lessons-learned, benefits-realization, handover-to-operations, closure-report.

### Cadence

- Status reporting: biweekly (project), weekly (compliance workstream)
- Retros: every 2 weeks
- Tailoring re-validation: week 8 (internal beta), week 12 (external beta)
- Risk review: weekly
- Steering: biweekly

### Principles in tension (and resolution)

- **Quality (8) vs. Adaptability (11):** the UX work will iterate, but the DoD is fixed (accessibility, instrumentation, accept-criteria coverage) regardless of which iteration. Resolution: lock DoD at kickoff; let scope per increment flex.
- **Value (4) vs. Stewardship (1):** speed of value pulls against the regulatory deadline. Resolution: the compliance workstream is a non-negotiable gate; adaptive product work that touches data flows must clear compliance review before each release.
- **Stakeholders (3) vs. Tailoring (7):** CS team anxiety could pull us toward over-communicating, slowing delivery. Resolution: an explicit OCM plan with a small set of high-quality touchpoints rather than ad-hoc reassurance.

---

## Sign-off

- Tailoring author: pm-tailor (reviewed by Marcus Lin, PM)
- Date: 2026-01-22
- Next re-validation due: 2026-03-19 (week 8, internal beta)
- Approved by sponsor: Priya Shah, 2026-01-26
