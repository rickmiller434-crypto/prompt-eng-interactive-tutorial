# Lessons Learned — Acme Onboarding Revamp

- **Owner:** pm-delivery
- **Compiled at:** closure — 2026-05-21
- **Sources:** 8 sprint retros, 1 mid-flight audit, 1 closing audit, all 5 variance analyses, full risk register.

## What we'd do again

| # | Lesson | Why it worked | Reuse |
|---|--------|----------------|--------|
| 1 | Co-own OCM with a respected stakeholder from the affected team (Jordan as CS co-owner) | Turned the largest threat (R1) from "manage at distance" to "owned from inside" | Any project changing how a team works |
| 2 | Adaptive delivery inside predictive governance | The hybrid let UX iterate while compliance had a hard backbone | Any project with mixed-stability workstreams + hard external dates |
| 3 | Instrumentation as a backlog item, not an afterthought (OB-6 early) | Benefits measurement was ready at internal beta, not invented at GA | Any value-driven project |
| 4 | Sandbox week before workshops | CSMs landed at workshops with hands-on context; engagement was real | Any OCM-heavy rollout |
| 5 | Pair newer experts with engineers to retire SPoF (Lin/Kemi) | Closed R7 inside 4 weeks; runbook now reusable | Whenever a critical role has a single person |

## What we'd change

| # | Lesson | What we did | What we'd do instead | Why |
|---|--------|--------------|-----------------------|------|
| 1 | DoD did not initially name accessibility evidence | Treated it as implicit; auditor flagged at week 7 | Make accessibility-evidence requirement explicit in DoD from day 1 | We did the work anyway, but a late-flagged gap creates avoidable rework risk |
| 2 | Compliance buffer was only inside management reserve | Surfaced after audit; rebuilt mid-flight | Surface workstream-specific buffers in `schedule-gantt.md` at baseline | Hidden buffers can't be defended in steering |
| 3 | Initial complexity assessment not re-done per workstream | Single project-level assessment carried for 7 weeks | Re-profile per workstream at every phase gate | Compliance complexity changed once regulator issued new guidance; we should have caught it earlier |
| 4 | Adoption metrics defined late | Defined at week 7 audit remediation | Define adoption metrics at planning, alongside delivery metrics | Pre-launch trend signal would have helped sponsor reporting |
| 5 | DE localization later than ideal | Carried over from Sprint 4 to Sprint 5 | Plan localization in parallel with EN copy, not after | 1 week of buffer lost; not a slip but tighter than needed |

## What surprised us

| # | Surprise | What it revealed |
|---|----------|------------------|
| 1 | Customer Advisory feedback flipped a UX assumption — they preferred sample data before integration, not after | We over-weighted "real data first" heuristic; the model of the customer was wrong |
| 2 | R6 (sales asset opportunity) materialized faster than expected; marketing wanted assets at limited beta | Successful UX work creates secondary obligations downstream; plan for them in OCM, not just delivery |
| 3 | CSMs' biggest concern was not their job — it was being asked to support a flow they didn't understand | The OCM plan we built fit; the OCM plan we initially imagined (job-security-focused) would not have |

## Pattern findings (Principle 9 — Complexity)

The compliance workstream behaved as "complicated" until the regulator issued clarifying guidance in March — at which point it briefly became "complex" until our team converged on a residency design. **Leading indicator we missed:** the regulator's consultation period was public 6 weeks earlier; tighter regulatory scanning would have surfaced the shift before it forced a re-profile.

## Cross-project recommendations

| # | Recommendation | Audience | Owner |
|---|----------------|----------|-------|
| 1 | Add accessibility-evidence requirement to org-wide DoD template | PMO + Design | Sam Reyes |
| 2 | Standardize "co-own OCM with affected-team leader" pattern in PMO playbook | PMO | Marcus |
| 3 | Surface workstream-specific buffers in baselines, not only in management reserve | PMO | Marcus |
| 4 | Maintain a regulatory-scan calendar for compliance-touching projects | Privacy / Legal | Aisha |

## How this document will travel

Posted to PMO Notion under "Lessons / FY26". Featured in next PM guild meeting (2026-06-04). Cross-referenced from the next two onboarding-adjacent projects in planning.

---

**Principles Honored:** 8 (quality), 9 (complexity), 11 (adaptability & resiliency), 12 (change).
