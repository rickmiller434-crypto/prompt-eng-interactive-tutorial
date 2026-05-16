# Project Charter — Acme Onboarding Revamp

- **Sponsor:** Priya Shah, VP Customer Success
- **PM:** Marcus Lin, Senior PM
- **Date authorized:** 2026-01-29
- **Version:** v1

## Vision

A new Acme customer signs up on Monday and is actively using the product daily by the following Monday — without needing a human-led kickoff call. The Customer Success team spends its time on expansion conversations, not on chasing stuck onboardings. The new flow honors GDPR and the 1 Sep 2026 data-residency rule from the day it ships.

## Business need

30-day drop-off has crept from 19% (2024) to 28% (Q4 2025). Customer Success time per onboarding has risen from 4.8 hr to 6.4 hr. Net retention is showing pressure in the mid-market segment. Without intervention, FY26 retention is forecast 4–6 points below plan.

## Objectives & success criteria

| # | Objective | Measure | Baseline | Target | Window |
|---|-----------|---------|----------|--------|--------|
| 1 | Faster time-to-value | Median TTFV | 14 days | ≤ 5 days | 60d post-launch |
| 2 | Lower drop-off | 30-day drop-off rate | 28% | < 15% | 90d post-launch |
| 3 | Reduced CS load | CS hr per onboarding | 6.4 hr | ≤ 2.0 hr | 60d post-launch |

## High-level scope

In scope:
- New customer onboarding UX (sign-up → first key action) for mid-market segment
- In-product guidance content (tours, video walkthroughs)
- Instrumentation for the three measures above
- Compliance workstream: data-residency, GDPR, SOC 2 review for changed flows

Out of scope:
- Enterprise tier (separate onboarding model)
- Self-serve free tier (not yet introduced)
- Mobile app onboarding (separate roadmap)

## Key stakeholders

(Full register: `stakeholder-register.md`.) Sponsor: Priya Shah. Primary user: mid-market admins. Key affected teams: CS, Sales, Eng, Data. Regulators: GDPR DPA; SOC 2 auditor (annual).

## High-level milestones

| Milestone | Target | Notes |
|---|---|---|
| Kickoff & tailoring sign-off | 2026-01-29 | Done |
| Discovery + first backlog | 2026-02-12 | |
| Internal beta | 2026-03-19 | Tailoring re-validation gate |
| External limited beta | 2026-04-16 | Tailoring re-validation gate |
| GA | 2026-05-14 | |
| 60-day post-launch review | 2026-07-13 | Benefits checkpoint |

## Budget envelope

$480k authorized (incl. ~12% contingency). Detail in `cost-baseline.md` (compliance workstream) and burn-down in `status-reports/`.

## Constraints

- Time: GA before summer freeze (end of May 2026)
- Compliance: 1 Sep 2026 data-residency rule
- Resource: shared engineering capacity with platform team
- Tools: Jira, Notion, GitHub

## Assumptions

(Detail in `assumption-constraint-log.md`.) Headline: existing instrumentation can be extended without a re-platforming effort; mid-market segment adopts a flow tested on a 200-customer beta.

## Risks at charter time

(Live in `risk-register.md`.) Top 3: CS-team resistance; shared-infra dependency; 1 Sep regulatory deadline.

## Obligations & constraints (Principle 1)

- GDPR (existing)
- SOC 2 (existing)
- Data-residency rule effective 1 Sep 2026 (new; non-negotiable)

## Authority

- PM authorized to: assign work within the team, spend within approved budget lines, approve change requests with cumulative cost ≤ $20k.
- PM must escalate to sponsor: scope changes that affect milestones, change requests > $20k, baseline re-baselines.

## Sign-off

| Role | Name | Date |
|---|---|---|
| Sponsor | Priya Shah | 2026-01-29 |
| PM | Marcus Lin | 2026-01-29 |
| Head of Eng | Dani Okonkwo | 2026-01-29 |

---

**Principles Honored:** 1 (stewardship), 3 (stakeholders), 4 (value), 6 (leadership), 7 (tailoring), 12 (change).
