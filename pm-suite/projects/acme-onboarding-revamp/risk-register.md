# Risk Register — Acme Onboarding Revamp

- **Owner:** pm-uncertainty
- **Last reviewed:** 2026-03-12

**Risk appetite statement:** The sponsor accepts up to 2 weeks of schedule slip if it preserves quality. Cost overrun tolerance ≤ 10% of envelope. Value risk (missing benefit targets) is **low** appetite — the case rests on these outcomes.

## Register

| ID | Type | Category | Description | Cause | Effect | P | I | Score | Response | Plan | Trigger | Owner | Status | Review |
|----|------|----------|-------------|-------|--------|---|---|-------|----------|------|---------|-------|--------|--------|
| R1 | Threat | Organizational | CS team resists role change; under-utilizes new flow | Anxiety over job changes | Adoption < target; benefit 2 not realized | 4 | 5 | 20 | Mitigate | OCM plan: co-design + sandbox week + measured rollout; Jordan as co-owner | CSM sentiment survey drops > 0.5 from baseline | pm-stakeholders | Open | weekly |
| R2 | Threat | Technical | Shared-infra dependency delays (platform team) | Competing roadmap priorities | Internal beta slips | 3 | 4 | 12 | Mitigate | Monthly review with platform; identified workarounds for 2 critical paths | Platform velocity > 1 sprint behind | pm-approach | Open | weekly |
| R3 | Threat | Regulatory | 1 Sep 2026 data-residency rule clarifications change late | Regulator behavior | Re-work in compliance workstream | 2 | 5 | 10 | Mitigate | Aisha embedded; weekly regulatory scan; design for residency optionality | New regulator guidance issued | pm-uncertainty | Open | weekly |
| R4 | Threat | External / market | Acquisition pause in mid-market segment shrinks cohort sizes | Macro / sales pipeline | Benefit measurements noisy | 2 | 3 | 6 | Accept (with monitor) | Pre-define minimum cohort sizes for valid measurement | Pipeline forecast drops > 25% | pm-measurement | Open | monthly |
| R5 | Threat | Quality | In-product video walkthroughs fail accessibility / localization | Newer team; novelty | DoD failures; legal exposure | 3 | 3 | 9 | Mitigate | Accessibility review pre-record; native-speaker QA per locale | First accessibility audit finds > 2 issues | pm-delivery | Open | per increment |
| R6 | Opportunity | Commercial | New flow becomes a sales asset (demo, case studies) | Visibility of UX improvement | Pipeline uplift in mid-market | 3 | 3 | 9 | Enhance | Coordinate with Marketing; case-study customer secured at limited beta | Customer advisory feedback "best experience" by ≥ 4/8 | pm-stakeholders | Open | per increment |
| R7 | Threat | Team | Single point of failure: Lin Wei (data analyst) — only instrumentation expert | Hiring not yet caught up | Instrumentation blocked if Lin unavailable | 3 | 4 | 12 | Mitigate | Pair Lin with eng for 4 weeks; runbook for instrumentation patterns; cross-train one engineer | Lin out > 1 week | pm-team | Open | monthly |
| R8 | Threat | Systemic | Sales→CS handoff process changes break for non-mid-market customers we didn't model | Out-of-scope segment piggybacks the flow | Customer complaints; segmentation logic patch | 2 | 3 | 6 | Mitigate | Segmentation guardrails in code; feature-flag rollout | Out-of-scope traffic > 5% of beta | pm-work | Open | per increment |

## Top risks (this reporting cycle)

| ID | Description | Score | Δ |
|---|---|---|---|
| R1 | CS resistance | 20 | flat |
| R2 | Shared-infra delays | 12 | flat |
| R7 | Lin SPoF | 12 | new |
| R3 | Late regulatory clarifications | 10 | flat |
| R5 | Video walkthrough quality | 9 | flat |

## Closed / realized risks (log)

(None yet at 2026-03-12.)

## Systemic risks check

R8 (out-of-scope segment behavior) and R4 (market segment shifts) are systemic / external. ✓

---

**Principles Honored:** 10 (risk), 5 (systems thinking), 11 (adaptability & resiliency).
