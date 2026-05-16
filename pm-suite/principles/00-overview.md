# The 12 Project Management Principles

The principles are the **why** of the suite. Every agent action — drafting an artifact, recommending a tailoring choice, auditing a deliverable — must be traceable back to one or more of these principles. The `pm-principles-auditor` agent (Phase 2) uses the audit checks at the bottom of each principle file to flag drift.

## The 12 principles

| # | Principle | Short form |
|---|-----------|------------|
| 1 | Stewardship | Be a diligent, respectful, and caring steward |
| 2 | Team | Create a collaborative project team environment |
| 3 | Stakeholders | Effectively engage with stakeholders |
| 4 | Value | Focus on value |
| 5 | Systems thinking | Recognize, evaluate, and respond to system interactions |
| 6 | Leadership | Demonstrate leadership behaviors |
| 7 | Tailoring | Tailor based on context |
| 8 | Quality | Build quality into processes and deliverables |
| 9 | Complexity | Navigate complexity |
| 10 | Risk | Optimize risk responses |
| 11 | Adaptability & resiliency | Embrace adaptability and resiliency |
| 12 | Change | Enable change to achieve the envisioned future state |

## How agents use these

1. **At action time** — Each agent declares which principles its current action serves (e.g., "drafting stakeholder register: principles 3, 5, 9").
2. **At artifact time** — Each generated artifact carries a `Principles Honored` footer listing the principles it supports.
3. **At audit time** — `pm-principles-auditor` reads artifacts and flags violations (e.g., a risk register with no mitigations cited fails Principle 10; a charter with no stakeholder analysis fails Principle 3).

## Relationship to performance domains

Principles are universal; performance domains are where work happens. A single principle (e.g., Value) shows up across multiple domains (Planning, Delivery, Measurement). The mapping is intentionally many-to-many.

## When principles conflict

They will. Speed (Adaptability) vs. predictability (Quality, baselines). Stakeholder X's value vs. stakeholder Y's. The principles do not rank themselves — **tailoring** (Principle 7) is the explicit mechanism for resolving conflicts in context. Document the call and move on.
