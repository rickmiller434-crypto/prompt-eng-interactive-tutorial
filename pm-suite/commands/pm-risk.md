---
description: Open the risk register for a review pass and update aging or status as needed
argument-hint: <project>
---

Run a risk review on `pm-suite/projects/$ARGUMENTS/risk-register.md`.

Steps:

1. Invoke the `pm-uncertainty` subagent via the `Agent` tool.
2. Goal: review every open risk, update scores if probability or impact has shifted, flag any aged-past-review-date items, and append at least one opportunity check (Principle 10 requires both threats and opportunities).
3. After review, summarize:
   - Top 5 risks with deltas
   - Risks closed this review (with closure rationale)
   - Risks aged past review date (require pm-lead attention)
   - Any new systemic / external risks (Principle 5 expectation)

If a risk has realized into an issue, instruct the user to open an issue entry (in RAID log) and consider whether a CR is needed.

Cite Principle 10 (Risk), 5 (Systems thinking), 11 (Adaptability & Resiliency).
