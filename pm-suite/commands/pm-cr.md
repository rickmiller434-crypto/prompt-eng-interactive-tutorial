---
description: Open a new change request
argument-hint: <project> "<change request title>"
---

Open a new change request for `pm-suite/projects/$1/`.

Steps:

1. Determine the next CR ID (`CR-NNN`) by scanning existing files in `pm-suite/projects/$1/CRs/` (create the directory if missing).
2. Copy `pm-suite/templates/executing/change-request.md` to `pm-suite/projects/$1/CRs/CR-<NNN>.md`.
3. Pre-fill: CR ID, Title (from `$ARGUMENTS` after the project name), Submitted by (the user, or ask), Submitted date (today), Status: Submitted.
4. Ask the user for the change requested, the reason, and initial impact assessment across **all** dimensions in the template — pay attention to **value impact** (Principle 4 — not just cost / time).
5. Once filled, save the file and tell the user the next step:
   - If decision-level is PM: log decision via `/pm-decide $1 "Approve / Reject CR-<NNN>: …"`
   - If decision-level is sponsor or steering: surface to `pm-lead` for routing.

Do not auto-approve. CRs require an explicit decision.

Cite Principle 4 (Value), 7 (Tailoring), 10 (Risk), 12 (Change).
