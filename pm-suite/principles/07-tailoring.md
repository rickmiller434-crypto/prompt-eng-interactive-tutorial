# Principle 7 — Tailoring

> Tailor based on context.

## Intent

No two projects are identical. The lifecycle, the cadence, the artifacts, the controls — all are tailored to the project's context (size, complexity, novelty, stakes, regulatory environment, team experience, organizational maturity). Tailoring is the meta-principle that resolves tension between the other 11.

## What it means in practice

- Tailoring is **explicit** — documented in the project's tailoring record, not implicit in "how we do things."
- Tailoring is **revisited** — the right choices at kickoff may be wrong six months in.
- The tailoring record names which artifacts will be produced, which agents will engage, and which controls will run.

## Agent behaviors

- `pm-tailor` (Phase 2) is invoked at initiation and at every major checkpoint. Its output — the **Tailoring Record** — is the source of truth for which other agents and templates the project uses.
- Every other agent reads the tailoring record before acting and refuses to produce artifacts the record has excluded.
- Changes to the tailoring record require a decision log entry.

## Audit checks

- [ ] Tailoring record exists and is dated.
- [ ] Tailoring record specifies lifecycle (predictive / adaptive / hybrid) with rationale.
- [ ] Tailoring record lists included/excluded artifacts and engaged agents.
- [ ] Tailoring record is re-validated at each phase gate or major increment.
