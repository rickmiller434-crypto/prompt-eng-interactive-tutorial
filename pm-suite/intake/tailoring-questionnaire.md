# Tailoring Intake Questionnaire

> Fill this out at project initiation. Re-validate at each phase gate or major increment. The output — the **Tailoring Record** — is the source of truth for which agents engage and which artifacts the project produces.

---

## Section A — Project identity

1. **Project name:**
2. **Sponsor / business owner:**
3. **PM / delivery lead:**
4. **One-line description (the elevator pitch):**
5. **Desired future state (one paragraph):**
6. **Hard deadline or target window (if any):**

---

## Section B — Value & success

7. **Primary value the project will deliver** (revenue, cost reduction, risk reduction, capability, compliance, customer outcome, other):
8. **Top 3 measurable benefits** (name, baseline, target, when realized):
9. **What does success look like at 30 / 90 / 365 days post-go-live?**
10. **What would make this project a clear failure?**

---

## Section C — Stakeholders

11. **Sponsor's level of engagement** (sets vision / approves only / hands-off):
12. **Primary user / customer group(s):**
13. **Regulators or external bodies involved (if any):**
14. **Other affected teams or projects:**
15. **Anticipated controversy / opposition** (none / low / moderate / high):

---

## Section D — Scope & uncertainty

16. **Requirements stability** (stable / mostly stable / volatile / unknown — discovery required):
17. **Technology novelty** (proven for us / proven elsewhere / experimental):
18. **Solution clarity at start** (we know the solution / we know the problem / we are still framing the problem):
19. **Number of distinct deliverables / increments anticipated:**
20. **Anticipated change rate during delivery** (low / medium / high):

---

## Section E — Size & resources

21. **Team size** (core + extended):
22. **Duration** (target months):
23. **Budget tier** (S / M / L / XL — define ranges per org):
24. **Geographic distribution** (co-located / hybrid / fully remote / multi-timezone):
25. **Team experience with this domain** (high / mixed / low):

---

## Section F — Constraints & obligations

26. **Regulatory / compliance obligations** (list):
27. **Contractual obligations driving deliverables or dates:**
28. **Security / data classification considerations:**
29. **Sustainability / ESG considerations:**
30. **Health & safety considerations:**

---

## Section G — Delivery cadence

31. **Delivery mode** (single big-bang / phased releases / continuous delivery):
32. **Acceptable batch size** (one big release / quarterly / monthly / per-increment):
33. **Customer ability to absorb change** (high / medium / low):

---

## Section H — Organizational context

34. **Organization's PM maturity** (ad-hoc / repeatable / defined / managed / optimizing):
35. **Preferred or mandated methodology** (none / waterfall / agile flavor / SAFe / custom — name it):
36. **Tools that are mandatory** (Jira / Notion / GitHub / MS Project / other):
37. **Reporting cadence expected by leadership** (weekly / biweekly / monthly / on milestone):

---

## Tailoring recommendation (produced by `pm-tailor`)

Once Section A–H are filled out, the `pm-tailor` agent produces this section.

### Recommended lifecycle

- [ ] **Predictive** — stable requirements, low novelty, well-understood solution, regulatory baselines, big-bang or phased delivery.
- [ ] **Adaptive** — volatile requirements, exploratory or high-novelty work, customer can absorb frequent change, short increments preferred.
- [ ] **Hybrid** — predictive backbone with adaptive increments inside specific workstreams; or adaptive delivery with predictive governance / reporting overlay.

**Rationale (3–5 sentences referencing the answers above):**

### Engaged agents

| Agent | Engaged? | Notes |
|---|---|---|
| `pm-lead` | ✅ always | Orchestrator |
| `pm-tailor` | ✅ always | Owns this record |
| `pm-principles-auditor` | ✅ always | Runs at gates |
| `pm-stakeholders` |  |  |
| `pm-team` |  |  |
| `pm-approach` |  |  |
| `pm-planning` |  |  |
| `pm-work` |  |  |
| `pm-delivery` |  |  |
| `pm-measurement` |  |  |
| `pm-uncertainty` |  |  |

### Required artifacts

Initiating:
- [ ] Project charter
- [ ] Business case / value statement
- [ ] Stakeholder register
- [ ] Assumption / constraint log

Planning (check the ones tailoring recommends):
- [ ] WBS / scope baseline
- [ ] Product backlog (adaptive)
- [ ] Schedule / Gantt
- [ ] Release plan (adaptive)
- [ ] Cost baseline
- [ ] Risk register
- [ ] Quality plan / Definition of Done
- [ ] Communications plan
- [ ] RAID log
- [ ] RACI matrix
- [ ] Procurement plan
- [ ] Resource plan
- [ ] Team charter / working agreement
- [ ] OCM plan
- [ ] Dependency map

Executing / Monitoring:
- [ ] Change request log
- [ ] Status report cadence (frequency: ____)
- [ ] EVM (predictive) / velocity & CFD (adaptive) / both (hybrid)
- [ ] Risk burndown
- [ ] Decision log
- [ ] Retro cadence (frequency: ____)

Closing:
- [ ] Acceptance / sign-off pack
- [ ] Lessons learned
- [ ] Benefits realization plan (post-close)
- [ ] Handover to operations
- [ ] Closure report

### Cadence

- Status reporting: __________
- Retros / phase reviews: __________
- Tailoring re-validation: __________
- Risk review: __________

### Principles in tension (and how this project resolves them)

Document any principles that pull against each other for this project and the chosen resolution (e.g., "Quality (8) vs. Adaptability (11): we will hold a fixed Definition of Done but allow scope to flex per increment").

---

## Sign-off

- Tailoring author:
- Date:
- Next re-validation due:
- Approved by sponsor:
