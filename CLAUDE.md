# Workspace Context — Rick Miller / RMM

This is Rick Miller's workspace for Ausenco engagements and parallel programs.
Author: Rick Miller, Senior Project Director.

## Active programs (May 2026)

| Program | Owner | Entity | Status |
|---|---|---|---|
| Goldboro FS 2026 (Job 107206-05) | NexGold Mining Corp. | Ausenco Engineering Canada ULC | Active — bid review |
| Green Bay (EOI) | FireFly Metals | Ausenco Engineering Canada ULC | On hold |
| Goderich follow-on | Compass Minerals | Ausenco Engineering Canada ULC | Follow-on scoping |
| Kemess | Centerra / Kemess Mining | Carter's / TMG combined | NOT active — flagged for future |

## Critical activation rule — Ausenco FD Presentation Standard

**Any output produced under Ausenco attribution must follow the
Ausenco FD Presentation Standard** documented at
`.claude/skills/ausenco-fd-presentation-standard.md`.

**Activates when:**
- Output is going under Ausenco attribution (Goldboro, Green Bay,
  Goderich, Ausenco BD)
- Request mentions: "Ausenco deliverable", "Ausenco memo", "Ausenco
  workbook", "Functional Director", "FD-grade", "letterhead-grade",
  "James", "Gallant", "Wilford", "Medley", "Sterling"
- Document type: benchmark, cost estimate, bid evaluation, executive
  briefing, gate readiness, PEP, schedule review, RFP audit, fee
  proposal, owner's rep memo
- Distribution: Ausenco internal leadership (Dan Wilford, Rob Medley,
  James Gallant, Grahame Sterling) OR NexGold / FireFly Metals /
  Compass under Ausenco letterhead

**Does NOT activate for:**
- Carter's Admin internal documents (separate standard)
- Centerra / LMC / Kemess work (Carter's / TMG combined standard
  with LMC privilege protocols — NOT YET DOCUMENTED HERE)
- Conversational responses or quick-turn analysis

**When activation is ambiguous: default to activating.** The seven
elements never hurt a deliverable; their absence does.

## When the standard is active — what Claude must do

1. **Read** `.claude/skills/ausenco-fd-presentation-standard.md`
   before producing output. The full SKILL spec with James Gallant's
   worked examples is there.

2. **Use the templates** in `mining_estimating_manual/templates/`
   as starting points:
   - `template_ausenco_memo.md` — memo header/footer with the 7
     elements wired in
   - `template_ausenco_workbook_cover.xlsx` — workbook Cover tab
     with classification stamp, confidence stamp, filing path
   - `template_audit_checklist.md` — 10-point pre-flight check

3. **Run the audit** before declaring any Ausenco deliverable
   ready to ship: `python3 mining_estimating_manual/build/fd_audit.py
   <workbook_path>`. Returns PASS / FAIL on each of the 10 elements.
   Any FAIL = rework before issuing.

4. **Mirror James Gallant's structure** for the deliverable type:
   - Benchmark workbook: 5 tabs (Cover & Summary, Rate Library,
     MTO × Rate, Area Roll-Up, Sensitivity & Flags)
   - Consolidation workbook: 8 tabs (Executive Dashboard, Financial,
     Q1-Q4 questions, Unresolved, References)
   - Memorandum: 10-section format (Purpose, Project Context,
     Result, Quantities by WBS, Methodology, Sensitivity, SME Flags,
     Limitations + Confidence, Supporting Deliverables, Distribution
     & Filing)

## Important — what is and is not durable

- This `CLAUDE.md` file auto-loads when Claude Code opens this
  workspace. That is real.
- Claude does NOT carry memory across sessions independently of
  files like this one. Each session starts fresh.
- The durable mechanism is workspace files (this `CLAUDE.md`, the
  SKILL file, the templates, the audit checker). When the workspace
  is open, those files give Claude full context. When the workspace
  is not open, Claude has no memory of any of this.
- For cross-workspace use (Carter's / Kemess project), the SKILL
  file must be explicitly loaded into the session.

## Naming conventions (Ausenco scheme)

- Job numbers: `107206-05` (Goldboro), one number per project
- Document numbers: `[Job]-[DocType]-[Package]-[Function]-[Seq]`
  - Example: `107206-RC-C5027-22320-001` (SoW for C5027)
  - Example: `107206-EC-C5027-40448-001` (Pricing schedule for C5027)
  - Example: `107206-MA-00000-99999-001` (Memorandum, generic)
- Document type codes: RC (SoW), EC (Estimate / Cost), MA (Memo),
  DT (Drawing), TR (Technical Report), SP (Specification)

## Author signature block

```
Rick Miller, Senior Project Director
Ausenco Engineering Canada ULC
1016B Sutton Drive, Suite 100
Burlington, ON L7L 6B8 | Canada
Mobile: 780-838-1653 | ausenco.com
```

NO P.Eng / NO Engineer credential — Rick is Senior Project Director, not an engineer.
