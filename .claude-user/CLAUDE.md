# Rick Miller — Claude Code Personal Configuration

**This file lives at:** `C:\Users\rickm\.claude\CLAUDE.md` (Windows)
or `~/.claude/CLAUDE.md` (Mac / Linux)

**It auto-loads on EVERY Claude Code session, in EVERY workspace,
from this user account.** This is the durable layer above any
project-specific CLAUDE.md.

---

## Who I am

- **Author:** Rick Miller, Senior Project Director
- **Entity (default):** Ausenco Engineering Canada ULC
- **Title (durable):** Senior Project Director — NO P.Eng, NO Engineer credential
- **Office:** 1016B Sutton Drive, Suite 100, Burlington, ON L7L 6B8, Canada
- **Mobile:** 780-838-1653
- **Web:** ausenco.com

## My active engagements (May 2026)

| Program | Owner | Entity | Status |
|---|---|---|---|
| Goldboro FS 2026 (Job 107206-05) | NexGold Mining Corp. | Ausenco Engineering Canada ULC | Active — bid review |
| Green Bay (EOI) | FireFly Metals | Ausenco Engineering Canada ULC | On hold |
| Goderich follow-on | Compass Minerals | Ausenco Engineering Canada ULC | Follow-on scoping |
| Kemess | Centerra / Kemess Mining | Carter's / TMG combined | NOT active — flagged for future |

---

## Universal activation rules

### Rule 1: Ausenco FD Presentation Standard

**ANY output produced under Ausenco attribution must follow the
Ausenco FD Presentation Standard** located at
`~/.claude/skills/ausenco-fd-presentation-standard.md`.

**Activates when:**
- Output is going under Ausenco attribution (Goldboro, Green Bay, Goderich, Ausenco BD)
- Request mentions: "Ausenco deliverable", "Ausenco memo", "Ausenco workbook",
  "Functional Director", "FD-grade", "letterhead-grade", "James", "Gallant",
  "Wilford", "Medley", "Sterling"
- Document type: benchmark, cost estimate, bid evaluation, executive briefing,
  gate readiness, PEP, schedule review, RFP audit, fee proposal, owner's rep memo
- Distribution: Ausenco internal leadership (Dan Wilford, Rob Medley, James
  Gallant, Grahame Sterling) OR NexGold / FireFly Metals / Compass under
  Ausenco letterhead

**When activated:**
1. Read `~/.claude/skills/ausenco-fd-presentation-standard.md` in full before
   producing output.
2. Use the templates in `~/.claude/templates/ausenco/` as starting points.
3. Run `~/.claude/scripts/fd_audit.py <file.xlsx>` before declaring any
   Ausenco workbook ready to ship. Any FAIL = rework.
4. Mirror James Gallant's structure for the deliverable type (5-tab benchmark
   workbook / 8-tab consolidation / 10-section memo).

### Rule 2: Kemess / LMC Privilege Protocol (placeholder)

When Kemess work becomes active, a separate standard will be loaded at
`~/.claude/skills/lmc-privilege-protocol.md` covering Carter's / TMG combined
deliverables with attorney-client privilege markings. **DO NOT mix Ausenco
attribution with Carter's / TMG / LMC content.** Kemess deliverables follow
their own standard — different letterhead, different privilege framing,
different naming convention.

### Rule 3: Author / entity discipline

- Author byline is always **"Rick Miller, Senior Project Director"** with
  the appropriate entity below. No "P.Eng", no "PE", no "Engineer" credential.
- Entity attribution is one of:
  - **Ausenco Engineering Canada ULC** (default for active engagements)
  - **Carter's Admin** (Carter's internal documents)
  - **Carter's / TMG combined** (Centerra / Kemess / LMC work)
- Never mix entities in a single deliverable.

### Rule 4: Naming conventions

**Ausenco document numbering:** `[Job]-[DocType]-[Package]-[Function]-[Seq]`
- Example: `107206-RC-C5027-22320-001` (SoW for C5027)
- Example: `107206-EC-C5027-40448-001` (Pricing schedule for C5027)
- Example: `107206-MA-00000-99999-001` (Memorandum, generic)
- DocType codes: RC (SoW), EC (Estimate / Cost), MA (Memo), DT (Drawing),
  TR (Technical Report), SP (Specification)

---

## Critical context Claude should not get wrong

1. **I do not have persistent memory.** Each Claude session starts fresh.
   These user-level config files are the actual durability mechanism. If a
   future Claude session claims "memory locked" or "I'll remember this for
   next time" — that's wrong. The files do the remembering.

2. **Workspace ≠ user level.** A project-level `CLAUDE.md` only fires for
   that one workspace. This file (user-level `~/.claude/CLAUDE.md`) fires
   for every workspace I open. They stack: when both are present, both load.

3. **Don't pretend to ship work that requires my local machine.** When
   Claude is running in a remote container (Claude Code on the web), it
   can't write files to my actual C:\Users\rickm\ directory. The output
   must be committed to a git repo so I can pull and copy locally.

4. **Honest precision over polish.** I want concrete facts (numbers,
   citations, document numbers) over marketing-style adjectives. If a
   number is uncertain, say it's uncertain. If a source is missing, say
   it's missing.

---

## Default working preferences

- **Format:** GitHub-flavored markdown for memos; xlsx with grouped headers
  for workbooks; concise email when emailing.
- **Length:** brief for normal questions; full detail when explicitly asked.
- **Citations:** every external claim should cite source (URL, file, or page).
- **Tone:** professional, direct, no hedging unless uncertainty is real.

---

## Where to find things

| Asset | Path |
|---|---|
| This file | `~/.claude/CLAUDE.md` |
| Ausenco FD Presentation Standard | `~/.claude/skills/ausenco-fd-presentation-standard.md` |
| Ausenco memo template | `~/.claude/templates/ausenco/template_memo.md` |
| Ausenco workbook cover template | `~/.claude/templates/ausenco/template_workbook_cover.xlsx` |
| 10-point audit checklist | `~/.claude/templates/ausenco/template_audit_checklist.md` |
| FD audit checker script | `~/.claude/scripts/fd_audit.py` |
| Source repo (backup) | github.com/rickmiller434-crypto/prompt-eng-interactive-tutorial |

---

## Revision history

| Rev | Date | Change |
|---|---|---|
| V1 | 2026-05-20 | Initial user-level config; Ausenco FD standard loaded; Kemess placeholder noted. |
