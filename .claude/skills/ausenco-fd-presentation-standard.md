# Ausenco Functional Director Presentation Standard

**Path:** `C:\Users\rickm\keystone\skills\ausenco-fd-presentation-standard.md`
**Activated by:** any Ausenco deliverable (Goldboro, Green Bay, Goderich, Ausenco BD)
**Trigger words:** "Ausenco deliverable", "Ausenco memo", "Ausenco workbook", "Functional Director", "FD-grade", "letterhead-grade"
**Last updated:** 2026-05-20 (V1.0 - absorbed from J. Gallant Goldboro C5027 parallel work)
**Author:** Rick Miller, Senior Project Director

---

## 1. Purpose

Every Ausenco-attributed deliverable Rick produces lands at Functional Director presentation level. This SKILL encodes the seven required elements observed in J. Gallant's (Functional Director Construction & Commissioning, North America) parallel C5027 benchmark + 5-bidder consolidation work issued 15-19 May 2026.

The analytical substance was already at peer level. The presentation gap was the level-up move. This SKILL closes it.

---

## 2. Standard Of Care

Ausenco deliverables are stamped under Ausenco Engineering Canada ULC. The reader assumes Director-level discipline:
- AACE 18R-97 estimate classification visible
- Methodology decomposed, not asserted
- Reusable library separated from project-specific application
- Flags are actionable, not informational
- Sensitivity quantified, not hand-waved
- Affirmative use case stated, not just limitations
- Filing path declared for project file integration

If any of the seven elements below is missing, the deliverable is NOT Ausenco letterhead grade and should not go out under Ausenco attribution.

---

## 3. The Seven Required Elements

### Element 1 - AACE Classification Stamp At The Top

**What it is:** One line at the top of every memo, workbook cover, executive summary that states what document class this is and what it is not.

**Format:**
> *"This is a [class] [type] prepared for [purpose]. It is NOT [misuse case 1, 2, 3]."*

**Example from James's C5027 benchmark memo:**
> *"This benchmark is NOT a tender evaluation, contractor selection recommendation, or commitment of estimated value. It is intended for internal Ausenco use to inform contingency setting, owner reporting bands, and SME-led review of the tender package prior to bid receipt."*

**AACE Class reference (from 18R-97):**
- Class 5: Concept screening (-50% / +100%)
- Class 4: Study or feasibility (-30% / +50%)
- Class 3: Budget authorization / control (-20% / +30%)
- Class 2: Control or bid / tender (-15% / +20%)
- Class 1: Check estimate or bid / tender (-10% / +15%)

**Why it matters:** Reader knows in 5 seconds what document class this is, what to use it for, and what to NOT use it for. Pre-empts misuse risk.

---

### Element 2 - Adjustments Decomposed Into Named Components

**What it is:** When a calibration, adjustment, or regional/site uplift is applied, the components are named individually with bidirectional ranges that net out.

**Wrong way:**
> *"Applied 5% adjustment for Atlantic Canada conditions."*

**Right way (from James's Section 5.3):**
> *"A composite -5% adjustment was applied to the Northern Ontario / Manitoba baseline to land at the Nova Scotia rate set. The components: -3% to -5% logistics and port access (Halifax within 175 km, reduced freight on bulk imports); +2% to +5% labour availability (smaller NS specialist mining earthworks pool); +3% to +5% climate seasonality (maritime freeze-thaw, higher rain shutdown days); +2% to +4% geomembrane window (Mid-May to Mid-Oct typical); +1% to +3% NSE Tier 1 EQS QA/QC burden. Logistics advantage dominates positive; labour and climate offset partially."*

**Standard form:**

| Component | Direction | Range | Driver |
|---|---|---|---|
| Logistics & port access | -3% to -5% | Halifax within 175 km, reduced freight on bulk imports |
| Labour availability | +2% to +5% | Smaller NS specialist mining earthworks pool |
| Climate seasonality | +3% to +5% | Maritime freeze-thaw, higher rain shutdown days |
| Geomembrane window | +2% to +4% | Mid-May to Mid-Oct typical |
| QA/QC regulatory burden | +1% to +3% | NSE Tier 1 EQS |
| **NET COMPOSITE** | **-5%** | **Logistics dominates positive; labour and climate offset partially** |

**Why it matters:** Anyone reviewing can challenge any single component without rejecting the whole adjustment. Senior peer review is structured component-by-component, not all-or-nothing.

---

### Element 3 - Rate Library Separated From Application

**What it is:** In any workbook deliverable that uses a rate library + applies it to a project, the library is a separate sheet from the application.

**Wrong way:** Single tab where rates, regional bands, project bidder data, project context notes all live together.

**Right way (from James's benchmark workbook):**
- Sheet 02: Rate Basis Library (223 rows, ~85 patterns × UoM combinations) - the **clean reusable asset**
- Sheet 03: MTO × Rate (898 rows, all 815 C5027 line items applying library rates) - the **project-specific application**

**Standard structure for a benchmark / rate workbook:**

| Sheet | Purpose | Reusable? |
|---|---|---|
| 00 Cover | Title, classification, revision, distribution | Per project |
| 01 Methodology | Rate basis, adjustments, decomposition | Reusable across projects |
| 02 Rate Library | Clean unit rates by category × UoM × region | **REUSABLE ASSET** |
| 03 Project Application | Library applied to project MTO with per-line context | Per project |
| 04 Area / WBS Roll-Up | Project-specific aggregation | Per project |
| 05 Sensitivity & Flags | Project-specific risk/sensitivity | Per project |
| 06 Sources & Provenance | Library citation chain | Reusable |

**Why it matters:** The library is the asset that compounds in value across projects. Mixing it with project-specific application destroys reusability. Carter's portfolio across Green Bay NL, Goderich ON, Kemess BC, future Ausenco BD pursuits all benefit from a clean library.

---

### Element 4 - FLAG Taxonomy Includes Resolution Required Column

**What it is:** Every flag in a flag register has a name, a severity, AND an explicit "what needs to happen to clear this" statement.

**Wrong way:**

| Flag | Severity |
|---|---|
| Pit pre-strip scope | HIGH |
| MSE wall design | MED |

Reader knows the issue exists. Does not know what to do about it.

**Right way (from James's Section 7):**

| Flag | Severity | Resolution required before SME sign-off |
|---|---|---|
| Pit pre-strip scope split | HIGH | 6.68 M m³ pre-prod is largest single bucket at $66-99 M direct. Confirm whether contractor or owner mining scope. If owner-mined, deduct from this package. |
| MSE wall at primary crusher | MED | SoW §3.2.3.1 specifies MSE wall is design-build with CAT 777 loading + 15 kPa surcharge. Confirm scope sits in this package or carried by structural contractor. |

**Standard form columns:**
- Flag (named)
- Severity (HIGH / MED / LOW / CLOSED)
- Resolution required (specific action, named owner if known, by-when if known)
- Status (Open / In Progress / Closed)

**Why it matters:** Flag is actionable, not informational. Reader knows the issue, the severity, AND what to do. Cuts the "what now?" follow-up email.

---

### Element 5 - Top 3 Sensitivities Pre-Calculated With Dollar Impact

**What it is:** Before the limitations section, the deliverable pre-calculates the 3 largest uncertainties with their dollar impact at the deliverable level.

**Example from James's Section 6:**
> *"Rate sensitivity: a ±5% uniform rate shift moves the rated direct cost by ±$13.3 M and the total carried cost by approximately ±$18.8 M. A 10% market escalation scenario adds approximately $37.6 M."*
>
> *"Pit pre-strip scope: if 2100 Pit Pre-prod (6.68 M m³, $82.4 M mid) is confirmed as owner-mined scope, the total carried reduces by approximately $116 M."*
>
> *"Geomembrane specification: if NSE compliance drives an LLDPE upgrade or double-liner system, the impact is approximately +$5.2 M direct (+$7.3 M total carried)."*

**Standard form:**

| Sensitivity | Trigger | Dollar Impact | Action if triggered |
|---|---|---|---|
| Rate market shift | ±5% uniform rate change | ±$X.X M total carried | Re-issue with updated escalation |
| Scope split decision | Owner-mined vs Contractor-mined Area X | $Y M reduction if Owner-mined | Tender scope clarification |
| Spec upgrade | NSE compliance LLDPE upgrade | +$Z M direct | Update geomembrane spec basis |

**For benchmark-style deliverables (like Rick's heavy civil workbook):** the sensitivity is at the rate-calibration level rather than the project total level. Example: *"If Goldboro rock geology turns out to be cleaner than weathered meta-sedimentary, C-10-007 should drop from $68 to $46 - this represents a $X impact on C5027 rock excavation volume of Y m³."*

**Why it matters:** Reader knows in advance what will move the number and by how much. The deliverable is robust to predictable challenges because they are pre-answered.

---

### Element 6 - Affirmative Confidence Stamp Before Limitations

**What it is:** Before the disclaimer / limitations text, an affirmative statement of what the deliverable IS good for. Limitations follow, not lead.

**Wrong way:**
> *"This document has limitations. It is not a substitute for X. It does not account for Y. The reader should be aware that Z..."*

Reader leaves with what the document cannot do. The use case is buried.

**Right way (from James's Section 8):**
> *"Confidence stamp (P50 mid-band): suitable for internal range planning, owner reporting bands, and contingency-setting discussions. The P10 and P90 boundaries represent reasonable scenarios but should not be treated as hard limits."*
>
> *"What this is NOT: a tender evaluation; a contractor selection recommendation; a commitment of estimated value; or a substitute for a formal Class 3 estimate."*

**Standard form:**

**CONFIDENCE STAMP**
- This deliverable is suitable for: [use case 1, use case 2, use case 3]
- The [P50 / midpoint / recommended value] is appropriate for: [decision type, audience]
- Boundary values represent reasonable scenarios but should not be treated as hard limits.

**WHAT THIS IS NOT:**
- Not [misuse case 1]
- Not [misuse case 2]
- Not [misuse case 3]

**Why it matters:** Reader leaves with the affirmative case for the document. Confidence is declared, not implied. Limitations frame the boundaries, not the deliverable.

---

### Element 7 - Filing Path At The Bottom

**What it is:** Every deliverable closes with explicit filing instructions per the Ausenco Standard Project Folder Structure (or equivalent for non-Ausenco programs).

**Example from James's Section 10:**
> *"Filing per Ausenco Standard Project Folder Structure:*
> *- 05 Reporting / 05.02 Benchmark Reports - primary location for this memorandum and supporting Excel model.*
> *- 02 Estimates / 02.04 Benchmarks - cross-referenced copy.*
> *- 10 Constructability / 10.01 C&C FEED Intake - C&C internal reference copy."*

**Ausenco Standard Project Folder Structure (typical):**

| Folder | Use |
|---|---|
| 01 Project Management | Charter, governance, PEP |
| 02 Estimates | Cost estimates, benchmarks, BoE |
| 03 Engineering | Discipline deliverables |
| 04 Procurement | Tenders, contracts, vendor data |
| 05 Reporting | Memos, status, executive briefings |
| 06 Schedule | IMS, lookaheads, baselines |
| 07 Risk | Registers, MC analysis |
| 08 Quality | ITPs, audits |
| 09 HSE | Plans, statistics |
| 10 Constructability | C&C reviews, FEED intake, peer reviews |
| 11 Commissioning | C0-C6 procedures |
| 12 Stakeholder | First Nations, regulatory, community |

**Standard form:**

**FILING:**
- Primary: [folder / subfolder]
- Cross-reference: [folder / subfolder]
- Internal reference: [folder / subfolder]

**Why it matters:** Document control is part of senior deliverable discipline. Reader knows where to file, where to find cross-references, where the source-of-truth copy lives. Cuts "where do I save this?" friction.

---

## 4. The Standard Template Header / Footer

Every Ausenco deliverable opens with this header block and closes with this footer block. Adapt for memo, workbook cover, executive briefing as needed.

### HEADER BLOCK (memo / report style)

```
AUSENCO ENGINEERING CANADA ULC
[Project / Engagement Name] | Job No. [######-##]

[DOCUMENT TITLE]

Document No.:    [Per Ausenco scheme: ######-DT-#####-#####-###]
Revision:        [A - Draft for SME Review / B - Issued / etc.]
Status:          [CONFIDENTIAL - INTERNAL / for client / etc.]
Date:            [DD MMM YYYY]
Prepared by:     Rick Miller, Senior Project Director
Document Class:  [AACE Class X or X→Y Transitional, per 18R-97]

CLASSIFICATION STAMP (Element 1):
[One line: "This is a [class] [type] prepared for [purpose]. It is NOT [misuse case 1, 2, 3]."]
```

### FOOTER BLOCK

```
SENSITIVITY (Element 5):
[Top 3 sensitivities table]

CONFIDENCE STAMP (Element 6):
[Affirmative use case statement]

WHAT THIS IS NOT (Element 6):
[Bulleted misuse cases]

FILING (Element 7):
Primary: [folder]
Cross-reference: [folder]
Internal reference: [folder]

— END OF [DOCUMENT TYPE] —

Ausenco Engineering Canada ULC | Job [######-##] | [DD MMM YYYY] | Page X of Y
```

### WORKBOOK COVER TAB STANDARD

```
Tab 00 Cover:
  [DOCUMENT TITLE]
  Document No.    [######-DT-#####-#####-###]
  Revision        [Rev / Date]
  Class           [AACE Class X / 18R-97]
  Prepared by     Rick Miller, Senior Project Director
  
  CLASSIFICATION STAMP:
  [One line per Element 1]
  
  CONFIDENCE STAMP:
  [Affirmative use case per Element 6]
  
  WHAT THIS IS NOT:
  [Misuse cases per Element 6]
  
  FILING:
  [Per Element 7]
  
  TAB INDEX:
  01 Methodology - rate basis, adjustments (Element 2 decomposition)
  02 Rate Library - clean reusable asset (Element 3)
  03 Project Application - library applied to MTO (Element 3)
  04 Roll-Up - WBS / area aggregation
  05 Sensitivity & Flags - Element 4 resolution table + Element 5 sensitivity
  06 Sources & Provenance - citation chain
  07+ Project-specific tabs as needed
```

---

## 5. Activation Rule

This SKILL activates automatically when:
- Output is going under Ausenco attribution (any Ausenco program: Goldboro, Green Bay, Goderich, Ausenco BD)
- Request mentions "Ausenco deliverable", "FD-grade", "letterhead-grade", "James", "Wilford", "Medley", "Sterling", "Gallant"
- Document type is one of: benchmark, cost estimate, bid evaluation, executive briefing, gate readiness, PEP, schedule review, RFP audit, fee proposal, owner's rep memo
- Document is going to Ausenco internal leadership (Dan Wilford, Rob Medley, James Gallant, Grahame Sterling, etc.) or to a NexGold / FireFly Metals / Compass client contact under Ausenco attribution

**Does NOT activate for:**
- Carter's Admin internal documents (use Carter's letterhead standard)
- Centerra / LMC / Kemess work (use Carter's/TMG combined standard with LMC privilege protocols)
- Conversational responses or quick-turn analysis

**When activation is ambiguous:** default to activating. The seven elements never hurt a deliverable; the absence hurts it.

---

## 6. Audit Checklist (Run Before Issuing Any Ausenco Deliverable)

Before any Ausenco-attributed document goes out:

| # | Element | Check | Pass / Fail |
|---|---|---|---|
| 1 | Classification stamp | One line at top stating class + purpose + NOT-for cases? | |
| 2 | Adjustments decomposed | Any rate/cost adjustment shown component-by-component with ranges? | |
| 3 | Library / Application separation | Reusable rate library on a separate tab from project application? | |
| 4 | FLAG taxonomy actionable | Every flag has Resolution Required column? | |
| 5 | Top 3 sensitivities | Pre-calculated with dollar impact at deliverable level? | |
| 6 | Confidence stamp | Affirmative use case stated BEFORE limitations text? | |
| 7 | Filing path | Standard Project Folder Structure path declared at bottom? | |
| 8 | Document number | Per Ausenco scheme ######-DT-#####-#####-###? | |
| 9 | Author | Rick Miller, Senior Project Director (no P.Eng, no Engineer)? | |
| 10 | Entity attribution | Ausenco Engineering Canada ULC (no Carter's/TMG mixing)? | |

All ten = ship. Any fail = re-work before issuing.

---

## 7. Worked Example Cross-Reference

The 2026 Heavy Civil Benchmark workbook (RMM-BENCHMARK-CANADA-2026-V4.2) was Rick's baseline before absorbing this standard. V5 of that workbook will be rebuilt to this standard:
- V5 Tab 00 Cover - adds classification stamp per Element 1
- V5 Tab 01 Methodology - adds adjustment decomposition table per Element 2
- V5 Tab 02 Rate Library (new, split from Tab 01 Rate Schedule) per Element 3
- V5 Tab 03 Project Application (the rate library applied to specific project context) per Element 3
- V5 Tab 04 Flag Register - adds Resolution Required column per Element 4
- V5 Tab 05 Sensitivity - pre-calculates top 3 with dollar impact per Element 5
- V5 Tab 00 Cover - adds Confidence Stamp + What This Is Not block per Element 6
- V5 Tab 00 Cover - adds Filing path per Element 7

V5 is the upgraded benchmark workbook produced AFTER this SKILL is in effect.

---

## 8. Source

Standard absorbed from J. Gallant (Functional Director Construction & Commissioning, North America, Ausenco) parallel Goldboro C5027 work issued 15-19 May 2026:
- Goldboro_Earthworks_Benchmark_Memorandum.docx (15 May 2026)
- Goldboro_Benchmark.xlsx (5-sheet supporting model)
- Goldboro_Bird_Evaluation_Memorandum_1.docx + workbook
- Goldboro_Dexter_Evaluation_Memorandum.docx + workbook
- Goldboro_GIP_Evaluation_Memorandum.docx + workbook
- Goldboro_C5027_Consolidation_Memo.docx (19 May 2026)
- Goldboro_C5027_Consolidation.xlsx (8-sheet consolidation workbook)

Rick reviewed against own 2026 Heavy Civil Benchmark V4.2 issued 18 May 2026 to Bob and observed seven presentation-level gaps. This SKILL closes those gaps as a durable standard.

---

— END OF SKILL —

Filing on Rick's side:
Primary: C:\Users\rickm\keystone\skills\ausenco-fd-presentation-standard.md
Cross-reference: C:\Users\rickm\OneDrive\Documents\Claude\Projects\AusencoBD\Program Management\Presentation Standards\
