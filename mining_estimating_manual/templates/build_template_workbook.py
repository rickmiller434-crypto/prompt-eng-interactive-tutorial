"""
Build the FD-grade Ausenco workbook cover template (Tab 00 Cover).

Mirrors James Gallant's Goldboro_Benchmark.xlsx Cover tab layout, which is the
canonical example of an Ausenco FD-grade workbook cover.

Output: mining_estimating_manual/templates/template_ausenco_workbook_cover.xlsx
"""

import os
import xlsxwriter

OUT = os.path.join(os.path.dirname(__file__), "template_ausenco_workbook_cover.xlsx")

wb = xlsxwriter.Workbook(OUT)

NAVY = "#0F2F4D"
LIGHT_GREY = "#F4F6F8"
SUB_GREY = "#D9D9D9"
GREEN = "#4F7942"
AMBER = "#C9A227"
RED = "#A63232"

f_co_title = wb.add_format({"bold": True, "font_size": 14, "font_color": "white",
                             "bg_color": NAVY, "align": "left", "valign": "vcenter"})
f_co_sub = wb.add_format({"italic": True, "font_size": 10, "font_color": NAVY,
                           "bg_color": LIGHT_GREY, "align": "left", "valign": "vcenter"})
f_doc_title = wb.add_format({"bold": True, "font_size": 18, "font_color": NAVY,
                              "align": "left", "valign": "vcenter"})
f_doc_sub = wb.add_format({"bold": True, "font_size": 12, "font_color": NAVY,
                            "align": "left", "valign": "vcenter"})
f_confidential = wb.add_format({"bold": True, "font_size": 12, "font_color": "white",
                                 "bg_color": RED, "align": "left", "valign": "vcenter"})
f_section = wb.add_format({"bold": True, "font_size": 11, "font_color": "white",
                            "bg_color": NAVY, "align": "left", "valign": "vcenter"})
f_label = wb.add_format({"bold": True, "font_size": 10, "align": "left", "valign": "top",
                          "border": 1, "border_color": SUB_GREY})
f_value = wb.add_format({"font_size": 10, "align": "left", "valign": "top",
                          "border": 1, "border_color": SUB_GREY, "text_wrap": True})
f_stamp = wb.add_format({"font_size": 10, "italic": True, "bg_color": "#FFFDE7",
                          "align": "left", "valign": "top", "border": 1,
                          "border_color": SUB_GREY, "text_wrap": True})
f_conf_stamp = wb.add_format({"font_size": 10, "italic": True, "bg_color": "#E8F5E9",
                               "align": "left", "valign": "top", "border": 1,
                               "border_color": SUB_GREY, "text_wrap": True})
f_filing = wb.add_format({"font_size": 10, "italic": True, "bg_color": "#E8F1FB",
                           "align": "left", "valign": "top", "border": 1,
                           "border_color": SUB_GREY, "text_wrap": True})


# ============================================================
# Tab 01 — Cover & Summary (Ausenco FD-grade template)
# ============================================================
ws = wb.add_worksheet("01 Cover & Summary")
ws.hide_gridlines(2)
ws.set_column("A:A", 3)
ws.set_column("B:B", 26)
ws.set_column("C:C", 80)

# Header block — entity attribution
ws.set_row(1, 22)
ws.merge_range("B2:C2", "AUSENCO ENGINEERING CANADA ULC", f_co_title)
ws.set_row(2, 18)
ws.merge_range("B3:C3", "[Sub-line: e.g. Construction & Commissioning — North America]", f_co_sub)

# Document title
ws.set_row(4, 26)
ws.merge_range("B5:C5", "[PROJECT NAME — e.g. NEXGOLD GOLDBORO GOLD PROJECT]", f_doc_title)
ws.set_row(5, 22)
ws.merge_range("B6:C6", "[Document descriptor — e.g. Site Earthworks Tender — Independent C&C Benchmark]", f_doc_sub)
ws.set_row(6, 22)
ws.merge_range("B7:C7", "CONFIDENTIAL — INTERNAL", f_confidential)

# Document control block
fields = [
    ("Project",          "[Project name + study phase]"),
    ("Job Number",       "[######-##]"),
    ("Owner",            "[Owner Mining Corp.]"),
    ("Location",         "[City, Province — distance from major centre]"),
    ("EPCM",             "Ausenco Engineering Canada ULC"),
    ("Tender Package",   "[Ausenco doc number — ######-DT-#####-#####-###]"),
    ("Tender Rev",       "[Rev letter — date]"),
    ("Document No.",     "[Ausenco scheme — ######-DT-#####-#####-###]"),
    ("Prepared by",      "Rick Miller, Senior Project Director"),
    ("Status",           "DRAFT — for SME review prior to issue"),
    ("Date",             "[DD MMM YYYY]"),
    ("Document Class",   "[AACE Class X or X→Y Transitional, per 18R-97]"),
]
r = 8
for label, value in fields:
    ws.set_row(r, 22)
    ws.write(r, 1, label, f_label)
    ws.write(r, 2, value, f_value)
    r += 1

# Element 1 — Classification Stamp
r += 1
ws.set_row(r, 22)
ws.merge_range(r, 1, r, 2, "CLASSIFICATION STAMP  (Element 1)", f_section)
r += 1
ws.set_row(r, 48)
ws.merge_range(r, 1, r, 2,
    "This [deliverable type] is NOT [misuse case 1], [misuse case 2], or [misuse case 3]. "
    "It is intended for [intended use case — e.g. internal Ausenco use to inform contingency setting, "
    "owner reporting bands, and SME-led review of the tender package prior to bid receipt].",
    f_stamp)
r += 2

# Element 6 — Confidence Stamp (affirmative)
ws.set_row(r, 22)
ws.merge_range(r, 1, r, 2, "CONFIDENCE STAMP  (Element 6 — affirmative use case)", f_section)
r += 1
ws.set_row(r, 60)
ws.merge_range(r, 1, r, 2,
    "Suitable for [affirmative use case 1], [affirmative use case 2], and [affirmative use case 3]. "
    "The [P50 / mid-band / recommended value] is appropriate for [decision type, audience]. "
    "Boundary values represent reasonable scenarios but should not be treated as hard limits. "
    "The expected accuracy band is consistent with AACE Class [X] narrowing toward Class [Y] as "
    "scope matures through the Estimate Review process.",
    f_conf_stamp)
r += 2

# What this is NOT (Element 6 continued)
ws.set_row(r, 22)
ws.merge_range(r, 1, r, 2, "WHAT THIS IS NOT", f_section)
r += 1
ws.set_row(r, 48)
ws.merge_range(r, 1, r, 2,
    "Not [misuse case 1, e.g. a tender evaluation]; "
    "not [misuse case 2, e.g. a contractor selection recommendation]; "
    "not [misuse case 3, e.g. a commitment of estimated value]; "
    "not [misuse case 4, e.g. a substitute for a formal Class [X] estimate].",
    f_stamp)
r += 2

# Element 7 — Filing path
ws.set_row(r, 22)
ws.merge_range(r, 1, r, 2, "FILING — Ausenco Standard Project Folder Structure  (Element 7)", f_section)
r += 1
ws.set_row(r, 22)
ws.write(r, 1, "Primary",          f_label)
ws.write(r, 2, "05 Reporting / 05.02 [Sub-folder]",     f_filing)
r += 1
ws.set_row(r, 22)
ws.write(r, 1, "Cross-reference",  f_label)
ws.write(r, 2, "02 Estimates / 02.04 [Sub-folder]",     f_filing)
r += 1
ws.set_row(r, 22)
ws.write(r, 1, "Internal reference", f_label)
ws.write(r, 2, "10 Constructability / 10.01 [Sub-folder]", f_filing)
r += 2

# Tab index
ws.set_row(r, 22)
ws.merge_range(r, 1, r, 2, "TAB INDEX  (typical 5-tab Ausenco benchmark workbook — per James Gallant pattern)", f_section)
r += 1
tab_index = [
    ("01 Cover & Summary",       "This tab — document control, classification, confidence, filing"),
    ("02 Rate Basis Library",    "REUSABLE ASSET — clean rates by category × UoM × region with derivation. No project-specific data."),
    ("03 MTO × Rate",            "PROJECT APPLICATION — line-item MTO with library rates applied, by WBS, with unrated flags"),
    ("04 Area / WBS Roll-Up",    "Project-specific aggregation; chart by area; rated coverage %"),
    ("05 Sensitivity & Flags",   "Top 3 sensitivities pre-calculated with $ impact + flag register with Resolution Required column + SME sign-off block"),
]
for tab, desc in tab_index:
    ws.set_row(r, 30)
    ws.write(r, 1, tab, f_label)
    ws.write(r, 2, desc, f_value)
    r += 1

# Footer
r += 1
ws.set_row(r, 18)
ws.merge_range(r, 1, r, 2, "— END OF COVER TAB —    Ausenco Engineering Canada ULC | Job [######-##] | [DD MMM YYYY]", f_co_sub)


# ============================================================
# Tab 02 — Rate Basis Library (skeleton)
# ============================================================
ws = wb.add_worksheet("02 Rate Basis Library")
ws.hide_gridlines(2)
ws.set_column("A:A", 12)
ws.set_column("B:B", 50)
ws.set_column("C:C", 8)
ws.set_column("D:F", 12)
ws.set_column("G:G", 40)
ws.set_column("H:H", 24)
ws.set_row(0, 26)
ws.merge_range("A1:H1", "RATE BASIS LIBRARY — REUSABLE ASSET (Element 3)", f_co_title)
ws.set_row(1, 18)
ws.merge_range("A2:H2",
    "Canada-wide rate library by category × UoM × region. Carry as Low / Mid / High band. "
    "Independent of any single project — this is the asset that compounds across engagements.",
    f_co_sub)
hdr = ["Category", "Description / Scope", "UoM", "Low (CAD)", "Mid (CAD)", "High (CAD)",
        "Derivation / source citation", "Region applicability"]
for c, h in enumerate(hdr):
    ws.set_row(2, 28)
    ws.write(2, c, h,
              wb.add_format({"bold": True, "font_size": 10, "font_color": "white",
                              "bg_color": NAVY, "align": "center", "border": 1, "text_wrap": True}))
ws.set_row(3, 22)
for c, v in enumerate(["[Example: Earthworks]", "[e.g. Common excavation — load + haul 1-3 km]",
                       "m³", "[low]", "[mid]", "[high]",
                       "[Citation: Marathon Cu-Pd 2024-26 + Lynn Lake LLGP 2025 avg; -5% NS composite adj per Element 2]",
                       "[NL / NS / NB; ON / MB-SK baseline; AB / BC; YT / NWT]"]):
    ws.write(3, c, v, f_value)


# ============================================================
# Tab 03 — MTO × Rate (skeleton)
# ============================================================
ws = wb.add_worksheet("03 MTO x Rate")
ws.hide_gridlines(2)
ws.set_column("A:A", 8)
ws.set_column("B:B", 10)
ws.set_column("C:C", 50)
ws.set_column("D:D", 8)
ws.set_column("E:E", 12)
ws.set_column("F:H", 12)
ws.set_column("I:I", 24)
ws.set_column("J:J", 32)
ws.set_row(0, 26)
ws.merge_range("A1:J1", "MTO × RATE — PROJECT APPLICATION (Element 3)", f_co_title)
ws.set_row(1, 18)
ws.merge_range("A2:J2",
    "Project-specific line-item MTO with library rates applied. Bidder columns live here, NOT in the library.",
    f_co_sub)
hdr = ["WBS Area", "Code", "Description", "UoM", "Quantity",
        "Library Rate (Mid)", "Direct Cost (Mid)", "Library Rate (High)",
        "Pattern Match", "Notes / Bidder Comparison"]
for c, h in enumerate(hdr):
    ws.set_row(2, 28)
    ws.write(2, c, h,
              wb.add_format({"bold": True, "font_size": 10, "font_color": "white",
                              "bg_color": NAVY, "align": "center", "border": 1, "text_wrap": True}))


# ============================================================
# Tab 04 — Area Roll-Up
# ============================================================
ws = wb.add_worksheet("04 Area Roll-Up")
ws.hide_gridlines(2)
ws.set_column("A:A", 12)
ws.set_column("B:B", 30)
ws.set_column("C:G", 14)
ws.set_row(0, 26)
ws.merge_range("A1:G1", "AREA / WBS ROLL-UP", f_co_title)
ws.set_row(1, 18)
ws.merge_range("A2:G2",
    "Project-specific aggregation by WBS area. Rated coverage % flagged. Chart at right.",
    f_co_sub)


# ============================================================
# Tab 05 — Sensitivity & Flags (with Element 4 Resolution Required + Element 5 top 3)
# ============================================================
ws = wb.add_worksheet("05 Sensitivity & Flags")
ws.hide_gridlines(2)
ws.set_column("A:A", 6)
ws.set_column("B:B", 32)
ws.set_column("C:C", 22)
ws.set_column("D:D", 18)
ws.set_column("E:E", 50)
ws.set_column("F:F", 14)
ws.set_row(0, 26)
ws.merge_range("A1:F1", "SENSITIVITY & FLAGS (Element 4 + Element 5)", f_co_title)

# Element 5 — Top 3 sensitivities
ws.set_row(2, 22)
ws.merge_range("A3:F3", "TOP 3 SENSITIVITIES — Pre-Calculated With $ Impact  (Element 5)", f_section)
hdr5 = ["#", "Sensitivity", "Trigger", "$ Impact", "Action if triggered", "Owner"]
for c, h in enumerate(hdr5):
    ws.set_row(3, 24)
    ws.write(3, c, h,
              wb.add_format({"bold": True, "font_size": 10, "font_color": "white",
                              "bg_color": NAVY, "align": "center", "border": 1}))
sens_rows = [
    ("1", "[Rate market shift]",   "[±X% uniform shift]",       "[±$X M total carried]",        "[Re-issue with updated escalation]", "[Owner]"),
    ("2", "[Scope split decision]","[Owner-mined vs Contractor]","[$Y M reduction if decision]", "[Tender scope clarification]",        "[Owner]"),
    ("3", "[Spec upgrade]",        "[NSE LLDPE compliance]",    "[+$Z M direct (+$W M carried)]","[Update geomembrane spec basis]",     "[Owner]"),
]
r = 4
for row in sens_rows:
    ws.set_row(r, 28)
    for c, v in enumerate(row):
        ws.write(r, c, v, f_value)
    r += 1

# Element 4 — Flag register with Resolution Required
r += 1
ws.set_row(r, 22)
ws.merge_range(r, 0, r, 5, "FLAG REGISTER — Each Flag Has Resolution Required  (Element 4)", f_section)
r += 1
hdr4 = ["#", "Flag", "Severity", "Status", "Resolution Required", "Owner / By-When"]
for c, h in enumerate(hdr4):
    ws.set_row(r, 24)
    ws.write(r, c, h,
              wb.add_format({"bold": True, "font_size": 10, "font_color": "white",
                              "bg_color": NAVY, "align": "center", "border": 1}))
r += 1
ws.set_row(r, 40)
for c, v in enumerate(["1", "[Named flag — e.g. Pit pre-strip scope split]",
                       "HIGH", "OPEN",
                       "[Specific action — e.g. Confirm whether contractor or owner mining scope. If owner-mined, deduct from this package. Largest single bucket at $X M direct.]",
                       "[Named owner — by date]"]):
    ws.write(r, c, v, f_value)
r += 2

# Element 6 — SME sign-off block
ws.set_row(r, 22)
ws.merge_range(r, 0, r, 5, "SME SIGN-OFF BLOCK", f_section)
r += 1
sign_hdr = ["Role", "Name", "Signature", "Date"]
hdr_fmt = wb.add_format({"bold": True, "font_size": 10, "font_color": "white",
                          "bg_color": NAVY, "align": "center", "border": 1})
ws.set_row(r, 22)
ws.merge_range(r, 0, r, 1, "Role", hdr_fmt)
ws.write(r, 2, "Name", hdr_fmt)
ws.write(r, 3, "Signature", hdr_fmt)
ws.write(r, 4, "Date", hdr_fmt)
r += 1
roles = ["Author — Senior Project Director", "Estimating SME", "Project Engineering SME",
          "Construction & Commissioning SME", "Project Manager"]
for role in roles:
    ws.set_row(r, 28)
    ws.merge_range(r, 0, r, 1, role, f_label)
    for c in range(2, 5):
        ws.write(r, c, "", f_value)
    r += 1


wb.close()
print(f"Wrote {OUT}")
print(f"Size: {os.path.getsize(OUT):,} bytes")
