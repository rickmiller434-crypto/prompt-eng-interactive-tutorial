"""
Build the formal PDF version of the Mining Heavy Civil Estimating Manual
& Rate Library (RMM-CIVIL-CANADA-2026-MANUAL-REV1.pdf) directly via
ReportLab — mirroring the content of the Word document.
"""

import os
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, PageBreak, KeepTogether)
from reportlab.pdfgen import canvas

OUT = os.path.join(os.path.dirname(__file__), "..", "RMM-CIVIL-CANADA-2026-MANUAL-REV1.pdf")
OUT = os.path.abspath(OUT)

NAVY = colors.HexColor("#0F2F4D")
ACCENT = colors.HexColor("#C9A227")
GREY = colors.HexColor("#555555")
LIGHT_GREY = colors.HexColor("#F4F6F8")
WHITE = colors.white

DOC_NUMBER = "RMM-CIVIL-CANADA-2026-MANUAL-REV1"

# Styles
ss = getSampleStyleSheet()
title_style    = ParagraphStyle("Title", parent=ss["Title"], fontName="Helvetica-Bold",
                                 fontSize=22, textColor=NAVY, alignment=TA_CENTER, spaceAfter=8)
subtitle_style = ParagraphStyle("Subtitle", parent=ss["Normal"], fontName="Helvetica-Oblique",
                                 fontSize=12, textColor=GREY, alignment=TA_CENTER, spaceAfter=4)
h1_style = ParagraphStyle("H1", parent=ss["Heading1"], fontName="Helvetica-Bold",
                           fontSize=16, textColor=NAVY, spaceBefore=14, spaceAfter=6)
h2_style = ParagraphStyle("H2", parent=ss["Heading2"], fontName="Helvetica-Bold",
                           fontSize=12, textColor=NAVY, spaceBefore=10, spaceAfter=4)
body_style = ParagraphStyle("Body", parent=ss["Normal"], fontName="Helvetica",
                             fontSize=9.5, leading=12.5, alignment=TA_JUSTIFY, spaceAfter=4)
bullet_style = ParagraphStyle("Bullet", parent=body_style, leftIndent=14,
                               bulletIndent=4, alignment=TA_LEFT)
small_style = ParagraphStyle("Small", parent=body_style, fontSize=8.5,
                              textColor=GREY, alignment=TA_LEFT)
table_cell  = ParagraphStyle("TableCell", parent=body_style, fontSize=8.5, leading=10.5,
                              alignment=TA_LEFT, spaceAfter=0)
table_cell_c= ParagraphStyle("TableCellC", parent=table_cell, alignment=TA_CENTER)
table_head  = ParagraphStyle("TableHead", parent=table_cell, fontName="Helvetica-Bold",
                              textColor=WHITE, alignment=TA_CENTER)

# Page footer / header
def on_page(canvas_obj, doc):
    canvas_obj.saveState()
    canvas_obj.setFont("Helvetica", 8)
    canvas_obj.setFillColor(GREY)
    canvas_obj.drawString(2*cm, 1.2*cm, DOC_NUMBER)
    canvas_obj.drawRightString(A4[0] - 2*cm, 1.2*cm, f"Page {canvas_obj.getPageNumber()}")
    canvas_obj.setStrokeColor(GREY)
    canvas_obj.line(2*cm, 1.6*cm, A4[0] - 2*cm, 1.6*cm)
    # Header
    canvas_obj.setFillColor(NAVY)
    canvas_obj.setFont("Helvetica-Bold", 8)
    canvas_obj.drawString(2*cm, A4[1] - 1.2*cm, "RICK MILLER SME STANDARD")
    canvas_obj.setFillColor(GREY)
    canvas_obj.setFont("Helvetica", 8)
    canvas_obj.drawRightString(A4[0] - 2*cm, A4[1] - 1.2*cm,
        "Canada 2026 Mining Heavy Civil Estimating Manual")
    canvas_obj.line(2*cm, A4[1] - 1.6*cm, A4[0] - 2*cm, A4[1] - 1.6*cm)
    canvas_obj.restoreState()

doc = SimpleDocTemplate(OUT, pagesize=A4,
                         leftMargin=2*cm, rightMargin=2*cm,
                         topMargin=2.0*cm, bottomMargin=2.0*cm,
                         title="RMM Mining Heavy Civil Estimating Manual",
                         author="Rick Miller")

story = []

def H1(t): story.append(Paragraph(t, h1_style))
def H2(t): story.append(Paragraph(t, h2_style))
def P(t):  story.append(Paragraph(t, body_style))
def B(t):  story.append(Paragraph(f"• {t}", bullet_style))
def SP(h=0.2): story.append(Spacer(1, h*cm))
def NP():  story.append(PageBreak())
def SMALL(t): story.append(Paragraph(t, small_style))

def TBL(headers, rows, col_widths=None, header_color=NAVY, alt_row=True, first_col_bold=False):
    data = [[Paragraph(h, table_head) for h in headers]]
    for row in rows:
        line = []
        for ci, val in enumerate(row):
            s = "" if val is None else str(val)
            style = table_cell
            if first_col_bold and ci == 0:
                style = ParagraphStyle("tcb", parent=table_cell, fontName="Helvetica-Bold")
            line.append(Paragraph(s, style))
        data.append(line)
    t = Table(data, colWidths=col_widths, repeatRows=1)
    ts = [
        ("BACKGROUND", (0,0), (-1,0), header_color),
        ("TEXTCOLOR",  (0,0), (-1,0), WHITE),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,0), 8.5),
        ("BOTTOMPADDING",(0,0),(-1,0), 5),
        ("TOPPADDING", (0,0), (-1,0), 5),
        ("GRID",       (0,0), (-1,-1), 0.4, colors.HexColor("#999999")),
        ("VALIGN",     (0,0), (-1,-1), "TOP"),
        ("FONTSIZE",   (0,1), (-1,-1), 8.5),
        ("TOPPADDING", (0,1), (-1,-1), 3),
        ("BOTTOMPADDING",(0,1),(-1,-1), 3),
    ]
    if alt_row:
        for i in range(1, len(data)):
            if i % 2 == 0:
                ts.append(("BACKGROUND", (0,i), (-1,i), LIGHT_GREY))
    t.setStyle(TableStyle(ts))
    story.append(t)

# =====================================================================
# COVER
# =====================================================================
story.append(Spacer(1, 4*cm))
story.append(Paragraph("RICK MILLER SME STANDARD", title_style))
story.append(Paragraph("CANADA 2026 MINING HEAVY CIVIL<br/>ESTIMATING MANUAL & RATE LIBRARY", title_style))
story.append(Spacer(1, 0.3*cm))
story.append(Paragraph("Consolidated Schedule of Rates, Bid Support Package & Estimating Standard", subtitle_style))
story.append(Paragraph("AACE 18R-97 Class 3 Baseline | Canada-wide Q2-2026 | All figures CAD", subtitle_style))
story.append(Spacer(1, 2*cm))

control_rows = [
    ("DOCUMENT NUMBER",        DOC_NUMBER),
    ("REVISION",               "Rev 1 — Consolidated"),
    ("PREPARED BY",            "Rick Miller, Senior Project Director / EPCM"),
    ("ISSUE DATE",             "Q2 2026"),
    ("BASE CURRENCY",          "Canadian Dollar (CAD)"),
    ("ESTIMATE CLASS BASIS",   "AACE 18R-97 Class 3 (±20% / +30%)"),
    ("ESCALATION BASIS",       "Ausenco go-by Q2-2020 × 1.25 BCPI → Q2-2026"),
    ("SOURCE DATA",            "RMM-CIVIL-CONTRACTOR-CANADA-2026-REV0 + RMM-CIVIL-CANADA-2026-REV0"),
    ("INTENDED USERS",         "Contractor estimators, EPCM cost engineers, owner cost teams"),
    ("CONFIDENTIALITY",        "Internal — not for public quotation issue without project validation"),
]
TBL(["Field", "Value"], control_rows, col_widths=[5.0*cm, 10.5*cm], first_col_bold=True)
NP()

# =====================================================================
# TOC
# =====================================================================
H1("TABLE OF CONTENTS")
toc = [
    ("1.",  "Executive Summary"),
    ("2.",  "Document Control, Basis & Source Data Provenance"),
    ("3.",  "Estimating Standards & Methodology (AACE 18R-97)"),
    ("4.",  "Crew Compositions & Production Assumptions"),
    ("5.",  "Labour Rates — Canada 2026 (Fully Burdened)"),
    ("6.",  "Equipment Rates — Canada 2026 (Fully Operated)"),
    ("7.",  "Equipment Spreads by Scope [TEMPLATE]"),
    ("8.",  "Master Schedule of Rates (Heavy Civil + Contractor Scope)"),
    ("9.",  "Unit Rate Build-Up Methodology"),
    ("10.", "Bid Screening — Floors, Ceilings & Sense-Checks"),
    ("11.", "Contractor Indirects (CDI) & Overhead Build-Up"),
    ("12.", "Bonding & Insurance"),
    ("13.", "Standby Rates, Acceleration, Winter Premiums & Delay"),
    ("14.", "SME Factors — Design Growth, Wastage, Regional, Escalation"),
    ("15.", "Contractor Bid Cost Structure — 6-Layer Markup Model"),
    ("16.", "Mobilization & Demobilization Logic"),
    ("17.", "Constructability & Execution Assumptions [TEMPLATE]"),
    ("18.", "Scope Inclusions, Exclusions & Clarifications [TEMPLATE]"),
    ("19.", "Bid Leveling & Tender Comparison Framework [TEMPLATE]"),
    ("20.", "Change Order Pricing Methodologies"),
    ("21.", "Project Estimate Roll-Up Template [TEMPLATE]"),
    ("22.", "Risk & Contingency Register [TEMPLATE]"),
    ("23.", "Glossary & Abbreviations"),
    ("24.", "Sources & References"),
]
TBL(["#", "Section"], toc, col_widths=[1.5*cm, 14.0*cm])
NP()

# =====================================================================
# 1. EXECUTIVE SUMMARY
# =====================================================================
H1("1.  EXECUTIVE SUMMARY")

H2("1.1  Purpose")
P("This document is the consolidated, internally-maintained estimating standard and rate library "
  "for Canadian mining heavy civil construction. It is designed to be used by contractors, EPCMs, "
  "and owner cost teams across the full project lifecycle: budgetary pricing, feasibility-class "
  "estimates, tender support, bid evaluation, change pricing, and unit-rate validation.")

H2("1.2  Baseline Content (Extracted from Source Workbooks)")
for line in [
    "12 crew codes (B1–N1) — labour + equipment all-in $/hr with 2020 → 2026 escalation",
    "27 fully-burdened labour positions covering supervision, civil trades, blasting, mechanical, electrical, instrumentation, PM and HSE",
    "30 fully-operated equipment items with fuel burn, mob/demob flags, and effective availability",
    "~78 Heavy Civil unit rates across 8 sections (earthworks → closure)",
    "~22 contractor-scope unit rates (blasting, dewatering, buried services, paving, fencing, signage, rock support, culverts, crushing)",
    "29 bid-screening floors/ceilings for tender evaluation",
    "19 bonding & insurance line items including WCB rates by province",
    "25 contractor overhead build-up items",
    "19 standby & schedule items",
    "SME factors: design growth (Class 4 → 1), wastage, 10 Canadian regional zones, escalation, AACE accuracy ranges",
]:
    B(line)

H2("1.3  Framework Templates (Authored in This Manual)")
P("The following sections are industry-standard scaffolding sections, clearly marked as [TEMPLATE] "
  "in the corresponding Excel tabs. They are not drawn from the source workbooks — they are blank "
  "or sample structures to be populated per project:")
for line in [
    "Equipment spreads by scope (Section 7)",
    "Unit rate build-up calculator (Section 9)",
    "Constructability and execution assumptions (Section 17)",
    "Scope inclusions, exclusions and clarifications register (Section 18)",
    "Bid leveling tender comparison (Section 19)",
    "Change order pricing methodologies (Section 20)",
    "Project estimate roll-up sheet (Section 21)",
    "Risk and contingency register (Section 22)",
]:
    B(line)

H2("1.4  Estimate Class Applicability")
TBL(["AACE Class", "Project Stage", "Application in This Manual"], [
    ("Class 5",  "Concept screening / ROM",          "Use floor/ceiling ranges; apply +15-20% design growth"),
    ("Class 4",  "Pre-feasibility / study",           "Mostly stochastic; equipment-factored; +15-20% DG"),
    ("Class 3",  "Feasibility / budget (baseline)",   "Mid-range unit rates; +10-15% DG"),
    ("Class 2",  "Sanction / AFE / pre-FID",          "Use mid-low range; +5-10% DG; firm quotes long-leads"),
    ("Class 1",  "Tender / check estimate",           "Project-specific RFQ; low end of range; +2-5% DG"),
], col_widths=[2.5*cm, 5.0*cm, 8.0*cm])

H2("1.5  Key 2026 Benchmarks (mid-range)")
for line in [
    "Mass earthworks rock (blasted): $55–$120 / m³",
    "Engineered fill compacted: $22–$45 / m³",
    "Haul road full build (sub-grade to surface): $65–$120 / m²",
    "TMF embankment Zone A/B: $22–$45 / m³",
    "HDPE 2 mm liner (full S+I+CQA): $28–$55 / m²",
    "CSP culvert 1200 mm: $1,000–$1,600 / m",
    "Concrete wall / grade beam: $2,500–$3,200 / m³",
    "Rebar Grade 400 (S+I): $4,000–$6,500 / tonne — CBSA 25% surtax exposure",
    "Performance bond medium contract ($5–$25M): 1.0–1.8% of contract value",
    "Regional adjustor range: 1.00 (MB/SK baseline) → 1.55 (Yukon/NWT/Nunavut)",
    "Winter construction premium: 10–25% (all Canada)",
]:
    B(line)

H2("1.6  Boundaries of Use")
P("This manual provides a defensible baseline for internal estimating and bid evaluation. It is not "
  "a public quotation. Do not issue any unit rate as fixed pricing without project-specific "
  "validation, including supplier RFQs for long-lead and bulk materials, confirmation of WCB "
  "classification by province, verification of CBSA surtax exposure on steel and rebar, and "
  "site-specific productivity assumptions.")
NP()

# =====================================================================
# 2. DOC CONTROL
# =====================================================================
H1("2.  DOCUMENT CONTROL, BASIS & SOURCE DATA PROVENANCE")
H2("2.1  Source Workbooks")
TBL(["File Code", "Title", "Content"], [
    ("RMM-CIVIL-CANADA-2026-REV0",            "EPCM Benchmark Estimator (Rev E)",
     "Crew Rates, Labour Rates, Equipment Rates, Heavy Civil Unit Rates, SME Factors"),
    ("RMM-CIVIL-CONTRACTOR-CANADA-2026-REV0", "Contractor Bid Tool (Rev G)",
     "Bid Summary, Additional Unit Rates, Bonding & Insurance, Overhead, Standby, Bid Screening"),
], col_widths=[5.0*cm, 5.0*cm, 5.5*cm])

H2("2.2  Currency, Base Period & Exchange")
P("All figures are in Canadian Dollars (CAD), basis Q2-2026. USD/CAD reference: 1.36–1.37 (Bank of Canada, May 2026).")

H2("2.3  Escalation Basis")
P("Unit rates inherited from a Q2-2020 go-by have been escalated by a single cumulative factor of "
  "x1.25, representing the BCPI midpoint to Q2-2026. Rates re-baselined to 2026 directly (per "
  "supplier quotes or current published WCB / fuel tables) are flagged in each tab's Source/Basis column.")

H2("2.4  Confidentiality & Use")
P("This document is intended for internal estimating, bid support and benchmarking. Distribution "
  "beyond a project team should be limited and approved. Any external use must include a current "
  "price-validity statement and exclusions register (Section 18).")
NP()

# =====================================================================
# 3. ESTIMATING STANDARDS
# =====================================================================
H1("3.  ESTIMATING STANDARDS & METHODOLOGY (AACE 18R-97)")
H2("3.1  Estimate Classification Framework")
P("This manual adopts AACE International Recommended Practice 18R-97 as its classification framework. "
  "The 5-class scheme is interpreted for mining heavy civil work as follows:")
TBL(["Class", "Stage", "Def.", "Methodology", "Accuracy", "Use"], [
    ("Class 5", "Concept screening",         "0–2%",   "Stochastic / parametric",          "-50% / +100%", "ROM"),
    ("Class 4", "Pre-feasibility",           "1–15%",  "Mostly stochastic; equip-factored","-30% / +50%",  "Concept evaluation"),
    ("Class 3", "Feasibility / budget",      "10–40%", "Semi-detailed; unit rates; MTOs", "-20% / +30%",  "Budget (this baseline)"),
    ("Class 2", "Sanction / AFE",            "30–75%", "Detailed; firm long-lead quotes", "-10% / +20%",  "FID approval"),
    ("Class 1", "Tender / check",            "65–100%","Detailed; IFC drawings; RFQs",    "-5% / +10%",   "Tender bid"),
], col_widths=[1.5*cm, 2.8*cm, 1.5*cm, 4.5*cm, 2.2*cm, 3.0*cm])

H2("3.2  Design Growth Allowances")
TBL(["Class", "Discipline", "Low", "High", "Notes"], [
    ("Class 5 / 4", "Earthworks / civil",                "15%", "20%", "Limited geotech; preliminary drawings"),
    ("Class 3",     "Earthworks / civil (this baseline)","10%", "15%", "Standard level for feasibility / budget"),
    ("Class 2",     "Earthworks / civil",                "5%",  "10%", "Pre-FID; IFC packages near complete"),
    ("Class 1",     "Earthworks / civil",                "2%",  "5%",  "Full detailed design completed"),
    ("All classes", "Concrete works (additive)",         "+5%", "+5%", "Add 5% incremental to earthworks DG"),
], col_widths=[2.0*cm, 4.5*cm, 1.5*cm, 1.5*cm, 6.0*cm])

H2("3.3  Wastage Factors")
TBL(["Material", "Low", "High", "Basis"], [
    ("Granular aggregate",  "5%",  "10%", "Rounding + delivery variance"),
    ("Concrete (formed)",   "8%",  "12%", "Pump waste + edge loss"),
    ("Rebar",               "3%",  "7%",  "Cutting waste; depends on bar size and complexity"),
    ("Geomembrane liner",   "5%",  "10%", "Panel overlaps + anchor trench + QA sampling"),
    ("Geotextile",         "10%",  "15%", "ASTM D4439 — 6-inch overlap minimum"),
], col_widths=[4.5*cm, 1.5*cm, 1.5*cm, 8.0*cm])

H2("3.4  Escalation & Market Basis (Q2-2026)")
TBL(["Item", "Low", "High", "Source"], [
    ("Diesel — NS (May 2026)",                 "$2.13/L", "$2.21/L", "NS Energy Regulatory Board"),
    ("Diesel — AB (May 2026)",                 "$1.68/L", "$1.78/L", "NRCan"),
    ("Diesel — BC Metro (May 2026)",           "$2.15/L", "$2.35/L", "NRCan"),
    ("USD/CAD (May 2026)",                     "1.36",    "1.37",    "Bank of Canada"),
    ("CBSA surtax — Chinese structural steel", "25%",     "25%",     "CBSA Notices 24-26 / 25-22"),
    ("CBSA surtax — Chinese rebar",            "25%",     "25%",     "Confirm supplier origin certificate"),
    ("Labour CBA escalation 2024–2026 (NL)",   "3.5%",    "5.5%",    "Per trade — see Section 5"),
], col_widths=[6.5*cm, 2.0*cm, 2.0*cm, 5.0*cm])
NP()

# =====================================================================
# 4. CREW COMPOSITIONS
# =====================================================================
H1("4.  CREW COMPOSITIONS & PRODUCTION ASSUMPTIONS")
P("Twelve standard crew codes are defined. Crew rates are reported both at the 2020 go-by basis and at the 2026 escalated basis (x1.25). Productivity (Prod Factor) is applied at the unit rate level (Section 8).")
crews = [
    ("B1", "Civil Works — Light",        "26.1", "$127.50", "$159.37"),
    ("B2", "Civil Works — Medium",       "28.1", "$158.90", "$198.63"),
    ("B3", "Civil Works — Heavy",        "37.7", "$161.30", "$201.63"),
    ("B4", "Civil Liner & Pipeline",     "28.1", "$116.10", "$145.12"),
    ("C1", "Concrete",                   "22.6", "$109.40", "$136.75"),
    ("D1", "Structural Steel",           "17.6", "$114.10", "$142.63"),
    ("E1", "Architectural",              "19.6", "$100.30", "$125.37"),
    ("G1", "Mechanical Services",        "17.7", "$111.30", "$139.13"),
    ("H1", "Mechanical Equipment",       "23.6", "$115.80", "$144.76"),
    ("J1", "Piping",                     "13.7", "$110.50", "$138.12"),
    ("K1", "Electrical",                 "10.8", "$117.80", "$147.25"),
    ("N1", "Instrumentation",            "10.4", "$123.80", "$154.75"),
]
TBL(["Code", "Description", "Avg Size", "2020 Total $/hr", "2026 Total $/hr"], crews,
    col_widths=[1.5*cm, 6.5*cm, 2.0*cm, 3.0*cm, 3.0*cm])
SMALL("Crew composition by trade is shown in Section 5. The crews above are SME averages for benchmarking. For tender or sanction-class estimates, build the project-specific crew from Section 5 labour rates and Section 6 equipment rates.")
NP()

# =====================================================================
# 5. LABOUR RATES
# =====================================================================
H1("5.  LABOUR RATES — CANADA 2026 (FULLY BURDENED)")
P("Labour rates are fully burdened: base wage + statutory burdens + employer benefits + WCB (province-specific) + small tools allowance + supervision ratio. Camp/LOA (rows O2/O3) is shown as additive per-person cost.")
labour = [
    ("S1","General Superintendent","Supervisor","$154","$172"),
    ("S2","Superintendent","Supervisor","$136","$155"),
    ("S3","General Foreman","Supervisor","$125","$140"),
    ("C1","Civil Work Lead","Civil","$116","$132"),
    ("C2","Equip Op — Heavy Duty","Civil","$110","$128"),
    ("C3","Equip Op — Medium Duty","Civil","$101","$116"),
    ("C4","Equip Op — Light Duty","Civil","$93","$106"),
    ("C5","Labourer — Journeyman","Civil","$82","$95"),
    ("C6","Labourer — 2yr Apprentice","Civil","$73","$84"),
    ("C7","Labourer — 1st yr Apprentice","Civil","$64","$74"),
    ("D1","Driller — Journeyman","Civil/Blast","$93","$108"),
    ("D2","Blaster — Journeyman","Civil/Blast","$93","$108"),
    ("M1","Mechanic — Heavy Duty","Mechanical","$84","$98"),
    ("M2","Mechanic — Medium Duty","Mechanical","$92","$106"),
    ("T1","Carpenter — Journeyman","Trades","$109","$125"),
    ("T3","Concrete Lead","Trades","$125","$140"),
    ("T4","Concrete Finisher","Trades","$95","$108"),
    ("T5","Ironworker (rebar/struct)","Trades","$115","$132"),
    ("T6","Welder / Pipefitter","Trades","$113","$130"),
    ("T8","Rodman — Reinforcement","Trades","$94","$108"),
    ("E1","Electrician — Journeyman","Electrical","$127","$145"),
    ("I1","Instrumentation Tech","Instrumentation","$127","$145"),
    ("P1","Project Engineer — Field","PM/Eng","$109","$125"),
    ("P2","Project Engineer — Senior","PM/Eng","$123","$140"),
    ("O1","Safety Officer (NCSO)","HSE","$110","$125"),
    ("O2","Camp — all-in per person/day","Indirect","$220","$300"),
    ("O3","LOA subsistence — non-camp","Indirect","$180","$250"),
]
TBL(["Ref", "Position", "Class", "2026 Low", "2026 High"], labour,
    col_widths=[1.3*cm, 6.5*cm, 3.0*cm, 2.5*cm, 2.5*cm])
NP()

# =====================================================================
# 6. EQUIPMENT RATES
# =====================================================================
H1("6.  EQUIPMENT RATES — CANADA 2026 (FULLY OPERATED)")
P("Equipment rates are fully operated. Fuel is L/hr; multiply by site diesel price for fuel-cost portion. Availability % drives effective hours per shift.")
equip = [
    ("HT","Tandem dump truck","18-20T","$140","$165","45-65","Required","85%"),
    ("HT","Tri-axle dump truck","25T","$155","$180","55-75","Required","85%"),
    ("HT","Wiggle wagon / Super B","45T","$185","$215","65-90","Required","80%"),
    ("HT","Off-road haul truck","40-100T","$320","$500","80-140","Included","85%"),
    ("EX","Hyd excavator Cat 320","20-25T","$250","$330","18-25","Required","85%"),
    ("EX","Hyd excavator Cat 336","35-40T","$340","$430","25-35","Required","85%"),
    ("EX","Hyd excavator Cat 390","50-65T","$460","$580","38-50","Required","82%"),
    ("EX","Long-reach excavator","20T 18m","$290","$380","18-25","Required","82%"),
    ("DO","Track dozer Cat D6N","17T","$250","$320","20-30","Required","88%"),
    ("DO","Track dozer Cat D8T","37T","$380","$480","35-48","Required","88%"),
    ("DO","Track dozer Cat D10T","58T","$520","$650","50-68","Required","85%"),
    ("GR","Motor grader Cat 140M","-","$280","$360","18-26","Required","88%"),
    ("GR","Motor grader Cat 16M","large","$360","$450","22-32","Required","88%"),
    ("CP","Vib roller CS56B","12T","$200","$270","15-22","Required","88%"),
    ("CP","Padfoot CS74B","17T","$240","$310","18-26","Required","88%"),
    ("CP","Plate compactor","0.5-1T","$45","$75","5-8","By truck","90%"),
    ("WL","Wheel loader Cat 930M","10T","$200","$270","18-26","Required","87%"),
    ("WL","Wheel loader Cat 950M","16T","$250","$320","22-30","Required","87%"),
    ("WL","Wheel loader Cat 980M","22T","$320","$400","28-38","Required","87%"),
    ("DR","Rotary drill rig","120mm hole","$550","$750","35-50","Included","80%"),
    ("SK","Skid steer Cat 262D","-","$95","$135","10-15","By truck","90%"),
    ("WP","Diesel submersible pump","4-inch","$55","$85","10-16","By truck","90%"),
    ("WP","Diesel submersible pump","8-inch","$120","$180","20-30","By truck","90%"),
    ("CR","RT crane Grove","RT 60T","$650","$900","30-45","Required","80%"),
    ("CR","Crawler crane Liebherr","150T","$1,800","$2,600","60-90","Included","78%"),
    ("MI","Rock breaker Cat 345+hammer","-","$480","$620","28-40","Required","80%"),
    ("MI","Air compressor diesel","375 cfm","$75","$110","12-18","By truck","90%"),
    ("MI","Light tower diesel","8-10 kW","$35","$55","5-8","By truck","95%"),
    ("MI","Fuel tanker site supply","10,000L","$180","$250","30-40","Required","90%"),
    ("MI","Water truck dust","10,000L","$160","$220","30-40","Required","90%"),
]
TBL(["Cat","Equipment","Size","Low $/hr","High $/hr","Fuel","Mob","Util"], equip,
    col_widths=[1.0*cm, 4.5*cm, 2.0*cm, 1.6*cm, 1.6*cm, 1.4*cm, 1.7*cm, 1.2*cm])
NP()

# =====================================================================
# 7. EQUIPMENT SPREADS
# =====================================================================
H1("7.  EQUIPMENT SPREADS BY SCOPE  [TEMPLATE]")
P("Typical SME spreads for Class 3 estimating. Validate fleet sizing against site access, haul distance, geotechnical conditions, climate window, and crew availability before committing to a bid.")
spreads = [
    ("Mass earthworks — rock",       "1× DR rotary, 1× EX Cat 390, 4–6× off-road 40-100T", "B3", "800–1,500 m³/shift"),
    ("Mass earthworks — common",     "1× EX Cat 336, 4–6× tandem/tri-axle, 1× D8T",         "B3", "1,200–2,000 m³/shift"),
    ("Engineered fill placement",    "1× EX Cat 320 or WL 950M, 2–4× tandem, 1× CS74B",    "B2", "600–1,000 m³/shift"),
    ("Haul road build",              "1× D8T (sub-base), 1× 140M grader (surface)",         "B2", "150–300 m/day"),
    ("HDPE liner install (TMF)",     "1× WL 930M, 1× SK, welding crew",                     "B4", "1,500–3,000 m²/day"),
    ("Culvert install (≥900 mm CSP)","1× EX Cat 320, 1× WL 930M, 2× tandem, dewatering",   "B4", "20–40 m/day"),
    ("Buried services trenching",    "1× EX Cat 320, 1× tandem, 1× WL 930M",                "B4", "60–120 m/day"),
    ("Concrete pour — foundation",   "1× crane RT 60T, 1× concrete pump, 1× vibrator",     "C1", "12–25 m³/day"),
    ("Rock support — shotcrete",     "1× shotcrete pump, 1× EX (mesh handling), 1× SK",    "C1", "120–250 m²/day"),
    ("Drill + blast rock anchors",   "1× drill rig, 1× EX, 1× grout plant",                 "D1", "8–15 anchors/day"),
    ("Buried HDPE forcemain",        "1× EX Cat 320, 1× WL, 1× fusion machine, 1× tandem", "B4", "80–150 m/day"),
    ("Demolition / strip and grub",  "1× D8T, 1× EX Cat 336, 1× tandem",                    "B3", "0.5–1.5 ha/day"),
]
TBL(["Scope", "Primary + Support Equipment", "Crew", "Production"], spreads,
    col_widths=[4.0*cm, 7.0*cm, 1.5*cm, 3.0*cm])
NP()

# =====================================================================
# 8. SoR
# =====================================================================
H1("8.  MASTER SCHEDULE OF RATES")
P("The full line-item Schedule of Rates is provided in the companion Excel workbook (sheet '08 SoR Master'). It contains ~110 line items with unit man-hours, productivity factor, labour / equipment / material / sub build-up columns, CDI %, design growth %, wastage %, and crew code per item. The PDF below summarises the section structure.")
H2("8.1  Section Structure")
TBL(["Section", "Scope", "Source", "Items"], [
    ("A — Earthworks & Mass Excavation",     "Common/rock excavation, fill, granular, drill+blast (holes only)", "Rev E",  "15"),
    ("B — Water Management / SWM & ESC",    "Riprap, geotextile, silt fence, temp dewatering","Rev E",  "10"),
    ("C — Roads, Pads & Working Surfaces",  "Sub-grade, granular base, asphalt, equip pads",         "Rev E",  "8"),
    ("D — WRSA / TMF Embankment",           "Embankment fill, HDPE liner, drainage layers, riprap ditching",   "Rev E",  "8"),
    ("E — Drainage, Culverts & Piping",     "CSP, HDPE, flow guards, discharge piping",                         "Rev E",  "15"),
    ("F — Concrete",                         "Walls, piers, footings, shotcrete, rebar",                         "Rev E",  "5"),
    ("G — Structural Steel",                 "Beams, columns, grating, handrail, stairs, deck, cladding",        "Rev E",  "8"),
    ("H — Closure & Reclamation",           "Cap, cover, topsoil, HDPE cap liner",                              "Rev E",  "4"),
    ("I — Blasting (full scope)",           "Drill, charge, blast, clear, controlled / presplit",                "Rev G",  "4"),
    ("J — Permanent Dewatering",            "Pump stations, forcemains, wellpoint systems",                     "Rev G",  "5"),
    ("K — Buried Services",                  "Water, sewer, manholes, thrust blocks, valves, hydrants",          "Rev G",  "7"),
    ("L — Concrete Paving & Curbing",       "Paving, curb & gutter, wheel wash",                                "Rev G",  "3"),
    ("M — Fencing, Security & Wildlife",    "Chain link, wildlife exclusion electrified, gates",                "Rev G",  "4"),
    ("N — Site Signage & Traffic Mgmt",     "Flagging, dust suppression",                                       "Rev G",  "2"),
    ("O — Rock Anchors & Stabilization",    "Rock anchors, wire mesh, shotcrete + mesh",                        "Rev G",  "4"),
    ("P — Culvert Structures",              "Precast box, arch, headwalls",                                     "Rev G",  "3"),
    ("Q — Aggregate Crushing On-Site",      "Granular A/B, riprap — crushing operating cost",                   "Rev G",  "2"),
], col_widths=[5.0*cm, 7.0*cm, 1.5*cm, 1.5*cm])
NP()

# =====================================================================
# 9. UNIT RATE BUILD-UP METHODOLOGY
# =====================================================================
H1("9.  UNIT RATE BUILD-UP METHODOLOGY")
P("The standard SME unit-rate composition is a six-element build-up:")
TBL(["Element", "Symbol", "Source", "Application"], [
    ("Unit Labour",              "L",  "Crew rate × (Unit MH / Productivity)",          "Section 4–5"),
    ("Unit Equipment",           "E",  "Same crew driver — included in crew total",     "Section 4 / 6"),
    ("Unit Material",            "M",  "Supplier quote × (1 + Wastage)",                "Section 3.3 / supplier RFQ"),
    ("Unit Subcontract",         "S",  "Sub quote (bare, before markup)",                "Per sub"),
    ("Contractor Indirect",     "I",  "% of (L+E+M+S) — CDI overhead",                  "Section 11 / 14"),
    ("Total direct unit rate",   "U",  "= (L + E + M + S) × (1 + I)",                    "Sheet 09 calculator"),
], col_widths=[4.0*cm, 1.5*cm, 6.0*cm, 4.0*cm])

P("Two further factors are applied at unit-rate level before roll-up:")
TBL(["Factor", "Symbol", "Range", "Notes"], [
    ("Design growth %", "DG", "2–20% depending on AACE class", "Apply to quantity, not to unit rate"),
    ("Regional adjustor","R",  "1.00 – 1.55",                  "Apply to direct unit rate × (1+DG)"),
], col_widths=[4.5*cm, 1.5*cm, 4.5*cm, 5.0*cm])
P("Final unit rate = U × (1 + DG) × R. The 6-layer markup stack (Section 15) is then applied at project roll-up, not at line level. The companion Excel workbook (sheet '09 Unit Rate Calculator') provides a live calculator with formulas.")
NP()

# =====================================================================
# 10. BID SCREENING
# =====================================================================
H1("10.  BID SCREENING — FLOORS, CEILINGS & SENSE-CHECKS")
P("Any bid line item priced more than 20% below the floor or 20% above the ceiling warrants a clarification request. Items below the floor most commonly indicate missing scope (insulation, CQA, supply, etc.).")
screen = [
    ("EARTHWORKS",      "Mass earthworks — rock (blasted)",         "$/m³",   "$55",   "$120"),
    ("EARTHWORKS",      "Engineered fill compacted Zone A/B",        "$/m³",   "$22",   "$45"),
    ("EARTHWORKS",      "Granular road base 150mm",                  "$/m²",   "$22",   "$42"),
    ("EARTHWORKS",      "Haul road full build",                       "$/m²",  "$65",   "$120"),
    ("EARTHWORKS",      "Strip and grub per hectare",                "$/ha",   "$8,500","$18,000"),
    ("WRSA / TMF",      "TMF embankment engineered fill",            "$/m³",   "$22",   "$45"),
    ("WRSA / TMF",      "HDPE liner 2mm supply+install",             "$/m²",   "$28",   "$55"),
    ("WRSA / TMF",      "TMF perimeter ditch riprap lined",          "$/m",    "$280",  "$500"),
    ("DRAINAGE",        "CSP culvert 1200mm supply+install",         "$/m",    "$1,000","$1,600"),
    ("DRAINAGE",        "HDPE 150mm above grade insulated",          "$/m",    "$230",  "$380"),
    ("DRAINAGE",        "Flow guard 600mm supply+install",           "$/ea",   "$2,200","$3,500"),
    ("DRAINAGE",        "Flow guard 1200mm supply+install",          "$/ea",   "$8,000","$12,000"),
    ("CONCRETE",        "Concrete foundation/wall",                  "$/m³",   "$2,200","$3,500"),
    ("CONCRETE",        "Shotcrete with fibre 75mm",                 "$/m²",   "$90",   "$160"),
    ("CONCRETE",        "Rebar supply+install",                      "$/tonne","$3,800","$6,500"),
    ("BLASTING",        "Pre-blast survey per structure",            "$/struct","$3,000","$9,000"),
    ("FENCING",         "Wildlife exclusion electrified 2.4m",       "$/m",    "$150",  "$280"),
    ("BURIED SERVICES", "Sanitary manhole precast 1200mm",           "$/ea",   "$4,000","$8,000"),
    ("OVERHEAD CHECK",  "Contractor OH+profit — $50M project",       "% dir.", "14%",   "22%"),
    ("OVERHEAD CHECK",  "Performance + L&M bond (>$25M)",            "% cont.","0.75%","1.25%"),
]
TBL(["Category", "Benchmark Item", "UOM", "Floor", "Ceiling"], screen,
    col_widths=[2.8*cm, 6.5*cm, 1.7*cm, 2.0*cm, 2.0*cm])
NP()

# =====================================================================
# 11. CDI
# =====================================================================
H1("11.  CONTRACTOR INDIRECTS (CDI) & OVERHEAD BUILD-UP")
P("Typical project total: 12–18% of direct cost. Detail below from Rev G source workbook.")
oh = [
    ("FIELD OFFICE",        "Site trailer — lunchroom/dryroom",          "month",  "$1,500–$2,500"),
    ("FIELD OFFICE",        "Temp power — 100 kW diesel generator",      "month",  "$4,500–$8,500"),
    ("FIELD OFFICE",        "Temp power — 250 kW generator large camp", "month",  "$9,500–$16,000"),
    ("FIELD OFFICE",        "Site IT, internet, comms (Starlink)",       "month",  "$850–$2,500"),
    ("FIELD OFFICE",        "Site vehicle — project pickup truck",      "month",  "$1,200–$1,800"),
    ("FIELD OFFICE",        "Site vehicle — 20-seat crew bus",          "month",  "$3,500–$6,500"),
    ("FIELD OFFICE",        "Mob/demob — site setup + teardown",        "LS",     "$35,000–$85,000"),
    ("PM STAFF",            "Senior Project Engineer",                   "month",  "$14,000–$20,000"),
    ("PM STAFF",            "Field / QC Engineer",                       "month",  "$11,000–$16,000"),
    ("PM STAFF",            "Site Safety Officer (NCSO certified)",      "month",  "$12,000–$17,000"),
    ("PM STAFF",            "Senior Superintendent",                     "month",  "$16,000–$24,000"),
    ("PM STAFF",            "General Foreman",                           "month",  "$13,000–$19,000"),
    ("PM STAFF",            "Document Controller / Field Admin",         "month",  "$7,500–$11,000"),
    ("QC & TESTING",        "Nuclear density testing — crew/day",        "day",    "$850–$1,500"),
    ("QC & TESTING",        "Concrete cylinder testing — per set",       "set",    "$185–$350"),
    ("QC & TESTING",        "Aggregate gradation testing",               "test",   "$250–$450"),
    ("QC & TESTING",        "Geomembrane weld testing (spark test)",     "m weld", "$2.50–$4.50"),
    ("QC & TESTING",        "Survey control — survey crew/day",          "day",    "$2,500–$4,500"),
    ("ENVIRONMENTAL",       "Spill kit and containment supplies",        "month",  "$350–$750"),
    ("ENVIRONMENTAL",       "Environmental water sampling",              "sample", "$450–$950"),
    ("ENVIRONMENTAL",       "Environmental coordinator (part-time)",     "month",  "$5,500–$9,500"),
    ("HEAD OFFICE",         "Warranty / deficiency holdback allowance",  "% cont.","0.5%–1.0%"),
]
TBL(["Category", "Item", "UOM", "2026 Range (CAD)"], oh,
    col_widths=[3.5*cm, 6.5*cm, 1.7*cm, 3.8*cm])
NP()

# =====================================================================
# 12. BONDING & INSURANCE
# =====================================================================
H1("12.  BONDING & INSURANCE")
P("WCB is province-specific — confirm classification and base rate against the relevant authority before pricing.")
bi = [
    ("SURETY",            "Performance Bond — small (<$5M)",         "% of contract",        "1.5%–2.5%"),
    ("SURETY",            "Performance Bond — medium ($5–$25M)",      "% of contract",        "1.0%–1.8%"),
    ("SURETY",            "Performance Bond — large (>$25M)",         "% of contract",        "0.75%–1.25%"),
    ("SURETY",            "Labour & Material Payment Bond",           "% of contract",        "0.5%–1.0%"),
    ("SURETY",            "Maintenance/Warranty Bond — 12 month",     "% of contract",        "0.5%–1.0%"),
    ("LIABILITY",         "Contractor's Pollution Liability",         "Annual policy",        "$1,500–$3,500"),
    ("WCB",               "Labourer — Alberta WCB Industry 41",       "Per $100 payroll",     "$3.80–$5.20"),
    ("WCB",               "Labourer — BC WorkSafeBC Class 711001",    "Per $100 payroll",     "$4.50–$6.00"),
    ("WCB",               "Labourer — NS WCB NS",                     "Per $100 payroll",     "$3.80–$5.50"),
    ("WCB",               "Labourer — NL WorkplaceNL",                "Per $100 payroll",     "$4.00–$5.80"),
    ("WCB",               "Labourer — Quebec CNESST",                 "Per $100 payroll",     "$5.50–$8.00"),
    ("WCB",               "Equipment Operator — all provinces",       "Per $100 payroll",     "$3.50–$5.00"),
    ("WCB",               "Blaster / Driller — elevated risk",        "Per $100 payroll",     "$5.50–$8.00"),
    ("PROPERTY",          "Contractor's Equipment Floater",           "% of fleet repl/yr",  "0.5%–1.2%"),
]
TBL(["Category", "Item", "Basis", "Rate (CAD)"], bi,
    col_widths=[2.5*cm, 6.5*cm, 3.5*cm, 3.0*cm])
NP()

# =====================================================================
# 13. STANDBY
# =====================================================================
H1("13.  STANDBY RATES, ACCELERATION & DELAY")
sb = [
    ("EQUIP STANDBY",   "Excavator Cat 320 — standby",      "hr",    "$110–$155"),
    ("EQUIP STANDBY",   "Excavator Cat 336 — standby",      "hr",    "$140–$200"),
    ("EQUIP STANDBY",   "Track dozer D8T — standby",        "hr",    "$155–$215"),
    ("EQUIP STANDBY",   "Motor grader Cat 140M — standby",  "hr",    "$110–$155"),
    ("EQUIP STANDBY",   "RT crane 60T — standby",           "hr",    "$320–$480 (4hr min)"),
    ("EQUIP STANDBY",   "Crawler crane 150T — standby",     "hr",    "$850–$1,300"),
    ("LABOUR STANDBY",  "Equipment operator — standby",     "hr",    "$93–$128"),
    ("LABOUR STANDBY",  "B3 crew of 10 — standby",          "day",   "$7,000–$10,000"),
    ("ACCELERATION",    "Overtime — double time (Sun/holiday)","% add","100% / 100%"),
    ("ACCELERATION",    "Afternoon shift premium",          "% add", "5%–10%"),
    ("ACCELERATION",    "Night shift premium",              "% add", "10%–15%"),
    ("WINTER",          "Cold weather heating — LP heaters","day",   "$450–$850"),
    ("WINTER",          "Cold weather enclosure — concrete","m² enc.","$8–$18"),
    ("WINTER",          "Winter freeze protection — earthworks","m³ thaw","$3.50–$8.50"),
    ("OWNER DELAY",     "Equip demob + remob per event",    "event", "$15,000–$65,000"),
    ("OWNER DELAY",     "Extended project overhead",        "week",  "$12,000–$35,000"),
    ("OWNER DELAY",     "Material restocking / escalation impact","% mat","2%–8% (>3 mo)"),
]
TBL(["Category", "Item", "UOM", "2026 Range (CAD)"], sb,
    col_widths=[3.0*cm, 6.0*cm, 3.0*cm, 3.5*cm])
NP()

# =====================================================================
# 14. SME FACTORS
# =====================================================================
H1("14.  SME FACTORS — REGIONAL ADJUSTORS & MARKUPS")
H2("14.1  Regional Adjustors (Canada 2026)")
P("Apply the regional adjustor to the direct unit rate (after design growth). Baseline is MB/SK = 1.00–1.08.")
reg = [
    ("Ontario (Greater Toronto / Ottawa)",         "1.05", "1.15"),
    ("Ontario (Northern — Thunder Bay / Sudbury)", "1.02", "1.10"),
    ("Quebec (Montreal)",                          "1.03", "1.12"),
    ("Quebec (Northern — Chibougamau / James Bay)","1.08", "1.22"),
    ("Manitoba / Saskatchewan",                    "1.00", "1.08"),
    ("Alberta (Edmonton / Calgary)",               "1.05", "1.18"),
    ("Alberta (Oil Sands / Fort McMurray)",        "1.15", "1.35"),
    ("British Columbia (Metro Vancouver)",         "1.10", "1.20"),
    ("British Columbia (Interior / Northern)",     "1.15", "1.30"),
    ("Yukon / NWT / Nunavut",                      "1.30", "1.55"),
    ("Winter construction premium — all Canada",   "10%",  "25%"),
]
TBL(["Region", "Low", "High"], reg, col_widths=[10.0*cm, 2.5*cm, 2.5*cm])

H2("14.2  Indirect / GC Markup Reference")
TBL(["Item", "Low", "High", "Notes"], [
    ("Mob/demob major earthwork fleet (% of equip)", "2%",      "5%",      "Or flat LS per fleet assembly"),
    ("Mob/demob RT crane 60T",                       "$4,000",  "$8,000",  "Flat rate per mob"),
    ("Mob/demob crawler crane 150T",                 "$25,000", "$55,000", "Incl. partial disassembly"),
    ("Camp/LOA markup on direct labour",             "8%",      "12%",     "Add for remote sites"),
    ("Small tools and consumables (on direct lab.)", "2.5%",     "4%",      "Blades, hoses, bits, PPE"),
    ("Supervision markup on direct labour",          "8%",      "12%",     "Foreman + super ratio"),
    ("Contractor distributable (CDI) on direct cost","6%",      "10%",     "Temp facilities, insurance, site admin"),
], col_widths=[6.5*cm, 1.8*cm, 1.8*cm, 5.0*cm])
NP()

# =====================================================================
# 15. BID COST STRUCTURE
# =====================================================================
H1("15.  CONTRACTOR BID COST STRUCTURE — 6-LAYER MARKUP MODEL")
P("All contractor bids in this manual are composed of six stacked cost layers. The first layer (Direct Field Costs) is built up from the Schedule of Rates (Section 8); the remaining five layers apply % markups in the following sequence:")
TBL(["Layer", "Scope", "Low (%)", "High (%)"], [
    ("1 — Direct Field Costs",                "Unit rates × MTO quantities", "—",       "100% (basis)"),
    ("2 — Contractor Field Indirects (CDI)",   "Temp facilities, supervision, small tools, QC, environmental", "12%", "18%"),
    ("3 — Company Overhead",                  "Bid & proposal, warranty, head office allocation",            "6%",  "12%"),
    ("4 — Bonding & Insurance",               "Performance bond, L&M bond, CGL, WCB, Builder's Risk",        "2%",  "4%"),
    ("5 — Contingency & Risk",                "Ground conditions, weather, escalation, sub default",          "8%",  "15%"),
    ("6 — Contractor Profit",                 "Profit / fee — varies by risk and competitive position",       "8%",  "15%"),
], col_widths=[5.5*cm, 7.0*cm, 1.5*cm, 1.5*cm])
P("Total bid = Direct + CDI + OH + Bond/Ins + Contingency + Profit. For mid-size mining civil projects (~$50M direct), total OH + profit typically falls in the 14–22% range — values outside this band warrant review.")
NP()

# =====================================================================
# 16. MOB / DEMOB
# =====================================================================
H1("16.  MOBILIZATION & DEMOBILIZATION LOGIC")
P("Mob/demob is typically priced as a separate lump-sum line, not absorbed in unit rates. For remote/fly-in sites, add freight allowance. For multi-phase work, charge re-mob per phase.")
md = [
    ("Earthwork fleet",     "Major earthwork fleet mob/demob",         "% of equip value or flat LS",   "2%–5%"),
    ("Earthwork fleet",     "Heavy haul move per low-bed trip",         "Per round-trip",                "$1,800–$3,500"),
    ("Earthwork fleet",     "Off-road haul truck (40–100T) move",       "Per unit / round-trip",         "$8,000–$18,000"),
    ("Cranes",              "Rough terrain crane RT 60T",                "Flat per mob",                  "$4,000–$8,000"),
    ("Cranes",              "Lattice boom crawler 150T",                 "Flat per mob (incl. disassembly)","$25,000–$55,000"),
    ("Camps",               "Camp setup — fly-in remote (40-bed)",       "LS",                            "$250,000–$500,000"),
    ("Camps",               "Camp setup — drive-in (40-bed)",            "LS",                            "$120,000–$220,000"),
    ("Specialty subs",      "HDPE liner sub (incl. CQA crew)",           "LS",                            "$25,000–$60,000"),
    ("Specialty subs",      "Blasting sub (drill + powder magazine)",    "LS",                            "$35,000–$90,000"),
    ("Specialty subs",      "Aggregate crushing plant",                  "LS",                            "$150,000–$400,000"),
    ("Insurance / WCB",     "Provincial registration + WCB account",     "LS",                            "$5,000–$15,000"),
]
TBL(["Category", "Asset / Activity", "Basis", "2026 Range (CAD)"], md,
    col_widths=[3.0*cm, 5.5*cm, 4.5*cm, 3.0*cm])
NP()

# =====================================================================
# 17. CONSTRUCTABILITY
# =====================================================================
H1("17.  CONSTRUCTABILITY & EXECUTION ASSUMPTIONS  [TEMPLATE]")
P("This section captures the standing assumptions that frame each estimate. Populate per project — the headings below are mandatory; the example detail is illustrative.")
cat_blocks = [
    ("17.1  Work Calendar",       [
        "Site works 10-hr shift, 14-on / 7-off (FIFO) OR 4-on / 3-off (drive-in)",
        "Effective hours per shift after breaks &amp; travel: 8.5–9.0 productive hr",
        "Winter weather window (no work): Nov–Mar productivity factor 0.70–0.85",
        "Spring break-up moratorium (haul road restrictions): early Apr to mid-May",
    ]),
    ("17.2  Access & Logistics", [
        "Primary access road condition (paved / gravel / seasonal) and weight rating",
        "Fly-in vs drive-in — confirm chartered air freight cost &amp; frequency",
        "Material laydown sufficient for 60–90 days inventory at site",
        "Fuel resupply schedule — confirm tanker turnaround and on-site storage capacity",
    ]),
    ("17.3  Sequencing",          [
        "Earthworks completion before liner install (no traffic post-liner)",
        "Concrete cure times: 7 days form-strip; 28 days full strength; cold weather double",
        "Steel erection cannot proceed until anchor bolts surveyed and accepted",
        "Culvert installs during low-flow window per provincial fisheries authorization",
    ]),
    ("17.4  Geotechnical",        [
        "Bedrock encountered at depth X m — drill+blast assumption from elevation Y",
        "Water table at depth X m — dewatering required for excavations deeper than Y m",
        "Acid-generating waste rock requires segregation and encapsulation",
    ]),
    ("17.5  Environmental",       [
        "Fish habitat compensation work in Year 1 — site-specific permits required",
        "SWPPP / ESCP plan controlling discharge to receiving water; weekly inspection minimum",
        "Wildlife protocol (bear / caribou / grizzly) — daily site clearance",
    ]),
    ("17.6  Construction Methods",[
        "Roll-on / roll-off compaction at 95% Std Proctor; 98% for process pads",
        "Concrete supplied from on-site batch plant (if vol &gt;2,000 m³) else ready-mix",
        "Aggregate supplied from on-site quarry/crusher (if vol &gt;50,000 t) else from X km",
        "Blasting confined per pattern; pre-blast survey for structures within 500m radius",
    ]),
    ("17.7  Interface",           [
        "Owner-supplied items (e.g. process plant equipment)",
        "Existing facility tie-ins coordinated during scheduled outage windows",
    ]),
    ("17.8  Labour / IR",         [
        "Open shop / CLAC / Building Trades — confirm market and CBA obligations",
        "Indigenous community engagement commitment — % local hire / IBA obligations",
    ]),
    ("17.9  HSE",                 [
        "Site induction 4 hr per worker; site-specific safety standards apply",
        "Fire watch protocol for hot work; permit-to-work system in operation",
        "Critical lifts &gt;75% chart capacity require lift plan + engineer approval",
    ]),
    ("17.10 Productivity",        [
        "Apply 0.85–0.95 productivity factor for first 4 weeks of project (learning curve)",
        "Confined working areas: derate output by 10–30% depending on access",
    ]),
    ("17.11 Quality",             [
        "All concrete pours: nuclear density + 4-cyl set per pour",
        "All HDPE liner seams: 100% non-destructive + destructive sample 1/150 m",
    ]),
]
for h, lines in cat_blocks:
    H2(h)
    for l in lines:
        B(l)
NP()

# =====================================================================
# 18. SCOPE INCL / EXCL
# =====================================================================
H1("18.  SCOPE INCLUSIONS, EXCLUSIONS & CLARIFICATIONS  [TEMPLATE]")
P("The following register defines a typical contractor scope position. Populate INCLUDED / EXCLUDED / CLARIFY per project. This register travels with the bid and is referenced in the contract documents.")
ix = [
    ("EARTHWORKS",   "Mass excavation common soil",                "INCLUDED"),
    ("EARTHWORKS",   "Rock excavation including drill &amp; blast","INCLUDED"),
    ("EARTHWORKS",   "Removal of contaminated soil",               "EXCLUDED"),
    ("EARTHWORKS",   "Temporary site dewatering",                  "CLARIFY"),
    ("WATER MGMT",   "Permanent stormwater pond construction",     "INCLUDED"),
    ("WATER MGMT",   "Operating sediment treatment chemicals",     "EXCLUDED"),
    ("ROADS",        "Haul road sub-base + surface course",        "INCLUDED"),
    ("ROADS",        "Permanent asphalt paving",                   "CLARIFY"),
    ("WRSA/TMF",     "TMF embankment construction Zones A, B, C",  "INCLUDED"),
    ("WRSA/TMF",     "HDPE 2mm liner with CQA",                    "INCLUDED"),
    ("WRSA/TMF",     "Dam Safety Officer (DSO) services",          "EXCLUDED"),
    ("DRAINAGE",     "CSP culverts up to 1200mm diameter",         "INCLUDED"),
    ("DRAINAGE",     "Precast box / arch culverts &gt; 1500mm",    "CLARIFY"),
    ("CONCRETE",     "Cast-in-place foundations, walls, footings", "INCLUDED"),
    ("CONCRETE",     "Precast concrete (process plant)",           "EXCLUDED"),
    ("STEEL",        "Building enclosures + equipment platforms",  "INCLUDED"),
    ("STEEL",        "Process plant structural steel",             "EXCLUDED"),
    ("GEOSYNTHETICS","HDPE liner systems",                          "INCLUDED"),
    ("GEOSYNTHETICS","Geomembrane CQA Engineer of Record",         "EXCLUDED"),
    ("BLASTING",     "Drill, blast, clear of rock excavation",     "INCLUDED"),
    ("BLASTING",     "Pre-blast condition surveys",                "INCLUDED"),
    ("BURIED SVCS",  "Site potable water, sewer, process water",   "INCLUDED"),
    ("BURIED SVCS",  "Electrical / instrumentation underground",   "EXCLUDED"),
    ("FENCING",      "Site perimeter chain link + wildlife excl.", "INCLUDED"),
    ("FENCING",      "Process plant area security fencing",        "EXCLUDED"),
    ("CLOSURE",      "Final cover system + topsoil + seeding",     "INCLUDED"),
    ("CLOSURE",      "Long-term water treatment infrastructure",   "EXCLUDED"),
    ("INDIRECTS",    "Site offices, lunchrooms, dryrooms",         "INCLUDED"),
    ("INDIRECTS",    "Camp accommodation (incl. food)",            "CLARIFY"),
    ("INDIRECTS",    "Owner inspection / EPCM offices",            "EXCLUDED"),
    ("PERMITS",      "Routine construction permits",               "INCLUDED"),
    ("PERMITS",      "Environmental authorization (EA, water lic.)","EXCLUDED"),
    ("PERMITS",      "Provincial blasting permits, magazine licenses","INCLUDED"),
    ("WARRANTY",     "12-month warranty",                          "INCLUDED"),
    ("WARRANTY",     "Extended warranty beyond 12 months",         "EXCLUDED"),
]
TBL(["Discipline", "Item", "Status"], ix, col_widths=[3.0*cm, 9.5*cm, 3.0*cm])
NP()

# =====================================================================
# 19. BID LEVELING
# =====================================================================
H1("19.  BID LEVELING — TENDER COMPARISON FRAMEWORK  [TEMPLATE]")
P("Use a side-by-side worksheet (companion Excel sheet '19 Bid Leveling') to compare bidder pricing against the SME benchmark. Convert each bidder's Form of Tender into a common UOM. Flag any item more than ±20% from the benchmark.")
H2("19.1  Process")
for step in [
    "1) Common UOM normalization — strip lump sums, convert per-unit, align UOM.",
    "2) Build the comparison table: Item | Qty | Bidder A | Bidder B | Bidder C | SME Benchmark.",
    "3) Compute variance % vs benchmark for each bidder line.",
    "4) Aggregate top 10 deviation items by absolute $ — these drive total bid difference.",
    "5) Compile clarification / deviation log per bidder; request written response.",
    "6) Re-level after clarifications: this is the 'normalized bid' that goes forward to commercial evaluation.",
    "7) Sense-check: bidder total ÷ direct quantity ($/m³ aggregate) against Section 10 ranges.",
]:
    story.append(Paragraph(step, body_style))
H2("19.2  Common Deviation Categories")
P("Typical sources of bid deviation: missing inclusions, different productivity assumption, different supplier source (especially steel/rebar after CBSA), regional adjustor misapplication, unbalanced bidding (front-loaded mobilization), and exclusion of CQA, surveys, or testing.")
NP()

# =====================================================================
# 20. CHANGE ORDER
# =====================================================================
H1("20.  CHANGE ORDER PRICING METHODOLOGIES")
P("Five methodologies are recognized. Methodology used should be agreed with the Contract Administrator at project start.")
co = [
    ("Unit-rate based",    "Quantum well-defined; existing rate covers scope",     "Qty × rate × (1+DG)"),
    ("Time &amp; Materials","Scope uncertain or schedule-critical","LEM hr × $/hr + Mat × (1+wastage) + Sub × (1+markup) + OH%"),
    ("Lump-sum quotation", "Discrete, well-defined; firm quote possible",          "Direct + Indirect 10–15% + OH+P 10–20%"),
    ("Negotiated rates",   "Recurring change category",                             "Pre-agreed rate schedule × Qty"),
    ("Force account",      "Disputed scope; owner directs work pending resolution","T&amp;M as above + auditable records"),
]
TBL(["Method", "When to Use", "Build-Up"], co, col_widths=[4.0*cm, 5.5*cm, 6.0*cm])

H2("20.1  Standard Markup Stack (CCDC 2 GC 6.2.4)")
TBL(["Layer", "Element", "% Range", "Apply To"], [
    ("1", "Direct labour + equip + material",          "100%",   "Cost of work"),
    ("2", "Contractor field overhead",                  "8–12%", "Direct labour"),
    ("3", "Head office overhead",                       "5–10%", "Total direct + CFO"),
    ("4", "Contractor profit / fee",                    "8–15%", "Total direct + OH"),
    ("5", "Bond + insurance adjustment",                "1.5–3%","Total contract increase"),
    ("6", "Sub-tier markup (flow-through subs)",        "5–10%", "Sub price"),
], col_widths=[1.5*cm, 6.5*cm, 2.0*cm, 5.5*cm])
SMALL("Daily T&M sheets must be co-signed within 48 hours. Lump-sum change quotations must include itemised breakdown. Time-impact analysis is required if the change affects critical path (CCDC 2 GC 6.5).")
NP()

# =====================================================================
# 21. ROLL-UP
# =====================================================================
H1("21.  PROJECT ESTIMATE ROLL-UP TEMPLATE  [TEMPLATE]")
P("Roll-up format mirroring the 6-layer bid model. Populate the Direct field cost lines from MTO × unit rate (Section 8). Apply markups within the bands shown in Section 15. Apply the regional adjustor (Section 14) to direct lines before summing if outside MB/SK baseline. Live formulas are in the companion Excel sheet '21 Project Roll-Up'.")
roll = [
    ("1 — DIRECT FIELD COSTS",                "Σ Earthworks + WaterMgmt + Roads + WRSA + Drainage + Concrete + Steel + Geosynth + Blast + Buried + Fencing + Dewater + Closure"),
    ("2 — CONTRACTOR FIELD INDIRECTS",        "= Direct × CDI% (typ. 12-18%)"),
    ("3 — COMPANY OVERHEAD",                  "= Direct × OH% (typ. 6-12%)"),
    ("4 — BONDING &amp; INSURANCE",           "= (Direct + CDI + OH) × B+I% (typ. 2-4%)"),
    ("5 — CONTINGENCY &amp; RISK",            "= Direct × Cont.% (typ. 8-15%)"),
    ("6 — CONTRACTOR PROFIT",                 "= (Direct + CDI + OH + B+I + Cont.) × Profit% (typ. 8-15%)"),
    ("TOTAL BID",                             "= Σ of layers 1-6"),
]
TBL(["Layer", "Formula"], roll, col_widths=[6.0*cm, 9.5*cm])
NP()

# =====================================================================
# 22. RISK REGISTER
# =====================================================================
H1("22.  RISK & CONTINGENCY REGISTER  [TEMPLATE]")
P("Typical risk categories for Canadian mining heavy civil with illustrative likelihood / impact ranking. Total cost-weighted exposure drives the contingency carry in Section 15 Layer 5.")
risks = [
    ("R-01","Geotechnical","Unforeseen rock below estimated quantity","4","4","3-8% earthworks contingency"),
    ("R-02","Geotechnical","Acid-generating waste rock encountered","3","5","Owner-led env mgmt; 1-3% allowance"),
    ("R-03","Weather","Extended winter shutdown beyond plan","4","3","2-5% weather contingency"),
    ("R-04","Market","Diesel >15% above bid basis","3","3","3-8% materials escalation; fuel surcharge"),
    ("R-05","Tariff","CBSA 25% surtax extended on steel/rebar","3","3","Source Canadian/US; 5-10% on steel"),
    ("R-06","Subcontractor","Specialty sub default or delay","2","4","1-3% sub default contingency"),
    ("R-07","Permits","Environmental permit delay","3","4","Phase work; engage authority early"),
    ("R-08","Labour","Skilled trades shortage","4","3","Lock supply via early subs; 5% premium"),
    ("R-09","Indigenous","IBA / community engagement delay","3","4","Engage early"),
    ("R-10","Owner","Owner-supplied materials late","3","4","Daily T&M / standby"),
    ("R-11","Quality","HDPE liner CQA failures","2","3","Pre-qualify installer; third-party CQA"),
    ("R-12","HSE","Major incident causing stop-work","1","5","Carry incident allowance"),
]
TBL(["ID", "Cat.", "Description", "L", "I", "Mitigation / Allocation"], risks,
    col_widths=[1.2*cm, 2.2*cm, 4.5*cm, 0.9*cm, 0.9*cm, 5.8*cm])
NP()

# =====================================================================
# 23. GLOSSARY
# =====================================================================
H1("23.  GLOSSARY & ABBREVIATIONS")
glossary = [
    ("AACE",            "Association for the Advancement of Cost Engineering International"),
    ("AACE 18R-97",     "Cost Estimate Classification System — process industries"),
    ("AFE",             "Authorization for Expenditure — owner approval at sanction"),
    ("BCPI",            "Building Construction Price Index (Statistics Canada)"),
    ("CAD",             "Canadian Dollar"),
    ("CBSA",            "Canada Border Services Agency"),
    ("CDI",             "Contractor Distributable Indirects (site indirects)"),
    ("CGL",             "Commercial General Liability insurance"),
    ("CLAC",            "Christian Labour Association of Canada (open shop)"),
    ("CCDC",            "Canadian Construction Documents Committee (CCDC 2 = stipulated sum)"),
    ("CQA",             "Construction Quality Assurance (typically third-party)"),
    ("DR-11 / DR-17",   "HDPE pipe Dimension Ratio (wall thickness class)"),
    ("EHT",             "Employer Health Tax (Ontario / BC)"),
    ("EPCM",            "Engineering, Procurement and Construction Management"),
    ("ESC",             "Erosion &amp; Sediment Control"),
    ("FID",             "Final Investment Decision"),
    ("FIFO",            "Fly-in / Fly-out (14-on / 7-off rotation)"),
    ("HDPE",            "High-Density Polyethylene"),
    ("HSE",             "Health, Safety, Environment"),
    ("IBA",             "Impact Benefit Agreement (Indigenous communities)"),
    ("IFC",             "Issued For Construction (drawing status)"),
    ("LOA",             "Living Out Allowance (per-diem) for non-camp remote work"),
    ("MTO",             "Material Take-Off"),
    ("NCSO",            "National Construction Safety Officer (Canadian certification)"),
    ("NRCan",           "Natural Resources Canada"),
    ("ROM",             "Rough Order of Magnitude (Class 5/4 estimate)"),
    ("SoR",             "Schedule of Rates"),
    ("SWM",             "Stormwater Management"),
    ("SWPPP",           "Stormwater Pollution Prevention Plan"),
    ("TMF",             "Tailings Management Facility"),
    ("WBS",             "Work Breakdown Structure"),
    ("WCB / WSIB",      "Workers' Compensation Board / Ontario WSIB"),
    ("WRSA",            "Waste Rock Storage Area"),
    ("S+I",             "Supply + Install"),
    ("S+F+P",           "Supply + Form + Pour"),
]
TBL(["Term", "Definition"], glossary, col_widths=[3.5*cm, 12.0*cm])
NP()

# =====================================================================
# 24. SOURCES
# =====================================================================
H1("24.  SOURCES & REFERENCES")
refs = [
    ("STANDARDS",     "AACE Recommended Practice 18R-97 — Cost Estimate Classification System"),
    ("STANDARDS",     "ASTM D698 — Standard Proctor; ASTM D1557 — Modified Proctor"),
    ("STANDARDS",     "ASTM D4439 — Geosynthetics terminology"),
    ("STANDARDS",     "CCDC 2 — Stipulated Sum Contract"),
    ("STANDARDS",     "Canadian Dam Association Dam Safety Guidelines"),
    ("INDICES",       "Statistics Canada — Building Construction Price Index (BCPI), quarterly"),
    ("INDICES",       "NRCan — Diesel Pricing, weekly"),
    ("INDICES",       "Bank of Canada — USD/CAD exchange rate"),
    ("AUTHORITIES",   "CBSA — Notices 24-26 and 25-22 (steel / rebar surtax)"),
    ("AUTHORITIES",   "Provincial WCB / WSIB / CNESST / WorkSafeBC / WorkplaceNL"),
    ("AUTHORITIES",   "NS Energy Regulatory Board — Atlantic diesel price"),
    ("INTERNAL",      "Rev G workbook — RMM-CIVIL-CONTRACTOR-CANADA-2026-REV0 (Contractor Bid Tool)"),
    ("INTERNAL",      "Rev E workbook — RMM-CIVIL-CANADA-2026-REV0 (EPCM Benchmark Estimator)"),
    ("INTERNAL",      "Ausenco go-by Q2-2020 (Senior PD / EPCM reference)"),
    ("SUPPLIERS",     "Armtec — CSP culvert and flow guard pricing (Dec-2025)"),
    ("SUPPLIERS",     "Various HDPE pipe suppliers — Jan-2026 quotes +15% install factor"),
    ("BENCHMARKS",    "SME benchmark observations 2024-2026 — Canadian mining civil contracts"),
    ("METHODOLOGY",   "Standard industry practice — contractor markup stack and CDI percentages"),
    ("METHODOLOGY",   "CCDC 2 GC 6.2 / 6.4 / 6.5 — change order provisions"),
]
TBL(["Category", "Reference"], refs, col_widths=[3.0*cm, 12.5*cm])

SP(2)
story.append(Paragraph("— END OF DOCUMENT —", subtitle_style))
story.append(Paragraph(f"{DOC_NUMBER} | Rev 1 — Consolidated | Q2-2026", small_style))

doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
print(f"Wrote {OUT}")
print(f"Size: {os.path.getsize(OUT):,} bytes")
