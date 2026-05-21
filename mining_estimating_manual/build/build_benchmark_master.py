"""
Build Goldboro 2026 Benchmark Master — GHD-style layout, full scope.

Combines every 2026 benchmark source for every available item:
  - 835 Ausenco / Goldboro C5027 template items (5 bidder coverage)
  - RMM Rev 5 specialty mining items (sections H, I, N, O, R-Y) for full scope
  - GHD Goldboro 2026 vendor-quoted overlay
  - Alberta Transportation 2026 UPA cross-references
  - Rick Q2 2026 operator anchor (tandem, wiggle, crushing)

Output: single GHD-style scannable spreadsheet with grouped headers.
"""

import json, os, re
from statistics import median
import xlsxwriter

OUT = "/home/user/prompt-eng-interactive-tutorial/mining_estimating_manual/Goldboro_2026_Benchmark_Master_v1.xlsx"

# Load bidder data
BIDS = json.load(open("/tmp/bid_compare.json"))
BIDDERS = ["Bird", "Dexter", "GIP", "Greenfields", "Nova"]

# Load GHD Goldboro overlay (unique items)
GHD_OVERLAY = json.load(open("/tmp/goldboro_unique.json"))

# RMM Rev 5 SME items — recreate from build_excel_v5 inline data
# These are the items in Sheet 08 SoR Master (sections A-Y)
RMM_REV5_SME = [
    # (wbs_ref, description, uom, sme_low, sme_high, crew, section)
    # A — EARTHWORKS & MASS EXCAVATION (Rev E source)
    ("C-10-002", "Excavation common soil — stockpile haul 1-3km",         "m3",   24, 32,   "B3", "A. Earthworks"),
    ("C-10-003", "Backfill common — spread and compact",                  "m3",   22, 28,   "B2", "A. Earthworks"),
    ("C-10-004", "Backfill Type 1 <250mm — spread and compact",           "m3",   40, 52,   "B2", "A. Earthworks"),
    ("C-10-005", "Backfill Type 2 <100mm — spread and compact",           "m3",   55, 68,   "B2", "A. Earthworks"),
    ("C-10-006", "Backfill Type 3 engineered — spread and compact",       "m3",   72, 88,   "B2", "A. Earthworks"),
    ("C-10-007", "Rock excavation — blasted and loaded to truck",         "m3",   38, 58,   "B3", "A. Earthworks"),
    ("C-10-008", "Overburden strip — doze + rip + stockpile",             "m3",   18, 28,   "B3", "A. Earthworks"),
    ("C-10-009", "Granular A base <19mm — supply+place+compact",          "m3",  105,135,   "B2", "A. Earthworks"),
    ("C-10-010", "Granular B sub-base <75mm — supply+place+compact",      "m3",   88,115,   "B2", "A. Earthworks"),
    ("C-10-011", "Sand bedding layer — supply+place+compact",             "m3",  115,145,   "B2", "A. Earthworks"),
    ("C-10-012", "Topsoil placement — supply+spread",                     "m3",   55, 80,   "B1", "A. Earthworks"),
    ("C-10-013", "Seeding and restoration — hydraulic",                   "m2",    4,  8,   "B1", "A. Earthworks"),
    ("C-10-014", "Cut-fill balanced earthworks — large-scale",            "m3",   22, 30,   "B3", "A. Earthworks"),
    ("C-10-015", "Settling pond / water management excavation",           "m3",   22, 28,   "B3", "A. Earthworks"),
    ("C-10-016", "Drill and blast — holes only (excav separate)",         "m3",   42, 65,   "B3", "A. Earthworks"),
    # B — WATER MANAGEMENT / SWM & ESC
    ("C-20-002", "Riprap D50 150mm entrance/exit — supply+place",         "m3",   65,100,   "B3", "B. Water Mgmt"),
    ("C-20-003", "Riprap D50 300mm — supply+place",                       "m3",   75,120,   "B3", "B. Water Mgmt"),
    ("C-20-004", "Riprap D50 600mm heavy — supply+place",                 "m3",   85,140,   "B3", "B. Water Mgmt"),
    ("C-20-005", "Energy dissipator riprap — full supply+place",          "m3",  120,200,   "B3", "B. Water Mgmt"),
    ("C-20-006", "10oz non-woven geotextile cushion — S+I",               "m2",    4,  9,   "B4", "B. Water Mgmt"),
    ("C-20-007", "6oz non-woven geotextile separator — S+I",              "m2",    3,  6,   "B4", "B. Water Mgmt"),
    ("C-20-008", "Silt fence heavy duty — supply+install",                "m",    18, 28,   "B1", "B. Water Mgmt"),
    ("C-20-009", "Straw bale flow check dam — supply+install",            "m",    22, 32,   "B1", "B. Water Mgmt"),
    ("C-20-010", "Sediment dewatering bag — supply+install",              "ea", 1200,1600,  "B4", "B. Water Mgmt"),
    ("C-20-011", "Temp pump diesel submersible 8-week rental",            "ea", 7500,9500,  "N/A","B. Water Mgmt"),
    # C — ROADS / PADS / WORKING SURFACES
    ("C-30-002", "Haul road — sub-grade preparation only",                "m2",   14, 22,   "B2", "C. Roads/Pads"),
    ("C-30-003", "Granular sub-base 300mm — supply+place+compact",        "m2",   32, 45,   "B2", "C. Roads/Pads"),
    ("C-30-004", "Granular surface course 150mm — supply+place",          "m2",   22, 32,   "B2", "C. Roads/Pads"),
    ("C-30-005", "Asphalt paving 50mm HL3 surface course",                "m2",   45, 62,   "B2", "C. Roads/Pads"),
    ("C-30-006", "Equipment pad granular 600mm — heavy duty",             "m2",   50, 70,   "B2", "C. Roads/Pads"),
    ("C-30-007", "Equipment pad concrete 300mm slab on grade",            "m2",  185,265,   "C1", "C. Roads/Pads"),
    ("C-30-008", "Geotextile separation under road base — S+I",           "m2",    3,  5,   "B4", "C. Roads/Pads"),
    ("C-30-009", "Road crown and ditching — motor grader finish",         "m",     5, 10,   "B1", "C. Roads/Pads"),
    # D — WRSA / TMF EMBANKMENT
    ("C-40-002", "TMF embankment common fill Zone A/B",                   "m3",   22, 32,   "B2", "D. WRSA/TMF"),
    ("C-40-003", "TMF embankment random rockfill Zone C",                 "m3",   18, 28,   "B3", "D. WRSA/TMF"),
    ("C-40-004", "HDPE 2mm (80mil) liner — full S+I+CQA",                 "m2",   30, 48,   "B4", "D. WRSA/TMF"),
    ("C-40-005", "HDPE 1.5mm liner — full S+I+CQA",                       "m2",   22, 34,   "B4", "D. WRSA/TMF"),
    ("C-40-006", "Filter zone drainage blanket 19mm — supply+place",      "m3",  120,175,   "B4", "D. WRSA/TMF"),
    ("C-40-007", "Geonet / geocomposite drainage layer — S+I",            "m2",   16, 24,   "B4", "D. WRSA/TMF"),
    ("C-40-008", "Protective cover soil over liner — place+compact",      "m3",   30, 45,   "B2", "D. WRSA/TMF"),
    ("C-40-009", "TMF perimeter ditch — cut+shape+riprap lined",          "m",   280,420,   "B3", "D. WRSA/TMF"),
    # E — DRAINAGE / CULVERTS / PIPING
    ("C-50-002", "CSP culvert 750mm — supply+install+bedding",            "m",   580,750,   "B4", "E. Drainage/Culverts"),
    ("C-50-003", "CSP culvert 900mm — supply+install+bedding",            "m",   680,880,   "B4", "E. Drainage/Culverts"),
    ("C-50-004", "CSP culvert 1050mm — supply+install+bedding",           "m",   800,1020,  "B4", "E. Drainage/Culverts"),
    ("C-50-005", "CSP culvert 1200mm — supply+install+bedding",           "m",  1100,1400,  "B4", "E. Drainage/Culverts"),
    ("C-50-006", "CSP flow guard 600mm — supply+install",                 "ea", 2500,3200,  "B4", "E. Drainage/Culverts"),
    ("C-50-007", "CSP flow guard 750mm — supply+install",                 "ea",10500,13000, "B4", "E. Drainage/Culverts"),
    ("C-50-008", "CSP flow guard 900mm — supply+install",                 "ea",14000,17500, "B4", "E. Drainage/Culverts"),
    ("C-50-009", "CSP flow guard 1200mm — supply+install",                "ea", 8500,10500, "B4", "E. Drainage/Culverts"),
    ("C-50-010", "HDPE 100mm DR11 buried — excav+install+backfill",       "m",   310,400,   "B4", "E. Drainage/Culverts"),
    ("C-50-011", "HDPE 150mm DR11 above grade + insulation",              "m",   250,330,   "B4", "E. Drainage/Culverts"),
    ("C-50-012", "HDPE 200mm DR11 above grade + insulation",              "m",   295,380,   "B4", "E. Drainage/Culverts"),
    ("C-50-013", "HDPE 600mm DR17 above grade + insulation",              "m",   280,370,   "B4", "E. Drainage/Culverts"),
    ("C-50-014", "HDPE 1200mm inlet piping — supply+install",             "m",  1500,1900,  "B4", "E. Drainage/Culverts"),
    ("C-50-015", "Clear stone discharge pipe 450mm — S+I",                "m",   400,520,   "B4", "E. Drainage/Culverts"),
    ("C-50-016", "Discharge trench clear stone — supply+place",           "m3",  135,175,   "B4", "E. Drainage/Culverts"),
    # F — CONCRETE
    ("C-60-002", "Concrete wall and grade beam — supply+pour+cure",       "m3", 2500,3200,  "C1", "F. Concrete"),
    ("C-60-003", "Concrete pier — drill+form+pour",                       "m3", 3600,4500,  "C1", "F. Concrete"),
    ("C-60-004", "Concrete footing — continuous/isolated",                "m3", 2400,3000,  "C1", "F. Concrete"),
    ("C-60-005", "Shotcrete rock reinforcement 50mm",                     "m2",   90,140,   "C1", "F. Concrete"),
    ("C-60-006", "Rebar Grade 400 — supply+install",                      "tonne",4000,6000,"D1", "F. Concrete"),
    # G — STRUCTURAL STEEL
    ("C-70-002", "Steel beams medium 26-65 kg/m — supply+erect",          "tonne",8000,9000,"D1", "G. Steel"),
    ("C-70-003", "Steel beams heavy 66-125 kg/m — supply+erect",          "tonne",7500,8500,"D1", "G. Steel"),
    ("C-70-004", "Steel columns extra-heavy — supply+erect",              "tonne",8000,9000,"D1", "G. Steel"),
    ("C-70-005", "Steel grating 32x4.8mm — supply+install",               "m2",  660,750,   "D1", "G. Steel"),
    ("C-70-006", "Handrail incl. toe plate and fasteners",                "m",   440,500,   "D1", "G. Steel"),
    ("C-70-007", "Stairs — stringers and treads (less handrail)",         "m",   540,620,   "D1", "G. Steel"),
    ("C-70-008", "Metal roof deck 22ga 76mm — supply+erect",              "m2",  320,380,   "D1", "G. Steel"),
    ("C-70-009", "Metal cladding 22ga prefinished — supply+erect",        "m2",  265,320,   "D1", "G. Steel"),
    # H — CLOSURE & RECLAMATION (SME-only)
    ("C-80-002", "Riprap closure cap D50 300mm — supply+place",           "m3",   90,150,   "B3", "H. Closure"),
    ("C-80-003", "Cover system compacted clay 600mm",                     "m3",   58, 95,   "B2", "H. Closure"),
    ("C-80-004", "Topsoil replacement — spread+seed",                     "m2",    5, 12,   "B1", "H. Closure"),
    ("C-80-005", "HDPE cap liner 2mm — supply+install+CQA",               "m2",   30, 48,   "B4", "H. Closure"),
    # I — BLASTING (Rev G source)
    ("B-01-002", "Drill blast hole 89-115mm — rotary drill",              "m",    25, 45,   "DR1","I. Blasting"),
    ("B-01-003", "Rock blasting — drill+charge+blast+clear (all-in)",     "m3",   45, 85,   "B3", "I. Blasting"),
    ("B-01-004", "Controlled blasting near structures — presplit",        "m",    55, 95,   "DR1","I. Blasting"),
    ("B-01-005", "Secondary breakage / pop shooting",                     "ea",  250,550,   "D2", "I. Blasting"),
    # J — PERMANENT DEWATERING
    ("B-02-002", "Submersible pump station — small <50 L/s",              "ea",45000,85000, "B4", "J. Dewatering"),
    ("B-02-003", "Submersible pump station — large >50 L/s",              "ea",120000,250000,"B4","J. Dewatering"),
    ("B-02-004", "Forcemain HDPE 150mm buried — S+I",                     "m",   320,420,   "B4", "J. Dewatering"),
    ("B-02-005", "Forcemain HDPE 200mm buried — S+I",                     "m",   380,500,   "B4", "J. Dewatering"),
    ("B-02-006", "Wellpoint dewatering system — install+operate 4wk",     "LS",35000,65000, "B4", "J. Dewatering"),
    # K — BURIED SERVICES
    ("B-03-002", "Water line HDPE 150mm buried — S+I",                    "m",   380,500,   "B4", "K. Buried Services"),
    ("B-03-003", "Water line HDPE 200mm buried — S+I",                    "m",   450,620,   "B4", "K. Buried Services"),
    ("B-03-004", "Sanitary sewer PVC 150mm buried — S+I",                 "m",   380,520,   "B4", "K. Buried Services"),
    ("B-03-005", "Sanitary manhole precast 1200mm — S+I",                 "ea", 4500,7500,  "B4", "K. Buried Services"),
    ("B-03-006", "Concrete thrust block — supply+form+pour",              "ea", 1200,2500,  "C1", "K. Buried Services"),
    ("B-03-007", "Gate valve 150mm buried — S+I",                         "ea", 2200,3500,  "B4", "K. Buried Services"),
    ("B-03-008", "Hydrant — supply+install",                              "ea", 8500,14000, "B4", "K. Buried Services"),
    # L — CONCRETE PAVING & CURBING
    ("B-04-002", "Concrete paving 250mm reinforced — S+P",                "m2",  220,300,   "C1", "L. Concrete Paving"),
    ("B-04-003", "Concrete curb and gutter — S+F+P",                      "m",   185,265,   "C1", "L. Concrete Paving"),
    ("B-04-004", "Concrete wheel wash pad — supply+construct",            "ea",18000,35000, "C1", "L. Concrete Paving"),
    # M — FENCING / SECURITY
    ("B-05-002", "Chain link fence 2.4m security — S+I",                  "m",   110,165,   "B1", "M. Fencing"),
    ("B-05-003", "Wildlife exclusion fence 2.4m electrified",             "m",   165,260,   "B1", "M. Fencing"),
    ("B-05-004", "Fence gate single swing 4m — S+I",                      "ea", 1800,3200,  "B1", "M. Fencing"),
    ("B-05-005", "Fence gate double swing 8m — S+I",                      "ea", 3500,6500,  "B1", "M. Fencing"),
    # N — SITE SIGNAGE & TRAFFIC MGMT
    ("B-06-002", "Traffic control — flagging (per shift)",                "shift",850,1200, "B1", "N. Signage/Traffic"),
    ("B-06-003", "Dust suppression — calcium chloride application",       "tonne",450,650,  "B1", "N. Signage/Traffic"),
    # O — ROCK ANCHORS / STABILIZATION
    ("B-07-002", "Rock anchor active tensioned 32mm strand",              "ea", 2500,5500,  "DR1","O. Rock Anchors"),
    ("B-07-003", "Wire mesh rockfall protection — supply+pin",            "m2",   65,120,   "B3", "O. Rock Anchors"),
    ("B-07-004", "Shotcrete with fibre 75mm — supply+apply",              "m2",  120,195,   "C1", "O. Rock Anchors"),
    ("B-07-005", "Shotcrete with wire mesh 100mm — supply+apply",         "m2",  165,250,   "C1", "O. Rock Anchors"),
    # P — CULVERT STRUCTURES
    ("B-08-002", "Precast box culvert 2400x1800mm — S+I",                 "m", 6500,10500,  "B4", "P. Culvert Structures"),
    ("B-08-003", "Precast arch culvert 3000mm span — S+I",                "m", 4500,8500,   "B4", "P. Culvert Structures"),
    ("B-08-004", "Headwall concrete — supply+form+pour",                  "ea", 8500,16000, "C1", "P. Culvert Structures"),
    # Q — AGGREGATE CRUSHING (Rev 5 reset)
    ("B-09-002", "Crushing — Granular A/B (on-site portable plant)",      "tonne",18, 35,   "B3", "Q. Crushing"),
    ("B-09-003", "Crushing — riprap D50 150-600mm (on-site)",             "tonne",22, 45,   "B3", "Q. Crushing"),
    ("B-09-004", "Crushing — fines / sand bedding (on-site)",             "tonne",25, 50,   "B3", "Q. Crushing"),
    ("B-09-005", "Trucked-in Granular A from off-site quarry",            "tonne",45, 70,   "N/A","Q. Crushing"),
    ("B-09-006", "Trucked-in riprap D50 300mm — supply+haul",             "tonne",55, 95,   "N/A","Q. Crushing"),
    ("B-09-007", "Mob/demob portable crushing plant (180-350 tph)",       "LS",150000,400000,"N/A","Q. Crushing"),
    # R — UNDERGROUND CIVIL (Rev 2 SME-only)
    ("U-10-001", "Shaft sinking 6m dia — D&B incl. ground support",       "m",  35000,65000,"UC1","R. Underground"),
    ("U-10-002", "Shaft sinking 7m dia — D&B incl. ground support",       "m",  45000,85000,"UC1","R. Underground"),
    ("U-10-003", "Lateral development 5×5 m drift — D&B incl. mucking",   "m",   6500,12500,"UC2","R. Underground"),
    ("U-10-004", "Lateral development 5×6 m drift — D&B incl. mucking",   "m",   8500,16000,"UC2","R. Underground"),
    ("U-10-005", "Raise bore 3m dia (vent / ore pass)",                   "m",   4500,8500, "UC3","R. Underground"),
    ("U-10-006", "Raise bore 5m dia (intake / exhaust raise)",            "m",   9500,18000,"UC3","R. Underground"),
    ("U-10-007", "Ground support — rock bolts 2.4m resin-grouted",        "ea",    85,165,  "UC2","R. Underground"),
    ("U-10-008", "Ground support — shotcrete with mesh 75-100mm UG",      "m2",  185,320,   "UC2","R. Underground"),
    ("U-10-009", "Ground support — fibre shotcrete 75mm UG",              "m2",  145,240,   "UC2","R. Underground"),
    ("U-10-010", "Cable bolts 12m twin-strand grouted",                   "ea",  650,1250,  "UC2","R. Underground"),
    ("U-10-011", "Mucking and tramming — load and haul UG 1-2 km",        "tonne", 18, 32,  "UC2","R. Underground"),
    ("U-10-012", "Stope backfill — paste fill placement",                 "m3",   45, 95,   "UC4","R. Underground"),
    # S — TAILINGS DAM ZONED CONSTRUCTION
    ("T-20-001", "Starter dam — Zone A core (low-perm clay)",             "m3",   38, 65,   "B2", "S. Tailings Dam"),
    ("T-20-002", "Starter dam — Zone B filter fine (transition)",         "m3",   55, 95,   "B2", "S. Tailings Dam"),
    ("T-20-003", "Starter dam — Zone C filter coarse (drainage)",         "m3",   65,115,   "B2", "S. Tailings Dam"),
    ("T-20-004", "Starter dam — Zone D random rockfill (downstream)",     "m3",   22, 38,   "B3", "S. Tailings Dam"),
    ("T-20-005", "Dam raise — centerline construction per lift",          "m3",   28, 48,   "B3", "S. Tailings Dam"),
    ("T-20-006", "Dam raise — downstream construction per lift",          "m3",   25, 42,   "B3", "S. Tailings Dam"),
    ("T-20-007", "Internal drainage — chimney drain construction",        "m3",  110,180,   "B4", "S. Tailings Dam"),
    ("T-20-008", "Toe drain — supply+install with geotextile sock",       "m",   165,285,   "B4", "S. Tailings Dam"),
    ("T-20-009", "Dam instrumentation — VW piezometer installation",      "ea", 3500,6500,  "DI1","S. Tailings Dam"),
    ("T-20-010", "Dam instrumentation — inclinometer installation",       "ea", 8500,16000, "DI1","S. Tailings Dam"),
    ("T-20-011", "Dam instrumentation — survey monument installation",    "ea",  850,1650,  "DI1","S. Tailings Dam"),
    ("T-20-012", "Spillway — concrete-lined emergency overflow",          "m2",  450,750,   "C1", "S. Tailings Dam"),
    ("T-20-013", "Reclaim barge — supply+install (tailings pond)",        "ea",750000,1500000,"MEC","S. Tailings Dam"),
    # T — HEAP LEACH PAD CONSTRUCTION
    ("H-30-001", "Pad sub-grade prep — cut/fill/compact to 95% Proctor",  "m2",   18, 32,   "B2", "T. Heap Leach"),
    ("H-30-002", "Low-perm clay liner 300mm (secondary containment)",     "m2",   28, 48,   "B2", "T. Heap Leach"),
    ("H-30-003", "Primary HDPE liner 80 mil — supply+install+CQA",         "m2",   32, 55,   "B4", "T. Heap Leach"),
    ("H-30-004", "Secondary HDPE liner 60 mil — supply+install+CQA",       "m2",   24, 42,   "B4", "T. Heap Leach"),
    ("H-30-005", "Geonet drainage layer between liners — S+I",             "m2",   18, 28,   "B4", "T. Heap Leach"),
    ("H-30-006", "Drain rock cover 300mm — supply+place over primary",     "m3",   55, 95,   "B4", "T. Heap Leach"),
    ("H-30-007", "HDPE perforated collection manifold 150-200mm — S+I",    "m",   220,360,   "B4", "T. Heap Leach"),
    ("H-30-008", "Pregnant solution (PLS) pond — multi-liner",             "m2",   85,145,   "B4", "T. Heap Leach"),
    ("H-30-009", "Solution piping HDPE 200mm — S+I (header to pad)",       "m",   295,445,   "B4", "T. Heap Leach"),
    ("H-30-010", "Drip irrigation network — supply+install per ha",        "ha",75000,140000,"B4", "T. Heap Leach"),
    # U — PROCESS PLANT CIVIL
    ("P-40-001", "Mill foundation — large mat (semi-autogenous)",          "m3", 2800,3800,  "C1", "U. Process Plant Civil"),
    ("P-40-002", "Mill foundation — ball mill (smaller)",                  "m3", 2650,3500,  "C1", "U. Process Plant Civil"),
    ("P-40-003", "Primary crusher foundation — pier + mat",                "m3", 3200,4500,  "C1", "U. Process Plant Civil"),
    ("P-40-004", "SAG mill apron feeder pit excavation + concrete",        "m3", 3000,4200,  "C1", "U. Process Plant Civil"),
    ("P-40-005", "Conveyor gallery footings — incl. excavation+concrete",  "ea",12000,22000, "C1", "U. Process Plant Civil"),
    ("P-40-006", "Conveyor head pulley structure — pad+drainage",          "ea",45000,95000, "C1", "U. Process Plant Civil"),
    ("P-40-007", "Ore stockpile reclaim tunnel — cast-in-place concrete",  "m", 18000,32000, "C1", "U. Process Plant Civil"),
    ("P-40-008", "Thickener foundation — circular ring beam + center",     "ea",180000,350000,"C1","U. Process Plant Civil"),
    ("P-40-009", "Flotation cell foundation — multi-cell platform",        "m2",  285,450,   "C1", "U. Process Plant Civil"),
    ("P-40-010", "Tank farm containment — secondary bund (oil/reagent)",   "m2",  220,380,   "C1", "U. Process Plant Civil"),
    ("P-40-011", "Process building slab on grade — 250mm with rebar mat",  "m2",  195,285,   "C1", "U. Process Plant Civil"),
    # V — SLURRY / CONCENTRATE / FUEL PIPELINE
    ("SP-50-001","Pipeline ROW clearing + grading per km",                 "km",28000,65000, "B3", "V. Pipeline"),
    ("SP-50-002","Pipeline trench 1.5m depth — excavation only",          "m",    45, 85,   "B3", "V. Pipeline"),
    ("SP-50-003","HDPE slurry pipe 250mm SDR 11 — supply",                "m",   165,265,   "B4", "V. Pipeline"),
    ("SP-50-004","HDPE slurry pipe 400mm SDR 11 — supply",                "m",   385,620,   "B4", "V. Pipeline"),
    ("SP-50-005","Steel slurry pipe 300mm Sch 80 (high pressure)",        "m",   450,720,   "B4", "V. Pipeline"),
    ("SP-50-006","Pipeline install incl. fusion/weld + lay",              "m",   140,240,   "B4", "V. Pipeline"),
    ("SP-50-007","Pipeline pigging launcher/receiver station",            "ea",85000,165000,"B4", "V. Pipeline"),
    ("SP-50-008","Pipeline valve station incl. concrete pad + fence",     "ea",55000,115000,"B4", "V. Pipeline"),
    ("SP-50-009","Pipeline cathodic protection system — per km",          "km",18000,35000, "E1", "V. Pipeline"),
    ("SP-50-010","River / road crossing — horizontal directional drill",  "m",  1200,2800,  "HDD","V. Pipeline"),
    # W — MINE SITE BRIDGES
    ("BR-60-001","Modular steel bridge ≤20 m span — S+I",                  "ea",180000,380000,"ST1","W. Bridges"),
    ("BR-60-002","Cast-in-place concrete bridge 15-25 m span — S+I",       "m2 deck", 4500,7500,"C1","W. Bridges"),
    ("BR-60-003","Mine haul road bridge 30-50 m — S+I",                    "m2 deck", 6500,11000,"ST1","W. Bridges"),
    ("BR-60-004","Bridge abutment — concrete cast-in-place (each)",        "m3", 2800,4200,  "C1", "W. Bridges"),
    ("BR-60-005","Bridge pier — concrete cast-in-place (each)",            "m3", 3200,4800,  "C1", "W. Bridges"),
    ("BR-60-006","Large precast box culvert 3000×2400 mm — S+I",            "m", 11000,18500, "B4", "W. Bridges"),
    # X — POWER LINE & SUBSTATION CIVIL
    ("EL-70-001","Tower foundation — concrete pier (138 kV transmission)", "ea",18000,32000, "C1", "X. Power Line"),
    ("EL-70-002","Tower foundation — concrete pier (240 kV transmission)", "ea",26000,48000, "C1", "X. Power Line"),
    ("EL-70-003","Distribution pole foundation — drilled + concrete",      "ea", 1850,3500,  "B1", "X. Power Line"),
    ("EL-70-004","Substation pad — granular sub-base + crushed surface",   "m2",   55, 90,   "B2", "X. Power Line"),
    ("EL-70-005","Substation control building slab + foundation",          "m2",  285,450,   "C1", "X. Power Line"),
    ("EL-70-006","Substation grounding grid — copper mat install",         "m2",   18, 32,   "E1", "X. Power Line"),
    ("EL-70-007","Transmission line ROW clearing 50 m wide per km",        "km",35000,75000, "B3", "X. Power Line"),
    # Y — HYDRAULIC STRUCTURES
    ("HY-80-001","Diversion channel excavation common soil — trapezoidal", "m3",   18, 32,   "B3", "Y. Hydraulic"),
    ("HY-80-002","Diversion channel — riprap class A (D50 150mm) lining",  "m3",   85,140,   "B3", "Y. Hydraulic"),
    ("HY-80-003","Diversion channel — riprap class B (D50 300mm) lining",  "m3",  105,175,   "B3", "Y. Hydraulic"),
    ("HY-80-004","Diversion channel — riprap class C (D50 600mm) lining",  "m3",  135,220,   "B3", "Y. Hydraulic"),
    ("HY-80-005","Energy dissipator — concrete chute with baffles",        "m3", 2600,4200,  "C1", "Y. Hydraulic"),
    ("HY-80-006","Spillway weir — concrete crest with rebar",              "m3", 2850,4500,  "C1", "Y. Hydraulic"),
    ("HY-80-007","Diversion inlet structure — concrete + trash rack",      "ea",65000,135000,"C1", "Y. Hydraulic"),
    ("HY-80-008","Geotextile underlayer under riprap (filter fabric)",     "m2",    5,  9,   "B4", "Y. Hydraulic"),
    ("HY-80-009","Diversion ditch — small unlined common section",         "m",    25, 45,   "B2", "Y. Hydraulic"),
]

# AB UPA cross-references (item: avg_unit_price)
AB_UPA = {
    "Common excavation":                     8.20,
    "Borrow Excavation Contractor-Supplied": 24.40,
    "Granular Base Course":                  34.27,
    "CSP 900mm":                             720.58,
    "CSP 1000mm":                            824.24,
    "Heavy Rock Riprap Class 1":             313.41,
    "Heavy Rock Riprap Class 2":             265.82,
    "Concrete Class C":                      1951.01,
    "Concrete Class HPC":                    2692.30,
    "Asphalt Superpave":                     147.89,
    "Erosion silt fence":                    15.16,
    "Non-woven geotextile":                  2.82,
    "Concrete curb":                         170.32,
    "Pre-cast street light base":            3480.95,
}

# Rick Q2 2026 intel anchors (item: rate)
RICK_INTEL = {
    "tandem dump truck":            150,    # $/hr
    "wiggle wagon 45T":              200,   # $/hr
    "crushing on-site":              "through the roof — $18-$45/tonne baseline",
}

# Helper: GHD overlay lookup by description keyword
def find_ghd_match(desc, cmt):
    text = ' '.join(str(x or '').lower() for x in [desc, cmt])
    candidates = []
    for it in GHD_OVERLAY:
        d = (it['desc'] or '').lower()
        # Heuristic match
        score = 0
        for word in re.findall(r'[a-z]+', text):
            if len(word) > 3 and word in d:
                score += 1
        if score >= 2:
            candidates.append((score, it['capex'], it['supplier']))
    if not candidates: return None, None
    candidates.sort(reverse=True)
    return candidates[0][1], candidates[0][2]


# ===================================
# Build combined item list
# ===================================
combined = []
# Start with Goldboro template items (the 835)
seen_codes = set()
for k, b in BIDS['Bird'].items():
    rec = {
        'item': k,
        'wbs': b.get('wbs'),
        'code': b.get('code'),
        'desc': b.get('desc'),
        'cmt': b.get('cmt'),
        'qty': b.get('qty'),
        'uom': b.get('uom'),
        'source_type': 'GOLDBORO BIDS',
        'rmm_low': None,
        'rmm_high': None,
        'bids': {},
    }
    for bidder in BIDDERS:
        tot = BIDS.get(bidder, {}).get(k, {}).get('tot')
        if isinstance(tot, (int, float)) and tot > 0:
            rec['bids'][bidder] = float(tot)
    combined.append(rec)
    seen_codes.add((b.get('code'), b.get('desc')))

# Append RMM Rev 5 specialty items not in bid template
for wbs, desc, uom, lo, hi, crew, sec in RMM_REV5_SME:
    # Check if already in bids
    found = False
    for r in combined:
        if r['code'] == wbs:
            r['rmm_low'] = lo
            r['rmm_high'] = hi
            r['rmm_section'] = sec
            found = True
            break
    if not found:
        combined.append({
            'item': wbs,
            'wbs': '',
            'code': wbs,
            'desc': desc,
            'cmt': sec,
            'qty': None,
            'uom': uom,
            'source_type': 'RMM REV 5 SME',
            'rmm_low': lo,
            'rmm_high': hi,
            'rmm_section': sec,
            'bids': {},
        })

# Augment all records with GHD match + AB UPA approximate match
for rec in combined:
    rec['ghd_capex'], rec['ghd_supplier'] = find_ghd_match(rec['desc'], rec['cmt'])

# Compute bid stats
for rec in combined:
    vals = list(rec['bids'].values())
    if vals:
        rec['bid_min'] = min(vals)
        rec['bid_max'] = max(vals)
        rec['bid_median'] = median(vals)
        rec['bid_count'] = len(vals)
    else:
        rec['bid_min'] = rec['bid_max'] = rec['bid_median'] = None
        rec['bid_count'] = 0

# Recommended 2026 benchmark
def recommend(rec):
    """Source priority:
    1. If 3+ bidder data — bidder median is strongest signal
    2. Else if GHD vendor quote — vendor pricing is real Atlantic NS market
    3. Else if AB UPA available — public-tender awarded baseline
    4. Else if RMM SME — judgment range mid-point
    5. Else None
    """
    if rec['bid_count'] >= 3:
        return rec['bid_median'], 'Bidder median (Goldboro)', 'HIGH'
    if rec['ghd_capex']:
        return rec['ghd_capex'], 'GHD Goldboro vendor quote', 'MED-HIGH'
    if rec['rmm_low'] is not None and rec['rmm_high'] is not None:
        return (rec['rmm_low'] + rec['rmm_high'])/2, 'RMM Rev 5 SME mid', 'LOW-MED'
    if rec['bid_count'] >= 1:
        return rec['bid_median'], f'Bidder median (n={rec["bid_count"]})', 'LOW'
    return None, '—', '—'

for rec in combined:
    rec['rec_value'], rec['rec_basis'], rec['rec_conf'] = recommend(rec)

# Source count tally
def src_count(rec):
    n = 0
    if rec['rmm_low'] is not None: n += 1
    if rec['ghd_capex']: n += 1
    if rec['bid_count'] >= 1: n += 1
    return n

for rec in combined:
    rec['src_count'] = src_count(rec)


# ===================================
# Build the spreadsheet
# ===================================
wb = xlsxwriter.Workbook(OUT)

NAVY = "#0F2F4D"; ACCENT = "#C9A227"
LIGHT_GREY = "#F4F6F8"; SUB_GREY = "#D9D9D9"
GREEN = "#4F7942"; RED = "#A63232"

# Formats — GHD-style grouped headers
f_title = wb.add_format({"bold": True, "font_size": 14, "font_color": "white",
                          "bg_color": NAVY, "align": "left", "valign": "vcenter"})
f_sub = wb.add_format({"italic": True, "font_size": 9, "font_color": NAVY,
                        "bg_color": LIGHT_GREY, "valign": "vcenter"})
# Group headers — distinct colours for each source group (GHD-style)
f_g_id   = wb.add_format({"bold": True, "font_size": 10, "font_color": "white",
                           "bg_color": "#1F4E79", "align": "center", "border": 1, "text_wrap": True})
f_g_qty  = wb.add_format({"bold": True, "font_size": 10, "font_color": "white",
                           "bg_color": "#5F8DBF", "align": "center", "border": 1, "text_wrap": True})
f_g_rmm  = wb.add_format({"bold": True, "font_size": 10, "font_color": "white",
                           "bg_color": "#8E44AD", "align": "center", "border": 1, "text_wrap": True})
f_g_ghd  = wb.add_format({"bold": True, "font_size": 10, "font_color": "white",
                           "bg_color": "#2E8B57", "align": "center", "border": 1, "text_wrap": True})
f_g_ab   = wb.add_format({"bold": True, "font_size": 10, "font_color": "white",
                           "bg_color": "#8B4513", "align": "center", "border": 1, "text_wrap": True})
f_g_bid  = wb.add_format({"bold": True, "font_size": 10, "font_color": "white",
                           "bg_color": "#D35400", "align": "center", "border": 1, "text_wrap": True})
f_g_rick = wb.add_format({"bold": True, "font_size": 10, "font_color": "white",
                           "bg_color": "#C9A227", "align": "center", "border": 1, "text_wrap": True})
f_g_rec  = wb.add_format({"bold": True, "font_size": 10, "font_color": "white",
                           "bg_color": "#0F2F4D", "align": "center", "border": 1, "text_wrap": True})
f_g_note = wb.add_format({"bold": True, "font_size": 10, "font_color": "white",
                           "bg_color": "#555555", "align": "center", "border": 1, "text_wrap": True})

f_subhdr = wb.add_format({"bold": True, "font_size": 9, "font_color": "#0F2F4D",
                           "bg_color": "#E8F1FB", "align": "center", "border": 1, "text_wrap": True})

f_txt = wb.add_format({"font_size": 9, "align": "left", "valign": "top",
                        "border": 1, "border_color": SUB_GREY, "text_wrap": True})
f_txt_c = wb.add_format({"font_size": 9, "align": "center", "valign": "top",
                          "border": 1, "border_color": SUB_GREY, "text_wrap": True})
f_qty = wb.add_format({"font_size": 9, "align": "right", "valign": "top",
                        "border": 1, "border_color": SUB_GREY, "num_format": "#,##0.0"})
f_money = wb.add_format({"font_size": 9, "align": "right", "valign": "top",
                          "border": 1, "border_color": SUB_GREY, "num_format": '"$"#,##0.00'})
f_money_rmm = wb.add_format({"font_size": 9, "align": "right", "valign": "top",
                              "border": 1, "border_color": SUB_GREY, "num_format": '"$"#,##0.00',
                              "bg_color": "#F2E6F7"})
f_money_ghd = wb.add_format({"font_size": 9, "align": "right", "valign": "top",
                              "border": 1, "border_color": SUB_GREY, "num_format": '"$"#,##0.00',
                              "bg_color": "#E8F5E9"})
f_money_ab = wb.add_format({"font_size": 9, "align": "right", "valign": "top",
                             "border": 1, "border_color": SUB_GREY, "num_format": '"$"#,##0.00',
                             "bg_color": "#FCEFE6"})
f_money_bid = wb.add_format({"font_size": 9, "align": "right", "valign": "top",
                              "border": 1, "border_color": SUB_GREY, "num_format": '"$"#,##0.00',
                              "bg_color": "#FEF0E1"})
f_money_rec = wb.add_format({"bold": True, "font_size": 10, "align": "right", "valign": "top",
                              "border": 1, "border_color": SUB_GREY, "num_format": '"$"#,##0.00',
                              "bg_color": "#FFF9C4", "font_color": "#0F2F4D"})

f_conf_h = wb.add_format({"bold": True, "font_size": 9, "font_color": "white",
                           "bg_color": GREEN, "align": "center", "border": 1})
f_conf_m = wb.add_format({"bold": True, "font_size": 9, "font_color": "white",
                           "bg_color": ACCENT, "align": "center", "border": 1})
f_conf_l = wb.add_format({"bold": True, "font_size": 9, "font_color": "white",
                           "bg_color": RED, "align": "center", "border": 1})

# Sheet 1: COVER
ws = wb.add_worksheet("00 Cover")
ws.hide_gridlines(2)
ws.set_column("A:A", 26)
ws.set_column("B:B", 100)
ws.set_row(0, 28)
ws.merge_range(0, 0, 0, 1, "GOLDBORO 2026 BENCHMARK MASTER — Full Scope, GHD-Style Layout, Multi-Source", f_title)
ws.set_row(1, 18)
ws.merge_range(1, 0, 1, 1, "Every available 2026 market data source per item, side-by-side. Recommended benchmark synthesized with source-count and confidence rating.", f_sub)

cover = [
    ("PURPOSE", "Single consolidated 2026 mining heavy civil benchmark master. Combines actual contractor bids, vendor-quoted public-tender data, EPCM reference rates, and SME judgement ranges into one scannable comparison table per item."),
    ("LAYOUT", "GHD-style grouped headers: IDENTIFICATION / QUANTITY / RMM Rev 5 SME / GHD Goldboro 2026 / AB UPA 2026 / BIDDERS (5) / Rick Intel / RECOMMENDED 2026 / Source Count / Confidence / Vetting Note."),
    ("SCOPE", f"Full scope. {len(combined)} line items: ~835 from Goldboro C5027 template (with 5 bidder coverage), plus ~155 RMM Rev 5 SME-only items (specialty mining sections: closure, blasting, signage, rock anchors, underground civil, tailings dam zoned construction, heap leach pads, process plant civil, slurry pipelines, mine bridges, power line civil, hydraulic structures)."),
    ("SOURCES",
     "1. RMM Rev 5 SME — author's SME ranges (purple).  "
     "2. GHD Goldboro 2026 — vendor-quoted real Atlantic NS pricing from GHD project file Apr-14-2026 (green).  "
     "3. AB UPA 2026 — Alberta Transportation weighted-avg public tender data (brown).  "
     "4. Goldboro C5027 bidders — 5 actual contractor April 2026 submissions: Bird, Dexter, GIP, Greenfields, Nova (orange).  "
     "5. Rick Q2 2026 intel — operator anchor for tandem ($150/hr), wiggle ($200/hr), crushing (through the roof) (gold)."),
    ("CONFIDENCE RATING",
     "HIGH (green) = 3+ bidder data points (real market submissions).  "
     "MED-HIGH (amber) = GHD vendor quote available.  "
     "LOW-MED (red) = SME-only or 1-2 sources only.  "
     "Sort the Confidence column descending to see items most defensible first."),
    ("RECOMMENDED 2026 BENCHMARK LOGIC",
     "Priority cascade: (1) Bidder median if 3+ bidders priced it → HIGH confidence.  "
     "(2) Else GHD Goldboro vendor capex → MED-HIGH confidence.  "
     "(3) Else RMM Rev 5 SME mid-point → LOW-MED confidence.  "
     "(4) Else bidder median (only 1-2 bidders) → LOW confidence.  "
     "Where multiple sources disagree by >50%, the spread itself is the signal — investigate scope before committing."),
    ("HOW TO USE",
     "Open '01 Benchmark Master'. Each row = one item. Each colour band = one source group. The Recommended 2026 column (yellow) is the synthesized benchmark. Apply Excel filter on Section, Confidence, or Source Count to focus your review. Sort descending on 'Bidder Spread %' to find scope-misalignment hotspots."),
    ("CELL COLOUR KEY",
     "Purple = RMM Rev 5 SME.  Green = GHD Goldboro vendor.  Brown = AB UPA.  Orange = Bidders.  Gold = Rick intel.  Yellow = Recommended.  Green/Amber/Red confidence tags."),
    ("DATA INTEGRITY",
     "Goldboro item descriptions, codes, and quantities are Ausenco-supplied (col 7 of source MTO). Bidder unit prices are 'TOTAL COST per UOM' (col 19 of source). RMM ranges are from RMM-CIVIL-CANADA-2026-MANUAL-REV5 Sheet 08. GHD overlay matched by description keyword (heuristic; manually verify high-value items)."),
    ("OUT OF SCOPE",
     "This file is a benchmark reference. Does not include: bid leveling and clarification log (see Goldboro_C5027_Bid_Comparison_v1.xlsx); contract pricing buildup with quantities (use this file's recommended values × project MTO); risk allocation or markup stack (see RMM Rev 5 Sheet 15)."),
]
r = 3
for k, v in cover:
    ws.set_row(r, 38)
    ws.write(r, 0, k, f_txt)
    ws.write(r, 1, v, f_txt)
    r += 1


# Sheet 2: BENCHMARK MASTER
ws = wb.add_worksheet("01 Benchmark Master")
ws.hide_gridlines(2)
ws.freeze_panes(4, 5)
# Column widths
col_widths = [
    7,     # WBS L4
    11,    # Code
    40,    # Desc
    18,    # Section / Comments
    9,     # Qty
    7,     # UOM
    10, 10, # RMM Low, RMM High
    10, 18, # GHD capex, GHD supplier
    10,    # AB UPA
    10, 10, 10, 10, 10, # Bird, Dexter, GIP, Greenfields, Nova
    10,    # Rick
    10, 10, 10, # Bidder Min, Med, Max
    8,     # Spread %
    11,    # Recommended
    18,    # Rec basis
    8,     # Src count
    10,    # Confidence
    30,    # Note
]
for i, w in enumerate(col_widths):
    ws.set_column(i, i, w)

# Row 0: title
ws.set_row(0, 26)
ws.merge_range(0, 0, 0, len(col_widths)-1,
    "GOLDBORO 2026 BENCHMARK MASTER — Full Scope (multi-source comparison)",
    f_title)
ws.set_row(1, 16)
ws.merge_range(1, 0, 1, len(col_widths)-1,
    "Row = item. Column groups = sources. Yellow = recommended 2026 benchmark. Green/amber/red = confidence.",
    f_sub)

# Row 2: GROUPED HEADERS
ws.set_row(2, 28)
ws.merge_range(2, 0, 2, 3, "IDENTIFICATION", f_g_id)
ws.merge_range(2, 4, 2, 5, "QUANTITY", f_g_qty)
ws.merge_range(2, 6, 2, 7, "RMM REV 5 SME", f_g_rmm)
ws.merge_range(2, 8, 2, 9, "GHD GOLDBORO 2026", f_g_ghd)
ws.write(2, 10, "AB UPA 2026", f_g_ab)
ws.merge_range(2, 11, 2, 15, "GOLDBORO C5027 BIDDERS (Apr 2026)", f_g_bid)
ws.write(2, 16, "Rick Q2-2026", f_g_rick)
ws.merge_range(2, 17, 2, 20, "BIDDER STATS", f_g_bid)
ws.merge_range(2, 21, 2, 25, "RECOMMENDED 2026 BENCHMARK", f_g_rec)
ws.write(2, 26, "VETTING NOTE", f_g_note)

# Row 3: sub-headers
sub_hdr = [
    "WBS L4", "Commodity Code", "Description / Scope", "Section",
    "Ausenco Qty", "UOM",
    "Low", "High",
    "Capex Unit", "Supplier / Basis",
    "Prov Avg",
    "Bird", "Dexter", "GIP", "Greenfields", "Nova",
    "Operator Anchor",
    "Bid Min", "Bid Median", "Bid Max", "Spread %",
    "2026 Value", "Basis", "# Sources", "Confidence", "—",
    "Vetting / Scope Note"
]
ws.set_row(3, 32)
for c, h in enumerate(sub_hdr):
    ws.write(3, c, h, f_subhdr)

# Sort items: by section (cat) then by description
# Determine cat for each row
def cat_of(rec):
    if 'rmm_section' in rec and rec['rmm_section']:
        return rec['rmm_section']
    # else categorize from desc
    text = ' '.join(str(x or '').lower() for x in [rec.get('desc'), rec.get('cmt')])
    if any(k in text for k in ['contact water', 'swm ditch', 'silt', 'runoff', 'sediment fence']): return "B. Water Mgmt"
    if any(k in text for k in ['excavation', 'clearing', 'grubbing', 'strip', 'organic']): return "A. Earthworks"
    if any(k in text for k in ['backfill', 'fill', 'compact', 'embank']): return "A. Earthworks"
    if any(k in text for k in ['granular', 'base course', 'subgrade', 'aggregate']): return "C. Roads/Pads"
    if any(k in text for k in ['culvert', 'csp', 'flow guard', 'coupler']): return "E. Drainage/Culverts"
    if any(k in text for k in ['riprap', 'rip-rap', 'rip rap', 'ballast', 'armour']): return "F. Riprap"
    if any(k in text for k in ['geotextile', 'geomembrane', 'liner', 'fabric']): return "G. Geosynthetics"
    if any(k in text for k in ['concrete', 'rebar', 'shotcrete']): return "H. Concrete"
    if any(k in text for k in ['hdpe pipe', 'piping', 'pipe', 'sdr', 'dr11']): return "I. Pipe"
    if any(k in text for k in ['fence', 'gate']): return "J. Fencing"
    if any(k in text for k in ['settling pond', 'polishing pond', 'pump station']): return "K. Ponds"
    return "Z. Other / Indirects"

for rec in combined:
    rec['cat'] = cat_of(rec)

# Sort
combined.sort(key=lambda r: (r['cat'], str(r.get('desc') or ''), str(r.get('code') or '')))

# Write data rows
row = 4
current_cat = None
for rec in combined:
    if rec['cat'] != current_cat:
        ws.set_row(row, 20)
        ws.merge_range(row, 0, row, len(col_widths)-1, rec['cat'],
                        wb.add_format({"bold": True, "font_size": 10, "font_color": NAVY,
                                        "bg_color": SUB_GREY, "valign": "vcenter", "border": 1}))
        row += 1
        current_cat = rec['cat']
    ws.set_row(row, 28)
    # IDENTIFICATION
    ws.write(row, 0, str(rec.get('wbs') or '')[:8], f_txt_c)
    ws.write(row, 1, str(rec.get('code') or '')[:12], f_txt_c)
    desc = str(rec.get('desc') or '')[:200]
    cmt = str(rec.get('cmt') or '')
    full = desc + ((' — ' + cmt) if cmt and cmt != desc and cmt != rec['cat'] else '')
    ws.write(row, 2, full[:150], f_txt)
    sec_label = rec.get('rmm_section') or rec.get('cat')
    ws.write(row, 3, sec_label[:18], f_txt_c)
    # QUANTITY
    qty = rec.get('qty')
    if isinstance(qty, (int, float)):
        ws.write_number(row, 4, qty, f_qty)
    else:
        ws.write(row, 4, "—", f_txt_c)
    ws.write(row, 5, str(rec.get('uom') or '—')[:6], f_txt_c)
    # RMM Rev 5 SME
    for ci, val in [(6, rec.get('rmm_low')), (7, rec.get('rmm_high'))]:
        if isinstance(val, (int, float)):
            ws.write_number(row, ci, val, f_money_rmm)
        else:
            ws.write(row, ci, "—", f_txt_c)
    # GHD Goldboro 2026
    if isinstance(rec.get('ghd_capex'), (int, float)):
        ws.write_number(row, 8, rec['ghd_capex'], f_money_ghd)
        ws.write(row, 9, str(rec.get('ghd_supplier') or '')[:30], f_txt)
    else:
        ws.write(row, 8, "—", f_txt_c)
        ws.write(row, 9, "—", f_txt_c)
    # AB UPA 2026 (heuristic — look for matching keywords)
    ab_val = None
    for k_ab, v_ab in AB_UPA.items():
        if k_ab.lower() in (rec.get('desc') or '').lower():
            ab_val = v_ab; break
    if ab_val:
        ws.write_number(row, 10, ab_val, f_money_ab)
    else:
        ws.write(row, 10, "—", f_txt_c)
    # BIDDERS (Bird, Dexter, GIP, Greenfields, Nova)
    for i, bidder in enumerate(BIDDERS):
        col = 11 + i
        val = rec['bids'].get(bidder)
        if isinstance(val, (int, float)):
            ws.write_number(row, col, val, f_money_bid)
        else:
            ws.write(row, col, "—", f_txt_c)
    # Rick intel
    rick_val = None
    desc_l = (rec.get('desc') or '').lower()
    if 'tandem' in desc_l: rick_val = 150
    elif 'wiggle' in desc_l or 'super b' in desc_l: rick_val = 200
    if isinstance(rick_val, (int, float)):
        ws.write_number(row, 16, rick_val, f_money)
    else:
        ws.write(row, 16, "—", f_txt_c)
    # BIDDER STATS
    if rec['bid_count'] >= 2:
        ws.write_number(row, 17, rec['bid_min'], f_money_bid)
        ws.write_number(row, 18, rec['bid_median'], f_money_bid)
        ws.write_number(row, 19, rec['bid_max'], f_money_bid)
        spread = (rec['bid_max'] - rec['bid_min']) / rec['bid_median'] if rec['bid_median'] > 0 else 0
        ws.write_number(row, 20, spread, wb.add_format({"font_size": 9, "align": "right", "border": 1, "num_format": "0%"}))
    else:
        for c in range(17, 21):
            ws.write(row, c, "—", f_txt_c)
    # RECOMMENDED
    if rec.get('rec_value'):
        ws.write_number(row, 21, rec['rec_value'], f_money_rec)
    else:
        ws.write(row, 21, "—", f_txt_c)
    ws.write(row, 22, rec.get('rec_basis') or '—', f_txt)
    ws.write(row, 23, rec.get('src_count') or 0, f_txt_c)
    # Confidence colour
    conf = rec.get('rec_conf') or '—'
    if 'HIGH' in conf and 'MED-HIGH' not in conf:
        ws.write(row, 24, conf, f_conf_h)
    elif 'MED' in conf:
        ws.write(row, 24, conf, f_conf_m)
    elif 'LOW' in conf:
        ws.write(row, 24, conf, f_conf_l)
    else:
        ws.write(row, 24, conf, f_txt_c)
    ws.write(row, 25, "—", f_txt_c)
    # Vetting note
    notes = []
    if rec.get('bid_count') == 0 and rec.get('rmm_low') is None:
        notes.append("⚠ NO DATA — SME judgment required")
    if rec.get('bid_count', 0) >= 4 and rec.get('bid_median'):
        spread_pct = (rec['bid_max'] - rec['bid_min']) / rec['bid_median'] * 100
        if spread_pct > 100:
            notes.append(f"⚠ Bidder spread {spread_pct:.0f}% — scope misalignment")
    if rec.get('rmm_low') and rec.get('bid_median') and rec.get('rmm_high'):
        bm = rec['bid_median']
        if bm < rec['rmm_low'] * 0.5:
            notes.append("Bidder median <50% of SME floor — bidders aggressive or missing scope")
        elif bm > rec['rmm_high'] * 2:
            notes.append("Bidder median >2x SME ceiling — premium or scope add")
    if not notes:
        notes.append("OK — sources align")
    ws.write(row, 26, '; '.join(notes), f_txt)
    row += 1

print(f"Wrote {row} rows in benchmark master")

# Sheet 3: SUMMARY METRICS
ws = wb.add_worksheet("02 Coverage Summary")
ws.hide_gridlines(2)
ws.set_column("A:A", 36)
ws.set_column("B:E", 14)
ws.set_row(0, 26)
ws.merge_range(0, 0, 0, 4, "BENCHMARK COVERAGE SUMMARY", f_title)

from collections import Counter
cat_stats = {}
for rec in combined:
    c = rec.get('cat') or 'Z'
    if c not in cat_stats:
        cat_stats[c] = {'total': 0, 'has_bidder': 0, 'has_ghd': 0, 'has_rmm': 0, 'no_data': 0}
    cat_stats[c]['total'] += 1
    if rec.get('bid_count', 0) >= 1: cat_stats[c]['has_bidder'] += 1
    if rec.get('ghd_capex'): cat_stats[c]['has_ghd'] += 1
    if rec.get('rmm_low') is not None: cat_stats[c]['has_rmm'] += 1
    if rec.get('bid_count', 0) == 0 and not rec.get('ghd_capex') and rec.get('rmm_low') is None:
        cat_stats[c]['no_data'] += 1

ws.set_row(2, 32)
ws.write(2, 0, "Section", f_g_id)
ws.write(2, 1, "Total Items", f_g_id)
ws.write(2, 2, "With Bidder Data", f_g_bid)
ws.write(2, 3, "With GHD Data", f_g_ghd)
ws.write(2, 4, "With RMM SME", f_g_rmm)

r = 3
for cat in sorted(cat_stats.keys()):
    s = cat_stats[cat]
    ws.set_row(r, 22)
    ws.write(r, 0, cat, f_txt)
    ws.write(r, 1, s['total'], f_txt_c)
    ws.write(r, 2, s['has_bidder'], f_txt_c)
    ws.write(r, 3, s['has_ghd'], f_txt_c)
    ws.write(r, 4, s['has_rmm'], f_txt_c)
    r += 1

# Totals
ws.set_row(r, 22)
ws.write(r, 0, "TOTAL", f_g_id)
ws.write(r, 1, sum(s['total'] for s in cat_stats.values()), f_g_id)
ws.write(r, 2, sum(s['has_bidder'] for s in cat_stats.values()), f_g_id)
ws.write(r, 3, sum(s['has_ghd'] for s in cat_stats.values()), f_g_id)
ws.write(r, 4, sum(s['has_rmm'] for s in cat_stats.values()), f_g_id)

wb.close()
print(f"Wrote {OUT}")
print(f"Size: {os.path.getsize(OUT):,} bytes")
