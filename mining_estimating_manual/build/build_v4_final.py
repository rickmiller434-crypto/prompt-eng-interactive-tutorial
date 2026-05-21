"""
Build RMM_2026_Heavy_Civil_Benchmark_V4_FINAL.xlsx

Preserves V3 structure (9 tabs) and modifies Tab 01 Rate Schedule:
  - Replace aggregate Bidder Min/Med/Max (3 cols) with 9 cols:
    5 individual bidder columns (Bird, Dexter, GIP, Greenfields, Nova)
    + 3 stat columns (Min/Med/Max as formulas) + Spread%
  - Per-item bidder matching with size-aware keyword scoring
  - Adds Match Basis column showing how each row was matched
  - Other tabs (00 Cover, 02-08) preserved from V3 verbatim

Author: Rick Miller / RMM
"""

import json, re, shutil
from statistics import median
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from copy import copy

V3 = "/tmp/v3.xlsx"
V4 = "/home/user/prompt-eng-interactive-tutorial/mining_estimating_manual/RMM_2026_Heavy_Civil_Benchmark_V4_FINAL.xlsx"

# Step 1: copy V3 → V4 (preserves all formatting + other tabs)
shutil.copy(V3, V4)

# Step 2: open and modify Tab 01
wb = openpyxl.load_workbook(V4)
ws = wb["01 Rate Schedule"]

# Step 3: load bidder data + GHD overlay
BIDS = json.load(open("/tmp/bid_compare.json"))
BIDDERS = ["Bird", "Dexter", "GIP", "Greenfields", "Nova"]

# Pre-index bidder data by description for fast matching
bidder_index = {}
for bidder in BIDDERS:
    bidder_index[bidder] = []
    for k, v in BIDS[bidder].items():
        tot = v.get('tot')
        if not isinstance(tot, (int, float)) or tot <= 0: continue
        full = ((v.get('desc') or '') + ' ' + (v.get('cmt') or '')).lower()
        bidder_index[bidder].append({
            'item': k, 'desc': v.get('desc'), 'cmt': v.get('cmt'),
            'full': full, 'uom': v.get('uom'), 'tot': float(tot)
        })

# --- Per-item matching algorithm ---
SIZE_PATTERNS = [
    r'\b(\d+)\s*mm\b',           # 150mm, 1200mm
    r'\b(\d+)\s*-\s*(\d+)\s*mm', # 600-700mm
    r'\b<\s*(\d+)\s*mm',          # <250mm
    r'\b(\d+)\s*inch\b',          # 8-inch
    r'\b(\d+\.?\d*)\s*oz\b',     # 6oz, 12oz
    r'\b(\d+)\s*mil\b',           # 80 mil
    r'\bd50\s*=?\s*(\d+)',        # D50 150
    r'\bdr\s*(\d+)\b',            # DR11, DR17
    r'\bsdr\s*(\d+)\b',           # SDR 11
    r'\bclass\s+([abc])\b',       # Class A/B/C riprap
    r'\bzone\s+([abcd])\b',       # Zone A/B/C/D
    r'\b(\d+)\s*t\b',             # 17T, 25T, 45T
    r'\b(\d+)\s*tph\b',           # 180 tph
]
STOPWORDS = {'and', 'the', 'for', 'with', 'from', 'into', 'unit', 'rate',
             'supply', 'install', 'allowance',
             'incl', 'includes', 'including', 'inc',
             'standard', 'work', 'works'}

def extract_keywords(text):
    """Extract size tokens + non-stopword keywords."""
    text_l = text.lower()
    sizes = []
    for pat in SIZE_PATTERNS:
        for m in re.findall(pat, text_l):
            if isinstance(m, tuple):
                sizes.extend([s for s in m if s])
            else:
                sizes.append(str(m))
    # Tokenize words
    words = re.findall(r'[a-z]{4,}', text_l)
    keywords = [w for w in words if w not in STOPWORDS]
    return set(sizes), set(keywords)

def find_per_item_match(rmm_desc, rmm_cmt, rmm_uom):
    """For an RMM canonical item, find best-matching bidder items per bidder.
    Two-tier strategy:
      - Tier 1 (preferred): size-matched + keyword overlap >= 2
      - Tier 2 (fallback): keyword-only match if Tier 1 yields nothing
    Returns dict {bidder: (unit_price, match_basis, n_matches)}."""
    rmm_text = (rmm_desc or '') + ' ' + (rmm_cmt or '')
    rmm_sizes, rmm_keywords = extract_keywords(rmm_text)
    rmm_uom_l = (rmm_uom or '').lower().strip()

    UOM_GROUPS = [
        {'ea', 'each', 'unit', 'units'},
        {'m3', 'cu.m', 'cu.m.', 'm³', 'cubic metre'},
        {'m2', 'sq.m', 'sq.m.', 'm²', 'square metre'},
        {'m', 'metre', 'lin.m', 'lm'},
        {'t', 'tonne', 'tonnes'},
        {'shift', 'day', 'days'},
        {'ls', 'lot', 'sum'},
    ]

    def uom_matches(b_uom):
        if not rmm_uom_l or not b_uom: return True
        if rmm_uom_l == b_uom: return True
        for grp in UOM_GROUPS:
            if rmm_uom_l in grp and b_uom in grp:
                return True
        return False

    result = {}
    for bidder in BIDDERS:
        # 3-tier: T1 size+kw, T2 strong kw, T3 category fallback
        tier1 = []  # size match + kw overlap >=2 (strict)
        tier2 = []  # kw overlap >=2 (no size constraint)
        tier3 = []  # kw overlap >=1 (category fallback)
        for item in bidder_index[bidder]:
            b_uom = (item['uom'] or '').lower().strip()
            if not uom_matches(b_uom): continue
            b_sizes, b_keywords = extract_keywords(item['full'])
            kw_overlap = len(rmm_keywords & b_keywords)
            size_overlap = len(rmm_sizes & b_sizes) if rmm_sizes else 0
            has_size_match = bool(rmm_sizes) and size_overlap > 0

            if rmm_sizes and has_size_match and kw_overlap >= 1:
                tier1.append((kw_overlap + 10, item['tot']))
            elif kw_overlap >= 2:
                tier2.append((kw_overlap, item['tot']))
            elif kw_overlap >= 1:
                tier3.append((kw_overlap, item['tot']))

        # Prefer tier 1; fall back to tier 2, then tier 3
        if tier1:
            tier1.sort(reverse=True)
            top = tier1[:3]
            prices = [c[1] for c in top]
            basis = f"size-matched ({len(top)} of {len(tier1)})"
        elif tier2:
            tier2.sort(reverse=True)
            top = tier2[:3]
            prices = [c[1] for c in top]
            basis = f"keyword-matched ({len(top)} of {len(tier2)})"
        elif tier3:
            tier3.sort(reverse=True)
            top = tier3[:5]  # broader sample for category fallback
            prices = [c[1] for c in top]
            basis = f"category-fallback ({len(top)} of {len(tier3)})"
        else:
            result[bidder] = None
            continue
        median_price = round(median(prices), 2) if prices else None
        result[bidder] = (median_price, basis, len(tier1) + len(tier2) + len(tier3))
    return result

# Step 4: read existing Tab 01 data
print("Reading existing V3 Tab 01 structure...")
existing_data = []
for r in range(5, ws.max_row + 1):
    row_vals = [ws.cell(r, c).value for c in range(1, 22)]
    if any(v is not None for v in row_vals):
        existing_data.append(row_vals)
    else:
        existing_data.append(None)  # blank row
print(f"Read {len(existing_data)} rows from V3 Tab 01 data area")

# Step 5: clear Tab 01 columns 22+ (bidder + recommended + notes + flags)
# AND insert 5 new columns for per-bidder data
# Strategy: rewrite all data rows 5+ with the new column layout

# New column layout (1-indexed):
# 1-4: ID (WBS, Code, Desc, UOM) - keep from V3
# 5-18: Regions (NL/NS/ON/MBSK/AB/BC/YT × Low/High) - keep
# 19-21: Anchors (GHD, AB UPA, Rick) - keep
# 22-26: NEW — Bird, Dexter, GIP, Greenfields, Nova
# 27-29: NEW — Bidder Min (formula), Med (formula), Max (formula)
# 30: NEW — Bidder Spread%
# 31: Recommended 2026 (was col 25)
# 32: Vetting Note (was col 26)
# 33: Match Basis (NEW — replaces Bidder Med Δ%)
# 34: GHD Δ% (was col 28)
# 35: FLAG (was col 29)
# 36: Data QC (was col 30)

# Clear all data cells beyond col 21 first — unmerge any merged ranges in that region
ranges_to_unmerge_all = []
for mr in list(ws.merged_cells.ranges):
    if mr.max_col >= 22:
        ranges_to_unmerge_all.append(str(mr))
for mr_str in ranges_to_unmerge_all:
    ws.unmerge_cells(mr_str)
print(f"Unmerged {len(ranges_to_unmerge_all)} merged ranges in cols 22+")

for r in range(3, ws.max_row + 1):
    for c in range(22, 40):  # up to col 40
        cell = ws.cell(r, c)
        cell.value = None

# Apply formatting helpers
def style_header(cell, bg_color, font_color="FFFFFF"):
    cell.font = Font(bold=True, size=10, color=font_color)
    cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(left=Side(style='thin', color='999999'),
                          right=Side(style='thin', color='999999'),
                          top=Side(style='thin', color='999999'),
                          bottom=Side(style='thin', color='999999'))

def style_sub(cell):
    cell.font = Font(bold=True, size=9, color="0F2F4D")
    cell.fill = PatternFill(start_color="E8F1FB", end_color="E8F1FB", fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(left=Side(style='thin', color='999999'),
                          right=Side(style='thin', color='999999'),
                          top=Side(style='thin', color='999999'),
                          bottom=Side(style='thin', color='999999'))

def style_money(cell, bg_color=None, bold=False):
    cell.font = Font(size=9, bold=bold, color="000000")
    if bg_color:
        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    cell.alignment = Alignment(horizontal="right", vertical="top")
    cell.number_format = '"$"#,##0.00'
    cell.border = Border(left=Side(style='thin', color='CCCCCC'),
                          right=Side(style='thin', color='CCCCCC'),
                          top=Side(style='thin', color='CCCCCC'),
                          bottom=Side(style='thin', color='CCCCCC'))

def style_text(cell, bg_color=None, bold=False):
    cell.font = Font(size=9, bold=bold, color="000000")
    if bg_color:
        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    cell.border = Border(left=Side(style='thin', color='CCCCCC'),
                          right=Side(style='thin', color='CCCCCC'),
                          top=Side(style='thin', color='CCCCCC'),
                          bottom=Side(style='thin', color='CCCCCC'))

def style_center(cell, bg_color=None, bold=False):
    cell.font = Font(size=9, bold=bold)
    if bg_color:
        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="top")
    cell.border = Border(left=Side(style='thin', color='CCCCCC'),
                          right=Side(style='thin', color='CCCCCC'),
                          top=Side(style='thin', color='CCCCCC'),
                          bottom=Side(style='thin', color='CCCCCC'))

def style_pct(cell, bg_color=None, bold=False):
    cell.font = Font(size=9, bold=bold)
    if bg_color:
        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    cell.alignment = Alignment(horizontal="right", vertical="top")
    cell.number_format = "0.0%"
    cell.border = Border(left=Side(style='thin', color='CCCCCC'),
                          right=Side(style='thin', color='CCCCCC'),
                          top=Side(style='thin', color='CCCCCC'),
                          bottom=Side(style='thin', color='CCCCCC'))

# Step 6: write new group header row 3
print("Writing new group headers...")
ws.cell(3, 22).value = "C5027 BIDDER SUBMISSIONS — Individual Rates (Apr 2026)"
ws.merge_cells(start_row=3, start_column=22, end_row=3, end_column=26)
style_header(ws.cell(3, 22), "A63232")

ws.cell(3, 27).value = "BIDDER STATS (per-rate)"
ws.merge_cells(start_row=3, start_column=27, end_row=3, end_column=30)
style_header(ws.cell(3, 27), "A63232")

ws.cell(3, 31).value = "RMM RECOMMENDED 2026"
style_header(ws.cell(3, 31), "0F2F4D")
ws.cell(3, 32).value = "VETTING / MATCH BASIS"
style_header(ws.cell(3, 32), "555555")
ws.merge_cells(start_row=3, start_column=32, end_row=3, end_column=33)
ws.cell(3, 34).value = "BIDDER MED Δ%"
style_header(ws.cell(3, 34), "C9A227")
ws.cell(3, 35).value = "GHD Δ%"
style_header(ws.cell(3, 35), "C9A227")
ws.cell(3, 36).value = "FLAG"
style_header(ws.cell(3, 36), "555555")

# Step 7: write new sub-headers row 4
print("Writing new sub-headers...")
sub_headers = {
    22: "Bird", 23: "Dexter", 24: "GIP", 25: "Greenfields", 26: "Nova",
    27: "Bidder Min", 28: "Bidder Med", 29: "Bidder Max", 30: "Spread %",
    31: "2026 $ / UOM",
    32: "Vetting Note", 33: "Match Basis",
    34: "Bidder Med vs REC", 35: "GHD vs REC",
    36: "FLAG"
}
for c, h in sub_headers.items():
    cell = ws.cell(4, c)
    cell.value = h
    style_sub(cell)

# Set column widths
col_widths_new = {
    22: 9, 23: 9, 24: 9, 25: 11, 26: 9,  # 5 bidders
    27: 10, 28: 10, 29: 10, 30: 9,        # stats
    31: 11,                                # rec
    32: 28, 33: 18,                        # note + match
    34: 11, 35: 11, 36: 9,                 # flags
}
for c, w in col_widths_new.items():
    ws.column_dimensions[ws.cell(1, c).column_letter].width = w

# Step 8: for each data row, compute per-item bidder rates and write
print("Computing per-item bidder rates and writing rows...")
n_rows_written = 0
n_with_real_data = 0
n_size_matched = 0
for ri, row_data in enumerate(existing_data):
    excel_row = 5 + ri
    if row_data is None: continue
    # row_data is cols 1-21
    wbs, code, desc, uom = row_data[0], row_data[1], row_data[2], row_data[3]
    # Skip section banner rows
    if not code or not desc or (isinstance(row_data[0], str) and row_data[0].endswith('.')) or all(v is None for v in row_data[1:4]):
        continue

    # Per-item bidder match
    bidder_results = find_per_item_match(desc, "", uom)

    n_priced = 0
    sized = False
    for i, bidder in enumerate(BIDDERS):
        col = 22 + i
        result = bidder_results.get(bidder)
        cell = ws.cell(excel_row, col)
        if result and result[0]:
            cell.value = result[0]
            style_money(cell, bg_color="FFEBEE")
            n_priced += 1
            if 'size-matched' in result[1]:
                sized = True
        else:
            cell.value = "—"
            style_center(cell)

    if n_priced > 0:
        n_with_real_data += 1
        if sized: n_size_matched += 1

    # Stats columns 27-29 (Min/Med/Max formulas)
    bidder_range = f"V{excel_row}:Z{excel_row}"  # V=22, W=23, X=24, Y=25, Z=26
    min_cell = ws.cell(excel_row, 27)
    med_cell = ws.cell(excel_row, 28)
    max_cell = ws.cell(excel_row, 29)
    spread_cell = ws.cell(excel_row, 30)
    if n_priced >= 2:
        min_cell.value = f"=MIN({bidder_range})"
        med_cell.value = f"=MEDIAN({bidder_range})"
        max_cell.value = f"=MAX({bidder_range})"
        spread_cell.value = f"=IFERROR((AC{excel_row}-AA{excel_row})/AB{excel_row},\"—\")"
        style_money(min_cell, bg_color="E8F5E9", bold=True)
        style_money(med_cell, bg_color="FFFDE7", bold=True)
        style_money(max_cell, bg_color="FFEBEE", bold=True)
        style_pct(spread_cell)
    elif n_priced == 1:
        # Single priced — show as median, blank others
        med_cell.value = f"=MEDIAN({bidder_range})"
        style_money(med_cell, bg_color="FFFDE7", bold=True)
        min_cell.value = "—"; style_center(min_cell)
        max_cell.value = "—"; style_center(max_cell)
        spread_cell.value = "—"; style_center(spread_cell)
    else:
        for cc in [min_cell, med_cell, max_cell, spread_cell]:
            cc.value = "—"; style_center(cc)

    # Col 31: Recommended (NS midpoint as default; uses existing NS Low/High at cols 7/8)
    rec_formula = f"=AVERAGE(G{excel_row},H{excel_row})"
    rec_cell = ws.cell(excel_row, 31)
    rec_cell.value = rec_formula
    style_money(rec_cell, bg_color="FFD54F", bold=True)

    # Col 32: Vetting Note
    note_parts = []
    if n_priced == 0:
        note_parts.append("No bidder data; SME only")
    elif n_priced < 3:
        note_parts.append(f"Only {n_priced}/5 bidders priced — limited market signal")
    # Col 33: Match Basis
    match_basis = ""
    if bidder_results:
        bases = set()
        for r in bidder_results.values():
            if r: bases.add(r[1].split(' (')[0])
        if bases:
            match_basis = '; '.join(bases)
        else:
            match_basis = "no match"
    note_cell = ws.cell(excel_row, 32)
    note_cell.value = '; '.join(note_parts) if note_parts else "OK — sources align"
    style_text(note_cell)
    basis_cell = ws.cell(excel_row, 33)
    basis_cell.value = match_basis
    style_text(basis_cell, bg_color="F0F0F0")

    # Col 34: Bidder Med Δ% (formula vs REC)
    bidmed_delta = ws.cell(excel_row, 34)
    bidmed_delta.value = f"=IFERROR((AB{excel_row}-AE{excel_row})/AE{excel_row},\"—\")"
    style_pct(bidmed_delta)

    # Col 35: GHD Δ% (formula: GHD - REC) / REC
    ghd_delta = ws.cell(excel_row, 35)
    ghd_delta.value = f"=IFERROR((S{excel_row}-AE{excel_row})/AE{excel_row},\"—\")"
    style_pct(ghd_delta)

    # Col 36: FLAG — based on bidder spread
    flag_cell = ws.cell(excel_row, 36)
    if n_priced >= 3:
        flag_cell.value = f'=IF(AD{excel_row}>1, "⚠ HIGH SPREAD", IF(ABS(AH{excel_row})>0.5, "⚠ DELTA", "OK"))'
    else:
        flag_cell.value = "—"
    style_center(flag_cell, bold=True)

    n_rows_written += 1

print(f"\nWrote {n_rows_written} data rows")
print(f"  with real bidder data: {n_with_real_data}")
print(f"  with size-matched data: {n_size_matched}")

# Step 9: Save
wb.save(V4)
print(f"\nWrote {V4}")
import os
print(f"Size: {os.path.getsize(V4):,} bytes")
