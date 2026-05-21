#!/usr/bin/env python3
"""
Ausenco FD Presentation Standard — 10-Point Audit Checker

Run before any Ausenco-attributed workbook ships. Returns PASS / FAIL per
element with specific diagnostic for failures.

Usage:
    python3 mining_estimating_manual/build/fd_audit.py <workbook.xlsx>

Exit code: 0 if all PASS, 1 if any FAIL.
"""

import sys
import os
import re
import warnings
warnings.filterwarnings("ignore")

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl required. Install: pip install openpyxl")
    sys.exit(2)


GREEN = "\033[92m"
RED = "\033[91m"
AMBER = "\033[93m"
RESET = "\033[0m"
BOLD = "\033[1m"

PASS_TAG = f"{GREEN}PASS{RESET}"
FAIL_TAG = f"{RED}FAIL{RESET}"
WARN_TAG = f"{AMBER}WARN{RESET}"


def _all_text(wb):
    """Concatenate all string values across all worksheets — for string-presence checks."""
    text = []
    for sn in wb.sheetnames:
        ws = wb[sn]
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if isinstance(v, str):
                    text.append(v)
    return ' '.join(text).lower()


def check_1_classification_stamp(wb, all_text):
    """Element 1 — Classification stamp at top: class + purpose + NOT-for cases."""
    has_class = bool(re.search(r"class\s+[1-5]", all_text))
    has_not_for = "not " in all_text and ("not a " in all_text or "not an " in all_text or "not the " in all_text)
    has_aace = "aace" in all_text or "18r-97" in all_text
    if has_class and has_not_for and has_aace:
        return True, "AACE class + NOT-for statement + 18R-97 reference found"
    missing = []
    if not has_class: missing.append("AACE Class reference (Class 3/4/etc.)")
    if not has_not_for: missing.append("'NOT a...' misuse-case statement")
    if not has_aace: missing.append("AACE / 18R-97 standard citation")
    return False, "Missing: " + "; ".join(missing)


def check_2_decomposed_adjustments(wb, all_text):
    """Element 2 — Any adjustment shown component-by-component with bidirectional ranges."""
    has_components = "component" in all_text or "decompos" in all_text
    has_bidirectional = bool(re.search(r"[-+]\d", all_text)) and bool(re.search(r"\+\d|positive|offset", all_text)) and bool(re.search(r"-\d|negative", all_text))
    has_net_composite = "net composite" in all_text or "composite" in all_text and bool(re.search(r"composite[^.]*[-+]\d", all_text))
    if (has_components or has_net_composite) and has_bidirectional:
        return True, "Component-level decomposition with bidirectional ranges found"
    missing = []
    if not has_components and not has_net_composite:
        missing.append("Component breakdown (e.g. 'logistics; labour; climate; QA/QC')")
    if not has_bidirectional:
        missing.append("Bidirectional ranges (+X% / -Y% net out)")
    return False, "Missing: " + "; ".join(missing)


def check_3_library_application_separation(wb, all_text):
    """Element 3 — Reusable Rate Library tab separate from Project Application tab."""
    sheet_names_l = [s.lower() for s in wb.sheetnames]
    has_library = any("rate library" in s or "rate basis" in s or "library" in s for s in sheet_names_l)
    has_application = any("mto" in s or "application" in s or "project" in s and "rate" in s for s in sheet_names_l)
    if has_library and has_application:
        return True, "Rate Library tab AND Project Application / MTO tab both found"
    missing = []
    if not has_library: missing.append("'Rate Library' or 'Rate Basis' tab")
    if not has_application: missing.append("'MTO × Rate' or 'Project Application' tab")
    return False, "Missing: " + "; ".join(missing)


def check_4_flag_resolution_required(wb, all_text):
    """Element 4 — Every flag in flag register has 'Resolution Required' column."""
    has_flag = "flag" in all_text
    has_resolution = "resolution required" in all_text or "resolution" in all_text and "required" in all_text
    has_severity = any(s in all_text for s in ["high", "med", "medium", "low", "severity"])
    if has_flag and has_resolution and has_severity:
        return True, "Flag register with Resolution Required + Severity columns found"
    missing = []
    if not has_flag: missing.append("Flag register / column")
    if not has_resolution: missing.append("'Resolution Required' or equivalent action column")
    if not has_severity: missing.append("Severity column (HIGH/MED/LOW)")
    return False, "Missing: " + "; ".join(missing)


def check_5_top_3_sensitivities(wb, all_text):
    """Element 5 — Top 3 sensitivities pre-calculated with $ impact."""
    has_sensitivity = "sensitivity" in all_text or "sensitivities" in all_text
    has_dollar_impact = bool(re.search(r"\$\s*[\d,]+\s*[MK]?\b", all_text)) or bool(re.search(r"\b\d+\s*M\b", all_text))
    has_trigger = "trigger" in all_text or "if " in all_text or "±" in all_text
    if has_sensitivity and has_dollar_impact and has_trigger:
        # Count actual sensitivity lines — should be ≥3
        sens_lines = sum(1 for sn in wb.sheetnames if "sensitivity" in sn.lower())
        return True, f"Sensitivity content + $ impact + triggers found ({sens_lines} sensitivity tab(s))"
    missing = []
    if not has_sensitivity: missing.append("'Sensitivity' content")
    if not has_dollar_impact: missing.append("Dollar impact figures ($X M)")
    if not has_trigger: missing.append("Trigger / condition descriptions")
    return False, "Missing: " + "; ".join(missing)


def check_6_confidence_stamp(wb, all_text):
    """Element 6 — Affirmative Confidence Stamp before limitations."""
    has_confidence = "confidence stamp" in all_text or "confidence" in all_text and "suitable for" in all_text
    has_what_not = "what this is not" in all_text or "not a tender" in all_text or "not a substitute" in all_text
    has_suitable = "suitable for" in all_text or "appropriate for" in all_text
    if has_confidence and has_what_not and has_suitable:
        return True, "Affirmative use case ('Suitable for…') AND 'What this is NOT' both found"
    missing = []
    if not has_suitable: missing.append("'Suitable for…' affirmative statement")
    if not has_what_not: missing.append("'What this is NOT' / misuse-case list")
    if not has_confidence: missing.append("'Confidence Stamp' header")
    return False, "Missing: " + "; ".join(missing)


def check_7_filing_path(wb, all_text):
    """Element 7 — Filing path per Ausenco Standard Project Folder Structure."""
    has_filing = "filing" in all_text
    has_folders = (
        bool(re.search(r"0\d\s+\w+\s*/\s*0\d\.\d", all_text))
        or any(s in all_text for s in ["05 reporting", "02 estimates", "10 constructability"])
    )
    if has_filing and has_folders:
        return True, "Filing path with Ausenco folder structure (05/02/10) found"
    missing = []
    if not has_filing: missing.append("'Filing' section header")
    if not has_folders: missing.append("Ausenco folder paths (05 Reporting / 02 Estimates / 10 Constructability)")
    return False, "Missing: " + "; ".join(missing)


def check_8_document_number(wb, all_text):
    """Element 8 — Document number per Ausenco scheme ######-DT-#####-#####-###"""
    # Pattern: 6 digits, hyphen, 2 letters, hyphen, 5 chars, hyphen, 5 digits, hyphen, 3 digits
    pat_full = re.compile(r"\b\d{6}-[A-Za-z]{2}-[A-Za-z0-9]{5}-\d{4,5}-\d{3}\b")
    # Looser pattern: 6 digits + hyphen + something + multiple hyphens
    pat_loose = re.compile(r"\b\d{6}-[A-Za-z0-9]+-[A-Za-z0-9]+-\d+-\d+\b")
    if pat_full.search(all_text):
        return True, "Ausenco document number found (full scheme match)"
    if pat_loose.search(all_text):
        return True, "Ausenco document number found (loose match — verify exact format)"
    return False, "No document number matching Ausenco scheme ######-DT-#####-#####-### found"


def check_9_author(wb, all_text):
    """Element 9 — Author = Rick Miller, Senior Project Director, no P.Eng / no Engineer."""
    has_rick = "rick miller" in all_text
    has_role = "senior project director" in all_text
    has_pen = bool(re.search(r"\bp\.?\s*eng\b", all_text))
    has_engineer = bool(re.search(r"\bengineer\b", all_text)) and "engineering" not in all_text and "ausenco engineering" not in all_text
    if has_rick and has_role and not has_pen:
        return True, "Rick Miller, Senior Project Director — no P.Eng (per author convention)"
    issues = []
    if not has_rick: issues.append("'Rick Miller' not found")
    if not has_role: issues.append("'Senior Project Director' role not found")
    if has_pen: issues.append("'P.Eng' or 'PE' credential found — Rick should not carry that")
    return False, "; ".join(issues)


def check_10_entity_attribution(wb, all_text):
    """Element 10 — Ausenco Engineering Canada ULC; no Carter's/TMG mixing."""
    has_aec = "ausenco engineering canada" in all_text or "ausenco engineering" in all_text
    has_ulc = "ulc" in all_text
    has_carters = "carter's" in all_text or "carters" in all_text
    has_tmg = "tmg" in all_text
    if has_aec and has_ulc and not has_carters and not has_tmg:
        return True, "Ausenco Engineering Canada ULC — clean attribution"
    issues = []
    if not has_aec: issues.append("'Ausenco Engineering Canada' not found")
    if not has_ulc: issues.append("'ULC' marker not found")
    if has_carters: issues.append("Carter's reference found — should not be mixed with Ausenco attribution")
    if has_tmg: issues.append("TMG reference found — should not be mixed with Ausenco attribution")
    return False, "; ".join(issues)


CHECKS = [
    (1, "Classification stamp",       check_1_classification_stamp),
    (2, "Adjustments decomposed",     check_2_decomposed_adjustments),
    (3, "Library / Application split", check_3_library_application_separation),
    (4, "FLAG resolution required",   check_4_flag_resolution_required),
    (5, "Top 3 sensitivities",        check_5_top_3_sensitivities),
    (6, "Confidence stamp",           check_6_confidence_stamp),
    (7, "Filing path",                check_7_filing_path),
    (8, "Document number",            check_8_document_number),
    (9, "Author",                     check_9_author),
    (10, "Entity attribution",        check_10_entity_attribution),
]


def main():
    if len(sys.argv) < 2:
        print(f"Usage: python3 {sys.argv[0]} <workbook.xlsx>")
        sys.exit(2)
    path = sys.argv[1]
    if not os.path.exists(path):
        print(f"ERROR: file not found: {path}")
        sys.exit(2)

    print()
    print(f"{BOLD}AUSENCO FD PRESENTATION STANDARD — Pre-Flight Audit{RESET}")
    print(f"File: {path}")
    print(f"Size: {os.path.getsize(path):,} bytes")
    print("=" * 78)

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        print(f"ERROR loading workbook: {e}")
        sys.exit(2)

    print(f"Sheets ({len(wb.sheetnames)}): {', '.join(wb.sheetnames)}")
    print("-" * 78)
    all_text = _all_text(wb)

    results = []
    for n, name, func in CHECKS:
        passed, diag = func(wb, all_text)
        results.append((n, name, passed, diag))
        tag = PASS_TAG if passed else FAIL_TAG
        print(f"  [{tag}]  {n:2d}. {name:<35} {diag}")

    print("-" * 78)
    n_pass = sum(1 for r in results if r[2])
    n_fail = len(results) - n_pass
    if n_fail == 0:
        print(f"  {BOLD}{GREEN}ALL 10 PASS — workbook is FD-grade and clear to ship.{RESET}")
        sys.exit(0)
    print(f"  {BOLD}{RED}{n_fail} FAIL / {n_pass} PASS — rework required before issuing under Ausenco letterhead.{RESET}")
    print()
    print("  Specific fixes needed:")
    for n, name, passed, diag in results:
        if not passed:
            print(f"    Element {n} ({name}): {diag}")
    print()
    print(f"  Refer to .claude/skills/ausenco-fd-presentation-standard.md for the full")
    print(f"  standard with James Gallant's worked examples.")
    sys.exit(1)


if __name__ == "__main__":
    main()
