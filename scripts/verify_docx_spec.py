#!/usr/bin/env python3
"""Verify a .docx against a machine-checkable spec sheet (spec.json).

This implements the "Class A" (static) half of the word-template-fidelity
verification loop: properties that can be read straight out of the XML and
asserted equal to the required value -- fonts, sizes, page margins, page size,
section count, table structure signature.

"Class B" properties (pagination, cross-page breaks, alignment, trailing
whitespace, visual overlap) CANNOT be checked here. Export the document to PDF,
render with `pdftoppm -png -r 100`, and compare page images against the
reference PDF.

Usage:
  python3 verify_docx_spec.py output.docx spec.json
  python3 verify_docx_spec.py output.docx spec.json --print-signature

Exit codes:
  0 = all checks PASS
  1 = at least one check FAIL
  2 = usage / IO / import error

spec.json schema (every key optional; omit a key to skip that check):

{
  "font": {
    "eastAsia": "仿宋_GB2312",        # CJK font name, from w:rFonts/@w:eastAsia
    "ascii": "Times New Roman",       # western font, from w:rFonts/@w:ascii
    "size_pt": 14,                    # dominant run size, in points
    "size_tolerance_pt": 0.5,
    "min_coverage": 0.5               # required share of runs using that font
  },
  "margins_cm": {"top": 2.54, "bottom": 2.54, "left": 3.17, "right": 3.17,
                 "tolerance_cm": 0.2},
  "page": {"width_cm": 21.0, "height_cm": 29.7, "tolerance_cm": 0.2},
  "sections": 1,
  "tables": [{"index": 0, "grid_cols": 4, "rows": 3}],
  "body_paragraphs_min": 20
}
"""
from __future__ import annotations

import argparse
import collections
import json
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.oxml.ns import qn
except ImportError:  # pragma: no cover
    print("ERROR: python-docx is not installed.  pip install python-docx",
          file=sys.stderr)
    sys.exit(2)

PASS, FAIL, SKIP = "PASS", "FAIL", "SKIP"
results: list[tuple[str, str, str, str]] = []


def record(status: str, name: str, expected, actual, note: str = "") -> None:
    results.append((status, name, str(expected), str(actual)))


def _runs(doc):
    """Yield every run in body paragraphs and inside tables."""
    for para in doc.paragraphs:
        yield from para.runs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    yield from para.runs


def collect_fonts(doc):
    east, ascii_, sizes = (collections.Counter() for _ in range(3))
    for run in _runs(doc):
        rPr = run._element.rPr
        if rPr is None:
            continue
        rf = rPr.rFonts
        if rf is not None:
            e = rf.get(qn("w:eastAsia"))
            a = rf.get(qn("w:ascii"))
            if e:
                east[e] += 1
            if a:
                ascii_[a] += 1
        sz = rPr.find(qn("w:sz"))
        if sz is not None:
            val = sz.get(qn("w:val"))
            if val:
                sizes[round(int(val) / 2, 1)] += 1
    return east, ascii_, sizes


def table_signature(table):
    """(grid_cols, ((gridSpan, vMerge), ...) per row) -- must be unchanged."""
    tbl = table._tbl
    grid_cols = len(tbl.tblGrid.gridCol_lst) if tbl.tblGrid is not None else 0
    rows = []
    for tr in tbl.tr_lst:
        row = []
        for tc in tr.tc_lst:
            span, vmerge = 1, None
            tcPr = tc.tcPr
            if tcPr is not None:
                if tcPr.gridSpan is not None:
                    span = int(tcPr.gridSpan.val)
                if tcPr.vMerge is not None:
                    vmerge = tcPr.vMerge.val
            row.append((span, vmerge))
        rows.append(tuple(row))
    return grid_cols, tuple(rows)


def dominant(counter):
    return counter.most_common(1)[0][0] if counter else None


def coverage(counter, value, total):
    if not total or value is None:
        return 0.0
    return counter.get(value, 0) / total


def check_font(doc, spec):
    f = spec.get("font")
    if not f:
        record(SKIP, "font", "-", "-", "not in spec")
        return
    east, ascii_, sizes = collect_fonts(doc)
    total = sum(east.values()) or 1

    for key, counter, label in (("eastAsia", east, "中文字体 eastAsia"),
                                ("ascii", ascii_, "西文字体 ascii")):
        if key not in f:
            continue
        want = f[key]
        got = dominant(counter)
        min_cov = float(f.get("min_coverage", 0.5))
        cov = coverage(counter, want, total)
        if got == want and cov >= min_cov:
            record(PASS, label, f"{want} (>= {min_cov:.0%})",
                   f"{got} ({cov:.0%})")
        else:
            record(FAIL, label, f"{want} (>= {min_cov:.0%})",
                   f"{got} ({cov:.0%})  top3={counter.most_common(3)}")

    if "size_pt" in f:
        want = float(f["size_pt"])
        tol = float(f.get("size_tolerance_pt", 0.5))
        got = dominant(sizes)
        if got is not None and abs(got - want) <= tol:
            record(PASS, "字号 size_pt", f"{want} ±{tol}", got)
        else:
            record(FAIL, "字号 size_pt", f"{want} ±{tol}",
                   f"{got}  all={sizes.most_common(5)}")


def check_margins(doc, spec):
    m = spec.get("margins_cm")
    if not m:
        record(SKIP, "margins", "-", "-", "not in spec")
        return
    tol = float(m.get("tolerance_cm", 0.2))
    sec = doc.sections[0]
    for key, attr in (("top", "top_margin"), ("bottom", "bottom_margin"),
                      ("left", "left_margin"), ("right", "right_margin")):
        if key not in m:
            continue
        want = float(m[key])
        got = round(getattr(sec, attr).cm, 3)
        if abs(got - want) <= tol:
            record(PASS, f"页边距 {key} (cm)", f"{want} ±{tol}", got)
        else:
            record(FAIL, f"页边距 {key} (cm)", f"{want} ±{tol}", got)


def check_page(doc, spec):
    p = spec.get("page")
    if not p:
        record(SKIP, "page size", "-", "-", "not in spec")
        return
    tol = float(p.get("tolerance_cm", 0.2))
    sec = doc.sections[0]
    for key, attr in (("width_cm", "page_width"), ("height_cm", "page_height")):
        if key not in p:
            continue
        want = float(p[key])
        got = round(getattr(sec, attr).cm, 3)
        if abs(got - want) <= tol:
            record(PASS, f"纸张 {key}", f"{want} ±{tol}", got)
        else:
            record(FAIL, f"纸张 {key}", f"{want} ±{tol}", got)


def check_sections(doc, spec):
    if "sections" not in spec:
        record(SKIP, "sections", "-", "-", "not in spec")
        return
    want = int(spec["sections"])
    got = len(doc.sections)
    record(PASS if got == want else FAIL, "分节数 sections", want, got)


def check_tables(doc, spec, print_sig=False):
    if "tables" not in spec:
        record(SKIP, "tables", "-", "-", "not in spec")
        return
    if print_sig:
        for i, t in enumerate(doc.tables):
            print(f"  table[{i}] signature = {table_signature(t)}")
    for want in spec["tables"]:
        idx = int(want.get("index", 0))
        if idx >= len(doc.tables):
            record(FAIL, f"表格[{idx}] 存在", "存在", f"仅有 {len(doc.tables)} 个表")
            continue
        t = doc.tables[idx]
        grid_cols, rows = table_signature(t)
        if "grid_cols" in want:
            record(PASS if grid_cols == int(want["grid_cols"]) else FAIL,
                   f"表格[{idx}] grid 列数", want["grid_cols"], grid_cols)
        if "rows" in want:
            record(PASS if len(rows) == int(want["rows"]) else FAIL,
                   f"表格[{idx}] 行数", want["rows"], len(rows))
        if "signature" in want:
            got = [list(r) for r in rows]
            record(PASS if got == want["signature"] else FAIL,
                   f"表格[{idx}] 结构签名", want["signature"], got)


def check_body(doc, spec):
    if "body_paragraphs_min" not in spec:
        record(SKIP, "body_paragraphs", "-", "-", "not in spec")
        return
    want = int(spec["body_paragraphs_min"])
    got = len([p for p in doc.paragraphs if p.text.strip()])
    record(PASS if got >= want else FAIL, "正文非空段落数 ≥", want, got)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx", type=Path)
    ap.add_argument("spec", type=Path)
    ap.add_argument("--print-signature", action="store_true",
                    help="dump every table structure signature")
    args = ap.parse_args()

    if not args.docx.exists():
        print(f"ERROR: no such file: {args.docx}", file=sys.stderr)
        return 2
    if not args.spec.exists():
        print(f"ERROR: no such file: {args.spec}", file=sys.stderr)
        return 2

    try:
        spec = json.loads(args.spec.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        print(f"ERROR: spec.json is not valid JSON: {e}", file=sys.stderr)
        return 2

    doc = Document(str(args.docx))

    if args.print_signature:
        print("table signatures:")

    check_font(doc, spec)
    check_margins(doc, spec)
    check_page(doc, spec)
    check_sections(doc, spec)
    check_tables(doc, spec, args.print_signature)
    check_body(doc, spec)

    width = max(len(r[1]) for r in results) + 2
    print(f"\nSpec check: {args.docx.name}  (vs {args.spec.name})")
    print("-" * (width + 46))
    for status, name, expected, actual in results:
        mark = {"PASS": "PASS", "FAIL": "FAIL", "SKIP": "skip"}[status]
        print(f"[{mark}] {name:<{width}} expected={expected:<22} actual={actual}")

    fails = [r for r in results if r[0] == FAIL]
    print("-" * (width + 46))
    print(f"{len(results) - len(fails)}/{len(results)} checks passed "
          f"({sum(1 for r in results if r[0] == SKIP)} skipped)")
    if fails:
        print("\nFAILED -- the document does NOT match the spec:")
        for _, name, expected, actual in fails:
            print(f"  - {name}: expected {expected}, got {actual}")
        print("\nRemember: Class-B items (pagination / alignment / trailing "
              "whitespace) still need a PDF render comparison.")
        return 1
    print("\nAll Class-A checks passed. Now do the Class-B render comparison:")
    print("  pdftoppm -png -r 100 out.pdf /tmp/pages/new")
    print("  pdftoppm -png -r 100 reference.pdf /tmp/pages/official")
    return 0


if __name__ == "__main__":
    sys.exit(main())
