#!/usr/bin/env python3
"""Verify a .docx against a machine-checkable spec sheet (spec.json).

This implements the "Class A" (static) half of the word-template-fidelity
verification loop: properties that can be read straight out of the XML and
asserted equal to the required value -- the four font slots, font size, page
margins, page size, section count, table structure signature, and whether the
required fonts are actually installed on this machine.

"Class B" properties (pagination, cross-page breaks, alignment, trailing
whitespace, silent font substitution, visual overlap) CANNOT be checked here.
Export to PDF, render page images, and compare against the reference PDF --
use render_pdf_pages.py.

Cross-platform: runs identically on Windows / macOS / Linux.

Usage:
  python3 verify_docx_spec.py output.docx spec.json
  python3 verify_docx_spec.py output.docx spec.json --print-signature
  python3 verify_docx_spec.py --list-fonts
  python3 verify_docx_spec.py --describe-size 12

Exit codes:
  0 = all checks PASS
  1 = at least one check FAIL
  2 = usage / IO / import error

spec.json schema (every key optional; omit a key to skip that check):

{
  "font": {
    "eastAsia": "仿宋_GB2312",     # CJK glyphs -- w:rFonts/@w:eastAsia
    "ascii": "Times New Roman",    # western      -- w:rFonts/@w:ascii
    "hAnsi": "Times New Roman",    # western ANSI -- w:rFonts/@w:hAnsi
    "cs": null,                    # complex script
    "size_pt": 14,                 # dominant run size in points (14pt = 四号)
    "size_tolerance_pt": 0.5,
    "min_coverage": 0.5,           # required share of runs using that font
    "require_installed": true      # fail if the font is not installed locally
  },
  "margins_cm": {"top": 2.54, "bottom": 2.54, "left": 3.17, "right": 3.17,
                 "tolerance_cm": 0.2},
  "page": {"width_cm": 21.0, "height_cm": 29.7, "tolerance_cm": 0.2},
  "sections": 1,
  "tables": [{"index": 0, "grid_cols": 4, "rows": 3}],
  "body_paragraphs_min": 20,
  "no_theme_fonts": true           # fail if any *Theme font attribute remains
}
"""
from __future__ import annotations

import argparse
import collections
import json
import os
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

FONT_SLOTS = ("ascii", "hAnsi", "eastAsia", "cs")
THEME_ATTRS = ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme")

# Chinese type sizes: 号 -> points.  w:sz stores half-points.
CJK_SIZES = [
    (44, "初号"), (36, "小初"), (26, "一号"), (24, "小一"), (22, "二号"),
    (18, "小二"), (16, "三号"), (15, "小三"), (14, "四号"), (12, "小四"),
    (10.5, "五号"), (9, "小五"), (7.5, "六号"), (6.5, "小六"), (5.5, "七号"),
    (5, "八号"),
]

results: list[tuple[str, str, str, str]] = []


def record(status: str, name: str, expected, actual, note: str = "") -> None:
    results.append((status, name, str(expected), str(actual)))


def describe_size(pt: float) -> str:
    for value, label in CJK_SIZES:
        if abs(value - pt) < 0.01:
            return f"{pt:g} pt = {label}"
    return f"{pt:g} pt (非标准中文字号)"


# --------------------------------------------------------------------------
# document inspection
# --------------------------------------------------------------------------

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
    """Return (slot counters, size counters, theme-attr counter)."""
    slots = {s: collections.Counter() for s in FONT_SLOTS}
    sizes = collections.Counter()
    themes = collections.Counter()
    for run in _runs(doc):
        rPr = run._element.rPr
        if rPr is None:
            continue
        rf = rPr.rFonts
        if rf is not None:
            for slot in FONT_SLOTS:
                val = rf.get(qn(f"w:{slot}"))
                if val:
                    slots[slot][val] += 1
            for attr in THEME_ATTRS:
                val = rf.get(qn(f"w:{attr}"))
                if val:
                    themes[f"{attr}={val}"] += 1
        sz = rPr.find(qn("w:sz"))
        if sz is not None:
            val = sz.get(qn("w:val"))
            if val:
                sizes[round(int(val) / 2, 1)] += 1
    return slots, sizes, themes


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


# --------------------------------------------------------------------------
# installed fonts (cross-platform)
# --------------------------------------------------------------------------

def _font_dirs() -> list[str]:
    home = os.path.expanduser("~")
    if sys.platform == "darwin":
        return [
            "/System/Library/Fonts",
            "/System/Library/Fonts/Supplemental",
            "/Library/Fonts",
            os.path.join(home, "Library", "Fonts"),
        ]
    if sys.platform.startswith("win"):
        windir = os.environ.get("WINDIR", r"C:\Windows")
        return [
            os.path.join(windir, "Fonts"),
            os.path.join(os.environ.get("LOCALAPPDATA", home),
                         "Microsoft", "Windows", "Fonts"),
        ]
    return ["/usr/share/fonts", "/usr/local/share/fonts",
            os.path.join(home, ".fonts"), os.path.join(home, ".local/share/fonts")]


_installed_cache: set[str] | None = None


def _family_names(font) -> set[str]:
    names: set[str] = set()
    try:
        for rec in font["name"].names:
            if rec.nameID in (1, 4, 16):
                try:
                    names.add(rec.toUnicode().strip())
                except Exception:
                    pass
    except Exception:
        pass
    return names


def installed_fonts(refresh: bool = False) -> set[str]:
    """All font family names installed on this machine (best effort)."""
    global _installed_cache
    if _installed_cache is not None and not refresh:
        return _installed_cache

    names: set[str] = set()
    try:
        from fontTools.ttLib import TTCollection, TTFont
    except ImportError:
        TTFont = TTCollection = None

    for d in _font_dirs():
        if not os.path.isdir(d):
            continue
        for fn in sorted(os.listdir(d)):
            if not fn.lower().endswith((".ttf", ".otf", ".ttc", ".otc")):
                continue
            path = os.path.join(d, fn)
            stem = os.path.splitext(fn)[0]
            if TTFont is None:
                names.add(stem)
                continue
            try:
                if fn.lower().endswith((".ttc", ".otc")):
                    for f in TTCollection(path, lazy=True).fonts:
                        names |= _family_names(f)
                else:
                    names |= _family_names(TTFont(path, lazy=True, fontNumber=0))
            except Exception:
                names.add(stem)
    _installed_cache = names
    return names


def normalize_font_name(name: str) -> str:
    return "".join(name.split()).lower().replace("-", "")


def font_is_installed(name: str, installed: set[str]) -> bool:
    target = normalize_font_name(name)
    if not target:
        return False
    return any(normalize_font_name(n) == target for n in installed)


# --------------------------------------------------------------------------
# checks
# --------------------------------------------------------------------------

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
    slots, sizes, _themes = collect_fonts(doc)
    total = max((sum(c.values()) for c in slots.values()), default=0) or 1
    min_cov = float(f.get("min_coverage", 0.5))
    installed = None

    for slot in FONT_SLOTS:
        if slot not in f or f[slot] is None:
            continue
        want = f[slot]
        counter = slots[slot]
        got = dominant(counter)
        cov = coverage(counter, want, total)
        label = f"字体槽 {slot}"
        if got == want and cov >= min_cov:
            record(PASS, label, f"{want} (>= {min_cov:.0%})", f"{got} ({cov:.0%})")
        else:
            record(FAIL, label, f"{want} (>= {min_cov:.0%})",
                   f"{got} ({cov:.0%})  top3={counter.most_common(3)}")

        if f.get("require_installed"):
            if installed is None:
                installed = installed_fonts()
            ok = font_is_installed(want, installed)
            record(PASS if ok else FAIL, f"字体已安装 {want}",
                   "本机已安装", "已安装" if ok else "未安装（Word 会静默回退！）")

    if "size_pt" in f:
        want = float(f["size_pt"])
        tol = float(f.get("size_tolerance_pt", 0.5))
        got = dominant(sizes)
        if got is not None and abs(got - want) <= tol:
            record(PASS, "字号 size_pt", f"{want} ±{tol} ({describe_size(want)})", got)
        else:
            record(FAIL, "字号 size_pt", f"{want} ±{tol} ({describe_size(want)})",
                   f"{got}  all={sizes.most_common(5)}")


def check_theme_fonts(doc, spec):
    if not spec.get("no_theme_fonts"):
        record(SKIP, "theme fonts", "-", "-", "not in spec")
        return
    _slots, _sizes, themes = collect_fonts(doc)
    if themes:
        record(FAIL, "无残留主题字体", "无 *Theme 属性",
               f"{dict(themes)}  (主题字体会在渲染时覆盖显式字体名)")
    else:
        record(PASS, "无残留主题字体", "无 *Theme 属性", "无")


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
        record(PASS if abs(got - want) <= tol else FAIL,
               f"页边距 {key} (cm)", f"{want} ±{tol}", got)


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
        record(PASS if abs(got - want) <= tol else FAIL,
               f"纸张 {key}", f"{want} ±{tol}", got)


def check_sections(doc, spec):
    if "sections" not in spec:
        record(SKIP, "sections", "-", "-", "not in spec")
        return
    want, got = int(spec["sections"]), len(doc.sections)
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
        grid_cols, rows = table_signature(doc.tables[idx])
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


# --------------------------------------------------------------------------
# entry points
# --------------------------------------------------------------------------

def cmd_list_fonts(filter_text: str | None) -> int:
    fonts = sorted(installed_fonts(), key=lambda s: s.lower())
    if filter_text:
        needle = normalize_font_name(filter_text)
        fonts = [f for f in fonts if needle in normalize_font_name(f)]
    if not fonts:
        print("未找到字体（fontTools 可能未安装，或字体目录为空）", file=sys.stderr)
        print("提示: pip install fonttools", file=sys.stderr)
        return 1
    print(f"本机已安装字体 {len(fonts)} 个"
          f"{f'（匹配 \"{filter_text}\"）' if filter_text else ''}:")
    for f in fonts:
        print(" ", f)
    return 0


def cmd_describe_size(pt: float) -> int:
    print(f"{describe_size(pt)}   ->  w:sz val=\"{int(round(pt * 2))}\"")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx", nargs="?", type=Path)
    ap.add_argument("spec", nargs="?", type=Path)
    ap.add_argument("--print-signature", action="store_true",
                    help="dump every table structure signature")
    ap.add_argument("--list-fonts", nargs="?", const="", metavar="FILTER",
                    help="list installed font families, optionally filtered")
    ap.add_argument("--describe-size", type=float, metavar="PT",
                    help="convert points to the Chinese type-size name")
    args = ap.parse_args()

    if args.list_fonts is not None:
        return cmd_list_fonts(args.list_fonts or None)
    if args.describe_size is not None:
        return cmd_describe_size(args.describe_size)

    if not args.docx or not args.spec:
        ap.print_help()
        return 2
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
    check_theme_fonts(doc, spec)
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
        print(f"[{mark}] {name:<{width}} expected={expected:<26} actual={actual}")

    fails = [r for r in results if r[0] == FAIL]
    print("-" * (width + 46))
    print(f"{len(results) - len(fails)}/{len(results)} checks passed "
          f"({sum(1 for r in results if r[0] == SKIP)} skipped)")
    if fails:
        print("\nFAILED -- the document does NOT match the spec:")
        for _, name, expected, actual in fails:
            print(f"  - {name}: expected {expected}, got {actual}")
        print("\nClass-B items (pagination / alignment / trailing whitespace / "
              "silent font substitution) still need a PDF render comparison.")
        return 1
    print("\nAll Class-A checks passed. Now do the Class-B render comparison:")
    print("  python3 render_pdf_pages.py out.pdf /tmp/pages/new")
    print("  python3 render_pdf_pages.py reference.pdf /tmp/pages/official")
    return 0


if __name__ == "__main__":
    sys.exit(main())
