#!/usr/bin/env python3
"""Render PDF pages to PNG for the Class-B (visual) half of the
word-template-fidelity verification loop, and optionally diff two PDFs
pixel-by-pixel.

Cross-platform by design: built on PyMuPDF (pure Python wheel), so it works
identically on Windows / macOS / Linux with no poppler or brew install.

Usage:
  # render every page of a PDF to PNG
  python3 render_pdf_pages.py out.pdf /tmp/pages/new
  # render at higher resolution
  python3 render_pdf_pages.py out.pdf /tmp/pages/new --dpi 150
  # render only selected pages (1-based)
  python3 render_pdf_pages.py out.pdf /tmp/pages/new --pages 1,2,4
  # list fonts used in the PDF + whether they are embedded
  python3 render_pdf_pages.py out.pdf --fonts
  # objective comparison against the reference PDF
  python3 render_pdf_pages.py out.pdf --diff reference.pdf

Exit codes:
  0 = success / pages identical within tolerance
  1 = diff found (page count, page size, or pixels differ)
  2 = usage / IO error
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

try:
    import pymupdf
except ImportError:  # pragma: no cover
    print("ERROR: PyMuPDF is not installed.  pip install pymupdf", file=sys.stderr)
    sys.exit(2)


def parse_pages(spec: str | None, count: int) -> list[int]:
    """'1,2,4' or '2-5' -> zero-based page indexes."""
    if not spec:
        return list(range(count))
    out: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out.extend(range(int(a) - 1, int(b)))
        else:
            out.append(int(part) - 1)
    return [i for i in out if 0 <= i < count]


def cmd_render(pdf: Path, prefix: Path, dpi: int, pages: str | None) -> int:
    doc = pymupdf.open(str(pdf))
    prefix.parent.mkdir(parents=True, exist_ok=True)
    idxs = parse_pages(pages, doc.page_count)
    for i in idxs:
        pix = doc[i].get_pixmap(dpi=dpi)
        out = prefix.parent / f"{prefix.name}-{i + 1:02d}.png"
        pix.save(str(out))
        print(f"page {i + 1}/{doc.page_count} -> {out}  ({pix.width}x{pix.height})")
    print(f"\n{len(idxs)} page(s) rendered from {pdf.name} at {dpi} dpi.")
    print("Now READ each PNG and compare against the reference page by page.")
    return 0


def cmd_fonts(pdf: Path) -> int:
    doc = pymupdf.open(str(pdf))
    print(f"{pdf.name}: {doc.page_count} page(s)")
    seen: set[tuple[str, str]] = set()
    for i in range(doc.page_count):
        for font in doc[i].get_fonts(full=True):
            # (xref, ext, type, basefont, name, encoding, referencer)
            xref, ext, basefont = font[0], font[1], font[3]
            if ext in (None, "", "n/a"):
                # no embedded font file -> confirm via the xref stream table
                try:
                    embedded = doc.xref_is_stream(xref)
                except Exception:
                    embedded = False
                ext = "embedded" if embedded else "NOT-EMBEDDED"
            key = (basefont, ext)
            if key not in seen:
                seen.add(key)
                flag = "" if ext != "NOT-EMBEDDED" else "   <-- 未内嵌！换机器会被替换字体"
                print(f"  p{i + 1}: {basefont}  [{ext}]{flag}")
    if not seen:
        print("  (未检出字体)")
    print("\n提示：交付 PDF 应内嵌全部字体；出现 NOT-EMBEDDED 必须重新导出。")
    return 0


def cmd_diff(a: Path, b: Path, dpi: int, tol: int, threshold: float) -> int:
    da, db = pymupdf.open(str(a)), pymupdf.open(str(b))
    print(f"A: {a.name}  {da.page_count} page(s)")
    print(f"B: {b.name}  {db.page_count} page(s)")
    bad = False
    if da.page_count != db.page_count:
        print(f"FAIL: 页数不同 ({da.page_count} vs {db.page_count})")
        bad = True

    for i in range(min(da.page_count, db.page_count)):
        pa, pb = da[i].get_pixmap(dpi=dpi), db[i].get_pixmap(dpi=dpi)
        if (pa.width, pa.height) != (pb.width, pb.height):
            print(f"FAIL: 第 {i + 1} 页尺寸不同 "
                  f"({pa.width}x{pa.height} vs {pb.width}x{pb.height})")
            bad = True
            continue
        sa, sb = pa.samples, pb.samples
        total = len(sa) or 1
        diff = 0
        for x, y in zip(sa, sb):
            if x - y > tol or y - x > tol:
                diff += 1
        ratio = diff / total
        status = "OK  " if ratio <= threshold else "FAIL"
        if ratio > threshold:
            bad = True
        print(f"  [{status}] 第 {i + 1} 页 差异像素 {ratio:.4%} "
              f"(阈值 {threshold:.2%}, dpi={dpi})")

    if bad:
        print("\n结论：两 PDF 不一致。逐页读图定位差异（分页/对齐/字体/页尾空白）。")
        return 1
    print("\n结论：所有页在阈值内一致 ✅")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdf", type=Path)
    ap.add_argument("prefix", nargs="?", type=Path,
                    help="output PNG prefix, e.g. /tmp/pages/new")
    ap.add_argument("--dpi", type=int, default=100)
    ap.add_argument("--pages", help="1-based selection, e.g. 1,2,4 or 2-5")
    ap.add_argument("--fonts", action="store_true",
                    help="list fonts used in the PDF and whether embedded")
    ap.add_argument("--diff", type=Path, metavar="REFERENCE_PDF",
                    help="compare this PDF against the reference")
    ap.add_argument("--tolerance", type=int, default=16,
                    help="per-channel delta counted as different (0-255)")
    ap.add_argument("--threshold", type=float, default=0.005,
                    help="max fraction of differing pixels to still pass")
    args = ap.parse_args()

    if not args.pdf.exists():
        print(f"ERROR: no such file: {args.pdf}", file=sys.stderr)
        return 2

    if args.fonts:
        return cmd_fonts(args.pdf)
    if args.diff:
        if not args.diff.exists():
            print(f"ERROR: no such file: {args.diff}", file=sys.stderr)
            return 2
        return cmd_diff(args.pdf, args.diff, args.dpi, args.tolerance, args.threshold)
    if not args.prefix:
        ap.print_help()
        return 2
    return cmd_render(args.pdf, args.prefix, args.dpi, args.pages)


if __name__ == "__main__":
    sys.exit(main())
