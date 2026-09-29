#!/usr/bin/env python3
"""Generate an AppleScript that drives Microsoft Word itself to do
paragraph-level find-and-replace on a master document, then save .doc and
export .pdf.

Why Word itself: complex .doc layouts (rotated text boxes, floating tables,
imposed/folded page setups, custom headers) only survive inside Word's own
layout engine. LibreOffice has a different engine (line spacing drifts, dotted
borders become solid, page count changes) and python-docx corrupts rotated text
boxes on a .doc -> .docx round trip. So for these documents the master must be
opened and edited in Microsoft Word.

Usage:
  python3 gen_word_replace_applescript.py pairs.json MASTER.doc NEW.doc NEW.pdf \
      > build.applescript
  osascript build.applescript      # prints "MISS:" + indexes that did not match

pairs.json -- either a bare list, or {"pairs": [...]}:
  [["old paragraph text", "new paragraph text"], ...]

Each old/new string must be ONE paragraph (no newlines) and <= 250 characters
(Word's find limit is 255). Split multi-paragraph changes into one pair per
paragraph so intra-paragraph formatting is preserved.

Every pair MUST match; the script reports the indexes that did not, so you can
fix the text and re-run rather than silently shipping a half-replaced document.
"""
from __future__ import annotations

import json
import sys

FIND_LIMIT = 250  # Word's hard limit is 255; stay under it.


def esc(s: str) -> str:
    return s.replace("\\", "\\\\").replace('"', '\\"')


def load_pairs(path: str) -> list[list[str]]:
    data = json.loads(open(path, encoding="utf-8").read())
    pairs = data["pairs"] if isinstance(data, dict) else data
    for pair in pairs:
        if len(pair) != 2:
            raise ValueError(f"each pair needs exactly 2 items: {pair!r}")
        for text in pair:
            if "\n" in text:
                raise ValueError(
                    f"newline in pair (split it into one pair per paragraph): "
                    f"{text[:40]!r}")
            if len(text) > FIND_LIMIT:
                raise ValueError(
                    f"text exceeds Word find limit ({FIND_LIMIT}): {text[:40]!r}")
    return pairs


def build(pairs, src, outdoc, outpdf) -> str:
    olds = ", ".join(f'"{esc(o)}"' for o, _ in pairs)
    news = ", ".join(f'"{esc(n)}"' for _, n in pairs)
    return f'''on run
  set pairsOld to {{{olds}}}
  set pairsNew to {{{news}}}
  set missList to {{}}
  with timeout of 600 seconds
    tell application "Microsoft Word"
      open POSIX file "{src}"
      set theDoc to active document
      repeat with i from 1 to count of pairsOld
        set findObj to find object of text object of theDoc
        tell findObj
          set match case to true
          set ok to execute find find text (item i of pairsOld) replace with (item i of pairsNew) replace replace all
        end tell
        if not ok then set end of missList to i
      end repeat
      -- save as renames the open document; later references use the new name
      save as theDoc file name "{outdoc}" file format format document
      save as theDoc file name "{outpdf}" file format format PDF
      close theDoc saving no
    end tell
  end timeout
  set AppleScript's text item delimiters to ","
  return "MISS:" & (missList as string)
end run'''


def main() -> int:
    if len(sys.argv) != 5:
        print(__doc__, file=sys.stderr)
        return 2
    pairs_path, src, outdoc, outpdf = sys.argv[1:5]
    try:
        pairs = load_pairs(pairs_path)
    except (ValueError, KeyError) as e:
        print(f"ERROR: bad pairs file: {e}", file=sys.stderr)
        return 2
    print(build(pairs, src, outdoc, outpdf))
    return 0


if __name__ == "__main__":
    sys.exit(main())
