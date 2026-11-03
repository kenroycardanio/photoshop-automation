#!/usr/bin/env python
"""Convert Excel (.xlsx) to UTF-8 TSV for Photoshop batch_fill.jsx.
Preserves LRN / long numbers as full digit strings (no scientific notation).
"""
from __future__ import print_function

import sys
import os

def cell_to_text(value, number_format=None):
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value == int(value) and abs(value) < 1e15:
            return str(int(round(value)))
        # avoid scientific in text
        s = ("%.15f" % value).rstrip("0").rstrip(".")
        return s
    s = str(value).strip()
    # Excel sometimes already stored sci-notation text
    import re
    m = re.match(r"^([+-]?)(\d+(?:\.\d+)?)[eE]([+-]?\d+)$", s)
    if m:
        sign, mant, exp = m.group(1), m.group(2), int(m.group(3))
        if "." in mant:
            whole, frac = mant.split(".", 1)
        else:
            whole, frac = mant, ""
        digits = whole + frac
        new_len = len(whole) + exp
        if new_len <= 0:
            return "0"
        if new_len > len(digits):
            digits = digits + ("0" * (new_len - len(digits)))
        else:
            digits = digits[:new_len]
        digits = digits.lstrip("0") or "0"
        return ("-" if sign == "-" else "") + digits
    return s


def escape_tsv(s):
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = s.replace("\\", "\\\\").replace("\t", "\\t").replace("\n", "\\n")
    return s


def main():
    if len(sys.argv) < 3:
        print("Usage: excel_to_rows.py input.xlsx output.tsv", file=sys.stderr)
        return 2

    src, dst = sys.argv[1], sys.argv[2]
    try:
        import openpyxl
    except ImportError:
        print("openpyxl not installed", file=sys.stderr)
        return 3

    wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
    ws = wb.active

    rows = []
    for row in ws.iter_rows(values_only=True):
        if row is None:
            continue
        cells = [cell_to_text(v) for v in row]
        # trim trailing empties
        while cells and cells[-1] == "":
            cells.pop()
        if any(c != "" for c in cells):
            rows.append(cells)

    wb.close()

    if len(rows) < 2:
        print("Excel needs header + at least one data row", file=sys.stderr)
        return 4

    # pad rows to same width as header
    width = max(len(r) for r in rows)
    out_lines = []
    for r in rows:
        padded = r + [""] * (width - len(r))
        out_lines.append("\t".join(escape_tsv(c) for c in padded))

    parent = os.path.dirname(os.path.abspath(dst))
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)

    with open(dst, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(out_lines))
        f.write("\n")

    print("OK rows=%d cols=%d -> %s" % (len(rows) - 1, width, dst))
    return 0


if __name__ == "__main__":
    sys.exit(main())
