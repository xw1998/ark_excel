# -*- coding: utf-8 -*-
"""对比原始 xlsx 与 ark_excel 复刻版的差异"""
import openpyxl, sys

src = r"C:\Users\14905\Desktop\2025-03-25 11_37.xlsx（合肥智旭）.xlsx"
dst = r"D:\工作\git_project\ark_excel\examples\gen_zhixin.xlsx"

a = openpyxl.load_workbook(src)['Sheet1']
b = openpyxl.load_workbook(dst)['Sheet1']

errs = 0

def chk(name, cond):
    global errs
    if not cond:
        errs += 1
        print('DIFF:', name)

# 1. 值
va = {c.coordinate: c.value for r in a.iter_rows() for c in r if c.value is not None}
vb = {c.coordinate: c.value for r in b.iter_rows() for c in r if c.value is not None}
chk('value count %d vs %d' % (len(va), len(vb)), len(va) == len(vb))
diff_cells = [k for k in va if va[k] != vb.get(k)]
chk('values mismatch at %s' % diff_cells[:10], not diff_cells)

# 2. 合并
ma = sorted(str(m) for m in a.merged_cells.ranges)
mb = sorted(str(m) for m in b.merged_cells.ranges)
sa, sb = set(ma), set(mb)
chk('merge count %d vs %d' % (len(sa), len(sb)), len(sa) == len(sb))
chk('merge only-in-src %s' % list(sa - sb)[:5], not (sa - sb))
chk('merge only-in-dst %s' % list(sb - sa)[:5], not (sb - sa))

# 3. 列宽
wa = {c: d.width for c, d in a.column_dimensions.items() if d.width}
wb_ = {c: d.width for c, d in b.column_dimensions.items() if d.width}
for c, w in wa.items():
    if c in wb_ and abs(wb_[c] - w) > 0.01:
        errs += 1; print('DIFF colwidth', c, w, wb_[c])

# 4. 行高
ha = {r: d.height for r, d in a.row_dimensions.items() if d.height}
hb = {r: d.height for r, d in b.row_dimensions.items() if d.height}
bad = [r for r in ha if abs((hb.get(r) or 0) - ha[r]) > 0.5]
chk('row heights %s' % bad[:8], not bad)

# 5. 边框: 仅比对非合并覆盖格(openpyxl 会用左上格重建覆盖格边框, 不可比)
covered = set()
for m in a.merged_cells.ranges:
    for r in range(m.min_row, m.max_row + 1):
        for cc in range(m.min_col, m.max_col + 1):
            if (r, cc) != (m.min_row, m.min_col):
                covered.add((r, cc))
def bs(side):
    return side.style if side and side.style else None
nb = 0
for row in a.iter_rows():
    for c in row:
        if (c.row, c.column) in covered:
            continue
        ea = (bs(c.border.left), bs(c.border.right), bs(c.border.top), bs(c.border.bottom))
        if not any(ea):
            continue
        cb = b.cell(row=c.row, column=c.column)
        ebb = (bs(cb.border.left), bs(cb.border.right), bs(cb.border.top), bs(cb.border.bottom))
        if ea != ebb:
            nb += 1
            if nb <= 5:
                errs += 1
                print('DIFF border', c.coordinate, ea, ebb)
chk('border diffs=%d' % nb, nb == 0)

# 6. 字体/对齐(有值格)
nf = 0
for k, v in va.items():
    ca, cb2 = a[k], b[k]
    if (ca.font.sz, ca.font.b, ca.font.name) != (cb2.font.sz, cb2.font.b, cb2.font.name):
        nf += 1
        if nf <= 5:
            errs += 1; print('DIFF font', k, ca.font.sz, cb2.font.sz)
chk('font diffs=%d' % nf, nf == 0)

print('==== %s | cells=%d merges=%d' % ('ALL IDENTICAL' if errs == 0 else 'DIFFS=%d' % errs, len(va), len(sa)))
