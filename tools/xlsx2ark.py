# -*- coding: utf-8 -*-
"""xlsx2ark.py —— 任意 xlsx 转 ark_excel C 生成器（通用工具）

用法:
    python tools/xlsx2ark.py 输入.xlsx 输出.c [sheet名或索引, 默认0]

读取输入 xlsx 的一个工作表, 提取:
  值(文本/数字/布尔) / 字体(名,字号,粗体) / 对齐(水平,垂直,换行)
  / 逐边边框 / 合并区域 / 列宽 / 行高
并生成一份独立可编译的 C 程序(含容量覆盖 #define),
用 ark_excel 逐格复刻原表。

编译:
    gcc -std=gnu99 -I. -Isrc 输出.c src/ark_excel.c src/ark_zip.c -o 输出.exe
"""
import sys, json, zipfile
from xml.etree import ElementTree as ET
import openpyxl
from openpyxl.utils import get_column_letter

NSC = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

def existing_cells(src):
    """源 XML 中真实存在 <c> 元素的坐标集合(含纯样式格)"""
    z = zipfile.ZipFile(src)
    # 取第一个 sheet 对应部件
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rid = wb.find('.//' + NSC + 'sheet').get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    tgt = [r.get('Target') for r in rels if r.get('Id') == rid][0]
    part = 'xl/' + tgt.lstrip('/')
    root = ET.fromstring(z.read(part))
    return {c.get('r') for c in root.iter(NSC + 'c')}

BORDER_MAP = {
    'thin': 1, 'medium': 2, 'thick': 3, 'dashed': 4,
    'dotted': 5, 'double': 6, 'hair': 7,
}
HALIGN = {'left': 1, 'center': 2, 'right': 3, None: 0}
VALIGN = {None: 0, 'top': 1, 'center': 2, 'bottom': 3, 'justify': 4}  # ARK_VALIGN_NONE=0,TOP=1,CENTER=2,BOTTOM=3


def c_str(s):
    out = []
    for ch in s:
        b = ch.encode('utf-8')
        if ch == '"':
            out.append('\\"')
        elif ch == '\\':
            out.append('\\\\')
        elif ch == '\n':
            out.append('\\n')
        elif ch == '\r':
            pass
        elif ch == '\t':
            out.append('\\t')
        elif 32 <= ord(ch) < 127:
            out.append(ch)
        else:
            for byte in b:
                out.append('\\x%02x' % byte)
    # 避免 \xNN 后紧跟十六进制字符被吞逃长(hex escape out of range)
    res = []
    hexdig = set('0123456789abcdefABCDEF')
    pending_hex = False
    for ch in out:
        if pending_hex and ch in hexdig:
            res.append('""')  # 拼接断开转义序列
            pending_hex = False
        if ch.startswith('\\x'):
            pending_hex = True
        else:
            pending_hex = False
        res.append(ch)
    return '"' + ''.join(res) + '"'


def bs(side):
    return BORDER_MAP.get(side.style, 0) if (side and side.style) else 0


def main():
    src, dst = sys.argv[1], sys.argv[2]
    sheet_sel = sys.argv[3] if len(sys.argv) > 3 else '0'
    wb = openpyxl.load_workbook(src, data_only=False)
    ws = wb[int(sheet_sel)] if sheet_sel.isdigit() else wb[sheet_sel]
    src_cell_set = existing_cells(src)

    # ---------- 提取 ----------
    cells = []      # (row, col, value_or_None, stylekey)
    styles = {}     # stylekey -> id
    fonts = set()
    for row in ws.iter_rows():
        for c in row:
            b = (bs(c.border.left), bs(c.border.right), bs(c.border.top), bs(c.border.bottom))
            v = c.value
            coord = '%s%d' % (get_column_letter(c.column), c.row)
            has_c = coord in src_cell_set
            if v is None and not any(b) and not has_c:
                continue
            fname = c.font.name or ''
            fsize = int(round(float(c.font.sz or 11) * 100))
            bold = 1 if c.font.b else 0
            ha = HALIGN.get(c.alignment.horizontal, 0)
            va = VALIGN.get(c.alignment.vertical, 0)
            wrap = 1 if c.alignment.wrap_text else 0
            dl = BORDER_MAP.get(c.border.diagonal.style, 0) if (c.border.diagonal and c.border.diagonal.style) else 0
            dup = 1 if c.border.diagonalUp else 0
            ddown = 1 if c.border.diagonalDown else 0
            nfid = c._style.numFmtId if hasattr(c, '_style') and c._style is not None else 0
            rot = int(c.alignment.textRotation or 0)
            key = (fname, fsize, bold, ha, va, wrap) + b + (dl, dup, ddown, nfid, rot)
            styles[key] = None
            fonts.add(fname)
            cells.append((c.row, c.column, v, key))

    style_keys = list(styles.keys())
    merges = [(m.min_row, m.min_col, m.max_row, m.max_col) for m in ws.merged_cells.ranges]
    colw = [(cd.min or idx, get_column_letter(idx), cd.width)
            for idx, cd in sorted(ws.column_dimensions.items())
            if cd.width] if False else None
    # 列宽: 按 min..max 展开为逐列
    colw = {}
    colhidden = set()
    colstyle = {}   # ci -> styledef key(用源 xf 索引映射到生成的 sid)
    for idx, cd in ws.column_dimensions.items():
        lo = cd.min or idx
        hi = cd.max or idx
        for ci in range(lo, hi + 1):
            if cd.width:
                colw[ci] = cd.width
            if cd.hidden:
                colhidden.add(ci)
            if getattr(cd, 'style', None) and cd.style:
                colstyle[ci] = cd.style  # 源 cellXf 索引, 后面映射
    rowh = {r: rd.height for r, rd in ws.row_dimensions.items() if rd.height}
    rowhidden = {r for r, rd in ws.row_dimensions.items() if rd.hidden}
    sheet_name = ws.title

    # 解析源 styles.xml: 列样式 xf 的 font/alignment 属性
    colstyle_defs = {}
    if colstyle:
        import zipfile as _zf
        from xml.etree import ElementTree as _ET
        NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
        _z = _zf.ZipFile(src)
        _st = _ET.fromstring(_z.read('xl/styles.xml'))
        _fonts = _st.find(NS + 'fonts')
        _fl = []
        for _f in _fonts:
            _sz = _f.find(NS + 'sz'); _nm = _f.find(NS + 'name')
            _fl.append((_nm.get('val') if _nm is not None else '',
                        int(round(float(_sz.get('val')) * 100)) if _sz is not None else 1100))
        _xfs = _st.find(NS + 'cellXfs')
        for ci, xfidx in colstyle.items():
            _xf = list(_xfs)[xfidx]
            _al = _xf.find(NS + 'alignment')
            va = _al.get('vertical') if _al is not None else None
            fname, fsize = _fl[int(_xf.get('fontId', 0))]
            key = (fname, fsize, 0, 0, VALIGN.get(va, 0), 0) + (0,0,0,0) + (0,0,0,0, 0)
            styles[key] = None
            colstyle_defs[ci] = key

    style_keys = list(styles.keys())
    for i, k in enumerate(style_keys):
        styles[k] = i

    # ---------- 生成容量覆盖头(编译时 -include 到所有翻译单元!) ----------
    over = dst.rsplit('.', 1)[0] + '_over.h'
    OV = []
    O = OV.append
    O('/* 由 tools/xlsx2ark.py 自动生成的容量覆盖, 必须对所有源文件统一使用:')
    O('   gcc ... -include %s ... */' % over.rsplit('\\', 1)[-1].rsplit('/', 1)[-1])
    O('#define ARK_MAX_CELLS   %d' % (len(cells) + 16))
    O('#define ARK_MAX_MERGES  %d' % (len(merges) + 8))
    O('#define ARK_MAX_STYLES  %d' % (len(style_keys) + 24))
    O('#define ARK_MAX_FONTS   %d' % (len(style_keys) + 24))
    O('#define ARK_MAX_BORDERS %d' % (len(style_keys) * 2 + 16))
    O('#define ARK_MAX_FILLS   %d' % (len(style_keys) + 24))
    O('#define ARK_MAX_COLDEFS %d' % (len(colw) + 4))
    O('#define ARK_MAX_ROWDEFS %d' % (len(rowh) + 4))
    O('#define ARK_MAX_STR     %d' % (sum(1 for c in cells if isinstance(c[2], str)) + 16))
    O('#define ARK_STR_ARENA   %d' % (sum(len(str(c[2])) for c in cells if c[2] is not None) * 3 + 1024))
    O('#define ARK_SST_ARENA  ARK_STR_ARENA')
    maxlen = max([len(str(c[2])) * 3 + 16 for c in cells if c[2] is not None] or [96])
    O('#define ARK_MAX_STR_LEN %d' % maxlen)
    O('#define ARK_MAX_SST     %d' % (sum(1 for c in cells if isinstance(c[2], str)) + 16))
    open(over, 'w', encoding='utf-8', newline='\n').write('\n'.join(OV) + '\n')

    # ---------- 生成 C ----------
    L = []
    A = L.append
    A('/* 由 tools/xlsx2ark.py 自动生成, 源: %s */' % src.split('\\')[-1].split('/')[-1])
    A('#include "%s"' % over.rsplit('\\', 1)[-1].rsplit('/', 1)[-1])
    A('#include <stdio.h>')
    A('#include "ark_excel.h"')
    A('')
    A('static const char F_0[] = %s;' % (c_str(sorted(fonts)[0]) if fonts else '"Calibri"'))
    A('')
    A('static void build(ark_wb_t *wb)')
    A('{')
    A('    int sh = ark_wb_add_sheet(wb, %s);' % c_str(sheet_name))
    # 样式注册
    A('    static const ark_style_def_t defs[] = {')
    for k in style_keys:
        fname, fsize, bold, ha, va, wrap, bl, br, bt, bb, dl, dup, ddown, nfid, rot = k
        A('        { F_0, %d, %d, 0, 0, 0, 0, 0, %d, %d, %d, %d, 0, %d, %d, %d, %d, %d, %d, 0, %d, %d, 0 },'
          % (fsize, bold, bl, br, bt, bb, dl, dup, ddown, ha, va, wrap, rot, nfid))
    A('    };')
    A('    int sid[%d];' % len(style_keys))
    A('    for (unsigned i = 0; i < sizeof defs / sizeof defs[0]; i++)')
    A('        sid[i] = ark_wb_add_style(wb, &defs[i]);')
    A('    (void)sid;')
    # 列宽 / 行高
    # 合并连续同宽且同隐藏属性的列为范围, 与源文件的 <col min max> 一致
    col_ranges = []
    for ci in sorted(set(colw) | colhidden | set(colstyle_defs)):
        w = colw.get(ci)
        hd = ci in colhidden
        ck = colstyle_defs.get(ci)
        if col_ranges and col_ranges[-1][1] == ci - 1 and col_ranges[-1][2] == w \
           and col_ranges[-1][3] == hd and col_ranges[-1][4] == ck:
            col_ranges[-1][1] = ci
        else:
            col_ranges.append([ci, ci, w, hd, ck])
    for lo, hi, w, hd, ck in col_ranges:
        if w is not None:
            if lo == 1:
                _st = z.read('xl/styles.xml').decode('utf-8')
                _f0 = re.search(r'<font>.*?</font>|<font/>', re.search(r'<fonts[^>]*>(.*?)</fonts>', _st, re.S).group(1), re.S).group(0)
                _nm = re.search(r'<name val="([^"]+)"', _f0)
                _sz = re.search(r'<sz val="([\d.]+)"', _f0)
                A('    ark_set_default_font("%s", %d);' % (_nm.group(1) if _nm else 'Calibri', int(float(_sz.group(1)) * 100) if _sz else 1100))
            A('    ark_set_col_width(wb, sh, %d, %d, %.9g);' % (lo, hi, w))
        if hd:
            A('    ark_set_col_hidden(wb, sh, %d, %d, 1);' % (lo, hi))
        if ck is not None:
            A('    ark_set_col_style(wb, sh, %d, %d, sid[%d]);' % (lo, hi, styles[ck]))
    for r in sorted(set(rowh) | rowhidden):
        if r in rowh:
            A('    ark_set_row_height(wb, sh, %d, %.9g);' % (r, rowh[r]))
        if r in rowhidden:
            A('    ark_set_row_hidden(wb, sh, %d, 1);' % r)
    # 单元格
    for r, c, v, key in cells:
        sid = styles[key]
        if v is None:
            A('    ark_set_blank(wb, sh, %d, %d, sid[%d]);' % (r, c, sid))
        elif isinstance(v, bool):
            A('    ark_set_bool(wb, sh, %d, %d, %d, sid[%d]);' % (r, c, int(v), sid))
        elif isinstance(v, (int, float)):
            if float(v).is_integer() and abs(v) < 2**31:
                A('    ark_set_int(wb, sh, %d, %d, %d, sid[%d]);' % (r, c, int(v), sid))
            else:
                A('    ark_set_num(wb, sh, %d, %d, %s, sid[%d]);' % (r, c, repr(float(v)), sid))
        else:
            s = str(v)
            if s.startswith('='):
                A('    ark_set_formula(wb, sh, %d, %d, %s, sid[%d]);' % (r, c, c_str(s[1:]), sid))
            else:
                A('    ark_set_str(wb, sh, %d, %d, %s, sid[%d]);' % (r, c, c_str(s), sid))
    # 合并
    for r1, c1, r2, c2 in merges:
        A('    ark_merge(wb, sh, %d, %d, %d, %d);' % (r1, c1, r2, c2))
    A('}')
    A('')
    A('static int fw(const void *d, size_t n, void *ctx)')
    A('{ return fwrite(d, 1, n, (FILE *)ctx) == n ? 0 : -1; }')
    A('')
    A('int main(void)')
    A('{')
    A('    static ark_wb_t wb;')
    out_xlsx = dst.rsplit('\\', 1)[-1].rsplit('/', 1)[-1].rsplit('.', 1)[0] + '.xlsx'
    A('    FILE *f = fopen("%s", "wb");' % out_xlsx)
    A('    if (!f) return 1;')
    A('    ark_wb_init(&wb, fw, f);')
    A('    build(&wb);')
    A('    ark_err_t e = ark_wb_save(&wb);')
    A('    fclose(f);')
    A('    printf("save: %d\\n", (int)e);')
    A('    return e != 0;')
    A('}')

    open(dst, 'w', encoding='utf-8', newline='\n').write('\n'.join(L) + '\n')
    print('generated %s: cells=%d styles=%d merges=%d colw=%d rowh=%d'
          % (dst, len(cells), len(style_keys), len(merges), len(colw), len(rowh)))


if __name__ == '__main__':
    main()
