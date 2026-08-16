#include "ark_excel_config.h"
#if ARK_ENABLE_READWRITE
/*!
 * @file    ark_reader.c
 * @brief   ark_excel 读改写模块实现 —— 打开/载入现有 xlsx
 *
 * 组成:
 *   1. ZIP 读取层: EOCD 定位 + 中央目录解析 + 条目解压(Store/Deflate)
 *   2. 微型 XML 工具: 标签遍历 / 属性提取(零拷贝, 直接在解压缓冲上)
 *   3. xlsx 载入: workbook/rels/SST/styles/worksheet -> ark_wb_t
 */
#include <string.h>
#include <stdlib.h>
#include <stdio.h>
#include "ark_excel_internal.h"
#include "ark_inflate.h"

/* ============================================================
 * 1. 读取回调上的小工具
 * ========================================================== */
static uint32_t rd_read(ark_reader_t *rd, uint32_t pos, uint8_t *buf, uint32_t len)
{
    if (rd->io.seek(rd->io.ctx, pos) != 0) return 0;
    return rd->io.read(rd->io.ctx, buf, len);
}

static uint16_t rd_u16(const uint8_t *p) { return (uint16_t)(p[0] | (p[1] << 8)); }
static uint32_t rd_u32(const uint8_t *p)
{
    return (uint32_t)p[0] | ((uint32_t)p[1] << 8) |
           ((uint32_t)p[2] << 16) | ((uint32_t)p[3] << 24);
}

/* ============================================================
 * 2. ZIP 中央目录解析
 * ========================================================== */
static int zip_find_eocd(ark_reader_t *rd, uint32_t *cd_off, uint16_t *count)
{
    /* 从文件尾部向前按块搜索 EOCD 签名 0x06054b50 */
    uint32_t search = rd->file_size > 66000u ? 66000u : rd->file_size;
    uint32_t base = rd->file_size - search;
    while (search) {
        uint32_t chunk = search > (uint32_t)sizeof(rd->buf)
                       ? (uint32_t)sizeof(rd->buf) : search;
        uint32_t off = base + search - chunk;
        uint32_t got = rd_read(rd, off, rd->buf, chunk);
        if (got < 4) return -1;
        for (int32_t i = (int32_t)got - 22; i >= 0; i--) {
            if (rd_u32(rd->buf + i) == 0x06054B50u) {
                *count = rd_u16(rd->buf + i + 10);
                *cd_off = rd_u32(rd->buf + i + 16);
                return 0;
            }
        }
        search -= chunk;
        if (search) search += 3;    /* 跨块边界回退 3 字节 */
    }
    return -1;
}

static int zip_parse_central(ark_reader_t *rd)
{
    uint32_t pos, end;
    uint8_t hdr[46];
    uint16_t i;
    uint32_t cd_off; uint16_t cnt;
    if (zip_find_eocd(rd, &cd_off, &cnt) != 0) return -1;
    if (cnt > ARK_MAX_ZIP_ENTRIES) return -2;
    rd->entry_count = cnt;
    pos = cd_off;
    for (i = 0; i < cnt; i++) {
        ark_zentry_t *e = &rd->entries[i];
        if (rd_read(rd, pos, hdr, 46) != 46) return -1;
        if (rd_u32(hdr) != 0x02014B50u) return -1;
        e->method    = rd_u16(hdr + 10);
        e->crc       = rd_u32(hdr + 16);
        e->csize     = rd_u32(hdr + 20);
        e->usize     = rd_u32(hdr + 24);
        {   /* 条目名(限制在 63 字节内, 超长视为不支持) */
            uint16_t nlen = rd_u16(hdr + 28);
            if (nlen == 0 || nlen >= sizeof(e->name)) return -3;
            if (rd_read(rd, pos + 46, rd->buf, nlen) != nlen) return -1;
            memcpy(e->name, rd->buf, nlen);
            e->name[nlen] = 0;
        }
        e->local_off = rd_u32(hdr + 42);
        pos += 46 + rd_u16(hdr + 28) + rd_u16(hdr + 30) + rd_u16(hdr + 32);
    }
    (void)end;
    return 0;
}

/* ---- 条目数据解压到 rd->buf(整体载入, 上限 ARK_OPEN_XML_BUF) ---- */
typedef struct {
    ark_reader_t *rd;
    uint32_t pos, remain;
} ent_src_t;

static uint32_t ent_pull(void *ctx, uint8_t *buf, uint32_t want)
{
    ent_src_t *s = (ent_src_t *)ctx;
    if (want > s->remain) want = s->remain;
    if (!want) return 0;
    {   /* seek 只在方向变化时需要; 顺序读仍显式 seek 保证正确 */
        uint32_t n = rd_read(s->rd, s->pos, buf, want);
        s->pos += n; s->remain -= n;
        return n;
    }
}

/*! 定位条目数据区(解析 local header), 返回 0 成功 */
static int entry_data_pos(ark_reader_t *rd, const ark_zentry_t *e, uint32_t *dpos)
{
    uint8_t lh[30];
    if (rd_read(rd, e->local_off, lh, 30) != 30) return -1;
    if (rd_u32(lh) != 0x04034B50u) return -1;
    *dpos = e->local_off + 30 + rd_u16(lh + 26) + rd_u16(lh + 28);
    return 0;
}

static int entry_extract(ark_reader_t *rd, const ark_zentry_t *e)
{
    uint32_t dpos;
    if (entry_data_pos(rd, e, &dpos) != 0) return -1;
    if (e->usize > (uint32_t)sizeof(rd->buf)) return -2;   /* 超出缓冲 */
    if (e->method == 0) {   /* Store */
        rd->buf_len = rd_read(rd, dpos, rd->buf, e->usize);
        return rd->buf_len == e->usize ? 0 : -1;
    } else if (e->method == 8) {   /* Deflate */
#if !ARK_ENABLE_INFLATE
        return -1;                  /* 解压器已裁剪, 仅支持 Store 模板 */
#else
        ent_src_t src = { rd, dpos, e->csize };
        size_t outn = 0;
        int r = ark_inflate(ent_pull, &src, rd->buf, sizeof(rd->buf), &outn);
        if (r != 0) return -1;
        rd->buf_len = (uint32_t)outn;
        return 0;
#endif
    }
    return -1;   /* 不支持的压缩方法 */
}

int ark_i_entry_extract(ark_reader_t *rd, uint16_t idx)
{
    return entry_extract(rd, &rd->entries[idx]);
}

static int entry_find(ark_reader_t *rd, const char *name)
{
    for (uint16_t i = 0; i < rd->entry_count; i++)
        if (strcmp(rd->entries[i].name, name) == 0) return i;
    return -1;
}

/* ============================================================
 * 3. 微型 XML 工具(在 rd->buf 上, NUL 终结化后使用)
 * ========================================================== */
static const char *xml_find(const char *s, const char *tag)
{ return strstr(s, tag); }

static const char *xml_find_next(const char *s, const char *tag, const char *from)
{
    if (from <= s) return NULL;
    const char *r = strstr(from, tag);
    return r;
}

/*! 提取属性值: 返回值指针写入 *len; 找不到返回 NULL */
static const char *xml_attr(const char *tagpos, const char *attr, int *len)
{
    size_t alen = strlen(attr);
    const char *p = tagpos;
    while (*p && *p != '>') {
        if (p[0] == ' ' || p[0] == '\t' || p[0] == '\n' || p[0] == '\r') {
            const char *q = ++p;
            while (*q && *q != '=' && *q != ' ' && *q != '>') q++;
            if ((size_t)(q - p) == alen && strncmp(p, attr, alen) == 0 && *q == '=') {
                const char *v = q + 1;
                if (*v == '"' || *v == '\'') {
                    const char *e = strchr(v + 1, *v);
                    if (!e) return NULL;
                    *len = (int)(e - v - 1);
                    return v + 1;
                }
            }
            p = q;
        } else p++;
    }
    return NULL;
}

/*! 拷贝属性值到 buf(解码 &amp; &lt; &gt; &quot;), 返回长度 */
static int xml_copy_decoded(const char *v, int len, char *buf, size_t cap)
{
    size_t o = 0;
    for (int i = 0; i < len && o + 1 < cap; i++) {
        if (v[i] == '&') {
            if (v[i+1] == '#') {   /* 数字字符引用 &#NNNN; 或 &#xHHHH; */
                long cp = -1; int adv = 0;
                const char *q = v + i + 2;
                if (q < v + len && (*q == 'x' || *q == 'X')) {
                    q++; cp = 0;
                    while (q < v + len && ((*q >= '0' && *q <= '9') ||
                           (*q >= 'a' && *q <= 'f') || (*q >= 'A' && *q <= 'F'))) {
                        cp = cp * 16 + (*q <= '9' ? *q - '0' :
                                        (*q | 32) - 'a' + 10);
                        q++; adv++;
                    }
                    if (!(q < v + len && *q == ';')) cp = -1;
                    else adv += 3;
                } else {
                    cp = 0;
                    while (q < v + len && *q >= '0' && *q <= '9') {
                        cp = cp * 10 + (*q - '0'); q++; adv++;
                    }
                    if (!(q < v + len && *q == ';')) cp = -1;
                    else adv += 2;
                }
                if (cp >= 0 && cp < 0x110000 && adv > 2) {
                    /* UTF-8 编码 */
                    if (cp < 0x80) { if (o + 1 < cap) buf[o++] = (char)cp; }
                    else if (cp < 0x800) {
                        if (o + 2 < cap) {
                            buf[o++] = (char)(0xC0 | (cp >> 6));
                            buf[o++] = (char)(0x80 | (cp & 0x3F)); }
                    } else if (cp < 0x10000) {
                        if (o + 3 < cap) {
                            buf[o++] = (char)(0xE0 | (cp >> 12));
                            buf[o++] = (char)(0x80 | ((cp >> 6) & 0x3F));
                            buf[o++] = (char)(0x80 | (cp & 0x3F)); }
                    } else {
                        if (o + 4 < cap) {
                            buf[o++] = (char)(0xF0 | (cp >> 18));
                            buf[o++] = (char)(0x80 | ((cp >> 12) & 0x3F));
                            buf[o++] = (char)(0x80 | ((cp >> 6) & 0x3F));
                            buf[o++] = (char)(0x80 | (cp & 0x3F)); }
                    }
                    i += adv;
                    continue;
                }
            }
            if (len - i >= 5 && strncmp(v + i, "&amp;", 5) == 0) { buf[o++]='&'; i+=4; }
            else if (len - i >= 4 && strncmp(v + i, "&lt;", 4) == 0) { buf[o++]='<'; i+=3; }
            else if (len - i >= 4 && strncmp(v + i, "&gt;", 4) == 0) { buf[o++]='>'; i+=3; }
            else if (len - i >= 6 && strncmp(v + i, "&quot;", 6) == 0) { buf[o++]='"'; i+=5; }
            else if (len - i >= 6 && strncmp(v + i, "&apos;", 6) == 0) { buf[o++]='\''; i+=5; }
            else buf[o++] = v[i];
        } else buf[o++] = v[i];
    }
    buf[o] = 0;
    return (int)o;
}

static double xml_atof(const char *v, int len)
{
    char tmp[40];
    if (len >= (int)sizeof(tmp)) len = (int)sizeof(tmp) - 1;
    memcpy(tmp, v, (size_t)len); tmp[len] = 0;
    return strtod(tmp, NULL);
}

static long xml_atoi(const char *v, int len)
{
    char tmp[24];
    if (len >= (int)sizeof(tmp)) len = (int)sizeof(tmp) - 1;
    memcpy(tmp, v, (size_t)len); tmp[len] = 0;
    return strtol(tmp, NULL, 10);
}

/*! "A1" -> col; 行号写入 *row */
static int ref_to_rc(const char *v, int len, uint32_t *row, uint16_t *col)
{
    uint16_t c = 0; uint32_t r = 0; int i = 0;
    for (; i < len && v[i] >= 'A' && v[i] <= 'Z'; i++)
        c = (uint16_t)(c * 26 + (v[i] - 'A' + 1));
    for (; i < len && v[i] >= '0' && v[i] <= '9'; i++)
        r = r * 10 + (v[i] - '0');
    if (i != len || c == 0 || r == 0) return -1;
    *row = r; *col = c;
    return 0;
}

/* ============================================================
 * 4. 载入各部件
 * ========================================================== */
static void rd_load_rels(ark_reader_t *rd, const char *xml,
                         char (*targets)[64], int max_rid)
{
    /* targets[rid-1] = Target; workbook.xml.rels 格式 */
    const char *p = xml;
    memset(targets, 0, (size_t)max_rid * 64);
    while ((p = xml_find(p, "<Relationship "))) {
        int lid = 0, lt = 0, ltp = 0;
        const char *id = xml_attr(p, "Id", &lid);
        const char *tg = xml_attr(p, "Target", &lt);
        const char *tp = xml_attr(p, "TargetMode", &ltp);
        if (id && tg && lid > 3 && strncmp(id, "rId", 3) == 0) {
            long rid = strtol(id + 3, NULL, 10);
            if (rid >= 1 && rid <= max_rid && (size_t)lt < 63) {
                if (!tp)   /* 只要内部部件 */
                    xml_copy_decoded(tg, lt, targets[rid - 1], 64);
            }
        }
        p += 15;
    }

}

/*! 解析 workbook.xml: 得到 sheet 名与 rId; 解析 rels 得到部件路径 */
static ark_err_t rd_load_workbook(ark_reader_t *rd, ark_wb_t *wb)
{
    int i;
    char (*targets)[64];
    if (entry_extract(rd, &rd->entries[rd->i_workbook]) != 0)
        return ARK_ERR_BAD_STATE;
    rd->buf[rd->buf_len] = 0;

    { static char s_targets[ARK_MAX_SHEETS][64]; targets = s_targets; }

    /* 第一遍: 数 sheet 并取 name + rId */
    {
        const char *p = (const char *)rd->buf;
        int sidx = 0;
        static char s_rids[ARK_MAX_SHEETS][16];
        char (*rids)[16] = s_rids;
        memset(rids, 0, sizeof(s_rids));
        while ((p = xml_find(p, "<sheet "))) {
            int ln = 0, li = 0;
            const char *name = xml_attr(p, "name", &ln);
            const char *id   = xml_attr(p, "r:id", &li);
            if (name && id && sidx < ARK_MAX_SHEETS) {
                char nm[ARK_SHEET_NAME_MAX];
                xml_copy_decoded(name, ln, nm, sizeof(nm));
                if (ark_wb_add_sheet(wb, nm) != sidx)
                    return ARK_ERR_TOO_MANY_SHEETS;
                if (li < 15) { memcpy(rids[sidx], id, (size_t)li); rids[sidx][li] = 0; }
                sidx++;
            }
            p += 7;
        }
        /* 第二遍: rels -> 部件路径 */
        if (rd->i_rels != 0xFFFF) {
            if (entry_extract(rd, &rd->entries[rd->i_rels]) != 0)
                return ARK_ERR_BAD_STATE;
            rd->buf[rd->buf_len] = 0;
            rd_load_rels(rd, (const char *)rd->buf, targets, ARK_MAX_SHEETS);
        }
        for (i = 0; i < sidx; i++) {
            const char *t = rids[i][0] ? targets[strtol(rids[i] + 3, NULL, 10) - 1] : NULL;
            if (!t || !t[0]) return ARK_ERR_BAD_STATE;
            /* 归一: "worksheets/sheet1.xml" -> "xl/worksheets/sheet1.xml" */
            if (t[0] == '/') snprintf(rd->sheet_part[i], 64, "%s", t + 1);
            else if (strncmp(t, "xl/", 3) == 0) snprintf(rd->sheet_part[i], 64, "%s", t);
            else snprintf(rd->sheet_part[i], 64, "xl/%s", t);
            rd->sheet_entry[i] = (uint16_t)entry_find(rd, rd->sheet_part[i]);
        }
    }
    return ARK_OK;
}

/*! 载入 SST */
static ark_err_t rd_load_sst(ark_reader_t *rd, ark_wb_t *wb)
{
    const char *p;
    char text[ARK_MAX_STR_LEN];
    if (rd->i_sst == 0xFFFF) return ARK_OK;
    if (entry_extract(rd, &rd->entries[rd->i_sst]) != 0) return ARK_ERR_BAD_STATE;
    rd->buf[rd->buf_len] = 0;
    p = (const char *)rd->buf;
    while ((p = xml_find(p, "<si>"))) {
        /* 简单 si: <si><t>text</t></si>; 富文本: 多个 <r><t>..</t></r> 拼接 */
        const char *end = xml_find(p, "</si>");
        const char *q = p + 4;
        size_t o = 0;
        if (!end) break;
        while (q < end && o + 1 < sizeof(text)) {
            const char *t = strstr(q, "<t");
            if (!t || t >= end) break;
            const char *v = strchr(t, '>');
            if (!v) break;
            v++;
            {   const char *e = strchr(v, '<');
                    if (!e || e > end) e = end;
                int n = (int)(e - v);
                char piece[ARK_MAX_STR_LEN];
                xml_copy_decoded(v, n, piece, sizeof(piece));
                size_t pl = strlen(piece);
                if (o + pl >= sizeof(text)) pl = sizeof(text) - 1 - o;
                memcpy(text + o, piece, pl); o += pl;
                q = e;
            }
        }
        text[o] = 0;
        {   /* 载入不去重(保留模板顺序, 引用序号才能对上) */
            int r = ark_str_add(wb, text);
            if (r >= 0) {
                wb->strs[r].is_sst = 1;
                wb->strs[r].sst_seq = wb->sst_count++;
            }
        }
        p = end + 4;
    }
    return ARK_OK;
}

/*! 载入 styles 计数 */
static ark_err_t rd_load_styles(ark_reader_t *rd)
{
    const char *p;
    if (rd->i_styles == 0xFFFF) return ARK_OK;
    if (entry_extract(rd, &rd->entries[rd->i_styles]) != 0) return ARK_ERR_BAD_STATE;
    rd->buf[rd->buf_len] = 0;
    p = (const char *)rd->buf;
    {
        static const char *tabs[4] = { "<fonts ", "<fills ", "<borders ", "<cellXfs " };
        uint16_t *dst[4] = { &rd->tpl_font_cnt, &rd->tpl_fill_cnt,
                             &rd->tpl_border_cnt, &rd->tpl_xf_cnt };
        for (int i = 0; i < 4; i++) {
            const char *t = xml_find(p, tabs[i]);
            if (t) { int ln; const char *v = xml_attr(t, "count", &ln);
                     *dst[i] = v ? (uint16_t)xml_atoi(v, ln) : 0; }
        }
    }
    /* numFmt: 条数 + 最大自定义 id */
    {
        const char *t = xml_find(p, "<numFmts ");
        rd->tpl_numfmt_cnt = 0;
        rd->tpl_numfmt_maxid = 164;
        if (t) {
            int ln2; const char *cv = xml_attr(t, "count", &ln2);
            rd->tpl_numfmt_cnt = cv ? (uint16_t)xml_atoi(cv, ln2) : 0;
            const char *q = t;
            while ((q = xml_find(q, "<numFmt "))) {
                int ln; const char *v = xml_attr(q, "numFmtId", &ln);
                long id = v ? xml_atoi(v, ln) : 0;
                if (id >= 164 && (uint16_t)id > rd->tpl_numfmt_maxid)
                    rd->tpl_numfmt_maxid = (uint16_t)id;
                q += 8;
            }
        }
    }
    /* dxfs 条数 */
    {
        const char *t = xml_find(p, "<dxfs");
        if (t) {
            int ln; const char *v = xml_attr(t, "count", &ln);
            rd->tpl_dxf_cnt = v ? (uint16_t)xml_atoi(v, ln) : 0;
        }
    }
    return ARK_OK;
}

/* cell_put 的裸版本(不做样式校验/touch, 供载入用) —— ark_excel.c 提供 */
extern ark_err_t ark_i_cell_put_raw(ark_wb_t *wb, int sheet, uint32_t row,
                                    uint16_t col, uint8_t type, uint16_t style,
                                    uint32_t sidx, double num);
static ark_err_t cell_put_raw(ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                              uint8_t type, uint16_t style, uint32_t sidx, double num)
{
    return ark_i_cell_put_raw(wb, sheet, row, col, type, style, sidx, num);
}

/*! 载入单个 worksheet */
static ark_err_t rd_load_sheet(ark_reader_t *rd, ark_wb_t *wb, uint8_t si)
{
    ark_sheet_t *sh = &wb->sheets[si];
    const char *p;
    if (rd->sheet_entry[si] == 0xFFFF) return ARK_ERR_BAD_STATE;
    if (entry_extract(rd, &rd->entries[rd->sheet_entry[si]]) != 0)
        return ARK_ERR_BAD_STATE;
    rd->buf[rd->buf_len] = 0;
    p = (const char *)rd->buf;

    /* tabColor */
    {   const char *t = xml_find(p, "<tabColor ");
        if (t) { int ln; const char *v = xml_attr(t, "rgb", &ln);
                 if (v && ln == 8) {
                     sh->tab_color = (uint32_t)strtoul(v, NULL, 16);
                     sh->tab_color_set = 1; } }
    }
    /* 冻结窗格 */
    {   const char *t = xml_find(p, "<pane ");
        if (t) { int ln;
            const char *xs = xml_attr(t, "xSplit", &ln);
            const char *ys = xml_attr(t, "ySplit", &ln);
            if (xs) sh->freeze_cols = (uint8_t)xml_atoi(xs, ln);
            if (ys) sh->freeze_rows = (uint8_t)xml_atoi(ys, ln);
        }
    }
    /* 自动筛选 */
    {   const char *t = xml_find(p, "<autoFilter ");
        if (t) { int ln; const char *v = xml_attr(t, "ref", &ln);
            if (v) { char ref[32]; xml_copy_decoded(v, ln, ref, sizeof(ref));
                char *c = strchr(ref, ':');
                if (c) { *c = 0;
                    uint32_t r1, r2; uint16_t c1, c2;
                    if (ref_to_rc(ref, (int)strlen(ref), &r1, &c1) == 0 &&
                        ref_to_rc(c + 1, (int)strlen(c + 1), &r2, &c2) == 0) {
                        sh->filter_r1 = (uint16_t)r1; sh->filter_c1 = c1;
                        sh->filter_r2 = (uint16_t)r2; sh->filter_c2 = c2;
                        sh->autofilter_set = 1;
                    }
                }
            }
        }
    }
    /* 列宽/隐藏 */
    {   const char *t = xml_find(p, "<cols>");
        if (t) { const char *q = t;
            const char *colend = xml_find(t, "</cols>");
            while ((q = xml_find(q, "<col ")) && (!colend || q < colend)) {
                int ln; uint32_t mn = 1, mx = 1;
                const char *v;
                v = xml_attr(q, "min", &ln); if (v) mn = (uint32_t)xml_atoi(v, ln);
                v = xml_attr(q, "max", &ln); if (v) mx = (uint32_t)xml_atoi(v, ln);
                v = xml_attr(q, "width", &ln);
                double w = v ? xml_atof(v, ln) : 0;
                v = xml_attr(q, "hidden", &ln);
                int hid = v && v[0] == '1';
                if (mn >= 1 && mn <= 16384 && mx >= mn && mx - mn < 64) {
                    for (uint32_t c = mn; c <= mx; c++) {
                        if (w > 0) ark_set_col_width(wb, si, (int)c, (int)c, w);
                        if (hid)   ark_set_col_hidden(wb, si, (int)c, (int)c, 1);
                    }
                }
                q += 5;
            }
        }
    }
    /* 行高/隐藏 + 单元格 */
    {   const char *q = p;
        while ((q = xml_find(q, "<row "))) {
            int ln;
            const char *v = xml_attr(q, "r", &ln);
            uint32_t rowno = v ? (uint32_t)xml_atoi(v, ln) : 0;
            v = xml_attr(q, "hidden", &ln);
            int rhid = v && v[0] == '1';
            v = xml_attr(q, "ht", &ln);
            double ht = v ? xml_atof(v, ln) : 0;
            if (rowno && (rhid || ht > 0)) {
                if (rhid) ark_set_row_hidden(wb, si, rowno, 1);
                if (ht > 0) ark_set_row_height(wb, si, rowno, ht);
            }
            /* 该行内的 <c> 单元格 */
            {   const char *ce = strstr(q, "</row>");
                const char *cq = q;
                if (!ce) ce = (const char *)rd->buf + rd->buf_len;
                while ((cq = xml_find(cq, "<c ")) && cq < ce) {
                    int lr = 0, ls = 0, lt = 0;
                    const char *ref = xml_attr(cq, "r", &lr);
                    const char *sty = xml_attr(cq, "s", &ls);
                    const char *typ = xml_attr(cq, "t", &lt);
                    const char *ve = strstr(cq, "</c>");
                    uint32_t row; uint16_t col;
                    if (ref && ref_to_rc(ref, lr, &row, &col) == 0) {
                        uint16_t style = sty ? (uint16_t)xml_atoi(sty, ls) : 0;
                        /* 值/公式 */
                        const char *fs = ve && strstr(cq, "<f") && strstr(cq, "<f") < ve
                                         ? strstr(cq, "<f") : NULL;
                        /* <v> 须在本单元格范围内 */
                        const char *vs = NULL;
                        {   const char *cand = strstr(cq, "<v");
                            if (cand && ve && cand < ve) vs = cand;
                        }
                        /* inlineStr: 文本在 <is><t>..</t>, 与 <v> 互斥 */
                        const char *istag = NULL;
                        {   const char *cand = strstr(cq, "<is>");
                            if (cand && ve && cand < ve) istag = cand;
                        }
                        char tcode = typ ? typ[0] : 0;
                        /* "inlineStr"/"str" 归一化为首个字母语义 */
                        if (tcode == 'i' && !istag) tcode = 0;

                        if (fs) {   /* 公式单元格(忽略缓存值) */
                            const char *fv = strchr(fs, '>') + 1;
                            const char *fe = strchr(fv, '<');
                            char formula[ARK_MAX_STR_LEN];
                            xml_copy_decoded(fv, (int)(fe - fv), formula, sizeof(formula));
                            if (formula[0])
                                ark_set_formula(wb, si, row, col, formula, style);
                        } else if (istag) {   /* t="inlineStr" (openpyxl 常用) */
                            const char *tv = strstr(istag, "<t");
                            if (tv) {
                                const char *te;
                                tv = strchr(tv, '>') + 1;
                                te = strchr(tv, '<');
                                if (te) {
                                    char sval[ARK_MAX_STR_LEN];
                                    xml_copy_decoded(tv, (int)(te - tv),
                                                     sval, sizeof(sval));
                                    cell_put_raw(wb, si, row, col, ARK_CELL_ISTR,
                                                 style,
                                                 (uint32_t)ark_str_add(wb, sval), 0);
                                }
                            }
                        } else if (vs && ve) {
                            const char *vv = strchr(vs, '>') + 1;
                            const char *vend = strchr(vv, '<');
                            int vn = (int)(vend - vv);
                            char val[ARK_MAX_STR_LEN];
                            xml_copy_decoded(vv, vn, val, sizeof(val));
                            if (!tcode || tcode == 'n') {  /* 数字/日期 */
                                cell_put_raw(wb, si, row, col, ARK_CELL_NUM, style,
                                             0, xml_atof(vv, vn));
                            } else if (tcode == 's') {   /* SST 引用 */
                                long idx = xml_atoi(vv, vn);
                                if (idx >= 0 && idx < wb->sst_count)
                                    cell_put_raw(wb, si, row, col, ARK_CELL_SSTR,
                                                 style, (uint32_t)idx, 0);
                            } else if (tcode == 'b') {
                                cell_put_raw(wb, si, row, col, ARK_CELL_BOOL, style,
                                             0, xml_atoi(vv, vn) ? 1 : 0);
                            } else if (tcode == 'e') {
                                cell_put_raw(wb, si, row, col, ARK_CELL_ERROR, style,
                                             (uint32_t)ark_str_add(wb, val), 0);
                            } else if (tcode == 'i' || tcode == 'd'
                                       || tcode == 't') {
                                cell_put_raw(wb, si, row, col, ARK_CELL_ISTR, style,
                                             (uint32_t)ark_str_add(wb, val), 0);
                            }
                        } else {   /* 空单元格(仅样式) */
                            if (style) ark_set_blank(wb, si, row, col, style);
                        }
                    }
                    cq = ve ? ve + 4 : cq + 3;
                }
            }
            q = strstr(q, "</row>");
            if (!q) break;
            q += 6;
        }
    }
    /* 合并单元格 */
    {   const char *t = xml_find(p, "<mergeCells ");
        if (t) { const char *q = t;
            while ((q = xml_find(q, "<mergeCell "))) {
                int ln; const char *v = xml_attr(q, "ref", &ln);
                if (v) { char ref[40]; xml_copy_decoded(v, ln, ref, sizeof(ref));
                    char *c = strchr(ref, ':');
                    if (c) { *c = 0;
                        uint32_t r1, r2; uint16_t c1, c2;
                        if (ref_to_rc(ref, (int)strlen(ref), &r1, &c1) == 0 &&
                            ref_to_rc(c + 1, (int)strlen(c + 1), &r2, &c2) == 0)
                            ark_merge(wb, si, r1, c1, r2, c2);
                    }
                }
                q += 11;
            }
        }
    }
    return ARK_OK;
}

/* ============================================================
 * 5. 公开 API
 * ========================================================== */
static ark_reader_t *s_rd = NULL;   /* touch 回调查找当前 reader */

void ark_rd_touch(ark_wb_t *wb, uint8_t sheet)
{
    if (wb && wb->reader) {
        ark_reader_t *rd = (ark_reader_t *)wb->reader;
        if (sheet < wb->sheet_count) rd->sheet_dirty[sheet] = 1;
    }
}
void ark_rd_touch_sst(ark_wb_t *wb)
{
    if (wb && wb->reader) ((ark_reader_t *)wb->reader)->sst_dirty = 1;
}
void ark_rd_touch_styles(ark_wb_t *wb)
{
    if (wb && wb->reader) ((ark_reader_t *)wb->reader)->styles_dirty = 1;
}
void ark_rd_touch_workbook(ark_wb_t *wb)
{
    if (wb && wb->reader) {
        ark_reader_t *rd = (ark_reader_t *)wb->reader;
        rd->wb_dirty = 1;
    }
}

static uint16_t s_style_map[ARK_STYLE_BUILTIN_COUNT];
uint16_t ark_style_map(uint16_t builtin_id)
{
    if (builtin_id < ARK_STYLE_BUILTIN_COUNT && s_rd && s_rd->style_map_ok)
        return s_style_map[builtin_id];
    return builtin_id;
}

ark_err_t ark_wb_open(ark_wb_t *wb, ark_reader_t *rd, const ark_reader_io_t *io)
{
    ark_err_t e;
    int i;
    if (!wb || !rd || !io || !io->read || !io->seek) return ARK_ERR_INVALID_ARG;
    memset(rd, 0, sizeof(*rd));
    rd->io = *io;
    rd->i_content_types = rd->i_workbook = rd->i_rels = 0xFFFF;
    rd->i_sst = rd->i_styles = 0xFFFF;
    for (i = 0; i < ARK_MAX_SHEETS; i++) rd->sheet_entry[i] = 0xFFFF;

    /* 文件大小: 从 0xFFFFFFFF seek 失败后逐步试出 —— 简化: 先读文件头
     * 判定 ZIP 签名, 再通过 EOCD 搜索界定(file_size 由 seek+read 探测) */
    {
        uint8_t t[4];
        if (io->read(io->ctx, t, 4) != 4 || rd_u32(t) != 0x04034B50u)
            return ARK_ERR_BAD_STATE;
        /* 探测文件大小: 二分 seek */
        uint32_t lo = 4, hi = 1u << 30;
        while (lo < hi) {
            uint32_t mid = lo + (hi - lo) / 2;
            if (io->seek(io->ctx, mid) != 0) { hi = mid; continue; }
            uint32_t n = io->read(io->ctx, t, 1);
            if (n == 1) lo = mid + 1; else hi = mid;
        }
        rd->file_size = lo;
    }
if (zip_parse_central(rd) != 0) return ARK_ERR_BAD_STATE;

    /* 定位关键部件 */
    {   int i;
        for (i = 0; i < rd->entry_count; i++) {
            const char *n = rd->entries[i].name;
            if (!strcmp(n, "[Content_Types].xml"))      rd->i_content_types = (uint16_t)i;
            else if (!strcmp(n, "xl/workbook.xml"))      rd->i_workbook = (uint16_t)i;
            else if (!strcmp(n, "xl/_rels/workbook.xml.rels")) rd->i_rels = (uint16_t)i;
            else if (!strcmp(n, "xl/sharedStrings.xml")) rd->i_sst = (uint16_t)i;
            else if (!strcmp(n, "xl/styles.xml"))        rd->i_styles = (uint16_t)i;
        }
    }
    if (rd->i_workbook == 0xFFFF || rd->i_rels == 0xFFFF)
        return ARK_ERR_BAD_STATE;

    wb->reader = rd;      /* 之后 add_sheet 等才能 touch 到位 */
    s_rd = rd;

    /* 载入顺序: workbook 结构 -> styles 计数 -> SST -> sheets */
    if ((e = rd_load_workbook(rd, wb))) return e;
    if ((e = rd_load_styles(rd))) return e;

    /* 样式表同步到模板计数, 内置样式追加注册 */
    ark_i_st_sync_open(rd->tpl_font_cnt, rd->tpl_fill_cnt, rd->tpl_border_cnt,
                       rd->tpl_xf_cnt, rd->tpl_numfmt_maxid);
    ark_i_st_register_builtins_open(s_style_map);
    rd->style_map_ok = 1;

    if ((e = rd_load_sst(rd, wb))) return e;
    for (i = 0; i < wb->sheet_count; i++)
        if ((e = rd_load_sheet(rd, wb, (uint8_t)i))) return e;

    /* 载入完成: 清除载入过程产生的 dirty 标记 */
    memset(rd->sheet_dirty, 0, sizeof(rd->sheet_dirty));
    rd->sst_dirty = rd->styles_dirty = rd->wb_dirty = 0;
    rd->tpl_sheet_cnt = wb->sheet_count;
    rd->opened = 1;
    return ARK_OK;
}

int ark_get_cell(ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                 uint8_t *type, double *num, char *text, size_t textcap,
                 uint16_t *style)
{
    uint32_t i;
    if (!wb || sheet < 0 || (uint32_t)sheet >= wb->sheet_count) return 1;
    for (i = 0; i < wb->cell_count; i++) {
        ark_cell_t *c = &wb->cells[i];
        if (c->sheet == (uint8_t)sheet && c->row == row && c->col == col) {
            if (type)  *type = c->type;
            if (num)   *num = c->num;
            if (style) *style = c->style;
            if (text && textcap) {
                text[0] = 0;
                switch (c->type) {
                case ARK_CELL_SSTR: {   /* sidx 是 SST 序号, 查 arena */
                    for (uint16_t k = 0; k < wb->str_count; k++)
                        if (wb->strs[k].is_sst && wb->strs[k].sst_seq == c->sidx) {
                            snprintf(text, textcap, "%s", wb->strs[k].text); break;
                        }
                    break;
                }
                case ARK_CELL_ISTR: case ARK_CELL_ERROR: case ARK_CELL_FORMULA:
                    snprintf(text, textcap, "%s", ark_str_get(wb, (uint16_t)c->sidx));
                    break;
                default: break;
                }
            }
            return 0;
        }
    }
    return 1;
}
#endif
