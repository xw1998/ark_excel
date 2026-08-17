/*!
 * @file    ark_excel.c
 * @brief   ark_excel 核心实现: 单元格/SST/样式表管理 + xlsx 全部 XML 流式生成
 *
 * 移植性: 仅依赖 <string.h> <stdio.h>(snprintf) <stdarg.h>, 无 OS 调用,
 *         无动态内存, 输出全部经用户回调.
 */
#include <string.h>
#include <stdarg.h>
#include <stdio.h>
#include "ark_excel.h"
#include "ark_excel_internal.h"
#if ARK_ENABLE_READWRITE
ark_err_t ark_save_open(ark_wb_t *wb);   /* 读改写增量保存(文件尾部) */
#endif
#include "ark_zip.h"

/* ============================================================
 * 内部: XML 发射器(分块缓冲, 两遍: 0=测量 1=输出)
 * ========================================================== */
typedef struct {
    char       buf[ARK_IO_CHUNK];
    uint16_t   len;
    int        pass;      /* 0 measure, 1 output */
    ark_zip_t *zip;
    ark_err_t  err;
} ark_xw_t;

static void xw_init(ark_xw_t *x, ark_zip_t *zip, int pass)
{
    x->len = 0; x->pass = pass; x->zip = zip; x->err = ARK_OK;
}
static void xw_flush(ark_xw_t *x)
{
    if (x->len == 0) return;
    if (x->err == ARK_OK) {
        if (x->pass == 0) ark_zip_measure_write(x->zip, x->buf, x->len);
        else if (ark_zip_write(x->zip, x->buf, x->len) != 0)
            x->err = ARK_ERR_WRITE;
    }
    x->len = 0;
}
static void xw_raw(ark_xw_t *x, const char *s, size_t n)
{
    while (n--) {
        if (x->len == ARK_IO_CHUNK) xw_flush(x);
        x->buf[x->len++] = *s++;
    }
}
static void xw_str(ark_xw_t *x, const char *s)   { xw_raw(x, s, strlen(s)); }
static void xw_ch(ark_xw_t *x, char c)           { xw_raw(x, &c, 1); }

static void xw_u32(ark_xw_t *x, uint32_t v)
{
    char t[10]; uint8_t i = 0;
    if (!v) { xw_ch(x, '0'); return; }
    while (v) { t[i++] = (char)('0' + (v % 10)); v /= 10; }
    while (i) xw_ch(x, t[--i]);
}
static void xw_i64(ark_xw_t *x, int64_t v)
{
    char t[24]; uint8_t i = 0; uint64_t u;
    if (v < 0) { xw_ch(x, '-'); u = (uint64_t)(-(v + 1)) + 1; }
    else u = (uint64_t)v;
    if (!u) { xw_ch(x, '0'); return; }
    while (u) { t[i++] = (char)('0' + (u % 10)); u /= 10; }
    while (i) xw_ch(x, t[--i]);
}
static void xw_dbl(ark_xw_t *x, double v)
{
    char t[40];
    int n = ARK_SNPRINTF(t, sizeof(t), "%.15g", v);
    if (n <= 0) { xw_str(x, "0"); return; }
    xw_raw(x, t, (size_t)n);
}

/*! XML 转义输出; 过滤非法控制字符 */
static void xw_esc(ark_xw_t *x, const char *s)
{
    while (*s) {
        unsigned char c = (unsigned char)*s++;
        switch (c) {
        case '&':  xw_str(x, "&amp;");  break;
        case '<':  xw_str(x, "&lt;");   break;
        case '>':  xw_str(x, "&gt;");   break;
        case '\"': xw_str(x, "&quot;"); break;
        case '\'': xw_str(x, "&apos;"); break;
        default:
            if (c >= 0x20 || c == 0x09 || c == 0x0A || c == 0x0D)
                xw_ch(x, (char)c);
            break;                       /* 其余控制字符丢弃 */
        }
    }
}

/*! printf 风格(仅用于短串, 如属性) */
static void xw_fmt(ark_xw_t *x, const char *fmt, ...)
{
    char t[ARK_IO_CHUNK];
    va_list ap;
    int n;
    va_start(ap, fmt);
    n = vsnprintf(t, sizeof(t), fmt, ap);
    va_end(ap);
    if (n > 0) xw_raw(x, t, (size_t)n);
}

/* ============================================================
 * 内部: 字符串 arena
 * ========================================================== */
#define ARK_IDX_NONE 0xFFFFu

int ark_str_add(ark_wb_t *wb, const char *s)
{
    size_t n;
    if (!s) return -(int)ARK_ERR_INVALID_ARG;
    n = strlen(s);
    if (n >= ARK_MAX_STR_LEN) return -(int)ARK_ERR_STRING_TOO_LONG;
    if (wb->str_count >= ARK_MAX_STR) return -(int)ARK_ERR_NO_MEMORY;
    wb->strs[wb->str_count].len = (uint16_t)n;
    wb->strs[wb->str_count].is_sst = 0;
    memcpy(wb->strs[wb->str_count].text, s, n + 1);
    return wb->str_count++;
}
const char *ark_str_get(ark_wb_t *wb, uint16_t idx)
{
    return wb->strs[idx].text;
}

/*! SST 添加(去重); 返回 SST 索引(即 arena 索引) 或负错误码 */
int ark_sst_add(ark_wb_t *wb, const char *s)
{
    uint16_t i;
    for (i = 0; i < wb->str_count; i++)
        if (wb->strs[i].is_sst && strcmp(wb->strs[i].text, s) == 0)
            return wb->strs[i].sst_seq;
    if (wb->sst_count >= ARK_MAX_SST)
        return -(int)ARK_ERR_SST_FULL;
    {
        int r = ark_str_add(wb, s);
        if (r < 0) return r;
        wb->strs[r].is_sst = 1;
        wb->strs[r].sst_seq = wb->sst_count++;
        ark_rd_touch_sst(wb);
        return wb->strs[r].sst_seq;
    }
}

/* ============================================================
 * 内部: 样式表
 * ========================================================== */
typedef struct {
    const char *name;
    uint16_t    size_x100;
    uint32_t    color;
    uint8_t     bold, italic, underline;
} ark_font_rec_t;
typedef struct { uint32_t argb; } ark_fill_rec_t;
typedef struct {
    uint8_t  l, r, t, b;      /* ARK_BORDER_xxx */
    uint8_t  dl;              /* 对角线样式 ARK_BORDER_xxx, 0=无 */
    uint8_t  dup, ddown;      /* diagonalUp / diagonalDown */
    uint32_t color;
} ark_border_rec_t;
typedef struct { uint16_t id; const char *fmt; } ark_numfmt_rec_t;
typedef struct {
    uint8_t  font, fill, border;
    uint8_t  halign, valign, wrap, indent;
    uint16_t numfmt;          /* 内置 id 或 164+ */
} ark_xf_rec_t;

typedef struct {
    ark_font_rec_t    fonts[ARK_MAX_FONTS];
    uint16_t          nfonts;
    ark_fill_rec_t    fills[ARK_MAX_FILLS];
    uint16_t          nfills;
    ark_border_rec_t  borders[ARK_MAX_BORDERS];
    uint16_t          nborders;
    ark_numfmt_rec_t  numfmts[ARK_MAX_NUMFMTS];
    uint16_t          nnumfmts;
    ark_xf_rec_t      xfs[ARK_MAX_STYLES];
    uint16_t          nxfs;
} ark_style_table_t;

static ark_style_table_t s_st;   /* 样式表与工作簿同步重建, 无并发场景 */
static uint16_t s_fmt_base = 164;   /* 新 numFmt id 起始(open 模式接模板之后) */

static uint8_t st_add_font(const char *name, uint16_t size_x100, uint32_t color,
                           uint8_t bold, uint8_t italic, uint8_t underline)
{
    uint16_t i;
    ark_font_rec_t f = { name, size_x100 ? size_x100 : 1100, color,
                         bold, italic, underline };
    if (!f.name) f.name = "Calibri";
    for (i = 0; i < s_st.nfonts; i++)
        if (s_st.fonts[i].name == f.name &&
            s_st.fonts[i].size_x100 == f.size_x100 &&
            s_st.fonts[i].color == f.color &&
            s_st.fonts[i].bold == f.bold && s_st.fonts[i].italic == f.italic &&
            s_st.fonts[i].underline == f.underline)
            return (uint8_t)i;
    if (s_st.nfonts >= ARK_MAX_FONTS) return 0;
    s_st.fonts[s_st.nfonts++] = f;
    return (uint8_t)(s_st.nfonts - 1);
}
static uint8_t st_add_fill(uint32_t argb)
{
    uint16_t i;
    for (i = 2; i < s_st.nfills; i++)
        if (s_st.fills[i].argb == argb) return (uint8_t)i;
    if (s_st.nfills >= ARK_MAX_FILLS) return 0;
    s_st.fills[s_st.nfills++].argb = argb;
    return (uint8_t)(s_st.nfills - 1);
}
static uint8_t st_add_border2(uint8_t l, uint8_t r, uint8_t t, uint8_t b,
                              uint8_t dl, uint8_t dup, uint8_t ddown,
                              uint32_t color)
{
    uint16_t i;
    for (i = 0; i < s_st.nborders; i++)
        if (s_st.borders[i].l == l && s_st.borders[i].r == r &&
            s_st.borders[i].t == t && s_st.borders[i].b == b &&
            s_st.borders[i].dl == dl && s_st.borders[i].dup == dup &&
            s_st.borders[i].ddown == ddown &&
            s_st.borders[i].color == color)
            return (uint8_t)i;
    if (s_st.nborders >= ARK_MAX_BORDERS) return 0;
    s_st.borders[s_st.nborders].l = l; s_st.borders[s_st.nborders].r = r;
    s_st.borders[s_st.nborders].t = t; s_st.borders[s_st.nborders].b = b;
    s_st.borders[s_st.nborders].dl = dl;
    s_st.borders[s_st.nborders].dup = dup;
    s_st.borders[s_st.nborders].ddown = ddown;
    s_st.borders[s_st.nborders].color = color;
    return (uint8_t)(s_st.nborders++);
}
#define st_add_border(l,r,t,b,color) st_add_border2(l,r,t,b,0,0,0,color)
static uint16_t st_add_numfmt(const char *fmt)
{
    uint16_t i;
    for (i = 0; i < s_st.nnumfmts; i++)
        if (strcmp(s_st.numfmts[i].fmt, fmt) == 0) return s_st.numfmts[i].id;
    if (s_st.nnumfmts >= ARK_MAX_NUMFMTS) return 0;
    s_st.numfmts[s_st.nnumfmts].id = (uint16_t)(s_fmt_base + s_st.nnumfmts);
    s_st.numfmts[s_st.nnumfmts].fmt = fmt;
    return s_st.numfmts[s_st.nnumfmts++].id;
}
static uint16_t st_add_xf(const ark_xf_rec_t *xf)
{
    if (s_st.nxfs >= ARK_MAX_STYLES) return 0;
    s_st.xfs[s_st.nxfs++] = *xf;
    return (uint16_t)(s_st.nxfs - 1);
}

/*! 注册预定义样式, 必须与 ark_excel.h 中 ARK_STYLE_* 枚举一一对应 */
static void st_register_builtins(void)
{
    uint8_t f_def   = st_add_font("Calibri", 1100, 0, 0, 0, 0);
    uint8_t f_title = st_add_font("Calibri", 1400, 0, 1, 0, 0);
    uint8_t f_hdr   = st_add_font("Calibri", 1100, ARK_COLOR_WHITE, 1, 0, 0);
    uint8_t f_bold  = st_add_font("Calibri", 1100, 0, 1, 0, 0);
    uint8_t f_warn  = st_add_font("Calibri", 1100, ARK_COLOR_WARNFG, 1, 0, 0);
    uint8_t f_ok    = st_add_font("Calibri", 1100, ARK_COLOR_OKFG, 1, 0, 0);
    uint8_t f_link  = st_add_font("Calibri", 1100, 0xFF0000FFu, 0, 0, 1);
    uint8_t fl_hdr  = st_add_fill(ARK_COLOR_HEADERBG);
    uint8_t fl_alt  = st_add_fill(ARK_COLOR_ALTBG);
    uint8_t fl_warn = st_add_fill(ARK_COLOR_WARNBG);
    uint8_t fl_ok   = st_add_fill(ARK_COLOR_OKBG);
    uint8_t b_all   = st_add_border(1,1,1,1, 0);
    uint8_t b_top   = st_add_border(0,0,1,0, 0);
    uint16_t nf_date = st_add_numfmt("yyyy-mm-dd");
    uint16_t nf_dt   = st_add_numfmt("yyyy-mm-dd hh:mm:ss");
    uint16_t nf_time = st_add_numfmt("hh:mm:ss");
    uint16_t nf_2d   = st_add_numfmt("0.00");
    uint16_t nf_int  = st_add_numfmt("0");
    uint16_t nf_pct  = st_add_numfmt("0.00%");
    uint16_t nf_mon  = st_add_numfmt("#,##0.00");
    ark_xf_rec_t xf;

    memset(&xf, 0, sizeof(xf)); xf.font = f_def;
    (void)st_add_xf(&xf);                                    /* 0 DEFAULT */

    xf.halign = ARK_HALIGN_CENTER; xf.valign = ARK_VALIGN_CENTER;
    xf.font = f_title; (void)st_add_xf(&xf);                 /* 1 TITLE */

    xf.font = f_hdr; xf.fill = fl_hdr; xf.border = b_all;
    xf.wrap = 1; (void)st_add_xf(&xf);                       /* 2 HEADER */

    xf.font = f_def; xf.fill = 0; xf.wrap = 0;
    xf.halign = ARK_HALIGN_LEFT; xf.border = b_all;
    (void)st_add_xf(&xf);                                    /* 3 DATA */
    xf.halign = ARK_HALIGN_CENTER; (void)st_add_xf(&xf);     /* 4 DATA_CENTER */
    xf.halign = ARK_HALIGN_RIGHT;  (void)st_add_xf(&xf);     /* 5 DATA_RIGHT */

    xf.numfmt = nf_2d; (void)st_add_xf(&xf);                 /* 6 NUM_2DEC */
    xf.numfmt = nf_int;(void)st_add_xf(&xf);                 /* 7 INT */
    xf.numfmt = nf_pct;(void)st_add_xf(&xf);                 /* 8 PERCENT */

    xf.numfmt = nf_date; xf.halign = ARK_HALIGN_CENTER;
    (void)st_add_xf(&xf);                                    /* 9 DATE */
    xf.numfmt = nf_dt;   (void)st_add_xf(&xf);               /* 10 DATETIME */
    xf.numfmt = nf_time; (void)st_add_xf(&xf);               /* 11 TIME */

    xf.numfmt = nf_mon; xf.halign = ARK_HALIGN_RIGHT;
    (void)st_add_xf(&xf);                                    /* 12 MONEY */

    xf.numfmt = 0; xf.halign = ARK_HALIGN_LEFT; xf.fill = fl_alt;
    xf.border = b_all; (void)st_add_xf(&xf);                 /* 13 ALT_ROW */

    xf.fill = fl_warn; xf.font = f_warn; xf.halign = ARK_HALIGN_CENTER;
    (void)st_add_xf(&xf);                                    /* 14 WARN */

    xf.fill = fl_ok; xf.font = f_ok; (void)st_add_xf(&xf);   /* 15 OK */

    xf.fill = 0; xf.font = f_bold; xf.border = b_top;
    (void)st_add_xf(&xf);                                    /* 16 TOTAL */

    xf.border = 0; xf.font = f_link;
    xf.halign = ARK_HALIGN_LEFT; (void)st_add_xf(&xf);       /* 17 HYPERLINK */

    xf.font = f_bold; xf.border = b_all;
    xf.halign = ARK_HALIGN_CENTER; xf.valign = ARK_VALIGN_CENTER;
    (void)st_add_xf(&xf);                                    /* 18 MERGE_CENTER */
}

void ark_i_st_sync_open(uint16_t fonts, uint16_t fills, uint16_t borders,
                        uint16_t xfs, uint16_t max_numfmt_id)
{
    memset(&s_st, 0, sizeof(s_st));
    s_st.nfonts = fonts; s_st.nfills = fills; s_st.nborders = borders;
    s_st.nxfs = xfs;
    /* 新 numFmt id 从模板最大 id 之后续编; 计数保持 0(去重表从空开始) */
    s_fmt_base = max_numfmt_id >= 164 ? (uint16_t)(max_numfmt_id + 1) : 164;
}

void ark_i_st_register_builtins_open(uint16_t *map)
{
    uint16_t base = s_st.nxfs, i;
    st_register_builtins();
    for (i = 0; i < ARK_STYLE_BUILTIN_COUNT && base + i < s_st.nxfs; i++)
        map[i] = (uint16_t)(base + i);
}

/* ============================================================
 * 工具: 列字母 / 日期序列号
 * ========================================================== */
void ark_col_to_letters(uint16_t col, char *buf)
{
    char t[4]; int i = 0;
    while (col > 0) {
        int rem = (col - 1) % 26;
        t[i++] = (char)('A' + rem);
        col = (uint16_t)((col - 1) / 26);
    }
    { int j; for (j = 0; j < i; j++) buf[j] = t[i - 1 - j]; }
    buf[i] = 0;
}

/*! 公历天数(Howard Hinnant 算法), 输入 y>=1 */
static int64_t days_from_civil(int y, int m, int d)
{
    y -= m <= 2;
    { int64_t era = (y >= 0 ? y : y - 399) / 400;
      uint32_t yoe = (uint32_t)(y - era * 400);
      uint32_t doy = (uint32_t)((153 * (m + (m > 2 ? -3 : 9)) + 2) / 5 + d - 1);
      uint32_t doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
      return era * 146097 + (int64_t)doe - 719468;
    }
}

double ark_date_to_serial(int y, int mo, int d, int h, int mi, int s)
{
    int64_t days = days_from_civil(y, mo, d) - days_from_civil(1899, 12, 31);
    double serial = (double)days;
    if (days >= 60) serial += 1.0;         /* 1900 伪闰日修正:
                                               1900-02-28=59, (伪02-29=60),
                                               1900-03-01 起真实天数+1 */
    serial += (h * 3600.0 + mi * 60.0 + s) / 86400.0;
    return serial;
}

const char *ark_errstr(ark_err_t e)
{
    switch (e) {
    case ARK_OK:                return "ok";
    case ARK_ERR_NO_MEMORY:     return "memory pool exhausted";
    case ARK_ERR_TOO_MANY_SHEETS: return "too many sheets";
    case ARK_ERR_TOO_MANY_CELLS:  return "too many cells";
    case ARK_ERR_TOO_MANY_MERGES: return "too many merges";
    case ARK_ERR_TOO_MANY_STYLES: return "too many styles";
    case ARK_ERR_STRING_TOO_LONG: return "string too long";
    case ARK_ERR_SST_FULL:      return "shared string table full";
    case ARK_ERR_TABLE_FULL:    return "feature table full";
    case ARK_ERR_INVALID_ARG:   return "invalid argument";
    case ARK_ERR_INVALID_ROW:   return "invalid row";
    case ARK_ERR_INVALID_COL:   return "invalid column";
    case ARK_ERR_INVALID_SHEET: return "invalid sheet";
    case ARK_ERR_INVALID_STYLE: return "invalid style";
    case ARK_ERR_INVALID_RANGE: return "invalid range";
    case ARK_ERR_INVALID_NAME:  return "invalid sheet name";
    case ARK_ERR_WRITE:         return "output callback failed";
    case ARK_ERR_BAD_STATE:     return "bad state";
    case ARK_ERR_SAVED:         return "workbook already saved";
    default:                    return "unknown";
    }
}

/* ============================================================
 * 生命周期
 * ========================================================== */
ark_err_t ark_wb_init(ark_wb_t *wb, ark_write_fn write, void *ctx)
{
    if (!wb || !write) return ARK_ERR_INVALID_ARG;
    memset(wb, 0, sizeof(*wb));
    wb->write = write;
    wb->ctx = ctx;
    /* 重建样式表 + 注册预定义样式 */
    memset(&s_st, 0, sizeof(s_st));
    s_st.fills[1].argb = 0;               /* fill1 = gray125 标记 */
    s_st.nfills = 2;                      /* 0:none 1:gray125 必需 */
    (void)st_add_border(0,0,0,0, 0);     /* border0 = 空边框(OOXML 惯例) */
    st_register_builtins();
    return ARK_OK;
}

void ark_wb_reset(ark_wb_t *wb)
{
    ark_write_fn w = wb->write;
    void *c = wb->ctx;
    memset(wb, 0, sizeof(*wb));
    wb->write = w; wb->ctx = c;
    memset(&s_st, 0, sizeof(s_st));
    s_st.nfills = 2;
    (void)st_add_border(0,0,0,0, 0);
    st_register_builtins();
}

int ark_wb_add_sheet(ark_wb_t *wb, const char *name)
{
    ark_rd_touch_workbook(wb);
    size_t n;
    uint8_t i;
    if (!wb || !name) return ARK_ERR_INVALID_ARG;
    if (wb->saved) return ARK_ERR_SAVED;
    if (wb->sheet_count >= ARK_MAX_SHEETS) return ARK_ERR_TOO_MANY_SHEETS;
    n = strlen(name);
    if (n == 0 || n >= ARK_SHEET_NAME_MAX) return ARK_ERR_INVALID_NAME;
    for (i = 0; i < wb->sheet_count; i++)
        if (strcmp(wb->sheets[i].name, name) == 0) return ARK_ERR_INVALID_NAME;
    memset(&wb->sheets[wb->sheet_count], 0, sizeof(ark_sheet_t));
    memcpy(wb->sheets[wb->sheet_count].name, name, n + 1);
    wb->sheets[wb->sheet_count].show_gridlines = 1;
    wb->sheets[wb->sheet_count].print_area_idx = ARK_IDX_NONE;
    wb->sheets[wb->sheet_count].header_idx = ARK_IDX_NONE;
    wb->sheets[wb->sheet_count].footer_idx = ARK_IDX_NONE;
    {
        ark_page_setup_t def = ARK_PAGE_SETUP_DEFAULT;
        wb->sheets[wb->sheet_count].page = def;
    }
    return wb->sheet_count++;
}

/* ============================================================
 * 单元格
 * ========================================================== */
#define ARK_MAX_ROW 1048576u
#define ARK_MAX_COL 16384u

static ark_err_t cell_check(ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                            uint16_t style)
{
    if (!wb) return ARK_ERR_INVALID_ARG;
    if (wb->saved) return ARK_ERR_SAVED;
    if (sheet < 0 || (uint32_t)sheet >= wb->sheet_count) return ARK_ERR_INVALID_SHEET;
    if (row < 1 || row > ARK_MAX_ROW) return ARK_ERR_INVALID_ROW;
    if (col < 1 || col > ARK_MAX_COL) return ARK_ERR_INVALID_COL;
    if (style >= s_st.nxfs) return ARK_ERR_INVALID_STYLE;
    return ARK_OK;
}

/*! 查找已有单元格(同址覆盖), 返回下标或 -1 */
static int32_t cell_find(ark_wb_t *wb, uint8_t sheet, uint32_t row, uint16_t col)
{
    uint32_t i;
    for (i = 0; i < wb->cell_count; i++)
        if (wb->cells[i].sheet == sheet && wb->cells[i].row == row &&
            wb->cells[i].col == col)
            return (int32_t)i;
    return -1;
}

static ark_err_t cell_put(ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                          uint8_t type, uint16_t style, uint32_t sidx, double num)
{
    ark_err_t e = cell_check(wb, sheet, row, col, style);
    ark_cell_t *c;
    int32_t idx;
    if (e != ARK_OK) return e;
    ark_rd_touch(wb, (uint8_t)sheet);
#if ARK_ENABLE_READWRITE
    /* open 模式: 引用模板之外的新样式(含内置重映射)需重生成 styles */
    if (wb->reader && style >= ((ark_reader_t *)wb->reader)->tpl_xf_cnt)
        ((ark_reader_t *)wb->reader)->styles_dirty = 1;
#endif
    idx = cell_find(wb, (uint8_t)sheet, row, col);
    if (idx >= 0) {
        c = &wb->cells[idx];
    } else {
        if (wb->cell_count >= ARK_MAX_CELLS) return ARK_ERR_TOO_MANY_CELLS;
        c = &wb->cells[wb->cell_count++];
        c->sheet = (uint8_t)sheet; c->row = row; c->col = col;
    }
    c->type = type; c->style = style; c->sidx = sidx; c->num = num;
    return ARK_OK;
}

ark_err_t ark_set_str(ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                      const char *text, uint16_t style)
{
    int idx;
    if (!wb || !text) return ARK_ERR_INVALID_ARG;
    idx = ark_sst_add(wb, text);
    if (idx >= 0)
        return cell_put(wb, sheet, row, col, ARK_CELL_SSTR, style,
                        (uint32_t)idx, 0);
    /* 降级: 内联字符串 */
    idx = ark_str_add(wb, text);
    if (idx < 0) return (ark_err_t)(-idx);
    return cell_put(wb, sheet, row, col, ARK_CELL_ISTR, style,
                    (uint32_t)(uint16_t)idx, 0);
}

ark_err_t ark_set_num(ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                      double value, uint16_t style)
{
    return cell_put(wb, sheet, row, col, ARK_CELL_NUM, style, 0, value);
}
ark_err_t ark_set_int(ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                      int64_t value, uint16_t style)
{
    return cell_put(wb, sheet, row, col, ARK_CELL_INT, style, 0, (double)value);
}
ark_err_t ark_set_bool(ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                       int value, uint16_t style)
{
    return cell_put(wb, sheet, row, col, ARK_CELL_BOOL, style, 0, value ? 1 : 0);
}
ark_err_t ark_set_formula(ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                          const char *formula, uint16_t style)
{
#if !ARK_ENABLE_FORMULA
    (void)wb; return ARK_ERR_BAD_STATE;
#else
    int idx;
    if (!formula) return ARK_ERR_INVALID_ARG;
    idx = ark_str_add(wb, formula);
    if (idx < 0) return (ark_err_t)(-idx);
    return cell_put(wb, sheet, row, col, ARK_CELL_FORMULA, style,
                    (uint32_t)(uint16_t)idx, 0);
#endif
}
ark_err_t ark_set_error(ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                        const char *errval, uint16_t style)
{
    int idx;
    if (!errval) return ARK_ERR_INVALID_ARG;
    idx = ark_str_add(wb, errval);
    if (idx < 0) return (ark_err_t)(-idx);
    return cell_put(wb, sheet, row, col, ARK_CELL_ERROR, style,
                    (uint32_t)(uint16_t)idx, 0);
}
ark_err_t ark_set_blank(ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                        uint16_t style)
{
    return cell_put(wb, sheet, row, col, ARK_CELL_BLANK, style, 0, 0);
}
ark_err_t ark_set_date(ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                       int y, int mo, int d, uint16_t style)
{
    return cell_put(wb, sheet, row, col, ARK_CELL_DATE, style, 0,
                    ark_date_to_serial(y, mo, d, 0, 0, 0));
}
ark_err_t ark_set_datetime(ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                           int y, int mo, int d, int h, int mi, int s,
                           uint16_t style)
{
    return cell_put(wb, sheet, row, col, ARK_CELL_DATE, style, 0,
                    ark_date_to_serial(y, mo, d, h, mi, s));
}

/* ============================================================
 * 特殊格式 API
 * ========================================================== */
ark_err_t ark_merge(ark_wb_t *wb, int sheet, uint32_t r1, uint16_t c1,
                    uint32_t r2, uint16_t c2)
{
    ark_rd_touch(wb, (uint8_t)sheet);
    ark_err_t e;
    if (!wb) return ARK_ERR_INVALID_ARG;
    if (wb->saved) return ARK_ERR_SAVED;
    if ((e = cell_check(wb, sheet, r1, c1, 0)) != ARK_OK) return e;
    if (r2 < r1 || c2 < c1 || r2 > 65535 || r1 < 1) return ARK_ERR_INVALID_RANGE;
    if (wb->merge_count >= ARK_MAX_MERGES) return ARK_ERR_TOO_MANY_MERGES;
    wb->merges[wb->merge_count].sheet = (uint8_t)sheet;
    wb->merges[wb->merge_count].r1 = (uint16_t)r1;
    wb->merges[wb->merge_count].c1 = c1;
    wb->merges[wb->merge_count].r2 = (uint16_t)r2;
    wb->merges[wb->merge_count].c2 = c2;
    wb->merge_count++;
    return ARK_OK;
}

static ark_err_t coldef_add(ark_wb_t *wb, int sheet, int first, int last,
                            uint16_t w100, uint8_t hidden)
{
    ark_rd_touch(wb, (uint8_t)sheet);
    if (!wb) return ARK_ERR_INVALID_ARG;
    if (wb->saved) return ARK_ERR_SAVED;
    if (sheet < 0 || (uint32_t)sheet >= wb->sheet_count) return ARK_ERR_INVALID_SHEET;
    if (first < 1 || last < first || (uint32_t)last > ARK_MAX_COL) return ARK_ERR_INVALID_COL;
    if (wb->coldef_count >= ARK_MAX_COLDEFS) return ARK_ERR_TABLE_FULL;
    wb->coldefs[wb->coldef_count].first = (uint16_t)first;
    wb->coldefs[wb->coldef_count].last = (uint16_t)last;
    wb->coldefs[wb->coldef_count].sheet = (uint8_t)sheet;
    wb->coldefs[wb->coldef_count].width_x100 = w100;
    wb->coldefs[wb->coldef_count].hidden = hidden;
    wb->coldef_count++;
    return ARK_OK;
}

ark_err_t ark_set_col_width(ark_wb_t *wb, int sheet, int first, int last,
                            double width)
{
    if (width < 0 || width > 255) return ARK_ERR_INVALID_ARG;
    return coldef_add(wb, sheet, first, last, (uint16_t)(width * 100 + 0.5), 0);
}
ark_err_t ark_set_col_hidden(ark_wb_t *wb, int sheet, int first, int last,
                             int hidden)
{
    return coldef_add(wb, sheet, first, last, 0, hidden ? 1 : 0);
}

static ark_rowdef_t *rowdef_find(ark_wb_t *wb, int sheet, uint32_t row)
{
    uint16_t i;
    for (i = 0; i < wb->rowdef_count; i++)
        if (wb->rowdefs[i].sheet == (uint8_t)sheet &&
            wb->rowdefs[i].row == row)
            return &wb->rowdefs[i];
    return NULL;
}
static ark_err_t rowdef_get(ark_wb_t *wb, int sheet, uint32_t row,
                            ark_rowdef_t **out)
{
    ark_rd_touch(wb, (uint8_t)sheet);
    if (!wb) return ARK_ERR_INVALID_ARG;
    if (wb->saved) return ARK_ERR_SAVED;
    if (sheet < 0 || (uint32_t)sheet >= wb->sheet_count) return ARK_ERR_INVALID_SHEET;
    if (row < 1 || row > ARK_MAX_ROW) return ARK_ERR_INVALID_ROW;
    *out = rowdef_find(wb, sheet, row);
    if (!*out) {
        if (wb->rowdef_count >= ARK_MAX_ROWDEFS) return ARK_ERR_TABLE_FULL;
        *out = &wb->rowdefs[wb->rowdef_count++];
        memset(*out, 0, sizeof(**out));
        (*out)->row = row;
        (*out)->sheet = (uint8_t)sheet;
    }
    return ARK_OK;
}
ark_err_t ark_set_row_height(ark_wb_t *wb, int sheet, uint32_t row, double pt)
{
    ark_rowdef_t *rd;
    ark_err_t e;
    if (pt < 0 || pt > 409) return ARK_ERR_INVALID_ARG;
    if ((e = rowdef_get(wb, sheet, row, &rd)) != ARK_OK) return e;
    rd->height_x20 = (uint16_t)(pt * 20 + 0.5);
    return ARK_OK;
}
ark_err_t ark_set_row_hidden(ark_wb_t *wb, int sheet, uint32_t row, int hidden)
{
    ark_rowdef_t *rd;
    ark_err_t e;
    if ((e = rowdef_get(wb, sheet, row, &rd)) != ARK_OK) return e;
    rd->hidden = hidden ? 1 : 0;
    return ARK_OK;
}

ark_err_t ark_freeze(ark_wb_t *wb, int sheet, int rows, int cols)
{
    ark_rd_touch(wb, (uint8_t)sheet);
    if (!wb) return ARK_ERR_INVALID_ARG;
    if (wb->saved) return ARK_ERR_SAVED;
    if (sheet < 0 || (uint32_t)sheet >= wb->sheet_count) return ARK_ERR_INVALID_SHEET;
    if (rows < 0 || rows > 1048575 || cols < 0 || cols > 16383) return ARK_ERR_INVALID_ARG;
    wb->sheets[sheet].freeze_rows = (uint8_t)(rows > 255 ? 255 : rows);
    wb->sheets[sheet].freeze_cols = (uint8_t)(cols > 255 ? 255 : cols);
    return ARK_OK;
}

ark_err_t ark_autofilter(ark_wb_t *wb, int sheet, uint32_t r1, uint16_t c1,
                         uint32_t r2, uint16_t c2)
{
    ark_rd_touch(wb, (uint8_t)sheet);
    ark_err_t e;
    if ((e = cell_check(wb, sheet, r1, c1, 0)) != ARK_OK) return e;
    if (r2 < r1 || c2 < c1 || r1 > 65535 || r2 > 65535) return ARK_ERR_INVALID_RANGE;
    wb->sheets[sheet].autofilter_set = 1;
    wb->sheets[sheet].filter_r1 = (uint16_t)r1;
    wb->sheets[sheet].filter_c1 = c1;
    wb->sheets[sheet].filter_r2 = (uint16_t)r2;
    wb->sheets[sheet].filter_c2 = c2;
    return ARK_OK;
}

ark_err_t ark_tab_color(ark_wb_t *wb, int sheet, uint32_t argb)
{
    ark_rd_touch(wb, (uint8_t)sheet);
    if (!wb) return ARK_ERR_INVALID_ARG;
    if (wb->saved) return ARK_ERR_SAVED;
    if (sheet < 0 || (uint32_t)sheet >= wb->sheet_count) return ARK_ERR_INVALID_SHEET;
    wb->sheets[sheet].tab_color = argb;
    wb->sheets[sheet].tab_color_set = 1;
    return ARK_OK;
}
ark_err_t ark_sheet_rtl(ark_wb_t *wb, int sheet, int enable)
{
    ark_rd_touch(wb, (uint8_t)sheet);
    if (!wb) return ARK_ERR_INVALID_ARG;
    if (sheet < 0 || (uint32_t)sheet >= wb->sheet_count) return ARK_ERR_INVALID_SHEET;
    wb->sheets[sheet].rtl = enable ? 1 : 0;
    return ARK_OK;
}
ark_err_t ark_sheet_gridlines(ark_wb_t *wb, int sheet, int show)
{
    ark_rd_touch(wb, (uint8_t)sheet);
    if (!wb) return ARK_ERR_INVALID_ARG;
    if (sheet < 0 || (uint32_t)sheet >= wb->sheet_count) return ARK_ERR_INVALID_SHEET;
    wb->sheets[sheet].show_gridlines = show ? 1 : 0;
    return ARK_OK;
}

ark_err_t ark_hyperlink(ark_wb_t *wb, int sheet, const char *ref,
                        const char *target, int internal, const char *tooltip)
{
#if !ARK_ENABLE_HYPERLINK
    (void)wb; return ARK_ERR_BAD_STATE;
#else
    int r, t, tp;
    if (!wb || !ref || !target) return ARK_ERR_INVALID_ARG;
    if (wb->saved) return ARK_ERR_SAVED;
    if (sheet < 0 || (uint32_t)sheet >= wb->sheet_count) return ARK_ERR_INVALID_SHEET;
    if (wb->hyperlink_count >= ARK_MAX_HYPERLINKS) return ARK_ERR_TABLE_FULL;
    ark_rd_touch(wb, (uint8_t)sheet);
    r = ark_str_add(wb, ref);
    t = ark_str_add(wb, target);
    if (r < 0) return (ark_err_t)(-r);
    if (t < 0) return (ark_err_t)(-t);
    tp = tooltip ? ark_str_add(wb, tooltip) : -1;
    {
        ark_hyperlink_t *h = &wb->hyperlinks[wb->hyperlink_count++];
        h->sheet = (uint8_t)sheet;
        h->internal = internal ? 1 : 0;
        h->ref_idx = (uint16_t)r;
        h->target_idx = (uint32_t)(uint16_t)t;
        h->tooltip_idx = (tp >= 0) ? (uint32_t)(uint16_t)tp : 0xFFFFFFFFu;
    }
    return ARK_OK;
#endif
}

ark_err_t ark_add_validation(ark_wb_t *wb, int sheet, const char *range,
                             int type, int op, const char *f1, const char *f2,
                             int allow_blank, const char *prompt,
                             const char *err)
{
#if !ARK_ENABLE_VALIDATION
    (void)wb; return ARK_ERR_BAD_STATE;
#else
    int ri, f1i, f2i, pi, ei;
    ark_validation_t *v;
    if (!wb || !range || (type != ARK_VAL_NONE && !f1)) return ARK_ERR_INVALID_ARG;
    if (wb->saved) return ARK_ERR_SAVED;
    if (sheet < 0 || (uint32_t)sheet >= wb->sheet_count) return ARK_ERR_INVALID_SHEET;
    if (wb->validation_count >= ARK_MAX_VALIDATIONS) return ARK_ERR_TABLE_FULL;
    ark_rd_touch(wb, (uint8_t)sheet);
    ri  = ark_str_add(wb, range);       if (ri  < 0) return (ark_err_t)(-ri);
    f1i = f1 ? ark_str_add(wb, f1) : -1; if (f1 && f1i < 0) return (ark_err_t)(-f1i);
    f2i = f2 ? ark_str_add(wb, f2) : -1; if (f2 && f2i < 0) return (ark_err_t)(-f2i);
    pi  = prompt ? ark_str_add(wb, prompt) : -1;
    ei  = err ? ark_str_add(wb, err) : -1;
    if (prompt && pi < 0) return (ark_err_t)(-pi);
    if (err && ei < 0) return (ark_err_t)(-ei);
    v = &wb->validations[wb->validation_count++];
    memset(v, 0, sizeof(*v));
    v->sheet = (uint8_t)sheet;
    v->type = (uint8_t)type;
    v->op = (uint8_t)op;
    v->allow_blank = allow_blank ? 1 : 0;
    v->ref_idx = (uint16_t)ri;
    v->f1_idx = f1i >= 0 ? (uint16_t)f1i : ARK_IDX_NONE;
    v->f2_idx = f2i >= 0 ? (uint16_t)f2i : ARK_IDX_NONE;
    v->prompt_idx = pi >= 0 ? (uint16_t)pi : ARK_IDX_NONE;
    v->error_idx = ei >= 0 ? (uint16_t)ei : ARK_IDX_NONE;
    return ARK_OK;
#endif
}

ark_err_t ark_add_condfmt(ark_wb_t *wb, int sheet, const char *range, int type,
                          int op, const char *f1, const char *f2,
                          uint32_t fill_argb, uint32_t font_argb, int bold)
{
#if !ARK_ENABLE_CONDFMT
    (void)wb; return ARK_ERR_BAD_STATE;
#else
    int ri, f1i, f2i;
    ark_condfmt_t *cf;
    if (!wb || !range || !f1) return ARK_ERR_INVALID_ARG;
    if (wb->saved) return ARK_ERR_SAVED;
    if (sheet < 0 || (uint32_t)sheet >= wb->sheet_count) return ARK_ERR_INVALID_SHEET;
    if (wb->condfmt_count >= ARK_MAX_CONDFMTS) return ARK_ERR_TABLE_FULL;
    ark_rd_touch(wb, (uint8_t)sheet);
    ri  = ark_str_add(wb, range);  if (ri < 0)  return (ark_err_t)(-ri);
    f1i = ark_str_add(wb, f1);     if (f1i < 0) return (ark_err_t)(-f1i);
    f2i = f2 ? ark_str_add(wb, f2) : -1;
    if (f2 && f2i < 0) return (ark_err_t)(-f2i);
    cf = &wb->condfmts[wb->condfmt_count++];
    memset(cf, 0, sizeof(*cf));
    cf->sheet = (uint8_t)sheet;
    cf->type = (uint8_t)type;
    cf->op = (uint8_t)op;
    cf->priority = (uint8_t)wb->condfmt_count;
    cf->ref_idx = (uint16_t)ri;
    cf->f1_idx = (uint16_t)f1i;
    cf->f2_idx = f2i >= 0 ? (uint16_t)f2i : ARK_IDX_NONE;
    cf->fill_color = fill_argb;
    cf->font_color = font_argb;
    cf->bold = bold ? 1 : 0;
    return ARK_OK;
#endif
}

ark_err_t ark_named_range(ark_wb_t *wb, const char *name, const char *refers_to)
{
    int ri;
    ark_namedref_t *n;
    if (!wb || !name || !refers_to) return ARK_ERR_INVALID_ARG;
    if (wb->saved) return ARK_ERR_SAVED;
    if (strlen(name) >= sizeof(n->name)) return ARK_ERR_INVALID_NAME;
    if (wb->name_count >= ARK_MAX_NAMES) return ARK_ERR_TABLE_FULL;
    ri = ark_str_add(wb, refers_to);
    if (ri < 0) return (ark_err_t)(-ri);
    n = &wb->names[wb->name_count++];
    memset(n, 0, sizeof(*n));
    memcpy(n->name, name, strlen(name));
    n->refers_idx = (uint16_t)ri;
    return ARK_OK;
}

ark_err_t ark_page_setup(ark_wb_t *wb, int sheet, const ark_page_setup_t *ps)
{
    if (!wb || !ps) return ARK_ERR_INVALID_ARG;
    if (wb->saved) return ARK_ERR_SAVED;
    if (sheet < 0 || (uint32_t)sheet >= wb->sheet_count) return ARK_ERR_INVALID_SHEET;
    wb->sheets[sheet].page = *ps;
    ark_rd_touch(wb, (uint8_t)sheet);
    return ARK_OK;
}
ark_err_t ark_print_area(ark_wb_t *wb, int sheet, const char *area)
{
    int ai;
    if (!wb || !area) return ARK_ERR_INVALID_ARG;
    if (sheet < 0 || (uint32_t)sheet >= wb->sheet_count) return ARK_ERR_INVALID_SHEET;
    ai = ark_str_add(wb, area);
    ark_rd_touch(wb, (uint8_t)sheet);
    if (ai < 0) return (ark_err_t)(-ai);
    wb->sheets[sheet].print_area_idx = (uint16_t)ai;
    return ARK_OK;
}
ark_err_t ark_header_footer(ark_wb_t *wb, int sheet, const char *header,
                            const char *footer)
{
    int hi, fi;
    if (!wb) return ARK_ERR_INVALID_ARG;
    if (sheet < 0 || (uint32_t)sheet >= wb->sheet_count) return ARK_ERR_INVALID_SHEET;
    hi = header ? ark_str_add(wb, header) : -1;
    ark_rd_touch(wb, (uint8_t)sheet);
    fi = footer ? ark_str_add(wb, footer) : -1;
    if (header && hi < 0) return (ark_err_t)(-hi);
    if (footer && fi < 0) return (ark_err_t)(-fi);
    wb->sheets[sheet].header_idx = hi >= 0 ? (uint16_t)hi : ARK_IDX_NONE;
    wb->sheets[sheet].footer_idx = fi >= 0 ? (uint16_t)fi : ARK_IDX_NONE;
    return ARK_OK;
}
ark_err_t ark_row_break(ark_wb_t *wb, int sheet, uint32_t row)
{
    if (!wb) return ARK_ERR_INVALID_ARG;
    if (wb->saved) return ARK_ERR_SAVED;
    if (sheet < 0 || (uint32_t)sheet >= wb->sheet_count) return ARK_ERR_INVALID_SHEET;
    if (row < 1 || row > 65535) return ARK_ERR_INVALID_ROW;
    wb->sheets[sheet].break_rows = (uint16_t)row;
    ark_rd_touch(wb, (uint8_t)sheet);
    return ARK_OK;
}

/* ============================================================
 * 自定义样式
 * ========================================================== */
int ark_wb_add_style(ark_wb_t *wb, const ark_style_def_t *def)
{
    uint8_t f, fl, b;
    uint16_t nf = 0;
    ark_xf_rec_t xf;
    uint8_t bl, br, bt, bb;
    if (!wb || !def) return ARK_ERR_INVALID_ARG;
    if (wb->saved) return ARK_ERR_SAVED;
    if (s_st.nxfs >= ARK_MAX_STYLES) return ARK_ERR_TOO_MANY_STYLES;
    ark_rd_touch_styles(wb);

    f = st_add_font(def->font_name, def->font_size, def->font_color,
                    def->bold, def->italic, def->underline);
    fl = def->fill_color ? st_add_fill(def->fill_color) : 0;
    bl = def->border_all ? def->border_all : def->border_left;
    br = def->border_all ? def->border_all : def->border_right;
    bt = def->border_all ? def->border_all : def->border_top;
    bb = def->border_all ? def->border_all : def->border_bottom;
    {
        uint8_t dl = def->diag_style;
        uint8_t dup = def->diag_up, ddown = def->diag_down;
        b = (bl || br || bt || bb || (dl && (dup || ddown)))
            ? st_add_border2(bl, br, bt, bb, dl, dup, ddown, def->border_color)
            : 0;
    }
    if (def->custom_fmt) nf = st_add_numfmt(def->custom_fmt);
    else if (def->numfmt_id) nf = def->numfmt_id;

    memset(&xf, 0, sizeof(xf));
    xf.font = f; xf.fill = fl; xf.border = b; xf.numfmt = nf;
    xf.halign = def->halign; xf.valign = def->valign;
    xf.wrap = def->wrap_text; xf.indent = def->indent;
    return (int)st_add_xf(&xf);
}

/* ============================================================
 * 保存: 排序
 * ========================================================== */
static uint32_t s_order[ARK_MAX_CELLS];   /* 排序后下标 */

static int cell_less(const ark_wb_t *wb, uint32_t a, uint32_t b)
{
    const ark_cell_t *x = &wb->cells[a], *y = &wb->cells[b];
    if (x->sheet != y->sheet) return x->sheet < y->sheet;
    if (x->row != y->row)   return x->row < y->row;
    return x->col < y->col;
}
static void cell_sort(ark_wb_t *wb)
{
    uint32_t n = wb->cell_count, i;
    /* 堆排序, O(n log n), 无递归无额外内存 */
    for (i = 0; i < n; i++) s_order[i] = i;
    if (n < 2) return;
    {   /* build heap */
        int64_t start = (int64_t)(n / 2 - 1);
        while (start >= 0) {
            uint32_t root = (uint32_t)start;
            for (;;) {
                uint32_t child = 2 * root + 1;
                if (child >= n) break;
                if (child + 1 < n &&
                    cell_less(wb, s_order[child], s_order[child + 1]))
                    child++;
                if (!cell_less(wb, s_order[root], s_order[child])) break;
                { uint32_t t = s_order[root]; s_order[root] = s_order[child];
                  s_order[child] = t; }
                root = child;
            }
            start--;
        }
    }
    {   /* sort down */
        uint32_t end = n;
        while (end > 1) {
            uint32_t root = 0;
            end--;
            { uint32_t t = s_order[0]; s_order[0] = s_order[end];
              s_order[end] = t; }
            for (;;) {
                uint32_t child = 2 * root + 1;
                if (child >= end) break;
                if (child + 1 < end &&
                    cell_less(wb, s_order[child], s_order[child + 1]))
                    child++;
                if (!cell_less(wb, s_order[root], s_order[child])) break;
                { uint32_t t = s_order[root]; s_order[root] = s_order[child];
                  s_order[child] = t; }
                root = child;
            }
        }
    }
}

/* ============================================================
 * 保存: XML 文件生成器
 * ========================================================== */
typedef void (*ark_gen_fn)(ark_wb_t *wb, ark_xw_t *x);

static void gen_content_types(ark_wb_t *wb, ark_xw_t *x)
{
    uint8_t i;
    xw_str(x, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
    for (i = 0; i < wb->sheet_count; i++)
        xw_fmt(x, "<Override PartName=\"/xl/worksheets/sheet%d.xml\" "
            "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>",
            i + 1);
    xw_str(x, "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
    if (wb->sst_count)
        xw_str(x, "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
#ifndef ARK_EXCEL_SKIP_DOCPROPS
    xw_str(x, "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>"
        "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
#endif
    xw_str(x, "</Types>");
}

static void gen_root_rels(ark_wb_t *wb, ark_xw_t *x)
{
    (void)wb;
    xw_str(x, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
#ifndef ARK_EXCEL_SKIP_DOCPROPS
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>"
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>"
#endif
        "</Relationships>");
}

static void gen_core(ark_wb_t *wb, ark_xw_t *x)
{
    (void)wb;
    xw_str(x, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
        "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" "
        "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
        "<dc:creator>ark_excel</dc:creator><cp:lastModifiedBy>ark_excel</cp:lastModifiedBy>"
        "<dcterms:created xsi:type=\"dcterms:W3CDTF\">2024-01-01T00:00:00Z</dcterms:created>"
        "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">2024-01-01T00:00:00Z</dcterms:modified>"
        "</cp:coreProperties>");
}

static void gen_app(ark_wb_t *wb, ark_xw_t *x)
{
    xw_str(x, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" "
        "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
        "<Application>ark_excel</Application>");
    xw_fmt(x, "<HeadingPairs><vt:vector size=\"2\" baseType=\"variant\">"
        "<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>"
        "<vt:variant><vt:i4>%d</vt:i4></vt:variant>"
        "</vt:vector></HeadingPairs>", wb->sheet_count);
    xw_fmt(x, "<TitlesOfParts><vt:vector size=\"%d\" baseType=\"lpstr\">",
           wb->sheet_count);
    { uint8_t i;
      for (i = 0; i < wb->sheet_count; i++) {
          xw_str(x, "<vt:lpstr>"); xw_esc(x, wb->sheets[i].name);
          xw_str(x, "</vt:lpstr>");
      } }
    xw_str(x, "</vt:vector></TitlesOfParts></Properties>");
}

static void gen_workbook(ark_wb_t *wb, ark_xw_t *x)
{
    uint8_t i;
    xw_str(x, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheets>");
    for (i = 0; i < wb->sheet_count; i++)
        xw_fmt(x, "<sheet name=\""), xw_esc(x, wb->sheets[i].name),
        xw_fmt(x, "\" sheetId=\"%d\" r:id=\"rId%d\"/>", i + 1, i + 1);
    xw_str(x, "</sheets>");
    {
        int any = (wb->name_count != 0);
        if (!any)
            for (i = 0; i < wb->sheet_count; i++)
                if (wb->sheets[i].print_area_idx != ARK_IDX_NONE) {
                    any = 1; break;
                }
        if (any) {
            xw_str(x, "<definedNames>");
            for (i = 0; i < wb->name_count; i++) {
                xw_str(x, "<definedName name=\"");
                xw_esc(x, wb->names[i].name);
                xw_str(x, "\">");
                xw_esc(x, ark_str_get(wb, wb->names[i].refers_idx));
                xw_str(x, "</definedName>");
            }
            /* 打印区域: _xlnm.Print_Area(带 localSheetId) */
            for (i = 0; i < wb->sheet_count; i++) {
                ark_sheet_t *sh = &wb->sheets[i];
                if (sh->print_area_idx == ARK_IDX_NONE) continue;
                xw_fmt(x, "<definedName name=\"_xlnm.Print_Area\" "
                    "localSheetId=\"%d\">'", i);
                xw_esc(x, sh->name);
                xw_str(x, "'!");
                xw_esc(x, ark_str_get(wb, sh->print_area_idx));
                xw_str(x, "</definedName>");
            }
            xw_str(x, "</definedNames>");
        }
    }
    xw_str(x, "</workbook>");
}

static void gen_workbook_rels(ark_wb_t *wb, ark_xw_t *x)
{
    uint8_t i;
    xw_str(x, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
    for (i = 0; i < wb->sheet_count; i++)
        xw_fmt(x, "<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet%d.xml\"/>",
               i + 1, i + 1);
    if (wb->sst_count)
        xw_fmt(x, "<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>",
               wb->sheet_count + 1);
    xw_fmt(x, "<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>",
           wb->sheet_count + 2);
    xw_str(x, "</Relationships>");
}

static void gen_sst(ark_wb_t *wb, ark_xw_t *x)
{
    uint16_t i;
    xw_fmt(x, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "count=\"%d\" uniqueCount=\"%d\">", wb->sst_count, wb->sst_count);
    for (i = 0; i < wb->str_count; i++) {
        if (!wb->strs[i].is_sst) continue;
        xw_str(x, "<si><t>");
        xw_esc(x, wb->strs[i].text);
        xw_str(x, "</t></si>");
    }
    xw_str(x, "</sst>");
}

static const char *s_align_h[6] = { "general","left","center","right","fill","justify" };
static const char *s_align_v[5] = { 0, "top","center","bottom","justify" };
static const char *s_border_st[8] = { "none","thin","medium","thick",
                                      "dashed","dotted","double","hair" };

static void gen_styles(ark_wb_t *wb, ark_xw_t *x)
{
    uint16_t i;
    (void)wb;
    xw_str(x, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
    /* numFmts */
    if (s_st.nnumfmts) {
        xw_fmt(x, "<numFmts count=\"%d\">", s_st.nnumfmts);
        for (i = 0; i < s_st.nnumfmts; i++) {
            xw_fmt(x, "<numFmt numFmtId=\"%d\" formatCode=\"",
                   s_st.numfmts[i].id);
            xw_esc(x, s_st.numfmts[i].fmt); xw_str(x, "\"/>");
        }
        xw_str(x, "</numFmts>");
    }
    /* fonts */
    xw_fmt(x, "<fonts count=\"%d\">", s_st.nfonts);
    for (i = 0; i < s_st.nfonts; i++) {
        const ark_font_rec_t *f = &s_st.fonts[i];
        xw_str(x, "<font>");
        if (f->size_x100 % 100)
            xw_fmt(x, "<sz val=\"%u.%02u\"/>", f->size_x100/100, f->size_x100%100);
        else
            xw_fmt(x, "<sz val=\"%u\"/>", f->size_x100 / 100);
        if (f->color)
            xw_fmt(x, "<color rgb=\"%08X\"/>", f->color);
        if (f->bold)     xw_str(x, "<b/>");
        if (f->italic)   xw_str(x, "<i/>");
        if (f->underline) xw_str(x, "<u val=\"single\"/>");
        xw_str(x, "<name val=\""), xw_esc(x, f->name), xw_str(x, "\"/></font>");
    }
    xw_str(x, "</fonts>");
    /* fills: 0=none 1=gray125 必需 */
    xw_fmt(x, "<fills count=\"%d\">", s_st.nfills);
    xw_str(x, "<fill><patternFill patternType=\"none\"/></fill>");
    xw_str(x, "<fill><patternFill patternType=\"gray125\"/></fill>");
    for (i = 2; i < s_st.nfills; i++)
        xw_fmt(x, "<fill><patternFill patternType=\"solid\">"
            "<fgColor rgb=\"%08X\"/><bgColor indexed=\"64\"/></patternFill></fill>",
            s_st.fills[i].argb);
    xw_str(x, "</fills>");
    /* borders */
    xw_fmt(x, "<borders count=\"%d\">", s_st.nborders);
    for (i = 0; i < s_st.nborders; i++) {
        const ark_border_rec_t *b = &s_st.borders[i];
        if (b->dup && b->ddown) xw_str(x, "<border diagonalUp=\"1\" diagonalDown=\"1\">");
        else if (b->dup)        xw_str(x, "<border diagonalUp=\"1\">");
        else if (b->ddown)      xw_str(x, "<border diagonalDown=\"1\">");
        else                    xw_str(x, "<border>");
        if (b->l) { xw_str(x, "<left style=\""); xw_str(x, s_border_st[b->l]);
                    if (b->color) xw_fmt(x, "\"><color rgb=\"%08X\"/></left>",
                                         b->color);
                    else xw_str(x, "\"/>"); }
        else xw_str(x, "<left/>");
        if (b->r) { xw_str(x, "<right style=\""); xw_str(x, s_border_st[b->r]);
                    if (b->color) xw_fmt(x, "\"><color rgb=\"%08X\"/></right>",
                                         b->color);
                    else xw_str(x, "\"/>"); }
        else xw_str(x, "<right/>");
        if (b->t) { xw_str(x, "<top style=\""); xw_str(x, s_border_st[b->t]);
                    if (b->color) xw_fmt(x, "\"><color rgb=\"%08X\"/></top>",
                                         b->color);
                    else xw_str(x, "\"/>"); }
        else xw_str(x, "<top/>");
        if (b->b) { xw_str(x, "<bottom style=\""); xw_str(x, s_border_st[b->b]);
                    if (b->color) xw_fmt(x, "\"><color rgb=\"%08X\"/></bottom>",
                                         b->color);
                    else xw_str(x, "\"/>"); }
        else xw_str(x, "<bottom/>");
        if (b->dl) { xw_str(x, "<diagonal style=\""); xw_str(x, s_border_st[b->dl]);
                     xw_ch(x, '\"'); xw_str(x, "/></border>"); }
        else xw_str(x, "<diagonal/></border>");
    }
    xw_str(x, "</borders>");
    /* cellStyleXfs + cellXfs */
    xw_str(x, "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" "
        "fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
    xw_fmt(x, "<cellXfs count=\"%d\">", s_st.nxfs);
    for (i = 0; i < s_st.nxfs; i++) {
        const ark_xf_rec_t *xf = &s_st.xfs[i];
        int al = xf->halign || xf->valign || xf->wrap || xf->indent;
        xw_fmt(x, "<xf numFmtId=\"%d\" fontId=\"%d\" fillId=\"%d\" borderId=\"%d\""
            " applyNumberFormat=\"%d\" applyFont=\"1\" applyFill=\"%d\" "
            "applyBorder=\"%d\" applyAlignment=\"%d\"",
            xf->numfmt, xf->font, xf->fill, xf->border,
            xf->numfmt ? 1 : 0, xf->fill ? 1 : 0, xf->border ? 1 : 0,
            al ? 1 : 0);
        if (al) {
            xw_str(x, ">");
            if (xf->halign || xf->valign || xf->wrap || xf->indent) {
                xw_str(x, "<alignment");
                if (xf->halign) { xw_str(x, " horizontal=\"");
                                  xw_str(x, s_align_h[xf->halign]); xw_ch(x, '\"'); }
                if (xf->valign) { xw_str(x, " vertical=\"");
                                  xw_str(x, s_align_v[xf->valign]); xw_ch(x, '\"'); }
                if (xf->wrap)   xw_str(x, " wrapText=\"1\"");
                if (xf->indent) xw_fmt(x, " indent=\"%d\"", xf->indent);
                xw_str(x, "/>");
            }
            xw_str(x, "</xf>");
        } else {
            xw_str(x, "/>");
        }
    }
    xw_str(x, "</cellXfs>");
    /* cellStyles 必须在 dxfs 之前(schema 顺序) */
    xw_str(x, "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" "
        "builtinId=\"0\"/></cellStyles>");
#if ARK_ENABLE_CONDFMT
    /* dxfs: 条件格式差异样式 */
    if (wb->condfmt_count) {
        xw_fmt(x, "<dxfs count=\"%d\">", wb->condfmt_count);
        for (i = 0; i < wb->condfmt_count; i++) {
            const ark_condfmt_t *cf = &wb->condfmts[i];
            xw_str(x, "<dxf><font>");
            if (cf->bold) xw_str(x, "<b/>");
            if (cf->font_color) xw_fmt(x, "<color rgb=\"%08X\"/>", cf->font_color);
            xw_str(x, "</font>");
            if (cf->fill_color)
                xw_fmt(x, "<fill><patternFill><bgColor rgb=\"%08X\"/>"
                    "</patternFill></fill>", cf->fill_color);
            xw_str(x, "</dxf>");
        }
        xw_str(x, "</dxfs>");
    }
#endif
    xw_str(x, "</styleSheet>");
}

/* ---------------- worksheet ---------------- */
static void emit_cell(ark_wb_t *wb, ark_xw_t *x, const ark_cell_t *c)
{
    char colbuf[5];
    ark_col_to_letters(c->col, colbuf);
    xw_str(x, "<c r=\"");
    xw_str(x, colbuf);
    xw_u32(x, c->row);
    xw_fmt(x, "\" s=\"%d\"", c->style);
    switch (c->type) {
    case ARK_CELL_SSTR:
        xw_fmt(x, " t=\"s\"><v>%d</v>", c->sidx);
        break;
    case ARK_CELL_ISTR:
        xw_str(x, " t=\"inlineStr\"><is><t>");
        xw_esc(x, ark_str_get(wb, (uint16_t)c->sidx));
        xw_str(x, "</t></is>");
        break;
    case ARK_CELL_NUM:
    case ARK_CELL_DATE:
        xw_str(x, "><v>");
        xw_dbl(x, c->num);
        xw_str(x, "</v>");
        break;
    case ARK_CELL_INT:
        xw_str(x, "><v>");
        xw_i64(x, (int64_t)c->num);
        xw_str(x, "</v>");
        break;
    case ARK_CELL_BOOL:
        xw_fmt(x, " t=\"b\"><v>%d</v>", c->num ? 1 : 0);
        break;
    case ARK_CELL_ERROR:
        xw_str(x, " t=\"e\"><v>");
        xw_esc(x, ark_str_get(wb, (uint16_t)c->sidx));
        xw_str(x, "</v>");
        break;
#if ARK_ENABLE_FORMULA
    case ARK_CELL_FORMULA:
        xw_str(x, "><f>");
        xw_esc(x, ark_str_get(wb, (uint16_t)c->sidx));
        xw_str(x, "</f>");
        break;
#endif
    default:
        xw_str(x, ">");   /* blank: 仅样式, 但开标签必须闭合 */
        break;
    }
    xw_str(x, "</c>");
}

static uint16_t s_dxf_base;   /* open 模式: 新 dxfId 的模板偏移 */

static void gen_sheet(ark_wb_t *wb, ark_xw_t *x, uint8_t si)
{
    ark_sheet_t *sh = &wb->sheets[si];
    uint32_t k;
    uint32_t minr = 0xFFFFFFFFu, maxr = 0, minc = 0xFFFF, maxc = 0;
    uint8_t first = 1;
    uint32_t cur_row = 0;
    uint16_t i;

    /* dimension */
    for (k = 0; k < wb->cell_count; k++) {
        const ark_cell_t *c = &wb->cells[s_order[k]];
        if (c->sheet != si) continue;
        if (c->row < minr) minr = c->row;
        if (c->row > maxr) maxr = c->row;
        if (c->col < minc) minc = c->col;
        if (c->col > maxc) maxc = c->col;
    }

    xw_str(x, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
    if (sh->tab_color_set || sh->page.fit_to_page) {
        /* sheetPr 必须是 worksheet 第一个子元素 */
        xw_str(x, "<sheetPr>");
        if (sh->tab_color_set)
            xw_fmt(x, "<tabColor rgb=\"%08X\"/>", sh->tab_color);
        if (sh->page.fit_to_page)
            xw_str(x, "<pageSetUpPr fitToPage=\"1\"/>");
        xw_str(x, "</sheetPr>");
    }
    if (minr != 0xFFFFFFFFu) {
        char c1[5], c2[5];
        ark_col_to_letters((uint16_t)minc, c1);
        ark_col_to_letters((uint16_t)maxc, c2);
        xw_fmt(x, "<dimension ref=\"%s%d:%s%d\"/>", c1, minr, c2, maxr);
    } else {
        xw_str(x, "<dimension ref=\"A1\"/>");
    }
    /* sheetViews: 冻结窗格 / RTL / 网格线 */
    xw_str(x, "<sheetViews><sheetView");
    if (sh->rtl) xw_str(x, " rightToLeft=\"1\"");
    if (!sh->show_gridlines) xw_str(x, " showGridLines=\"0\"");
    xw_str(x, " workbookViewId=\"0\"");
    if (sh->freeze_rows || sh->freeze_cols) {
        char fc[5];
        xw_str(x, ">");
        if (sh->freeze_rows && !sh->freeze_cols)
            xw_fmt(x, "<pane ySplit=\"%d\" topLeftCell=\"A%d\" "
                "activePane=\"bottomLeft\" state=\"frozen\"/>",
                sh->freeze_rows, sh->freeze_rows + 1);
        else if (sh->freeze_cols && !sh->freeze_rows)
            xw_fmt(x, "<pane xSplit=\"%d\" topLeftCell=\"%s1\" "
                "activePane=\"topRight\" state=\"frozen\"/>",
                sh->freeze_cols,
                (ark_col_to_letters((uint16_t)(sh->freeze_cols + 1), fc), fc));
        else
            xw_fmt(x, "<pane xSplit=\"%d\" ySplit=\"%d\" topLeftCell=\"%s%d\" "
                "activePane=\"bottomRight\" state=\"frozen\"/>",
                sh->freeze_cols, sh->freeze_rows,
                (ark_col_to_letters((uint16_t)(sh->freeze_cols + 1), fc), fc),
                sh->freeze_rows + 1);
        xw_str(x, "</sheetView>");
    } else {
        xw_str(x, "/>");
    }
    xw_str(x, "</sheetViews>");
    xw_str(x, "<sheetFormatPr defaultRowHeight=\"15\"/>");
    /* cols */
    first = 1;
    for (i = 0; i < wb->coldef_count; i++) {
        const ark_coldef_t *cd = &wb->coldefs[i];
        if (cd->sheet != si) continue;
        if (first) { xw_str(x, "<cols>"); first = 0; }
        xw_fmt(x, "<col min=\"%d\" max=\"%d\"", cd->first, cd->last);
        if (cd->width_x100)
            xw_fmt(x, " width=\"%u.%02u\" customWidth=\"1\"",
                   cd->width_x100 / 100, cd->width_x100 % 100);
        if (cd->hidden) xw_str(x, " hidden=\"1\"");
        xw_str(x, "/>");
    }
    if (!first) xw_str(x, "</cols>");

    /* sheetData */
    xw_str(x, "<sheetData>");
    cur_row = 0;
    /* 无单元格但定义了行高/隐藏的行也要输出 <row/>(OOXML 要求行号递增,
     * 与单元格行归并输出: 先按行号排序本 sheet 的 rowdefs) */
    {
        uint16_t nr = 0, ri, rj;
        static uint16_t rd_idx[ARK_MAX_ROWDEFS];
        for (ri = 0; ri < wb->rowdef_count; ri++)
            if (wb->rowdefs[ri].sheet == (uint8_t)si) rd_idx[nr++] = ri;
        for (ri = 1; ri < nr; ri++) {          /* 插入排序(项数少) */
            uint16_t key = rd_idx[ri];
            for (rj = ri; rj > 0 &&
                 wb->rowdefs[rd_idx[rj-1]].row > wb->rowdefs[key].row; rj--)
                rd_idx[rj] = rd_idx[rj-1];
            rd_idx[rj] = key;
        }
        uint32_t last_out_row = 0;   /* 已输出的最大行号(含空行) */
        for (k = 0; k <= wb->cell_count; k++) {
            const ark_cell_t *c = NULL;
            uint32_t next_row;
            if (k < wb->cell_count) {
                c = &wb->cells[s_order[k]];
                if (c->sheet != si) continue;
            }
            if (!c) { next_row = 0xFFFFFFFFu; }
            else if (cur_row && c->row == cur_row) next_row = cur_row;
            else next_row = c->row;
            /* 先冲刷行号小于下一单元格行的空 rowdef 行 */
            for (ri = 0; ri < nr; ri++) {
                ark_rowdef_t *rd = &wb->rowdefs[rd_idx[ri]];
                if (rd->row >= next_row) break;
                if (rd->row <= last_out_row) continue;    /* 已随内容行输出 */
                if (cur_row) { xw_str(x, "</row>"); cur_row = 0; }
                if (rd->row > last_out_row) last_out_row = rd->row;
                xw_fmt(x, "<row r=\"%u\"", rd->row);
                if (rd->height_x20)
                    xw_fmt(x, " ht=\"%u.%02u\" customHeight=\"1\"",
                           rd->height_x20 / 20, (rd->height_x20 % 20) * 5);
                if (rd->hidden) xw_str(x, " hidden=\"1\"/>");
                else xw_str(x, "/>");
            }
            if (!c) {
                if (cur_row) { xw_str(x, "</row>"); cur_row = 0; }
                break;
            }
            if (cur_row && c->row != cur_row) {
                xw_str(x, "</row>");
                cur_row = 0;
            }
            if (!cur_row) {
            cur_row = c->row;
            if (cur_row > last_out_row) last_out_row = cur_row;
            xw_fmt(x, "<row r=\"%d\"", cur_row);
            /* 行属性 */
            {
                ark_rowdef_t *rd = rowdef_find(wb, si, cur_row);
                if (rd) {
                    if (rd->height_x20)
                        xw_fmt(x, " ht=\"%u.%02u\" customHeight=\"1\"",
                               rd->height_x20 / 20, (rd->height_x20 % 20) * 5);
                    if (rd->hidden) xw_str(x, " hidden=\"1\"");
                }
            }
            xw_str(x, ">");
        }
            emit_cell(wb, x, c);
    }   /* for k */
    }   /* 空行归并块 */
    xw_str(x, "</sheetData>");

    /* autoFilter(必须在 mergeCells 之前) */
    if (sh->autofilter_set) {
        char c1[5], c2[5];
        ark_col_to_letters(sh->filter_c1, c1);
        ark_col_to_letters(sh->filter_c2, c2);
        xw_fmt(x, "<autoFilter ref=\"%s%d:%s%d\"/>",
               c1, sh->filter_r1, c2, sh->filter_r2);
    }
    /* mergeCells */
    {
        uint16_t cnt = 0;
        for (i = 0; i < wb->merge_count; i++)
            if (wb->merges[i].sheet == si) cnt++;
        if (cnt) {
            xw_fmt(x, "<mergeCells count=\"%d\">", cnt);
            for (i = 0; i < wb->merge_count; i++) {
                const ark_merge_t *m = &wb->merges[i];
                char c1[5], c2[5];
                if (m->sheet != si) continue;
                ark_col_to_letters(m->c1, c1);
                ark_col_to_letters(m->c2, c2);
                xw_fmt(x, "<mergeCell ref=\"%s%d:%s%d\"/>",
                       c1, m->r1, c2, m->r2);
            }
            xw_str(x, "</mergeCells>");
        }
    }
    /* conditionalFormatting */
    for (i = 0; i < wb->condfmt_count; i++) {
        const ark_condfmt_t *cf = &wb->condfmts[i];
        static const char *cfop[8] = { "between","notBetween","equal","notEqual",
                                       "greaterThan","lessThan","greaterThanOrEqual",
                                       "lessThanOrEqual" };
        if (cf->sheet != si) continue;
        xw_str(x, "<conditionalFormatting sqref=\"");
        xw_esc(x, ark_str_get(wb, cf->ref_idx));
        xw_fmt(x, "\"><cfRule type=\"%s\" dxfId=\"%d\" priority=\"%d\"",
               cf->type == ARK_CF_CELLIS ? "cellIs" : "expression",
               i + s_dxf_base, cf->priority);
        if (cf->type == ARK_CF_CELLIS)
            xw_fmt(x, " operator=\"%s\"", cfop[cf->op & 7]);
        xw_str(x, ">");
        xw_str(x, "<formula>");
        xw_esc(x, ark_str_get(wb, cf->f1_idx));
        xw_str(x, "</formula>");
        if (cf->type == ARK_CF_CELLIS && cf->op <= ARK_OP_NOT_BETWEEN &&
            cf->f2_idx != ARK_IDX_NONE) {
            xw_str(x, "<formula>");
            xw_esc(x, ark_str_get(wb, cf->f2_idx));
            xw_str(x, "</formula>");
        }
        xw_str(x, "</cfRule></conditionalFormatting>");
    }
    /* dataValidations */
    {
        uint16_t cnt = 0;
        static const char *vtype[7] = { "none","whole","decimal","list",
                                        "textLength","date","time" };
        static const char *vop[8] = { "between","notBetween","equal","notEqual",
                                      "greaterThan","lessThan","greaterThanOrEqual",
                                      "lessThanOrEqual" };
        for (i = 0; i < wb->validation_count; i++)
            if (wb->validations[i].sheet == si) cnt++;
        if (cnt) {
            xw_fmt(x, "<dataValidations count=\"%d\">", cnt);
            for (i = 0; i < wb->validation_count; i++) {
                const ark_validation_t *v = &wb->validations[i];
                if (v->sheet != si) continue;
                xw_str(x, "<dataValidation type=\"");
                xw_str(x, vtype[v->type % 7]);
                if (v->type != ARK_VAL_LIST)
                    xw_fmt(x, "\" operator=\"%s", vop[v->op & 7]);
                xw_fmt(x, "\" allowBlank=\"%d\"", v->allow_blank);
                if (v->prompt_idx != ARK_IDX_NONE) {
                    xw_str(x, " prompt=\"");
                    xw_esc(x, ark_str_get(wb, v->prompt_idx));
                    xw_ch(x, '\"');
                }
                if (v->error_idx != ARK_IDX_NONE) {
                    xw_str(x, " error=\"");
                    xw_esc(x, ark_str_get(wb, v->error_idx));
                    xw_ch(x, '\"');
                }
                xw_str(x, " showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"");
                xw_esc(x, ark_str_get(wb, v->ref_idx));
                xw_ch(x, '"');
                xw_str(x, ">");   /* sqref 是属性, 不能作子元素 */
                if (v->f1_idx != ARK_IDX_NONE) {
                    xw_str(x, "<formula1>");
                    xw_esc(x, ark_str_get(wb, v->f1_idx));
                    xw_str(x, "</formula1>");
                }
                if (v->f2_idx != ARK_IDX_NONE) {
                    xw_str(x, "<formula2>");
                    xw_esc(x, ark_str_get(wb, v->f2_idx));
                    xw_str(x, "</formula2>");
                }
                xw_str(x, "</dataValidation>");
            }
            xw_str(x, "</dataValidations>");
        }
    }
    /* hyperlinks */
    {
        uint16_t cnt = 0;
        for (i = 0; i < wb->hyperlink_count; i++)
            if (wb->hyperlinks[i].sheet == si) cnt++;
        if (cnt) {
            xw_fmt(x, "<hyperlinks>");
            for (i = 0; i < wb->hyperlink_count; i++) {
                const ark_hyperlink_t *h = &wb->hyperlinks[i];
                if (h->sheet != si) continue;
                xw_fmt(x, "<hyperlink ref=\"%s\" ", ark_str_get(wb, h->ref_idx));
                if (h->internal) {
                    const char *loc = ark_str_get(wb, (uint16_t)h->target_idx);
                    if (loc[0] == '#') loc++;   /* location 属性不带 # */
                    xw_fmt(x, "location=\""); xw_esc(x, loc); xw_ch(x, '"');
                } else {
                    xw_fmt(x, "r:id=\"rId%d\"", i + 1);
                }
                if (h->tooltip_idx != 0xFFFFFFFFu) {
                    xw_str(x, " tooltip=\"");
                    xw_esc(x, ark_str_get(wb, (uint16_t)h->tooltip_idx));
                    xw_ch(x, '\"');
                }
                xw_str(x, "/>");
            }
            xw_str(x, "</hyperlinks>");
        }
    }
    /* 打印: pageMargins + pageSetup + headerFooter + rowBreaks */
    {
        const ark_page_setup_t *p = &sh->page;
        uint32_t ml = p->margin_left   == 0xFFFF ? 700  : p->margin_left;
        uint32_t mr = p->margin_right  == 0xFFFF ? 700  : p->margin_right;
        uint32_t mt = p->margin_top    == 0xFFFF ? 750  : p->margin_top;
        uint32_t mb = p->margin_bottom == 0xFFFF ? 750  : p->margin_bottom;
        uint32_t mh = p->margin_header == 0xFFFF ? 300  : p->margin_header;
        uint32_t mf = p->margin_footer == 0xFFFF ? 300  : p->margin_footer;
        /* printOptions 必须位于 pageMargins 之前(schema 顺序) */
        if (p->fit_to_page)
            xw_str(x, "<printOptions horizontalCentered=\"1\"/>");
        xw_fmt(x, "<pageMargins left=\"%u.%03u\" right=\"%u.%03u\" "
            "top=\"%u.%03u\" bottom=\"%u.%03u\" header=\"%u.%03u\" "
            "footer=\"%u.%03u\"/>",
            ml / 1000, ml % 1000, mr / 1000, mr % 1000,
            mt / 1000, mt % 1000, mb / 1000, mb % 1000,
            mh / 1000, mh % 1000, mf / 1000, mf % 1000);
        xw_fmt(x, "<pageSetup paperSize=\"%d\" orientation=\"%s\"",
               p->paper, p->orientation == ARK_LANDSCAPE ? "landscape"
                                                          : "portrait");
        if (p->fit_to_page)
            xw_fmt(x, " fitToWidth=\"%d\" fitToHeight=\"%d\"/>",
                   p->fit_width ? p->fit_width : 1, p->fit_height);
        else
            xw_fmt(x, " scale=\"%d\"/>", p->scale ? p->scale : 100);
        if (sh->header_idx != ARK_IDX_NONE) {
            xw_str(x, "<headerFooter><oddHeader>");
            xw_esc(x, ark_str_get(wb, sh->header_idx));
            xw_str(x, "</oddHeader>");
            if (sh->footer_idx != ARK_IDX_NONE) {
                xw_str(x, "<oddFooter>");
                xw_esc(x, ark_str_get(wb, sh->footer_idx));
                xw_str(x, "</oddFooter>");
            }
            xw_str(x, "</headerFooter>");
        } else if (sh->footer_idx != ARK_IDX_NONE) {
            xw_str(x, "<headerFooter><oddFooter>");
            xw_esc(x, ark_str_get(wb, sh->footer_idx));
            xw_str(x, "</oddFooter></headerFooter>");
        }
        if (sh->break_rows)
            xw_fmt(x, "<rowBreaks count=\"1\" manualBreakCount=\"1\">"
                "<brk id=\"%d\" max=\"16383\" man=\"1\"/></rowBreaks>",
                sh->break_rows);
    }
    xw_str(x, "</worksheet>");
}

static void gen_sheet_rels(ark_wb_t *wb, ark_xw_t *x, uint8_t si)
{
    uint16_t i;
    xw_str(x, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
    for (i = 0; i < wb->hyperlink_count; i++) {
        const ark_hyperlink_t *h = &wb->hyperlinks[i];
        if (h->sheet != si || h->internal) continue;
        xw_fmt(x, "<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"",
               i + 1);
        xw_esc(x, ark_str_get(wb, (uint16_t)h->target_idx));  /* URL 中 & 等须转义 */
        xw_str(x, "\" TargetMode=\"External\"/>");
    }
    xw_str(x, "</Relationships>");
}

/* ============================================================
 * 保存
 * ========================================================== */
static ark_err_t save_file(ark_wb_t *wb, ark_zip_t *zip, const char *name,
                           ark_gen_fn gen)
{
    ark_xw_t x;
    /* 第一遍: 测量 CRC + 大小 */
    ark_zip_measure_start(zip);
    xw_init(&x, zip, 0);
    gen(wb, &x);
    xw_flush(&x);
    if (x.err) return x.err;
    /* 第二遍: 真实输出 */
    if (ark_zip_open_stream(zip, name) != 0) return ARK_ERR_WRITE;
    xw_init(&x, zip, 1);
    gen(wb, &x);
    xw_flush(&x);
    if (x.err) return x.err;
    if (ark_zip_close_stream(zip) != 0) return ARK_ERR_WRITE;
    return ARK_OK;
}

typedef struct { uint8_t si; } sheet_ctx_t;
/* 闭包不可行, 用带参生成器: */
typedef void (*ark_gen2_fn)(ark_wb_t *wb, ark_xw_t *x, uint8_t si);
static ark_err_t save_sheet_file(ark_wb_t *wb, ark_zip_t *zip, const char *name,
                                 ark_gen2_fn gen, uint8_t si)
{
    ark_xw_t x;
    ark_zip_measure_start(zip);
    xw_init(&x, zip, 0);
    gen(wb, &x, si);
    xw_flush(&x);
    if (x.err) return x.err;
    if (ark_zip_open_stream(zip, name) != 0) return ARK_ERR_WRITE;
    xw_init(&x, zip, 1);
    gen(wb, &x, si);
    xw_flush(&x);
    if (x.err) return x.err;
    if (ark_zip_close_stream(zip) != 0) return ARK_ERR_WRITE;
    return ARK_OK;
}
static void gen_sheet_rels2(ark_wb_t *wb, ark_xw_t *x, uint8_t si)
{ gen_sheet_rels(wb, x, si); }
static void gen_sheet2(ark_wb_t *wb, ark_xw_t *x, uint8_t si)
{ gen_sheet(wb, x, si); }

ark_err_t ark_wb_save(ark_wb_t *wb)
{
    ark_zip_t zip;
    ark_err_t e;
    uint8_t i;
    char name[48];
    if (!wb) return ARK_ERR_INVALID_ARG;
    if (wb->saved) return ARK_ERR_SAVED;
    if (wb->sheet_count == 0) return ARK_ERR_BAD_STATE;
#if ARK_ENABLE_READWRITE
    if (wb->reader) return ark_save_open(wb);   /* 读改写增量保存 */
#endif

    cell_sort(wb);
    ark_zip_begin(&zip, wb->write, wb->ctx);

    if ((e = save_file(wb, &zip, "[Content_Types].xml", gen_content_types)))
        goto fail;
    if ((e = save_file(wb, &zip, "_rels/.rels", gen_root_rels)))
        goto fail;
#ifndef ARK_EXCEL_SKIP_DOCPROPS
    if ((e = save_file(wb, &zip, "docProps/core.xml", gen_core)))
        goto fail;
    if ((e = save_file(wb, &zip, "docProps/app.xml", gen_app)))
        goto fail;
#endif
    if ((e = save_file(wb, &zip, "xl/workbook.xml", gen_workbook)))
        goto fail;
    if ((e = save_file(wb, &zip, "xl/_rels/workbook.xml.rels", gen_workbook_rels)))
        goto fail;
    if ((e = save_file(wb, &zip, "xl/styles.xml", gen_styles)))
        goto fail;
    if (wb->sst_count &&
        (e = save_file(wb, &zip, "xl/sharedStrings.xml", gen_sst)))
        goto fail;
    for (i = 0; i < wb->sheet_count; i++) {
        /* 先输出本表超链接 rels(rId 从 1 起, 与 gen_sheet 中一致) */
        {
            uint16_t h;
            int has_ext = 0;
            for (h = 0; h < wb->hyperlink_count; h++)
                if (wb->hyperlinks[h].sheet == i &&
                    !wb->hyperlinks[h].internal) { has_ext = 1; break; }
            if (has_ext) {
                if (ARK_SNPRINTF(name, sizeof(name),
                                 "xl/worksheets/_rels/sheet%d.xml.rels", i + 1) < 0)
                    goto fail;
                if ((e = save_sheet_file(wb, &zip, name, gen_sheet_rels2, i)))
                    goto fail;
            }
        }
        if (ARK_SNPRINTF(name, sizeof(name), "xl/worksheets/sheet%d.xml",
                         i + 1) < 0)
            goto fail;
        if ((e = save_sheet_file(wb, &zip, name, gen_sheet2, i)))
            goto fail;
    }
    if (ark_zip_end(&zip) != 0) { e = ARK_ERR_WRITE; goto fail; }
    wb->saved = 1;
    return ARK_OK;
fail:
    wb->last_err = e;
    return e;
}

#if ARK_ENABLE_READWRITE
/* ============================================================
 * 读改写: 增量保存 —— 未修改条目原样字节复制, 被改部件重新生成
 * ========================================================== */
ark_err_t ark_i_cell_put_raw(ark_wb_t *wb, int sheet, uint32_t row,
                             uint16_t col, uint8_t type, uint16_t style,
                             uint32_t sidx, double num)
{
    return cell_put(wb, sheet, row, col, type, style, sidx, num);
}

/*! 原样字节复制模板条目(压缩数据不重压缩) */
static ark_err_t copy_raw_entry(ark_zip_t *zip, ark_reader_t *rd,
                                const ark_zentry_t *e)
{
    uint8_t tmp[ARK_IO_CHUNK];
    uint8_t lh[30];
    uint32_t dpos, left;
    if (e->method != 0 && e->method != 8) return ARK_ERR_BAD_STATE;
    if (rd->io.seek(rd->io.ctx, e->local_off) != 0) return ARK_ERR_WRITE;
    if (rd->io.read(rd->io.ctx, lh, 30) != 30) return ARK_ERR_WRITE;
    dpos = e->local_off + 30 + (uint32_t)(lh[26] | (lh[27] << 8))
                          + (uint32_t)(lh[28] | (lh[29] << 8));
    if (ark_zip_open_raw(zip, e->name, e->method, e->crc,
                         e->usize, e->csize) != 0)
        return ARK_ERR_WRITE;
    if (rd->io.seek(rd->io.ctx, dpos) != 0) return ARK_ERR_WRITE;
    left = e->csize;
    while (left) {
        uint32_t n = left > sizeof(tmp) ? (uint32_t)sizeof(tmp) : left;
        if (rd->io.read(rd->io.ctx, tmp, n) != n) return ARK_ERR_WRITE;
        if (ark_zip_write(zip, tmp, n) != 0) return ARK_ERR_WRITE;
        left -= n;
    }
    return ark_zip_close_stream(zip) == 0 ? ARK_OK : ARK_ERR_WRITE;
}

/*! 拷贝 XML 段 [begin,end), 把段内首个 count="N" 替换为 new_count */
static void seg_copy(ark_xw_t *x, const char *begin, const char *end,
                     int new_count)
{
    const char *p = begin;
    while (p < end) {
        if (end - p >= 7 && strncmp(p, "count=\"", 7) == 0) {
            xw_raw(x, "count=\"", 7);
            xw_fmt(x, "%d", new_count);
            p += 7;
            while (p < end && *p != '"') p++;
            if (p < end) { xw_ch(x, '"'); p++; }
        } else { xw_raw(x, p, 1); p++; }
    }
}

static void emit_new_fonts(ark_xw_t *x, uint16_t from)
{
    uint16_t i;
    for (i = from; i < s_st.nfonts; i++) {
        const ark_font_rec_t *f = &s_st.fonts[i];
        xw_str(x, "<font>");
        if (f->size_x100 % 100)
            xw_fmt(x, "<sz val=\"%u.%02u\"/>", f->size_x100/100, f->size_x100%100);
        else
            xw_fmt(x, "<sz val=\"%u\"/>", f->size_x100 / 100);
        if (f->color)     xw_fmt(x, "<color rgb=\"%08X\"/>", f->color);
        if (f->bold)      xw_str(x, "<b/>");
        if (f->italic)    xw_str(x, "<i/>");
        if (f->underline) xw_str(x, "<u val=\"single\"/>");
        xw_str(x, "<name val=\"");
        xw_esc(x, f->name);
        xw_str(x, "\"/></font>");
    }
}
static void emit_new_fills(ark_xw_t *x, uint16_t from)
{
    uint16_t i;
    for (i = from; i < s_st.nfills; i++)
        xw_fmt(x, "<fill><patternFill patternType=\"solid\">"
            "<fgColor rgb=\"%08X\"/><bgColor indexed=\"64\"/></patternFill></fill>",
            s_st.fills[i].argb);
}
static void emit_new_borders(ark_xw_t *x, uint16_t from)
{
    uint16_t i;
    for (i = from; i < s_st.nborders; i++) {
        const ark_border_rec_t *b = &s_st.borders[i];
        if (b->dup && b->ddown) xw_str(x, "<border diagonalUp=\"1\" diagonalDown=\"1\">");
        else if (b->dup)        xw_str(x, "<border diagonalUp=\"1\">");
        else if (b->ddown)      xw_str(x, "<border diagonalDown=\"1\">");
        else                    xw_str(x, "<border>");
        if (b->l) { xw_str(x, "<left style=\""); xw_str(x, s_border_st[b->l]);
                    if (b->color) xw_fmt(x, "\"><color rgb=\"%08X\"/></left>", b->color);
                    else xw_str(x, "\"/>"); }
        else xw_str(x, "<left/>");
        if (b->r) { xw_str(x, "<right style=\""); xw_str(x, s_border_st[b->r]);
                    if (b->color) xw_fmt(x, "\"><color rgb=\"%08X\"/></right>", b->color);
                    else xw_str(x, "\"/>"); }
        else xw_str(x, "<right/>");
        if (b->t) { xw_str(x, "<top style=\""); xw_str(x, s_border_st[b->t]);
                    if (b->color) xw_fmt(x, "\"><color rgb=\"%08X\"/></top>", b->color);
                    else xw_str(x, "\"/>"); }
        else xw_str(x, "<top/>");
        if (b->b) { xw_str(x, "<bottom style=\""); xw_str(x, s_border_st[b->b]);
                    if (b->color) xw_fmt(x, "\"><color rgb=\"%08X\"/></bottom>", b->color);
                    else xw_str(x, "\"/>"); }
        else xw_str(x, "<bottom/>");
        if (b->dl) { xw_str(x, "<diagonal style=\""); xw_str(x, s_border_st[b->dl]);
                     xw_ch(x, '\"'); xw_str(x, "/></border>"); }
        else xw_str(x, "<diagonal/></border>");
    }
}
static void emit_new_xfs(ark_xw_t *x, uint16_t from)
{
    uint16_t i;
    for (i = from; i < s_st.nxfs; i++) {
        const ark_xf_rec_t *xf = &s_st.xfs[i];
        int al = xf->halign || xf->valign || xf->wrap || xf->indent;
        xw_fmt(x, "<xf numFmtId=\"%d\" fontId=\"%d\" fillId=\"%d\" borderId=\"%d\""
            " applyNumberFormat=\"%d\" applyFont=\"1\" applyFill=\"%d\" "
            "applyBorder=\"%d\" applyAlignment=\"%d\" xfId=\"0\"",
            xf->numfmt, xf->font, xf->fill, xf->border,
            xf->numfmt ? 1 : 0, xf->fill ? 1 : 0, xf->border ? 1 : 0,
            al ? 1 : 0);
        if (al) {
            xw_str(x, "><alignment");
            if (xf->halign) { xw_str(x, " horizontal=\"");
                              xw_str(x, s_align_h[xf->halign]); xw_ch(x, '"'); }
            if (xf->valign) { xw_str(x, " vertical=\"");
                              xw_str(x, s_align_v[xf->valign]); xw_ch(x, '"'); }
            if (xf->wrap)   xw_str(x, " wrapText=\"1\"");
            if (xf->indent) xw_fmt(x, " indent=\"%d\"", xf->indent);
            xw_str(x, "/></xf>");
        } else xw_str(x, "/>");
    }
}
static void emit_new_dxf(ark_xw_t *x, const ark_condfmt_t *cf)
{
    xw_str(x, "<dxf><font>");
    if (cf->bold) xw_str(x, "<b/>");
    if (cf->font_color) xw_fmt(x, "<color rgb=\"%08X\"/>", cf->font_color);
    xw_str(x, "</font>");
    if (cf->fill_color)
        xw_fmt(x, "<fill><patternFill><bgColor rgb=\"%08X\"/>"
                  "</patternFill></fill>", cf->fill_color);
    xw_str(x, "</dxf>");
}
/*! 段落定位: 找 <tag ...> 的结束(显式 </tag> 或自闭合 />)
 *  返回值: 段落后续扫描起点; *seg_end = 模板段落输出终点(不含闭合标签)
 *  *need_close = 1 表示模板自闭合, 输出新条目后需补闭合标签 */
static const char *seg_end_of(const char *p, const char *tag,
                              const char **seg_end, int *need_close)
{
    const char *close = strstr(p, "</");
    /* 在 p 后第一个 "</" 且以 tag 开头 */
    while (close) {
        if (strncmp(close + 2, tag, strlen(tag)) == 0 &&
            close[2 + strlen(tag)] == '>')
            break;
        close = strstr(close + 1, "</");
    }
    if (close) {
        *seg_end = close;
        *need_close = 0;
        return close + 2 + strlen(tag) + 1;
    }
    /* 自闭合: 找标签内第一个 '>' */
    {
        const char *gt = strchr(p, '>');
        /* 自闭合 '<tag .../>': 剥掉 '/' 转为开标签, 输出到 gt 前 */
        if (gt > p && gt[-1] == '/') *seg_end = gt - 1;
        else                          *seg_end = gt;
        *need_close = 1;      /* 补闭合标签 */
        return gt + 1;
    }
}

static void emit_new_numfmts(ark_xw_t *x, uint16_t from)
{
    uint16_t i;
    for (i = from; i < s_st.nnumfmts; i++) {
        xw_fmt(x, "<numFmt numFmtId=\"%d\" formatCode=\"", s_st.numfmts[i].id);
        xw_esc(x, s_st.numfmts[i].fmt);
        xw_str(x, "\"/>");
    }
}

/*! 模板 styles.xml 文本 + 追加新增样式 -> 完整 styles.xml(两遍确定性输出) */
static void gen_styles_open(ark_wb_t *wb, ark_xw_t *x)
{
    ark_reader_t *rd = (ark_reader_t *)wb->reader;
    const char *t = (const char *)rd->buf;
    const char *p = t;
    rd->buf[rd->buf_len] = 0;
    while (*p) {
        if (strncmp(p, "<numFmts", 8) == 0) {
            const char *seg_end; int need_close;
            const char *next = seg_end_of(p, "numFmts", &seg_end, &need_close);
            seg_copy(x, p, seg_end, s_st.nnumfmts);
            if (need_close) xw_str(x, ">");   /* 自闭合 count="0/ -> 开标签 */
            emit_new_numfmts(x, rd->tpl_numfmt_cnt);
            xw_str(x, "</numFmts>");
            p = next;
        } else if (strncmp(p, "<fonts", 6) == 0) {
            /* 模板无 numFmts 而新增了自定义格式: 按 schema 在 fonts 前插入 */
            if (!strstr(t, "<numFmts") && s_st.nnumfmts > rd->tpl_numfmt_cnt) {
                xw_fmt(x, "<numFmts count=\"%d\">",
                       s_st.nnumfmts - rd->tpl_numfmt_cnt);
                emit_new_numfmts(x, rd->tpl_numfmt_cnt);
                xw_str(x, "</numFmts>");
            }
            {   const char *seg_end; int need_close;
                const char *next = seg_end_of(p, "fonts", &seg_end, &need_close);
                seg_copy(x, p, seg_end, s_st.nfonts);
                if (need_close) xw_str(x, ">");
                emit_new_fonts(x, rd->tpl_font_cnt);
                xw_str(x, "</fonts>");
                p = next;
            }
        } else if (strncmp(p, "<fills", 6) == 0) {
            const char *seg_end; int need_close;
            const char *next = seg_end_of(p, "fills", &seg_end, &need_close);
            seg_copy(x, p, seg_end, s_st.nfills);
            if (need_close) xw_str(x, ">");
            emit_new_fills(x, rd->tpl_fill_cnt);
            xw_str(x, "</fills>");
            p = next;
        } else if (strncmp(p, "<borders", 8) == 0) {
            const char *seg_end; int need_close;
            const char *next = seg_end_of(p, "borders", &seg_end, &need_close);
            seg_copy(x, p, seg_end, s_st.nborders);
            if (need_close) xw_str(x, ">");
            emit_new_borders(x, rd->tpl_border_cnt);
            xw_str(x, "</borders>");
            p = next;
        } else if (strncmp(p, "<cellXfs", 8) == 0) {
            const char *seg_end; int need_close;
            const char *next = seg_end_of(p, "cellXfs", &seg_end, &need_close);
            seg_copy(x, p, seg_end, s_st.nxfs);
            if (need_close) xw_str(x, ">");
            emit_new_xfs(x, rd->tpl_xf_cnt);
            xw_str(x, "</cellXfs>");
            p = next;
        } else if (strncmp(p, "<dxfs", 5) == 0) {
            const char *seg_end; int need_close;
            const char *next = seg_end_of(p, "dxfs", &seg_end, &need_close);
            seg_copy(x, p, seg_end, rd->tpl_dxf_cnt + wb->condfmt_count);
            if (need_close) xw_str(x, ">");
            {   uint16_t i;
                for (i = 0; i < wb->condfmt_count; i++)
                    emit_new_dxf(x, &wb->condfmts[i]);
            }
            xw_str(x, "</dxfs>");
            p = next;
        } else if (strncmp(p, "</styleSheet>", 13) == 0) {
            if (!strstr(t, "<dxfs") && wb->condfmt_count) {
                xw_fmt(x, "<dxfs count=\"%d\">", wb->condfmt_count);
                {   uint16_t i;
                    for (i = 0; i < wb->condfmt_count; i++)
                        emit_new_dxf(x, &wb->condfmts[i]);
                }
                xw_str(x, "</dxfs>");
            }
            xw_str(x, "</styleSheet>");
            p += 13;
        } else { xw_ch(x, *p); p++; }
    }
}

/*! sheet 部件名 -> 其 rels 部件名: "xl/worksheets/sheet1.xml"
 *  -> "xl/worksheets/_rels/sheet1.xml.rels" */
static void sheet_rels_name(const char *part, char *buf, size_t cap)
{
    const char *slash = strrchr(part, '/');
    size_t dirlen = slash ? (size_t)(slash - part + 1) : 0;
    snprintf(buf, cap, "%.*s_rels/%s.rels", (int)dirlen, part,
             slash ? slash + 1 : part);
}

/*! 判断条目名是否某部件的 rels */
static int is_sheet_rels_of(const char *name, const char *part)
{
    char tmp[80];
    sheet_rels_name(part, tmp, sizeof(tmp));
    return strcmp(name, tmp) == 0;
}

static int sheet_has_ext_links(ark_wb_t *wb, uint8_t si)
{
    uint16_t h;
    for (h = 0; h < wb->hyperlink_count; h++)
        if (wb->hyperlinks[h].sheet == si && !wb->hyperlinks[h].internal)
            return 1;
    return 0;
}

ark_err_t ark_save_open(ark_wb_t *wb)
{
    ark_reader_t *rd = (ark_reader_t *)wb->reader;
    ark_zip_t zip;
    ark_err_t e = ARK_OK;
    uint16_t i;
    char name[96];
    uint8_t regen_ct = 0, regen_rels = 0, regen_wb = 0, regen_wbrels = 0;
    uint8_t regen_core = 0, regen_app = 0;

    cell_sort(wb);
    s_dxf_base = rd->tpl_dxf_cnt;

    /* 前置检查: 新增 sheet 需模板部件名规范 */
    if (wb->sheet_count > rd->tpl_sheet_cnt) {
        char expect[64];
        for (i = 0; i < rd->tpl_sheet_cnt; i++) {
            snprintf(expect, sizeof(expect), "xl/worksheets/sheet%d.xml", i + 1);
            if (strcmp(rd->sheet_part[i], expect) != 0)
                return ARK_ERR_BAD_STATE;
        }
    }
    /* sst 新增但模板无 sst 部件 -> workbook rels 需重生成 */
    if (rd->sst_dirty && rd->i_sst == 0xFFFF && wb->sst_count)
        rd->wb_dirty = 1;

    ark_zip_begin(&zip, wb->write, wb->ctx);

    /* --- 决定重新生成的结构部件 --- */
    if (rd->wb_dirty) {
        regen_ct = regen_rels = regen_wb = regen_wbrels = 1;
        regen_core = regen_app = 1;
        if ((e = save_file(wb, &zip, "[Content_Types].xml", gen_content_types))
         || (e = save_file(wb, &zip, "_rels/.rels", gen_root_rels))
         || (e = save_file(wb, &zip, "docProps/core.xml", gen_core))
         || (e = save_file(wb, &zip, "docProps/app.xml", gen_app))
         || (e = save_file(wb, &zip, "xl/workbook.xml", gen_workbook))
         || (e = save_file(wb, &zip, "xl/_rels/workbook.xml.rels",
                           gen_workbook_rels)))
            goto fail;
    }

    /* --- styles: 增量拼接或原样复制 --- */
    if (rd->styles_dirty || wb->condfmt_count) {
        if (rd->i_styles == 0xFFFF) {
            /* 模板无 styles(罕见): 整体生成 */
            if ((e = save_file(wb, &zip, "xl/styles.xml", gen_styles)))
                goto fail;
        } else {
            if (ark_i_entry_extract(rd, rd->i_styles) != 0) { e = ARK_ERR_BAD_STATE; goto fail; }
            {   ark_xw_t x;
                ark_zip_measure_start(&zip);
                xw_init(&x, &zip, 0);
                gen_styles_open(wb, &x);
                xw_flush(&x);
                if (x.err) { e = x.err; goto fail; }
                if (ark_zip_open_stream(&zip, "xl/styles.xml") != 0)
                    { e = ARK_ERR_WRITE; goto fail; }
                xw_init(&x, &zip, 1);
                gen_styles_open(wb, &x);
                xw_flush(&x);
                if (x.err) { e = x.err; goto fail; }
                if (ark_zip_close_stream(&zip) != 0) { e = ARK_ERR_WRITE; goto fail; }
            }
        }
    } else if (rd->i_styles != 0xFFFF) {
        if ((e = copy_raw_entry(&zip, rd, &rd->entries[rd->i_styles])))
            goto fail;
    }

    /* --- SST: 重新生成或原样复制 --- */
    if (wb->sst_count && (rd->sst_dirty || rd->wb_dirty)) {
        if ((e = save_file(wb, &zip, "xl/sharedStrings.xml", gen_sst)))
            goto fail;
    } else if (rd->i_sst != 0xFFFF) {
        if ((e = copy_raw_entry(&zip, rd, &rd->entries[rd->i_sst])))
            goto fail;
    }

    /* --- 遍历模板条目 --- */
    for (i = 0; i < rd->entry_count; i++) {
        const char *n = rd->entries[i].name;
        int handled = 0;
        if (!strcmp(n, "[Content_Types].xml")) {
            if (!regen_ct) { e = copy_raw_entry(&zip, rd, &rd->entries[i]); handled = 1; }
            else handled = 1;
        }
        else if (!strcmp(n, "_rels/.rels")) {
            if (regen_rels) handled = 1;
        }
        else if (!strcmp(n, "docProps/core.xml") || !strcmp(n, "docProps/app.xml")) {
            if (regen_core) handled = 1;
        }
        else if (!strcmp(n, "xl/workbook.xml")) {
            if (regen_wb) handled = 1;
        }
        else if (!strcmp(n, "xl/_rels/workbook.xml.rels")) {
            if (regen_wbrels) handled = 1;
        }
        else if (!strcmp(n, "xl/styles.xml")) {
            handled = 1;    /* 上面已处理 */
        }
        else if (!strcmp(n, "xl/sharedStrings.xml")) {
            handled = 1;    /* 上面已处理 */
        }
        else {
            /* worksheet 部件? */
            uint8_t si;
            int is_sheet = 0;
            for (si = 0; si < rd->tpl_sheet_cnt; si++)
                if (strcmp(n, rd->sheet_part[si]) == 0) { is_sheet = 1; break; }
            if (is_sheet) {
                if (rd->sheet_dirty[si]) {
                    if ((e = save_sheet_file(wb, &zip, n, gen_sheet2, si)))
                        goto fail;
                } else {
                    if ((e = copy_raw_entry(&zip, rd, &rd->entries[i])))
                        goto fail;
                }
                handled = 1;
            } else {
                /* 该部件是否某 dirty sheet 的 rels */
                uint8_t k;
                for (k = 0; k < rd->tpl_sheet_cnt; k++) {
                    if (rd->sheet_dirty[k] && is_sheet_rels_of(n, rd->sheet_part[k])) {
                        /* 跳过; 若有外部超链接则重新生成 */
                        if (sheet_has_ext_links(wb, k)) {
                            if ((e = save_sheet_file(wb, &zip, n, gen_sheet_rels2, k)))
                                goto fail;
                        }
                        handled = 1;
                        break;
                    }
                }
            }
        }
        if (handled) continue;
        if ((e = copy_raw_entry(&zip, rd, &rd->entries[i]))) goto fail;
    }

    /* --- 用户新增 sheet 的部件 --- */
    for (i = rd->tpl_sheet_cnt; i < wb->sheet_count; i++) {
        snprintf(name, sizeof(name), "xl/worksheets/sheet%d.xml", i + 1);
        if (sheet_has_ext_links(wb, (uint8_t)i)) {
            char rname[96];
            sheet_rels_name(name, rname, sizeof(rname));
            if ((e = save_sheet_file(wb, &zip, rname, gen_sheet_rels2, (uint8_t)i)))
                goto fail;
        }
        if ((e = save_sheet_file(wb, &zip, name, gen_sheet2, (uint8_t)i)))
            goto fail;
    }

    if (ark_zip_end(&zip) != 0) { e = ARK_ERR_WRITE; goto fail; }
    wb->saved = 1;
    return ARK_OK;
fail:
    wb->last_err = e;
    return e;
}
#endif
