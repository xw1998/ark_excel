/*!
 * @file    test_features.c
 * @brief   ark_excel 全特性测试集: 按类别生成多个 xlsx,
 *          由 validate_features.py 逐一核对, 全部特性各成独立文件,
 *          任何一类损坏都能精确定位(特性二分思想, 也可单特性排查).
 *
 * 构建(PowerShell/cmd):  gcc -std=gnu99 -I.. -I../port test_features.c
 *                            ../src/ark_excel.c ../src/ark_zip.c -o test_features.exe
 * 运行: ./test_features.exe   (在 tests 目录下生成 out/*.xlsx)
 * 校验: python validate_features.py
 */
#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <direct.h>
#include <sys/stat.h>
#include "ark_excel.h"

static ark_wb_t wb;
static char out_dir[256] = "out";

static int w_stdio(const void *d, size_t n, void *ctx)
{
    return fwrite(d, 1, n, (FILE *)ctx) == n ? 0 : -1;
}

static int save_as(const char *fname)
{
    char path[300];
    FILE *f;
    snprintf(path, sizeof(path), "%s/%s", out_dir, fname);
    f = fopen(path, "wb");
    if (!f) { printf("open fail %s\n", path); return -1; }
    /* 把当前输出流切到新文件: 重绑回调上下文 */
    if (ark_wb_save(&wb) != ARK_OK) { fclose(f); return -1; }
    fclose(f);
    return 0;
}

/* 由于 ark_wb_save 在 init 时绑定回调, 这里用小包装: 每个用例独立 init */
static int run_case(void (*fill)(ark_wb_t *, int), const char *fname)
{
    char path[300];
    FILE *f;
    int sh;
    snprintf(path, sizeof(path), "%s/%s", out_dir, fname);
    f = fopen(path, "wb");
    if (!f) return -1;
    if (ark_wb_init(&wb, w_stdio, f) != ARK_OK) { fclose(f); return -1; }
    sh = ark_wb_add_sheet(&wb, "S1");
    if (sh < 0) { fclose(f); return -1; }
    fill(&wb, sh);
    if (ark_wb_save(&wb) != ARK_OK) { fclose(f); return -1; }
    fclose(f);
    return 0;
}

/* ---------------- 各特性填充函数 ---------------- */

static void f_basic(ark_wb_t *wb, int sh)
{
    ark_set_str(wb, sh, 1, 1, "文本", ARK_STYLE_DATA);
    ark_set_int(wb, sh, 2, 1, 42, ARK_STYLE_INT);
    ark_set_num(wb, sh, 3, 1, 3.14159, ARK_STYLE_NUM_2DEC);
    ark_set_num(wb, sh, 4, 1, -0.125, ARK_STYLE_PERCENT);
    ark_set_bool(wb, sh, 5, 1, 1, ARK_STYLE_DATA_CENTER);
    ark_set_error(wb, sh, 6, 1, "#DIV/0!", ARK_STYLE_DATA);
    ark_set_blank(wb, sh, 7, 1, ARK_STYLE_HEADER);
}
static void f_escape(ark_wb_t *wb, int sh)
{
    ark_set_str(wb, sh, 1, 1, "a<b>&c\"d'e", ARK_STYLE_DATA);
    ark_set_str(wb, sh, 2, 1, "中文内容 UTF-8", ARK_STYLE_DATA);
    ark_set_str(wb, sh, 3, 1, "tab\there", ARK_STYLE_DATA);
}
static void f_merge(ark_wb_t *wb, int sh)
{
    ark_set_str(wb, sh, 1, 1, "大标题", ARK_STYLE_TITLE);
    ark_merge(wb, sh, 1, 1, 1, 6);
    ark_set_str(wb, sh, 2, 1, "二级A", ARK_STYLE_HEADER);
    ark_merge(wb, sh, 2, 1, 2, 3);
    ark_set_str(wb, sh, 2, 4, "二级B", ARK_STYLE_HEADER);
    ark_merge(wb, sh, 2, 4, 2, 6);
    ark_set_str(wb, sh, 3, 1, "横向", ARK_STYLE_MERGE_CENTER);
    ark_merge(wb, sh, 3, 1, 3, 2);
    ark_set_str(wb, sh, 4, 3, "纵向", ARK_STYLE_MERGE_CENTER);
    ark_merge(wb, sh, 4, 3, 6, 3);
    ark_set_str(wb, sh, 3, 4, "九宫格", ARK_STYLE_MERGE_CENTER);
    ark_merge(wb, sh, 3, 4, 5, 6);
}
static void f_merge_stagger(ark_wb_t *wb, int sh)
{
    /* 双列合并 + 与前面列表错开(交错布局):
     * 行2: A:B 合并 |   C:D 合并      (同排两组双列)
     * 行3:    B:C 合并(整体右错开一列, 与上行交错)
     * 行4: A:B 合并 |   C:D 合中      (回归对齐) */
    ark_set_str(wb, sh, 1, 1, "序号", ARK_STYLE_HEADER);
    ark_set_str(wb, sh, 1, 2, "数值", ARK_STYLE_HEADER);
    ark_set_str(wb, sh, 1, 3, "状态", ARK_STYLE_HEADER);
    ark_set_str(wb, sh, 1, 4, "备注", ARK_STYLE_HEADER);
    ark_set_str(wb, sh, 2, 1, "左组", ARK_STYLE_MERGE_CENTER);
    ark_merge(wb, sh, 2, 1, 2, 2);
    ark_set_str(wb, sh, 2, 3, "右组", ARK_STYLE_MERGE_CENTER);
    ark_merge(wb, sh, 2, 3, 2, 4);
    ark_set_str(wb, sh, 3, 2, "中部错开", ARK_STYLE_MERGE_CENTER);
    ark_merge(wb, sh, 3, 2, 3, 3);
    ark_set_str(wb, sh, 4, 1, "回归A", ARK_STYLE_MERGE_CENTER);
    ark_merge(wb, sh, 4, 1, 4, 2);
    ark_set_str(wb, sh, 4, 3, "回归B", ARK_STYLE_MERGE_CENTER);
    ark_merge(wb, sh, 4, 3, 4, 4);
    /* 数据行紧随交错表头 */
    ark_set_int(wb, sh, 5, 1, 1, ARK_STYLE_DATA);
    ark_set_int(wb, sh, 5, 2, 100, ARK_STYLE_NUM_2DEC);
}
static void f_merge_rowspan(ark_wb_t *wb, int sh)
{
    /* 双行(跨两行)合并 + 与前面列表错开(交错布局):
     * 行1: 普通列表头(不合并)
     * 行2-3: A2:A3 纵向双行合并 | C2:D3 2x2 块合并(双行)
     * 行4-5: B4:B5 右错开一列的双行合并, 与上方列表交错
     * 行6: 数据行回归列表对齐 */
    ark_set_str(wb, sh, 1, 1, "项目", ARK_STYLE_HEADER);
    ark_set_str(wb, sh, 1, 2, "分组", ARK_STYLE_HEADER);
    ark_set_str(wb, sh, 1, 3, "单位", ARK_STYLE_HEADER);
    ark_set_str(wb, sh, 1, 4, "备注", ARK_STYLE_HEADER);
    /* 行2-3: 两组双行合并 */
    ark_set_str(wb, sh, 2, 1, "竖向双行", ARK_STYLE_MERGE_CENTER);
    ark_merge(wb, sh, 2, 1, 3, 1);          /* A2:A3 */
    ark_set_str(wb, sh, 2, 3, "2x2块", ARK_STYLE_MERGE_CENTER);
    ark_merge(wb, sh, 2, 3, 3, 4);          /* C2:D3 */
    /* 行4-5: 整体右错开一列的双行合并, 与上方组交错 */
    ark_set_str(wb, sh, 4, 2, "错开双行", ARK_STYLE_MERGE_CENTER);
    ark_merge(wb, sh, 4, 2, 5, 2);          /* B4:B5 */
    ark_set_str(wb, sh, 4, 3, "右错块", ARK_STYLE_MERGE_CENTER);
    ark_merge(wb, sh, 4, 3, 5, 4);          /* C4:D5 */
    /* 行6: 回归列表对齐的数据行 */
    ark_set_int(wb, sh, 6, 1, 42, ARK_STYLE_DATA);
}
static void f_altrow(ark_wb_t *wb, int sh)
{
    uint32_t r;
    ark_set_str(wb, sh, 1, 1, "交错报表", ARK_STYLE_TITLE);
    ark_merge(wb, sh, 1, 1, 1, 3);
    for (r = 2; r <= 11; r++) {
        int st = (r & 1) ? ARK_STYLE_ALT_ROW : ARK_STYLE_DATA;
        ark_set_int(wb, sh, r, 1, (int64_t)r, st);
        ark_set_num(wb, sh, r, 2, r * 1.5, ARK_STYLE_NUM_2DEC);
    }
}
static void f_date(ark_wb_t *wb, int sh)
{
    ark_set_date(wb, sh, 1, 1, 2024, 3, 15, ARK_STYLE_DATE);
    ark_set_date(wb, sh, 2, 1, 1900, 2, 28, ARK_STYLE_DATE);  /* 闰年bug边界 */
    ark_set_date(wb, sh, 3, 1, 1900, 3, 1, ARK_STYLE_DATE);
    ark_set_date(wb, sh, 4, 1, 1980, 1, 1, ARK_STYLE_DATE);
    ark_set_datetime(wb, sh, 5, 1, 2026, 8, 16, 14, 30, 5, ARK_STYLE_DATETIME);
}
static void f_formula(ark_wb_t *wb, int sh)
{
    ark_set_int(wb, sh, 1, 1, 10, ARK_STYLE_INT);
    ark_set_int(wb, sh, 2, 1, 20, ARK_STYLE_INT);
    ark_set_formula(wb, sh, 3, 1, "A1+A2", ARK_STYLE_INT);
    ark_set_formula(wb, sh, 4, 1, "SUM(A1:A2)*2", ARK_STYLE_INT);
    ark_set_formula(wb, sh, 5, 1, "IF(A1>5,\"big\",\"small\")", ARK_STYLE_DATA);
}
static void f_freeze(ark_wb_t *wb, int sh)
{
    uint32_t r;
    for (r = 1; r <= 20; r++) {
        ark_set_int(wb, sh, r, 1, (int64_t)r, ARK_STYLE_DATA);
        ark_set_int(wb, sh, r, 2, (int64_t)(r * 2), ARK_STYLE_DATA);
    }
    ark_freeze(wb, sh, 1, 1);   /* 同时冻结首行首列 */
    ark_autofilter(wb, sh, 1, 1, 20, 2);
}
static void f_colrow(ark_wb_t *wb, int sh)
{
    ark_set_col_width(wb, sh, 1, 1, 30.5);
    ark_set_col_width(wb, sh, 2, 5, 8.25);
    ark_set_row_height(wb, sh, 1, 40);
    ark_set_row_height(wb, sh, 2, 6.5);
    ark_set_str(wb, sh, 1, 1, "宽列", ARK_STYLE_DATA);
    ark_set_str(wb, sh, 2, 2, "常规", ARK_STYLE_DATA);
}
static void f_hidden(ark_wb_t *wb, int sh)
{
    uint32_t r;
    for (r = 1; r <= 10; r++)
        ark_set_int(wb, sh, r, 1, (int64_t)r, ARK_STYLE_DATA);
    for (r = 2; r <= 5; r++)
        ark_set_int(wb, sh, r, 2, (int64_t)(r * 10), ARK_STYLE_DATA);
    ark_set_row_hidden(wb, sh, 3, 1);
    ark_set_row_hidden(wb, sh, 4, 1);
    ark_set_col_hidden(wb, sh, 2, 2, 1);
}
static void f_styles(ark_wb_t *wb, int sh)
{
    ark_style_def_t sd = ARK_STYLE_DEF_INIT;
    int a, b, c;
    sd.bold = 1; sd.font_size = 1600; sd.font_color = ARK_COLOR_RED;
    sd.fill_color = ARK_COLOR_ALTBG; sd.border_all = ARK_BORDER_THICK;
    sd.halign = ARK_HALIGN_CENTER; sd.valign = ARK_VALIGN_CENTER;
    a = ark_wb_add_style(wb, &sd);                       /* 组合样式 */
    memset(&sd, 0, sizeof(sd));
    sd.font_name = "Consolas"; sd.numfmt_id = 9;          /* 内置千分位0% */
    b = ark_wb_add_style(wb, &sd);
    memset(&sd, 0, sizeof(sd));
    sd.custom_fmt = "0.000\"kg\""; sd.border_bottom = ARK_BORDER_DOUBLE;
    sd.halign = ARK_HALIGN_RIGHT; sd.wrap_text = 1; sd.indent = 2;
    c = ark_wb_add_style(wb, &sd);
    ark_set_str(wb, sh, 1, 1, "组合", (uint16_t)a);
    ark_set_num(wb, sh, 2, 1, 0.456, (uint16_t)b);
    ark_set_num(wb, sh, 3, 1, 12.34567, (uint16_t)c);
    /* 预定义样式遍历 */
    ark_set_str(wb, sh, 5, 1, "T", ARK_STYLE_TITLE);
    ark_set_str(wb, sh, 6, 1, "H", ARK_STYLE_HEADER);
    ark_set_num(wb, sh, 7, 1, 1234.5, ARK_STYLE_MONEY);
    ark_set_str(wb, sh, 8, 1, "W", ARK_STYLE_WARN);
    ark_set_str(wb, sh, 9, 1, "O", ARK_STYLE_OK);
}
static void f_condfmt(ark_wb_t *wb, int sh)
{
    uint32_t r;
    for (r = 1; r <= 10; r++)
        ark_set_int(wb, sh, r, 1, (int64_t)(r * 10), ARK_STYLE_DATA);
    ark_add_condfmt(wb, sh, "A1:A10", ARK_CF_CELLIS, ARK_OP_GREATER,
                    "50", NULL, ARK_COLOR_WARNBG, ARK_COLOR_WARNFG, 1);
    ark_add_condfmt(wb, sh, "A1:A10", ARK_CF_CELLIS, ARK_OP_LESS,
                    "20", NULL, ARK_COLOR_OKBG, ARK_COLOR_OKFG, 0);
}
static void f_valid(ark_wb_t *wb, int sh)
{
    ark_set_str(wb, sh, 1, 1, "正常", ARK_STYLE_DATA_CENTER);
    ark_add_validation(wb, sh, "A1:A5", ARK_VAL_LIST, 0,
                       "\"正常,告警,停机\"", NULL, 1,
                       "请选择状态", "必须从下拉列表选择");
    ark_set_int(wb, sh, 1, 2, 50, ARK_STYLE_INT);
    ark_add_validation(wb, sh, "B1:B5", ARK_VAL_WHOLE, ARK_OP_BETWEEN,
                       "0", "100", 0, "0-100", "超出范围");
}
static void f_hyperlink(ark_wb_t *wb, int sh)
{
    int s2;
    ark_set_str(wb, sh, 1, 1, "外部链接", ARK_STYLE_HYPERLINK);
    ark_hyperlink(wb, sh, "A1", "https://gitee.com/xw19981010/svcrtos_new.git",
                  0, "svcrtos 仓库");
    ark_set_str(wb, sh, 2, 1, "带特殊字符?a=1&b=2", ARK_STYLE_HYPERLINK);
    ark_hyperlink(wb, sh, "A2", "https://e.com/?a=1&b=2", 0, NULL);
    s2 = ark_wb_add_sheet(wb, "目标页");
    ark_set_str(wb, s2, 1, 1, "内部跳转目标", ARK_STYLE_DATA);
    ark_set_str(wb, sh, 3, 1, "跳到目标页", ARK_STYLE_HYPERLINK);
    ark_hyperlink(wb, sh, "A3", "#目标页!A1", 1, "内部跳转");
}
static void f_names(ark_wb_t *wb, int sh)
{
    ark_set_int(wb, sh, 1, 1, 1, ARK_STYLE_INT);
    ark_set_int(wb, sh, 2, 1, 2, ARK_STYLE_INT);
    ark_named_range(wb, "DataZone", "S1!$A$1:$A$2");
    ark_named_range(wb, "Answer", "S1!$A$1");
    ark_set_formula(wb, sh, 4, 1, "SUM(DataZone)", ARK_STYLE_INT);
}
static void f_multisheet(ark_wb_t *wb, int sh)
{
    int i;
    char name[16];
    ark_set_str(wb, sh, 1, 1, "第一个", ARK_STYLE_DATA);
    for (i = 1; i < 5; i++) {
        snprintf(name, sizeof(name), "表%d", i + 1);
        {
            int s = ark_wb_add_sheet(wb, name);
            if (s >= 0) {
                ark_set_str(wb, s, 1, 1, name, ARK_STYLE_HEADER);
                ark_set_num(wb, s, 2, 1, i * 1.25, ARK_STYLE_NUM_2DEC);
                ark_tab_color(wb, s, ARK_RGB(i * 40, 0x80, 0xC0));
            }
        }
    }
}
static void f_print(ark_wb_t *wb, int sh)
{
    ark_page_setup_t ps = ARK_PAGE_SETUP_DEFAULT;
    uint32_t r;
    for (r = 1; r <= 30; r++)
        ark_set_int(wb, sh, r, 1, (int64_t)r, ARK_STYLE_DATA);
    ps.orientation = ARK_LANDSCAPE;
    ps.paper = 8;              /* A3 */
    ps.fit_to_page = 1;
    ps.fit_width = 1;
    ps.fit_height = 0;
    ps.margin_left = 500;
    ark_page_setup(wb, sh, &ps);
    ark_print_area(wb, sh, "A1:A25");
    ark_header_footer(wb, sh, "&L左&C中&R&P/&N", "&C第&P页");
    ark_row_break(wb, sh, 15);
}
static void f_view(ark_wb_t *wb, int sh)
{
    uint32_t r;
    for (r = 1; r <= 10; r++)
        ark_set_int(wb, sh, r, 1, (int64_t)r, ARK_STYLE_DATA);
    ark_sheet_rtl(wb, sh, 1);
    ark_sheet_gridlines(wb, sh, 0);
    ark_tab_color(wb, sh, ARK_COLOR_ORANGE);
}
static void f_sst_flood(ark_wb_t *wb, int sh)
{
    /* 大量不同字符串: 考验 SST 去重/上限降级为内联字符串 */
    uint32_t r;
    char buf[40];
    for (r = 1; r <= 60; r++) {
        snprintf(buf, sizeof(buf), "字符串条目编号%03u号", r);
        ark_set_str(wb, sh, r, 1, buf, ARK_STYLE_DATA);
    }
    for (r = 1; r <= 20; r++)                 /* 重复串应被去重 */
        ark_set_str(wb, sh, r, 2, "字符串条目编号001号", ARK_STYLE_DATA);
}
static void f_overwrite(ark_wb_t *wb, int sh)
{
    /* 乱序 + 覆盖写: 生成器必须按行列排序且后写覆盖先写 */
    uint32_t r;
    for (r = 20; r >= 1; r--)
        ark_set_int(wb, sh, r, 3, (int64_t)r, ARK_STYLE_DATA);
    for (r = 1; r <= 20; r += 2)
        ark_set_int(wb, sh, r, 1, (int64_t)(r + 100), ARK_STYLE_DATA);
    ark_set_str(wb, sh, 10, 3, "覆盖", ARK_STYLE_WARN);  /* 覆盖数值 */
}

static const struct { const char *name; void (*fn)(ark_wb_t *, int); } CASES[] = {
    { "basic.xlsx",     f_basic },
    { "escape.xlsx",    f_escape },
    { "merge.xlsx",     f_merge },
    { "merge_stagger.xlsx", f_merge_stagger },
    { "merge_rowspan.xlsx",  f_merge_rowspan },
    { "altrow.xlsx",    f_altrow },
    { "date.xlsx",      f_date },
    { "formula.xlsx",   f_formula },
    { "freeze.xlsx",    f_freeze },
    { "colrow.xlsx",    f_colrow },
    { "hidden.xlsx",    f_hidden },
    { "styles.xlsx",    f_styles },
    { "condfmt.xlsx",   f_condfmt },
    { "valid.xlsx",     f_valid },
    { "hyperlink.xlsx", f_hyperlink },
    { "names.xlsx",     f_names },
    { "multisheet.xlsx",f_multisheet },
    { "print.xlsx",     f_print },
    { "view.xlsx",      f_view },
    { "sst.xlsx",       f_sst_flood },
    { "overwrite.xlsx", f_overwrite },
};
#define N_CASES (int)(sizeof(CASES) / sizeof(CASES[0]))

int main(int argc, char **argv)
{
    int i, bad = 0;
    if (argc > 1) snprintf(out_dir, sizeof(out_dir), "%s", argv[1]);
#ifdef _WIN32
    _mkdir(out_dir);
#else
    mkdir(out_dir, 0755);
#endif
    for (i = 0; i < N_CASES; i++) {
        if (run_case(CASES[i].fn, CASES[i].name) != 0) {
            printf("GEN FAIL %s\n", CASES[i].name);
            bad++;
        }
    }
    printf("generated %d cases, %d failed\n", N_CASES, bad);
    return bad;
}
