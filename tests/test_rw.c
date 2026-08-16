/*!
 * @file    test_rw.c
 * @brief   ark_excel 读改写(打开/修改/保存)冒烟测试
 *
 * 用例:
 *   1. rw_basic        : 单表读改写 —— 载入校验 + 覆盖/追加单元格 + 样式重映射
 *   2. rw_zero_loss    : 未修改 sheet 原样字节复制(零损失)
 *   3. rw_multisheet   : 多 sheet, 仅改 sheet1, sheet2 不动
 *   4. rw_deflate_tpl  : Deflate 压缩模板(第三方生成)的读改写
 *   5. rw_add_sheet    : 在已有工作簿上新增 sheet
 *   6. rw_custom_style : 自定义样式追加进模板样式表
 *   7. rw_condfmt      : 载入后新增条件格式(dxf 追加)
 *   8. rw_validation   : 载入后新增数据验证
 *
 * 运行: ./tests/test_rw.exe  (输出到 tests/out/rw_*.xlsx)
 */
#include <stdio.h>
#include <string.h>
#include "ark_excel.h"
#include "ark_reader.h"

static int g_fail = 0, g_total = 0;
#define CHECK(cond, msg) do { \
    g_total++; \
    if (cond) { printf("  ok   %s\n", msg); } \
    else      { printf("  FAIL %s\n", msg); g_fail++; } \
} while (0)

/* ---------------- 文件读写回调 ---------------- */
static int wr(const void *d, size_t n, void *ctx)
{ return fwrite(d, 1, n, (FILE *)ctx) == n ? 0 : 1; }
static int wr_nop(const void *d, size_t n, void *ctx)
{ (void)d; (void)n; (void)ctx; return 0; }
static uint32_t fread_cb(void *ctx, uint8_t *buf, uint32_t len)
{ return (uint32_t)fread(buf, 1, len, (FILE *)ctx); }
static int fseek_cb(void *ctx, uint32_t pos)
{ return fseek((FILE *)ctx, (long)pos, SEEK_SET); }

static ark_wb_t  g_wb;
static ark_reader_t g_rd;

/* ---------------- 模板生成 ---------------- */

/*! 生成单表模板(rw_basic/rw_zero_loss/rw_custom_style 用) */
static void make_tpl_basic(const char *path)
{
    FILE *f = fopen(path, "wb");
    int s;
    ark_wb_init(&g_wb, wr, f);
    s = ark_wb_add_sheet(&g_wb, "Data");
    ark_set_str(&g_wb, s, 1, 1, "设备", ARK_STYLE_HEADER);
    ark_set_str(&g_wb, s, 1, 2, "数值", ARK_STYLE_HEADER);
    ark_set_str(&g_wb, s, 2, 1, "温度传感器A", ARK_STYLE_DATA);
    ark_set_num(&g_wb, s, 2, 2, 25.5, ARK_STYLE_NUM_2DEC);
    ark_set_int(&g_wb, s, 3, 1, 100, ARK_STYLE_INT);
    ark_set_formula(&g_wb, s, 3, 2, "B2*2", ARK_STYLE_NUM_2DEC);
    ark_merge(&g_wb, s, 5, 1, 5, 2);
    ark_set_str(&g_wb, s, 5, 1, "合并区", ARK_STYLE_MERGE_CENTER);
    ark_freeze(&g_wb, s, 1, 0);
    ark_set_col_width(&g_wb, s, 1, 1, 18);
    ark_wb_save(&g_wb);
    fclose(f);
}

/*! 生成双表模板(rw_multisheet/rw_add_sheet 用) */
static void make_tpl_multi(const char *path)
{
    FILE *f = fopen(path, "wb");
    int s1, s2;
    ark_wb_init(&g_wb, wr, f);
    s1 = ark_wb_add_sheet(&g_wb, "Sheet1");
    s2 = ark_wb_add_sheet(&g_wb, "Sheet2");
    ark_set_str(&g_wb, s1, 1, 1, "表一A1", ARK_STYLE_HEADER);
    ark_set_num(&g_wb, s1, 2, 1, 11, ARK_STYLE_INT);
    ark_set_str(&g_wb, s2, 1, 1, "表二A1", ARK_STYLE_TITLE);
    ark_set_num(&g_wb, s2, 2, 1, 22, ARK_STYLE_MONEY);
    ark_freeze(&g_wb, s2, 1, 1);
    ark_wb_save(&g_wb);
    fclose(f);
}

/*! 通用 open 流程: 成功返回 0 */
static int open_tpl(const char *path)
{
    FILE *src = fopen(path, "rb");
    ark_reader_io_t io = { fread_cb, fseek_cb, NULL };
    ark_err_t e;
    if (!src) return -1;
    io.ctx = src;
    e = ark_wb_init(&g_wb, wr_nop, NULL);
    if (e == ARK_OK) e = ark_wb_open(&g_wb, &g_rd, &io);
    if (e != ARK_OK) { fclose(src); return (int)e; }
    /* 注意: 读取源需保持打开到 save 完成(增量保存仍需读模板字节);
     * 本测试用全局 FILE* 保存 */
    {
        static FILE *keep;
        keep = src;   /* 防优化; 实际生命周期由调用方控制 */
    }
    return 0;
}

/* 为简化生命周期管理: 每个用例在 save 前不关闭源 */
static FILE *g_src = NULL;

static int open_tpl2(const char *path)
{
    ark_reader_io_t io = { fread_cb, fseek_cb, NULL };
    ark_err_t e;
    g_src = fopen(path, "rb");
    if (!g_src) return -1;
    io.ctx = g_src;
    e = ark_wb_init(&g_wb, wr_nop, NULL);
    if (e == ARK_OK) e = ark_wb_open(&g_wb, &g_rd, &io);
    return e == ARK_OK ? 0 : (int)e;
}

static int save_out(const char *path)
{
    FILE *f = fopen(path, "wb");
    ark_err_t e;
    if (!f) return -1;
    g_wb.write = wr; g_wb.ctx = f;
    e = ark_wb_save(&g_wb);
    fclose(f);
    if (g_src) { fclose(g_src); g_src = NULL; }
    return e == ARK_OK ? 0 : (int)e;
}

/* ---------------- 用例 ---------------- */

static void case_basic(void)
{
    uint8_t type; double num; char text[96]; uint16_t style;
    printf("case 1: rw_basic\n");
    make_tpl_basic("tests/out/rw_tpl_basic.xlsx");
    CHECK(open_tpl2("tests/out/rw_tpl_basic.xlsx") == 0, "open template");

    CHECK(g_wb.sheet_count == 1, "sheet count = 1");
    CHECK(g_wb.cell_count == 7, "cell count = 7");
    CHECK(g_wb.sst_count == 4, "sst count = 4");

    ark_get_cell(&g_wb, 0, 2, 1, &type, &num, text, sizeof(text), &style);
    CHECK(type == ARK_CELL_SSTR && strcmp(text, "温度传感器A") == 0,
          "load B2 string");
    CHECK(style == ARK_STYLE_DATA, "load B2 style");

    ark_get_cell(&g_wb, 0, 2, 2, &type, &num, NULL, 0, NULL);
    CHECK(type == ARK_CELL_NUM && num == 25.5, "load B2 number");

    ark_get_cell(&g_wb, 0, 3, 2, &type, &num, text, sizeof(text), NULL);
    CHECK(type == ARK_CELL_FORMULA && strcmp(text, "B2*2") == 0,
          "load B3 formula");

    CHECK(g_wb.merge_count == 1, "load merge");
    CHECK(g_wb.sheets[0].freeze_rows == 1, "load freeze");
    CHECK(g_wb.coldef_count == 1 && g_wb.coldefs[0].width_x100 == 1800,
          "load col width");

    /* 修改: 覆盖 + 追加(带重映射样式) */
    ark_set_num(&g_wb, 0, 2, 2, 99.9, ark_style_map(ARK_STYLE_NUM_2DEC));
    ark_set_str(&g_wb, 0, 4, 1, "新插入行", ark_style_map(ARK_STYLE_WARN));
    ark_set_str(&g_wb, 0, 4, 2, "附加数据", ark_style_map(ARK_STYLE_OK));
    CHECK(save_out("tests/out/rw_basic.xlsx") == 0, "save modified");
}

static void case_zero_loss(void)
{
    printf("case 2: rw_zero_loss (不改内容直接另存, 部件应原样复制)\n");
    make_tpl_basic("tests/out/rw_tpl_zl.xlsx");
    CHECK(open_tpl2("tests/out/rw_tpl_zl.xlsx") == 0, "open");
    ark_set_num(&g_wb, 0, 2, 2, 1.5, 0);   /* 仅一个修改点 */
    CHECK(save_out("tests/out/rw_zl.xlsx") == 0, "save");
}

static void case_multisheet(void)
{
    printf("case 3: rw_multisheet (改 sheet1 不动 sheet2)\n");
    make_tpl_multi("tests/out/rw_tpl_multi.xlsx");
    CHECK(open_tpl2("tests/out/rw_tpl_multi.xlsx") == 0, "open");
    CHECK(g_wb.sheet_count == 2, "two sheets loaded");
    ark_set_str(&g_wb, 0, 3, 1, "sheet1新增", ark_style_map(ARK_STYLE_DATA));
    CHECK(save_out("tests/out/rw_multi.xlsx") == 0, "save");
}

static void case_add_sheet(void)
{
    printf("case 4: rw_add_sheet\n");
    make_tpl_multi("tests/out/rw_tpl_add.xlsx");
    CHECK(open_tpl2("tests/out/rw_tpl_add.xlsx") == 0, "open");
    {
        int s = ark_wb_add_sheet(&g_wb, "NewSheet");
        CHECK(s == 2, "add sheet index=2");
        ark_set_str(&g_wb, s, 1, 1, "新表", ark_style_map(ARK_STYLE_TITLE));
        ark_set_num(&g_wb, s, 2, 1, 3.14, ark_style_map(ARK_STYLE_NUM_2DEC));
    }
    CHECK(save_out("tests/out/rw_addsheet.xlsx") == 0, "save with new sheet");
}

static void case_custom_style(void)
{
    printf("case 5: rw_custom_style\n");
    make_tpl_basic("tests/out/rw_tpl_cs.xlsx");
    CHECK(open_tpl2("tests/out/rw_tpl_cs.xlsx") == 0, "open");
    {
        ark_style_def_t d = ARK_STYLE_DEF_INIT;
        int id;
        d.font_size = 1200; d.bold = 1;
        d.fill_color = ARK_RGB(0xFF, 0xD9, 0x66);
        d.border_all = ARK_BORDER_THIN;
        d.halign = ARK_HALIGN_CENTER;
        id = ark_wb_add_style(&g_wb, &d);
        CHECK(id > 0, "custom style registered");
        ark_set_str(&g_wb, 0, 7, 1, "自定义样式行", (uint16_t)id);
        ark_set_str(&g_wb, 0, 8, 1, "内置样式行", ark_style_map(ARK_STYLE_HEADER));
    }
    CHECK(save_out("tests/out/rw_custstyle.xlsx") == 0, "save");
}

static void case_condfmt(void)
{
    printf("case 6: rw_condfmt (模板无 dxfs, 追加条件格式)\n");
    make_tpl_basic("tests/out/rw_tpl_cf.xlsx");
    CHECK(open_tpl2("tests/out/rw_tpl_cf.xlsx") == 0, "open");
    ark_add_condfmt(&g_wb, 0, "B2:B10", ARK_CF_CELLIS, ARK_OP_GREATER,
                    "50", NULL, ARK_COLOR_WARNBG, ARK_COLOR_WARNFG, 1);
    ark_set_num(&g_wb, 0, 2, 2, 88.0, ark_style_map(ARK_STYLE_NUM_2DEC));
    CHECK(save_out("tests/out/rw_condfmt.xlsx") == 0, "save");
}

static void case_validation(void)
{
    printf("case 7: rw_validation\n");
    make_tpl_basic("tests/out/rw_tpl_v.xlsx");
    CHECK(open_tpl2("tests/out/rw_tpl_v.xlsx") == 0, "open");
    ark_add_validation(&g_wb, 0, "A10:A20", ARK_VAL_LIST, 0,
                       "\"合格,不合格,待检\"", NULL, 1, NULL, NULL);
    CHECK(save_out("tests/out/rw_valid.xlsx") == 0, "save");
}

static void case_deflate_tpl(void)
{
    printf("case 8: rw_deflate_tpl (Deflate 压缩模板, 见 run 前 python 生成)\n");
    /* 模板由测试驱动脚本用 python(openpyxl) 生成: tests/out/rw_tpl_openpyxl.xlsx */
    CHECK(open_tpl2("tests/out/rw_tpl_openpyxl.xlsx") == 0, "open openpyxl file");
    CHECK(g_wb.sheet_count >= 1, "sheet loaded from 3rd-party file");
    {
        uint8_t type; double num; char text[96];
        int r = ark_get_cell(&g_wb, 0, 1, 1, &type, &num, text, sizeof(text), NULL);
        CHECK(r == 0, "cell A1 exists");
        printf("       (A1 = %s)\n", text);
    }
    ark_set_str(&g_wb, 0, 5, 1, "写入第三方模板", ark_style_map(ARK_STYLE_HEADER));
    { int r = save_out("tests/out/rw_openpyxl.xlsx");
      CHECK(r == 0, "save modified 3rd-party"); }
}

int main(void)
{
    setvbuf(stdout, NULL, _IONBF, 0);
    printf("==== ark_excel read-modify-write smoke ====\n");
    case_basic();
    case_zero_loss();
    case_multisheet();
    case_add_sheet();
    case_custom_style();
    case_condfmt();
    case_validation();
    case_deflate_tpl();
    printf("==== %d checks, %d failed ====\n", g_total, g_fail);
    return g_fail ? 1 : 0;
}
