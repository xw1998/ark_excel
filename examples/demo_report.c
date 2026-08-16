/*!
 * @file    demo_report.c
 * @brief   ark_excel 完整功能演示: 生成交错行 + 多级表头合并 + 冻结 +
 *          筛选 + 条件格式 + 数据验证 + 超链接 + 公式 + 打印设置 的报表
 *
 * 编译(与 src/ark_excel.c src/ark_zip.c 一起):
 *   cc -I. demo_report.c port/port_examples.c src/ark_excel.c src/ark_zip.c -o demo
 */
#include <stdio.h>
#include "ark_excel.h"
#include "../port/port_examples.c"

static ark_wb_t wb;   /* 静态分配, 也可放堆/特定 RAM 段 */

int main(void)
{
    FILE *f;
    int sh, i;
    ark_style_def_t sd = ARK_STYLE_DEF_INIT;   /* 必须清零初始化! */
    int style_warn_bold;
    static const char *names[5] = { "温度传感器A", "压力变送器B", "流量计C",
                                    "液位开关D", "振动探头E" };
    static const double temps[5] = { 25.5, 88.7, 42.1, 15.0, 60.3 };

    f = fopen("demo_report.xlsx", "wb");
    if (!f) { perror("fopen"); return 1; }

    if (ark_wb_init(&wb, ark_write_stdio, f) != ARK_OK) return 2;
    sh = ark_wb_add_sheet(&wb, "设备巡检");
    if (sh < 0) return 3;

    /* --- 自定义样式: 橙色粗体斜体 + 黄色填充 + 中粗边框 --- */
    sd.font_size = 1200; sd.bold = 1; sd.italic = 1; sd.font_color = ARK_COLOR_ORANGE;
    sd.fill_color = ARK_RGB(0xFF,0xF2,0xCC);
    sd.border_all = ARK_BORDER_MEDIUM;
    sd.halign = ARK_HALIGN_CENTER; sd.valign = ARK_VALIGN_CENTER;
    sd.wrap_text = 1; sd.numfmt_id = 0; sd.custom_fmt = NULL;
    style_warn_bold = ark_wb_add_style(&wb, &sd);

    /* --- 三级表头(交错合并) --- */
    ark_set_str(&wb, sh, 1, 1, "车间设备巡检日报表", ARK_STYLE_TITLE);
    ark_merge(&wb, sh, 1, 1, 1, 6);
    ark_set_str(&wb, sh, 2, 1, "采集时间:", ARK_STYLE_DATA_RIGHT);
    ark_set_date(&wb, sh, 2, 2, 2026, 8, 16, ARK_STYLE_DATE);
    ark_set_str(&wb, sh, 2, 4, "值班员:", ARK_STYLE_DATA_RIGHT);
    ark_set_str(&wb, sh, 2, 5, "张三", ARK_STYLE_DATA_CENTER);

    ark_set_str(&wb, sh, 4, 1, "基本信息", ARK_STYLE_HEADER);
    ark_merge(&wb, sh, 4, 1, 4, 2);
    ark_set_str(&wb, sh, 4, 3, "测量数据", ARK_STYLE_HEADER);
    ark_merge(&wb, sh, 4, 3, 4, 6);
    ark_set_str(&wb, sh, 5, 1, "设备名称", ARK_STYLE_HEADER);
    ark_set_str(&wb, sh, 5, 2, "安装位置", ARK_STYLE_HEADER);
    ark_set_str(&wb, sh, 5, 3, "温度(℃)", ARK_STYLE_HEADER);
    ark_set_str(&wb, sh, 5, 4, "压力(kPa)", ARK_STYLE_HEADER);
    ark_set_str(&wb, sh, 5, 5, "状态", ARK_STYLE_HEADER);
    ark_set_str(&wb, sh, 5, 6, "备注", ARK_STYLE_HEADER);

    /* --- 数据行(交错样式) --- */
    for (i = 0; i < 5; i++) {
        int r = 6 + i;
        int st = (i & 1) ? ARK_STYLE_ALT_ROW : ARK_STYLE_DATA;
        ark_set_str(&wb, sh, r, 1, names[i], st);
        ark_set_str(&wb, sh, r, 2, i == 0 ? "1#线" : "2#线", st);
        ark_set_num(&wb, sh, r, 3, temps[i], ARK_STYLE_NUM_2DEC);
        ark_set_int(&wb, sh, r, 4, 100 + i * 12, ARK_STYLE_INT);
        if (temps[i] > 80.0) {
            ark_set_str(&wb, sh, r, 5, "超温告警", style_warn_bold);
            ark_set_error(&wb, sh, r, 6, "#N/A", st);
        } else {
            ark_set_str(&wb, sh, r, 5, "正常", st);
        }
    }
    /* 合计行: 公式 */
    ark_set_str(&wb, sh, 11, 1, "平均温度", ARK_STYLE_TOTAL);
    ark_set_formula(&wb, sh, 11, 3, "AVERAGE(C6:C10)", ARK_STYLE_TOTAL);
    ark_merge(&wb, sh, 11, 1, 11, 2);

    /* --- 特殊格式 --- */
    ark_freeze(&wb, sh, 5, 0);                       /* 冻结表头 */
    ark_autofilter(&wb, sh, 5, 1, 10, 6);            /* 自动筛选 */
    ark_set_col_width(&wb, sh, 1, 2, 16);
    ark_set_col_width(&wb, sh, 3, 6, 12);
    ark_set_row_height(&wb, sh, 1, 24);

    /* 条件格式: 压力 > 130 标红 */
    ark_add_condfmt(&wb, sh, "D6:D10", ARK_CF_CELLIS, ARK_OP_GREATER,
                    "130", NULL, ARK_COLOR_WARNBG, ARK_COLOR_WARNFG, 1);
    /* 数据验证: 状态列下拉框 */
    ark_add_validation(&wb, sh, "E6:E20", ARK_VAL_LIST, 0,
                       "\"正常,超温告警,维修中,停机\"", NULL, 1,
                       "请选择设备状态", "取值必须来自下拉列表");
    /* 超链接: 内部跳转 + 外部 URL */
    ark_set_str(&wb, sh, 12, 1, "查看历史趋势", ARK_STYLE_HYPERLINK);
    ark_hyperlink(&wb, sh, "A12", "#Sheet2!A1", 1, "跳转");
    ark_set_str(&wb, sh, 12, 3, "规格书下载", ARK_STYLE_HYPERLINK);
    ark_hyperlink(&wb, sh, "C12", "https://example.com/spec.pdf", 0, NULL);

    /* 打印: A4 横向, 页宽自适应, 页眉页脚 */
    {
        ark_page_setup_t ps = ARK_PAGE_SETUP_DEFAULT;
        ps.orientation = ARK_LANDSCAPE;
        ps.fit_to_page = 1;
        ark_page_setup(&wb, sh, &ps);
    }
    ark_print_area(&wb, sh, "A1:F12");
    ark_header_footer(&wb, sh, "&L车间巡检&C&R第 &P 页",
                      "&C生成于嵌入式设备");

    /* --- 第二个 Sheet: 历史趋势(命名区域 + Tab 颜色) --- */
    {
        int s2 = ark_wb_add_sheet(&wb, "Sheet2");
        ark_set_str(&wb, s2, 1, 1, "历史趋势数据(示意)", ARK_STYLE_TITLE);
        for (i = 0; i < 5; i++)
            ark_set_num(&wb, s2, 2 + i, 1, 20.0 + i, ARK_STYLE_NUM_2DEC);
        ark_named_range(&wb, "TrendData", "Sheet2!$A$2:$A$6");
        ark_tab_color(&wb, s2, ARK_COLOR_GREEN);
    }

    /* --- 保存 --- */
    if (ark_wb_save(&wb) != ARK_OK) {
        printf("save failed: %s\n", ark_errstr(wb.last_err ?
                                               (ark_err_t)wb.last_err : ARK_ERR_BAD_STATE));
        fclose(f);
        return 4;
    }
    fclose(f);
    printf("demo_report.xlsx generated, %d cells\n", (int)wb.cell_count);
    return 0;
}
