/* 乱序写入 + 覆盖写 + 大批量 回归测试 */
#include <stdio.h>
#include <string.h>
#include "ark_excel.h"
#include "port_examples.c"

static ark_wb_t wb;
static int fail = 0;
#define CK(e) do { ark_err_t _e=(e); if(_e!=ARK_OK){ \
    printf("FAIL %s -> %s\n", #e, ark_errstr(_e)); fail=1; } } while(0)

int main(void)
{
    FILE *f = fopen("stress.xlsx", "wb");
    uint32_t i;
    int sh;
    char buf[32];
    if (!f) return 1;
    CK(ark_wb_init(&wb, ark_write_stdio, f));
    sh = ark_wb_add_sheet(&wb, "S");
    /* 乱序写入: 倒序 + 跳行 */
    for (i = 500; i >= 1; i--) {
        sprintf(buf, "row%u", i);
        CK(ark_set_int(&wb, sh, i, 3, (int64_t)i, ARK_STYLE_INT));
        CK(ark_set_str(&wb, sh, i, 1, buf, i & 1 ? ARK_STYLE_ALT_ROW : ARK_STYLE_DATA));
        CK(ark_set_num(&wb, sh, i, 2, (double)i / 7.0, ARK_STYLE_NUM_2DEC));
    }
    /* 覆盖写 */
    CK(ark_set_str(&wb, sh, 250, 1, "OVERWRITE", ARK_STYLE_HEADER));
    CK(ark_merge(&wb, sh, 1, 1, 1, 3));
    CK(ark_freeze(&wb, sh, 1, 0));
    CK(ark_wb_save(&wb));
    fclose(f);
    printf(fail ? "STRESS FAIL\n" : "STRESS OK cells=%u\n", wb.cell_count);
    return fail;
}
