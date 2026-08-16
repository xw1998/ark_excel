/*!
 * @file    demo_baremetal.c
 * @brief   裸机/无文件系统 示例: 生成到内存缓冲, 再经 UART 发送
 *          (伪代码风格, uart_send_blocking 换成你的 BSP 函数)
 */
#include <string.h>
#include <stdio.h>
#include "ark_excel.h"

/* ---- 内存缓冲输出(缓冲大小需 >= 生成文件大小, 或改用 UART 直发) ---- */
static uint8_t  s_outbuf[32 * 1024];
static size_t   s_outlen;

static int write_membuf(const void *data, size_t len, void *ctx)
{
    (void)ctx;
    if (s_outlen + len > sizeof(s_outbuf)) return -1;
    memcpy(s_outbuf + s_outlen, data, len);
    s_outlen += len;
    return 0;
}

/* 若无文件系统可直接流到 UART, 零 RAM 缓冲:
 * static int write_uart(const void *d, size_t n, void *ctx)
 * { (void)ctx; return uart_send_blocking(d, (unsigned)n); } */

static ark_wb_t s_wb;   /* 放指定 RAM 段: __attribute__((section(".sdram"))) */

void make_report(void)
{
    ark_wb_t *wb = &s_wb;
    int sh = ark_wb_add_sheet(wb, "Data");
    if (sh < 0) return;

    ark_set_str(wb, sh, 1, 1, "Sensor Log", ARK_STYLE_TITLE);
    ark_merge(wb, sh, 1, 1, 1, 4);
    ark_set_str(wb, sh, 2, 1, "Time",   ARK_STYLE_HEADER);
    ark_set_str(wb, sh, 2, 2, "Ch0",    ARK_STYLE_HEADER);
    ark_set_str(wb, sh, 2, 3, "Ch1",    ARK_STYLE_HEADER);
    ark_set_str(wb, sh, 2, 4, "OK?",    ARK_STYLE_HEADER);
    ark_freeze(wb, sh, 2, 0);

    {
        uint32_t r;
        for (r = 3; r <= 20; r++) {
            int st = (r & 1) ? ARK_STYLE_ALT_ROW : ARK_STYLE_DATA;
            ark_set_int(wb, sh, r, 1, (int64_t)(r - 2) * 100, ARK_STYLE_INT);
            ark_set_num(wb, sh, r, 2, 2.5 * (r - 2), ARK_STYLE_NUM_2DEC);
            ark_set_num(wb, sh, r, 3, 1.25 * (r - 2), ARK_STYLE_NUM_2DEC);
            ark_set_bool(wb, sh, r, 4, (r % 3) != 0, st);
        }
    }

    ark_err_t sv = ark_wb_save(wb);
    printf("save=%s outlen=%lu\n", ark_errstr(sv), (unsigned long)s_outlen);
    if (sv == ARK_OK) {
        /* 此时 s_outlen 字节的合法 xlsx 位于 s_outbuf:
         * - 写入外部 SPI Flash / SD
         * - 经 UART 'wb->write=write_uart' 版本则早已流式发出 */
    }
}

int main(void)
{
    ark_wb_init(&s_wb, write_membuf, NULL);
    make_report();
    return 0;   /* 文件已在 s_outbuf, 共 s_outlen 字节 */
}
