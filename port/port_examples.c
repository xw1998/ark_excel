/*!
 * @file    port_examples.c
 * @brief   ark_excel 输出回调移植示例: stdio / 内存缓冲 / UART / FatFS / svcrtos
 *
 * 本文件仅是参考实现, 按需拷贝到你的工程. 库本身不依赖本文件.
 */
#include <stdio.h>
#include <string.h>
#include "ark_excel.h"

/* ------------------------------------------------------------
 * 1. 标准 C 文件(POSIX / Windows / 嵌入式 Linux)
 * ------------------------------------------------------------ */
int ark_write_stdio(const void *data, size_t len, void *ctx)
{
    FILE *f = (FILE *)ctx;
    return (fwrite(data, 1, len, f) == len) ? 0 : -1;
}
/* 用法: FILE *f = fopen("report.xlsx", "wb");
 *       ark_wb_init(&wb, ark_write_stdio, f); ... fclose(f); */

/* ------------------------------------------------------------
 * 2. 内存缓冲(便于先压缩/加密或经网络发送)
 * ------------------------------------------------------------ */
typedef struct {
    uint8_t *buf;
    size_t   len, cap;
} ark_membuf_t;

int ark_write_membuf(const void *data, size_t len, void *ctx)
{
    ark_membuf_t *m = (ark_membuf_t *)ctx;
    if (m->len + len > m->cap) return -1;
    memcpy(m->buf + m->len, data, len);
    m->len += len;
    return 0;
}

/* ------------------------------------------------------------
 * 3. 裸机 UART / DMA 流式发送
 * ------------------------------------------------------------ */
#if 0   /* 按你的芯片修改 */
extern int uart_send_blocking(const void *d, unsigned n);  /* BSP 提供 */

int ark_write_uart(const void *data, size_t len, void *ctx)
{
    (void)ctx;
    return uart_send_blocking(data, (unsigned)len);
}
/* 用法: ark_wb_init(&wb, ark_write_uart, NULL); */
#endif

/* ------------------------------------------------------------
 * 4. FatFS (SD 卡 / SPI Flash + elm-chan FatFS)
 * ------------------------------------------------------------ */
#if 0
#include "ff.h"

int ark_write_fatfs(const void *data, size_t len, void *ctx)
{
    UINT bw;
    FRESULT r = f_write((FIL *)ctx, data, (UINT)len, &bw);
    return (r == FR_OK && bw == len) ? 0 : -1;
}
/* 用法: FIL fil; f_open(&fil, "0:/report.xlsx", FA_WRITE|FA_CREATE_ALWAYS);
 *       ark_wb_init(&wb, ark_write_fatfs, &fil); ... f_close(&fil); */
#endif

/* ------------------------------------------------------------
 * 5. svcrtos (自研 RTOS, https://gitee.com/xw19981010/svcrtos_new)
 *    思路: 生成任务持有队列, 回调把数据块投递给写任务/存储服务,
 *    生成逻辑永不阻塞存储慢速设备, 也天然线程安全(单消费者).
 * ------------------------------------------------------------ */
#if 0
#include "svc_os.h"    /* 按实际头文件名调整 */

#define WRQ_MSG_BYTES 128   /* 每条消息负载, 建议与 ARK_IO_CHUNK 匹配 */

typedef struct {
    uint32_t len;
    uint8_t  data[WRQ_MSG_BYTES];
} wrq_msg_t;

static svc_queue_t s_wrq;
static wrq_msg_t   s_pool[16];
static svc_mempool_t s_pool_hdl;

int ark_write_svcrtos(const void *data, size_t len, void *ctx)
{
    (void)ctx;
    while (len) {
        wrq_msg_t *m = (wrq_msg_t *)svc_mempool_alloc(&s_pool_hdl);
        size_t n = len > WRQ_MSG_BYTES ? WRQ_MSG_BYTES : len;
        if (!m) return -1;                       /* 池空, 视为背压 */
        memcpy(m->data, data, n);
        m->len = (uint32_t)n;
        svc_queue_send(&s_wrq, m, SVC_WAIT_FOREVER);
        data = (const char *)data + n;
        len -= n;
    }
    return 0;
}

/* 写任务: 从队列取出数据块写入文件系统 */
static void writer_task(void *arg)
{
    (void)arg;
    for (;;) {
        wrq_msg_t *m = (wrq_msg_t *)svc_queue_recv(&s_wrq, SVC_WAIT_FOREVER);
        storage_append(m->data, m->len);         /* 你自己的落盘接口 */
        svc_mempool_free(&s_pool_hdl, m);
    }
}
#endif
