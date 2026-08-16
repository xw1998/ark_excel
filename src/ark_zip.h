/*!
 * @file    ark_zip.h
 * @brief   最小流式 ZIP 写入器(Store 模式, 无压缩, 无 libc 依赖)
 *
 * 两遍法流式输出(对 UART 等不可回写的只追加流同样正确):
 *   ark_zip_begin(&z, write, ctx, NULL, 0);
 *   对每个文件:
 *     ark_zip_measure_start(&z);               // 第一遍: 仅测量
 *     ... 生成数据, 调 ark_zip_measure_write() ...
 *     ark_zip_open_stream(&z, "a.xml");        // 用测量值写 local header
 *     ... 再次生成同样数据, 调 ark_zip_write() // 第二遍: 真实输出
 *     ark_zip_close_stream(&z);
 *   ark_zip_end(&z);                           // 中央目录 + EOCD
 *
 * 优点: 每文件数据零缓冲, ZIP 层内存占用为 O(条目数).
 */
#ifndef ARK_ZIP_H
#define ARK_ZIP_H

#include <stdint.h>
#include <stddef.h>

#ifdef __cplusplus
extern "C" {
#endif

#ifndef ARK_ZIP_MAX_ENTRIES
#define ARK_ZIP_MAX_ENTRIES 32
#endif

typedef int (*ark_zip_write_fn)(const void *data, size_t len, void *ctx);

typedef struct {
    char     name[64];
    uint32_t crc;
    uint32_t size;                 /* Store: 内容大小 */
    uint32_t usize;                /* Deflate: 原始大小(size 为压缩后) */
    uint32_t offset;       /* local header 在 ZIP 中的偏移 */
    uint16_t dostime, dosdate;
    uint16_t method;               /* 0=Store 8=Deflate */
} ark_zip_entry_t;

typedef struct {
    ark_zip_write_fn write;
    void            *ctx;
    int              out_failed;          /* 输出回调曾返回失败 */
    uint32_t         out_pos;             /* 已输出总字节数 */
    ark_zip_entry_t  entries[ARK_ZIP_MAX_ENTRIES];
    uint16_t         entry_count;
    /* 第一遍测量状态 */
    uint32_t         m_crc, m_size;
    /* 第二遍流式状态 */
    uint16_t         cur;                 /* 当前条目下标, 0xFFFF=无 */
} ark_zip_t;

void    ark_zip_begin(ark_zip_t *z, ark_zip_write_fn write, void *ctx);

/* ---- 第一遍: 测量 ---- */
void    ark_zip_measure_start(ark_zip_t *z);
void    ark_zip_measure_write(ark_zip_t *z, const void *data, size_t len);

/* ---- 第二遍: 真实输出 ---- */
/* 用最近一次测量结果打开条目并立即写出 local header; 0 成功 */
int     ark_zip_open_stream (ark_zip_t *z, const char *name);
/*! 底层入口: 指定 method(0=Store 8=Deflate)/crc/size 直接开条目,
 *  用于读改写时原样字节复制模板条目(无需 measure) */
int     ark_zip_open_raw    (ark_zip_t *z, const char *name,
                             uint16_t method, uint32_t crc,
                             uint32_t usize, uint32_t csize);
/* 流式写出数据(必须与测量遍字节一致); 0 成功 */
int     ark_zip_write       (ark_zip_t *z, const void *data, size_t len);
/* 关闭条目(登记中央目录, 数据已在 open/write 中输出); 0 成功 */
int     ark_zip_close_stream(ark_zip_t *z);

/* 写中央目录 + EOCD; 0 成功 */
int     ark_zip_end         (ark_zip_t *z);

/* CRC32 (IEEE 802.3 反射 0xEDB88320), 增量接口 */
uint32_t ark_crc32(uint32_t crc, const void *data, size_t len);

#ifdef __cplusplus
}
#endif
#endif
