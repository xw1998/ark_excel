/*!
 * @file    ark_inflate.h
 * @brief   自包含 DEFLATE (RFC 1951) 解压 - puff 风格, 无 zlib 依赖
 *
 * 输入采用"拉取"回调, 压缩数据分块读入, RAM 占用 =
 * 调用方输出缓冲 + 内部 1KB 输入暂存。
 */
#ifndef ARK_INFLATE_H
#define ARK_INFLATE_H

#include <stdint.h>
#include <stddef.h>

#ifdef __cplusplus
extern "C" {
#endif

/*! 拉取回调: 向 buf 填入最多 len 字节的后续压缩数据, 返回实际字节数(0=EOF) */
typedef uint32_t (*ark_inflate_pull)(void *ctx, uint8_t *buf, uint32_t len);

/*!
 * @brief 解压完整 deflate 流(gzip 的 payload 部分, 即 zip method=8 的数据)
 * @return 0 成功; 负值错误(-1 数据错, -2 输出不够)
 * @param outlen 输出实际字节数
 */
int ark_inflate(ark_inflate_pull pull, void *ctx,
                uint8_t *out, size_t outcap, size_t *outlen);

#ifdef __cplusplus
}
#endif
#endif
