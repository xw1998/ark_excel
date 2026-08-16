/*!
 * @file    ark_excel_config.h
 * @brief   ark_excel 编译期配置 —— 按目标 MCU 资源裁剪
 *
 * 全部容量为静态分配, 内存峰值可在编译期精确确定.
 * 典型内存占用估算(默认配置):
 *   单元格池   2048 * 24B  = 48 KB
 *   字符串池   1024 * (2+96)B ≈ 100 KB  <- 通常是最大头, 小内存时调小
 *   其余表     < 4 KB
 * 如需更低占用: 减小 ARK_MAX_STR / ARK_MAX_CELLS / ARK_MAX_SST.
 */
#ifndef ARK_EXCEL_CONFIG_H
#define ARK_EXCEL_CONFIG_H

/* ---------- 容量配置 ---------- */
#ifndef ARK_MAX_SHEETS
#define ARK_MAX_SHEETS          8       /* 最大 Sheet 数 */
#endif
#ifndef ARK_MAX_CELLS
#define ARK_MAX_CELLS           2048    /* 最大单元格总数(所有 Sheet 共享) */
#endif
#ifndef ARK_MAX_MERGES
#define ARK_MAX_MERGES          64      /* 最大合并区域数 */
#endif
#ifndef ARK_MAX_COLDEFS
#define ARK_MAX_COLDEFS         32      /* 列定义条目(宽度/隐藏) */
#endif
#ifndef ARK_MAX_ROWDEFS
#define ARK_MAX_ROWDEFS         64      /* 行定义条目(高度/隐藏) */
#endif
#ifndef ARK_MAX_HYPERLINKS
#define ARK_MAX_HYPERLINKS      16
#endif
#ifndef ARK_MAX_VALIDATIONS
#define ARK_MAX_VALIDATIONS     16
#endif
#ifndef ARK_MAX_CONDFMTS
#define ARK_MAX_CONDFMTS        16
#endif
#ifndef ARK_MAX_NAMES
#define ARK_MAX_NAMES           8
#endif
#ifndef ARK_MAX_STYLES
#define ARK_MAX_STYLES          64      /* cellXfs 最大数(含预定义 19 个) */
#endif
#ifndef ARK_MAX_FONTS
#define ARK_MAX_FONTS           32
#endif
#ifndef ARK_MAX_FILLS
#define ARK_MAX_FILLS           32
#endif
#ifndef ARK_MAX_BORDERS
#define ARK_MAX_BORDERS         32
#endif
#ifndef ARK_MAX_NUMFMTS
#define ARK_MAX_NUMFMTS         16      /* 自定义数字格式 */
#endif

/* 字符串池: SST(共享字符串) + 引用串(范围/公式/超链接目标等) 共用 */
#ifndef ARK_MAX_STR
#define ARK_MAX_STR             1024    /* 条目数 */
#endif
#ifndef ARK_MAX_SST
#define ARK_MAX_SST             768     /* 其中可用作共享字符串的上限 */
#endif
#ifndef ARK_MAX_STR_LEN
#define ARK_MAX_STR_LEN         96      /* 单条字符串最大字节(含\0, UTF-8) */
#endif
#ifndef ARK_SHEET_NAME_MAX
#define ARK_SHEET_NAME_MAX      28      /* Sheet 名最大字节 */
#endif

/* 读改写(ark_reader)配置 */
#ifndef ARK_MAX_ZIP_ENTRIES
#define ARK_MAX_ZIP_ENTRIES      64      /* 模板 ZIP 最大条目数 */
#endif
#ifndef ARK_ZIP_MAX_ENTRIES
#define ARK_ZIP_MAX_ENTRIES      64      /* 写出 ZIP 最大条目数 */
#endif
#ifndef ARK_OPEN_XML_BUF
#define ARK_OPEN_XML_BUF         16384   /* 单个 XML 部件解压缓冲(字节) */
#endif

/* ZIP 流式写出时的分块缓冲 */
#ifndef ARK_IO_CHUNK
#define ARK_IO_CHUNK            256
#endif

/* ---------- 功能裁剪开关(省 flash; 深度裁内存用上面的 ARK_MAX_*) ----------
 *
 * 置 0 即从链接中剔除对应模块代码:
 *   ARK_ENABLE_READWRITE   读改写(打开/修改/保存现有 xlsx, ark_reader.c)
 *   ARK_ENABLE_INFLATE     DEFLATE 解压(读改写打开 Deflate 压缩模板用;
 *                           置 0 后只能打开 Store(未压缩)模板, 可省 ~3KB flash)
 *   ARK_ENABLE_FORMULA     公式写入与载入
 *   ARK_ENABLE_CONDFMT     条件格式(含 dxf 样式生成)
 *   ARK_ENABLE_VALIDATION  数据验证(下拉框等)
 *   ARK_ENABLE_HYPERLINK   超链接
 *
 * 裁剪组合建议:
 *   极小 ROM 裸机(只写数据报表):
 *     ARK_ENABLE_READWRITE=0 ARK_ENABLE_FORMULA=0 ARK_ENABLE_CONDFMT=0
 *     ARK_ENABLE_VALIDATION=0 ARK_ENABLE_HYPERLINK=0
 *   只写不改但模板是 Excel/openpyxl 生成的(默认 Deflate 压缩):
 *     保持 READWRITE=1 INFLATE=1, 其余可关
 *   读改写但模板确定是 ark_excel 自己生成的(Store 模式):
 *     INFLATE=0 可省掉解压器
 * 注意: READWRITE=0 时 INFLATE 自动失效; 被裁功能的 API 仍存在,
 *       调用返回 ARK_ERR_BAD_STATE, 不会静默丢数据.
 */
#ifndef ARK_ENABLE_READWRITE
#define ARK_ENABLE_READWRITE     1
#endif
#ifndef ARK_ENABLE_INFLATE
#define ARK_ENABLE_INFLATE       1
#endif
#ifndef ARK_ENABLE_FORMULA
#define ARK_ENABLE_FORMULA       1
#endif
#ifndef ARK_ENABLE_CONDFMT
#define ARK_ENABLE_CONDFMT       1
#endif
#ifndef ARK_ENABLE_VALIDATION
#define ARK_ENABLE_VALIDATION    1
#endif
#ifndef ARK_ENABLE_HYPERLINK
#define ARK_ENABLE_HYPERLINK     1
#endif

/* 兼容旧开关: 定义后跳过 docProps(core.xml/app.xml), 少 2 个 ZIP 条目, 更省时 */
#ifndef ARK_EXCEL_SKIP_DOCPROPS
/* #define ARK_EXCEL_SKIP_DOCPROPS 1 */
#endif

/* 调试日志: 定义后通过 ARK_LOG 宏输出. 默认提供 printf 版, 可在
 * 编译选项中重定义 ARK_LOG 为 RTOS 日志接口 */
#if defined(ARK_EXCEL_DEBUG)
  #ifndef ARK_LOG
    #include <stdio.h>
    #define ARK_LOG(fmt, ...)  printf("[ark_excel] " fmt "\n", ##__VA_ARGS__)
  #endif
#else
  #ifndef ARK_LOG
    #define ARK_LOG(fmt, ...)  ((void)0)
  #endif
#endif

/* 数值格式化: 默认用 snprintf. 若工具链不支持浮点 printf,
 * 全部用 ark_set_int() 写整数单元格即可(内部手写转换, 不依赖 printf) */
#ifndef ARK_SNPRINTF
#include <stdio.h>
#define ARK_SNPRINTF  snprintf
#endif

#endif /* ARK_EXCEL_CONFIG_H */
