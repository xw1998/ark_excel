#include "ark_excel_config.h"
#if ARK_ENABLE_READWRITE
/*!
 * @file    ark_reader.h
 * @brief   ark_excel 读改写模块 —— 打开/修改/保存现有 xlsx
 *
 * 架构:
 *   ark_wb_open()  : 通过读取回调解析现有 xlsx(ZIP+XML) 进 ark_wb_t
 *   正常 API 修改  : ark_set_* / ark_merge / ark_wb_add_style 等
 *   ark_wb_save()  : 增量保存 —— 未修改 ZIP 条目原样字节复制,
 *                    被修改部件(目标 sheet / SST / styles)重新生成,
 *                    模板其余 sheet 的所有特性零损失保留
 *
 * 依赖: 自带 DEFLATE 解压(ark_inflate), 无 zlib 依赖
 */
#ifndef ARK_READER_H
#define ARK_READER_H

#include <stdint.h>
#include <stddef.h>
#include "ark_excel.h"

#ifdef __cplusplus
extern "C" {
#endif

/* ============================================================
 * 1. 读取回调 —— 与 ark_write_fn 对应
 * ========================================================== */

/*! 从 pos 处读取最多 len 字节到 buf, 返回实际读取数(0=EOF/失败) */
typedef uint32_t (*ark_read_fn)(void *ctx, uint8_t *buf, uint32_t len);

/*! 绝对偏移定位, 返回 0 成功 */
typedef int (*ark_seek_fn)(void *ctx, uint32_t pos);

typedef struct {
    ark_read_fn  read;
    ark_seek_fn  seek;
    void        *ctx;
} ark_reader_io_t;

/* ============================================================
 * 2. 内部结构(透明, 便于静态分配)
 * ========================================================== */

typedef struct {
    char     name[64];       /* ZIP 条目名, 如 "xl/worksheets/sheet1.xml" */
    uint32_t local_off;      /* 本地文件头偏移 */
    uint32_t crc;            /* 原始内容 CRC32 */
    uint32_t csize;          /* 压缩后大小 */
    uint32_t usize;          /* 原始大小 */
    uint16_t method;         /* 0=Store 8=Deflate */
    uint8_t  dirty;          /* 保存时需重新生成 */
} ark_zentry_t;

typedef struct {
    ark_reader_io_t io;

    ark_zentry_t entries[ARK_MAX_ZIP_ENTRIES];
    uint16_t     entry_count;
    uint32_t     file_size;

    /* 部件索引(指向 entries 下标, 0xFFFF=不存在) */
    uint16_t i_content_types;
    uint16_t i_workbook;
    uint16_t i_rels;
    uint16_t i_sst;
    uint16_t i_styles;
    uint16_t sheet_entry[ARK_MAX_SHEETS];  /* 每个 sheet 的 worksheet 部件 */

    char     sheet_part[ARK_MAX_SHEETS][64]; /* workbook 顺序的部件路径 */

    /* 模板样式计数(新增样式追加在其后) */
    uint16_t tpl_font_cnt, tpl_fill_cnt, tpl_border_cnt, tpl_xf_cnt;
    uint16_t tpl_numfmt_cnt;      /* 模板 numFmt 条数 */
    uint16_t tpl_numfmt_maxid;    /* 模板 numFmt 最大自定义 id */
    uint16_t tpl_cellstyle_cnt;   /* cellStyles 条目数 */
    uint16_t tpl_dxf_cnt;         /* 模板 dxfs 条目数(新 dxfId 从此续编) */

    /* 内置样式重映射: ARK_STYLE_xxx -> 模板追加后的实际 id */
    uint16_t style_map[ARK_STYLE_BUILTIN_COUNT];
    uint8_t  style_map_ok;

    /* sheet 修改标志(索引对应 wb->sheets) */
    uint8_t  sheet_dirty[ARK_MAX_SHEETS];
    uint8_t  sst_dirty;           /* 新增了共享字符串 */
    uint8_t  wb_dirty;            /* 结构变化(新增 sheet), 需重生成 workbook */
    uint8_t  styles_dirty;        /* 注册了新样式 */

    /* 共享工作缓冲(解压/解析复用) */
    uint8_t  buf[ARK_OPEN_XML_BUF];
    uint32_t buf_len;

    uint8_t  tpl_sheet_cnt;       /* 模板原有 sheet 数(其后为用户新增) */
    uint8_t  opened;              /* open 模式标志 */
} ark_reader_t;

/* ============================================================
 * 3. API
 * ========================================================== */

/*!
 * @brief 打开现有 xlsx 并载入 ark_wb_t
 *
 * 载入内容: 所有 sheet 名称/单元格/合并/列宽行高/冻结/筛选/
 *           超链接/验证/条件格式, SST, 模板样式计数。
 * 未解析的部件(chart/theme/pivotTable 等)保存时原样复制, 零损失。
 *
 * 载入后:
 *   - 可用 ark_set_* 覆盖/追加单元格(原样式自动保留)
 *   - 新样式用 ark_wb_add_style 注册(追加在模板样式之后)
 *   - 内置样式 id 经 ark_style_map() 重映射
 *   - ark_wb_save() 自动走增量保存路径
 *
 * @param wb 工作簿(先 ark_wb_init 或手动挂载 write 回调)
 * @param rd 读取器(静态分配)
 * @param io 读取回调(seek+read, 文件系统/内存均可包装)
 * @return ARK_OK 或错误码(ARK_ERR_NO_MEMORY=XML 超出 ARK_OPEN_XML_BUF)
 */
ark_err_t ark_wb_open(ark_wb_t *wb, ark_reader_t *rd, const ark_reader_io_t *io);

/*! 内置样式 id -> open 模式实际 id; 未 open 或未注册时原样返回 */
uint16_t ark_style_map(uint16_t builtin_id);

/*! 查询单元格: 命中返回 0 并填充 *type/*num/*text/*style, 未命中返回 1 */
int ark_get_cell(ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                 uint8_t *type, double *num, char *text, size_t textcap,
                 uint16_t *style);

#ifdef __cplusplus
}
#endif
#endif /* ARK_READER_H */
#endif
