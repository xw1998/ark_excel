/*!
 * @file    ark_excel_internal.h
 * @brief   ark_excel 内部跨编译单元接口(不对外发布)
 */
#ifndef ARK_EXCEL_INTERNAL_H
#define ARK_EXCEL_INTERNAL_H

#include "ark_excel.h"
#if ARK_ENABLE_READWRITE
#include "ark_reader.h"
#else
/* 读改写被裁剪: touch 变为空操作, 无需链接 ark_reader.c */
#define ark_rd_touch(wb, sheet)          ((void)0)
#define ark_rd_touch_sst(wb)             ((void)0)
#define ark_rd_touch_styles(wb)          ((void)0)
#define ark_rd_touch_workbook(wb)        ((void)0)
#endif

/* ---- arena/SST(ark_excel.c 实现, reader 复用) ---- */
int         ark_str_add(ark_wb_t *wb, const char *s);
const char *ark_str_get(ark_wb_t *wb, uint16_t idx);
int         ark_sst_add(ark_wb_t *wb, const char *s);

/* ---- 样式表同步(ark_excel.c 实现) ---- */
/*! open 模式: 把 s_st 计数置为模板值, 随后注册的样式追加在其后.
 *  max_numfmt_id: 模板 numFmt 最大自定义 id(无则 164), 内部换算 nnumfmts */
void ark_i_st_sync_open(uint16_t fonts, uint16_t fills, uint16_t borders,
                        uint16_t xfs, uint16_t max_numfmt_id);
/*! open 模式下注册全部内置样式(追加), 填充 map[ARK_STYLE_BUILTIN_COUNT] */
void ark_i_st_register_builtins_open(uint16_t *map);

/* ---- 读改写标记(ark_reader.c 实现, ark_excel.c 各写 API 调用) ---- */
#if ARK_ENABLE_READWRITE
void ark_rd_touch(ark_wb_t *wb, uint8_t sheet);   /* sheet 部件需重新生成 */
void ark_rd_touch_sst(ark_wb_t *wb);              /* 新增共享字符串 */
void ark_rd_touch_styles(ark_wb_t *wb);           /* 新增样式 */
void ark_rd_touch_workbook(ark_wb_t *wb);         /* 结构变化(新 sheet 等) */
#endif

#if ARK_ENABLE_READWRITE
/*! 解压第 idx 个 ZIP 条目到 rd->buf(供增量保存读 styles 文本) */
int ark_i_entry_extract(ark_reader_t *rd, uint16_t idx);
#endif

#endif /* ARK_EXCEL_INTERNAL_H */
