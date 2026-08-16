/*!
 * @file    ark_excel.h
 * @brief   ark_excel - 嵌入式专用 Excel(.xlsx) 生成库
 *
 * 特性:
 *   - 纯 C99, 无操作系统依赖, 支持 RTOS / 裸机 / PC
 *   - 流式 ZIP(Store 模式) 输出, 内存占用与表格规模基本无关(除单元格池)
 *   - 全静态内存分配(可在编译期确定峰值), 也可切换动态分配
 *   - 输出通过回调: 文件系统 / UART / 网络socket / 内存缓冲 均可
 *   - 支持特殊格式: 合并单元格、交错行、多级表头、冻结窗格、自动筛选、
 *     条件格式、数据验证(下拉框等)、超链接、命名区域、富样式(字体/填充/
 *     边框/对齐/数字格式)、日期时间、公式、打印设置、页眉页脚、Tab颜色、
 *     隐藏行列、自定义列宽行高、RTL 布局等
 *
 * 依据《嵌入式系统中实现 Excel (.xlsx) 生成库》设计规范实现
 */
#ifndef ARK_EXCEL_H
#define ARK_EXCEL_H

#include <stdint.h>
#include <stddef.h>
#include "ark_excel_config.h"

#ifdef __cplusplus
extern "C" {
#endif

/* ============================================================
 * 1. 错误码
 * ========================================================== */
typedef enum {
    ARK_OK = 0,

    /* 内存/容量错误 */
    ARK_ERR_NO_MEMORY,          /* 内存池耗尽 */
    ARK_ERR_TOO_MANY_SHEETS,    /* Sheet 数量超限 */
    ARK_ERR_TOO_MANY_CELLS,     /* 单元格数量超限 */
    ARK_ERR_TOO_MANY_MERGES,    /* 合并区域超限 */
    ARK_ERR_TOO_MANY_STYLES,    /* 样式超限 */
    ARK_ERR_STRING_TOO_LONG,    /* 字符串太长 */
    ARK_ERR_SST_FULL,           /* 共享字符串表已满 */
    ARK_ERR_TABLE_FULL,         /* 其他表满(列定义/超链接/验证等) */

    /* 参数错误 */
    ARK_ERR_INVALID_ARG,        /* 无效参数(NULL 指针等) */
    ARK_ERR_INVALID_ROW,        /* 无效行号(有效范围 1..1048576) */
    ARK_ERR_INVALID_COL,        /* 无效列号(有效范围 1..16384) */
    ARK_ERR_INVALID_SHEET,      /* 无效 Sheet 索引 */
    ARK_ERR_INVALID_STYLE,      /* 无效样式索引 */
    ARK_ERR_INVALID_RANGE,      /* 无效范围(r1>r2 或 c1>c2) */
    ARK_ERR_INVALID_NAME,       /* 无效 Sheet 名称 */

    /* 输出错误 */
    ARK_ERR_WRITE,              /* 输出回调返回失败 */

    /* 状态错误 */
    ARK_ERR_BAD_STATE,          /* 状态错误(如重复保存) */
    ARK_ERR_SAVED,              /* 已保存, 工作簿已锁定 */

    ARK_ERR_LIMIT = 0x7FFFFFFF
} ark_err_t;

/* ============================================================
 * 2. 输出回调 —— 可移植性核心
 *    返回 0 成功; 非 0 失败(将中断保存并透传错误)
 * ========================================================== */
typedef int (*ark_write_fn)(const void *data, size_t len, void *ctx);

/* ============================================================
 * 3. 单元格类型
 * ========================================================== */
enum {
    ARK_CELL_BLANK = 0,   /* 空单元格(可携带样式) */
    ARK_CELL_SSTR,        /* 共享字符串 */
    ARK_CELL_ISTR,        /* 内联字符串(SST 满时自动降级) */
    ARK_CELL_NUM,         /* 浮点数 */
    ARK_CELL_INT,         /* 整数(无需浮点 printf 支持) */
    ARK_CELL_DATE,        /* 日期序列号(配合日期样式) */
    ARK_CELL_FORMULA,     /* 公式(字符串) */
    ARK_CELL_BOOL,        /* 布尔 */
    ARK_CELL_ERROR        /* 错误值(#DIV/0! 等) */
};

/* ============================================================
 * 4. 对齐/颜色常量
 * ========================================================== */
enum { ARK_HALIGN_GENERAL = 0, ARK_HALIGN_LEFT, ARK_HALIGN_CENTER,
       ARK_HALIGN_RIGHT, ARK_HALIGN_FILL, ARK_HALIGN_JUSTIFY };
enum { ARK_VALIGN_TOP = 0, ARK_VALIGN_CENTER, ARK_VALIGN_BOTTOM,
       ARK_VALIGN_JUSTIFY };
enum { ARK_BORDER_NONE = 0, ARK_BORDER_THIN, ARK_BORDER_MEDIUM,
       ARK_BORDER_THICK, ARK_BORDER_DASHED, ARK_BORDER_DOTTED,
       ARK_BORDER_DOUBLE, ARK_BORDER_HAIR };

/* 常用颜色 (0xAARRGGBB, AR=FF 不透明) */
#define ARK_RGB(r,g,b)   ((uint32_t)0xFF000000u | \
                          ((uint32_t)(r) << 16) | \
                          ((uint32_t)(g) << 8) | (uint32_t)(b))
#define ARK_COLOR_BLACK    ARK_RGB(0x00,0x00,0x00)
#define ARK_COLOR_WHITE    ARK_RGB(0xFF,0xFF,0xFF)
#define ARK_COLOR_RED      ARK_RGB(0xFF,0x00,0x00)
#define ARK_COLOR_GREEN    ARK_RGB(0x00,0xB0,0x50)
#define ARK_COLOR_BLUE     ARK_RGB(0x00,0x70,0xC0)
#define ARK_COLOR_ORANGE   ARK_RGB(0xFF,0xC0,0x00)
#define ARK_COLOR_GRAY     ARK_RGB(0x80,0x80,0x80)
#define ARK_COLOR_HEADERBG ARK_RGB(0x44,0x72,0xC4)  /* 经典表头蓝 */
#define ARK_COLOR_ALTBG    ARK_RGB(0xF2,0xF2,0xF2)  /* 交错行浅灰 */
#define ARK_COLOR_WARNBG   ARK_RGB(0xFF,0xC7,0xCE)  /* 警告浅红 */
#define ARK_COLOR_WARNFG   ARK_RGB(0x9C,0x00,0x06)
#define ARK_COLOR_OKBG     ARK_RGB(0xC6,0xEF,0xCE)  /* 成功浅绿 */
#define ARK_COLOR_OKFG     ARK_RGB(0x00,0x61,0x00)

/* ============================================================
 * 5. 预定义样式索引(内部已注册, 直接使用)
 * ========================================================== */
enum {
    ARK_STYLE_DEFAULT = 0,
    ARK_STYLE_TITLE,          /* 标题: 粗体 14 居中 */
    ARK_STYLE_HEADER,         /* 表头: 粗体白字 蓝底 细边框 居中 */
    ARK_STYLE_DATA,           /* 数据: 细边框 左对齐 */
    ARK_STYLE_DATA_CENTER,    /* 数据: 细边框 居中 */
    ARK_STYLE_DATA_RIGHT,     /* 数据: 细边框 右对齐 */
    ARK_STYLE_NUM_2DEC,       /* 数字 0.00 右对齐带边框 */
    ARK_STYLE_INT,            /* 整数 0 右对齐带边框 */
    ARK_STYLE_PERCENT,        /* 百分比 0.00% */
    ARK_STYLE_DATE,           /* 日期 yyyy-mm-dd */
    ARK_STYLE_DATETIME,       /* 日期时间 yyyy-mm-dd hh:mm:ss */
    ARK_STYLE_TIME,           /* 时间 hh:mm:ss */
    ARK_STYLE_MONEY,          /* 千分位两位小数 #,##0.00 */
    ARK_STYLE_ALT_ROW,        /* 交错行: 浅灰底 细边框 */
    ARK_STYLE_WARN,           /* 警告: 浅红底 深红字 */
    ARK_STYLE_OK,             /* 成功: 浅绿底 深绿字 */
    ARK_STYLE_TOTAL,          /* 合计: 粗体 上边框 */
    ARK_STYLE_HYPERLINK,      /* 超链接: 蓝色下划线 */
    ARK_STYLE_MERGE_CENTER,   /* 合并居中: 粗体 居中 边框 */
    ARK_STYLE_BUILTIN_COUNT
};

/* ============================================================
 * 6. 自定义样式定义(用于 ark_wb_add_style)
 * ========================================================== */
typedef struct {
    /* 字体; font_name 为 NULL 时用默认 "Calibri" */
    const char *font_name;    /* 字体名(指向静态字符串, 库内不拷贝) */
    uint16_t    font_size;    /* 字号(*100, 1100 = 11.00; 0 = 默认 1100) */
    uint8_t     bold;
    uint8_t     italic;
    uint8_t     underline;    /* 0 无 1 单下划线 */
    uint32_t    font_color;   /* 0 = 默认黑 */
    /* 填充 */
    uint32_t    fill_color;   /* 0 = 无填充(透明) */
    /* 边框: 四边同风格时用 border_all; 想逐边控制见下 */
    uint8_t     border_all;   /* ARK_BORDER_xxx 应用到四边 */
    uint8_t     border_left, border_right, border_top, border_bottom;
    uint32_t    border_color; /* 边框颜色, 0 = 黑 */
    /* 对齐 */
    uint8_t     halign;       /* ARK_HALIGN_xxx */
    uint8_t     valign;       /* ARK_VALIGN_xxx */
    uint8_t     wrap_text;
    uint8_t     indent;       /* 缩进级别 0-15 */
    /* 数字格式: 内置 id 直接填(如 9/14), 或 custom_fmt 指向静态格式串 */
    uint16_t    numfmt_id;    /* 0 = General */
    const char *custom_fmt;   /* 如 "#,##0.0%;[RED]-#,##0.0%" */
} ark_style_def_t;

/* 样式定义初始化为全零的辅助宏 */
#define ARK_STYLE_DEF_INIT  { 0 }

/* ============================================================
 * 7. 数据验证类型/操作符
 * ========================================================== */
enum {
    ARK_VAL_NONE = 0,
    ARK_VAL_WHOLE,      /* 整数 */
    ARK_VAL_DECIMAL,    /* 小数 */
    ARK_VAL_LIST,       /* 列表(下拉框) */
    ARK_VAL_TEXTLEN,    /* 文本长度 */
    ARK_VAL_DATE,       /* 日期 */
    ARK_VAL_TIME        /* 时间 */
};
enum {
    ARK_OP_BETWEEN = 0, ARK_OP_NOT_BETWEEN, ARK_OP_EQUAL, ARK_OP_NOT_EQUAL,
    ARK_OP_GREATER, ARK_OP_LESS, ARK_OP_GREATER_EQUAL, ARK_OP_LESS_EQUAL
};

/* ============================================================
 * 8. 条件格式操作符(cellIs) / 类型
 * ========================================================== */
enum {
    ARK_CF_CELLIS = 0,   /* 与常量比较 */
    ARK_CF_EXPR           /* 公式表达式 */
};

/* ============================================================
 * 9. 打印设置
 * ========================================================== */
enum { ARK_PORTRAIT = 0, ARK_LANDSCAPE };

typedef struct {
    uint8_t  orientation;    /* ARK_PORTRAIT / ARK_LANDSCAPE */
    uint8_t  paper;          /* 纸张: 9=A4 8=A3 11=A5 1=Letter(默认9) */
    uint16_t scale;          /* 缩放百分比 10-400; 0 = 100 */
    uint8_t  fit_to_page;    /* 1 = 调整为页宽(忽略 scale) */
    uint8_t  fit_width;      /* fit 模式下页宽页数(默认1) */
    uint8_t  fit_height;     /* 0 = 自动 */
    /* 页边距, 单位: 英寸(*1000, 如 700 = 0.7"), 0xFFFF = 默认 */
    uint16_t margin_left, margin_right, margin_top, margin_bottom;
    uint16_t margin_header, margin_footer;
} ark_page_setup_t;

#define ARK_PAGE_SETUP_DEFAULT { ARK_PORTRAIT, 9, 100, 0, 1, 0, \
                                 0xFFFF,0xFFFF,0xFFFF,0xFFFF,0xFFFF,0xFFFF }

/* ============================================================
 * 10. 工作簿结构(透明定义, 便于静态分配)
 * ========================================================== */

/* 字符串 arena 条目: [u16 len][bytes] */
typedef struct {
    uint16_t len;
    uint16_t sst_seq;     /* 若属于 SST: 其在 SST 中的序号 */
    uint8_t  is_sst;      /* 该条目属于共享字符串表 */
    char     text[ARK_MAX_STR_LEN - 3];
} ark_str_t;

typedef struct {
    uint32_t row;        /* 1..1048576 */
    uint16_t col;        /* 1..16384 */
    uint8_t  sheet;
    uint8_t  type;       /* ARK_CELL_xxx */
    uint16_t style;      /* cellXf 索引 */
    uint16_t _pad;
    uint32_t sidx;       /* 字符串索引(arena 下标) 或 SST 索引 或 错误码 */
    double   num;        /* 数值 / 日期序列号 / 布尔(0/1) / 整数 */
} ark_cell_t;

typedef struct {
    uint8_t  sheet;
    uint8_t  _pad[3];
    uint16_t r1, c1, r2, c2;   /* 合并区域(行号最大 65535 内使用便利) */
} ark_merge_t;

typedef struct {
    uint16_t first, last;      /* 列范围 */
    uint8_t  sheet;
    uint8_t  hidden;
    uint16_t width_x100;       /* 列宽(字符单位 *100; 0=默认) */
} ark_coldef_t;

typedef struct {
    uint32_t row;
    uint16_t height_x20;       /* 行高(磅*20; 0=默认) */
    uint8_t  hidden;
    uint8_t  sheet;
} ark_rowdef_t;

typedef struct {
    uint8_t  sheet;
    uint8_t  internal;         /* 1=内部跳转(#Sheet!A1), 0=URL */
    uint16_t ref_idx;          /* 显示文本单元格引用, 如 "B3" (存 arena) */
    uint32_t target_idx;       /* 目标字符串(arena) */
    uint32_t tooltip_idx;      /* 提示(arena, 0xFFFFFFFF=无) */
} ark_hyperlink_t;

typedef struct {
    uint8_t  sheet;
    uint8_t  type;             /* ARK_VAL_xxx */
    uint8_t  op;               /* ARK_OP_xxx */
    uint8_t  allow_blank;
    uint16_t ref_idx;          /* 应用范围 sqref(arena) */
    uint16_t f1_idx;           /* formula1(arena) */
    uint16_t f2_idx;           /* formula2(arena, 0xFFFF=无) */
    uint16_t prompt_idx;       /* 输入提示(0xFFFF=无) */
    uint16_t error_idx;        /* 错误提示(0xFFFF=无) */
} ark_validation_t;

typedef struct {
    uint8_t  sheet;
    uint8_t  type;             /* ARK_CF_xxx */
    uint8_t  op;               /* ARK_OP_xxx(仅 cellIs) */
    uint8_t  priority;
    uint16_t ref_idx;          /* 应用范围 sqref(arena) */
    uint16_t f1_idx;           /* formula1(arena) */
    uint16_t f2_idx;           /* formula2(arena, 0xFFFF=无) */
    /* dxf: 差异格式 */
    uint32_t fill_color;       /* 0 = 不改填充 */
    uint32_t font_color;       /* 0 = 不改字色 */
    uint8_t  bold;
    uint8_t  _pad[3];
} ark_condfmt_t;

typedef struct {
    char     name[32];         /* 命名区域名 */
    uint16_t refers_idx;       /* 引用串(arena), 如 "Sheet1!$A$1:$D$20" */
    uint16_t _pad;
} ark_namedref_t;

typedef struct {
    char     name[ARK_SHEET_NAME_MAX];
    /* 视图 */
    uint8_t  freeze_rows, freeze_cols;  /* 冻结窗格 */
    uint8_t  rtl;              /* 从右到左布局 */
    uint8_t  show_gridlines;   /* 1 显示网格线(默认) */
    uint8_t  autofilter_set;
    uint8_t  tab_color_set;
    uint8_t  _pad[2];
    uint16_t filter_r1, filter_c1, filter_r2, filter_c2; /* 自动筛选范围 */
    uint32_t tab_color;
    uint16_t print_area_idx;   /* 打印区域(arena, 0xFFFF=无) */
    uint16_t break_rows;       /* 分页符所在行(0=无, 在此行前分页) */
    uint16_t header_idx;       /* 页眉(arena, 0xFFFF=无) */
    uint16_t footer_idx;       /* 页脚(arena, 0xFFFF=无) */
    ark_page_setup_t page;
} ark_sheet_t;

typedef struct {
    ark_write_fn write;        /* 输出回调(不可为 NULL) */
    void        *ctx;          /* 回调上下文 */

    ark_sheet_t  sheets[ARK_MAX_SHEETS];
    uint8_t      sheet_count;
    uint8_t      saved;        /* 保存后置 1, 锁定 */

    ark_cell_t   cells[ARK_MAX_CELLS];
    uint32_t     cell_count;

    ark_merge_t  merges[ARK_MAX_MERGES];
    uint16_t     merge_count;

    ark_coldef_t coldefs[ARK_MAX_COLDEFS];
    uint16_t     coldef_count;

    ark_rowdef_t rowdefs[ARK_MAX_ROWDEFS];
    uint16_t     rowdef_count;

    ark_hyperlink_t hyperlinks[ARK_MAX_HYPERLINKS];
    uint16_t     hyperlink_count;

    ark_validation_t validations[ARK_MAX_VALIDATIONS];
    uint16_t     validation_count;

    ark_condfmt_t condfmts[ARK_MAX_CONDFMTS];
    uint16_t     condfmt_count;

    ark_namedref_t names[ARK_MAX_NAMES];
    uint16_t     name_count;

    /* 读改写: 非 NULL 时 save 走增量保存(见 ark_reader.h) */
    void        *reader;

    /* 字符串 arena: SST 与普通引用字符串共用 */
    ark_str_t    strs[ARK_MAX_STR];
    uint16_t     str_count;
    uint16_t     sst_count;    /* 前 sst_count 条为共享字符串 */

    /* 样式表 */
    uint16_t     font_count, fill_count, border_count, xf_count;
    uint16_t     numfmt_count;             /* 自定义 numFmt 个数 */
    uint16_t     cellstylexf_count;        /* 固定为 1 */
    ark_err_t    last_err;
} ark_wb_t;

/* ============================================================
 * 11. API —— 生命周期
 * ========================================================== */

/*!
 * @brief 初始化工作簿(重复调用前先 ark_wb_reset)
 *   读改写: ark_wb_init 后可再调 ark_wb_open 载入现有文件
 * @param wb    工作簿(建议静态分配或由调用者管理生命周期)
 * @param write 输出回调, 保存时所有字节流经此函数
 * @param ctx   传递给回调的用户上下文(如 FILE* / UART 句柄)
 * @return ARK_OK 或错误码
 */
ark_err_t ark_wb_init(ark_wb_t *wb, ark_write_fn write, void *ctx);

/*! 重置工作簿(不清零回调), 可复用结构体生成下一个文件 */
void ark_wb_reset(ark_wb_t *wb);

/*! 添加 Sheet, 返回索引(>=0) 或负错误码 */
int ark_wb_add_sheet(ark_wb_t *wb, const char *name);

/*!
 * @brief 保存并输出 xlsx 全部字节到回调流
 * 保存成功后工作簿锁定, 需 ark_wb_reset 后才能复用
 */
ark_err_t ark_wb_save(ark_wb_t *wb);

/* ============================================================
 * 12. API —— 单元格写入
 *    row/col 均从 1 开始; style 传 ARK_STYLE_xxx 或自定义 id
 * ========================================================== */
ark_err_t ark_set_str    (ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                          const char *text, uint16_t style);
ark_err_t ark_set_num    (ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                          double value, uint16_t style);
ark_err_t ark_set_int    (ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                          int64_t value, uint16_t style);  /* 无浮点printf也可用 */
ark_err_t ark_set_bool   (ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                          int value, uint16_t style);
ark_err_t ark_set_formula(ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                          const char *formula, uint16_t style); /* 不带'=' */
ark_err_t ark_set_error  (ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                          const char *errval, uint16_t style);  /* "#DIV/0!" 等 */
ark_err_t ark_set_blank  (ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                          uint16_t style);   /* 仅样式, 用于合并区域边框 */

/*! 日期/时间: 自动转 Excel 序列号, 搭配 ARK_STYLE_DATE 等 */
ark_err_t ark_set_date    (ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                           int y, int mo, int d, uint16_t style);
ark_err_t ark_set_datetime(ark_wb_t *wb, int sheet, uint32_t row, uint16_t col,
                           int y, int mo, int d, int h, int mi, int s,
                           uint16_t style);

/* ============================================================
 * 13. API —— 特殊格式
 * ========================================================== */

/*! 合并单元格(闭区间); Excel 要求仅左上角单元格有值 */
ark_err_t ark_merge(ark_wb_t *wb, int sheet,
                    uint32_t r1, uint16_t c1, uint32_t r2, uint16_t c2);

/*! 列宽(字符单位, 如 12.5); first/last 为列范围 */
ark_err_t ark_set_col_width(ark_wb_t *wb, int sheet, int first, int last,
                            double width);
ark_err_t ark_set_col_hidden(ark_wb_t *wb, int sheet, int first, int last,
                             int hidden);
/*! 行高(磅); 每行一条记录 */
ark_err_t ark_set_row_height(ark_wb_t *wb, int sheet, uint32_t row,
                             double height_pt);
ark_err_t ark_set_row_hidden(ark_wb_t *wb, int sheet, uint32_t row, int hidden);

/*! 冻结窗格: 冻结前 freeze_rows 行 + freeze_cols 列 */
ark_err_t ark_freeze(ark_wb_t *wb, int sheet, int freeze_rows, int freeze_cols);

/*! 自动筛选范围(通常为表头行) */
ark_err_t ark_autofilter(ark_wb_t *wb, int sheet,
                         uint32_t r1, uint16_t c1, uint32_t r2, uint16_t c2);

/*! Sheet Tab 颜色 */
ark_err_t ark_tab_color(ark_wb_t *wb, int sheet, uint32_t argb);

/*! RTL / 网格线 */
ark_err_t ark_sheet_rtl(ark_wb_t *wb, int sheet, int enable);
ark_err_t ark_sheet_gridlines(ark_wb_t *wb, int sheet, int show);

/*! 超链接: ref 为单元格地址如 "B3"; internal=1 时 target 为 "#Sheet2!A1" */
ark_err_t ark_hyperlink(ark_wb_t *wb, int sheet, const char *ref,
                        const char *target, int internal, const char *tooltip);

/*! 数据验证. list 类型时 f1 形如 "\"是,否,待定\""(带引号逗号分隔);
 *  整数/小数时 f1/f2 为数值串, 如 "0" "100" */
ark_err_t ark_add_validation(ark_wb_t *wb, int sheet, const char *range,
                             int type, int op,
                             const char *f1, const char *f2,
                             int allow_blank,
                             const char *prompt, const char *error);

/*! 条件格式(cellIs / expression), dxf 仅支持填充色+字色+粗体 */
ark_err_t ark_add_condfmt(ark_wb_t *wb, int sheet, const char *range,
                          int type, int op,
                          const char *f1, const char *f2,
                          uint32_t fill_argb, uint32_t font_argb, int bold);

/*! 命名区域: refers_to 如 "Sheet1!$A$1:$D$20" */
ark_err_t ark_named_range(ark_wb_t *wb, const char *name, const char *refers_to);

/*! 打印: 页面设置 / 打印区域(如 "A1:F30") / 页眉页脚 */
ark_err_t ark_page_setup(ark_wb_t *wb, int sheet, const ark_page_setup_t *ps);
ark_err_t ark_print_area (ark_wb_t *wb, int sheet, const char *area);
ark_err_t ark_header_footer(ark_wb_t *wb, int sheet,
                            const char *header, const char *footer);
/*! 手动行分页符: 在该行之前分页(仅支持一处) */
ark_err_t ark_row_break(ark_wb_t *wb, int sheet, uint32_t row);

/* ============================================================
 * 14. API —— 自定义样式
 * ========================================================== */

/*! 注册自定义样式, 返回样式 id(即 cellXf 索引, >=ARK_STYLE_BUILTIN_COUNT) */
int ark_wb_add_style(ark_wb_t *wb, const ark_style_def_t *def);

/* ============================================================
 * 15. 工具函数
 * ========================================================== */

/*! 公历日期 -> Excel 序列号(处理 1900 闰年 bug, 基准 1899-12-31) */
double ark_date_to_serial(int y, int mo, int d, int h, int mi, int s);

/*! 列号(1..16384) -> 列字母("A".."XFD"); buf 至少 5 字节 */
void ark_col_to_letters(uint16_t col, char *buf);

const char *ark_errstr(ark_err_t e);

#ifdef __cplusplus
}
#endif
#endif /* ARK_EXCEL_H */
