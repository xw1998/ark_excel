# ark_excel —— 嵌入式专用 Excel(.xlsx) 生成库

纯 C99、零动态内存、可静态全分配、输出全回调的 Excel 生成库。
面向 RTOS（FreeRTOS / RT-Thread / svcrtos…）与裸机环境，也可在 PC 上使用。

依据《嵌入式系统中实现 Excel (.xlsx) 生成库》设计规范实现。

## 特性

- **合法 .xlsx**：ZIP(Store) + OOXML，Excel / WPS / LibreOffice 直接打开
- **流式输出**：两遍法（先测 CRC/大小，再流式写出），对 UART 等不可回写流也正确，ZIP 层零数据缓冲
- **可移植**：不依赖任何 OS/文件系统；所有字节经 `ark_write_fn` 回调输出
- **读改写**：打开现有 .xlsx（含 Excel/openpyxl 生成的 Deflate 压缩文件）→ 修改 → 保存；
  自带 DEFLATE 解压（无 zlib 依赖）；未修改部件原样字节复制，模板其余特性（图表/主题/透视表等）零损失保留；
  新样式以增量方式追加进模板样式表
- **资源可控**：所有容量在 `ark_excel_config.h` 编译期确定，无隐藏 malloc
- **特殊格式支持**：
  合并单元格（多级表头/交错合并）、交错行变色、自定义样式（字体/填充/边框/对齐/缩进/数字格式）、
  19 个预定义样式、冻结窗格、自动筛选、隐藏行列、列宽行高、条件格式（cellIs/expression）、
  数据验证（下拉框等）、超链接（URL/内部跳转/提示）、命名区域、公式、错误值、布尔值、
  日期时间（含 1900 闰年 bug 修正）、Tab 颜色、RTL、网格线开关、
  打印设置（纸张/方向/缩放/适应页宽/边距/页眉页脚/打印区域/分页符）、多 Sheet、SST 去重与内联字符串自动降级

## 目录结构

```
ark_excel/
├── ark_excel.h            公共 API（含完整结构定义，便于静态分配）
├── ark_excel_config.h     编译期容量/行为配置（按 MCU 裁剪）
├── src/
│   ├── ark_excel.c        核心：数据模型 + 全部 XML 生成 + 增量保存
│   ├── ark_excel_internal.h  内部跨单元接口（不对外发布）
│   ├── ark_zip.h/.c       最小流式 ZIP 写入器 + CRC32 + 原样条目复制
│   ├── ark_reader.h/.c    读改写：ZIP 解析 + xlsx 载入（可整体裁剪）
│   └── ark_inflate.h/.c   自包含 DEFLATE 解压（puff 风格，可裁剪）
├── port/
│   └── port_examples.c    各平台输出回调示例(stdio/membuf/UART/FatFS/svcrtos)
├── examples/
│   ├── demo_report.c      PC 全功能演示（生成 demo_report.xlsx）
│   └── demo_baremetal.c   裸机/内存缓冲/UART 示例
└── tests/
    ├── test_features.c    20 个单特性生成用例
    ├── validate_features.py  openpyxl 76 断言 + Excel COM 实测
    ├── test_rw.c          读改写冒烟测试（8 用例 31 断言）
    ├── validate_rw.py     读改写输出 openpyxl 校验
    ├── com_rw.py          读改写输出 Excel COM 批量实测
    ├── stress_test.c + validate.py  压力回归
    └── validate.py        （旧版保留）
```

## 快速上手（3 步）

```c
#include "ark_excel.h"

static ark_wb_t wb;                     /* 静态分配 */

/* 1. 初始化：给一个输出回调（这里是 stdio 文件） */
ark_wb_init(&wb, ark_write_stdio, f);

/* 2. 填数据 */
int sh = ark_wb_add_sheet(&wb, "Report");
ark_set_str(&wb, sh, 1, 1, "标题", ARK_STYLE_TITLE);
ark_merge(&wb, sh, 1, 1, 1, 6);
ark_set_num(&wb, sh, 2, 1, 3.14, ARK_STYLE_NUM_2DEC);

/* 3. 保存：xlsx 全部字节流经回调 */
ark_wb_save(&wb);
```

## 移植指南

### 1. 接入文件系统（FatFS / LittleFS / POSIX…）

只需实现一个回调：

```c
int my_write(const void *data, size_t len, void *ctx)
{
    UINT bw;
    return (f_write((FIL *)ctx, data, len, &bw) == FR_OK && bw == len) ? 0 : -1;
}
/* f_open(...); ark_wb_init(&wb, my_write, &fil); ... f_close(...); */
```

### 2. 无文件系统（裸机）

- **UART/网络直发**：回调里调用你的发送函数，库是纯流式，零缓冲；
- **外部 Flash**：回调里按块写入你的 Flash 写接口。

### 3. RTOS 集成

单线程生成（推荐）或由调用方加互斥锁。长报表可分批 `ark_set_*` 填充，
`ark_wb_save` 的输出是分块的（默认 256B/次回调），不会长时间独占 CPU。

svcrtos（https://gitee.com/xw19981010/svcrtos_new）参考接线见 `port/port_examples.c`
第 5 节：生成任务 → 内存池+队列 → 写任务落盘，天然背压与线程安全。

### 4. 裁剪功能与内存

`ark_excel_config.h` 提供两类裁剪：

**功能开关（省 flash，`ARK_ENABLE_*`，默认全开）**

| 开关 | 控制内容 | 关闭效果 |
|------|---------|---------|
| `ARK_ENABLE_READWRITE` | 读改写（`ark_reader.c`） | 生成-only，省 ~15KB flash |
| `ARK_ENABLE_INFLATE` | DEFLATE 解压 | 只能打开 Store(未压缩)模板，省 ~3KB |
| `ARK_ENABLE_FORMULA` | 公式写入/载入 | `ark_set_formula` 返回 `ARK_ERR_BAD_STATE` |
| `ARK_ENABLE_CONDFMT` | 条件格式 | 同上，API 显式报错不静默丢数据 |
| `ARK_ENABLE_VALIDATION` | 数据验证 | 同上 |
| `ARK_ENABLE_HYPERLINK` | 超链接 | 同上 |

实测代码体积（gcc -Os，库本体 .text）：

| 配置 | 大小 | 说明 |
|------|------|------|
| 全功能（默认） | 48.8 KB | 含读改写 + inflate |
| `READWRITE=0` | 33.1 KB | 只写，省 32% |
| `READWRITE=0 FORMULA=0 CONDFMT=0 VALIDATION=0 HYPERLINK=0` | 31.6 KB | 极小 ROM 裸机数据报表 |

常用组合：
```c
#define ARK_ENABLE_READWRITE   0   // 只写不改
#define ARK_ENABLE_INFLATE     0   // 模板确定是 ark_excel 自己生成的(Store)
// 极小 ROM: 再关 FORMULA/CONDFMT/VALIDATION/HYPERLINK
```

**容量裁剪（省 RAM，`ARK_MAX_*`）。默认配置 RAM 估算：

| 项 | 大小 |
|---|---|
| 单元格池 2048×24B | 48 KB |
| 字符串池 1024×98B | ≈98 KB |
| 其余表 | < 4 KB |

STM32F0/F1 级别建议：`ARK_MAX_CELLS 256`、`ARK_MAX_STR 256`、`ARK_MAX_STR_LEN 48`，
整体可压到 < 30 KB；配合 UART 直发，RAM 峰值 ≈ 单元格池 + 3 KB。

### 5. 无浮点 printf 的平台

用 `ark_set_int()`（内部手写 int64 转换，不依赖 `%f/%g`）与 `ark_set_str()`。
`ark_set_num/ark_set_date` 需要工具链支持 `snprintf("%.15g")`。

## 构建与测试

### 快速验证（PC）

```bash
# gcc/clang/zig cc 任一, gcc 4.8 请加 -std=gnu99
gcc -std=gnu99 -I. examples/demo_report.c src/ark_excel.c src/ark_zip.c \
    src/ark_reader.c src/ark_inflate.c -o demo      # 读改写需要后两个文件
./demo && python tests/validate.py demo_report.xlsx

# 嵌入式工程: src/*.c 全部 4 个 .c + 4 个头文件加入工程即可;
# 若 -DARK_ENABLE_READWRITE=0 可只编 ark_excel.c + ark_zip.c
```

CMake / Makefile 已附带。

### 测试体系（tests/ 目录）

四层测试架构，当前状态：**openpyxl 76/76 | Excel COM 20/20 | 读改写 31/31+8/8+10/10 | 压力回归 PASS**。

| 层 | 文件 | 内容 |
|----|------|------|
| ① 单特性生成 | `test_features.c` | 19 个特性用例各生成独立 xlsx（基础类型/转义/合并/交错行/日期/公式/冻结筛选/列宽行高/隐藏/样式/条件格式/数据验证/超链接/命名区域/多Sheet/打印/视图/SST洪泛/乱序覆盖） |
| ② 语义校验 | `validate_features.py` | openpyxl 打开逐项断言（68 项：值、样式、合并范围、日期序列号、公式、冻结、下拉框、条件格式、命名区域、打印…） |
| ③ 真机校验 | `validate_features.py`（同上第二层） | **Excel COM 批量静默打开**——严格 schema 校验，任何损坏立即暴露（Windows + 已装 Excel） |
| 回归 | `stress_test.c` + `validate.py` | 1500 单元格乱序+覆盖写；行/列引用升序逐节点断言 |
| ④ 读改写 | `test_rw.c` | 8 用例：载入校验/零损失另存/多表/新增表/自定义样式追加/条件格式追加/验证追加/**打开 openpyxl 生成的 Deflate 压缩文件** |
| ④ 验证 | `validate_rw.py` + `com_rw.py` | openpyxl 8 项内容断言 + Excel COM 10 文件实测 |

```bash
# 跑完整测试(在 tests 目录下)
gcc -std=gnu99 -I.. test_features.c ../src/ark_excel.c ../src/ark_zip.c -o test_features.exe
./test_features.exe                                     # 生成 out/*.xlsx (19 个)
python validate_features.py                             # openpyxl 断言 + Excel COM 实测
python validate_features.py --no-com                    # 无 Excel 环境时跳过 COM 层

# 读改写测试(先用 openpyxl 生成第三方压缩模板)
python -c "from openpyxl import Workbook; wb=Workbook(); wb.active['A1']='tpl'; wb.save('tests/out/rw_tpl_openpyxl.xlsx')"
gcc -std=gnu99 -I.. -I../src test_rw.c ../src/ark_excel.c ../src/ark_zip.c \
    ../src/ark_reader.c ../src/ark_inflate.c -o test_rw.exe
./test_rw.exe && python validate_rw.py && python com_rw.py

gcc -std=gnu99 -I.. -I../port stress_test.c ../src/ark_excel.c ../src/ark_zip.c -o stress.exe
./stress.exe && python validate.py ../stress.xlsx
```

> **重要经验**：openpyxl 能打开 ≠ Excel 能打开。openpyxl 容错性强，会放过元素顺序/sqref 位置等 schema 违规，而 Excel 严格校验直接判损坏。因此本库以 Excel COM 实测作为最终验收标准。遇到“文件打不开”时，用单特性二分法（每个特性独立成文件让 Excel 逐个打开）可快速定位损坏源。详见设计文档第 99 章。

## API 概览

| 类别 | 函数 |
|---|---|
| 生命周期 | `ark_wb_init` `ark_wb_add_sheet` `ark_wb_save` `ark_wb_reset` |
| 读改写 | `ark_wb_open`（打开现有 xlsx）`ark_style_map`（内置样式 id 重映射）`ark_get_cell`（查单元格） |
| 单元格 | `ark_set_str/num/int/bool/formula/error/blank/date/datetime` |
| 合并/行列 | `ark_merge` `ark_set_col_width/hidden` `ark_set_row_height/hidden` |
| 视图 | `ark_freeze` `ark_autofilter` `ark_tab_color` `ark_sheet_rtl` `ark_sheet_gridlines` |
| 交互 | `ark_hyperlink` `ark_add_validation` `ark_add_condfmt` `ark_named_range` |
| 打印 | `ark_page_setup` `ark_print_area` `ark_header_footer` `ark_row_break` |
| 样式 | `ark_wb_add_style`（配合 `ARK_STYLE_DEF_INIT` 清零初始化）+ 19 个预定义样式 |

注意：`ark_style_def_t` 必须以 `= ARK_STYLE_DEF_INIT` 清零后再填字段。

## 已知限制

- 写出为 Store 模式（不压缩）：文件体积约为压缩版的 3~6 倍，换 CPU/内存为零压缩成本；
  如需压缩可在输出层外接 miniz（见设计文档第 46 章）。
  **读取方向无此限制**：`ark_wb_open` 可正常打开任何 Deflate 压缩的 xlsx（自带解压器）
- 不支持：图表、图片、批注(VML)、Shape、Sparklines、加密（设计文档中标注嵌入式性价比低的特性）
- 每条字符串 ≤ `ARK_MAX_STR_LEN` 字节；行号 ≤ 65535（合并/筛选区域）
- 读改写：单个 XML 部件解压后需 ≤ `ARK_OPEN_XML_BUF`（默认 16KB）；载入的是数据层
  （单元格/合并/行列/冻结/筛选），模板中的图表/批注等**不解析但原样保留**；
  `ark_wb_open` 后读取源需保持可读至 `ark_wb_save` 完成（增量保存仍需读模板字节）
