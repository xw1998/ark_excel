# ark_excel 移植指南（小白友好版）

> 本指南面向第一次接触 ark_excel 的开发者，手把手教你把库跑在 **RT-Thread**、**svcrtos** 和 **裸机（无 OS）** 三种环境上。读完本文，你应当能：
>
> 1. 知道 ark_excel 需要哪几个文件、依赖什么、不依赖什么；
> 2. 写出适合你平台的"输出回调"（写文件 / UART / 内存缓冲）；
> 3. （可选）写出"读取回调"，实现打开并修改现有 xlsx；
> 4. 按设备资源裁剪功能，省下宝贵的 Flash 和 RAM；
> 5. 编译、运行第一个例程，用 Excel 打开你生成的文件。
>
> 全程不需要动态内存（malloc），不需要操作系统，不需要文件系统——**库的核心只有一个回调函数**。

---

## 目录

- [0. 开始之前：ark_excel 是如何工作的](#0-开始之前ark-excel-是如何工作的)
- [1. 通用移植步骤（三种平台都要做）](#1-通用移植步骤三种平台都要做)
  - [1.1 拷贝文件](#11-拷贝文件)
  - [1.2 理解回调：库与外部世界的唯一接口](#12-理解回调库与外部世界的唯一接口)
  - [1.3 用配置头裁剪功能](#13-用配置头裁剪功能)
  - [1.4 内存预算：ark_wb_t 到底多大](#14-内存预算ark_wb_t-到底多大)
- [2. RT-Thread 移植](#2-rt-thread-移植)
- [3. svcrtos 移植](#3-svcrtos-移植)
- [4. 裸机（无 OS）移植](#4-裸机无-os-移植)
  - [4.1 有文件系统（SD 卡 / SPI Flash + FatFS）](#41-有文件系统sd-卡--spi-flash--fatfs)
  - [4.2 无文件系统：直接从 UART 发出](#42-无文件系统直接从-uart-发出)
  - [4.3 无文件系统：先存内存再一次性发送](#43-无文件系统先存内存再一次性发送)
- [5. 读改写（打开并修改现有 xlsx）的移植要点](#5-读改写打开并修改现有-xlsx的移植要点)
- [6. 常见编译错误与解决](#6-常见编译错误与解决)
- [7. 验收清单](#7-验收清单)

---

## 0. 开始之前：ark_excel 是如何工作的

先建立一个直观印象。生成一个 xlsx 文件的本质是：**按 ZIP 格式把若干 XML 文件打包成一个字节流**。ark_excel 内部做了三件事：

1. 你调用 `ark_set_str` / `ark_set_num` 等 API 时，数据被记进 `ark_wb_t` 这个大结构体（静态内存，不用 malloc）；
2. 调用 `ark_wb_save()` 时，库用"两遍法"流式生成 ZIP：先算一遍长度和 CRC，再输出一遍字节；
3. **所有输出的字节，都通过你注册的一个函数送出去**——这个函数写到哪，xlsx 就"存"到哪。

```c
/* 你唯一必须实现的函数原型 */
int your_write(const void *data, size_t len, void *ctx);
```

它可以写文件、写 UART、写网络 socket、写内存缓冲——库完全不关心。这就是"极其便于移植"的原因：**库本身是纯 C99 计算 + 一个回调，零平台依赖**。

| 依赖项 | 是否需要 | 说明 |
|---|---|---|
| 动态内存 malloc/free | 否 | 全静态池 |
| 操作系统 / 线程 | 否 | 库内无线程，无全局锁 |
| 文件系统 | 否 | 只在你需要时经回调使用 |
| C 标准库 | 仅 `<stdint.h>` `<stddef.h>` `<string.h>`（如 strlen/memcpy） | `printf`/`double` 输出均为手写实现，可无 libc printf |
| 编译器 | C99 及以上 | gcc / armcc(AC5 需 --c99) / armclang / IAR / MSVC 均可 |

---

## 1. 通用移植步骤（三种平台都要做）

### 1.1 拷贝文件

最小只需要 **2 个文件**；全功能是 **8 个文件**：

| 文件 | 作用 | 必须？ |
|---|---|---|
| `ark_excel.h` | 公共 API 头 | 必须 |
| `src/ark_excel.c` | 核心：数据模型 + XML 生成 | 必须 |
| `src/ark_zip.c/.h` | ZIP 容器（两遍法流式） | 必须 |
| `src/ark_reader.c/.h` | 读改写：解析现有 xlsx | 用"打开/修改/保存"才需要 |
| `src/ark_inflate.c/.h` | DEFLATE 解压（reader 依赖） | 用读改写才需要 |
| `ark_excel_config.h` | 容量与功能裁剪 | 必须（默认值即可编译） |
| `src/ark_excel_internal.h` | 内部接口（只需和 .c 放一起） | 随源码带上 |

操作：把整个 `src/` 目录、`ark_excel.h`、`ark_excel_config.h` 拷进你的工程，加入编译即可。**不需要修改库的任何一行代码**。

### 1.2 理解回调：库与外部世界的唯一接口

**写回调**（生成 xlsx 时必用）：

```c
/* 原型见 ark_excel.h: typedef int (*ark_write_fn)(const void*, size_t, void*); */
/* 返回 0 表示成功，非 0 表示失败（保存立即中止并返回 ARK_ERR_WRITE） */

int my_write(const void *data, size_t len, void *ctx)
{
    /* 把 data 指向的 len 个字节送往目的地，然后 return 0 */
}
```

注册方式：

```c
ark_wb_t wb;                              /* 建议静态分配或由你管理生命周期 */
ark_wb_init(&wb, my_write, my_ctx);       /* my_ctx 会被原样传回给回调 */
```

最常用的现成实现在 `port/port_examples.c` 里，涵盖 stdio / 内存缓冲 / UART / FatFS / svcrtos 五种，可直接拷贝修改。

**读回调**（仅"打开现有 xlsx 修改"时需要，见第 5 章）：

```c
uint32_t my_read(void *ctx, uint8_t *buf, uint32_t len); /* 返回实际读取字节数 */
int      my_seek(void *ctx, uint32_t pos);               /* 定位到绝对偏移，0 成功 */
```

### 1.3 用配置头裁剪功能

打开 `ark_excel_config.h`，两类参数：

- **容量类**：`ARK_MAX_CELLS`、`ARK_MAX_SHEETS`、`ARK_MAX_STR` 等。容量↑内存↑，按你的表格规模设。
- **功能类**（置 0 即从编译中剔除该功能，对应 API 返回 `ARK_ERR_BAD_STATE(17)`）：

| 开关 | 关掉后省的 Flash | 适合场景 |
|---|---|---|
| `ARK_ENABLE_READWRITE` | ~15KB（连 reader+inflate 一起剔除） | 只生成、不读取的设备 |
| `ARK_ENABLE_INFLATE` | ~2KB（只解压压缩过的 xlsx 时用） | 只读自家生成的 Store 文件 |
| `ARK_ENABLE_FORMULA` | 少量 | 无公式需求 |
| `ARK_ENABLE_CONDFMT` | 少量 | 无条件格式需求 |
| `ARK_ENABLE_VALIDATION` | 少量 | 无数据验证需求 |
| `ARK_ENABLE_HYPERLINK` | 少量 | 无超链接需求 |

实测体积（gcc -Os，.text）：全功能 48.8KB；只写（READWRITE=0）33.1KB；全关 31.6KB。

**小白常见误区**：把开关设为 0 后编译报"未定义的 `ark_wb_open`"是正常的——头文件里该 API 已被条件编译包裹，恢复 `1` 即可。

### 1.4 内存预算：ark_wb_t 到底多大

`ark_wb_t` 是一个纯静态数组结构体，大小完全由 config 决定。默认配置下典型值约 **十几 KB**（含单元格池、字符串 arena、样式表）。

估算方法（粗略）：

```
sizeof(ark_wb_t) ≈ ARK_MAX_CELLS × sizeof(ark_cell_t)
                + ARK_MAX_STR × sizeof(ark_str_t) + arena 总量
                + ARK_MAX_SHEETS × sizeof(ark_sheet_t)
                + 合并/列宽/行高/验证/条件格式等小数组
```

想知道准确值，在 PC 上打印一次即可：

```c
printf("sizeof(ark_wb_t) = %zu\n", sizeof(ark_wb_t));
```

> 提示：`ark_wb_t` 很大，**不要**放在栈上（`void task(){ ark_wb_t wb; }` 会爆栈），用 `static ark_wb_t wb;` 或从你的内存池分配。

---

## 2. RT-Thread 移植

RT-Thread（RTT）是国内最流行的 RTOS，包管理器丰富。ark_excel 不需要任何 RTT 组件，但配合 DFS 文件系统最顺手。

### 步骤 1：加入源码

任选一种：

- **方式 A（推荐，手工加入）**：把 `src/`、`ark_excel.h`、`ark_excel_config.h` 拷到工程，例如 `applications/ark_excel/` 下，在 RT-Studio / Keil / CMake 里把两个（或四个）`.c` 加入编译，头文件路径加上该目录。
- **方式 B（SCons）**：如果你的工程用 SCons，在 `SConscript` 里加：

```python
from building import *

cwd = GetCurrentDir()
src = Glob('applications/ark_excel/src/*.c')
path = [cwd + '/applications/ark_excel', cwd + '/applications/ark_excel/src']

group = DefineGroup('ArkExcel', src, depend = [''], CPPPATH = path)
Return('group')
```

> 注意：`src/` 下 `ark_reader.c` 与 `ark_inflate.c` 只有开 `ARK_ENABLE_READWRITE` 才有内容，平时一起加入编译也无害（内置条件编译）。

### 步骤 2：配置裁剪

拷贝一份 `ark_excel_config.h`，小内存设备（如 STM32F103，20KB RAM）建议：

```c
#define ARK_MAX_CELLS     200
#define ARK_MAX_SHEETS     2
#define ARK_MAX_STR       150
#define ARK_STR_ARENA   4096      /* 字符串总池(字节) */
#define ARK_ENABLE_READWRITE   0  /* 只写不读，省 15KB Flash */
```

### 步骤 3：实现写回调（经 DFS 写 SD 卡）

RTT 开启 DFS + FatFS 后，直接用 POSIX 文件接口：

```c
#include <stdio.h>
#include "ark_excel.h"

static int sd_write(const void *data, size_t len, void *ctx)
{
    FILE *f = (FILE *)ctx;
    return (fwrite(data, 1, len, f) == len) ? 0 : -1;
}

/* 在线程里（注意栈别太小，建议 >= 2KB，ark_wb_t 本身放静态区） */
static ark_wb_t wb;
void excel_demo_thread(void *param)
{
    (void)param;
    FILE *f = fopen("/sdcard/report.xlsx", "wb");   /* DFS 挂载路径 */
    if (!f) return;

    ark_wb_init(&wb, sd_write, f);
    ark_wb_add_sheet(&wb, "数据");
    ark_set_str(&wb, 0, 1, 1, "温度(℃)", ARK_STYLE_TABLE_HEADER);
    ark_set_num(&wb, 0, 2, 1, 25.5, ARK_STYLE_NUM_2DP);
    ark_wb_save(&wb);          /* 所有字节经 sd_write 进入文件 */

    fclose(f);
    /* 拔卡 → 电脑 Excel 打开验证 */
}
```

### 步骤 4：常见 RTT 专属问题

- **`fopen` 返回 NULL**：DFS 未挂载，先 `df_mount()`（或开 PKG_FAL/自动挂载），用 `list_device` 确认块设备存在。
- **写一半断电文件损坏**：xlsx 是 ZIP，必须完整写入。建议保存成功后再 `fclose`（库内顺序输出，只要回调都返回 0，文件即完整）。
- **FinSH 打印中文乱码**：与本库无关，检查终端编码为 UTF-8；库生成的 XML 均为 UTF-8。

---

## 3. svcrtos 移植

svcrtos（https://gitee.com/xw19981010/svcrtos_new.git）是自研 RTOS，接口风格接近 POSIX/-lite。本库已在 `port/port_examples.c` 提供 svcrtos 参考回调。

### 步骤 1：加入源码

把 `src/`、`ark_excel.h`、`ark_excel_config.h` 拷入 svcrtos 工程的用户源码目录（如 `user/` 或 `app/`），加入 Makefile/CMake 源列表。svcrtos 工具链（gcc/arm-none-eabi）为 C99，无需额外编译参数。

### 步骤 2：实现写回调

svcrtos 若带 VFS，用其文件 API 包装；否则走设备驱动（如 UART/SPI Flash）：

```c
/* 假设 svcrtos 提供 svc_fopen/svc_fwrite 风格接口（按你的实际 API 改名） */
static int svc_write(const void *data, size_t len, void *ctx)
{
    svc_file_t *fp = (svc_file_t *)ctx;
    return (svc_fwrite(fp, data, len) == len) ? 0 : -1;
}

/* UART 直接输出（无文件系统时） */
extern int svc_uart_send(const void *d, unsigned n);   /* BSP 提供 */
static int uart_write(const void *data, size_t len, void *ctx)
{
    (void)ctx;
    return svc_uart_send(data, (unsigned)len) == (int)len ? 0 : -1;
}
```

### 步骤 3：任务里使用

```c
static ark_wb_t wb;    /* 静态分配，避免占任务栈 */

void app_excel_task(void)
{
    ark_wb_init(&wb, uart_write, NULL);
    ark_wb_add_sheet(&wb, "log");
    ark_set_str(&wb, 0, 1, 1, "hello svcrtos", 0);
    ark_wb_save(&wb);          /* PC 端用串口工具收流存为 .xlsx 即可打开 */
    ark_wb_reset(&wb);         /* 复用结构体生成下一个文件 */
}
```

### 步骤 4：svcrtos 专属建议

- **接收端技巧**：UART 场景下，PC 用 `python -m serial.tools.miniterm` 或自写脚本把收到的原始字节存成 `.xlsx`，Excel 即可打开。注意关闭串口工具的"换行转换"选项，保证字节流原样落盘。
- **SPI Flash 落盘**：若 Flash 上有简易文件层，用你的驱动实现 `write` 回调即可；若没有，可先写满一片 Flash 区域，再用烧录器读出存为文件。

---

## 4. 裸机（无 OS）移植

裸机 = 无线程、无文件系统也可能没有。ark_excel 天然适配：整个保存过程是**同步、顺序、可重入于单个调用栈**的，在 main 的 while(1) 里调用即可。

### 步骤 1：编译

任何支持 C99 的编译器。Keil AC5 记得勾选 `--c99`；IAR 选 C99 方言；gcc 直接编译：

```
gcc -std=gnu99 -Os -I. -Isrc -c src/ark_excel.c -o ark_excel.o
gcc -std=gnu99 -Os -I. -Isrc -c src/ark_zip.c   -o ark_zip.o
```

### 步骤 2：选择输出方式

裸机三种典型接法，按硬件条件选：

#### 4.1 有文件系统（SD 卡 / SPI Flash + FatFS）

```c
#include "ff.h"

static FIL g_fil;
static int fatfs_write(const void *data, size_t len, void *ctx)
{
    UINT bw;
    if (f_write((FIL *)ctx, data, len, &bw) != FR_OK || bw != len) return -1;
    return 0;
}

void make_report(void)
{
    if (f_open(&g_fil, "0:report.xlsx", FA_WRITE | FA_CREATE_ALWAYS) != FR_OK)
        return;
    ark_wb_init(&wb, fatfs_write, &g_fil);
    /* ...填数据... */
    ark_wb_save(&wb);
    f_close(&g_fil);
}
```

> FatFS 注意：`_FS_MINIMIZE` 保持默认 0（需要 `f_lseek` 支持时确认 `_FS_MINIMIZE <= 1`；本库写路径只用 f_write，不 seek）。`_USE_LFN` 与本库无关（文件名是你自己起的）。

#### 4.2 无文件系统：直接从 UART 发出

保存期间约需调用你的 UART 发送函数上百次，每次几十~几百字节。阻塞发送最简单：

```c
static int uart_write(const void *data, size_t len, void *ctx)
{
    (void)ctx;
    const uint8_t *p = (const uint8_t *)data;
    while (len--) {
        while (!(UART->SR & UART_SR_TXE));   /* 等发送寄存器空，按芯片改 */
        UART->DR = *p++;
    }
    return 0;
}
```

115200 波特率下约 50KB 的 xlsx 需要 ~5 秒，期间 CPU 忙等。若不能接受，用 4.3 的内存方案 + DMA 事后发送。

#### 4.3 无文件系统：先存内存再一次性发送

```c
static uint8_t g_xlsx_buf[ARK_XLSX_MAX_SIZE];   /* 按你的表格规模定，如 32KB */
static ark_membuf_t g_mb = { g_xlsx_buf, 0, sizeof(g_xlsx_buf) };

/* 回调用 port_examples.c 里的 ark_write_membuf */
ark_wb_init(&wb, ark_write_membuf, &g_mb);
ark_wb_save(&wb);
/* 此后任何时机：uart_dma_send(g_xlsx_buf, g_mb.len); */

/* 也可以就地在 PC 上查看：把 g_xlsx_buf 通过调试器导出为 .xlsx 文件 */
```

### 步骤 3：裸机专属注意

- **中断安全**：库不是线程/中断安全的。裸机单循环顺序调用没有问题；若在中断里也要写数据，用主循环标志位代替直接调用。
- **看门狗**：大表保存（尤其 UART 忙等）可能超时喂狗，在 write 回调里顺手 `IWDG_Refresh()`。
- **`sizeof(ark_wb_t)` 很大**：务必全局/静态分配，绝不放栈上。

---

## 5. 读改写（打开并修改现有 xlsx）的移植要点

读改写需要实现 `read + seek` 两个回调（`src/ark_reader.h`）：

```c
typedef uint32_t (*ark_read_fn)(void *ctx, uint8_t *buf, uint32_t len);
typedef int      (*ark_seek_fn)(void *ctx, uint32_t pos);

typedef struct {
    ark_read_fn  read;
    ark_seek_fn  seek;
    void        *ctx;
} ark_reader_io_t;
```

FatFS 版示例：

```c
static uint32_t ff_read(void *ctx, uint8_t *buf, uint32_t len)
{
    UINT br;
    f_read((FIL *)ctx, buf, len, &br);
    return br;                       /* 返回实际读到的字节数 */
}
static int ff_seek(void *ctx, uint32_t pos)
{
    return f_lseek((FIL *)ctx, pos) == FR_OK ? 0 : -1;
}

static ark_wb_t    wb;
static ark_reader_t rd;              /* 读取器，也要静态分配，数 KB */

void modify_existing(void)
{
    FIL f;
    ark_reader_io_t io = { ff_read, ff_seek, &f };

    f_open(&f, "0:tpl.xlsx", FA_READ);
    if (ark_wb_open(&wb, &rd, &io) == ARK_OK) {
        ark_set_num(&wb, 0, 5, 2, 3.14, 0);        /* 改 B5 */
        ark_wb_save(&wb);                          /* 增量保存：未改部分原样拷贝 */
    }
    f_close(&f);
}
```

移植要点（小白最容易踩的三个坑）：

1. **seek 必须支持绝对定位**，从文件头数起；只支持顺序读的设备（流式 UART）不能当读取源。
2. **读取源必须保持打开直到 `ark_wb_save` 返回**：增量保存时会回头再拷贝未修改的 ZIP 条目，提前关闭会得到 `ARK_ERR_READ`（16 是 WRITE，读取相关错误码见 ark_excel.h）。
3. **裁剪联动**：`ARK_ENABLE_READWRITE=0` 时 `ark_wb_open` 不存在；reader 依赖 `ARK_ENABLE_INFLATE` 才能读 Deflate 压缩的第三方文件（Excel/openpyxl 生成的默认就是 Deflate）。

---

## 6. 常见编译错误与解决

| 现象 | 原因 | 解决 |
|---|---|---|
| `undefined reference to 'ark_wb_open'` | `ARK_ENABLE_READWRITE=0` 或没编入 ark_reader.c/ark_inflate.c | 开开关或补源文件 |
| 调用被裁剪 API 返回 17 | 功能开关置 0，API 保留但显式报错 | 打开对应 `ARK_ENABLE_*` |
| 链接体积超 Flash | 全功能 48.8KB | 参见 1.3 裁剪；最低 31.6KB |
| HardFault / 栈溢出 | `ark_wb_t` 放在了栈上 | 改 `static` 全局 |
| Excel 提示"内容有问题" | 回调中途返回非 0，文件截断 | 检查 write 回调错误路径与介质是否写满 |
| 中文乱码 | 你的字符串不是 UTF-8 | 库按 UTF-8 原样写入，源头转好编码 |
| Keil 报 C99 语法错 | AC5 默认 C90 | Options → C/C++ → 勾选 C99 Mode |
| FatFS 写失败 | 卷满或 `_FS_READONLY` | 检查 `_FS_READONLY == 0` 与剩余空间 |

## 7. 验收清单

移植完成后逐项打勾：

- [ ] `sizeof(ark_wb_t)` 已打印并确认放得下（静态区）
- [ ] 最小程序：init → add_sheet → set_str → save → 回调成功返回 0
- [ ] 生成文件拷到 PC，Excel 双击能打开、内容正确
- [ ] （裁剪场景）被裁 API 返回 17 而不是链接错误
- [ ] （读改写）模板文件能 open，改一格后 save，Excel 打开无修复提示
- [ ] （UART）串口收流存文件，Excel 可打开
- [ ] 断电/复位后重跑一遍，确认无残留状态影响（每次 init 或 reset）

遇到问题优先对照 `port/port_examples.c` 与 `examples/` 目录的完整可编译示例，再查阅 `docs/embedded_excel_guide.md` 第 93–104 章的设计细节。
