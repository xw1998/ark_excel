# 嵌入式系统中实现 Excel (.xlsx) 生成库 —— 从原理到实现的完整指南

> 本文档面向有 C/C++ 基础、但可能对 Excel 文件格式完全不了解的开发者。从"一个 .xlsx 文件到底是什么"讲起，逐步深入到合并单元格、交错行格式等复杂表格的实现，最终给出嵌入式环境下的完整代码方案。

---

## 目录

- [1. 你以为的 Excel 文件 vs 真实的 Excel 文件](#1-你以为的-excel-文件-vs-真实的-excel-文件)
- [2. ZIP 容器：xlsx 的最外层结构](#2-zip-容器xlsx-的最外层结构)
- [3. 核心 XML 文件详解](#3-核心-xml-文件详解)
- [4. 单元格、行、列：表格的基本积木](#4-单元格行列表格的基本积木)
- [5. 合并单元格：两行变一行的核心机制](#5-合并单元格两行变一行的核心机制)
- [6. 交错行格式：让每一行"长得不一样"](#6-交错行格式让每一行长得不一样)
- [7. 样式系统：颜色、边框、字体、对齐](#7-样式系统颜色边框字体对齐)
- [8. 共享字符串表：省内存的关键技巧](#8-共享字符串表省内存的关键技巧)
- [9. 嵌入式实现策略：在资源受限环境下的工程取舍](#9-嵌入式实现策略在资源受限环境下的工程取舍)
- [10. 最小 ZIP 写入器：不压缩也能生成合法 xlsx](#10-最小-zip-写入器不压缩也能生成合法-xlsx)
- [11. 完整代码方案：一个可用的嵌入式 Excel 生成器](#11-完整代码方案一个可用的嵌入式-excel-生成器)
- [12. 高级：各种"奇葩"格式的实现配方](#12-高级各种奇葩格式的实现配方)
- [13. 常见陷阱与排错指南](#13-常见陷阱与排错指南)
- [14. 性能优化与内存预估](#14-性能优化与内存预估)
- [附录 A：OOXML 速查表](#附录-aooxml-速查表)
- [附录 B：单元格地址编码算法](#附录-b单元格地址编码算法)

---

## 1. 你以为的 Excel 文件 vs 真实的 Excel 文件

### 1.1 直觉认知（错误）

```
❌ "Excel 文件就是一种特殊的二进制格式，有固定偏移的表头..."
```

### 1.2 真相

**一个 .xlsx 文件就是一个 ZIP 压缩包，里面装着一堆 XML 文件和一些辅助文件。**

没错，就这么简单。你可以用任何解压工具（7-Zip、WinRAR、甚至 `unzip`）把一个 .xlsx 文件解压，会看到这样的目录结构：

```
Book1.xlsx（解压后）
├── [Content_Types].xml          ← 告诉打开它的程序："我是 Excel 文件"
├── _rels/
│   └── .rels                    ← 关系文件：说明根文件是谁
├── docProps/
│   ├── app.xml                  ← 应用程序属性（作者、公司等）
│   └── core.xml                 ← 核心属性（创建时间、修改时间等）
└── xl/
    ├── workbook.xml             ← 工作簿定义：有哪些 Sheet
    ├── _rels/
    │   └── workbook.xml.rels    ← 工作簿的关系：Sheet 对应哪个文件
    ├── sharedStrings.xml        ← 共享字符串表（所有文本都存这里）
    ├── styles.xml               ← 样式表（颜色、字体、边框定义）
    └── worksheets/
        ├── sheet1.xml           ← 第一个工作表的实际数据
        ├── sheet2.xml           ← 第二个工作表（如果有）
        └── _rels/
            └── sheet1.xml.rels   ← Sheet 的关系（超链接等）
```

### 1.3 这对嵌入式开发者意味着什么

| 传统认知 | 实际情况 |
|---------|---------|
| Excel 是复杂的二进制格式 | Excel 是 ZIP + XML，纯文本可读 |
| 需要专门的库才能生成 | 理论上手写 XML 字符串 + 打个 ZIP 就行 |
| 格式细节被微软保密 | OOXML 是国际标准（ISO/IEC 29500），公开可查 |
| 嵌入式上不可能实现 | 完全可能，核心逻辑不到 1000 行 C 代码 |

> **关键理解**：生成 .xlsx 的本质就是——**拼 XML 字符串 → 打成 ZIP 包 → 改后缀为 .xlsx**。理解了这一点，后面的所有复杂格式都不再神秘。

---

## 2. ZIP 容器：xlsx 的最外层结构

### 2.1 为什么是 ZIP

微软在 Office 2007 中引入了 OOXML（Office Open XML）格式，用 ZIP + XML 替代了老的二进制格式（.xls）。原因很简单：

- **压缩节省空间**：XML 是纯文本，压缩率高
- **模块化**：不同部分（内容、样式、元数据）分文件存储，互不干扰
- **可恢复性**：文件损坏时，至少能手动解压取出部分数据

### 2.2 ZIP 文件格式精解

一个 ZIP 文件由三部分组成：

```
┌──────────────────────────────────────────────┐
│  文件条目 1 (Local File Header + 数据)        │
├──────────────────────────────────────────────┤
│  文件条目 2 (Local File Header + 数据)        │
├──────────────────────────────────────────────┤
│  ...                                          │
├──────────────────────────────────────────────┤
│  文件条目 N (Local File Header + 数据)        │
├──────────────────────────────────────────────┤
│  中央目录 (Central Directory)                 │
│  - 包含所有文件的索引信息                       │
├──────────────────────────────────────────────┤
│  中央目录结束记录 (End of Central Directory)  │
│  - 指向中央目录的起始位置                       │
└──────────────────────────────────────────────┘
```

### 2.3 Local File Header 结构（每个文件条目的开头）

这是 ZIP 中每个文件条目的头部，**固定 30 字节**（文件名长度另算）：

```
偏移   长度   含义                  嵌入式实现时怎么填
───────────────────────────────────────────────────────────
0      4     签名 (固定)            0x04034b50 (小端: PK\x03\x04)
4      2     版本所需               填 20 (2.0)，表示最基本
6      2     通用标志位             填 0（无加密、无特殊标志）
8      2     压缩方式               0 = 不压缩(Stored)，8 = Deflate
10     2     最后修改时间            编码后的时间（见下文）
12     2     最后修改日期            编码后的日期（见下文）
14     4     CRC32 校验             数据的 CRC32 值（必须算对！）
18     4     压缩后大小             不压缩时 = 原始大小
22     4     未压缩大小             = 原始数据长度
26     2     文件名长度              文件名的字节数
28     2     额外字段长度            填 0（我们不需要）
30     变长  文件名                  如 "xl/worksheets/sheet1.xml"
       变长  额外字段                无（长度为0）
       变长  文件数据                XML 内容本体
```

### 2.4 Central Directory 结构（中央目录）

每个文件在中央目录中也有一条记录，**固定 46 字节**（文件名等另算）：

```
偏移   长度   含义
───────────────────────────────────────────────────────────
0      4     签名 (固定)            0x02014b50 (PK\x01\x02)
4      2     版本制作               填 20
6      2     版本所需               填 20
8      2     通用标志位             0
10     2     压缩方法               0 或 8
12     2     最后修改时间            同 Local Header
14     2     最后修改日期            同 Local Header
16     4     CRC32                  同 Local Header
20     4     压缩后大小              同 Local Header
24     4     未压缩大小              同 Local Header
28     2     文件名长度              同 Local Header
30     2     额外字段长度            0
32     2     文件注释长度            0
34     2     磁盘编号开始            0
36     2     内部文件属性            0
38     4     外部文件属性            0
42     4     Local Header 偏移       ← 关键！指向该文件在ZIP中的字节位置
46     变长  文件名                  同 Local Header
```

### 2.5 End of Central Directory Record（中央目录结束记录）

放在 ZIP 文件最末尾，**固定 22 字节**（注释另算）：

```
偏移   长度   含义
───────────────────────────────────────────────────────────
0      4     签名 (固定)            0x06054b50 (PK\x05\x06)
4      2     磁盘编号               0
6      2     中央目录起始磁盘        0
8      2     本磁盘上的条目数        文件总数
10     2     总条目数               文件总数
12     4     中央目录大小            中央目录总字节数
16     4     中央目录起始偏移        中央目录在 ZIP 中的位置
20     2     注释长度               0
```

### 2.6 时间和日期的编码

ZIP 用一种特殊的 MS-DOS 格式编码时间日期，每个 2 字节：

```
时间 (2 bytes):
  位 15-11: 小时 (0-23)
  位 10-5:  分钟 (0-59)
  位 4-0:   秒/2  (0-29，实际秒 = 值 × 2)

日期 (2 bytes):
  位 15-9:  年 - 1980 (0-127，即 1980-2107)
  位 8-5:   月 (1-12)
  位 4-0:   日 (1-31)

示例：2024年3月15日 14:30:00
  日期 = (2024-1980)<<9 | 3<<5 | 15 = 44<<9 | 96 | 15 = 22592 + 96 + 15 = 22703 → 0x58AF
  时间 = 14<<11 | 30<<5 | 0 = 28672 + 960 + 0 = 29632 → 0x7380
```

### 2.7 CRC32 校验

这是 ZIP 格式中**唯一可能卡住你的数学运算**。CRC32 是一种校验和算法，用于检测数据是否被损坏。Excel 在打开 xlsx 时会校验 CRC32，如果不对会报"文件已损坏"。

CRC32 的原理：
- 用一个固定的 256 项查找表（预计算）
- 对数据逐字节异或运算
- 最终得到一个 4 字节的校验值

> **嵌入式实现提示**：CRC32 查找表占 1024 字节 ROM，如果你系统里已有 CRC32 实现（很多 MCU 的硬件 CRC 模块也可用），直接复用即可。查找表是固定的，放在 `const` 数组里不占 RAM。

---

## 3. 核心 XML 文件详解

### 3.1 `[Content_Types].xml` —— 身份证明

这个文件位于 ZIP 根目录，告诉打开程序"我是什么类型的文件"。**没有它，Excel 直接拒绝打开。**

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <!-- 默认类型：按扩展名识别 -->
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  
  <!-- 覆盖类型：按文件路径精确指定 -->
  <Override PartName="/xl/workbook.xml" 
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" 
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" 
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" 
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>
```

**为什么有两种类型声明？**
- `<Default>` 按文件扩展名匹配，如所有 `.rels` 文件都是关系文件
- `<Override>` 按完整路径精确匹配，因为同样叫 `.xml`，workbook 和 worksheet 的 ContentType 不一样

### 3.2 `_rels/.rels` —— 根关系文件

关系文件（.rels）的作用是建立"谁引用谁"的关系。根关系文件告诉程序：这个包的入口文件是哪个。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" 
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" 
                Target="xl/workbook.xml"/>
  <Relationship Id="rId2" 
                Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" 
                Target="docProps/core.xml"/>
  <Relationship Id="rId3" 
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" 
                Target="docProps/app.xml"/>
</Relationships>
```

> **通俗理解**：关系文件就像书的目录页——"第1章在哪个页码"一样，`rId1` 对应 `xl/workbook.xml`，是整个 Excel 文件的真正入口。

### 3.3 `xl/workbook.xml` —— 工作簿

定义有哪些工作表（Sheet），以及它们的名字和顺序。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="数据报表" sheetId="1" r:id="rId1"/>
    <sheet name="汇总" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
```

- `name`：显示在 Excel 底部 Tab 上的名字
- `sheetId`：内部 ID（随便填，不重复即可）
- `r:id`：指向 `workbook.xml.rels` 中定义的关系 ID

### 3.4 `xl/_rels/workbook.xml.rels` —— 工作簿关系

告诉程序：`rId1` 对应哪个 sheet 文件。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" 
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" 
                Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" 
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" 
                Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" 
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" 
                Target="styles.xml"/>
  <Relationship Id="rId4" 
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" 
                Target="sharedStrings.xml"/>
</Relationships>
```

### 3.5 `xl/sharedStrings.xml` —— 共享字符串表

这是**省内存的关键设计**。Excel 不会在每个单元格里重复存储文本，而是把所有文本放在一个"共享字符串表"中，单元格里只存索引号。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
     count="5" uniqueCount="3">
  <si><t>姓名</t></si>
  <si><t>张三</t></si>
  <si><t>工程师</t></si>
</sst>
```

- `count`：所有单元格中文本的总引用次数（包括重复引用）
- `uniqueCount`：不同文本的数量
- `<si>`：一个字符串项 (Shared Item)
- `<t>`：实际的文本内容
- 索引从 0 开始：`姓名` = 索引 0，`张三` = 索引 1，`工程师` = 索引 2

> **为什么这样设计？** 假设有 1000 个单元格都写着"待定"两个字。如果不共享，就要在 XML 里写 1000 遍"待定"。共享后，只在 sharedStrings.xml 里写一次，每个单元格只需写一个数字 `0`。对于嵌入式系统，这意味着 **大幅减少输出数据量和内存占用**。

### 3.6 `xl/styles.xml` —— 样式表

定义所有样式：字体、颜色、边框、对齐方式、数字格式等。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <!-- 数字格式 -->
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="0.00"/>
  </numFmts>
  
  <!-- 字体定义 -->
  <fonts count="2">
    <font>
      <sz val="11"/>
      <name val="等线"/>
    </font>
    <font>
      <b/>                          <!-- 粗体 -->
      <sz val="14"/>
      <color rgb="FFFF0000"/>       <!-- 红色 -->
      <name val="微软雅黑"/>
    </font>
  </fonts>
  
  <!-- 填充色定义 -->
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>                    <!-- 索引0：无填充 -->
    <fill><patternFill patternType="gray125"/></fill>                  <!-- 索引1：灰色125 -->
    <fill><patternFill patternType="solid">
      <fgColor rgb="FFFFC000"/>                                      <!-- 索引2：橙色 -->
    </patternFill></fill>
  </fills>
  
  <!-- 边框定义 -->
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>        <!-- 索引0：无边框 -->
    <border>
      <left style="thin"><color rgb="FF000000"/></left>              <!-- 索引1：细黑边框 -->
      <right style="thin"><color rgb="FF000000"/></right>
      <top style="thin"><color rgb="FF000000"/></top>
      <bottom style="thin"><color rgb="FF000000"/></bottom>
      <diagonal/>
    </border>
  </borders>
  
  <!-- 单元格样式组合 (xf = cell Format) -->
  <cellXfs count="3">
    <!-- 索引0：默认样式 -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <!-- 索引1：粗体红色大字 + 橙色背景 -->
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" 
        applyFont="1" applyFill="1"/>
    <!-- 索引2：细黑边框 + 默认字体 -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" 
        applyBorder="1"/>
  </cellXfs>
</styleSheet>
```

**样式的工作方式像菜单点菜：**
1. `fonts` 里定义了有哪些"字体"可选（索引 0、1、...）
2. `fills` 里定义了有哪些"填充色"可选
3. `borders` 里定义了有哪些"边框"可选
4. `cellXfs` 是"套餐组合"——把字体 + 填充 + 边框 + 数字格式组合在一起
5. 每个单元格引用一个 `cellXfs` 的索引号（`s` 属性），就决定了它的全部外观

> **重要规则**：`fonts` 的前 2 项、`fills` 的前 3 项是保留的，不能随便改！
> - `fonts[0]`：默认字体（必须存在）
> - `fills[0]`：无填充（必须存在）
> - `fills[1]`：灰色125样式（必须存在，Excel 打开时会检查）
> - 你的自定义样式从索引 1（字体）、2（填充）开始

### 3.7 `xl/worksheets/sheet1.xml` —— 工作表数据（最核心的文件）

这里存放实际的表格数据。这是我们要重点操作的文件。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  
  <!-- 列宽定义 -->
  <cols>
    <col min="1" max="1" width="20" customWidth="1"/>   <!-- A列宽20 -->
    <col min="2" max="2" width="15" customWidth="1"/>   <!-- B列宽15 -->
  </cols>
  
  <!-- 行和单元格数据 -->
  <sheetData>
    <row r="1" ht="30" customHeight="1">                <!-- 第1行，行高30 -->
      <c r="A1" t="s" s="1"><v>0</v></c>                <!-- A1：共享字符串索引0，样式1 -->
      <c r="B1" t="s" s="1"><v>1</v></c>                <!-- B1：共享字符串索引1，样式1 -->
    </row>
    <row r="2">                                          <!-- 第2行，默认行高 -->
      <c r="A2" t="s"><v>2</v></c>                      <!-- A2：共享字符串索引2 -->
      <c r="B2"><v>42.5</v></c>                         <!-- B2：数字42.5 -->
    </row>
  </sheetData>
  
  <!-- 合并单元格定义 -->
  <mergeCells count="1">
    <mergeCell ref="A1:B1"/>                            <!-- A1到B1合并 -->
  </mergeCells>
</worksheet>
```

### 3.8 文件之间的关系图

```
[Content_Types].xml
       │ (声明所有文件类型)
       ▼
_rels/.rels
       │ (指向根文件)
       ▼
xl/workbook.xml
       │                    │                    │
       ▼                    ▼                    ▼
xl/_rels/workbook.xml.rels  │                    │
       │ (指向具体文件)      │                    │
       ├───────────────────► │                    │
       │              xl/worksheets/     xl/styles.xml   xl/sharedStrings.xml
       │              sheet1.xml            ▲                    ▲
       │                   │                │                    │
       │                   └────────────────┴────────────────────┘
       │                (单元格通过 s 属性引用样式，通过 t="s" 的 v 值引用字符串)
```

> **一句话总结关系链**：`[Content_Types].xml` 声明身份 → `.rels` 找到入口 → `workbook.xml` 列出 Sheet → `workbook.xml.rels` 指向具体 sheet 文件 → `sheetN.xml` 里存数据 → 数据中的字符串指向 `sharedStrings.xml`，样式指向 `styles.xml`。

---

## 4. 单元格、行、列：表格的基本积木

### 4.1 单元格 (`<c>` 元素)

每个单元格是 XML 中的一个 `<c>` 元素：

```xml
<c r="B3" t="s" s="2"><v>5</v></c>
```

| 属性 | 含义 | 示例值 |
|------|------|--------|
| `r` | 单元格地址（列字母+行号） | `B3` = 第B列第3行 |
| `t` | 数据类型（可选，默认 `n` 数字） | `s`=共享字符串, `n`=数字, `b`=布尔, `str`=公式结果字符串 |
| `s` | 样式索引（指向 styles.xml 中 cellXfs 的索引） | `0`=默认, `1`=样式1, ... |

| 子元素 | 含义 |
|--------|------|
| `<v>` | 值 (Value)。对于 `t="s"` 类型，这里是 sharedStrings 的索引；对于数字，直接是数字 |
| `<f>` | 公式 (Formula)。如果单元格包含公式，写在这里 |
| `<is>` | 内联字符串 (Inline String)。用 `<c t="inlineStr"><is><t>文字</t></is></c>` 直接写字符串 |

### 4.2 数据类型详解

```xml
<!-- 数字 -->
<c r="A1"><v>123</v></c>
<c r="A2"><v>-45.67</v></c>

<!-- 文本（通过共享字符串，最常用） -->
<c r="B1" t="s"><v>0</v></c>        <!-- 0 是 sharedStrings.xml 中的索引 -->

<!-- 文本（内联字符串，不推荐大量使用，但简单省事） -->
<c r="B2" t="inlineStr"><is><t>直接写的文字</t></is></c>

<!-- 布尔值 -->
<c r="C1" t="b"><v>1</v></c>        <!-- 1=true, 0=false -->

<!-- 公式 -->
<c r="D1"><f>SUM(A1:A10)</f><v>550</v></c>  <!-- f 是公式，v 是计算结果 -->
```

> **嵌入式实现建议**：如果不想搞 sharedStrings（省事但浪费空间），可以直接用 `t="inlineStr"` 内联字符串。数据量小的时候完全可以。数据量大时强烈建议用共享字符串表。

### 4.3 行 (`<row>` 元素)

```xml
<row r="3" ht="25.5" customHeight="1" hidden="0" outlineLevel="0">
  <!-- 这一行的所有单元格 -->
</row>
```

| 属性 | 含义 | 默认值 |
|------|------|--------|
| `r` | 行号（从1开始） | 必填 |
| `ht` | 行高（单位：磅，1磅≈0.353mm） | 不填=默认行高 |
| `customHeight` | 是否使用自定义行高 | `0` |
| `hidden` | 是否隐藏行 | `0` |
| `spans` | 跨度范围（如 "1:5" 表示第1-5列有数据） | 可省略 |

### 4.4 列定义 (`<col>` 元素)

列宽在 `<cols>` 中定义，不是每列都必须定义，只定义需要改宽度的列：

```xml
<cols>
  <col min="1" max="1" width="30" customWidth="1"/>    <!-- A列：宽30 -->
  <col min="2" max="3" width="15" customWidth="1"/>    <!-- B-C列：宽15 -->
  <col min="5" max="5" width="8.43"/>                   <!-- E列：宽8.43 -->
</cols>
```

- `min`/`max`：列范围（用数字，1=A, 2=B, ...）
- `width`：列宽（单位约等于字符数）
- `customWidth`：设为 1 表示自定义宽度

### 4.5 单元格地址编码（列字母 ↔ 列数字）

Excel 用字母表示列：A=1, B=2, ..., Z=26, AA=27, AB=28, ..., AZ=52, BA=53, ...

这是一个**26进制**系统，但没有"零"——就像没有"第0列"一样：

```
A → 1      (1 * 26^0)
B → 2
...
Z → 26
AA → 27    (1 * 26^1 + 1 * 26^0)
AB → 28    (1 * 26^1 + 2 * 26^0)
...
AZ → 52    (1 * 26^1 + 26 * 26^0)
BA → 53    (2 * 26^1 + 1 * 26^0)
...
ZZ → 702
AAA → 703
```

**C 语言转换算法**（数字转字母）：

```c
// 列号(1-based) → 列字母，如 27 → "AA"
void col_to_letters(int col, char *buf) {
    int i = 0;
    while (col > 0) {
        col--;                    // 关键：先减1，因为A=1但映射到0
        buf[i++] = 'A' + (col % 26);
        col /= 26;
    }
    buf[i] = '\0';
    // 反转字符串（因为是从低位开始写的）
    for (int j = 0; j < i / 2; j++) {
        char tmp = buf[j];
        buf[j] = buf[i - 1 - j];
        buf[i - 1 - j] = tmp;
    }
}

// 列字母 → 列号，如 "AA" → 27
int letters_to_col(const char *letters) {
    int col = 0;
    while (*letters) {
        col = col * 26 + (*letters - 'A' + 1);
        letters++;
    }
    return col;
}
```

---

## 5. 合并单元格：两行变一行的核心机制

### 5.1 什么是合并

合并单元格就是把多个相邻的格子"粘"成一个大的格子。在 Excel 中你选中多个格子，点"合并居中"，效果如下：

```
合并前：                    合并后：
┌─────┬─────┐              ┌───────────┐
│ A1  │ B1  │              │    A1:B1  │
├─────┼─────┤              ├─────┬─────┤
│ A2  │ B2  │              │ A2  │ B2  │
└─────┴─────┘              └─────┴─────┘
```

### 5.2 XML 中的合并实现

合并通过 `<mergeCells>` 和 `<mergeCell>` 元素实现，**放在 `<sheetData>` 之后**：

```xml
<worksheet xmlns="...">
  <sheetData>
    <!-- 行和单元格数据 -->
    <row r="1">
      <c r="A1" t="s" s="1"><v>0</v></c>    <!-- 只有左上角A1有值 -->
      <c r="B1" s="1"/>                      <!-- B1必须存在但可以没值 -->
    </row>
    <row r="2">
      <c r="A2" t="s"><v>1</v></c>
      <c r="B2"><v>100</v></c>
    </row>
  </sheetData>
  
  <!-- 合并定义 -->
  <mergeCells count="1">
    <mergeCell ref="A1:B1"/>
  </mergeCells>
</worksheet>
```

### 5.3 合并的规则（非常重要）

| 规则 | 说明 |
|------|------|
| **`ref` 格式** | `左上角:右下角`，如 `A1:C3` 表示从A1到C3的矩形区域合并 |
| **`count` 必须准确** | `<mergeCells>` 的 `count` 属性必须等于 `<mergeCell>` 的数量 |
| **左上角单元格存值** | 合并后，只有左上角单元格的值会显示。其他被合并的单元格**应该为空或不存在** |
| **所有被合并的单元格必须有样式** | 即使 B1 在合并区域中但没有值，它也应该设置 `s` 属性（样式索引），否则边框可能显示异常 |
| **不能交叉合并** | 两个合并区域不能重叠，如 `A1:B2` 和 `B1:C2` 是非法的 |
| **`mergeCells` 位置** | 必须在 `<sheetData>` 之后，在 `<conditionalFormatting>` 之前 |

### 5.4 两行合并成一行 —— 具体实现

用户需求：**"两行合并成一行"**，实际就是跨行的垂直合并。

假设要实现这样的效果：
```
┌───────────┬───────┐
│           │  B1   │    ← A1和A2合并（第1、2行的A列合并成一个高格子）
│   A1:A2   ├───────┤
│           │  B2   │
├───────────┼───────┤
│    A3     │  B3   │    ← 第3行正常
└───────────┴───────┘
```

XML 实现：
```xml
<sheetData>
  <row r="1">
    <c r="A1" t="s" s="1"><v>0</v></c>     <!-- A1存值，样式1（有边框） -->
    <c r="B1" t="s" s="1"><v>1</v></c>     <!-- B1正常 -->
  </row>
  <row r="2">
    <c r="A2" s="1"/>                       <!-- A2：被合并，无值但有样式 -->
    <c r="B2" t="s" s="1"><v>2</v></c>     <!-- B2正常 -->
  </row>
  <row r="3">
    <c r="A3" t="s" s="1"><v>3</v></c>
    <c r="B3" t="s" s="1"><v>4</v></c>
  </row>
</sheetData>

<mergeCells count="1">
  <mergeCell ref="A1:A2"/>                 <!-- 垂直合并A1和A2 -->
</mergeCells>
```

### 5.5 水平合并 vs 垂直合并 vs 区域合并

```
水平合并（同行多列）：        垂直合并（同列多行）：        区域合并（多行多列）：
┌───────────┐               ┌───────┐                     ┌───────────┐
│  A1:C1    │               │       │                     │           │
├───┬───┬───┤               │  A1   │                     │  A1:C2   │
│A2 │B2 │C2 │               │  :A3  │                     │           │
└───┴───┴───┘               │       │                     ├───────────┤
                            ├───────┤                     │    A3     │
                            │  A4   │                     ├─────┬─────┤
                            └───────┘                     │ A4  │ B4  │
                                                          └─────┴─────┘

ref="A1:C1"                  ref="A1:A3"                    ref="A1:C2"
```

### 5.6 多个合并区域

```xml
<mergeCells count="3">
  <mergeCell ref="A1:C1"/>    <!-- 标题行：A1到C1合并 -->
  <mergeCell ref="A2:A4"/>    <!-- A列2-4行合并 -->
  <mergeCell ref="B3:C3"/>    <!-- 第3行B-C列合并 -->
</mergeCells>
```

> **注意**：`count="3"` 必须和实际的 `<mergeCell>` 数量一致，否则 Excel 会报错。

---

## 6. 交错行格式：让每一行"长得不一样"

### 6.1 什么是交错格式

用户说的"两行合并成一行和其他行交错"，本质上是**不同行有不同的合并模式**，形成视觉上的交错效果。

```
┌───────────┬───────┬───────┐
│           │  B1   │  C1   │    第1行：A1:A2合并（竖向跨2行）
│   A1:A2   ├───────┼───────┤
│           │  B2   │  C2   │
├───────────┼───────┴───────┤
│    A3     │    B3:C3      │    第3行：A3正常，B3:C3合并（横向）
├───────────┼───────┬───────┤
│           │  B4   │       │    第4行：A4:A5合并，C4:C5合并
│   A4:A5   ├───────┤ C4:C5 │
│           │  B5   │       │
└───────────┴───────┴───────┘
```

### 6.2 XML 实现

```xml
<sheetData>
  <row r="1">
    <c r="A1" t="s" s="1"><v>0</v></c>     <!-- "设备名称" -->
    <c r="B1" t="s" s="1"><v>1</v></c>     <!-- "型号" -->
    <c r="C1" t="s" s="1"><v>2</v></c>     <!-- "状态" -->
  </row>
  <row r="2">
    <c r="A2" s="1"/>                       <!-- 被A1:A2合并 -->
    <c r="B2" t="s" s="1"><v>3</v></c>
    <c r="C2" t="s" s="1"><v>4</v></c>
  </row>
  <row r="3">
    <c r="A3" t="s" s="1"><v>5</v></c>
    <c r="B3" t="s" s="1"><v>6</v></c>     <!-- 被B3:C3合并，左上角存值 -->
    <c r="C3" s="1"/>                       <!-- 被B3:C3合并 -->
  </row>
  <row r="4">
    <c r="A4" t="s" s="1"><v>7</v></c>     <!-- 被A4:A5合并，左上角存值 -->
    <c r="B4" t="s" s="1"><v>8</v></c>
    <c r="C4" t="s" s="1"><v>9</v></c>     <!-- 被C4:C5合并，左上角存值 -->
  </row>
  <row r="5">
    <c r="A5" s="1"/>                       <!-- 被A4:A5合并 -->
    <c r="B5" t="s" s="1"><v>10</v></c>
    <c r="C5" s="1"/>                       <!-- 被C4:C5合并 -->
  </row>
</sheetData>

<mergeCells count="4">
  <mergeCell ref="A1:A2"/>
  <mergeCell ref="B3:C3"/>
  <mergeCell ref="A4:A5"/>
  <mergeCell ref="C4:C5"/>
</mergeCells>
```

### 6.3 交错格式的设计思路

在代码中实现交错格式，关键是**建立一张"合并映射表"**：

```
步骤1：确定每行每列的合并情况
步骤2：遍历时，遇到合并区域的左上角就写值；遇到被合并的格子就留空但有样式
步骤3：所有合并区域记录到 mergeCells 列表
```

可以用一个二维数组来表示合并信息：

```c
// merge_info[row][col] 的值表示：
//   0  = 普通单元格
//   1  = 合并区域的左上角（存值）
//   -1 = 被合并的格子（不存值，但要有样式）
// 用一个结构体记录合并范围更清晰：

typedef struct {
    int start_row, start_col;
    int end_row, end_col;
} merge_range_t;

// 示例：上面的交错格式
merge_range_t merges[] = {
    {1, 1, 2, 1},    // A1:A2
    {3, 2, 3, 3},    // B3:C3
    {4, 1, 5, 1},    // A4:A5
    {4, 3, 5, 3},    // C4:C5
};
```

### 6.4 交错行数据的程序化填充

实际应用中，数据往往是动态的。以下是一个典型场景——**多传感器数据报表**，每个传感器有2个通道，用交错格式展示：

```
┌────────┬──────────┬──────────┐
│        │ 通道1    │ 通道2    │     ← 传感器名合并跨2列
│ 传感器A ├──────────┼──────────┤
│        │ 25.3°C   │ 24.8°C   │
├────────┼──────────┴──────────┤
│        │     均值: 25.05°C    │     ← 均值行合并B-C列
│ 传感器B ├──────────┬──────────┤
│        │ 30.1°C   │ 29.7°C   │
├────────┼──────────┴──────────┤
│        │     均值: 29.90°C    │
└────────┴──────────────────────┘
```

程序化生成的伪代码：

```c
void generate_sensor_report(sensor_data_t *sensors, int count) {
    int row = 1;
    
    // 写表头
    write_cell(row, 1, "传感器", STYLE_HEADER);
    write_cell(row, 2, "通道1", STYLE_HEADER);
    write_cell(row, 3, "通道2", STYLE_HEADER);
    row++;
    
    for (int i = 0; i < count; i++) {
        int name_start_row = row;
        
        // 传感器名（跨2行合并）
        write_cell(row, 1, sensors[i].name, STYLE_NORMAL);
        write_cell(row, 2, sensors[i].ch1_value, STYLE_DATA);
        write_cell(row, 3, sensors[i].ch2_value, STYLE_DATA);
        row++;
        
        // 均值行（B-C列合并）
        write_cell(row, 1, "", STYLE_NORMAL);       // 空，但要有样式
        write_cell(row, 2, avg_string, STYLE_AVG);   // 左上角存值
        write_cell(row, 3, "", STYLE_AVG);            // 被合并，有样式
        row++;
        
        // 记录合并区域
        add_merge(name_start_row, 1, name_start_row + 1, 1);  // 传感器名垂直合并
        add_merge(name_start_row + 1, 2, name_start_row + 1, 3); // 均值行水平合并
    }
}
```

---

## 7. 样式系统：颜色、边框、字体、对齐

### 7.1 样式的层级结构

```
单元格 s="5"
         │
         ▼
cellXfs[5] (样式组合)
  ├── fontId=2    → fonts[2]    → {粗体, 14号, 红色, 微软雅黑}
  ├── fillId=3    → fills[3]   → {实心填充, #FFC000}
  ├── borderId=1  → borders[1] → {四周细黑边框}
  ├── numFmtId=0  → 默认数字格式
  └── alignment   → {水平居中, 垂直居中, 自动换行}
```

### 7.2 对齐方式

对齐属性直接写在 `cellXfs` 的 `<xf>` 元素内：

```xml
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1">
  <alignment horizontal="center" vertical="center" wrapText="1"/>
</xf>
```

| 属性 | 可选值 | 说明 |
|------|--------|------|
| `horizontal` | `left`, `center`, `right`, `justify`, `distributed` | 水平对齐 |
| `vertical` | `top`, `center`, `bottom`, `justify`, `distributed` | 垂直对齐 |
| `wrapText` | `0` 或 `1` | 是否自动换行 |
| `textRotation` | 0-180 | 文字旋转角度（0-90顺时针，91-180逆时针） |
| `indent` | 0-250 | 缩进量 |
| `shrinkToFit` | `0` 或 `1` | 是否缩小字体以适应列宽 |

> **合并单元格的对齐**：合并后的单元格，对齐方式取**左上角单元格**的样式。所以如果你想让合并后的文字居中，需要给左上角单元格设置居中样式。

### 7.3 颜色编码

颜色在 OOXML 中用 ARGB 格式（注意不是 RGBA）：

```
FF000000
││└┴─────── RGB 部分（000000 = 黑色）
│└───────── GG 部分（00）
└────────── AA 部分（FF = 完全不透明）

常见颜色：
FF000000 = 黑色      FFFFFF00 = 黄色
FFFFFFFF = 白色      FFFF0000 = 红色
FF0000FF = 蓝色      FF00FF00 = 绿色
FF808080 = 灰色      FFFFFC000 = 橙色
```

前两位 `FF` 是 Alpha 通道（透明度），`FF` = 完全不透明。在 Excel 中一般固定用 `FF`。

### 7.4 边框样式

```xml
<border>
  <left style="thin"><color rgb="FF000000"/></left>
  <right style="medium"><color rgb="FF0000FF"/></right>
  <top style="thick"><color rgb="FFFF0000"/></top>
  <bottom style="dashed"><color rgb="FF00FF00"/></bottom>
  <diagonal style="dotted"><color rgb="FF808080"/></diagonal>
</border>
```

| style 值 | 效果 |
|----------|------|
| `thin` | 细线 |
| `medium` | 中等线 |
| `thick` | 粗线 |
| `dashed` | 虚线 |
| `dotted` | 点线 |
| `double` | 双线 |
| `mediumDashed` | 中等虚线 |
| `slantDashDot` | 斜划线点线 |

### 7.5 数字格式

```xml
<numFmts count="2">
  <numFmt numFmtId="164" formatCode="0.00"/>           <!-- 两位小数 -->
  <numFmt numFmtId="165" formatCode="#,##0"/>          <!-- 千分位 -->
  <numFmt numFmtId="166" formatCode="0.00%"/>          <!-- 百分比 -->
  <numFmt numFmtId="167" formatCode="yyyy-mm-dd"/>     <!-- 日期 -->
  <numFmt numFmtId="168" formatCode="0.00°C"/>          <!-- 带文字 -->
</numFmts>
```

**内置格式 ID（不需要在 numFmts 中定义，直接用）**：

| ID | 格式 | 示例 |
|----|------|------|
| 0 | 常规 | 1234.56 |
| 1 | 0 | 1235 |
| 2 | 0.00 | 1234.56 |
| 3 | #,##0 | 1,235 |
| 4 | #,##0.00 | 1,234.56 |
| 9 | 0% | 123456% |
| 10 | 0.00% | 123456.00% |
| 14 | yyyy/m/d | 2024/3/15 |
| 22 | yyyy/m/d h:mm | 2024/3/15 14:30 |

自定义格式 ID 从 **164** 开始（0-163 保留）。

---

## 8. 共享字符串表：省内存的关键技巧

### 8.1 为什么要共享

假设你的报表有 500 行数据，每行的"状态"列都写着"正常运行"：

```
不用共享：sheet1.xml 中重复出现 500 次 "正常运行"，浪费大量空间
使用共享：sharedStrings.xml 中只存 1 次 "正常运行"，sheet1.xml 中只写 <v>3</v>
```

### 8.2 实现方案

```c
#define MAX_SHARED_STRINGS 256   // 根据内存调整

typedef struct {
    char strings[MAX_SHARED_STRINGS][64];  // 字符串池
    int count;                              // 当前字符串数量
} shared_string_table_t;

// 查找或添加字符串，返回索引
int shared_string_add(shared_string_table_t *table, const char *str) {
    // 先查找是否已存在
    for (int i = 0; i < table->count; i++) {
        if (strcmp(table->strings[i], str) == 0) {
            return i;  // 已存在，返回索引
        }
    }
    // 不存在，添加新的
    if (table->count < MAX_SHARED_STRINGS) {
        strncpy(table->strings[table->count], str, 63);
        table->strings[table->count][63] = '\0';
        return table->count++;
    }
    return -1;  // 满了
}
```

### 8.3 输出 sharedStrings.xml

```c
void write_shared_strings(shared_string_table_t *table, FILE *f) {
    fprintf(f, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    fprintf(f, "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
    fprintf(f, "count=\"%d\" uniqueCount=\"%d\">\n", table->count, table->count);
    for (int i = 0; i < table->count; i++) {
        fprintf(f, "  <si><t>%s</t></si>\n", table->strings[i]);
    }
    fprintf(f, "</sst>");
}
```

### 8.4 XML 特殊字符转义

字符串写入 XML 时，必须转义以下字符：

```c
void xml_escape(const char *src, char *dst, int dst_size) {
    int j = 0;
    for (int i = 0; src[i] && j < dst_size - 8; i++) {
        switch (src[i]) {
            case '<':  dst[j++]=0x26; dst[j++]=0x6C; dst[j++]=0x74; dst[j++]=0x3B; break; // &lt;
            case '>':  dst[j++]=0x26; dst[j++]=0x67; dst[j++]=0x74; dst[j++]=0x3B; break; // &gt;
            case '&':  dst[j++]=0x26; dst[j++]=0x61; dst[j++]=0x6D; dst[j++]=0x70; dst[j++]=0x3B; break; // &amp;
            case '"':  dst[j++]=0x26; dst[j++]=0x71; dst[j++]=0x75; dst[j++]=0x6F; dst[j++]=0x74; dst[j++]=0x3B; break; // &quot;
            case '\'': dst[j++]=0x26; dst[j++]=0x61; dst[j++]=0x70; dst[j++]=0x6F; dst[j++]=0x73; dst[j++]=0x3B; break; // &apos;
            default:   dst[j++] = src[i]; break;
        }
    }
    dst[j] = '\0';
}
```

> **注意**：上面的转义用了字符的十六进制值而不是直接写字符串，是为了避免在 C 代码中写大量转义字符。实际可以用 `strcat` 和字符串字面量，更易读。

---

## 9. 嵌入式实现策略：在资源受限环境下的工程取舍

### 9.1 核心挑战

| 挑战 | 说明 | 对策 |
|------|------|------|
| **内存有限** | MCU 可能只有几十 KB RAM | 流式生成 XML，不缓存全部内容 |
| **无文件系统** | 可能没有 SD 卡或 Flash 文件系统 | 直接输出到串口/UART/SPI Flash/网络 |
| **无 ZIP 库** | 嵌入式环境通常没有 zlib | 用 Store 模式（不压缩），自己写 ZIP 头 |
| **无浮点支持** | 部分 MCU 无 FPU | 数字格式化用整数模拟 |
| **实时性要求** | 不能长时间阻塞 | 分块生成，每块控制在几 ms 内 |

### 9.2 三种实现方案对比

#### 方案 A：全量内存生成（适合 RAM > 256KB 的系统）

```
内存中构建全部 XML → 打包 ZIP → 一次输出
```

- 优点：实现简单，逻辑清晰
- 缺点：内存占用大，与表格大小成正比
- 适合：STM32F4/H7 系列、嵌入式 Linux

#### 方案 B：流式生成（适合 RAM 64KB-256KB）

```
逐行生成 XML → 逐个文件写入 ZIP → 边生成边输出
```
- 优点：内存占用恒定（不随表格大小增长）
- 缺点：实现复杂，需要处理 ZIP 的 CRC32 和偏移
- 适合：STM32F1/F3/G0 系列

#### 方案 C：最小化方案（适合 RAM < 64KB）

```
只生成核心文件 → 用内联字符串（跳过 sharedStrings）→ 最简样式 → Store 模式 ZIP
```
- 优点：极低内存占用（< 4KB RAM）
- 缺点：文件体积大（不压缩+不共享字符串），样式有限
- 适合：STM32F0/L0 系列、8位 MCU

### 9.3 流式 ZIP 生成的关键

流式生成的核心难点是：**ZIP 的中央目录需要知道每个文件的偏移量和 CRC32，但这些值在流式写入时还没算完。**

解决方案——**两遍法**：

```
第一遍：计算每个文件的 CRC32 和大小（不实际输出数据）
第二遍：输出 Local Header + 数据，同时记录偏移
第三遍：输出中央目录（此时已有所有偏移和 CRC32）
```

或者更优雅的**一遍法**（推荐）：

```
维护一个文件条目数组，每写入一个文件时：
  1. 记录当前偏移
  2. 写 Local Header（先填占位值）
  3. 流式写入数据，同时计算 CRC32
  4. 数据写完后，更新 Local Header 中的 CRC32 和大小
  5. 记录到中央目录条目数组

最后：
  6. 写中央目录
  7. 写 EOCD
```

> **注意**：如果输出目标是只读流（如 UART 发送），无法回头修改已写的 Local Header。此时只能在内存中先缓存完整的文件数据，算好 CRC32 后再输出。对于大文件，需分块处理。

### 9.4 省内存的技巧汇总

| 技巧 | 节省效果 | 实现难度 |
|------|---------|---------|
| 用共享字符串表 | 文本越多越省 | 低 |
| 用 Store 模式免压缩 | 省去压缩缓冲区 | 无 |
| 精简 XML 命名空间 | 减少每行冗余字符 | 低 |
| 跳过 docProps | 少 2 个文件 | 无 |
| 用短变量名（不推荐，Excel 可能不认） | 轻微 | 高风险 |
| 分块输出，不缓存完整 XML | 省大量 RAM | 中 |
| 预计算固定 XML 模板 | 省 sprintf 开销 | 中 |

### 9.5 精简 XML 命名空间

标准的命名空间声明很长：
```xml
xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
```

这个字符串必须完整保留（Excel 会检查），但你可以**只在一个文件中声明一次**，后续引用不需要重复。如果内存极度紧张，可以把这些长字符串存在 Flash/ROM 中，通过指针引用。

---

## 10. 最小 ZIP 写入器：不压缩也能生成合法 xlsx

### 10.1 为什么不需要压缩

Excel 打开 xlsx 时，**不要求 ZIP 中的文件必须被压缩**。Store 模式（method=0）是完全合法的。这意味着：

- 不需要 zlib / miniz
- 不需要 Deflate 算法
- 只需要 CRC32 + 文件头

### 10.2 CRC32 查找表

```c
// CRC32 查找表（固定值，放在 Flash 中）
static const uint32_t crc32_table[256] = {
    0x00000000, 0x77073096, 0xEE0E612C, 0x990951BA,
    0x076DC419, 0x706AF48F, 0xE963A535, 0x9E6495A3,
    0x0EDB8832, 0x79DCB8A4, 0xE0D5E91E, 0x97D2D988,
    0x09B64C2B, 0x7EB17CBD, 0xE7B82D07, 0x90BF1D91,
    0x1DB71064, 0x6AB020F2, 0xF3B97148, 0x84BE41DE,
    0x1ADAD47D, 0x6DDDE4EB, 0xF4D4B551, 0x83D385C7,
    0x136C9856, 0x646BA8C0, 0xFD62F97A, 0x8A65C9EC,
    0x14015C4F, 0x63066CD9, 0xFA0F3D63, 0x8D080DF5,
    0x3B6E20C8, 0x4C69105E, 0xD56041E4, 0xA2677172,
    0x3C03E4D1, 0x4B04D447, 0xD20D85FD, 0xA50AB56B,
    0x35B5A8FA, 0x42B2986C, 0xDBBBC9D6, 0xACBCF940,
    0x32D86CE3, 0x45DF5C75, 0xDCD60DCF, 0xABD13D59,
    0x26D930AC, 0x51DE003A, 0xC8D75180, 0xBFD06116,
    0x21B4F4B5, 0x56B3C423, 0xCFBA9599, 0xB8BDA50F,
    0x2802B89E, 0x5F058808, 0xC60CD9B2, 0xB10BE924,
    0x2F6F7C87, 0x58684C11, 0xC1611DAB, 0xB6662D3D,
    0x76DC4190, 0x01DB7106, 0x98D220BC, 0xEFD5102A,
    0x71B18589, 0x06B6B51F, 0x9FBFE4A5, 0xE8B8D433,
    0x7807C9A2, 0x0F00F934, 0x9609A88E, 0xE10E9818,
    0x7F6A0DBB, 0x086D3D2D, 0x91646C97, 0xE6635C01,
    0x6B6B51F4, 0x1C6C6162, 0x856530D8, 0xF262004E,
    0x6C0695ED, 0x1B01A57B, 0x8208F4C1, 0xF50FC457,
    0x65B0D9C6, 0x12B7E950, 0x8BBEB8EA, 0xFCB9887C,
    0x62DD1DDF, 0x15DA2D49, 0x8CD37CF3, 0xFBD44C65,
    0x4DB26158, 0x3AB551CE, 0xA3BC0074, 0xD4BB30E2,
    0x4ADFA541, 0x3DD895D7, 0xA4D1C46D, 0xD3D6F4FB,
    0x4369E96A, 0x346ED9FC, 0xAD678846, 0xDA60B8D0,
    0x44042D73, 0x33031DE5, 0xAA0A4C5F, 0xDD0D7CC9,
    0x5005713C, 0x270241AA, 0xBE0B1010, 0xC90C2086,
    0x5768B525, 0x206F85B3, 0xB966D409, 0xCE61E49F,
    0x5EDEF90E, 0x29D9C998, 0xB0D09822, 0xC7D7A8B4,
    0x59B33D17, 0x2EB40D81, 0xB7BD5C3B, 0xC0BA6CAD,
    0xEDB88320, 0x9ABFB3B6, 0x03B6E20C, 0x74B1D29A,
    0xEAD54739, 0x9DD277AF, 0x04DB2615, 0x73DC1683,
    0xE3630B12, 0x94643B84, 0x0D6D6A3E, 0x7A6A5AA8,
    0xE40ECF0B, 0x9309FF9D, 0x0A00AE27, 0x7D079EB1,
    0xF00F9344, 0x8708A3D2, 0x1E01F268, 0x6906C2FE,
    0xF7625756, 0x806567CB, 0x196C3671, 0x6E6B06E7,
    0xFED41B72, 0x89D32BE0, 0x10DA7A5A, 0x67DD4ACC,
    0xF9B9DF6F, 0x8EBEEFF9, 0x17B7BE43, 0x60B08ED5,
    0xD6D6A3E8, 0xA1D1937E, 0x38D8C2C4, 0x4FDF25E3,
    0xD1BB67F1, 0xA6BC5767, 0x3FB506DD, 0x48B2364B,
    0xD80D2BDA, 0xAF0A1B4C, 0x36034AF6, 0x41047A60,
    0xDF60EFC3, 0xA8670855, 0x316E8EEF, 0x4669BE79,
    0xCB61B38C, 0xBC66831A, 0x256FD2A0, 0x5268E236,
    0xCC0C7795, 0xBB0B4703, 0x220216B9, 0x5505262F,
    0xC5BA3BBE, 0xB2BD0B28, 0x2BB45A92, 0x5CB36A04,
    0xC2D7FFA7, 0xB5D0CF31, 0x2CD99E8B, 0x5BDEAE1D,
    0x9B64C2B0, 0xEC63F226, 0x756AA39C, 0x026D930A,
    0x9C0906A9, 0xEB0E363F, 0x72076785, 0x05005713,
    0x95BF4A82, 0xE2B87A14, 0x7BB12BAE, 0x0CB61B38,
    0x92D28E9B, 0xE5D5BE0D, 0x7CDCEFB7, 0x0BDBDF21,
    0x86D3D2D4, 0xF1D4E242, 0x68DDB3F8, 0x1FDA836E,
    0x81BE16CD, 0xF6B9265B, 0x6FB077E1, 0x18B74777,
    0x88085AE6, 0xFF0F6A70, 0x66063BCA, 0x11010B5C,
    0x8F659EFF, 0xF862AE69, 0x616BFFD3, 0x166CCF45,
    0xA00AE278, 0xD70DD2EE, 0x4E048354, 0x3903B3C2,
    0xA7672661, 0xD06016F7, 0x4969474D, 0x3E6E77DB,
    0xAED16A4A, 0xD9D65ADC, 0x40DF0B66, 0x37D83BF0,
    0xA9BCAE53, 0xDEBB9EC5, 0x47B2CF7F, 0x30B5FFE9,
    0xBDBDF21C, 0xCABAC28A, 0x53B39330, 0x24B4A3A6,
    0xBAD03605, 0xCDD70693, 0x54DE5729, 0x23D967BF,
    0xB3667A2E, 0xC4614AB8, 0x5D681B02, 0x2A6F2B94,
    0xB40BBE37, 0xC30C8EA1, 0x5A05DF1B, 0x2D02EF8D,
};

uint32_t crc32_compute(const uint8_t *data, size_t len) {
    uint32_t crc = 0xFFFFFFFF;
    for (size_t i = 0; i < len; i++) {
        crc = crc32_table[(crc ^ data[i]) & 0xFF] ^ (crc >> 8);
    }
    return crc ^ 0xFFFFFFFF;
}
```

> **硬件 CRC 加速**：STM32 等 MCU 自带 CRC 硬件模块，但注意它的多项式可能与 ZIP 要求的不同（ZIP 用标准 CRC-32，多项式 0xEDB88320）。需要确认硬件 CRC 模块是否兼容，否则用软件表驱动更安全。

---

## 11. 完整代码方案：一个可用的嵌入式 Excel 生成器

以下是一个完整的、可直接编译运行的 C 语言实现，包含 ZIP 打包、XML 生成、合并单元格等全部功能。

> 代码较长，但每一部分都有注释。建议先通读结构，再逐段理解。

### 11.1 数据结构定义

```c
#include <stdint.h>
#include <string.h>
#include <stdio.h>
#include <stdlib.h>

/* ======================== 数据结构 ======================== */

// 合并区域定义
typedef struct {
    int start_row;  // 起始行 (1-based)
    int start_col;  // 起始列 (1-based)
    int end_row;    // 结束行
    int end_col;    // 结束列
} merge_range_t;

#define MAX_MERGES 64

// 单元格数据
typedef struct {
    int row;
    int col;
    int style_index;     // 样式索引
    int is_string;       // 1=字符串(用共享字符串索引), 0=数字
    int string_index;    // 共享字符串索引 (is_string=1时有效)
    double number;       // 数字值 (is_string=0时有效)
    int has_value;       // 是否有值(被合并的单元格没有值)
} cell_t;

#define MAX_CELLS 1024

// 列宽定义
typedef struct {
    int col;
    double width;
} col_width_t;

#define MAX_COL_WIDTHS 32

// 共享字符串表
#define MAX_SHARED_STRINGS 256
#define MAX_STRING_LEN 64

typedef struct {
    char strings[MAX_SHARED_STRINGS][MAX_STRING_LEN];
    int count;
} sst_t;

// 工作表数据
typedef struct {
    cell_t cells[MAX_CELLS];
    int cell_count;
    merge_range_t merges[MAX_MERGES];
    int merge_count;
    col_width_t col_widths[MAX_COL_WIDTHS];
    int col_width_count;
    int max_row;
    int max_col;
} worksheet_t;

// 整个工作簿
typedef struct {
    sst_t sst;                // 共享字符串表
    worksheet_t sheets[4];    // 最多4个工作表
    int sheet_count;
    char sheet_names[4][32];  // 工作表名称
    
    // 样式相关 (简化版：预定义几种样式)
    // 样式索引说明：
    // 0 = 默认
    // 1 = 标题（粗体、居中、灰色背景、边框）
    // 2 = 数据（边框、左对齐）
    // 3 = 合并单元格的左上角（粗体、居中、边框）
    // 4 = 强调（红色字体、边框）
    // 5 = 百分比格式
} workbook_t;
```

### 11.2 列字母转换

```c
/* ======================== 列字母转换 ======================== */

// 列号(1-based) → 列字母，如 1→"A", 27→"AA"
void col_to_letters(int col, char *buf, int buf_size) {
    int i = 0;
    while (col > 0 && i < buf_size - 1) {
        col--;
        buf[i++] = 'A' + (col % 26);
        col /= 26;
    }
    buf[i] = '\0';
    // 反转
    for (int j = 0; j < i / 2; j++) {
        char tmp = buf[j];
        buf[j] = buf[i - 1 - j];
        buf[i - 1 - j] = tmp;
    }
}

// 生成单元格地址，如 (3, 2) → "B3"
void cell_address(int row, int col, char *buf, int buf_size) {
    char col_str[8];
    col_to_letters(col, col_str, sizeof(col_str));
    snprintf(buf, buf_size, "%s%d", col_str, row);
}
```

### 11.3 共享字符串表操作

```c
/* ======================== 共享字符串表 ======================== */

void sst_init(sst_t *sst) {
    sst->count = 0;
}

// XML 特殊字符转义
void xml_escape(const char *src, char *dst, int dst_size) {
    int j = 0;
    for (int i = 0; src[i] && j < dst_size - 7; i++) {
        switch (src[i]) {
            case '<':  dst[j++]='&'; dst[j++]='l'; dst[j++]='t'; dst[j++]=';'; break;
            case '>':  dst[j++]='&'; dst[j++]='g'; dst[j++]='t'; dst[j++]=';'; break;
            case '&':  dst[j++]='&'; dst[j++]='a'; dst[j++]='m'; dst[j++]='p'; dst[j++]=';'; break;
            case '"':  dst[j++]='&'; dst[j++]='q'; dst[j++]='u'; dst[j++]='o'; dst[j++]='t'; dst[j++]=';'; break;
            default:   dst[j++] = src[i]; break;
        }
    }
    dst[j] = '\0';
}

// 查找或添加字符串，返回索引
int sst_add(sst_t *sst, const char *str) {
    for (int i = 0; i < sst->count; i++) {
        if (strncmp(sst->strings[i], str, MAX_STRING_LEN) == 0) {
            return i;
        }
    }
    if (sst->count < MAX_SHARED_STRINGS) {
        strncpy(sst->strings[sst->count], str, MAX_STRING_LEN - 1);
        sst->strings[sst->count][MAX_STRING_LEN - 1] = '\0';
        return sst->count++;
    }
    return 0; // 满了，返回0
}
```

### 11.4 工作表操作

```c
/* ======================== 工作表操作 ======================== */

void ws_init(worksheet_t *ws) {
    memset(ws, 0, sizeof(*ws));
}

// 添加字符串单元格
void ws_set_string(worksheet_t *ws, sst_t *sst, int row, int col, 
                   const char *str, int style) {
    if (ws->cell_count >= MAX_CELLS) return;
    cell_t *c = &ws->cells[ws->cell_count++];
    c->row = row;
    c->col = col;
    c->style_index = style;
    c->is_string = 1;
    c->string_index = sst_add(sst, str);
    c->has_value = 1;
    if (row > ws->max_row) ws->max_row = row;
    if (col > ws->max_col) ws->max_col = col;
}

// 添加数字单元格
void ws_set_number(worksheet_t *ws, int row, int col, 
                   double num, int style) {
    if (ws->cell_count >= MAX_CELLS) return;
    cell_t *c = &ws->cells[ws->cell_count++];
    c->row = row;
    c->col = col;
    c->style_index = style;
    c->is_string = 0;
    c->number = num;
    c->has_value = 1;
    if (row > ws->max_row) ws->max_row = row;
    if (col > ws->max_col) ws->max_col = col;
}

// 添加空单元格（用于被合并的格子，需要样式）
void ws_set_empty(worksheet_t *ws, int row, int col, int style) {
    if (ws->cell_count >= MAX_CELLS) return;
    cell_t *c = &ws->cells[ws->cell_count++];
    c->row = row;
    c->col = col;
    c->style_index = style;
    c->has_value = 0;
    if (row > ws->max_row) ws->max_row = row;
    if (col > ws->max_col) ws->max_col = col;
}

// 添加合并区域
void ws_add_merge(worksheet_t *ws, int r1, int c1, int r2, int c2) {
    if (ws->merge_count >= MAX_MERGES) return;
    ws->merges[ws->merge_count].start_row = r1;
    ws->merges[ws->merge_count].start_col = c1;
    ws->merges[ws->merge_count].end_row = r2;
    ws->merges[ws->merge_count].end_col = c2;
    ws->merge_count++;
}

// 设置列宽
void ws_set_col_width(worksheet_t *ws, int col, double width) {
    if (ws->col_width_count >= MAX_COL_WIDTHS) return;
    ws->col_widths[ws->col_width_count].col = col;
    ws->col_widths[ws->col_width_count].width = width;
    ws->col_width_count++;
}
```

### 11.5 XML 生成

```c
/* ======================== XML 生成 ======================== */

// 生成 sharedStrings.xml
void gen_shared_strings(sst_t *sst, FILE *f) {
    fprintf(f, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    fprintf(f, "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
    fprintf(f, "count=\"%d\" uniqueCount=\"%d\">\n", sst->count, sst->count);
    char escaped[MAX_STRING_LEN * 2];
    for (int i = 0; i < sst->count; i++) {
        xml_escape(sst->strings[i], escaped, sizeof(escaped));
        fprintf(f, "<si><t>%s</t></si>\n", escaped);
    }
    fprintf(f, "</sst>");
}

// 生成 styles.xml（预定义样式）
void gen_styles(FILE *f) {
    fprintf(f,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
        "  <fonts count=\"3\">\n"
        "    <font><sz val=\"11\"/><name val=\"等线\"/></font>\n"                              // 0: 默认
        "    <font><b/><sz val=\"12\"/><color rgb=\"FF000000\"/><name val=\"等线\"/></font>\n"  // 1: 粗体
        "    <font><b/><sz val=\"12\"/><color rgb=\"FFFF0000\"/><name val=\"等线\"/></font>\n"  // 2: 粗体红色
        "  </fonts>\n"
        "  <fills count=\"3\">\n"
        "    <fill><patternFill patternType=\"none\"/></fill>\n"
        "    <fill><patternFill patternType=\"gray125\"/></fill>\n"
        "    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFD9E1F2\"/></patternFill></fill>\n" // 淡蓝
        "  </fills>\n"
        "  <borders count=\"2\">\n"
        "    <border><left/><right/><top/><bottom/><diagonal/></border>\n"                    // 0: 无边框
        "    <border><left style=\"thin\"><color rgb=\"FF000000\"/></left>"                   // 1: 细黑边框
        "      <right style=\"thin\"><color rgb=\"FF000000\"/></right>"
        "      <top style=\"thin\"><color rgb=\"FF000000\"/></top>"
        "      <bottom style=\"thin\"><color rgb=\"FF000000\"/></bottom>"
        "      <diagonal/></border>\n"
        "  </borders>\n"
        "  <cellStyleXfs count=\"1\">\n"
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n"
        "  </cellStyleXfs>\n"
        "  <cellXfs count=\"7\">\n"
        // 索引0: 默认
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>\n"
        // 索引1: 标题（粗体+居中+淡蓝背景+边框）
        "    <xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"1\" xfId=\"0\" "
        "       applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\">"
        "      <alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\"/></xf>\n"
        // 索引2: 数据（边框+左对齐）
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" "
        "       applyBorder=\"1\" applyAlignment=\"1\">"
        "      <alignment horizontal=\"left\" vertical=\"center\"/></xf>\n"
        // 索引3: 合并区域左上角（粗体+居中+边框）
        "    <xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"1\" xfId=\"0\" "
        "       applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\">"
        "      <alignment horizontal=\"center\" vertical=\"center\"/></xf>\n"
        // 索引4: 强调（红色粗体+边框）
        "    <xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"1\" xfId=\"0\" "
        "       applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\">"
        "      <alignment horizontal=\"center\" vertical=\"center\"/></xf>\n"
        // 索引5: 数字格式（两位小数+边框+右对齐）
        "    <xf numFmtId=\"2\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" "
        "       applyBorder=\"1\" applyAlignment=\"1\" applyNumberFormat=\"1\">"
        "      <alignment horizontal=\"right\" vertical=\"center\"/></xf>\n"
        // 索引6: 数据居中（边框+居中）
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" "
        "       applyBorder=\"1\" applyAlignment=\"1\">"
        "      <alignment horizontal=\"center\" vertical=\"center\"/></xf>\n"
        "  </cellXfs>\n"
        "  <cellStyles count=\"1\">\n"
        "    <cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>\n"
        "  </cellStyles>\n"
        "</styleSheet>"
    );
}

// 生成 worksheet XML
void gen_worksheet(worksheet_t *ws, FILE *f) {
    fprintf(f, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    fprintf(f, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
    fprintf(f, "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n");
    
    // 列宽定义
    if (ws->col_width_count > 0) {
        fprintf(f, "  <cols>\n");
        for (int i = 0; i < ws->col_width_count; i++) {
            fprintf(f, "    <col min=\"%d\" max=\"%d\" width=\"%.1f\" customWidth=\"1\"/>\n",
                    ws->col_widths[i].col, ws->col_widths[i].col, 
                    ws->col_widths[i].width);
        }
        fprintf(f, "  </cols>\n");
    }
    
    // 逐行输出
    fprintf(f, "  <sheetData>\n");
    for (int row = 1; row <= ws->max_row; row++) {
        fprintf(f, "    <row r=\"%d\">\n", row);
        
        // 找当前行的所有单元格
        for (int col = 1; col <= ws->max_col; col++) {
            cell_t *cell = NULL;
            for (int i = 0; i < ws->cell_count; i++) {
                if (ws->cells[i].row == row && ws->cells[i].col == col) {
                    cell = &ws->cells[i];
                    break;
                }
            }
            if (!cell) continue;
            
            char addr[16];
            cell_address(row, col, addr, sizeof(addr));
            
            if (cell->has_value) {
                if (cell->is_string) {
                    fprintf(f, "      <c r=\"%s\" t=\"s\" s=\"%d\"><v>%d</v></c>\n",
                            addr, cell->style_index, cell->string_index);
                } else {
                    fprintf(f, "      <c r=\"%s\" s=\"%d\"><v>%g</v></c>\n",
                            addr, cell->style_index, cell->number);
                }
            } else {
                fprintf(f, "      <c r=\"%s\" s=\"%d\"/>\n", addr, cell->style_index);
            }
        }
        fprintf(f, "    </row>\n");
    }
    fprintf(f, "  </sheetData>\n");
    
    // 合并单元格
    if (ws->merge_count > 0) {
        fprintf(f, "  <mergeCells count=\"%d\">\n", ws->merge_count);
        for (int i = 0; i < ws->merge_count; i++) {
            char a1[16], a2[16];
            cell_address(ws->merges[i].start_row, ws->merges[i].start_col, 
                       a1, sizeof(a1));
            cell_address(ws->merges[i].end_row, ws->merges[i].end_col, 
                       a2, sizeof(a2));
            fprintf(f, "    <mergeCell ref=\"%s:%s\"/>\n", a1, a2);
        }
        fprintf(f, "  </mergeCells>\n");
    }
    
    fprintf(f, "</worksheet>");
}
```

### 11.6 ZIP 打包器

```c
/* ======================== ZIP 打包器（Store 模式） ======================== */

// CRC32 查找表（见第10节，此处省略，假设已定义）
extern const uint32_t crc32_table[256];

uint32_t crc32_compute(const uint8_t *data, size_t len) {
    uint32_t crc = 0xFFFFFFFF;
    for (size_t i = 0; i < len; i++) {
        crc = crc32_table[(crc ^ data[i]) & 0xFF] ^ (crc >> 8);
    }
    return crc ^ 0xFFFFFFFF;
}

// ZIP 文件条目记录
typedef struct {
    char filename[128];
    uint32_t offset;        // 在 ZIP 中的偏移
    uint32_t compressed_size;
    uint32_t uncompressed_size;
    uint32_t crc32;
} zip_entry_t;

#define MAX_ZIP_ENTRIES 32

typedef struct {
    FILE *file;                        // 输出文件
    zip_entry_t entries[MAX_ZIP_ENTRIES];
    int entry_count;
    uint32_t current_offset;           // 当前写入位置
} zip_writer_t;

// MS-DOS 时间编码
uint16_t dos_time(int h, int m, int s) {
    return (h << 11) | (m << 5) | (s / 2);
}

// MS-DOS 日期编码
uint16_t dos_date(int y, int mo, int d) {
    return ((y - 1980) << 9) | (mo << 5) | d;
}

// 写入小端 uint16
void write_u16(FILE *f, uint16_t v) {
    fputc(v & 0xFF, f);
    fputc((v >> 8) & 0xFF, f);
}

// 写入小端 uint32
void write_u32(FILE *f, uint32_t v) {
    fputc(v & 0xFF, f);
    fputc((v >> 8) & 0xFF, f);
    fputc((v >> 16) & 0xFF, f);
    fputc((v >> 24) & 0xFF, f);
}

// 添加一个文件到 ZIP
void zip_add_file(zip_writer_t *zw, const char *filename, 
                  const char *data, size_t size) {
    if (zw->entry_count >= MAX_ZIP_ENTRIES) return;
    
    zip_entry_t *e = &zw->entries[zw->entry_count++];
    strncpy(e->filename, filename, sizeof(e->filename) - 1);
    e->filename[sizeof(e->filename) - 1] = '\0';
    e->offset = zw->current_offset;
    e->uncompressed_size = size;
    e->compressed_size = size;  // Store 模式，大小不变
    e->crc32 = crc32_compute((const uint8_t *)data, size);
    
    size_t filename_len = strlen(filename);
    
    // 写 Local File Header
    write_u32(zw->file, 0x04034b50);           // 签名
    write_u16(zw->file, 20);                    // 版本
    write_u16(zw->file, 0);                     // 标志
    write_u16(zw->file, 0);                     // 压缩方式: 0=Store
    write_u16(zw->file, dos_time(14, 30, 0));   // 时间
    write_u16(zw->file, dos_date(2024, 1, 1));  // 日期
    write_u32(zw->file, e->crc32);              // CRC32
    write_u32(zw->file, e->compressed_size);    // 压缩后大小
    write_u32(zw->file, e->uncompressed_size);  // 未压缩大小
    write_u16(zw->file, filename_len);          // 文件名长度
    write_u16(zw->file, 0);                     // 额外字段长度
    
    // 写文件名
    fwrite(filename, 1, filename_len, zw->file);
    
    // 写数据
    fwrite(data, 1, size, zw->file);
    
    zw->current_offset += 30 + filename_len + size;
}

// 完成 ZIP（写中央目录和 EOCD）
void zip_finalize(zip_writer_t *zw) {
    uint32_t cd_start = zw->current_offset;
    uint32_t cd_size = 0;
    
    // 写中央目录
    for (int i = 0; i < zw->entry_count; i++) {
        zip_entry_t *e = &zw->entries[i];
        size_t filename_len = strlen(e->filename);
        
        write_u32(zw->file, 0x02014b50);        // 签名
        write_u16(zw->file, 20);                 // 版本制作
        write_u16(zw->file, 20);                 // 版本所需
        write_u16(zw->file, 0);                   // 标志
        write_u16(zw->file, 0);                   // 压缩方法
        write_u16(zw->file, dos_time(14, 30, 0));
        write_u16(zw->file, dos_date(2024, 1, 1));
        write_u32(zw->file, e->crc32);
        write_u32(zw->file, e->compressed_size);
        write_u32(zw->file, e->uncompressed_size);
        write_u16(zw->file, filename_len);
        write_u16(zw->file, 0);                  // 额外字段长度
        write_u16(zw->file, 0);                  // 文件注释长度
        write_u16(zw->file, 0);                  // 磁盘编号
        write_u16(zw->file, 0);                  // 内部属性
        write_u32(zw->file, 0);                  // 外部属性
        write_u32(zw->file, e->offset);          // Local Header 偏移
        
        fwrite(e->filename, 1, filename_len, zw->file);
        
        cd_size += 46 + filename_len;
    }
    
    // 写 EOCD
    write_u32(zw->file, 0x06054b50);             // 签名
    write_u16(zw->file, 0);                       // 磁盘编号
    write_u16(zw->file, 0);                       // 中央目录起始磁盘
    write_u16(zw->file, zw->entry_count);          // 本磁盘条目数
    write_u16(zw->file, zw->entry_count);          // 总条目数
    write_u32(zw->file, cd_size);                 // 中央目录大小
    write_u32(zw->file, cd_start);                // 中央目录偏移
    write_u16(zw->file, 0);                        // 注释长度
}

void zip_init(zip_writer_t *zw, FILE *f) {
    zw->file = f;
    zw->entry_count = 0;
    zw->current_offset = 0;
}
```

### 11.7 最终打包函数

```c
/* ======================== 最终打包 ======================== */

// 将内存中的数据写入临时文件，再读回来添加到 ZIP
// 实际实现中，可以直接用内存缓冲区

typedef struct {
    char *data;
    size_t size;
    size_t capacity;
} membuf_t;

void membuf_init(membuf_t *buf, size_t initial) {
    buf->capacity = initial;
    buf->data = malloc(initial);
    buf->size = 0;
}

void membuf_write(membuf_t *buf, const char *fmt, ...) {
    va_list ap;
    va_start(ap, fmt);
    int needed = vsnprintf(NULL, 0, fmt, ap);
    va_end(ap);
    
    if (buf->size + needed + 1 > buf->capacity) {
        while (buf->size + needed + 1 > buf->capacity) {
            buf->capacity *= 2;
        }
        buf->data = realloc(buf->data, buf->capacity);
    }
    
    va_start(ap, fmt);
    vsnprintf(buf->data + buf->size, needed + 1, fmt, ap);
    va_end(ap);
    buf->size += needed;
}

// 主函数：生成 xlsx 文件
void generate_xlsx(workbook_t *wb, const char *output_path) {
    FILE *f = fopen(output_path, "wb");
    if (!f) return;
    
    zip_writer_t zw;
    zip_init(&zw, f);
    
    // 1. [Content_Types].xml
    const char *content_types =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
        "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
        "</Types>";
    zip_add_file(&zw, "[Content_Types].xml", content_types, strlen(content_types));
    
    // 2. _rels/.rels
    const char *root_rels =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
        "</Relationships>";
    zip_add_file(&zw, "_rels/.rels", root_rels, strlen(root_rels));
    
    // 3. xl/workbook.xml
    char workbook_xml[1024];
    snprintf(workbook_xml, sizeof(workbook_xml),
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheets><sheet name=\"%s\" sheetId=\"1\" r:id=\"rId1\"/></sheets>"
        "</workbook>",
        wb->sheet_names[0]);
    zip_add_file(&zw, "xl/workbook.xml", workbook_xml, strlen(workbook_xml));
    
    // 4. xl/_rels/workbook.xml.rels
    const char *wb_rels =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>"
        "</Relationships>";
    zip_add_file(&zw, "xl/_rels/workbook.xml.rels", wb_rels, strlen(wb_rels));
    
    // 5. xl/styles.xml
    membuf_t style_buf;
    membuf_init(&style_buf, 4096);
    // 用 fopen 到内存的方式... 实际上用 membuf 模拟 FILE*
    // 简化：直接拼字符串（见 gen_styles 函数，内容固定）
    // 这里用一个临时文件
    FILE *tmp = tmpfile();
    gen_styles(tmp);
    fseek(tmp, 0, SEEK_END);
    long styles_size = ftell(tmp);
    fseek(tmp, 0, SEEK_SET);
    char *styles_data = malloc(styles_size);
    fread(styles_data, 1, styles_size, tmp);
    fclose(tmp);
    zip_add_file(&zw, "xl/styles.xml", styles_data, styles_size);
    free(styles_data);
    
    // 6. xl/sharedStrings.xml
    tmp = tmpfile();
    gen_shared_strings(&wb->sst, tmp);
    fseek(tmp, 0, SEEK_END);
    long sst_size = ftell(tmp);
    fseek(tmp, 0, SEEK_SET);
    char *sst_data = malloc(sst_size);
    fread(sst_data, 1, sst_size, tmp);
    fclose(tmp);
    zip_add_file(&zw, "xl/sharedStrings.xml", sst_data, sst_size);
    free(sst_data);
    
    // 7. xl/worksheets/sheet1.xml
    tmp = tmpfile();
    gen_worksheet(&wb->sheets[0], tmp);
    fseek(tmp, 0, SEEK_END);
    long sheet_size = ftell(tmp);
    fseek(tmp, 0, SEEK_SET);
    char *sheet_data = malloc(sheet_size);
    fread(sheet_data, 1, sheet_size, tmp);
    fclose(tmp);
    zip_add_file(&zw, "xl/worksheets/sheet1.xml", sheet_data, sheet_size);
    free(sheet_data);
    
    // 完成 ZIP
    zip_finalize(&zw);
    
    fclose(f);
}
```

### 11.8 使用示例：生成交错格式报表

```c
/* ======================== 使用示例 ======================== */

int main() {
    workbook_t wb;
    memset(&wb, 0, sizeof(wb));
    strcpy(wb.sheet_names[0], "传感器数据");
    wb.sheet_count = 1;
    
    worksheet_t *ws = &wb.sheets[0];
    ws_init(ws);
    
    // 设置列宽
    ws_set_col_width(ws, 1, 18);   // A列
    ws_set_col_width(ws, 2, 15);   // B列
    ws_set_col_width(ws, 3, 15);   // C列
    
    // === 标题行 ===
    ws_set_string(ws, &wb.sst, 1, 1, "传感器数据报表", 1);
    ws_set_empty(ws, 1, 2, 1);
    ws_set_empty(ws, 1, 3, 1);
    ws_add_merge(ws, 1, 1, 1, 3);  // A1:C1 合并标题
    
    // === 表头 ===
    ws_set_string(ws, &wb.sst, 2, 1, "传感器名称", 1);
    ws_set_string(ws, &wb.sst, 2, 2, "通道1", 1);
    ws_set_string(ws, &wb.sst, 2, 3, "通道2", 1);
    
    // === 传感器A（交错格式：名称跨2行合并） ===
    ws_set_string(ws, &wb.sst, 3, 1, "传感器A", 3);     // 名称，跨2行
    ws_set_number(ws, 3, 2, 25.3, 5);                     // 通道1数据
    ws_set_number(ws, 3, 3, 24.8, 5);                     // 通道2数据
    ws_set_empty(ws, 4, 1, 3);                            // A4被合并
    ws_set_string(ws, &wb.sst, 4, 2, "均值: 25.05°C", 6); // B4:C4 合并
    ws_set_empty(ws, 4, 3, 6);                            // C4被合并
    ws_add_merge(ws, 3, 1, 4, 1);  // A3:A4 垂直合并
    ws_add_merge(ws, 4, 2, 4, 3);  // B4:C4 水平合并
    
    // === 传感器B ===
    ws_set_string(ws, &wb.sst, 5, 1, "传感器B", 3);
    ws_set_number(ws, 5, 2, 30.1, 5);
    ws_set_number(ws, 5, 3, 29.7, 5);
    ws_set_empty(ws, 6, 1, 3);
    ws_set_string(ws, &wb.sst, 6, 2, "均值: 29.90°C", 6);
    ws_set_empty(ws, 6, 3, 6);
    ws_add_merge(ws, 5, 1, 6, 1);
    ws_add_merge(ws, 6, 2, 6, 3);
    
    // === 传感器C（异常数据用红色标记） ===
    ws_set_string(ws, &wb.sst, 7, 1, "传感器C", 3);
    ws_set_number(ws, 7, 2, 85.2, 4);    // 异常值，用红色样式
    ws_set_number(ws, 7, 3, 22.1, 5);
    ws_set_empty(ws, 8, 1, 3);
    ws_set_string(ws, &wb.sst, 8, 2, "均值: 53.65°C [异常!]", 4);
    ws_set_empty(ws, 8, 3, 4);
    ws_add_merge(ws, 7, 1, 8, 1);
    ws_add_merge(ws, 8, 2, 8, 3);
    
    // 生成文件
    generate_xlsx(&wb, "report.xlsx");
    
    printf("report.xlsx 生成成功！\n");
    return 0;
}
```

---

## 12. 高级：各种"奇葩"格式的实现配方

### 12.1 配方速查表

| 格式效果 | 实现方法 | 关键 XML |
|---------|---------|----------|
| 两行合并成一行 | 垂直合并 | `<mergeCell ref="A1:A2"/>` |
| 一行跨多列 | 水平合并 | `<mergeCell ref="A1:C1"/>` |
| 交错行 | 不同行用不同的合并模式 | 多个 `<mergeCell>` |
| 标题居中大字 | 粗体字体 + 居中对齐 + 合并 | `fontId` + `alignment` |
| 隔行变色 | 偶数行用不同 `fillId` 的样式 | 交替 `s` 属性 |
| 合并区域加边框 | 所有被合并的单元格都设置有边框的样式 | 每个被合并的 cell 都设 `s` |
| 冻结首行 | `<sheetView>` 中的 `<pane>` | 见下文 |
| 自动列宽 | 不设 `customWidth`，让 Excel 自适应 | 省略 `<col>` |
| 超大行高 | `ht` 属性设大值 | `<row ht="100" customHeight="1">` |
| 隐藏行 | `hidden="1"` | `<row r="5" hidden="1">` |
| 文字旋转 | `textRotation` 属性 | `<alignment textRotation="45"/>` |

### 12.2 冻结窗格

```xml
<sheetViews>
  <sheetView workbookViewId="0">
    <!-- 冻结前2行 -->
    <pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/>
    <selection pane="bottomLeft" activeCell="A3" sqref="A3"/>
  </sheetView>
</sheetViews>
```

> **注意**：`<sheetViews>` 必须放在 `<sheetData>` 之前，是 `<worksheet>` 的第一个子元素。

### 12.3 隔行变色

实现效果（斑马纹）：

```
┌───────────┬───────────┐
│  数据1    │  数据2    │    ← 白色背景 (style A)
├───────────┼───────────┤
│  数据3    │  数据4    │    ← 浅蓝背景 (style B)
├───────────┼───────────┤
│  数据5    │  数据6    │    ← 白色背景 (style A)
└───────────┴───────────┘
```

代码中交替使用样式索引：
```c
for (int i = 0; i < data_count; i++) {
    int style = (i % 2 == 0) ? STYLE_NORMAL : STYLE_ALT_ROW;
    ws_set_string(ws, &sst, i + 2, 1, data[i].name, style);
    ws_set_number(ws, i + 2, 2, data[i].value, style);
}
```

需要在 styles.xml 中多定义一个带背景色的样式。

### 12.4 多级表头（三层交错合并）

```
┌───────────────────┬───────────────────────────────┐
│                   │          测量数据              │  ← B1:D1 合并
│     设备信息       ├───────────┬───────────────────┤
│                   │  温度     │     湿度           │  ← B2, C2:D2 合并
├───────┬───────────┼───────────┼───────┬───────────┤
│ 编号  │ 名称      │ 通道1     │ 通道1 │ 通道2     │  ← 第3行普通
├───────┼───────────┼───────────┼───────┼───────────┤
│ 001   │ 传感器A   │ 25.3      │ 45.2  │ 46.1      │
└───────┴───────────┴───────────┴───────┴───────────┘
```

XML 核心：
```xml
<sheetData>
  <row r="1">
    <c r="A1" t="s" s="1"><v>0</v></c>     <!-- "设备信息" -->
    <c r="B1" t="s" s="1"><v>1</v></c>     <!-- "测量数据" -->
    <c r="C1" s="1"/>                       <!-- 被合并 -->
    <c r="D1" s="1"/>                       <!-- 被合并 -->
  </row>
  <row r="2">
    <c r="A2" s="1"/>                       <!-- 被A1:A2合并 -->
    <c r="B2" t="s" s="1"><v>2</v></c>     <!-- "温度" -->
    <c r="C2" t="s" s="1"><v>3</v></c>     <!-- "湿度" -->
    <c r="D2" s="1"/>                       <!-- 被C2:D2合并 -->
  </row>
  <row r="3">
    <c r="A3" t="s" s="1"><v>4</v></c>     <!-- "编号" -->
    <c r="B3" t="s" s="1"><v>5</v></c>     <!-- "名称" -->
    <c r="C3" t="s" s="1"><v>6</v></c>     <!-- "通道1" -->
    <c r="D3" t="s" s="1"><v>7</v></c>     <!-- "通道2" -->
  </row>
  <!-- 数据行... -->
</sheetData>

<mergeCells count="4">
  <mergeCell ref="A1:A2"/>    <!-- "设备信息" 垂直合并2行 -->
  <mergeCell ref="B1:D1"/>    <!-- "测量数据" 水平合并3列 -->
  <mergeCell ref="C2:D2"/>    <!-- "湿度" 水平合并2列 -->
</mergeCells>
```

### 12.5 大量小表格拼在一页

在同一个 sheet 中放置多个独立表格，用空行隔开，每个表格有自己的合并区域：

```
┌───────────────────────────┐
│      [表格1]  传感器数据   │
├───────────────────────────┤
│  (空行隔开)                │
├───────────────────────────┤
│      [表格2]  报警记录     │
├───────────────────────────┤
│  (空行隔开)                │
├───────────────────────────┤
│      [表格3]  统计摘要     │
└───────────────────────────┘
```

实现要点：
- 行号不连续，中间跳过几行（如表格1在1-5行，表格2在8-12行，6-7行留空）
- 每个表格的标题行各自合并
- 列宽如果不同，需要折中设置或使用区域内合并

### 12.6 条件格式（颜色标记异常值）

虽然嵌入式实现中不一定要用完整的条件格式（Conditional Formatting），但可以预先在代码中判断并指定样式：

```c
double value = sensor_read(i);
int style;
if (value > THRESHOLD_HIGH) {
    style = STYLE_WARNING;   // 红色背景
} else if (value < THRESHOLD_LOW) {
    style = STYLE_CAUTION;  // 黄色背景
} else {
    style = STYLE_NORMAL;   // 正常
}
ws_set_number(ws, row, col, value, style);
```

### 12.7 冻结首行 + 自动筛选

```xml
<sheetViews>
  <sheetView workbookViewId="0">
    <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
  </sheetView>
</sheetViews>

<!-- 在 sheetData 之后 -->
<autoFilter ref="A1:C100"/>
```

### 12.8 打印设置

```xml
<!-- 放在 mergeCells 之后 -->
<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" 
             header="0.3" footer="0.3"/>

<pageSetup orientation="landscape"           <!-- 横向 -->
           horizontalDpi="300"
           fitToWidth="1"                    <!-- 缩放到1页宽 -->
           fitToHeight="0"/>                 <!-- 高度不限制 -->

<printOptions horizontalCentered="1"/>       <!-- 水平居中打印 -->
```

---

## 13. 常见陷阱与排错指南

### 13.1 "文件已损坏"错误

| 原因 | 检查方法 | 修复 |
|------|---------|------|
| CRC32 计算错误 | 用 7-Zip 测试 ZIP 完整性 | 检查 CRC32 表是否正确、是否反转了最终值 |
| ZIP 结构不完整 | 解压查看是否有所有必需文件 | 确保有 `[Content_Types].xml`、`_rels/.rels` 等 |
| XML 格式错误 | 用浏览器打开 XML 看是否报错 | 检查标签是否闭合、特殊字符是否转义 |
| 缺少必需文件 | 对比标准 xlsx 解压结果 | 至少需要 5 个核心文件 |

### 13.2 "需要修复"提示但能打开

| 原因 | 修复 |
|------|------|
| `count` 属性值不匹配 | 确保 `<mergeCells count="N">` 中 N 等于实际的 `<mergeCell>` 数量 |
| `fills` 少于 3 项 | 必须至少有 none + gray125 + 一个自定义 = 3 项 |
| `fonts` 少于 1 项 | 至少要有 1 个默认字体 |
| `cellXfs` 索引超出 | 单元格的 `s` 属性值不能超过 `cellXfs` 的 `count - 1` |
| sharedStrings 的 count 不对 | `count` 可以大于 `uniqueCount`（表示总引用次数），但不能小于 |

### 13.3 合并单元格的边框不显示

**原因**：被合并的单元格没有设置样式。

**错误做法**：
```xml
<!-- B1 被合并但没有 s 属性 -->
<c r="B1"/>
```

**正确做法**：
```xml
<!-- B1 被合并但设置了与 A1 相同的边框样式 -->
<c r="B1" s="1"/>
```

> **核心规则**：合并区域中的**每一个**单元格都必须设置有边框的样式索引，否则 Excel 不会为合并后的区域绘制完整边框。

### 13.4 数字显示为文本

**原因**：单元格设置了 `t="s"`（字符串类型）但内容是数字。

**修复**：数字单元格不要设置 `t` 属性（默认就是数字 `n`），或者设 `t="n"`：
```xml
<!-- 正确：数字 -->
<c r="A1"><v>123.45</v></c>

<!-- 错误：把数字当字符串存了 -->
<c r="A1" t="s"><v>5</v></c>   <!-- 这会显示 sharedStrings[5] 的文本，不是数字 -->
```

### 13.5 中文乱码

**原因**：XML 声明的编码与实际编码不一致。

**修复**：
- XML 头必须声明 `encoding="UTF-8"`
- 实际写入文件时也必须用 UTF-8 编码
- 在 C 中直接用 UTF-8 字符串字面量即可（大部分编译器默认 UTF-8）
- 如果源码文件不是 UTF-8，需要转换

### 13.6 Excel 提示"内容有问题"但可恢复

通常是 XML 的小问题，比如：
- 属性值中有未转义的 `&`、`<`、`>`、`"`、`'`
- 自闭合标签写法不一致（`<c/>` vs `<c></c>`，两种都可以，但不能混用导致标签不匹配）
- 命名空间 URI 拼写错误（必须与标准完全一致）

---

## 14. 性能优化与内存预估

### 14.1 内存占用估算

假设生成一个 100 行 × 5 列的报表：

| 组成部分 | 估算大小 | 说明 |
|---------|---------|------|
| worksheet XML | ~15 KB | 500 个单元格，每个约 30 字节 |
| sharedStrings XML | ~2 KB | 约 50 个不同字符串 |
| styles XML | ~2 KB | 固定模板 |
| workbook + rels + content_types | ~1 KB | 固定模板 |
| ZIP 开销 | ~3 KB | 文件头和中央目录 |
| **总文件大小** | **~23 KB** | |
| 运行时 RAM（全量模式） | ~30 KB | 需缓存所有 XML |
| 运行时 RAM（流式模式） | ~4 KB | 只需当前行 + 元数据 |

### 14.2 优化建议

1. **减少 `sprintf` 调用**：对于重复的 XML 结构，用 `memcpy` + 替换占位符
2. **复用缓冲区**：所有 XML 生成共用一个大缓冲区
3. **预计算字符串索引**：在生成数据时先算好 sharedStrings 索引
4. **分块输出**：如果支持流式输出（如 UART/SPI），每生成一行就输出一行
5. **使用 Store 模式**：跳过压缩，用 CPU 时间换内存
6. **静态分配**：所有数组用静态大小，避免 `malloc`
7. **精简命名空间**：长 URI 存在 Flash 中，用指针引用

### 14.3 流式生成的伪代码

```c
// 流式生成模式：不需要缓存整个 XML
void generate_xlsx_stream(workbook_t *wb, write_fn_t output, void *ctx) {
    // Phase 1: 输出固定文件（Content_Types, rels, workbook等）
    // 这些文件内容固定且小，可以直接输出
    
    // Phase 2: 输出 styles.xml（固定模板）
    
    // Phase 3: 流式输出 sharedStrings.xml
    //   先输出 <sst ...> 开头
    //   遍历字符串表，逐个输出 <si><t>...</t></si>
    //   输出 </sst> 结尾
    //   同时计算 CRC32
    
    // Phase 4: 流式输出 sheet1.xml
    //   输出 <worksheet> 和 <sheetData> 开头
    //   逐行输出 <row> 和 <c>
    //   输出 </sheetData>
    //   输出 <mergeCells>
    //   输出 </worksheet>
    //   同时计算 CRC32
    
    // Phase 5: 输出中央目录和 EOCD
    //   此时有所有文件的偏移和CRC32
}
```

---

## 附录 A：OOXML 速查表

### A.1 必需文件清单

| 文件路径 | 作用 | 可省略 |
|---------|------|--------|
| `[Content_Types].xml` | 声明文件类型 | 否 |
| `_rels/.rels` | 根关系 | 否 |
| `xl/workbook.xml` | 工作簿 | 否 |
| `xl/_rels/workbook.xml.rels` | 工作簿关系 | 否 |
| `xl/worksheets/sheet1.xml` | 工作表数据 | 否 |
| `xl/sharedStrings.xml` | 共享字符串 | 是（用内联字符串替代） |
| `xl/styles.xml` | 样式 | 是（但所有 `s` 属性需设为0） |
| `docProps/app.xml` | 应用属性 | 是 |
| `docProps/core.xml` | 核心属性 | 是 |

### A.2 XML 元素速查

| 元素 | 作用 | 出现位置 |
|------|------|---------|
| `<sst>` | 共享字符串表根元素 | sharedStrings.xml |
| `<si>` | 一个字符串项 | sst 的子元素 |
| `<t>` | 文本内容 | si 的子元素 |
| `<worksheet>` | 工作表根元素 | sheetN.xml |
| `<sheetData>` | 数据区域 | worksheet 的子元素 |
| `<row>` | 一行 | sheetData 的子元素 |
| `<c>` | 一个单元格 | row 的子元素 |
| `<v>` | 单元格值 | c 的子元素 |
| `<f>` | 公式 | c 的子元素 |
| `<mergeCells>` | 合并单元格集合 | sheetData 之后 |
| `<mergeCell>` | 一个合并区域 | mergeCells 的子元素 |
| `<cols>` | 列定义集合 | sheetData 之前 |
| `<col>` | 列宽定义 | cols 的子元素 |
| `<styleSheet>` | 样式表根元素 | styles.xml |
| `<fonts>` | 字体集合 | styleSheet 的子元素 |
| `<fills>` | 填充集合 | styleSheet 的子元素 |
| `<borders>` | 边框集合 | styleSheet 的子元素 |
| `<cellXfs>` | 样式组合集合 | styleSheet 的子元素 |
| `<xf>` | 一个样式组合 | cellXfs 的子元素 |

### A.3 Content Type URI 速查

| 文件 | ContentType |
|------|-------------|
| workbook.xml | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml` |
| sheetN.xml | `application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml` |
| styles.xml | `application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml` |
| sharedStrings.xml | `application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml` |
| .rels | `application/vnd.openxmlformats-package.relationships+xml` |

### A.4 关系类型 URI 速查

| 关系 | Type URI |
|------|----------|
| 文档入口 | `.../officeDocument/2006/relationships/officeDocument` |
| 工作表 | `.../officeDocument/2006/relationships/worksheet` |
| 样式表 | `.../officeDocument/2006/relationships/styles` |
| 共享字符串 | `.../officeDocument/2006/relationships/sharedStrings` |
| 核心属性 | `.../package/2006/relationships/metadata/core-properties` |
| 扩展属性 | `.../officeDocument/2006/relationships/extended-properties` |

> **省略技巧**：URI 中的 `http://schemas.openxmlformats.org/...` 部分很长，但**必须完整**。可以将其定义为 C 宏：
> ```c
> #define NS_MAIN "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
> #define NS_R "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
> #define NS_PKG "http://schemas.openxmlformats.org/package/2006/relationships"
> #define NS_CT "http://schemas.openxmlformats.org/package/2006/content-types"
> ```

---

## 附录 B：单元格地址编码算法

### B.1 完整的地址转换实现

```c
#include <string.h>
#include <stdio.h>

// 列号(1-based) → 列字母
// 1→A, 26→Z, 27→AA, 52→AZ, 53→BA, 702→ZZ, 703→AAA
void col_to_letters(int col, char *buf, int buf_size) {
    if (col < 1 || buf_size < 2) {
        buf[0] = '?';
        buf[1] = '\0';
        return;
    }
    int i = 0;
    while (col > 0 && i < buf_size - 1) {
        col--;  // 1-based → 0-based for modulo
        buf[i++] = 'A' + (col % 26);
        col /= 26;
    }
    buf[i] = '\0';
    // 反转（因为我们是从低位写到高位的）
    for (int j = 0; j < i / 2; j++) {
        char tmp = buf[j];
        buf[j] = buf[i - 1 - j];
        buf[i - 1 - j] = tmp;
    }
}

// 列字母 → 列号
int letters_to_col(const char *letters) {
    int col = 0;
    while (*letters && *letters >= 'A' && *letters <= 'Z') {
        col = col * 26 + (*letters - 'A' + 1);
        letters++;
    }
    return col;
}

// 解析单元格地址，如 "AB23" → row=23, col=28
void parse_address(const char *addr, int *row, int *col) {
    *col = 0;
    *row = 0;
    int i = 0;
    // 解析列字母部分
    while (addr[i] >= 'A' && addr[i] <= 'Z') {
        *col = *col * 26 + (addr[i] - 'A' + 1);
        i++;
    }
    // 解析行号部分
    while (addr[i] >= '0' && addr[i] <= '9') {
        *row = *row * 10 + (addr[i] - '0');
        i++;
    }
}

// 测试
void test_address() {
    char buf[10];
    
    col_to_letters(1, buf, sizeof(buf));   // → "A"
    col_to_letters(26, buf, sizeof(buf));  // → "Z"
    col_to_letters(27, buf, sizeof(buf));  // → "AA"
    col_to_letters(52, buf, sizeof(buf));  // → "AZ"
    col_to_letters(53, buf, sizeof(buf));  // → "BA"
    col_to_letters(702, buf, sizeof(buf)); // → "ZZ"
    col_to_letters(703, buf, sizeof(buf)); // → "AAA"
    
    // 反向测试
    letters_to_col("A");    // → 1
    letters_to_col("Z");    // → 26
    letters_to_col("AA");   // → 27
    letters_to_col("ZZ");  // → 702
    
    // 地址解析
    int row, col;
    parse_address("B3", &row, &col);      // → row=3, col=2
    parse_address("AA100", &row, &col);   // → row=100, col=27
}
```

### B.2 地址编码的直觉理解

可以把列字母看作一种特殊的进制系统——**没有零的 26 进制**：

```
十进制         字母          理解方式
1              A             1 × 26^0 = 1
2              B             2 × 26^0 = 2
...
26             Z             26 × 26^0 = 26
27             AA            1 × 26^1 + 1 × 26^0 = 27
28             AB            1 × 26^1 + 2 × 26^0 = 28
...
52             AZ            1 × 26^1 + 26 × 26^0 = 52
53             BA            2 × 26^1 + 1 × 26^0 = 53
...
702            ZZ            26 × 26^1 + 26 × 26^0 = 702
703            AAA           1 × 26^2 + 1 × 26^1 + 1 × 26^0 = 703
```

与普通进制不同的是，这里的"数字"从 1 到 26（对应 A-Z），没有 0。这就是为什么转换时要先减 1：

```
普通进制：0, 1, 2, ..., 9, 10
Excel 列：A, B, C, ..., Z, AA

A 对应数字 1，但在模运算中我们需要 0
所以转换时：col - 1 → 取模 → 除法 → col - 1 → ...
```

---

## 总结

在嵌入式系统中实现 Excel 生成库，核心知识点可以浓缩为以下几句话：

1. **xlsx = ZIP(无压缩也行) + XML**，不需要任何特殊二进制格式知识
2. **5 个核心 XML 文件**就能生成合法的 xlsx：Content_Types、root rels、workbook、styles、sheet
3. **合并单元格**就是 `<mergeCell ref="左上:右下"/>`，被合并的格子要有样式但可以没值
4. **交错格式**就是不同行用不同的合并模式，在代码中维护一个合并列表即可
5. **样式系统**是"字体×填充×边框×对齐"的组合套餐，通过索引引用
6. **共享字符串表**能大幅减少重复文本的体积
7. **CRC32** 是 ZIP 中唯一需要计算的数学运算，用 1024 字节查找表搞定
8. **嵌入式优化**的核心是流式生成 + 静态分配 + Store 模式

理解了这些，你就能在资源受限的嵌入式系统上生成任何格式的 Excel 文件了。


---

# 补充篇：标准 Excel 库的完整能力清单

> 上一篇文档覆盖了 xlsx 生成的基础原理和核心代码。但如果要做一个**功能完备的标准 Excel 库**，以下所有内容都是不可或缺的。每一章对应一个独立功能模块，按重要性和常用程度排序。

---

## 目录（补充篇）

- [15. 日期时间系统：Excel 的"秘密历法"](#15-日期时间系统excel-的秘密历法)
- [16. 数字精度与格式化陷阱](#16-数字精度与格式化陷阱)
- [17. 多工作表完整管理](#17-多工作表完整管理)
- [18. 公式系统：让单元格"会算术"](#18-公式系统让单元格会算术)
- [19. 数据验证：限制用户能填什么](#19-数据验证限制用户能填什么)
- [20. 条件格式：自动变色规则](#20-条件格式自动变色规则)
- [21. 图表：在表格里画图](#21-图表在表格里画图)
- [22. 图片插入：把图片嵌进单元格](#22-图片插入把图片嵌进单元格)
- [23. 超链接：可点击的链接](#23-超链接可点击的链接)
- [24. 单元格批注/注释](#24-单元格批注注释)
- [25. 自动筛选详解](#25-自动筛选详解)
- [26. 表格对象：不只是数据范围](#26-表格对象不只是数据范围)
- [27. 命名区域：给单元格起名字](#27-命名区域给单元格起名字)
- [28. 分组与大纲：折叠/展开行和列](#28-分组与大纲折叠展开行和列)
- [29. 工作表保护与工作簿保护](#29-工作表保护与工作簿保护)
- [30. 打印设置完整版](#30-打印设置完整版)
- [31. 页眉页脚](#31-页眉页脚)
- [32. 分页符](#32-分页符)
- [33. 工作表视图设置](#33-工作表视图设置)
- [34. 富文本：一个单元格多种格式](#34-富文本一个单元格多种格式)
- [35. 错误值](#35-错误值)
- [36. Sheet Tab 颜色](#36-sheet-tab-颜色)
- [37. 隐藏行/列与行列高级操作](#37-隐藏行列与行列高级操作)
- [38. 冻结窗格与拆分窗格完整版](#38-冻结窗格与拆分窗格完整版)
- [39. 列宽行高单位换算精解](#39-列宽行高单位换算精解)
- [40. 自定义数字格式语法大全](#40-自定义数字格式语法大全)
- [41. Sparklines：单元格里的迷你图](#41-sparklines单元格里的迷你图)
- [42. 外部链接](#42-外部链接)
- [43. 自定义文档属性](#43-自定义文档属性)
- [44. VBA 宏文件(.xlsm)](#44-vba-宏文件xlsm)
- [45. 文件加密与密码保护](#45-文件加密与密码保护)
- [46. Deflate 压缩可选实现](#46-deflate-压缩可选实现)
- [47. API 设计模式](#47-api-设计模式)
- [48. 线程安全与 RTOS 集成](#48-线程安全与-rtos-集成)
- [49. 错误处理框架](#49-错误处理框架)
- [50. 兼容性矩阵](#50-兼容性矩阵)
- [51. 测试与验证策略](#51-测试与验证策略)
- [52. 内存缓冲区管理进阶](#52-内存缓冲区管理进阶)
- [53. 数据类型完整参考](#53-数据类型完整参考)
- [54. 完整文件清单总表](#54-完整文件清单总表)

---

## 15. 日期时间系统：Excel 的"秘密历法"

### 15.1 为什么必须搞懂日期系统

在嵌入式系统中，传感器数据往往带时间戳。如果你把时间戳当作字符串写入 Excel，用户就没法按日期排序、筛选或做日期运算。正确做法是**把日期转为 Excel 的序列号**写入单元格，再用数字格式让它显示成日期样子。

### 15.2 Excel 日期序列号原理

Excel 内部**不存储日期字符串**，而是存储一个整数，代表"从 1900 年 1 月 1 日开始的天数"：

```
序列号 1  = 1900年1月1日
序列号 2  = 1900年1月2日
...
序列号 44927 = 2023年1月1日
序列号 45292 = 2024年1月1日
```

时间的小数部分表示一天中的比例：

```
序列号 44927.5  = 2023年1月1日 中午12:00
序列号 44927.25 = 2023年1月1日 早上6:00
序列号 44927.75 = 2023年1月1日 下午6:00
```

### 15.3 1900 闰年 Bug（必须处理！）

Excel 有一个著名的 bug：它**错误地认为 1900 年是闰年**（实际上 1900 年不是闰年，因为能被 100 整除但不能被 400 整除的年份不是闰年）。

这意味着 Excel 认为存在 "1900年2月29日" 这一天。结果是：

```
Excel 序列号 60 = 1900年2月29日（不存在的日期！）
```

因此，对于 1900 年 3 月 1 日及以后的所有日期，Excel 的序列号比实际天数**多 1**。

### 15.4 转换算法（C 语言）

```c
// 将年月日转为 Excel 序列号（1900 系统）
// 适用于 1900年3月1日之后的日期
uint32_t date_to_excel_serial(int year, int month, int day) {
    // 算法：基于 1900-01-01 = 1 的约定
    // 需要补偿 1900 闰年 bug
    
    if (year < 1900) return 0;  // 不支持1900年之前
    
    // 先计算从 1900-01-01 到 (year-01-01) 的天数
    int days = 0;
    for (int y = 1900; y < year; y++) {
        days += is_leap_year(y) ? 366 : 365;
    }
    
    // 加上当年已过天数
    static const int month_days[] = {0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
    for (int m = 1; m < month; m++) {
        days += month_days[m];
        if (m == 2 && is_leap_year(year)) {
            days += 1;  // 闰年2月多一天
        }
    }
    days += day;
    
    // 补偿 1900 闰年 bug
    // 1900年3月1日之前不需要补偿，之后需要+1
    if (year > 1900 || (year == 1900 && month >= 3)) {
        days += 1;
    }
    
    return (uint32_t)days;
}

// Excel 序列号转年月日
void excel_serial_to_date(uint32_t serial, int *year, int *month, int *day) {
    // 补偿 1900 闰年 bug
    if (serial > 59) serial -= 1;
    
    // 从 1900-01-01 = 1 开始反推
    int y = 1900;
    uint32_t remaining = serial;
    
    while (1) {
        int year_days = is_leap_year(y) ? 366 : 365;
        if (remaining > year_days) {
            remaining -= year_days;
            y++;
        } else {
            break;
        }
    }
    *year = y;
    
    // 逐月减
    static const int month_days[] = {0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
    int m = 1;
    while (m <= 12) {
        int md = month_days[m];
        if (m == 2 && is_leap_year(y)) md = 29;
        if (remaining > md) {
            remaining -= md;
            m++;
        } else {
            break;
        }
    }
    *month = m;
    *day = remaining;
}

// 闰年判断（正确的规则）
int is_leap_year(int year) {
    return (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0);
}
```

### 15.5 时间转小数

```c
// 时分秒 → 一天中的小数部分
// 14:30:00 → 0.604167
double time_to_fraction(int hour, int minute, int second) {
    return (hour * 3600.0 + minute * 60.0 + second) / 86400.0;
}

// 完整日期时间 → Excel 序列号
double datetime_to_excel(int year, int month, int day,
                          int hour, int minute, int second) {
    return (double)date_to_excel_serial(year, month, day) 
         + time_to_fraction(hour, minute, second);
}
```

### 15.6 在单元格中写入日期

日期本质上是一个浮点数，用 `t="n"`（数字类型）写入，但需要搭配日期格式的样式：

```xml
<!-- styles.xml 中定义日期格式 -->
<numFmts count="2">
  <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>           <!-- 仅日期 -->
  <numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm:ss"/>  <!-- 日期+时间 -->
</numFmts>

<!-- cellXfs 中定义日期样式 -->
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" 
    applyNumberFormat="1" applyBorder="1"/>

<!-- sheet1.xml 中写入日期 -->
<c r="A1" s="5"><v>44927</v></c>   <!-- 2023年1月1日，s=5 是日期格式的样式索引 -->
```

### 15.7 1904 日期系统（Mac 版）

Excel 还支持另一种日期系统——1904 系统，以 1904 年 1 月 1 日为起点。这主要影响旧版 Mac 版 Excel。在现代文件中很少使用，但需要知道它的存在：

```xml
<!-- workbook.xml 中声明使用 1904 系统 -->
<workbookPr date1904="1"/>
```

> **嵌入式建议**：统一使用 1900 系统（默认），不用管 1904。但如果你的文件要在老 Mac 上打开，需要注意。

### 15.8 日期格式的 formatCode 速查

| formatCode | 显示效果 | 说明 |
|------------|---------|------|
| `yyyy-mm-dd` | 2024-03-15 | ISO 标准 |
| `yyyy/m/d` | 2024/3/15 | 常见中文格式 |
| `mm/dd/yyyy` | 03/15/2024 | 美式格式 |
| `yyyy"年"m"月"d"日"` | 2024年3月15日 | 中文正式格式 |
| `yyyy-mm-dd hh:mm:ss` | 2024-03-15 14:30:00 | 完整日期时间 |
| `hh:mm:ss` | 14:30:00 | 仅时间 |
| `hh:mm` | 14:30 | 时分 |
| `yyyy-mm-dd hh:mm` | 2024-03-15 14:30 | 日期+时分 |
| `dddd` | Friday | 星期全称 |
| `ddd` | Fri | 星期缩写 |
| `yyyy"第"q"季度"` | 2024第1季度 | 季度（需公式配合） |
| `yyyy-mm-dd;@` | 2024-03-15 | 带文本占位符 |
| `AM/PM hh:mm` | PM 02:30 | 12小时制 |

---

## 16. 数字精度与格式化陷阱

### 16.1 IEEE 754 浮点数的精度限制

Excel 内部使用 IEEE 754 双精度浮点数（64 位 double）存储所有数值。这意味着：

```
精度：约 15-17 位有效数字
范围：约 ±1.7×10^308 到 ±5×10^-324
```

超出 15 位有效数字的数字会被截断：

```
实际输入：12345678901234567
Excel 存储：12345678901234500  ← 后几位丢失
```

### 16.2 嵌入式系统中的浮点问题

很多嵌入式 MCU 没有硬件浮点单元（FPU），软件模拟浮点既慢又占代码空间。解决方案：

**方案 A：整数模拟定点数**

```c
// 用整数存储"放大后的值"，如温度精确到 0.01°C
// 25.37°C 存储为整数 2537
// 写入 Excel 时除以 100

int32_t temp_raw = 2537;  // 传感器读数
double temp_value = temp_raw / 100.0;

// 生成 XML 时用 %.2f 格式化
fprintf(f, "<v>%.2f</v>", temp_value);  // 输出 "25.37"
```

**方案 B：直接用整数，让 Excel 的数字格式处理小数**

```xml
<!-- styles.xml: 定义"除以100"的数字格式 -->
<numFmt numFmtId="164" formatCode="0.00"/>

<!-- sheet1.xml: 写入原始整数 -->
<c r="A1" s="5"><v>2537</v></c>

<!-- Excel 显示为 "25.37" -->
```

等等，这个方法不对。Excel 的数字格式只是显示层面的转换，`0.00` 不会自动除以 100。如果需要缩放，要用公式或者直接写入正确的浮点值。

**方案 C：使用 `%g` 或手工格式化避免浮点库**

```c
// 手动将定点数转为字符串，避免依赖 sprintf 的 %f
// 适用于无浮点库的极端环境

void fixed_to_string(int32_t value, int scale, char *buf, int buf_size) {
    // value = 原始整数, scale = 小数位数
    // 例: value=2537, scale=2 → "25.37"
    
    int32_t int_part = value / scale_power(scale);  // 25
    int32_t frac_part = value % scale_power(scale);  // 37
    
    if (frac_part < 0) {
        frac_part = -frac_part;
        if (int_part == 0) {
            snprintf(buf, buf_size, "-0.%0*d", scale, frac_part);
        } else {
            snprintf(buf, buf_size, "-%d.%0*d", int_part, scale, frac_part);
        }
    } else {
        snprintf(buf, buf_size, "%d.%0*d", int_part, scale, frac_part);
    }
}

// 辅助：计算 10^scale
int32_t scale_power(int scale) {
    int32_t p = 1;
    while (scale-- > 0) p *= 10;
    return p;
}
```

### 16.3 精度丢失的常见场景

| 场景 | 问题 | 对策 |
|------|------|------|
| 大数加小数 | 1e15 + 1 = 1e15（小数被丢失） | 拆分计算或用整数 |
| 小数减法 | 0.1 + 0.2 = 0.30000000000000004 | 用 `%.15g` 格式化或四舍五入 |
| 日期计算 | 跨闰年天数可能有 ±1 误差 | 使用整数天数，不用浮点 |
| 货币计算 | 浮点精度不足 | 用整数"分"存储 |
| 科学计数法 | 大数自动变成科学计数法 | 用数字格式 `0` 强制普通显示 |

### 16.4 sprintf 的 %g vs %f vs %.15g

```c
double val = 25.37;

// %.2f — 固定2位小数，适合已知精度
snprintf(buf, sizeof(buf), "%.2f", val);  // "25.37"

// %g — 自动选择最短表示，适合通用
snprintf(buf, sizeof(buf), "%g", val);     // "25.37"

// %.15g — 保留15位有效数字，最大精度
snprintf(buf, sizeof(buf), "%.15g", val);  // "25.37"

// 整数直接用 %d
snprintf(buf, sizeof(buf), "%d", 1234);    // "1234"
```

> **推荐**：嵌入式系统中用 `%.15g` 格式化浮点数，能兼顾精度和简洁。整数直接用 `%d`。

### 16.5 特殊数字值

| 值 | XML 中的写法 | 说明 |
|----|------------|------|
| 正零 | `<v>0</v>` | |
| 负零 | `<v>-0</v>` | Excel 支持，显示为 0 |
| 正无穷 | 不支持直接写入 | Excel 显示为 `#NUM!` |
| 负无穷 | 不支持直接写入 | Excel 显示为 `#NUM!` |
| NaN | 不支持直接写入 | Excel 显示为 `#NUM!` |
| 极大数 | `<v>1e308</v>` | 接近 double 上限 |
| 极小数 | `<v>1e-308</v>` | 接近 double 下限 |

---

## 17. 多工作表完整管理

### 17.1 多 Sheet 的文件结构

一个工作簿可以有多个 Sheet，每个 Sheet 是一个独立的 XML 文件：

```
xl/
├── workbook.xml                  ← 列出所有 Sheet
├── _rels/
│   └── workbook.xml.rels         ← Sheet → 文件名的映射
├── worksheets/
│   ├── sheet1.xml
│   ├── sheet2.xml
│   └── sheet3.xml
```

### 17.2 workbook.xml 中的多 Sheet 定义

```xml
<workbook xmlns="..." xmlns:r="...">
  <sheets>
    <sheet name="数据" sheetId="1" r:id="rId1"/>
    <sheet name="图表" sheetId="2" r:id="rId2" state="hidden"/>
    <sheet name="汇总" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>
```

| 属性 | 说明 |
|------|------|
| `name` | 显示名称（最多 31 字符，不能包含 `\/?*[]:`） |
| `sheetId` | 内部 ID（正整数，不必连续） |
| `r:id` | 关系 ID，指向 workbook.xml.rels |
| `state` | `visible`（默认）、`hidden`、`veryHidden` |

### 17.3 Sheet 可见性

| state 值 | 效果 | 用户能否取消隐藏 |
|---------|------|---------------|
| `visible` | 正常显示 | — |
| `hidden` | 隐藏，但在右键菜单中可以看到 | 可以 |
| `veryHidden` | 深度隐藏，右键菜单看不到 | 不能（需要 VBA 或修改 XML） |

### 17.4 活动工作表

指定文件打开时哪个 Sheet 显示在最前面：

```xml
<workbook xmlns="...">
  <sheets>...</sheets>
  
  <!-- activeTab 的值是 sheets 列表中的索引（从0开始） -->
  <bookViews>
    <workbookView activeTab="2"/>   <!-- 第3个Sheet（索引2）为活动工作表 -->
  </bookViews>
</workbook>
```

### 17.5 Sheet 排序

Sheet 的显示顺序由 `<sheets>` 中 `<sheet>` 的排列顺序决定，与 `sheetId` 无关：

```xml
<!-- sheetId 乱序，但显示顺序是：数据→汇总→图表 -->
<sheets>
  <sheet name="数据" sheetId="1" r:id="rId1"/>
  <sheet name="汇总" sheetId="5" r:id="rId3"/>
  <sheet name="图表" sheetId="3" r:id="rId2"/>
</sheets>
```

### 17.6 多 Sheet 的 workbook.xml.rels

```xml
<Relationships xmlns="...">
  <Relationship Id="rId1" Type=".../worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type=".../worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type=".../worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type=".../styles" Target="styles.xml"/>
  <Relationship Id="rId5" Type=".../sharedStrings" Target="sharedStrings.xml"/>
</Relationships>
```

### 17.7 多 Sheet 的 Content_Types

每个 Sheet 都需要在 `[Content_Types].xml` 中声明：

```xml
<Types xmlns="...">
  <Default Extension="rels" ContentType="..."/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="...sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="...worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="...worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="...worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="...styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="...sharedStrings+xml"/>
</Types>
```

### 17.8 共享字符串和样式的跨 Sheet 共享

`sharedStrings.xml` 和 `styles.xml` 是**全工作簿共享**的。所有 Sheet 引用同一张字符串表和样式表。这意味着：

- Sheet1 中 "温度" 是索引 3，Sheet2 中引用 "温度" 也是索引 3
- 所有样式定义集中在一个 `styles.xml` 中

### 17.9 Sheet 名称的命名规则

| 规则 | 说明 |
|------|------|
| 最大长度 | 31 个字符 |
| 禁止字符 | `\` `/` `?` `*` `[` `]` `:` |
| 不能同名 | 同一工作簿内不能有两个同名的 Sheet |
| 不能为空 | 至少 1 个字符 |
| 大小写不敏感 | "Sheet1" 和 "SHEET1" 视为同名 |

---

## 18. 公式系统：让单元格"会算术"

### 18.1 公式的 XML 结构

公式存储在 `<c>` 元素的 `<f>` 子元素中，计算结果存在 `<v>` 中：

```xml
<!-- 简单公式 -->
<c r="C1"><f>SUM(A1:A10)</f><v>550</v></c>

<!-- 带数字格式的公式 -->
<c r="C2" s="5"><f>AVERAGE(B1:B10)</f><v>55.5</v></c>
```

| 子元素 | 说明 |
|--------|------|
| `<f>` | 公式文本（不含等号 `=`） |
| `<v>` | 公式计算后的结果值 |

> **关键**：OOXML 规范要求同时提供 `<f>`（公式）和 `<v>`（结果）。Excel 打开文件时会重新计算，但如果 `<v>` 缺失，某些阅读器会显示 0 或空白。

### 18.2 公式中的单元格引用

```
A1     相对引用
$A$1   绝对引用（行列都固定）
$A1    混合引用（列固定，行相对）
A$1    混合引用（行固定，列相对）
Sheet2!A1   跨工作表引用
Sheet1:Sheet3!A1  3D引用（跨多个Sheet的同一位置）
```

### 18.3 共享公式(Shared Formula)

当多个单元格用相似的公式时，可以只写一次公式，其他单元格引用它：

```xml
<!-- 主公式单元格 -->
<c r="C1"><f t="shared" ref="C1:C10">A1*B1</f><v>100</v></c>

<!-- 引用主公式的单元格 -->
<c r="C2"><f t="shared" si="0"/><v>200</v></c>
<c r="C3"><f t="shared" si="0"/><v>300</v></c>
```

- `t="shared"`：标识这是共享公式
- `ref`：共享公式覆盖的范围
- `si`：共享公式索引（从 0 开始）
- 引用单元格的 `<f>` 只有 `t="shared"` 和 `si`，没有公式文本

### 18.4 数组公式(Array Formula)

```xml
<c r="C1">
  <f t="array" ref="C1:C3">A1:A3*B1:B3</f>
  <v>100</v>
</c>
<c r="C2"><v>200</v></c>
<c r="C3"><v>300</v></c>
```

### 18.5 嵌入式系统中的公式策略

在嵌入式环境中，通常**无法执行 Excel 公式计算引擎**（太复杂）。因此有两种策略：

**策略 A：只写公式，不写结果值**

```xml
<c r="C1"><f>SUM(A1:A10)</f></c>
```

Excel 打开时会自动计算。但如果用 WPS 或 LibreOffice 打开，可能需要手动触发重算。

**策略 B：预计算结果，同时写公式和值（推荐）**

```c
// 嵌入式系统中先算好结果
double sum = 0;
for (int i = 0; i < 10; i++) sum += values[i];

// 写入公式和预计算结果
fprintf(f, "<c r=\"C1\"><f>SUM(A1:A10)</f><v>%.15g</v></c>", sum);
```

**策略 C：强制 Excel 打开时重算**

```xml
<!-- workbook.xml 中添加 -->
<calcPr fullCalcOnLoad="1"/>
```

这样 Excel 打开文件时会强制重新计算所有公式，确保结果正确。

### 18.6 常用公式速查

| 公式 | 说明 | 示例 |
|------|------|------|
| `SUM(range)` | 求和 | `SUM(A1:A10)` |
| `AVERAGE(range)` | 平均值 | `AVERAGE(B1:B5)` |
| `MAX(range)` | 最大值 | `MAX(A1:Z1)` |
| `MIN(range)` | 最小值 | `MIN(A1:Z1)` |
| `COUNT(range)` | 计数 | `COUNT(A1:A100)` |
| `IF(cond,a,b)` | 条件 | `IF(A1>100,"高","低")` |
| `ROUND(num,d)` | 四舍五入 | `ROUND(A1,2)` |
| `CONCATENATE(a,b)` | 字符串拼接 | `CONCATENATE(A1,"-",B1)` |
| `VLOOKUP(v,r,c,f)` | 垂直查找 | `VLOOKUP(A1,B1:D10,2,FALSE)` |
| `TEXT(v,fmt)` | 数字转文本 | `TEXT(A1,"0.00")` |
| `NOW()` | 当前日期时间 | `NOW()` |
| `TODAY()` | 当前日期 | `TODAY()` |
| `ABS(num)` | 绝对值 | `ABS(A1-100)` |
| `SQRT(num)` | 平方根 | `SQRT(A1)` |
| `POWER(b,e)` | 幂运算 | `POWER(A1,2)` |

---

## 19. 数据验证：限制用户能填什么

### 19.1 什么是数据验证

数据验证就是给单元格设置"规则"，用户填入不符合规则的数据时弹出警告。常见场景：

- 下拉列表选择
- 只能填数字范围
- 只能填日期范围
- 限制文本长度

### 19.2 XML 结构

数据验证放在 `<sheetData>` 和 `<mergeCells>` 之后：

```xml
<worksheet xmlns="...">
  <sheetData>...</sheetData>
  <mergeCells count="...">...</mergeCells>
  
  <!-- 数据验证 -->
  <dataValidations count="2">
    <!-- 下拉列表 -->
    <dataValidation type="list" allowBlank="1" showInputMessage="1" 
                    showErrorMessage="1" sqref="B2:B10">
      <formula1>"正常,警告,故障,维修"</formula1>
      <promptTitle="请选择状态" prompt="从下拉列表中选择设备状态"/>
      <errorTitle="输入错误" error="请从列表中选择有效状态"/>
    </dataValidation>
    
    <!-- 数字范围验证 -->
    <dataValidation type="decimal" operator="between" allowBlank="1"
                    showErrorMessage="1" sqref="C2:C100">
      <formula1>0</formula1>
      <formula2>100</formula2>
      <errorTitle="超出范围" error="温度必须在 0 到 100 之间"/>
    </dataValidation>
  </dataValidations>
</worksheet>
```

### 19.3 验证类型(type)完整列表

| type 值 | 说明 | formula1 | formula2 |
|---------|------|----------|----------|
| `none` | 无验证 | — | — |
| `whole` | 整数 | 最小值或值列表 | 最大值(范围时) |
| `decimal` | 小数 | 最小值 | 最大值(范围时) |
| `list` | 列表选择 | `"a,b,c"` 或 `=$D$1:$D$5` | — |
| `date` | 日期 | 最小日期 | 最大日期(范围时) |
| `time` | 时间 | 最小时间 | 最大时间(范围时) |
| `textLength` | 文本长度 | 最小长度 | 最大长度(范围时) |
| `custom` | 自定义公式 | 公式（返回 TRUE/FALSE） | — |

### 19.4 操作符(operator)列表

| operator 值 | 说明 | 需要 formula2 |
|------------|------|-------------|
| `between` | 在...之间 | 是 |
| `notBetween` | 不在...之间 | 是 |
| `equal` | 等于 | 否 |
| `notEqual` | 不等于 | 否 |
| `greaterThan` | 大于 | 否 |
| `lessThan` | 小于 | 否 |
| `greaterThanOrEqual` | 大于等于 | 否 |
| `lessThanOrEqual` | 小于等于 | 否 |

### 19.5 下拉列表的两种实现方式

**方式 A：内联列表（简单，适合少量选项）**

```xml
<formula1>"待机,运行,停机,故障"</formula1>
```

注意：整个列表用双引号包裹，选项间用逗号分隔。如果选项本身包含逗号，会有问题。

**方式 B：引用单元格区域（适合大量选项或动态选项）**

```xml
<!-- 引用同工作表的区域 -->
<formula1>$F$1:$F$10</formula1>

<!-- 引用命名区域 -->
<formula1>=状态列表</formula1>

<!-- 引用其他工作表 -->
<formula1>Sheet2!$A$1:$A$20</formula1>
```

### 19.6 错误提示样式

| errorStyle 值 | 效果 |
|--------------|------|
| `stop`（默认） | 阻止输入，必须改 |
| `warning` | 警告，用户可选择"是"继续或"否"修改 |
| `information` | 提示信息，用户可选择"确定"继续或"取消" |

```xml
<dataValidation type="decimal" operator="between" 
                errorStyle="warning"   <!-- 警告而非阻止 -->
                sqref="C2:C100">
  <formula1>0</formula1>
  <formula2>100</formula2>
</dataValidation>
```

### 19.7 输入提示

当用户选中设置了验证的单元格时，显示帮助信息：

```xml
<dataValidation type="list" 
                showInputMessage="1"
                showErrorMessage="1"
                sqref="B2:B10">
  <formula1>"高,中,低"</formula1>
  <promptTitle="优先级"/>
  <prompt="请选择任务的优先级：高、中、低"/>
</dataValidation>
```

---

## 20. 条件格式：自动变色规则

### 20.1 什么是条件格式

条件格式让单元格根据其值自动变色，如：
- 温度高于 50°C 显示红色背景
- 数值排名前 10% 加粗
- 用数据条直观显示大小
- 用色阶表示渐变

### 20.2 XML 结构

```xml
<worksheet xmlns="...">
  <sheetData>...</sheetData>
  
  <!-- 条件格式 -->
  <conditionalFormatting sqref="B2:B100">
    <!-- 规则1：大于80的标红 -->
    <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
      <formula>80</formula>
    </cfRule>
    <!-- 规则2：小于20的标蓝 -->
    <cfRule type="cellIs" dxfId="1" priority="2" operator="lessThan">
      <formula>20</formula>
    </cfRule>
  </conditionalFormatting>
</worksheet>
```

- `sqref`：条件格式应用的范围
- `priority`：规则优先级（数字越小优先级越高）
- `dxfId`：引用 styles.xml 中的差异格式(Differential Format)索引

### 20.3 差异格式(dxf)定义

差异格式定义在 styles.xml 中，与普通样式不同，它只包含"变化的部分"：

```xml
<styleSheet xmlns="...">
  <!-- ... fonts, fills, borders, cellXfs ... -->
  
  <!-- 差异格式 -->
  <dxfs count="3">
    <!-- dxfId=0：红色背景 -->
    <dxf>
      <fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill>
      <font><color rgb="FF9C0006"/><b/></font>
    </dxf>
    <!-- dxfId=1：蓝色背景 -->
    <dxf>
      <fill><patternFill><bgColor rgb="FFFFEB9C"/></patternFill></fill>
      <font><color rgb="FF9C6500"/></font>
    </dxf>
    <!-- dxfId=2：绿色背景 -->
    <dxf>
      <fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill>
      <font><color rgb="FF006100"/></font>
    </dxf>
  </dxfs>
</styleSheet>
```

> **注意**：差异格式中的 `fill` 用 `bgColor`（背景色）而不是 `fgColor`（前景色），这与普通样式相反！

### 20.4 规则类型(type)完整列表

| type | 说明 | 示例 |
|------|------|------|
| `cellIs` | 单元格值比较 | 大于/小于/等于某值 |
| `expression` | 公式判断 | 公式返回 TRUE 时应用 |
| `colorScale` | 色阶 | 三色渐变 |
| `dataBar` | 数据条 | 在单元格内画进度条 |
| `iconSet` | 图标集 | 交通灯、箭头等 |
| `top10` | 前/后 N 名 | 前 10 名标红 |
| `aboveAverage` | 高于/低于平均值 | 高于平均值的标绿 |
| `duplicateValues` | 重复值 | 重复值高亮 |
| `uniqueValues` | 唯一值 | 唯一值高亮 |
| `containsText` | 包含文本 | 包含"故障"的标红 |
| `notContainsText` | 不包含文本 | |
| `beginsWith` | 以...开头 | |
| `endsWith` | 以...结尾 | |
| `containsErrors` | 包含错误 | 有错误的标红 |
| `containsBlanks` | 包含空格 | |
| `timePeriod` | 时间段 | 今天/昨天/上周等 |

### 20.5 常用规则示例

**单元格值比较：**

```xml
<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
  <formula>80</formula>
</cfRule>
<cfRule type="cellIs" dxfId="1" priority="2" operator="between">
  <formula>50</formula>
  <formula>80</formula>
</cfRule>
<cfRule type="cellIs" dxfId="2" priority="3" operator="lessThan">
  <formula>20</formula>
</cfRule>
```

**公式判断：**

```xml
<!-- 偶数行标灰 -->
<cfRule type="expression" dxfId="0" priority="1">
  <formula>MOD(ROW(),2)=0</formula>
</cfRule>

<!-- 如果A列值大于100，整行标红 -->
<cfRule type="expression" dxfId="1" priority="2">
  <formula>$A1>100</formula>
</cfRule>
```

**色阶：**

```xml
<cfRule type="colorScale" priority="1">
  <colorScale>
    <cfvo type="min"/>                    <!-- 最小值 -->
    <cfvo type="percentile" val="50"/>    <!-- 50百分位 -->
    <cfvo type="max"/>                    <!-- 最大值 -->
    <color rgb="FFF8696B"/>                <!-- 红 -->
    <color rgb="FFFFEB84"/>               <!-- 黄 -->
    <color rgb="FF63BE7B"/>               <!-- 绿 -->
  </colorScale>
</cfRule>
```

**数据条：**

```xml
<cfRule type="dataBar" priority="1">
  <dataBar>
    <cfvo type="min"/>
    <cfvo type="max"/>
    <color rgb="FF638EC6"/>               <!-- 数据条颜色 -->
  </dataBar>
</cfRule>
```

**图标集：**

```xml
<cfRule type="iconSet" priority="1">
  <iconSet iconSet="3TrafficLights1">
    <cfvo type="percent" val="0"/>
    <cfvo type="percent" val="33"/>
    <cfvo type="percent" val="67"/>
  </iconSet>
</cfRule>
```

可用图标集：`3TrafficLights1`、`3TrafficLights2`、`3Arrows`、`3ArrowsGray`、`4Arrows`、`4ArrowsGray`、`5Arrows`、`5ArrowsGray`、`3Flags`、`3Symbols`、`3Symbols2`、`4Rating`、`5Rating`、`5Quarters`

**前 N 名：**

```xml
<!-- 前 5 名标红 -->
<cfRule type="top10" dxfId="0" priority="1" rank="5" percent="0" bottom="0">
</cfRule>

<!-- 后 10% 标蓝 -->
<cfRule type="top10" dxfId="1" priority="2" rank="10" percent="1" bottom="1">
</cfRule>
```

**包含特定文本：**

```xml
<cfRule type="containsText" dxfId="0" priority="1" operator="containsText" text="故障">
  <formula>NOT(ISERROR(SEARCH("故障",B2)))</formula>
</cfRule>
```

---

## 21. 图表：在表格里画图

### 21.1 图表的文件结构

图表需要额外的文件：

```
xl/
├── charts/
│   ├── chart1.xml          ← 图表定义
│   ├── chart2.xml
│   └── colors1.xml         ← 可选：颜色方案
├── drawings/
│   └── drawing1.xml        ← 图表在工作表中的位置
├── worksheets/
│   └── sheet1.xml
└── _rels/
    └── sheet1.xml.rels     ← sheet → drawing 的关系
```

### 21.2 关系链

```
sheet1.xml
  → (通过 sheet1.xml.rels) → drawing1.xml
    → (通过 drawing1.xml.rels) → chart1.xml
```

### 21.3 在 Sheet 中引用图表

sheet1.xml 的末尾添加：
```xml
<!-- 放在 mergeCells 之后，conditionalFormatting 之前 -->
<drawing r:id="rId10"/>
```

sheet1.xml.rels:
```xml
<Relationship Id="rId10" Type=".../drawing" Target="../drawings/drawing1.xml"/>
```

### 21.4 drawing1.xml：图表位置

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from>
      <xdr:col>3</xdr:col>          <!-- 起始列 (0-based) -->
      <xdr:colOff>0</xdr:colOff>     <!-- 列内偏移 (EMUs) -->
      <xdr:row>0</xdr:row>           <!-- 起始行 (0-based) -->
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>10</xdr:col>          <!-- 结束列 -->
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>20</xdr:row>          <!-- 结束行 -->
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" 
                   xmlns:r="..." r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
```

### 21.5 chart1.xml：柱状图示例

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartspace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  
  <c:chart>
    <!-- 标题 -->
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r><a:t>温度趋势</a:t></a:r>
          </a:p>
        </c:rich>
      </c:tx>
      <c:overlay val="0"/>
    </c:title>
    
    <!-- 自动调整绘图区 -->
    <c:plotArea>
      <c:layout/>
      
      <!-- 柱状图(barChart) -->
      <c:barChart>
        <c:barDir val="col"/>              <!-- col=竖柱, bar=横条 -->
        <c:grouping val="clustered"/>       <!-- clustered=簇状, stacked=堆积 -->
        
        <!-- 数据系列 -->
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx>
            <c:strRef>
              <c:f>Sheet1!$B$1</c:f>       <!-- 系列名：引用B1单元格 -->
              <c:strCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>温度(°C)</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          
          <!-- 类别（X轴标签） -->
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A$6</c:f>
              <c:strCache>
                <c:ptCount val="5"/>
                <c:pt idx="0"><c:v>周一</c:v></c:pt>
                <c:pt idx="1"><c:v>周二</c:v></c:pt>
                <c:pt idx="2"><c:v>周三</c:v></c:pt>
                <c:pt idx="3"><c:v>周四</c:v></c:pt>
                <c:pt idx="4"><c:v>周五</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          
          <!-- 数据值 -->
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$6</c:f>
              <c:numCache>
                <c:formatCode>0.0</c:formatCode>
                <c:ptCount val="5"/>
                <c:pt idx="0"><c:v>25.3</c:v></c:pt>
                <c:pt idx="1"><c:v>26.1</c:v></c:pt>
                <c:pt idx="2"><c:v>24.8</c:v></c:pt>
                <c:pt idx="3"><c:v>28.2</c:v></c:pt>
                <c:pt idx="4"><c:v>27.5</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      
      <!-- X轴(类别轴) -->
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>                 <!-- b=bottom -->
        <c:crossAx val="2"/>
      </c:catAx>
      
      <!-- Y轴(值轴) -->
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>                 <!-- l=left -->
        <c:crossAx val="1"/>
      </c:valAx>
    </c:plotArea>
    
    <!-- 图例 -->
    <c:legend>
      <c:legendPos val="r"/>               <!-- r=right -->
    </c:legend>
  </c:chart>
</c:chartspace>
```

### 21.6 图表类型对照表

| XML 元素 | 图表类型 | 说明 |
|---------|---------|------|
| `<c:barChart>` | 柱状图/条形图 | `barDir="col"`=竖柱, `"bar"`=横条 |
| `<c:lineChart>` | 折线图 | 趋势变化 |
| `<c:pieChart>` | 饼图 | 占比分布 |
| `<c:scatterChart>` | 散点图 | XY 坐标关系 |
| `<c:areaChart>` | 面积图 | 累积趋势 |
| `<c:doughnutChart>` | 圆环图 | 饼图中间有洞 |
| `<c:radarChart>` | 雷达图 | 多维对比 |
| `<c:bubbleChart>` | 气泡图 | 三维数据 |
| `<c:stockChart>` | K线图 | 开盘/最高/最低/收盘 |
| `<c:surfaceChart>` | 曲面图 | 三维表面 |

### 21.7 嵌入式实现要点

| 要点 | 说明 |
|------|------|
| **数据缓存** | chart1.xml 中的 `<c:strCache>` 和 `<c:numCache>` 是数据缓存，即使源数据变了，图表仍显示缓存值。嵌入式中直接写缓存即可 |
| **不需要公式引擎** | `<c:f>` 中写引用地址，但实际数据在 `<c:strCache>`/`<c:numCache>` 中，所以不需要计算公式 |
| **EMU 单位** | 偏移量用 EMU（English Metric Unit），1 英寸 = 914400 EMU，1 厘米 = 360000 EMU |
| **图表复杂度** | 最简图表约 50 行 XML，复杂图表可达数百行 |
| **内存占用** | 单个图表约 2-5 KB，需预算内存 |

---

## 22. 图片插入：把图片嵌进单元格

### 22.1 文件结构

```
xl/
├── media/
│   └── image1.png              ← 实际图片文件（PNG/JPEG/GIF等）
├── drawings/
│   └── drawing1.xml            ← 图片位置定义
├── worksheets/
│   └── sheet1.xml
└── _rels/
    └── sheet1.xml.rels         ← sheet → drawing 关系
    
xl/drawings/
└── _rels/
    └── drawing1.xml.rels       ← drawing → image 关系
```

### 22.2 关系链

```
sheet1.xml → sheet1.xml.rels → drawing1.xml → drawing1.xml.rels → media/image1.png
```

### 22.3 Content_Types 声明

需要在 `[Content_Types].xml` 中声明图片类型：

```xml
<Default Extension="png" ContentType="image/png"/>
<Default Extension="jpeg" ContentType="image/jpeg"/>
<Default Extension="jpg" ContentType="image/jpeg"/>
<Default Extension="gif" ContentType="image/gif"/>
<Default Extension="bmp" ContentType="image/bmp"/>
```

### 22.4 drawing1.xml：图片位置

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="..."
          xmlns:a="...">
  
  <!-- 方式1：锚定到两个单元格之间（推荐，随单元格移动） -->
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from>
      <xdr:col>0</xdr:col>        <!-- A列 -->
      <xdr:colOff>50000</xdr:colOff>
      <xdr:row>0</xdr:row>         <!-- 第1行 -->
      <xdr:rowOff>50000</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>2</xdr:col>         <!-- C列 -->
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>10</xdr:row>        <!-- 第11行 -->
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="图片1"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>   <!-- 引用 drawing1.xml.rels 中的 rId1 -->
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  
  <!-- 方式2：绝对位置（固定像素大小） -->
  <xdr:absoluteAnchor>
    <xdr:pos x="100000" y="50000"/>     <!-- 位置(EMU) -->
    <xdr:ext cx="2000000" cy="1500000"/> <!-- 大小(EMU) -->
    <xdr:pic>...</xdr:pic>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>
```

### 22.5 drawing1.xml.rels

```xml
<Relationships xmlns="...">
  <Relationship Id="rId1" 
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" 
                Target="../media/image1.png"/>
</Relationships>
```

### 22.6 锚定方式对比

| 锚定方式 | XML 元素 | 行为 |
|---------|---------|------|
| 两点锚定 | `twoCellAnchor` | 图片随两个单元格移动，`editAs` 控制调整行为 |
| 单点锚定 | `oneCellAnchor` | 图片随一个单元格移动，大小固定 |
| 绝对位置 | `absoluteAnchor` | 图片固定在指定位置，不随单元格移动 |

`editAs` 属性（仅 `twoCellAnchor`）：

| 值 | 行为 |
|----|------|
| `twoCell` | 图片随两个锚点移动和缩放 |
| `oneCell` | 图片随起始锚点移动，大小不变 |
| `absolute` | 图片不随单元格移动 |

### 22.7 EMU 单位换算

```
1 英寸 = 914400 EMU
1 厘米 = 360000 EMU
1 毫米 = 36000 EMU
1 像素(96dpi) = 9525 EMU
```

### 22.8 嵌入式实现要点

- 图片文件以原始格式（PNG/JPEG）直接放入 ZIP 包，不需要转码
- 嵌入式系统通常有摄像头或截图功能，生成的 PNG 可直接嵌入
- 注意 ZIP 的 Store 模式不压缩图片（图片本身已是压缩格式）
- 如果图片很大，考虑在嵌入前先缩小尺寸

---

## 23. 超链接：可点击的链接

### 23.1 URL 超链接

需要两个部分：单元格中的链接标记 + 关系文件中的目标 URL。

**sheet1.xml 中的单元格：**
```xml
<c r="A1" t="s" s="4">
  <v>0</v>
  <!-- 超链接放在 c 元素后面 -->
</c>
<!-- ... 其他单元格 ... -->

<!-- 超链接定义（在 sheetData 之后） -->
<hyperlinks>
  <hyperlink ref="A1" r:id="rId10"/>
</hyperlinks>
```

**sheet1.xml.rels 中的目标：**
```xml
<Relationship Id="rId10" 
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" 
              Target="https://www.example.com" 
              TargetMode="External"/>
```

> **注意**：`TargetMode="External"` 表示外部链接，没有这个属性会被当作内部链接。

### 23.2 带显示文本和提示的超链接

```xml
<hyperlinks>
  <hyperlink ref="B5" r:id="rId10" 
             display="点击访问" 
             tooltip="访问我们的官网"/>
</hyperlinks>
```

### 23.3 内部跳转（跳到其他 Sheet 或单元格）

不需要关系文件，直接用 `location` 属性：

```xml
<hyperlinks>
  <!-- 跳到 Sheet2 的 C5 单元格 -->
  <hyperlink ref="A1" location="Sheet2!C5" display="跳转到汇总表"/>
  
  <!-- 跳到 Sheet2 的 A1 并选中区域 -->
  <hyperlink ref="A2" location="Sheet2!A1:D10"/>
</hyperlinks>
```

### 23.4 邮件超链接

```xml
<Relationship Id="rId10" 
              Type=".../hyperlink" 
              Target="mailto:support@example.com?subject=技术支持&body=请描述您的问题" 
              TargetMode="External"/>
```

### 23.5 多个超链接

```xml
<hyperlinks>
  <hyperlink ref="A1" r:id="rId10"/>
  <hyperlink ref="A2" r:id="rId11" display="文档"/>
  <hyperlink ref="A3" location="Sheet2!A1"/>
  <hyperlink ref="B1:B5" r:id="rId12"/>   <!-- 区域超链接 -->
</hyperlinks>
```

---

## 24. 单元格批注/注释

### 24.1 文件结构

```
xl/
├── worksheets/
│   └── sheet1.xml
├── comments/
│   └── sheet1.xml          ← 批注内容
├── drawings/
│   └── vmlDrawing1.vml     ← 批注的显示位置（VML格式）
└── _rels/
    └── sheet1.xml.rels     ← sheet → comments 和 vmlDrawing 的关系
```

> **注意**：批注需要 VML（Vector Markup Language）文件来定义显示位置。VML 是一种旧格式，但在 OOXML 中仍用于批注定位。现代 Excel 也支持 `threadedComments`（线程批注）。

### 24.2 sheet1.xml 中的标记

```xml
<c r="A1" t="s" s="1">
  <v>0</v>
</c>

<!-- 在 sheetData 之后 -->
<legacyDrawing r:id="rId10"/>
```

### 24.3 comments/sheet1.xml：批注内容

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <comment ref="A1" authorId="0">
    <text>
      <r>
        <rPr>
          <sz val="9"/>
          <color rgb="FF000000"/>
          <rFont val="等线"/>
        </rPr>
        <t>这是批注内容。注意此数据为异常值。</t>
      </r>
    </text>
  </comment>
  <comment ref="B5" authorId="1">
    <text>
      <r><t>多行批注第一行。</t></r>
      <r><rPr><b/></rPr><t>第二行加粗。</t></r>
    </text>
  </comment>
</comments>
```

### 24.4 作者列表

```xml
<comments xmlns="...">
  <authors>
    <author>系统</author>
    <author>管理员</author>
  </authors>
  <comment ref="A1" authorId="0">...</comment>
  <comment ref="B5" authorId="1">...</comment>
</comments>
```

### 24.5 VML 文件（批注框位置）

VML 是一种 XML 格式的矢量标记语言，用于定义批注框的位置和大小：

```xml
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <o:shapelayout v:ext="edit">
    <o:idmap v:ext="edit" data="1"/>
  </o:shapelayout>
  <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m0,0l0,21600,21600,21600,21600,0xe">
    <v:stroke joinstyle="miter"/>
    <v:path gradientshapeok="t" o:connecttype="rect"/>
  </v:shapetype>
  <v:shape id="_x0000_s1025" type="#_x0000_t202" 
           style="position:absolute;margin-left:60pt;margin-top:15pt;width:120pt;height:60pt;z-index:1;visibility:hidden" 
           fillcolor="#ffffe1" o:insetmode="auto">
    <v:fill color2="#ffffe1"/>
    <v:shadow on="t" color="black" obscured="t"/>
    <v:path o:connecttype="none"/>
    <v:textbox style="mso-direction-alt:auto">
      <div style="text-align:left"/>
    </v:textbox>
    <x:ClientData ObjectType="Note">
      <x:MoveWithCells/>
      <x:SizeWithCells/>
      <x:Anchor>1, 0, 0, 0, 3, 0, 4, 0</x:Anchor>
      <x:AutoFill>False</x:AutoFill>
    </x:ClientData>
  </v:shape>
</xml>
```

### 24.6 嵌入式简化方案

VML 格式复杂，如果只需要批注功能（不含显示位置），可以：
1. 写 comments XML + 关系文件
2. VML 写一个固定模板，只修改锚点坐标
3. 或跳过 VML，Excel 仍能读取批注内容，只是不显示批注框

---

## 25. 自动筛选详解

### 25.1 基础自动筛选

```xml
<worksheet xmlns="...">
  <sheetData>...</sheetData>
  
  <!-- 自动筛选（放在 sheetData 之后） -->
  <autoFilter ref="A1:C100"/>
</worksheet>
```

这会在 A1:C100 的第一行显示下拉箭头，允许用户筛选。

### 25.2 带筛选条件的自动筛选

```xml
<autoFilter ref="A1:C100">
  <!-- 对A列（第一列）筛选 -->
  <filterColumn colId="0">
    <filters>
      <filter val="正常运行"/>
      <filter val="待机"/>
    </filters>
  </filterColumn>
  
  <!-- 对B列（第二列）进行自定义筛选 -->
  <filterColumn colId="1">
    <customFilters>
      <customFilter operator="greaterThan" val="50"/>
    </customFilters>
  </filterColumn>
  
  <!-- 对C列进行日期筛选 -->
  <filterColumn colId="2">
    <filters>
      <dateGroupItem year="2024" month="3"/>
      <dateGroupItem year="2024" month="4"/>
    </filters>
  </filterColumn>
</autoFilter>
```

### 25.3 筛选类型

| 子元素 | 说明 | 示例 |
|--------|------|------|
| `filters` | 值列表筛选 | 筛选包含特定值的行 |
| `customFilters` | 自定义条件 | 大于、小于、包含等 |
| `colorFilter` | 按颜色筛选 | 按单元格背景色 |
| `iconFilter` | 按图标筛选 | 按条件格式的图标 |
| `dynamicFilter` | 动态筛选 | 本周、上月、前10等 |
| `top10` | 前/后 N 名 | 前 5 大/小值 |

### 25.4 customFilters 的 AND/OR

```xml
<!-- AND：大于50且小于100 -->
<customFilters and="1">
  <customFilter operator="greaterThan" val="50"/>
  <customFilter operator="lessThan" val="100"/>
</customFilters>

<!-- OR：等于"高"或等于"紧急" -->
<customFilters and="0">
  <customFilter val="高"/>
  <customFilter val="紧急"/>
</customFilters>
```

### 25.5 排序状态

```xml
<autoFilter ref="A1:C100">
  <sortState>
    <!-- 按B列升序排序 -->
    <sortCondition ref="B1:B100" descending="0"/>
  </sortState>
</autoFilter>
```

---

## 26. 表格对象：不只是数据范围

### 26.1 表格对象 vs 普通数据范围

Excel 的"表格(Table)"是一种特殊对象，提供：
- 自动斑马纹
- 标题行筛选按钮
- 结构化引用公式
- 自动扩展公式列

### 26.2 文件结构

表格对象需要额外文件：

```
xl/
├── tables/
│   └── table1.xml          ← 表格定义
├── worksheets/
│   └── sheet1.xml
└── _rels/
    └── sheet1.xml.rels     ← sheet → table 关系
```

### 26.3 sheet1.xml 中的引用

```xml
<worksheet xmlns="...">
  <sheetData>...</sheetData>
  
  <!-- 表格对象引用 -->
  <tableParts count="1">
    <tablePart r:id="rId10"/>
  </tableParts>
</worksheet>
```

### 26.4 table1.xml：表格定义

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
       id="1" name="SensorTable" displayName="SensorTable" 
       ref="A1:D100" totalsRowCount="1">
  <!-- ref: 表格范围 -->
  <!-- totalsRowCount: 是否有汇总行 -->
  
  <tableColumns count="4">
    <tableColumn id="1" name="编号"/>
    <tableColumn id="2" name="传感器名"/>
    <tableColumn id="3" name="数值" 
                 dataDxfId="0"/>           <!-- 数据差异格式 -->
    <tableColumn id="4" name="单位" 
                 totalsRowLabel="平均"/>     <!-- 汇总行标签 -->
  </tableColumns>
  
  <!-- 表格样式 -->
  <tableStyleInfo name="TableStyleMedium2" 
                  showFirstColumn="0" 
                  showLastColumn="0" 
                  showRowStripes="1"          <!-- 斑马纹 -->
                  showColumnStripes="0"/>
</table>
```

### 26.5 表格样式名称

| 名称前缀 | 效果 |
|---------|------|
| `TableStyleLight1-28` | 浅色样式 |
| `TableStyleMedium1-28` | 中等样式 |
| `TableStyleDark1-11` | 深色样式 |

常用：`TableStyleMedium2`（蓝色中浅）、`TableStyleMedium9`（绿色）、`TableStyleLight1`（极简）

### 26.6 汇总行(Total Row)

```xml
<table ... totalsRowCount="1">
  <tableColumns count="3">
    <tableColumn id="1" name="名称" totalsRowLabel="合计"/>
    <tableColumn id="2" name="数量" totalsRowFunction="sum"/>     <!-- 求和 -->
    <tableColumn id="3" name="均价" totalsRowFunction="average"/>  <!-- 平均 -->
  </tableColumns>
</table>
```

汇总函数：

| totalsRowFunction | 说明 |
|-------------------|------|
| `sum` | 求和 |
| `average` | 平均值 |
| `count` | 计数 |
| `countNums` | 数字计数 |
| `max` | 最大值 |
| `min` | 最小值 |
| `stdDev` | 标准差 |
| `var` | 方差 |
| `custom` | 自定义公式 |

---

## 27. 命名区域：给单元格起名字

### 27.1 工作簿级命名区域

```xml
<!-- workbook.xml 中定义 -->
<workbook xmlns="...">
  <sheets>...</sheets>
  
  <definedNames>
    <!-- 给单个单元格起名 -->
    <definedName name="阈值">Sheet1!$D$1</definedName>
    
    <!-- 给区域起名 -->
    <definedName name="数据范围">Sheet1!$A$2:$A$100</definedName>
    
    <!-- 常量命名（不含单元格引用） -->
    <definedName name="PI">3.14159265358979</definedName>
    
    <!-- 命名公式 -->
    <definedName name="摄氏转华氏">9/5*Sheet1!$A$1+32</definedName>
  </definedNames>
</workbook>
```

### 27.2 工作表级命名区域

```xml
<!-- localSheetId 指定属于哪个Sheet（0-based索引） -->
<definedName name="本表数据" localSheetId="0">Sheet1!$A$2:$A$100</definedName>
```

### 27.3 命名区域的用途

```xml
<!-- 在公式中使用命名区域，更可读 -->
<c r="B1"><f>IF(A1>阈值,"超标","正常")</f><v>正常</v></c>
<c r="C1"><f>SUM(数据范围)</f><v>5500</v></c>
<c r="D1"><f>摄氏转华氏</f><v>77</v></c>
```

### 27.4 隐藏的命名区域

```xml
<definedName name="内部参数" hidden="1">Sheet1!$Z$1</definedName>
```

---

## 28. 分组与大纲：折叠/展开行和列

### 28.1 行分组（大纲）

```xml
<sheetData>
  <!-- outlineLevel 属性表示大纲级别 -->
  <row r="1" outlineLevel="0">          <!-- 顶级 -->
    <c r="A1" t="s"><v>0</v></c>        <!-- "2024年" -->
  </row>
  <row r="2" outlineLevel="1">          <!-- 一级子分组 -->
    <c r="A2" t="s"><v>1</v></c>        <!-- "Q1" -->
  </row>
  <row r="3" outlineLevel="2" collapsed="1">  <!-- 二级子分组，已折叠 -->
    <c r="A3" t="s"><v>2</v></c>        <!-- "1月" -->
  </row>
  <row r="4" outlineLevel="2" collapsed="1">
    <c r="A4" t="s"><v>3</v></c>        <!-- "2月" -->
  </row>
  <row r="5" outlineLevel="2" collapsed="1">
    <c r="A5" t="s"><v>4</v></c>        <!-- "3月" -->
  </row>
  <row r="6" outlineLevel="1">          <!-- Q2 -->
    <c r="A6" t="s"><v>5</v></c>
  </row>
</sheetData>

<!-- 大纲设置 -->
<sheetFormatPr outlineLevelRow="2" outlineLevelCol="0"/>
<!-- summaryBelow="0" 表示分组汇总在上方 -->
<sheetPr>
  <outlinePr summaryBelow="0" summaryRight="0"/>
</sheetPr>
```

### 28.2 列分组

```xml
<cols>
  <col min="1" max="1" width="15" customWidth="1"/>
  <col min="2" max="4" width="10" outlineLevel="1" collapsed="1" hidden="1"/>
  <col min="5" max="6" width="10" outlineLevel="1"/>
</cols>
```

### 28.3 大纲属性

| 属性 | 位置 | 说明 |
|------|------|------|
| `outlineLevel` | `<row>` 或 `<col>` | 大纲级别（0-7） |
| `collapsed` | `<row>` 或 `<col>` | 是否已折叠 |
| `hidden` | `<row>` 或 `<col>` | 是否隐藏（折叠时子项为隐藏） |
| `summaryBelow` | `<outlinePr>` | 汇总行在下方(1)还是上方(0) |
| `summaryRight` | `<outlinePr>` | 汇总列在右侧(1)还是左侧(0) |

### 28.4 折叠/展开的层级关系

```
级别0: ┌── 2024年 ─────────────────┐
       │                              │
级别1: │   ├── Q1 ──────────┐        │
       │   │                  │        │
级别2: │   │   ├── 1月  ◄────┤ 折叠   │
       │   │   ├── 2月  ◄────┤        │
       │   │   └── 3月  ◄────┘        │
       │   │                            │
级别1: │   ├── Q2 ──────────┐        │
       │   │   ├── 4月       │        │
       │   │   ├── 5月       │        │
       │   │   └── 6月       │        │
       │   └── ...            │        │
       └──────────────────────────────┘
```

---

## 29. 工作表保护与工作簿保护

### 29.1 工作表保护

```xml
<worksheet xmlns="...">
  <!-- sheetPr 必须在 sheetData 之前 -->
  <sheetPr>
    <!-- 保护设置 -->
    <sheetProtection 
      sheet="1"                    ← 1=启用保护
      objects="1"                  ← 禁止编辑对象（图片等）
      scenarios="1"                ← 禁止编辑方案
      formatCells="1"              ← 禁止格式化单元格
      formatColumns="1"            ← 禁止格式化列
      formatRows="1"               ← 禁止格式化行
      insertColumns="1"            ← 禁止插入列
      insertRows="1"               ← 禁止插入行
      insertHyperlinks="1"         ← 禁止插入超链接
      deleteColumns="1"            ← 禁止删除列
      deleteRows="1"               ← 禁止删除行
      selectLockedCells="0"        ← 0=允许选中锁定单元格
      selectUnlockedCells="0"      ← 0=允许选中未锁定单元格
      sort="1"                     ← 禁止排序
      autoFilter="1"               ← 禁止自动筛选
      pivotTables="1"              ← 禁止数据透视表
      password="..."               ← 密码哈希（可选）
    />
  </sheetPr>
  <sheetData>...</sheetData>
</worksheet>
```

> **注意**：属性值为 `1` 表示**禁止**该操作，`0` 表示**允许**。这是反直觉的！

### 29.2 单元格锁定

单元格默认是"锁定"的，但只在启用工作表保护后才生效。通过样式控制单元格是否锁定：

```xml
<!-- styles.xml 中的 cellXfs -->
<!-- locked="0" 表示未锁定，启用保护后仍可编辑 -->
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" 
    protectionId="0" applyProtection="1">
  <protection locked="0"/>          ← 未锁定
</xf>

<!-- locked="1" 表示已锁定（默认值，可省略） -->
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0">
  <!-- 默认 locked="1"，无需显式声明 -->
</xf>
```

### 29.3 密码哈希算法

工作表保护密码不是明文存储，而是用特定算法哈希：

```c
// Excel 工作表保护密码哈希算法
// 这是一个简单的哈希，不是加密！
uint16_t excel_password_hash(const char *password) {
    uint16_t hash = 0;
    int len = strlen(password);
    
    for (int i = len - 1; i >= 0; i--) {
        char ch = password[i];
        hash = ((hash >> 14) & 0x0001) | ((hash << 1) & 0x7FFF);
        hash ^= (uint16_t)ch;
    }
    
    // 再异或长度和魔数
    hash ^= len;
    hash ^= 0xCE4B;
    
    return hash;
}
```

生成的哈希值以十六进制写入 `password` 属性：

```xml
<sheetProtection sheet="1" password="ABCD" .../>
```

> **安全警告**：这个哈希极其脆弱，可以轻易破解。它只是防止误操作，不是真正的安全措施。

### 29.4 工作簿保护

```xml
<!-- workbook.xml 中 -->
<workbook xmlns="...">
  <sheets>...</sheets>
  
  <!-- 保护工作簿结构（禁止增删移动Sheet）和窗口（禁止调整窗口） -->
  <workbookProtection 
    lockStructure="1"       ← 禁止修改工作簿结构
    lockWindows="0"         ← 不禁止窗口操作
    lockRevision="0"        ← 不禁止修订跟踪
  />
</workbook>
```

---

## 30. 打印设置完整版

### 30.1 页面边距

```xml
<!-- 放在 sheetData 之后，mergeCells 之后 -->
<pageMargins 
  left="0.7"      ← 左边距（英寸）
  right="0.7"     ← 右边距
  top="0.75"      ← 上边距
  bottom="0.75"   ← 下边距
  header="0.3"    ← 页眉到顶部的距离
  footer="0.3"    ← 页脚到底部的距离
/>
```

常用预设值（英寸）：

| 预设 | left | right | top | bottom | header | footer |
|------|------|-------|-----|--------|--------|--------|
| 普通 | 0.7 | 0.7 | 0.75 | 0.75 | 0.3 | 0.3 |
| 窄 | 0.25 | 0.25 | 0.75 | 0.75 | 0.3 | 0.3 |
| 宽 | 1.0 | 1.0 | 1.0 | 1.0 | 0.5 | 0.5 |

### 30.2 页面设置

```xml
<pageSetup
  orientation="landscape"          ← portrait=纵向, landscape=横向
  paperSize="9"                   ← 纸张大小（见下表）
  scale="100"                     ← 缩放比例（1-100，与fitToWidth/Height互斥）
  fitToWidth="1"                  ← 缩放到N页宽（0=不限制）
  fitToHeight="0"                 ← 缩放到N页高（0=不限制）
  horizontalDpi="300"             ← 水平DPI
  verticalDpi="300"               ← 垂直DPI
  copies="1"                      ← 打印份数
  firstPageNumber="1"             ← 起始页码
  paperHeight="297mm"            ← 纸张高度（可选）
  paperWidth="210mm"             ← 纸张宽度（可选）
/>
```

### 30.3 纸张大小(paperSize)常用值

| 值 | 纸张 |
|----|------|
| 1 | Letter (8.5" × 11") |
| 8 | A3 (297mm × 420mm) |
| 9 | A4 (210mm × 297mm) |
| 11 | A5 (148mm × 210mm) |
| 5 | Legal (8.5" × 14") |
| 67 | B4 (250mm × 353mm) |
| 68 | B5 (176mm × 250mm) |
| 0 | 自定义 |

### 30.4 打印选项

```xml
<printOptions
  horizontalCentered="1"    ← 水平居中
  verticalCentered="0"       ← 垂直居中
  headings="0"              ← 是否打印行列号
  gridLines="0"             ← 是否打印网格线
  gridLinesSet="1"          ← 网格线设置标记
/>
```

### 30.5 打印区域

```xml
<!-- 定义打印范围 -->
<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$F$50</definedName>
```

### 30.6 打印标题行

每页重复打印指定的行或列：

```xml
<!-- 每页重复第1-2行作为标题行 -->
<definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$2</definedName>

<!-- 每页重复A列作为标题列 -->
<definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$A:$A</definedName>
```

---

## 31. 页眉页脚

### 31.1 基本结构

```xml
<headerFooter differentFirst="1" differentOddEven="0">
  <!-- 首页页眉页脚 -->
  <firstHeader>
    <oddHeader>
      &L                              ← 左侧
      &C                              ← 居中
      &R                              ← 右侧
    </oddHeader>
  </firstHeader>
  
  <!-- 默认页眉页脚 -->
  <oddHeader>
    <oddHeader>
      &L&"等线,常规"传感器数据报表&C&D&R第 &P 页，共 &N 页
    </oddHeader>
  </oddHeader>
  
  <oddFooter>
    &L&D &T&C&F&R第 &P / &N 页
  </oddFooter>
</headerFooter>
```

### 31.2 特殊代码

| 代码 | 含义 |
|------|------|
| `&L` | 左对齐 |
| `&C` | 居中 |
| `&R` | 右对齐 |
| `&P` | 当前页码 |
| `&N` | 总页数 |
| `&D` | 当前日期 |
| `&T` | 当前时间 |
| `&F` | 文件名 |
| `&A` | 工作表名 |
| `&Z` | 文件路径 |
| `&B` | 粗体开关 |
| `&I` | 斜体开关 |
| `&U` | 下划线开关 |
| `&S` | 删除线开关 |
| `&"字体名,样式"` | 字体设置 |
| `&数字` | 字号（如 `&12` = 12号字） |
| `&&` | 输出 & 符号本身 |

### 31.3 嵌入式实现

```xml
<!-- 简化版页眉页脚 -->
<headerFooter>
  <oddHeader>&C&"等线,12"B传感器数据报表&R&D &T</oddHeader>
  <oddFooter>&L&F&R第 &P 页 / 共 &N 页</oddFooter>
</headerFooter>
```

---

## 32. 分页符

### 32.1 手动分页符

```xml
<!-- 放在 sheetData 之后 -->
<rowBreaks count="2" manualBreakCount="2">
  <brk id="5" min="0" max="255" man="1"/>    ← 在第5行后插入分页符
  <brk id="20" min="0" max="255" man="1"/>   ← 在第20行后插入分页符
</rowBreaks>

<colBreaks count="1" manualBreakCount="1">
  <brk id="3" min="0" max="65535" man="1"/>  ← 在第3列(D列)后插入分页符
</colBreaks>
```

| 属性 | 说明 |
|------|------|
| `id` | 分页符位置（行号或列号，0-based） |
| `min`/`max` | 分页范围（通常 min=0, max=最大行列） |
| `man` | 1=手动分页符，0=自动分页符 |

---

## 33. 工作表视图设置

### 33.1 视图元素位置

`<sheetViews>` 是 `<worksheet>` 的第一个子元素，在 `<sheetData>` 之前：

```xml
<worksheet xmlns="...">
  <sheetViews>
    <sheetView ...>...</sheetView>
  </sheetViews>
  
  <sheetFormatPr .../>
  
  <cols>...</cols>
  
  <sheetData>...</sheetData>
  
  ...
</worksheet>
```

### 33.2 常用视图设置

```xml
<sheetViews>
  <sheetView 
    workbookViewId="0"              ← 关联的工作簿视图ID
    showGridLines="1"               ← 是否显示网格线（0=隐藏）
    showRowColHeaders="1"           ← 是否显示行列号
    showZeros="1"                  ← 是否显示零值
    rightToLeft="0"                 ← 是否从右到左（阿拉伯语等）
    tabSelected="1"                 ← 是否选中此Sheet的Tab
    zoomScale="100"                 ← 缩放比例（10-400）
    zoomScaleNormal="100"           ← 普通视图缩放
    zoomScalePageLayoutView="100"   ← 页面布局视图缩放
    view="normal"                   ← normal|pageBreakPreview|pageLayout
  >
    
    <!-- 冻结窗格（见第38章） -->
    <pane .../>
    
    <!-- 选中区域 -->
    <selection activeCell="A1" sqref="A1"/>
    
  </sheetView>
</sheetViews>
```

### 33.3 视图模式

| view 值 | 说明 |
|---------|------|
| `normal` | 普通视图（默认） |
| `pageBreakPreview` | 分页预览 |
| `pageLayout` | 页面布局视图 |

### 33.4 网格线控制

```xml
<!-- 隐藏网格线 -->
<sheetView showGridLines="0">
  <selection activeCell="A1" sqref="A1"/>
</sheetView>
```

### 33.5 从右到左布局（RTL）

```xml
<sheetView rightToLeft="1">
  <selection activeCell="A1" sqref="A1"/>
</sheetView>
```

---

## 34. 富文本：一个单元格多种格式

### 34.1 什么是富文本

普通单元格只能有一种样式（字体、颜色等）。富文本允许一个单元格内有多种格式：

```
┌──────────────────────────────────────────┐
│  温度: 25.3°C  状态: 正常  偏差: +0.3   │
│  ↑黑    ↑蓝(粗体)  ↑黑    ↑绿  ↑红      │
└──────────────────────────────────────────┘
```

### 34.2 共享字符串中的富文本

富文本存储在 sharedStrings.xml 中，使用 `<r>` (Run) 元素：

```xml
<sst xmlns="...">
  <!-- 普通字符串 -->
  <si><t>普通文本</t></si>
  
  <!-- 富文本 -->
  <si>
    <r>
      <rPr>                          ← 文本属性(Run Properties)
        <b/>                          ← 粗体
        <sz val="11"/>                ← 字号
        <color rgb="FF000000"/>       ← 黑色
        <rFont val="等线"/>           ← 字体
      </rPr>
      <t>温度: </t>
    </r>
    <r>
      <rPr>
        <sz val="14"/>
        <color rgb="FF0000FF"/>       ← 蓝色
        <rFont val="等线"/>
      </rPr>
      <t>25.3°C</t>
    </r>
    <r>
      <rPr>
        <sz val="11"/>
        <color rgb="FF000000"/>
        <rFont val="等线"/>
      </rPr>
      <t>  状态: </t>
    </r>
    <r>
      <rPr>
        <color rgb="FF00B050"/>       ← 绿色
        <rFont val="等线"/>
      </rPr>
      <t>正常</t>
    </r>
  </si>
</sst>
```

### 34.3 rPr 中的可用属性

```xml
<rPr>
  <b/>                           ← 粗体
  <i/>                           ← 斜体
  <u/>                           ← 下划线
  <strike/>                      ← 删除线
  <vertAlign val="superscript"/> ← 上标/下标
  <sz val="12"/>                 ← 字号
  <color rgb="FFFF0000"/>        ← 颜色
  <rFont val="微软雅黑"/>          ← 字体名
  <family val="2"/>              ← 字体族
</rPr>
```

### 34.4 嵌入式实现要点

- 富文本比普通文本占更多 XML 空间（每个格式变化都需要一组标签）
- 嵌入式中通常不需要太多富文本，几处标题或状态标记即可
- 富文本不影响单元格的样式索引 `s`，只影响文本内部的显示
- 对齐方式仍然由单元格样式控制，富文本只控制文字格式

---

## 35. 错误值

### 35.1 Excel 错误类型

| 错误值 | 含义 | XML 写法 |
|--------|------|---------|
| `#NULL!` | 交集为空 | `<c r="A1" t="e"><v>#NULL!</v></c>` |
| `#DIV/0!` | 除以零 | `<c r="A1" t="e"><v>#DIV/0!</v></c>` |
| `#VALUE!` | 值类型错误 | `<c r="A1" t="e"><v>#VALUE!</v></c>` |
| `#REF!` | 引用无效 | `<c r="A1" t="e"><v>#REF!</v></c>` |
| `#NAME?` | 名称未识别 | `<c r="A1" t="e"><v>#NAME?</v></c>` |
| `#NUM!` | 数字无效 | `<c r="A1" t="e"><v>#NUM!</v></c>` |
| `#N/A` | 值不可用 | `<c r="A1" t="e"><v>#N/A</v></c>` |
| `#GETTING_DATA` | 正在获取数据 | `<c r="A1" t="e"><v>#GETTING_DATA</v></c>` |
| `#SPILL!` | 溢出错误 | `<c r="A1" t="e"><v>#SPILL!</v></c>` |
| `#CALC!` | 计算错误 | `<c r="A1" t="e"><v>#CALC!</v></c>` |

### 35.2 使用场景

```c
// 嵌入式中主动输出错误值的情况
if (sensor_value == ERROR_READ_FAILED) {
    fprintf(f, "<c r=\"%s\" t=\"e\"><v>#N/A</v></c>", addr);  // 数据不可用
} else if (sensor_value == ERROR_DIV_ZERO) {
    fprintf(f, "<c r=\"%s\" t=\"e\"><v>#DIV/0!</v></c>", addr);  // 除零错误
} else {
    fprintf(f, "<c r=\"%s\"><v>%.2f</v></c>", addr, sensor_value);
}
```

### 35.3 公式中的错误值

如果公式计算结果为错误，`<v>` 中写入错误字符串，`t` 设为 `e`：

```xml
<c r="C1"><f>A1/B1</f><v>#DIV/0!</v></c>
```

---

## 36. Sheet Tab 颜色

### 36.1 设置 Tab 颜色

```xml
<!-- workbook.xml 中 -->
<sheets>
  <sheet name="数据" sheetId="1" r:id="rId1">
    <!-- Tab 颜色不在这里设置 -->
  </sheet>
</sheets>

<!-- 在 sheetN.xml 中设置 -->
<worksheet xmlns="...">
  <sheetPr>
    <tabColor rgb="FFFF0000"/>    ← 红色 Tab
  </sheetPr>
  <sheetViews>...</sheetViews>
  ...
</worksheet>
```

### 36.2 常用 Tab 颜色方案

| 颜色 | ARGB | 用途 |
|------|------|------|
| 红色 | `FFFF0000` | 警报/异常 Sheet |
| 绿色 | `FF00B050` | 正常/汇总 Sheet |
| 蓝色 | `FF0070C0` | 数据 Sheet |
| 黄色 | `FFFFFF00` | 临时/草稿 Sheet |
| 灰色 | `FF808080` | 归档 Sheet |
| 紫色 | `FF7030A0` | 配置 Sheet |

---

## 37. 隐藏行/列与行列高级操作

### 37.1 隐藏行

```xml
<row r="5" hidden="1">
  <c r="A5" t="s"><v>3</v></c>
</row>
```

### 37.2 隐藏列

```xml
<cols>
  <col min="3" max="3" width="0" hidden="1"/>     ← C列隐藏
  <col min="5" max="6" width="0" hidden="1"/>     ← E-F列隐藏
</cols>
```

### 37.3 自定义行高

```xml
<!-- 精确行高（磅为单位） -->
<row r="3" ht="28.5" customHeight="1">...</row>

<!-- 自动行高（不设 ht 和 customHeight） -->
<row r="4">...</row>
```

### 37.4 最佳列宽（手动估算）

Excel 没有"自动列宽"的 XML 属性——列宽是在客户端渲染时计算的。在嵌入式生成时，需要自行估算：

```c
// 估算字符串显示宽度（字符数，近似值）
int estimate_width(const char *str) {
    int width = 0;
    while (*str) {
        if ((unsigned char)*str >= 0x80) {
            // 中文字符约占2个字符宽
            width += 2;
            // 跳过 UTF-8 多字节字符的剩余字节
            unsigned char c = (unsigned char)*str;
            if (c >= 0xF0) str += 3;
            else if (c >= 0xE0) str += 2;
            else if (c >= 0xC0) str += 1;
        } else {
            width += 1;
        }
        str++;
    }
    return width;
}

// 设置列宽为最佳宽度
void auto_col_width(worksheet_t *ws, int col, const char *longest_str) {
    double width = estimate_width(longest_str) + 2.0;  // 加2字符的边距
    if (width > 255) width = 255;   // Excel 最大列宽
    if (width < 8.43) width = 8.43; // 默认列宽
    ws_set_col_width(ws, col, width);
}
```

### 37.5 行高单位换算

```
1 磅(pt) = 1/72 英寸 = 0.353 毫米
1 像素(96dpi) = 0.75 磅

常用行高：
15 磅  = 默认行高（11号字）
20 磅  = 稍大
30 磅  = 双倍
45 磅  = 三倍
```

---

## 38. 冻结窗格与拆分窗格完整版

### 38.1 冻结首行

```xml
<sheetView workbookViewId="0">
  <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
  <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
</sheetView>
```

### 38.2 冻结前 N 行

```xml
<!-- 冻结前3行 -->
<sheetView workbookViewId="0">
  <pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/>
  <selection pane="bottomLeft" activeCell="A4" sqref="A4"/>
</sheetView>
```

### 38.3 冻结首列

```xml
<sheetView workbookViewId="0">
  <pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>
  <selection pane="topRight" activeCell="B1" sqref="B1"/>
</sheetView>
```

### 38.4 冻结前 N 列

```xml
<!-- 冻结前2列 -->
<sheetView workbookViewId="0">
  <pane xSplit="2" topLeftCell="C1" activePane="topRight" state="frozen"/>
  <selection pane="topRight" activeCell="C1" sqref="C1"/>
</sheetView>
```

### 38.5 同时冻结行和列

```xml
<!-- 冻结前2行和前1列 -->
<sheetView workbookViewId="0">
  <pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>
  <selection pane="bottomRight" activeCell="B3" sqref="B3"/>
</sheetView>
```

### 38.6 拆分窗格（非冻结）

```xml
<sheetView workbookViewId="0">
  <pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="split"/>
</sheetView>
```

`state="split"` 与 `state="frozen"` 的区别：拆分窗格的分隔线可以拖动，冻结的不可以。

### 38.7 参数说明

| 参数 | 说明 |
|------|------|
| `xSplit` | 垂直分隔线位置（列数，0-based） |
| `ySplit` | 水平分隔线位置（行数，0-based） |
| `topLeftCell` | 右下窗格的左上角单元格 |
| `activePane` | 活动窗格：`topLeft`/`topRight`/`bottomLeft`/`bottomRight` |
| `state` | `frozen`=冻结, `split`=拆分, 不写=无 |

---

## 39. 列宽行高单位换算精解

### 39.1 列宽的"字符单位"

Excel 的列宽单位不是像素也不是磅，而是**"0号字体的数字字符宽度"**：

```
列宽 = 1 表示可以放下约 1 个 "0" 字符的宽度
列宽 = 8.43 = 默认列宽（可放下约 8 个字符）
列宽 = 20 = 可放下约 20 个字符
```

### 39.2 换算公式

```
像素宽度 = Truncate((256 * 列宽 + Truncate(128/7)) / 256 * 7)

近似公式：
像素 = 列宽 * 7 + 5      （96 DPI 下）

磅 = 像素 * 72 / 96
英寸 = 像素 / 96
厘米 = 像素 * 2.54 / 96
毫米 = 像素 * 25.4 / 96
```

### 39.3 列宽与像素对照表

| 列宽 | 像素 | 厘米 | 说明 |
|------|------|------|------|
| 8.43 | 64 | 1.69 | 默认列宽 |
| 10 | 75 | 1.98 | |
| 15 | 110 | 2.91 | |
| 20 | 145 | 3.84 | |
| 30 | 215 | 5.68 | 宽列 |
| 50 | 355 | 9.39 | 非常宽 |
| 255 | 1785 | 47.2 | 最大列宽 |

### 39.4 行高换算

行高单位是**磅(pt)**：

```
磅 = 行高值
像素 = 行高 * 96 / 72 = 行高 * 4 / 3
厘米 = 行高 * 2.54 / 72
毫米 = 行高 * 25.4 / 72
```

### 39.5 行高与像素对照表

| 行高(磅) | 像素 | 厘米 | 说明 |
|---------|------|------|------|
| 15 | 20 | 0.53 | 默认行高 |
| 20 | 27 | 0.71 | |
| 30 | 40 | 1.06 | 双倍 |
| 45 | 60 | 1.59 | 三倍 |
| 60 | 80 | 2.12 | |
| 409 | 545 | 14.4 | 最大行高 |

### 39.6 C 语言换算函数

```c
// 列宽 → 像素（96 DPI）
int col_width_to_pixels(double width) {
    return (int)(width * 7 + 5);
}

// 像素 → 列宽
double pixels_to_col_width(int pixels) {
    return (pixels - 5.0) / 7.0;
}

// 行高 → 像素
int row_height_to_pixels(double height) {
    return (int)(height * 96.0 / 72.0);
}

// 像素 → 行高
double pixels_to_row_height(int pixels) {
    return pixels * 72.0 / 96.0;
}

// EMU → 像素
int emu_to_pixels(int emu) {
    return emu / 9525;
}

// 像素 → EMU
int pixels_to_emu(int pixels) {
    return pixels * 9525;
}
```

---

## 40. 自定义数字格式语法大全

### 40.1 四段式格式

自定义数字格式最多有 4 段，用分号分隔：

```
正数格式;负数格式;零值格式;文本格式
```

示例：
```
#,##0.00;[Red]-#,##0.00;0.00;"—"
```

效果：
- 正数 1234.5 → `1,234.50`
- 负数 -1234.5 → `1,234.50`（红色）
- 零 0 → `0.00`
- 文本 "abc" → `—`

### 40.2 格式代码详解

| 代码 | 含义 | 示例输入 | 显示效果 |
|------|------|---------|---------|
| `0` | 数字占位（不足补零） | `25.3` | `25` (格式 `0`) |
| `#` | 数字占位（不补零） | `25` | `25` (格式 `#`) |
| `0.00` | 固定2位小数 | `25.3` | `25.30` |
| `#,##0` | 千分位 | `1234567` | `1,234,567` |
| `#,##0.00` | 千分位+2位小数 | `1234567.8` | `1,234,567.80` |
| `0%` | 百分比（整数） | `0.25` | `25%` |
| `0.00%` | 百分比（2位小数） | `0.255` | `25.50%` |
| `0.00E+00` | 科学计数法 | `12345` | `1.23E+04` |
| `?` | 数字占位（空格对齐） | `25.3` | `25.3 ` (对齐小数点) |
| `??/??` | 分数 | `0.25` | `1/4` |
| `# ??/??` | 带整数的分数 | `1.25` | `1 1/4` |

### 40.3 颜色和条件

```
[颜色]格式
[条件]格式
```

| 代码 | 效果 |
|------|------|
| `[Red]` | 红色 |
| `[Blue]` | 蓝色 |
| `[Green]` | 绿色 |
| `[Yellow]` | 黄色 |
| `[Black]` | 黑色 |
| `[Cyan]` | 青色 |
| `[Magenta]` | 品红 |
| `[White]` | 白色 |
| `[Color1]`~`[Color56]` | 调色板颜色 |

条件格式：
```
[>50][Green]0;[<20][Red]0;[Black]0
```

效果：
- 大于 50：绿色
- 小于 20：红色
- 其他：黑色

### 40.4 文字嵌入

```
0.00"°C"
#,##0"元"
0.00" "(元)
```

用双引号包裹文字部分。

### 40.5 重复字符

```
* = 重复下一个字符填满列宽
```

示例：`$* #` → `$      1234`（$ 后面用空格填满）

### 40.6 常用自定义格式速查

| formatCode | 用途 | 示例 |
|------------|------|------|
| `0.00"°C"` | 温度 | `25.37°C` |
| `0.00"V"` | 电压 | `3.30V` |
| `0.00"A"` | 电流 | `1.25A` |
| `0.0"Hz"` | 频率 | `50.0Hz` |
| `0"ms"` | 毫秒 | `150ms` |
| `0.00"%"` | 百分比（直接值） | `12.50%` |
| `[>100][Red]0;[<0][Blue]0;0` | 范围标色 | |
| `#,##0.00" kg"` | 重量 | `1,234.50 kg` |
| `0.00E+00` | 科学计数 | `1.23E+04` |
| `??/??` | 分数 | `1/4` |

---

## 41. Sparklines：单元格里的迷你图

### 41.1 什么是 Sparkline

Sparkline 是嵌入在单个单元格中的微型图表，通常用于显示趋势：

```
┌────────┬──────────────────┐
│ 一月   │ 25.3             │
│ 二月   │ 26.1             │
│ 三月   │ 24.8             │
│ 四月   │ 28.2             │
│ 趋势   │ ▁▂▁▃ (迷你折线)   │  ← Sparkline
└────────┴──────────────────┘
```

### 41.2 文件结构

```
xl/
├── sparklines/
│   └── sparkline1.xml     ← Sparkline 组定义
├── worksheets/
│   └── sheet1.xml
└── _rels/
    └── sheet1.xml.rels
```

### 41.3 sparkline1.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x14:sparklineGroups xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
                      xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <x14:sparklineGroup 
    xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    type="line"              ← line=折线, column=柱状, stacked=盈亏
    displayEmptyCellsAs="gap"
    markers="1"              ← 显示数据点标记
    high="1"                 ← 标记最高点
    low="1"                  ← 标记最低点
    negative="1">            ← 标记负值
    
    <x14:colorSeries rgb="FF3266CC"/>     ← 线条颜色
    <x14:colorNegative rgb="FFFF0000"/>   ← 负值颜色
    <x14:colorMarkers rgb="FFD00000"/>   ← 标记颜色
    <x14:colorHigh rgb="FF00B050"/>      ← 最高点颜色
    <x14:colorLow rgb="FFFF0000"/>       ← 最低点颜色
    
    <x14:sparklines>
      <x14:sparkline>
        <xm:f>Sheet1!A1:A5</xm:f>         ← 数据范围
        <xm:sqref>C1</xm:sqref>           ← 显示位置（单元格地址）
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>Sheet1!B1:B5</xm:f>
        <xm:sqref>C2</xm:sqref>
      </x14:sparkline>
    </x14:sparklines>
  </x14:sparklineGroup>
</x14:sparklineGroups>
```

### 41.4 Sparkline 类型

| type | 说明 |
|------|------|
| `line` | 折线图 |
| `column` | 柱状图 |
| `stacked` | 盈亏图（正值向上、负值向下） |

### 41.5 命名空间注意

Sparkline 使用 `x14` 命名空间（Microsoft 扩展），不是标准 OOXML 的一部分。但所有现代 Excel 版本都支持。

---

## 42. 外部链接

### 42.1 引用外部工作簿

```xml
<!-- sheet1.xml 中的公式引用外部文件 -->
<c r="A1"><f>'[外部数据.xlsx]Sheet1'!A1</f><v>100</v></c>
```

### 42.2 外部链接文件

```
xl/
├── externalLinks/
│   ├── externalLink1.xml      ← 外部链接定义
│   └── _rels/
│       └── externalLink1.xml.rels  ← 指向实际外部文件
├── worksheets/
│   └── sheet1.xml
└── _rels/
    └── workbook.xml.rels      ← workbook → externalLink 关系
```

### 42.3 externalLink1.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:mc="...">
  <externalBook>
    <sheetNames>
      <sheetName val="Sheet1"/>
    </sheetNames>
    <sheetDataSet>
      <sheetData sheetId="0" refreshError="0">
        <cell ref="A1" type="num">
          <v>100</v>
        </cell>
      </sheetData>
    </sheetDataSet>
  </externalBook>
</externalLink>
```

> **嵌入式建议**：外部链接在嵌入式系统中很少使用，因为通常所有数据都在同一个文件中。如果需要跨文件引用，建议在代码中合并数据后写入一个文件。

---

## 43. 自定义文档属性

### 43.1 文件结构

```
docProps/
├── core.xml     ← 标准属性（标题、作者、日期等）
├── app.xml      ← 应用属性（公司、版本等）
└── custom.xml   ← 自定义属性
```

### 43.2 core.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties 
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>传感器数据报表</dc:title>
  <dc:creator>嵌入式采集系统</dc:creator>
  <cp:lastModifiedBy>自动生成</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-03-15T14:30:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-03-15T14:30:00Z</dcterms:modified>
  <cp:keywords>传感器,温度,报表</cp:keywords>
  <dc:description>由 STM32 数据采集系统自动生成</dc:description>
  <cp:category>数据报表</cp:category>
  <cp:version>1.0</cp:version>
</cp:coreProperties>
```

### 43.3 app.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Embedded Excel Generator v1.0</Application>
  <Company>MyCompany</Company>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>工作表</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>3</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="3" baseType="lpstr">
      <vt:lpstr>数据</vt:lpstr>
      <vt:lpstr>图表</vt:lpstr>
      <vt:lpstr>汇总</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
</Properties>
```

### 43.4 custom.xml（自定义属性）

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="设备ID">
    <vt:lpwstr>STM32_001</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="固件版本">
    <vt:lpwstr>v2.1.3</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="采样率">
    <vt:i4>1000</vt:i4>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="导出时间戳">
    <vt:filetime>2024-03-15T14:30:00Z</vt:filetime>
  </property>
</Properties>
```

### 43.5 自定义属性的数据类型

| vt 类型 | 说明 | 示例 |
|---------|------|------|
| `vt:lpwstr` | 字符串 | `<vt:lpwstr>文本</vt:lpwstr>` |
| `vt:i4` | 32位整数 | `<vt:i4>42</vt:i4>` |
| `vt:i8` | 64位整数 | `<vt:i8>9223372036854775807</vt:i8>` |
| `vt:r8` | 双精度浮点 | `<vt:r8>3.14</vt:r8>` |
| `vt:bool` | 布尔 | `<vt:bool>true</vt:bool>` |
| `vt:filetime` | 日期时间 | `<vt:filetime>2024-03-15T14:30:00Z</vt:filetime>` |

---

## 44. VBA 宏文件(.xlsm)

### 44.1 .xlsm vs .xlsx

| 特性 | .xlsx | .xlsm |
|------|-------|-------|
| 包含 VBA 宏 | 否 | 是 |
| Content_Types | `spreadsheetml.sheet.main+xml` | `spreadsheetml.sheet.macroEnabled.main+xml` |
| 额外文件 | 无 | `xl/vbaProject.bin` |
| 安全性 | 可安全打开 | 需用户确认启用宏 |

### 44.2 Content_Types 差异

```xml
<!-- .xlsx -->
<Override PartName="/xl/workbook.xml" 
          ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>

<!-- .xlsm -->
<Override PartName="/xl/workbook.xml" 
          ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.macroEnabled.main+xml"/>

<!-- 额外声明 vbaProject.bin -->
<Override PartName="/xl/vbaProject.bin" 
          ContentType="application/vnd.ms-office.vbaProject"/>
```

### 44.3 vbaProject.bin

VBA 宏以二进制格式（.bin）存储，不是 XML。这个文件包含：
- VBA 代码（编译后的 P-Code）
- 模块结构
- 引用信息

> **嵌入式限制**：在嵌入式系统中通常不生成 VBA 宏（太复杂且二进制格式难以手工构造）。如果需要宏功能，建议在 PC 端后处理，或使用公式替代宏。

### 44.4 workbook.xml 中的宏声明

```xml
<workbook xmlns="..." xmlns:r="...">
  <!-- 声明包含宏代码 -->
  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
    <mc:Choice Requires="vba">
      <codeName>ThisWorkbook</codeName>
    </mc:Choice>
  </mc:AlternateContent>
  <sheets>...</sheets>
</workbook>
```

---

## 45. 文件加密与密码保护

### 45.1 加密概述

Excel 文件加密使用 AGILE 加密（ECMA-376 标准），包含：
- 密码密钥派生（PBKDF2）
- AES-128/192/256 加密
- SHA-1/SHA-256 哈希

### 45.2 加密后的文件结构

加密的 xlsx 文件不再是标准 ZIP 结构，而是一个 OLE2 复合文档：

```
[OLE2 头]
├── EncryptionInfo     ← 加密参数（算法、密钥派生参数等）
├── EncryptedPackage    ← 加密后的 ZIP 数据
└── SummaryInformation
```

### 45.3 嵌入式可行性

| 方面 | 可行性 | 说明 |
|------|--------|------|
| AES 加密 | 可行 | 很多 MCU 有硬件 AES |
| SHA-1/256 | 可行 | 可用软件实现 |
| PBKDF2 | 可行 | 但迭代次数高(10000+)时较慢 |
| OLE2 容器 | 困难 | 需要额外实现 OLE2 格式 |

> **嵌入式建议**：文件加密在嵌入式上实现成本较高。如果需要保护数据，建议在传输层（TLS/HTTPS）或存储层（Flash 加密）解决，而不是在文件格式层加密。

---

## 46. Deflate 压缩可选实现

### 46.1 什么时候需要压缩

Store 模式（不压缩）的 xlsx 文件比压缩版大 3-10 倍。在以下场景可能需要压缩：

- 存储空间有限（如 SPI Flash）
- 传输带宽有限（如通过 UART/LoRa 传输）
- 文件包含大量文本数据（XML 压缩率高）

### 46.2 压缩效果对比

| 文件类型 | Store 大小 | Deflate 大小 | 压缩比 |
|---------|-----------|-------------|--------|
| 纯文本数据 | 100 KB | ~15 KB | 6.7:1 |
| 混合数据 | 50 KB | ~8 KB | 6.3:1 |
| 图片为主 | 200 KB | ~195 KB | 1.03:1 |
| XML 样式表 | 5 KB | ~1 KB | 5:1 |

### 46.3 集成 miniz（推荐方案）

miniz 是一个单文件的 Deflate 实现，适合嵌入式：

```c
// miniz 是一个单头文件库
// 下载: https://github.com/richgel999/miniz
#include "miniz.h"

// 压缩数据
int compress_data(const uint8_t *input, size_t input_size, 
                  uint8_t *output, size_t output_size) {
    mz_ulong compressed_size = output_size;
    int ret = mz_compress2(output, &compressed_size, input, input_size, 
                            MZ_DEFAULT_COMPRESSION);
    if (ret != MZ_OK) return -1;
    return (int)compressed_size;
}
```

### 46.4 在 ZIP 中使用 Deflate

```c
void zip_add_file_deflate(zip_writer_t *zw, const char *filename,
                           const char *data, size_t size) {
    // 压缩数据
    size_t compressed_size_max = size + size / 100 + 512;  // Deflate 上界
    uint8_t *compressed = malloc(compressed_size_max);
    int compressed_size = compress_data((const uint8_t*)data, size, 
                                          compressed, compressed_size_max);
    
    // 计算 CRC32（用未压缩的数据算）
    uint32_t crc = crc32_compute((const uint8_t*)data, size);
    
    // 写 Local Header
    // ... 同 Store 模式，但 compression method = 8 (Deflate) ...
    write_u16(zw->file, 8);  // ← 8 = Deflate
    
    // 写压缩后的数据
    fwrite(compressed, 1, compressed_size, zw->file);
    
    free(compressed);
}
```

### 46.5 miniz 的资源占用

| 资源 | 占用 | 说明 |
|------|------|------|
| ROM（代码） | ~15 KB | miniz 压缩+解压 |
| RAM | ~32 KB | 压缩缓冲区（可调整） |
| ROM（CRC表） | 1 KB | 与 Store 模式共享 |

---

## 47. API 设计模式

### 47.1 库的 API 层次设计

```
┌─────────────────────────────────────┐
│  应用层 (用户代码)                    │
├─────────────────────────────────────┤
│  高级 API (wb_add_data, wb_save)     │  ← 简单易用
├─────────────────────────────────────┤
│  中级 API (ws_set_cell, ws_merge)    │  ← 灵活控制
├─────────────────────────────────────┤
│  低级 API (xml_write, zip_add)       │  ← 底层操作
├─────────────────────────────────────┤
│  硬件层 (输出函数回调)                │  ← 可移植
└─────────────────────────────────────┘
```

### 47.2 输出回调设计（可移植性核心）

```c
// 输出回调函数类型
typedef int (*excel_write_fn)(const void *data, size_t size, void *ctx);

// 初始化时传入输出回调
typedef struct {
    excel_write_fn write;        // 写入回调
    void *write_ctx;             // 用户上下文（如文件句柄、UART 句柄）
    
    // 内存配置
    size_t max_memory;           // 最大内存使用
    int compression_enabled;      // 是否启用压缩
    
    // 功能开关
    int use_shared_strings;       // 是否使用共享字符串
    int max_sheets;              // 最大 Sheet 数
    int max_styles;               // 最大样式数
} excel_config_t;

// 文件输出回调实现
int file_write_cb(const void *data, size_t size, void *ctx) {
    FILE *f = (FILE *)ctx;
    return fwrite(data, 1, size, f) == size ? 0 : -1;
}

// UART 输出回调实现
int uart_write_cb(const void *data, size_t size, void *ctx) {
    uart_handle_t *uart = (uart_handle_t *)ctx;
    return uart_send(uart, data, size);
}

// 内存缓冲区输出回调实现
typedef struct {
    uint8_t *buf;
    size_t size;
    size_t capacity;
} membuf_t;

int membuf_write_cb(const void *data, size_t size, void *ctx) {
    membuf_t *mb = (membuf_t *)ctx;
    if (mb->size + size > mb->capacity) return -1;
    memcpy(mb->buf + mb->size, data, size);
    mb->size += size;
    return 0;
}
```

### 47.3 流式 API 设计

```c
// 创建工作簿
excel_handle_t *excel_create(excel_config_t *config);

// 添加 Sheet
int excel_add_sheet(excel_handle_t *h, const char *name);

// 设置单元格（字符串）
int excel_set_str(excel_handle_t *h, int sheet, int row, int col, 
                   const char *value, int style);

// 设置单元格（数字）
int excel_set_num(excel_handle_t *h, int sheet, int row, int col, 
                   double value, int style);

// 设置单元格（日期）
int excel_set_date(excel_handle_t *h, int sheet, int row, int col,
                    int year, int month, int day, int style);

// 合并单元格
int excel_merge(excel_handle_t *h, int sheet, 
                 int r1, int c1, int r2, int c2);

// 设置列宽
int excel_set_col_width(excel_handle_t *h, int sheet, int col, double width);

// 设置行高
int excel_set_row_height(excel_handle_t *h, int sheet, int row, double height);

// 添加数据验证
int excel_add_validation(excel_handle_t *h, int sheet, 
                          const char *range, int type, ...);

// 添加条件格式
int excel_add_cond_format(excel_handle_t *h, int sheet,
                           const char *range, int rule_type, ...);

// 冻结窗格
int excel_freeze_panes(excel_handle_t *h, int sheet, int rows, int cols);

// 最终保存
int excel_save(excel_handle_t *h);

// 销毁
void excel_destroy(excel_handle_t *h);
```

### 47.4 预定义样式常量

```c
// 样式索引常量，方便使用
enum {
    STYLE_DEFAULT = 0,        // 默认
    STYLE_TITLE,              // 标题（粗体居中大字）
    STYLE_HEADER,             // 表头（粗体居中+背景色+边框）
    STYLE_DATA,               // 数据（边框左对齐）
    STYLE_DATA_CENTER,        // 数据居中
    STYLE_DATA_RIGHT,         // 数据右对齐
    STYLE_NUMBER_2DEC,        // 数字2位小数
    STYLE_PERCENT,            // 百分比
    STYLE_DATE,               // 日期
    STYLE_DATETIME,           // 日期时间
    STYLE_MONEY,              // 货币
    STYLE_MERGE_CENTER,       // 合并居中（粗体+边框）
    STYLE_WARNING,            // 警告（红色背景）
    STYLE_SUCCESS,            // 成功（绿色背景）
    STYLE_ALT_ROW,            // 隔行变色
    STYLE_TOTAL,              // 合计行（粗体+上边框）
    STYLE_FROZEN_HEADER,      // 冻结表头样式
    STYLE_HYPERLINK,          // 超链接样式（蓝色下划线）
    STYLE_ERROR,              // 错误值样式
    STYLE_MAX
};
```

### 47.5 Builder 模式（可选）

```c
// 链式调用风格
excel_set_str(h, 0, 1, 1, "标题", STYLE_TITLE);
excel_merge(h, 0, 1, 1, 1, 3);

// 可以封装为宏，支持链式调用
#define EXCEL_SET_STR(h, s, r, c, v, st) \
    do { excel_set_str(h, s, r, c, v, st); } while(0)

// 或者用宏实现简化的表格构建
#define EXCEL_ROW_BEGIN(ws, row)  int __row = row;
#define EXCEL_CELL(ws, col, val, st)  excel_set_str(h, ws, __row, col, val, st)
#define EXCEL_ROW_END()  __row++
```

---

## 48. 线程安全与 RTOS 集成

### 48.1 线程安全问题

如果 Excel 生成库在 RTOS（FreeRTOS/RTX/Zephyr）环境中使用，需要考虑：

| 问题 | 说明 | 对策 |
|------|------|------|
| 共享状态 | 工作簿数据结构被多线程访问 | 加互斥锁 |
| 静态变量 | 全局变量被多线程修改 | 改为实例化（每个 handle 独立） |
| 内存分配 | malloc 线程安全性 | 使用 RTOS 提供的 malloc 或内存池 |
| 输出回调 | 写入回调被并发调用 | 在 handle 中加锁 |

### 48.2 线程安全实现

```c
typedef struct {
    // ... 数据字段 ...
    
    // RTOS 互斥锁
    SemaphoreHandle_t mutex;
} excel_handle_t;

// 所有 API 函数加锁
int excel_set_str(excel_handle_t *h, int sheet, int row, int col, 
                   const char *value, int style) {
    xSemaphoreTake(h->mutex, portMAX_DELAY);    // 加锁
    
    // ... 实际操作 ...
    
    xSemaphoreGive(h->mutex);                     // 解锁
    return 0;
}
```

### 48.3 无锁设计（单线程生成）

如果生成过程在单个任务中完成（推荐），不需要加锁：

```c
// RTOS 任务：单线程生成 Excel
void excel_export_task(void *param) {
    // 在此任务中完成所有 Excel 操作
    // 不需要加锁，因为没有并发访问
    
    excel_config_t config = {
        .write = file_write_cb,
        .write_ctx = &file_handle,
    };
    
    excel_handle_t *h = excel_create(&config);
    excel_add_sheet(h, "数据");
    // ... 填充数据 ...
    excel_save(h);
    excel_destroy(h);
    
    vTaskDelete(NULL);  // 任务完成，删除自身
}
```

### 48.4 栈大小估算

```c
// Excel 生成任务的栈大小估算
// 基础栈：512 字节
// 每个 Sheet：+ 256 字节
// 流式生成缓冲区：+ 512 字节
// sprintf 等库函数：+ 1024 字节
// 总计：约 2-4 KB

#define EXCEL_TASK_STACK_SIZE  (4 * 1024)  // 4KB 栈

xTaskCreate(excel_export_task, "Excel", 
            EXCEL_TASK_STACK_SIZE / sizeof(StackType_t), 
            NULL, tskIDLE_PRIORITY + 1, NULL);
```

### 48.5 任务优先级

Excel 生成不是实时关键任务，建议低优先级：

```c
// 低于数据采集任务，避免影响实时性
#define EXCEL_TASK_PRIORITY  (tskIDLE_PRIORITY + 1)

// 如果生成过程中不能被低优先级任务抢占
// 但可以被高优先级任务（如通信任务）抢占
```

### 48.6 分块生成（避免长时间阻塞）

```c
void excel_export_task(void *param) {
    excel_handle_t *h = excel_create(&config);
    
    // 每次只生成一部分行，然后让出 CPU
    for (int i = 0; i < total_rows; i += CHUNK_SIZE) {
        int end = (i + CHUNK_SIZE < total_rows) ? i + CHUNK_SIZE : total_rows;
        
        for (int j = i; j < end; j++) {
            excel_set_num(h, 0, j, 0, sensor_data[j], STYLE_DATA);
        }
        
        vTaskDelay(1);  // 让出 CPU 1 个 tick
    }
    
    excel_save(h);
    excel_destroy(h);
}
```

---

## 49. 错误处理框架

### 49.1 错误码定义

```c
typedef enum {
    EXCEL_OK = 0,
    
    // 内存错误
    EXCEL_ERR_NO_MEMORY,           // 内存不足
    EXCEL_ERR_BUFFER_FULL,         // 缓冲区已满
    EXCEL_ERR_TOO_MANY_SHEETS,     // Sheet 数量超限
    EXCEL_ERR_TOO_MANY_CELLS,      // 单元格数量超限
    EXCEL_ERR_TOO_MANY_MERGES,     // 合并区域超限
    EXCEL_ERR_TOO_MANY_STYLES,     // 样式超限
    EXCEL_ERR_STRING_TOO_LONG,     // 字符串太长
    EXCEL_ERR_SST_FULL,            // 共享字符串表已满
    
    // 参数错误
    EXCEL_ERR_INVALID_ROW,         // 无效行号
    EXCEL_ERR_INVALID_COL,         // 无效列号
    EXCEL_ERR_INVALID_SHEET,       // 无效 Sheet 索引
    EXCEL_ERR_INVALID_STYLE,       // 无效样式索引
    EXCEL_ERR_INVALID_RANGE,       // 无效合并范围（如交叉合并）
    EXCEL_ERR_INVALID_NAME,        // 无效 Sheet 名称
    
    // 文件错误
    EXCEL_ERR_FILE_OPEN,           // 文件打开失败
    EXCEL_ERR_FILE_WRITE,          // 文件写入失败
    EXCEL_ERR_FILE_READ,           // 文件读取失败
    
    // 格式错误
    EXCEL_ERR_XML_ESCAPE,          // XML 转义失败
    EXCEL_ERR_CRC32,               // CRC32 计算失败
    EXCEL_ERR_ZIP,                 // ZIP 打包失败
    
    // 状态错误
    EXCEL_ERR_NOT_INITIALIZED,     // 未初始化
    EXCEL_ERR_ALREADY_SAVED,       // 已保存，不能再修改
    EXCEL_ERR_SHEET_EXISTS,        // Sheet 名称已存在
    
} excel_error_t;
```

### 49.2 错误处理策略

```c
// 返回错误码，调用者决定如何处理
int excel_set_str(excel_handle_t *h, int sheet, int row, int col, 
                   const char *value, int style) {
    if (!h) return EXCEL_ERR_NOT_INITIALIZED;
    if (sheet < 0 || sheet >= h->sheet_count) return EXCEL_ERR_INVALID_SHEET;
    if (row < 1 || row > MAX_ROWS) return EXCEL_ERR_INVALID_ROW;
    // ...
    
    if (h->sheets[sheet].cell_count >= MAX_CELLS) {
        return EXCEL_ERR_TOO_MANY_CELLS;
    }
    
    // ... 添加单元格 ...
    return EXCEL_OK;
}

// 应用层错误处理
void generate_report() {
    int ret;
    
    ret = excel_set_str(h, 0, 1, 1, "标题", STYLE_TITLE);
    if (ret != EXCEL_OK) {
        log_error("设置单元格失败: %d", ret);
        // 降级处理：跳过或使用默认值
    }
    
    ret = excel_save(h);
    if (ret != EXCEL_OK) {
        log_error("保存失败: %d", ret);
        // 尝试恢复或通知用户
    }
}
```

### 49.3 优雅降级

```c
// 当资源不足时自动降级
int excel_set_str(excel_handle_t *h, int sheet, int row, int col,
                   const char *value, int style) {
    // 如果共享字符串表满了，降级为内联字符串
    int idx = sst_add(&h->sst, value);
    if (idx < 0) {
        // 降级：使用内联字符串
        return set_cell_inline_str(h, sheet, row, col, value, style);
    }
    return set_cell_shared_str(h, sheet, row, col, idx, style);
}
```

### 49.4 调试支持

```c
// 调试模式：输出详细日志
#ifdef EXCEL_DEBUG
    #define EXCEL_LOG(fmt, ...)  printf("[EXCEL] " fmt "\n", ##__VA_ARGS__)
    #define EXCEL_LOG_CELL(sheet, row, col, val) \
        printf("[EXCEL] Sheet%d R%dC%d = %s\n", sheet, row, col, val)
#else
    #define EXCEL_LOG(fmt, ...)
    #define EXCEL_LOG_CELL(sheet, row, col, val)
#endif
```

---

## 50. 兼容性矩阵

### 50.1 各软件兼容性

| 功能 | Excel 2007+ | Excel 2016+ | WPS Office | LibreOffice | Google Sheets |
|------|------------|------------|------------|-------------|---------------|
| 基本 xlsx | ✅ | ✅ | ✅ | ✅ | ✅ |
| 合并单元格 | ✅ | ✅ | ✅ | ✅ | ✅ |
| 条件格式 | ✅ | ✅ | ✅ | ⚠️ 部分 | ⚠️ 部分 |
| 图表 | ✅ | ✅ | ✅ | ⚠️ 部分 | ⚠️ 部分 |
| 图片 | ✅ | ✅ | ✅ | ✅ | ⚠️ 仅查看 |
| 数据验证 | ✅ | ✅ | ✅ | ✅ | ⚠️ 部分 |
| 公式 | ✅ | ✅ | ⚠️ 部分 | ⚠️ 部分 | ⚠️ 部分 |
| 批注 | ✅ | ✅ | ✅ | ✅ | ⚠️ 仅查看 |
| Sparklines | ✅ | ✅ | ❌ | ❌ | ❌ |
| 表格对象 | ✅ | ✅ | ✅ | ✅ | ⚠️ 部分 |
| 数据条/色阶/图标集 | ✅ | ✅ | ⚠️ 部分 | ⚠️ 部分 | ⚠️ 部分 |
| VBA 宏 | ✅ | ✅ | ✅ | ⚠️ 部分 | ❌ |

### 50.2 Excel 版本特性

| 特性 | 最低版本 | 说明 |
|------|---------|------|
| 基本 OOXML | 2007 | 所有基础功能 |
| 条件格式增强 | 2010 | 数据条、色阶、图标集 |
| Sparklines | 2010 | 迷你图 |
| 切片器 | 2010 | 数据筛选 |
| 快速填充 | 2013 | 智能数据填充 |
| 新图表类型 | 2016 | 瀑布图、树形图等 |
| 3D 地图 | 2016 | 3D 数据可视化 |
| 动态数组 | 365 | SORT, FILTER 等函数 |

### 50.3 兼容性建议

| 目标 | 建议 |
|------|------|
| 最大兼容性 | 只用基本功能：单元格、合并、样式、数字格式 |
| 包含图表 | 确保同时提供数据缓存(strCache/numCache) |
| 包含条件格式 | 用最基本的 cellIs + dxf，避免高级特性 |
| 包含数据验证 | 用 list 和 decimal 类型，避免 custom |
| 跨平台使用 | 避免使用 VBA、Sparklines、OLAP 等专有特性 |

### 50.4 生成文件时的版本标记

```xml
<!-- workbook.xml 中 -->
<workbook xmlns="...">
  <!-- fileVersion 声明文件版本 -->
  <fileVersion appName="xl" lastEdited="6" lowestEdited="6" 
               rupBuild="14420"/>
  <!-- lastEdited: 最后编辑的版本号 -->
  <!-- lowestEdited: 能打开此文件的最低版本 -->
</workbook>
```

| lastEdited | 版本 |
|-----------|------|
| 3 | Excel 2007 |
| 4 | Excel 2010 |
| 5 | Excel 2013 |
| 6 | Excel 2016+ |

---

## 51. 测试与验证策略

### 51.1 验证方法

| 方法 | 说明 | 工具 |
|------|------|------|
| 解压检查 | 解压 ZIP，检查文件结构是否完整 | 7-Zip, unzip |
| XML 校验 | 检查 XML 是否合法 | 浏览器, xmllint |
| Excel 打开 | 用 Excel 打开，查看是否有修复提示 | Excel, WPS |
| 对比标准文件 | 与 Excel 生成的文件逐文件对比 | diff, Beyond Compare |
| 自动化测试 | 批量生成+验证 | 脚本 |

### 51.2 自动化验证脚本(Python)

```python
import zipfile
import xml.etree.ElementTree as ET
import os

def validate_xlsx(filepath):
    """验证 xlsx 文件的基本完整性"""
    errors = []
    
    # 1. 检查是否能解压
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            namelist = z.namelist()
            
            # 2. 检查必需文件
            required = [
                '[Content_Types].xml',
                '_rels/.rels',
                'xl/workbook.xml',
                'xl/_rels/workbook.xml.rels',
                'xl/worksheets/sheet1.xml',
            ]
            for req in required:
                if req not in namelist:
                    errors.append(f"缺少必需文件: {req}")
            
            # 3. 检查每个 XML 文件是否合法
            for name in namelist:
                if name.endswith('.xml') or name.endswith('.rels'):
                    try:
                        content = z.read(name)
                        ET.fromstring(content)
                    except ET.ParseError as e:
                        errors.append(f"XML 解析错误 ({name}): {e}")
            
            # 4. 检查 ZIP 完整性
            bad = z.testzip()
            if bad:
                errors.append(f"ZIP 损坏: {bad}")
                
    except zipfile.BadZipFile:
        errors.append("不是合法的 ZIP 文件")
    
    return errors

# 使用
errors = validate_xlsx("report.xlsx")
if errors:
    print("❌ 验证失败:")
    for e in errors:
        print(f"  - {e}")
else:
    print("✅ 验证通过")
```

### 51.3 测试用例设计

| 测试类别 | 测试点 | 验证方法 |
|---------|--------|---------|
| 基本功能 | 空文件能打开 | Excel 打开无报错 |
| 基本功能 | 单个单元格有值 | 值正确显示 |
| 基本功能 | 数字/字符串/日期 | 类型正确 |
| 合并 | 水平/垂直/区域合并 | 显示正确 |
| 合并 | 合并区域边框完整 | 四周都有边框 |
| 合并 | 交叉合并检测 | 应报错 |
| 样式 | 字体/颜色/大小 | 显示正确 |
| 样式 | 数字格式 | 显示正确 |
| 样式 | 样式索引边界 | s=0 和 s=max 都不报错 |
| 大数据 | 最大行数 | 不崩溃 |
| 大数据 | 最大单元格数 | 不崩溃 |
| 特殊字符 | 含 < > & " ' 的文本 | 正确转义 |
| 特殊字符 | 中文/Unicode | 正确显示 |
| 边界 | 空字符串 | 不报错 |
| 边界 | 超长字符串 | 截断或报错 |
| 边界 | 空工作表 | 能打开 |

### 51.4 持续集成测试

```c
// 测试框架示例
void test_basic_xlsx() {
    workbook_t wb;
    memset(&wb, 0, sizeof(wb));
    strcpy(wb.sheet_names[0], "Test");
    
    ws_set_string(&wb.sheets[0], &wb.sst, 1, 1, "Hello", 0);
    ws_set_number(&wb.sheets[0], 1, 2, 3.14, 0);
    
    generate_xlsx(&wb, "test_basic.xlsx");
    
    // 验证
    assert(file_exists("test_basic.xlsx"));
    assert(file_size("test_basic.xlsx") > 0);
    
    // 用 Python 脚本验证
    system("python validate.py test_basic.xlsx");
}

void test_merge_cells() {
    workbook_t wb;
    memset(&wb, 0, sizeof(wb));
    
    ws_set_string(&wb.sheets[0], &wb.sst, 1, 1, "Merged", 1);
    ws_set_empty(&wb.sheets[0], 1, 2, 1);
    ws_add_merge(&wb.sheets[0], 1, 1, 1, 2);
    
    generate_xlsx(&wb, "test_merge.xlsx");
    
    // 验证 XML 中有 mergeCell
    // ...
}

void test_special_chars() {
    workbook_t wb;
    memset(&wb, 0, sizeof(wb));
    
    ws_set_string(&wb.sheets[0], &wb.sst, 1, 1, "A<B & C>D", 0);
    ws_set_string(&wb.sheets[0], &wb.sst, 1, 2, "\"引号\"", 0);
    ws_set_string(&wb.sheets[0], &wb.sst, 1, 3, "中文测试", 0);
    
    generate_xlsx(&wb, "test_special.xlsx");
    
    // 验证 XML 中字符已正确转义
    // ...
}
```

---

## 52. 内存缓冲区管理进阶

### 52.1 静态分配 vs 动态分配

| 方式 | 优点 | 缺点 | 适用场景 |
|------|------|------|---------|
| 静态分配 | 无碎片、确定性强 | 大小固定、浪费内存 | MCU、RTOS |
| 动态分配 | 灵活、不浪费 | 碎片风险、不确定 | Linux、大内存系统 |
| 内存池 | 折中 | 实现复杂 | 中等资源系统 |

### 52.2 静态分配方案

```c
// 所有数据结构静态分配，编译时确定大小
typedef struct {
    // 数据存储（静态数组）
    cell_t cells[MAX_CELLS];
    int cell_count;
    
    merge_range_t merges[MAX_MERGES];
    int merge_count;
    
    // 共享字符串
    char sst_strings[MAX_SHARED_STRINGS][MAX_STRING_LEN];
    int sst_count;
    
    // 输出缓冲区
    char xml_buffer[XML_BUFFER_SIZE];
    int xml_pos;
    
} excel_static_t;

// 全局单例（适合单线程）
static excel_static_t g_excel;

// 或作为局部变量（注意栈大小）
void generate_report() {
    excel_static_t excel;  // 确保栈足够大
    // ...
}
```

### 52.3 内存池方案

```c
// 自定义内存池
#define POOL_SIZE  (32 * 1024)  // 32KB
static uint8_t memory_pool[POOL_SIZE];
static size_t pool_offset = 0;

void *pool_alloc(size_t size) {
    // 对齐到 4 字节
    size = (size + 3) & ~3;
    if (pool_offset + size > POOL_SIZE) return NULL;
    void *ptr = &memory_pool[pool_offset];
    pool_offset += size;
    return ptr;
}

void pool_reset(void) {
    pool_offset = 0;  // 重置内存池，"释放"所有内存
}

// 使用
void generate_report() {
    pool_reset();  // 重置到起始位置
    
    workbook_t *wb = pool_alloc(sizeof(workbook_t));
    // ... 使用 wb ...
    
    // 不需要 free，下次 pool_reset 时自动回收
}
```

### 52.4 环形缓冲区（流式输出）

```c
// 用于流式输出，避免大缓冲区
typedef struct {
    uint8_t buf[OUTPUT_BUF_SIZE];  // 小缓冲区，如 512 字节
    int head;
    int tail;
    excel_write_fn write_fn;
    void *write_ctx;
} ring_buffer_t;

void rb_flush(ring_buffer_t *rb) {
    while (rb->tail != rb->head) {
        int chunk = (rb->head - rb->tail + OUTPUT_BUF_SIZE) % OUTPUT_BUF_SIZE;
        if (chunk > OUTPUT_BUF_SIZE - rb->tail) {
            chunk = OUTPUT_BUF_SIZE - rb->tail;
        }
        rb->write_fn(&rb->buf[rb->tail], chunk, rb->write_ctx);
        rb->tail = (rb->tail + chunk) % OUTPUT_BUF_SIZE;
    }
}

void rb_write(ring_buffer_t *rb, const void *data, size_t size) {
    const uint8_t *p = data;
    for (size_t i = 0; i < size; i++) {
        rb->buf[rb->head] = p[i];
        rb->head = (rb->head + 1) % OUTPUT_BUF_SIZE;
        if (rb->head == rb->tail) {
            rb_flush(rb);  // 缓冲区满，自动刷新
        }
    }
}

void rb_printf(ring_buffer_t *rb, const char *fmt, ...) {
    char temp[256];
    va_list ap;
    va_start(ap, fmt);
    int n = vsnprintf(temp, sizeof(temp), fmt, ap);
    va_end(ap);
    rb_write(rb, temp, n);
}
```

### 52.5 内存用量配置表

| 配置等级 | RAM 需求 | 功能限制 |
|---------|---------|---------|
| 极简（8KB RAM） | ~4 KB | 1 Sheet, 50 单元格, 无合并, 内联字符串 |
| 基础（32KB RAM） | ~12 KB | 1 Sheet, 500 单元格, 10 合并, 共享字符串 |
| 标准（128KB RAM） | ~48 KB | 4 Sheet, 5000 单元格, 100 合并, 完整样式 |
| 高级（256KB+ RAM） | ~128 KB | 多 Sheet, 无限单元格, 图表, 条件格式 |

---

## 53. 数据类型完整参考

### 53.1 单元格类型(t属性)完整列表

| t 值 | 类型 | `<v>` 内容 | 说明 |
|------|------|-----------|------|
| `n`（默认） | 数字 | 数值文本 | 整数或浮点数 |
| `s` | 共享字符串 | 索引号 | 指向 sharedStrings.xml |
| `str` | 公式结果字符串 | 文本 | 公式计算结果是文本时 |
| `inlineStr` | 内联字符串 | 无（用 `<is>`） | 直接在单元格中写文本 |
| `b` | 布尔 | 0 或 1 | 0=false, 1=true |
| `e` | 错误 | 错误字符串 | 如 `#N/A`, `#DIV/0!` |
| `d` | 日期 | ISO 8601 日期 | 如 `2024-03-15T14:30:00`（较少使用） |

### 53.2 内联字符串的两种写法

```xml
<!-- 方式 A：推荐 -->
<c r="A1" t="inlineStr"><is><t>文本内容</t></is></c>

<!-- 方式 B：带格式 -->
<c r="A1" t="inlineStr">
  <is>
    <r>
      <rPr><b/><color rgb="FFFF0000"/></rPr>
      <t>红色粗体文本</t>
    </r>
  </is>
</c>
```

### 53.3 数字类型的特殊情况

```xml
<!-- 整数 -->
<c r="A1"><v>42</v></c>

<!-- 浮点数 -->
<c r="A2"><v>3.14159</v></c>

<!-- 负数 -->
<c r="A3"><v>-273.15</v></c>

<!-- 科学计数法 -->
<c r="A4"><v>6.022E23</v></c>

<!-- 零 -->
<c r="A5"><v>0</v></c>

<!-- 负零 -->
<c r="A6"><v>-0</v></c>

<!-- 极大数 -->
<c r="A7"><v>1.7976931348623157E308</v></c>

<!-- 极小数 -->
<c r="A8"><v>4.9E-324</v></c>
```

### 53.4 空单元格的三种状态

```xml
<!-- 状态 A：单元格不存在（XML 中没有对应的 <c> 元素） -->
<!-- 效果：空单元格，默认样式 -->

<!-- 状态 B：单元格存在但无值 -->
<c r="A1"/>                         <!-- 无样式无值 -->
<c r="A1" s="2"/>                   <!-- 有样式无值（用于被合并的格子） -->

<!-- 状态 C：单元格存在且值为空字符串 -->
<c r="A1" t="s"><v></v></c>         <!-- 共享字符串空项 -->
<c r="A1" t="inlineStr"><is><t></t></is></c>  <!-- 内联空字符串 -->
```

---

## 54. 完整文件清单总表

### 54.1 按功能分类的文件清单

| 功能 | 需要的文件 | 说明 |
|------|-----------|------|
| **基本 xlsx** | `[Content_Types].xml` | 必需 |
| | `_rels/.rels` | 必需 |
| | `xl/workbook.xml` | 必需 |
| | `xl/_rels/workbook.xml.rels` | 必需 |
| | `xl/worksheets/sheetN.xml` | 必需 |
| **样式** | `xl/styles.xml` | 可选但强烈推荐 |
| **共享字符串** | `xl/sharedStrings.xml` | 可选（可用内联替代） |
| **文档属性** | `docProps/core.xml` | 可选 |
| | `docProps/app.xml` | 可选 |
| **自定义属性** | `docProps/custom.xml` | 可选 |
| **图表** | `xl/charts/chartN.xml` | 按需 |
| | `xl/drawings/drawingN.xml` | 按需 |
| | `xl/drawings/_rels/drawingN.xml.rels` | 按需 |
| | `xl/worksheets/_rels/sheetN.xml.rels` | 按需 |
| **图片** | `xl/media/imageN.xxx` | 按需 |
| | `xl/drawings/drawingN.xml` | 同图表 |
| **批注** | `xl/comments/sheetN.xml` | 按需 |
| | `xl/drawings/vmlDrawingN.vml` | 按需 |
| **表格对象** | `xl/tables/tableN.xml` | 按需 |
| **Sparklines** | `xl/sparklines/sparklineN.xml` | 按需 |
| **外部链接** | `xl/externalLinks/externalLinkN.xml` | 按需 |
| | `xl/externalLinks/_rels/externalLinkN.xml.rels` | 按需 |
| **VBA 宏** | `xl/vbaProject.bin` | 仅 .xlsm |
| **VBA 签名** | `xl/vbaProjectSignature.bin` | 仅签名宏 |

### 54.2 按优先级的实现路线图

```
第一阶段（最小可用）:
  ✅ [Content_Types].xml
  ✅ _rels/.rels
  ✅ xl/workbook.xml
  ✅ xl/_rels/workbook.xml.rels
  ✅ xl/worksheets/sheet1.xml
  ✅ 最小 styles.xml（1个默认样式）
  ✅ 内联字符串（不使用 sharedStrings）

第二阶段（基础功能）:
  ✅ xl/sharedStrings.xml
  ✅ 完整 styles.xml（多种样式）
  ✅ 合并单元格
  ✅ 列宽/行高
  ✅ 数字格式
  ✅ 多 Sheet

第三阶段（增强功能）:
  ✅ 数据验证
  ✅ 自动筛选
  ✅ 条件格式
  ✅ 冻结窗格
  ✅ 打印设置
  ✅ 页眉页脚
  ✅ 分组大纲
  ✅ 工作表保护

第四阶段（高级功能）:
  ✅ 图表
  ✅ 图片插入
  ✅ 批注
  ✅ 超链接
  ✅ 表格对象
  ✅ 命名区域
  ✅ Sparklines
  ✅ 富文本

第五阶段（专业功能）:
  ✅ 自定义文档属性
  ✅ 外部链接
  ✅ Deflate 压缩
  ✅ VBA 宏(.xlsm)
  ✅ 文件加密
  ✅ 分页符
  ✅ Sheet Tab 颜色
```

### 54.3 各阶段的内存需求

| 阶段 | ROM (代码) | RAM (运行时) | 文件大小 |
|------|-----------|-------------|---------|
| 第一阶段 | ~4 KB | ~2 KB | ~3 KB |
| 第二阶段 | ~8 KB | ~12 KB | ~10 KB |
| 第三阶段 | ~16 KB | ~32 KB | ~15 KB |
| 第四阶段 | ~32 KB | ~64 KB | ~30 KB |
| 第五阶段 | ~48 KB | ~128 KB | ~50 KB |

### 54.4 ZIP 文件条目统计

| 功能组合 | ZIP 文件条目数 |
|---------|--------------|
| 最小 xlsx（5个核心文件） | 5 |
| + 样式 + 共享字符串 | 7 |
| + 文档属性 | 9 |
| + 多 Sheet（3个） | 11 |
| + 图表（1个） | 14 |
| + 图片（1张） | 16 |
| + 批注 | 18 |
| + 表格对象 | 19 |
| 完整功能 | 20-30 |

> **注意**：ZIP 写入器需要支持足够的文件条目数（建议至少 32 个）。


---

# 终极篇：全覆盖格式与高级特性

> 前面的章节已经覆盖了 Excel 库的核心功能。本篇补齐**所有剩余格式和特性**，从多种文件格式输出、图形对象、表单控件到数据透视表、切片器、动态数组公式等，确保不遗漏任何标准 Excel 库应具备的能力。

---

## 目录（终极篇）

- [55. 全部文件输出格式](#55-全部文件输出格式)
- [56. CSV / TSV 格式](#56-csv--tsv-格式)
- [57. HTML 表格导出](#57-html-表格导出)
- [58. SpreadsheetML 2003 XML 格式](#58-spreadsheetml-2003-xml-格式)
- [59. ODS (OpenDocument) 格式](#59-ods-opendocument-格式)
- [60. .xlsb 二进制格式概述](#60-xlsb-二进制格式概述)
- [61. .xls (BIFF8) 旧二进制格式概述](#61-xls-biff8-旧二进制格式概述)
- [62. 绘图形状（Shape）](#62-绘图形状shape)
- [63. 文本框与标注](#63-文本框与标注)
- [64. 表单控件](#64-表单控件)
- [65. ActiveX 控件](#65-activex-控件)
- [66. 数据透视表(Pivot Table)](#66-数据透视表pivot-table)
- [67. 切片器(Slicer)](#67-切片器slicer)
- [68. 时间线(Timeline)](#68-时间线timeline)
- [69. 图表工作表(Chart Sheet)](#69-图表工作表chart-sheet)
- [70. 线程批注(Threaded Comments)](#70-线程批注threaded-comments)
- [71. 自定义 XML 部件(Custom XML Parts)](#71-自定义-xml-部件custom-xml-parts)
- [72. 外部数据连接与查询表](#72-外部数据连接与查询表)
- [73. 计算链与工作簿计算属性](#73-计算链与工作簿计算属性)
- [74. 假设分析方案(Scenarios)](#74-假设分析方案scenarios)
- [75. 拼音信息(Phonetic)](#75-拼音信息phonetic)
- [76. 富数据类型(Rich Data Types)](#76-富数据类型rich-data-types)
- [77. 3D 引用与跨表公式](#77-3d-引用与跨表公式)
- [78. 动态数组与溢出范围](#78-动态数组与溢出范围)
- [79. LAMBDA 与自定义函数](#79-lambda-与自定义函数)
- [80. OLE 对象与嵌入内容](#80-ole-对象与嵌入内容)
- [81. 智能标记(Smart Tags)](#81-智能标记smart-tags)
- [82. 旧式绘图(Legacy Drawing)](#82-旧式绘图legacy-drawing)
- [83. 墨迹注释(Ink Annotations)](#83-墨迹注释ink-annotations)
- [84. Web 发布设置](#84-web-发布设置)
- [85. 文件恢复信息](#85-文件恢复信息)
- [86. 工作簿元数据与未来元数据](#86-工作簿元数据与未来元数据)
- [87. 单元格批注完整版（旧式+线程式）](#87-单元格批注完整版旧式线程式)
- [88. 完整数字格式令牌参考](#88-完整数字格式令牌参考)
- [89. 完整样式属性参考](#89-完整样式属性参考)
- [90. 工作簿结构完整 XML Schema](#90-工作簿结构完整-xml-schema)
- [91. 工作表元素排序规则（关键！）](#91-工作表元素排序规则关键)
- [92. 嵌入式全格式实现路线图](#92-嵌入式全格式实现路线图)
- [93. ark_excel：总体架构与文件结构](#93-ark_excel总体架构与文件结构)
- [94. 两遍法流式 ZIP：对 UART 也正确的零缓冲输出](#94-两遍法流式-zip对-uart-也正确的零缓冲输出)
- [95. 数据模型：静态池 + arena 字符串](#95-数据模型静态池--arena-字符串)
- [96. XML 生成器：模板直拼 + 手写数字转换](#96-xml-生成器模板直拼--手写数字转换)
- [97. 样式表：去重注册 + 索引复用](#97-样式表去重注册--索引复用)
- [98. 可移植性：输出回调与 svcrtos 集成](#98-可移植性输出回调与-svcrtos-集成)
- [99. 测试体系（tests/ 目录）](#99-测试体系tests-目录)
- [100. 内存与性能](#100-内存与性能)
- [101. API 速查](#101-api-速查)
- [102. 已验证特性矩阵](#102-已验证特性矩阵)
- [103. 读改写：打开/修改/保存现有 xlsx](#103-读改写打开修改保存现有-xlsx)
- [104. 功能裁剪：用 config.h 应对不同资源的嵌入式设备](#104-功能裁剪用-configh-应对不同资源的嵌入式设备)

---

## 55. 全部文件输出格式

### 55.1 格式全景表

| 格式 | 扩展名 | 本质 | 嵌入式可行性 | 典型大小 |
|------|--------|------|------------|---------|
| Office Open XML | `.xlsx` | ZIP + XML | ★★★★★ 完全可行 | 5-50 KB |
| 含宏 OOXML | `.xlsm` | ZIP + XML + VBA bin | ★★★☆☆ VBA 二进制难构造 | 10-100 KB |
| 模板 | `.xltx` | ZIP + XML（同 xlsx，ContentType 不同） | ★★★★★ | 5-50 KB |
| 含宏模板 | `.xltm` | ZIP + XML + VBA bin | ★★★☆☆ | 10-100 KB |
| 二进制 | `.xlsb` | ZIP + 二进制记录 | ★★☆☆☆ 二进制格式复杂 | 3-30 KB |
| 旧二进制 | `.xls` | OLE2 + BIFF8 | ★★☆☆☆ OLE2 容器复杂 | 10-100 KB |
| CSV | `.csv` | 纯文本 | ★★★★★ 最简单 | 最小 |
| TSV | `.tsv` | 纯文本（Tab 分隔） | ★★★★★ | 最小 |
| HTML | `.html` | HTML 表格 | ★★★★☆ | 中等 |
| SpreadsheetML | `.xml` | 纯 XML（2003 格式） | ★★★★☆ | 较大 |
| ODS | `.ods` | ZIP + XML（ODF 标准） | ★★★★☆ 类似 xlsx | 5-50 KB |
| 纯文本 | `.txt` / `.prn` | 纯文本（固定宽度） | ★★★★★ | 最小 |
| PDF | `.pdf` | PDF 格式 | ★★★☆☆ 需 PDF 库 | 中等 |

### 55.2 格式选择决策树

```
是否需要 Excel 公式？
├── 是 → 是否需要宏？
│   ├── 是 → .xlsm
│   └── 否 → .xlsx
└── 否 → 是否需要复杂数据（合并、样式）？
    ├── 是 → .xlsx
    └── 否 → 是否需要其他软件打开？
        ├── 是 → .ods (ODF标准) 或 .html
        └── 否 → .csv 或 .tsv (最简单)
```

### 55.3 嵌入式推荐格式优先级

```
第一优先：.xlsx (Store模式)  — 最大兼容性，无需额外库
第二优先：.csv               — 极简场景，数据量小
第三优先：.html              — 需要简单样式但没有 Excel
第四优先：.ods               — 需要 ODF 标准兼容
```

---

## 56. CSV / TSV 格式

### 56.1 CSV 格式规范

CSV（Comma-Separated Values）是最简单的表格格式，纯文本，每行一条记录，字段用逗号分隔。

### 56.2 CSV 编码规则

```
规则1：每行一条记录
规则2：字段用逗号分隔
规则3：含逗号、换行、双引号的字段必须用双引号包裹
规则4：字段内的双引号要转义为两个双引号 ""
规则5：空字段表示为 ,（两个逗号之间无内容）或 ""（带引号的空）
```

### 56.3 CSV 转义实现

```c
// 将一个字段写入 CSV，自动处理转义
void csv_write_field(FILE *f, const char *value) {
    if (value == NULL) {
        fprintf(f, ",");
        return;
    }
    
    // 检查是否需要引号
    int need_quote = 0;
    for (const char *p = value; *p; p++) {
        if (*p == ',' || *p == '"' || *p == '\n' || *p == '\r') {
            need_quote = 1;
            break;
        }
    }
    
    if (need_quote) {
        fputc('"', f);
        for (const char *p = value; *p; p++) {
            if (*p == '"') {
                fputc('"', f);  // 双引号转义为两个
                fputc('"', f);
            } else {
                fputc(*p, f);
            }
        }
        fputc('"', f);
    } else {
        fputs(value, f);
    }
    fputc(',', f);
}

// 写一行 CSV
void csv_write_row(FILE *f, const char **fields, int count) {
    for (int i = 0; i < count; i++) {
        csv_write_field(f, fields[i]);
    }
    // 去掉最后一个逗号，换行
    fseek(f, -1, SEEK_CUR);  // 回退一个字节（去掉最后逗号）
    fprintf(f, "\r\n");       // CSV 标准用 CRLF
}
```

### 56.4 BOM 处理

Excel 打开 UTF-8 编码的 CSV 时，如果没有 BOM（Byte Order Mark），中文会乱码。

```c
// 写入 UTF-8 BOM（3字节: EF BB BF）
void csv_write_bom(FILE *f) {
    fputc(0xEF, f);
    fputc(0xBB, f);
    fputc(0xBF, f);
}

// 使用
FILE *f = fopen("data.csv", "wb");
csv_write_bom(f);
csv_write_row(f, fields, count);
fclose(f);
```

### 56.5 CSV 分隔符变体

| 变体 | 分隔符 | 地区 | 说明 |
|------|--------|------|------|
| 标准 CSV | `,` (逗号) | 英文环境 | 最常见 |
| 分号 CSV | `;` (分号) | 欧洲/中文环境 | Excel 中文版默认 |
| TSV | `\t` (Tab) | 通用 | 粘贴友好 |
| 竖线 | `\|` | 自定义 | 某些数据库 |

### 56.6 嵌入式 CSV 生成器（完整）

```c
typedef struct {
    FILE *f;
    char delimiter;     // 分隔符
    int col_index;      // 当前列索引
} csv_writer_t;

void csv_init(csv_writer_t *w, FILE *f, char delimiter) {
    w->f = f;
    w->delimiter = delimiter;
    w->col_index = 0;
    // 写 BOM
    fputc(0xEF, f);
    fputc(0xBB, f);
    fputc(0xBF, f);
}

void csv_field_str(csv_writer_t *w, const char *value) {
    if (w->col_index > 0) fputc(w->delimiter, w->f);
    
    int need_quote = 0;
    for (const char *p = value; *p; p++) {
        if (*p == w->delimiter || *p == '"' || *p == '\n' || *p == '\r') {
            need_quote = 1; break;
        }
    }
    
    if (need_quote) {
        fputc('"', w->f);
        for (const char *p = value; *p; p++) {
            if (*p == '"') { fputc('"', w->f); fputc('"', w->f); }
            else fputc(*p, w->f);
        }
        fputc('"', w->f);
    } else {
        fputs(value, w->f);
    }
    w->col_index++;
}

void csv_field_num(csv_writer_t *w, double value) {
    char buf[32];
    snprintf(buf, sizeof(buf), "%.15g", value);
    csv_field_str(w, buf);
}

void csv_field_int(csv_writer_t *w, int value) {
    if (w->col_index > 0) fputc(w->delimiter, w->f);
    fprintf(w->f, "%d", value);
    w->col_index++;
}

void csv_newline(csv_writer_t *w) {
    fprintf(w->f, "\r\n");
    w->col_index = 0;
}
```

---

## 57. HTML 表格导出

### 57.1 基本 HTML 表格

```html
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid black; padding: 5px; text-align: left; }
th { background-color: #4CAF50; color: white; }
.merged { text-align: center; font-weight: bold; }
</style>
</head>
<body>
<table>
  <tr><th>姓名</th><th>年龄</th><th>城市</th></tr>
  <tr><td>张三</td><td>28</td><td>北京</td></tr>
  <tr><td>李四</td><td>35</td><td>上海</td></tr>
</table>
</body>
</html>
```

### 57.2 HTML 合并单元格

```html
<table border="1">
  <!-- colspan: 水平合并 -->
  <tr>
    <td colspan="2" class="merged">合并标题</td>
  </tr>
  <!-- rowspan: 垂直合并 -->
  <tr>
    <td rowspan="2">垂直合并</td>
    <td>B1</td>
  </tr>
  <tr>
    <td>B2</td>
  </tr>
</table>
```

### 57.3 嵌入式 HTML 生成器

```c
typedef struct {
    FILE *f;
    int in_row;
    int col_count;
} html_writer_t;

void html_init(html_writer_t *w, FILE *f) {
    w->f = f;
    w->in_row = 0;
    w->col_count = 0;
    fprintf(f,
        "<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"UTF-8\">\n"
        "<style>\n"
        "table{border-collapse:collapse;}\n"
        "th,td{border:1px solid #000;padding:4px 8px;}\n"
        "th{background:#4472C4;color:#fff;}\n"
        ".title{font-weight:bold;background:#D9E1F2;text-align:center;}\n"
        ".warn{color:red;font-weight:bold;}\n"
        "</style>\n</head>\n<body>\n<table>\n");
}

void html_row_begin(html_writer_t *w) {
    fprintf(w->f, "  <tr>\n");
    w->in_row = 1;
    w->col_count = 0;
}

void html_cell_str(html_writer_t *w, const char *value, const char *css_class) {
    fprintf(w->f, "    <td");
    if (css_class) fprintf(w->f, " class=\"%s\"", css_class);
    fprintf(w->f, ">");
    // HTML 转义
    for (const char *p = value; *p; p++) {
        switch (*p) {
            case '<': fprintf(w->f, "&lt;"); break;
            case '>': fprintf(w->f, "&gt;"); break;
            case '&': fprintf(w->f, "&amp;"); break;
            case '"': fprintf(w->f, "&quot;"); break;
            default: fputc(*p, w->f);
        }
    }
    fprintf(w->f, "</td>\n");
    w->col_count++;
}

// 合并单元格
void html_cell_merge(html_writer_t *w, const char *value, 
                      int colspan, int rowspan, const char *css_class) {
    fprintf(w->f, "    <td");
    if (colspan > 1) fprintf(w->f, " colspan=\"%d\"", colspan);
    if (rowspan > 1) fprintf(w->f, " rowspan=\"%d\"", rowspan);
    if (css_class) fprintf(w->f, " class=\"%s\"", css_class);
    fprintf(w->f, ">");
    for (const char *p = value; *p; p++) {
        switch (*p) {
            case '<': fprintf(w->f, "&lt;"); break;
            case '>': fprintf(w->f, "&gt;"); break;
            case '&': fprintf(w->f, "&amp;"); break;
            default: fputc(*p, w->f);
        }
    }
    fprintf(w->f, "</td>\n");
    w->col_count += colspan;
}

void html_row_end(html_writer_t *w) {
    fprintf(w->f, "  </tr>\n");
    w->in_row = 0;
}

void html_finish(html_writer_t *w) {
    fprintf(w->f, "</table>\n</body>\n</html>");
}
```

---

## 58. SpreadsheetML 2003 XML 格式

### 58.1 什么是 SpreadsheetML 2003

在 OOXML (.xlsx) 之前，微软曾用一种纯 XML 格式——SpreadsheetML 2003（也叫 "XML Spreadsheet"）。它的优势是**不需要 ZIP 容器**，整个文件就是一个 XML。

### 58.2 基本结构

```xml
<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
          xmlns:o="urn:schemas-microsoft-com:office:office"
          xmlns:x="urn:schemas-microsoft-com:office:excel"
          xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
          xmlns:html="http://www.w3.org/TR/REC-html40">
  
  <Styles>
    <Style ss:ID="Default" ss:Name="Normal">
      <Font ss:FontName="等线" x:CharSet="134" ss:Size="11"/>
      <Interior/>
      <Borders/>
    </Style>
    <Style ss:ID="Header">
      <Font ss:FontName="微软雅黑" ss:Size="12" ss:Bold="1"/>
      <Interior ss:Color="#4472C4" ss:Pattern="Solid"/>
      <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
      </Borders>
    </Style>
    <Style ss:ID="Data">
      <NumberFormat ss:Format="0.00"/>
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
      </Borders>
    </Style>
  </Styles>
  
  <Worksheet ss:Name="数据">
    <Table>
      <!-- 列宽 -->
      <Column ss:Width="100"/>
      <Column ss:Width="80"/>
      
      <Row>
        <Cell ss:StyleID="Header"><Data ss:Type="String">传感器</Data></Cell>
        <Cell ss:StyleID="Header"><Data ss:Type="String">温度</Data></Cell>
      </Row>
      <Row>
        <Cell ss:StyleID="Data"><Data ss:Type="String">传感器A</Data></Cell>
        <Cell ss:StyleID="Data"><Data ss:Type="Number">25.3</Data></Cell>
      </Row>
      <!-- 合并单元格 -->
      <Row>
        <Cell ss:StyleID="Header" ss:MergeAcross="2">
          <Data ss:Type="String">合并标题</Data>
        </Cell>
      </Row>
    </Table>
    
    <WorksheetOptions>
      <FreezePanes/>
      <FrozenNoSplit/>
      <SplitHorizontal>1</SplitHorizontal>
      <TopRowBottomPane>1</TopRowBottomPane>
    </WorksheetOptions>
  </Worksheet>
</Workbook>
```

### 58.3 SpreadsheetML 2003 的优劣

| 优点 | 缺点 |
|------|------|
| 不需要 ZIP 库 | 文件体积大（无压缩） |
| 单文件 XML，简单直观 | 微软专有格式，非 ISO 标准 |
| Excel 和 WPS 都支持 | 不支持现代特性（Sparklines等） |
| 嵌入式极友好 | 公式支持有限 |

### 58.4 数据类型

```xml
<Data ss:Type="String">文本</Data>
<Data ss:Type="Number">123.45</Data>
<Data ss:Type="Boolean">1</Data>
<Data ss:Type="DateTime">2024-03-15T14:30:00</Data>  <!-- ISO 8601 -->
<Data ss:Type="Error">#N/A</Data>
```

### 58.5 合并属性

```xml
<!-- MergeAcross: 向右合并N列 -->
<Cell ss:MergeAcross="2"><Data ss:Type="String">合并3列</Data></Cell>

<!-- MergeDown: 向下合并N行 -->
<Cell ss:MergeDown="1"><Data ss:Type="String">合并2行</Data></Cell>
```

### 58.6 嵌入式实现建议

如果不想实现 ZIP 打包，SpreadsheetML 2003 是**最简单的方案**：
- 只需一个 `fprintf` 生成 XML
- 不需要 CRC32、不需要 ZIP 头
- Excel 双击直接打开
- 文件后缀 `.xml`，但可以用 `<?mso-application progid="Excel.Sheet"?>` 指令让 Excel 识别

---

## 59. ODS (OpenDocument) 格式

### 59.1 ODS 与 XLSX 的对比

| 特性 | XLSX (OOXML) | ODS (ODF) |
|------|-------------|-----------|
| 标准 | ISO/IEC 29500 | ISO/IEC 26300 |
| 容器 | ZIP | ZIP |
| 内容 | XML | XML |
| 主命名空间 | `...spreadsheetml/2006/main` | `urn:oasis:names:tc:opendocument:xmlns:...` |
| 兼容软件 | Excel, WPS, LibreOffice | LibreOffice, OpenOffice, Google Docs |

### 59.2 ODS 文件结构

```
report.ods (ZIP)
├── mimetype                  ← 固定内容（必须放在ZIP第一个条目，Store模式）
├── META-INF/
│   ├── manifest.xml          ← 类似 Content_Types
│   └── manifest.rdf          ← RDF 元数据
├── content.xml               ← 工作表数据+样式
├── styles.xml                ← 全局样式
├── meta.xml                  ← 文档属性
├── settings.xml             ← 应用设置
└── Thumbnails/
    └── thumbnail.png         ← 缩略图（可选）
```

### 59.3 mimetype 文件

```
application/vnd.oasis.opendocument.spreadsheet
```

> **关键**：mimetype 文件必须是 ZIP 中的第一个条目，且**必须用 Store 模式（不压缩）**。这是 ODS 标准的要求，用于让操作系统通过文件头部判断类型。

### 59.4 content.xml 结构

```xml
<?xml version="1.0" encoding="UTF-8"?>
<office:document-content 
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  office:version="1.2">
  
  <office:automatic-styles>
    <!-- 列样式 -->
    <style:style style:name="co1" style:family="table-column">
      <style:table-column-properties style:column-width="3cm"/>
    </style:style>
    <!-- 单元格样式 -->
    <style:style style:name="ce1" style:family="table-cell" 
                 style:parent-style-name="Default">
      <style:text-properties fo:font-weight="bold"/>
      <style:table-cell-properties fo:background-color="#4472C4"/>
    </style:style>
  </office:automatic-styles>
  
  <office:body>
    <office:spreadsheet>
      <table:table table:name="数据">
        <!-- 列定义 -->
        <table:table-column table:style-name="co1" 
                          table:default-cell-style-name="Default"/>
        
        <table:table-row>
          <table:table-cell table:style-name="ce1" office:value-type="string">
            <text:p>传感器</text:p>
          </table:table-cell>
          <table:table-cell table:style-name="ce1" office:value-type="string">
            <text:p>温度</text:p>
          </table:table-cell>
        </table:table-row>
        
        <table:table-row>
          <table:table-cell office:value-type="string">
            <text:p>传感器A</text:p>
          </table:table-cell>
          <table:table-cell office:value-type="float" office:value="25.3">
            <text:p>25.3</text:p>
          </table:table-cell>
        </table:table-row>
        
        <!-- 合并单元格 -->
        <table:table-row>
          <table:table-cell table:number-columns-repeated="2" 
                            table:style-name="ce1" office:value-type="string">
            <text:p>合并标题</text:p>
          </table:table-cell>
        </table:table-row>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
```

### 59.5 ODS 合并单元格

```xml
<!-- 水平合并（跨N列） -->
<table:table-cell table:number-columns-repeated="3" office:value-type="string">
  <text:p>跨3列</text:p>
</table:table-cell>

<!-- 垂直合并（跨N行） -->
<table:table-cell table:number-rows-repeated="2" office:value-type="string">
  <text:p>跨2行</text:p>
</table:table-cell>

<!-- 指定合并范围 -->
<table:table-cell table:number-matrix-columns-spanned="3" 
                  table:number-matrix-rows-spanned="2"
                  office:value-type="string">
  <text:p>3列2行合并</text:p>
</table:table-cell>
```

### 59.6 ODS 数据类型

| office:value-type | 说明 | office:value 属性 |
|-------------------|------|-------------------|
| `string` | 字符串 | 无（文本在 `<text:p>` 中） |
| `float` | 数字 | `office:value="123.45"` |
| `boolean` | 布尔 | `office:boolean-value="true"` |
| `date` | 日期 | `office:date-value="2024-03-15"` |
| `time` | 时间 | `office:time-value="PT14H30M00S"` |
| `percentage` | 百分比 | `office:value="0.25"` |
| `currency` | 货币 | `office:value="99.50" office:currency="CNY"` |

### 59.7 META-INF/manifest.xml

```xml
<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest 
  xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
  manifest:version="1.2">
  <manifest:file-entry manifest:full-path="/" 
                       manifest:version="1.2"
                       manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
  <manifest:file-entry manifest:full-path="content.xml"
                       manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="styles.xml"
                       manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml"
                       manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="settings.xml"
                       manifest:media-type="text/xml"/>
</manifest:manifest>
```

---

## 60. .xlsb 二进制格式概述

### 60.1 什么是 .xlsb

.xlsb（Excel Binary Workbook）是 Excel 的二进制格式，使用 ZIP 容器但内部用二进制记录替代 XML。优点是文件更小、读写更快。

### 60.2 文件结构

```
report.xlsb (ZIP)
├── [Content_Types].xml
├── _rels/.rels
├── docProps/
│   ├── app.xml
│   └── core.xml
├── xl/
│   ├── workbook.bin
│   ├── _rels/workbook.bin.rels
│   ├── sharedStrings.bin
│   ├── styles.bin
│   └── worksheets/
│       └── sheet1.bin
```

### 60.3 二进制记录格式

.xlsb 内部使用 BIFF12（Binary Interchange File Format 12）记录结构：

```
每个记录:
┌──────────────┐
│ 记录类型 (2B) │  ← 标识记录种类
├──────────────┤
│ 记录长度 (4B) │  ← 数据部分字节数
├──────────────┤
│ 数据 (变长)   │  ← 二进制编码的数据
└──────────────┘
```

### 60.4 主要记录类型

| 记录类型 ID | 名称 | 说明 |
|------------|------|------|
| 0x0080 | BrBeginSheet | Sheet 开始 |
| 0x0081 | BrEndSheet | Sheet 结束 |
| 0x0087 | BrBeginRow | 行开始 |
| 0x0088 | BrEndRow | 行结束 |
| 0x0089 | BrCellBlank | 空单元格 |
| 0x008A | BrCellBool | 布尔单元格 |
| 0x008B | BrCellError | 错误单元格 |
| 0x008C | BrCellFloat | 浮点单元格 |
| 0x008D | BrCellString | 字符串单元格 |
| 0x008E | BrCellSt | 共享字符串单元格 |

### 60.5 嵌入式可行性

| 方面 | 评估 |
|------|------|
| 二进制编码 | ★★★★☆ 比纯文本更省空间 |
| 记录结构 | ★★★☆☆ 需实现 BIFF12 读写器 |
| 兼容性 | ★★★☆☆ 仅 Excel 支持，WPS/LibreOffice 部分支持 |
| 推荐度 | ★★☆☆☆ 不推荐，实现成本高于 xlsx |

> **结论**：除非有极端的文件大小要求，否则不推荐在嵌入式系统中实现 .xlsb。xlsx 的 Store 模式已经足够紧凑。

---

## 61. .xls (BIFF8) 旧二进制格式概述

### 61.1 BIFF8 格式特点

| 特性 | 说明 |
|------|------|
| 容器 | OLE2 复合文档（不是 ZIP） |
| 内部 | BIFF8 二进制记录流 |
| 最大行数 | 65536 |
| 最大列数 | 256 |
| 兼容性 | Excel 97-2003，所有现代软件都支持读取 |

### 61.2 OLE2 复合文档结构

```
report.xls (OLE2)
├── Workbook               ← BIFF8 记录流（主数据）
├── SummaryInformation      ← 属性摘要
└── DocumentSummaryInformation ← 文档属性
```

OLE2 文件结构复杂：
- 扇区（Sector）为单位分配存储
- FAT（File Allocation Table）管理扇区链
- Directory 管理内部文件列表
- Mini-Stream 管理小文件

### 61.3 BIFF8 记录类型

| 类型 ID | 名称 | 说明 |
|---------|------|------|
| 0x0080 | BOF | 文件/Sheet 开始 |
| 0x000A | EOF | 文件/Sheet 结束 |
| 0x0085 | BOUNDSHEET | Sheet 定义 |
| 0x0020 | DIMENSION | 数据范围 |
| 0x020B | ROW | 行定义 |
| 0x00FD | LABELSST | 共享字符串单元格 |
| 0x0203 | NUMBER | 数字单元格 |
| 0x0006 | FORMULA | 公式单元格 |
| 0x00E0 | MERGEDCELLS | 合并单元格 |
| 0x0293 | STYLE | 样式 |
| 0x041E | FORMAT | 数字格式 |
| 0x0031 | FONT | 字体 |
| 0x00E5 | XF | 单元格格式 |

### 61.4 嵌入式实现难度

| 难点 | 说明 |
|------|------|
| OLE2 容器 | 需实现 FAT、Directory、扇区管理，约 2000 行代码 |
| BIFF8 记录 | 二进制编码/解码，约 1500 行代码 |
| 最大收益 | 文件体积可能比 xlsx Store 小，但比 xlsx Deflate 大 |

> **结论**：不推荐在嵌入式系统中实现 .xls。OLE2 容器的复杂度远超 ZIP，且收益有限。如果需要与旧系统兼容，建议在 PC 端做格式转换。

---

## 62. 绘图形状（Shape）

### 62.1 什么是 Shape

Excel 支持在工作表上放置各种形状（Shape），包括：
- 矩形、圆角矩形
- 椭圆、圆形
- 箭头（各种方向）
- 三角形、菱形
- 流程图符号
- 星形、标注
- 自由形状

### 62.2 文件结构

```
xl/
├── drawings/
│   ├── drawing1.xml           ← 形状定义
│   └── _rels/
│       └── drawing1.xml.rels   ← 关系（如需要）
├── worksheets/
│   └── sheet1.xml
└── _rels/
    └── sheet1.xml.rels          ← sheet → drawing 关系
```

### 62.3 drawing1.xml：矩形示例

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from>
      <xdr:col>2</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>2</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>5</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>5</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    
    <!-- 矩形形状 -->
    <xdr:sp macro="" textlink="">
      <xdr:nvSpPr>
        <xdr:cNvPr id="3" name="矩形 1"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
        </a:xfrm>
        <!-- 形状几何 -->
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
        <!-- 填充色 -->
        <a:solidFill>
          <a:srgbClr val="4472C4"/>
        </a:solidFill>
        <!-- 边框 -->
        <a:ln w="9525">
          <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
      </xdr:spPr>
      
      <!-- 形状内文本 -->
      <xdr:txBody>
        <a:bodyPr vertOverflow="clip" horzOverflow="clip"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r>
            <a:rPr lang="zh-CN" sz="1200" b="1">
              <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
            </a:rPr>
            <a:t>重要提示</a:t>
          </a:r>
        </a:p>
      </xdr:txBody>
    </xdr:sp>
    
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
```

### 62.4 预设形状(prst)完整列表

| prst 值 | 形状 | prst 值 | 形状 |
|---------|------|---------|------|
| `rect` | 矩形 | `roundRect` | 圆角矩形 |
| `ellipse` | 椭圆 | `triangle` | 三角形 |
| `rtTriangle` | 直角三角形 | `diamond` | 菱形 |
| `pentagon` | 五边形 | `hexagon` | 六边形 |
| `heptagon` | 七边形 | `octagon` | 八边形 |
| `star4` | 四角星 | `star5` | 五角星 |
| `star6` | 六角星 | `star8` | 八角星 |
| `heart` | 心形 | `lightningBolt` | 闪电 |
| `sun` | 太阳 | `moon` | 月亮 |
| `cloud` | 云形 | `arrow` | 箭头 |
| `leftArrow` | 左箭头 | `rightArrow` | 右箭头 |
| `upArrow` | 上箭头 | `downArrow` | 下箭头 |
| `leftRightArrow` | 左右箭头 | `upDownArrow` | 上下箭头 |
| `bentArrow` | 弯曲箭头 | `curvedUpArrow` | 曲线上箭头 |
| `curvedDownArrow` | 曲线下箭头 | `stripedRightArrow` | 条纹右箭头 |
| `notchedRightArrow` | 凹口右箭头 | `homePlate` | 五角大楼形 |
| `chevron` | V形 | `rightArrowCallout` | 右箭头标注 |
| `leftArrowCallout` | 左箭头标注 | `upArrowCallout` | 上箭头标注 |
| `downArrowCallout` | 下箭头标注 | `flowChartProcess` | 流程图处理 |
| `flowChartDecision` | 流程图决策 | `flowChartTerminator` | 流程图终止 |
| `flowChartConnector` | 流程图连接 | `flowChartDocument` | 流程图文档 |
| `flowChartMagneticDisk` | 流程图磁盘 | `flowChartOnlineStorage` | 流程图存储 |
| `wedgeRectCallout` | 矩形标注 | `wedgeRoundRectCallout` | 圆角矩形标注 |
| `wedgeEllipseCallout` | 椭圆标注 | `cloudCallout` | 云形标注 |
| `borderCallout1` | 边框标注1 | `borderCallout2` | 边框标注2 |
| `roundedRectangularCallout` | 圆角矩形标注 | `ovalCallout` | 椭圆标注 |
| `rectCallout` | 矩形标注 | `line` | 直线 |
| `lineInv` | 反向直线 | `plus` | 加号 |
| `plaque` | 牌匾 | `can` | 圆柱 |
| `cube` | 立方体 | `bevel` | 斜面 |
| `frame` | 框架 | `halfFrame` | 半框 |
| `corner` | 角 | `bracePair` | 大括号对 |
| `bracketPair` | 方括号对 | `blockArc` | 弧形 |
| `pie` | 扇形 | `donut` | 圆环 |
| `noSmoking` | 禁止符号 | `paperclip` | 回形针 |

### 62.5 连接线(Connector)

```xml
<!-- 连接两个形状的箭头线 -->
<xdr:cxnSp>
  <xdr:nvCxnSpPr>
    <xdr:cNvPr id="4" name="连接线 1"/>
    <xdr:cNvCxnSpPr/>
  </xdr:nvCxnSpPr>
  <xdr:spPr>
    <a:xfrm>
      <a:off x="1000000" y="500000"/>
      <a:ext cx="2000000" cy="0"/>
    </a:xfrm>
    <a:prstGeom prst="line">
      <a:avLst/>
    </a:prstGeom>
    <a:ln w="12700">
      <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
      <!-- 箭头 -->
      <a:tailEnd type="triangle" w="med" len="med"/>
    </a:ln>
  </xdr:spPr>
</xdr:cxnSp>
```

---

## 63. 文本框与标注

### 63.1 文本框

文本框是一种特殊的 Shape，主要用来显示文字：

```xml
<xdr:sp macro="" textlink="">
  <xdr:nvSpPr>
    <xdr:cNvPr id="5" name="文本框 1"/>
    <xdr:cNvSpPr txBox="1"/>     ← 标识为文本框
  </xdr:nvSpPr>
  <xdr:spPr>
    <a:xfrm>
      <a:off x="0" y="0"/>
      <a:ext cx="0" cy="0"/>
    </a:xfrm>
    <a:prstGeom prst="rect">
      <a:avLst/>
    </a:prstGeom>
    <a:noFill/>                   ← 无填充色
    <a:ln>
      <a:noFill/>                 ← 无边框
    </a:ln>
  </xdr:spPr>
  <xdr:txBody>
    <a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="l"/>
      <a:r>
        <a:rPr lang="zh-CN" sz="1100"/>
        <a:t>这是一段说明文字。</a:t>
      </a:r>
    </a:p>
    <a:p>
      <a:pPr algn="l"/>
      <a:r>
        <a:rPr lang="zh-CN" sz="1100" b="1">
          <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
        </a:rPr>
        <a:t>注意：红色粗体</a:t>
      </a:r>
    </a:p>
  </xdr:txBody>
</xdr:sp>
```

### 63.2 标注(Callout)

标注是一种带"指向尾巴"的形状，常用于添加说明：

```xml
<xdr:sp macro="" textlink="">
  <xdr:nvSpPr>
    <xdr:cNvPr id="6" name="标注 1"/>
    <xdr:cNvSpPr/>
  </xdr:nvSpPr>
  <xdr:spPr>
    <a:xfrm>
      <a:off x="0" y="0"/>
      <a:ext cx="0" cy="0"/>
    </a:xfrm>
    <!-- 圆角矩形标注 -->
    <a:prstGeom prst="wedgeRoundRectCallout">
      <a:avLst>
        <!-- 调整尾巴位置和大小 -->
        <a:gd name="adj1" fmla="val 50000"/>   <!-- 尾巴水平位置 -->
        <a:gd name="adj2" fmla="val 50000"/>   <!-- 尾巴垂直位置 -->
      </a:avLst>
    </a:prstGeom>
    <a:solidFill>
      <a:srgbClr val="FFF2CC"/>                <!-- 浅黄色背景 -->
    </a:solidFill>
    <a:ln w="9525">
      <a:solidFill><a:srgbClr val="BF8F00"/></a:solidFill>
    </a:ln>
  </xdr:spPr>
  <xdr:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="zh-CN" sz="1100"/>
        <a:t>异常值：请检查</a:t>
      </a:r>
    </a:p>
  </xdr:txBody>
</xdr:sp>
```

### 63.3 文本框中的富文本段落

```xml
<xdr:txBody>
  <a:bodyPr wrap="square" anchor="t"/>    <!-- anchor: t=top, ctr=center, b=bottom -->
  <a:lstStyle/>
  
  <!-- 段落1：标题（粗体大字居中） -->
  <a:p>
    <a:pPr algn="ctr"/>
    <a:r>
      <a:rPr lang="zh-CN" sz="1600" b="1">
        <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
        <a:latin typeface="微软雅黑"/>
      </a:rPr>
      <a:t>数据采集报告</a:t>
    </a:r>
  </a:p>
  
  <!-- 段落2：正文 -->
  <a:p>
    <a:pPr algn="l" indent="0"/>
    <a:r>
      <a:rPr lang="zh-CN" sz="1100"/>
      <a:t>采集时间：2024-03-15 14:30</a:t>
    </a:r>
  </a:p>
  
  <!-- 段落3：带项目符号 -->
  <a:p>
    <a:pPr marL="19050" indent="-19050">
      <a:buFont typeface="Arial"/>
      <a:buChar char="•"/>
    </a:pPr>
    <a:r>
      <a:rPr lang="zh-CN" sz="1100"/>
      <a:t>传感器A：正常</a:t>
    </a:r>
  </a:p>
  
  <!-- 段落4：带颜色标记 -->
  <a:p>
    <a:pPr marL="19050" indent="-19050">
      <a:buFont typeface="Arial"/>
      <a:buChar char="•"/>
    </a:pPr>
    <a:r>
      <a:rPr lang="zh-CN" sz="1100">
        <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
      </a:rPr>
      <a:t>传感器C：异常！</a:t>
    </a:r>
  </a:p>
</xdr:txBody>
```

---

## 64. 表单控件

### 64.1 表单控件类型

Excel 支持两种控件系统：
- **表单控件(Form Controls)**：老式控件，简单、兼容性好
- **ActiveX 控件**：新式控件，功能强大但复杂

### 64.2 表单控件的文件结构

表单控件需要 VML 文件来定义位置和外观：

```
xl/
├── worksheets/
│   └── sheet1.xml
├── ctrlProps/
│   ├── ctrlProp1.xml        ← 控件属性
│   └── ctrlProp2.xml
├── drawings/
│   ├── vmlDrawing1.vml      ← 控件的位置和外观（VML格式）
│   └── drawing1.xml         ← ActiveX控件使用DrawingML
└── _rels/
    └── sheet1.xml.rels
```

### 64.3 按钮控件

```xml
<!-- vmlDrawing1.vml 中的按钮 -->
<v:shape id="_x0000_s1025" type="#_x0000_t201" 
         style="position:absolute;margin-left:100pt;margin-top:50pt;
                width:80pt;height:25pt;z-index:1" 
         o:button="t" fillcolor="#4472C4" strokecolor="#FFFFFF">
  <v:fill color2="#4472C4" o:detectmouseclick="t"/>
  <v:textbox style="mso-direction-alt:auto" inset="2pt,2pt,2pt,2pt">
    <div style="text-align:center">
      <font face="微软雅黑" size="3" color="#FFFFFF">
        <b>导出数据</b>
      </font>
    </div>
  </v:textbox>
  <x:ClientData ObjectType="Button">
    <x:Anchor>2, 0, 1, 0, 4, 0, 3, 0</x:Anchor>
    <x:PrintObject>True</x:PrintObject>
    <x:AutoFill>False</x:AutoFill>
  </x:ClientData>
</v:shape>
```

### 64.4 复选框控件

```xml
<v:shape id="_x0000_s1026" type="#_x0000_t201" 
         style="position:absolute;margin-left:50pt;margin-top:80pt;
                width:100pt;height:20pt;z-index:2"
         o:button="t">
  <v:textbox>
    <div style="text-align:left">
      <font face="等线" size="2">启用报警</font>
    </div>
  </v:textbox>
  <x:ClientData ObjectType="Checkbox">
    <x:Anchor>1, 0, 3, 0, 3, 0, 4, 0</x:Anchor>
    <x:Checked>1</x:Checked>              <!-- 0=未选中, 1=选中, 2=混合 -->
    <x:FmlaLink>Sheet1!$A$10</x:FmlaLink>  <!-- 关联的单元格 -->
    <x:PrintObject>True</x:PrintObject>
  </x:ClientData>
</v:shape>
```

### 64.5 下拉框(Combobox)控件

```xml
<v:shape id="_x0000_s1027" type="#_x0000_t201" 
         style="position:absolute;margin-left:200pt;margin-top:80pt;
                width:80pt;height:18pt;z-index:3">
  <x:ClientData ObjectType="Drop">
    <x:Anchor>4, 0, 3, 0, 6, 0, 4, 0</x:Anchor>
    <x:FmlaLink>Sheet1!$B$10</x:FmlaLink>           <!-- 输出单元格 -->
    <x:FmlaRange>Sheet1!$F$1:$F$5</x:FmlaRange>     <!-- 数据源 -->
    <x:Sel>0</x:Sel>                                  <!-- 默认选中索引 -->
    <x:NoThreeD2>True</x:NoThreeD2>
    <x:PrintObject>True</x:PrintObject>
  </x:ClientData>
</v:shape>
```

### 64.6 微调按钮(Spinner)控件

```xml
<v:shape id="_x0000_s1028" type="#_x0000_t201" 
         style="position:absolute;margin-left:300pt;margin-top:80pt;
                width:50pt;height:18pt;z-index:4">
  <x:ClientData ObjectType="Spin">
    <x:Anchor>6, 0, 3, 0, 7, 0, 4, 0</x:Anchor>
    <x:FmlaLink>Sheet1!$C$10</x:FmlaLink>    <!-- 关联单元格 -->
    <x:Min>0</x:Min>                          <!-- 最小值 -->
    <x:Max>100</x:Max>                         <!-- 最大值 -->
    <x:Inc>1</x:Inc>                           <!-- 步进值 -->
    <x:PrintObject>True</x:PrintObject>
  </x:ClientData>
</v:shape>
```

### 64.7 表单控件完整列表

| ObjectType | 说明 | 特有属性 |
|------------|------|---------|
| `Button` | 按钮 | `Macro`(宏名) |
| `Checkbox` | 复选框 | `Checked`, `FmlaLink` |
| `Radio` | 单选按钮 | `Checked`, `FmlaLink`, `FmlaRange` |
| `Drop` | 下拉框 | `FmlaLink`, `FmlaRange`, `Sel` |
| `List` | 列表框 | `FmlaLink`, `FmlaRange`, `SelType` |
| `Label` | 标签 | — |
| `Spin` | 微调按钮 | `FmlaLink`, `Min`, `Max`, `Inc`, `Page` |
| `Scroll` | 滚动条 | `FmlaLink`, `Min`, `Max`, `Inc`, `Page` |
| `GBox` | 分组框 | — |
| `EditBox` | 编辑框 | `FmlaLink` |

### 64.8 控件关联单元格的运作方式

```
用户操作控件 → 控件自动更新关联单元格的值 → 其他单元格通过公式引用
```

例如：复选框关联 A10，A10 会显示 TRUE/FALSE：
```xml
<!-- C列公式根据A10的值显示不同内容 -->
<c r="C1"><f>IF(A10,"报警已启用","报警已禁用")</f><v>报警已启用</v></c>
```

---

## 65. ActiveX 控件

### 65.1 ActiveX 与表单控件的对比

| 特性 | 表单控件 | ActiveX 控件 |
|------|---------|-------------|
| 格式 | VML | DrawingML + 二进制 |
| 事件支持 | 有限（需宏） | 完整（Click, Change等） |
| 外观定制 | 有限 | 丰富 |
| 兼容性 | 好 | 仅 Windows Excel |
| 嵌入式推荐 | ★★★☆☆ | ★☆☆☆☆ |

### 65.2 文件结构

```
xl/
├── activeX/
│   ├── activeX1.bin          ← ActiveX 控件二进制数据
│   ├── activeX2.bin
│   └── _rels/
│       └── activeX1.xml.rels
├── drawings/
│   └── drawing1.xml          ← DrawingML 定义位置
└── ctrlProps/
    └── ctrlProp1.xml         ← 控件属性
```

> **嵌入式建议**：ActiveX 控件涉及二进制 ActiveX 格式，实现极其复杂，不推荐在嵌入式系统中使用。如需交互控件，使用表单控件。

---

## 66. 数据透视表(Pivot Table)

### 66.1 什么是数据透视表

数据透视表是一种交互式报表，可以动态汇总、分析大量数据。它从源数据中提取信息，按行/列分组聚合。

### 66.2 文件结构

```
xl/
├── pivotTables/
│   ├── pivotTable1.xml         ← 透视表定义
│   ├── pivotCacheDefinition1.xml ← 透视缓存定义（数据源信息）
│   ├── pivotCacheRecords1.xml    ← 透视缓存记录（缓存的数据）
│   └── _rels/
│       └── pivotTable1.xml.rels
├── worksheets/
│   └── sheet1.xml
└── _rels/
    └── sheet1.xml.rels          ← sheet → pivotTable 关系
```

### 66.3 pivotCacheDefinition1.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                       xmlns:r="..."
                       r:id="rId1"
                       refreshOnLoad="1"
                       refreshedBy="System"
                       refreshedDate="45366.60417"
                       createdVersion="6"
                       refreshedVersion="6"
                       minRefreshableVersion="3"
                       recordCount="100">
  
  <!-- 数据源定义 -->
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:D100" sheet="数据"/>
  </cacheSource>
  
  <!-- 缓存字段定义 -->
  <cacheFields count="4">
    <cacheField name="传感器" numFmtId="0">
      <sharedItems count="3">
        <s v="传感器A"/>
        <s v="传感器B"/>
        <s v="传感器C"/>
      </sharedItems>
    </cacheField>
    <cacheField name="通道" numFmtId="0">
      <sharedItems count="2">
        <s v="通道1"/>
        <s v="通道2"/>
      </sharedItems>
    </cacheField>
    <cacheField name="数值" numFmtId="0">
      <sharedItems containsSemiMixedTypes="0" containsString="0" 
                   containsNumber="1" containsInteger="0" 
                   minValue="20" maxValue="85" 
                   count="0"/>
    </cacheField>
    <cacheField name="状态" numFmtId="0">
      <sharedItems count="3">
        <s v="正常"/>
        <s v="警告"/>
        <s v="故障"/>
      </sharedItems>
    </cacheField>
  </cacheFields>
</pivotCacheDefinition>
```

### 66.4 pivotTable1.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="..." name="数据透视表1" 
                       cacheId="1" 
                       applyNumberFormats="0" applyBorderFormats="0"
                       applyFontFormats="0" applyPatternFormats="0"
                       applyAlignmentFormats="0" applyWidthHeightFormats="1"
                       dataCaption="值" 
                       updatedVersion="6" 
                       minRefreshableVersion="3"
                       useAutoFormatting="1" 
                       itemPrintTitles="1" 
                       createdVersion="6"
                       indent="0" outline="1" outlineData="1" 
                       multipleFieldFilters="0">
  
  <!-- 位置 -->
  <location ref="A3:D20" firstHeaderRow="1" firstDataRow="2" 
            firstDataCol="1" rowPageCount="1" colPageCount="1"/>
  
  <!-- 字段列表 -->
  <pivotFields count="4">
    <!-- 字段0：传感器 → 放在行区域 -->
    <pivotField axis="axisRow" showAll="0">
      <items>
        <item x="0"/>     <!-- 传感器A -->
        <item x="1"/>     <!-- 传感器B -->
        <item x="2"/>     <!-- 传感器C -->
        <item t="default"/>
      </items>
    </pivotField>
    
    <!-- 字段1：通道 → 放在列区域 -->
    <pivotField axis="axisCol" showAll="0">
      <items>
        <item x="0"/>
        <item x="1"/>
        <item t="default"/>
      </items>
    </pivotField>
    
    <!-- 字段2：数值 → 放在数据区域 -->
    <pivotField dataField="1" showAll="0"/>
    
    <!-- 字段3：状态 → 不使用 -->
    <pivotField showAll="0"/>
  </pivotFields>
  
  <!-- 行字段 -->
  <rowFields count="1">
    <field x="0"/>
  </rowFields>
  
  <!-- 列字段 -->
  <colFields count="1">
    <field x="1"/>
  </colFields>
  
  <!-- 数据字段 -->
  <dataFields count="1">
    <dataField name="求和项:数值" fld="2" subtotal="sum" 
               baseField="0" baseItem="0"/>
  </dataFields>
  
  <!-- 格式 -->
  <pivotTableStyleInfo name="PivotStyleLight16" 
                       showRowHeaders="1" showColHeaders="1"
                       showRowStripes="0" showColStripes="0" 
                       showLastColumn="1"/>
</pivotTableDefinition>
```

### 66.5 数据字段汇总方式(subtotal)

| subtotal 值 | 说明 |
|-------------|------|
| `sum` | 求和 |
| `count` | 计数 |
| `average` | 平均值 |
| `max` | 最大值 |
| `min` | 最小值 |
| `product` | 乘积 |
| `countNums` | 数字计数 |
| `stdDev` | 标准差(样本) |
| `stdDevp` | 标准差(总体) |
| `var` | 方差(样本) |
| `varp` | 方差(总体) |

### 66.6 透视表轴(axis)

| axis 值 | 说明 |
|---------|------|
| `axisRow` | 行区域 |
| `axisCol` | 列区域 |
| `axisPage` | 筛选区域（页字段） |
| `axisValues` | 数据区域 |

### 66.7 嵌入式可行性

| 方面 | 评估 |
|------|------|
| XML 生成 | ★★★☆☆ 格式复杂但可模板化 |
| 数据预计算 | ★★★★★ 嵌入式可直接写缓存数据 |
| 实时刷新 | ★★☆☆☆ 需要 Excel 端刷新 |
| 推荐 | ★★★☆☆ 如需动态报表，值得实现 |

> **嵌入式策略**：预计算所有聚合结果，写入 pivotCacheRecords，Excel 打开后 `refreshOnLoad="1"` 会自动刷新。

---

## 67. 切片器(Slicer)

### 67.1 什么是切片器

切片器是一种可视化的数据筛选器，通常配合数据透视表或表格对象使用，提供按钮式快速筛选。

### 67.2 文件结构

```
xl/
├── slicers/
│   └── slicer1.xml             ← 切片器定义
├── slicerCaches/
│   └── slicerCache1.xml        ← 切片器缓存
├── drawings/
│   └── drawing1.xml            ← 切片器位置
├── worksheets/
│   └── sheet1.xml
└── _rels/
    └── sheet1.xml.rels
```

### 67.3 slicer1.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
         xmlns:r="...">
  <slicer name="传感器筛选器" cache="传感器" 
          caption="选择传感器" 
          startItem="0" columnCount="1"
          showCaption="1" level="0"
          style="SlicerStyleLight1"
          lockedPosition="0">
    <items>
      <item i="0" s="1"/>     <!-- 第0项，选中 -->
      <item i="1" s="0"/>     <!-- 第1项，未选中 -->
      <item i="2" s="1"/>     <!-- 第2项，选中 -->
    </items>
  </slicer>
</slicers>
```

### 67.4 slicerCache1.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
                        xmlns:r="..."
                        name="传感器"
                        sourceName="传感器"
                        showMissing="1"
                        sort="ascending"
                        crossFilter="showItemsWithNoData"
                        showNoDataItems="1">
  
  <!-- 数据源引用 -->
  <pivotTables>
    <pivotTable tabId="1" name="数据透视表1"/>
  </pivotTables>
  
  <!-- 切片器项 -->
  <data>
    <item v="传感器A"/>
    <item v="传感器B"/>
    <item v="传感器C"/>
  </data>
  
  <!-- 选项卡排序 -->
  <tabs>
    <tab showTabHeader="1"/>
  </tabs>
</slicerCacheDefinition>
```

---

## 68. 时间线(Timeline)

### 68.1 什么是时间线

时间线是专门用于日期字段筛选的切片器，提供类似时间轴的可视化界面。

### 68.2 文件结构

```
xl/
├── timelines/
│   └── timeline1.xml          ← 时间线定义
├── slicerCaches/
│   └── slicerCache1.xml       ← 共享切片器缓存
├── drawings/
│   └── drawing1.xml
└── worksheets/
    └── sheet1.xml
```

### 68.3 timeline1.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
           xmlns:r="...">
  <timeline name="时间线1" cache="日期缓存"
            caption="选择日期范围"
            level="days"
            selectionLevel="2"
            start="2024-01-01T00:00:00"
            end="2024-12-31T00:00:00"
            min="2024-01-01T00:00:00"
            max="2024-12-31T00:00:00"
            scrollingStart="2024-01-01T00:00:00"
            scrollingEnd="2024-12-31T00:00:00"
            showHeader="1"
            showSelectionLabel="1"
            showTimeLevel="1"
            showHorizontalScrollbar="1"
            style="TimeSlicerStyleLight1">
    <options showMissing="1"/>
    <extLst>
      <!-- 扩展列表 -->
    </extLst>
  </timeline>
</timelines>
```

### 68.4 时间线级别(level)

| level 值 | 说明 |
|---------|------|
| `years` | 按年 |
| `quarters` | 按季度 |
| `months` | 按月 |
| `days` | 按日 |

---

## 69. 图表工作表(Chart Sheet)

### 69.1 什么是图表工作表

普通图表是嵌入在工作表中的"浮动对象"。图表工作表是一种特殊的 Sheet，整个 Sheet 就是一个图表（占满整个页面）。

### 69.2 文件结构

```
xl/
├── chartsheets/
│   └── sheet2.xml             ← 图表工作表（不是 worksheets/）
├── _rels/
│   └── workbook.xml.rels      ← r:id 指向 chartsheets/sheet2.xml
└── charts/
    └── chart1.xml              ← 图表定义
```

### 69.3 workbook.xml 中的声明

```xml
<sheets>
  <sheet name="数据" sheetId="1" r:id="rId1"/>
  <sheet name="图表" sheetId="2" r:id="rId2"/>   <!-- 这是图表工作表 -->
</sheets>
```

> **注意**：在 workbook.xml 中，图表工作表和普通工作表用相同的 `<sheet>` 元素，区别在于 workbook.xml.rels 中 `r:id` 指向的文件路径不同（`worksheets/` vs `chartsheets/`）。

### 69.4 workbook.xml.rels

```xml
<Relationships xmlns="...">
  <Relationship Id="rId1" Type=".../worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type=".../chartsheet" 
                Target="chartsheets/sheet2.xml"/>   <!-- 注意 Type 不同 -->
</Relationships>
```

### 69.5 Content_Types 声明

```xml
<!-- 普通工作表 -->
<Override PartName="/xl/worksheets/sheet1.xml" 
          ContentType="...worksheet+xml"/>

<!-- 图表工作表（不同的 ContentType） -->
<Override PartName="/xl/chartsheets/sheet2.xml" 
          ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>
```

### 69.6 chartsheets/sheet2.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="...">
  <!-- 图表工作表没有 sheetData，只有 drawing 引用 -->
  <sheetViews>
    <sheetView workbookViewId="0">
      <zoomScale scale="100"/>
    </sheetView>
  </sheetViews>
  
  <drawing r:id="rId1"/>     <!-- 指向 drawing1.xml -->
</chartsheet>
```

### 69.7 chartsheets/_rels/sheet2.xml.rels

```xml
<Relationships xmlns="...">
  <Relationship Id="rId1" Type=".../drawing" 
                Target="../drawings/drawing2.xml"/>
</Relationships>
```

---

## 70. 线程批注(Threaded Comments)

### 70.1 旧式批注 vs 线程批注

| 特性 | 旧式批注(Legacy Comment) | 线程批注(Threaded Comment) |
|------|------------------------|--------------------------|
| 引入版本 | Excel 2007 | Excel 2019 / Microsoft 365 |
| 格式 | comments/*.xml + VML | threadheedComments/*.xml |
| 线程 | 不支持回复 | 支持多级回复 |
| 显示 | 黄色批注框 | 右侧面板 |
| 人员 | 仅作者 | 多人参与讨论 |

### 70.2 文件结构

```
xl/
├── threadedComments/
│   └── sheet1.xml              ← 线程批注定义
├── people/
│   └── people1.xml             ← 参与者信息
├── worksheets/
│   └── sheet1.xml              ← 包含单元格引用标记
└── _rels/
    └── sheet1.xml.rels         ← sheet → threadedComments 关系
```

### 70.3 people/people1.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<people xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
         xmlns:r="...">
  <person displayName="系统" providerId="None" id="0"/>
  <person displayName="管理员" providerId="None" id="1"/>
</people>
```

### 70.4 threadedComments/sheet1.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<threadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
                  xmlns:r="...">
  
  <!-- 线程1：A1单元格 -->
  <threadedComment ref="A1" personId="0" id="{GUID-1}" done="0">
    <text>此数值为异常读数，需要排查。</text>
    <creationTime>2024-03-15T14:30:00Z</creationTime>
  </threadedComment>
  
  <!-- 回复 -->
  <threadedComment ref="A1" personId="1" id="{GUID-2}" 
                   parentId="{GUID-1}" done="0">
    <text>已确认，传感器C硬件故障，已更换。</text>
    <creationTime>2024-03-15T15:00:00Z</creationTime>
  </threadedComment>
  
  <!-- 线程2：B5单元格 -->
  <threadedComment ref="B5" personId="0" id="{GUID-3}" done="1">
    <text>此公式已验证正确。</text>
    <creationTime>2024-03-15T16:00:00Z</creationTime>
  </threadedComment>
</threadedComments>
```

### 70.5 sheet1.xml 中的引用

```xml
<!-- 单元格需要标记有线程批注 -->
<c r="A1" t="s" s="1">
  <v>0</v>
  <!-- th="1" 表示此单元格有线程批注 -->
</c>

<!-- 注意: th 属性加在 c 元素上 -->
<c r="A1" t="s" s="1" th="1"><v>0</v></c>
```

### 70.6 同时支持旧式和新式批注

现代 Excel 文件可以同时包含旧式批注（`comments/`）和线程批注（`threadedComments/`）。旧版 Excel 只读取旧式批注，新版 Excel 优先读取线程批注。

---

## 71. 自定义 XML 部件(Custom XML Parts)

### 71.1 什么是自定义 XML 部件

自定义 XML 部件允许在工作簿中存储任意 XML 数据，不显示在表格中，但可以被 VBA 或外部程序读取。

### 71.2 用途

- 存储嵌入式设备的配置信息
- 存储数据采集的元数据
- 存储版本控制信息
- 存储自定义数据结构

### 71.3 文件结构

```
xl/
├── customXml/
│   ├── item1.xml              ← 自定义 XML 数据
│   ├── itemProps1.xml         ← 自定义 XML 属性
│   └── _rels/
│       └── itemProps1.xml.rels
├── _rels/
│   └── workbook.xml.rels      ← workbook → customXml 关系
└── worksheets/
    └── sheet1.xml
```

### 71.4 customXml/item1.xml

```xml
<!-- 任意 XML 内容 -->
<?xml version="1.0" encoding="UTF-8"?>
<DeviceConfig>
  <Device>
    <ID>STM32_001</ID>
    <Firmware>v2.1.3</Firmware>
    <Location>生产线A-工位3</Location>
    <Sensors>
      <Sensor name="温度" type="PT100" channel="1"/>
      <Sensor name="湿度" type="DHT22" channel="2"/>
      <Sensor name="压力" type="BMP280" channel="3"/>
    </Sensors>
    <Calibration>
      <Offset temp="-0.3" humid="2.1" pressure="0.0"/>
    </Calibration>
    <ExportTime>2024-03-15T14:30:00</ExportTime>
  </Device>
</DeviceConfig>
```

### 71.5 customXml/itemProps1.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ds:datastoreItem ds:itemID="{GUID-HERE}"
  xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">
  <ds:schemaRefs>
    <ds:schemaRef ds:uri="http://mycompany.com/schemas/deviceconfig"/>
  </ds:schemaRefs>
</ds:datastoreItem>
```

---

## 72. 外部数据连接与查询表

### 72.1 什么是外部数据连接

Excel 可以连接到外部数据源（数据库、Web API、文本文件等），导入数据并定期刷新。

### 72.2 文件结构

```
xl/
├── connections.xml             ← 连接定义
├── _rels/
│   └── workbook.xml.rels       ← workbook → connections 关系
├── queryTables/
│   └── queryTable1.xml         ← 查询表定义
├── worksheets/
│   └── sheet1.xml              ← 包含查询表引用
└── _rels/
    └── sheet1.xml.rels
```

### 72.3 connections.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<connections xmlns="..." xmlns:r="...">
  <connection id="1" name="传感器数据" 
              type="3"              <!-- 3=Text/CSV, 1=ODBC, 2=OLE DB -->
              refreshedVersion="6" 
              minRefreshableVersion="3" 
              refresh="1" 
              saveData="1"
              background="0" 
              onlyUseConnectionFile="0">
    <dbPr connection="Data Source=\\device\\sensors" 
          command="SELECT * FROM readings" 
          commandType="2"/>  <!-- 1=SQL, 2=Table, 4=StoredProc -->
    <textPr characterSet="UTF-8" 
            delimiter=","
            firstRow="1"
            sourceFile="\\device\\data.csv">
      <textFields count="3">
        <textField type="text"/>
        <textField type="double"/>
        <textField type="date" dec="."/>
      </textFields>
    </textPr>
  </connection>
</connections>
```

### 72.4 queryTables/queryTable1.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<queryTable xmlns="..." xmlns:r="..." 
            name="传感器数据查询" 
            connectionId="1"
            autoFormatFont="1" 
            applyNumberFormats="1"
            applyBorderFormats="1" 
            applyPatternFormats="1"
            applyAlignmentFormats="1" 
            applyWidthHeightFormats="1"
            headers="1" 
            autoRefresh="0"
            refreshOnLoad="1"
            backgroundQuery="0" 
            preserveFormatting="1"
            adjustColumnWidth="1" 
            intermediate="0">
  <queryTableRefresh nextId="100" preserveFormatting="1" 
                     adjustColumnWidth="1" 
                     enableRefresh="1" 
                     refreshOnLoad="1">
    <queryTableFields count="3">
      <queryTableField id="1" name="传感器" tableColumnId="1"/>
      <queryTableField id="2" name="数值" tableColumnId="2"/>
      <queryTableField id="3" name="时间" tableColumnId="3"/>
    </queryTableFields>
  </queryTableRefresh>
</queryTable>
```

### 72.5 连接类型(type)

| type | 说明 |
|------|------|
| 1 | ODBC 数据源 |
| 2 | OLE DB 数据源 |
| 3 | 文本/CSV 文件 |
| 4 | Web 查询 |
| 5 | 数据检索服务 |
| 6 | XML 数据 |
| 7 | 数据库查询 |
| 12 | 内部数据源 |

### 72.6 嵌入式场景

嵌入式系统可以通过 CSV 连接让 Excel 自动导入设备上的数据文件：

```xml
<connection id="1" name="设备数据" type="3" refreshOnLoad="1">
  <textPr sourceFile="\\device\\export.csv" delimiter="," characterSet="UTF-8"/>
</connection>
```

---

## 73. 计算链与工作簿计算属性

### 73.1 工作簿计算属性

```xml
<!-- workbook.xml 中 -->
<workbook xmlns="...">
  <sheets>...</sheets>
  
  <!-- 计算属性 -->
  <calcPr calcId="0"               ← 计算引擎版本（0=使用默认）
          calcMode="auto"            ← auto=自动, manual=手动, autoNoTable=自动(不含表)
          fullCalcOnLoad="1"        ← 打开时全量重算
          refMode="A1"             ← A1引用模式(默认) 或 R1C1
          iterate="0"               ← 是否迭代计算
          iterateCount="100"        ← 迭代次数上限
          iterateDelta="0.001"      ← 迭代精度
          calcCompleted="1"         ← 上次计算是否完成
          calcOnSave="1"            ← 保存时重算
          concurrentCalc="0"        ← 是否多线程计算
          concurrentManualCount="0"
          forceFullCalc="0"/>
</workbook>
```

### 73.2 计算模式详解

| calcMode | 说明 |
|---------|------|
| `auto` | 自动计算（默认）。任何单元格变化都触发重算 |
| `autoNoTable` | 自动计算，但数据表除外 |
| `manual` | 手动计算。按 F9 才计算 |

### 73.3 迭代计算

用于处理循环引用：

```xml
<calcPr iterate="1" iterateCount="100" iterateDelta="0.001"/>
```

- `iterate="1"`：启用迭代计算
- `iterateCount`：最多迭代次数
- `iterateDelta`：精度阈值，变化量小于此值则停止

### 73.4 计算链(Calculation Chain)

计算链记录公式的计算顺序（依赖关系决定的先后顺序）：

```xml
<!-- xl/calcChain.xml（可选文件） -->
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <c r="C1" i="1"/>          <!-- C1 先算 -->
  <c r="C2" i="1"/>          <!-- C2 依赖C1，后算 -->
  <c r="C3" i="2" s="1"/>    <!-- C3 在另一个 Sheet -->
  <c r="D1" l="1"/>          <!-- D1 是最后一次计算 -->
</calcChain>
```

| 属性 | 说明 |
|------|------|
| `r` | 单元格地址 |
| `i` | 计算级别 |
| `s` | Sheet 索引（0-based） |
| `l` | 是否为最后计算的单元格 |

> **注意**：calcChain.xml 是可选的。Excel 打开文件时会自动重建计算链。嵌入式实现中可以省略此文件。

### 73.5 R1C1 引用模式

```xml
<calcPr refMode="R1C1"/>
```

在 R1C1 模式下，单元格地址用行列数字表示：
- `R1C1` = A1（第1行第1列）
- `R5C3` = C5（第5行第3列）
- `R[-1]C` = 当前单元格上方一行

---

## 74. 假设分析方案(Scenarios)

### 74.1 什么是方案

Excel 的"方案管理器"允许保存多组假设值，在不同方案间切换查看结果。

### 74.2 文件结构

```
xl/
├── scenarios/
│   └── sheet1.xml              ← 方案定义（按 Sheet 分文件）
├── worksheets/
│   └── sheet1.xml
└── _rels/
    └── sheet1.xml.rels
```

### 74.3 scenarios/sheet1.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<scenarios xmlns="..." 
           current="1"           ← 当前方案索引
           show="0">
  
  <!-- 方案1：乐观估计 -->
  <scenario name="乐观" 
            locked="1" 
            user="System"
            comment="乐观估计：温度+10%"
            count="3">
    <inputCells r="B2" val="55.0"/>     <!-- 传感器A温度 -->
    <inputCells r="B3" val="66.0"/>     <!-- 传感器B温度 -->
    <inputCells r="B4" val="93.7"/>     <!-- 传感器C温度 -->
  </scenario>
  
  <!-- 方案2：悲观估计 -->
  <scenario name="悲观" 
            locked="1" 
            user="System"
            comment="悲观估计：温度-10%"
            count="3">
    <inputCells r="B2" val="45.0"/>
    <inputCells r="B3" val="54.0"/>
    <inputCells r="B4" val="76.7"/>
  </scenario>
  
  <!-- 方案3：正常情况 -->
  <scenario name="正常" 
            locked="1" 
            user="System"
            comment="正常工作温度"
            count="3">
    <inputCells r="B2" val="50.0"/>
    <inputCells r="B3" val="60.0"/>
    <inputCells r="B4" val="85.0"/>
  </scenario>
</scenarios>
```

---

## 75. 拼音信息(Phonetic)

### 75.1 什么是拼音信息

在中文和日文 Excel 中，可以为单元格中的文字添加拼音/假名标注。显示时文字上方出现小字注音。

### 75.2 共享字符串中的拼音

```xml
<sst xmlns="...">
  <si>
    <t>传感器</t>
    <rPh sb="0" eb="3">           ← 拼音标注
      <t>chuán gǎn qì</t>         ← 拼音文本
    </rPh>
    <phoneticPr fontId="1" type="noConversion" alignment="left"/>
  </si>
</sst>
```

| 属性 | 说明 |
|------|------|
| `sb` | 起始字符位置（0-based） |
| `eb` | 结束字符位置 |
| `fontId` | 拼音字体索引（指向 styles.xml 中的字体） |
| `type` | 拼音类型：`noConversion`、`Hiragana`、`Katakana`、`HalfWidthKatakana` |
| `alignment` | 对齐方式：`left`、`center`、`distributed` |

### 75.3 多段拼音

```xml
<si>
  <t>数据采集系统</t>
  <rPh sb="0" eb="2"><t>shù jù</t></rPh>         <!-- "数据"的拼音 -->
  <rPh sb="2" eb="4"><t>cǎi jí</t></rPh>           <!-- "采集"的拼音 -->
  <rPh sb="4" eb="6"><t>xì tǒng</t></rPh>           <!-- "系统"的拼音 -->
  <phoneticPr fontId="1" type="noConversion" alignment="center"/>
</si>
```

---

## 76. 富数据类型(Rich Data Types)

### 76.1 什么是富数据类型

Excel 365 引入了"富数据类型"，允许单元格存储结构化数据（如股票信息、地理数据），显示时只显示摘要，但包含完整数据。

### 76.2 XML 结构

```xml
<!-- 单元格引用富数据类型 -->
<c r="A1" t="7" s="1">         <!-- t="7" 表示富数据类型 -->
  <v>{"Type":"Stock","Value":"AAPL","Data":{"Name":"Apple Inc.","Price":175.43,"Change":2.15}}</v>
  <vm v="7"/>                   <!-- 值元数据 -->
</c>
```

### 76.3 元数据文件

```
xl/
├── richData/
│   ├── richValueRel.xml       ← 富数据值关系
│   ├── rdrichvalue.xml        ← 富数据值定义
│   └── richValueRel.xml.rels
├── metadata.xml               ← 元数据定义
└── worksheets/
    └── sheet1.xml
```

### 76.4 嵌入式可行性

富数据类型是 Excel 365 专有特性，依赖云端数据源，不适合嵌入式实现。

---

## 77. 3D 引用与跨表公式

### 77.1 3D 引用语法

```xml
<!-- 引用多个Sheet的同一单元格 -->
<c r="A1"><f>SUM(Sheet1:Sheet3!B1)</f><v>300</v></c>

<!-- 引用多个Sheet的区域 -->
<c r="A2"><f>AVERAGE(Sheet1:Sheet3!A1:C10)</f><v>45.6</v></c>
```

### 77.2 跨表引用

```xml
<!-- 引用另一个Sheet的单元格 -->
<c r="A3"><f>Sheet2!B5</f><v>100</v></c>

<!-- 跨表VLOOKUP -->
<c r="A4"><f>VLOOKUP(A1,Sheet2!A1:D100,2,FALSE)</f><v>42</v></c>

<!-- 引用带空格的Sheet名（需用单引号） -->
<c r="A5"><f>'数据 Sheet'!A1</f><v>50</v></c>
```

### 77.3 INDIRECT 函数（动态引用）

```xml
<!-- 根据B1的值动态引用 -->
<!-- 如B1="Sheet2"，则引用 Sheet2!A1 -->
<c r="C1"><f>INDIRECT(B1&amp;"!A1")</f><v>100</v></c>
```

---

## 78. 动态数组与溢出范围

### 78.1 什么是动态数组

Excel 365 引入了"动态数组"（Dynamic Arrays），一个公式可以自动向相邻单元格"溢出"(Spill)多个结果值。

### 78.2 动态数组公式

```xml
<!-- SORT 函数：排序后溢出 -->
<c r="A1"><f t="array" ref="A1:A10">SORT(Sheet2!A1:A10)</f>
  <v>5</v>
</c>
<!-- A2-A10 是溢出的值 -->
<c r="A2"><v>8</v></c>
<c r="A3"><v>12</v></c>
<!-- ... -->

<!-- UNIQUE 函数：去重后溢出 -->
<c r="B1"><f t="array" ref="B1:B5">UNIQUE(Sheet2!B1:B100)</f>
  <v>正常</v>
</c>
<c r="B2"><v>警告</v></c>
<c r="B3"><v>故障</v></c>
```

### 78.3 动态数组函数列表

| 函数 | 说明 |
|------|------|
| `FILTER(range, include)` | 筛选满足条件的行 |
| `SORT(range, [index], [order])` | 排序 |
| `SORTBY(range, by_array)` | 按另一列排序 |
| `UNIQUE(range)` | 去重 |
| `SEQUENCE(rows, [cols], [start], [step])` | 生成序列 |
| `RANDARRAY(rows, cols)` | 生成随机数组 |
| `XLOOKUP(val, lookup, result)` | 增强版 VLOOKUP |
| `XMATCH(val, lookup)` | 增强版 MATCH |
| `LET(name, value, calc)` | 定义局部变量 |

### 78.4 溢出标记

```xml
<!-- 溢出区域的非起始单元格标记为 spill -->
<c r="A2" vm="1"><v>8</v></c>      <!-- vm="1" 表示元数据引用 -->
```

> **嵌入式注意**：动态数组是 Excel 365 特性，旧版 Excel 不支持。在嵌入式生成中，建议预计算所有值并写入每个单元格，而不是依赖动态数组。

---

## 79. LAMBDA 与自定义函数

### 79.1 LAMBDA 函数

Excel 365 允许用户使用 LAMBDA 定义自定义函数：

```xml
<!-- 定义：摄氏转华氏 = LAMBDA(c, c*9/5+32) -->
<!-- 使用：=LAMBDA(25, 25*9/5+32) -->

<c r="A1">
  <f>LAMBDA(c, c*9/5+32)(25)</f>
  <v>77</v>
</c>
```

### 79.2 命名 LAMBDA

在 workbook.xml 中定义命名 LAMBDA：

```xml
<definedNames>
  <definedName name="摄氏转华氏">
    LAMBDA(c, c*9/5+32)
  </definedName>
</definedNames>
```

然后在工作表中直接使用：

```xml
<c r="A1"><f>摄氏转华氏(25)</f><v>77</v></c>
```

### 79.3 递归 LAMBDA

```xml
<!-- 阶乘函数：=IF(n<=1, 1, n*FACT(n-1)) -->
<definedName name="FACT">
  LAMBDA(n, IF(n&lt;=1, 1, n*FACT(n-1)))
</definedName>
```

---

## 80. OLE 对象与嵌入内容

### 80.1 什么是 OLE 对象

OLE（Object Linking and Embedding）允许在 Excel 中嵌入其他应用程序的对象，如嵌入 Word 文档、PDF 等。

### 80.2 文件结构

```
xl/
├── embeddings/
│   └── oleObject1.bin         ← OLE 对象二进制数据
├── drawings/
│   └── drawing1.xml           ← 对象位置
├── oleObjects/
│   └── oleObject1.xml         ← OLE 对象定义
└── worksheets/
    └── sheet1.xml
```

### 80.3 oleObjects/oleObject1.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<oleObjects xmlns="..." xmlns:r="..." xmlns:mc="..."
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <oleObject shapeId="1025" r:id="rId1" progId="Word.Document.12"
             name="文档" showAsIcon="0"/>
</oleObjects>
```

> **嵌入式建议**：OLE 对象需要 OLE2 二进制格式，极其复杂，不推荐在嵌入式系统中实现。

---

## 81. 智能标记(Smart Tags)

### 81.1 什么是智能标记

智能标记是 Excel 2002-2010 的特性，识别特定文本模式并提供操作菜单。在新版 Excel 中已移除。

### 81.2 文件结构

```
xl/
├── smartTags/
│   └── sheet1.xml             ← 智能标记
└── worksheets/
    └── sheet1.xml
```

> **注意**：智能标记已被微软弃用，不建议在新文件中使用。

---

## 82. 旧式绘图(Legacy Drawing)

### 82.1 什么是旧式绘图

旧式绘图（Legacy Drawing）使用 VML 格式，用于：
- 批注的位置定义
- 表单控件的外观
- 旧式自选图形

### 82.2 文件结构

```
xl/
├── drawings/
│   └── vmlDrawing1.vml        ← VML 格式的旧式绘图
├── worksheets/
│   └── sheet1.xml
└── _rels/
    └── sheet1.xml.rels
```

### 82.3 sheet1.xml 中的引用

```xml
<worksheet xmlns="...">
  <sheetData>...</sheetData>
  
  <!-- 旧式绘图引用 -->
  <legacyDrawing r:id="rId5"/>
</worksheet>
```

### 82.4 VML 基本结构

```xml
<xml xmlns:v="urn:schemas-microsoft-com:vml"
     xmlns:o="urn:schemas-microsoft-com:office:office"
     xmlns:x="urn:schemas-microsoft-com:office:excel">
  
  <!-- 形状类型定义 -->
  <v:shapetype id="_x0000_t202" coordsize="21600,21600" 
               o:spt="202" path="m0,0l0,21600,21600,21600,21600,0xe">
    <v:stroke joinstyle="miter"/>
    <v:path gradientshapeok="t" o:connecttype="rect"/>
  </v:shapetype>
  
  <!-- 具体形状 -->
  <v:shape id="_x0000_s1025" type="#_x0000_t202" 
           style="position:absolute;
                  margin-left:60pt;margin-top:15pt;
                  width:120pt;height:60pt;
                  z-index:1;
                  visibility:hidden" 
           fillcolor="#ffffe1" 
           o:insetmode="auto">
    <v:fill color2="#ffffe1"/>
    <v:shadow on="t" color="black" obscured="t"/>
    <v:path o:connecttype="none"/>
    <v:textbox style="mso-direction-alt:auto">
      <div style="text-align:left"/>
    </v:textbox>
    <x:ClientData ObjectType="Note">
      <x:MoveWithCells/>
      <x:SizeWithCells/>
      <x:Anchor>1, 0, 0, 0, 3, 0, 4, 0</x:Anchor>
      <x:AutoFill>False</x:AutoFill>
    </x:ClientData>
  </v:shape>
</xml>
```

### 82.5 VML Anchor 格式

`<x:Anchor>` 的值由 8 个数字组成，用逗号分隔：

```
colLeft, dxLeft, rowTop, dyTop, colRight, dxRight, rowBottom, dyBottom
```

| 参数 | 说明 |
|------|------|
| `colLeft` | 起始列 (0-based) |
| `dxLeft` | 列内偏移 (像素) |
| `rowTop` | 起始行 (0-based) |
| `dyTop` | 行内偏移 (像素) |
| `colRight` | 结束列 (0-based) |
| `dxRight` | 列内偏移 |
| `rowBottom` | 结束行 (0-based) |
| `dyBottom` | 行内偏移 |

---

## 83. 墨迹注释(Ink Annotations)

### 83.1 什么是墨迹注释

墨迹注释是用户用触控笔或手指在 Excel 中手写的注释。在 Surface 等触摸设备上使用。

### 83.2 文件结构

```
xl/
├── drawings/
│   └── vmlDrawing1.vml       ← 墨迹以 VML 格式存储
├── ink/
│   └── ink1.xml               ← 墨迹数据（可选，MSInk 格式）
└── worksheets/
    └── sheet1.xml
```

> **嵌入式建议**：墨迹注释需要触摸屏和 MSInk 格式支持，不适合嵌入式系统。

---

## 84. Web 发布设置

### 84.1 文件结构

```
xl/
├── worksheets/
│   └── sheet1.xml
├── _rels/
│   └── sheet1.xml.rels        ← 包含 web 发布关系
```

### 84.2 sheet1.xml 中的 Web 发布

```xml
<worksheet xmlns="...">
  <sheetData>...</sheetData>
  
  <!-- Web 发布项 -->
  <webPublishItems count="1">
    <webPublishItem id="1" divId="RPT1" 
                    sourceType="range" 
                    sourceRef="A1:F50"
                    destinationFile="report.html"
                    title="传感器数据"
                    autoRepublish="0"/>
  </webPublishItems>
</worksheet>
```

---

## 85. 文件恢复信息

### 85.1 文件恢复

```xml
<!-- workbook.xml 中 -->
<workbook xmlns="...">
  <fileRecoveryPr autoRecover="0" 
                  crashSave="0" 
                  dataOnly="0" 
                  repairLoad="0"/>
</workbook>
```

| 属性 | 说明 |
|------|------|
| `autoRecover` | 是否启用了自动恢复 |
| `crashSave` | 是否因崩溃而保存的文件 |
| `dataOnly` | 是否只保存数据 |
| `repairLoad` | 是否在修复模式下加载 |

---

## 86. 工作簿元数据与未来元数据

### 86.1 元数据文件

```
xl/
├── metadata.xml               ← 元数据定义
├── _rels/
│   └── workbook.xml.rels      ← 包含 metadata 关系
└── worksheets/
    └── sheet1.xml
```

### 86.2 metadata.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">
  
  <!-- 元数据类型 -->
  <metadataTypes count="2">
    <!-- 类型1：单元格元数据 -->
    <metadataType name="XDA" 
                  minSupportedVersion="120000" 
                  copy="1" paste="1" pasteAll="1"
                  pasteValues="1" remove="1" 
                  assign="0" coerce="0"
                  cellMeta="1"/>
    
    <!-- 类型2：值元数据 -->
    <metadataType name="XLDAPROP" 
                  minSupportedVersion="120000"
                  copy="1" paste="1" pasteAll="1"
                  pasteValues="1" remove="0"
                  assign="1" coerce="1"
                  valueMeta="1"/>
  </metadataTypes>
  
  <!-- 元数据字符串 -->
  <futureMetadata name="XDA" count="1">
    <bk>
      <extLst>
        <ext uri="{GUID-HERE}" 
             xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">
          <xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  
  <!-- 单元格元数据 -->
  <cellMetadata count="1">
    <bk>
      <rc t="XDA" v="0"/>
    </bk>
  </cellMetadata>
  
  <!-- 值元数据 -->
  <valueMetadata count="1">
    <bk>
      <rc t="XLDAPROP" v="0"/>
    </bk>
  </valueMetadata>
</metadata>
```

---

## 87. 单元格批注完整版（旧式+线程式）

### 87.1 同时支持两种批注的完整文件结构

```
xl/
├── comments/
│   └── sheet1.xml              ← 旧式批注内容
├── threadedComments/
│   └── sheet1.xml              ← 线程批注内容
├── people/
│   └── people.xml              ← 参与者（仅线程批注）
├── drawings/
│   └── vmlDrawing1.vml         ← 旧式批注的显示位置
├── worksheets/
│   └── sheet1.xml
└── _rels/
    └── sheet1.xml.rels          ← 包含所有关系
```

### 87.2 sheet1.xml.rels（完整）

```xml
<Relationships xmlns="...">
  <!-- 样式 -->
  <Relationship Id="rId1" Type=".../styles" Target="styles.xml"/>
  <!-- 共享字符串 -->
  <Relationship Id="rId2" Type=".../sharedStrings" Target="sharedStrings.xml"/>
  <!-- 旧式批注 -->
  <Relationship Id="rId3" Type=".../comments" Target="../comments/sheet1.xml"/>
  <!-- 旧式绘图（批注位置） -->
  <Relationship Id="rId4" Type=".../vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
  <!-- 线程批注 -->
  <Relationship Id="rId5" Type=".../threadedComment" 
                Target="../threadedComments/sheet1.xml"/>
  <!-- 绘图（形状/图片/图表） -->
  <Relationship Id="rId6" Type=".../drawing" Target="../drawings/drawing1.xml"/>
  <!-- 表格对象 -->
  <Relationship Id="rId7" Type=".../table" Target="../tables/table1.xml"/>
  <!-- 数据透视表 -->
  <Relationship Id="rId8" Type=".../pivotTable" Target="../pivotTables/pivotTable1.xml"/>
  <!-- 超链接 -->
  <Relationship Id="rId9" Type=".../hyperlink" Target="https://..." TargetMode="External"/>
</Relationships>
```

### 87.3 Content_Types 完整声明

```xml
<Types xmlns="...">
  <!-- 默认类型 -->
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="gif" ContentType="image/gif"/>
  <Default Extension="bmp" ContentType="image/bmp"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  
  <!-- Override 类型 -->
  <Override PartName="/xl/workbook.xml" ContentType="...sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="...worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="...worksheet+xml"/>
  <Override PartName="/xl/chartsheets/sheet3.xml" ContentType="...chartsheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="...styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="...sharedStrings+xml"/>
  <Override PartName="/xl/comments/sheet1.xml" ContentType="...comments+xml"/>
  <Override PartName="/xl/threadedComments/sheet1.xml" ContentType="...threadedComments+xml"/>
  <Override PartName="/xl/people/people.xml" ContentType="...person+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="...drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="...chart+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="...table+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="...pivotTable+xml"/>
  <Override PartName="/xl/pivotCacheDefinition.xml" ContentType="...pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCacheRecords1.xml" ContentType="...pivotCacheRecords+xml"/>
  <Override PartName="/xl/connections.xml" ContentType="...connections+xml"/>
  <Override PartName="/xl/queryTables/queryTable1.xml" ContentType="...queryTable+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="...sheetMetadata+xml"/>
  <Override PartName="/xl/slicers/slicer1.xml" ContentType="...slicer+xml"/>
  <Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="...slicerCache+xml"/>
  <Override PartName="/xl/timelines/timeline1.xml" ContentType="...timeline+xml"/>
  <Override PartName="/xl/scenarios/sheet1.xml" ContentType="...scenarios+xml"/>
  <Override PartName="/xl/customXml/item1.xml" ContentType="application/xml"/>
  <Override PartName="/xl/customXml/itemProps1.xml" ContentType="...customXmlProps+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="...core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="...extended-properties+xml"/>
  <Override PartName="/docProps/custom.xml" ContentType="...custom-properties+xml"/>
</Types>
```

---

## 88. 完整数字格式令牌参考

### 88.1 全部格式令牌

| 令牌 | 含义 | 示例输入 | 显示 |
|------|------|---------|------|
| `0` | 数字占位（补零） | `3.1` (fmt: `0.00`) | `3.10` |
| `#` | 数字占位（不补零） | `3` (fmt: `#.##`) | `3` |
| `?` | 数字占位（空格补齐） | `3.1` (fmt: `?.?`) | `3.1 ` |
| `.` | 小数点 | — | — |
| `,` | 千分位/千位缩放 | `1234567` (fmt: `#,##0`) | `1,234,567` |
| `%` | 百分比 | `0.25` (fmt: `0%`) | `25%` |
| `E+` | 科学计数法(正指数) | `12345` (fmt: `0.00E+00`) | `1.23E+04` |
| `E-` | 科学计数法(负指数) | `0.0001` (fmt: `0.00E-00`) | `1.00E-04` |
| `/` | 分数分隔 | `0.5` (fmt: `?/?`) | `1/2` |
| `\` | 转义下一个字符 | `123` (fmt: `0\!`) | `123!` |
| `"text"` | 文本嵌入 | `123` (fmt: `0"元"`) | `123元` |
| `*` | 重复字符填满 | `123` (fmt: `0*#`) | `123####` |
| `_` | 空格(等于下一个字符宽度) | `123` (fmt: `_0`) | ` 123` |
| `@` | 文本占位 | `abc` (fmt: `"!"@`) | `!abc` |
| `[Color]` | 颜色 | `123` (fmt: `[Red]0`) | 红色123 |
| `[condition]` | 条件 | `60` (fmt: `[>50][Red]0;0`) | 红色60 |
| `[DBNum1]` | 中文小写 | `123` (fmt: `[DBNum1]0`) | 一百二十三 |
| `[DBNum2]` | 中文大写 | `123` (fmt: `[DBNum2]0`) | 壹佰贰拾叁 |
| `[DBNum3]` | 全角数字 | `123` (fmt: `[DBNum3]0`) | １２３ |
| `[Red]` | 红色 | — | — |
| `[Black]` | 黑色 | — | — |
| `[Blue]` | 蓝色 | — | — |
| `[Green]` | 绿色 | — | — |
| `[White]` | 白色 | — | — |
| `[Cyan]` | 青色 | — | — |
| `[Magenta]` | 品红 | — | — |
| `[Yellow]` | 黄色 | — | — |
| `[ColorN]` | 调色板颜色(N=1-56) | — | — |

### 88.2 日期/时间令牌

| 令牌 | 含义 | 示例 |
|------|------|------|
| `yyyy` | 四位年 | 2024 |
| `yy` | 两位年 | 24 |
| `mmmm` | 月份全称 | March |
| `mmm` | 月份缩写 | Mar |
| `mm` | 月份数字(补零) | 03 |
| `m` | 月份数字 | 3 |
| `dd` | 日期(补零) | 15 |
| `d` | 日期 | 15 |
| `dddd` | 星期全称 | Friday |
| `ddd` | 星期缩写 | Fri |
| `hh` | 小时(补零) | 14 |
| `h` | 小时 | 14 |
| `mm` | 分钟(补零，跟在h后面时) | 30 |
| `ss` | 秒(补零) | 00 |
| `AM/PM` | 上午/下午 | PM |
| `am/pm` | 上午/下午(小写) | pm |
| `A/P` | 上/下(缩写) | P |
| `[h]` | 总小时数 | 38 |
| `[m]` | 总分钟数 | 2280 |
| `[s]` | 总秒数 | 136800 |
| `.000` | 毫秒 | 14:30:00.123 |

### 88.3 特殊格式示例

```
正数;负数;零;文本                    ← 四段格式
[>50][Green]0;[<20][Red]0;[Blue]0    ← 条件格式
0.00;[Red]-0.00                      ← 负数红色
0.00" °C";-0.00" °C"                ← 带单位
0"°";0"°"                            ← 度数
#,##0.00,,                           ← 百万(千位缩放两次)
#,##0.00,                            ← 千(千位缩放一次)
[>=1000000]0.00,,"M";[>=1000]0.00,"K";0    ← 自适应K/M
hh:mm:ss.000                         ← 带毫秒时间
yyyy"年"m"月"d"日";@                  ← 中文日期+文本段
[$-404]e"年"m"月"d"日"               ← 台湾地区格式
[$-804]yyyy"年"m"月"d"日"            ← 中国大陆格式
[DBNum2][$-804]0"元"                 ← 中文大写金额
```

---

## 89. 完整样式属性参考

### 89.1 字体(font)全部属性

```xml
<font>
  <!-- 基本属性 -->
  <sz val="11"/>                          <!-- 字号 -->
  <name val="微软雅黑"/>                    <!-- 字体名 -->
  <color rgb="FF0000FF"/>                 <!-- 颜色 -->
  
  <!-- 样式开关 -->
  <b val="1"/>                             <!-- 粗体 -->
  <i val="1"/>                             <!-- 斜体 -->
  <u val="single"/>                        <!-- 下划线类型 -->
  <strike/>                                <!-- 删除线 -->
  <outline/>                               <!-- 轮廓 -->
  <shadow/>                                <!-- 阴影 -->
  <condense/>                              <!-- 压缩 -->
  <extend/>                                <!-- 扩展 -->
  
  <!-- 垂直对齐 -->
  <vertAlign val="superscript"/>          <!-- 上标/下标 -->
  
  <!-- 字体族/方案 -->
  <family val="2"/>                        <!-- 字体族 -->
  <scheme val="minor"/>                    <!-- 字体方案(minor/major) -->
  
  <!-- 字符集 -->
  <charset val="134"/>                     <!-- GB2312(简体中文) -->
</font>
```

### 89.2 下划线类型

| u val | 效果 |
|-------|------|
| `none` | 无 |
| `single` | 单线 |
| `double` | 双线 |
| `singleAccounting` | 会计单线 |
| `doubleAccounting` | 会计双线 |

### 89.3 字符集(charset)常用值

| 值 | 字符集 |
|----|--------|
| 0 | ANSI (西欧) |
| 128 | Shift-JIS (日文) |
| 134 | GB2312 (简体中文) |
| 136 | Big5 (繁体中文) |
| 129 | EUC-KR (韩文) |
| 163 | Vietnamese |
| 177 | Hebrew |
| 178 | Arabic |
| 186 | Baltic |
| 204 | Russian |
| 222 | Thai |
| 238 | Central European |

### 89.4 填充(fill)全部属性

```xml
<!-- 纯色填充 -->
<fill>
  <patternFill patternType="solid">
    <fgColor rgb="FFFFC000"/>       <!-- 前景色 -->
    <bgColor rgb="FF000000"/>       <!-- 背景色(solid模式下无效) -->
  </patternFill>
</fill>

<!-- 渐变填充 -->
<fill>
  <gradientFill degree="90" type="linear">
    <stop position="0">
      <color rgb="FF4472C4"/>
    </stop>
    <stop position="1">
      <color rgb="FFED7D31"/>
    </stop>
  </gradientFill>
</fill>

<!-- 图案填充 -->
<fill>
  <patternFill patternType="darkGrid">
    <fgColor rgb="FF000000"/>
    <bgColor rgb="FFFFFFFF"/>
  </patternFill>
</fill>
```

### 89.5 图案类型(patternType)

| 值 | 效果 | 值 | 效果 |
|----|------|----|------|
| `none` | 无填充 | `solid` | 纯色 |
| `gray125` | 12.5%灰 | `gray0625` | 6.25%灰 |
| `darkHorizontal` | 深色横线 | `lightHorizontal` | 浅色横线 |
| `darkVertical` | 深色竖线 | `lightVertical` | 浅色竖线 |
| `darkDown` | 深色下斜线 | `lightDown` | 浅色下斜线 |
| `darkUp` | 深色上斜线 | `lightUp` | 浅色上斜线 |
| `darkGrid` | 深色网格 | `lightGrid` | 浅色网格 |
| `darkTrellis` | 深色菱形 | `lightTrellis` | 浅色菱形 |
| `mediumGray` | 中灰 | `darkGray` | 深灰 |
| `lightGray` | 浅灰 | `lightTrellis` | 浅色菱形 |

### 89.6 边框(border)完整属性

```xml
<border diagonalUp="1" diagonalDown="1" outline="0">
  <left style="thin"><color rgb="FF000000"/></left>
  <right style="medium"><color rgb="FF0000FF"/></right>
  <top style="thick"><color rgb="FFFF0000"/></top>
  <bottom style="dashed"><color rgb="FF00FF00"/></bottom>
  <diagonal style="dotted"><color rgb="FF808080"/></diagonal>
  
  <!-- 水平/垂直内边线（仅多单元格区域有效） -->
  <horizontal style="thin"><color rgb="FF000000"/></horizontal>
  <vertical style="thin"><color rgb="FF000000"/></vertical>
</border>
```

| 属性 | 说明 |
|------|------|
| `diagonalUp` | 从左下到右上的对角线 |
| `diagonalDown` | 从左上到右下的对角线 |
| `outline` | 是否绘制外边框（0=不绘制） |

### 89.7 对齐(alignment)完整属性

```xml
<alignment
  horizontal="center"              <!-- 水平对齐 -->
  vertical="center"                <!-- 垂直对齐 -->
  textRotation="0"                 <!-- 文字旋转角度(0-180) -->
  wrapText="1"                     <!-- 自动换行 -->
  indent="0"                       <!-- 缩进(0-250) -->
  relativeIndent="0"               <!-- 相对缩进 -->
  justifyLastLine="1"              <!-- 两端对齐末行也分散 -->
  shrinkToFit="0"                  <!-- 缩小字体适应列宽 -->
  readingOrder="0"                 <!-- 阅读顺序(0=上下文,1=LTR,2=RTL) -->
  mergeCell="1"                    <!-- 是否合并单元格(只读) -->
/>
```

### 89.8 textRotation 特殊值

| 值 | 效果 |
|----|------|
| 0 | 无旋转 |
| 1-90 | 逆时针旋转 1-90 度 |
| 91-180 | 顺时针旋转 (180-91) 度 |
| 255 | 垂直文字（每个字符堆叠） |

---

## 90. 工作簿结构完整 XML Schema

### 90.1 worksheet.xml 元素完整层级

```
<worksheet>
├── <sheetPr>                      ← 工作表属性（Tab颜色、大纲、过滤模式）
│   ├── <tabColor>
│   ├── <outlinePr>
│   └── <pageSetUpPr>
├── <dimension ref="A1:Z100"/>     ← 数据范围
├── <sheetViews>                   ← 视图设置
│   └── <sheetView>
│       ├── <pane>                  ← 冻结/拆分窗格
│       └── <selection>             ← 选中区域
├── <sheetFormatPr>                ← 格式默认值
│   (defaultRowHeight, defaultColWidth, outlineLevelRow, outlineLevelCol)
├── <cols>                         ← 列定义
│   └── <col>
├── <sheetData>                    ← 数据（必需）
│   └── <row>
│       └── <c>
│           ├── <v>                ← 值
│           ├── <f>                ← 公式
│           └── <is>               ← 内联字符串
├── <sheetCalcPr>                  ← 计算属性
├── <sheetProtection>              ← 工作表保护
├── <protectedRanges>              ← 受保护区域
├── <scenarios>                    ← 方案
├── <autoFilter>                   ← 自动筛选
│   └── <filterColumn>
├── <sortState>                    ← 排序状态
├── <dataConsolidate>              ← 数据合并
├── <customSheetViews>             ← 自定义视图
├── <mergeCells>                   ← 合并单元格
│   └── <mergeCell>
├── <phoneticPr>                   ← 拼音属性
├── <conditionalFormatting>        ← 条件格式
│   └── <cfRule>
├── <dataValidations>              ← 数据验证
│   └── <dataValidation>
├── <hyperlinks>                   ← 超链接
│   └── <hyperlink>
├── <printOptions>                 ← 打印选项
├── <pageMargins>                  ← 页面边距
├── <pageSetup>                    ← 页面设置
├── <headerFooter>                 ← 页眉页脚
│   ├── <oddHeader>
│   ├── <oddFooter>
│   ├── <evenHeader>
│   ├── <evenFooter>
│   ├── <firstHeader>
│   └── <firstFooter>
├── <rowBreaks>                    ← 行分页符
├── <colBreaks>                    ← 列分页符
├── <customProperties>             ← 自定义属性
├── <cellWatches>                  ← 监视单元格
├── <ignoredErrors>                ← 忽略的错误
├── <smartTags>                    ← 智能标记
├── <drawing>                      ← 绘图引用（形状、图片、图表）
├── <legacyDrawing>                ← 旧式绘图（VML）
├── <legacyDrawingHF>              ← 页眉页脚旧式绘图
├── <picture>                      ← 背景图片
├── <oleObjects>                   ← OLE 对象
├── <controls>                     ← ActiveX 控件
├── <webPublishItems>              ← Web 发布
├── <tableParts>                   ← 表格对象引用
├── <extLst>                       ← 扩展列表（Sparklines、Slicers等）
│   └── <ext>
└── <dataBoundObjects>            ← 数据绑定对象
```

### 90.2 workbook.xml 元素完整层级

```
<workbook>
├── <fileVersion>                   ← 文件版本
├── <fileSharing>                   ← 文件共享
├── <workbookPr>                    ← 工作簿属性（日期系统、编码等）
├── <workbookProtection>            ← 工作簿保护
├── <bookViews>                     ← 工作簿视图
│   └── <workbookView>
├── <sheets>                        ← 工作表列表（必需）
│   └── <sheet>
├── <functionGroups>               ← 函数组
├── <externalReferences>            ← 外部引用
├── <definedNames>                 ← 命名区域
│   └── <definedName>
├── <calcPr>                        ← 计算属性
├── <oleSize>                       ← OLE 对象大小
├── <customWorkbookViews>           ← 自定义工作簿视图
├── <pivotCaches>                   ← 透视缓存
│   └── <pivotCache>
├── <smartTagPr>                    ← 智能标记属性
├── <smartTagTypes>                 ← 智能标记类型
├── <webPublishing>                 ← Web 发布
├── <fileRecoveryPr>                ← 文件恢复
├── <webPublishObjects>             ← Web 发布对象
└── <extLst>                        ← 扩展列表
    └── <ext>                       ← (Slicer缓存、Timeline缓存等)
```

---

## 91. 工作表元素排序规则（关键！）

### 91.1 XML 元素的顺序必须严格遵守

OOXML 规范要求 `<worksheet>` 的子元素**必须按特定顺序排列**，否则 Excel 会报错。以下是正确的排列顺序：

```
 1. sheetPr                      ← 工作表属性
 2. dimension                    ← 数据范围
 3. sheetViews                    ← 视图设置
 4. sheetFormatPr                 ← 格式默认值
 5. cols                          ← 列宽定义
 6. sheetData                     ← 数据（唯一必需的元素）
 7. sheetCalcPr                   ← 计算属性
 8. sheetProtection               ← 工作表保护
 9. protectedRanges               ← 受保护区域
10. scenarios                     ← 方案
11. autoFilter                    ← 自动筛选
12. sortState                     ← 排序状态
13. dataConsolidate               ← 数据合并
14. customSheetViews              ← 自定义视图
15. mergeCells                    ← 合并单元格
16. phoneticPr                    ← 拼音属性
17. conditionalFormatting         ← 条件格式
18. dataValidations               ← 数据验证
19. hyperlinks                    ← 超链接
20. printOptions                  ← 打印选项
21. pageMargins                   ← 页面边距
22. pageSetup                     ← 页面设置
23. headerFooter                  ← 页眉页脚
24. rowBreaks                      ← 行分页符
25. colBreaks                      ← 列分页符
26. customProperties               ← 自定义属性
27. cellWatches                    ← 监视单元格
28. ignoredErrors                 ← 忽略的错误
29. smartTags                     ← 智能标记
30. drawing                       ← 绘图（形状/图片/图表）
31. legacyDrawing                  ← 旧式绘图（VML）
32. legacyDrawingHF                ← 页眉页脚旧式绘图
33. picture                        ← 背景图片
34. oleObjects                     ← OLE 对象
35. controls                       ← ActiveX 控件
36. webPublishItems                ← Web 发布项
37. tableParts                     ← 表格对象
38. extLst                         ← 扩展列表（Sparklines等）
```

### 91.2 常用元素的排列速记

```
sheetPr → dimension → sheetViews → sheetFormatPr → cols → 
sheetData → mergeCells → conditionalFormatting → dataValidations → 
hyperlinks → printOptions → pageMargins → pageSetup → 
headerFooter → rowBreaks → colBreaks → drawing → tableParts → extLst
```

### 91.3 最常见的排序错误

```
❌ 错误：mergeCells 在 sheetData 之前
✅ 正确：sheetData → mergeCells

❌ 错误：conditionalFormatting 在 mergeCells 之前
✅ 正确：mergeCells → conditionalFormatting

❌ 错误：dataValidations 在 hyperlinks 之后
✅ 正确：dataValidations → hyperlinks

❌ 错误：pageSetup 在 pageMargins 之前
✅ 正确：pageMargins → pageSetup

❌ 错误：drawing 在 printOptions 之前
✅ 正确：printOptions → ... → drawing
```

---

## 92. 嵌入式全格式实现路线图

### 92.1 完整实现优先级矩阵

```
                    高使用频率
                        │
        ┌───────────────┼───────────────┐
        │               │               │
        │  第一优先级     │  第二优先级     │
        │  (必须实现)     │  (推荐实现)     │
        │               │               │
        │  • 基本 xlsx   │  • 数据验证     │
        │  • 单元格/行/列 │  • 条件格式     │
        │  • 合并单元格   │  • 冻结窗格     │
        │  • 样式系统     │  • 自动筛选     │
        │  • 共享字符串   │  • 图表        │
        │  • 数字格式     │  • 图片插入     │
        │  • 日期时间     │  • 表格对象     │
        │  • 多 Sheet    │  • 页眉页脚     │
        │  • 打印设置     │  • 打印区域     │
        │  • ZIP 打包     │  • 分组大纲     │
        │  • CSV 输出    │  • 命名区域     │
        │  • HTML 输出   │  • 批注        │
        │               │  • 超链接       │
        │               │  • 错误值       │
        │               │  • 公式        │
 低─────┼───────────────┼───────────────┼─────高
复杂度   │  第三优先级     │  第四优先级     │  复杂度
        │  (可选实现)     │  (按需实现)     │
        │               │               │
        │  • 工作表保护   │  • 数据透视表   │
        │  • Sheet Tab色 │  • 切片器      │
        │  • 隐藏行列     │  • 时间线      │
        │  • 分页符      │  • 线程批注     │
        │  • 富文本      │  • 自定义XML   │
        │  • 视图设置     │  • 外部连接     │
        │  • ODS 格式    │  • 表单控件     │
        │  • SpreadsheetML│ • 图表工作表   │
        │  • 文档属性     │  • 绘图形状     │
        │  • 自定义属性   │  • 文本框/标注  │
        │  • 假设方案     │  • OLE 对象    │
        │  • 拼音信息     │  • ActiveX    │
        │               │  • Sparklines  │
        │               │  • 富数据类型   │
        │               │  • LAMBDA     │
        │               │  • 墨迹注释     │
        │               │  • VBA 宏     │
        │               │  • 文件加密     │
        └───────────────┴───────────────┘
```

### 92.2 各格式对应的 ZIP 文件条目数

| 功能组合 | 文件数 | 累计 |
|---------|--------|------|
| 最小 xlsx | 5 | 5 |
| + 样式 + 共享字符串 | 2 | 7 |
| + 文档属性 | 2 | 9 |
| + 多 Sheet(3个) | 2 | 11 |
| + 图表 | 3 | 14 |
| + 图片 | 2 | 16 |
| + 批注(旧式+线程) | 4 | 20 |
| + 表格对象 | 2 | 22 |
| + 数据验证 | 0(内嵌sheet) | 22 |
| + 条件格式 | 0(内嵌sheet) | 22 |
| + 数据透视表 | 4 | 26 |
| + 切片器 | 3 | 29 |
| + 绘图形状 | 1 | 30 |
| + 表单控件 | 2 | 32 |
| + 外部连接 | 2 | 34 |
| + 自定义XML | 3 | 37 |
| + 线程批注 | 2 | 39 |
| **全功能最大** | — | **~40** |

> **ZIP 写入器建议**：`MAX_ZIP_ENTRIES` 设为 **64**，确保全功能时不溢出。

### 92.3 全格式内存需求

| 实现级别 | ROM(代码) | RAM(运行时) | 支持的格式 |
|---------|----------|-----------|-----------|
| 极简 | ~4 KB | ~2 KB | xlsx(Store) + CSV |
| 基础 | ~12 KB | ~16 KB | + HTML + 基本样式 + 合并 |
| 标准 | ~24 KB | ~48 KB | + 公式 + 数据验证 + 条件格式 + 图表 + 图片 + 批注 + 表格对象 + 冻结 + 打印 |
| 高级 | ~48 KB | ~96 KB | + ODS + SpreadsheetML + 透视表 + 切片器 + 形状 + 控件 + 线程批注 + 自定义XML |
| 全格式 | ~80 KB | ~128 KB | + .xlsb + VBA + 加密 + 外部连接 + OLE + ActiveX + Sparklines + 墨迹 |

### 92.4 格式选择决策表

```
需要 Excel 打开？
├── 是 → 需要公式？
│   ├── 是 → 需要宏？
│   │   ├── 是 → .xlsm
│   │   └── 否 → .xlsx
│   └── 否 → .xlsx
└── 否 → 需要其他办公软件打开？
    ├── 是 → LibreOffice/OpenOffice → .ods
    │       浏览器 → .html
    │       任何文本编辑器 → .csv / .tsv
    └── 否 → 纯数据交换 → .csv
            旧系统兼容 → .xls (BIFF8)
            需要XML解析 → SpreadsheetML 2003 (.xml)
            最小体积 → .xlsb
```

### 92.5 全格式特性覆盖检查表

| # | 特性 | 状态 | 章节 |
|---|------|------|------|
| 1 | 基本 xlsx 结构 | ✅ | 1-11 |
| 2 | 合并单元格（水平/垂直/区域/交错） | ✅ | 5-6 |
| 3 | 样式系统（字体/颜色/边框/对齐） | ✅ | 7, 89 |
| 4 | 共享字符串 | ✅ | 8 |
| 5 | ZIP 打包(Store/Deflate) | ✅ | 10, 46 |
| 6 | 日期时间系统 | ✅ | 15 |
| 7 | 数字精度与格式化 | ✅ | 16, 40, 88 |
| 8 | 多工作表 | ✅ | 17 |
| 9 | 公式（共享/数组/3D引用） | ✅ | 18, 77 |
| 10 | 数据验证 | ✅ | 19 |
| 11 | 条件格式 | ✅ | 20 |
| 12 | 图表（10种类型） | ✅ | 21 |
| 13 | 图片插入 | ✅ | 22 |
| 14 | 超链接 | ✅ | 23 |
| 15 | 批注（旧式+线程式） | ✅ | 24, 70, 87 |
| 16 | 自动筛选 | ✅ | 25 |
| 17 | 表格对象 | ✅ | 26 |
| 18 | 命名区域 | ✅ | 27 |
| 19 | 分组与大纲 | ✅ | 28 |
| 20 | 工作表/工作簿保护 | ✅ | 29 |
| 21 | 打印设置（边距/纸张/区域/标题行） | ✅ | 30 |
| 22 | 页眉页脚 | ✅ | 31 |
| 23 | 分页符 | ✅ | 32 |
| 24 | 视图设置 | ✅ | 33 |
| 25 | 富文本 | ✅ | 34 |
| 26 | 错误值 | ✅ | 35 |
| 27 | Sheet Tab 颜色 | ✅ | 36 |
| 28 | 隐藏行列 | ✅ | 37 |
| 29 | 冻结/拆分窗格 | ✅ | 38 |
| 30 | 单位换算 | ✅ | 39 |
| 31 | 自定义数字格式 | ✅ | 40, 88 |
| 32 | Sparklines | ✅ | 41 |
| 33 | 外部链接 | ✅ | 42, 72 |
| 34 | 自定义文档属性 | ✅ | 43 |
| 35 | VBA 宏(.xlsm) | ✅ | 44 |
| 36 | 文件加密 | ✅ | 45 |
| 37 | Deflate 压缩 | ✅ | 46 |
| 38 | API 设计 | ✅ | 47 |
| 39 | 线程安全 | ✅ | 48 |
| 40 | 错误处理 | ✅ | 49 |
| 41 | 兼容性矩阵 | ✅ | 50 |
| 42 | 测试验证 | ✅ | 51 |
| 43 | 内存管理 | ✅ | 52 |
| 44 | 数据类型完整参考 | ✅ | 53, 76 |
| 45 | 文件清单总表 | ✅ | 54 |
| 46 | CSV / TSV | ✅ | 56 |
| 47 | HTML 导出 | ✅ | 57 |
| 48 | SpreadsheetML 2003 | ✅ | 58 |
| 49 | ODS 格式 | ✅ | 59 |
| 50 | .xlsb 二进制 | ✅ | 60 |
| 51 | .xls (BIFF8) | ✅ | 61 |
| 52 | 绘图形状 | ✅ | 62 |
| 53 | 文本框与标注 | ✅ | 63 |
| 54 | 表单控件 | ✅ | 64 |
| 55 | ActiveX 控件 | ✅ | 65 |
| 56 | 数据透视表 | ✅ | 66 |
| 57 | 切片器 | ✅ | 67 |
| 58 | 时间线 | ✅ | 68 |
| 59 | 图表工作表 | ✅ | 69 |
| 60 | 线程批注 | ✅ | 70 |
| 61 | 自定义 XML 部件 | ✅ | 71 |
| 62 | 外部数据连接 | ✅ | 72 |
| 63 | 计算链与计算属性 | ✅ | 73 |
| 64 | 假设分析方案 | ✅ | 74 |
| 65 | 拼音信息 | ✅ | 75 |
| 66 | 富数据类型 | ✅ | 76 |
| 67 | 3D 引用 | ✅ | 77 |
| 68 | 动态数组 | ✅ | 78 |
| 69 | LAMBDA | ✅ | 79 |
| 70 | OLE 对象 | ✅ | 80 |
| 71 | 智能标记 | ✅ | 81 |
| 72 | 旧式绘图(VML) | ✅ | 82 |
| 73 | 墨迹注释 | ✅ | 83 |
| 74 | Web 发布 | ✅ | 84 |
| 75 | 文件恢复 | ✅ | 85 |
| 76 | 工作簿元数据 | ✅ | 86 |
| 77 | 完整样式属性 | ✅ | 89 |
| 78 | 元素排序规则 | ✅ | 91 |
| 79 | 全格式路线图 | ✅ | 92 |
| 80 | ark_excel 总体架构 | ✅ | 93 |
| 81 | 两遍法流式 ZIP | ✅ | 94 |
| 82 | 数据模型与 SST 踩坑 | ✅ | 95 |
| 83 | XML 生成与元素顺序 | ✅ | 96 |
| 84 | 样式表与日期算法 | ✅ | 97 |
| 85 | 可移植性与 svcrtos | ✅ | 98 |
| 86 | 测试体系 | ✅ | 99 |
| 87 | 内存与性能 | ✅ | 100 |
| 88 | API 速查 | ✅ | 101 |
| 89 | 特性验证矩阵 | ✅ | 102 |

---

# 实现篇：ark_excel 库的具体实现

> 以下章节（93–102）记录 `ark_excel` 库的落地实现——一个依据本文档前 92 章规范写成的、可在 RTOS 与裸机上运行的 xlsx 生成库。代码位于 `D:\工作\git_project\ark_excel`。

---

## 93. ark_excel：总体架构与文件结构

### 93.1 库的三层架构

```
┌─────────────────────────────────────────────┐
│  应用层: ark_set_str / ark_merge / ark_save  │  ← 60+ 个 API
├─────────────────────────────────────────────┤
│  生成层: 数据模型 + XML 逐文件流式生成        │  ← ark_excel.c
├─────────────────────────────────────────────┤
│  容器层: 两遍法流式 ZIP + CRC32              │  ← ark_zip.c
├─────────────────────────────────────────────┤
│  输出层: ark_write_fn 回调(用户实现)          │  ← 文件/UART/网络
└─────────────────────────────────────────────┘
```

### 93.2 文件清单

| 文件 | 职责 | 依赖 |
|------|------|------|
| `ark_excel.h` | 全部公共 API + 数据结构透明定义（便于静态分配） | 无 |
| `ark_excel_config.h` | 编译期容量裁剪（单元格数、字符串数、样式数…） | 无 |
| `src/ark_excel.c` | 数据模型、样式表、全部 OOXML XML 生成、保存流程 | string.h / stdio.h(snprintf) |
| `src/ark_zip.h/.c` | 两遍法流式 ZIP 写入器 + CRC32 查表 | string.h |
| `port/port_examples.c` | stdio / 内存缓冲 / UART / FatFS / svcrtos 五套输出回调（仅示例） | 按平台 |

**零 OS 依赖、零动态内存、单个 `ark_wb_t` 结构承载全部状态**——工作簿可放静态区、指定 RAM 段或堆，库不关心。

### 93.3 与设计规范的对应

| 本文档章节 | 库中的实现 |
|-----------|-----------|
| 第 2 章 ZIP 容器 | `ark_zip.c`：Store 模式 Local Header/Central Directory/EOCD |
| 第 9.3 节流式生成 | 两遍法（见第 94 章），ZIP 层零数据缓冲 |
| 第 47 章 API 设计 | `ark_write_fn` 回调 + 分层 API + 预定义样式常量 |
| 第 49 章错误处理 | `ark_err_t` 枚举（22 个错误码）+ `ark_errstr()` |
| 第 91 章元素排序 | 生成器按 schema 顺序输出（曾被 Excel 拒开的关键） |

---

## 94. 两遍法流式 ZIP：对 UART 也正确的零缓冲输出

### 94.1 问题

ZIP 的 Local Header 要求在**数据之前**写入 CRC32 和大小，但流式生成时数据还没算完。md 第 9.3 节给出两个方案：两遍法与一遍回写法。嵌入式输出流（UART、网络 socket）**不可回写**，所以 ark_excel 实现了两遍法：

```c
/* 第一遍: 只测量, 不输出 */
ark_zip_measure_start(z);
for (每块数据) ark_zip_measure_write(z, buf, len);   /* 累计 CRC32 + 尺寸 */

/* 第二遍: 用测量值写 Local Header, 再流式输出同样数据 */
ark_zip_open_stream(z, "xl/worksheets/sheet1.xml");
for (同样的数据块) ark_zip_write(z, buf, len);
ark_zip_close_stream(z);
```

### 94.2 XML 层的配合：确定性重放

两遍法成立的前提是**同一生成函数跑两遍产出完全相同的字节**。ark_excel 的做法是把每个 XML 文件封装成纯函数：

```c
static ark_err_t save_file(ark_wb_t *wb, ark_zip_t *zip,
                           const char *name, ark_gen_fn gen)
{
    ark_xw_t x;
    ark_zip_measure_start(zip);
    xw_init(&x, zip, 0);          /* pass=0: 只测量 */
    gen(wb, &x);  xw_flush(&x);

    ark_zip_open_stream(zip, name);
    xw_init(&x, zip, 1);          /* pass=1: 真实输出 */
    gen(wb, &x);  xw_flush(&x);
    ark_zip_close_stream(zip);
}
```

`gen` 只读工作簿状态、不修改它，因此两遍结果逐字节一致。生成器内部用一个 256 字节的分块缓冲（`ARK_IO_CHUNK`），写满即转发给 ZIP 层——输出回调的调用粒度恒定，不会长时间阻塞 RTOS 调度。

### 94.3 CRC32

预计算 256 项查找表（存 Flash，第 10.2 节原表），增量接口：

```c
uint32_t ark_crc32(uint32_t crc, const void *data, size_t len);
```

---

## 95. 数据模型：静态池 + arena 字符串

### 95.1 核心结构（透明定义）

```c
typedef struct {
    uint32_t row;          /* 1..1048576 */
    uint16_t col;          /* 1..16384 */
    uint8_t  sheet;
    uint8_t  type;         /* SSTR/ISTR/NUM/INT/DATE/FORMULA/BOOL/ERROR/BLANK */
    uint16_t style;        /* cellXf 索引 */
    uint32_t sidx;         /* SST 序号 / arena 索引 */
    double   num;          /* 数值 / 日期序列号 / 布尔 */
} ark_cell_t;              /* 恰好 24 字节 */
```

工作簿 = 若干固定容量池的集合：`cells[]`（默认 2048）、`merges[]`、`coldefs[]`、`rowdefs[]`、`hyperlinks[]`、`validations[]`、`condfmts[]`、`names[]`。所有容量在 `ark_excel_config.h` 一处裁剪，RAM 峰值编译期即可确定（见第 14 章内存预估方法的直接应用）。

### 95.2 字符串 arena 与 SST 的共存设计（踩坑记录）

初版把 SST 和引用串（公式文本、范围地址、超链接目标）混在一个数组，且假设 SST 必须是数组前缀——公式一旦先于字符串写入，SST 区被"打断"，索引全面错位。修复后的设计：

```c
typedef struct {
    uint16_t len;
    uint16_t sst_seq;      /* 属于 SST 时的 SST 序号 */
    uint8_t  is_sst;       /* 是否共享字符串 */
    char     text[ARK_MAX_STR_LEN - 3];
} ark_str_t;
```

- 添加字符串时置 `is_sst=1`，记下 `sst_seq`（即第 8 章的 SST 索引）
- 单元格引用存 `sst_seq`，生成 `sharedStrings.xml` 时按 arena 顺序输出带 `is_sst` 的条目——**序号天然一致**
- SST 满时 `ark_set_str` 自动降级为内联字符串（第 49.3 节优雅降级）

> **教训**：SST 去重函数失败时返回的错误码若为正值，会被调用方当成索引使用——错误码必须统一返回负值。这是真实发生过的 bug：SST 满后所有字符串悄悄变成"第 7 条"。

### 95.3 保存时的排序

用户可以任意顺序调用 `ark_set_*`（乱序、覆盖写），保存前对单元格按 `sheet → row → col` 做无递归堆排序（`O(n log n)`、零额外内存），保证 `<row>`/`<c>` 严格升序——这是 Excel 的硬性要求，也是第 13 章陷阱清单的第一条。同址重复写取最后一次（后写覆盖先写，与直觉一致）。

---

## 96. XML 生成器：模板直拼 + 手写数字转换

### 96.1 发射器（xw_*）

为避免 snprintf 的栈开销和格式串解析开销，常用输出走手写路径：

| 函数 | 用途 |
|------|------|
| `xw_str/xw_ch/xw_raw` | 字面量与原始字节 |
| `xw_u32/xw_i64` | 手写整数转十进制（比 printf 快且不依赖库） |
| `xw_dbl` | `%.15g`（浮点必须走 snprintf，见第 16.4 节） |
| `xw_esc` | XML 转义：`& < > " '` + 丢弃非法控制字符（第 8.4 节） |

长 XML 片段（命名空间 URL、固定头部）作为字符串字面量拼接，编译器合并存放于 Flash（第 9.5 节技巧）。

### 96.2 单元格输出（emit_cell）

第 4 章与第 53 章类型的完整落地：

```c
<c r="C6" s="6">          <!-- r=地址 s=样式 -->
  <v>25.5</v>             <!-- NUM -->
</c>
<c r="A1" s="2" t="s"><v>0</v></c>      <!-- SST 字符串 -->
<c r="B4" t="inlineStr"><is><t>…</t></is></c>  <!-- SST 满时降级 -->
<c r="C11"><f>AVERAGE(C6:C10)</f></c>   <!-- 公式，只存 f 不存 v -->
<c r="F7" t="e"><v>#N/A</v></c>         <!-- 错误值 -->
<c r="A7" s="2"></c>                    <!-- BLANK: 仅样式(合并区边框用) -->
```

### 96.3 元素顺序：Excel 拒开文件的三大元凶（实测）

开发中被真实 Excel 拒绝过的问题，全部是**顺序/位置**错误——openpyxl 全部能容错打开，Excel 严格按 schema 校验：

| 问题 | 规范要求 | 后果 |
|------|---------|------|
| `<sqref>` 写成子元素 | dataValidation 的 sqref 是**属性** | 文件损坏 |
| `printOptions` 在 `pageSetup` 后 | 必须在 `pageMargins` **之前** | 需要修复 |
| `dxfs` 在 `cellStyles` 前 | cellStyles 必须在前 | 需要修复 |
| `HeadingPairs` size=1 | variant 向量必须成对（lpstr+i4），size=2 | 需要修复 |
| blank 单元格漏 `>` | `<c r="A7" s="2"` 后必须闭合标签 | XML 解析失败 |
| rels 中 URL 未转义 | `Target="…?a=1&b=2"` 的 `&` 须写 `&amp;` | XML 解析失败 |

worksheet 子元素最终顺序（第 91 章的实测版）：

```
sheetPr → dimension → sheetViews → sheetFormatPr → cols → sheetData
→ autoFilter → mergeCells → conditionalFormatting → dataValidations
→ hyperlinks → printOptions → pageMargins → pageSetup → headerFooter
→ rowBreaks
```

**结论：验证 xlsx 合法性必须用真实 Excel（可用 COM 自动化批量静默打开），openpyxl 通过 ≠ Excel 能打开。**

### 96.4 fitToPage 的隐藏依赖

`pageSetup` 里写 `fitToWidth` 不够，Excel 还要求 `sheetPr` 内声明：

```xml
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
```

否则缩放设置不生效。sheetPr 是 worksheet 的**第一个**子元素，tabColor 也在其中。

---

## 97. 样式表：去重注册 + 索引复用

### 97.1 三张注册表

`fonts[]` / `fills[]` / `borders[]` 各自做内容去重（比较全部字段，相同返回已有索引），fill 表前两项固定为 `none`/`gray125`（规范强制）。`cellXfs` 是组合索引：`{font, fill, border, numfmt, halign, valign, wrap, indent}`，同样去重。19 个预定义样式（标题/表头/交错行/警告/成功/金额/日期…）在 `ark_wb_init` 时注册，索引即 `ARK_STYLE_*` 枚举值，用户零成本可用。

### 97.2 自定义样式

```c
ark_style_def_t sd = ARK_STYLE_DEF_INIT;   /* 必须清零! */
sd.bold = 1; sd.fill_color = 0xFF00FF00u;
sd.border_all = ARK_BORDER_THICK;
sd.custom_fmt = "0.000\"kg\"";             /* 自定义数字格式 → numFmt id≥164 */
int st = ark_wb_add_style(&wb, &sd);       /* 返回样式 id, 可直接用于 ark_set_* */
```

`ARK_STYLE_DEF_INIT` 清零不可省——未清零的结构体中 `font_name` 是野指针，会直接写进 styles.xml 导致文件损坏。

### 97.3 日期序列号（1900 闰年 bug 的正确补法）

第 15 章算法的实测修正版：

```c
serial = days_from_civil(y,m,d) - days_from_civil(1899,12,31);
if (days >= 60) serial += 1.0;   /* 1900-03-01 起补偿伪闰日 */
serial += 时分秒 / 86400.0;
```

**边界必须是 `>= 60` 而不是 `>= 61`**：真实日历 1900-03-01 距 1899-12-31 是 60 天，加上伪闰日正好等于 Excel 的 61；用 61 判断会让 1900-03-01 输出成 60（Excel 显示为 1900-02-28）。

---

## 98. 可移植性：输出回调与 svcrtos 集成

### 98.1 一个回调适配所有目标

```c
typedef int (*ark_write_fn)(const void *data, size_t len, void *ctx);
```

| 目标 | 回调实现要点 |
|------|-------------|
| FatFS/LittleFS | `f_write` 返回值校验，ctx 传 `FIL*` |
| POSIX/Win32 | `fwrite` 比较 len |
| 裸机 UART | 直接调 BSP 阻塞发送；库是纯流式，零 RAM 缓冲 |
| 内存缓冲 | memcpy + 越界检查，可后续加密/压缩/入网络栈 |

### 98.2 svcrtos（自研 RTOS）推荐接线

```
生成任务(高优先级)                 写任务(低优先级)
  ark_set_* 填数据                   storage_append(...)
  ark_wb_save()                        ↑
    └→ 回调: 内存池 alloc + 队列 send ─┘
```

- 回调里 `mempool_alloc` 定长块（建议 128B，与 `ARK_IO_CHUNK` 匹配）→ 队列投递 → 写任务落盘
- 池满即返回失败 → 生成自然背压，不需要额外锁——**单生产者单消费者模型天然线程安全**
- 生成大量数据时分块回调不会长时间独占 CPU（第 48.6 节分块生成的落地）

### 98.3 无浮点 printf 的平台

`ark_set_int()` 内部手写 int64→十进制，完全不依赖 `%f/%g`；只有 `ark_set_num/date` 需要 `snprintf("%.15g")`。纯整数报表在无 FPU 的 MCU 上零障碍。

---

## 99. 测试体系（tests/ 目录）

### 99.1 三层测试

| 层 | 文件 | 内容 |
|----|------|------|
| 单特性生成 | `test_features.c` | 19 个用例，每个特性独立成文件（basic/escape/merge/altrow/date/formula/freeze/colrow/hidden/styles/condfmt/valid/hyperlink/names/multisheet/print/view/sst/overwrite） |
| 语义校验 | `validate_features.py` | openpyxl 打开每个文件，68 项断言（值、样式、合并范围、日期序列号、公式、冻结、下拉框、条件格式、命名区域、打印…） |
| 真机校验 | 同上（第二层） | **Excel COM 批量静默打开**——Excel 严格 schema 校验，任何损坏立即暴露 |
| 回归 | `stress_test.c` + `validate.py` | 1500 单元格乱序+覆盖写；行/列引用升序逐节点断言 |

### 99.2 特性二分定位法

"文件打不开"类问题的高效排查：把特性拆成 N 个独立文件让 Excel 逐个打开，损坏特性立即定位到文件——曾用该方法把"文件损坏"在 19 分钟内定位到 `dataValidation` 的 sqref 位置错误。

```bash
# 完整跑一遍
gcc -std=gnu99 -I.. tests/test_features.c src/ark_excel.c src/ark_zip.c -o t
cd tests && ./t && python validate_features.py
```

### 99.3 当前状态

```
openpyxl 68/68 | Excel COM 19/19 | 压力回归 PASS
```

本轮测试修复的库 bug：blank 单元格标签未闭合、rels URL 未转义、1900-03-01 日期边界。

---

## 100. 内存与性能

### 100.1 默认配置 RAM 预算

| 池 | 默认容量 | 大小 |
|----|---------|------|
| cells | 2048 × 24B | 48 KB |
| strs | 1024 × ~96B | ≈96 KB |
| merges/coldefs/rowdefs/… | 各 8–64 条 | < 4 KB |
| ZIP 条目表 + IO 缓冲 | 32 × ~76B + 256B | ≈2.7 KB |

裁剪到 STM32F1 级（256 单元格 / 256 串 / 串长 48）约 28 KB；配合 UART 直发，保存期增量仅需 IO 缓冲 256B。

### 100.2 输出量

Store 模式（不压缩）下 demo（49 单元格、19 样式、多特性）约 17 KB；500 行×3 列报表约 87 KB。需要更小体积时按第 46 章外接 miniz，ZIP 层已预留 method 字段。

---

## 101. API 速查

```c
/* 生命周期 */
ark_wb_init(wb, write_cb, ctx);  ark_wb_add_sheet(wb, "名字");  ark_wb_save(wb);

/* 单元格 */
ark_set_str / set_num / set_int / set_bool / set_formula / set_error
ark_set_blank / set_date / set_datetime

/* 结构 */
ark_merge / set_col_width / set_col_hidden / set_row_height / set_row_hidden

/* 视图 */
ark_freeze / autofilter / tab_color / sheet_rtl / sheet_gridlines

/* 交互 */
ark_hyperlink / add_validation / add_condfmt / named_range

/* 打印 */
ark_page_setup / print_area / header_footer / row_break

/* 样式 */
ARK_STYLE_* 19 个预定义;  ark_wb_add_style(&sd) 自定义
```

完整签名与错误码见 `ark_excel.h`（含注释约 500 行）。

---

## 102. 已验证特性矩阵

| 特性 | openpyxl | Excel 实测 |
|------|:--:|:--:|
| 文本/整数/浮点/布尔/错误值/空样式格 | ✅ | ✅ |
| XML 特殊字符转义、UTF-8 中文、制表符 | ✅ | ✅ |
| 合并（横/纵/区域/三级表头） | ✅ | ✅ |
| 交错行、19 预定义样式、自定义样式（字体/填充/粗边框/自定义数字格式） | ✅ | ✅ |
| 日期/时间（含 1900 边界） | ✅ | ✅ |
| 公式、命名区域 | ✅ | ✅ |
| 冻结窗格、自动筛选 | ✅ | ✅ |
| 列宽/行高/隐藏行列 | ✅ | ✅ |
| 条件格式（cellIs 双规则）、数据验证（list/whole between） | ✅ | ✅ |
| 超链接（URL/带参数 URL/内部跳转/提示） | ✅ | ✅ |
| 多 Sheet（5 个，各自数据+Tab 颜色） | ✅ | ✅ |
| 打印（A3 横向/适应页宽/打印区域/页眉页脚/分页符） | ✅ | ✅ |
| RTL、网格线开关、SST 去重与降级、乱序覆盖写 | ✅ | ✅ |

---

## 103. 读改写：打开/修改/保存现有 xlsx

### 103.1 设计目标与总体思路

只写的库不算完整的 Excel 库——现场设备最常见的场景是"出厂模板 + 运行数据填充"：
模板由 Excel 制作（含图表、主题、Logo），固件只需往指定单元格写数据。
为此 ark_excel 实现了完整的**读 + 改 + 写**链路，设计原则：

1. **自带 DEFLATE 解压**（`ark_inflate.c`，puff 风格 ~400 行）：不引入 zlib 依赖，
   能直接打开 Excel / openpyxl 生成的压缩 xlsx
2. **未修改部件原样字节复制**：模板中的图表/主题/图片等库本身不支持的特性，
   保存时按 ZIP 条目逐字节复制（含压缩数据，不重压缩），**零损失**
3. **被改部件重新生成**：仅被 touch 的 sheet、SST、styles 走正常生成路径
4. **样式增量追加**：新样式（含 19 个内置样式）追加在模板样式表之后，
   id 通过 `ark_style_map()` 重映射，模板原样式全部保留

### 103.2 读取层架构

```
ark_reader_io_t (seek+read 回调)     ← 文件系统 / 内存缓冲均可包装
        │
        ▼
┌─ ZIP 读取层 ────────────────────────┐
│ EOCD 尾部搜索 → 中央目录解析(46B/条) │
│ → 条目定位 → Store 直读 / Deflate 解压│
└──────────────┬─────────────────────┘
               ▼
┌─ xlsx 载入层 ───────────────────────┐
│ workbook.xml: sheet 名 + rId        │
│ workbook.xml.rels: rId → 部件路径    │
│ sharedStrings.xml: SST(保序不去重)   │
│ styles.xml: 仅统计计数(fonts/fills/  │
│   borders/cellXfs/numFmt/dxfs)      │
│ sheetN.xml: 单元格/合并/列宽行高/    │
│   冻结/筛选/公式 → ark_wb_t          │
└────────────────────────────────────┘
```

关键实现细节：

- **文件大小探测**：从 0 到 1GB 二分 seek，无需 fstat，Flash 移植零成本
- **SST 保序载入**：共享字符串按模板顺序进 arena，`sst_seq` 与模板索引一一对应，
  单元格的 `t="s"` 引用号才能直接映射
- **XML 实体解码**：`&amp; &lt; &gt; &quot; &apos;` + 数字字符引用 `&#NNNN;`/`&#xHHHH;`
  （openpyxl 默认用数字实体写中文，不解码就是乱码）
- **单元格类型全覆盖**：无 t(数字/日期) / `t="n"` / `t="s"` / `t="inlineStr"`
  (openpyxl 常用，文本在 `<is><t>`) / `t="b"` / `t="e"` / 公式(取 `<f>` 忽略缓存值)

### 103.3 读取回调

```c
typedef uint32_t (*ark_read_fn) (void *ctx, uint8_t *buf, uint32_t len);
typedef int      (*ark_seek_fn) (void *ctx, uint32_t pos);

typedef struct {
    ark_read_fn  read;
    ark_seek_fn  seek;
    void        *ctx;
} ark_reader_io_t;
```

与写出回调对称。FatFS 上就是 `f_read`/`f_lseek` 的两行包装（见 `port/port_examples.c`）。

### 103.4 增量保存：dirty 标记体系

每个写 API 在 open 模式下自动 touch 对应部件：

| 操作 | 标记 | 保存时的处理 |
|------|------|-------------|
| `ark_set_*` 写单元格 | `sheet_dirty[si]` | 该 sheet XML 重新生成 |
| 新增共享字符串(SST) | `sst_dirty` | sharedStrings.xml 重新生成 |
| `ark_wb_add_style` 或引用重映射样式 | `styles_dirty` | styles.xml 增量拼接 |
| `ark_wb_add_sheet` | `wb_dirty` | workbook/rels/[Content_Types] 全部重生成 |
| 未 touch 的任何条目 | — | **原样字节复制**(含 Deflate 数据) |

样式增量拼接的核心是 `seg_copy()`：拷贝模板 fonts/fills/borders/cellXfs/dxfs 段，
替换段内 `count="N"` 为新计数，再把新条目追加在段尾。模板段落可能是自闭合
（如 openpyxl 的 `<numFmts count="0"/>`），处理时剥掉 `/` 转为开标签再补闭合。

### 103.5 使用示例

```c
static ark_wb_t wb;
static ark_reader_t rd;           /* ~17KB, 建议静态分配 */

/* 1. 初始化 + 打开(读取源保持打开到 save 完成!) */
ark_wb_init(&wb, wr_nop, NULL);
ark_reader_io_t io = { my_read, my_seek, &file };
ark_wb_open(&wb, &rd, &io);

/* 2. 查询载入结果 */
double v; char txt[96]; uint16_t st;
ark_get_cell(&wb, 0, 2, 2, NULL, &v, NULL, 0, NULL);

/* 3. 修改: 注意内置样式要经 ark_style_map() 重映射 */
ark_set_num(&wb, 0, 2, 2, 99.9, ark_style_map(ARK_STYLE_NUM_2DEC));
ark_set_str(&wb, 0, 4, 1, "新数据", ark_style_map(ARK_STYLE_WARN));

/* 4. 保存(此时才需要真正的输出流) */
wb.write = my_write; wb.ctx = &outfile;
ark_wb_save(&wb);     /* 自动走增量路径 */
```

### 103.6 限制与边界

- 单个 XML 部件解压后 ≤ `ARK_OPEN_XML_BUF`(默认 16KB)，超限返回错误而非截断
- 载入的是数据层；模板里的图表/批注/透视表**不解析但零损失保留**
- 新增 sheet 要求模板部件名符合 `sheet1..N` 规范（ark_excel/Excel 生成均满足）
- `ark_wb_open` 后读取源必须保持可读至 `ark_wb_save` 完成（增量保存要读模板字节）

## 104. 功能裁剪：用 config.h 应对不同资源的嵌入式设备

### 104.1 两级裁剪模型

```
功能开关 (ARK_ENABLE_*)     → 裁 flash/代码, 模块整块进出链接
容量参数 (ARK_MAX_*)        → 裁 RAM/数据池, 编译期定死峰值
```

### 104.2 功能开关总表

| 开关 | 默认 | 控制内容 | 关闭效果 |
|------|:--:|---------|---------|
| `ARK_ENABLE_READWRITE` | 1 | 读改写(ark_reader.c) | 生成-only, 省 ~15KB flash |
| `ARK_ENABLE_INFLATE` | 1 | DEFLATE 解压 | 只能打开 Store 模板, 省 ~3KB |
| `ARK_ENABLE_FORMULA` | 1 | 公式写入/载入 | API 返回 `ARK_ERR_BAD_STATE` |
| `ARK_ENABLE_CONDFMT` | 1 | 条件格式(dxf) | 同上 |
| `ARK_ENABLE_VALIDATION` | 1 | 数据验证 | 同上 |
| `ARK_ENABLE_HYPERLINK` | 1 | 超链接 | 同上 |

设计要点：

- **裁剪不静默**：被裁功能的 API 保留原型，调用返回 `ARK_ERR_BAD_STATE`——
  误调用在代码审查/单元测试阶段就暴露，而不是文件里悄悄缺数据
- **依赖自动折叠**：`READWRITE=0` 时 `INFLATE` 自动失效, touch 钩子变空宏，
  `ark_reader.c` 整个编译为空翻译单元，无需从工程里摘文件
- **仅写场景连 reader 都不用编**：链接列表从 4 个 .c 缩到 2 个

### 104.3 实测代码体积（gcc 4.8 -Os，库本体 .text）

| 配置 | 大小 | 适用场景 |
|------|------|---------|
| 全功能(默认) | 48.8 KB | 读改写 + 全特性, Cortex-M3/M4 起步 |
| `READWRITE=0` | 33.1 KB (-32%) | 只写报表, M0/M0+ 舒适 |
| 全部特性关(仅数据+样式) | 31.6 KB | 极小 ROM 裸机数据上报 |

读改写方向的 RAM 代价只有 `ark_reader_t`（≈`ARK_OPEN_XML_BUF` + 条目表 ≈ 17KB），
且按需分配——不开 `ark_wb_open` 就不占（静态分配由用户决定放哪个 RAM 段）。

### 104.4 典型设备配置档案

```c
/* --- 档案 A: STM32F407 + FatFS, 全功能 --- */
/* 默认配置, 无需改动 */

/* --- 档案 B: STM32F103 (64KB RAM), 只写巡检报表 --- */
#define ARK_ENABLE_READWRITE  0
#define ARK_MAX_CELLS         512
#define ARK_MAX_STR           384
#define ARK_MAX_STR_LEN       64

/* --- 档案 C: STM32F030 (16KB RAM), 模板填充(读改写) --- */
#define ARK_ENABLE_FORMULA    0     /* 模板里没有公式 */
#define ARK_ENABLE_CONDFMT    0
#define ARK_ENABLE_VALIDATION 0
#define ARK_ENABLE_HYPERLINK  0
#define ARK_MAX_CELLS         256
#define ARK_MAX_STR           256
#define ARK_MAX_STR_LEN       48
#define ARK_OPEN_XML_BUF      8192  /* 模板部件较小 */

/* --- 档案 D: 裸机 + UART 直发, 无文件系统 --- */
#define ARK_ENABLE_READWRITE  0
#define ARK_EXCEL_SKIP_DOCPROPS 1
#define ARK_MAX_CELLS         128
```

### 104.5 验证裁剪配置

测试矩阵（4 种配置全部编译通过 + 默认配置全量回归）：

```bash
# A: 默认全开 —— 31/31 + 76/76 + COM 全过
gcc -std=gnu99 -I. -Isrc tests/test_rw.c src/*.c -o test_rw.exe && ./test_rw.exe

# B/C: 裁剪配置编译 + 运行验证
gcc -std=gnu99 -DARK_ENABLE_READWRITE=0 ... examples/demo_report.c src/ark_excel.c src/ark_zip.c
# 被裁 API 返回 ARK_ERR_BAD_STATE(17), 可用小驱动程序断言
```

---

## 最终总结

至此，本文档已覆盖 **102 个章节**（其中 93–102 为 ark_excel 库的完整落地实现），涵盖了一个完整 Excel 库所需的**全部格式和特性**。从最基础的 xlsx 结构到最前沿的动态数组和 LAMBDA 函数，从嵌入式 C 代码实现到 API 设计、线程安全和测试策略。

**核心认知框架：**

1. **xlsx 的本质**：ZIP + XML，Store 模式不需要任何压缩库
2. **所有格式的选择**：按场景选择——简单数据用 CSV，丰富格式用 xlsx，跨平台用 ODS，旧系统兼容用 SpreadsheetML
3. **合并与交错格式**：`<mergeCell ref="A1:C2"/>`，被合并格子要有样式但可无值
4. **样式系统**：字体×填充×边框×对齐的组合套餐
5. **元素顺序**：XML 子元素必须按规范排序（第 91 章）
6. **嵌入式实现哲学**：静态分配 + 流式生成 + 回调输出 + 优雅降级
