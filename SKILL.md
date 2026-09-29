---
name: word-template-fidelity
description: >-
  让生成的 Word 文档与指定母版模板「一模一样」，并逐条满足外部要求文件
  （字体、字号、行距、页边距、缩进、结构、分页等）的硬性规范。核心机制：
  母版即产物（绝不从空白重建）、按版式复杂度选对排版内核、只做最小文本改动、
  把要求文件转成可机检规格表、用静态属性断言 + 渲染比对做闭环验收。
  触发：「按某个 Word 模板/格式要求生成文档」「必须和模板完全一致，不能差不多」
  「参考这些要求文件（字体多少号、页边距多少）产出文档」「生成的 doc/docx 格式跑偏了」
  「和空白模板一模一样只填内容」。
version: 1.0.0
author: Hermes Agent
metadata:
  hermes:
    tags: [word, docx, doc, template, fidelity, format-spec, layout, applescript, python-docx, verification, forms]
    prerequisites: [Microsoft Word (macOS, 复杂 .doc 版式必选), python-docx, poppler/pdftoppm, LibreOffice (仅简单 .doc→.docx)]
---

# Word 模板保真 — 与母版「一模一样」+ 逐条满足外部格式要求

## 这份 skill 解决什么

目标不是「看起来差不多」，而是**可验证的一模一样**：产物与母版在版式上逐项相等，
且满足所有外部要求文件写明的格式规范。

一句话心法：

> **母版是唯一真值源（ground truth）。产物 = 母版的副本 + 最小文本改动；
> 验收 = 逐项 diff，不是肉眼看「像不像」。**

「差不多像了就行」为什么必然失败：版式的关键信息（旋转文本框、浮动表格、合并单元格、
边框、列宽、样式定义、制表位、段后距、中文字体）**只存在于原文件的 XML 与排版内核里**，
肉眼在屏幕上根本看不全（字体名、字号、eastAsia 属性、gridSpan 都看不见）。
所以「像」是错觉，「逐项断言相等」才是判据。

## 触发条件

- 「按这个 Word 模板/格式要求生成一份文档」
- 「必须和模板完全一致，不能差不多像」
- 「这些是格式要求文件（字体、字号、页边距、行距…），照它产出」
- 「和空白模板一模一样，只填内容」（官方表单/推荐书/计划书/封面）
- 「生成的 doc/docx 在 Word 里打开格式不对 / 边框没了 / 表格错位」

## 四条铁律（不可违背）

1. **母版即产物** — 永远 `Document(template)` / `shutil.copy` / Word 打开母版再另存，
   **绝不从空白文档重建版式**。重建 = 边框/列宽/合并/字体定义全丢。
2. **用对排版内核** — 复杂 `.doc` 版式（旋转文本框、浮动表格、折页拼版、密封线）
   必须用 **Microsoft Word 本体**（AppleScript）改；纯表格 `.docx` 表单用
   `python-docx` 复制母版 + 只改 run。**内核选错 = 必然失真**。
3. **最小改动** — 只碰文本，不碰结构。保留段落标记、`pPr`、`rPr`、合并单元格。
4. **验收必须机器可判** — 每条要求都绑定一个可执行的校验（读属性断言 / 渲染读图），
   不靠「我感觉像了」。

## 第 0 步：路线选择（先定内核，再动手）

| 母版特征 | 内核 | 手段 |
|:---------|:-----|:-----|
| `.doc`（旧格式）、含旋转文本框/浮动表格/折页拼版/复杂页眉 | **Microsoft Word 本体** | AppleScript 查找替换 + `save as`，见 `scripts/gen_word_replace_applescript.py` |
| `.docx`、以表格表单为主、无旋转元素 | **python-docx** | `Document(template)` → 只改 run → 结构签名校验 |
| `.docx`、纯段落型（封面/公文/报告） | **python-docx** | 只改正文 run，**禁止改 styles.xml** |

**绝对不要**：用 LibreOffice 转存作为最终产物；用 `python-docx` 做 `.doc→.docx` 往返
处理含旋转文本框的文件（会损坏旋转文本框/浮动表格）。

> ⚠️ 两个易混点，别踩：
> - **LibreOffice 的 UNO/Python 桥**（`soffice --headless` + python 脚本）在 macOS 26+
>   会被代码签名 Launch Constraint 直接 SIGKILL（exit 137）。但 **`soffice --convert-to`
>   命令行转换本身可用** —— 只是别把它当最终产物的排版内核，仅用于 `.doc→.docx` 一次性取样。
> - `soffice --convert-to` 对**含旋转文本框的 `.doc`** 同样有损。这类文件请直接走 Word 本体路线。

## 第 1 步 ★核心★：把「要求文件」转成可机检规格表

用户一次性给的可能是**一叠异构文件**：母版模板、格式规范/排版说明文档、样例 PDF、
评审/评分标准、甚至口述要求。它们必须被**归一成一张规格表**，而不是读一遍凭记忆写。

### 1.1 抽取

| 来源 | 抽取方法 |
|:-----|:---------|
| `.docx` / `.doc` | `textutil -convert txt -stdout`（macOS）或 python-docx 读段落+表格 |
| `.pdf`（规范/样例） | `pdftotext -layout x.pdf -`；样例版式还要 `pdftoppm` 出图 |
| `.xlsx` / `.csv` | 读表（规范常以表格列出「对象 / 属性 / 要求」） |
| 口述/聊天记录 | 直接抄进规格表，标注「口头要求」 |

### 1.2 归一成规格表

**每条要求 = 一行**，字段固定：`对象 | 属性 | 期望值 | 容差 | 校验方式`。

```markdown
| 对象 | 属性 | 期望值 | 容差 | 校验方式 |
|:-----|:-----|:-------|:-----|:---------|
| 全文正文 | 中文字体 (eastAsia) | 仿宋_GB2312 | 精确 | A 静态 |
| 全文正文 | 字号 | 14 pt | ±0.5 | A 静态 |
| 标题 1 | 加粗 | 是 | — | A 静态 |
| 全文 | 行距 | 固定值 28 pt | ±1 | A 静态 |
| 页面 | 页边距 上/下/左/右 | 2.54 / 2.54 / 3.17 / 3.17 cm | ±0.2 | A 静态 |
| 表格 1 | 列数 / gridSpan / vMerge | 与母版一致 | 精确 | A 静态 |
| 全文 | 总页数 | 4 | 精确 | B 渲染 |
| 页脚 | 「第 N 页（共 M 页）」 | 与样例一致 | — | B 渲染 |
| 正文 | 大题标题不落在页尾 | 成立 | — | B 渲染 |
```

**属性清单（照抄不漏）**：
- 字体：`eastAsia`（中文字体！）+ `ascii`（西文）、字号、加粗、倾斜、下划线、颜色
- 段落：行距（倍数/固定值）、段前距、段后距、首行缩进、左缩进、对齐、制表位位置
- 页面：纸张尺寸、页边距、页眉页脚距离、分节
- 结构：表格数、每表行列数、`gridSpan`、`vMerge`、列宽、边框
- 分页：总页数、各标题所在页、表格是否跨页、页尾空白
- 内容：编号体系、必填字段、占位符

### 1.3 校验方式只有两类（决定用什么工具）

- **A 类 · 可静态读取**（字体、字号、边距、结构、缩进）→ 读 XML 属性后**断言相等**。
  用 `scripts/verify_docx_spec.py`。
- **B 类 · 只能渲染判断**（分页、跨页、对齐、页尾空白、视觉重合）→ 导出 PDF →
  `pdftoppm -png -r 100` → 逐页读图，与基准 PDF 并排比对。

### 1.4 冲突仲裁（必须显式记录，不许静默选一个）

要求文件之间、或要求文件与母版实物冲突时，默认优先级：

```
母版实物  >  明确的文字规范  >  行业/学校惯例
```

若用户另有指定，以用户为准。**任何冲突都写进规格表并告知用户**，说明你选了哪个、为什么。

### 1.5 落盘

把规格表写成 `spec.json`（机器可读，供脚本断言）+ `spec.md`（人可读，供交付说明）。
**验收只依据规格表，不依据记忆。**

## 第 2 步：抽取母版的隐式规格（规范没写、但必须保真的部分）

要求文件不会写「制表位在第 120.5pt」这种细节，但它们是母版版式的一部分，必须原样保留：

- **结构签名**：`(grid 列数, [(gridSpan, vMerge), ...])` —— 填表前后必须完全相等。
- **制表位**：`tab stops of paragraph format`，选项/字段靠 tab 对齐，母版每题可能不同。
- **段后距 / 缩进**：母版靠 `space after` 与 `first line indent` 控制分页与对齐。
- **中文字体**：在 `w:rFonts/@w:eastAsia`，不是 `@w:ascii`。clone 字体必须 deepcopy 整段 `w:rFonts` XML。
- **自动编号**：段落若挂 Word 自动编号，编号不在文本里，删段会重排后续编号。

## 第 3 步：执行改动（最小改动原则）

### 3.1 文本替换（保留段落结构）

```python
# 填值格：只改 run 文本，不动段落
def set_cell_value(cell, text: str) -> None:
    para = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    if para.runs:
        para.runs[0].text = text or ""
        for r in para.runs[1:]:
            r.text = ""
    else:
        para.add_run(text or "")
```

```applescript
-- 段落挂自动编号时，set content 会把该段踢出列表（题号消失 + 后续重排）
-- 安全做法：只替换段内文本，保留段落标记
set tr to text object of paragraph i of theDoc
set innerR to create range theDoc start (start of content of tr) end ((end of content of tr) - 1)
set content of innerR to "新文本"
```

### 3.2 硬约束

- **禁止** `doc.add_table(rows, cols)` + 手动 merge —— 边框/列宽/字体全丢。
- **禁止** `cell.text = "..."` 整格覆盖 —— 清掉格内段落/run 结构。
- **禁止**全局设字体（如全文宋体 14pt）—— 会覆盖母版的标题行/正文差异样式。
- **禁止**改 `styles.xml` —— 改样式定义会波及其它所有引用处。
- **禁止** `cell._tc.remove` 删母版预留空段落 —— 官方表单用空行撑开独立「框」，
  删掉后 Word 里会像合并成一个大格。
- 段落缩进用 AppleScript 设置**可能静默失败**（设完读回还是旧值）→ 先
  `remove numbers (list format of ...)` 解除列表绑定，再设缩进。
- 删除段落要**从大到小的索引顺序**，否则索引漂移删错段。
- `page break before` 慎用 —— 折页/拼页母版插分页符会补出空白页（4 页变 8 页）。
  优先用**加大上一题的段后距**让内容自然流到下一页。

### 3.3 结构改动后必须全量重验

合并/删行/换长文本会让后面内容整体上移，之前调好的分页**全部漂移**，
且被删段落上挂的段后距设置会一并消失。**每次结构改动 → 重新导出 → 重新跑 A+B 全量校验。**

## 第 4 步：验收闭环

```
抽取要求文件 → 规格表 spec.json/spec.md
      ↓
选内核（Word 本体 / python-docx）
      ↓
复制母版 → 最小改动 → 输出
      ↓
A 类：verify_docx_spec.py 逐条断言（字体/字号/边距/结构）
B 类：导出 PDF → pdftoppm -png -r 100 → 逐页读图对比基准 PDF
      ↓
任一不绿 → 回到改动步骤（不是回到「凭感觉再调调」）
      ↓
全绿 ✅ 交付
```

### B 类渲染比对命令

```bash
pdftoppm -png -r 100 产物.pdf /tmp/pages/new
pdftoppm -png -r 100 基准.pdf /tmp/pages/official
# 然后逐页读 PNG，检查：页数、标题位置、表格完整性、对齐、页尾空白
```

**逐页、逐项检查，一页都不要跳过。** 结构性改动后要重跑全部页。

## 反模式表（导致版式崩）

| 做法 | 后果 |
|:-----|:-----|
| 从空白文档重建版式 | 边框/列宽/合并/字体定义全丢 |
| 用 LibreOffice 作最终排版内核 | 行距漂移、虚线变实线、页数变化 |
| `python-docx` 做含旋转文本框的 `.doc→.docx` 往返 | 旋转文本框/浮动表格损坏 |
| `add_table()` 重建表格 | 与母版边框/列宽/字体完全不同 |
| `cell.text = "..."` | 清掉格内段落/run 结构 |
| 全局设字体字号 | 覆盖母版标题行/正文样式差异 |
| 不校验 `gridSpan`/`vMerge` | 静默破坏合并，Word 里对不齐 |
| 肉眼「差不多」判定 | 字体名、eastAsia、gridSpan、分页差异全部漏判 |
| 改完只看第 1 页 | 分页漂移发生在后续页 |

## 调试清单

1. Word 打开产物 vs 母版**并排对比**（同缩放）看边框与对齐。
2. 产物文件大小应接近母版（±30%）；**暴增通常意味着重建了内容**。
3. `python3 scripts/verify_docx_spec.py 产物.docx spec.json`。
4. 某格错位 → 检查是否用了 `cell.text =` 或误删了 vMerge 行。
5. 字体不对 → 检查是不是只设了 `ascii` 没设 `eastAsia`。
6. 分页不对 → 看该段 `space after` 是否与同节其它段一致。

## 文件组织建议

```
project/
  官方模板/《xxx》.doc(x)        # 只读，永不覆盖
  要求文件/                      # 格式规范、样例 PDF、评审标准
  spec.json                      # 归一后的可机检规格表
  spec.md                        # 人可读规格表 + 冲突记录
  build_xxx.py / build.applescript
  _template_work/xxx.docx        # LO 转换缓存（若需要）
  scripts/verify_docx_spec.py
```

## 依赖

- macOS + Microsoft Word（复杂 `.doc` 版式必选）
- `python-docx`（`pip install python-docx`）
- poppler（`brew install poppler`，提供 `pdftoppm` / `pdftotext`）
- LibreOffice（可选，仅用于简单 `.doc→.docx` 一次性取样）

## 脚本

- `scripts/verify_docx_spec.py` — A 类静态校验：按 `spec.json` 断言字体/字号/边距/结构
- `scripts/gen_word_replace_applescript.py` — 生成驱动 Word 本体做段落级查找替换的 AppleScript
