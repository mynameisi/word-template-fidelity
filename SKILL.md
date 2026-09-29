---
name: word-template-fidelity
description: >-
  让生成的 Word 文档与指定母版模板「一模一样」，并逐条满足外部要求文件
  （字体、字号、行距、页边距、缩进、结构、分页等）的硬性规范。Windows / macOS 通用，
  不依赖 Mac。核心机制：母版即产物（绝不从空白重建）、按版式复杂度选对排版内核、
  只做最小文本改动、把要求文件转成可机检规格表、静态属性断言 + 渲染比对闭环验收，
  并对中文字体（eastAsia）做专门保真。触发：「按某个 Word 模板/格式要求生成文档」
  「必须和模板完全一致，不能差不多」「参考这些要求文件（字体多少号、页边距多少）产出文档」
  「生成的 doc/docx 格式跑偏了」「和空白模板一模一样只填内容」。
version: 2.0.0
author: Hermes Agent
metadata:
  hermes:
    tags: [word, docx, doc, template, fidelity, format-spec, layout, chinese-font, eastasia, cross-platform, windows, macos, python-docx, word-com, applescript, verification]
    prerequisites: [python-docx, pymupdf, fonttools(可选), Microsoft Word (仅复杂 .doc 版式需要)]
---

# Word 模板保真 — 与母版「一模一样」+ 逐条满足外部格式要求

**Windows / macOS 通用，不依赖 Mac。** 中文排版体系下尤其关注字体。

## 这份 skill 解决什么

目标不是「看起来差不多」，而是**可验证的一模一样**：产物与母版在版式上逐项相等，
且满足所有外部要求文件写明的格式规范。

一句话心法：

> **母版是唯一真值源（ground truth）。产物 = 母版的副本 + 最小文本改动；
> 验收 = 逐项 diff，不是肉眼看「像不像」。**

「差不多像了就行」为什么必然失败：版式的关键信息（旋转文本框、浮动表格、合并单元格、
边框、列宽、样式定义、制表位、段后距、**中文字体**）**只存在于原文件的 XML 与排版内核里**，
肉眼在屏幕上根本看不全（字体名、字号、`eastAsia` 属性、`gridSpan` 都看不见）。
所以「像」是错觉，「逐项断言相等」才是判据。

## 触发条件

- 「按这个 Word 模板/格式要求生成一份文档」
- 「必须和模板完全一致，不能差不多像」
- 「这些是格式要求文件（字体、字号、页边距、行距…），照它产出」
- 「和空白模板一模一样，只填内容」（官方表单/推荐书/计划书/封面）
- 「生成的 doc/docx 在 Word 里打开格式不对 / 字体不对 / 边框没了 / 表格错位」

## 四条铁律（不可违背）

1. **母版即产物** — 永远 `Document(template)` / 复制文件 / Word 打开母版再另存，
   **绝不从空白文档重建版式**。重建 = 边框/列宽/合并/字体定义全丢。
2. **用对排版内核** — 复杂 `.doc` 版式（旋转文本框、浮动表格、折页拼版）必须用
   **Microsoft Word 本体**改；纯表格/段落型 `.docx` 用 `python-docx` 复制母版 + 只改 run。
   **内核选错 = 必然失真。**
3. **最小改动** — 只碰文本，不碰结构。保留段落标记、`pPr`、`rPr`、合并单元格。
4. **验收必须机器可判** — 每条要求都绑定一个可执行的校验（读属性断言 / 渲染读图），
   不靠「我感觉像了」。

## 第 0 步：路线选择（先定内核，再动手）

### 平台能力矩阵（同一能力，两端各有等价实现）

| 能力 | Windows | macOS | 跨平台首选 |
|:-----|:--------|:------|:-----------|
| 驱动 Word 本体做查找替换/导出 | **PowerShell + COM**（`New-Object -ComObject Word.Application`） | AppleScript（`osascript`） | 二者等价，按平台选 |
| 读 `.docx` XML、改 run | `python-docx` | `python-docx` | **`python-docx`** |
| 抽取文本（`.docx`） | `python-docx` | `python-docx` | **`python-docx`** |
| 抽取文本（`.doc`） | Word COM | Word AppleScript 或 `textutil` | 用 Word 本体另存 `.docx` 再读 |
| PDF → PNG 渲染 | **PyMuPDF** | **PyMuPDF** | **PyMuPDF**（纯 Python，免装 poppler） |
| 枚举已装字体 | `C:\Windows\Fonts` + 用户字体目录 | `/System/Library/Fonts` 等 | **`fontTools` 读 name 表** |

> **不要写死平台命令。** 任何脚本都先 `sys.platform` 判断，或直接选跨平台方案。
> PyMuPDF 与 python-docx 在 Windows/macOS/Linux 行为一致，是默认路径。

### 路线表

| 母版特征 | 内核 | 手段 |
|:---------|:-----|:-----|
| `.doc`（旧格式）、含旋转文本框/浮动表格/折页拼版/复杂页眉 | **Microsoft Word 本体** | Windows：`scripts/gen_word_replace_powershell.py`；macOS：`scripts/gen_word_replace_applescript.py` |
| `.docx`、以表格表单为主、无旋转元素 | **python-docx** | `Document(template)` → 只改 run → 结构签名校验 |
| `.docx`、纯段落型（封面/公文/报告） | **python-docx** | 只改正文 run，**禁止改 `styles.xml`** |

**绝对不要**：用 LibreOffice 转存作为最终产物；用 `python-docx` 做 `.doc→.docx` 往返
处理含旋转文本框的文件（会损坏旋转文本框/浮动表格）。

> ⚠️ 两个易混点：
> - **LibreOffice 的 UNO/Python 桥**在 macOS 26+ 会被代码签名 Launch Constraint 直接
>   SIGKILL（exit 137）。`soffice --convert-to` 命令行转换本身可用，但别当最终排版内核。
> - `soffice --convert-to` 对**含旋转文本框的 `.doc`** 同样有损 → 走 Word 本体路线。

## 第 1 步 ★核心★：把「要求文件」转成可机检规格表

用户一次性给的可能是**一叠异构文件**：母版模板、格式规范/排版说明文档、样例 PDF、
评审/评分标准、甚至口述要求。它们必须被**归一成一张规格表**，而不是读一遍凭记忆写。

### 1.1 抽取（跨平台）

| 来源 | 抽取方法 |
|:-----|:---------|
| `.docx` | `python-docx` 读段落 + 表格 |
| `.doc` | 先用 Word 本体另存为 `.docx`（Windows COM / macOS AppleScript），再按 `.docx` 读 |
| `.pdf`（规范/样例） | `pymupdf`：`page.get_text("text")`；样例版式还要 `page.get_pixmap()` 出图 |
| `.xlsx` / `.csv` | 读表（规范常以表格列出「对象 / 属性 / 要求」） |
| 口述/聊天记录 | 直接抄进规格表，标注「口头要求」 |

### 1.2 归一成规格表

**每条要求 = 一行**，字段固定：`对象 | 属性 | 期望值 | 容差 | 校验方式`。

```markdown
| 对象 | 属性 | 期望值 | 容差 | 校验方式 |
|:-----|:-----|:-------|:-----|:---------|
| 全文正文 | 中文字体 (eastAsia) | 仿宋_GB2312 | 精确 | A 静态 + 字体已装 |
| 全文正文 | 字号 | 14 pt（四号） | ±0.5 | A 静态 |
| 标题 1 | 加粗 | 是 | — | A 静态 |
| 全文 | 行距 | 固定值 28 pt | ±1 | A 静态 |
| 页面 | 页边距 上/下/左/右 | 2.54 / 2.54 / 3.17 / 3.17 cm | ±0.2 | A 静态 |
| 表格 1 | 列数 / gridSpan / vMerge | 与母版一致 | 精确 | A 静态 |
| 全文 | 总页数 | 4 | 精确 | B 渲染 |
| 页脚 | 「第 N 页（共 M 页）」 | 与样例一致 | — | B 渲染 |
| 正文 | 大题标题不落在页尾 | 成立 | — | B 渲染 |
```

**属性清单（照抄不漏）**：
- 字体：**四个槽全查** —— `eastAsia`（中文）、`ascii`、`hAnsi`、`cs`；字号（`sz`/`szCs`）、加粗、倾斜、下划线、颜色
- 段落：行距（倍数/固定值）、段前距、段后距、首行缩进、左缩进、对齐、制表位位置
- 页面：纸张尺寸、页边距、页眉页脚距离、分节
- 结构：表格数、每表行列数、`gridSpan`、`vMerge`、列宽、边框
- 分页：总页数、各标题所在页、表格是否跨页、页尾空白
- 内容：编号体系、必填字段、占位符

### 1.3 校验方式只有两类（决定用什么工具）

- **A 类 · 可静态读取**（字体、字号、边距、结构、缩进）→ 读 XML 属性后**断言相等**。
  用 `scripts/verify_docx_spec.py`。
- **B 类 · 只能渲染判断**（分页、跨页、对齐、页尾空白、**字体是否被静默替换**）→
  导出 PDF → 渲染成 PNG → 逐页读图，与基准 PDF 并排比对。用 `scripts/render_pdf_pages.py`。

### 1.4 冲突仲裁（必须显式记录，不许静默选一个）

```
母版实物  >  明确的文字规范  >  行业/学校惯例
```

若用户另有指定，以用户为准。**任何冲突都写进规格表并告知用户**，说明你选了哪个、为什么。

### 1.5 落盘

写成 `spec.json`（机器可读，供脚本断言）+ `spec.md`（人可读，供交付说明）。
**验收只依据规格表，不依据记忆。**

## 第 2 步 ★核心★：中文字体保真（中文体系下最关键）

中文文档（公文、论文、试卷、合同、标书）对字体有硬性规定，而 Word 的字体模型比想象复杂。
**字体不对 = 直接不合格**，且往往肉眼不易发现。

### 2.1 四个字体槽 —— 中文字形只认 `eastAsia`

`w:rFonts` 有四个属性：

| 属性 | 管哪些字符 | 典型值 |
|:-----|:-----------|:-------|
| `w:ascii` | 西文（A-Z 0-9） | Times New Roman |
| `w:hAnsi` | 西文高 ANSI | 通常同 `ascii` |
| `w:eastAsia` | **中日韩字符（中文！）** | 仿宋_GB2312 |
| `w:cs` | 复杂文种 | — |

> **只设 `w:ascii` 中文不会变。** `python-docx` 的 `run.font.name = "仿宋"` 只写
> `w:ascii` + `w:hAnsi`，**不写 `w:eastAsia`** —— 这是「设了字体但中文没变」的头号原因。

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_cjk_font(run, cjk: str, western: str | None = None) -> None:
    """同时设置四个槽：中文走 eastAsia，西文走 ascii/hAnsi。"""
    rPr = run._element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    rf.set(qn("w:eastAsia"), cjk)
    rf.set(qn("w:ascii"), western or cjk)
    rf.set(qn("w:hAnsi"), western or cjk)
    # 若母版用了主题字体，显式字体名会被主题覆盖 -> 一并清掉
    for attr in ("w:eastAsiaTheme", "w:asciiTheme", "w:hAnsiTheme"):
        if rf.get(qn(attr)) is not None:
            del rf.attrib[qn(attr)]
```

### 2.2 主题字体陷阱

`w:eastAsiaTheme="minorEastAsia"` 这类属性表示「字体由主题解析」，**渲染时才决定**。
静态读到 `w:eastAsia` 为空 **≠** 没有字体。校验时必须二选一：
要么解析 theme（复杂），要么**统一改成显式字体名**（推荐，可控）。

### 2.3 字号：中文「号」↔ 磅，且中文用 `w:sz`

| 号 | 磅(pt) | `w:sz`(半磅) | 号 | 磅(pt) | `w:sz`(半磅) |
|:---|:------:|:-----------:|:---|:------:|:-----------:|
| 初号 | 44 | 88 | 小三 | 15 | 30 |
| 小初 | 36 | 72 | 四号 | 14 | 28 |
| 一号 | 26 | 52 | 小四 | 12 | 24 |
| 小一 | 24 | 48 | 五号 | 10.5 | 21 |
| 二号 | 22 | 44 | 小五 | 9 | 18 |
| 小二 | 18 | 36 | 六号 | 7.5 | 15 |
| 三号 | 16 | 32 | 七号 | 5.5 | 11 |

- **中文正文**：小四（12pt）或 五号（10.5pt）
- **公文正文**：三号仿宋（16pt）
- **论文正文**：小四宋体（12pt）
- 中文用 `w:sz`（半磅值）；`w:szCs` 是复杂文种字号，别搞混。

`python-docx` 的 `run.font.size = Pt(14)` 会写 `w:sz val="28"`，正确。

### 2.4 跨平台字体名不一致（最致命的坑）

**同一个「宋体」在 Windows 和 macOS 上不是同一个字体文件**，字形与字宽都有差异：

| 中文字体 | Windows 名 | macOS 名 | 字体文件 |
|:---------|:-----------|:---------|:---------|
| 宋体 | `SimSun` / 宋体 | `Songti SC` / 宋体-简 | simsun.ttc |
| 黑体 | `SimHei` / 黑体 | `Heiti SC` / 黑体-简 | simhei.ttf |
| 楷体 | `KaiTi` / 楷体 | `Kaiti SC` | simkai.ttf |
| 仿宋 | `FangSong` / 仿宋 | `STFangsong` | simfang.ttf |
| 微软雅黑 | `Microsoft YaHei` | **macOS 无**（回退） | msyh.ttc |
| 仿宋_GB2312 | **需单独安装** | **通常无** | FZFSK.TTF |
| 方正小标宋简体 | 需单独安装 | 通常无 | — |

后果：
- 字形/字宽不同 → **分页会漂移**。
- `仿宋_GB2312`、`方正小标宋简体`、`华文中宋` 等**很多 Windows 也没装**。
- macOS 上**没有微软雅黑**。

**结论**：
1. **基准 PDF 必须与产物在同一平台产出**，否则比对无意义。
2. 跨平台交付前，先确认对方装了哪些字体；缺字体就用**两端都有的等价字体**并在规格表里写明替代关系。
3. 规格表里记录**「要求的字体名」+「平台实际字体名」**两列。

### 2.5 字体缺失 = 静默回退（必须靠渲染发现）

Word 找不到字体时**不报错**，静默替换成别的字体（通常退到宋体/系统默认）。
静态检查读到的仍是 `仿宋_GB2312`，但渲染出来是别的 —— **只有 B 类渲染比对能发现**。

```bash
# 先确认本机到底装了哪些字体
python3 scripts/verify_docx_spec.py --list-fonts
# 再让校验脚本断言规格要求的字体确实已安装
#   spec.json: "font": {"eastAsia": "仿宋_GB2312", "require_installed": true}
```

### 2.6 PDF 必须内嵌字体

导出 PDF 时若不内嵌字体，别人机器上打开会被替换 → 版式变化。
Word 本体导出 PDF 默认内嵌；用第三方工具转换时必须显式确认（PyMuPDF 可查 `page.get_fonts()`）。

### 2.7 中文标点与字距

中文标点（，。「」）与全角字符的宽度由 `w:eastAsia` 字体决定；`w:kern`（字距调整）
影响中西文混排效果。**母版若开了 `w:kern`，别关掉。**

## 第 3 步：抽取母版的隐式规格（规范没写、但必须保真的部分）

- **结构签名**：`(grid 列数, [(gridSpan, vMerge), ...])` —— 填表前后必须完全相等。
- **制表位**：`tab stops`，字段靠 tab 对齐，母版每段可能不同。
- **段后距 / 缩进**：母版靠 `space after` 与 `first line indent` 控制分页与对齐。
- **中文字体**：在 `w:rFonts/@w:eastAsia`，clone 字体必须 deepcopy 整段 `w:rFonts` XML。
- **自动编号**：段落若挂 Word 自动编号，编号不在文本里，删段会重排后续编号。

## 第 4 步：执行改动（最小改动原则）

### 4.1 文本替换（保留段落结构，跨平台）

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

### 4.2 用 Word 本体做整段查找替换

**Windows（PowerShell + COM）** — `scripts/gen_word_replace_powershell.py`：

```powershell
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($src)
$f = $doc.Content.Find
# Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
#         MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace)
$ok = $f.Execute($old, $true, $false, $false, $false, $false, $true, 1, $false, $new, 2)
$doc.SaveAs2($outDoc, 0)              # wdFormatDocument = .doc
$doc.ExportAsFixedFormat($outPdf, 17) # wdExportFormatPDF
$doc.Close(0); $word.Quit()
```

**macOS（AppleScript）** — `scripts/gen_word_replace_applescript.py`：见脚本头注释。

> 段落若挂 Word 自动编号，直接整段覆盖会把该段踢出列表（编号消失 + 后续重排）。
> 安全做法：**只替换段内文本，保留段落标记**。
> - AppleScript：建一个 `start..(end-1)` 的内部 range 再 `set content`
> - python-docx：只改 `runs[i].text`，不要 `paragraph.text = ...`

### 4.3 硬约束

- **禁止** `doc.add_table(rows, cols)` + 手动 merge —— 边框/列宽/字体全丢。
- **禁止** `cell.text = "..."` 整格覆盖 —— 清掉格内段落/run 结构。
- **禁止**全局设字体（如全文宋体 14pt）—— 会覆盖母版的标题行/正文差异样式。
- **禁止**改 `styles.xml` —— 改样式定义会波及其它所有引用处。
- **禁止** `cell._tc.remove` 删母版预留空段落 —— 官方表单用空行撑开独立「框」，
  删掉后 Word 里会像合并成一个大格。
- 段落缩进用 Word 自动化设置**可能静默失败**（设完读回还是旧值）→ 先解除列表绑定
  （AppleScript `remove numbers` / COM 清 `ListFormat`），再设缩进。
- 删除段落要**从大到小的索引顺序**，否则索引漂移删错段。
- `page break before` 慎用 —— 折页/拼页母版插分页符会补出空白页（4 页变 8 页）。
  优先用**加大上一段的段后距**让内容自然流到下一页。

### 4.4 结构改动后必须全量重验

合并/删行/换长文本会让后面内容整体上移，之前调好的分页**全部漂移**，
且被删段落上挂的段后距设置会一并消失。**每次结构改动 → 重新导出 → 重新跑 A+B 全量校验。**

## 第 5 步：验收闭环

```
抽取要求文件 → 规格表 spec.json / spec.md
      ↓
选内核（Word 本体 / python-docx）
      ↓
复制母版 → 最小改动 → 输出
      ↓
A 类：verify_docx_spec.py 逐条断言（四槽字体/字号/边距/结构/字体已装）
B 类：render_pdf_pages.py 渲染 PNG → 逐页读图对比基准 PDF
      ↓
任一不绿 → 回到改动步骤（不是回到「凭感觉再调调」）
      ↓
全绿 ✅ 交付
```

```bash
# A 类
python3 scripts/verify_docx_spec.py 产物.docx spec.json --print-signature
# B 类（跨平台，纯 Python）
python3 scripts/render_pdf_pages.py 产物.pdf /tmp/pages/new
python3 scripts/render_pdf_pages.py 基准.pdf /tmp/pages/official
# 然后逐页读 PNG：页数、标题位置、表格完整性、对齐、页尾空白、字体是否被替换
```

**逐页、逐项检查，一页都不要跳过。** 结构性改动后要重跑全部页。

## 反模式表（导致版式崩）

| 做法 | 后果 |
|:-----|:-----|
| 从空白文档重建版式 | 边框/列宽/合并/字体定义全丢 |
| 用 LibreOffice 作最终排版内核 | 行距漂移、虚线变实线、页数变化 |
| `python-docx` 做含旋转文本框的 `.doc→.docx` 往返 | 旋转文本框/浮动表格损坏 |
| 只设 `w:ascii` 不设 `w:eastAsia` | **中文没换字体**（最隐蔽） |
| 保留 `w:eastAsiaTheme` 主题字体 | 显式字体名被主题覆盖 |
| `add_table()` 重建表格 | 与母版边框/列宽/字体完全不同 |
| `cell.text = "..."` | 清掉格内段落/run 结构 |
| 全局设字体字号 | 覆盖母版标题行/正文样式差异 |
| 不校验 `gridSpan`/`vMerge` | 静默破坏合并，Word 里对不齐 |
| 假设字体已安装 | Word 静默回退，静态检查查不出来 |
| 跨平台拿不同平台产出的 PDF 比对 | 字体/字宽不同，比对无意义 |
| 肉眼「差不多」判定 | 字体名、eastAsia、gridSpan、分页差异全部漏判 |
| 改完只看第 1 页 | 分页漂移发生在后续页 |

## 调试清单

1. Word 打开产物 vs 母版**并排对比**（同缩放）看边框与对齐。
2. 产物文件大小应接近母版（±30%）；**暴增通常意味着重建了内容**。
3. `python3 scripts/verify_docx_spec.py 产物.docx spec.json`。
4. 某格错位 → 检查是否用了 `cell.text =` 或误删了 vMerge 行。
5. **中文字体不对** → 先查 `w:eastAsia`（不是 `w:ascii`），再查有没有 `*Theme` 残留，
   最后 `--list-fonts` 确认该字体本机装没装。
6. 分页不对 → 看该段 `space after` 是否与同节其它段一致；确认基准 PDF 与产物同平台产出。

## 文件组织建议

```
project/
  官方模板/《xxx》.doc(x)        # 只读，永不覆盖
  要求文件/                      # 格式规范、样例 PDF、评审标准
  spec.json                      # 归一后的可机检规格表
  spec.md                        # 人可读规格表 + 冲突记录 + 字体替代关系
  build_xxx.py                   # 填表脚本（python-docx，跨平台）
  build_xxx.ps1 / build.applescript   # 复杂 .doc 时的 Word 自动化
  _template_work/xxx.docx        # 转换缓存（若需要）
  scripts/
```

## 依赖

| 依赖 | 用途 | 是否必需 |
|:-----|:-----|:---------|
| `python-docx` | 读写 `.docx`、改 run、结构签名 | **必需** |
| `pymupdf` | PDF → PNG 渲染、PDF 文本抽取（跨平台，免 poppler） | **必需** |
| `fonttools` | 枚举本机已安装字体、校验字体是否存在 | 推荐 |
| Microsoft Word | 仅当母版是复杂 `.doc`（旋转文本框/浮动表格/折页）时必需 | 按需 |

```bash
pip install python-docx pymupdf fonttools
```

> 不再需要 `poppler` / `brew` —— PDF 渲染走 PyMuPDF，Windows 上同样可用。

## 脚本

- `scripts/verify_docx_spec.py` — A 类静态校验：四槽字体 / 字号 / 边距 / 结构 / 字体是否已安装；
  `--list-fonts` 枚举本机字体，`--describe-size 12` 换算中文「号」
- `scripts/render_pdf_pages.py` — B 类渲染：PDF 逐页导出 PNG（跨平台，基于 PyMuPDF）
- `scripts/gen_word_replace_powershell.py` — 生成 Windows PowerShell + Word COM 替换脚本
- `scripts/gen_word_replace_applescript.py` — 生成 macOS AppleScript 替换脚本
