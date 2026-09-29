# word-template-fidelity

让生成的 Word 文档与指定**母版模板一模一样**，并逐条满足**外部要求文件**（字体、字号、行距、页边距、缩进、结构、分页等）的硬性规范。

**Windows / macOS 通用，不依赖 Mac。** 中文排版体系下尤其关注字体。

不是「差不多像了就行」，而是**可验证的一模一样**。

## 核心心法

> **母版是唯一真值源。产物 = 母版的副本 + 最小文本改动；验收 = 逐项 diff，不是肉眼看「像不像」。**

「差不多」为什么必然失败：版式的关键信息（旋转文本框、浮动表格、合并单元格、边框、列宽、样式定义、制表位、段后距、**中文字体**）**只存在于原文件的 XML 与排版内核里**，屏幕上根本看不全。

## 四条铁律

1. **母版即产物** — 从模板复制/打开再另存，绝不从空白文档重建版式。
2. **用对排版内核** — 复杂 `.doc` 版式（旋转文本框/浮动表格/折页拼版）必须用 **Microsoft Word 本体**；纯表格/段落型 `.docx` 用 `python-docx` 复制母版 + 只改 run。
3. **最小改动** — 只碰文本，不碰结构。
4. **验收必须机器可判** — 每条要求绑定一个可执行校验，不靠感觉。

## 触发条件

- 「按这个 Word 模板 / 格式要求生成一份文档」
- 「必须和模板完全一致，不能差不多像」
- 「这些是格式要求文件（字体多少号、页边距多少），照它产出」
- 「和空白模板一模一样，只填内容」（官方表单 / 推荐书 / 计划书 / 封面）
- 「生成的 doc/docx 在 Word 里打开格式不对 / 字体不对 / 边框没了 / 表格错位」

## 关键机制一：要求文件 → 可机检规格表

用户一次给的可能是一叠异构文件（母版、格式规范、样例 PDF、评审标准、口述要求），必须**归一成一张规格表**：

| 对象 | 属性 | 期望值 | 容差 | 校验方式 |
|:-----|:-----|:-------|:-----|:---------|
| 全文正文 | 中文字体 (eastAsia) | 仿宋_GB2312 | 精确 | A 静态 + 字体已装 |
| 全文正文 | 字号 | 14 pt（四号） | ±0.5 | A 静态 |
| 页面 | 页边距 | 2.54 / 3.17 cm | ±0.2 | A 静态 |
| 表格 1 | 列数 / gridSpan / vMerge | 与母版一致 | 精确 | A 静态 |
| 全文 | 总页数 | 4 | 精确 | B 渲染 |

- **A 类（可静态读取）** → 读 XML 属性断言相等 → `scripts/verify_docx_spec.py`
- **B 类（只能渲染判断）** → 导出 PDF → 渲染 PNG → 逐页读图 / 像素比对 → `scripts/render_pdf_pages.py`

冲突仲裁优先级：**母版实物 > 明确的文字规范 > 惯例**（用户可覆盖，冲突必须显式记录）。

## 关键机制二：中文字体保真

- **四个字体槽**：`w:ascii` / `w:hAnsi` / `w:eastAsia` / `w:cs`。**中文字形只认 `w:eastAsia`** —— `python-docx` 的 `run.font.name` 不写 `eastAsia`，这是「设了字体但中文没变」的头号原因。
- **主题字体陷阱**：`w:eastAsiaTheme` 会在渲染时覆盖显式字体名，必须清掉或解析。
- **字号换算**：中文「号」↔ 磅（小四 = 12pt = `w:sz val="24"`），脚本可查 `--describe-size 12`。
- **跨平台字体名不一致**：同一个「宋体」在 Windows（`SimSun`）与 macOS（`Songti SC`）**不是同一个字体文件**，字宽不同会导致分页漂移；`仿宋_GB2312`、`方正小标宋简体` 很多机器根本没装。
- **字体缺失 = 静默回退**：Word 找不到字体不报错，静态检查查不出来 —— 必须靠渲染比对 + `--list-fonts` 已装字体检查。
- **PDF 必须内嵌字体**：否则换机器打开会被替换，`--fonts` 可检查。

## 文件结构

```
word-template-fidelity/
├── SKILL.md
├── README.md
└── scripts/
    ├── verify_docx_spec.py                 # A 类静态校验（四槽字体/字号/边距/结构/字体已装）
    ├── render_pdf_pages.py                 # B 类渲染 + 像素比对 + 字体内嵌检查（跨平台）
    ├── gen_word_replace_powershell.py      # Windows：生成 Word COM 替换脚本
    └── gen_word_replace_applescript.py     # macOS：生成 Word AppleScript 替换脚本
```

## 用法

```bash
pip install python-docx pymupdf fonttools

# A 类：按 spec.json 逐条断言
python3 scripts/verify_docx_spec.py output.docx spec.json --print-signature

# 辅助：查本机字体 / 换算中文字号
python3 scripts/verify_docx_spec.py --list-fonts 仿宋
python3 scripts/verify_docx_spec.py --describe-size 12

# B 类：渲染成 PNG 逐页读图，或直接像素比对
python3 scripts/render_pdf_pages.py out.pdf /tmp/pages/new --dpi 100
python3 scripts/render_pdf_pages.py out.pdf --diff reference.pdf
python3 scripts/render_pdf_pages.py out.pdf --fonts

# 复杂 .doc 母版：用 Word 本体做替换
python3 scripts/gen_word_replace_powershell.py pairs.json master.doc new.doc new.pdf -o build.ps1   # Windows
powershell -NoProfile -ExecutionPolicy Bypass -File build.ps1

python3 scripts/gen_word_replace_applescript.py pairs.json master.doc new.doc new.pdf > build.applescript  # macOS
osascript build.applescript
```

> Windows 提示：生成的 `.ps1` 以 **UTF-8 with BOM** 写入 —— Windows PowerShell 5.1 没有 BOM 会把中文读成乱码，导致查找替换静默失败。

## 依赖

| 依赖 | 用途 | 是否必需 |
|:-----|:-----|:---------|
| `python-docx` | 读写 `.docx`、改 run、结构签名 | 必需 |
| `pymupdf` | PDF → PNG 渲染、PDF 文本抽取（跨平台，免 poppler） | 必需 |
| `fonttools` | 枚举本机已安装字体 | 推荐 |
| Microsoft Word | 仅当母版是复杂 `.doc`（旋转文本框/浮动表格/折页）时必需 | 按需 |

## 仓库

https://github.com/mynameisi/word-template-fidelity
