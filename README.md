# word-template-fidelity

让生成的 Word 文档与指定**母版模板一模一样**，并逐条满足**外部要求文件**（字体、字号、行距、页边距、缩进、结构、分页等）的硬性规范。

不是「差不多像了就行」，而是**可验证的一模一样**。

## 核心心法

> **母版是唯一真值源。产物 = 母版的副本 + 最小文本改动；验收 = 逐项 diff，不是肉眼看「像不像」。**

「差不多」为什么必然失败：版式的关键信息（旋转文本框、浮动表格、合并单元格、边框、列宽、样式定义、制表位、段后距、中文字体）**只存在于原文件的 XML 与排版内核里**，屏幕上根本看不全。

## 四条铁律

1. **母版即产物** — 从模板复制/打开再另存，绝不从空白文档重建版式。
2. **用对排版内核** — 复杂 `.doc` 版式（旋转文本框/浮动表格/折页拼版）必须用 **Microsoft Word 本体**；纯表格 `.docx` 用 `python-docx` 复制母版 + 只改 run。
3. **最小改动** — 只碰文本，不碰结构。
4. **验收必须机器可判** — 每条要求绑定一个可执行校验，不靠感觉。

## 触发条件

- 「按这个 Word 模板 / 格式要求生成一份文档」
- 「必须和模板完全一致，不能差不多像」
- 「这些是格式要求文件（字体多少号、页边距多少），照它产出」
- 「和空白模板一模一样，只填内容」（官方表单 / 推荐书 / 计划书 / 封面）
- 「生成的 doc/docx 在 Word 里打开格式不对 / 边框没了 / 表格错位」

## 关键机制：要求文件 → 可机检规格表

用户一次给的可能是一叠异构文件（母版、格式规范、样例 PDF、评审标准、口述要求），必须**归一成一张规格表**：

| 对象 | 属性 | 期望值 | 容差 | 校验方式 |
|:-----|:-----|:-------|:-----|:---------|
| 全文正文 | 中文字体 (eastAsia) | 仿宋_GB2312 | 精确 | A 静态 |
| 全文正文 | 字号 | 14 pt | ±0.5 | A 静态 |
| 页面 | 页边距 | 2.54 / 3.17 cm | ±0.2 | A 静态 |
| 表格 1 | 列数 / gridSpan / vMerge | 与母版一致 | 精确 | A 静态 |
| 全文 | 总页数 | 4 | 精确 | B 渲染 |

- **A 类（可静态读取）** → 读 XML 属性断言相等 → `scripts/verify_docx_spec.py`
- **B 类（只能渲染判断）** → 导出 PDF → `pdftoppm -png -r 100` → 逐页读图对比基准 PDF

冲突仲裁优先级：**母版实物 > 明确的文字规范 > 惯例**（用户可覆盖，冲突必须显式记录）。

## 文件结构

```
word-template-fidelity/
├── SKILL.md                              # 完整工作流 + 铁律 + 反模式表
├── README.md
└── scripts/
    ├── verify_docx_spec.py               # A 类静态校验（字体/字号/边距/结构）
    └── gen_word_replace_applescript.py   # 生成驱动 Word 本体的 AppleScript
```

## 用法

```bash
# A 类：按 spec.json 逐条断言
python3 scripts/verify_docx_spec.py output.docx spec.json --print-signature

# 生成 Word 查找替换脚本
python3 scripts/gen_word_replace_applescript.py pairs.json master.doc new.doc new.pdf > build.applescript
osascript build.applescript   # 输出 MISS:<未命中的索引>
```

## 依赖

- macOS + Microsoft Word（复杂 `.doc` 版式必选）
- `python-docx`（`pip install python-docx`）
- poppler（`brew install poppler`，提供 `pdftoppm` / `pdftotext`）
- LibreOffice（可选，仅用于简单 `.doc→.docx` 一次性取样）

## 仓库

https://github.com/mynameisi/word-template-fidelity （private）
