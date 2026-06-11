---
name: md2thesis
description: 把 Markdown 转成中文学位论文风格的正式 DOCX(可选转 PDF)。纯 Python 实现, 不依赖 pandoc。标题套 Word 内置"标题1/2/3"样式(进大纲、可一键生成目录、纯黑不变蓝)。公式 $..$/$$..$$ 自写 LaTeX→OMML 可编辑公式。支持用样式预设(preset)定制排版: 用户用自然语言描述要求, 助手生成 preset 再驱动生成。当用户想把 .md 转成排版正式的学术论文 Word、需要宋体小四正文+黑体标题+三线表+可编辑公式+目录时使用。
---

# md2thesis — Markdown 转中文学位论文 DOCX

把任意 Markdown 转成排版正式的中文学位论文 docx。**纯 Python, 不用 pandoc**。

## 关键特性

- **标题进样式库 + 大纲**: 重定义 Word 内置 Heading1/2/3 样式(改黑体、纯黑、去主题字体引用,
  保留 outlineLvl)。标题段落套这些样式 → 进导航窗格、可一键生成目录、不变等线、不变蓝。
- **目录不变蓝**: 同时把 Hyperlink 样式改为黑色无下划线, 生成目录后链接是黑色。
- **可编辑公式**: `$..$` / `$$..$$` 自写 LaTeX→OMML 转换, 字号字体写死, 不变大。
- **三线表不染蓝**: 表格不引用任何内置样式, 自画黑色三线、无填充。
- **样式预设可定制**: 全部排版参数集中在 preset, 可按用户要求定制。

## 用法

```bash
bash ~/.claude/skills/md2thesis/scripts/build.sh <输入.md> [-o 输出.docx] [--title T] [--author A] [--pdf]
# 或直接:
python3 scripts/build_docx.py <输入.md> -o 输出.docx --title T --author A [--preset 预设.json]
```

生成 docx 后, 在 Word 里「引用 → 目录 → 自动目录」即可生成黑色可跳转目录(标题已进大纲)。

## 提示词 → 预设 → 文档(定制排版的工作流)

当用户用自然语言描述排版要求(如"正文用仿宋三号, 一级标题居中楷体二号, 行距1.5"),
助手应:
1. 读 `scripts/preset.py` 的 `DEFAULT_PRESET` 了解可配置项与默认值。
2. 把用户要求翻译成一个**覆盖项 JSON**(只写要改的, 其余继承默认), 字号用 pt、长度用 cm。
   中文字号对照: 初号42 小初36 一号26 二号22 小二18 三号16 小三15 四号14 小四12 五号10.5 小五9。
3. 用 `preset.load()` 会自动深合并默认值; 把覆盖 JSON 存成文件, 传 `--preset` 生成。
4. 生成后让用户在 Word 里确认观感, 按反馈调整 JSON 再重生成。

预设可配置块: `page`(边距) `body`(正文) `title`/`author` `headings`(1/2/3级)
`table`(三线表) `code` `quote` `list` `math`(公式字号) `caption` `hyperlink`。

示例覆盖 JSON:
```json
{
  "body": {"chinese": "仿宋", "size": 14, "line": 1.5},
  "headings": {"1": {"chinese": "楷体", "size": 22, "align": "center"}}
}
```
`presets/thesis_zh.json` 是导出的默认学位论文预设, 可作为修改起点。

## 支持的 Markdown 元素

标题 `#`~`###` · 段落 · 行内 `**粗**`/`*斜*`/`` `等宽` `` · 行内公式 `$..$` ·
独立公式 `$$..$$` · 有序/无序列表 · 管道表格(→三线表) · 代码块 · 引用 `>` · 图片 `![题注](路径)`

## 默认排版规范(中文学位论文)

正文 宋体/Times New Roman 小四 1.25倍 首行缩进2字符 · 一级标题 黑体三号居中 ·
二级 黑体四号 · 三级 黑体小四 · 三线表 · 代码 Consolas 五号 · 公式 Cambria Math 小四 ·
页面 A4 上下2.54cm 左右3.17cm

## 公式支持范围(latex2omml.py)

上下标 · `\frac` · `\sqrt[n]{}` · `\int \sum \prod \lim`(含上下限与被作用式) ·
希腊字母 · `\times \leq \rightarrow \infty` 等运算符 · `\left( \right)` 定界符 ·
`\sin \cos \log` 等函数 · `\hat \bar \vec` 等重音符。每个公式 run 写死 Cambria Math + 字号。

## 文件结构

- `scripts/build.sh` — 命令行入口
- `scripts/build_docx.py` — 组装 docx(重定义标题样式、读 preset、块渲染、行内解析)
- `scripts/md_parser.py` — Markdown 解析器
- `scripts/latex2omml.py` — LaTeX→OMML 公式转换器(可单独运行自测)
- `scripts/preset.py` — 样式预设定义/合并/读写
- `scripts/fontcheck.py` — 字体存在性检测(生成后扫描系统字体, 缺失则警告, 不阻断)
- `presets/thesis_zh.json` — 默认学位论文预设

## 依赖

- `python-docx` (`pip3 install --user python-docx`) — 必需
- LibreOffice (`brew install --cask libreoffice`) — 仅 `--pdf` 需要
