# Markdown → Docx (md2thesis)

把 Markdown 转成**排版正式的中文学位论文风格 Word 文档**。纯 Python 实现，**不依赖 pandoc**。

为解决用 pandoc 转换学术文档时的几个老问题而生：公式字号变大/字体乱、表格被渲染成蓝色、标题不进大纲无法生成目录。本工具改为自写解析与组装，每个文本片段直接写死字体字号，标题套用并重定义 Word 内置样式，从根上避免这些问题。

## 特性

- **标题进样式库 + 大纲**：重定义 Word 内置「标题 1/2/3」样式（改黑体、纯黑、去除主题字体引用，保留大纲级别）。标题进导航窗格，可在 Word 里一键「引用 → 目录」生成目录，且不会变成蓝色或等线体。
- **目录不变蓝**：同时把超链接样式改为黑色无下划线，生成的目录链接是黑色。
- **可编辑公式**：`$...$` 行内、`$$...$$` 独立公式自写 LaTeX→OMML 转换，生成 Word 原生可编辑公式，字号字体完全可控、不会变大。
- **三线表不染蓝**：表格不引用任何内置样式，自画黑色三线、无填充无内框。
- **样式预设可定制**：全部排版参数集中在预设（preset）中，可用一份 JSON 覆盖默认值，按需定制字体、字号、行距、边距等。
- **字体缺失检测**：生成后扫描系统字体，若预设用到的字体未安装则警告（不阻断），避免 Word 静默替换字体却无人察觉。

## 依赖

- Python 3.9+
- [`python-docx`](https://pypi.org/project/python-docx/)：`pip3 install --user python-docx`
- （可选，仅 `--pdf` 需要）LibreOffice：`brew install --cask libreoffice`

## 安装

作为 Claude Code skill 使用，把整个文件夹放到 skills 目录：

```bash
git clone https://github.com/Liuuoo/Markdown-to-Docx.git ~/.claude/skills/md2thesis
pip3 install --user python-docx
```

之后在 Claude Code 里直接说「把 X.md 转成论文格式」即可自动调用。

也可脱离 Claude，直接当命令行工具用。

## 用法

```bash
# 通过入口脚本
bash scripts/build.sh 论文.md -o 论文.docx --title "标题" --author "作者"

# 或直接调 Python
python3 scripts/build_docx.py 论文.md -o 论文.docx --title "标题" --author "作者"

# 同时输出 PDF（需 LibreOffice）
bash scripts/build.sh 论文.md --pdf

# 使用自定义样式预设
python3 scripts/build_docx.py 论文.md -o 论文.docx --preset 我的预设.json
```

生成 docx 后，在 Word 里「引用 → 目录 → 自动目录」即可生成黑色可跳转目录（标题已进大纲）。

## 支持的 Markdown 元素

| 元素 | 写法 | 输出 |
|------|------|------|
| 标题 | `#` `##` `###` | 黑体三号/四号/小四，套内置标题样式 |
| 行内强调 | `**粗**` `*斜*` `` `等宽` `` | 对应字形 |
| 行内公式 | `$E=mc^2$` | Word 可编辑公式 |
| 独立公式 | `$$...$$` | 居中可编辑公式 |
| 列表 | `- 项` / `1. 项` | 无序/有序列表 |
| 表格 | `\| a \| b \|` | 三线表 |
| 代码块 | ` ``` ` 围栏 | Consolas 等宽 |
| 引用 | `> 文本` | 缩进引用 |
| 图片 | `![题注](路径)` | 居中图片 + 题注 |

## 默认排版规范（中文学位论文）

- 正文：中文宋体 / 西文 Times New Roman，小四（12pt），1.25 倍行距，首行缩进 2 字符
- 一级标题：黑体三号（16pt）居中
- 二级标题：黑体四号（14pt）顶格
- 三级标题：黑体小四（12pt）
- 表格：三线表（黑色顶/底线，无填充无内框）
- 代码块：Consolas 五号（10.5pt）
- 公式：Cambria Math 小四
- 页面：A4，上下 2.54cm，左右 3.17cm

## 自定义样式预设

所有排版参数在 `scripts/preset.py` 的 `DEFAULT_PRESET` 中定义。写一份只含**要修改项**的 JSON，其余自动继承默认值：

```json
{
  "body": {"chinese": "仿宋", "size": 14, "line": 1.5},
  "headings": {"1": {"chinese": "楷体", "size": 22, "align": "center"}}
}
```

中文字号对照：初号 42 · 小初 36 · 一号 26 · 二号 22 · 小二 18 · 三号 16 · 小三 15 · 四号 14 · 小四 12 · 五号 10.5 · 小五 9。

可配置块：`page`（边距）· `body`（正文）· `title` / `author` · `headings`（1/2/3 级）· `table` · `code` · `quote` · `list` · `math`（公式字号）· `caption` · `hyperlink`。

`presets/thesis_zh.json` 是导出的默认预设，可作为修改起点。

## 公式支持范围

上下标 · `\frac` · `\sqrt[n]{}` · `\int` `\sum` `\prod` `\lim`（含上下限与被作用式）· 希腊字母 · `\times` `\leq` `\rightarrow` `\infty` 等运算符 · `\left( \right)` 定界符 · `\sin` `\cos` `\log` 等函数 · `\hat` `\bar` `\vec` 等重音符。

## 字体注意事项

预设默认使用宋体、黑体、楷体、仿宋、Times New Roman、Cambria Math、Consolas。**Word 找不到字体时会静默替换**，因此：

- 工具会在生成后检测并警告缺失的字体。
- 不同系统字体可用性不同：Windows 通常有宋体/黑体，楷体/仿宋有时缺；macOS 自带宋体/黑体（映射到 Songti/STHeiti），但常缺楷体与 Cambria Math；Linux 多数中文字体需自行安装。
- 若提示缺字体，可安装对应字体，或用 `--preset` 把字体改成本机已有的。

## 文件结构

```
SKILL.md                  Claude Code skill 描述
scripts/
  build.sh                命令行入口
  build_docx.py           组装 docx（重定义标题样式、读预设、块渲染、行内解析）
  md_parser.py            Markdown 解析器
  latex2omml.py           LaTeX→OMML 公式转换器（可单独运行自测）
  preset.py               样式预设定义/合并/读写
  fontcheck.py            字体存在性检测
presets/
  thesis_zh.json          默认学位论文预设
```

## 许可

MIT
