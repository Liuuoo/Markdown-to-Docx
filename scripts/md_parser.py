#!/usr/bin/env python3
"""轻量 Markdown 解析器(面向中文学位论文)。

把 Markdown 解析成结构化块列表, 供 build_docx.py 组装。
支持的块:
- 标题:        # / ## / ###            -> {"t":"h", "level":n, "text":..}
- 独立公式:    $$ ... $$ (可跨行)       -> {"t":"mathblock", "latex":..}
- 代码块:      ``` ... ```              -> {"t":"code", "lang":.., "text":..}
- 表格:        | a | b |  (含分隔行)     -> {"t":"table", "rows":[[..],..]}
- 无序/有序列表: - / 1.                  -> {"t":"list", "ordered":bool, "items":[..]}
- 引用:        > ...                     -> {"t":"quote", "text":..}
- 图片:        ![alt](path)              -> {"t":"image", "alt":.., "path":..}
- 普通段落                                -> {"t":"p", "text":..}

段落/列表项/标题内部的行内标记($..$ 公式、**粗**、*斜*)由 build_docx 二次解析。
"""
import re


def parse(md: str):
    lines = md.replace("\r\n", "\n").split("\n")
    blocks = []
    i, n = 0, len(lines)

    while i < n:
        line = lines[i]
        stripped = line.strip()

        # 空行
        if not stripped:
            i += 1
            continue

        # 独立公式 $$ ... $$
        if stripped.startswith("$$"):
            # 单行 $$..$$ 或多行
            body = stripped[2:]
            if body.endswith("$$") and len(body) >= 2:
                latex = body[:-2].strip()
                blocks.append({"t": "mathblock", "latex": latex})
                i += 1
                continue
            buf = [body]
            i += 1
            while i < n and "$$" not in lines[i]:
                buf.append(lines[i])
                i += 1
            if i < n:
                tail = lines[i].split("$$")[0]
                buf.append(tail)
                i += 1
            blocks.append({"t": "mathblock", "latex": "\n".join(buf).strip()})
            continue

        # 代码块 ```
        if stripped.startswith("```"):
            lang = stripped[3:].strip()
            buf = []
            i += 1
            while i < n and not lines[i].strip().startswith("```"):
                buf.append(lines[i])
                i += 1
            if i < n:
                i += 1  # 跳过结束 ```
            blocks.append({"t": "code", "lang": lang, "text": "\n".join(buf)})
            continue

        # 标题
        m = re.match(r"^(#{1,6})\s+(.*)$", stripped)
        if m:
            level = len(m.group(1))
            blocks.append({"t": "h", "level": level, "text": m.group(2).strip()})
            i += 1
            continue

        # 图片(整行)
        m = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)\s*$", stripped)
        if m:
            blocks.append({"t": "image", "alt": m.group(1), "path": m.group(2)})
            i += 1
            continue

        # 表格: 当前行含 | 且下一行是分隔行 |---|---|
        if "|" in line and i + 1 < n and re.match(r"^\s*\|?[\s:|-]+\|?\s*$", lines[i + 1]) and "-" in lines[i + 1]:
            rows = []
            header = _split_row(line)
            rows.append(header)
            i += 2  # 跳过表头和分隔行
            while i < n and "|" in lines[i] and lines[i].strip():
                rows.append(_split_row(lines[i]))
                i += 1
            blocks.append({"t": "table", "rows": rows})
            continue

        # 引用
        if stripped.startswith(">"):
            buf = []
            while i < n and lines[i].strip().startswith(">"):
                buf.append(re.sub(r"^\s*>\s?", "", lines[i]))
                i += 1
            blocks.append({"t": "quote", "text": " ".join(s.strip() for s in buf)})
            continue

        # 列表(连续的列表项)
        if re.match(r"^(\s*)([-*+]|\d+\.)\s+", line):
            items = []
            ordered = bool(re.match(r"^\s*\d+\.\s+", line))
            while i < n and re.match(r"^(\s*)([-*+]|\d+\.)\s+", lines[i]):
                txt = re.sub(r"^(\s*)([-*+]|\d+\.)\s+", "", lines[i])
                items.append(txt.strip())
                i += 1
            blocks.append({"t": "list", "ordered": ordered, "items": items})
            continue

        # 普通段落: 收集连续非空行(直到空行或其他块标记)
        buf = [stripped]
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if not nxt:
                break
            if re.match(r"^(#{1,6}\s|```|>|\s*([-*+]|\d+\.)\s)", lines[i]):
                break
            if nxt.startswith("$$"):
                break
            if "|" in lines[i] and i + 1 < n and re.match(r"^\s*\|?[\s:|-]+\|?\s*$", lines[i + 1]):
                break
            buf.append(nxt)
            i += 1
        blocks.append({"t": "p", "text": " ".join(buf)})

    return blocks


def _split_row(line: str):
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


if __name__ == "__main__":
    import json
    import sys
    src = open(sys.argv[1], encoding="utf-8").read() if len(sys.argv) > 1 else """# 标题1
## 标题2
正文段落, 含公式 $E=mc^2$ 和 **粗体**。

$$\\eta = \\frac{a}{b}$$

- 项一
- 项二

1. 第一
2. 第二

| 列A | 列B |
|-----|-----|
| 1   | 2   |

```c
int x = 0;
```

> 引用内容
"""
    for b in parse(src):
        print(json.dumps(b, ensure_ascii=False))
