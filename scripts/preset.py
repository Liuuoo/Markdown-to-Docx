#!/usr/bin/env python3
"""文档样式预设(preset)定义与默认值。

preset 是一个 dict, 描述整篇文档的排版规范。build_docx.py 读取它来驱动生成。
我(助手)可以根据用户的自然语言要求, 生成或修改一份 preset, 再交给 build_docx 使用。

字号单位: pt。长度单位: cm。字体分中文(chinese)与西文(western)。
"""
import copy
import json

# 中文学位论文默认预设
DEFAULT_PRESET = {
    "page": {
        "width_cm": 21.0, "height_cm": 29.7,
        "top_cm": 2.54, "bottom_cm": 2.54, "left_cm": 3.17, "right_cm": 3.17,
    },
    "body": {  # 正文
        "chinese": "宋体", "western": "Times New Roman",
        "size": 12, "line": 1.25, "first_indent_chars": 2,
    },
    "title": {  # 文档大标题(--title)
        "chinese": "黑体", "western": "Times New Roman",
        "size": 18, "bold": False, "align": "center",
    },
    "author": {
        "chinese": "宋体", "western": "Times New Roman", "size": 12, "align": "center",
    },
    # 各级标题: 重定义 Word 内置 Heading1/2/3, 保留大纲级别(进目录), 去主题引用与蓝色
    "headings": {
        1: {"chinese": "黑体", "western": "Times New Roman", "size": 16,
            "bold": False, "align": "center", "before": 9, "after": 9,
            "first_indent": False, "color": "000000"},
        2: {"chinese": "黑体", "western": "Times New Roman", "size": 14,
            "bold": False, "align": "left", "before": 6, "after": 6,
            "first_indent": False, "color": "000000"},
        3: {"chinese": "黑体", "western": "Times New Roman", "size": 12,
            "bold": False, "align": "left", "before": 3, "after": 3,
            "first_indent": False, "color": "000000"},
    },
    "table": {  # 三线表
        "chinese": "宋体", "western": "Times New Roman", "size": 10.5,
        "header_bold": True, "line_color": "000000",
        "top_sz": 12, "header_sz": 6, "bottom_sz": 12,  # 边框磅(1/8pt)
        "align": "center", "cell_line": 1.15,
    },
    "code": {
        "chinese": "宋体", "western": "Consolas", "size": 10.5,
        "line": 1.15, "left_indent_cm": 0.74,
    },
    "quote": {
        "chinese": "宋体", "western": "Times New Roman", "size": 12,
        "line": 1.25, "left_indent_cm": 0.74,
    },
    "list": {
        "chinese": "宋体", "western": "Times New Roman", "size": 12,
        "line": 1.25, "left_indent_cm": 0.74, "hanging_cm": 0.37,
    },
    "math": {
        "font": "Cambria Math", "size": 12,  # 与正文同号
    },
    "caption": {  # 图表题注
        "chinese": "楷体", "western": "Times New Roman", "size": 10.5,
    },
    "hyperlink": {  # 目录/超链接: 黑色无下划线, 避免变蓝
        "color": "000000", "underline": False,
    },
}


def get_default():
    return copy.deepcopy(DEFAULT_PRESET)


def merge(base, override):
    """深合并 override 到 base 的拷贝。"""
    out = copy.deepcopy(base)

    def _m(a, b):
        for k, v in b.items():
            if isinstance(v, dict) and isinstance(a.get(k), dict):
                _m(a[k], v)
            else:
                a[k] = v
    _m(out, override)
    return out


def load(path):
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    # headings 的 key 在 JSON 里是字符串, 转回 int
    if "headings" in data:
        data["headings"] = {int(k): v for k, v in data["headings"].items()}
    return merge(DEFAULT_PRESET, data)


def save(preset, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(preset, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "dump":
        print(json.dumps(DEFAULT_PRESET, ensure_ascii=False, indent=2))
