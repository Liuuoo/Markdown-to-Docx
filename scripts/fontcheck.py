#!/usr/bin/env python3
"""字体存在性检测(跨平台, 纯 Python, 无第三方依赖)。

Word 找不到字体时会静默替换成别的字体, 用户不会收到任何提示, 排版就悄悄错了。
本模块在生成后扫描系统已装字体, 对比 preset 用到的字体, 缺失则警告(不阻断)。

做法: 扫描各平台字体目录, 解析 TTF/TTC/OTF 的 name 表(name ID 1/4/16),
取出字体内部记录的中英文名(含中文名), 汇成"已装字体名"集合。
比文件名匹配可靠(文件名常是英文如 Songti.ttc, 而 preset 写中文"宋体")。
"""
import os
import struct
import sys
from pathlib import Path


def _font_dirs():
    home = Path.home()
    if sys.platform == "darwin":
        return [home / "Library/Fonts", Path("/Library/Fonts"),
                Path("/System/Library/Fonts"),
                Path("/System/Library/Fonts/Supplemental")]
    if sys.platform.startswith("win"):
        win = Path(os.environ.get("WINDIR", r"C:\Windows"))
        return [win / "Fonts", home / "AppData/Local/Microsoft/Windows/Fonts"]
    # linux / 其他
    return [home / ".fonts", home / ".local/share/fonts",
            Path("/usr/share/fonts"), Path("/usr/local/share/fonts")]


def _read_name_table(data, name_off):
    """从 sfnt 的 name 表偏移处解析出所有 name 记录字符串。"""
    names = set()
    try:
        fmt, count, str_off = struct.unpack(">HHH", data[name_off:name_off + 6])
        rec = name_off + 6
        storage = name_off + str_off
        for i in range(count):
            pid, eid, lid, nid, length, off = struct.unpack(
                ">HHHHHH", data[rec + i * 12: rec + i * 12 + 12])
            if nid not in (1, 4, 16):  # 1=family 4=full 16=typographic family
                continue
            s = data[storage + off: storage + off + length]
            # platformID 3(Windows)/0(Unicode) 用 UTF-16BE; 1(Mac) 多为 latin/mac
            try:
                if pid in (0, 3):
                    txt = s.decode("utf-16-be", "ignore")
                else:
                    txt = s.decode("latin-1", "ignore")
            except Exception:
                continue
            txt = txt.strip("\x00").strip()
            if txt:
                names.add(txt)
    except Exception:
        pass
    return names


def _names_from_sfnt(data, base=0):
    """解析单个 sfnt(ttf/otf)字体, 返回其 name 集合。"""
    try:
        num_tables = struct.unpack(">H", data[base + 4: base + 6])[0]
        rec = base + 12
        for i in range(num_tables):
            tag = data[rec + i * 16: rec + i * 16 + 4]
            if tag == b"name":
                off = struct.unpack(">I", data[rec + i * 16 + 8: rec + i * 16 + 12])[0]
                return _read_name_table(data, off)
    except Exception:
        pass
    return set()


def _names_from_file(path):
    names = set()
    try:
        data = path.read_bytes()
    except Exception:
        return names
    tag = data[:4]
    if tag == b"ttcf":  # TrueType Collection: 多个字体
        try:
            n = struct.unpack(">I", data[8:12])[0]
            offs = struct.unpack(">" + "I" * n, data[12:12 + 4 * n])
            for o in offs:
                names |= _names_from_sfnt(data, o)
        except Exception:
            pass
    else:  # 单字体 ttf/otf
        names |= _names_from_sfnt(data, 0)
    return names


_CACHE = None


def installed_font_names():
    """返回系统已装字体的所有内部名(中英文)集合, 带缓存。"""
    global _CACHE
    if _CACHE is not None:
        return _CACHE
    names = set()
    exts = {".ttf", ".ttc", ".otf", ".otc"}
    for d in _font_dirs():
        if not d.exists():
            continue
        for root, _, files in os.walk(d):
            for f in files:
                if Path(f).suffix.lower() in exts:
                    names |= _names_from_file(Path(root) / f)
    _CACHE = names
    return names


# preset 里可能用的"逻辑名" -> 系统中可能的等价名(别名), 任一命中即视为可用
_ALIASES = {
    "宋体": ["宋体", "SimSun", "Songti SC", "Songti TC", "STSong", "Songti", "NSimSun"],
    "黑体": ["黑体", "SimHei", "STHeiti", "Heiti SC", "Heiti TC", "Hiragino Sans GB"],
    "楷体": ["楷体", "楷体_GB2312", "KaiTi", "STKaiti", "Kaiti SC", "Kaiti TC"],
    "仿宋": ["仿宋", "仿宋_GB2312", "FangSong", "STFangsong", "FangSong_GB2312"],
    "Times New Roman": ["Times New Roman", "Times", "TimesNewRomanPSMT"],
    "Cambria Math": ["Cambria Math", "Cambria"],
    "Consolas": ["Consolas", "Menlo", "DejaVu Sans Mono", "Courier New"],
}


def _collect_preset_fonts(preset):
    fonts = set()
    for key, val in preset.items():
        if isinstance(val, dict):
            for fk in ("chinese", "western", "font"):
                if fk in val and val[fk]:
                    fonts.add(val[fk])
            # headings 是 {1:{..},2:{..}} 嵌套
            for sub in val.values():
                if isinstance(sub, dict):
                    for fk in ("chinese", "western", "font"):
                        if fk in sub and sub[fk]:
                            fonts.add(sub[fk])
    return fonts


def check_fonts(preset):
    """返回 (missing, installed_count)。missing 为缺失字体名列表(已考虑别名)。"""
    installed = installed_font_names()
    want = _collect_preset_fonts(preset)
    missing = []
    for f in sorted(want):
        candidates = _ALIASES.get(f, [f])
        if not any(c in installed for c in candidates):
            missing.append(f)
    return missing, len(installed)


def warn_if_missing(preset, stream=sys.stderr):
    """打印警告(不阻断)。返回缺失列表。"""
    missing, n = check_fonts(preset)
    if missing:
        print("⚠ 字体提醒: 以下预设字体在本机未检测到, Word/WPS 打开时可能被静默替换:",
              file=stream)
        for f in missing:
            alts = "、".join(_ALIASES.get(f, [f])[:3])
            print(f"    · {f}  (可接受的等价名: {alts} …)", file=stream)
        print("  解决: 安装对应字体, 或用 --preset 把该字体改成本机已有的字体。",
              file=stream)
    return missing


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).parent))
    import preset as pm
    ps = pm.get_default()
    miss, n = check_fonts(ps)
    print(f"已扫描系统字体名 {n} 个")
    print("预设用到的字体:", sorted(_collect_preset_fonts(ps)))
    print("缺失:", miss if miss else "无, 全部可用")
