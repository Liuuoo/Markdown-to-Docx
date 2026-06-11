#!/usr/bin/env python3
"""LaTeX 数学公式 -> Word OMML (Office Math Markup Language)。

自写转换器, 不依赖 pandoc。支持论文常见语法:
- 上标 ^ / 下标 _ (含 {} 分组)
- 分式 \\frac{a}{b}
- 根号 \\sqrt{x} / \\sqrt[n]{x}
- 积分 \\int \\iint \\oint, 求和 \\sum, 连乘 \\prod, 极限 \\lim (含上下限)
- 希腊字母 \\alpha \\beta ... 和常见运算符 \\times \\leq \\rightarrow \\infty ...
- \\left( \\right) 等定界符
- 函数名 \\sin \\cos \\log 等

生成的每个 <m:r> 都写死字号(默认12pt)与数学字体(Cambria Math),
不依赖文档主题或样式, 因此 Word 渲染稳定、字号不会变大。

对外接口: latex_to_omath(latex:str, size_pt=12.0) -> str  返回 <m:oMath>...</m:oMath>
"""
import re

MATH_FONT = "Cambria Math"

# 希腊字母与符号 -> Unicode
SYMBOLS = {
    r"\alpha": "α", r"\beta": "β", r"\gamma": "γ", r"\delta": "δ",
    r"\epsilon": "ε", r"\varepsilon": "ε", r"\zeta": "ζ", r"\eta": "η",
    r"\theta": "θ", r"\vartheta": "ϑ", r"\iota": "ι", r"\kappa": "κ",
    r"\lambda": "λ", r"\mu": "μ", r"\nu": "ν", r"\xi": "ξ", r"\pi": "π",
    r"\rho": "ρ", r"\sigma": "σ", r"\tau": "τ", r"\upsilon": "υ",
    r"\phi": "φ", r"\varphi": "φ", r"\chi": "χ", r"\psi": "ψ", r"\omega": "ω",
    r"\Gamma": "Γ", r"\Delta": "Δ", r"\Theta": "Θ", r"\Lambda": "Λ",
    r"\Xi": "Ξ", r"\Pi": "Π", r"\Sigma": "Σ", r"\Phi": "Φ", r"\Psi": "Ψ",
    r"\Omega": "Ω",
    r"\times": "×", r"\div": "÷", r"\pm": "±", r"\mp": "∓",
    r"\cdot": "·", r"\ast": "∗", r"\star": "⋆",
    r"\leq": "≤", r"\le": "≤", r"\geq": "≥", r"\ge": "≥",
    r"\neq": "≠", r"\ne": "≠", r"\approx": "≈", r"\equiv": "≡",
    r"\sim": "∼", r"\propto": "∝", r"\cong": "≅",
    r"\rightarrow": "→", r"\to": "→", r"\leftarrow": "←",
    r"\Rightarrow": "⇒", r"\Leftarrow": "⇐", r"\leftrightarrow": "↔",
    r"\Leftrightarrow": "⇔", r"\uparrow": "↑", r"\downarrow": "↓",
    r"\infty": "∞", r"\partial": "∂", r"\nabla": "∇",
    r"\forall": "∀", r"\exists": "∃", r"\in": "∈", r"\notin": "∉",
    r"\subset": "⊂", r"\subseteq": "⊆", r"\cup": "∪", r"\cap": "∩",
    r"\emptyset": "∅", r"\angle": "∠", r"\perp": "⊥", r"\parallel": "∥",
    r"\cdots": "⋯", r"\ldots": "…", r"\dots": "…", r"\vdots": "⋮",
    r"\degree": "°", r"\circ": "∘", r"\bullet": "•",
    r"\Box": "□", r"\square": "□", r"\triangle": "△",
    r"\%": "%", r"\&": "&", r"\#": "#", r"\$": "$", r"\_": "_",
    r"\{": "{", r"\}": "}", r"\backslash": "\\",
    r"\quad": " ", r"\qquad": "  ", r"\,": " ", r"\;": " ",
}

# 大型算符(带上下限)
BIGOPS = {
    r"\sum": "∑", r"\prod": "∏", r"\coprod": "∐",
    r"\int": "∫", r"\iint": "∬", r"\iiint": "∭", r"\oint": "∮",
    r"\bigcup": "⋃", r"\bigcap": "⋂", r"\bigoplus": "⨁", r"\bigotimes": "⨂",
    r"\lim": "lim", r"\max": "max", r"\min": "min", r"\sup": "sup", r"\inf": "inf",
}

# 函数名(直立)
FUNCS = [r"\sin", r"\cos", r"\tan", r"\cot", r"\sec", r"\csc",
         r"\arcsin", r"\arccos", r"\arctan", r"\sinh", r"\cosh", r"\tanh",
         r"\log", r"\ln", r"\lg", r"\exp", r"\det", r"\dim", r"\gcd",
         r"\deg", r"\arg", r"\ker", r"\mod"]


def _esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ===== 词法分析: 把 latex 串切成 token 列表 =====
def tokenize(s: str):
    toks = []
    i, n = 0, len(s)
    # 命令名: 先匹配长的(\iiint 在 \int 前), 再按字典序
    cmds = sorted(list(SYMBOLS) + list(BIGOPS) + FUNCS +
                  [r"\frac", r"\sqrt", r"\left", r"\right", r"\hat", r"\bar",
                   r"\vec", r"\dot", r"\ddot", r"\tilde", r"\overline",
                   r"\text", r"\mathrm", r"\mathbf", r"\boldsymbol"],
                  key=len, reverse=True)
    while i < n:
        c = s[i]
        if c.isspace():
            i += 1
            continue
        if c == "\\":
            matched = None
            for cmd in cmds:
                if s.startswith(cmd, i):
                    # 命令后若是字母, 须保证不是更长命令的前缀被截断
                    nxt = i + len(cmd)
                    if cmd[-1].isalpha() and nxt < n and s[nxt].isalpha():
                        continue
                    matched = cmd
                    break
            if matched:
                toks.append(("cmd", matched))
                i += len(matched)
            else:
                # 未知命令: 跳过反斜杠, 原样输出后续字母
                m = re.match(r"\\([a-zA-Z]+)", s[i:])
                if m:
                    toks.append(("text", m.group(1)))
                    i += m.end()
                else:
                    i += 1
            continue
        if c == "{":
            toks.append(("{", "{")); i += 1; continue
        if c == "}":
            toks.append(("}", "}")); i += 1; continue
        if c == "^":
            toks.append(("^", "^")); i += 1; continue
        if c == "_":
            toks.append(("_", "_")); i += 1; continue
        if c == "&":
            toks.append(("&", "&")); i += 1; continue
        if c.isdigit():
            j = i
            while j < n and (s[j].isdigit() or s[j] == "."):
                j += 1
            toks.append(("num", s[i:j])); i = j; continue
        if c.isalpha():
            toks.append(("var", c)); i += 1; continue
        # 其他单字符(运算符、括号、标点)
        toks.append(("op", c)); i += 1
    return toks


# ===== 语法分析: token -> 抽象节点 =====
# 节点用 dict 表示: {"t": 类型, ...}
class Parser:
    def __init__(self, toks):
        self.toks = toks
        self.i = 0

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def next(self):
        t = self.toks[self.i]
        self.i += 1
        return t

    def parse_group(self):
        """解析一个 {} 分组或单个原子, 返回节点列表。"""
        kind, val = self.peek()
        if kind == "{":
            self.next()
            nodes = self.parse_until("}")
            if self.peek()[0] == "}":
                self.next()
            return nodes
        return [self.parse_atom()]

    def parse_until(self, stop):
        nodes = []
        while True:
            kind, val = self.peek()
            if kind is None or kind == stop:
                break
            nodes.append(self.parse_scripted())
        return nodes

    def parse_scripted(self):
        """解析一个原子, 并附加其上标/下标。"""
        base = self.parse_atom()
        sub = sup = None
        while True:
            kind, val = self.peek()
            if kind == "_":
                self.next()
                sub = self.parse_group()
            elif kind == "^":
                self.next()
                sup = self.parse_group()
            else:
                break
        if sub is None and sup is None:
            return base
        return {"t": "script", "base": base, "sub": sub, "sup": sup}

    def parse_atom(self):
        kind, val = self.next()
        if kind == "cmd":
            return self.parse_cmd(val)
        if kind in ("num", "var", "op", "text", "&"):
            return {"t": "run", "kind": kind, "text": val}
        if kind == "{":
            # 裸 { 当分组
            nodes = self.parse_until("}")
            if self.peek()[0] == "}":
                self.next()
            return {"t": "group", "nodes": nodes}
        return {"t": "run", "kind": "op", "text": val or ""}

    def parse_cmd(self, cmd):
        if cmd == r"\frac":
            num = self.parse_group()
            den = self.parse_group()
            return {"t": "frac", "num": num, "den": den}
        if cmd == r"\sqrt":
            deg = None
            if self.peek()[0] == "op" and self.peek()[1] == "[":
                # \sqrt[n]{x}
                self.next()
                deg = []
                while self.peek()[0] is not None and not (self.peek()[0] == "op" and self.peek()[1] == "]"):
                    deg.append(self.parse_scripted())
                if self.peek()[0] == "op" and self.peek()[1] == "]":
                    self.next()
            rad = self.parse_group()
            return {"t": "sqrt", "deg": deg, "rad": rad}
        if cmd in (r"\hat", r"\bar", r"\vec", r"\dot", r"\ddot", r"\tilde", r"\overline"):
            acc = {r"\hat": "0302", r"\bar": "0304", r"\overline": "0304",
                   r"\vec": "20D7", r"\dot": "0307", r"\ddot": "0308",
                   r"\tilde": "0303"}[cmd]
            base = self.parse_group()
            return {"t": "accent", "chr": acc, "base": base}
        if cmd in (r"\text", r"\mathrm", r"\mathbf", r"\boldsymbol"):
            base = self.parse_group()
            return {"t": "upright", "nodes": base}
        if cmd in BIGOPS:
            return {"t": "bigop", "sym": BIGOPS[cmd], "cmd": cmd}
        if cmd in FUNCS:
            return {"t": "func", "name": cmd[1:]}
        if cmd == r"\left":
            op = self.next()[1]
            inner = []
            while True:
                kind, val = self.peek()
                if kind is None:
                    break
                if kind == "cmd" and val == r"\right":
                    self.next()
                    close = self.next()[1] if self.peek()[0] is not None else ")"
                    break
                inner.append(self.parse_scripted())
            else:
                close = ")"
            return {"t": "delim", "open": op, "close": close, "nodes": inner}
        if cmd in SYMBOLS:
            return {"t": "run", "kind": "sym", "text": SYMBOLS[cmd]}
        return {"t": "run", "kind": "text", "text": cmd[1:]}


# ===== OMML 渲染: 节点 -> XML =====
class Renderer:
    def __init__(self, size_pt=12.0):
        self.sz = str(int(round(size_pt * 2)))

    def rpr(self, italic=True):
        """每个 m:r 的属性: 写死字体+字号, 不依赖主题/样式。
        默认变量用斜体(数学惯例), 数字/函数名直立。"""
        sty = "" if italic else '<w:i w:val="0"/>'
        return (f'<w:rPr>{sty}'
                f'<w:rFonts w:ascii="{MATH_FONT}" w:hAnsi="{MATH_FONT}" w:cs="{MATH_FONT}"/>'
                f'<w:sz w:val="{self.sz}"/><w:szCs w:val="{self.sz}"/></w:rPr>')

    def run(self, text, italic=True):
        return f"<m:r>{self.rpr(italic)}<m:t>{_esc(text)}</m:t></m:r>"

    def render_nodes(self, nodes):
        return "".join(self.render(n) for n in nodes)

    def render(self, node):
        if isinstance(node, list):
            return self.render_nodes(node)
        t = node["t"]
        if t == "run":
            kind = node["kind"]
            txt = node["text"]
            # 数字、运算符、符号直立; 变量斜体
            italic = (kind == "var")
            return self.run(txt, italic=italic)
        if t == "group":
            return self.render_nodes(node["nodes"])
        if t == "upright":
            return "".join(
                self.run(n.get("text", ""), italic=False) if n.get("t") == "run"
                else self.render(n) for n in node["nodes"])
        if t == "frac":
            return (f'<m:f><m:fPr><m:type m:val="bar"/></m:fPr>'
                    f'<m:num>{self.render_nodes(node["num"])}</m:num>'
                    f'<m:den>{self.render_nodes(node["den"])}</m:den></m:f>')
        if t == "sqrt":
            inner = self.render_nodes(node["rad"])
            if node["deg"]:
                return (f'<m:rad><m:deg>{self.render_nodes(node["deg"])}</m:deg>'
                        f'<m:e>{inner}</m:e></m:rad>')
            return (f'<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr>'
                    f'<m:deg/><m:e>{inner}</m:e></m:rad>')
        if t == "script":
            base = self.render(node["base"])
            sub, sup = node["sub"], node["sup"]
            if sub is not None and sup is not None:
                return (f'<m:sSubSup><m:e>{base}</m:e>'
                        f'<m:sub>{self.render_nodes(sub)}</m:sub>'
                        f'<m:sup>{self.render_nodes(sup)}</m:sup></m:sSubSup>')
            if sub is not None:
                return (f'<m:sSub><m:e>{base}</m:e>'
                        f'<m:sub>{self.render_nodes(sub)}</m:sub></m:sSub>')
            return (f'<m:sSup><m:e>{base}</m:e>'
                    f'<m:sup>{self.render_nodes(sup)}</m:sup></m:sSup>')
        if t == "accent":
            return (f'<m:acc><m:accPr><m:chr m:val="&#x{node["chr"]};"/></m:accPr>'
                    f'<m:e>{self.render_nodes(node["base"])}</m:e></m:acc>')
        if t == "func":
            return self.run(node["name"], italic=False)
        if t == "bigop":
            return self.run(node["sym"], italic=False)
        if t == "delim":
            return (f'<m:d><m:dPr>'
                    f'<m:begChr m:val="{_esc(node["open"])}"/>'
                    f'<m:endChr m:val="{_esc(node["close"])}"/></m:dPr>'
                    f'<m:e>{self.render_nodes(node["nodes"])}</m:e></m:d>')
        return ""


def _attach_bigop_limits(nodes):
    """把大型算符(sum/int/lim...)与其上下限及【被作用式】收拢成 nary/limLow 结构。
    被作用式 = 算符之后、直到遇到顶层关系运算符(= < > ≤ ≥ ≠ ≈)或结尾 的所有节点。
    这样 \\int_0^t I d\\tau 的被积式 I d\\tau 会正确落入积分号内, 不会断开。
    """
    REL = {"=", "<", ">", "≤", "≥", "≠", "≈", "≡", "→", "⇒", "∝"}
    LIMFUNCS = (r"\lim", r"\max", r"\min", r"\sup", r"\inf")

    def is_rel(node):
        return node.get("t") == "run" and node.get("text") in REL

    out = []
    i = 0
    while i < len(nodes):
        node = nodes[i]
        op = sub = sup = None
        if node.get("t") == "script" and node["base"].get("t") == "bigop":
            op, sub, sup = node["base"], node.get("sub"), node.get("sup")
        elif node.get("t") == "bigop":
            op = node
        if op is not None:
            # 收集被作用式
            j = i + 1
            body = []
            while j < len(nodes) and not is_rel(nodes[j]):
                body.append(nodes[j])
                j += 1
            kind = "limlow" if op["cmd"] in LIMFUNCS else "nary"
            out.append({"t": kind, "op": op, "sub": sub, "sup": sup, "body": body})
            i = j
        else:
            out.append(node)
            i += 1
    return out


class Renderer2(Renderer):
    def render(self, node):
        if isinstance(node, dict) and node.get("t") == "nary":
            op, sub, sup, body = node["op"], node.get("sub"), node.get("sup"), node.get("body")
            subxml = f'<m:sub>{self.render_nodes(sub)}</m:sub>' if sub else '<m:sub/>'
            supxml = f'<m:sup>{self.render_nodes(sup)}</m:sup>' if sup else '<m:sup/>'
            loc = "undOvr" if op["cmd"] in (r"\sum", r"\prod", r"\coprod") else "subSup"
            sub_hide = '<m:subHide m:val="1"/>' if not sub else ""
            sup_hide = '<m:supHide m:val="1"/>' if not sup else ""
            pr = ('<m:naryPr>'
                  f'<m:chr m:val="{op["sym"]}"/>'
                  f'<m:limLoc m:val="{loc}"/>'
                  '<m:grow m:val="1"/>'
                  f'{sub_hide}{sup_hide}'
                  '</m:naryPr>')
            e = f'<m:e>{self.render_nodes(body)}</m:e>'
            return f'<m:nary>{pr}{subxml}{supxml}{e}</m:nary>'
        if isinstance(node, dict) and node.get("t") == "limlow":
            op, sub, body = node["op"], node.get("sub"), node.get("body")
            funcname = self.run(op["sym"], italic=False)
            limpart = self.render_nodes(sub) if sub else ""
            ll = (f'<m:limLow><m:limLowPr/><m:e>{funcname}</m:e>'
                  f'<m:lim>{limpart}</m:lim></m:limLow>')
            return ll + self.render_nodes(body)
        return super().render(node)


def latex_to_omath(latex: str, size_pt: float = 12.0) -> str:
    """LaTeX 公式串 -> <m:oMath>...</m:oMath> 字符串。"""
    toks = tokenize(latex)
    parser = Parser(toks)
    nodes = parser.parse_until(None)
    nodes = _attach_bigop_limits(nodes)
    body = Renderer2(size_pt).render_nodes(nodes)
    return f'<m:oMath>{body}</m:oMath>'


if __name__ == "__main__":
    import sys
    tests = [
        "E=mc^2",
        "x_{i}",
        r"\eta = \frac{P_{out}}{P_{in}} \times 100\%",
        r"SOC_t = SOC_0 - \frac{1}{C_N}\int_{0}^{t} I(\tau)\,d\tau",
        r"\sqrt{a^2+b^2}",
        r"\sum_{i=1}^{n} x_i",
        r"\alpha + \beta \leq \gamma",
        r"\lim_{x \to \infty} \frac{1}{x} = 0",
    ]
    for t in tests:
        out = latex_to_omath(t)
        print(f"\n[LaTeX] {t}\n[OMML ] {out[:200]}{'...' if len(out)>200 else ''}")

