"""Низкоуровневая работа с WordprocessingML: текст абзацев, эффективное форматирование, правки XML.

python-docx отдаёт только прямое форматирование, а Word вычисляет итоговое значение по цепочке:
прямое форматирование → нумерация → стиль (и его базовые стили) → docDefaults. Здесь эта цепочка
воспроизводится, чтобы проверки видели то же, что видит пользователь в Word.
"""

from __future__ import annotations

from copy import deepcopy
from typing import Callable, Iterable, Iterator

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "v": "urn:schemas-microsoft-com:vml",
}


def clark(prefixed: str) -> str:
    prefix, local = prefixed.split(":")
    return f"{{{NS[prefix]}}}{local}"


W_P = qn("w:p")
W_R = qn("w:r")
W_T = qn("w:t")
W_TBL = qn("w:tbl")
W_SDT = qn("w:sdt")
W_SYM = qn("w:sym")
W_TAB = qn("w:tab")
W_BR = qn("w:br")
W_CR = qn("w:cr")
W_NBHYPHEN = qn("w:noBreakHyphen")
W_FLDCHAR = qn("w:fldChar")
W_INSTR = qn("w:instrText")
W_FLDSIMPLE = qn("w:fldSimple")
W_DRAWING = qn("w:drawing")
W_PICT = qn("w:pict")
W_OBJECT = qn("w:object")
M_OMATH = clark("m:oMath")
M_OMATHPARA = clark("m:oMathPara")
M_T = clark("m:t")
WP_ANCHOR = clark("wp:anchor")
WP_INLINE = clark("wp:inline")

_EXCLUDED_ANCESTORS = {
    qn("w:txbxContent"),
    qn("w:del"),
    qn("w:moveFrom"),
    clark("mc:Fallback"),
}

# Порядок дочерних элементов по схеме OOXML — Word не открывает файлы с нарушенным порядком.
PPR_SEQ = (
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr", "w:widowControl",
    "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens",
    "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing",
    "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
    "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
)
LVL_SEQ = (
    "w:start", "w:numFmt", "w:lvlRestart", "w:pStyle", "w:isLgl", "w:suff", "w:lvlText",
    "w:lvlPicBulletId", "w:legacy", "w:lvlJc", "w:pPr", "w:rPr",
)

# Символы шрифта Symbol, которыми Word часто вставляет тире и маркеры (w:sym w:char="F02D").
_SYMBOL_CHARS = {"2D": "–", "BE": "—", "B7": "•", "A7": "▪", "D8": "►", "FC": "✓"}


def _is_excluded(el, stop) -> bool:
    parent = el.getparent()
    while parent is not None and parent is not stop:
        if parent.tag in _EXCLUDED_ANCESTORS:
            return True
        parent = parent.getparent()
    return False


def sym_char(sym) -> str:
    code = (sym.get(qn("w:char")) or "").upper()
    return _SYMBOL_CHARS.get(code[-2:], "?")


def paragraph_runs(p) -> list:
    """Все видимые w:r абзаца (включая гиперссылки, поля, вставки), кроме текстовых полей и удалённого."""
    return [r for r in p.iter(W_R) if not _is_excluded(r, p)]


def run_text(r) -> str:
    parts = []
    for ch in r:
        tag = ch.tag
        if tag == W_T:
            parts.append(ch.text or "")
        elif tag == W_SYM:
            parts.append(sym_char(ch))
        elif tag == W_TAB:
            parts.append("\t")
        elif tag in (W_BR, W_CR):
            parts.append("\n")
        elif tag == W_NBHYPHEN:
            parts.append("-")
    return "".join(parts)


def paragraph_text(p) -> str:
    return "".join(run_text(r) for r in paragraph_runs(p))


def math_text(el) -> str:
    return "".join(t.text or "" for t in el.iter(M_T))


def has_picture(p) -> bool:
    for tag in (W_DRAWING, W_PICT, W_OBJECT):
        for el in p.iter(tag):
            if not _is_excluded(el, p):
                return True
    return False


def has_math(p) -> bool:
    return next(p.iter(M_OMATH), None) is not None


def text_runs(p) -> list:
    """Прогоны абзаца, содержащие непробельный текст."""
    return [r for r in paragraph_runs(p) if run_text(r).strip()]


def iter_paragraphs(el) -> Iterator:
    for p in el.iter(W_P):
        if not _is_excluded(p, el):
            yield p


# --- Редактирование текста ------------------------------------------------------------


def _text_segments(p) -> list[list]:
    segs = []
    for r in paragraph_runs(p):
        for ch in r:
            if ch.tag == W_T:
                segs.append([ch, "t", ch.text or ""])
            elif ch.tag == W_SYM:
                segs.append([ch, "sym", sym_char(ch)])
            elif ch.tag == W_TAB:
                segs.append([ch, "tab", "\t"])
            elif ch.tag in (W_BR, W_CR):
                segs.append([ch, "br", "\n"])
            elif ch.tag == W_NBHYPHEN:
                segs.append([ch, "nbh", "-"])
    return segs


def _set_t(t, text: str) -> None:
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def replace_span(p, start: int, end: int, new: str) -> None:
    """Заменяет символы [start, end) текста абзаца (как его возвращает paragraph_text) на new.

    Форматирование прогонов сохраняется: новый текст попадает в первый затронутый прогон.
    """
    inserted = not new
    pos = 0
    for el, kind, txt in _text_segments(p):
        s, e = pos, pos + len(txt)
        pos = e
        touches = (s < end and e > start) or (start == end and s <= start <= e and kind == "t")
        if not touches:
            continue
        if kind == "t":
            before = txt[: start - s] if start > s else ""
            after = txt[end - s:] if end < e else ""
            _set_t(el, before + ("" if inserted else new) + after)
            inserted = True
        else:
            if not inserted:
                t = OxmlElement("w:t")
                _set_t(t, new)
                el.addprevious(t)
                inserted = True
            el.getparent().remove(el)
    if not inserted:
        runs = paragraph_runs(p)
        if runs:
            t = OxmlElement("w:t")
            _set_t(t, new)
            runs[-1].append(t)


def map_text_nodes(p, func: Callable[[str], str]) -> None:
    for t in p.iter(W_T):
        if not _is_excluded(t, p) and t.text:
            _set_t(t, func(t.text))


# --- Работа с pPr / rPr в правильном порядке ----------------------------------------------


def insert_in_order(parent, child, sequence: tuple[str, ...]) -> None:
    tag = child.tag
    names = [qn(n) for n in sequence]
    if tag not in names:
        parent.append(child)
        return
    successors = set(names[names.index(tag) + 1:])
    for existing in parent:
        if existing.tag in successors:
            existing.addprevious(child)
            return
    parent.append(child)


def get_or_add_ppr(p):
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p.insert(0, ppr)
    return ppr


def set_ppr_child(p, tag: str, attrs: dict[str, str]):
    ppr = get_or_add_ppr(p)
    el = ppr.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        insert_in_order(ppr, el, PPR_SEQ)
    for k, v in attrs.items():
        el.set(qn(k), v)
    return el


def set_lvl_child(lvl, tag: str, attrs: dict[str, str]):
    el = lvl.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        insert_in_order(lvl, el, LVL_SEQ)
    for k, v in attrs.items():
        el.set(qn(k), v)
    return el


def inline_from_anchor(anchor):
    """Строит wp:inline из wp:anchor (рисунок «В тексте» вместо обтекания)."""
    inline = OxmlElement("wp:inline")
    for attr in ("distT", "distB", "distL", "distR"):
        inline.set(attr, "0")
    for local in ("extent", "effectExtent", "docPr", "cNvGraphicFramePr"):
        child = anchor.find(clark(f"wp:{local}"))
        if child is not None:
            inline.append(deepcopy(child))
    graphic = anchor.find(clark("a:graphic"))
    if graphic is not None:
        inline.append(deepcopy(graphic))
    return inline


# --- Эффективные свойства ------------------------------------------------------------------

_TOGGLE_OFF = {"0", "false", "off", "none"}


def _toggle_value(el) -> bool:
    return (el.get(qn("w:val")) or "true").lower() not in _TOGGLE_OFF


def _int(v) -> int:
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return 0


class Resolver:
    """Вычисляет итоговое форматирование абзацев и прогонов с учётом стилей."""

    def __init__(self, document):
        self.document = document
        root = document.styles.element
        self.styles: dict[str, object] = {}
        self.default_pstyle = None
        self.style_names: dict[str, str] = {}
        for s in root.findall(qn("w:style")):
            sid = s.get(qn("w:styleId"))
            self.styles[sid] = s
            name_el = s.find(qn("w:name"))
            self.style_names[sid] = (name_el.get(qn("w:val")) if name_el is not None else sid) or sid
            if s.get(qn("w:type")) == "paragraph" and s.get(qn("w:default")) in ("1", "true"):
                self.default_pstyle = sid
        dd = root.find(qn("w:docDefaults"))
        self.def_rpr = dd.find(f"{qn('w:rPrDefault')}/{qn('w:rPr')}") if dd is not None else None
        self.def_ppr = dd.find(f"{qn('w:pPrDefault')}/{qn('w:pPr')}") if dd is not None else None
        self.theme_major, self.theme_minor = self._load_theme_fonts()
        self.numbering = self._load_part_element("/numbering")
        self._chain_cache: dict[str | None, list] = {}
        self._num_cache: dict[tuple[str, int], object] = {}

    # -- загрузка частей --
    def _related_part(self, reltype_suffix: str):
        for rel in self.document.part.rels.values():
            if rel.reltype.endswith(reltype_suffix) and not rel.is_external:
                return rel.target_part
        return None

    def _load_part_element(self, suffix: str):
        part = self._related_part(suffix)
        if part is None:
            return None
        element = getattr(part, "element", None)
        if element is not None:
            return element
        try:
            return etree.fromstring(part.blob)
        except Exception:
            return None

    def _load_theme_fonts(self) -> tuple[str, str]:
        part = self._related_part("/theme")
        major = minor = "Times New Roman"
        if part is None:
            return major, minor
        try:
            root = etree.fromstring(part.blob)
            mj = root.find(f".//{clark('a:majorFont')}/{clark('a:latin')}")
            mn = root.find(f".//{clark('a:minorFont')}/{clark('a:latin')}")
            if mj is not None and mj.get("typeface"):
                major = mj.get("typeface")
            if mn is not None and mn.get("typeface"):
                minor = mn.get("typeface")
        except Exception:
            pass
        return major, minor

    # -- стили --
    def style_chain(self, sid: str | None) -> list:
        if sid in self._chain_cache:
            return self._chain_cache[sid]
        chain, seen = [], set()
        cur = sid
        while cur and cur in self.styles and cur not in seen:
            seen.add(cur)
            st = self.styles[cur]
            chain.append(st)
            based = st.find(qn("w:basedOn"))
            cur = based.get(qn("w:val")) if based is not None else None
        self._chain_cache[sid] = chain
        return chain

    def pstyle_id(self, p) -> str | None:
        ppr = p.find(qn("w:pPr"))
        if ppr is not None:
            ps = ppr.find(qn("w:pStyle"))
            if ps is not None and ps.get(qn("w:val")) in self.styles:
                return ps.get(qn("w:val"))
        return self.default_pstyle

    def pstyle_name(self, p) -> str:
        sid = self.pstyle_id(p)
        return self.style_names.get(sid, "") if sid else ""

    # -- нумерация --
    def num_ref(self, p) -> tuple[str, int] | None:
        """(numId, ilvl) абзаца с учётом стиля; None, если абзац не в списке."""
        num_id = ilvl = None
        ppr = p.find(qn("w:pPr"))
        sources = [ppr] if ppr is not None else []
        sources += [st.find(qn("w:pPr")) for st in self.style_chain(self.pstyle_id(p))]
        for src in sources:
            if src is None:
                continue
            numpr = src.find(qn("w:numPr"))
            if numpr is None:
                continue
            if num_id is None and numpr.find(qn("w:numId")) is not None:
                num_id = numpr.find(qn("w:numId")).get(qn("w:val"))
            if ilvl is None and numpr.find(qn("w:ilvl")) is not None:
                ilvl = _int(numpr.find(qn("w:ilvl")).get(qn("w:val")))
            if num_id is not None and ilvl is not None:
                break
        if num_id is None or num_id == "0" or self.numbering is None:
            return None
        return num_id, ilvl or 0

    def num_level(self, num_id: str, ilvl: int, _depth: int = 0):
        """Возвращает (lvl, abstractNum) для нумерации или (None, None)."""
        key = (num_id, ilvl)
        if key in self._num_cache:
            return self._num_cache[key]
        result = (None, None)
        numbering = self.numbering
        num = None
        for n in numbering.findall(qn("w:num")):
            if n.get(qn("w:numId")) == num_id:
                num = n
                break
        if num is not None:
            override_lvl = None
            for ov in num.findall(qn("w:lvlOverride")):
                if _int(ov.get(qn("w:ilvl"))) == ilvl and ov.find(qn("w:lvl")) is not None:
                    override_lvl = ov.find(qn("w:lvl"))
            abs_id_el = num.find(qn("w:abstractNumId"))
            abs_id = abs_id_el.get(qn("w:val")) if abs_id_el is not None else None
            abstract = None
            for a in numbering.findall(qn("w:abstractNum")):
                if a.get(qn("w:abstractNumId")) == abs_id:
                    abstract = a
                    break
            if abstract is not None:
                link = abstract.find(qn("w:numStyleLink"))
                if link is not None and _depth < 2:
                    st = self.styles.get(link.get(qn("w:val")))
                    numpr = st.find(f"{qn('w:pPr')}/{qn('w:numPr')}") if st is not None else None
                    nid = numpr.find(qn("w:numId")) if numpr is not None else None
                    if nid is not None:
                        result = self.num_level(nid.get(qn("w:val")), ilvl, _depth + 1)
                        self._num_cache[key] = result
                        return result
                lvl = override_lvl
                if lvl is None:
                    for l in abstract.findall(qn("w:lvl")):
                        if _int(l.get(qn("w:ilvl"))) == ilvl:
                            lvl = l
                            break
                result = (lvl, abstract)
        self._num_cache[key] = result
        return result

    # -- свойства абзаца --
    def _ppr_chain(self, p, with_numbering: bool = True) -> list:
        chain = []
        ppr = p.find(qn("w:pPr"))
        if ppr is not None:
            chain.append(ppr)
        if with_numbering:
            ref = self.num_ref(p)
            if ref:
                lvl, _ = self.num_level(*ref)
                if lvl is not None and lvl.find(qn("w:pPr")) is not None:
                    chain.append(lvl.find(qn("w:pPr")))
        for st in self.style_chain(self.pstyle_id(p)):
            sp = st.find(qn("w:pPr"))
            if sp is not None:
                chain.append(sp)
        if self.def_ppr is not None:
            chain.append(self.def_ppr)
        return chain

    def alignment(self, p) -> str:
        for ppr in self._ppr_chain(p, with_numbering=False):
            jc = ppr.find(qn("w:jc"))
            if jc is not None and jc.get(qn("w:val")):
                v = jc.get(qn("w:val"))
                if v in ("both", "distribute", "lowKashida", "mediumKashida", "highKashida", "thaiDistribute"):
                    return "both"
                return {"start": "left", "end": "right"}.get(v, v)
        return "left"

    def indent(self, p) -> tuple[int, int, int]:
        """(left, first_line, right) в твипах; висячий отступ — отрицательный first_line."""
        left = first = right = None
        for ppr in self._ppr_chain(p):
            ind = ppr.find(qn("w:ind"))
            if ind is None:
                continue
            if left is None:
                v = ind.get(qn("w:left"))
                v = v if v is not None else ind.get(qn("w:start"))
                if v is not None:
                    left = _int(v)
            if right is None:
                v = ind.get(qn("w:right"))
                v = v if v is not None else ind.get(qn("w:end"))
                if v is not None:
                    right = _int(v)
            if first is None:
                if ind.get(qn("w:hanging")) is not None:
                    first = -_int(ind.get(qn("w:hanging")))
                elif ind.get(qn("w:firstLine")) is not None:
                    first = _int(ind.get(qn("w:firstLine")))
        return left or 0, first or 0, right or 0

    def spacing(self, p) -> tuple[int, int, float, str]:
        """(before, after) в твипах, межстрочный множитель (или пункты для exact/atLeast), правило."""
        before = after = line = rule = None
        for ppr in self._ppr_chain(p, with_numbering=False):
            sp = ppr.find(qn("w:spacing"))
            if sp is None:
                continue
            if before is None:
                if (sp.get(qn("w:beforeAutospacing")) or "").lower() in ("1", "true", "on"):
                    before = 280
                elif sp.get(qn("w:before")) is not None:
                    before = _int(sp.get(qn("w:before")))
            if after is None:
                if (sp.get(qn("w:afterAutospacing")) or "").lower() in ("1", "true", "on"):
                    after = 280
                elif sp.get(qn("w:after")) is not None:
                    after = _int(sp.get(qn("w:after")))
            if line is None and sp.get(qn("w:line")) is not None:
                line = _int(sp.get(qn("w:line")))
                rule = sp.get(qn("w:lineRule")) or "auto"
        line = 240 if line is None else line
        rule = rule or "auto"
        value = line / 240 if rule == "auto" else line / 20
        return before or 0, after or 0, value, rule

    def outline_level(self, p) -> int | None:
        for ppr in self._ppr_chain(p, with_numbering=False):
            ol = ppr.find(qn("w:outlineLvl"))
            if ol is not None:
                v = _int(ol.get(qn("w:val")))
                return v if v < 9 else None
        return None

    # -- свойства прогона --
    def _rpr_chain(self, r, p) -> list:
        chain = []
        rpr = r.find(qn("w:rPr"))
        if rpr is not None:
            chain.append(rpr)
            rs = rpr.find(qn("w:rStyle"))
            if rs is not None:
                for st in self.style_chain(rs.get(qn("w:val"))):
                    if st.find(qn("w:rPr")) is not None:
                        chain.append(st.find(qn("w:rPr")))
        for st in self.style_chain(self.pstyle_id(p)):
            if st.find(qn("w:rPr")) is not None:
                chain.append(st.find(qn("w:rPr")))
        if self.def_rpr is not None:
            chain.append(self.def_rpr)
        return chain

    def _theme_font(self, theme: str) -> str:
        return self.theme_major if theme.startswith("major") else self.theme_minor

    def font_name(self, r, p) -> str:
        text = run_text(r)
        which = "hAnsi" if any(ord(c) > 127 for c in text) else "ascii"
        for rpr in self._rpr_chain(r, p):
            rf = rpr.find(qn("w:rFonts"))
            if rf is None:
                continue
            theme = rf.get(qn(f"w:{which}Theme"))
            if theme:
                return self._theme_font(theme)
            name = rf.get(qn(f"w:{which}"))
            if name:
                return name
        return "Times New Roman"

    def font_size(self, r, p) -> float:
        for rpr in self._rpr_chain(r, p):
            sz = rpr.find(qn("w:sz"))
            if sz is not None and sz.get(qn("w:val")):
                return _int(sz.get(qn("w:val"))) / 2
        return 10.0

    def toggle(self, r, p, tag: str) -> bool:
        for rpr in self._rpr_chain(r, p):
            el = rpr.find(qn(tag))
            if el is not None:
                return _toggle_value(el)
        return False

    def underline(self, r, p) -> bool:
        for rpr in self._rpr_chain(r, p):
            u = rpr.find(qn("w:u"))
            if u is not None:
                return (u.get(qn("w:val")) or "single") != "none"
        return False

    def color(self, r, p) -> str | None:
        """None — цвет чёрный/авто, иначе строка с описанием цвета."""
        for rpr in self._rpr_chain(r, p):
            c = rpr.find(qn("w:color"))
            if c is None:
                continue
            theme = c.get(qn("w:themeColor"))
            if theme:
                return None if theme in ("text1", "dark1") else f"цвет темы «{theme}»"
            val = (c.get(qn("w:val")) or "auto").upper()
            if val in ("AUTO", "000000"):
                return None
            return f"#{val}"
        return None


# --- Утилиты форматирования чисел ------------------------------------------------------------


def cm(twips: float) -> str:
    return f"{twips / 566.929:.2f}".replace(".", ",") + " см"


def num_ru(value: float) -> str:
    return f"{value:g}".replace(".", ",")


def shorten(text: str, limit: int = 45) -> str:
    text = " ".join(text.split())
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


def uniq(items: Iterable[str]) -> list[str]:
    seen, out = set(), []
    for it in items:
        if it not in seen:
            seen.add(it)
            out.append(it)
    return out
