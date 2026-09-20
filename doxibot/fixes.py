"""Примитивы автоисправления: прямое форматирование прогонов и абзацев, стили, колонтитулы."""

from __future__ import annotations

from typing import Iterable

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Twips
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from .ooxml import get_or_add_ppr, insert_in_order, LVL_SEQ, PPR_SEQ
from .profile import DEFAULT, Profile

_ALIGN = {
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}
_THEME_FONT_ATTRS = ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme", "csTheme")
_CHAR_IND_ATTRS = (
    "start", "end", "firstLineChars", "hangingChars", "leftChars", "rightChars", "startChars", "endChars",
)


# --- прогоны ---------------------------------------------------------------------------------


def _fonts_element(r):
    return r.get_or_add_rPr().get_or_add_rFonts()


def set_font_name(runs: Iterable, name: str = DEFAULT.font_name) -> None:
    for r in runs:
        rf = _fonts_element(r)
        for attr in ("ascii", "hAnsi", "cs"):
            rf.set(qn(f"w:{attr}"), name)
        for attr in _THEME_FONT_ATTRS:
            rf.attrib.pop(qn(f"w:{attr}"), None)
        rf.attrib.pop(qn("w:hint"), None)


def set_font_size(runs: Iterable, size_pt: float = DEFAULT.font_size) -> None:
    for r in runs:
        Run(r, None).font.size = Pt(size_pt)
        szcs = r.rPr.find(qn("w:szCs"))
        if szcs is not None:
            szcs.set(qn("w:val"), str(int(size_pt * 2)))


def set_black(runs: Iterable) -> None:
    for r in runs:
        Run(r, None).font.color.rgb = RGBColor(0, 0, 0)
        color = r.rPr.find(qn("w:color"))
        if color is not None:
            for attr in ("themeColor", "themeTint", "themeShade"):
                color.attrib.pop(qn(f"w:{attr}"), None)


def remove_underline(runs: Iterable) -> None:
    for r in runs:
        Run(r, None).font.underline = False


def set_bold(runs: Iterable, on: bool) -> None:
    for r in runs:
        font = Run(r, None).font
        font.bold = on
        if r.rPr.find(qn("w:bCs")) is not None:
            font.complex_script_bold = on


def caps_off(runs: Iterable) -> None:
    for r in runs:
        rpr = r.get_or_add_rPr()
        for tag in ("w:caps", "w:smallCaps"):
            el = rpr.find(qn(tag))
            if el is not None:
                rpr.remove(el)
        Run(r, None).font.all_caps = False


def normalize_runs(runs: Iterable, size_pt: float | None = DEFAULT.font_size,
                   name: str = DEFAULT.font_name) -> None:
    runs = list(runs)
    set_font_name(runs, name)
    if size_pt is not None:
        set_font_size(runs, size_pt)
    set_black(runs)
    remove_underline(runs)


def normalize_paragraph_mark(p, profile: Profile = DEFAULT, bold: bool | None = None) -> None:
    """Приводит формат знака абзаца (pPr/rPr) — от него Word берёт вид номера/маркера перечисления."""
    ppr = get_or_add_ppr(p)
    rpr = ppr.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        insert_in_order(ppr, rpr, PPR_SEQ)
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("ascii", "hAnsi", "cs"):
        rfonts.set(qn(f"w:{attr}"), profile.font_name)
    for attr in _THEME_FONT_ATTRS:
        rfonts.attrib.pop(qn(f"w:{attr}"), None)
    rfonts.attrib.pop(qn("w:hint"), None)
    _set_mark_child(rpr, "w:sz", str(profile.font_size_half))
    _set_mark_child(rpr, "w:szCs", str(profile.font_size_half))
    _set_mark_child(rpr, "w:color", "000000", attr="w:val")
    color = rpr.find(qn("w:color"))
    for attr in ("themeColor", "themeTint", "themeShade"):
        color.attrib.pop(qn(f"w:{attr}"), None)
    _set_mark_child(rpr, "w:u", "none")
    if bold is not None:
        _set_mark_child(rpr, "w:b", "1" if bold else "0")
        _set_mark_child(rpr, "w:bCs", "1" if bold else "0")


_RPR_SEQ = (
    "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps", "w:strike", "w:dstrike",
    "w:outline", "w:shadow", "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish", "w:webHidden",
    "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect",
    "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout",
    "w:specVanish", "w:oMath",
)


def _set_mark_child(rpr, tag: str, value: str, attr: str = "w:val") -> None:
    el = rpr.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        insert_in_order(rpr, el, _RPR_SEQ)
    el.set(qn(attr), value)


# --- абзацы ----------------------------------------------------------------------------------


def _pf(p):
    return Paragraph(p, None).paragraph_format


def set_alignment(p, key: str) -> None:
    _pf(p).alignment = _ALIGN[key]


def set_indent(p, first: int = 0, left: int = 0, right: int = 0) -> None:
    pf = _pf(p)
    pf.left_indent = Twips(left)
    pf.right_indent = Twips(right)
    pf.first_line_indent = Twips(first)
    ind = p.pPr.find(qn("w:ind"))
    if ind is not None:
        for attr in _CHAR_IND_ATTRS:
            ind.attrib.pop(qn(f"w:{attr}"), None)


def set_spacing(p, before_pt: float = 0.0, after_pt: float = 0.0) -> None:
    pf = _pf(p)
    pf.space_before = Pt(before_pt)
    pf.space_after = Pt(after_pt)
    sp = p.pPr.find(qn("w:spacing"))
    if sp is not None:
        for attr in ("beforeAutospacing", "afterAutospacing", "beforeLines", "afterLines"):
            sp.attrib.pop(qn(f"w:{attr}"), None)


def set_spacing_zero(p) -> None:
    set_spacing(p, 0, 0)


def normalize_numbering_level(lvl, profile: Profile = DEFAULT, set_font: bool = True) -> None:
    """Приводит вид номера/маркера в описании нумерации к параметрам профиля."""
    rpr = lvl.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        insert_in_order(lvl, rpr, LVL_SEQ)
    if set_font:
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        for attr in ("ascii", "hAnsi", "cs"):
            rfonts.set(qn(f"w:{attr}"), profile.font_name)
        for attr in _THEME_FONT_ATTRS:
            rfonts.attrib.pop(qn(f"w:{attr}"), None)
        rfonts.attrib.pop(qn("w:hint"), None)
    _set_mark_child(rpr, "w:sz", str(profile.font_size_half))
    _set_mark_child(rpr, "w:szCs", str(profile.font_size_half))
    _set_mark_child(rpr, "w:color", "000000")
    color = rpr.find(qn("w:color"))
    for attr in ("themeColor", "themeTint", "themeShade"):
        color.attrib.pop(qn(f"w:{attr}"), None)
    for tag in ("w:b", "w:bCs", "w:u"):
        el = rpr.find(qn(tag))
        if el is not None:
            el.set(qn("w:val"), "none" if tag == "w:u" else "0")


def set_line_spacing(p, value: float) -> None:
    _pf(p).line_spacing = value


def set_style(p, style_id: str) -> None:
    ppr = get_or_add_ppr(p)
    ps = ppr.find(qn("w:pStyle"))
    if ps is None:
        ps = OxmlElement("w:pStyle")
        insert_in_order(ppr, ps, PPR_SEQ)
    ps.set(qn("w:val"), style_id)


# --- стили -----------------------------------------------------------------------------------


def _style_ppr_child(style, tag: str, attrs: dict[str, str]) -> None:
    ppr = style.element.get_or_add_pPr()
    el = ppr.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        insert_in_order(ppr, el, PPR_SEQ)
    for k, v in attrs.items():
        el.set(qn(k), v)


def _normalize_style_font(style, profile: Profile = DEFAULT, bold: bool | None = None) -> None:
    font = style.font
    font.name = profile.font_name
    rpr = style.element.get_or_add_rPr()
    rf = rpr.get_or_add_rFonts()
    rf.set(qn("w:cs"), profile.font_name)
    for attr in _THEME_FONT_ATTRS:
        rf.attrib.pop(qn(f"w:{attr}"), None)
    font.size = Pt(profile.font_size)
    font.color.rgb = RGBColor(0, 0, 0)
    color = rpr.find(qn("w:color"))
    if color is not None:
        for attr in ("themeColor", "themeTint", "themeShade"):
            color.attrib.pop(qn(f"w:{attr}"), None)
    if bold is not None:
        font.bold = bold


def ensure_heading_style(document, level: int, profile: Profile = DEFAULT) -> str:
    """Возвращает styleId стиля «Заголовок N», при необходимости создавая его."""
    name = f"Heading {level}"
    try:
        style = document.styles[name]
    except KeyError:
        style = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        try:
            style.base_style = document.styles["Normal"]
        except KeyError:
            pass
        style.quick_style = True
        _style_ppr_child(style, "w:keepNext", {})
        _style_ppr_child(style, "w:outlineLvl", {"w:val": str(level - 1)})
        _normalize_style_font(style, profile, bold=True)
    return style.style_id


def normalize_styles(document, profile: Profile = DEFAULT) -> None:
    """Заголовки и оглавление: Times New Roman 14, чёрный — чтобы автосодержание собиралось правильно."""
    for level in (1, 2, 3):
        try:
            _normalize_style_font(document.styles[f"Heading {level}"], profile, bold=True)
        except KeyError:
            pass
    for level in (1, 2, 3):
        try:
            style = document.styles[f"TOC {level}"]
        except KeyError:
            continue
        _normalize_style_font(style, profile)
        style.paragraph_format.first_line_indent = Twips(0)


def request_fields_update(document) -> None:
    """Просит Word обновить поля (содержание) при открытии файла."""
    settings = document.settings.element
    el = settings.find(qn("w:updateFields"))
    if el is None:
        el = OxmlElement("w:updateFields")
        settings.append(el)
    el.set(qn("w:val"), "true")


# --- колонтитулы ------------------------------------------------------------------------------


def _field_run(kind: str, payload: str | None = None, profile: Profile = DEFAULT):
    r = OxmlElement("w:r")
    if kind in ("begin", "separate", "end"):
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), kind)
        r.append(fc)
    elif kind == "instr":
        it = OxmlElement("w:instrText")
        it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        it.text = payload
        r.append(it)
    else:
        t = OxmlElement("w:t")
        t.text = payload
        r.append(t)
    normalize_runs([r], profile.font_size, profile.font_name)
    return r


def add_page_numbers(document, profile: Profile = DEFAULT) -> None:
    section = document.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    paragraph = footer.paragraphs[0]
    p = paragraph._p
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)
    set_alignment(p, "center")
    set_indent(p, 0, 0, 0)
    set_spacing_zero(p)
    for kind, payload in (("begin", None), ("instr", " PAGE "), ("separate", None), ("text", "3"), ("end", None)):
        p.append(_field_run(kind, payload, profile))
    section.different_first_page_header_footer = True
