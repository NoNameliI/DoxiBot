"""Анализ DOCX на соответствие стандартным требованиям к оформлению.

Документ разбирается на блоки (абзацы и таблицы), каждый блок классифицируется (заголовок, рисунок,
подпись, формула, перечисление, основной текст…), после чего к нему применяются свои правила.
Каждое найденное замечание может нести функцию автоисправления.
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Callable

from docx.document import Document as DocxDocument
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm

from . import fixes as F
from . import rules as R
from .model import Analysis, Issue
from .profile import DEFAULT, Profile
from .ooxml import (
    _SYMBOL_CHARS, M_OMATHPARA, W_FLDCHAR, W_FLDSIMPLE, W_INSTR, W_P, W_SDT, W_T, W_TBL, WP_ANCHOR, Resolver, _is_excluded,
    clark, cm, has_math, has_picture, inline_from_anchor, iter_paragraphs, map_text_nodes, math_text,
    num_ru, paragraph_runs, paragraph_text, replace_span, set_lvl_child, shorten, text_runs,
)

TOL = R.INDENT_TOLERANCE_TWIPS
LAST_RENDERED = qn("w:lastRenderedPageBreak")
W_BR = qn("w:br")

ALIGN_RU = {"both": "по ширине", "center": "по центру", "left": "по левому краю", "right": "по правому краю"}
CAPTION_WORD = {"fig": "Рисунок", "tbl": "Таблица", "lst": "Листинг"}

CAPTION_RE = re.compile(
    r"^(?P<lead>\s*)(?P<word>[А-Яа-яЁё]+\.?)[ \t\xa0]*(?:№[ \t\xa0]*)?"
    r"(?P<num>(?:\d+(?:\.\d+)*|[А-ЯЁA-Z](?:\.\d+)+)\.?)"
    r"(?P<sep>[ \t\xa0]*(?P<dash>[–—−‒―\-:])?[ \t\xa0]*)"
    r"(?P<name>.*?)(?P<tail>[\s.]*)$",
    re.S,
)
VALID_NUM = re.compile(r"^(?:\d+|[А-ЯЁA-Z])\.\d+$")
CONT_RE = re.compile(r"^\s*продолжение\s+таблицы\b", re.I)
CONT_FULL = re.compile(r"^Продолжение таблицы ((?:\d+|[А-ЯЁA-Z])(?:\.\d+)*)$")
FORMULA_NUM_RE = re.compile(r"\(\s*((?:\d+|[А-ЯЁA-Z])(?:\.\d+)*)\s*\)")
NUM_HEADING_RE = re.compile(r"^(?P<num>\d{1,2}(?:\.\d{1,2}){0,3})(?P<dot>\.?)\s+(?P<title>\S.*)$", re.S)
GLAVA_RE = re.compile(r"^глава\s+(?P<num>\d+)\.?\s*(?P<title>.*)$", re.I | re.S)
APPENDIX_RE = re.compile(r"^ПРИЛОЖЕНИЕ(?:\s+(?P<letter>[А-ЯЁA-Z])(?![А-ЯЁA-Z]))?[.\s]*(?P<title>.*)$", re.S)
TOC_LINE_RE = re.compile(r"(\.{3,}|…|\t)\s*\d+\s*$")
REF_NUMS = r"(?:\d+|[А-ЯЁA-Z])(?:\.\d+)+"
FIG_REF_RE = re.compile(
    rf"(?<![А-Яа-яЁё])(?:рисун[а-яё]*|рис\.)\s*(?:№\s*)?(?P<nums>{REF_NUMS}(?:\s*(?:,|и|или|–|—|-)\s*{REF_NUMS})*)",
    re.I,
)
TBL_REF_RE = re.compile(
    rf"(?<![А-Яа-яЁё])(?:таблиц[а-яё]*|табл\.)\s*(?:№\s*)?(?P<nums>{REF_NUMS}(?:\s*(?:,|и|или|–|—|-)\s*{REF_NUMS})*)",
    re.I,
)
NUM_TOKEN_RE = re.compile(REF_NUMS, re.I)
FIG_ABBR_RE = re.compile(r"(?<![А-Яа-яЁё])рис\.\s*(?:№\s*)?\d", re.I)
URL_RE = re.compile(r"(https?://|www\.|\bURL\b)", re.I)


def parse_caption(text: str) -> dict | None:
    m = CAPTION_RE.match(text)
    if not m:
        return None
    word = m.group("word").lower().rstrip(".")
    kind = {"рисунок": "fig", "рис": "fig", "таблица": "tbl", "табл": "tbl", "листинг": "lst"}.get(word)
    if kind is None:
        return None
    d = m.groupdict()
    d["kind"] = kind
    d["match"] = m
    return d


@dataclass
class Block:
    idx: int
    kind: str                       # 'p' | 'tbl'
    el: object
    text: str = ""
    ptype: str = ""
    in_toc: bool = False
    heading: tuple | None = None    # (вид, уровень, оформлен_стилем, доп. данные)
    caption: dict | None = None
    section: str = ""
    chapter: str | None = None
    in_sources: bool = False
    in_appendix: bool = False
    special: str = ""
    expected: str | None = None
    extra: dict = field(default_factory=dict)


class Analyzer:
    def __init__(self, document: DocxDocument, profile: Profile = DEFAULT, title_pages: int | None = None):
        self.doc = document
        self.p = profile
        self.indent = profile.indent_twips
        self.title_pages = profile.title_pages if title_pages is None else title_pages
        self.res = Resolver(document)
        self.blocks: list[Block] = []
        self.analysis = Analysis(profile=profile)
        self.has_toc_field = False
        self._fields: list[dict] = []
        self._section_no = 0
        self._page = 1
        self._seen_text = True
        self.renumber: dict[str, dict[str, str]] = {"fig": {}, "tbl": {}}
        self.renumber_conflicts: dict[str, set[str]] = {"fig": set(), "tbl": set()}
        self.table_profiles: list[tuple[Block, tuple[float, float]]] = []
        self.table_target: tuple[float, float] = (14.0, 1.0)

    # ================================================================== public
    def run(self) -> Analysis:
        self._collect()
        self._classify()
        self._mark_front()
        self._resolve_captions()
        self._context_pass()
        for b in self.blocks:
            self._check_block(b)
        self._check_tables_uniform()
        self._check_lists()
        self._check_references()
        self._check_toc()
        self._check_page_setup()
        self._check_page_numbers()
        self._register_post_fixes()
        self._stats()
        return self.analysis

    # ================================================================== issues
    def add(self, code: str, b: Block | None, detail: str = "", fix: Callable | None = None,
            where: str | None = None, anchor=None) -> None:
        group = R.RULES[code].group
        if group and not getattr(self.p, group, True):
            return
        if anchor is None and b is not None:
            anchor = b.el if b.kind == "p" else next(iter_paragraphs(b.el), None)
        self.analysis.issues.append(
            Issue(code=code, where=where or (self._where(b) if b else ""), detail=detail, anchor=anchor, fix=fix)
        )

    def _where(self, b: Block) -> str:
        sec = shorten(b.section, 30)
        if b.kind == "tbl":
            prev = self._prev(b)
            if prev is not None and prev.ptype in ("CAP_TBL", "CAP_CONT"):
                return f"таблица «{shorten(prev.text, 45)}»"
            return f"таблица в разделе «{sec}»" if sec else "таблица"
        if b.ptype == "HEADING":
            return f"заголовок «{shorten(b.text, 50)}»"
        if b.ptype == "PICTURE":
            nxt = self._next(b)
            if nxt is not None and nxt.ptype == "CAP_FIG":
                return f"рисунок над «{shorten(nxt.text, 40)}»"
            return f"рисунок в разделе «{sec}»" if sec else "рисунок"
        if b.ptype == "FORMULA":
            label = "формула"
            if b.caption and b.caption.get("num"):
                label += f" ({b.caption['num']})"
            return f"{label} в разделе «{sec}»" if sec else label
        snippet = shorten(b.text.strip(), 50) or "пустой абзац"
        return f"«{sec}» → «{snippet}»" if sec else f"«{snippet}»"

    # ================================================================== collect
    def _collect(self) -> None:
        for child in self.doc.element.body:
            self._walk(child, False)

    def _walk(self, el, in_toc: bool) -> None:
        tag = el.tag
        if tag == W_P:
            toc = self._scan_fields(el) or in_toc
            self.blocks.append(Block(len(self.blocks), "p", el, in_toc=toc,
                                     extra={"sect": self._section_no, "page": self._page}))
            breaks = self._page_breaks(el)
            ppr = el.find(qn("w:pPr"))
            sect = ppr.find(qn("w:sectPr")) if ppr is not None else None
            if sect is not None:
                self._section_no += 1
                kind = sect.find(qn("w:type"))
                if (kind is None or kind.get(qn("w:val")) in ("nextPage", "oddPage", "evenPage")) and not breaks:
                    breaks = 1
            self._page += breaks
        elif tag == W_TBL:
            self._scan_fields(el)
            self.blocks.append(Block(len(self.blocks), "tbl", el,
                                     extra={"sect": self._section_no, "page": self._page}))
            self._page += self._page_breaks(el)
        elif tag in (W_SDT, qn("w:customXml")):
            content = el
            is_toc = False
            if tag == W_SDT:
                gal = el.find(f"{qn('w:sdtPr')}/{qn('w:docPartObj')}/{qn('w:docPartGallery')}")
                is_toc = gal is not None and "contents" in (gal.get(qn("w:val")) or "").lower()
                content = el.find(qn("w:sdtContent"))
            if content is not None:
                for child in content:
                    self._walk(child, in_toc or is_toc)

    def _page_breaks(self, el) -> int:
        """Сколько страниц начинается внутри блока: явные разрывы и разметка Word (lastRenderedPageBreak).

        Два маркера подряд без текста между ними — один и тот же разрыв.
        """
        count = 0
        for node in el.iter(W_T, W_BR, LAST_RENDERED):
            if node.tag == W_T:
                if (node.text or "").strip():
                    self._seen_text = True
            elif node.tag == W_BR:
                if node.get(qn("w:type")) == "page" and self._seen_text:
                    count += 1
                    self._seen_text = False
            elif self._seen_text:
                count += 1
                self._seen_text = False
        return count

    def _scan_fields(self, el) -> bool:
        stack = self._fields
        in_toc = any(f["toc"] for f in stack)
        for node in el.iter(W_FLDCHAR, W_INSTR, W_FLDSIMPLE):
            if node.tag == W_FLDSIMPLE:
                if re.match(r"\s*TOC\b", node.get(qn("w:instr")) or ""):
                    in_toc = self.has_toc_field = True
            elif node.tag == W_INSTR:
                if stack:
                    stack[-1]["instr"] += node.text or ""
                    if re.match(r"\s*TOC\b", stack[-1]["instr"]):
                        stack[-1]["toc"] = True
                        in_toc = self.has_toc_field = True
            else:
                kind = node.get(qn("w:fldCharType"))
                if kind == "begin":
                    stack.append({"instr": "", "toc": False})
                elif kind == "end" and stack:
                    stack.pop()
        return in_toc

    # ================================================================== classify
    def _classify(self) -> None:
        for b in self.blocks:
            if b.kind == "tbl":
                b.ptype = "TABLE"
                continue
            p = b.el
            b.text = paragraph_text(p)
            s = b.text.strip()
            style = self.res.pstyle_name(p).lower()
            heading = self._heading(p, b.text)
            if b.in_toc or style.startswith("toc ") or style.startswith("оглавление"):
                if heading and heading[0] == "toc":
                    b.ptype, b.heading = "HEADING", heading
                else:
                    b.ptype = "TOC" if s else "EMPTY"
            elif heading:
                b.ptype, b.heading = "HEADING", heading
            elif has_picture(p):
                cap = parse_caption(s) if s else None
                if len(s) <= 3 or (cap and cap["kind"] == "fig"):
                    b.ptype = "PICTURE"
                    b.caption = cap
                else:
                    b.ptype = "BODY"
                    b.extra["picture_in_text"] = True
            elif has_math(p) and self._formula_like(p, b.text):
                b.ptype = "FORMULA"
            elif not s:
                b.ptype = "EMPTY"
            elif FORMULA_NUM_RE.fullmatch(s):
                b.ptype = "FORMULA_NUM"
            elif (cap := parse_caption(s)) is not None:
                b.ptype, b.caption = "PRECAP", cap
            elif CONT_RE.match(s) and len(s) <= 40:
                b.ptype = "CAP_CONT"
            elif self.res.num_ref(p) is not None:
                b.ptype = "LIST"
            else:
                b.ptype = "BODY"

    def _heading(self, p, text: str):
        t = " ".join(text.split())
        if not t or len(t) > 200:
            return None
        up = t.upper().rstrip(".").strip()
        outline = self.res.outline_level(p)
        styled = outline is not None
        if up in R.TOC_HEADINGS:
            return ("toc", 1, styled, {})
        if up in R.STRUCTURAL_HEADINGS:
            return ("structural", 1, styled, {})
        m = APPENDIX_RE.match(up)
        if m and (styled or len(t) <= 60) and (m.group("letter") or up == "ПРИЛОЖЕНИЕ"):
            return ("appendix", 1, styled, {"letter": m.group("letter"), "title": m.group("title").strip()})
        if TOC_LINE_RE.search(text):
            return None
        g = GLAVA_RE.match(t)
        n = NUM_HEADING_RE.match(t)
        number = g.group("num") if g else (n.group("num") if n else None)
        if styled:
            level = min(outline + 1, 3)
            if g:
                level = 1
            elif n:  # номер в тексте надёжнее стиля: «2.1 …» — подглава, даже со стилем «Заголовок 1»
                level = min(n.group("num").count(".") + 1, 3)
            return ("chapter" if level == 1 else "sub", level, True,
                    {"num": number and number.split(".")[0], "outline": outline, "numbered": bool(g or n)})
        # Заголовок, набранный без стиля: номер + полужирный текст без точки в конце
        if (g or n) and len(t) <= 150 and not t.endswith((";", ",", ":")) and self.res.num_ref(p) is None:
            runs = text_runs(p)
            if runs and all(self.res.toggle(r, p, "w:b") for r in runs):
                level = 1 if g else min(n.group("num").count(".") + 1, 3)
                return ("chapter" if level == 1 else "sub", level, False, {"num": number.split(".")[0]})
        return None

    def _formula_like(self, p, text: str) -> bool:
        rest = FORMULA_NUM_RE.sub("", text)
        rest = re.sub(r"[\s,.;:#]", "", rest)
        return len(rest) <= 2

    def _mark_front(self) -> None:
        if self.title_pages is not None and self.title_pages >= 0:
            if self.title_pages == 0:
                return
            pages = max((b.extra.get("page", 1) for b in self.blocks), default=1)
            if pages > self.title_pages:
                for b in self.blocks:
                    if b.extra.get("page", 1) <= self.title_pages:
                        b.extra["orig"] = b.ptype
                        b.ptype = "FRONT"
                return
            self.analysis.notes.append(
                "Границы страниц в файле не размечены — титульный лист определён автоматически, "
                "по первому разделу отчёта."
            )
        headings = [b for b in self.blocks if b.ptype == "HEADING"]
        # Бланк задания содержит «1. ЦЕЛЕВАЯ УСТАНОВКА» и т.п. — неоформленные «заголовки» не считаем началом отчёта.
        known = [b for b in headings if b.heading[0] in ("toc", "structural", "appendix") or b.heading[2]]
        first = (known or headings)[0].idx if headings else None
        if first is None:
            self.analysis.notes.append(
                "Не найдено ни одного заголовка — проверены только общие параметры текста."
            )
            return
        front = self.blocks[:first]
        if sum(1 for b in front if b.kind == "p" and b.text.strip()) > 80:
            return  # слишком много текста до первого заголовка — это не титульный лист
        for b in front:
            b.extra["orig"] = b.ptype
            b.ptype = "FRONT"

    def _prev(self, b: Block, skip=("EMPTY",)) -> Block | None:
        for i in range(b.idx - 1, -1, -1):
            if self.blocks[i].ptype not in skip:
                return self.blocks[i]
        return None

    def _next(self, b: Block, skip=("EMPTY",)) -> Block | None:
        for i in range(b.idx + 1, len(self.blocks)):
            if self.blocks[i].ptype not in skip:
                return self.blocks[i]
        return None

    def _is_picture_block(self, b: Block | None) -> bool:
        if b is None:
            return False
        if b.ptype == "PICTURE":
            return True
        return b.kind == "tbl" and any(has_picture(p) for p in iter_paragraphs(b.el))

    def _is_monospace(self, p) -> bool:
        runs = text_runs(p)
        return bool(runs) and all(self.res.font_name(r, p).lower() in R.MONOSPACE_FONTS for r in runs)

    def _is_code_table(self, tbl) -> bool:
        paras = [p for p in iter_paragraphs(tbl) if paragraph_text(p).strip()]
        return bool(paras) and all(self._is_monospace(p) for p in paras)

    def _resolve_captions(self) -> None:
        for b in self.blocks:
            if b.ptype == "FORMULA_NUM":
                prev = self._prev(b)
                if prev is None or prev.ptype != "FORMULA":
                    b.ptype = "BODY"
            if b.ptype != "PRECAP":
                continue
            cap = b.caption
            prev, nxt = self._prev(b), self._next(b)
            has_dash = cap["dash"] is not None
            kind = cap["kind"]
            short = len(b.text) <= 300
            if kind == "fig" and (self._is_picture_block(prev) or self._is_picture_block(nxt) or (has_dash and short)):
                b.ptype = "CAP_FIG"
            elif kind == "tbl" and ((nxt and nxt.ptype == "TABLE") or (prev and prev.ptype == "TABLE") or (has_dash and short)):
                b.ptype = "CAP_TBL"
            elif kind == "lst" and ((has_dash and short) or (nxt and (nxt.ptype == "TABLE" or (nxt.kind == "p" and self._is_monospace(nxt.el))))):
                b.ptype = "CAP_LST"
            else:
                b.ptype = "LIST" if self.res.num_ref(b.el) is not None else "BODY"
                b.caption = None
        # номера формул
        for b in self.blocks:
            if b.ptype != "FORMULA":
                continue
            m = FORMULA_NUM_RE.search(b.text) or FORMULA_NUM_RE.search(math_text(b.el))
            num_block = b if m else None
            if m is None:
                nxt = self._next(b)
                if nxt is not None and nxt.ptype == "FORMULA_NUM":
                    m = FORMULA_NUM_RE.search(nxt.text)
                    num_block = nxt
            b.caption = {"num": m.group(1), "block": num_block, "in_text": bool(num_block and FORMULA_NUM_RE.search(num_block.text))} if m else None

    # ================================================================== context
    def _context_pass(self) -> None:
        section, chapter, chapter_counter = "", None, 0
        counters: Counter = Counter()
        in_sources = in_appendix = listing = appendix_title_pending = False
        manual_toc = False
        last_table_num = None
        for b in self.blocks:
            if b.ptype == "FRONT":
                continue
            if b.ptype == "HEADING":
                hk, level, _, info = b.heading
                section = b.text.strip()
                listing = appendix_title_pending = False
                manual_toc = hk == "toc"
                if hk in ("structural", "toc"):
                    chapter, counters = None, Counter()
                    in_sources = " ".join(b.text.split()).upper().rstrip(".") in R.SOURCES_HEADINGS
                    in_appendix = False
                elif hk == "appendix":
                    chapter, counters = info.get("letter"), Counter()
                    in_sources, in_appendix = False, True
                    appendix_title_pending = not info.get("title")
                elif hk == "chapter":
                    num = info.get("num")
                    chapter_counter = int(num) if num else chapter_counter + 1
                    chapter, counters = str(chapter_counter), Counter()
                    in_sources = in_appendix = False
                b.section, b.chapter = section, chapter
                continue
            b.section, b.chapter, b.in_sources, b.in_appendix = section, chapter, in_sources, in_appendix
            # Содержание, набранное вручную, — строки до следующего заголовка
            if manual_toc and not self.has_toc_field and b.ptype in ("BODY", "LIST"):
                b.ptype = "TOC"
                continue
            if appendix_title_pending and b.ptype != "EMPTY":
                appendix_title_pending = False
                if b.ptype == "BODY" and len(b.text.strip()) <= 200:
                    b.ptype = "APPX_TITLE"
                    continue
            if b.ptype == "CAP_LST":
                listing = True
            elif b.ptype in ("BODY", "LIST") and self._is_monospace(b.el):
                b.ptype = "CODE"
                b.extra["no_listing"] = not (listing or in_appendix)
            elif b.ptype == "TABLE" and (listing or self._is_code_table(b.el)):
                b.special = "code"
                b.extra["no_listing"] = not (listing or in_appendix)
                listing = False
            elif b.ptype not in ("EMPTY", "CODE"):
                listing = False

            if b.ptype in ("CAP_FIG", "CAP_TBL", "CAP_LST") or (b.ptype == "PICTURE" and b.caption):
                kind = b.caption["kind"]
                counters[kind] += 1
                b.expected = f"{chapter}.{counters[kind]}" if chapter else None
                if kind == "tbl":
                    last_table_num = b.expected or b.caption["num"].rstrip(".")
            elif b.ptype == "CAP_CONT":
                b.expected = last_table_num
            elif b.ptype == "FORMULA" and b.caption:
                counters["eq"] += 1
                b.expected = f"{chapter}.{counters['eq']}" if chapter else None

    # ================================================================== dispatch
    def _check_block(self, b: Block) -> None:
        t = b.ptype
        if t == "CODE" or (t == "TABLE" and b.special == "code"):
            prev = self._prev(b)
            if b.extra.get("no_listing") and not (prev is not None and prev.ptype == "CODE"):
                self.add("code_no_listing", b)
            return
        if t in ("FRONT", "EMPTY", "PRECAP"):
            return
        if t == "HEADING":
            self._check_heading(b)
        elif t == "TOC":
            self._check_toc_entry(b)
        elif t == "BODY":
            self._check_body(b)
        elif t == "LIST":
            self._check_body(b, list_item=True)
        elif t == "PICTURE":
            self._check_picture(b)
        elif t == "FORMULA":
            self._check_formula(b)
        elif t == "FORMULA_NUM":
            self._check_formula_number_line(b)
        elif t in ("CAP_FIG", "CAP_TBL", "CAP_LST", "CAP_CONT"):
            self._check_caption(b)
        elif t == "APPX_TITLE":
            self._check_runs(b, allow_bold=True)
            self._check_alignment(b, "center", {"center"}, 0, "heading_align", "heading_indent")
        elif t == "TABLE":
            self._check_table(b)
        if b.kind == "p" and b.ptype != "PICTURE" and b.ptype not in ("HEADING",):
            self._check_floating_pictures(b)

    # ================================================================== общие проверки
    def _check_runs(self, b: Block, allow_bold: bool = False, size: float | None = -1.0,
                    font_code: str = "font_name") -> None:
        p, res = b.el, self.res
        if size == -1.0:
            size = self.p.font_size
        runs = text_runs(p)
        if not runs:
            return
        fonts, sizes, colors = set(), set(), set()
        bad_font = bad_size = bad_color = bad_u = bad_b = False
        for r in runs:
            name = res.font_name(r, p)
            if name.strip().lower() != self.p.font_name.lower():
                fonts.add(name)
                bad_font = True
            if size is not None:
                sz = res.font_size(r, p)
                if abs(sz - size) > 0.01:
                    sizes.add(sz)
                    bad_size = True
            c = res.color(r, p)
            if c:
                colors.add(c)
                bad_color = True
            if res.underline(r, p):
                bad_u = True
            if not allow_bold and res.toggle(r, p, "w:b"):
                bad_b = True
        all_runs = paragraph_runs(p)
        if bad_font:
            self.add(font_code, b, ", ".join(sorted(fonts)),
                     fix=lambda: F.set_font_name(all_runs, self.p.font_name))
        if bad_size:
            self.add("font_size", b, ", ".join(f"{num_ru(s)} пт" for s in sorted(sizes)),
                     fix=lambda: F.set_font_size(all_runs, size))
        if bad_color:
            self.add("font_color", b, ", ".join(sorted(colors)), fix=lambda: F.set_black(all_runs))
        if bad_u:
            self.add("underline", b, fix=lambda: F.remove_underline(all_runs))
        if bad_b:
            self.add("bold_body", b, fix=lambda: F.set_bold(all_runs, False))

    def _check_alignment(self, b: Block, target: str, allowed: set[str], first: int | None,
                         align_code: str, indent_code: str) -> None:
        p = b.el
        al = self.res.alignment(p)
        if al not in allowed:
            self.add(align_code, b, f"сейчас {ALIGN_RU.get(al, al)}, нужно {ALIGN_RU[target]}",
                     fix=lambda: F.set_alignment(p, target))
        if first is None:
            return
        left, cur_first, right = self.res.indent(p)
        if abs(cur_first - first) > TOL or abs(left) > TOL or abs(right) > TOL:
            parts = [f"первая строка {cm(cur_first)}" if cur_first >= 0 else f"выступ {cm(-cur_first)}"]
            if abs(left) > TOL:
                parts.append(f"слева {cm(left)}")
            if abs(right) > TOL:
                parts.append(f"справа {cm(right)}")
            self.add(indent_code, b, f"{', '.join(parts)}; нужно {cm(first) if first else 'без отступа'}",
                     fix=lambda: F.set_indent(p, first, 0, 0))

    def _check_spacing(self, b: Block, line: float | None = -1.0) -> None:
        p = b.el
        if line == -1.0:
            line = self.p.line_spacing
        before, after, value, rule = self.res.spacing(p)
        want_before, want_after = round(self.p.space_before_pt * 20), round(self.p.space_after_pt * 20)
        if abs(before - want_before) > 5 or abs(after - want_after) > 5:
            self.add("spacing_body", b,
                     f"перед {num_ru(before / 20)} пт, после {num_ru(after / 20)} пт; "
                     f"нужно {num_ru(self.p.space_before_pt)}/{num_ru(self.p.space_after_pt)} пт",
                     fix=lambda: F.set_spacing(p, self.p.space_before_pt, self.p.space_after_pt))
        if line is not None and (rule != "auto" or abs(value - line) > 0.02):
            cur = num_ru(round(value, 2)) if rule == "auto" else f"{'точно' if rule == 'exact' else 'минимум'} {num_ru(value)} пт"
            self.add("line_spacing_body", b, f"сейчас {cur}", fix=lambda: F.set_line_spacing(p, line))

    def _check_floating_pictures(self, b: Block) -> None:
        anchors = [a for a in b.el.iter(WP_ANCHOR) if not _is_excluded(a, b.el)]
        if anchors:
            self.add("picture_wrap", b, "рисунок «плавает» поверх текста — нужно «В тексте»")
        if b.extra.get("picture_in_text"):
            big = any(_int_attr(ext, "cx") > 1_080_000 for ext in b.el.iter(clark("wp:extent")))
            if big:
                self.add("inline_picture_in_text", b)

    # ================================================================== основной текст
    def _check_body(self, b: Block, list_item: bool = False) -> None:
        align = self.p.body_align
        self._check_runs(b)
        self._check_alignment(b, align, {align}, None, "align_body", "indent_body")
        if not list_item and not self._is_where_line(b):
            self._check_alignment(b, align, {"both", "left", "center", "right"}, self.indent,
                                  "align_body", "indent_body")
        self._check_spacing(b)
        if b.in_sources and URL_RE.search(b.text) and "дата обращения" not in b.text.lower():
            self.add("source_access_date", b)

    def _is_where_line(self, b: Block) -> bool:
        """Строка «где …» сразу после формулы пишется без отступа — не считаем это ошибкой."""
        prev = self._prev(b)
        return bool(prev and prev.ptype in ("FORMULA", "FORMULA_NUM") and b.text.strip().lower().startswith("где"))

    # ================================================================== заголовки
    def _check_heading(self, b: Block) -> None:
        hk, level, styled, info = b.heading
        p, res = b.el, self.res
        text = b.text.strip()
        self._check_runs(b, allow_bold=True)
        runs = text_runs(p)
        all_runs = paragraph_runs(p)
        if runs and not all(res.toggle(r, p, "w:b") for r in runs):
            self.add("heading_bold", b, fix=lambda: F.set_bold(all_runs, True))

        letters = [c for c in text if c.isalpha()]
        typed_upper = bool(letters) and all(c.isupper() for c in letters)
        caps_attr = bool(runs) and all(res.toggle(r, p, "w:caps") for r in runs)

        if hk in ("structural", "toc", "appendix"):
            sa = self.p.structural_align
            self._check_alignment(b, sa, {sa}, 0, "heading_align", "heading_indent")
            if not (typed_upper or caps_attr):
                self.add("heading_caps", b, fix=lambda: map_text_nodes(p, str.upper))
        else:
            ca = self.p.chapter_align
            self._check_alignment(b, ca, {ca}, self.indent if self.p.chapter_indent else 0,
                                  "heading_align", "heading_indent")
            if hk == "chapter" and self.p.chapter_caps and not (typed_upper or caps_attr):
                self.add("heading_caps", b, fix=lambda: map_text_nodes(p, str.upper))
            if hk == "sub" and self.p.sub_caps and not (typed_upper or caps_attr):
                self.add("heading_caps", b, fix=lambda: map_text_nodes(p, str.upper))
            if hk == "sub" and not self.p.sub_caps and len(letters) > 3 and (typed_upper or caps_attr):
                fix = (lambda: F.caps_off(all_runs)) if caps_attr and not typed_upper else None
                self.add("heading_not_caps", b, "исправьте регистр вручную" if fix is None else "", fix=fix)
            n = NUM_HEADING_RE.match(" ".join(b.text.split()))
            if n and n.group("dot"):
                self.add("heading_number_dot", b, f"«{n.group('num')}.» → «{n.group('num')}»",
                         fix=lambda: _remove_number_dot(p))

        if text.endswith(".") and not text.endswith("..") and hk != "toc":
            self.add("heading_end_dot", b, fix=lambda: _remove_trailing_dot(p))

        def apply_style(p=p, level=level):
            F.set_style(p, F.ensure_heading_style(self.doc, level, self.p))

        if not styled and hk != "toc":
            self.add("heading_style", b, f"нужен стиль «Заголовок {level}»", fix=apply_style)
        elif styled and hk in ("chapter", "sub") and info.get("numbered") and info["outline"] + 1 != level:
            self.add("heading_style", b,
                     f"стиль «Заголовок {info['outline'] + 1}», а для {'главы' if level == 1 else 'подглавы'} "
                     f"нужен «Заголовок {level}» — иначе содержание соберётся с неверной вложенностью",
                     fix=apply_style)

    def _check_toc_entry(self, b: Block) -> None:
        p, res = b.el, self.res
        runs = text_runs(p)
        bad = sorted({f"{res.font_name(r, p)} {num_ru(res.font_size(r, p))} пт" for r in runs
                      if res.font_name(r, p).lower() != self.p.font_name.lower()
                      or abs(res.font_size(r, p) - self.p.font_size) > 0.01})
        if bad:
            all_runs = paragraph_runs(p)
            self.add("toc_font", b, ", ".join(bad),
                     fix=lambda: F.normalize_runs(all_runs, self.p.font_size, self.p.font_name))

    # ================================================================== рисунки
    def _check_picture(self, b: Block) -> None:
        p = b.el
        self._check_alignment(b, "center", {"center"}, 0, "picture_align", "picture_align")
        anchors = [a for a in p.iter(WP_ANCHOR) if not _is_excluded(a, p)]
        if anchors:
            def to_inline(anchors=anchors):
                for a in anchors:
                    parent = a.getparent()
                    if parent is not None:
                        parent.replace(a, inline_from_anchor(a))
            self.add("picture_wrap", b, fix=to_inline if len(b.text.strip()) <= 3 else None)
        if b.caption:
            self.add("fig_caption_text", b, "подпись в одном абзаце с рисунком — вынесите её в отдельный абзац под рисунком")
            return
        nxt, prev = self._next(b), self._prev(b)
        if nxt is not None and nxt.ptype in ("CAP_FIG", "PICTURE"):
            return
        if prev is not None and prev.ptype == "CAP_FIG" and not self._is_picture_block(self._prev(prev)):
            self.add("fig_caption_position", prev)
            return
        self.add("picture_no_caption", b)

    # ================================================================== подписи
    def _check_caption(self, b: Block) -> None:
        p = b.el
        t = b.ptype
        self._check_runs(b)
        if t == "CAP_FIG":
            self._check_alignment(b, "center", {"center"}, 0, "fig_caption_align", "fig_caption_align")
        else:
            code = "lst_caption_align" if t == "CAP_LST" else "tbl_caption_align"
            self._check_alignment(b, "left", {"left", "both"}, 0, code, code)
        if t == "CAP_CONT":
            self._check_continuation(b)
            return

        kind = b.caption["kind"]
        text_code = {"fig": "fig_caption_text", "tbl": "tbl_caption_text", "lst": "lst_caption_text"}[kind]
        num_code = {"fig": "fig_number", "tbl": "tbl_number", "lst": "lst_number"}[kind]
        if kind == "tbl":
            nxt = self._next(b)
            if (nxt is None or nxt.ptype != "TABLE") and self._prev(b) is not None and self._prev(b).ptype == "TABLE":
                self.add("tbl_caption_position", b)

        cap = b.caption
        word, num, name, tail, sep, dash = cap["word"], cap["num"], cap["name"], cap["tail"], cap["sep"], cap["dash"]
        nnum = num.rstrip(".")
        canonical = CAPTION_WORD[kind]
        has_fields = next(p.iter(W_FLDCHAR, W_FLDSIMPLE), None) is not None

        problems = []
        if word != canonical:
            problems.append(f"«{word}» → «{canonical}»")
        if num.endswith("."):
            problems.append("точка после номера")
        if not name.strip():
            problems.append("нет названия")
        elif dash is None:
            problems.append("нет тире между номером и названием")
        elif dash not in "–—":
            problems.append(f"«{dash}» вместо тире «–»")
        elif sep.replace("\xa0", " ") != f" {dash} ":
            problems.append("лишние/недостающие пробелы вокруг тире")
        if "." in tail:
            problems.append("точка в конце")
        if cap["lead"]:
            problems.append("пробелы в начале")

        new_num = nnum
        number_detail = None
        if not VALID_NUM.match(nnum):
            number_detail = f"«{nnum}» — нужен формат «номер главы.номер»"
            if b.expected:
                number_detail += f", ожидается «{b.expected}»"
                new_num = b.expected
        elif b.expected and nnum != b.expected:
            number_detail = f"«{nnum}» → «{b.expected}»"
            new_num = b.expected
        renumber = new_num != nnum and not has_fields

        m = cap["match"]
        spans = {k: (m.start(k), m.end(k)) for k in ("lead", "word", "num", "sep", "name", "tail")}
        name_ok = bool(name.strip())

        def fix_caption(p=p, spans=spans, dash=dash, name_ok=name_ok):
            edits = []
            if spans["tail"][1] > spans["tail"][0]:
                edits.append((spans["tail"], ""))
            if name_ok:
                good_dash = dash if dash in ("–", "—") else self.p.caption_dash
                edits.append((spans["sep"], f" {good_dash} "))
            if renumber or num.endswith("."):
                edits.append((spans["num"], new_num if renumber else nnum))
            if word != canonical:
                edits.append((spans["word"], canonical))
            if spans["lead"][1] > spans["lead"][0]:
                edits.append((spans["lead"], ""))
            for (s, e), new in sorted(edits, key=lambda x: x[0][0], reverse=True):
                replace_span(p, s, e, new)

        text_fixable = name_ok and not (has_fields and num.endswith("."))
        if problems:
            self.add(text_code, b, "; ".join(problems), fix=fix_caption if text_fixable else None)
        if number_detail:
            self.add(num_code, b, number_detail + (" (подпись использует поля — обновите их в Word)" if has_fields else ""),
                     fix=fix_caption if renumber else None)
            if renumber and kind in self.renumber:
                mapping = self.renumber[kind]
                if nnum in mapping and mapping[nnum] != new_num:
                    self.renumber_conflicts[kind].add(nnum)
                mapping[nnum] = new_num
        b.extra["number"] = nnum

    def _check_continuation(self, b: Block) -> None:
        p = b.el
        s = b.text.strip()
        expected = b.expected
        m = CONT_FULL.match(s)
        if m is None:
            canonical = f"Продолжение таблицы {expected}" if expected else None
            fix = None
            if canonical and canonical != s:   # без этой проверки правка переписывала текст сама в себя
                fix = lambda: replace_span(p, 0, len(paragraph_text(p)), canonical)
            self.add("tbl_caption_text", b, "формат: «Продолжение таблицы 1.1»", fix=fix)
            return
        num = m.group(1)
        if expected and num != expected:
            start = b.text.rfind(num)
            self.add("tbl_number", b, f"«{num}» → «{expected}» (продолжение последней таблицы)",
                     fix=lambda: replace_span(p, start, start + len(num), expected))

    # ================================================================== формулы
    def _check_formula(self, b: Block) -> None:
        p = b.el
        omp = p.find(f".//{M_OMATHPARA}")
        centered, math_jc = True, None
        if omp is not None:
            math_jc = omp.find(f"{clark('m:oMathParaPr')}/{clark('m:jc')}")
            if math_jc is not None and math_jc.get(clark("m:val")) in ("left", "right"):
                centered = False
        elif "\t" not in b.text and self.res.alignment(p) != "center":
            centered = False
        left, first, _ = self.res.indent(p)
        if not centered or abs(first) > TOL or abs(left) > TOL:
            def fix(p=p, math_jc=math_jc):
                if math_jc is not None:
                    math_jc.set(clark("m:val"), "centerGroup")
                if "\t" not in paragraph_text(p):
                    F.set_alignment(p, "center")
                F.set_indent(p, 0, 0, 0)
            self.add("formula_align", b, fix=fix)

        cap = b.caption
        if cap is None:
            self.add("formula_no_number", b)
        else:
            num = cap["num"]
            valid = VALID_NUM.match(num)
            if not valid or (b.expected and num != b.expected):
                detail = f"({num}) → ({b.expected})" if b.expected else f"({num}) — нужен формат (глава.номер)"
                fix = None
                block = cap["block"]
                if b.expected and cap["in_text"] and block is not None:
                    m = FORMULA_NUM_RE.search(block.text)
                    s, e = m.span(1)
                    fix = lambda p2=block.el, s=s, e=e, new=b.expected: replace_span(p2, s, e, new)
                self.add("formula_number", b, detail, fix=fix)

        prev = self.blocks[b.idx - 1] if b.idx > 0 else None
        if prev is not None and prev.ptype not in ("EMPTY", "FORMULA", "FORMULA_NUM", "HEADING", "FRONT"):
            def insert_blank(p=p):
                p.addprevious(OxmlElement("w:p"))
            self.add("formula_blank", b, fix=insert_blank)

    def _check_formula_number_line(self, b: Block) -> None:
        p = b.el
        al = self.res.alignment(p)
        if al != "right" and "\t" not in b.text:
            self.add("formula_number_align", b, f"сейчас {ALIGN_RU.get(al, al)}",
                     fix=lambda: (F.set_alignment(p, "right"), F.set_indent(p, 0, 0, 0)))
        self._check_runs(b)

    # ================================================================== таблицы
    def _check_table(self, b: Block) -> None:
        tbl, res = b.el, self.res
        paras = list(iter_paragraphs(tbl))
        rows = tbl.findall(qn("w:tr"))
        prev, nxt = self._prev(b), self._next(b)
        text_total = sum(len(paragraph_text(p).strip()) for p in paras)
        if any(has_math(p) for p in paras) and len(rows) <= 3 and text_total <= 20:
            b.special = "formula"
            return
        if any(has_picture(p) for p in paras) and (
            (nxt is not None and nxt.ptype == "CAP_FIG")
            or any((parse_caption(paragraph_text(p).strip()) or {}).get("kind") == "fig" for p in paras)
        ):
            b.special = "figure"
            return
        if b.special == "code":
            return

        if prev is None or prev.ptype not in ("CAP_TBL", "CAP_CONT"):
            if not (nxt is not None and nxt.ptype == "CAP_TBL"):
                self.add("table_no_caption", b)

        fonts, colors = set(), set()
        underline = False
        sizes: Counter = Counter()
        lines: Counter = Counter()
        for p in paras:
            for r in text_runs(p):
                name = res.font_name(r, p)
                if name.lower() != self.p.font_name.lower() and name.lower() not in R.MONOSPACE_FONTS:
                    fonts.add(name)
                c = res.color(r, p)
                if c:
                    colors.add(c)
                underline = underline or res.underline(r, p)
                sizes[res.font_size(r, p)] += len(paragraph_text(p))
            if paragraph_text(p).strip():
                _, _, value, rule = res.spacing(p)
                lines[(round(value, 2), rule)] += 1
        all_runs = [r for p in paras for r in paragraph_runs(p)]
        if fonts or colors or underline:
            parts = sorted(fonts) + sorted(colors) + (["подчёркивание"] if underline else [])

            def fix_fonts(all_runs=all_runs):
                F.set_font_name(all_runs, self.p.font_name)
                F.set_black(all_runs)
                F.remove_underline(all_runs)
            self.add("table_font", b, ", ".join(parts), fix=fix_fonts)

        if sizes:
            lo, hi = self.p.table_font_min, self.p.table_font_max
            lo_line, hi_line = self.p.table_line_min, self.p.table_line_max
            bad_sizes = sorted(s for s in sizes if not lo <= s <= hi)
            bad_lines = sorted(v for (v, rule) in lines if rule != "auto" or not lo_line <= v <= hi_line)
            dom_size = sizes.most_common(1)[0][0]
            dom_line = lines.most_common(1)[0][0] if lines else (lo_line, "auto")
            profile = (min(max(dom_size, lo), hi),
                       min(max(dom_line[0], lo_line), hi_line) if dom_line[1] == "auto" else lo_line)
            self.table_profiles.append((b, profile))
            if bad_sizes or bad_lines:
                detail = []
                if bad_sizes:
                    detail.append("шрифт " + ", ".join(f"{num_ru(s)} пт" for s in bad_sizes))
                if bad_lines:
                    detail.append("интервал " + ", ".join(num_ru(v) for v in bad_lines))
                self.add("table_size", b, "; ".join(detail), fix=self._table_format_fix(tbl, None))

        if len(rows) >= 2 and not (prev is not None and prev.ptype == "CAP_CONT"):
            first_cells = [
                " ".join(paragraph_text(p) for p in iter_paragraphs(tc)).strip()
                for tc in rows[0].findall(qn("w:tc"))
            ]
            filled = [c for c in first_cells if c]
            if not filled or all(re.fullmatch(r"[\d\s.,\-–]+", c) for c in filled):
                self.add("table_header", b)

    def _table_format_fix(self, tbl, profile: tuple[float, float] | None):
        def fix():
            size, line = profile or self.table_target
            for p in iter_paragraphs(tbl):
                F.set_font_size(paragraph_runs(p), size)
                F.set_line_spacing(p, line)
        return fix

    def _check_tables_uniform(self) -> None:
        if not self.table_profiles:
            return
        counts = Counter(profile for _, profile in self.table_profiles)
        majority = counts.most_common(1)[0][0]
        self.table_target = majority
        if len(counts) == 1:
            return
        for b, profile in self.table_profiles:
            if profile != majority:
                self.add(
                    "table_uniform", b,
                    f"шрифт {num_ru(profile[0])} пт, интервал {num_ru(profile[1])}; "
                    f"в большинстве таблиц — {num_ru(majority[0])} пт, {num_ru(majority[1])}",
                    fix=self._table_format_fix(b.el, majority),
                )

    # ================================================================== перечисления
    def _check_lists(self) -> None:
        res = self.res
        reported_levels: set[tuple] = set()
        self._checked_list_numbers: set[int] = set()
        groups: list[list[Block]] = []
        current: list[Block] = []
        current_num = None
        for b in self.blocks:
            ref = res.num_ref(b.el) if b.ptype == "LIST" else None
            if ref is None:
                if b.ptype != "EMPTY" and current:
                    groups.append(current)
                    current, current_num = [], None
                continue
            if current and ref[0] != current_num:
                groups.append(current)
                current = []
            current.append(b)
            current_num = ref[0]
        if current:
            groups.append(current)

        for group in groups:
            first = group[0]
            bad_indent = []
            for b in group:
                num_id, ilvl = res.num_ref(b.el)
                lvl, abstract = res.num_level(num_id, ilvl)
                if lvl is None:
                    continue
                key = (id(abstract), ilvl)
                if key not in reported_levels:
                    reported_levels.add(key)
                    suff = lvl.find(qn("w:suff"))
                    if suff is None or suff.get(qn("w:val")) == "tab":
                        self.add("list_tab", b,
                                 fix=lambda abstract=abstract: [set_lvl_child(l, "w:suff", {"w:val": "space"})
                                                                for l in abstract.findall(qn("w:lvl"))])
                    fmt = lvl.find(qn("w:numFmt"))
                    lvl_text = lvl.find(qn("w:lvlText"))
                    marker = lvl_text.get(qn("w:val")) if lvl_text is not None else ""
                    if fmt is not None and fmt.get(qn("w:val")) == "bullet" and marker not in ("–", "—", "-", "−", ""):
                        def fix_marker(lvl=lvl):
                            set_lvl_child(lvl, "w:lvlText", {"w:val": self.p.caption_dash})
                            rpr = lvl.find(qn("w:rPr"))
                            if rpr is None:
                                rpr = OxmlElement("w:rPr")
                                lvl.append(rpr)
                            rf = rpr.find(qn("w:rFonts"))
                            if rf is None:
                                rf = OxmlElement("w:rFonts")
                                rpr.insert(0, rf)
                            for attr in ("ascii", "hAnsi", "cs"):
                                rf.set(qn(f"w:{attr}"), self.p.font_name)
                            rf.attrib.pop(qn("w:hint"), None)
                        shown = "".join(
                            _SYMBOL_CHARS.get(f"{ord(ch):X}"[-2:], "?") if 0xF000 <= ord(ch) <= 0xF0FF else ch
                            for ch in marker
                        )
                        self.add("list_marker", b, f"сейчас «{shown or '?'}»", fix=fix_marker)
                self._check_list_number(b, group, lvl)
                if ilvl == 0:
                    left, first_line, _ = res.indent(b.el)
                    if abs(left) > TOL or abs(first_line - self.indent) > TOL:
                        bad_indent.append((b, left, first_line))
            if bad_indent:
                b0, left, first_line = bad_indent[0]
                items = [b for b, _, _ in bad_indent]
                self.add(
                    "list_indent", b0,
                    f"номер на {cm(left + first_line)}, текст с {cm(left)}; "
                    f"нужно {cm(self.indent)} и 0 см"
                    + (f" (пунктов: {len(items)})" if len(items) > 1 else ""),
                    fix=lambda items=items: [F.set_indent(x.el, self.indent, 0, 0) for x in items],
                )

    def _check_list_number(self, b: Block, group: list[Block], lvl) -> None:
        """Номер/маркер перечисления оформляется знаком абзаца — сам текст пункта его не задаёт."""
        if id(lvl) in self._checked_list_numbers:
            return
        self._checked_list_numbers.add(id(lvl))
        res, p = self.res, b.el
        lvl_rpr = lvl.find(qn("w:rPr"))
        chain = ([lvl_rpr] if lvl_rpr is not None else []) + res.mark_chain(p)
        runs = text_runs(p)
        text_bold = bool(runs) and all(res.toggle(r, p, "w:b") for r in runs)
        fmt = lvl.find(qn("w:numFmt"))
        lvl_text = lvl.find(qn("w:lvlText"))
        marker = (lvl_text.get(qn("w:val")) or "") if lvl_text is not None else ""
        # маркер-символ (например, тире из шрифта Symbol) рисуется своим шрифтом — это не ошибка
        symbol_marker = (fmt is not None and fmt.get(qn("w:val")) == "bullet"
                         and bool(marker) and 0xF000 <= ord(marker[0]) <= 0xF0FF)
        name, size = res.font_name_in(chain), res.font_size_in(chain)
        color, bold = res.color_in(chain), res.toggle_in(chain, "w:b")
        problems = []
        if not symbol_marker and name.strip().lower() != self.p.font_name.lower():
            problems.append(name)
        if abs(size - self.p.font_size) > 0.01:
            problems.append(f"{num_ru(size)} пт")
        if color:
            problems.append(color)
        if bold and not text_bold:
            problems.append("полужирный")
        if not problems:
            return

        def fix_numbers(items=list(group), lvl=lvl, text_bold=text_bold, symbol_marker=symbol_marker):
            for item in items:
                F.normalize_paragraph_mark(item.el, self.p, bold=text_bold)
            F.normalize_numbering_level(lvl, self.p, set_font=not symbol_marker)

        self.add("list_number_format", b, ", ".join(problems)
                 + (f" (пунктов: {len(group)})" if len(group) > 1 else ""), fix=fix_numbers)

    # ================================================================== ссылки
    def _ref_paragraphs(self):
        for b in self.blocks:
            if b.ptype in ("BODY", "LIST", "APPX_TITLE"):
                yield b, b.el, b.text
            elif b.ptype == "TABLE" and not b.special:
                for p in iter_paragraphs(b.el):
                    yield b, p, paragraph_text(p)

    @staticmethod
    def _numbers_in(match) -> list[str]:
        nums = match.group("nums")
        tokens = list(NUM_TOKEN_RE.finditer(nums))
        found = [t.group(0).upper() for t in tokens]
        for a, c in zip(tokens, tokens[1:]):
            between = nums[a.end():c.start()].strip()
            pa, pc = a.group(0).upper().rsplit(".", 1), c.group(0).upper().rsplit(".", 1)
            if between in ("–", "—", "-") and pa[0] == pc[0] and pa[1].isdigit() and pc[1].isdigit():
                found += [f"{pa[0]}.{i}" for i in range(int(pa[1]) + 1, int(pc[1]))]
        return found

    def _check_references(self) -> None:
        fig_refs: dict[str, int] = {}
        tbl_refs: dict[str, int] = {}
        for b, p, text in self._ref_paragraphs():
            for m in FIG_REF_RE.finditer(text):
                for n in self._numbers_in(m):
                    fig_refs.setdefault(n, b.idx)
            for m in TBL_REF_RE.finditer(text):
                for n in self._numbers_in(m):
                    tbl_refs.setdefault(n, b.idx)
            if b.ptype in ("BODY", "LIST") and FIG_ABBR_RE.search(text):
                self.add("fig_ref_abbrev", b)

        for b in self.blocks:
            if b.ptype == "CAP_FIG" or (b.ptype == "PICTURE" and b.caption):
                num = b.caption["num"].rstrip(".").upper()
                if num not in fig_refs:
                    self.add("fig_ref_missing", b, f"рисунок {num}")
                elif fig_refs[num] > b.idx:
                    self.add("fig_ref_order", b, f"рисунок {num}")
            elif b.ptype == "CAP_TBL":
                num = b.caption["num"].rstrip(".").upper()
                if num not in tbl_refs:
                    self.add("tbl_ref_missing", b, f"таблица {num}")

    def _register_post_fixes(self) -> None:
        mappings = {
            kind: {old: new for old, new in mapping.items() if old not in self.renumber_conflicts[kind]}
            for kind, mapping in self.renumber.items()
        }
        if not any(mappings.values()):
            return

        def update_references():
            for kind, regex in (("fig", FIG_REF_RE), ("tbl", TBL_REF_RE)):
                mapping = mappings[kind]
                if not mapping:
                    continue
                for _, p, _ in list(self._ref_paragraphs()):
                    text = paragraph_text(p)
                    edits = []
                    for m in regex.finditer(text):
                        base = m.start("nums")
                        for t in NUM_TOKEN_RE.finditer(m.group("nums")):
                            old = t.group(0)
                            if old in mapping:
                                edits.append((base + t.start(), base + t.end(), mapping[old]))
                    for s, e, new in sorted(edits, reverse=True):
                        replace_span(p, s, e, new)
        self.analysis.post_fixes.append(update_references)

    # ================================================================== содержание
    def _check_toc(self) -> None:
        headings = [b for b in self.blocks if b.ptype == "HEADING"]
        toc_heads = [b for b in headings if b.heading[0] == "toc"]
        chapters = [b for b in headings if b.heading[0] == "chapter"]
        glava = [b for b in chapters if GLAVA_RE.match(b.text.strip())]
        if glava and len(glava) != len(chapters):
            glava_ids = {id(b) for b in glava}
            odd = (next(b for b in chapters if id(b) not in glava_ids)
                   if len(glava) >= len(chapters) / 2 else glava[0])
            self.add("chapter_style_mixed", odd)
        if chapters:
            expected = "ОГЛАВЛЕНИЕ" if glava and len(glava) >= len(chapters) / 2 else "СОДЕРЖАНИЕ"
            for b in toc_heads:
                current = " ".join(b.text.split()).upper().rstrip(".")
                if current != expected:
                    reason = "главы названы «ГЛАВА N»" if expected == "ОГЛАВЛЕНИЕ" else "главы пронумерованы без слова «ГЛАВА»"
                    start = b.text.upper().find(current)
                    p = b.el
                    self.add("toc_title", b, f"«{current}» → «{expected}», т.к. {reason}",
                             fix=lambda p=p, s=start, e=start + len(current), new=expected: replace_span(p, s, e, new))
        if toc_heads and not self.has_toc_field:
            self.add("toc_manual", toc_heads[0])
        if not toc_heads and not self.has_toc_field and len(headings) >= 3:
            self.add("toc_missing", headings[0], where="начало документа")

    # ================================================================== страница
    def _first_anchor(self):
        for b in self.blocks:
            if b.kind == "p" and b.ptype != "FRONT":
                return b.el
        return self.blocks[0].el if self.blocks and self.blocks[0].kind == "p" else None

    def _check_page_setup(self) -> None:
        anchor = self._first_anchor()
        many = len(self.doc.sections) > 1
        report_sections = {b.extra.get("sect") for b in self.blocks if b.ptype != "FRONT"}
        for i, s in enumerate(self.doc.sections, start=1):
            if many and (i - 1) not in report_sections:
                continue  # раздел целиком состоит из титульного листа/задания
            where = f"раздел документа {i}" if many else "параметры страницы"
            wrong = []
            margins = {"left": self.p.margin_left_mm, "right": self.p.margin_right_mm,
                       "top": self.p.margin_top_mm, "bottom": self.p.margin_bottom_mm}
            for side, target in margins.items():
                value = getattr(s, f"{side}_margin")
                if value is None or abs(value.mm - target) > R.MARGIN_TOLERANCE_MM:
                    names = {"left": "левое", "right": "правое", "top": "верхнее", "bottom": "нижнее"}
                    cur = f"{value.mm:.0f}" if value is not None else "?"
                    wrong.append(f"{names[side]} {cur} мм (нужно {target:.0f})")
            if wrong:
                def fix_margins(s=s):
                    s.left_margin, s.right_margin = Mm(self.p.margin_left_mm), Mm(self.p.margin_right_mm)
                    s.top_margin, s.bottom_margin = Mm(self.p.margin_top_mm), Mm(self.p.margin_bottom_mm)
                self.add("margins", None, "; ".join(wrong), fix=fix_margins, where=where, anchor=anchor)
            hf = []
            for attr, label in (("header_distance", "верхний"), ("footer_distance", "нижний")):
                value = getattr(s, attr)
                if value is None or abs(value.cm - self.p.hf_distance_cm) > 0.03:
                    hf.append(f"{label} {num_ru(round(value.cm, 2)) if value is not None else '?'} см")
            if hf:
                def fix_hf(s=s):
                    s.header_distance = s.footer_distance = Cm(self.p.hf_distance_cm)
                self.add("hf_distance", None, "; ".join(hf) + f" (нужно {num_ru(self.p.hf_distance_cm)} см)",
                         fix=fix_hf, where=where, anchor=anchor)

    def _check_page_numbers(self) -> None:
        anchor = self._first_anchor()
        found_any = False
        for i, s in enumerate(self.doc.sections, start=1):
            for kind, footer in (("default", s.footer), ("first", s.first_page_footer), ("even", s.even_page_footer)):
                if footer.is_linked_to_previous:
                    continue
                root = footer._element
                page_paras = [p for p in root.iter(W_P) if _has_page_field(p)]
                if not page_paras:
                    continue
                if kind == "first" and not s.different_first_page_header_footer:
                    continue  # колонтитул первой страницы не используется
                if kind == "first" and i == 1:
                    def clear_first(page_paras=page_paras):
                        for p in page_paras:
                            for child in list(p):
                                if child.tag != qn("w:pPr"):
                                    p.remove(child)
                    self.add("page_number_title", None, "номер стоит в колонтитуле первой страницы",
                             fix=clear_first, where="титульный лист", anchor=anchor)
                    continue
                found_any = True
                has_title_page = any(b.ptype == "FRONT" for b in self.blocks)
                if i == 1 and kind == "default" and not s.different_first_page_header_footer and has_title_page:
                    def set_title_page(s=s):
                        s.different_first_page_header_footer = True
                    self.add("page_number_title", None, fix=set_title_page, where="титульный лист", anchor=anchor)
                for p in page_paras:
                    self._check_page_number_paragraph(p, i, anchor)
                empties = [p for p in root.iter(W_P) if p not in page_paras and _is_blank(p)]
                brs = [br for p in page_paras for br in p.iter(qn("w:br"))]
                if empties or brs:
                    def remove_blank(empties=empties, brs=brs):
                        for p in empties:
                            parent = p.getparent()
                            if parent is not None and len(parent.findall(W_P)) > 1:
                                parent.remove(p)
                        for br in brs:
                            br.getparent().remove(br)
                    self.add("footer_empty", None, f"пустых строк: {len(empties) + len(brs)}", fix=remove_blank,
                             where=f"нижний колонтитул (раздел {i})", anchor=anchor)
        if not found_any:
            self.add("page_numbers_missing", None, fix=lambda: F.add_page_numbers(self.doc, self.p),
                     where="весь документ", anchor=anchor)

    def _check_page_number_paragraph(self, p, section_no: int, anchor) -> None:
        res = self.res
        problems = []
        al = res.alignment(p)
        text = paragraph_text(p)
        leading_tabs = len(text) - len(text.lstrip("\t"))
        if al != "center":
            problems.append(f"выравнивание {ALIGN_RU.get(al, al)}")
        left, first, _ = res.indent(p)
        if abs(first) > TOL or abs(left) > TOL:
            problems.append(f"отступ {cm(first)}")
        runs = [r for r in paragraph_runs(p)]
        bad_font = {f"{res.font_name(r, p)} {num_ru(res.font_size(r, p))} пт" for r in runs
                    if run_has_content(r) and (res.font_name(r, p).lower() != self.p.font_name.lower()
                                               or abs(res.font_size(r, p) - self.p.font_size) > 0.01)}
        if bad_font:
            problems.append("шрифт " + ", ".join(sorted(bad_font)))
        if problems:
            def fix(p=p, runs=runs, leading_tabs=leading_tabs):
                if leading_tabs:
                    replace_span(p, 0, leading_tabs, "")
                F.set_alignment(p, "center")
                F.set_indent(p, 0, 0, 0)
                F.set_font_name(runs, self.p.font_name)
                F.set_font_size(runs, self.p.font_size)
            self.add("page_number_format", None, "; ".join(problems), fix=fix,
                     where=f"нижний колонтитул (раздел {section_no})", anchor=anchor)

    # ================================================================== статистика
    def _stats(self) -> None:
        c = Counter(b.ptype for b in self.blocks)
        self.analysis.stats = {
            "paragraphs": sum(1 for b in self.blocks if b.kind == "p" and b.text.strip()),
            "headings": c["HEADING"],
            "pictures": c["PICTURE"],
            "tables": c["TABLE"],
            "formulas": c["FORMULA"],
            "front": c["FRONT"],
        }


def _remove_number_dot(p) -> None:
    text = paragraph_text(p)
    m = re.match(r"^\s*\d{1,2}(?:\.\d{1,2}){0,3}(\.)\s", text)
    if m:
        replace_span(p, m.start(1), m.end(1), "")


def _remove_trailing_dot(p) -> None:
    text = paragraph_text(p)
    end = len(text.rstrip())
    if end and text[end - 1] == "." and not text[:end].endswith(".."):
        replace_span(p, end - 1, end, "")


def _int_attr(el, name: str) -> int:
    try:
        return int(el.get(name) or 0)
    except ValueError:
        return 0


def run_has_content(r) -> bool:
    return any(ch.tag in (qn("w:t"), qn("w:fldChar"), qn("w:instrText")) for ch in r)


def _has_page_field(p) -> bool:
    for node in p.iter(W_INSTR):
        if re.search(r"\bPAGE\b", node.text or ""):
            return True
    for node in p.iter(W_FLDSIMPLE):
        if re.search(r"\bPAGE\b", node.get(qn("w:instr")) or ""):
            return True
    return False


def _is_blank(p) -> bool:
    if paragraph_text(p).strip():
        return False
    if has_picture(p) or next(p.iter(W_FLDCHAR, W_FLDSIMPLE), None) is not None:
        return False
    return True


def analyze_document(document) -> Analysis:
    return Analyzer(document).run()
