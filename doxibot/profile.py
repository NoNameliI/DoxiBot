"""Профиль требований к оформлению: по умолчанию — стандартные требования, но пользователь может менять.

Профиль хранится по user_id в JSON-файле и передаётся в анализатор, автоисправление и тексты подсказок.
"""

from __future__ import annotations

import json
import logging
import threading
from dataclasses import asdict, dataclass, fields, replace
from pathlib import Path

log = logging.getLogger(__name__)

ASK = -1        # спрашивать про титульный лист при каждой загрузке
AUTO = -2       # определять титульный лист автоматически


@dataclass(frozen=True)
class Profile:
    # текст
    font_name: str = "Times New Roman"
    font_size: float = 14.0
    line_spacing: float = 1.5
    indent_cm: float = 1.25
    body_align: str = "both"
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    # страница
    margin_left_mm: float = 30.0
    margin_right_mm: float = 15.0
    margin_top_mm: float = 20.0
    margin_bottom_mm: float = 20.0
    hf_distance_cm: float = 1.25
    # заголовки
    structural_align: str = "center"
    chapter_align: str = "both"
    chapter_indent: bool = True
    chapter_caps: bool = True
    sub_caps: bool = False
    # таблицы
    table_font_min: float = 12.0
    table_line_min: float = 1.0
    # подписи
    caption_dash: str = "–"
    # что проверять
    check_page: bool = True
    check_page_numbers: bool = True
    check_headings: bool = True
    check_toc: bool = True
    check_figures: bool = True
    check_tables: bool = True
    check_formulas: bool = True
    check_lists: bool = True
    check_listings: bool = True
    check_sources: bool = True
    # титульный лист: ASK / AUTO / число страниц
    title_pages: int = ASK

    # --- производные значения ---
    @property
    def indent_twips(self) -> int:
        return round(self.indent_cm * 566.929)

    @property
    def font_size_half(self) -> int:
        return round(self.font_size * 2)

    @property
    def table_font_max(self) -> float:
        return max(self.font_size, self.table_font_min)

    @property
    def table_line_max(self) -> float:
        return max(self.line_spacing, self.table_line_min)

    @property
    def placeholders(self) -> dict[str, str]:
        return {
            "font": self.font_name,
            "size": _num(self.font_size),
            "line": _num(self.line_spacing),
            "indent": _num(self.indent_cm),
            "ml": _num(self.margin_left_mm),
            "mr": _num(self.margin_right_mm),
            "mt": _num(self.margin_top_mm),
            "mb": _num(self.margin_bottom_mm),
            "hf": _num(self.hf_distance_cm),
            "tfont": _num(self.table_font_min),
            "tline": _num(self.table_line_min),
            "dash": self.caption_dash,
            "before": _num(self.space_before_pt),
            "after": _num(self.space_after_pt),
            "body_spacing": f"{_num(self.space_before_pt)}/{_num(self.space_after_pt)} пт",
            "body_align": ALIGN_RU[self.body_align],
            "structural_align": ALIGN_RU[self.structural_align],
            "chapter_align": ALIGN_RU[self.chapter_align],
            "chapter_indent": _num(self.indent_cm) + " см" if self.chapter_indent else "без отступа",
        }

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "Profile":
        known = {f.name: f.type for f in fields(cls)}
        clean = {}
        for key, value in (data or {}).items():
            if key not in known:
                continue
            try:
                if known[key] == "float":
                    clean[key] = float(value)
                elif known[key] == "int":
                    clean[key] = int(value)
                elif known[key] == "bool":
                    clean[key] = bool(value)
                else:
                    clean[key] = str(value)
            except (TypeError, ValueError):
                continue
        return cls(**clean)

    def with_value(self, key: str, value) -> "Profile":
        return replace(self, **{key: value})


DEFAULT = Profile()
ALIGN_RU = {"both": "по ширине", "center": "по центру", "left": "по левому краю", "right": "по правому краю"}


def _num(value: float) -> str:
    return f"{value:g}".replace(".", ",")


# --------------------------------------------------------------------------- описание полей для меню


@dataclass(frozen=True)
class Field:
    key: str
    label: str
    kind: str                       # float | str | bool | choice
    group: str
    unit: str = ""
    choices: tuple[tuple[str, str], ...] = ()      # (значение, подпись)
    low: float = 0.0
    high: float = 100.0
    prompt: str = ""

    def show(self, profile: Profile) -> str:
        value = getattr(profile, self.key)
        if self.kind == "bool":
            return "да" if value else "нет"
        if self.kind == "choice":
            return dict(self.choices).get(str(value), str(value))
        if self.kind == "float":
            return f"{_num(value)}{(' ' + self.unit) if self.unit else ''}"
        return str(value)


GROUPS = {
    "text": "🔤 Текст",
    "page": "📐 Страница",
    "head": "🔠 Заголовки",
    "table": "🧮 Таблицы и подписи",
    "checks": "🧩 Что проверять",
    "title": "📑 Титульный лист",
}

FIELDS: tuple[Field, ...] = (
    Field("font_name", "Шрифт", "str", "text", prompt="Введите название шрифта, например: Times New Roman"),
    Field("font_size", "Размер шрифта", "float", "text", "пт", low=6, high=36,
          prompt="Введите размер шрифта в пунктах, например: 14"),
    Field("line_spacing", "Межстрочный интервал", "float", "text", low=0.8, high=3,
          prompt="Введите межстрочный интервал, например: 1,5"),
    Field("indent_cm", "Абзацный отступ", "float", "text", "см", low=0, high=5,
          prompt="Введите абзацный отступ в сантиметрах, например: 1,25"),
    Field("body_align", "Выравнивание текста", "choice", "text",
          choices=(("both", "по ширине"), ("left", "по левому краю"))),
    Field("space_before_pt", "Интервал перед абзацем", "float", "text", "пт", low=0, high=48,
          prompt="Введите интервал перед абзацем в пунктах, например: 0"),
    Field("space_after_pt", "Интервал после абзаца", "float", "text", "пт", low=0, high=48,
          prompt="Введите интервал после абзаца в пунктах, например: 0"),

    Field("margin_left_mm", "Левое поле", "float", "page", "мм", low=5, high=60, prompt="Левое поле в мм, например: 30"),
    Field("margin_right_mm", "Правое поле", "float", "page", "мм", low=5, high=60, prompt="Правое поле в мм, например: 15"),
    Field("margin_top_mm", "Верхнее поле", "float", "page", "мм", low=5, high=60, prompt="Верхнее поле в мм, например: 20"),
    Field("margin_bottom_mm", "Нижнее поле", "float", "page", "мм", low=5, high=60, prompt="Нижнее поле в мм, например: 20"),
    Field("hf_distance_cm", "Колонтитулы от края", "float", "page", "см", low=0, high=5,
          prompt="Расстояние до колонтитула в см, например: 1,25"),
    Field("check_page_numbers", "Проверять нумерацию страниц", "bool", "page"),

    Field("structural_align", "ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ и т.п.", "choice", "head",
          choices=(("center", "по центру"), ("both", "по ширине"), ("left", "по левому краю"))),
    Field("chapter_align", "Главы", "choice", "head",
          choices=(("both", "по ширине"), ("center", "по центру"), ("left", "по левому краю"))),
    Field("chapter_indent", "Отступ у глав и подглав", "bool", "head"),
    Field("chapter_caps", "Главы КАПСОМ", "bool", "head"),
    Field("sub_caps", "Подглавы КАПСОМ", "bool", "head"),

    Field("table_font_min", "Минимальный шрифт в таблицах", "float", "table", "пт", low=6, high=20,
          prompt="Минимальный размер шрифта в таблицах, например: 12"),
    Field("table_line_min", "Минимальный интервал в таблицах", "float", "table", low=0.8, high=3,
          prompt="Минимальный межстрочный интервал в таблицах, например: 1"),
    Field("caption_dash", "Знак в подписях", "choice", "table",
          choices=(("–", "– (среднее тире)"), ("—", "— (длинное тире)"))),

    Field("check_headings", "Заголовки", "bool", "checks"),
    Field("check_toc", "Содержание/оглавление", "bool", "checks"),
    Field("check_figures", "Рисунки и подписи", "bool", "checks"),
    Field("check_tables", "Таблицы", "bool", "checks"),
    Field("check_formulas", "Формулы", "bool", "checks"),
    Field("check_lists", "Перечисления", "bool", "checks"),
    Field("check_listings", "Листинги кода", "bool", "checks"),
    Field("check_sources", "Список источников", "bool", "checks"),
    Field("check_page", "Поля и колонтитулы", "bool", "checks"),

    Field("title_pages", "Титульный лист", "choice", "title",
          choices=((str(ASK), "спрашивать каждый раз"), (str(AUTO), "определять автоматически"),
                   ("0", "титульного листа нет"), ("1", "1 страница"), ("2", "2 страницы"), ("3", "3 страницы"))),
)

FIELD_BY_KEY = {f.key: f for f in FIELDS}


def parse_value(field: Field, text: str) -> float | str:
    """Разбирает введённое пользователем значение; бросает ValueError с понятным текстом."""
    text = text.strip()
    if field.kind == "float":
        try:
            value = float(text.replace(",", ".").replace(" ", ""))
        except ValueError:
            raise ValueError("Нужно число, например 1,25") from None
        if not field.low <= value <= field.high:
            raise ValueError(f"Значение должно быть от {_num(field.low)} до {_num(field.high)}")
        return value
    if not text or len(text) > 64:
        raise ValueError("Слишком длинное значение")
    return text


# --------------------------------------------------------------------------- хранилище


class SettingsStore:
    """Профили пользователей в одном JSON-файле."""

    def __init__(self, path: Path):
        self.path = Path(path)
        self._lock = threading.Lock()
        self._data: dict[str, dict] = {}
        self._load()

    def _load(self) -> None:
        try:
            self._data = json.loads(self.path.read_text(encoding="utf-8"))
        except FileNotFoundError:
            self._data = {}
        except Exception:
            log.exception("Не удалось прочитать настройки, используются значения по умолчанию")
            self._data = {}

    def _save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self._data, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp.replace(self.path)

    def get(self, user_id: int) -> Profile:
        return Profile.from_dict(self._data.get(str(user_id), {}))

    def set(self, user_id: int, profile: Profile) -> None:
        with self._lock:
            diff = {k: v for k, v in profile.to_dict().items() if v != getattr(DEFAULT, k)}
            if diff:
                self._data[str(user_id)] = diff
            else:
                self._data.pop(str(user_id), None)
            self._save()

    def update(self, user_id: int, key: str, value) -> Profile:
        profile = self.get(user_id).with_value(key, value)
        self.set(user_id, profile)
        return profile

    def reset(self, user_id: int) -> Profile:
        self.set(user_id, DEFAULT)
        return DEFAULT

    def is_custom(self, user_id: int) -> bool:
        return bool(self._data.get(str(user_id)))
