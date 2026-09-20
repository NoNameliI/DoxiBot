"""Обработчики Telegram: приём DOCX, вопрос о титульном листе, отчёт, исправление, настройки."""

from __future__ import annotations

import asyncio
import logging
from io import BytesIO
from pathlib import Path

from aiogram import Bot, F, Router
from aiogram.enums import ChatAction
from aiogram.filters import Command, CommandStart, StateFilter
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    BufferedInputFile, CallbackQuery, InlineKeyboardButton, InlineKeyboardMarkup, Message,
)

from . import report, service
from .model import Analysis
from .profile import ASK, AUTO, FIELD_BY_KEY, FIELDS, GROUPS, Profile, SettingsStore, parse_value
from .rules import rules_summary_html
from .storage import FileStore

log = logging.getLogger(__name__)
router = Router(name="doxibot")

START_TEXT = (
    "👋 Привет! Я <b>DoxiBot</b> — проверяю оформление отчётов в формате <b>.docx</b> "
    "по стандартным требованиям.\n\n"
    "📎 Пришлите мне файл отчёта, и я:\n"
    "• найду ошибки оформления и покажу, где они;\n"
    "• 🛠 исправлю автоматически всё, что можно исправить;\n"
    "• 📝 верну копию документа с примечаниями прямо на проблемных местах.\n\n"
    "/settings — свои требования к оформлению\n"
    "/rules — по каким правилам проверяю\n"
    "/help — что именно я проверяю"
)

HELP_TEXT = (
    "<b>Как пользоваться</b>\n"
    "Отправьте отчёт файлом (📎 → Файл). Поддерживается только <b>.docx</b>, до 20 МБ.\n"
    "Перед проверкой я спрошу, сколько первых страниц занимает титульный лист и задание — "
    "их я не проверяю и не исправляю (ответ можно закрепить в /settings).\n\n"
    "<b>Что проверяю</b>\n"
    "• поля, колонтитулы и нумерацию страниц (в т.ч. лишний «enter» в колонтитуле)\n"
    "• шрифт, размер, цвет, интервалы, отступы, выравнивание, полужирный и подчёркивание\n"
    "• заголовки: структурные элементы, главы, подглавы, стили для автособираемого содержания\n"
    "• «СОДЕРЖАНИЕ» или «ОГЛАВЛЕНИЕ» — в зависимости от названий глав\n"
    "• рисунки: обтекание, выравнивание, подписи, нумерацию по главам, ссылки в тексте, «рис.»\n"
    "• таблицы и листинги: подписи, нумерацию, шапку, единый шрифт и интервал\n"
    "• формулы: выравнивание, номер, пустую строку сверху\n"
    "• перечисления: табуляцию после номера («стрелочки»), отступы, маркеры и формат самих номеров\n"
    "• список источников: дату обращения у электронных ресурсов\n\n"
    "<b>Что меняет автоисправление</b>\n"
    "Только оформление: шрифты, отступы, интервалы, стили, поля, номера страниц, формат подписей "
    "и их нумерацию (ссылки в тексте обновляются). Текст отчёта не переписывается.\n\n"
    "<b>Свои требования</b>\n"
    "/settings — шрифт, размер, интервалы, поля, выравнивание заголовков и набор проверок. "
    "По умолчанию стоят стандартные требования."
)

TITLE_QUESTION = (
    "📑 <b>Титульный лист</b>\n"
    "Сколько первых страниц файла занимают титульный лист и бланк задания? "
    "Я не буду их проверять и исправлять."
)

_busy: set[tuple[int, str]] = set()
_semaphore: asyncio.Semaphore | None = None


class SettingInput(StatesGroup):
    value = State()


def setup_workers(workers: int) -> None:
    global _semaphore
    _semaphore = asyncio.Semaphore(max(1, workers))


async def _run(func, *args):
    assert _semaphore is not None
    async with _semaphore:
        return await asyncio.to_thread(func, *args)


def _stem(filename: str) -> str:
    return Path(filename).stem or "document"


def _result_keyboard(key: str, analysis: Analysis, offer_fix: bool = True) -> InlineKeyboardMarkup | None:
    if not analysis.issues:
        return None
    rows = []
    fixable = len(analysis.fixable)
    if offer_fix and fixable:
        rows.append([InlineKeyboardButton(text=f"🛠 Исправить автоматически ({fixable})", callback_data=f"fix:{key}")])
    rows.append([
        InlineKeyboardButton(text="📝 Примечания в файле", callback_data=f"notes:{key}"),
        InlineKeyboardButton(text="📋 Полный отчёт", callback_data=f"txt:{key}"),
    ])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def _title_keyboard(key: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Нет титульного листа", callback_data=f"tp:{key}:0")],
        [
            InlineKeyboardButton(text="1 страница", callback_data=f"tp:{key}:1"),
            InlineKeyboardButton(text="2 страницы", callback_data=f"tp:{key}:2"),
            InlineKeyboardButton(text="3 страницы", callback_data=f"tp:{key}:3"),
        ],
        [InlineKeyboardButton(text="🤖 Определить автоматически", callback_data=f"tp:{key}:auto")],
    ])


# ---------------------------------------------------------------------------- команды


@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer(START_TEXT)


@router.message(Command("help"))
async def cmd_help(message: Message) -> None:
    await message.answer(HELP_TEXT)


@router.message(Command("rules"))
async def cmd_rules(message: Message, settings: SettingsStore) -> None:
    await message.answer(rules_summary_html(settings.get(message.from_user.id)))


@router.message(Command("cancel"))
async def cmd_cancel(message: Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer("Отменено.")


# ---------------------------------------------------------------------------- настройки


def _settings_text(profile: Profile, group: str | None) -> str:
    if group is None:
        head = ["⚙️ <b>Требования к оформлению</b>",
                "<i>По умолчанию — стандартные требования. Выберите раздел, чтобы изменить значения.</i>", ""]
        head.append(f"🔤 {profile.font_name}, {profile.placeholders['size']} пт, "
                    f"интервал {profile.placeholders['line']}, отступ {profile.placeholders['indent']} см")
        head.append(f"📐 Поля {profile.placeholders['ml']}/{profile.placeholders['mr']}/"
                    f"{profile.placeholders['mt']}/{profile.placeholders['mb']} мм")
        off = [f.label for f in FIELDS if f.group == "checks" and not getattr(profile, f.key)]
        head.append("🧩 Отключены проверки: " + (", ".join(off) if off else "нет"))
        return "\n".join(head)
    lines = [f"⚙️ <b>{GROUPS[group]}</b>", ""]
    for field in FIELDS:
        if field.group == group:
            lines.append(f"• {field.label}: <b>{field.show(profile)}</b>")
    lines.append("\n<i>Нажмите на параметр, чтобы изменить.</i>")
    return "\n".join(lines)


def _settings_keyboard(profile: Profile, group: str | None) -> InlineKeyboardMarkup:
    rows: list[list[InlineKeyboardButton]] = []
    if group is None:
        buttons = [InlineKeyboardButton(text=title, callback_data=f"set:g:{key}") for key, title in GROUPS.items()]
        rows = [buttons[i:i + 2] for i in range(0, len(buttons), 2)]
        rows.append([InlineKeyboardButton(text="♻️ Вернуть стандартные требования", callback_data="set:reset")])
        rows.append([InlineKeyboardButton(text="✅ Готово", callback_data="set:close")])
        return InlineKeyboardMarkup(inline_keyboard=rows)
    for field in FIELDS:
        if field.group != group:
            continue
        mark = "✅" if field.kind == "bool" and getattr(profile, field.key) else ("❌" if field.kind == "bool" else "")
        text = f"{mark} {field.label}" if mark else f"{field.label}: {field.show(profile)}"
        rows.append([InlineKeyboardButton(text=text, callback_data=f"set:f:{field.key}")])
    rows.append([InlineKeyboardButton(text="⬅️ Назад", callback_data="set:main")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def _choice_keyboard(field, profile: Profile) -> InlineKeyboardMarkup:
    current = str(getattr(profile, field.key))
    rows = [[InlineKeyboardButton(text=("• " if value == current else "") + label,
                                  callback_data=f"set:v:{field.key}:{n}")]
            for n, (value, label) in enumerate(field.choices)]
    rows.append([InlineKeyboardButton(text="⬅️ Назад", callback_data=f"set:g:{field.group}")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


@router.message(Command("settings"))
async def cmd_settings(message: Message, settings: SettingsStore, state: FSMContext) -> None:
    await state.clear()
    profile = settings.get(message.from_user.id)
    await message.answer(_settings_text(profile, None), reply_markup=_settings_keyboard(profile, None))


@router.callback_query(F.data.startswith("set:"))
async def on_settings(cb: CallbackQuery, settings: SettingsStore, state: FSMContext) -> None:
    parts = cb.data.split(":")
    action = parts[1]
    user_id = cb.from_user.id
    profile = settings.get(user_id)

    if action == "close":
        await state.clear()
        await cb.message.edit_text("⚙️ Настройки сохранены. Присылайте файл — проверю по ним.")
        await cb.answer()
        return
    if action == "reset":
        profile = settings.reset(user_id)
        await cb.answer("Возвращены стандартные требования")
    elif action == "main":
        await state.clear()
    elif action == "g":
        await state.clear()
        await cb.message.edit_text(_settings_text(profile, parts[2]),
                                   reply_markup=_settings_keyboard(profile, parts[2]))
        await cb.answer()
        return
    elif action == "f":
        field = FIELD_BY_KEY[parts[2]]
        if field.kind == "bool":
            profile = settings.update(user_id, field.key, not getattr(profile, field.key))
            await cb.message.edit_text(_settings_text(profile, field.group),
                                       reply_markup=_settings_keyboard(profile, field.group))
            await cb.answer()
            return
        if field.kind == "choice":
            await cb.message.edit_text(f"⚙️ <b>{field.label}</b>\nТекущее значение: <b>{field.show(profile)}</b>",
                                       reply_markup=_choice_keyboard(field, profile))
            await cb.answer()
            return
        await state.set_state(SettingInput.value)
        await state.update_data(key=field.key, message_id=cb.message.message_id)
        await cb.message.edit_text(
            f"⚙️ <b>{field.label}</b>\nТекущее значение: <b>{field.show(profile)}</b>\n\n"
            f"{field.prompt}\nОтправьте новое значение сообщением или /cancel."
        )
        await cb.answer()
        return
    elif action == "v":
        field = FIELD_BY_KEY[parts[2]]
        value = field.choices[int(parts[3])][0]
        if field.kind == "choice" and field.key == "title_pages":
            value = int(value)
        profile = settings.update(user_id, field.key, value)
        await cb.message.edit_text(_settings_text(profile, field.group),
                                   reply_markup=_settings_keyboard(profile, field.group))
        await cb.answer("Сохранено")
        return

    await cb.message.edit_text(_settings_text(profile, None), reply_markup=_settings_keyboard(profile, None))
    if action != "reset":
        await cb.answer()


@router.message(SettingInput.value, F.text)
async def on_setting_value(message: Message, settings: SettingsStore, state: FSMContext) -> None:
    data = await state.get_data()
    field = FIELD_BY_KEY[data["key"]]
    try:
        value = parse_value(field, message.text)
    except ValueError as exc:
        await message.answer(f"⚠️ {exc}\nПопробуйте ещё раз или отправьте /cancel.")
        return
    profile = settings.update(message.from_user.id, field.key, value)
    await state.clear()
    await message.answer(_settings_text(profile, field.group),
                         reply_markup=_settings_keyboard(profile, field.group))


# ---------------------------------------------------------------------------- документы


@router.message(F.document)
async def on_document(message: Message, bot: Bot, store: FileStore, settings: SettingsStore,
                      state: FSMContext, max_file_mb: int) -> None:
    await state.clear()
    document = message.document
    filename = document.file_name or "document.docx"
    lower = filename.lower()
    if lower.endswith((".doc", ".rtf", ".odt")):
        await message.answer(
            "⚠️ Этот формат я не читаю. Откройте файл в Word и сохраните как <b>.docx</b> "
            "(Файл → Сохранить как → «Документ Word (*.docx)»), затем пришлите снова."
        )
        return
    if not lower.endswith(".docx"):
        await message.answer("📎 Я проверяю только документы Word в формате <b>.docx</b>.")
        return
    if document.file_size and document.file_size > max_file_mb * 1024 * 1024:
        await message.answer(f"⚠️ Файл больше {max_file_mb} МБ — Telegram не позволяет ботам скачивать такие файлы. "
                             "Сожмите рисунки (Файл → Сжать рисунки) и пришлите снова.")
        return

    await bot.send_chat_action(message.chat.id, ChatAction.TYPING)
    buffer = BytesIO()
    await bot.download(document, destination=buffer)
    key = store.put(buffer.getvalue(), filename, message.from_user.id)
    profile = settings.get(message.from_user.id)

    if profile.title_pages == ASK:
        await message.answer(TITLE_QUESTION, reply_markup=_title_keyboard(key))
        return
    store.set_title_pages(key, None if profile.title_pages == AUTO else profile.title_pages)
    await _check_and_report(message, key, store, profile)


@router.callback_query(F.data.startswith("tp:"))
async def on_title_pages(cb: CallbackQuery, store: FileStore, settings: SettingsStore) -> None:
    _, key, value = cb.data.split(":")
    item = store.get(key, cb.from_user.id)
    if item is None:
        await cb.answer("Файл уже не хранится — пришлите документ ещё раз.", show_alert=True)
        return
    store.set_title_pages(key, None if value == "auto" else int(value))
    label = {"0": "без титульного листа", "auto": "определяю сам"}.get(value, f"пропускаю первые {value} стр.")
    await cb.answer()
    await cb.message.edit_text(f"📑 Титульный лист: {label}\n⏳ Проверяю оформление…", reply_markup=None)
    await _check_and_report(cb.message, key, store, settings.get(cb.from_user.id), status=cb.message)


async def _check_and_report(target: Message, key: str, store: FileStore, profile: Profile,
                            status: Message | None = None) -> None:
    item = store.get_any(key)
    if item is None:
        await target.answer("Файл уже не хранится — пришлите документ ещё раз.")
        return
    if status is None:
        status = await target.answer("⏳ Проверяю оформление…")
    try:
        analysis = await _run(service.check, item.data, profile, item.title_pages)
    except service.BadDocument:
        await status.edit_text("❌ Не получилось открыть файл. Убедитесь, что это настоящий .docx "
                               "(не переименованный .doc) и он не повреждён.")
        return
    except Exception:
        log.exception("Ошибка проверки %s", item.filename)
        await status.edit_text("❌ Во время проверки произошла ошибка. Попробуйте ещё раз или пришлите другой файл.")
        return

    messages, _ = report.chat_messages(analysis, item.filename)
    keyboard = _result_keyboard(key, analysis)
    await status.delete()
    for i, text in enumerate(messages):
        await target.answer(text, reply_markup=keyboard if i == len(messages) - 1 else None)


@router.message(F.photo)
async def on_photo(message: Message) -> None:
    await message.answer("🖼 Это картинка. Пришлите отчёт <b>файлом .docx</b> (📎 → Файл).")


@router.message(StateFilter(None))
async def on_other(message: Message) -> None:
    await message.answer("📎 Пришлите отчёт в формате <b>.docx</b> — я проверю оформление.\n"
                         "/settings — свои требования · /help — подробности")


# ---------------------------------------------------------------------------- кнопки


async def _take(cb: CallbackQuery, store: FileStore, action: str):
    key = cb.data.split(":", 1)[1]
    item = store.get(key, cb.from_user.id)
    if item is None:
        await cb.answer("Файл уже не хранится — пришлите документ ещё раз.", show_alert=True)
        return None, None
    busy_key = (cb.from_user.id, f"{action}:{key}")
    if busy_key in _busy:
        await cb.answer("Уже обрабатываю, подождите…")
        return None, None
    _busy.add(busy_key)
    return item, busy_key


@router.callback_query(F.data.startswith("fix:"))
async def on_fix(cb: CallbackQuery, bot: Bot, store: FileStore, settings: SettingsStore) -> None:
    item, busy_key = await _take(cb, store, "fix")
    if item is None:
        return
    try:
        await cb.answer("🛠 Исправляю…")
        await bot.send_chat_action(cb.message.chat.id, ChatAction.UPLOAD_DOCUMENT)
        profile = settings.get(cb.from_user.id)
        result = await _run(service.fix, item.data, profile, item.title_pages)
        new_name = f"{_stem(item.filename)}_DoxiBot.docx"
        new_key = store.put(result.data, new_name, cb.from_user.id)
        store.set_title_pages(new_key, item.title_pages)
        await cb.message.answer_document(
            BufferedInputFile(result.data, filename=new_name),
            caption="🛠 Исправленный документ",
        )
        await cb.message.answer(
            report.fix_summary(result.applied, result.failed, result.after),
            reply_markup=_result_keyboard(new_key, result.after, offer_fix=False),
        )
    except Exception:
        log.exception("Ошибка автоисправления %s", item.filename)
        await cb.message.answer("❌ Не удалось исправить документ. Воспользуйтесь примечаниями и исправьте вручную.")
    finally:
        _busy.discard(busy_key)


@router.callback_query(F.data.startswith("notes:"))
async def on_notes(cb: CallbackQuery, bot: Bot, store: FileStore, settings: SettingsStore) -> None:
    item, busy_key = await _take(cb, store, "notes")
    if item is None:
        return
    try:
        await cb.answer("📝 Расставляю примечания…")
        await bot.send_chat_action(cb.message.chat.id, ChatAction.UPLOAD_DOCUMENT)
        data, analysis = await _run(service.annotate, item.data, settings.get(cb.from_user.id), item.title_pages)
        await cb.message.answer_document(
            BufferedInputFile(data, filename=f"{_stem(item.filename)}_примечания.docx"),
            caption=f"📝 Примечания DoxiBot: {len(analysis.issues)}. Откройте в Word → вкладка «Рецензирование».",
        )
    except Exception:
        log.exception("Ошибка расстановки примечаний %s", item.filename)
        await cb.message.answer("❌ Не удалось добавить примечания.")
    finally:
        _busy.discard(busy_key)


@router.callback_query(F.data.startswith("txt:"))
async def on_text_report(cb: CallbackQuery, store: FileStore, settings: SettingsStore) -> None:
    item, busy_key = await _take(cb, store, "txt")
    if item is None:
        return
    try:
        await cb.answer()
        analysis = await _run(service.check, item.data, settings.get(cb.from_user.id), item.title_pages)
        text = report.full_text_report(analysis, item.filename)
        await cb.message.answer_document(
            BufferedInputFile(text.encode("utf-8-sig"), filename=f"{_stem(item.filename)}_отчёт.txt"),
            caption="📋 Полный список замечаний",
        )
    except Exception:
        log.exception("Ошибка формирования отчёта %s", item.filename)
        await cb.message.answer("❌ Не удалось сформировать отчёт.")
    finally:
        _busy.discard(busy_key)
