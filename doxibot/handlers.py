"""Обработчики Telegram: приём DOCX, отчёт, кнопки исправления и примечаний."""

from __future__ import annotations

import asyncio
import logging
from io import BytesIO
from pathlib import Path

from aiogram import Bot, F, Router
from aiogram.enums import ChatAction
from aiogram.filters import Command, CommandStart
from aiogram.types import (
    BufferedInputFile, CallbackQuery, InlineKeyboardButton, InlineKeyboardMarkup, Message,
)

from . import report, service
from .model import Analysis
from .rules import RULES_SUMMARY_HTML
from .storage import FileStore

log = logging.getLogger(__name__)
router = Router(name="doxibot")

START_TEXT = (
    "👋 Привет! Я <b>DoxiBot</b> — проверяю оформление отчётов в формате <b>.docx</b> "
    "по требованиям ККСО-XX-24.\n\n"
    "📎 Пришлите мне файл отчёта, и я:\n"
    "• найду ошибки оформления и покажу, где они;\n"
    "• 🛠 исправлю автоматически всё, что можно исправить;\n"
    "• 📝 верну копию документа с примечаниями прямо на проблемных местах.\n\n"
    "/rules — требования к оформлению\n"
    "/help — что именно я проверяю"
)

HELP_TEXT = (
    "<b>Как пользоваться</b>\n"
    "Отправьте отчёт файлом (📎 → Файл). Поддерживается только <b>.docx</b>, до 20 МБ.\n\n"
    "<b>Что проверяю</b>\n"
    "• поля, колонтитулы и нумерацию страниц (в т.ч. лишний «enter» в колонтитуле)\n"
    "• шрифт, размер, цвет, интервалы, отступы, выравнивание, полужирный и подчёркивание\n"
    "• заголовки: структурные элементы, главы, подглавы, стили для автособираемого содержания\n"
    "• «СОДЕРЖАНИЕ» или «ОГЛАВЛЕНИЕ» — в зависимости от названий глав\n"
    "• рисунки: обтекание, выравнивание, подписи, нумерацию по главам, ссылки в тексте, «рис.»\n"
    "• таблицы и листинги: подписи, нумерацию, шапку, единый шрифт и интервал\n"
    "• формулы: выравнивание, номер, пустую строку сверху\n"
    "• перечисления: табуляцию после номера («стрелочки»), отступы, маркеры\n"
    "• список источников: дату обращения у электронных ресурсов\n\n"
    "<b>Что меняет автоисправление</b>\n"
    "Только оформление: шрифты, отступы, интервалы, стили, поля, номера страниц, формат подписей "
    "и их нумерацию (ссылки в тексте обновляются). Текст отчёта не переписывается.\n\n"
    "Титульный лист и бланк задания не проверяются — проверка начинается с первого раздела "
    "(АННОТАЦИЯ, СОДЕРЖАНИЕ, ВВЕДЕНИЕ…)."
)

_busy: set[tuple[int, str]] = set()
_semaphore: asyncio.Semaphore | None = None


def setup_workers(workers: int) -> None:
    global _semaphore
    _semaphore = asyncio.Semaphore(max(1, workers))


async def _run(func, *args):
    assert _semaphore is not None
    async with _semaphore:
        return await asyncio.to_thread(func, *args)


def _stem(filename: str) -> str:
    return Path(filename).stem or "document"


def _keyboard(key: str, analysis: Analysis) -> InlineKeyboardMarkup | None:
    if not analysis.issues:
        return None
    rows = []
    fixable = len(analysis.fixable)
    if fixable:
        rows.append([InlineKeyboardButton(text=f"🛠 Исправить автоматически ({fixable})", callback_data=f"fix:{key}")])
    rows.append([
        InlineKeyboardButton(text="📝 Примечания в файле", callback_data=f"notes:{key}"),
        InlineKeyboardButton(text="📋 Полный отчёт", callback_data=f"txt:{key}"),
    ])
    return InlineKeyboardMarkup(inline_keyboard=rows)


# ---------------------------------------------------------------------------- команды


@router.message(CommandStart())
async def cmd_start(message: Message) -> None:
    await message.answer(START_TEXT)


@router.message(Command("help"))
async def cmd_help(message: Message) -> None:
    await message.answer(HELP_TEXT)


@router.message(Command("rules"))
async def cmd_rules(message: Message) -> None:
    await message.answer(RULES_SUMMARY_HTML)


# ---------------------------------------------------------------------------- документы


@router.message(F.document)
async def on_document(message: Message, bot: Bot, store: FileStore, max_file_mb: int) -> None:
    document = message.document
    filename = document.file_name or "document.docx"
    lower = filename.lower()
    if lower.endswith(".doc") or lower.endswith(".rtf") or lower.endswith(".odt"):
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

    status = await message.answer("⏳ Проверяю оформление…")
    await bot.send_chat_action(message.chat.id, ChatAction.TYPING)
    buffer = BytesIO()
    await bot.download(document, destination=buffer)
    data = buffer.getvalue()
    try:
        analysis = await _run(service.check, data)
    except service.BadDocument:
        await status.edit_text("❌ Не получилось открыть файл. Убедитесь, что это настоящий .docx "
                               "(не переименованный .doc) и он не повреждён.")
        return
    except Exception:
        log.exception("Ошибка проверки %s", filename)
        await status.edit_text("❌ Во время проверки произошла ошибка. Попробуйте ещё раз или пришлите другой файл.")
        return

    key = store.put(data, filename, message.from_user.id)
    messages, _ = report.chat_messages(analysis, filename)
    keyboard = _keyboard(key, analysis)
    await status.delete()
    for i, text in enumerate(messages):
        await message.answer(text, reply_markup=keyboard if i == len(messages) - 1 else None)


@router.message(F.photo)
async def on_photo(message: Message) -> None:
    await message.answer("🖼 Это картинка. Пришлите отчёт <b>файлом .docx</b> (📎 → Файл).")


@router.message()
async def on_other(message: Message) -> None:
    await message.answer("📎 Пришлите отчёт в формате <b>.docx</b> — я проверю оформление.\n/help — подробности")


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
async def on_fix(cb: CallbackQuery, bot: Bot, store: FileStore) -> None:
    item, busy_key = await _take(cb, store, "fix")
    if item is None:
        return
    try:
        await cb.answer("🛠 Исправляю…")
        await bot.send_chat_action(cb.message.chat.id, ChatAction.UPLOAD_DOCUMENT)
        result = await _run(service.fix, item.data)
        new_name = f"{_stem(item.filename)}_DoxiBot.docx"
        new_key = store.put(result.data, new_name, cb.from_user.id)
        await cb.message.answer_document(
            BufferedInputFile(result.data, filename=new_name),
            caption="🛠 Исправленный документ",
        )
        await cb.message.answer(
            report.fix_summary(result.applied, result.failed, result.after),
            reply_markup=_keyboard(new_key, result.after),
        )
    except Exception:
        log.exception("Ошибка автоисправления %s", item.filename)
        await cb.message.answer("❌ Не удалось исправить документ. Воспользуйтесь примечаниями и исправьте вручную.")
    finally:
        _busy.discard(busy_key)


@router.callback_query(F.data.startswith("notes:"))
async def on_notes(cb: CallbackQuery, bot: Bot, store: FileStore) -> None:
    item, busy_key = await _take(cb, store, "notes")
    if item is None:
        return
    try:
        await cb.answer("📝 Расставляю примечания…")
        await bot.send_chat_action(cb.message.chat.id, ChatAction.UPLOAD_DOCUMENT)
        data, analysis = await _run(service.annotate, item.data)
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
async def on_text_report(cb: CallbackQuery, store: FileStore) -> None:
    item, busy_key = await _take(cb, store, "txt")
    if item is None:
        return
    try:
        await cb.answer()
        analysis = await _run(service.check, item.data)
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
