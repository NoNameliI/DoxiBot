"""DoxiBot — Telegram-бот проверки оформления DOCX по требованиям ККСО-XX-24.

Запуск:  python bot.py   (токен берётся из .env → BOT_TOKEN)
"""

from __future__ import annotations

import asyncio
import logging

from aiogram import Bot, Dispatcher
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.exceptions import TelegramAPIError
from aiogram.types import BotCommand

from doxibot.config import load_config
from doxibot.handlers import router, setup_workers
from doxibot.storage import FileStore

BOT_NAME = "DoxiBot"
DESCRIPTION = (
    "DoxiBot проверяет оформление отчётов .docx по требованиям ККСО-XX-24: поля, шрифты, интервалы, "
    "заголовки, рисунки, таблицы, формулы, перечисления и нумерацию страниц. Показывает ошибки, "
    "исправляет их автоматически и расставляет примечания в документе.\n\nПросто пришлите файл .docx."
)
SHORT_DESCRIPTION = "Проверка и автоисправление оформления отчётов .docx"


async def on_startup(bot: Bot) -> None:
    try:
        await bot.set_my_commands([
            BotCommand(command="start", description="Начать"),
            BotCommand(command="rules", description="Требования к оформлению"),
            BotCommand(command="help", description="Что проверяет бот"),
        ])
        if (await bot.get_my_name()).name != BOT_NAME:
            await bot.set_my_name(BOT_NAME)
        if (await bot.get_my_description()).description != DESCRIPTION:
            await bot.set_my_description(DESCRIPTION)
        if (await bot.get_my_short_description()).short_description != SHORT_DESCRIPTION:
            await bot.set_my_short_description(SHORT_DESCRIPTION)
    except TelegramAPIError as exc:  # например, лимит частоты на смену имени — не критично
        logging.warning("Не удалось обновить профиль бота: %s", exc)
    me = await bot.get_me()
    logging.info("%s запущен: https://t.me/%s", BOT_NAME, me.username)


async def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    config = load_config()
    setup_workers(config.workers)
    bot = Bot(config.token, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher(store=FileStore(config.store_ttl_minutes), max_file_mb=config.max_file_mb)
    dp.include_router(router)
    dp.startup.register(on_startup)
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit) as exc:
        if isinstance(exc, SystemExit) and exc.code not in (None, 0):
            print(exc.code)
