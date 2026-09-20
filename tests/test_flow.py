"""Сценарии бота на заглушках Telegram: загрузка файла, вопрос о титульном листе, отчёт, исправление, настройки."""

from __future__ import annotations

import asyncio
import sys
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from doxibot import handlers as H  # noqa: E402
from doxibot.profile import ASK, AUTO, Profile, SettingsStore  # noqa: E402
from doxibot.storage import FileStore  # noqa: E402
from tests.make_samples import build_bad, build_good  # noqa: E402

USER = 777


@dataclass
class FakeUser:
    id: int = USER


@dataclass
class FakeChat:
    id: int = 1


@dataclass
class FakeDocument:
    file_name: str
    file_size: int


@dataclass
class FakeMessage:
    text: str | None = None
    document: FakeDocument | None = None
    from_user: FakeUser = field(default_factory=FakeUser)
    chat: FakeChat = field(default_factory=FakeChat)
    message_id: int = 1
    sent: list = field(default_factory=list)
    documents: list = field(default_factory=list)
    edits: list = field(default_factory=list)
    deleted: bool = False

    async def answer(self, text, reply_markup=None, **kw):
        child = FakeMessage(text=text, sent=self.sent, documents=self.documents, edits=self.edits)
        self.sent.append((text, reply_markup))
        return child

    async def answer_document(self, document, caption=None, **kw):
        self.documents.append((document, caption))
        return self

    async def edit_text(self, text, reply_markup=None, **kw):
        self.edits.append((text, reply_markup))
        self.text = text
        return self

    async def delete(self):
        self.deleted = True


@dataclass
class FakeCallback:
    data: str
    message: FakeMessage
    from_user: FakeUser = field(default_factory=FakeUser)
    answers: list = field(default_factory=list)

    async def answer(self, text=None, show_alert=False):
        self.answers.append(text)


class FakeBot:
    def __init__(self, payload: bytes):
        self.payload = payload

    async def send_chat_action(self, *a, **kw):
        pass

    async def download(self, document, destination):
        destination.write(self.payload)


class FakeState:
    def __init__(self):
        self.state = None
        self.data = {}

    async def set_state(self, state):
        self.state = state

    async def update_data(self, **kw):
        self.data.update(kw)

    async def get_data(self):
        return dict(self.data)

    async def clear(self):
        self.state, self.data = None, {}


def buttons(markup) -> list[str]:
    return [b.callback_data for row in markup.inline_keyboard for b in row]


def texts(markup) -> list[str]:
    return [b.text for row in markup.inline_keyboard for b in row]


@pytest.fixture(scope="module")
def sample() -> bytes:
    buf = BytesIO()
    build_bad().save(buf)
    return buf.getvalue()


@pytest.fixture
def env(tmp_path, sample):
    H.setup_workers(1)
    return {
        "bot": FakeBot(sample),
        "store": FileStore(),
        "settings": SettingsStore(tmp_path / "settings.json"),
        "state": FakeState(),
    }


def run(coro):
    return asyncio.run(coro)


def test_document_asks_about_title_pages(env):
    message = FakeMessage(document=FakeDocument("отчёт.docx", 5000))
    run(H.on_document(message, env["bot"], env["store"], env["settings"], env["state"], 20))
    text, markup = message.sent[-1]
    assert "Титульный лист" in text
    assert [b.split(":")[-1] for b in buttons(markup)] == ["0", "1", "2", "3", "auto"]


def test_title_answer_runs_check_and_offers_fix(env):
    message = FakeMessage(document=FakeDocument("отчёт.docx", 5000))
    run(H.on_document(message, env["bot"], env["store"], env["settings"], env["state"], 20))
    key = buttons(message.sent[-1][1])[0].split(":")[1]

    cb = FakeCallback(data=f"tp:{key}:1", message=FakeMessage(sent=message.sent, edits=message.edits))
    run(H.on_title_pages(cb, env["store"], env["settings"]))
    assert env["store"].get(key, USER).title_pages == 1
    assert "пропускаю первые 1" in cb.message.edits[0][0]

    assert any("Ошибок" in text for text, _ in message.sent)
    markup = [m for _, m in message.sent if m is not None][-1]
    assert any(b.startswith("fix:") for b in buttons(markup))


def test_auto_and_saved_choice(env):
    env["settings"].update(USER, "title_pages", AUTO)
    message = FakeMessage(document=FakeDocument("отчёт.docx", 5000))
    run(H.on_document(message, env["bot"], env["store"], env["settings"], env["state"], 20))
    assert all("Титульный лист" not in text for text, _ in message.sent)
    assert any("Ошибок" in text for text, _ in message.sent)


def test_fix_does_not_offer_fix_again(env):
    message = FakeMessage(document=FakeDocument("отчёт.docx", 5000))
    run(H.on_document(message, env["bot"], env["store"], env["settings"], env["state"], 20))
    key = buttons(message.sent[-1][1])[0].split(":")[1]
    run(H.on_title_pages(FakeCallback(f"tp:{key}:0", FakeMessage(sent=message.sent)), env["store"], env["settings"]))

    cb = FakeCallback(data=f"fix:{key}", message=FakeMessage(sent=[], documents=[]))
    run(H.on_fix(cb, env["bot"], env["store"], env["settings"]))
    assert cb.message.documents, "исправленный файл не отправлен"
    summary, markup = cb.message.sent[-1]
    assert "Исправлено замечаний" in summary
    assert not any(b.startswith("fix:") for b in buttons(markup)), "повторное исправление не должно предлагаться"


def test_settings_toggle_and_choice(env):
    message = FakeMessage(text="/settings")
    run(H.cmd_settings(message, env["settings"], env["state"]))
    text, markup = message.sent[-1]
    assert "Требования к оформлению" in text
    assert "set:g:text" in buttons(markup)

    cb = FakeCallback(data="set:g:checks", message=FakeMessage())
    run(H.on_settings(cb, env["settings"], env["state"]))
    assert "set:f:check_figures" in buttons(cb.message.edits[-1][1])

    cb = FakeCallback(data="set:f:check_figures", message=FakeMessage())
    run(H.on_settings(cb, env["settings"], env["state"]))
    assert env["settings"].get(USER).check_figures is False
    assert "❌ Рисунки и подписи" in texts(cb.message.edits[-1][1])

    cb = FakeCallback(data="set:f:body_align", message=FakeMessage())
    run(H.on_settings(cb, env["settings"], env["state"]))
    assert "set:v:body_align:1" in buttons(cb.message.edits[-1][1])
    cb = FakeCallback(data="set:v:body_align:1", message=FakeMessage())
    run(H.on_settings(cb, env["settings"], env["state"]))
    assert env["settings"].get(USER).body_align == "left"


def test_settings_custom_value_and_reset(env):
    cb = FakeCallback(data="set:f:font_size", message=FakeMessage())
    run(H.on_settings(cb, env["settings"], env["state"]))
    assert env["state"].state is H.SettingInput.value

    bad = FakeMessage(text="много")
    run(H.on_setting_value(bad, env["settings"], env["state"]))
    assert "Нужно число" in bad.sent[-1][0]

    good = FakeMessage(text="12,5")
    run(H.on_setting_value(good, env["settings"], env["state"]))
    assert env["settings"].get(USER).font_size == 12.5
    assert env["state"].state is None

    cb = FakeCallback(data="set:reset", message=FakeMessage())
    run(H.on_settings(cb, env["settings"], env["state"]))
    assert env["settings"].get(USER) == Profile()


def test_profile_changes_what_is_checked(env, sample):
    from doxibot import service

    strict = service.check(sample, Profile(title_pages=0))
    relaxed = service.check(sample, Profile(title_pages=0, check_figures=False, check_tables=False))
    assert len(relaxed.issues) < len(strict.issues)
    assert not [i for i in relaxed.issues if i.code.startswith(("fig_", "tbl_", "picture_", "table_"))]

    buf = BytesIO()
    build_good().save(buf)
    good = buf.getvalue()
    assert service.check(good, Profile()).issues == []
    arial = service.check(good, Profile(font_name="Arial", font_size=12))
    assert any(i.code == "font_name" and "Times New Roman" in i.detail for i in arial.issues)
    assert any(i.code == "font_size" and "14" in i.detail for i in arial.issues)


def test_wrong_format_message(env):
    message = FakeMessage(document=FakeDocument("отчёт.doc", 5000))
    run(H.on_document(message, env["bot"], env["store"], env["settings"], env["state"], 20))
    assert "Сохранить как" in message.sent[-1][0]
