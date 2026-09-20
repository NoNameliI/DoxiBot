"""Текст отчётов для Telegram (HTML) и полный отчёт в .txt."""

from __future__ import annotations

from collections import OrderedDict
from html import escape

from .model import Analysis, Issue

TG_LIMIT = 3900
EXAMPLES_IN_CHAT = 4


def _plural(n: int, one: str, few: str, many: str) -> str:
    if n % 10 == 1 and n % 100 != 11:
        return one
    if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        return few
    return many


def group_issues(issues: list[Issue]) -> list[tuple[str, list[Issue]]]:
    groups: OrderedDict[str, list[Issue]] = OrderedDict()
    for issue in issues:
        groups.setdefault(issue.code, []).append(issue)
    return sorted(groups.items(), key=lambda kv: (not kv[1][0].is_error,))


def _stats_line(analysis: Analysis) -> str:
    s = analysis.stats
    parts = [f"{s.get('paragraphs', 0)} абз."]
    for key, label in (("headings", "загол."), ("pictures", "рис."), ("tables", "табл."), ("formulas", "формул")):
        if s.get(key):
            parts.append(f"{s[key]} {label}")
    return " · ".join(parts)


def summary_header(analysis: Analysis, filename: str) -> str:
    errors, warnings = len(analysis.errors), len(analysis.warnings)
    fixable = len(analysis.fixable)
    lines = [f"📄 <b>{escape(filename)}</b>", f"<i>Проверено: {_stats_line(analysis)}</i>", ""]
    if not analysis.issues:
        lines.append("✅ <b>Нарушений оформления не найдено!</b>")
        lines.append("Проверьте вручную то, что бот не видит: номера страниц титульного листа и задания, "
                     "смысл подписей и полноту ссылок.")
        return "\n".join(lines)
    lines.append(
        f"❌ Ошибок: <b>{errors}</b>   ⚠️ Замечаний: <b>{warnings}</b>\n"
        f"🛠 Исправлю автоматически: <b>{fixable}</b> из {len(analysis.issues)}"
    )
    for note in analysis.notes:
        lines.append(f"ℹ️ {escape(note)}")
    return "\n".join(lines)


def _group_html(code: str, issues: list[Issue], limit: int, profile) -> str:
    first = issues[0]
    icon = "❌" if first.is_error else "⚠️"
    fix_mark = " 🛠" if all(i.fixable for i in issues) else (" 🛠½" if any(i.fixable for i in issues) else "")
    head = f"{icon} <b>{escape(first.title(profile))}</b> — {len(issues)}{fix_mark}"
    body = []
    for issue in issues[:limit]:
        line = f"• {escape(issue.where)}"
        if issue.detail:
            line += f": <i>{escape(issue.detail)}</i>"
        body.append(line)
    if len(issues) > limit:
        rest = len(issues) - limit
        body.append(f"…и ещё {rest} {_plural(rest, 'место', 'места', 'мест')}")
    body.append(f"💡 {escape(first.hint(profile))}")
    return f"{head}\n<blockquote expandable>" + "\n".join(body) + "</blockquote>"


def chat_messages(analysis: Analysis, filename: str) -> tuple[list[str], bool]:
    """Сообщения для чата и признак, что часть примеров не поместилась (нужен файл отчёта)."""
    messages: list[str] = []
    truncated = False
    current = summary_header(analysis, filename)
    groups = group_issues(analysis.issues)
    for n, (code, issues) in enumerate(groups):
        block = _group_html(code, issues, EXAMPLES_IN_CHAT, analysis.profile)
        truncated = truncated or len(issues) > EXAMPLES_IN_CHAT
        if len(current) + len(block) + 2 > TG_LIMIT:
            messages.append(current)
            if len(messages) >= 3:
                rest = len(groups) - n
                messages[-1] += f"\n\n…и ещё {rest} {_plural(rest, 'правило', 'правила', 'правил')} — см. «📋 Полный отчёт»"
                return messages, True
            current = ""
        current += ("\n\n" if current else "") + block
    messages.append(current)
    return messages, truncated


def full_text_report(analysis: Analysis, filename: str) -> str:
    out = [
        f"DoxiBot — отчёт о проверке оформления",
        f"Файл: {filename}",
        f"Проверено: {_stats_line(analysis)}",
        f"Ошибок: {len(analysis.errors)}, замечаний: {len(analysis.warnings)}, "
        f"исправляется автоматически: {len(analysis.fixable)}",
        "",
    ]
    for note in analysis.notes:
        out.append(f"* {note}")
    for code, issues in group_issues(analysis.issues):
        kind = "ОШИБКА" if issues[0].is_error else "ЗАМЕЧАНИЕ"
        out.append("=" * 70)
        out.append(f"[{kind}] {issues[0].title(analysis.profile)} — {len(issues)}")
        out.append(f"Как исправить: {issues[0].hint(analysis.profile)}")
        out.append("-" * 70)
        for n, issue in enumerate(issues, 1):
            line = f"{n:>3}. {issue.where}"
            if issue.detail:
                line += f" — {issue.detail}"
            if issue.fixable:
                line += "  [автоисправление]"
            out.append(line)
        out.append("")
    return "\n".join(out)


def fix_summary(applied: int, failed: int, after: Analysis) -> str:
    lines = [f"🛠 <b>Готово!</b> Исправлено замечаний: <b>{applied}</b>"]
    if failed:
        lines.append(f"Не удалось применить: {failed}")
    remaining = after.issues
    if not remaining:
        lines.append("✅ Повторная проверка нарушений не нашла.")
    else:
        lines.append(
            f"Осталось поправить вручную: ❌ {len(after.errors)}  ⚠️ {len(after.warnings)}"
        )
        for code, issues in group_issues(remaining)[:12]:
            icon = "❌" if issues[0].is_error else "⚠️"
            lines.append(f"{icon} {escape(issues[0].title(after.profile))} — {len(issues)}")
        if len(group_issues(remaining)) > 12:
            lines.append("…")
    lines.append("\n<i>Меняется только оформление, регистр заголовков и номера подписей (ссылки на них "
                 "обновляются) — сам текст не переписывается. Проверьте результат; если есть автособираемое "
                 "содержание — обновите его (ПКМ → Обновить поле).</i>")
    return "\n".join(lines)
