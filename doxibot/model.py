from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable

from .rules import ERROR, RULES, Rule


@dataclass
class Issue:
    code: str
    where: str
    detail: str = ""
    anchor: Any = None                      # w:p, к которому привязывается примечание
    fix: Callable[[], None] | None = None

    @property
    def rule(self) -> Rule:
        return RULES[self.code]

    @property
    def is_error(self) -> bool:
        return self.rule.severity == ERROR

    @property
    def fixable(self) -> bool:
        return self.fix is not None


@dataclass
class Analysis:
    issues: list[Issue] = field(default_factory=list)
    stats: dict[str, int] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)
    post_fixes: list[Callable[[], None]] = field(default_factory=list)

    @property
    def errors(self) -> list[Issue]:
        return [i for i in self.issues if i.is_error]

    @property
    def warnings(self) -> list[Issue]:
        return [i for i in self.issues if not i.is_error]

    @property
    def fixable(self) -> list[Issue]:
        return [i for i in self.issues if i.fixable]
