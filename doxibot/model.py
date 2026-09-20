from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable

from .profile import DEFAULT, Profile
from .rules import ERROR, RULES, Rule, hint_for, title_for


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

    def title(self, profile: Profile = DEFAULT) -> str:
        return title_for(self.code, profile)

    def hint(self, profile: Profile = DEFAULT) -> str:
        return hint_for(self.code, profile)

    @property
    def is_error(self) -> bool:
        return self.rule.severity == ERROR

    @property
    def fixable(self) -> bool:
        return self.fix is not None


@dataclass
class Analysis:
    profile: Profile = DEFAULT
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
