from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Any, Literal


@dataclass(frozen=True)
class Anchor:
    anchor_id: str
    kind: Literal[
        "content_control",
        "heading",
        "bookmark",
        "paragraph",
        "table_cell",
        "image",
        "comment",
        "hyperlink",
        "field",
    ]
    label: str
    path: str
    text_preview: str = ""
    style_id: str | None = None
    level: int | None = None
    extra: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class ValidationIssue:
    severity: Literal["info", "warning", "error"]
    code: str
    message: str
    path: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class ValidationReport:
    ok: bool
    source_path: str
    target_path: str
    issues: list[ValidationIssue] = field(default_factory=list)
    metrics: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "ok": self.ok,
            "source_path": self.source_path,
            "target_path": self.target_path,
            "issues": [i.to_dict() for i in self.issues],
            "metrics": self.metrics,
        }
