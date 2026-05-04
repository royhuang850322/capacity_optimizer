from __future__ import annotations

MAX_BASIS = "Max"
PLANNED_BASIS = "Planned"
LEGACY_PLANNER_BASIS = "Planner"

CAPACITY_BASES: tuple[str, str] = (MAX_BASIS, PLANNED_BASIS)


def normalize_capacity_basis(value: str | None) -> str:
    text = str(value or "").strip()
    if not text:
        return PLANNED_BASIS
    if text == LEGACY_PLANNER_BASIS:
        return PLANNED_BASIS
    return text
