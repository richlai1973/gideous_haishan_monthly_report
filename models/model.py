"""統一資料模型（設計計畫書 §4）。解析層寫入、產出引擎只讀。"""

from __future__ import annotations

import json
import os
from typing import Any

EMPTY_MODEL: dict[str, Any] = {
    "meta": {},
    "annual": {},
    "distribution_plan": {},
    "prayer_rota": {},
    "ministry_stats": {},
    "offerings": [],
    "finance": {"income": [], "expense": []},
    "bible_giving": {"general": [], "schools": [], "schedule": []},
    "prayers": {"members": [], "new_members": [], "churches": []},
    "church_testimony": [],
    "agenda_todo": {"motions": [], "guest_reports": [], "next_events": []},
    "analysis": {},
    "_sources": [],
}

MODULE_KEYS = {
    "finance": ["finance", "offerings", "ministry_stats", "church_testimony", "analysis"],
    "line": ["annual", "bible_giving", "prayer_rota", "agenda_todo"],
    "prayer": ["prayers"],
    "other": ["distribution_plan", "bible_giving", "prayer_rota", "agenda_todo"],
}

AFFECTED_DOCS = {
    "generate": list(range(1, 11)),
    "finance": [2, 3, 4, 9],
    "line": [1, 5, 6],
    "prayer": [7],
    "other": [1, 5, 8],
}


def model_path(work_dir: str) -> str:
    return os.path.join(work_dir, "_inputs", "model.json")


def load_model(work_dir: str) -> dict:
    p = model_path(work_dir)
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            m = json.load(f)
        return {**json.loads(json.dumps(EMPTY_MODEL)), **m}
    return json.loads(json.dumps(EMPTY_MODEL))


def save_model(work_dir: str, model: dict) -> str:
    p = model_path(work_dir)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(model, f, ensure_ascii=False, indent=2, default=str)
    return p


def merge(model: dict, patch: dict, source: str = "") -> dict:
    """淺層合併：dict 逐鍵覆寫、list 直接取代，並記錄來源。"""
    for k, v in (patch or {}).items():
        if isinstance(v, dict) and isinstance(model.get(k), dict):
            model[k] = {**model[k], **v}
        else:
            model[k] = v
    if source:
        model.setdefault("_sources", []).append(source)
    return model
