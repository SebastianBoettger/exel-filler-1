from __future__ import annotations

import json
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List

@dataclass
class AppSettings:
    col_links: Dict[str, str]
    cuts: Dict[str, bool]
    country_default_value: str

    t1_hidden: List[str]
    t2_hidden: List[str]
    t1_colors: Dict[str, str]  # col -> "#RRGGBB"
    t2_colors: Dict[str, str]
    t1_order: List[str]
    t2_order: List[str]

    @staticmethod
    def defaults() -> "AppSettings":
        return AppSettings(
            col_links={},
            cuts={
                "split_street_house": True,
                "normalize_phone": True,
                "fill_country_default": False,
                "infer_state_from_zip": False,
            },
            country_default_value="Deutschland",
            t1_hidden=[],
            t2_hidden=[],
            t1_colors={},
            t2_colors={},
            t1_order=[],
            t2_order=[],
        )

def settings_path() -> Path:
    base = Path.home() / ".excel_filler_gui"
    base.mkdir(parents=True, exist_ok=True)
    return base / "settings.json"

def load_settings() -> AppSettings:
    p = settings_path()
    if not p.exists():
        return AppSettings.defaults()
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return AppSettings.defaults()

    d = AppSettings.defaults()
    return AppSettings(
        col_links=data.get("col_links", d.col_links) or {},
        cuts={**d.cuts, **(data.get("cuts", {}) or {})},
        country_default_value=str(data.get("country_default_value", d.country_default_value) or d.country_default_value),
        t1_hidden=list(data.get("t1_hidden", d.t1_hidden) or []),
        t2_hidden=list(data.get("t2_hidden", d.t2_hidden) or []),
        t1_colors=dict(data.get("t1_colors", d.t1_colors) or {}),
        t2_colors=dict(data.get("t2_colors", d.t2_colors) or {}),
        t1_order=list(data.get("t1_order", d.t1_order) or []),
        t2_order=list(data.get("t2_order", d.t2_order) or []),
    )

def save_settings(s: AppSettings) -> None:
    p = settings_path()
    p.write_text(json.dumps(asdict(s), ensure_ascii=False, indent=2), encoding="utf-8")
