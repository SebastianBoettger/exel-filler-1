from __future__ import annotations

import json
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List


@dataclass
class AppSettings:
    # Matching profile (fields)
    t1_match_fields: List[str]
    src_match_fields: Dict[str, List[str]]  # src_id -> fields

    # Existing logic settings
    col_links: Dict[str, str]
    cuts: Dict[str, bool]
    country_default_value: str

    # Table preferences
    t1_hidden: List[str]
    t1_colors: Dict[str, str]
    t1_order: List[str]

    src_hidden: Dict[str, List[str]]     # src_id -> hidden list
    src_colors: Dict[str, Dict[str, str]]  # src_id -> {col: "#RRGGBB"}
    src_order: Dict[str, List[str]]      # src_id -> order list

    @staticmethod
    def defaults() -> "AppSettings":
        return AppSettings(
            t1_match_fields=[],
            src_match_fields={"src2": [], "src3": [], "src4": []},
            col_links={},
            cuts={
                "split_street_house": True,
                "normalize_phone": True,
                "fill_country_default": False,
                "infer_state_from_zip": False,
            },
            country_default_value="Deutschland",
            t1_hidden=[],
            t1_colors={},
            t1_order=[],
            src_hidden={"src2": [], "src3": [], "src4": []},
            src_colors={"src2": {}, "src3": {}, "src4": {}},
            src_order={"src2": [], "src3": [], "src4": []},
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
        t1_match_fields=list(data.get("t1_match_fields", d.t1_match_fields) or []),
        src_match_fields={**d.src_match_fields, **(data.get("src_match_fields", {}) or {})},

        col_links=dict(data.get("col_links", d.col_links) or {}),
        cuts={**d.cuts, **(data.get("cuts", {}) or {})},
        country_default_value=str(data.get("country_default_value", d.country_default_value) or d.country_default_value),

        t1_hidden=list(data.get("t1_hidden", d.t1_hidden) or []),
        t1_colors=dict(data.get("t1_colors", d.t1_colors) or {}),
        t1_order=list(data.get("t1_order", d.t1_order) or []),

        src_hidden={**d.src_hidden, **(data.get("src_hidden", {}) or {})},
        src_colors={**d.src_colors, **(data.get("src_colors", {}) or {})},
        src_order={**d.src_order, **(data.get("src_order", {}) or {})},
    )


def save_settings(s: AppSettings) -> None:
    p = settings_path()
    p.write_text(json.dumps(asdict(s), ensure_ascii=False, indent=2), encoding="utf-8")
