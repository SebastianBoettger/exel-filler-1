from __future__ import annotations
import pandas as pd
import re

MISSING_TOKENS = {"", "nan", "none", "-", "n/a", "null"}

def norm_text(v: str | None) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    s = re.sub(r"\s+", " ", s)
    return s

def norm_key(v: str | None, keep_leading_zeros: bool = True) -> str:
    s = norm_text(v)
    if not keep_leading_zeros:
        s = s.lstrip("0")
    return s

def is_missing(v: str | None) -> bool:
    s = norm_text(v).lower()
    return s in MISSING_TOKENS

def missing_mask(df: pd.DataFrame) -> pd.DataFrame:
    # True = fehlt
    return df.applymap(is_missing)
