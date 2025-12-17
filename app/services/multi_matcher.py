from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
import pandas as pd


def _norm(v: object) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return " ".join(s.split()).lower()


def make_key(row: pd.Series, fields: List[str]) -> str:
    parts = []
    for f in fields:
        if f not in row.index:
            parts.append("")
        else:
            parts.append(_norm(row.get(f)))
    return "|".join(parts)


@dataclass
class MatchProfile:
    t1_fields: List[str]
    src_fields: Dict[str, List[str]]  # src_id -> fields


class MultiMatcher:
    """
    Matching: T1 row -> per source list of matching row indices.
    Keys: build_key(T1, t1_fields) == build_key(src, src_fields[src_id])
    """

    def __init__(self, t1_df: pd.DataFrame, sources: Dict[str, pd.DataFrame], profile: MatchProfile):
        self.t1_df = t1_df
        self.sources = sources
        self.profile = profile

        # Build indexes for sources: src_id -> key -> [row_index]
        self.src_index: Dict[str, Dict[str, List[int]]] = {}
        for src_id, df in sources.items():
            fields = profile.src_fields.get(src_id, [])
            idx: Dict[str, List[int]] = {}
            if not fields:
                self.src_index[src_id] = idx
                continue

            for i, row in df.iterrows():
                k = make_key(row, fields)
                if k:
                    idx.setdefault(k, []).append(int(i))
            self.src_index[src_id] = idx

    def match_for_t1_index(self, t1_index: int) -> Dict[str, List[int]]:
        if t1_index not in self.t1_df.index:
            return {sid: [] for sid in self.sources.keys()}

        row = self.t1_df.loc[t1_index]
        k = make_key(row, self.profile.t1_fields)
        out: Dict[str, List[int]] = {}
        for src_id in self.sources.keys():
            idx = self.src_index.get(src_id, {})
            out[src_id] = idx.get(k, []).copy() if k else []
        return out
