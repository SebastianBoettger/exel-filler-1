from __future__ import annotations
import pandas as pd
from .normalize import norm_key, norm_text, is_missing

class MatchEngine:
    def __init__(self, df1: pd.DataFrame, key1: str, df2: pd.DataFrame, key2: str, keep_zeros=True):
        self.df1 = df1
        self.df2 = df2
        self.key1 = key1
        self.key2 = key2
        self.keep_zeros = keep_zeros

        self.df1["_KEY_"] = self.df1[key1].map(lambda x: norm_key(x, keep_zeros))
        self.df2["_KEY_"] = self.df2[key2].map(lambda x: norm_key(x, keep_zeros))

        # Index: KEY -> alle Zeilen in Tabelle 2
        self.t2_groups = self.df2.groupby("_KEY_", dropna=False)

    def keys_with_missing(self, columns_to_check: list[str]) -> list[str]:
        keys = []
        for _, row in self.df1.iterrows():
            k = row["_KEY_"]
            if not k:
                continue
            for col in columns_to_check:
                if col in self.df1.columns and is_missing(row.get(col)):
                    keys.append(k)
                    break
        return keys

    def get_row1_by_key(self, key: str) -> pd.Series | None:
        hit = self.df1[self.df1["_KEY_"] == key]
        if hit.empty:
            return None
        return hit.iloc[0]

    def suggestions_for_key(self, key: str, target_cols: list[str], mapping: dict[str, str]) -> dict[str, list[dict]]:
        """
        mapping: T1_col -> T2_col
        returns: {T1_col: [{value, source_col, score, row_index}]}
        """
        out: dict[str, list[dict]] = {c: [] for c in target_cols}

        if key not in self.t2_groups.groups:
            return out

        grp = self.t2_groups.get_group(key)

        for t1_col in target_cols:
            t2_col = mapping.get(t1_col)
            if not t2_col or t2_col not in grp.columns:
                continue

            for idx, r in grp.iterrows():
                v = norm_text(r.get(t2_col))
                if v and not is_missing(v):
                    out[t1_col].append({
                        "value": v,
                        "source_col": t2_col,
                        "score": 100,
                        "row_index": int(idx),
                    })
        return out
    
    def t2_rows_for_key(self, key: str) -> pd.DataFrame:
        if key not in self.t2_groups.groups:
            return self.df2.iloc[0:0].copy()
        return self.t2_groups.get_group(key).copy()

    def t1_row_index_for_key(self, key: str) -> int | None:
        hit = self.df1.index[self.df1["_KEY_"] == key]
        if len(hit) == 0:
            return None
        return int(hit[0])
