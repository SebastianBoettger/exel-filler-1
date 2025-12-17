from __future__ import annotations
import pandas as pd
from .normalize import norm_key, is_missing

class MatchEngine:
    def __init__(self, df1: pd.DataFrame, key1: str, df2: pd.DataFrame, key2: str, keep_zeros=True):
        self.df1 = df1
        self.df2 = df2
        self.key1 = key1
        self.key2 = key2
        self.keep_zeros = keep_zeros

        self.df1["_KEY_"] = self.df1[key1].map(lambda x: norm_key(x, keep_zeros))
        self.df2["_KEY_"] = self.df2[key2].map(lambda x: norm_key(x, keep_zeros))

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

    def t1_row_index_for_key(self, key: str) -> int | None:
        hit = self.df1.index[self.df1["_KEY_"] == key]
        if len(hit) == 0:
            return None
        return int(hit[0])

    def t2_rows_for_key(self, key: str) -> pd.DataFrame:
        if key not in self.t2_groups.groups:
            return self.df2.iloc[0:0].copy()
        return self.t2_groups.get_group(key).copy()
