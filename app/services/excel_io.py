from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
import pandas as pd

@dataclass
class ExcelTable:
    path: Path
    sheet: str
    df: pd.DataFrame

def list_sheets(path: str) -> list[str]:
    xls = pd.ExcelFile(path, engine="openpyxl")
    return xls.sheet_names

def load_table(path: str, sheet: str, header_row_1based: int) -> ExcelTable:
    # header_row_1based: 1 = erste Zeile
    df = pd.read_excel(
        path,
        sheet_name=sheet,
        engine="openpyxl",
        header=header_row_1based - 1,
        dtype=str,          # wichtig: Kundennummer etc. als TEXT
    )
    df.columns = [str(c).strip() for c in df.columns]
    return ExcelTable(Path(path), sheet, df)
