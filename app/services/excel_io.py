from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
import pandas as pd

@dataclass
class ExcelTable:
    path: Path
    sheet: str
    df: pd.DataFrame

def _engine_for(path: str) -> str | None:
    p = path.lower()
    if p.endswith(".xlsx"):
        return "openpyxl"
    if p.endswith(".xls"):
        return "xlrd"
    return None

def list_sheets(path: str) -> list[str]:
    eng = _engine_for(path)
    try:
        xls = pd.ExcelFile(path, engine=eng)
        return xls.sheet_names
    except ImportError as e:
        # xlrd fehlt für .xls
        raise ImportError("Für .xls bitte 'pip install xlrd' ausführen.") from e

def load_table(path: str, sheet: str, header_row_1based: int) -> ExcelTable:
    eng = _engine_for(path)
    try:
        df = pd.read_excel(
            path,
            sheet_name=sheet,
            engine=eng,
            header=header_row_1based - 1,
            dtype=str,
        )
    except ImportError as e:
        raise ImportError("Für .xls bitte 'pip install xlrd' ausführen.") from e

    df.columns = [str(c).strip() for c in df.columns]
    return ExcelTable(Path(path), sheet, df)
