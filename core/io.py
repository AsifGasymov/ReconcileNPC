"""Common file readers with encoding fallback."""
from __future__ import annotations

from typing import Optional

import pandas as pd


def read_table(path: str, *, sheet_name: Optional[str] = None,
               dtype: Optional[type | dict] = str) -> pd.DataFrame:
    """Read CSV/XLS/XLSX into a DataFrame.

    Tries utf-8 → latin-1 for CSV. Selects sheet only when given.
    """
    p = path.lower()
    if p.endswith(".csv"):
        try:
            df = pd.read_csv(path, dtype=dtype, encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(path, dtype=dtype, encoding="latin-1")
    elif p.endswith(".xls"):
        df = pd.read_excel(path, dtype=dtype, engine="xlrd")
    else:
        if sheet_name is not None:
            df = pd.read_excel(path, sheet_name=sheet_name, dtype=dtype)
        else:
            df = pd.read_excel(path, dtype=dtype)

    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]
    return df
