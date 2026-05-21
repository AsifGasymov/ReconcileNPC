"""Stage 2 — ERP Merger: Transactions × Statement (FX) × Orders → ERP Upload."""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

import openpyxl
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from .excel_styles import (
    ALT, DARK_BLUE, GOLD, GOLD_ALT, GREEN_FILL, MID_BLUE, ORANGE_ALT, ORANGE_FILL, WHITE,
    data_font, header_font, style_cell,
)


LogFn = Callable[[str], None]


@dataclass
class ERPResult:
    rows: int
    rates_count: int
    total_net_gbp: float
    out_path: str
    out_paths: list[str]


FEE_COLS = [
    "EU Fee", "Non-EU Fee", "IC++ Fee", "IC++ Interchange",
    "IC++ Scheme Fee", "Approve Fee", "Decline Fee", "Refund Fee",
]

OUTPUT_COLS: list[tuple[str, int]] = [
    ("Company", 18), ("IPS", 7), ("Region", 10),
    ("RRN", 16), ("Payment ID", 32),
    ("Statement Date", 16),
    ("Transaction Date/Time", 20), ("Processing Date/Time", 20),
    ("Currency", 10), ("Trn Amount", 14), ("FX Rate to GBP", 15),
    ("Trn Amount GBP", 16), ("Total Fees", 12),
    ("Net Amount", 12), ("Net Amount GBP", 16),
]

GOLD_COLS = {"FX Rate to GBP", "Trn Amount GBP", "Net Amount GBP"}
TEXT_COLS = {"RRN"}
NUMBER_FORMATS = {
    "Trn Amount": "#,##0.00", "FX Rate to GBP": "0.00000",
    "Trn Amount GBP": "#,##0.00", "Total Fees": "#,##0.0000",
    "Net Amount": "#,##0.0000", "Net Amount GBP": "#,##0.00",
}


def _sanitize_filename(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]', '', name)
    return name.strip()[:40]


def _is_priority_row(pid) -> bool:
    s = str(pid).strip() if pd.notna(pid) else ""
    return s == "" or s.lower().startswith("charge for")


def _extract_rates(stmt_path: str) -> tuple[dict[tuple[str, str], float], str]:
    wb = openpyxl.load_workbook(stmt_path, data_only=True)
    rates: dict[tuple[str, str], float] = {}
    stmt_date = ""

    for sheet_name in wb.sheetnames:
        ccy = sheet_name.upper()
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        if not stmt_date:
            for row in rows[:6]:
                for cell in row:
                    s = str(cell or "")
                    if re.match(r"\d{2}\.\d{2}\.\d{4}", s):
                        stmt_date = s
                        break
                if stmt_date:
                    break

        mc_rate = visa_rate = None
        for row in rows:
            a0 = str(row[0] or "")
            a2 = str(row[2] or "").replace(",", ".")
            if "Mastercard Payout" in a0:
                try:
                    mc_rate = float(a2)
                except ValueError:
                    pass
            if "Visa Payout" in a0:
                try:
                    visa_rate = float(a2)
                except ValueError:
                    pass

        if ccy == "EUR":
            rates[("EUR", "MC")] = 1.0
            rates[("EUR", "VISA")] = 1.0
        if mc_rate is not None:
            rates[(ccy, "MC")] = mc_rate
        if visa_rate is not None:
            rates[(ccy, "VISA")] = visa_rate

    return rates, stmt_date


def _read_transactions(trn_path: str) -> pd.DataFrame:
    p = trn_path.lower()
    if p.endswith(".csv"):
        try:
            df = pd.read_csv(trn_path, encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(trn_path, encoding="latin-1")
    else:
        df = pd.read_excel(trn_path)
    return df.dropna(how="all")


def _read_orders(orders_path: str) -> pd.DataFrame:
    df = pd.read_excel(orders_path, sheet_name="Orders",
                       dtype={"ARN": str, "RRN": str})
    df = df.dropna(how="all")
    df["ARN"] = df["ARN"].apply(
        lambda x: str(x).split(".")[0].strip() if pd.notna(x) else "")
    df["RRN"] = df["RRN"].apply(
        lambda x: str(x).split(".")[0].strip() if pd.notna(x) else "")
    df["Payment ID"] = df["Product name or description"].apply(
        lambda x: str(x)[:28].strip() if pd.notna(x) else "")
    return df[["ARN", "RRN", "Payment ID"]]


def _build_workbook(trn_df: pd.DataFrame, rates: dict, stmt_date: str,
                    out_path: str, orders_df: Optional[pd.DataFrame]) -> tuple[int, int, float]:
    trn_df = trn_df.copy()

    if orders_df is not None and not orders_df.empty:
        trn_df["_RRN"] = trn_df["RRN"].apply(
            lambda x: str(x).split(".")[0].strip() if pd.notna(x) else "")
        rrn_map = orders_df[orders_df["RRN"] != ""].set_index("RRN")["Payment ID"].to_dict()
        trn_df["Payment ID"] = trn_df["_RRN"].map(rrn_map).fillna("")
        trn_df.drop(columns=["_RRN"], inplace=True)
    else:
        trn_df["Payment ID"] = ""

    trn_df["Statement Date"] = stmt_date

    def get_rate(row):
        ccy = str(row.get("Currency", "")).upper().strip()
        ips = str(row.get("IPS", "")).upper().strip()
        return rates.get((ccy, ips))

    trn_df["FX Rate to GBP"] = trn_df.apply(get_rate, axis=1)

    trn_df["_priority"] = trn_df["Payment ID"].apply(_is_priority_row)
    trn_df = trn_df.sort_values("_priority", ascending=False).reset_index(drop=True)

    present_fees = [c for c in FEE_COLS if c in trn_df.columns]
    for c in present_fees:
        trn_df[c] = pd.to_numeric(trn_df[c], errors="coerce").fillna(0)
    trn_df["Trn Amount"] = pd.to_numeric(trn_df["Trn Amount"], errors="coerce").fillna(0)
    trn_df["Total Fees"] = trn_df[present_fees].sum(axis=1) if present_fees else 0
    trn_df["Net Amount"] = trn_df["Trn Amount"] - trn_df["Total Fees"]
    trn_df["Trn Amount GBP"] = trn_df.apply(
        lambda r: round(r["Trn Amount"] * r["FX Rate to GBP"], 2)
        if pd.notna(r["FX Rate to GBP"]) else None, axis=1)
    trn_df["Net Amount GBP"] = trn_df.apply(
        lambda r: round(r["Net Amount"] * r["FX Rate to GBP"], 2)
        if pd.notna(r["FX Rate to GBP"]) else None, axis=1)

    wb = Workbook()

    # Sheet 1: ERP Upload
    ws1 = wb.active
    ws1.title = "ERP Upload"
    ws1.row_dimensions[1].height = 28
    ws1.row_dimensions[2].height = 18

    ws1.merge_cells(f"A1:{get_column_letter(len(OUTPUT_COLS))}1")
    title_cell = ws1["A1"]
    title_cell.value = f"Transaction Report — ERP Upload Ready   |   Statement: {stmt_date}"
    title_cell.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    title_cell.fill = DARK_BLUE
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    for ci, (name, width) in enumerate(OUTPUT_COLS, 1):
        style_cell(ws1, 2, ci, name, fill=MID_BLUE,
                   font=header_font(size=9), align="center")
        ws1.column_dimensions[get_column_letter(ci)].width = width

    src_cols = [c[0] for c in OUTPUT_COLS]
    for ri, (_, row) in enumerate(trn_df.iterrows(), 3):
        even = ri % 2 == 0
        is_priority = bool(row.get("_priority", False))
        for ci, col in enumerate(src_cols, 1):
            val = row.get(col)
            if not isinstance(val, str) and pd.isna(val):
                val = None
            if col in TEXT_COLS and val is not None:
                val = str(val).split(".")[0]
            is_gold = col in GOLD_COLS
            if is_priority:
                fill = ORANGE_FILL if even else ORANGE_ALT
            elif is_gold:
                fill = GOLD if even else GOLD_ALT
            else:
                fill = ALT if even else WHITE
            cell = style_cell(ws1, ri, ci, val, fill=fill,
                              font=data_font(),
                              number_format=NUMBER_FORMATS.get(col))
            if col in TEXT_COLS:
                cell.number_format = "@"

    ws1.freeze_panes = "A3"
    trn_df.drop(columns=["_priority"], inplace=True, errors="ignore")

    # Sheet 2: FX Rates
    ws2 = wb.create_sheet("FX Rates")
    ws2.merge_cells("A1:D1")
    fx_title = ws2["A1"]
    fx_title.value = f"FX Rates Applied   |   Source: Statement {stmt_date}"
    fx_title.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    fx_title.fill = DARK_BLUE
    fx_title.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 26

    for ci, (h, w) in enumerate(
            [("Currency", 14), ("IPS", 10), ("Rate to GBP", 16), ("Note", 34)], 1):
        style_cell(ws2, 2, ci, h, fill=MID_BLUE,
                   font=header_font(size=9), align="center")
        ws2.column_dimensions[get_column_letter(ci)].width = w

    note_map = {
        ("EUR", "MC"): "Settlement currency",
        ("EUR", "VISA"): "Settlement currency",
    }
    for ri, ((ccy, ips), rate) in enumerate(sorted(rates.items()), 3):
        fill = ALT if ri % 2 == 0 else WHITE
        style_cell(ws2, ri, 1, ccy, fill=fill, font=data_font(True))
        style_cell(ws2, ri, 2, ips, fill=fill, font=data_font())
        style_cell(ws2, ri, 3, rate, fill=GOLD, font=data_font(),
                   number_format="0.00000", align="center")
        style_cell(ws2, ri, 4,
                   note_map.get((ccy, ips), "From statement Payout row"),
                   fill=fill, font=data_font())

    # Sheet 3: Summary
    ws3 = wb.create_sheet("Summary")
    ws3.merge_cells("A1:E1")
    sum_title = ws3["A1"]
    sum_title.value = f"Summary by Currency & IPS   |   {stmt_date}"
    sum_title.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    sum_title.fill = DARK_BLUE
    sum_title.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 26

    for ci, (h, w) in enumerate(
            [("Currency", 12), ("IPS", 10), ("Trn Count", 12),
             ("Total Trn Amount", 18), ("Total Net GBP", 16)], 1):
        style_cell(ws3, 2, ci, h, fill=MID_BLUE,
                   font=header_font(size=9), align="center")
        ws3.column_dimensions[get_column_letter(ci)].width = w

    summary = trn_df.groupby(["Currency", "IPS"]).agg(
        Count=("Trn Amount", "count"),
        Total_Trn=("Trn Amount", "sum"),
        Total_Net_GBP=("Net Amount GBP", "sum"),
    ).reset_index()

    for ri, (_, row) in enumerate(summary.iterrows(), 3):
        fill = ALT if ri % 2 == 0 else WHITE
        style_cell(ws3, ri, 1, row["Currency"], fill=fill, font=data_font(True))
        style_cell(ws3, ri, 2, row["IPS"], fill=fill, font=data_font())
        style_cell(ws3, ri, 3, int(row["Count"]), fill=fill,
                   font=data_font(), align="center")
        style_cell(ws3, ri, 4, round(row["Total_Trn"], 2), fill=fill,
                   font=data_font(), number_format="#,##0.00")
        style_cell(ws3, ri, 5, round(float(row["Total_Net_GBP"] or 0), 2),
                   fill=GREEN_FILL, font=data_font(True),
                   number_format="#,##0.00")

    last = 3 + len(summary)
    ws3.merge_cells(f"A{last}:C{last}")
    style_cell(ws3, last, 1, "TOTAL", fill=DARK_BLUE,
               font=header_font(), align="right")
    style_cell(ws3, last, 4, round(summary["Total_Trn"].sum(), 2),
               fill=DARK_BLUE, font=header_font(), number_format="#,##0.00")
    style_cell(ws3, last, 5, round(float(summary["Total_Net_GBP"].sum()), 2),
               fill=DARK_BLUE,
               font=Font(name="Arial", size=10, bold=True, color="FFD700"),
               number_format="#,##0.00")

    wb.save(out_path)
    return len(trn_df), len(rates), float(summary["Total_Net_GBP"].sum())


def run_erp_merger(*, transactions_paths: list[str], statement_paths: list[str],
                   out_dir: str, orders_path: Optional[str] = None,
                   log: Optional[LogFn] = None) -> ERPResult:
    """Execute the ERP merger pipeline for one or more transaction/statement pairs."""
    def lg(msg: str) -> None:
        if log:
            log(msg)

    orders_df = None
    if orders_path:
        lg("Reading order statement…")
        orders_df = _read_orders(orders_path)

    pairs = list(zip(transactions_paths, statement_paths))
    total_rows = 0
    total_rates = 0
    total_gbp = 0.0
    out_paths: list[str] = []

    for i, (trn_path, stmt_path) in enumerate(pairs, 1):
        prefix = f"[{i}/{len(pairs)}] " if len(pairs) > 1 else ""

        lg(f"{prefix}Reading statement…")
        rates, stmt_date = _extract_rates(stmt_path)
        lg(f"  rates: {len(rates)} · date: {stmt_date or 'n/a'}")

        lg(f"{prefix}Reading transactions…")
        trn_df = _read_transactions(trn_path)
        lg(f"  rows: {len(trn_df)}")

        companies = trn_df["Company"].dropna() if "Company" in trn_df.columns else pd.Series(dtype=str)
        companies = companies[companies.astype(str).str.strip().ne("")]
        company = _sanitize_filename(str(companies.mode()[0])) if not companies.empty else "Unknown"

        date_tag = stmt_date.replace(".", "-") if stmt_date else datetime.now().strftime("%d-%m-%Y")
        out_name = f"ERP_{company}_{date_tag}.xlsx"
        out_path = str(Path(out_dir) / out_name)

        lg(f"{prefix}Writing {out_name}…")
        rows, n_rates, gbp = _build_workbook(trn_df, rates, stmt_date, out_path, orders_df)

        total_rows += rows
        total_rates = max(total_rates, n_rates)
        total_gbp += gbp
        out_paths.append(out_path)

    return ERPResult(
        rows=total_rows,
        rates_count=total_rates,
        total_net_gbp=total_gbp,
        out_path=out_dir,
        out_paths=out_paths,
    )
