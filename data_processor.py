# data_processor.py
"""
香港出入境流量数据处理、SOP 校验与 Excel 报表生成引擎
"""

from pathlib import Path
import io
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from download_data import clean_and_parse_dates

# =====================================================
# File Path Configurations
# =====================================================
BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
TEMPLATE_DIR = BASE_DIR / "template"

RAW_FILE = INPUT_DIR / "latest_immd_raw.csv"
MAPPING_FILE = TEMPLATE_DIR / "port_mapping.xlsx"

INPUT_DIR.mkdir(exist_ok=True)


# =====================================================
# Data Processing Pipeline (Pandas)
# =====================================================
def load_data():
    df = pd.read_csv(RAW_FILE)
    df.columns = df.columns.str.strip()

    # 统一调用单一点日期解析逻辑
    df["Date"] = clean_and_parse_dates(df["Date"])
    df = df.dropna(subset=["Date"])

    numeric_cols = ["Hong Kong Residents", "Mainland Visitors", "Other Visitors", "Total"]
    for col in numeric_cols:
        if col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].astype(str).str.replace(",", "")
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def calculate_port_totals(df):
    arrival = df[df["Arrival / Departure"] == "Arrival"]
    departure = df[df["Arrival / Departure"] == "Departure"]

    mainland = arrival.groupby("Control Point")["Mainland Visitors"].sum()
    local = departure.groupby("Control Point")["Hong Kong Residents"].sum()

    result = pd.concat(
        [mainland.rename("Mainland Arrivals"), local.rename("Local Departures")],
        axis=1,
    )
    return result.fillna(0).reset_index()


def apply_mapping(df):
    mapping = pd.read_excel(MAPPING_FILE)
    required = ["Control Point", "Display Name", "Group"]
    missing = [x for x in required if x not in mapping.columns]
    if missing:
        raise ValueError(f"Missing columns in port_mapping.xlsx: {missing}")

    df = df.merge(mapping, on="Control Point", how="left")
    df["Display Name"] = df["Display Name"].fillna(df["Control Point"])
    df["Group"] = df["Group"].fillna("Other")
    return df


def merge_other(df):
    keep = ["HK-SZ", "Rail", "Airport", "Bridge"]
    main = df[df["Group"].isin(keep)].copy()
    other = df[~df["Group"].isin(keep)].copy()

    if other.empty:
        return main

    other_row = pd.DataFrame(
        [
            {
                "Control Point": "Other",
                "Display Name": "Other",
                "Group": "Other",
                "Mainland Arrivals": other["Mainland Arrivals"].sum(),
                "Local Departures": other["Local Departures"].sum(),
            }
        ]
    )
    return pd.concat([main, other_row], ignore_index=True)


def merge_current_previous(current, previous):
    previous = previous.rename(
        columns={
            "Mainland Arrivals": "Previous Mainland Arrivals",
            "Local Departures": "Previous Local Departures",
        }
    )
    previous = previous[["Display Name", "Previous Mainland Arrivals", "Previous Local Departures"]]
    result = current.merge(previous, on="Display Name", how="outer")

    for col in ["Mainland Arrivals", "Local Departures", "Previous Mainland Arrivals", "Previous Local Departures"]:
        result[col] = result[col].fillna(0)

    return result


def calculate_metrics(df):
    total_arrival = df["Mainland Arrivals"].sum()
    total_departure = df["Local Departures"].sum()

    df["Arrival Share"] = df["Mainland Arrivals"] / total_arrival if total_arrival > 0 else 0
    df["Departure Share"] = df["Local Departures"] / total_departure if total_departure > 0 else 0

    df["Arrival WoW %"] = np.where(
        df["Previous Mainland Arrivals"] == 0,
        0,
        (df["Mainland Arrivals"] - df["Previous Mainland Arrivals"]) / df["Previous Mainland Arrivals"],
    )

    df["Departure WoW %"] = np.where(
        df["Previous Local Departures"] == 0,
        0,
        (df["Local Departures"] - df["Previous Local Departures"]) / df["Previous Local Departures"],
    )

    ordered_cols = [
        "Group",
        "Display Name",
        "Mainland Arrivals",
        "Arrival Share",
        "Arrival WoW %",
        "Local Departures",
        "Departure Share",
        "Departure WoW %",
        "Previous Mainland Arrivals",
        "Previous Local Departures",
    ]
    return df[ordered_cols]


def sort_report(df):
    group_order = {"HK-SZ": 1, "Rail": 2, "Airport": 2, "Bridge": 2, "Other": 3}
    df["Group Order"] = df["Group"].map(group_order)

    return (
        df.sort_values(["Group Order", "Arrival Share"], ascending=[True, False])
        .drop(columns="Group Order")
        .reset_index(drop=True)
    )


def add_total(df):
    prev_mainland_sum = df["Previous Mainland Arrivals"].sum()
    prev_local_sum = df["Previous Local Departures"].sum()

    arrival_wow = (
        (df["Mainland Arrivals"].sum() - prev_mainland_sum) / prev_mainland_sum
        if prev_mainland_sum != 0 else 0.0
    )
    departure_wow = (
        (df["Local Departures"].sum() - prev_local_sum) / prev_local_sum
        if prev_local_sum != 0 else 0.0
    )

    total = pd.DataFrame(
        [
            {
                "Group": "Grand Total",
                "Display Name": "Grand Total",
                "Mainland Arrivals": df["Mainland Arrivals"].sum(),
                "Arrival Share": np.nan,
                "Arrival WoW %": arrival_wow,
                "Local Departures": df["Local Departures"].sum(),
                "Departure Share": np.nan,
                "Departure WoW %": departure_wow,
                "Previous Mainland Arrivals": prev_mainland_sum,
                "Previous Local Departures": prev_local_sum,
            }
        ]
    )
    return pd.concat([df, total], ignore_index=True)


def run_sop_checks(df):
    data_rows = df[df["Group"] != "Grand Total"]
    total_row = df[df["Group"] == "Grand Total"].iloc[0]

    assert data_rows["Mainland Arrivals"].sum() == total_row["Mainland Arrivals"], "Mainland Arrivals sum mismatch"
    assert data_rows["Local Departures"].sum() == total_row["Local Departures"], "Local Departures sum mismatch"
    assert abs(data_rows["Arrival Share"].sum() - 1.0) < 1e-5, "Arrival Share sum != 100%"
    assert abs(data_rows["Departure Share"].sum() - 1.0) < 1e-5, "Departure Share sum != 100%"


# =====================================================
# OpenPyXL Styling & Bytes Export (供网页直接下载)
# =====================================================
def get_openpyxl_percentage_format(val):
    if val == 0 or val is None or np.isnan(val):
        return '0%'

    abs_val = abs(val)
    decimals = 0
    while decimals <= 6:
        formatted_str = f"{abs_val * 100:.{decimals}f}%"
        if float(formatted_str.replace("%", "")) != 0:
            break
        decimals += 1

    return '0%' if decimals == 0 else '0.' + ('0' * decimals) + '%'


def export_styled_excel(df, weeks, output_target):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Weekly Report"

    font_family = "Arial"
    header_dark_fill = PatternFill(start_color="1E5631", end_color="1E5631", fill_type="solid")
    header_light_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    audit_fill = PatternFill(start_color="F3F3F3", end_color="F3F3F3", fill_type="solid")

    font_white_bold = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    font_bold = Font(name=font_family, size=10, bold=True)
    font_regular = Font(name=font_family, size=10)
    font_blue_bold = Font(name=font_family, size=10, bold=True, color="0000FF")
    font_blue_regular = Font(name=font_family, size=10, color="0000FF")
    font_red_bold = Font(name=font_family, size=10, bold=True, color="FF0000")
    font_red_regular = Font(name=font_family, size=10, color="FF0000")

    thin_border_side = Side(style="thin", color="D3D3D3")
    border_all_thin = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    date_str = f"{weeks['current_start'].strftime('%d/%m/%Y')} - {weeks['current_end'].strftime('%d/%m/%Y')}"
    ws["A1"] = date_str
    ws["A1"].font = Font(name=font_family, size=11, bold=True)
    ws.merge_cells("A1:B1")

    ws["C1"] = "Mainland Arrivals"
    ws["F1"] = "Local Departures"
    ws["I1"] = "Audit Data (Previous Week)"

    ws.merge_cells("C1:E1")
    ws.merge_cells("F1:H1")
    ws.merge_cells("I1:J1")

    for col_idx in range(3, 11):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_dark_fill if col_idx <= 8 else audit_fill
        cell.alignment = align_center
        cell.font = font_white_bold if col_idx <= 8 else font_bold

    sub_headers = [
        "Port", "",
        "Weekly Total", "% Share", "WoW %",
        "Weekly Total", "% Share", "WoW %",
        "Prev Mainland", "Prev Local"
    ]
    ws.append(sub_headers)
    ws.merge_cells("A2:B2")

    for col_idx in range(1, 11):
        cell = ws.cell(row=2, column=col_idx)
        cell.fill = header_light_fill if col_idx <= 8 else audit_fill
        cell.font = font_bold
        cell.alignment = align_center

    for row_idx, row_data in enumerate(df.to_dict("records"), start=3):
        is_total_row = (row_data["Group"] == "Grand Total")

        ws.cell(row=row_idx, column=1, value=row_data["Group"])
        ws.cell(row=row_idx, column=2, value=row_data["Display Name"])
        ws.cell(row=row_idx, column=3, value=row_data["Mainland Arrivals"])
        ws.cell(row=row_idx, column=4, value=row_data["Arrival Share"])
        ws.cell(row=row_idx, column=5, value=row_data["Arrival WoW %"])
        ws.cell(row=row_idx, column=6, value=row_data["Local Departures"])
        ws.cell(row=row_idx, column=7, value=row_data["Departure Share"])
        ws.cell(row=row_idx, column=8, value=row_data["Departure WoW %"])
        ws.cell(row=row_idx, column=9, value=row_data["Previous Mainland Arrivals"])
        ws.cell(row=row_idx, column=10, value=row_data["Previous Local Departures"])

        for col_idx in range(1, 11):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = border_all_thin
            cell.font = font_bold if is_total_row else font_regular
            cell.alignment = align_left if col_idx <= 2 else align_center

            if col_idx in [3, 6, 9, 10]:
                cell.number_format = "#,##0"
            elif col_idx in [4, 7]:
                val = cell.value
                if is_total_row or val is None or np.isnan(val):
                    cell.value = "-"
                    cell.number_format = '@'
                else:
                    cell.value = float(val)
                    cell.number_format = get_openpyxl_percentage_format(val)
            elif col_idx in [5, 8]:
                val = cell.value
                if val is not None and not np.isnan(val):
                    val_float = float(val)
                    cell.value = val_float
                    num_format_str = get_openpyxl_percentage_format(val_float)

                    if val_float == 0:
                        cell.number_format = num_format_str
                        cell.font = font_bold if is_total_row else font_regular
                    elif val_float > 0:
                        cell.number_format = f'"↑"{num_format_str}'
                        cell.font = font_blue_bold if is_total_row else font_blue_regular
                    else:
                        cell.number_format = f';"↓"{num_format_str}'
                        cell.font = font_red_bold if is_total_row else font_red_regular

    hksz_rows = [r for r in range(3, len(df) + 3) if ws.cell(row=r, column=1).value == "HK-SZ"]
    if hksz_rows:
        min_r, max_r = min(hksz_rows), max(hksz_rows)
        ws.merge_cells(start_row=min_r, start_column=1, end_row=max_r, end_column=1)
        ws.cell(row=min_r, column=1, value="HK - SZ ports").alignment = align_center

    for r in range(3, len(df) + 3):
        grp = ws.cell(row=r, column=1).value
        if grp in ["Rail", "Airport", "Bridge", "Other", "Grand Total"]:
            disp_name = ws.cell(row=r, column=2).value
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
            cell = ws.cell(row=r, column=1)
            cell.value = "Total" if grp == "Grand Total" else disp_name
            cell.alignment = align_left
            if grp == "Grand Total":
                cell.font = font_bold

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or "")
            if cell.row == 1:
                continue
            cell_len = sum(2 if ord(char) > 127 else 1 for char in val_str)
            max_len = max(max_len, cell_len)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 10)

    wb.save(output_target)


def export_report_to_excel_bytes(df, current_start, current_end):
    """直接在内存中生成带样式的 Excel 字节流，供 Streamlit 网页下载"""
    buffer = io.BytesIO()
    weeks = {"current_start": current_start, "current_end": current_end}
    export_styled_excel(df, weeks, buffer)
    buffer.seek(0)
    return buffer.getvalue()