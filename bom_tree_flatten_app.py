"""
bom_tree_flatten_app.py
========================
Streamlit app: uploads a multi-level BOM tree export (LEVEL / ID / NEXT_ASSY
pattern) and produces a flat workbook where each part appears once, with its
QPA (Quantity Per Assembly) broken out into one column per parent assembly.

Deploy:
  1. Push this file + requirements.txt to a GitHub repo.
  2. On https://share.streamlit.io -> New app -> pick the repo/branch ->
     main file path: bom_tree_flatten_app.py -> Deploy.

Run locally:
  pip install -r requirements.txt
  streamlit run bom_tree_flatten_app.py
"""

import io
import re

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="BOM Tree Flattener", page_icon="\U0001F5C2\uFE0F", layout="wide")

INFO_COLS_PRIORITY = [
    "ID", "REV", "CN", "DESCRIPTION", "UOM", "TYPE", "STATUS",
    "OWNER", "GROUP", "BUY_STATUS", "PRD_TYPE", "CATEGORY",
]
REQUIRED_COLS = {"LEVEL", "ID", "NEXT_ASSY"}


# ----------------------------------------------------------------------
# Core logic (same engine as the standalone script)
# ----------------------------------------------------------------------

def sanitize_sheet_name(name: str, used_names: set) -> str:
    name = str(name)
    name = re.sub(r'[:\\/?*\[\]]', "_", name)
    name = name.strip() or "Sheet"
    name = name[:31]
    base = name
    n = 1
    while name in used_names:
        suffix = f"_{n}"
        name = base[: 31 - len(suffix)] + suffix
        n += 1
    used_names.add(name)
    return name


def load_sheet_as_table(raw: pd.DataFrame) -> pd.DataFrame:
    header = [str(h).strip() for h in raw.iloc[0].tolist()]
    data = raw.iloc[1:].copy()
    data.columns = header
    data = data.reset_index(drop=True)
    data = data.dropna(how="all")
    return data


def flatten_single_tree(block: pd.DataFrame, qty_col: str):
    comps = block[pd.to_numeric(block["LEVEL"], errors="coerce") > 0].copy()

    if qty_col not in comps.columns:
        raise ValueError(f"Quantity column '{qty_col}' not found in source data.")

    comps[qty_col] = pd.to_numeric(comps[qty_col], errors="coerce").fillna(0)
    comps["NEXT_ASSY"] = comps["NEXT_ASSY"].astype(str)
    comps["ID"] = comps["ID"].astype(str)

    info_cols = [c for c in INFO_COLS_PRIORITY if c in comps.columns]

    agg = (
        comps.groupby(["ID", "NEXT_ASSY"], dropna=False)[qty_col]
        .sum()
        .reset_index()
    )

    pivot = agg.pivot(index="ID", columns="NEXT_ASSY", values=qty_col)
    parent_ids = list(pivot.columns)

    item_info = comps.drop_duplicates(subset="ID", keep="first")[info_cols].set_index("ID")
    usage_count = comps.groupby("ID")["NEXT_ASSY"].nunique().rename("USED_IN_N_ASSEMBLIES")

    flat = item_info.join(usage_count).join(pivot).reset_index()

    qpa_rename = {p: f"QPA_in_{p}" for p in parent_ids}
    flat = flat.rename(columns=qpa_rename)
    flat = flat.sort_values("ID").reset_index(drop=True)

    all_ids = block[["ID", "DESCRIPTION", "REV"]].astype(str).drop_duplicates(subset="ID")
    legend = all_ids[all_ids["ID"].isin(parent_ids)].copy()
    legend = legend.rename(columns={"ID": "ASSEMBLY_ID"})
    legend = legend.sort_values("ASSEMBLY_ID").reset_index(drop=True)

    return flat, legend


def format_workbook(wb):
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    body_font = Font(name="Arial", size=10)
    qpa_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        if ws.max_row == 0:
            continue
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            header_val = str(cell.value or "")
            is_qpa = header_val.startswith("QPA_in_")
            max_len = len(header_val)
            for row_idx in range(2, ws.max_row + 1):
                body_cell = ws.cell(row=row_idx, column=col_idx)
                body_cell.font = body_font
                if is_qpa:
                    body_cell.alignment = Alignment(horizontal="center")
                    if body_cell.value not in (None, ""):
                        body_cell.fill = qpa_fill
                val_len = len(str(body_cell.value)) if body_cell.value is not None else 0
                max_len = max(max_len, val_len)
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 10), 40)
        ws.row_dimensions[1].height = 30
    return wb


def flatten_bom_tree(file_bytes: bytes, filename: str, qty_col: str = "ORIG_QTY"):
    """Returns (output_xlsx_bytes, summary_rows, preview_dict)."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    used_names = set()
    summary_rows = []
    preview = {}

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name in xls.sheet_names:
            raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            if raw.empty:
                continue
            data = load_sheet_as_table(raw)

            missing = REQUIRED_COLS - set(data.columns)
            if missing:
                st.warning(f"Sheet '{sheet_name}' skipped — missing columns: {missing}")
                continue

            level_numeric = pd.to_numeric(data["LEVEL"], errors="coerce")
            root_idx = data.index[level_numeric == 0].tolist()
            if not root_idx:
                st.warning(f"Sheet '{sheet_name}' skipped — no LEVEL=0 root row found")
                continue
            boundaries = root_idx + [len(data)]

            for i in range(len(root_idx)):
                start, end = boundaries[i], boundaries[i + 1]
                block = data.iloc[start:end].copy()
                root_row = block.iloc[0]
                root_id = str(root_row["ID"])
                root_desc = str(root_row.get("DESCRIPTION", ""))

                try:
                    flat_df, legend_df = flatten_single_tree(block, qty_col=qty_col)
                except ValueError as e:
                    st.warning(f"Root '{root_id}' in sheet '{sheet_name}' skipped: {e}")
                    continue

                out_sheet = sanitize_sheet_name(root_id, used_names)
                flat_df.to_excel(writer, sheet_name=out_sheet, index=False)
                preview[out_sheet] = flat_df

                if not legend_df.empty:
                    legend_sheet = sanitize_sheet_name(f"{root_id}_legend", used_names)
                    legend_df.to_excel(writer, sheet_name=legend_sheet, index=False)
                else:
                    legend_sheet = ""

                summary_rows.append(
                    {
                        "ROOT_ID": root_id,
                        "DESCRIPTION": root_desc,
                        "SOURCE_SHEET": sheet_name,
                        "UNIQUE_PARTS": len(flat_df),
                        "OUTPUT_SHEET": out_sheet,
                        "LEGEND_SHEET": legend_sheet,
                    }
                )

        summary_df = pd.DataFrame(summary_rows)
        summary_sheet = sanitize_sheet_name("Summary", used_names)
        summary_df.to_excel(writer, sheet_name=summary_sheet, index=False)

    buf.seek(0)
    wb = load_workbook(buf)
    if "Summary" in wb.sheetnames:
        wb.move_sheet("Summary", offset=-len(wb.sheetnames) + 1)
    wb = format_workbook(wb)

    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)

    return out_buf.getvalue(), summary_rows, preview


# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------

st.title("\U0001F5C2\uFE0F BOM Tree Flattener")
st.caption(
    "Upload a multi-level BOM tree export (LEVEL / ID / NEXT_ASSY columns). "
    "Each level-0 root becomes its own output sheet, with every part appearing "
    "once and its QPA per parent assembly shown in separate columns."
)

with st.sidebar:
    st.header("Settings")
    qty_col = st.selectbox(
        "QPA source column",
        options=["ORIG_QTY", "REQ_QTY"],
        index=0,
        help=(
            "ORIG_QTY = nominal designed BOM quantity per assembly (recommended). "
            "REQ_QTY = point-in-time MRP net requirement, net of stock/orders."
        ),
    )

uploaded = st.file_uploader("Upload BOM tree file (.xls / .xlsx)", type=["xls", "xlsx"])

if uploaded is not None:
    with st.spinner("Flattening BOM tree..."):
        try:
            output_bytes, summary_rows, preview = flatten_bom_tree(
                uploaded.getvalue(), uploaded.name, qty_col=qty_col
            )
        except Exception as e:
            st.error(f"Failed to process file: {e}")
            st.stop()

    if not summary_rows:
        st.error("No valid LEVEL=0 root products were found in this file.")
        st.stop()

    st.success(f"Done — {len(summary_rows)} root product(s) flattened.")
    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)

    out_name = uploaded.name.rsplit(".", 1)[0] + "_flattened.xlsx"
    st.download_button(
        "\u2b07\uFE0F Download flattened workbook",
        data=output_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.subheader("Preview")
    sheet_choice = st.selectbox("Choose a root product to preview", options=list(preview.keys()))
    if sheet_choice:
        st.dataframe(preview[sheet_choice], use_container_width=True)
else:
    st.info("Upload a file to get started.")
