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
REQUIRED_COLS = {"LEVEL", "ID"}

# Different ERP exports spell the same column differently. Map known
# variants to the canonical name the rest of the app expects.
HEADER_ALIASES = {
    "DESC": "DESCRIPTION",
    "DESC.": "DESCRIPTION",
    "DESCR": "DESCRIPTION",
    "PART_DESCRIPTION": "DESCRIPTION",
    "PARENT": "NEXT_ASSY",
    "PARENT_ID": "NEXT_ASSY",
    "NEXTASSY": "NEXT_ASSY",
    "NEXT_ASSEMBLY": "NEXT_ASSY",
    "PART_NUMBER": "ID",
    "PART_NO": "ID",
    "PN": "ID",
    "ITEM_ID": "ID",
    "QTY": "ORIG_QTY",
    "QUANTITY": "ORIG_QTY",
}


def canonicalize_header(h) -> str:
    """Normalize a raw header cell to the canonical column name we expect."""
    raw = str(h).strip()
    stripped = raw.rstrip(".").strip()
    key = re.sub(r"\s+", "_", stripped.upper())
    return HEADER_ALIASES.get(key, key)


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
    header = [canonicalize_header(h) for h in raw.iloc[0].tolist()]
    data = raw.iloc[1:].copy()
    data.columns = header
    data = data.reset_index(drop=True)
    data = data.dropna(how="all")
    # If a canonicalization collision produced duplicate column names, keep
    # the first occurrence of each so downstream selection stays unambiguous.
    data = data.loc[:, ~data.columns.duplicated()]
    return data


def parse_tree_nodes(block: pd.DataFrame, qty_col: str):
    """
    Parse the raw indented BOM export into an actual tree of row-nodes,
    using LEVEL + row order (the same signal any indented-tree viewer
    uses) rather than matching by ID text. This means a sub-assembly used
    under two different parents becomes two distinct node objects — each
    with its own children and its own quantities exactly as the source
    file lists them, with nothing summed or blended together.
    """
    block = block.reset_index(drop=True)
    nodes = []
    stack = []  # (level, node_idx)
    for i in range(len(block)):
        row = block.iloc[i]
        lvl_raw = pd.to_numeric(row.get("LEVEL"), errors="coerce")
        if pd.isna(lvl_raw):
            continue
        lvl = int(lvl_raw)
        while stack and stack[-1][0] >= lvl:
            stack.pop()
        parent_idx = stack[-1][1] if stack else None
        qty_val = pd.to_numeric(row.get(qty_col), errors="coerce") if qty_col in row.index else None
        seq_val = pd.to_numeric(row.get("SEQ"), errors="coerce") if "SEQ" in row.index else None
        node = {
            "idx": len(nodes),
            "id": str(row.get("ID")),
            "level": lvl,
            "parent_idx": parent_idx,
            "qty": qty_val,
            "seq": seq_val,
            "row": row,
        }
        nodes.append(node)
        stack.append((lvl, node["idx"]))
    return nodes


def flatten_single_tree(block: pd.DataFrame, qty_col: str):
    if qty_col not in block.columns:
        raise ValueError(f"Quantity column '{qty_col}' not found in source data.")

    nodes = parse_tree_nodes(block, qty_col)
    if not nodes:
        raise ValueError("No valid LEVEL/ID rows found in this tree block.")

    children_of = {}
    for n in nodes:
        if n["parent_idx"] is not None:
            children_of.setdefault(n["parent_idx"], []).append(n["idx"])
    for parent_idx, kids in children_of.items():
        children_of[parent_idx] = sorted(
            kids, key=lambda idx: (nodes[idx]["seq"] if pd.notna(nodes[idx]["seq"]) else 0, nodes[idx]["id"])
        )

    assembly_idxs = list(children_of.keys())  # nodes that have >=1 child

    # How many distinct occurrence-nodes share the same assembly ID? Only
    # those need a disambiguating column label; a uniquely-placed assembly
    # keeps the simple "QPA_in_<id> (L<level>)" label as before.
    id_occurrence_count = {}
    for idx in assembly_idxs:
        aid = nodes[idx]["id"]
        id_occurrence_count[aid] = id_occurrence_count.get(aid, 0) + 1

    labels, used_labels = {}, set()
    for idx in assembly_idxs:
        n = nodes[idx]
        if id_occurrence_count[n["id"]] == 1:
            label = f"QPA_in_{n['id']} (L{n['level']})"
        else:
            parent_id = nodes[n["parent_idx"]]["id"] if n["parent_idx"] is not None else "ROOT"
            label = f"QPA_in_{n['id']} [under {parent_id}] (L{n['level']})"
        final = label
        i = 2
        while final in used_labels:
            final = f"{label} #{i}"
            i += 1
        used_labels.add(final)
        labels[idx] = final

    # Depth-first, pre-order walk of the REAL tree: root, then each direct
    # child assembly immediately followed by its own children, recursively.
    order_idx = []

    def dfs(node_idx):
        if node_idx in children_of:
            order_idx.append(node_idx)
        for c in children_of.get(node_idx, []):
            dfs(c)

    dfs(0)  # nodes[0] is always this block's LEVEL-0 root

    # Per-occurrence child quantities — no cross-occurrence summing. If the
    # exact same child ID appears twice under the EXACT same occurrence
    # (duplicate SEQ under one specific parent instance), those are summed;
    # different occurrences of the parent are never mixed together.
    child_qty = {}
    for idx in assembly_idxs:
        d = {}
        for c_idx in children_of[idx]:
            c = nodes[c_idx]
            q = c["qty"] if c["qty"] is not None and pd.notna(c["qty"]) else 0
            d[c["id"]] = d.get(c["id"], 0) + q
        child_qty[idx] = d

    info_cols = [c for c in INFO_COLS_PRIORITY if c in block.columns]
    part_first_row = {}
    for n in nodes:
        if n["level"] > 0 and n["id"] not in part_first_row:
            part_first_row[n["id"]] = n["row"]

    all_part_ids = sorted(part_first_row.keys())
    data_rows = []
    for pid in all_part_ids:
        row = part_first_row[pid]
        rec = {c: row.get(c) for c in info_cols}
        rec["ID"] = pid  # always the normalized string key, never the raw (possibly numeric) cell
        data_rows.append(rec)
    flat = pd.DataFrame(data_rows)

    usage_count = [sum(1 for idx in order_idx if pid in child_qty[idx]) for pid in all_part_ids]
    flat["USED_IN_N_ASSEMBLIES"] = usage_count

    for idx in order_idx:
        d = child_qty[idx]
        flat[labels[idx]] = [d.get(pid, None) for pid in all_part_ids]

    flat = flat.sort_values("ID").reset_index(drop=True)

    legend_rows = []
    for idx in order_idx:
        n = nodes[idx]
        row = n["row"]
        parent_id = nodes[n["parent_idx"]]["id"] if n["parent_idx"] is not None else ""
        legend_rows.append(
            {
                "LEVEL": n["level"],
                "ASSEMBLY_ID": n["id"],
                "DESCRIPTION": row.get("DESCRIPTION", ""),
                "REV": row.get("REV", ""),
                "PARENT_ID": parent_id,
                "COLUMN_LABEL": labels[idx],
            }
        )
    legend = pd.DataFrame(legend_rows)

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
