"""
Streamlit app that:
- Accepts an uploaded .xlsx file
- Runs the Excel extraction/join/clean logic
- Produces output files (XLSX/CSV/Auth debug optional)
- Lets users download each file directly in the UI

Run:
  pip install -r requirements.txt
  streamlit run app_v2.py
"""

import os
import re
import tempfile
from datetime import datetime, date
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook
import streamlit as st


# ============================================================
# 1) Helpers (normalize, empty checks, date parsing)
# ============================================================

def norm(v) -> str:
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).strip())


def clean_person_name(v) -> str:
    return norm(v).replace("*", "")


def is_empty(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def try_parse_date(v):
    if isinstance(v, (datetime, date)):
        return pd.to_datetime(v).normalize()

    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None
        ts = pd.to_datetime(s, errors="coerce")
        return None if pd.isna(ts) else ts.normalize()

    return None


def to_number_or_none(v):
    n = pd.to_numeric(pd.Series([v]), errors="coerce").iloc[0]
    return None if pd.isna(n) else n


def build_grid(ws):
    max_row, max_col = ws.max_row, ws.max_column
    return [
        [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
        for r in range(1, max_row + 1)
    ]


def row_has_date_label(row) -> bool:
    return any(norm(v).lower() == "date" for v in row)


def count_dates_in_row(row) -> int:
    return sum(1 for v in row if try_parse_date(v) is not None)


# ============================================================
# 2) Sheet 1 extractor: Initial Forms blocks (Sent/Signed)
# ============================================================

def extract_initial_forms_blocks(
    wb,
    sheet_name: str,
    min_dates_in_group_row: int = 2,
    stop_blank_streak: int = 2,
) -> pd.DataFrame:
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")

    ws = wb[sheet_name]
    grid = build_grid(ws)

    records = []
    r = 0
    block_index = 0

    while r < len(grid):
        if not row_has_date_label(grid[r]):
            r += 1
            continue

        if count_dates_in_row(grid[r]) < min_dates_in_group_row:
            r += 1
            continue

        title_row_idx = r - 1 if r - 1 >= 0 else None
        group_row_idx = r
        sub_row_idx = r + 1
        if sub_row_idx >= len(grid):
            break

        title_row = grid[title_row_idx] if title_row_idx is not None else []
        group_row = grid[group_row_idx]
        sub_row = grid[sub_row_idx]

        week_label = next((norm(v) for v in title_row if not is_empty(v)), "")

        try:
            date_label_col = next(i for i, v in enumerate(group_row) if norm(v).lower() == "date")
        except StopIteration:
            r += 1
            continue

        first_data_col = date_label_col + 1

        date_groups = []
        c = first_data_col
        while c < len(group_row):
            if norm(group_row[c]).lower() == "total":
                break

            g_date = try_parse_date(group_row[c])
            if g_date is not None:
                if (
                    norm(sub_row[c]).lower() == "sent"
                    and c + 1 < len(sub_row)
                    and norm(sub_row[c + 1]).lower() == "signed"
                ):
                    date_groups.append((g_date, c, c + 1))
                    c += 2
                    continue
            c += 1

        rr = sub_row_idx + 1
        blank_streak = 0

        while rr < len(grid):
            row = grid[rr]

            if row_has_date_label(row) and count_dates_in_row(row) >= min_dates_in_group_row:
                break

            non_empty = sum(0 if is_empty(v) else 1 for v in row)
            if non_empty <= 1:
                blank_streak += 1
                if blank_streak >= stop_blank_streak:
                    break
            else:
                blank_streak = 0
                name = clean_person_name(row[date_label_col]) if date_label_col < len(row) else ""
                if name and not name.lower().startswith("total"):
                    for d, sent_c, signed_c in date_groups:
                        records.append({
                            "source_sheet": sheet_name,
                            "block_index": block_index,
                            "week_label": week_label,
                            "Welcome Specialist": name,
                            "Date": d,
                            "Sent": row[sent_c] if sent_c < len(row) else None,
                            "Signed": row[signed_c] if signed_c < len(row) else None,
                        })

            rr += 1

        block_index += 1
        r = rr

    return pd.DataFrame(records)


# ============================================================
# 3) Availability: move OFF/LEAVE-like strings to availability
# ============================================================

def extract_availability(df: pd.DataFrame) -> pd.DataFrame:
    def is_string_value(v):
        return isinstance(v, str) and v.strip() != ""

    df = df.copy()
    df["availability"] = None

    for idx, row in df.iterrows():
        sent_val = row.get("Sent")
        signed_val = row.get("Signed")

        availability_val = None
        if is_string_value(signed_val):
            availability_val = signed_val
        elif is_string_value(sent_val):
            availability_val = sent_val

        if availability_val is not None:
            df.at[idx, "availability"] = availability_val
            df.at[idx, "Sent"] = None
            df.at[idx, "Signed"] = None

    return df


# ============================================================
# 4) Sheet 2 extractor: Auth Report blocks (Received)
# ============================================================

def find_best_metric_header_row(grid, group_row_idx: int, metric_key: str, header_search_depth: int):
    best_hdr_row_idx = None
    best_hits = 0
    for k in range(1, header_search_depth + 1):
        cand_idx = group_row_idx + k
        if cand_idx >= len(grid):
            break
        cand = grid[cand_idx]
        hits = sum(1 for v in cand if metric_key in norm(v).lower())
        if hits > best_hits:
            best_hits = hits
            best_hdr_row_idx = cand_idx
    return best_hdr_row_idx, best_hits


def build_metric_date_map(group_row, metric_hdr_row, first_data_col: int, metric_key: str):
    date_ff = []
    last_date = None
    for c in range(len(group_row)):
        d = try_parse_date(group_row[c])
        if d is not None:
            last_date = d
        date_ff.append(last_date)

    date_to_col = {}
    c = first_data_col
    while c < len(group_row):
        if norm(group_row[c]).lower() == "total":
            break

        d = date_ff[c]
        if d is None:
            c += 1
            continue

        hdr = norm(metric_hdr_row[c]).lower()
        if metric_key in hdr:
            date_to_col[d] = c

        c += 1

    return date_to_col


def build_combined_date_map(group_row, category_hdr_row, first_data_col: int, combine_headers: set[str]):
    date_to_cols = {}
    current_date = None
    c = first_data_col

    while c < len(group_row):
        if norm(group_row[c]).lower() == "total":
            break

        maybe_date = try_parse_date(group_row[c])
        if maybe_date is not None:
            current_date = maybe_date

        if current_date is not None:
            hdr = norm(category_hdr_row[c]).lower()
            if hdr in combine_headers:
                date_to_cols.setdefault(current_date, []).append(c)

        c += 1

    return date_to_cols


def combine_row_values(row, col_indexes):
    total = 0
    found_numeric = False

    for col_idx in col_indexes:
        if col_idx >= len(row):
            continue
        num = to_number_or_none(row[col_idx])
        if num is not None:
            total += num
            found_numeric = True

    return total if found_numeric else None


def extract_combined_header_values(row, column_map, combine_headers):
    values = {}
    total = 0
    found_any = False

    for header in combine_headers:
        col_indexes = column_map.get(header, [])
        value = combine_row_values(row, col_indexes)
        values[header] = value
        if value is not None:
            total += value
            found_any = True

    values["total"] = total if found_any else None
    return values


def extract_auth_received_blocks(
    wb,
    sheet_name: str,
    metric_name: str = "Received",
    name_header: str = "Welcome Specialist",
    header_search_depth: int = 3,
    min_dates_in_group_row: int = 2,
    stop_blank_streak: int = 2,
    combine_headers: tuple[str, ...] = ("PMC", "Spr"),
    debug: bool = False,
) -> pd.DataFrame:
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")

    ws = wb[sheet_name]
    grid = build_grid(ws)

    metric_key = norm(metric_name).lower()
    name_key = norm(name_header).lower()
    normalized_combine_headers = {norm(v).lower(): norm(v) for v in combine_headers}

    records = []
    r = 0
    block_index = 0

    while r < len(grid):
        if not row_has_date_label(grid[r]) or count_dates_in_row(grid[r]) < min_dates_in_group_row:
            r += 1
            continue

        title_row_idx = r - 1 if r - 1 >= 0 else None
        group_row_idx = r

        title_row = grid[title_row_idx] if title_row_idx is not None else []
        group_row = grid[group_row_idx]
        week_label = next((norm(v) for v in title_row if not is_empty(v)), "")

        try:
            date_label_col = next(i for i, v in enumerate(group_row) if norm(v).lower() == "date")
        except StopIteration:
            r += 1
            continue

        first_data_col = date_label_col + 1

        best_hdr_row_idx, best_hits = find_best_metric_header_row(
            grid, group_row_idx, metric_key, header_search_depth
        )

        value_mode = "metric"
        date_map = {}

        if best_hdr_row_idx is not None and best_hits > 0:
            metric_hdr_row = grid[best_hdr_row_idx]
            date_map = build_metric_date_map(group_row, metric_hdr_row, first_data_col, metric_key)
        else:
            category_hdr_idx = group_row_idx + 1
            if category_hdr_idx < len(grid):
                category_hdr_row = grid[category_hdr_idx]
                date_map = build_combined_date_map(
                    group_row,
                    category_hdr_row,
                    first_data_col,
                    set(normalized_combine_headers.keys()),
                )
                if date_map:
                    best_hdr_row_idx = category_hdr_idx
                    value_mode = "combined_headers"

        if best_hdr_row_idx is None or not date_map:
            if debug:
                print(
                    f"[DEBUG] No auth mapping found for sheet '{sheet_name}' "
                    f"using metric '{metric_name}' or combined headers {sorted(normalized_combine_headers.keys())}."
                )
            r += 1
            continue

        header_row = grid[best_hdr_row_idx]

        name_col = None
        for i, v in enumerate(header_row):
            if norm(v).lower() == name_key:
                name_col = i
                break
        if name_col is None:
            name_col = date_label_col

        rr = best_hdr_row_idx + 1
        blank_streak = 0

        while rr < len(grid):
            row = grid[rr]

            if row_has_date_label(row) and count_dates_in_row(row) >= min_dates_in_group_row:
                break

            non_empty = sum(0 if is_empty(v) else 1 for v in row)
            if non_empty <= 1:
                blank_streak += 1
                if blank_streak >= stop_blank_streak:
                    break
            else:
                blank_streak = 0
                name = clean_person_name(row[name_col]) if name_col < len(row) else ""
                if name and not name.lower().startswith("total"):
                    for d, col_info in date_map.items():
                        if value_mode == "metric":
                            value = row[col_info] if col_info < len(row) else None
                            record = {
                                "source_sheet": sheet_name,
                                "block_index": block_index,
                                "week_label": week_label,
                                "Welcome Specialist": name,
                                "Date": d,
                                metric_name: value,
                            }
                        else:
                            by_header = {}
                            for col_idx in col_info:
                                header_name = norm(header_row[col_idx]).lower() if col_idx < len(header_row) else ""
                                if header_name:
                                    by_header.setdefault(header_name, []).append(col_idx)

                            extracted = extract_combined_header_values(
                                row,
                                by_header,
                                list(normalized_combine_headers.keys()),
                            )
                            record = {
                                "source_sheet": sheet_name,
                                "block_index": block_index,
                                "week_label": week_label,
                                "Welcome Specialist": name,
                                "Date": d,
                                metric_name: extracted["total"],
                            }
                            for header_key, original_name in normalized_combine_headers.items():
                                record[original_name] = extracted[header_key]

                        records.append(record)

            rr += 1

        block_index += 1
        r = rr

    return pd.DataFrame(records)


# ============================================================
# 5) Join + clean
# ============================================================

def attach_auth_to_initial(initial_df: pd.DataFrame, auth_df: pd.DataFrame, metric_name: str = "Received") -> pd.DataFrame:
    out = initial_df.copy()
    if auth_df.empty:
        out[metric_name] = None
        if "PMC" in auth_df.columns:
            out["PMC"] = None
        if "Spr" in auth_df.columns:
            out["Spr"] = None
        return out

    a = auth_df.copy()
    a["Date"] = pd.to_datetime(a["Date"], errors="coerce").dt.normalize()
    out["Date"] = pd.to_datetime(out["Date"], errors="coerce").dt.normalize()

    def agg_series(s: pd.Series):
        numeric = pd.to_numeric(s, errors="coerce")
        if numeric.notna().any():
            return numeric.sum()
        return s.dropna().iloc[0] if s.dropna().shape[0] > 0 else None

    agg_map = {metric_name: agg_series}
    for extra_col in ("PMC", "Spr"):
        if extra_col in a.columns:
            agg_map[extra_col] = agg_series

    a_keyed = a.groupby(["Welcome Specialist", "Date"], as_index=False).agg(agg_map)
    return out.merge(a_keyed, on=["Welcome Specialist", "Date"], how="left")


def clean_received_column(df: pd.DataFrame, column_name: str = "Received") -> pd.DataFrame:
    df = df.copy()
    if column_name in df.columns:
        df[column_name] = pd.to_numeric(df[column_name], errors="coerce")
    return df


def clean_welcome_specialist_column(df: pd.DataFrame, column_name: str = "Welcome Specialist") -> pd.DataFrame:
    df = df.copy()
    if column_name in df.columns:
        df[column_name] = df[column_name].apply(clean_person_name)
    return df


def process_workbook(excel_path: str, initial_sheet: str, auth_sheet: str, auth_metric_name: str):
    wb = load_workbook(excel_path, data_only=True)

    df_initial = extract_initial_forms_blocks(wb, initial_sheet)
    if df_initial.empty:
        raise ValueError("No blocks found in Initial Forms sheet (could not detect 'Date' row with multiple dates).")

    df_initial = extract_availability(df_initial)
    df_initial = clean_welcome_specialist_column(df_initial)

    df_auth = extract_auth_received_blocks(
        wb,
        sheet_name=auth_sheet,
        metric_name=auth_metric_name,
        debug=False
    )
    df_auth = clean_welcome_specialist_column(df_auth)

    df_final = attach_auth_to_initial(df_initial, df_auth, metric_name=auth_metric_name)
    df_final = clean_received_column(df_final, column_name=auth_metric_name)
    df_final = clean_welcome_specialist_column(df_final)

    return df_final, df_auth


# ============================================================
# 6) Streamlit UI
# ============================================================

ALLOWED_EXTENSIONS = {"xlsx"}


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def dataframe_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    return buffer.getvalue()


def render_streamlit_app() -> None:
    st.set_page_config(
        page_title="Welcome Dashboard Processor",
        page_icon=":bar_chart:",
        layout="wide",
    )

    st.title("Welcome Dashboard Processor")
    st.write(
        "Upload your source workbook, choose the source sheet names, and generate "
        "a cleaned output with the auth metric matched by specialist and date."
    )

    upload_col, settings_col = st.columns([1.3, 1.0], gap="large")

    with upload_col:
        st.subheader("Upload Workbook")
        uploaded_file = st.file_uploader(
            "Choose a single Excel workbook",
            type=["xlsx"],
            help="The app reads the initial forms block, extracts the auth metric, and builds a merged output.",
        )

    with settings_col:
        st.subheader("Settings")
        initial_sheet = st.text_input("Initial Forms sheet name", value="January Weekly + MTD")
        auth_sheet = st.text_input("Auth sheet name", value="Auth Report Weekly + MTD")
        auth_metric = st.text_input("Auth metric label to extract", value="Received")
        include_auth_debug = st.checkbox(
            "Generate auth debug workbook",
            value=True,
            help="Useful when validating that the Auth extraction found the correct rows and dates.",
        )
        st.caption("Input: .xlsx | Output: .xlsx, .csv, optional auth debug .xlsx")

    st.info(
        "Matching is performed on cleaned `Welcome Specialist` values and normalized `Date` "
        "values. Text entries like OFF/LEAVE are moved into `availability`."
    )

    process_clicked = st.button("Process workbook", type="primary", use_container_width=True)

    if not process_clicked:
        return

    if uploaded_file is None:
        st.error("Please upload an .xlsx file.")
        return

    if not allowed_file(uploaded_file.name):
        st.error("Please upload an .xlsx file.")
        return

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.getbuffer())
        input_path = tmp.name

    try:
        with st.spinner("Processing workbook..."):
            df_final, df_auth = process_workbook(
                input_path,
                initial_sheet=initial_sheet,
                auth_sheet=auth_sheet,
                auth_metric_name=auth_metric,
            )
    except Exception as e:
        st.error(str(e))
        return
    finally:
        if os.path.exists(input_path):
            os.unlink(input_path)

    final_xlsx_bytes = dataframe_to_xlsx_bytes(df_final)
    final_csv_bytes = df_final.to_csv(index=False).encode("utf-8")
    auth_xlsx_bytes = dataframe_to_xlsx_bytes(df_auth) if include_auth_debug else None

    st.success("Your files are ready. Download them below.")

    metric_col1, metric_col2, metric_col3 = st.columns(3)
    metric_col1.metric("Merged rows", len(df_final))
    metric_col2.metric("Auth rows", len(df_auth))
    metric_col3.metric("Auth metric", auth_metric)

    st.subheader("Downloads")
    download_col1, download_col2, download_col3 = st.columns(3)

    with download_col1:
        st.download_button(
            "Download XLSX",
            data=final_xlsx_bytes,
            file_name="structured_output_with_auth4.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with download_col2:
        st.download_button(
            "Download CSV",
            data=final_csv_bytes,
            file_name="structured_output_with_auth4.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with download_col3:
        if auth_xlsx_bytes is not None:
            st.download_button(
                "Download Auth Debug",
                data=auth_xlsx_bytes,
                file_name="auth_extracted_debug.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.caption("Auth debug workbook not generated.")

    with st.expander("Preview merged output"):
        st.dataframe(df_final, use_container_width=True)

    if include_auth_debug:
        with st.expander("Preview auth debug output"):
            st.dataframe(df_auth, use_container_width=True)


render_streamlit_app()