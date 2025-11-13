
import re
from io import BytesIO
from difflib import SequenceMatcher
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# -----------------------------
# Utilities
# -----------------------------
def _norm_key(s: str) -> str:
    """Normalize keys for robust matching:
    - lowercase
    - & -> and
    - remove non-alphanumerics
    """
    if s is None:
        return ""
    s = str(s).strip().lower().replace("&", "and")
    return re.sub(r"[^a-z0-9]+", "", s)

# -----------------------------
# Page config & constants
# -----------------------------
st.set_page_config(page_title="WALMART Masterfile Filler — Simple Mapping", layout="wide")

TARGET_SHEET_TITLES = [
    "Product Content And Site Exp",
    "Trade Item Configurations",
]
# Canonical (normalized) -> display title
TARGET_SHEETS_CANON: Dict[str, str] = { _norm_key(t): t for t in TARGET_SHEET_TITLES }

HEADER_ROW_INDEX = 6      # mapping is based on this row
DATA_START_ROW = 7        # write data from here (row 7)

# Excel highlight fill for duplicates
YELLOW_DUP_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# -----------------------------
# Helper functions
# -----------------------------
def _enforce_extension(filename: str, is_xlsm: bool) -> str:
    if not filename:
        return "filled_template.xlsm" if is_xlsm else "filled_template.xlsx"
    filename = filename.strip()
    if is_xlsm and not filename.lower().endswith(".xlsm"):
        filename = re.sub(r"\.(xlsx|xls)$", "", filename, flags=re.IGNORECASE) + ".xlsm"
    if not is_xlsm and not filename.lower().endswith(".xlsx"):
        filename = re.sub(r"\.(xlsm|xls)$", "", filename, flags=re.IGNORECASE) + ".xlsx"
    return filename

def _find_target_sheets(actual_names: List[str]) -> Dict[str, str]:
    """Return mapping canon_key (normalized) -> actual_name for present target sheets, tolerant to minor name variations."""
    present = {}
    norm_actual = { _norm_key(n): n for n in actual_names }
    for canon_norm, display in TARGET_SHEETS_CANON.items():
        # Exact normalized match
        if canon_norm in norm_actual:
            present[canon_norm] = norm_actual[canon_norm]
            continue
        # Fallback: partial contains either way
        for n_norm, real in norm_actual.items():
            if canon_norm in n_norm or n_norm in canon_norm:
                present[canon_norm] = real
                break
    return present

def _extract_headers_row(ws, header_row: int) -> Tuple[List[str], Dict[str, int]]:
    """Read the header row into list and normalized->col_index map (1-based col). Empty headers ignored."""
    headers: List[str] = []
    norm_to_col: Dict[str, int] = {}
    for col_idx, cell in enumerate(ws[header_row], start=1):
        val = cell.value
        if val is None or str(val).strip() == "":
            continue
        val_str = str(val).strip()
        headers.append(val_str)
        norm_to_col[_norm_key(val_str)] = col_idx
    return headers, norm_to_col

def _touch_tick(sheet_key: str, templ_norm: str):
    st.session_state["last_edit_tick"] = st.session_state.get("last_edit_tick", 0) + 1
    st.session_state[f"tick_{sheet_key}_{templ_norm}"] = st.session_state["last_edit_tick"]

def _make_on_change(sheet_key: str, templ_norm: str):
    def _cb():
        _touch_tick(sheet_key, templ_norm)
    return _cb

def _build_mapping_for_sheet(sheet_key: str, templ_headers: List[str]) -> List[Dict[str, str]]:
    """Gather mapping from session_state for given sheet; skip blanks."""
    records = []
    for t in templ_headers:
        templ_norm = _norm_key(t)
        k = f"map_{sheet_key}_{templ_norm}"
        raw = st.session_state.get(k, "") or ""
        if raw:
            records.append({
                "template_header": t,
                "template_norm": templ_norm,
                "raw_header": raw,
                "tick": st.session_state.get(f"tick_{sheet_key}_{templ_norm}", 0),
            })
    return records

def _auto_map_exact(sheet_key: str, templ_headers: List[str], raw_headers: List[str]):
    raw_norm_index = { _norm_key(r): r for r in raw_headers }
    for t in templ_headers:
        templ_norm = _norm_key(t)
        k = f"map_{sheet_key}_{templ_norm}"
        if not st.session_state.get(k, ""):
            match = raw_norm_index.get(templ_norm)
            if match:
                st.session_state[k] = match
                _touch_tick(sheet_key, templ_norm)

def _auto_map_fuzzy(sheet_key: str, templ_headers: List[str], raw_headers: List[str], threshold: int = 80):
    thr = max(0, min(100, threshold)) / 100.0
    raw_norms = [(r, _norm_key(r)) for r in raw_headers]
    for t in templ_headers:
        templ_norm = _norm_key(t)
        k = f"map_{sheet_key}_{templ_norm}"
        if st.session_state.get(k, ""):
            continue  # fill blanks only
        best_raw = None
        best_score = 0.0
        for r, rnorm in raw_norms:
            score = SequenceMatcher(None, templ_norm, rnorm).ratio()
            if score > best_score:
                best_score, best_raw = score, r
        if best_raw and best_score >= thr:
            st.session_state[k] = best_raw
            _touch_tick(sheet_key, templ_norm)

def _import_mapping(df: pd.DataFrame, apply_sheet_key: Optional[str],
                    present_sheets: Dict[str, str],
                    templ_headers_by_sheet: Dict[str, List[str]],
                    raw_headers: List[str]) -> int:
    """Import mapping rows.
    Accept columns like Template/Raw (case-insensitive). Optional `Sheet` to target specific sheet.
    Only applies rows whose headers exist and raw headers exist.
    """
    cols = { _norm_key(c): c for c in df.columns }
    # find template col
    tcol = None
    for candidate in ["template", "templateheader", "templ", "target", "targetheader"]:
        if candidate in cols:
            tcol = cols[candidate]; break
    rcol = None
    for candidate in ["raw", "rawheader", "source", "sourceheader"]:
        if candidate in cols:
            rcol = cols[candidate]; break
    scol = None
    for candidate in ["sheet", "tab", "worksheet"]:
        if candidate in cols:
            scol = cols[candidate]; break
    if not tcol or not rcol:
        return 0

    raw_set = set(raw_headers)
    count = 0
    for _, row in df.iterrows():
        templ = str(row[tcol]).strip() if pd.notna(row[tcol]) else ""
        raw = str(row[rcol]).strip() if pd.notna(row[rcol]) else ""
        if not templ or not raw or raw not in raw_set:
            continue

        target_sheet_keys: List[str]
        if scol and pd.notna(row[scol]):
            s = str(row[scol]).strip()
            norm_s = _norm_key(s)
            # map human provided to one of present_sheets keys
            target_sheet_keys = []
            for sk in present_sheets.keys():
                if norm_s == sk or norm_s in sk or sk in norm_s:
                    target_sheet_keys.append(sk)
            if not target_sheet_keys:
                continue
        else:
            # no sheet specified -> apply to current selected only
            target_sheet_keys = [apply_sheet_key] if apply_sheet_key else list(present_sheets.keys())

        for sk in target_sheet_keys:
            # only if template header exists in that sheet
            templ_set = set(templ_headers_by_sheet.get(sk, []))
            if templ in templ_set:
                map_key = f"map_{sk}_{_norm_key(templ)}"
                st.session_state[map_key] = raw
                _touch_tick(sk, _norm_key(templ))
                count += 1
    return count

def _export_mapping_df(present_sheets: Dict[str, str],
                       templ_headers_by_sheet: Dict[str, List[str]]) -> pd.DataFrame:
    rows = []
    for sk, actual_name in present_sheets.items():
        templ_list = templ_headers_by_sheet.get(sk, [])
        for t in templ_list:
            v = st.session_state.get(f"map_{sk}_{_norm_key(t)}", "")
            rows.append({"Sheet": actual_name, "Template": t, "Raw": v})
    return pd.DataFrame(rows)

def _resolve_duplicate_raw_mappings(records: List[Dict[str, str]], auto_resolve: bool) -> Tuple[List[Dict[str, str]], List[str]]:
    """Detect duplicates in selected raw headers across template rows."""
    bucket: Dict[str, List[Dict[str, str]]] = {}
    for r in records:
        bucket.setdefault(r["raw_header"], []).append(r)

    duplicates = [raw for raw, lst in bucket.items() if len(lst) > 1]
    if not duplicates:
        return records, []

    if auto_resolve:
        resolved = []
        for raw, lst in bucket.items():
            if len(lst) == 1:
                resolved.append(lst[0])
            else:
                # Keep the one with highest tick (latest edit wins)
                keep = sorted(lst, key=lambda x: x.get("tick", 0))[-1]
                resolved.append(keep)
        return resolved, []
    else:
        return records, duplicates

def _write_sheet_data(ws, mapping: List[Dict[str, str]],
                      header_row: int,
                      start_row: int,
                      raw_df: pd.DataFrame,
                      dup_headers_to_highlight: List[str]) -> Tuple[int, List[str]]:
    """Write mapped data into ws and highlight duplicates for requested columns."""
    # Build column index map from header row
    _, norm_to_col = _extract_headers_row(ws, header_row)
    # Filter mapping to only headers present in the worksheet
    to_write = []
    missing_template_headers = []
    for m in mapping:
        templ_norm = _norm_key(m["template_header"])
        if templ_norm in norm_to_col:
            to_write.append((norm_to_col[templ_norm], m["raw_header"]))
        else:
            missing_template_headers.append(m["template_header"])

    # Write data
    nrows = len(raw_df)
    for i in range(nrows):
        excel_row = start_row + i
        for col_idx, raw_name in to_write:
            val = raw_df.iloc[i][raw_name] if raw_name in raw_df.columns else None
            ws.cell(row=excel_row, column=col_idx, value=val)

    # Highlight duplicates in requested columns (by header text in the sheet)
    dup_norms = [_norm_key(x) for x in dup_headers_to_highlight if str(x).strip()]
    for want_norm in dup_norms:
        if want_norm not in norm_to_col:
            continue
        cidx = norm_to_col[want_norm]
        # Build frequency map for values in this column (row start_row .. start_row + nrows - 1)
        counts: Dict[str, int] = {}
        for i in range(nrows):
            v = ws.cell(row=start_row + i, column=cidx).value
            if v is None or str(v).strip() == "":
                continue
            key = str(v)
            counts[key] = counts.get(key, 0) + 1
        dup_values = {k for k, c in counts.items() if c > 1}
        if not dup_values:
            continue
        # Apply fill
        for i in range(nrows):
            cell = ws.cell(row=start_row + i, column=cidx)
            v = cell.value
            if v is None:
                continue
            if str(v) in dup_values:
                cell.fill = YELLOW_DUP_FILL

    return nrows, missing_template_headers

def _build_output_bytes(template_bytes: bytes,
                        template_is_xlsm: bool,
                        present_sheets: Dict[str, str],
                        templ_headers_by_sheet: Dict[str, List[str]],
                        raw_df: pd.DataFrame,
                        auto_resolve_dupe_mappings: bool,
                        dup_columns_to_highlight: List[str]) -> bytes:
    bio_in = BytesIO(template_bytes)
    wb = load_workbook(bio_in, read_only=False, keep_vba=template_is_xlsm, data_only=False)

    for sk, actual_name in present_sheets.items():
        ws = wb[actual_name]
        mapping_records = _build_mapping_for_sheet(sk, templ_headers_by_sheet.get(sk, []))
        # Resolve duplicate RAW column selections at download time
        mapping_resolved, dups = _resolve_duplicate_raw_mappings(mapping_records, auto_resolve_dupe_mappings)
        if dups:
            dup_list = ", ".join(sorted(set(dups)))
            raise ValueError(
                f"Duplicate RAW column selections for sheet '{actual_name}': {dup_list}. "
                f"Auto-resolve is OFF; please resolve manually or turn it ON."
            )
        # Write
        _write_sheet_data(
            ws=ws,
            mapping=mapping_resolved,
            header_row=HEADER_ROW_INDEX,
            start_row=DATA_START_ROW,
            raw_df=raw_df,
            dup_headers_to_highlight=dup_columns_to_highlight,
        )

    bio_out = BytesIO()
    wb.save(bio_out)
    return bio_out.getvalue()

# -----------------------------
# Session bootstrap
# -----------------------------
if "template_bytes" not in st.session_state:
    st.session_state["template_bytes"] = None
if "template_ext" not in st.session_state:
    st.session_state["template_ext"] = None
if "present_sheets" not in st.session_state:
    st.session_state["present_sheets"] = {}  # canon_key -> actual name
if "templ_headers_by_sheet" not in st.session_state:
    st.session_state["templ_headers_by_sheet"] = {}  # canon_key -> [headers]
if "templ_header_sigs" not in st.session_state:
    st.session_state["templ_header_sigs"] = {}  # canon_key -> signature string
if "raw_headers" not in st.session_state:
    st.session_state["raw_headers"] = []
if "raw_sig" not in st.session_state:
    st.session_state["raw_sig"] = ""
if "raw_prev_headers" not in st.session_state:
    st.session_state["raw_prev_headers"] = []
if "current_sheet_key" not in st.session_state:
    st.session_state["current_sheet_key"] = None
if "last_edit_tick" not in st.session_state:
    st.session_state["last_edit_tick"] = 0
if "download_payload" not in st.session_state:
    st.session_state["download_payload"] = None

# -----------------------------
# UI: Scroll position retention (no jump to top on interaction)
# -----------------------------
st.components.v1.html(
    """
    <script>
    const KEY = "scrollY";
    window.addEventListener("beforeunload", () => {
      sessionStorage.setItem(KEY, window.scrollY);
    });
    window.addEventListener("load", () => {
      const y = sessionStorage.getItem(KEY);
      if (y !== null) window.scrollTo(0, parseFloat(y));
    });
    </script>
    """,
    height=0,
)

# -----------------------------
# Tabs
# -----------------------------
tab1, tab2, tab3 = st.tabs(["Upload Masterfile Template", "Upload Raw / Onboarding Data", "Simple Mapping & Download"])

# -----------------------------
# Tab 1: Upload Masterfile Template
# -----------------------------
with tab1:
    st.markdown("### 1) Upload Masterfile Template (.xlsx / .xlsm)")
    tpl = st.file_uploader("Upload template", type=["xlsx", "xlsm"], key="template_uploader")

    if tpl is not None:
        raw_bytes = tpl.read()
        is_xlsm = tpl.name.lower().endswith(".xlsm")
        # Load workbook and discover target sheets & their row-6 headers
        wb = load_workbook(BytesIO(raw_bytes), read_only=False, keep_vba=is_xlsm, data_only=False)
        all_names = wb.sheetnames
        present = _find_target_sheets(all_names)

        # Build headers per present sheet (from row 6)
        templ_headers_by_sheet: Dict[str, List[str]] = {}
        templ_header_sigs: Dict[str, str] = {}

        for sk, actual in present.items():
            ws = wb[actual]
            headers, _ = _extract_headers_row(ws, HEADER_ROW_INDEX)
            templ_headers_by_sheet[sk] = headers
            templ_header_sigs[sk] = "|".join(headers)

        # Save to session_state
        prev_sigs = st.session_state.get("templ_header_sigs", {})
        st.session_state["template_bytes"] = raw_bytes
        st.session_state["template_ext"] = "xlsm" if is_xlsm else "xlsx"
        st.session_state["present_sheets"] = present
        st.session_state["templ_headers_by_sheet"] = templ_headers_by_sheet
        st.session_state["templ_header_sigs"] = templ_header_sigs

        # Clear only mappings for sheets whose header signature changed
        for sk, sig in templ_header_sigs.items():
            if prev_sigs.get(sk, None) != sig:
                # clear mapping + tick keys for this sheet
                prefix_map = f"map_{sk}_"
                prefix_tick = f"tick_{sk}_"
                for k in list(st.session_state.keys()):
                    if k.startswith(prefix_map) or k.startswith(prefix_tick):
                        del st.session_state[k]

        # Default sheet selection
        if st.session_state["current_sheet_key"] not in present:
            st.session_state["current_sheet_key"] = next(iter(present.keys()), None)

        # Status
        st.success("Template loaded.")
        if present:
            st.write("Detected target sheets:")
            for sk, actual in present.items():
                display = TARGET_SHEETS_CANON.get(sk, actual)
                st.write(f"- **{display}** → actual name: `{actual}` "
                         f"(headers found in row {HEADER_ROW_INDEX}: {len(templ_headers_by_sheet.get(sk, []))})")
        else:
            st.warning("No target sheets found. Expected one or both of: "
                       "'Product Content And Site Exp', 'Trade Item Configurations'.")

# -----------------------------
# Tab 2: Upload Raw / Onboarding Data
# -----------------------------
with tab2:
    st.markdown("### 2) Upload Raw / Onboarding Data (.csv / .xlsx)")
    raw_file = st.file_uploader("Upload raw/onboarding file", type=["csv", "xlsx"], key="raw_uploader")

    raw_df = None
    if raw_file is not None:
        if raw_file.name.lower().endswith(".csv"):
            raw_df = pd.read_csv(raw_file)
        else:
            # Excel: let the user pick a sheet
            xl = pd.ExcelFile(raw_file)
            sheetname = st.selectbox("Choose a sheet", xl.sheet_names, index=0)
            raw_df = xl.parse(sheetname)

        # Normalize headers to strings
        raw_df.columns = [str(c) for c in raw_df.columns]
        new_headers = list(raw_df.columns)
        new_sig = "|".join(new_headers)

        # Update session; clear only selections that reference removed raw headers
        prev_headers = st.session_state.get("raw_prev_headers", [])
        removed = set(prev_headers) - set(new_headers)
        if removed:
            for sk in st.session_state.get("present_sheets", {}).keys():
                templ_headers = st.session_state.get("templ_headers_by_sheet", {}).get(sk, [])
                for t in templ_headers:
                    mk = f"map_{sk}_{_norm_key(t)}"
                    if st.session_state.get(mk, "") in removed:
                        st.session_state[mk] = ""

        st.session_state["raw_prev_headers"] = new_headers
        st.session_state["raw_headers"] = new_headers
        st.session_state["raw_sig"] = new_sig
        st.session_state["raw_df_payload"] = raw_df  # store DF for next tab

        st.success(f"Raw data loaded ({len(raw_df)} rows, {len(new_headers)} columns).")
        with st.expander("Preview (first 20 rows)"):
            st.dataframe(raw_df.head(20), use_container_width=True)

# -----------------------------
# Tab 3: Mapping & Download
# -----------------------------
with tab3:
    st.markdown("### 3) Simple Mapping & Download")

    # Guardrails
    if not st.session_state.get("template_bytes", None):
        st.info("Please upload a masterfile template in tab 1.")
        st.stop()
    if "raw_df_payload" not in st.session_state:
        st.info("Please upload raw/onboarding data in tab 2.")
        st.stop()

    present = st.session_state.get("present_sheets", {})
    templ_headers_by_sheet = st.session_state.get("templ_headers_by_sheet", {})
    raw_headers = st.session_state.get("raw_headers", [])
    raw_df: pd.DataFrame = st.session_state["raw_df_payload"]

    if not present:
        st.error("Template has no target sheets to map (expected PCSE and/or TIC).")
        st.stop()

    # Target sheet switcher (mapping applies to one sheet at a time)
    sheet_options = [(sk, TARGET_SHEETS_CANON.get(sk, st.session_state["present_sheets"][sk])) for sk in present.keys()]
    label_to_key = { display: sk for (sk, display) in sheet_options }
    display_names = [disp for _, disp in sheet_options]

    # Determine index for radio (default to previously selected sheet)
    current_key = st.session_state.get("current_sheet_key")
    if current_key in present:
        current_display = TARGET_SHEETS_CANON.get(current_key, present[current_key])
        try:
            default_index = display_names.index(current_display)
        except ValueError:
            default_index = 0
    else:
        default_index = 0

    selected_display = st.radio("Choose target sheet to map", options=display_names, index=default_index, horizontal=True)
    sheet_key = label_to_key[selected_display]
    st.session_state["current_sheet_key"] = sheet_key

    templ_headers = templ_headers_by_sheet.get(sheet_key, [])

    # 3-column layout: left (template headers), center (dropdowns), right (tools)
    left_col, mid_col, right_col = st.columns((1, 2, 1), gap="large")

    with left_col:
        st.markdown("**Template Headers (from row 6)**")
        if templ_headers:
            st.dataframe(pd.DataFrame({"Template Header": templ_headers}), use_container_width=True, height=400)
        else:
            st.write("No headers detected on row 6 of this sheet.")

    with mid_col:
        st.markdown("**Mapping (select one raw column per template header)**")
        if not raw_headers:
            st.warning("Upload raw data first.")
        else:
            # One selectbox per template header
            options = [""] + raw_headers
            for t in templ_headers:
                t_norm = _norm_key(t)
                key = f"map_{sheet_key}_{t_norm}"
                current_val = st.session_state.get(key, "")
                try:
                    idx = options.index(current_val) if current_val in options else 0
                except ValueError:
                    idx = 0
                st.selectbox(
                    label=t,
                    options=options,
                    index=idx,
                    key=key,
                    on_change=_make_on_change(sheet_key, t_norm),
                    help=None,
                )

    with right_col:
        st.markdown("**Tools**")

        # Auto-map exact (fill blanks only)
        if st.button("Auto‑map (exact)", use_container_width=True):
            _auto_map_exact(sheet_key, templ_headers, raw_headers)
            st.toast("Exact auto‑map applied (blanks only).")

        # Auto-map fuzzy (fill blanks only)
        fuzz_thr = st.slider("Fuzzy threshold", min_value=0, max_value=100, value=80, step=1)
        if st.button("Auto‑map (fuzzy)", use_container_width=True):
            _auto_map_fuzzy(sheet_key, templ_headers, raw_headers, threshold=fuzz_thr)
            st.toast(f"Fuzzy auto‑map applied at threshold {fuzz_thr} (blanks only).")

        # Import mapping
        st.markdown("---")
        st.markdown("**Import mapping (CSV/XLSX with columns: Template, Raw, [Sheet])**")
        map_in = st.file_uploader("Import mapping file", type=["csv", "xlsx"], key="map_import_upl")
        if map_in is not None:
            try:
                if map_in.name.lower().endswith(".csv"):
                    mdf = pd.read_csv(map_in)
                else:
                    mdf = pd.read_excel(map_in)
                applied = _import_mapping(
                    mdf,
                    apply_sheet_key=sheet_key,
                    present_sheets=present,
                    templ_headers_by_sheet=templ_headers_by_sheet,
                    raw_headers=raw_headers,
                )
                st.success(f"Imported {applied} mapping rows.")
            except Exception as e:
                st.error(f"Failed to import mapping: {e}")

        # Export mapping
        st.markdown("---")
        st.markdown("**Export mapping (CSV)**")
        export_df = _export_mapping_df(present, templ_headers_by_sheet)
        st.download_button(
            "Export current mapping",
            data=export_df.to_csv(index=False).encode("utf-8"),
            file_name="mapping_export.csv",
            mime="text/csv",
            use_container_width=True,
        )

        st.markdown("---")
        auto_resolve = st.checkbox(
            "Auto‑resolve duplicate *Raw* selections on download (latest edit wins)",
            value=True,
            key="auto_resolve_dupe_mappings"
        )
        # DEFAULT DUPLICATE HIGHLIGHT COLUMNS UPDATED PER REQUEST:
        dup_cols_text = st.text_input(
            "Duplicate columns to highlight (comma‑separated)",
            value="SKU, productId, manufacturerPartNumber"
        )
        dup_cols = [c.strip() for c in dup_cols_text.split(",") if c.strip()]

        # Filename enforcement (xlsm keeps .xlsm; else .xlsx)
        tpl_ext = st.session_state.get("template_ext", "xlsx")
        is_xlsm = (tpl_ext == "xlsm")
        suggested_name = "filled_template.xlsm" if is_xlsm else "filled_template.xlsx"
        file_name_input = st.text_input("Output filename", value=suggested_name)
        final_name = _enforce_extension(file_name_input, is_xlsm=is_xlsm)

        st.markdown("---")
        # Prepare and download
        if st.button("Build file", use_container_width=True):
            try:
                payload = _build_output_bytes(
                    template_bytes=st.session_state["template_bytes"],
                    template_is_xlsm=is_xlsm,
                    present_sheets=present,
                    templ_headers_by_sheet=templ_headers_by_sheet,
                    raw_df=raw_df,
                    auto_resolve_dupe_mappings=auto_resolve,
                    dup_columns_to_highlight=dup_cols,
                )
                st.session_state["download_payload"] = payload
                st.success("File prepared. Use the download button below.")
            except ValueError as ve:
                st.error(str(ve))
            except Exception as e:
                st.error(f"Failed to build file: {e}")

        if st.session_state.get("download_payload", None):
            st.download_button(
                "Download Excel",
                data=st.session_state["download_payload"],
                file_name=final_name,
                mime="application/vnd.ms-excel" if is_xlsm else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

# -----------------------------
# Notes for operators:
# - Mapping is per target sheet (radio switch at the top of tab 3).
# - Row 6 = headers; Row 7+ = data writes.
# - Only the PCSE/TIC sheets are edited; other sheets (like Data Definitions) remain untouched.
# - Macros in .xlsm are preserved (keep_vba=True).
# - No live clearing or uniqueness enforcement while editing; duplicate RAW selections are handled at download time.
# - Scroll position is preserved across interactions.
# -----------------------------
