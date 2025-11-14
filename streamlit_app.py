
import re
from io import BytesIO
from difflib import SequenceMatcher
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# =============================
# Page config
# =============================
st.set_page_config(page_title="Walmart Masterfile Filler", layout="wide")

# Polished heading + subheading (no custom colors)
st.markdown(
    """
    <div style="text-align:center; margin-top:-0.5rem; margin-bottom:0.5rem;">
      <h1 style="margin-bottom:0.25rem;">Walmart Masterfile Filler</h1>
      <p style="margin-top:0; font-size:1.05rem;">Empowering Innovation ⏐ Maximizing Performance</p>
    </div>
    """,
    unsafe_allow_html=True,
)
st.divider()

# =============================
# Canonical target sheets (normalized key -> display name)
# =============================
TARGET_SHEETS_CANON = {
    "productcontentandsiteexp": "Product Content And Site Exp",
    "tradeitemconfigurations": "Trade Item Configurations",
}
CANON_ORDER = ["productcontentandsiteexp", "tradeitemconfigurations"]

# Synonyms for mapping workbook tab detection (import)
MAPPING_TAB_SYNONYMS = {
    "productcontentandsiteexp": ["product content and site exp", "product site exp", "pcse"],
    "tradeitemconfigurations": ["trade item configurations", "trade item", "tic"],
}

# Walmart template specifics
HEADER_ROW_INDEX = 5  # titles in row 5
DATA_START_ROW = 7    # write from row 7

# Highlight style for duplicate cells
YELLOW_DUP_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# =============================
# Helpers
# =============================
def _norm_key(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower().replace("&", "and")
    return re.sub(r"[^a-z0-9]+", "", s)

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
    """Return mapping canon_key -> actual sheet name for present target sheets, tolerant to variations."""
    present: Dict[str, str] = {}
    norm_actual = { _norm_key(n): n for n in actual_names }
    for canon_norm, display in TARGET_SHEETS_CANON.items():
        if canon_norm in norm_actual:
            present[canon_norm] = norm_actual[canon_norm]
            continue
        # Fallback: partial contains
        for n_norm, real in norm_actual.items():
            if canon_norm in n_norm or n_norm in canon_norm:
                present[canon_norm] = real
                break
    return present

def _extract_headers_and_positions(ws, header_row: int) -> Tuple[List[str], Dict[int, int], Dict[str, List[int]]]:
    """
    Read the header row into:
        - headers: ordered list of non-empty header strings (left->right)
        - pos_to_col: mapping 1-based header position -> column index
        - norm_to_cols: mapping normalized header -> list of column indexes (handles duplicates)
    """
    headers: List[str] = []
    pos_to_col: Dict[int, int] = {}
    norm_to_cols: Dict[str, List[int]] = {}
    for col_idx, cell in enumerate(ws[header_row], start=1):
        val = cell.value
        if val is None or str(val).strip() == "":
            continue
        text = str(val).strip()
        headers.append(text)
        pos = len(headers)  # 1-based
        pos_to_col[pos] = col_idx
        norm = _norm_key(text)
        norm_to_cols.setdefault(norm, []).append(col_idx)
    return headers, pos_to_col, norm_to_cols

# ----- widget keys per position (prevents duplicate-key errors) -----
def _map_key(sheet_key: str, templ_norm: str, pos: int) -> str:
    return f"map_{sheet_key}_{pos}_{templ_norm}"

def _tick_key(sheet_key: str, templ_norm: str, pos: int) -> str:
    return f"tick_{sheet_key}_{pos}_{templ_norm}"

def _iter_template_positions(sheet_key: str, templ_headers: List[str]):
    """Yield (header_text, normalized, pos, map_key, tick_key) for each header position."""
    for idx, t in enumerate(templ_headers, start=1):
        templ_norm = _norm_key(t)
        yield t, templ_norm, idx, _map_key(sheet_key, templ_norm, idx), _tick_key(sheet_key, templ_norm, idx)

def _touch_tick(sheet_key: str, templ_norm: str, pos: int):
    st.session_state["last_edit_tick"] = st.session_state.get("last_edit_tick", 0) + 1
    st.session_state[_tick_key(sheet_key, templ_norm, pos)] = st.session_state["last_edit_tick"]

def _make_on_change(sheet_key: str, templ_norm: str, pos: int):
    def _cb():
        _touch_tick(sheet_key, templ_norm, pos)
    return _cb

def _build_live_mapping_for_sheet(sheet_key: str, templ_headers: List[str]) -> List[Dict[str, str]]:
    """Capture current UI selections for a sheet as mapping records (skip blanks)."""
    records = []
    for t, templ_norm, pos, mkey, tkey in _iter_template_positions(sheet_key, templ_headers):
        raw = st.session_state.get(mkey, "") or ""
        if raw:
            records.append({
                "template_header": t,
                "template_norm": templ_norm,
                "templ_pos": pos,      # critical for disambiguation
                "raw_header": raw,
                "tick": st.session_state.get(tkey, 0),
            })
    return records

def _commit_current_sheet_mapping(sheet_key: str, templ_headers: List[str]):
    """Snapshot the current (live) mapping for the selected sheet into a saved store."""
    recs = _build_live_mapping_for_sheet(sheet_key, templ_headers)
    saved = st.session_state.setdefault("saved_mapping", {})
    saved[sheet_key] = recs
    st.session_state["saved_mapping"] = saved

def _hydrate_live_from_saved(sheet_key: str, templ_headers: List[str]):
    """Populate blank UI selects from saved snapshot (does not overwrite existing user edits)."""
    saved = st.session_state.get("saved_mapping", {})
    recs = saved.get(sheet_key, [])
    by_pos = { r.get("templ_pos"): r.get("raw_header") for r in recs if r.get("templ_pos") is not None }
    for t, templ_norm, pos, mkey, _ in _iter_template_positions(sheet_key, templ_headers):
        if not st.session_state.get(mkey, "") and pos in by_pos:
            st.session_state[mkey] = by_pos[pos]

# ----- automap -----
def _auto_map_exact(sheet_key: str, templ_headers: List[str], raw_headers: List[str]):
    norm_idx = { _norm_key(r): r for r in raw_headers }
    for t, templ_norm, pos, mkey, _ in _iter_template_positions(sheet_key, templ_headers):
        if not st.session_state.get(mkey, ""):
            match = norm_idx.get(templ_norm)
            if match:
                st.session_state[mkey] = match
                _touch_tick(sheet_key, templ_norm, pos)

def _auto_map_fuzzy(sheet_key: str, templ_headers: List[str], raw_headers: List[str], threshold: int = 80):
    thr = max(0, min(100, threshold)) / 100.0
    raw_norms = [(r, _norm_key(r)) for r in raw_headers]
    for t, templ_norm, pos, mkey, _ in _iter_template_positions(sheet_key, templ_headers):
        if st.session_state.get(mkey, ""):
            continue
        best_raw, best_score = None, 0.0
        for r, rnorm in raw_norms:
            score = SequenceMatcher(None, templ_norm, rnorm).ratio()
            if score > best_score:
                best_score, best_raw = score, r
        if best_raw and best_score >= thr:
            st.session_state[mkey] = best_raw
            _touch_tick(sheet_key, templ_norm, pos)

# ----- import/export mapping -----
def _import_mapping_df_for_sheet(df: pd.DataFrame,
                                 target_sheet_key: str,
                                 templ_headers_by_sheet: Dict[str, List[str]],
                                 raw_headers: List[str]) -> int:
    """Import mapping for a single sheet. Columns accepted: Template, Raw (case-insensitive)."""
    cols = { _norm_key(c): c for c in df.columns }
    tcol = next((cols[c] for c in ["template","templateheader","templ","target","targetheader"] if c in cols), None)
    rcol = next((cols[c] for c in ["raw","rawheader","source","sourceheader"] if c in cols), None)
    if not tcol or not rcol:
        return 0

    templ_headers = templ_headers_by_sheet.get(target_sheet_key, [])
    raw_set = set(raw_headers)
    applied = 0
    # For duplicates: assign to the first blank occurrence, else first occurrence.
    for _, row in df.iterrows():
        templ = str(row[tcol]).strip() if pd.notna(row[tcol]) else ""
        raw = str(row[rcol]).strip() if pd.notna(row[rcol]) else ""
        if not templ or not raw or raw not in raw_set:
            continue
        # first try blank occurrences
        assigned = False
        for t, templ_norm, pos, mkey, _ in _iter_template_positions(target_sheet_key, templ_headers):
            if t == templ and not st.session_state.get(mkey, ""):
                st.session_state[mkey] = raw
                _touch_tick(target_sheet_key, templ_norm, pos)
                applied += 1
                assigned = True
                break
        if assigned:
            continue
        # fallback: first occurrence of this header
        for t, templ_norm, pos, mkey, _ in _iter_template_positions(target_sheet_key, templ_headers):
            if t == templ:
                st.session_state[mkey] = raw
                _touch_tick(target_sheet_key, templ_norm, pos)
                applied += 1
                break
    return applied

def _export_mapping_df(present_sheets: Dict[str, str],
                       templ_headers_by_sheet: Dict[str, List[str]],
                       only_sheet_key: Optional[str] = None) -> pd.DataFrame:
    rows = []
    for sk in _ordered_keys(present_sheets):
        if only_sheet_key and sk != only_sheet_key:
            continue
        actual_name = present_sheets[sk]
        templ_headers = templ_headers_by_sheet.get(sk, [])
        for t, templ_norm, pos, mkey, _ in _iter_template_positions(sk, templ_headers):
            v = st.session_state.get(mkey, "")
            rows.append({"Sheet": TARGET_SHEETS_CANON.get(sk, actual_name), "Position": pos, "Template": t, "Raw": v})
    return pd.DataFrame(rows)

# ----- duplicate RAW selection handling -----
def _resolve_duplicate_raw_mappings(records: List[Dict[str, str]], auto_resolve: bool) -> Tuple[List[Dict[str, str]], List[str]]:
    bucket: Dict[str, List[Dict[str, str]]] = {}
    for r in records:
        bucket.setdefault(r["raw_header"], []).append(r)
    dups = [raw for raw, lst in bucket.items() if len(lst) > 1]
    if not dups:
        return records, []
    if auto_resolve:
        resolved = []
        for raw, lst in bucket.items():
            if len(lst) == 1:
                resolved.append(lst[0])
            else:
                keep = sorted(lst, key=lambda x: x.get("tick", 0))[-1]
                resolved.append(keep)
        return resolved, []
    else:
        return records, dups

# ----- writing -----
def _write_sheet_data(ws, mapping: List[Dict[str, str]],
                      header_row: int,
                      start_row: int,
                      raw_df: pd.DataFrame,
                      dup_headers_to_highlight: List[str]) -> Tuple[int, List[str]]:
    """Write mapped data and highlight duplicate values in specified columns (by header)."""
    headers, pos_to_col, norm_to_cols = _extract_headers_and_positions(ws, header_row)

    # Build write plan (prefer explicit templ_pos; fallback to first column by header norm)
    write_plan: List[Tuple[int, str]] = []  # (col_idx, raw_header)
    missing: List[str] = []
    for m in mapping:
        pos = m.get("templ_pos")
        if isinstance(pos, int) and pos in pos_to_col:
            col_idx = pos_to_col[pos]
            write_plan.append((col_idx, m["raw_header"]))
        else:
            norm = _norm_key(m["template_header"])
            if norm in norm_to_cols and len(norm_to_cols[norm]) > 0:
                write_plan.append((norm_to_cols[norm][0], m["raw_header"]))  # fallback: first matching column
            else:
                missing.append(m["template_header"])

    nrows = len(raw_df)
    for i in range(nrows):
        excel_row = start_row + i
        for col_idx, raw_name in write_plan:
            val = raw_df.iloc[i][raw_name] if raw_name in raw_df.columns else None
            ws.cell(row=excel_row, column=col_idx, value=val)

    # Highlight duplicates for all columns whose header matches the requested names
    dup_norms = [_norm_key(x) for x in dup_headers_to_highlight if str(x).strip()]
    for want_norm in dup_norms:
        for cidx in norm_to_cols.get(want_norm, []):
            # count occurrences
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
            for i in range(nrows):
                cell = ws.cell(row=start_row + i, column=cidx)
                v = cell.value
                if v is None:
                    continue
                if str(v) in dup_values:
                    cell.fill = YELLOW_DUP_FILL

    return nrows, missing

def _build_output_bytes(template_bytes: bytes,
                        template_is_xlsm: bool,
                        sheets_to_process: Dict[str, str],
                        templ_headers_by_sheet: Dict[str, List[str]],
                        raw_df: pd.DataFrame,
                        auto_resolve_dupe_mappings: bool,
                        dup_columns_to_highlight: List[str],
                        saved_mapping_store: Optional[Dict[str, List[Dict[str, str]]]] = None) -> bytes:
    bio_in = BytesIO(template_bytes)
    wb = load_workbook(bio_in, read_only=False, keep_vba=template_is_xlsm, data_only=False)

    for sk, actual_name in sheets_to_process.items():
        ws = wb[actual_name]
        if saved_mapping_store and sk in saved_mapping_store:
            mapping_records = list(saved_mapping_store[sk])
        else:
            mapping_records = _build_live_mapping_for_sheet(sk, templ_headers_by_sheet.get(sk, []))

        mapping_resolved, dup_raw = _resolve_duplicate_raw_mappings(mapping_records, auto_resolve_dupe_mappings)
        if dup_raw:
            raise ValueError(
                f"Duplicate RAW column selections for sheet '{TARGET_SHEETS_CANON.get(sk, actual_name)}': {', '.join(sorted(set(dup_raw)))}. "
                "Turn ON auto‑resolve or change your selections."
            )

        _write_sheet_data(
            ws=ws,
            mapping=mapping_resolved,
            header_row=HEADER_ROW_INDEX,
            start_row=DATA_START_ROW,
            raw_df=raw_df,
            dup_headers_to_highlight=dup_columns_to_highlight,
        )

    bio_out = BytesIO()
    wb.save(bio_out)  # preserves formatting, heights, colors, validations; keep_vba preserves macros
    return bio_out.getvalue()

# ----- ordering helper -----
def _ordered_keys(present_sheets: Dict[str, str]) -> List[str]:
    ordered = [k for k in CANON_ORDER if k in present_sheets]
    ordered += [k for k in present_sheets.keys() if k not in ordered]
    return ordered

def _match_mapping_tab(sheet_names: List[str], target_key: str) -> Optional[str]:
    """Find a mapping tab name inside the uploaded workbook that corresponds to target_key."""
    candidates = [(_norm_key(n), n) for n in sheet_names]
    targets = [target_key] + MAPPING_TAB_SYNONYMS.get(target_key, [])
    targets_norm = [_norm_key(x) for x in targets]
    for norm, name in candidates:
        if norm in targets_norm:
            return name
    for norm, name in candidates:
        for t in targets_norm:
            if t in norm or norm in t:
                return name
    return None

# =============================
# Session bootstrap
# =============================
for k, v in [
    ("template_bytes", None),
    ("template_ext", None),
    ("present_sheets", {}),
    ("templ_headers_by_sheet", {}),
    ("templ_header_sigs", {}),
    ("raw_headers", []),
    ("raw_sig", ""),
    ("raw_prev_headers", []),
    ("current_sheet_key", None),
    ("prev_sheet_key", None),
    ("last_edit_tick", 0),
    ("download_payload", None),
    ("saved_mapping", {}),
]:
    st.session_state.setdefault(k, v)

# =============================
# Scroll position retention (avoid page jump)
# =============================
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

# =============================
# Tabs
# =============================
tab1, tab2, tab3 = st.tabs(["Upload Masterfile Template", "Upload Raw / Onboarding Data", "Simple Mapping & Download"])

# -----------------------------
# Tab 1: Upload Masterfile Template
# -----------------------------
with tab1:
    st.markdown("#### Step 1 of 3 — Upload Masterfile Template (.xlsx / .xlsm)")
    tpl = st.file_uploader("Upload template", type=["xlsx", "xlsm"], key="template_uploader")

    if tpl is not None:
        raw_bytes = tpl.read()
        is_xlsm = tpl.name.lower().endswith(".xlsm")

        wb = load_workbook(BytesIO(raw_bytes), read_only=False, keep_vba=is_xlsm, data_only=False)
        all_names = wb.sheetnames
        present = _find_target_sheets(all_names)

        templ_headers_by_sheet: Dict[str, List[str]] = {}
        templ_header_sigs: Dict[str, str] = {}
        for sk, actual in present.items():
            ws = wb[actual]
            headers, _, _ = _extract_headers_and_positions(ws, HEADER_ROW_INDEX)
            templ_headers_by_sheet[sk] = headers
            templ_header_sigs[sk] = "|".join(headers)

        prev_sigs = st.session_state.get("templ_header_sigs", {})
        st.session_state["template_bytes"] = raw_bytes
        st.session_state["template_ext"] = "xlsm" if is_xlsm else "xlsx"
        st.session_state["present_sheets"] = present
        st.session_state["templ_headers_by_sheet"] = templ_headers_by_sheet
        st.session_state["templ_header_sigs"] = templ_header_sigs

        # Clear only on header signature change per sheet (both live & saved)
        for sk, sig in templ_header_sigs.items():
            if prev_sigs.get(sk) != sig:
                prefix_map = f"map_{sk}_"
                prefix_tick = f"tick_{sk}_"
                for k in list(st.session_state.keys()):
                    if k.startswith(prefix_map) or k.startswith(prefix_tick):
                        del st.session_state[k]
                if sk in st.session_state["saved_mapping"]:
                    del st.session_state["saved_mapping"][sk]

        # default sheet selection
        ordered = _ordered_keys(present)
        if ordered:
            if st.session_state["current_sheet_key"] not in present:
                st.session_state["current_sheet_key"] = ordered[0]

        st.success("Template loaded.")
        if present:
            st.caption(f"Headers read from row {HEADER_ROW_INDEX}. Data will be written from row {DATA_START_ROW}.")
            st.write("Detected target sheets:")
            for sk in _ordered_keys(present):
                actual = present[sk]
                display = TARGET_SHEETS_CANON.get(sk, actual)
                st.write(f"- **{display}** → actual name: `{actual}`")
        else:
            st.warning("No target sheets found. Expected one or both of: "
                       "'Product Content And Site Exp', 'Trade Item Configurations'.")

# -----------------------------
# Tab 2: Upload Raw / Onboarding Data
# -----------------------------
with tab2:
    st.markdown("#### Step 2 of 3 — Upload Raw / Onboarding Data (.csv / .xlsx)")
    raw_file = st.file_uploader("Upload raw/onboarding file", type=["csv", "xlsx"], key="raw_uploader")

    raw_df = None
    if raw_file is not None:
        if raw_file.name.lower().endswith(".csv"):
            raw_df = pd.read_csv(raw_file)
        else:
            xl = pd.ExcelFile(raw_file)
            sheetname = st.selectbox("Choose a sheet", xl.sheet_names, index=0)
            raw_df = xl.parse(sheetname)

        raw_df.columns = [str(c) for c in raw_df.columns]
        new_headers = list(raw_df.columns)
        new_sig = "|".join(new_headers)

        # Clear only selections that reference removed raw headers (live + saved)
        removed = set(st.session_state.get("raw_prev_headers", [])) - set(new_headers)
        if removed:
            for sk in st.session_state.get("present_sheets", {}).keys():
                # Clear all live mapping entries that reference removed columns
                for k in list(st.session_state.keys()):
                    if k.startswith(f"map_{sk}_"):
                        if st.session_state.get(k, "") in removed:
                            st.session_state[k] = ""
                # Prune saved recs
                if "saved_mapping" in st.session_state and sk in st.session_state["saved_mapping"]:
                    st.session_state["saved_mapping"][sk] = [
                        r for r in st.session_state["saved_mapping"][sk] if r.get("raw_header") not in removed
                    ]

        st.session_state["raw_prev_headers"] = new_headers
        st.session_state["raw_headers"] = new_headers
        st.session_state["raw_sig"] = new_sig
        st.session_state["raw_df_payload"] = raw_df

        st.success(f"Raw data loaded ({len(raw_df)} rows, {len(new_headers)} columns).")
        with st.expander("Preview (first 20 rows)"):
            st.dataframe(raw_df.head(20), use_container_width=True)

# -----------------------------
# Tab 3: Mapping & Download
# -----------------------------
with tab3:
    st.markdown("#### Step 3 of 3 — Simple Mapping & Download")

    if not st.session_state.get("template_bytes"):
        st.info("Please upload a masterfile template in tab 1.")
        st.stop()
    if "raw_df_payload" not in st.session_state:
        st.info("Please upload raw/onboarding data in tab 2.")
        st.stop()

    present = st.session_state["present_sheets"]
    templ_headers_by_sheet = st.session_state["templ_headers_by_sheet"]
    raw_headers = st.session_state["raw_headers"]
    raw_df: pd.DataFrame = st.session_state["raw_df_payload"]

    if not present:
        st.error("Template has no target sheets to map (expected PCSE and/or TIC).")
        st.stop()

    ordered_keys = _ordered_keys(present)
    display_names = [TARGET_SHEETS_CANON.get(sk, present[sk]) for sk in ordered_keys]
    key_by_display = { TARGET_SHEETS_CANON.get(sk, present[sk]): sk for sk in ordered_keys }

    # Sheet switcher (auto-save previous sheet on change)
    current_key_before = st.session_state.get("current_sheet_key")
    default_index = 0
    if current_key_before in present:
        current_display_before = TARGET_SHEETS_CANON.get(current_key_before, present.get(current_key_before, ""))
        if current_display_before in display_names:
            default_index = display_names.index(current_display_before)
    selected_display = st.radio("Choose target sheet to map", options=display_names, index=default_index, horizontal=True)
    sheet_key = key_by_display[selected_display]

    if current_key_before and current_key_before != sheet_key and current_key_before in present:
        _commit_current_sheet_mapping(current_key_before, templ_headers_by_sheet.get(current_key_before, []))
    st.session_state["current_sheet_key"] = sheet_key

    templ_headers = templ_headers_by_sheet.get(sheet_key, [])
    _hydrate_live_from_saved(sheet_key, templ_headers)

    left, mid, right = st.columns((1, 2, 1), gap="large")

    with left:
        st.markdown(f"**Template Headers (from row {HEADER_ROW_INDEX})**")
        if templ_headers:
            st.dataframe(pd.DataFrame({"Template Header": templ_headers}), use_container_width=True, height=400)
        else:
            st.write(f"No headers detected on row {HEADER_ROW_INDEX} of this sheet.")

        # Saved mapping status across sheets
        st.markdown("**Saved mapping status**")
        for sk in ordered_keys:
            total = len(templ_headers_by_sheet.get(sk, []))
            saved_count = len(st.session_state.get("saved_mapping", {}).get(sk, []))
            st.caption(f"{TARGET_SHEETS_CANON.get(sk, present[sk])}: {saved_count} saved of {total} headers")

    with mid:
        st.markdown("**Mapping (select one raw column per template header)**")
        if not raw_headers:
            st.warning("Upload raw data first.")
        else:
            options = [""] + raw_headers
            for t, templ_norm, pos, mkey, tkey in _iter_template_positions(sheet_key, templ_headers):
                current_val = st.session_state.get(mkey, "")
                try:
                    idx = options.index(current_val) if current_val in options else 0
                except ValueError:
                    idx = 0
                st.selectbox(
                    label=t,
                    options=options,
                    index=idx,
                    key=mkey,
                    on_change=_make_on_change(sheet_key, templ_norm, pos),
                )

    with right:
        st.markdown("**Tools**")
        if st.button("Auto‑map (exact)", use_container_width=True):
            _auto_map_exact(sheet_key, templ_headers, raw_headers)
            st.toast("Exact auto‑map applied (blanks only).")

        fuzz_thr = st.slider("Fuzzy threshold", 0, 100, 80, 1)
        if st.button("Auto‑map (fuzzy)", use_container_width=True):
            _auto_map_fuzzy(sheet_key, templ_headers, raw_headers, threshold=fuzz_thr)
            st.toast(f"Fuzzy auto‑map applied at threshold {fuzz_thr} (blanks only).")

        st.markdown("---")
        # -------- Single Import area: ONE uploader for a workbook with two mapping tabs --------
        st.markdown("**Import mapping workbook (.xlsx with two tabs: PCSE & TIC)**")
        st.caption("Provide one Excel file containing two sheets — one for **Product Content And Site Exp** and one for **Trade Item Configurations**. Each tab must have columns: **Template**, **Raw**.")
        imp_wb = st.file_uploader("Upload mapping workbook (.xlsx)", type=["xlsx"], key="map_workbook")
        if imp_wb is not None:
            try:
                xl = pd.ExcelFile(imp_wb)
                applied_summary = []
                for sk in ordered_keys:
                    tab_name = _match_mapping_tab(xl.sheet_names, sk)
                    if not tab_name:
                        applied_summary.append(f"{TARGET_SHEETS_CANON.get(sk, present[sk])}: no matching tab found")
                        continue
                    df = xl.parse(tab_name)
                    cnt = _import_mapping_df_for_sheet(df, sk, templ_headers_by_sheet, raw_headers)
                    applied_summary.append(f"{TARGET_SHEETS_CANON.get(sk, present[sk])}: {cnt} rows applied from '{tab_name}'")
                st.success("Import complete:\n- " + "\n- ".join(applied_summary))
            except Exception as e:
                st.error(f"Failed to import mapping workbook: {e}")

        st.markdown("---")
        export_df_all = _export_mapping_df(present, templ_headers_by_sheet)
        st.download_button(
            "Export mapping (all sheets)",
            data=export_df_all.to_csv(index=False).encode("utf-8"),
            file_name="mapping_export_all.csv",
            mime="text/csv",
            use_container_width=True,
        )
        export_df_selected = _export_mapping_df(present, templ_headers_by_sheet, only_sheet_key=sheet_key)
        st.download_button(
            "Export mapping (selected sheet)",
            data=export_df_selected.to_csv(index=False).encode("utf-8"),
            file_name=f"mapping_export_{TARGET_SHEETS_CANON.get(sheet_key, sheet_key)}.csv",
            mime="text/csv",
            use_container_width=True,
        )

        st.markdown("---")
        auto_resolve = st.checkbox(
            "Auto‑resolve duplicate *Raw* selections on download (latest edit wins)",
            value=True,
            key="auto_resolve_dupe_mappings"
        )
        dup_cols_text = st.text_input(
            "Duplicate columns to highlight (comma‑separated)",
            value="SKU, productId, manufacturerPartNumber"
        )
        dup_cols = [c.strip() for c in dup_cols_text.split(",") if c.strip()]

        tpl_ext = st.session_state.get("template_ext", "xlsx")
        is_xlsm = (tpl_ext == "xlsm")
        suggested_name = "filled_template.xlsm" if is_xlsm else "filled_template.xlsx"
        file_name_input = st.text_input("Output filename", value=suggested_name)
        final_name = _enforce_extension(file_name_input, is_xlsm=is_xlsm)

        st.markdown("---")
        if st.button("Save mapping (selected sheet)", use_container_width=True):
            _commit_current_sheet_mapping(sheet_key, templ_headers)
            st.success(f"Saved mapping for '{TARGET_SHEETS_CANON.get(sheet_key, present.get(sheet_key, sheet_key))}'.")

        if st.button("Save & Build (process both sheets)", use_container_width=True):
            try:
                _commit_current_sheet_mapping(sheet_key, templ_headers)
                payload = _build_output_bytes(
                    template_bytes=st.session_state["template_bytes"],
                    template_is_xlsm=is_xlsm,
                    sheets_to_process=present,  # process all detected target sheets
                    templ_headers_by_sheet=templ_headers_by_sheet,
                    raw_df=raw_df,
                    auto_resolve_dupe_mappings=auto_resolve,
                    dup_columns_to_highlight=dup_cols,
                    saved_mapping_store=st.session_state.get("saved_mapping", {}),
                )
                st.session_state["download_payload"] = payload
                st.success("File prepared. Use the download button below.")
            except ValueError as ve:
                st.error(str(ve))
            except Exception as e:
                st.error(f"Failed to build file: {e}")

        if st.button("Build (selected sheet only)", use_container_width=True):
            try:
                payload = _build_output_bytes(
                    template_bytes=st.session_state["template_bytes"],
                    template_is_xlsm=is_xlsm,
                    sheets_to_process={sheet_key: present[sheet_key]},
                    templ_headers_by_sheet=templ_headers_by_sheet,
                    raw_df=raw_df,
                    auto_resolve_dupe_mappings=auto_resolve,
                    dup_columns_to_highlight=dup_cols,
                    saved_mapping_store=None,  # use live mapping for selected sheet
                )
                st.session_state["download_payload"] = payload
                st.success("File prepared. Use the download button below.")
            except ValueError as ve:
                st.error(str(ve))
            except Exception as e:
                st.error(f"Failed to build file: {e}")

        if st.session_state.get("download_payload"):
            st.download_button(
                "Download Excel",
                data=st.session_state["download_payload"],
                file_name=final_name,
                mime="application/vnd.ms-excel" if is_xlsm else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
