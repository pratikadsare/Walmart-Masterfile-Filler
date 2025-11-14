# streamlit_app.py
# Walmart Masterfile Filler — Horizontal Mapping with Styled Excel Preview (rows 1–6)
# - Shows a near-pixel preview of the first 6 rows *with* merged cells, fills, fonts, alignments, widths/heights (approx)
# - Horizontal mapping row (one select per template column) sits directly beneath the preview
# - Supports multi-row/composite headers (e.g., 3,4,5) for mapping
# - Preserves template macros/formatting/validations in the exported Excel
# - Writes starting from row 7 by default (configurable)
# - Duplicate highlighting defaults to: SKU, productId, manufacturerPartNumber
#
# Requirements:
#   streamlit>=1.32
#   pandas>=2.0
#   openpyxl>=3.1

import re
import html
from io import BytesIO
from difflib import SequenceMatcher
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter


# =============================
# Page config & hero
# =============================
st.set_page_config(page_title="Walmart Masterfile Filler", layout="wide")
st.markdown(
    """
    <div style="text-align:center; margin-top:-0.5rem; margin-bottom:0.5rem;">
      <h1 style="margin-bottom:0.25rem;">Walmart Masterfile Filler 2</h1>
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

# Defaults
DEFAULT_HEADER_ROWS = [5]   # can be customized (e.g., 3,4,5)
DEFAULT_DATA_START_ROW = 7  # write from row 7 by default

# Highlight style for duplicate cells in the exported Excel
YELLOW_DUP_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


# =============================
# Helpers
# =============================
def _norm_key(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower().replace("&", "and")
    return re.sub(r"[^a-z0-9]+", "", s)


def _parse_rows_input(s: str, default: List[int]) -> List[int]:
    try:
        vals = [int(x.strip()) for x in s.split(",") if x.strip()]
        vals = sorted(set(v for v in vals if v > 0))
        return vals if vals else default
    except Exception:
        return default


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
    present: Dict[str, str] = {}
    norm_actual = {_norm_key(n): n for n in actual_names}
    for canon_norm, _display in TARGET_SHEETS_CANON.items():
        if canon_norm in norm_actual:
            present[canon_norm] = norm_actual[canon_norm]
            continue
        for n_norm, real in norm_actual.items():
            if canon_norm in n_norm or n_norm in canon_norm:
                present[canon_norm] = real
                break
    return present


def _get_cell_value_with_merge(ws, row: int, col: int):
    cell = ws.cell(row=row, column=col)
    if cell.value not in (None, ""):
        return cell.value
    # If cell is part of a merged range, return top-left value
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return ws.cell(row=rng.min_row, column=rng.min_col).value
    return None


def _extract_headers_composite(ws, header_rows: List[int]) -> Tuple[List[str], Dict[int, int], Dict[str, List[int]]]:
    """Return (composite headers list, pos->col map, normalized header -> [col]) built from header_rows (multi-row headers)."""
    headers: List[str] = []
    pos_to_col: Dict[int, int] = {}
    norm_to_cols: Dict[str, List[int]] = {}
    max_col = ws.max_column or 0

    for col_idx in range(1, max_col + 1):
        parts: List[str] = []
        for r in header_rows:
            v = _get_cell_value_with_merge(ws, r, col_idx)
            if v is not None and str(v).strip() != "":
                parts.append(str(v).strip())
        if not parts:
            continue
        composite = " | ".join(parts)
        headers.append(composite)
        pos = len(headers)  # 1-based position among "valid" columns
        pos_to_col[pos] = col_idx
        norm = _norm_key(composite)
        norm_to_cols.setdefault(norm, []).append(col_idx)

    return headers, pos_to_col, norm_to_cols


# ---------- Styled Excel preview (rows 1–6) as HTML ----------
def _color_to_hex(color_obj) -> Optional[str]:
    """
    Convert openpyxl Color to CSS hex (#RRGGBB).
    Handles rgb (ARGB) primarily; theme/indexed are returned as None (best-effort fallback).
    """
    if color_obj is None:
        return None
    try:
        if getattr(color_obj, "type", None) == "rgb" and color_obj.rgb:
            rgb = color_obj.rgb
            if isinstance(rgb, str) and len(rgb) == 8:  # ARGB
                return f"#{rgb[-6:]}"
            if isinstance(rgb, str) and len(rgb) == 6:
                return f"#{rgb}"
        # theme/indexed fallback: None (we'll leave default styling)
        return None
    except Exception:
        return None


def _pt_to_px(points: float) -> int:
    try:
        return int(round(points * 96.0 / 72.0))  # 1pt = 1/72 in; 96 DPI
    except Exception:
        return 0


def _col_width_to_px(width_chars: Optional[float]) -> int:
    # Excel width ~ number of '0' characters; common rough conversion:
    # px ≈ trunc(7 * width + 5)
    if width_chars is None:
        return 0
    try:
        return int(max(32, round(7 * float(width_chars) + 5)))
    except Exception:
        return 64


def _build_preview_html(ws, max_rows: int = 6) -> str:
    """
    Build an HTML table for the first `max_rows` rows that respects:
      - merged cells (rowspan/colspan)
      - solid fills (background-color)
      - font bold/italic/size/color
      - horizontal/vertical alignment
      - approximate column widths and row heights
    """
    # Precompute merged ranges for quick checks
    merged_top_left = {}  # (r, c) -> merged range object
    merged_cells = set()  # (r, c) coordinates that are within a merge
    for rng in ws.merged_cells.ranges:
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                merged_cells.add((r, c))
        merged_top_left[(rng.min_row, rng.min_col)] = rng

    max_col = ws.max_column or 0

    # Column widths
    col_widths_px = []
    for c in range(1, max_col + 1):
        letter = get_column_letter(c)
        dim = ws.column_dimensions.get(letter)
        w = _col_width_to_px(getattr(dim, "width", None))
        col_widths_px.append(w if w > 0 else 80)

    # Row heights (first max_rows)
    row_heights_px = []
    for r in range(1, max_rows + 1):
        dim = ws.row_dimensions.get(r)
        if dim is not None and getattr(dim, "height", None):
            row_heights_px.append(_pt_to_px(dim.height))
        else:
            row_heights_px.append(22)  # a reasonable default

    # Build table header (colgroup for widths)
    html_parts = []
    html_parts.append(
        """
        <div style="overflow-x:auto;border:1px solid #e5e7eb;border-radius:6px;">
          <table style="border-collapse:collapse; font-family:system-ui, -apple-system, Segoe UI, Roboto, Arial; font-size:13px; width:max-content;">
            <colgroup>
        """
    )
    for w in col_widths_px:
        html_parts.append(f'<col style="width:{w}px;">')
    html_parts.append("</colgroup>")

    # Build rows 1..max_rows
    for r in range(1, max_rows + 1):
        tr_style = f"height:{row_heights_px[r-1]}px;"
        html_parts.append(f'<tr style="{tr_style}">')

        c = 1
        while c <= max_col:
            # Skip cells that are interior of a merge (not top-left)
            if (r, c) in merged_cells and (r, c) not in merged_top_left:
                c += 1
                continue

            cell = ws.cell(row=r, column=c)

            # Determine rowspan/colspan
            rowspan = 1
            colspan = 1
            if (r, c) in merged_top_left:
                rng = merged_top_left[(r, c)]
                # clamp to preview window vertically
                rng_max_row = min(rng.max_row, max_rows)
                rowspan = max(1, rng_max_row - r + 1)
                colspan = max(1, rng.max_col - c + 1)

            # Cell styles
            styles = []
            # Fill
            try:
                fill = cell.fill
                if isinstance(fill, PatternFill) and fill.fill_type == "solid":
                    bg = _color_to_hex(fill.fgColor) or _color_to_hex(fill.start_color)
                    if bg:
                        styles.append(f"background-color:{bg};")
            except Exception:
                pass

            # Font
            try:
                font = cell.font
                if font is not None:
                    if font.bold:
                        styles.append("font-weight:600;")
                    if font.italic:
                        styles.append("font-style:italic;")
                    if font.sz:
                        styles.append(f"font-size:{_pt_to_px(float(font.sz))}px;")
                    if font.color:
                        fg = _color_to_hex(font.color)
                        if fg:
                            styles.append(f"color:{fg};")
            except Exception:
                pass

            # Alignment
            try:
                align = cell.alignment
                if align is not None:
                    if align.horizontal:
                        styles.append(f"text-align:{align.horizontal};")
                    if align.vertical:
                        styles.append(f"vertical-align:{align.vertical};")
            except Exception:
                pass

            # Borders (simple hairline)
            styles.append("border:1px solid #d1d5db; padding:2px 6px;")

            # Text value
            val = _get_cell_value_with_merge(ws, r, c)
            text = "" if val is None else html.escape(str(val))

            # Emit TD
            span = ""
            if rowspan > 1:
                span += f' rowspan="{rowspan}"'
            if colspan > 1:
                span += f' colspan="{colspan}"'
            html_parts.append(f'<td{span} style="{"".join(styles)}">{text}</td>')

            c += colspan

        html_parts.append("</tr>")

    html_parts.append("</table></div>")
    return "".join(html_parts)


# ----- widget keys per position (prevents duplicate-key errors) -----
def _map_key(sheet_key: str, templ_norm: str, pos: int) -> str:
    return f"map_{sheet_key}_{pos}_{templ_norm}"


def _tick_key(sheet_key: str, templ_norm: str, pos: int) -> str:
    return f"tick_{sheet_key}_{pos}_{templ_norm}"


def _iter_template_positions(sheet_key: str, templ_headers: List[str]):
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
    """Capture current selections as mapping records (skip blanks)."""
    records = []
    for t, templ_norm, pos, mkey, tkey in _iter_template_positions(sheet_key, templ_headers):
        raw = st.session_state.get(mkey, "") or ""
        if raw:
            records.append({
                "template_header": t,
                "template_norm": templ_norm,
                "templ_pos": pos,
                "raw_header": raw,
                "tick": st.session_state.get(tkey, 0),
            })
    return records


def _commit_current_sheet_mapping(sheet_key: str, templ_headers: List[str]):
    recs = _build_live_mapping_for_sheet(sheet_key, templ_headers)
    saved = st.session_state.setdefault("saved_mapping", {})
    saved[sheet_key] = recs
    st.session_state["saved_mapping"] = saved


def _hydrate_live_from_saved(sheet_key: str, templ_headers: List[str]):
    saved = st.session_state.get("saved_mapping", {})
    recs = saved.get(sheet_key, [])
    by_pos = {r.get("templ_pos"): r.get("raw_header") for r in recs if r.get("templ_pos") is not None}
    for t, templ_norm, pos, mkey, _ in _iter_template_positions(sheet_key, templ_headers):
        if not st.session_state.get(mkey, "") and pos in by_pos:
            st.session_state[mkey] = by_pos[pos]


# ----- automap -----
def _auto_map_exact(sheet_key: str, templ_headers: List[str], raw_headers: List[str]):
    norm_idx = {_norm_key(r): r for r in raw_headers}
    for t, templ_norm, pos, mkey, _ in _iter_template_positions(sheet_key, templ_headers):
        if not st.session_state.get(mkey, ""):
            # try exact composite match; else try last token
            match = norm_idx.get(templ_norm)
            if not match:
                tokens = [tok.strip() for tok in t.split("|") if tok.strip()]
                if tokens:
                    match = norm_idx.get(_norm_key(tokens[-1]))
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
        # try composite
        for r, rnorm in raw_norms:
            score = SequenceMatcher(None, templ_norm, rnorm).ratio()
            if score > best_score:
                best_score, best_raw = score, r
        # try last token if composite insufficient
        if (best_raw is None) or (best_score < thr):
            tokens = [tok.strip() for tok in t.split("|") if tok.strip()]
            if tokens:
                last_norm = _norm_key(tokens[-1])
                for r, rnorm in raw_norms:
                    score = SequenceMatcher(None, last_norm, rnorm).ratio()
                    if score > best_score:
                        best_score, best_raw = score, r
        if best_raw and best_score >= thr:
            st.session_state[mkey] = best_raw
            _touch_tick(sheet_key, templ_norm, pos)


# ----- import/export mapping -----
def _header_match(import_templ: str, display_header: str) -> bool:
    a = _norm_key(import_templ)
    b = _norm_key(display_header)
    if a == b:
        return True
    tokens = re.split(r"\s*\|\s*|\s*>\s*", display_header)
    for tok in tokens:
        if _norm_key(tok) == a:
            return True
    return False


def _import_mapping_df_for_sheet(df: pd.DataFrame,
                                 target_sheet_key: str,
                                 templ_headers_by_sheet: Dict[str, List[str]],
                                 raw_headers: List[str]) -> int:
    cols = {_norm_key(c): c for c in df.columns}
    tcol = next((cols[c] for c in ["template", "templateheader", "templ", "target", "targetheader"] if c in cols), None)
    rcol = next((cols[c] for c in ["raw", "rawheader", "source", "sourceheader"] if c in cols), None)
    if not tcol or not rcol:
        return 0
    templ_headers = templ_headers_by_sheet.get(target_sheet_key, [])
    raw_set = set(raw_headers)
    applied = 0
    for _, row in df.iterrows():
        templ_in = str(row[tcol]).strip() if pd.notna(row[tcol]) else ""
        raw = str(row[rcol]).strip() if pd.notna(row[rcol]) else ""
        if not templ_in or not raw or raw not in raw_set:
            continue
        assigned = False
        for t, templ_norm, pos, mkey, _ in _iter_template_positions(target_sheet_key, templ_headers):
            if _header_match(templ_in, t) and not st.session_state.get(mkey, ""):
                st.session_state[mkey] = raw
                _touch_tick(target_sheet_key, templ_norm, pos)
                applied += 1
                assigned = True
                break
        if assigned:
            continue
        for t, templ_norm, pos, mkey, _ in _iter_template_positions(target_sheet_key, templ_headers):
            if _header_match(templ_in, t):
                st.session_state[mkey] = raw
                _touch_tick(target_sheet_key, templ_norm, pos)
                applied += 1
                break
    return applied


def _export_mapping_df(present_sheets: Dict[str, str],
                       templ_headers_by_sheet: Dict[str, List[str]],
                       pos_to_col_by_sheet: Dict[str, Dict[int, int]],
                       only_sheet_key: Optional[str] = None) -> pd.DataFrame:
    rows = []
    for sk in _ordered_keys(present_sheets):
        if only_sheet_key and sk != only_sheet_key:
            continue
        actual_name = present_sheets[sk]
        templ_headers = templ_headers_by_sheet.get(sk, [])
        pos_to_col = pos_to_col_by_sheet.get(sk, {})
        for t, templ_norm, pos, mkey, _ in _iter_template_positions(sk, templ_headers):
            v = st.session_state.get(mkey, "")
            col_letter = get_column_letter(pos_to_col.get(pos, 0)) if pos_to_col.get(pos, 0) else ""
            rows.append({
                "Sheet": TARGET_SHEETS_CANON.get(sk, actual_name),
                "Position": pos,
                "Column": col_letter,
                "Template": t,
                "Raw": v
            })
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
                      header_rows: List[int],
                      start_row: int,
                      raw_df: pd.DataFrame,
                      dup_headers_to_highlight: List[str]) -> Tuple[int, List[str]]:
    headers, pos_to_col, norm_to_cols = _extract_headers_composite(ws, header_rows)

    write_plan: List[Tuple[int, str]] = []
    missing: List[str] = []
    for m in mapping:
        pos = m.get("templ_pos")
        if isinstance(pos, int) and pos in pos_to_col:
            col_idx = pos_to_col[pos]
            write_plan.append((col_idx, m["raw_header"]))
        else:
            norm = _norm_key(m["template_header"])
            if norm in norm_to_cols and len(norm_to_cols[norm]) > 0:
                write_plan.append((norm_to_cols[norm][0], m["raw_header"]))
            else:
                missing.append(m["template_header"])

    nrows = len(raw_df)
    for i in range(nrows):
        excel_row = start_row + i
        for col_idx, raw_name in write_plan:
            val = raw_df.iloc[i][raw_name] if raw_name in raw_df.columns else None
            ws.cell(row=excel_row, column=col_idx, value=val)

    # Duplicate highlighting
    dup_norms = [_norm_key(x) for x in dup_headers_to_highlight if str(x).strip()]
    for want_norm in dup_norms:
        match_cols = set()
        for norm_h, cols in norm_to_cols.items():
            if want_norm == norm_h or want_norm in norm_h:
                match_cols.update(cols)
        for cidx in match_cols:
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
                        header_rows: List[int],
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
            header_rows=header_rows,
            start_row=st.session_state.get("data_start_row", DEFAULT_DATA_START_ROW),
            raw_df=raw_df,
            dup_headers_to_highlight=dup_columns_to_highlight,
        )

    bio_out = BytesIO()
    wb.save(bio_out)  # preserves formatting, heights, colors, validations; keep_vba preserves macros
    return bio_out.getvalue()


# ----- ordering / mapping workbook helpers -----
def _ordered_keys(present_sheets: Dict[str, str]) -> List[str]:
    ordered = [k for k in CANON_ORDER if k in present_sheets]
    ordered += [k for k in present_sheets.keys() if k not in ordered]
    return ordered


def _match_mapping_tab(sheet_names: List[str], target_key: str) -> Optional[str]:
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
    ("pos_to_col_by_sheet", {}),
    ("templ_header_sigs", {}),
    ("raw_headers", []),
    ("raw_sig", ""),
    ("raw_prev_headers", []),
    ("current_sheet_key", None),
    ("last_edit_tick", 0),
    ("download_payload", None),
    ("saved_mapping", {}),
    ("header_rows", DEFAULT_HEADER_ROWS),
    ("data_start_row", DEFAULT_DATA_START_ROW),
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
tab1, tab2, tab3 = st.tabs([
    "Upload Masterfile Template",
    "Upload Raw / Onboarding Data",
    "Mapping (Horizontal) & Download"
])


# -----------------------------
# Tab 1: Upload Masterfile Template
# -----------------------------
with tab1:
    st.markdown("#### Step 1 of 3 — Upload Masterfile Template (.xlsx / .xlsm)")

    cols_hdr = st.columns((2, 1))
    with cols_hdr[0]:
        hdr_str_default = ",".join(str(x) for x in st.session_state.get("header_rows", DEFAULT_HEADER_ROWS))
        hdr_str = st.text_input(
            "Header rows to combine (comma‑separated, top→bottom)",
            value=hdr_str_default,
            help="Example: 3,4,5 — the app will join the texts in those rows (same column) to form one header."
        )
        st.session_state["header_rows"] = _parse_rows_input(hdr_str, DEFAULT_HEADER_ROWS)
    with cols_hdr[1]:
        dsr = st.number_input("Data start row", min_value=2, step=1,
                              value=st.session_state.get("data_start_row", DEFAULT_DATA_START_ROW))
        st.session_state["data_start_row"] = int(dsr)

    tpl = st.file_uploader("Upload template", type=["xlsx", "xlsm"], key="template_uploader")
    do_rescan = False
    if st.session_state.get("template_bytes"):
        do_rescan = st.button("Re‑scan headers using current rows")

    if tpl is not None or do_rescan:
        if tpl is not None:
            raw_bytes = tpl.read()
            st.session_state["template_bytes"] = raw_bytes
            st.session_state["template_ext"] = "xlsm" if tpl.name.lower().endswith(".xlsm") else "xlsx"
        else:
            raw_bytes = st.session_state["template_bytes"]

        is_xlsm = st.session_state["template_ext"] == "xlsm"
        wb = load_workbook(BytesIO(raw_bytes), read_only=False, keep_vba=is_xlsm, data_only=False)
        all_names = wb.sheetnames
        present = _find_target_sheets(all_names)

        templ_headers_by_sheet: Dict[str, List[str]] = {}
        pos_to_col_by_sheet: Dict[str, Dict[int, int]] = {}
        templ_header_sigs: Dict[str, str] = {}

        for sk, actual in present.items():
            ws = wb[actual]
            headers, pos_to_col, _ = _extract_headers_composite(ws, st.session_state["header_rows"])
            templ_headers_by_sheet[sk] = headers
            pos_to_col_by_sheet[sk] = pos_to_col
            templ_header_sigs[sk] = "|".join(headers)

        prev_sigs = st.session_state.get("templ_header_sigs", {})
        st.session_state["present_sheets"] = present
        st.session_state["templ_headers_by_sheet"] = templ_headers_by_sheet
        st.session_state["pos_to_col_by_sheet"] = pos_to_col_by_sheet
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

        ordered = _ordered_keys(present)
        if ordered and st.session_state.get("current_sheet_key") not in present:
            st.session_state["current_sheet_key"] = ordered[0]

        st.success("Template loaded & headers scanned.")
        if present:
            st.caption(f"Composite headers from rows: {', '.join(str(r) for r in st.session_state['header_rows'])}. "
                       f"Data writes from row {st.session_state['data_start_row']}.")
            st.write("Detected target sheets:")
            for sk in _ordered_keys(present):
                display = TARGET_SHEETS_CANON.get(sk, present[sk])
                st.write(f"- **{display}**")
        else:
            st.warning("No target sheets found. Expected PCSE and/or TIC.")


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
                for k in list(st.session_state.keys()):
                    if k.startswith(f"map_{sk}_"):
                        if st.session_state.get(k, "") in removed:
                            st.session_state[k] = ""
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
# Tab 3: Mapping (Horizontal) & Download
# -----------------------------
with tab3:
    st.markdown("#### Step 3 of 3 — Mapping (Horizontal) & Download")

    if not st.session_state.get("template_bytes"):
        st.info("Please upload a masterfile template in tab 1.")
        st.stop()
    if "raw_df_payload" not in st.session_state:
        st.info("Please upload raw/onboarding data in tab 2.")
        st.stop()

    present = st.session_state["present_sheets"]
    templ_headers_by_sheet = st.session_state["templ_headers_by_sheet"]
    pos_to_col_by_sheet = st.session_state.get("pos_to_col_by_sheet", {})
    raw_headers = st.session_state["raw_headers"]
    raw_df: pd.DataFrame = st.session_state["raw_df_payload"]

    if not present:
        st.error("Template has no target sheets to map (expected PCSE and/or TIC).")
        st.stop()

    ordered_keys_list = _ordered_keys(present)
    display_names = [TARGET_SHEETS_CANON.get(sk, present[sk]) for sk in ordered_keys_list]
    key_by_display = {TARGET_SHEETS_CANON.get(sk, present[sk]): sk for sk in ordered_keys_list}

    # Sheet radio
    current_key = st.session_state.get("current_sheet_key", ordered_keys_list[0])
    default_index = display_names.index(TARGET_SHEETS_CANON.get(current_key, present.get(current_key, ""))) \
        if current_key in present else 0
    selected_display = st.radio("Choose target sheet", options=display_names, index=default_index, horizontal=True)
    sheet_key = key_by_display[selected_display]
    st.session_state["current_sheet_key"] = sheet_key

    # Prepare preview
    wb_prev = load_workbook(BytesIO(st.session_state["template_bytes"]),
                            read_only=False,
                            keep_vba=(st.session_state["template_ext"] == "xlsm"),
                            data_only=False)
    ws_prev = wb_prev[present[sheet_key]]

    # Styled preview (rows 1–6)
    st.markdown("**Template preview (rows 1–6)**")
    preview_html = _build_preview_html(ws_prev, max_rows=6)
    # Heuristic height: sum of row heights + some margin, with cap
    height_guess = min(560, max(240, sum([22]*6) + 80))
    st.components.v1.html(preview_html, height=height_guess, scrolling=True)

    # Build composite headers for mapping (using selected header rows)
    headers_for_map, pos_to_col, _ = _extract_headers_composite(ws_prev, st.session_state["header_rows"])
    templ_headers_by_sheet[sheet_key] = headers_for_map
    pos_to_col_by_sheet[sheet_key] = pos_to_col
    st.session_state["templ_headers_by_sheet"] = templ_headers_by_sheet
    st.session_state["pos_to_col_by_sheet"] = pos_to_col_by_sheet

    templ_headers = templ_headers_by_sheet.get(sheet_key, [])
    _hydrate_live_from_saved(sheet_key, templ_headers)

    st.markdown("**Mapping row (choose a raw column for each template column below)**")
    if not raw_headers:
        st.warning("Upload raw data first to enable mapping.")
    else:
        # Horizontal one-row editor: a selectbox per column
        options = [""] + raw_headers
        initial = {}
        col_config = {}
        for t, templ_norm, pos, mkey, _ in _iter_template_positions(sheet_key, templ_headers):
            excel_letter = get_column_letter(pos_to_col.get(pos, 0)) if pos_to_col.get(pos, 0) else ""
            label = f"{excel_letter} {t}".strip()
            initial[label] = st.session_state.get(mkey, "")
            col_config[label] = st.column_config.SelectboxColumn(
                label=label,
                options=options,
                required=False
            )
        map_df_seed = pd.DataFrame([initial])
        edited_df = st.data_editor(
            map_df_seed,
            column_config=col_config,
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            key=f"map_editor_{sheet_key}"
        )
        # push edits back
        if isinstance(edited_df, pd.DataFrame) and len(edited_df) >= 1:
            row0 = edited_df.iloc[0].to_dict()
            # reverse lookup to session keys
            # row0 keys look like "G Product Identifiers | Product ID"
            # We'll match by suffix (the composite header text)
            for t, templ_norm, pos, mkey, _ in _iter_template_positions(sheet_key, templ_headers):
                excel_letter = get_column_letter(pos_to_col.get(pos, 0)) if pos_to_col.get(pos, 0) else ""
                key_label = f"{excel_letter} {t}".strip()
                new_val = row0.get(key_label, "")
                if new_val != st.session_state.get(mkey, ""):
                    st.session_state[mkey] = new_val or ""
                    if new_val:
                        _touch_tick(sheet_key, templ_norm, pos)

    st.markdown("---")
    cols_tools = st.columns((1, 1, 1))
    with cols_tools[0]:
        if st.button("Auto‑map (exact)", use_container_width=True):
            _auto_map_exact(sheet_key, templ_headers, raw_headers)
            st.toast("Exact auto‑map applied (blanks only).")
    with cols_tools[1]:
        fuzz_thr = st.slider("Fuzzy threshold", 0, 100, 80, 1)
        if st.button("Auto‑map (fuzzy)", use_container_width=True):
            _auto_map_fuzzy(sheet_key, templ_headers, raw_headers, threshold=fuzz_thr)
            st.toast(f"Fuzzy auto‑map applied at threshold {fuzz_thr} (blanks only).")
    with cols_tools[2]:
        st.caption("")

    st.markdown("---")
    st.markdown("**Import mapping workbook (.xlsx with two tabs: PCSE & TIC)**")
    st.caption("Provide one Excel file containing two sheets — one for **Product Content And Site Exp** and one for **Trade Item Configurations**. Each tab must have columns: **Template**, **Raw**.")
    imp_wb = st.file_uploader("Upload mapping workbook (.xlsx)", type=["xlsx"], key="map_workbook")
    if imp_wb is not None:
        try:
            xl = pd.ExcelFile(imp_wb)
            applied_summary = []
            for sk in ordered_keys_list:
                tab_name = _match_mapping_tab(xl.sheet_names, sk)
                if not tab_name:
                    applied_summary.append(f"{TARGET_SHEETS_CANON.get(sk, present[sk])}: no matching tab found")
                    continue
                df_imp = xl.parse(tab_name)
                cnt = _import_mapping_df_for_sheet(df_imp, sk, templ_headers_by_sheet, raw_headers)
                applied_summary.append(f"{TARGET_SHEETS_CANON.get(sk, present[sk])}: {cnt} rows applied from '{tab_name}'")
            st.success("Import complete:\n- " + "\n- ".join(applied_summary))
        except Exception as e:
            st.error(f"Failed to import mapping workbook: {e}")

    st.markdown("---")
    export_df_all = _export_mapping_df(present, templ_headers_by_sheet, pos_to_col_by_sheet)
    st.download_button(
        "Export mapping (all sheets)",
        data=export_df_all.to_csv(index=False).encode("utf-8"),
        file_name="mapping_export_all.csv",
        mime="text/csv",
        use_container_width=True,
    )
    export_df_selected = _export_mapping_df(present, templ_headers_by_sheet, pos_to_col_by_sheet, only_sheet_key=sheet_key)
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

    cols_btn = st.columns((1, 1, 1))
    with cols_btn[0]:
        if st.button("Save mapping (selected sheet)", use_container_width=True):
            _commit_current_sheet_mapping(sheet_key, templ_headers)
            st.success(f"Saved mapping for '{TARGET_SHEETS_CANON.get(sheet_key, present.get(sheet_key, sheet_key))}'.")
    with cols_btn[1]:
        if st.button("Save & Build (process both sheets)", use_container_width=True):
            try:
                _commit_current_sheet_mapping(sheet_key, templ_headers)
                payload = _build_output_bytes(
                    template_bytes=st.session_state["template_bytes"],
                    template_is_xlsm=is_xlsm,
                    sheets_to_process=present,
                    templ_headers_by_sheet=templ_headers_by_sheet,
                    header_rows=st.session_state.get("header_rows", DEFAULT_HEADER_ROWS),
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
    with cols_btn[2]:
        if st.button("Build (selected sheet only)", use_container_width=True):
            try:
                payload = _build_output_bytes(
                    template_bytes=st.session_state["template_bytes"],
                    template_is_xlsm=is_xlsm,
                    sheets_to_process={sheet_key: present[sheet_key]},
                    templ_headers_by_sheet=templ_headers_by_sheet,
                    header_rows=st.session_state.get("header_rows", DEFAULT_HEADER_ROWS),
                    raw_df=raw_df,
                    auto_resolve_dupe_mappings=auto_resolve,
                    dup_columns_to_highlight=dup_cols,
                    saved_mapping_store=None,  # use live mapping
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
