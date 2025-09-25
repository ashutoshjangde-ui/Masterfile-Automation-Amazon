# app_masterfile.py
import io
import json
import re
import time
from difflib import SequenceMatcher
from textwrap import dedent
from pathlib import Path
import tempfile
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# Try fast path
try:
    import xlwings as xw  # Windows + Excel only
    XLWINGS_AVAILABLE = True
except Exception:
    XLWINGS_AVAILABLE = False

st.set_page_config(page_title="Masterfile Automation (Template fast write)", page_icon="📦", layout="wide")

# -------- Masterfile layout (your template) --------
MASTER_TEMPLATE_SHEET = "Template"   # write only here
MASTER_DISPLAY_ROW    = 2            # mapping row in master
MASTER_DATA_START_ROW = 4            # first data row in master

# -------- Helpers --------
def norm(s: str) -> str:
    if s is None:
        return ""
    x = str(s).strip().lower()
    x = re.sub(r"\s*-\s*en\s*[-_ ]\s*us\s*$", "", x)
    x = x.replace("–","-").replace("—","-").replace("−","-")
    x = re.sub(r"[._/\\-]+", " ", x)
    x = re.sub(r"[^0-9a-z\s]+", " ", x)
    return re.sub(r"\s+", " ", x).strip()

def top_matches(query, candidates, k=3):
    q = norm(query)
    scored = [(SequenceMatcher(None, q, norm(c)).ratio(), c) for c in candidates]
    scored.sort(key=lambda t: t[0], reverse=True)
    return scored[:k]

def worksheet_used_cols(ws, header_rows=(1,), hard_cap=2048, empty_streak_stop=8):
    max_try = min(ws.max_column, hard_cap)
    last_nonempty, streak = 0, 0
    for c in range(1, max_try + 1):
        any_val = any((ws.cell(row=r, column=c).value not in (None, "")) for r in header_rows)
        if any_val:
            last_nonempty, streak = c, 0
        else:
            streak += 1
            if streak >= empty_streak_stop:
                break
    return max(last_nonempty, 1)

def nonempty_rows(df: pd.DataFrame) -> int:
    if df.empty: return 0
    return df.replace("", pd.NA).dropna(how="all").shape[0]

def pick_best_onboarding_sheet(uploaded_file, mapping_aliases_by_master):
    uploaded_file.seek(0)
    xl = pd.ExcelFile(uploaded_file)
    best, best_score, best_info = None, -1, ""
    for sheet in xl.sheet_names:
        try:
            df = xl.parse(sheet_name=sheet, header=0, dtype=str).fillna("")
            df.columns = [str(c).strip() for c in df.columns]
        except Exception:
            continue
        header_set = {norm(c) for c in df.columns}
        matches = sum(any(norm(a) in header_set for a in aliases)
                      for aliases in mapping_aliases_by_master.values())
        rows = nonempty_rows(df)
        score = matches + (0.01 if rows > 0 else 0.0)
        if score > best_score:
            best, best_score = (df, sheet), score
            best_info = f"matched headers: {matches}, non-empty rows: {rows}"
    if best is None:
        raise ValueError("No readable onboarding sheet found.")
    return best[0], best[1], best_info

SENTINEL_LIST = object()

# -------- UI --------
st.title("📦 Masterfile Automation — Template fast writer")
st.caption("Fills **only** the Template sheet and preserves all other sheets/styles. Use the Excel-fast writer for huge speedups on Windows.")

with st.expander("ℹ️ Instructions", expanded=True):
    st.markdown(dedent(f"""
    **Masterfile (.xlsx)**  
    - Sheet: **{MASTER_TEMPLATE_SHEET}**  
    - Row 1 & 3 = internal keys (preserved)  
    - Row **{MASTER_DISPLAY_ROW}** = mapping headers  
    - Data written from **Row {MASTER_DATA_START_ROW}**.

    **Onboarding (.xlsx)**  
    - Row 1 = headers; Row 2+ = data (best sheet auto-selected)

    **Mapping JSON** (keys = master display headers; values = ordered aliases):
    ```json
    {{
      "Partner SKU": ["Target SKU","Seller SKU","SKU","item_sku"],
      "Barcode": ["UPC/EAN","UPC","Product ID","barcode","barcode.value"],
      "Brand": ["Brand Name","brand_name","Walmart Brand Name - en-US"],
      "Product Title": ["Item Name","Product Name","Title"],
      "Description": ["Long Description","Product Description","Description"]
    }}
    ```
    """))

c1, c2 = st.columns([1, 1])
with c1:
    masterfile_file = st.file_uploader("📄 Masterfile Template (.xlsx)", type=["xlsx"])
with c2:
    onboarding_file = st.file_uploader("🧾 Onboarding (.xlsx)", type=["xlsx"])

st.markdown("#### 🔗 Mapping JSON")
tab1, tab2 = st.tabs(["Paste JSON", "Upload JSON"])
mapping_json_text, mapping_json_file = "", None
with tab1:
    mapping_json_text = st.text_area("Paste mapping JSON", height=200,
                                     placeholder='{\n  "Partner SKU": ["Seller SKU", "item_sku"]\n}')
with tab2:
    mapping_json_file = st.file_uploader("Or upload mapping.json", type=["json"], key="mapping_file")

use_fast = st.checkbox("⚡ Use Excel-fast writer (xlwings, Windows + Excel)", value=XLWINGS_AVAILABLE, disabled=not XLWINGS_AVAILABLE,
                       help="Writes the whole data block in one shot via Excel. Falls back to openpyxl if unavailable.")

st.divider()
go = st.button("🚀 Generate Final Masterfile", type="primary")

# -------- Main --------
if go:
    log = st.empty()
    def slog(msg): log.markdown(msg)

    if not masterfile_file or not onboarding_file:
        st.error("Please upload both **Masterfile Template** and **Onboarding**.")
        st.stop()

    # Parse mapping JSON
    try:
        mapping_raw = json.loads(mapping_json_text) if mapping_json_text.strip() else json.load(mapping_json_file)
    except Exception as e:
        st.error(f"Mapping JSON parse error: {e}")
        st.stop()
    if not isinstance(mapping_raw, dict):
        st.error("Mapping JSON must be an object: {\"Master header\": [aliases...]}.")
        st.stop()

    # Normalize mapping { master_norm: [alias1, alias2, ..., fallback master display] }
    mapping_aliases = {}
    for k, v in mapping_raw.items():
        aliases = v[:] if isinstance(v, list) else [v]
        if k not in aliases:
            aliases.append(k)
        mapping_aliases[norm(k)] = aliases

    # Read Template headers quickly
    masterfile_file.seek(0)
    master_bytes = masterfile_file.read()

    slog("⏳ Reading Template headers…")
    t0 = time.time()
    wb_ro = load_workbook(io.BytesIO(master_bytes), read_only=True, data_only=True, keep_links=True)
    if MASTER_TEMPLATE_SHEET not in wb_ro.sheetnames:
        st.error(f"Sheet **'{MASTER_TEMPLATE_SHEET}'** not found in the masterfile.")
        st.stop()
    ws_ro = wb_ro[MASTER_TEMPLATE_SHEET]
    used_cols = worksheet_used_cols(ws_ro, header_rows=(MASTER_DISPLAY_ROW,), hard_cap=2048, empty_streak_stop=8)
    display_headers = [ws_ro.cell(row=MASTER_DISPLAY_ROW, column=c).value or "" for c in range(1, used_cols+1)]
    wb_ro.close()
    slog(f"✅ Template headers loaded (cols={used_cols}) in {time.time()-t0:.2f}s")

    # Pick best onboarding sheet
    try:
        best_df, best_sheet, info = pick_best_onboarding_sheet(onboarding_file, mapping_aliases)
    except Exception as e:
        st.error(f"Onboarding error: {e}")
        st.stop()
    on_df = best_df.fillna("")
    on_df.columns = [str(c).strip() for c in on_df.columns]
    on_headers = list(on_df.columns)
    st.success(f"Using onboarding sheet: **{best_sheet}** ({info})")

    # Build mapping: master col -> source Series (or SENTINEL_LIST)
    series_by_alias = {norm(h): on_df[h] for h in on_headers}
    master_to_source, report_lines, unmatched = {}, [], []
    report_lines.append("#### 🔎 Mapping Summary (Template)")

    for c, disp in enumerate(display_headers, start=1):
        disp_norm = norm(disp)
        if not disp_norm:
            continue
        aliases = mapping_aliases.get(disp_norm, [disp])
        resolved = None
        for a in aliases:
            s = series_by_alias.get(norm(a))
            if s is not None:
                resolved = s
                report_lines.append(f"- ✅ **{disp}** ← `{a}`")
                break
        if resolved is not None:
            master_to_source[c] = resolved
        else:
            if disp_norm == norm("Listing Action (List or Unlist)"):
                master_to_source[c] = SENTINEL_LIST
                report_lines.append(f"- 🟨 **{disp}** ← (will fill `'List'`)")
            else:
                unmatched.append(disp or f"Col {c}")
                sugg = top_matches(disp, on_headers, 3)
                sug_txt = ", ".join(f"`{n}` ({round(sc*100,1)}%)" for sc, n in sugg) if sugg else "*none*"
                report_lines.append(f"- ❌ **{disp}** ← *no match*. Suggestions: {sug_txt}")

    st.markdown("\n".join(report_lines))

    n_rows = len(on_df)

    # -------- FAST WRITE PATH (xlwings) --------
    if use_fast and XLWINGS_AVAILABLE:
        slog("⚡ Using Excel-fast writer (xlwings)…")
        t_write = time.time()

        # Build 2-D block once (num_rows x used_cols)
        block = [[""] * used_cols for _ in range(n_rows)]
        for col, src in master_to_source.items():
            if src is SENTINEL_LIST:
                for i in range(n_rows):
                    block[i][col-1] = "List"
            else:
                vals = src.astype(str).tolist()
                m = min(len(vals), n_rows)
                for i in range(m):
                    v = vals[i].strip()
                    if v and v.lower() not in ("nan", "none"):
                        block[i][col-1] = v

        # Save master bytes to a temp file; open with Excel invisibly; paste the whole block once
        with tempfile.TemporaryDirectory() as td:
            src_path = Path(td) / "master.xlsx"
            dst_path = Path(td) / "final_masterfile.xlsx"
            src_path.write_bytes(master_bytes)

            app = xw.App(visible=False)  # no Excel window
            try:
                wb = xw.Book(str(src_path))
                ws = wb.sheets[MASTER_TEMPLATE_SHEET]

                # Top-left cell for the data block
                start_cell = f"A{MASTER_DATA_START_ROW}"
                # If Template has headers before column A, still safe: we filled block width = used_cols
                ws.range(start_cell).options(expand=False).value = block  # one-shot write

                wb.save(str(dst_path))
                wb.close()
            finally:
                app.quit()

            out_bytes = dst_path.read_bytes()

        slog(f"✅ Wrote & saved via Excel in {time.time()-t_write:.2f}s")
        st.download_button(
            "⬇️ Download Final Masterfile",
            data=out_bytes,
            file_name="final_masterfile.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_fast",
        )

    # -------- Fallback: openpyxl (slower) --------
    else:
        slog("🛠️ Writing via openpyxl (fallback)…")
        t_write = time.time()
        wb = load_workbook(io.BytesIO(master_bytes), read_only=False, data_only=False, keep_links=True)
        ws = wb[MASTER_TEMPLATE_SHEET]

        # Convert series once
        col_value_lists = {}
        for col, src in master_to_source.items():
            if src is SENTINEL_LIST:
                continue
            col_value_lists[col] = src.astype(str).tolist()

        # Write row by row
        prog = st.progress(0)
        total = max(1, n_rows)
        for i in range(n_rows):
            row_idx = MASTER_DATA_START_ROW + i
            for col, src in master_to_source.items():
                if src is SENTINEL_LIST:
                    ws.cell(row=row_idx, column=col, value="List")
                else:
                    vals = col_value_lists[col]
                    if i < len(vals):
                        v = vals[i].strip()
                        if v and v.lower() not in ("nan", "none", ""):
                            ws.cell(row=row_idx, column=col, value=v)
            if (i+1) % max(1, n_rows // 50) == 0:
                prog.progress((i+1)/total)

        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)
        out_bytes = bio.getvalue()
        slog(f"✅ Wrote & saved via openpyxl in {time.time()-t_write:.2f}s")

        st.download_button(
            "⬇️ Download Final Masterfile",
            data=out_bytes,
            file_name="final_masterfile.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_fallback",
        )
