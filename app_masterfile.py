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

# Try fast path (desktop Excel via COM)
try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except Exception:
    XLWINGS_AVAILABLE = False

# ─────────────────────────────────────────────────────────────────────
# Page meta + subtle theming (visuals only; core logic unchanged)
# ─────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Masterfile Automation - Amazon",
    page_icon="🧾",
    layout="wide"
)

# Soft palette + “cards”
st.markdown("""
<style>
.stApp { background: linear-gradient(180deg, #f8fafc 0%, #ffffff 80%); }
.block-container { padding-top: 1rem; }
.section {
  border: 1px solid #eef2f7; background: #ffffff; border-radius: 14px;
  padding: 18px 20px; box-shadow: 0 4px 18px rgba(2, 6, 23, 0.04); margin-bottom: 18px;
}
.badge { display:inline-block; padding:3px 10px; border-radius:999px; font-size:.82rem; font-weight:600; letter-spacing:.2px; }
.badge-info { background:#eef2ff; color:#1e40af; }
.badge-ok { background:#ecfdf5; color:#065f46; }
.badge-warn { background:#fff7ed; color:#9a3412; }
.small-note { color:#64748b; font-size:.92rem; }
h1, h2, h3 { color:#0f172a; } hr { border-color:#f1f5f9; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────
# Masterfile layout (your template)
# ─────────────────────────────────────────────────────────────────────
MASTER_TEMPLATE_SHEET = "Template"   # write only here
MASTER_DISPLAY_ROW    = 2            # mapping row in master (normal headers)
MASTER_SECONDARY_ROW  = 3            # ONLY used to disambiguate bullet points
MASTER_DATA_START_ROW = 4            # first data row in master

# ─────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────
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

# ─────────────────────────────────────────────────────────────────────
# UI Header
# ─────────────────────────────────────────────────────────────────────
st.title("🧾 Masterfile Automation – Amazon")
st.caption("Fills **only** the Template sheet and preserves all other sheets/styles. Use the Excel-fast writer for big files on Windows.")

fast_badge = "badge-ok" if XLWINGS_AVAILABLE else "badge-warn"
fast_text  = "Excel-fast writer available" if XLWINGS_AVAILABLE else "Excel-fast writer unavailable"
st.markdown(f"<span class='badge {fast_badge}'>{fast_text}</span>  "
            f"<span class='badge badge-info'>Template-only writer</span>",
            unsafe_allow_html=True)

with st.expander("ℹ️ Instructions", expanded=True):
    st.markdown(dedent(f"""
    **Masterfile (.xlsx)**  
    - Sheet: **{MASTER_TEMPLATE_SHEET}**  
    - Row 1 & 3 = internal keys (preserved)  
    - Row **{MASTER_DISPLAY_ROW}** = mapping headers  
    - For **Key Product Features** only, we use **Row {MASTER_SECONDARY_ROW}** (e.g., `bullet_point1..5`) as the mapping header.  
    - Data written from **Row {MASTER_DATA_START_ROW}**.

    **Onboarding (.xlsx)**  
    - Row 1 = headers; Row 2+ = data *(best sheet auto-selected)*
    """))

# ─────────────────────────────────────────────────────────────────────
# Upload + Mapping inputs (in a “card”)
# ─────────────────────────────────────────────────────────────────────
st.markdown("<div class='section'>", unsafe_allow_html=True)
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

# ➜ Leave the toggle enabled so you can force fast path on Windows with Excel
use_fast = st.checkbox(
    "⚡ Use Excel-fast writer (xlwings, Windows + Excel)",
    value=True,
    help="If Excel isn’t available, the app will fall back to openpyxl automatically."
)
st.markdown("</div>", unsafe_allow_html=True)

st.divider()
go = st.button("🚀 Generate Final Masterfile", type="primary")

# ─────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────
if go:
    st.markdown("<div class='section'>", unsafe_allow_html=True)
    st.markdown("### 📝 Log")
    log = st.empty()
    def slog(msg): log.markdown(msg)

    if not masterfile_file or not onboarding_file:
        st.error("Please upload both **Masterfile Template** and **Onboarding**.")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()

    # Parse mapping JSON
    try:
        mapping_raw = json.loads(mapping_json_text) if mapping_json_text.strip() else json.load(mapping_json_file)
    except Exception as e:
        st.error(f"Mapping JSON parse error: {e}")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()
    if not isinstance(mapping_raw, dict):
        st.error("Mapping JSON must be an object: {\"Master header\": [aliases...]}.")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()

    # Normalize mapping { master_norm: [alias1, alias2, ..., fallback master display] }
    mapping_aliases = {}
    for k, v in mapping_raw.items():
        aliases = v[:] if isinstance(v, list) else [v]
        if k not in aliases:
            aliases.append(k)
        mapping_aliases[norm(k)] = aliases

    # Read Template headers
    masterfile_file.seek(0)
    master_bytes = masterfile_file.read()

    slog("⏳ Reading Template headers…")
    t0 = time.time()
    wb_ro = load_workbook(io.BytesIO(master_bytes), read_only=True, data_only=True, keep_links=True)
    if MASTER_TEMPLATE_SHEET not in wb_ro.sheetnames:
        st.error(f"Sheet **'{MASTER_TEMPLATE_SHEET}'** not found in the masterfile.")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()
    ws_ro = wb_ro[MASTER_TEMPLATE_SHEET]
    used_cols = worksheet_used_cols(ws_ro, header_rows=(MASTER_DISPLAY_ROW, MASTER_SECONDARY_ROW), hard_cap=2048, empty_streak_stop=8)
    display_headers   = [ws_ro.cell(row=MASTER_DISPLAY_ROW,   column=c).value or "" for c in range(1, used_cols+1)]
    secondary_headers = [ws_ro.cell(row=MASTER_SECONDARY_ROW, column=c).value or "" for c in range(1, used_cols+1)]
    wb_ro.close()
    slog(f"✅ Template headers loaded (cols={used_cols}) in {time.time()-t0:.2f}s")

    # Pick best onboarding sheet
    try:
        best_df, best_sheet, info = pick_best_onboarding_sheet(onboarding_file, mapping_aliases)
    except Exception as e:
        st.error(f"Onboarding error: {e}")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()
    on_df = best_df.fillna("")
    on_df.columns = [str(c).strip() for c in on_df.columns]
    on_headers = list(on_df.columns)
    st.success(f"Using onboarding sheet: **{best_sheet}** ({info})")

    # Build mapping: master col -> source Series (or SENTINEL_LIST)
    series_by_alias = {norm(h): on_df[h] for h in on_headers}
    master_to_source, report_lines, unmatched = {}, [], []
    report_lines.append("#### 🔎 Mapping Summary (Template)")

    BULLET_DISP_N = norm("Key Product Features")

    for c, (disp, sec) in enumerate(zip(display_headers, secondary_headers), start=1):
        disp_norm = norm(disp)
        sec_norm  = norm(sec)

        if disp_norm == BULLET_DISP_N and sec_norm:
            effective_header = sec  # e.g., 'bullet_point1'
            label_for_log = f"{disp} ({sec})"
        else:
            effective_header = disp
            label_for_log = disp

        eff_norm = norm(effective_header)
        if not eff_norm:
            continue

        aliases = mapping_aliases.get(eff_norm, [effective_header])
        resolved = None
        for a in aliases:
            s = series_by_alias.get(norm(a))
            if s is not None:
                resolved = s
                report_lines.append(f"- ✅ **{label_for_log}** ← `{a}`")
                break

        if resolved is not None:
            master_to_source[c] = resolved
        else:
            if disp_norm == norm("Listing Action (List or Unlist)"):
                master_to_source[c] = SENTINEL_LIST
                report_lines.append(f"- 🟨 **{label_for_log}** ← (will fill `'List'`)")
            else:
                unmatched.append(label_for_log or f"Col {c}")
                sugg = top_matches(effective_header, on_headers, 3)
                sug_txt = ", ".join(f"`{name}` ({round(sc*100,1)}%)" for sc, name in sugg) if sugg else "*none*"
                report_lines.append(f"- ❌ **{label_for_log}** ← *no match*. Suggestions: {sug_txt}")

    st.markdown("\n".join(report_lines))

    n_rows = len(on_df)

    # ── Excel-fast path (xlwings) ─────────────────────────────────────
    if use_fast:
        if not XLWINGS_AVAILABLE:
            st.warning("Excel-fast writer requested, but xlwings/Excel isn’t available in this runtime. Falling back to openpyxl.")
        else:
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

            with tempfile.TemporaryDirectory() as td:
                src_path = Path(td) / "master.xlsx"
                dst_path = Path(td) / "final_masterfile.xlsx"
                src_path.write_bytes(master_bytes)

                app = xw.App(visible=False)
                try:
                    wb = xw.Book(str(src_path))
                    ws = wb.sheets[MASTER_TEMPLATE_SHEET]
                    start_cell = f"A{MASTER_DATA_START_ROW}"
                    ws.range(start_cell).options(expand=False).value = block
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
            st.markdown("</div>", unsafe_allow_html=True)
            st.stop()  # done

    # ── Fallback: openpyxl (fast enough on cloud) ─────────────────────
    slog("🛠️ Writing via openpyxl (fallback)…")
    t_write = time.time()
    wb = load_workbook(io.BytesIO(master_bytes), read_only=False, data_only=False, keep_links=True)
    ws = wb[MASTER_TEMPLATE_SHEET]

    col_value_lists = {}
    for col, src in master_to_source.items():
        if src is SENTINEL_LIST:
            continue
        col_value_lists[col] = src.astype(str).tolist()

    prog = st.progress(0)
    total = max(1, n_rows)
    step = max(1, n_rows // 50)

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
        if (i + 1) % step == 0:
            prog.progress((i + 1) / total)

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
    st.markdown("</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────
# Footer
# ─────────────────────────────────────────────────────────────────────
st.markdown(
    "<div class='section small-note'>"
    "Tip: The **Excel-fast writer** requires Windows + desktop Excel with permissions to launch. "
    "If unavailable (e.g., Streamlit Cloud), the app automatically uses openpyxl."
    "</div>",
    unsafe_allow_html=True
)
