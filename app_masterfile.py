# app_masterfile.py  — Aspose.Cells fast path only (no xlwings, no openpyxl, no XML patch)
import io
import json
import re
import time
from difflib import SequenceMatcher
from pathlib import Path
from textwrap import dedent

import pandas as pd
import streamlit as st

# ── Require Aspose.Cells ─────────────────────────────────────────────
try:
    import aspose.cells as ac
except Exception as e:
    st.error(
        "Aspose.Cells is required. Add **aspose-cells** to requirements.txt "
        "or run `pip install aspose-cells` locally.\n\n"
        f"Import error: {e}"
    )
    st.stop()

# ── UI chrome ────────────────────────────────────────────────────────
st.set_page_config(page_title="Masterfile Automation - Amazon", page_icon="🧾", layout="wide")
st.markdown("""
<style>
:root{--bg1:#f6f9fc;--bg2:#fff;--card:#fff;--card-border:#e8eef6;--ink:#0f172a;--muted:#64748b;--accent:#2563eb}
.stApp{background:linear-gradient(180deg,var(--bg1) 0%,var(--bg2) 70%)}
.block-container{padding-top:.75rem}
.section{border:1px solid var(--card-border);background:var(--card);border-radius:16px;padding:18px 20px;box-shadow:0 6px 24px rgba(2,6,23,.05);margin-bottom:18px}
h1,h2,h3{color:var(--ink)}hr{border-color:#eef2f7}
.badge{display:inline-block;padding:4px 10px;border-radius:999px;font-size:.82rem;font-weight:600;letter-spacing:.2px;margin-right:.25rem}
.badge-info{background:#eef2ff;color:#1e40af}.badge-ok{background:#ecfdf5;color:#065f46}.small-note{color:var(--muted);font-size:.92rem}
.stDownloadButton>button,div.stButton>button{background:var(--accent)!important;color:#fff!important;border-radius:10px!important;border:0!important;box-shadow:0 8px 18px rgba(37,99,235,.18)}
</style>
""", unsafe_allow_html=True)

# ── Template layout ──────────────────────────────────────────────────
MASTER_TEMPLATE_SHEET = "Template"
MASTER_DISPLAY_ROW    = 2
MASTER_SECONDARY_ROW  = 3
MASTER_DATA_START_ROW = 4

# ── Helpers ──────────────────────────────────────────────────────────
_INVALID_XML_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF]")
def sanitize_text(x):
    if x is None: return ""
    s = str(x).strip()
    if not s: return ""
    return _INVALID_XML_CHARS.sub("", s)

def norm(s: str) -> str:
    if s is None: return ""
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

def used_cols_on_rows_aspose(ws, header_rows=(MASTER_DISPLAY_ROW, MASTER_SECONDARY_ROW),
                             hard_cap=2048, empty_streak_stop=8):
    cells = ws.cells
    last_nonempty, streak = 0, 0
    for c in range(1, hard_cap + 1):
        any_val = False
        for r in header_rows:
            v = cells.get(r-1, c-1).value
            if v not in (None, ""):
                any_val = True
                break
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

# ── UI ───────────────────────────────────────────────────────────────
st.title("🧾 Masterfile Automation – Amazon")
st.caption("Fills only the Template sheet and preserves all other sheets/styles.")
st.markdown("<div class='section'><span class='badge badge-info'>Template-only writer</span> "
            "<span class='badge badge-ok'>Aspose.Cells fast engine</span></div>", unsafe_allow_html=True)

st.markdown("<div class='section'>", unsafe_allow_html=True)
c1, c2 = st.columns([1, 1])
with c1:
    masterfile_file = st.file_uploader("📄 Masterfile Template (.xlsx / .xlsm)", type=["xlsx", "xlsm"])
with c2:
    onboarding_file = st.file_uploader("🧾 Onboarding (.xlsx)", type=["xlsx"])

st.markdown("#### 🔗 Mapping JSON")
tab1, tab2 = st.tabs(["Paste JSON", "Upload JSON"])
mapping_json_text, mapping_json_file = "", None
with tab1:
    mapping_json_text = st.text_area("Paste mapping JSON", height=200,
                                     placeholder='\n{\n  "Partner SKU": ["Seller SKU", "item_sku"]\n}\n')
with tab2:
    mapping_json_file = st.file_uploader("Or upload mapping.json", type=["json"], key="mapping_file")
st.markdown("</div>", unsafe_allow_html=True)

st.divider()
go = st.button("🚀 Generate Final Masterfile", type="primary")

# ── Main ─────────────────────────────────────────────────────────────
if go:
    st.markdown("<div class='section'>", unsafe_allow_html=True)
    st.markdown("### 📝 Log")
    log = st.empty()
    def slog(msg): log.markdown(msg)

    if not masterfile_file or not onboarding_file:
        st.error("Please upload both Masterfile Template and Onboarding.")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()

    ext = (Path(masterfile_file.name).suffix or ".xlsx").lower()
    out_mime = {
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
    }.get(ext, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Parse mapping JSON
    try:
        mapping_raw = json.loads(mapping_json_text) if mapping_json_text.strip() else json.load(mapping_json_file)
    except Exception as e:
        st.error(f"Mapping JSON parse error: {e}")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()
    if not isinstance(mapping_raw, dict):
        st.error("Mapping JSON must be an object: {\"Master header\": [aliases...]}")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()

    mapping_aliases = {}
    for k, v in mapping_raw.items():
        aliases = v[:] if isinstance(v, list) else [v]
        if k not in aliases:
            aliases.append(k)
        mapping_aliases[norm(k)] = aliases

    # Read template headers with Aspose
    masterfile_file.seek(0)
    master_bytes = masterfile_file.read()

    slog("⏳ Reading Template headers (Aspose)…")
    t0 = time.time()
    wb = ac.Workbook(io.BytesIO(master_bytes))
    if wb.worksheets.get_count() == 0 or wb.worksheets.get(MASTER_TEMPLATE_SHEET) is None:
        st.error(f"Sheet '{MASTER_TEMPLATE_SHEET}' not found in masterfile.")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()
    ws = wb.worksheets.get(MASTER_TEMPLATE_SHEET)
    cells = ws.cells

    used_cols = used_cols_on_rows_aspose(ws)
    display_headers   = [(cells.get(MASTER_DISPLAY_ROW-1,   c-1).value or "") for c in range(1, used_cols+1)]
    secondary_headers = [(cells.get(MASTER_SECONDARY_ROW-1, c-1).value or "") for c in range(1, used_cols+1)]
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

    # Map columns
    series_by_alias = {norm(h): on_df[h] for h in on_headers}
    master_to_source, report_lines, unmatched = {}, [], []
    report_lines.append("#### 🔎 Mapping Summary (Template)")
    BULLET_DISP_N = norm("Key Product Features")

    for c, (disp, sec) in enumerate(zip(display_headers, secondary_headers), start=1):
        disp_norm = norm(disp); sec_norm = norm(sec)
        if disp_norm == BULLET_DISP_N and sec_norm:
            effective_header = sec; label_for_log = f"{disp} ({sec})"
        else:
            effective_header = disp; label_for_log = disp

        eff_norm = norm(effective_header)
        if not eff_norm: continue

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
                report_lines.append(f"- 🟨 **{label_for_log}** ← (will fill 'List')")
            else:
                unmatched.append(label_for_log or f"Col {c}")
                sugg = top_matches(effective_header, on_headers, 3)
                sug_txt = ", ".join(f"`{name}` ({round(sc*100,1)}%)" for sc, name in sugg) if sugg else "*none*"
                report_lines.append(f"- ❌ **{label_for_log}** ← no match. Suggestions: {sug_txt}")

    st.markdown("\n".join(report_lines))

    # Build data block
    n_rows = len(on_df)
    block = [[""] * used_cols for _ in range(n_rows)]
    for col, src in master_to_source.items():
        if src is SENTINEL_LIST:
            for i in range(n_rows): block[i][col-1] = "List"
        else:
            vals = src.astype(str).tolist()
            m = min(len(vals), n_rows)
            for i in range(m):
                v = sanitize_text(vals[i])
                if v and v.lower() not in ("nan", "none", ""):
                    block[i][col-1] = v

    # Write with Aspose.Cells (very fast, compiled)
    slog("🚀 Writing via Aspose.Cells…")
    t_write = time.time()

    # Clear old data block first (rows >= start)
    cells.delete_rows(MASTER_DATA_START_ROW-1, max(0, ws.cells.max_data_row - (MASTER_DATA_START_ROW-1) + 1), True)

    # Bulk put values (row by row is fine; engine is fast)
    start_r, start_c = MASTER_DATA_START_ROW-1, 0
    for i, row in enumerate(block):
        if not any(row): 
            continue
        for j, v in enumerate(row):
            if v:
                cells.get(start_r + i, start_c + j).put_value(v)

    # Resize Excel tables (ListObjects) to cover header->last row/col
    try:
        last_row_idx = max(MASTER_DISPLAY_ROW-1, start_r + max(0, n_rows-1))
        last_col_idx = max(0, used_cols-1)
        rng = cells.create_range(MASTER_DISPLAY_ROW-1, 0, last_row_idx-(MASTER_DISPLAY_ROW-1)+1, used_cols)
        los = ws.list_objects
        for k in range(los.get_count()):
            lo = los.get(k)
            lo.resize(rng)
    except Exception:
        pass

    # Make Excel rebuild calculation chain silently (prevents repair prompt)
    wb.settings.set_recalculate_before_save(True)

    # Save to bytes, preserving macro format if .xlsm
    bio = io.BytesIO()
    if ext == ".xlsm":
        wb.save(bio, ac.SaveFormat.XLSM)
    else:
        wb.save(bio, ac.SaveFormat.XLSX)
    bio.seek(0)
    out_bytes = bio.getvalue()

    slog(f"✅ Wrote & saved in {time.time()-t_write:.2f}s")

    st.download_button(
        "⬇️ Download Final Masterfile",
        data=out_bytes,
        file_name=f"final_masterfile{ext}",
        mime=out_mime,
        key="dl_aspose",
    )
    st.markdown("</div>", unsafe_allow_html=True)

with st.expander("📘 How to use", expanded=False):
    st.markdown(dedent(f"""
    - Writes only into `{MASTER_TEMPLATE_SHEET}`; preserves all other tabs, styles, formulas, macros.
    - Headers live in row {MASTER_DISPLAY_ROW}; bullet-point subheaders in row {MASTER_SECONDARY_ROW}.
    - Data starts at row {MASTER_DATA_START_ROW}.
    - Tables on **Template** are automatically resized to the new data area.
    - Workbook is flagged `RecalculateBeforeSave` so Excel opens without *repair* prompts.
    """))
