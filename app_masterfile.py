# app_masterfile.py
import io
import json
import re
import time
import zipfile
import xml.etree.ElementTree as ET
from difflib import SequenceMatcher
from textwrap import dedent
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# ─────────────────────────────────────────────────────────────────────
# Linux-fast XML writer (no xlwings). Patches sheet + tables and strips calcChain.
# ─────────────────────────────────────────────────────────────────────
XL_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XL_NS_REL  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", XL_NS_MAIN)
ET.register_namespace("r", XL_NS_REL)
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")

# Remove characters illegal in XML 1.0
_INVALID_XML_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF]")
CELL_RE  = re.compile(r"^([A-Z]+)(\d+)$")
RANGE_RE = re.compile(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$")

def sanitize_xml_text(s) -> str:
    if s is None:
        return ""
    s = str(s)
    return _INVALID_XML_CHARS.sub("", s)

def _col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n-1, 26)
        s = chr(65+r) + s
    return s

def _col_number(letters: str) -> int:
    n = 0
    for ch in letters:
        if not ch.isalpha():
            break
        n = n * 26 + (ord(ch.upper()) - 64)
    return n

def _parse_range(ref: str):
    m = RANGE_RE.match(ref or "")
    if not m:
        return ("A", 1, "A", 1)
    return m.group(1), int(m.group(2)), m.group(3), int(m.group(4))

def _find_sheet_part_path(z: zipfile.ZipFile, sheet_name: str) -> str:
    wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
    rels_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid = None
    for sh in wb_xml.find(f"{{{XL_NS_MAIN}}}sheets"):
        if sh.attrib.get("name") == sheet_name:
            rid = sh.attrib.get(f"{{{XL_NS_REL}}}id")
            break
    if not rid:
        raise ValueError(f"Sheet '{sheet_name}' not found in workbook.xml")

    target = None
    for rel in rels_xml:
        if rel.attrib.get("Id") == rid:
            target = rel.attrib.get("Target")
            break
    if not target:
        raise ValueError(f"Relationship for sheet '{sheet_name}' not found.")
    target = target.replace("\\", "/")
    if target.startswith("../"):
        target = target[3:]
    if not target.startswith("xl/"):
        target = "xl/" + target
    return target  # e.g., xl/worksheets/sheet1.xml

def _get_table_paths_for_sheet(z: zipfile.ZipFile, sheet_path: str) -> list:
    rels_path = sheet_path.replace("worksheets/", "worksheets/_rels/").replace(".xml", ".xml.rels")
    if rels_path not in z.namelist():
        return []
    root = ET.fromstring(z.read(rels_path))
    out = []
    for rel in root:
        t = rel.attrib.get("Type", "")
        if t.endswith("/table"):
            target = rel.attrib.get("Target", "").replace("\\", "/")
            if target.startswith("../"):
                target = target[3:]
            if not target.startswith("xl/"):
                target = "xl/" + target
            out.append(target)
    return out

def _read_table_info(table_xml_bytes: bytes):
    """Return (start_col_letter, header_row_from_ref, width_by_tableColumns)."""
    root = ET.fromstring(table_xml_bytes)
    ref = root.attrib.get("ref", "A1:A1")
    sc, sr, ec, _ = _parse_range(ref)
    tcols = root.find(f"{{{XL_NS_MAIN}}}tableColumns")
    width = sum(1 for _ in tcols) if tcols is not None else (_col_number(ec) - _col_number(sc) + 1)
    return sc, sr, width

def _union_dimension(orig_dim_ref: str, used_cols: int, last_row: int) -> str:
    try:
        _, right = (orig_dim_ref or "A1:A1").split(":", 1)
        m = CELL_RE.match(right)
        if m:
            orig_last_col = _col_number(m.group(1))
            orig_last_row = int(m.group(2))
        else:
            orig_last_col, orig_last_row = used_cols, last_row
    except Exception:
        orig_last_col, orig_last_row = used_cols, last_row
    u_last_col = max(orig_last_col, used_cols)
    u_last_row = max(orig_last_row, last_row)
    return f"A1:{_col_letter(u_last_col)}{u_last_row}"

def _patch_sheet_xml(sheet_xml_bytes: bytes, start_row: int, used_cols_final: int, block_2d: list) -> bytes:
    root = ET.fromstring(sheet_xml_bytes)
    sheetData = root.find(f"{{{XL_NS_MAIN}}}sheetData")
    if sheetData is None:
        sheetData = ET.SubElement(root, f"{{{XL_NS_MAIN}}}sheetData")

    # Remove existing rows at/after start_row
    for row in list(sheetData):
        try:
            r = int(row.attrib.get("r", "0") or "0")
        except Exception:
            r = 0
        if r >= start_row:
            sheetData.remove(row)

    row_span = f"1:{used_cols_final}" if used_cols_final > 0 else "1:1"
    for i, row_vals in enumerate(block_2d):
        r = start_row + i
        row_el = ET.Element(f"{{{XL_NS_MAIN}}}row", r=str(r))
        row_el.set("spans", row_span)
        row_el.set("{http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac}dyDescent", "0.25")
        any_val = False
        for j in range(used_cols_final):
            v = row_vals[j] if j < len(row_vals) else ""
            if not v:
                continue
            txt = sanitize_xml_text(v)
            if txt == "":
                continue
            any_val = True
            col = _col_letter(j+1)
            c = ET.Element(f"{{{XL_NS_MAIN}}}c", r=f"{col}{r}", t="inlineStr")
            is_el = ET.SubElement(c, f"{{{XL_NS_MAIN}}}is")
            t_el = ET.SubElement(is_el, f"{{{XL_NS_MAIN}}}t")
            t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t_el.text = txt
            row_el.append(c)
        if any_val:
            sheetData.append(row_el)

    dim = root.find(f"{{{XL_NS_MAIN}}}dimension")
    if dim is None:
        dim = ET.SubElement(root, f"{{{XL_NS_MAIN}}}dimension")
        dim.set("ref", "A1:A1")
    last_row = start_row + max(0, len(block_2d) - 1)
    new_ref = _union_dimension(dim.attrib.get("ref", "A1:A1"), used_cols_final, last_row)
    dim.set("ref", new_ref)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)

def _patch_table_xml(table_xml_bytes: bytes, header_row_from_table: int, last_row: int, start_col_letter: str, width_cols: int) -> bytes:
    root = ET.fromstring(table_xml_bytes)
    last_col_letter = _col_letter(_col_number(start_col_letter) + width_cols - 1)
    new_ref = f"{start_col_letter}{header_row_from_table}:{last_col_letter}{last_row}"
    root.set("ref", new_ref)
    af = root.find(f"{{{XL_NS_MAIN}}}autoFilter")
    if af is None:
        af = ET.SubElement(root, f"{{{XL_NS_MAIN}}}autoFilter")
    af.set("ref", new_ref)
    # DO NOT touch <tableColumns>; Excel expects the count to match children.
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)

def _strip_calcchain_from_content_types(ct_bytes: bytes) -> bytes:
    try:
        root = ET.fromstring(ct_bytes)
        # Remove the calcChain override if present so missing part isn't treated as an error
        for child in list(root):
            if child.tag.endswith('Override') and child.attrib.get('PartName') == '/xl/calcChain.xml':
                root.remove(child)
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    except Exception:
        return ct_bytes

def fast_patch_template(master_bytes: bytes, sheet_name: str, header_row: int, start_row: int, used_cols: int, block_2d: list) -> bytes:
    """
    Replace Template sheet data, synchronize all table ranges, and strip calcChain.
    This prevents Excel repair prompts while preserving macros/styles.
    """
    zin = zipfile.ZipFile(io.BytesIO(master_bytes), "r")
    sheet_path = _find_sheet_part_path(zin, sheet_name)
    table_paths = _get_table_paths_for_sheet(zin, sheet_path)

    # If there are tables, clamp writing width to the SMALLEST table width
    final_cols = used_cols
    table_infos = []
    for tp in table_paths:
        try:
            sc, sr, width = _read_table_info(zin.read(tp))
            table_infos.append((tp, sc, sr, width))
            final_cols = min(final_cols, width)
        except Exception:
            pass

    original_sheet_xml = zin.read(sheet_path)
    new_sheet_xml = _patch_sheet_xml(original_sheet_xml, start_row, final_cols, block_2d)

    last_row = max(header_row, start_row + max(0, len(block_2d) - 1))

    patched_tables = {}
    for (tp, sc, sr, width) in table_infos:
        try:
            patched_tables[tp] = _patch_table_xml(zin.read(tp), sr, last_row, sc, width)
        except Exception:
            pass

    out_bio = io.BytesIO()
    with zipfile.ZipFile(out_bio, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            # Skip calcChain completely (Excel will rebuild silently)
            if item.filename == 'xl/calcChain.xml':
                continue
            data = zin.read(item.filename)
            if item.filename == sheet_path:
                data = new_sheet_xml
            elif item.filename in patched_tables:
                data = patched_tables[item.filename]
            elif item.filename == '[Content_Types].xml':
                data = _strip_calcchain_from_content_types(data)
            zout.writestr(item, data)
    zin.close()
    out_bio.seek(0)
    return out_bio.getvalue()

# ─────────────────────────────────────────────────────────────────────
# Page meta + theming
# ─────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Masterfile Automation - Amazon", page_icon="🧾", layout="wide")

st.markdown("""
<style>
:root{--bg1:#f6f9fc;--bg2:#fff;--card:#fff;--card-border:#e8eef6;--ink:#0f172a;--muted:#64748b;--accent:#2563eb}
.stApp{background:linear-gradient(180deg,var(--bg1) 0%,var(--bg2) 70%)}
.block-container{padding-top:.75rem}
.section{border:1px solid var(--card-border);background:var(--card);border-radius:16px;padding:18px 20px;box-shadow:0 6px 24px rgba(2,6,23,.05);margin-bottom:18px}
h1,h2,h3{color:var(--ink)}hr{border-color:#eef2f7}
.badge{display:inline-block;padding:4px 10px;border-radius:999px;font-size:.82rem;font-weight:600;letter-spacing:.2px;margin-right:.25rem}
.badge-info{background:#eef2ff;color:#1e40af}.badge-ok{background:#ecfdf5;color:#065f46}
.stDownloadButton>button,div.stButton>button{background:var(--accent)!important;color:#fff!important;border-radius:10px!important;border:0!important;box-shadow:0 8px 18px rgba(37,99,235,.18)}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────
# Masterfile layout
# ─────────────────────────────────────────────────────────────────────
MASTER_TEMPLATE_SHEET = "Template"
MASTER_DISPLAY_ROW    = 2
MASTER_SECONDARY_ROW  = 3
MASTER_DATA_START_ROW = 4

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
# UI
# ─────────────────────────────────────────────────────────────────────
st.title("🧾 Masterfile Automation – Amazon")
st.caption("Fills only the Template sheet and preserves all other tabs/styles/macros.")

st.markdown("<div class='section'><span class='badge badge-info'>Template-only writer</span> <span class='badge badge-ok'>Linux-fast XML patch</span></div>", unsafe_allow_html=True)

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
    mapping_json_text = st.text_area("Paste mapping JSON", height=200, placeholder='\n{\n  "Partner SKU": ["Seller SKU", "item_sku"]\n}\n')
with tab2:
    mapping_json_file = st.file_uploader("Or upload mapping.json", type=["json"], key="mapping_file")

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
        st.error("Please upload both Masterfile Template and Onboarding.")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()

    ext = (Path(masterfile_file.name).suffix or ".xlsx").lower()
    mime_map = {
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
    }
    out_mime = mime_map.get(ext, mime_map[".xlsx"])

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

    masterfile_file.seek(0)
    master_bytes = masterfile_file.read()

    slog("⏳ Reading Template headers…")
    t0 = time.time()
    wb_ro = load_workbook(io.BytesIO(master_bytes), read_only=True, data_only=True, keep_links=True)
    if MASTER_TEMPLATE_SHEET not in wb_ro.sheetnames:
        st.error(f"Sheet '{MASTER_TEMPLATE_SHEET}' not found in the masterfile.")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()
    ws_ro = wb_ro[MASTER_TEMPLATE_SHEET]
    used_cols = worksheet_used_cols(ws_ro, header_rows=(MASTER_DISPLAY_ROW, MASTER_SECONDARY_ROW), hard_cap=2048, empty_streak_stop=8)
    display_headers   = [ws_ro.cell(row=MASTER_DISPLAY_ROW,   column=c).value or "" for c in range(1, used_cols+1)]
    secondary_headers = [ws_ro.cell(row=MASTER_SECONDARY_ROW, column=c).value or "" for c in range(1, used_cols+1)]
    wb_ro.close()
    slog(f"✅ Template headers loaded (cols={used_cols}) in {time.time()-t0:.2f}s")

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

    # Build mapping: master col -> source Series (or sentinel)
    series_by_alias = {norm(h): on_df[h] for h in on_headers}
    master_to_source, report_lines = {}, []
    report_lines.append("#### 🔎 Mapping Summary (Template)")

    BULLET_DISP_N = norm("Key Product Features")

    for c, (disp, sec) in enumerate(zip(display_headers, secondary_headers), start=1):
        disp_norm = norm(disp)
        sec_norm  = norm(sec)
        if disp_norm == BULLET_DISP_N and sec_norm:
            effective_header = sec
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
                master_to_source[c] = object()  # sentinel handled below
                report_lines.append(f"- 🟨 **{label_for_log}** ← (will fill 'List')")
            else:
                report_lines.append(f"- ❌ **{label_for_log}** ← no match.")

    st.markdown("\n".join(report_lines))

    n_rows = len(on_df)
    block = [[""] * used_cols for _ in range(n_rows)]
    for col, src in master_to_source.items():
        if isinstance(src, object) and src.__class__ is object:  # sentinel
            for i in range(n_rows):
                block[i][col-1] = "List"
        else:
            vals = src.astype(str).tolist()
            m = min(len(vals), n_rows)
            for i in range(m):
                v = sanitize_xml_text(vals[i].strip())
                if v and v.lower() not in ("nan", "none", ""):
                    block[i][col-1] = v

    slog("🚀 Writing via Linux-fast XML (sheet + tables + strip calcChain)…")
    out_bytes = fast_patch_template(
        master_bytes=master_bytes,
        sheet_name=MASTER_TEMPLATE_SHEET,
        header_row=MASTER_DISPLAY_ROW,
        start_row=MASTER_DATA_START_ROW,
        used_cols=used_cols,
        block_2d=block
    )

    st.download_button(
        "⬇️ Download Final Masterfile",
        data=out_bytes,
        file_name=f"final_masterfile{Path(masterfile_file.name).suffix}",
        mime=("application/vnd.ms-excel.sheet.macroEnabled.12" if ext == ".xlsm"
              else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    )

    st.markdown("</div>", unsafe_allow_html=True)

with st.expander("📘 Notes", expanded=False):
    st.markdown(dedent(f"""
    - Writes only into `{MASTER_TEMPLATE_SHEET}`; preserves other tabs, macros & styles.
    - **All tables** on the sheet are resized without changing tableColumns.
    - The **calc chain** is removed so Excel rebuilds it silently (no warnings).
    - Illegal control characters in values are stripped automatically.
    """))
