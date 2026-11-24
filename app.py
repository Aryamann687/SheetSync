"""SheetSync — Streamlit interface (plain-language edition).

UI only. Calls core.reader / core.matcher / core.updater as implemented; the backend is unchanged.
"""
import hashlib
import io
import os

import pandas as pd
import streamlit as st

from core.reader import list_sheets, load_table
from core.matcher import run_match
from core.updater import build_plan, apply_plan

st.set_page_config(page_title="SheetSync", layout="wide", initial_sidebar_state="collapsed")

# ---------------------------------------------------------------- styling
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Instrument+Sans:wght@400;500;600&display=swap');
:root{
  --bg:#F8F6F1; --surface:#FFFFFF; --line:#E4E0D6; --ink:#2A2C2D; --muted:#6F7371;
  --accent:#3F7F7A; --accent-ink:#2F635F; --accent-soft:#E6EFED; --r:10px;
}
html, body, .stApp, [class*="css"]{font-family:'Instrument Sans',-apple-system,'Segoe UI',sans-serif;}
.stApp{background:var(--bg); color:var(--ink);}
#MainMenu, footer, header[data-testid="stHeader"]{visibility:hidden; height:0;}
.block-container{max-width:920px; padding:3.2rem 1.5rem 5rem;}
h1{font-size:1.55rem !important; font-weight:600 !important; letter-spacing:-0.01em; margin:0 !important; padding:0 !important;}
h3{font-size:1.05rem !important; font-weight:600 !important; margin:0 0 .15rem !important; padding:0 !important;}
.sub{color:var(--muted); font-size:.95rem; margin:.35rem 0 0;}
.hint{color:var(--muted); font-size:.85rem; margin:0 0 .9rem;}
h3 .n{color:var(--accent);}
hr{border:none; border-top:1px solid var(--line); margin:2.1rem 0 1.6rem !important;}
[data-testid="stFileUploaderDropzone"]{background:var(--surface); border:1px dashed #D3CEC1; border-radius:var(--r);}
[data-baseweb="select"] > div, .stNumberInput input, .stTextInput input{
  background:var(--surface) !important; border-radius:var(--r) !important; border-color:var(--line) !important;}
[data-testid="stExpander"]{background:var(--surface); border:1px solid var(--line); border-radius:var(--r);}
[data-testid="stExpander"] summary{font-size:.92rem;}
[data-testid="stDataFrame"], [data-testid="stDataEditor"]{border:1px solid var(--line); border-radius:var(--r); overflow:hidden;}
.stButton > button, .stDownloadButton > button{
  border-radius:var(--r); border:1px solid var(--line); background:var(--surface); color:var(--ink);
  padding:.5rem 1.1rem; font-weight:500; box-shadow:0 1px 0 rgba(42,44,45,.04); transition:background .15s, border-color .15s;}
.stButton > button:hover, .stDownloadButton > button:hover{border-color:var(--accent); color:var(--accent-ink);}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"]{
  background:var(--accent); border-color:var(--accent); color:#fff;}
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover{
  background:var(--accent-ink); border-color:var(--accent-ink); color:#fff;}
.chips{display:flex; flex-wrap:wrap; gap:.5rem; margin:.4rem 0 1.2rem;}
.chip{display:inline-flex; align-items:center; gap:.45rem; background:var(--surface); border:1px solid var(--line);
  border-radius:999px; padding:.28rem .75rem; font-size:.85rem;}
.chip i{width:.55rem; height:.55rem; border-radius:50%; display:inline-block;}
.chip b{font-weight:600;}
.note{background:var(--accent-soft); border-radius:var(--r); padding:.7rem 1rem; font-size:.9rem; color:var(--accent-ink);}
:focus-visible{outline:2px solid var(--accent) !important; outline-offset:2px;}
</style>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------- constants
# core.matcher status code -> (result label, dot colour, detail wording)
STATUS = {
    "EXACT": ("Matched", "#3F7F7A", "Exact match"),
    "NORMALIZED": ("Matched", "#3F7F7A", "Match after ignoring formatting differences"),
    "FUZZY": ("Needs review", "#C49A4A", "Possible match"),
    "AMBIGUOUS": ("Unclear", "#B77E5C", "Several similar records found"),
    "DUPLICATE_SOURCE": ("Duplicate", "#8C9196", "Duplicate in source file"),
    "DUPLICATE_DEST": ("Duplicate", "#8C9196", "Duplicate in destination file"),
    "INVALID_KEY": ("Invalid information", "#A65F5F", "Identifying information is missing"),
    "UNMATCHED": ("Not found", "#B3AEA3", "No match found"),
}
RESULTS = {  # result label -> (dot colour, what it means)
    "Matched": ("#3F7F7A", "A corresponding record was found."),
    "Needs review": ("#C49A4A", "A possible match was found and needs your confirmation."),
    "Unclear": ("#B77E5C", "Several similar records were found, so none was chosen. Nothing is copied."),
    "Not found": ("#B3AEA3", "No corresponding record was found."),
    "Duplicate": ("#8C9196", "Multiple records have the same identifying information. Nothing is copied."),
    "Invalid information": ("#A65F5F", "Some identifying information is missing."),
}
STRATEGIES = {  # shown to the user -> core.updater strategy
    "Only fill empty cells": "fill_only_blank",
    "Replace existing values": "overwrite",
}
HIDDEN = "_xl_row"  # internal column added by core.reader; never shown or mapped
MODES = ["Recommended", "Strict matching"]
MATCH_MODE = {"Recommended": "recommended", "Strict matching": "strict"}  # -> run_match(mode=...)

SKIP_REASONS = {  # core.updater skip reason -> plain wording
    "Missing source row": "No matching record in the source file",
    "Blank source value": "The source cell is empty",
    "Merged cell is not writable": "The destination cell is part of a merged cell",
    "Destination contains a formula": "The destination cell contains a formula",
    "Destination already contains a value": "The destination cell already has a value",
}
ERRORS = [  # (text found in the backend message, instruction for the user)
    ("column headers are missing", "We couldn't identify all column names. Check that the row you chose under Row containing column names has a name in every column."),
    ("duplicate column headers", "Two or more columns have the same name. Give every column a unique name, then upload the file again."),
    ("reserved column name", "One column is named _xl_row, which SheetSync uses internally. Rename it, then upload the file again."),
    ("contains no data", "We couldn't find any data. Check that you chose the right sheet and that the column names are above your data."),
    ("unsupported file format", "This file type isn't supported. Upload an Excel (.xlsx) or CSV file."),
    ("worksheet not found", "We couldn't find the destination sheet. Upload the workbook again and choose the sheet."),
    ("merged cell", "A cell that needs to be filled is part of a merged cell. Unmerge it in Excel or leave that column out, then review your changes again."),
    ("cannot overwrite formula", "A cell that needs to be filled now contains a formula, and SheetSync never overwrites formulas. Review your changes again."),
    ("at least one transfer column", "Choose at least one column to copy."),
    ("equal counts", "Select the same number of columns on both sides."),
]


# ---------------------------------------------------------------- helpers
def _wide(fn, *args, **kwargs):
    """Full-width tables across Streamlit versions."""
    try:
        return fn(*args, width="stretch", **kwargs)
    except Exception:  # older Streamlit versions use use_container_width
        return fn(*args, use_container_width=True, **kwargs)


def _fh(name, data):
    f = io.BytesIO(data)
    f.name = name
    return f


def _sig(*parts):
    return hashlib.md5(repr(parts).encode()).hexdigest()


def _friendly(e, fallback):
    """Turn a backend error into an instruction; unknown errors keep their details."""
    msg = str(e).lower()
    for needle, text in ERRORS:
        if needle in msg:
            return text
    return f"{fallback} (Details: {e})"


def _skip_reason(reason):
    if reason.startswith("Match status: "):
        code = reason.split(": ", 1)[1]
        return "Possible match not confirmed" if code == "FUZZY" else STATUS.get(code, (0, 0, reason))[2]
    return SKIP_REASONS.get(reason, reason)


@st.cache_data(show_spinner=False)
def _sheets(name, data):
    return list_sheets(_fh(name, data))


@st.cache_data(show_spinner=False)
def _load(name, data, sheet, header_row):
    return load_table(_fh(name, data), sheet=sheet, header_row=header_row)


def _text(df, idx, cols):
    """Readable key text for one row; blank if the index is missing."""
    if idx is None or pd.isna(idx):
        return ""
    try:
        row = df.loc[idx, cols]
    except (KeyError, TypeError):
        try:
            row = df.loc[int(idx), cols]
        except Exception:
            return ""
    return " / ".join("" if pd.isna(v) else str(v) for v in row)


def _get(item, key, default=None):
    return item.get(key, default) if isinstance(item, dict) else getattr(item, key, default)


def _cols(df):
    return [c for c in df.columns if c != HIDDEN]


def _xl_row(df, idx):
    """Excel row number for a DataFrame index label; blank when there is none."""
    if idx is None or pd.isna(idx):
        return ""
    try:
        return int(df.loc[idx, HIDDEN])
    except (KeyError, TypeError, ValueError):
        pass
    try:
        return int(df.loc[int(idx), HIDDEN])
    except (KeyError, TypeError, ValueError):
        return ""


@st.cache_data(show_spinner=False)
def _report_csv(sig, approved, _res, _src_df, _dst_df, src_keys, dst_keys):
    """Matching report: one row per destination row. `sig` and `approved` identify the inputs for caching."""
    accepted = set(approved)
    rows = []
    for _, r in _res.iterrows():
        status = r["status"]
        dst_index = r["_dst_index"]
        src_index = r["_src_index"] if status in ("EXACT", "NORMALIZED") else r["_suggested_src_index"]
        score = r["match_score"]
        if status == "FUZZY":
            decision = "Confirmed" if dst_index in accepted else "Not confirmed"
        else:
            decision = "Not needed"
        rows.append({
            "Row in destination file": _xl_row(_dst_df, dst_index),
            "Record in destination file": _text(_dst_df, dst_index, dst_keys),
            "Result": STATUS[status][0],
            "Details": STATUS[status][2],
            "Row in source file": _xl_row(_src_df, src_index),
            "Record in source file": _text(_src_df, src_index, src_keys),
            "Similarity (0-100)": "" if pd.isna(score) else round(float(score), 1),
            "Your decision": decision,
        })
    return pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig")


def _blank(v):
    return "(empty)" if v is None or (not isinstance(v, str) and pd.isna(v)) else v


def _step(n, title, hint):
    st.markdown(f'<h3><span class="n">{n}.</span> {title}</h3>', unsafe_allow_html=True)
    st.markdown(f'<p class="hint">{hint}</p>', unsafe_allow_html=True)


def _note(text, top=False):
    style = ' style="margin-top:.8rem"' if top else ""
    st.markdown(f'<div class="note"{style}>{text}</div>', unsafe_allow_html=True)


def _upload_block(title, help_text, types, key):
    """Uploader, sheet picker and header row for one file. Returns (name, bytes, sheet, header_row, df), or Nones."""
    st.markdown(f"**{title}**")
    st.caption(help_text)
    up = st.file_uploader(title, type=types, key=f"up_{key}", label_visibility="collapsed")
    if up is None:
        return None, None, None, None, None
    name, data = up.name, up.getvalue()
    try:
        sheets = _sheets(name, data)
        sheet = st.selectbox("Sheet", sheets, key=f"sheet_{key}") if len(sheets) > 1 else sheets[0]
        header_row = st.number_input("Row containing column names", min_value=1, value=1, step=1, key=f"hdr_{key}",
                                     help="Usually row 1. Change this if your column names start further down.")
        df = _load(name, data, sheet, int(header_row))
    except Exception as e:
        st.error(_friendly(e, f"We couldn't read {name}. Check that the file isn't password-protected "
                              "and that the sheet and header row are correct."))
        return None, None, None, None, None
    st.caption(f"{len(df):,} rows and {len(_cols(df))} columns found")
    with st.expander("Preview the first rows"):
        _wide(st.dataframe, df[_cols(df)].head(8), hide_index=True)
    return name, data, sheet, int(header_row), df


# ---------------------------------------------------------------- header
st.markdown("# SheetSync")
st.markdown(
    '<p class="sub">Fill in an Excel workbook with information from another file by matching records.</p>',
    unsafe_allow_html=True,
)
st.divider()

# ---------------------------------------------------------------- 1. upload
_step(1, "Upload your files",
      "Choose the file with the information you want to copy and the Excel workbook you want to fill in. "
      "Your original files are never changed.")
c1, c2 = st.columns(2, gap="large")
with c1:
    src_name, src_bytes, src_sheet, src_hdr, src_df = _upload_block(
        "Source file", "The file that has the information to copy. Excel or CSV.", ["xlsx", "csv"], "src")
with c2:
    dst_name, dst_bytes, dst_sheet, dst_hdr, dst_df = _upload_block(
        "Destination workbook", "The Excel workbook you want to fill in.", ["xlsx"], "dst")

if src_df is None or dst_df is None:
    st.stop()

# ---------------------------------------------------------------- 2. matching
st.divider()
_step(2, "Identify matching records",
      "Choose the columns that help us identify the same record in both files. "
      "Select them in the same order on each side. For example, choose District in both files.")
k1, k2 = st.columns(2, gap="large")
src_keys = k1.multiselect("Columns that identify a record in the source file", _cols(src_df), key="src_keys")
dst_keys = k2.multiselect("Matching columns in the destination file", _cols(dst_df), key="dst_keys")

keys_ok = bool(src_keys) and len(src_keys) == len(dst_keys)
if src_keys and dst_keys and len(src_keys) != len(dst_keys):
    st.warning("Select the same number of columns on both sides.")

mode = st.radio("How should records be matched?", MODES, horizontal=True, key="mode")
strict = mode != "Recommended"
st.markdown(
    '<p class="hint">' + (
        "Only matches records whose selected information is exactly the same, including capital letters, spaces and punctuation. "
        "Nothing is suggested or guessed."
        if strict else
        "Also finds records that look slightly different, such as spelling or formatting variations. "
        "You confirm each of these before anything is copied.") + "</p>",
    unsafe_allow_html=True,
)
with st.expander("Advanced matching settings"):
    if strict:
        st.caption("These settings apply to Recommended matching only.")
    loose = st.slider("How closely should records match?", 50, 100, 90, disabled=strict,
                      help="Higher: only very similar records are suggested. Lower: more records are suggested.")
    gap = st.slider("Difference between possible matches", 0, 20, 5, disabled=strict,
                    help="If two records look almost equally likely, none is chosen for you. A higher number makes this more cautious.")
match_mode = MATCH_MODE[mode]

match_sig = _sig(src_name, hashlib.md5(src_bytes).hexdigest(), src_sheet, dst_name, hashlib.md5(dst_bytes).hexdigest(), dst_sheet,
                 tuple(src_keys), tuple(dst_keys), match_mode, None if strict else (loose, gap),
                 src_hdr, dst_hdr)

m = st.session_state.get("match")
if m and m["sig"] != match_sig:  # files or settings changed: drop the results and any plan built from them
    st.session_state.pop("match", None)
    st.session_state.pop("plan", None)
    st.session_state["stale"], st.session_state["plan_stale"] = True, False

if st.button("Find matching records", type="primary", disabled=not keys_ok):
    try:
        with st.spinner("Looking for matching records"):
            result_df, summary = run_match(src_df, dst_df, src_keys, dst_keys, threshold=loose, margin=gap, mode=match_mode)
        st.session_state["match"] = {"sig": match_sig, "result": result_df, "summary": summary}
        st.session_state["stale"] = False
    except Exception as e:
        st.error(_friendly(e, "We couldn't compare the two files. Check the columns you selected and try again."))

m = st.session_state.get("match")
if not m:
    if st.session_state.get("stale"):
        _note("Your files or settings have changed. Select Find matching records to update the results.", top=True)
    st.stop()

res = m["result"]
labels = res["status"].map(lambda s: STATUS[s][0])
counts = labels.value_counts()
st.markdown(f"**Found a match for {int(counts.get('Matched', 0)):,} of {len(res):,} records in the destination file.**")
chips = "".join(
    f'<span class="chip"><i style="background:{col}"></i>{label} <b>{int(counts[label])}</b></span>'
    for label, (col, _) in RESULTS.items() if counts.get(label, 0)
)
st.markdown(f'<div class="chips">{chips}</div>', unsafe_allow_html=True)

with st.expander("What do these results mean?"):
    st.markdown("\n".join(f"- **{label}**: {text}" for label, (_, text) in RESULTS.items()))

with st.expander("See every record"):
    present = [l for l in RESULTS if counts.get(l, 0)]
    chosen = st.multiselect("Show", present, default=[l for l in present if l != "Matched"])
    view = res[labels.isin(chosen)]
    _wide(st.dataframe, pd.DataFrame({
        "Row": [int(dst_df.loc[i, HIDDEN]) for i in view["_dst_index"]],
        "Record": [_text(dst_df, i, dst_keys) for i in view["_dst_index"]],
        "Result": [STATUS[s][0] for s in view["status"]],
        "Details": [STATUS[s][2] for s in view["status"]],
        "Similarity": view["match_score"].values,
    }), hide_index=True)

# ---------------------------------------------------------------- 3. what to copy
st.divider()
_step(3, "Choose what to copy",
      "Select the columns to copy from the source file, then choose where each one goes in the destination workbook. "
      "Cells with formulas are never changed, and empty source cells are never copied.")
picked = st.multiselect("Columns to copy from the source file", _cols(src_df), key="map_src")
dst_cols = _cols(dst_df)
mapping = {}
if picked:
    h1, h2 = st.columns(2, gap="large")
    h1.caption("Source column")
    h2.caption("Copy into this destination column")
for col in picked:
    a, b = st.columns(2, gap="large")
    a.markdown(f"<div style='padding-top:.45rem'>{col}</div>", unsafe_allow_html=True)
    mapping[col] = b.selectbox(f"Destination column for {col}", dst_cols,
                               index=dst_cols.index(col) if col in dst_cols else 0,
                               key=f"map_{col}", label_visibility="collapsed")

dup_targets = len(set(mapping.values())) != len(mapping)
if dup_targets:
    st.warning("Each destination column can receive values from only one source column.")

strategy_label = st.radio("If a cell in the destination already has a value", list(STRATEGIES), horizontal=True)
strategy = STRATEGIES[strategy_label]
if strategy == "overwrite":
    st.caption("Existing values will be replaced. You will see every change before anything is saved.")

# ---------------------------------------------------------------- 4. review
st.divider()
_step(4, "Review your changes", "Confirm any possible matches, then check exactly what will change.")
fuzzy = res[res["status"] == "FUZZY"]
accepted = set()
if len(fuzzy):
    st.markdown("**Records that need your confirmation**")
    st.caption("These records look similar but are not identical. Tick Confirm when the two records are the same. "
               "Nothing is copied for records you leave unticked.")
    review = pd.DataFrame(
        {
            "Record in destination file": [_text(dst_df, i, dst_keys) for i in fuzzy["_dst_index"]],
            "Possible match in source file": [_text(src_df, i, src_keys) for i in fuzzy["_suggested_src_index"]],
            "Similarity (0-100)": fuzzy["match_score"].values,
            "Confirm": False,
        },
        index=fuzzy["_dst_index"].values,
    )
    edited = _wide(
        st.data_editor, review, key=f"review_{m['sig']}",
        disabled=["Record in destination file", "Possible match in source file", "Similarity (0-100)"],
        column_config={"Similarity (0-100)": st.column_config.NumberColumn(format="%.0f"),
                       "Confirm": st.column_config.CheckboxColumn(default=False)},
    )
    accepted = set(edited.index[edited["Confirm"]])
    st.caption(f"{len(accepted)} of {len(fuzzy)} confirmed")
else:
    st.caption("No records need your confirmation.")

plan_sig = _sig(m["sig"], strategy, tuple(sorted(mapping.items(), key=str)), tuple(sorted(accepted, key=str)))

p = st.session_state.get("plan")
if p and p["sig"] != plan_sig:  # columns, existing-value choice or confirmations changed since the review
    st.session_state.pop("plan", None)
    st.session_state["plan_stale"] = True

if not mapping:
    st.caption("Choose at least one column to copy in step 3, then review your changes.")
if st.button("Review changes", type="primary", disabled=not mapping or dup_targets):
    try:
        updates, skips = build_plan(dst_bytes, dst_sheet, res, src_df, dst_df, mapping,
                                    strategy=strategy, accepted_rows=accepted)
        st.session_state["plan"] = {"sig": plan_sig, "updates": updates, "skips": skips, "out": None}
        st.session_state["plan_stale"] = False
    except Exception as e:
        st.error(_friendly(e, "We couldn't work out the changes. Check the columns you chose and try again."))

p = st.session_state.get("plan")
updates, skips = (p["updates"], p["skips"]) if p else ([], [])
if not p and st.session_state.get("plan_stale"):
    _note("Your selections have changed. Select Review changes to update this page.", top=True)

if p:
    if updates:
        n = len(updates)
        st.markdown(f"**{n:,} cell{'s' if n != 1 else ''} will be updated**")
        _wide(st.dataframe, pd.DataFrame({
            "Row": [_get(u, "row") for u in updates],
            "Column": [_get(u, "column_name") for u in updates],
            "Current value": [_blank(_get(u, "old_value")) for u in updates],
            "New value": [_get(u, "new_value") for u in updates],
        }).astype(str), hide_index=True)
    else:
        _note("No cells will be updated with these choices. The list below explains why.")

    if skips:
        with st.expander(f"Records and cells left unchanged ({len(skips):,})"):
            st.caption("These records or cells will not be updated, and why.")
            skipped = pd.DataFrame({
                "Row": [int(dst_df.loc[s["destination_index"], HIDDEN]) for s in skips],
                "Record": [_text(dst_df, s["destination_index"], dst_keys) for s in skips],
                "Column": [s.get("column", "") for s in skips],
                "Reason": [_skip_reason(s["reason"]) for s in skips],
            })
            _wide(st.dataframe, skipped["Reason"].value_counts().rename("Count").reset_index(), hide_index=True)
            _wide(st.dataframe, skipped, hide_index=True)

# ---------------------------------------------------------------- 5. download
st.divider()
_step(5, "Download your workbook", "Download the updated Excel file and a report showing how every record was matched.")
d1, d2 = st.columns(2, gap="large")
with d1:
    st.markdown("**Updated workbook**")
    if p and updates:
        st.caption("A copy of your destination workbook with the changes above. Your original file stays as it is.")
        if st.button("Create updated workbook", type="primary"):
            try:
                with st.spinner("Creating your workbook"):
                    p["out"] = apply_plan(dst_bytes, dst_sheet, p["updates"])
            except Exception as e:
                st.error(_friendly(e, "We couldn't create the workbook. Review your changes and try again."))
        if p.get("out"):
            _note("Your updated workbook is ready.")
            st.download_button("Download updated workbook", p["out"],
                               file_name=f"{os.path.splitext(dst_name)[0]}_updated.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    elif p:
        st.caption("There are no changes to save, so no updated workbook was created. You can still download the matching report.")
    else:
        st.caption("Review your changes in step 4 to create the updated workbook.")
with d2:
    st.markdown("**Matching report**")
    st.caption("A spreadsheet (CSV) listing every destination record, how it was matched and your decision on possible "
               "matches. Possible matches you did not confirm are marked Not confirmed.")
    st.download_button(
        "Download matching report",
        _report_csv(m["sig"], tuple(sorted(int(i) for i in accepted)), res, src_df, dst_df, src_keys, dst_keys),
        file_name=f"{os.path.splitext(dst_name)[0]}_matching_report.csv",
        mime="text/csv",
    )