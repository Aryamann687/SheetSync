"""SheetSync — Streamlit interface (simple column-transfer edition).

UI only. One source file and one destination workbook; matching, planning, safety checks and applying
all go through core.consolidator (a single source), and the backend is unchanged.
"""
import hashlib
import html
import io
import os

import pandas as pd
import streamlit as st

from core.reader import list_sheets, load_table
from core.consolidator import (
    STATUS_LABELS, MatchSettings, SourceEntry, SourceSpec, apply_consolidated, build_source_plan, consolidate,
    make_destination, match_source, matching_report_csv, unused_source_records, validate_setup,
)

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
.stats{display:flex; gap:.75rem; flex-wrap:wrap; margin:.4rem 0 1rem;}
.stat{background:var(--surface); border:1px solid var(--line); border-radius:var(--r); padding:.8rem 1.1rem; min-width:9.5rem;}
.stat b{display:block; font-size:1.5rem; font-weight:600; line-height:1.2;}
.stat span{color:var(--muted); font-size:.85rem;}
.note{background:var(--accent-soft); border-radius:var(--r); padding:.7rem 1rem; font-size:.9rem; color:var(--accent-ink);}
:focus-visible{outline:2px solid var(--accent) !important; outline-offset:2px;}
</style>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------- constants
# core.matcher status code -> (result label, dot colour, detail wording)
STRATEGIES = {  # shown to the user -> core.updater strategy
    "Only fill empty cells": "fill_only_blank",
    "Replace existing values": "overwrite",
}
HIDDEN = "_xl_row"  # internal column added by core.reader; never shown or mapped

SKIP_REASONS = {  # core.updater skip reason -> plain wording
    "Missing source row": "No matching record in the source file",
    "Blank source value": "The source cell is empty",
    "Merged cell is not writable": "The destination cell is part of a merged cell",
    "Destination contains a formula": "The destination cell contains a formula",
    "Destination already contains a value": "The destination cell already has a value",
}
ERRORS = [  # (text found in the backend message, instruction for the user)
    ("column headers are missing", "We couldn't identify all column names. Open \"My column names are not in row 1\" and enter the row that has a name for every column."),
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
        return "Possible match not confirmed" if code == "FUZZY" else STATUS_LABELS.get(code, (None, reason))[1]
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






def _blank(v):
    return "(empty)" if v is None or (not isinstance(v, str) and pd.isna(v)) else v


def _step(n, title, hint):
    st.markdown(f'<h3><span class="n">{n}.</span> {title}</h3>', unsafe_allow_html=True)
    st.markdown(f'<p class="hint">{hint}</p>', unsafe_allow_html=True)


def _note(text, top=False):
    style = ' style="margin-top:.8rem"' if top else ""
    st.markdown(f'<div class="note"{style}>{text}</div>', unsafe_allow_html=True)


# ---------------------------------------------------------------- helpers for the simple page
PH = "Choose a column"      # first option of every column picker
SOURCE_ID = "s1"            # the simple page uses exactly one source


def _rerun():
    (getattr(st, "rerun", None) or st.experimental_rerun)()


def _digest(data):
    return hashlib.md5(data).hexdigest()


def _tag(text):
    return f'<span class="chip">{html.escape(str(text))}</span>'


def _download(label, data, **kw):
    try:
        return st.download_button(label, data, type="primary", **kw)
    except TypeError:  # older Streamlit versions have no button types
        return st.download_button(label, data, **kw)


# Results depend only on the files and choices, so they are cached by a signature of those inputs.
@st.cache_data(show_spinner=False)
def _match(sig, _dest, _spec, _table):
    return match_source(_dest, _spec, _table)


@st.cache_data(show_spinner=False)
def _plan(sig, _dest, _entry):
    return build_source_plan(_dest, _entry)


@st.cache_data(show_spinner=False)
def _workbook(sig, _dest, _cons):
    return apply_consolidated(_dest, _cons)


@st.cache_data(show_spinner=False)
def _report(sig, _dest, _entries):
    return matching_report_csv(_dest, _entries)


@st.cache_data(show_spinner=False)
def _unused(sig, _entry):
    return unused_source_records(_entry)


def _file_inputs(key, up):
    """Sheet, column-name row and loaded table for one uploaded file; shows the columns found."""
    name, data = up.name, up.getvalue()
    try:
        sheets = _sheets(name, data)
        sheet = st.selectbox("Sheet", sheets, key=f"{key}_sheet") if len(sheets) > 1 else sheets[0]
        with st.expander("My column names are not in row 1"):
            hdr = int(st.number_input("Row that contains the column names", min_value=1, value=1, step=1, key=f"{key}_hdr"))
        table = _load(name, data, sheet, hdr)
    except Exception as e:
        msg = _friendly(e, f"We couldn't read {name}. Check that the file isn't password-protected "
                           "and that the sheet and column-name row are correct.")
        st.error(msg)
        return {"name": name, "data": data, "error": msg}
    cols = _cols(table)
    st.markdown(f"Found **{len(table):,} row{'s' if len(table) != 1 else ''}** and "
                f"**{len(cols)} column{'s' if len(cols) != 1 else ''}**:")
    st.markdown('<div class="chips">' + "".join(_tag(c) for c in cols) + "</div>", unsafe_allow_html=True)
    return {"name": name, "data": data, "sheet": sheet, "hdr": hdr, "table": table, "error": None}


def _row_picker(prefix, rid, src_cols, dst_cols, removable, default=None, auto_dest=False):
    """One row of two column pickers. Returns (source column, destination column, remove clicked)."""
    ss, ks, kd = st.session_state, f"{prefix}s_{rid}", f"{prefix}d_{rid}"
    if auto_dest:  # when the source column changes, pre-pick the destination column with the same name
        now = ss.get(ks, PH)
        if ss.get(f"{prefix}prev_{rid}") != now:
            ss[f"{prefix}prev_{rid}"] = now
            if now != PH:
                ss[kd] = now if now in dst_cols else PH
    c1, c2, c3 = st.columns([5, 5, 1], gap="medium")
    s = c1.selectbox("Column in the source file", [PH] + src_cols, key=ks, label_visibility="collapsed",
                     index=1 + src_cols.index(default) if default else 0)
    d = c2.selectbox("Column in the destination file", [PH] + dst_cols, key=kd, label_visibility="collapsed",
                     index=1 + dst_cols.index(default) if default else 0)
    gone = bool(removable and c3.button("Remove", key=f"{prefix}rm_{rid}"))
    return s, d, gone


def _forget_row(ss, kind, prefix, rid):
    ss[kind].remove(rid)
    for k in (f"{prefix}s_{rid}", f"{prefix}d_{rid}", f"{prefix}prev_{rid}"):
        ss.pop(k, None)


def _stats(matched, cells, unmatched):
    box = lambda n, label: f'<div class="stat"><b>{n:,}</b><span>{label}</span></div>'
    st.markdown('<div class="stats">' + box(matched, "Records matched") + box(cells, "Cells to update")
                + box(unmatched, "Records not matched") + "</div>", unsafe_allow_html=True)


# ---------------------------------------------------------------- header
ss = st.session_state
ss.setdefault("match_rows", [1])   # ids of the "match rows using" pickers
ss.setdefault("copy_rows", [2])    # ids of the "copy these columns" pickers
ss.setdefault("next_row", 3)
st.markdown("# SheetSync")
st.markdown('<p class="sub">Copy columns from one Excel file into another, matching rows by a column they share.</p>',
            unsafe_allow_html=True)
st.divider()

# ---------------------------------------------------------------- 1. source file
_step(1, "Upload the file you want to copy from", "Choose the Excel or CSV file that has the information you want to copy.")
src_up = st.file_uploader("Source file", type=["xlsx", "csv"], key="src_up", label_visibility="collapsed")
src = _file_inputs("src", src_up) if src_up is not None else None
if src is None or src["error"]:
    st.stop()

# ---------------------------------------------------------------- 2. destination file
st.divider()
_step(2, "Upload the workbook you want to update",
      "Choose the Excel workbook to fill in. Your original is never changed. You will download an updated copy.")
dst_up = st.file_uploader("Destination workbook", type=["xlsx"], key="dst_up", label_visibility="collapsed")
dst = _file_inputs("dst", dst_up) if dst_up is not None else None
if dst is None or dst["error"]:
    st.stop()
dest = make_destination(dst["data"], dst["sheet"], dst["hdr"], dst["table"])
src_cols, dst_cols = _cols(src["table"]), _cols(dest.table)

# ---------------------------------------------------------------- 3. select columns
st.divider()
_step(3, "Select columns", "Choose how to find the same row in both files, then which columns to copy.")
head = st.columns([5, 5, 1], gap="medium")
head[0].caption("Column in the source file")
head[1].caption("Column in the destination file")

st.markdown("**Match rows using**")
st.caption("Pick the column that identifies the same row in both files, such as District or Employee ID.")
common = next((c for c in src_cols if c in dst_cols), None)
match_raw, gone = [], None
for n, rid in enumerate(ss["match_rows"]):
    s, d, rm = _row_picker("m", rid, src_cols, dst_cols, n > 0, default=common if n == 0 else None)
    match_raw.append((s, d))
    gone = ("match_rows", "m", rid) if rm else gone
if st.button("+ Add another matching column", key="add_match"):
    ss["match_rows"].append(ss["next_row"])
    ss["next_row"] += 1
    _rerun()

st.markdown("**Copy these columns**")
st.caption("Pick a column to copy from the source file and the column to fill in the destination.")
copy_raw = []
for n, rid in enumerate(ss["copy_rows"]):
    s, d, rm = _row_picker("c", rid, src_cols, dst_cols, n > 0, auto_dest=True)
    copy_raw.append((s, d))
    gone = ("copy_rows", "c", rid) if rm else gone
if st.button("+ Add another column", key="add_copy"):
    ss["copy_rows"].append(ss["next_row"])
    ss["next_row"] += 1
    _rerun()
if gone:
    _forget_row(ss, *gone)
    _rerun()

with st.expander("More options"):
    replace = st.checkbox("Replace values that are already in the destination", key="opt_replace",
                          help="Leave this unticked to fill empty cells only. Existing values are then never touched.")
    exact = st.checkbox("Only match rows that are exactly the same", key="opt_exact",
                        help="Otherwise rows that differ only in capital letters, spaces or small spelling differences "
                             "can still be matched. Rows that only look similar are always shown for you to confirm first.")

match_pairs = [(s, d) for s, d in match_raw if s != PH and d != PH]
copy_pairs = [(s, d) for s, d in copy_raw if s != PH and d != PH]
problems = []
if any((s == PH) != (d == PH) for s, d in match_raw + copy_raw):
    problems.append("Choose a column on both sides of each row, or remove the row.")
if len({s for s, _ in copy_pairs}) != len(copy_pairs):
    problems.append("Each source column can only be copied once.")
if not match_pairs or not copy_pairs:
    if not problems:
        st.caption("Choose the matching columns and at least one column to copy to see a preview.")
    for message in problems:
        st.error(message)
    st.stop()

spec = SourceSpec(SOURCE_ID, src["name"], src["sheet"], src["hdr"], tuple(s for s, _ in match_pairs),
                  tuple(d for _, d in match_pairs), MatchSettings(90, 5, "strict" if exact else "recommended"),
                  dict(copy_pairs), "overwrite" if replace else "fill_only_blank")
for i in validate_setup([spec], dst_cols, {SOURCE_ID: src_cols}):
    if i.severity == "error":
        problems.append(i.message.split(": ", 1)[1] if i.message.startswith(spec.label + ": ") else i.message)
if problems:
    for message in dict.fromkeys(problems):
        st.error(message)
    st.stop()

# ---------------------------------------------------------------- 4. preview
st.divider()
_step(4, "Preview", "This is what will change. Nothing is saved until you download.")
msig = _sig(dest.signature, _digest(src["data"]), src["sheet"], src["hdr"], spec.src_keys, spec.dst_keys, spec.settings.mode)
try:
    with st.spinner("Looking for matching rows"):
        res = _match(msig, dest, spec, src["table"])
except Exception as e:
    st.error(_friendly(e, "We couldn't compare the two files. Check the columns you selected and try again."))
    st.stop()

summary_box = st.container()
fuzzy = res[res["status"] == "FUZZY"]
approved = frozenset()
if len(fuzzy):
    with st.expander("1 row looks similar. Check it." if len(fuzzy) == 1 else f"{len(fuzzy)} rows look similar. Check them."):
        st.caption("These rows are not exactly the same, but they look alike. Tick Same row for the ones that match. "
                   "Nothing is copied for rows you leave unticked.")
        review = pd.DataFrame({
            "Row in the destination file": [_text(dest.table, i, spec.dst_keys) for i in fuzzy["_dst_index"]],
            "Looks like this row in the source file": [_text(src["table"], i, spec.src_keys) for i in fuzzy["_suggested_src_index"]],
            "Same row": False,
        }, index=fuzzy["_dst_index"].values)
        edited = _wide(st.data_editor, review, key=f"review_{msig}",
                       disabled=["Row in the destination file", "Looks like this row in the source file"],
                       column_config={"Same row": st.column_config.CheckboxColumn(default=False)})
        approved = frozenset(edited.index[edited["Same row"]])

psig = _sig(msig, spec.strategy, tuple(sorted(spec.mapping.items())), tuple(sorted(approved, key=str)))
entry0 = SourceEntry(spec, src["table"], res, approved)
try:
    with st.spinner("Working out the changes"):
        entry = SourceEntry(spec, src["table"], res, approved, _plan(psig, dest, entry0))
except Exception as e:
    st.error(_friendly(e, "We couldn't work out the changes. Check the columns you chose and try again."))
    st.stop()
cons = consolidate(dest, [entry])

not_matched = []
for _, r in res.iterrows():
    status, d = r["status"], r["_dst_index"]
    if status in ("EXACT", "NORMALIZED") or (status == "FUZZY" and d in approved):
        continue
    not_matched.append({
        "Row in the destination file": int(dest.table.loc[d, HIDDEN]),
        "Record": _text(dest.table, d, spec.dst_keys),
        "Why": "Looks similar to a row in the source file, but not confirmed" if status == "FUZZY" else STATUS_LABELS[status][1]})
with summary_box:
    _stats(len(res) - len(not_matched), len(cons.updates), len(not_matched))
    if not cons.updates and not cons.blockers:
        _note("Nothing would change with these choices. Open the details to see why.")
    for b in cons.blockers:
        st.error(b.reason)
    if cons.conflicts:
        st.error("Some cells would receive different values. Check the columns you chose.")

with st.expander("Show details"):
    tabs = st.tabs(["Changes", "Rows not matched", "Cells left unchanged", "Unused rows in the source file"])
    with tabs[0]:
        if cons.updates:
            _wide(st.dataframe, pd.DataFrame({
                "Row": [u.row for u in cons.updates], "Column": [u.column_name for u in cons.updates],
                "Current value": [_blank(u.old_value) for u in cons.updates], "New value": [u.new_value for u in cons.updates],
            }).astype(str), hide_index=True)
        else:
            st.caption("No cells will be updated.")
    with tabs[1]:
        if not_matched:
            st.caption("Nothing is copied for these rows.")
            _wide(st.dataframe, pd.DataFrame(not_matched), hide_index=True)
        else:
            st.caption("Every row in the destination file found a match.")
    with tabs[2]:
        cells = [s for s in entry.plan.skips if "column" in s]
        if cells:
            _wide(st.dataframe, pd.DataFrame({
                "Row": [int(dest.table.loc[s["destination_index"], HIDDEN]) for s in cells],
                "Record": [_text(dest.table, s["destination_index"], spec.dst_keys) for s in cells],
                "Column": [s["column"] for s in cells],
                "Why": [_skip_reason(s["reason"]) for s in cells]}), hide_index=True)
        else:
            st.caption("Every matched cell will be updated.")
    with tabs[3]:
        unused = _unused(_sig(msig, tuple(sorted(approved, key=str))), entry)
        if len(unused):
            st.caption("These rows in the source file were not used for any row in the destination file.")
            _wide(st.dataframe, unused, hide_index=True)
        else:
            st.caption("Every row in the source file was used.")

# ---------------------------------------------------------------- 5. download
st.divider()
_step(5, "Download", "You get an updated copy of your workbook. Your original file stays as it is.")
stem = os.path.splitext(dst["name"])[0]
if cons.can_apply:
    out_sig = _sig(dest.signature, tuple((u.row, u.column, repr(u.new_value)) for u in cons.updates))
    _download("Download updated workbook", _workbook(out_sig, dest, cons), file_name=f"{stem}_updated.xlsx",
              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("Cell values and formatting are kept as far as possible and formulas are never changed, but charts, "
               "images and other embedded objects may not carry over, so check the result in Excel.")
else:
    st.caption("There is nothing to download yet. Choose columns that match and have values to copy.")
report_sig = _sig(psig, tuple(sorted(approved, key=str)))
st.download_button("Download matching report (CSV)", _report(report_sig, dest, [entry]), file_name=f"{stem}_matching_report.csv",
                   mime="text/csv", key="report")
st.caption("A list of every row in the destination file and how it was matched.")