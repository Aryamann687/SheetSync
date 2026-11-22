"""SheetSync — Streamlit interface.

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
STATUS = {  # status -> (label, dot colour)
    "EXACT": ("Exact", "#3F7F7A"),
    "NORMALIZED": ("Normalized", "#6FA39D"),
    "FUZZY": ("Fuzzy", "#C49A4A"),
    "AMBIGUOUS": ("Ambiguous", "#B77E5C"),
    "DUPLICATE_SOURCE": ("Duplicate in source", "#8C9196"),
    "DUPLICATE_DEST": ("Duplicate in destination", "#8C9196"),
    "INVALID_KEY": ("Invalid key", "#A65F5F"),
    "UNMATCHED": ("Unmatched", "#B3AEA3"),
}
STRATEGIES = {
    "Keep existing values": "fill_only_blank",
    "Replace existing values": "overwrite",
}
HIDDEN = "_xl_row"  # internal column added by core.reader; never shown or mapped


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
    """Match report: one row per destination row. `sig` and `approved` identify the inputs for caching."""
    accepted = set(approved)
    rows = []
    for _, r in _res.iterrows():
        status = r["status"]
        dst_index = r["_dst_index"]
        src_index = r["_src_index"] if status in ("EXACT", "NORMALIZED") else r["_suggested_src_index"]
        score = r["match_score"]
        if status == "FUZZY":
            decision = "Accepted" if dst_index in accepted else "Not accepted"
        else:
            decision = "Not applicable"
        rows.append({
            "Destination row": _xl_row(_dst_df, dst_index),
            "Destination key": _text(_dst_df, dst_index, dst_keys),
            "Status": STATUS[status][0],
            "Source row (matched or suggested)": _xl_row(_src_df, src_index),
            "Source key (matched or suggested)": _text(_src_df, src_index, src_keys),
            "Score": "" if pd.isna(score) else round(float(score), 1),
            "Fuzzy decision": decision,
        })
    return pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig")


def _blank(v):
    return "(blank)" if v is None or (not isinstance(v, str) and pd.isna(v)) else v


def _upload_block(label, types, key):
    """Uploader, sheet picker and header row for one file. Returns (name, bytes, sheet, header_row, df), or Nones."""
    up = st.file_uploader(label, type=types, key=f"up_{key}")
    if up is None:
        return None, None, None, None, None
    name, data = up.name, up.getvalue()
    sheet = None
    try:
        sheets = _sheets(name, data)
        sheet = st.selectbox("Sheet", sheets, key=f"sheet_{key}") if len(sheets) > 1 else sheets[0]
        header_row = st.number_input("Header row", min_value=1, value=1, step=1, key=f"hdr_{key}")
        df = _load(name, data, sheet, int(header_row))
    except Exception as e:
        st.error(f"Couldn't read {name}. {e} Check the sheet and header row.")
        return None, None, None, None, None
    st.caption(f"{len(df):,} rows, {len(df.columns)} columns")
    with st.expander("Preview"):
        _wide(st.dataframe, df[_cols(df)].head(8), hide_index=True)
    return name, data, sheet, int(header_row), df


# ---------------------------------------------------------------- header
st.markdown("# SheetSync")
st.markdown(
    '<p class="sub">Copy values from a source file into a destination workbook by matching rows.</p>',
    unsafe_allow_html=True,
)
st.divider()

# ---------------------------------------------------------------- files
st.markdown("### Files")
st.markdown('<p class="hint">Your destination workbook is never modified. You download an updated copy.</p>', unsafe_allow_html=True)
c1, c2 = st.columns(2, gap="large")
with c1:
    st.markdown("**Source**")
    src_name, src_bytes, src_sheet, src_hdr, src_df = _upload_block("Source file", ["xlsx", "csv"], "src")
with c2:
    st.markdown("**Destination**")
    dst_name, dst_bytes, dst_sheet, dst_hdr, dst_df = _upload_block("Destination workbook", ["xlsx"], "dst")

if src_df is None or dst_df is None:
    st.stop()

# ---------------------------------------------------------------- match
st.divider()
st.markdown("### Match rows")
st.markdown(
    '<p class="hint">Choose the columns that identify a row. List them in the same order on both sides.</p>',
    unsafe_allow_html=True,
)
k1, k2 = st.columns(2, gap="large")
src_keys = k1.multiselect("Source key columns", _cols(src_df), key="src_keys")
dst_keys = k2.multiselect("Destination key columns", _cols(dst_df), key="dst_keys")

with st.expander("Matching sensitivity"):
    threshold = st.slider("Fuzzy match threshold", 50, 100, 90, help="Minimum similarity score for a fuzzy suggestion.")
    margin = st.slider("Ambiguity margin", 0, 20, 5, help="How close two candidates can score before the row is marked ambiguous.")

keys_ok = bool(src_keys) and len(src_keys) == len(dst_keys)
if src_keys and dst_keys and len(src_keys) != len(dst_keys):
    st.warning("Pick the same number of key columns on both sides.")

match_sig = _sig(src_name, hashlib.md5(src_bytes).hexdigest(), src_sheet, dst_name, hashlib.md5(dst_bytes).hexdigest(), dst_sheet,
                 tuple(src_keys), tuple(dst_keys), threshold, margin,
                 src_hdr, dst_hdr)

m = st.session_state.get("match")
if m and m["sig"] != match_sig:  # files or settings changed: drop the results and any plan built from them
    st.session_state.pop("match", None)
    st.session_state.pop("plan", None)
    st.session_state["stale"], st.session_state["plan_stale"] = True, False

if st.button("Find matches", type="primary", disabled=not keys_ok):
    try:
        with st.spinner("Matching rows"):
            result_df, summary = run_match(src_df, dst_df, src_keys, dst_keys, threshold=threshold, margin=margin)
        st.session_state["match"] = {"sig": match_sig, "result": result_df, "summary": summary}
        st.session_state["stale"] = False
    except Exception as e:
        st.error(f"Matching failed. {e}")

m = st.session_state.get("match")
if not m:
    if st.session_state.get("stale"):
        st.markdown('<div class="note" style="margin-top:.8rem">Files or settings changed. Select Find matches to update the results.</div>',
                    unsafe_allow_html=True)
    st.stop()

res = m["result"]
counts = res["status"].value_counts()
chips = "".join(
    f'<span class="chip"><i style="background:{col}"></i>{label} <b>{int(counts[s])}</b></span>'
    for s, (label, col) in STATUS.items() if counts.get(s, 0)
)
st.markdown(f'<div class="chips">{chips}</div>', unsafe_allow_html=True)

with st.expander("Browse all rows"):
    chosen = st.multiselect("Show statuses", [STATUS[s][0] for s in STATUS if counts.get(s, 0)],
                            default=[STATUS[s][0] for s in STATUS if counts.get(s, 0) and s not in ("EXACT", "NORMALIZED")])
    keep = [s for s, (label, _) in STATUS.items() if label in chosen]
    view = res[res["status"].isin(keep)]
    _wide(st.dataframe, pd.DataFrame({
        "Row": [int(dst_df.loc[i, HIDDEN]) for i in view["_dst_index"]],
        "Destination": [_text(dst_df, i, dst_keys) for i in view["_dst_index"]],
        "Status": [STATUS[s][0] for s in view["status"]],
        "Score": view["match_score"].values,
    }), hide_index=True)

# ---------------------------------------------------------------- review fuzzy
fuzzy = res[res["status"] == "FUZZY"]
accepted = set()
if len(fuzzy):
    st.divider()
    st.markdown("### Review fuzzy matches")
    st.markdown('<p class="hint">These are suggestions only. Nothing is transferred for a fuzzy row unless you accept it.</p>',
                unsafe_allow_html=True)
    review = pd.DataFrame(
        {
            "Destination": [_text(dst_df, i, dst_keys) for i in fuzzy["_dst_index"]],
            "Suggested source": [_text(src_df, i, src_keys) for i in fuzzy["_suggested_src_index"]],
            "Score": fuzzy["match_score"].values,
            "Accept": False,
        },
        index=fuzzy["_dst_index"].values,
    )
    edited = _wide(
        st.data_editor, review, key=f"review_{m['sig']}", disabled=["Destination", "Suggested source", "Score"],
        column_config={"Score": st.column_config.NumberColumn(format="%.0f"),
                       "Accept": st.column_config.CheckboxColumn(default=False)},
    )
    accepted = set(edited.index[edited["Accept"]])
    st.caption(f"{len(accepted)} of {len(fuzzy)} accepted")

st.markdown('<div style="height:.5rem"></div>', unsafe_allow_html=True)
st.download_button(
    "Download match report",
    _report_csv(m["sig"], tuple(sorted(int(i) for i in accepted)), res, src_df, dst_df, src_keys, dst_keys),
    file_name=f"{os.path.splitext(dst_name)[0]}_match_report.csv",
    mime="text/csv",
)
st.markdown('<p class="hint" style="margin-top:.5rem">One row per destination row. Fuzzy suggestions you have not ticked are reported as not accepted.</p>',
            unsafe_allow_html=True)

# ---------------------------------------------------------------- update
st.divider()
st.markdown("### Transfer values")
st.markdown('<p class="hint">Formulas are never overwritten and blank source values are never copied.</p>',
            unsafe_allow_html=True)
strategy = STRATEGIES[st.radio("If a destination cell already has a value", list(STRATEGIES), horizontal=True)]

picked = st.multiselect("Source columns to transfer", _cols(src_df), key="map_src")
dst_cols = _cols(dst_df)
mapping = {}
for col in picked:
    a, b = st.columns([1, 1], gap="large")
    a.markdown(f"<div style='padding-top:.45rem'>{col}</div>", unsafe_allow_html=True)
    mapping[col] = b.selectbox(f"Destination column for {col}", dst_cols,
                               index=dst_cols.index(col) if col in dst_cols else 0,
                               key=f"map_{col}", label_visibility="collapsed")

dup_targets = len(set(mapping.values())) != len(mapping)
if dup_targets:
    st.warning("Each destination column can receive only one source column.")
plan_sig = _sig(m["sig"], strategy, tuple(sorted(mapping.items(), key=str)), tuple(sorted(accepted, key=str)))

p = st.session_state.get("plan")
if p and p["sig"] != plan_sig:  # mapping, strategy or approvals changed since the preview
    st.session_state.pop("plan", None)
    st.session_state["plan_stale"] = True

if st.button("Preview changes", type="primary", disabled=not mapping or dup_targets):
    try:
        updates, skips = build_plan(dst_bytes, dst_sheet, res, src_df, dst_df, mapping,
                                    strategy=strategy, accepted_rows=accepted)
        st.session_state["plan"] = {"sig": plan_sig, "updates": updates, "skips": skips, "out": None}
        st.session_state["plan_stale"] = False
    except Exception as e:
        st.error(f"Couldn't build the update plan. {e}")

p = st.session_state.get("plan")
if not p:
    if st.session_state.get("plan_stale"):
        st.markdown('<div class="note" style="margin-top:.8rem">Your selection changed. Select Preview changes to update the preview.</div>',
                    unsafe_allow_html=True)
    st.stop()

updates, skips = p["updates"], p["skips"]
if updates:
    st.markdown(f"**{len(updates):,} cells will change**")
    _wide(st.dataframe, pd.DataFrame({
        "Row": [_get(u, "row") for u in updates],
        "Column": [_get(u, "column_name") for u in updates],
        "Current": [_blank(_get(u, "old_value")) for u in updates],
        "New": [_get(u, "new_value") for u in updates],
    }).astype(str), hide_index=True)
else:
    st.markdown('<div class="note">Nothing to change with the current selection.</div>', unsafe_allow_html=True)

if skips:
    with st.expander(f"Skipped ({len(skips):,})"):
        skipped = pd.DataFrame({
            "Row": [int(dst_df.loc[s["destination_index"], HIDDEN]) for s in skips],
            "Key": [_text(dst_df, s["destination_index"], dst_keys) for s in skips],
            "Column": [s.get("column", "") for s in skips],
            "Reason": [s["reason"] for s in skips],
        })
        st.caption("By reason")
        _wide(st.dataframe, skipped["Reason"].value_counts().rename("Count").reset_index(), hide_index=True)
        _wide(st.dataframe, skipped, hide_index=True)

if updates:
    if st.button("Create updated workbook", type="primary"):
        try:
            with st.spinner("Writing workbook"):
                p["out"] = apply_plan(dst_bytes, dst_sheet, p["updates"])
        except Exception as e:
            st.error(f"Couldn't write the workbook. {e}")
    if p.get("out"):
        stem = os.path.splitext(dst_name)[0]
        st.markdown('<div class="note">Updated workbook is ready. Your original file was not changed.</div>', unsafe_allow_html=True)
        st.download_button("Download updated workbook", p["out"], file_name=f"{stem}_updated.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")