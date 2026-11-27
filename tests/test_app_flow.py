"""Scenarios for the simplified page: upload source, upload destination, select columns, preview, download.

They run the real app.py top to bottom on a fake Streamlit runtime (tests/fake_streamlit.py),
the way Streamlit re-runs it after every interaction.
"""
import io
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import pandas as pd
from openpyxl import Workbook, load_workbook

import core.consolidator as cons
from core.matcher import run_match
from core.reader import load_table
from fake_streamlit import AppDriver

APP = ROOT / "app.py"
PH = "Choose a column"
COPY = 2                    # id of the first "copy these columns" picker (the first "match rows using" picker is 1)


# ---------------------------------------------------------------- workbook helpers
def book(headers, rows, formulas=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(list(headers))
    for r in rows:
        ws.append(r)
    for ref, f in (formulas or {}).items():
        ws[ref] = f
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def load(data, name="f.xlsx"):
    f = io.BytesIO(data)
    f.name = name
    return load_table(f, sheet="Data")


def cells(data):
    ws = load_workbook(io.BytesIO(data))["Data"]
    return {(c.row, c.column): c.value for row in ws.iter_rows() for c in row}


# the example from the brief
SRC_HEAD = ["District", "Population", "Literacy"]
SRC_ROWS = [["Kolar", 650000, 0.82], ["Mysuru", 350000, 0.85]]
DST_HEAD = ["District", "Total Population", "Literacy Rate"]
DST_ROWS = [["Kolar", None, None], ["Mysuru", None, None], ["Hassan", None, None]]


# ---------------------------------------------------------------- driving the page
def start(src_headers=SRC_HEAD, src_rows=SRC_ROWS, dest=None, src_name="Source.xlsx"):
    app = AppDriver(APP).run()
    data = src_headers if isinstance(src_headers, bytes) else book(src_headers, src_rows)
    app.upload("src_up", src_name, data).run()
    return app.upload("dst_up", "Master.xlsx", dest or book(DST_HEAD, DST_ROWS)).run()


def pick(app, src, dst=None, rid=COPY):
    app.set(f"cs_{rid}", src).run()
    if dst is not None:
        app.set(f"cd_{rid}", dst).run()
    return app


def stats(app):
    found = re.findall(r'<div class="stat"><b>([\d,]+)</b><span>([^<]+)</span>', app.text())
    return {label: int(n) for n, label in found}


def workbook(app):
    return cells(app.st.downloads["Download updated workbook"]["data"])


def reasons(app, columns=("Row in the destination file", "Record", "Why")):
    t = app.table(*columns)
    return dict(zip(t["Record"], t["Why"]))


def cell_reasons(app):
    return set(app.table("Row", "Record", "Column", "Why")["Why"])


def rich_run():
    """A run that shows every part of the page, including rows that only look similar."""
    dest = book(["District", "A"], [["Kolar", None], ["Chikkamagaluru", None], ["Tumkur", 9]])
    app = start(["District", "A"], [["Kolar", 1], ["Chikkamagalur", 2], ["Tumkur", 3]], dest)
    return pick(app, "A", "A")


# ---------------------------------------------------------------- the five steps
def test_the_page_reveals_one_step_at_a_time():
    app = AppDriver(APP).run()
    assert app.has_text("Upload the file you want to copy from") and not app.has_text("Upload the workbook you want to update")
    app.upload("src_up", "Source.xlsx", book(SRC_HEAD, SRC_ROWS)).run()
    assert app.has_text("Upload the workbook you want to update") and not app.has_text("Select columns")
    app.upload("dst_up", "Master.xlsx", book(DST_HEAD, DST_ROWS)).run()
    assert app.has_text("Select columns") and not app.has_text("Preview")
    pick(app, "Population", "Total Population")
    assert app.has_text("Preview") and app.has_text("Download")


def test_each_file_shows_the_columns_it_contains():
    app = start()
    for column in SRC_HEAD + DST_HEAD:
        assert f'<span class="chip">{column}</span>' in app.text()
    assert app.has_text("Found **2 rows** and **3 columns**") and app.has_text("Found **3 rows** and **3 columns**")


def test_the_brief_example_works_with_two_copied_columns():
    app = start()
    assert (app.state["ms_1"], app.state["md_1"]) == ("District", "District")           # same-named columns are pre-picked
    pick(app, "Population", "Total Population")
    app.click("add_copy").run()
    assert app.state["copy_rows"] == [2, 3]
    pick(app, "Literacy", None, rid=3)
    assert app.state["cd_3"] == PH                                                      # no same-named column to pre-pick
    pick(app, "Literacy", "Literacy Rate", rid=3)
    assert stats(app) == {"Records matched": 2, "Cells to update": 4, "Records not matched": 1}
    out = workbook(app)
    assert [out[(2, c)] for c in (1, 2, 3)] == ["Kolar", 650000, 0.82]
    assert [out[(3, c)] for c in (1, 2, 3)] == ["Mysuru", 350000, 0.85]
    assert out[(4, 2)] is None and out[(4, 3)] is None                                  # Hassan is not in the source
    assert app.st.downloads["Download updated workbook"]["file_name"] == "Master_updated.xlsx"


def test_the_updated_workbook_is_built_once_per_result_and_the_original_is_untouched(monkeypatch):
    calls, real = [], cons.apply_plan
    monkeypatch.setattr(cons, "apply_plan", lambda *a: calls.append(a) or real(*a))
    dest = book(DST_HEAD, DST_ROWS)
    app = start(dest=dest)
    pick(app, "Population", "Total Population")
    for _ in range(3):
        app.run()                                                                       # more interactions, same choices
    assert len(calls) == 1
    assert app.state["dst_up"].getvalue() == dest


def test_the_matching_report_is_the_existing_single_file_report():
    app = start()
    pick(app, "Population", "Total Population")
    report = app.st.downloads["Download matching report (CSV)"]["data"]
    src, dst_bytes = load(book(SRC_HEAD, SRC_ROWS)), book(DST_HEAD, DST_ROWS)
    dest = cons.make_destination(dst_bytes, "Data", 1, load(dst_bytes))
    res, _ = run_match(src, dest.table, ["District"], ["District"], 90, 5)
    spec = cons.SourceSpec("s1", "x", "Data", 1, ("District",), ("District",), cons.MatchSettings(90, 5), {})
    assert report == cons.matching_report_csv(dest, [cons.SourceEntry(spec, src, res)]) and b"Source file" not in report
    assert pd.read_csv(io.BytesIO(report))["Result"].tolist() == ["Matched", "Matched", "Not found"]


# ---------------------------------------------------------------- choosing columns
def test_nothing_is_previewed_until_the_choices_are_complete():
    app = start()
    assert stats(app) == {} and "Download updated workbook" not in app.st.downloads
    app.set(f"cs_{COPY}", "Population").run()                                           # destination column still missing
    assert app.has_text("Choose a column on both sides") and stats(app) == {}
    assert "Download updated workbook" not in app.st.downloads


def test_a_destination_column_used_for_matching_cannot_be_filled():
    app = start()
    pick(app, "Population", "District")
    assert app.has_text("used to identify records") and stats(app) == {} and not app.st.downloads


def test_a_source_column_can_only_be_copied_once_and_a_destination_column_only_filled_once():
    app = start()
    pick(app, "Population", "Total Population")
    app.click("add_copy").run()
    pick(app, "Population", "Literacy Rate", rid=3)
    assert app.has_text("Each source column can only be copied once") and stats(app) == {}
    pick(app, "Literacy", "Total Population", rid=3)
    assert app.has_text("more than one column is being copied into Total Population") and stats(app) == {}


def test_rows_can_be_added_and_removed_without_leaving_state_behind():
    app = start()
    app.click("add_copy").run().click("add_match").run()
    assert app.state["copy_rows"] == [2, 3] and app.state["match_rows"] == [1, 4]
    pick(app, "Literacy", "Literacy Rate", rid=3)
    app.click("crm_3").run().click("mrm_4").run()
    assert app.state["copy_rows"] == [2] and app.state["match_rows"] == [1]
    assert not any(k.endswith(("_3", "_4")) for k in app.state)


def test_several_matching_columns_must_all_agree():
    head = ["First", "Last", "Score"]
    app = start(head, [["Ann", "Lee", 1], ["Ann", "Kim", 2]], book(head, [["Ann", "Kim", None]]))
    app.click("add_match").run()
    app.set("ms_3", "Last").set("md_3", "Last").run()
    pick(app, "Score")
    assert app.state["cd_2"] == "Score"
    assert stats(app) == {"Records matched": 1, "Cells to update": 1, "Records not matched": 0}
    assert workbook(app)[(2, 3)] == 2                                                   # Ann Kim, not Ann Lee


def test_a_header_row_other_than_one_can_be_chosen():
    data = book(["Population report"], [[], ["District", "Population"], ["Kolar", 650000]])
    app = AppDriver(APP).run().upload("src_up", "Report.xlsx", data).run()
    assert app.has_text("My column names are not in row 1") and app.has_text("We couldn't identify all column names")
    app.set("src_hdr", 3).run()
    assert app.has_text("Found **1 row** and **2 columns**") and app.has_text("Upload the workbook you want to update")


def test_csv_files_can_be_the_source():
    csv = "District,Population\nKolar,650000\nMysuru,350000\n".encode()
    app = start(csv, src_name="data.csv")
    pick(app, "Population", "Total Population")
    assert stats(app) == {"Records matched": 2, "Cells to update": 2, "Records not matched": 1}


# ---------------------------------------------------------------- safety is unchanged
def test_only_empty_cells_are_filled_unless_replacing_is_chosen():
    app = start(dest=book(DST_HEAD, [["Kolar", 111, None], ["Mysuru", None, None]]))
    pick(app, "Population", "Total Population")
    assert stats(app)["Cells to update"] == 1 and workbook(app)[(2, 2)] == 111
    assert "The destination cell already has a value" in cell_reasons(app)
    app.set("opt_replace", True).run()
    assert stats(app)["Cells to update"] == 2 and workbook(app)[(2, 2)] == 650000


def test_formulas_are_never_overwritten_even_when_replacing():
    app = start(dest=book(DST_HEAD, DST_ROWS, {"B2": "=1+1"}))
    app.set("opt_replace", True).run()
    pick(app, "Population", "Total Population")
    assert workbook(app)[(2, 2)] == "=1+1" and "The destination cell contains a formula" in cell_reasons(app)
    assert workbook(app)[(3, 2)] == 350000


def test_duplicate_rows_in_the_source_are_not_copied_and_are_explained():
    app = start(SRC_HEAD, [["Kolar", 1, 0.1], ["Kolar", 2, 0.2], ["Mysuru", 350000, 0.85]])
    pick(app, "Population", "Total Population")
    assert stats(app) == {"Records matched": 1, "Cells to update": 1, "Records not matched": 2}
    assert reasons(app) == {"Kolar": "Duplicate in source file", "Hassan": "No match found"}
    assert workbook(app)[(2, 2)] is None and workbook(app)[(3, 2)] == 350000


def test_rows_that_only_look_similar_are_copied_only_after_confirmation():
    app = rich_run()
    assert stats(app) == {"Records matched": 2, "Cells to update": 1, "Records not matched": 1}   # Kolar; Tumkur already has a value
    key = next(k for k in app.st.editors if k.startswith("review_"))
    assert any("similar" in e for e in app.st.expanders) and app.state.get(key, {}) == {}
    app.edit(key, {1: True}).run()
    assert stats(app) == {"Records matched": 3, "Cells to update": 2, "Records not matched": 0}
    assert workbook(app)[(3, 2)] == 2
    report = pd.read_csv(io.BytesIO(app.st.downloads["Download matching report (CSV)"]["data"]))
    assert report.loc[report["Record in destination file"] == "Chikkamagaluru", "Your decision"].tolist() == ["Confirmed"]


def test_unconfirmed_similar_rows_are_reported_as_not_confirmed_and_copy_nothing():
    app = rich_run()
    report = pd.read_csv(io.BytesIO(app.st.downloads["Download matching report (CSV)"]["data"]))
    assert report.loc[report["Record in destination file"] == "Chikkamagaluru", "Your decision"].tolist() == ["Not confirmed"]
    assert workbook(app).get((3, 2)) is None


def test_matching_exactly_is_a_simple_option():
    ids = [f"EMP00{i}" for i in range(1, 7)]
    app = start(["Employee ID", "Salary"], [[i, n * 10] for n, i in enumerate(ids, 1) if i != "EMP005"],
                book(["Employee ID", "Salary"], [[i, None] for i in ids]))
    pick(app, "Salary", "Salary")
    assert reasons(app) == {"EMP005": "Several similar records found"}
    app.set("opt_exact", True).run()
    assert reasons(app) == {"EMP005": "No match found"} and stats(app)["Records matched"] == 5


def test_nothing_to_download_when_nothing_would_change():
    app = start(["District", "Population"], [["Nowhere", 1]])
    pick(app, "Population", "Total Population")
    assert stats(app) == {"Records matched": 0, "Cells to update": 0, "Records not matched": 3}
    assert "Download updated workbook" not in app.st.downloads and "Download matching report (CSV)" in app.st.downloads
    assert app.has_text("Nothing would change")


# ---------------------------------------------------------------- stale choices
def test_replacing_the_source_file_discards_choices_that_no_longer_apply():
    app = start()
    pick(app, "Population", "Total Population")
    assert "Download updated workbook" in app.st.downloads
    app.upload("src_up", "Other.xlsx", book(["District", "Households"], [["Kolar", 5]])).run()
    assert app.state["cs_2"] == PH and stats(app) == {} and "Download updated workbook" not in app.st.downloads
    pick(app, "Households", "Total Population")
    assert workbook(app)[(2, 2)] == 5


def test_replacing_the_destination_recalculates_the_preview():
    app = start()
    pick(app, "Population", "Total Population")
    assert stats(app)["Records not matched"] == 1
    app.upload("dst_up", "Bigger.xlsx", book(DST_HEAD, DST_ROWS + [["Kolkata", None, None], ["Tumkur", None, None]])).run()
    assert stats(app) == {"Records matched": 2, "Cells to update": 2, "Records not matched": 3}


def test_changing_columns_updates_the_download_and_never_reuses_an_old_one():
    app = start()
    pick(app, "Population", "Total Population")
    first = workbook(app)
    pick(app, "Literacy", "Literacy Rate")
    second = workbook(app)
    assert first[(2, 2)] == 650000 and first[(2, 3)] is None
    assert second[(2, 3)] == 0.82 and second[(2, 2)] is None


# ---------------------------------------------------------------- the page stays simple
FORBIDDEN = r"\b(fuzzy|thresholds?|margins?|blockers?|consolidat\w*|plans?|ambiguous|normali[sz]ed|status(es)?)\b"


def everything_the_user_can_read(app):
    texts = [t for _, t in app.st.texts if "<style>" not in t]
    tables = [" ".join(map(str, list(t.columns) + t.to_numpy().ravel().tolist())) for t in app.st.tables]
    return "\n".join(texts + app.st.expanders + tables)


def test_no_technical_terms_appear_anywhere_on_the_page():
    exact = rich_run().set("opt_exact", True).run()
    bad = pick(start(), "Population", "District")
    for page in (rich_run(), start(), AppDriver(APP).run(), exact, bad):
        found = re.search(FORBIDDEN, everything_the_user_can_read(page), re.I)
        assert not found, found.group(0)


def test_details_and_options_stay_collapsed_until_the_user_opens_them():
    app = rich_run()
    for label in ("Show details", "More options", "My column names are not in row 1"):
        assert app.st.expanded[label] is False
    assert not any(app.st.expanded[e] for e in app.st.expanded if "similar" in e)


def test_there_are_no_source_cards_and_only_one_primary_download():
    app = rich_run()
    assert not any(e.startswith("Source file") for e in app.st.expanders)
    assert [b for b in app.st.downloads if "workbook" in b] == ["Download updated workbook"]