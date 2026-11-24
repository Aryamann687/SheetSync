"""Regression tests for Bug Report 01: Strict matching must never produce possible matches.

Strict mode means exact matches only: no normalized pass, no fuzzy pass, no similarity scores.
Recommended mode (the default) keeps the existing exact, normalized and fuzzy behaviour.
"""
import ast
import io
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import pandas as pd
import pytest
from openpyxl import Workbook

from core.consolidator import (
    MatchSettings, SourceEntry, SourceSpec, consolidate, make_destination, matching_report_csv,
)
from core.matcher import MATCH_MODES, run_match
from core.reader import load_table
from core.updater import build_plan

KEY = ["Key"]


def table(*values):
    return pd.DataFrame({"Key": list(values)})


def match(src, dst, **kw):
    return run_match(src, dst, KEY, KEY, **kw)[0]


def statuses(src, dst, **kw):
    return list(match(src, dst, **kw)["status"])


def assert_no_suggestions(res):
    assert res["match_score"].isna().all() and res["_suggested_src_index"].isna().all()


# One dataset that exercises every status the matcher can produce.
MIXED_DST = table("EMP001", " emp002 ", "Chikkamagaluru", "Mandya", "DUPKEY", None)
MIXED_SRC = table("EMP001", "EMP002", "Chikkamagalur", "Tumkur", "DUPKEY", "DUPKEY")

# The Bug Report 01 data: EMP005 exists only in the destination; every other ID is one character away.
HR_DST = table(*[f"EMP00{i}" for i in range(1, 7)])
HR_SRC = table("EMP001", "EMP002", "EMP003", "EMP004", "EMP006")


# ---------------------------------------------------------------- API and defaults
def test_modes_and_invalid_mode():
    assert MATCH_MODES == ("recommended", "strict")
    for bad in ("loose", "Strict", "", None):
        with pytest.raises(ValueError, match="Invalid matching mode"):
            run_match(HR_SRC, HR_DST, KEY, KEY, mode=bad)


def test_default_mode_is_recommended_and_behaviour_is_unchanged():
    default, summary = run_match(MIXED_SRC, MIXED_DST, KEY, KEY)
    explicit, _ = run_match(MIXED_SRC, MIXED_DST, KEY, KEY, mode="recommended")
    pd.testing.assert_frame_equal(default, explicit)
    assert list(default["status"]) == ["EXACT", "NORMALIZED", "FUZZY", "UNMATCHED", "DUPLICATE_SOURCE", "INVALID_KEY"]
    assert default.loc[2, "_suggested_src_index"] == 2 and round(default.loc[2, "match_score"], 1) == 96.3
    assert summary["FUZZY"] == 1


def test_positional_calls_from_before_the_change_still_work():
    res, _ = run_match(HR_SRC, HR_DST, KEY, KEY, 90, 5)
    assert res.loc[4, "status"] == "AMBIGUOUS"


# ---------------------------------------------------------------- Strict: what it accepts
def test_strict_accepts_exact_matches():
    res = match(table("EMP002", "EMP001"), table("EMP001", "EMP002"), mode="strict")
    assert list(res["status"]) == ["EXACT", "EXACT"] and list(res["_src_index"]) == [1, 0]
    assert_no_suggestions(res)


@pytest.mark.parametrize("dst_value, src_value", [
    ("EMP005", "emp005"), ("EMP005", " EMP005 "), ("EMP-005", "EMP 005"), ("Kolar", "KOLAR"),
])
def test_strict_rejects_matches_that_differ_only_in_formatting(dst_value, src_value):
    res = match(table(src_value), table(dst_value), mode="strict")
    assert list(res["status"]) == ["UNMATCHED"]
    assert_no_suggestions(res)
    # the same pair is an ordinary "match after ignoring formatting" in Recommended mode
    assert statuses(table(src_value), table(dst_value)) == ["NORMALIZED"]


# ---------------------------------------------------------------- Strict: what it must not produce
def test_strict_reports_the_hr_record_as_not_found_without_any_suggestion():
    res = match(HR_SRC, HR_DST, mode="strict")
    assert list(res["status"]) == ["EXACT"] * 4 + ["UNMATCHED", "EXACT"]
    assert res.loc[4, "status"] == "UNMATCHED"      # EMP005
    assert_no_suggestions(res)


def test_recommended_still_reports_the_hr_record_as_unclear_with_a_score():
    res = match(HR_SRC, HR_DST, mode="recommended")
    assert res.loc[4, "status"] == "AMBIGUOUS" and round(res.loc[4, "match_score"], 1) == 83.3


@pytest.mark.parametrize("threshold, margin", [(90, 5), (80, 5), (100, 5), (101, 5), (50, 0), (0, 20)])
def test_strict_ignores_threshold_and_margin_and_never_matches_similar_ids(threshold, margin):
    res = match(table("EMP006"), table("EMP005"), threshold=threshold, margin=margin, mode="strict")
    assert list(res["status"]) == ["UNMATCHED"]
    assert_no_suggestions(res)


def test_similar_ids_are_a_possible_match_only_in_recommended_mode():
    src, dst = table("EMP006"), table("EMP005")
    assert statuses(src, dst, threshold=80) == ["FUZZY"]
    assert statuses(src, dst, threshold=80, mode="strict") == ["UNMATCHED"]
    assert statuses(src, dst) == ["UNMATCHED"]      # default threshold 90: 83.3 is below it


def test_strict_never_produces_fuzzy_or_ambiguous_statuses_or_scores():
    res = match(MIXED_SRC, MIXED_DST, mode="strict")
    assert list(res["status"]) == ["EXACT", "UNMATCHED", "UNMATCHED", "UNMATCHED", "DUPLICATE_SOURCE", "INVALID_KEY"]
    assert_no_suggestions(res)
    assert not {"FUZZY", "AMBIGUOUS", "NORMALIZED"} & set(res["status"])


def test_strict_rows_it_does_decide_agree_with_recommended():
    strict, rec = match(MIXED_SRC, MIXED_DST, mode="strict"), match(MIXED_SRC, MIXED_DST)
    decided = strict["status"].isin(["EXACT", "DUPLICATE_SOURCE", "DUPLICATE_DEST", "INVALID_KEY"])
    assert decided.sum() == 3
    pd.testing.assert_series_equal(strict.loc[decided, "status"], rec.loc[decided, "status"])
    pd.testing.assert_series_equal(strict.loc[decided, "_src_index"], rec.loc[decided, "_src_index"])


# ---------------------------------------------------------------- duplicates and invalid keys
def test_duplicate_source_keys_under_strict():
    src, dst = table("EMP001", "EMP001", "EMP002"), table("EMP001", "EMP002")
    assert statuses(src, dst, mode="strict") == ["DUPLICATE_SOURCE", "EXACT"] == statuses(src, dst)


def test_duplicate_destination_keys_under_strict():
    src, dst = table("EMP001"), table("EMP001", "EMP001")
    assert statuses(src, dst, mode="strict") == ["EXACT", "DUPLICATE_DEST"] == statuses(src, dst)


def test_strict_does_not_treat_formatting_variants_as_duplicates():
    src, dst = table("Kolar", "kolar"), table("KOLAR")
    assert statuses(src, dst, mode="strict") == ["UNMATCHED"]
    assert statuses(src, dst) == ["DUPLICATE_SOURCE"]    # Recommended compares ignoring formatting


@pytest.mark.parametrize("mode", ["strict", "recommended"])
def test_blank_keys_are_invalid_in_the_destination_and_ignored_in_the_source(mode):
    res = match(table("", None, "EMP001"), table(None, "", "   ", "EMP001"), mode=mode)
    assert list(res["status"]) == ["INVALID_KEY"] * 3 + ["EXACT"]
    assert res.loc[3, "_src_index"] == 2


# ---------------------------------------------------------------- approval workflow, end to end
def workbook(headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(headers)
    for r in rows:
        ws.append(r)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def load(data, name="f.xlsx"):
    f = io.BytesIO(data)
    f.name = name
    return load_table(f, sheet="Data")


FUZZY_DST = workbook(["Key", "Value"], [["Chikkamagaluru", None], ["Mandya", None]])
FUZZY_SRC = workbook(["Key", "Value"], [["Chikkamagalur", 5], ["Tumkur", 6]])


def plan_for(mode, approved):
    dst, src = load(FUZZY_DST, "d.xlsx"), load(FUZZY_SRC, "s.xlsx")
    res, _ = run_match(src, dst, KEY, KEY, mode=mode)
    plan, skips = build_plan(FUZZY_DST, "Data", res, src, dst, {"Value": "Value"}, accepted_rows=approved)
    return res, plan, skips


def test_recommended_possible_matches_still_need_explicit_approval():
    res, plan, skips = plan_for("recommended", set())
    assert res.loc[0, "status"] == "FUZZY"
    assert plan == [] and {"destination_index": 0, "reason": "Match status: FUZZY"} in skips
    _, plan, _ = plan_for("recommended", {0})
    assert [(u["row"], u["new_value"]) for u in plan] == [(2, 5)]


def test_strict_results_cannot_be_turned_into_matches_by_an_approval():
    res, plan, skips = plan_for("strict", {0, 1})
    assert list(res["status"]) == ["UNMATCHED", "UNMATCHED"] and plan == []
    assert all(s["reason"] == "Match status: UNMATCHED" for s in skips)


# ---------------------------------------------------------------- reports reflect the real statuses
def test_matching_report_and_consolidation_show_strict_results_correctly():
    dst_bytes = workbook(["Key", "Value"], [[f"EMP00{i}", None] for i in range(1, 7)])
    src_bytes = workbook(["Key", "Value"], [[f"EMP00{i}", i * 10] for i in (1, 2, 3, 4, 6)])
    dest = make_destination(dst_bytes, "Data", 1, load(dst_bytes, "d.xlsx"))
    src = load(src_bytes, "s.xlsx")
    res, _ = run_match(src, dest.table, KEY, KEY, mode="strict")
    spec = SourceSpec("s", "Employees.xlsx", "Data", 1, KEY, KEY, MatchSettings(90, 5), {"Value": "Value"})
    plan, skips = build_plan(dest.file_bytes, "Data", res, src, dest.table, {"Value": "Value"})
    from core.consolidator import SourcePlan
    entry = SourceEntry(spec, src, res, frozenset(), SourcePlan("s", dest.signature, tuple(plan), tuple(skips)))

    report = pd.read_csv(io.BytesIO(matching_report_csv(dest, [entry])))
    hr = report[report["Record in destination file"] == "EMP005"].iloc[0]
    assert hr["Result"] == "Not found" and hr["Details"] == "No match found"
    assert pd.isna(hr["Similarity (0-100)"]) and pd.isna(hr["Record in source file"])
    assert hr["Your decision"] == "Not needed"
    assert not {"Needs review", "Unclear"} & set(report["Result"])

    c = consolidate(dest, [entry])
    assert c.can_apply and len(c.updates) == 5 and c.summaries[0].needs_review == 0


# ---------------------------------------------------------------- the UI passes the mode explicitly
def test_every_run_match_call_in_app_py_passes_the_matching_mode():
    app = ROOT / "app.py"
    if not app.exists():
        return
    tree = ast.parse(app.read_text())
    calls = [n for n in ast.walk(tree) if isinstance(n, ast.Call)
             and getattr(n.func, "id", getattr(n.func, "attr", None)) == "run_match"]
    assert calls, "app.py should call run_match"
    assert all("mode" in {k.arg for k in c.keywords} for c in calls)
    assert "STRICT" not in {n.id for n in ast.walk(tree) if isinstance(n, ast.Name)}