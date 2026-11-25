"""Matching-mode propagation in core/consolidator.py.

Each source carries its own MatchSettings (mode, threshold, margin). `match_source()` runs the
matcher with those settings; `consolidate()` refuses a strict source whose results contain
anything a strict match cannot produce.
"""
import io
import sys
from dataclasses import replace
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd
import pytest
from openpyxl import Workbook

import core.consolidator as cons
from core.consolidator import (
    MatchSettings, SourceEntry, SourceSpec, build_source_plan, consolidate, make_destination,
    match_source, matching_report_csv,
)
from core.matcher import MATCH_MODES, run_match
from core.reader import load_table
from core.updater import build_plan

KEY = ["Employee ID"]


# ---------------------------------------------------------------- helpers
def book(rows, headers=("Employee ID", "Salary")):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(list(headers))
    for r in rows:
        ws.append(r)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def load(data, name="f.xlsx"):
    f = io.BytesIO(data)
    f.name = name
    return load_table(f, sheet="Data")


def destination(ids):
    data = book([[i, None] for i in ids])
    return make_destination(data, "Data", 1, load(data, "dest.xlsx"))


def source_table(rows):
    return load(book(rows), "src.xlsx")


def spec(sid, label, mode=None, settings=None):
    settings = settings or (MatchSettings(90, 5) if mode is None else MatchSettings(90, 5, mode))
    return SourceSpec(sid, label, "Data", 1, KEY, KEY, settings, {"Salary": "Salary"})


def entry(dest, sid, label, table, settings=None):
    s = spec(sid, label, settings=settings)
    e = SourceEntry(s, table, match_source(dest, s, table))
    return replace(e, plan=build_source_plan(dest, e))


def fake_result(*rows):
    """One (status, src_index, suggested_index, score) per destination row, in the matcher's format."""
    return pd.DataFrame([
        {"_dst_index": i, "_src_index": s, "_suggested_src_index": g, "match_score": sc, "status": st}
        for i, (st, s, g, sc) in enumerate(rows)])


# The Bug Report 01 data: EMP005 exists only in the destination; every other ID is one character away.
HR_IDS = [f"EMP00{i}" for i in range(1, 7)]
HR_SRC = source_table([["EMP001", 10], ["EMP002", 20], ["EMP003", 30], ["EMP004", 40], ["EMP006", 60]])


# ---------------------------------------------------------------- MatchSettings
def test_default_mode_is_recommended():
    assert MatchSettings(90, 5).mode == "recommended"
    assert MatchSettings(90, 5) == MatchSettings(90, 5, "recommended") == MatchSettings(threshold=90, margin=5)
    assert spec("a", "A.xlsx").settings.mode == "recommended"


@pytest.mark.parametrize("mode", list(MATCH_MODES))
def test_every_mode_the_matcher_supports_is_accepted(mode):
    assert MatchSettings(90, 5, mode).mode == mode


@pytest.mark.parametrize("bad", ["loose", "Strict", "RECOMMENDED", "", None, 1])
def test_invalid_modes_are_rejected(bad):
    with pytest.raises(ValueError, match="Invalid matching mode"):
        MatchSettings(90, 5, bad)


def test_the_error_names_the_supported_modes_and_replace_revalidates():
    with pytest.raises(ValueError) as exc:
        MatchSettings(90, 5, "loose")
    assert all(mode in str(exc.value) for mode in MATCH_MODES)
    with pytest.raises(ValueError, match="Invalid matching mode"):
        replace(MatchSettings(90, 5), mode="loose")
    assert replace(MatchSettings(90, 5), mode="strict").mode == "strict"


# ---------------------------------------------------------------- propagation to the matcher
@pytest.mark.parametrize("settings, expected_mode", [
    (MatchSettings(90, 5), "recommended"),
    (MatchSettings(90, 5, "recommended"), "recommended"),
    (MatchSettings(80, 2, "strict"), "strict"),
])
def test_match_source_passes_the_sources_own_settings_to_the_matcher(monkeypatch, settings, expected_mode):
    dest, calls = destination(HR_IDS), []
    monkeypatch.setattr(cons, "run_match", lambda *a, **k: calls.append((a, k)) or (pd.DataFrame(), {}))
    match_source(dest, spec("a", "A.xlsx", settings=settings), HR_SRC)
    (args, kwargs), = calls
    assert args[0] is HR_SRC and args[1] is dest.table and args[2] == KEY and args[3] == KEY
    assert kwargs == {"threshold": settings.threshold, "margin": settings.margin, "mode": expected_mode}


def test_strict_and_recommended_sources_behave_differently_on_the_same_data():
    dest = destination(HR_IDS)
    strict = match_source(dest, spec("s", "S.xlsx", "strict"), HR_SRC)
    rec = match_source(dest, spec("r", "R.xlsx", "recommended"), HR_SRC)
    assert strict.loc[4, "status"] == "UNMATCHED"       # EMP005
    assert strict["match_score"].isna().all() and strict["_suggested_src_index"].isna().all()
    assert rec.loc[4, "status"] == "AMBIGUOUS" and round(rec.loc[4, "match_score"], 1) == 83.3


def test_threshold_applies_to_recommended_sources_only():
    dest, table = destination(["EMP005"]), source_table([["EMP006", 6]])
    assert list(match_source(dest, spec("r", "R.xlsx", settings=MatchSettings(80, 5)), table)["status"]) == ["FUZZY"]
    assert list(match_source(dest, spec("s", "S.xlsx", settings=MatchSettings(80, 5, "strict")), table)["status"]) == ["UNMATCHED"]


def test_two_sources_use_their_own_modes_independently_and_in_any_order():
    dest = destination(HR_IDS)
    ss, sr = spec("s", "S.xlsx", "strict"), spec("r", "R.xlsx", "recommended")
    alone_s = run_match(HR_SRC, dest.table, KEY, KEY, mode="strict")[0]
    alone_r = run_match(HR_SRC, dest.table, KEY, KEY, mode="recommended")[0]
    s1, r1 = match_source(dest, ss, HR_SRC), match_source(dest, sr, HR_SRC)
    r2, s2 = match_source(dest, sr, HR_SRC), match_source(dest, ss, HR_SRC)
    for got, expected in ((s1, alone_s), (s2, alone_s), (r1, alone_r), (r2, alone_r)):
        pd.testing.assert_frame_equal(got, expected)


def test_consolidation_and_report_keep_each_sources_mode_results_separate():
    dest = destination(HR_IDS)
    es = [entry(dest, "s", "Strict.xlsx", HR_SRC, MatchSettings(90, 5, "strict")),
          entry(dest, "r", "Recommended.xlsx", HR_SRC, MatchSettings(90, 5))]
    c = consolidate(dest, es)
    summaries = {x.source_id: x for x in c.summaries}
    assert (summaries["s"].results["Not found"], summaries["s"].results["Unclear"]) == (1, 0)
    assert (summaries["r"].results["Not found"], summaries["r"].results["Unclear"]) == (0, 1)
    assert c.can_apply and not c.blockers and not c.conflicts
    assert len(c.updates) == 5 and all(u.source_ids == ("s", "r") for u in c.updates)

    report = pd.read_csv(io.BytesIO(matching_report_csv(dest, es)))
    hr = report[report["Record in destination file"] == "EMP005"].set_index("Source file")
    assert hr.loc["Strict.xlsx", "Result"] == "Not found" and pd.isna(hr.loc["Strict.xlsx", "Similarity (0-100)"])
    assert hr.loc["Recommended.xlsx", "Result"] == "Unclear" and hr.loc["Recommended.xlsx", "Similarity (0-100)"] == 83.3


# ---------------------------------------------------------------- strict sources never carry possible matches
@pytest.mark.parametrize("status", ["FUZZY", "AMBIGUOUS", "NORMALIZED"])
def test_a_strict_source_with_non_strict_results_is_blocked(status):
    dest = destination(HR_IDS)
    rows = [("EXACT", 0, None, None), (status, 1 if status == "NORMALIZED" else None,
                                       1 if status == "FUZZY" else None, 91.0 if status != "NORMALIZED" else None)]
    rows += [("UNMATCHED", None, None, None)] * 4
    res = fake_result(*rows)
    for mode, blocked in (("strict", True), ("recommended", False)):
        s = spec("a", "A.xlsx", mode)
        e = SourceEntry(s, HR_SRC, res)
        c = consolidate(dest, [replace(e, plan=build_source_plan(dest, e))])
        assert bool(c.blockers) is blocked and c.can_apply is (not blocked)
        if blocked:
            assert "different matching settings" in c.blockers[0].reason and c.blockers[0].label == "A.xlsx"


def test_genuine_strict_results_are_never_blocked():
    dest = destination(HR_IDS)
    table = source_table([["EMP001", 1], ["EMP001", 2], ["EMP002", 3], ["emp003", 4]])      # duplicate and formatting variant
    c = consolidate(dest, [entry(dest, "s", "S.xlsx", table, MatchSettings(90, 5, "strict"))])
    assert not c.blockers and c.included == ("s",)


# ---------------------------------------------------------------- existing behaviour is unchanged
def test_one_source_with_default_settings_behaves_exactly_as_before():
    dest = destination(HR_IDS)
    sp = spec("a", "A.xlsx")                                                    # no mode given
    via = match_source(dest, sp, HR_SRC)
    direct, _ = run_match(HR_SRC, dest.table, KEY, KEY, 90, 5)                  # the call made before this change
    pd.testing.assert_frame_equal(via, direct)
    e = SourceEntry(sp, HR_SRC, via)
    c = consolidate(dest, [replace(e, plan=build_source_plan(dest, e))])
    plan, _ = build_plan(dest.file_bytes, "Data", direct, HR_SRC, dest.table, {"Salary": "Salary"})
    assert [(u.row, u.column, u.new_value) for u in c.updates] == [(u["row"], u["column"], u["new_value"]) for u in plan]
    assert c.can_apply and not c.blockers


def test_build_source_plan_still_consumes_existing_results_and_never_rematches(monkeypatch):
    dest = destination(HR_IDS)
    sp = spec("a", "A.xlsx", "strict")
    e = SourceEntry(sp, HR_SRC, match_source(dest, sp, HR_SRC))

    def forbidden(*a, **k):
        raise AssertionError("build_source_plan must not run the matcher")

    monkeypatch.setattr(cons, "run_match", forbidden)
    assert len(build_source_plan(dest, e).updates) == 5