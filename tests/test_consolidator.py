"""Unit and integration tests for core/consolidator.py.

Uses the real core.reader, core.matcher and core.updater (never modified) with small in-memory
workbooks, plus hand-built match results where a specific status (for example a possible match)
is needed regardless of the fuzzy-matching scores.
"""
import hashlib
import io
import sys
from dataclasses import replace
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

import core.consolidator as cons
from core.consolidator import (
    ConsolidationBlocked, MatchSettings, SourceEntry, SourcePlan, SourceSpec, apply_consolidated,
    build_source_plan, conflict_table, consolidate, make_destination, matching_report_csv,
    unused_source_records, validate_setup,
)
from core.matcher import run_match
from core.reader import load_table
from core.updater import apply_plan, build_plan

HEAD = ["District", "Target", "Population", "Literacy"]
TARGET = [["Kolar", 750000], ["Mysuru", 400000], ["Hassan", 100]]
POP = [["Kolar", 650000], ["Mysuru", 350000], ["Tumkur", 5], ["Mandya", 10]]
LIT = [["Kolar", 0.82], ["Mysuru", 0.85], ["Hassan", 0.9]]


# ---------------------------------------------------------------- helpers
def plain_book(headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(headers)
    for r in rows:
        ws.append(r)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def dst_book(extra_rows=()):
    """Destination with formatting, one formula (D4) and one pre-filled value (C5)."""
    rows = [["Kolar", None, None, None], ["Mysuru", None, None, None],
            ["Hassan", None, None, None], ["Tumkur", None, 999, None], *extra_rows]
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(HEAD)
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.font = Font(bold=True)
    for (c,) in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        c.number_format = "#,##0"
        c.fill = PatternFill("solid", fgColor="FFF2CC")
    ws["D4"] = "=1+1"
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def fh(name, data):
    f = io.BytesIO(data)
    f.name = name
    return f


def load(data, name="f.xlsx"):
    return load_table(fh(name, data), sheet="Data")


def destination(data=None):
    data = data or dst_book()
    return make_destination(data, "Data", 1, load(data, "dest.xlsx"))


def spec(sid, label, mapping, **kw):
    base = dict(sheet="Data", header_row=1, src_keys=("District",), dst_keys=("District",),
                settings=MatchSettings(90, 5), strategy="fill_only_blank")
    base.update(kw)
    return SourceSpec(sid, label, mapping=mapping, **base)


def entry(dest, sid, label, headers, rows, mapping, *, approve=(), build=True, **kw):
    s = spec(sid, label, mapping, **kw)
    table = load(plain_book(headers, rows), label)
    res, _ = run_match(table, dest.table, list(s.src_keys), list(s.dst_keys),
                       s.settings.threshold, s.settings.margin)
    e = SourceEntry(s, table, res, frozenset(approve))
    return replace(e, plan=build_source_plan(dest, e)) if build else e


def three(dest):
    return [
        entry(dest, "t", "Target.xlsx", ["District", "Target"], TARGET, {"Target": "Target"}),
        entry(dest, "p", "Population.xlsx", ["District", "Population"], POP, {"Population": "Population"}),
        entry(dest, "l", "Literacy.xlsx", ["District", "Literacy"], LIT, {"Literacy": "Literacy"}),
    ]


def fake_result(*rows):
    """One (status, src_index, suggested_index, score) per destination row, in the matcher's format."""
    return pd.DataFrame([
        {"_dst_index": i, "_src_index": s, "_suggested_src_index": g, "match_score": sc, "status": st}
        for i, (st, s, g, sc) in enumerate(rows)])


def cell_map(data):
    ws = load_workbook(io.BytesIO(data))["Data"]
    return {(c.row, c.column): c.value for row in ws.iter_rows() for c in row}


def updates_by_cell(consolidation):
    return {(u.row, u.column): u.new_value for u in consolidation.updates}


# ---------------------------------------------------------------- backend contract
def test_backend_contract_build_plan_and_apply_plan():
    dest = destination()
    e = entry(dest, "p", "Population.xlsx", ["District", "Population"], POP, {"Population": "Population"}, build=False)
    out = build_plan(dest.file_bytes, "Data", e.match_result, e.table, dest.table, {"Population": "Population"})
    assert isinstance(out, tuple) and len(out) == 2 and all(isinstance(x, list) for x in out)
    plan, skips = out
    assert set(plan[0]) == {"row", "column", "column_name", "old_value", "new_value", "action"}
    assert all("destination_index" in s and "reason" in s for s in skips)
    assert isinstance(apply_plan(dest.file_bytes, "Data", plan), bytes)


def test_consolidator_has_no_streamlit_dependency():
    assert "import streamlit" not in Path(cons.__file__).read_text()


# ---------------------------------------------------------------- destination
def test_destination_signature_depends_on_bytes_sheet_and_header_row():
    data = dst_book()
    base = make_destination(data, "Data", 1, load(data)).signature
    assert make_destination(data, "Data", 1, load(data)).signature == base
    assert make_destination(data, "Other", 1, load(data)).signature != base
    assert make_destination(data, "Data", 2, load(data)).signature != base
    other = dst_book([["Extra", None, None, None]])
    assert make_destination(other, "Data", 1, load(other)).signature != base


def test_make_destination_rejects_bad_input():
    data = dst_book()
    with pytest.raises(ValueError):
        make_destination(data, "Data", 1, load(data).drop(columns="_xl_row"))
    with pytest.raises(TypeError):
        make_destination("not bytes", "Data", 1, load(data))


# ---------------------------------------------------------------- validate_setup
BASE = spec("s", "S.xlsx", {"Target": "Target"})
SRC_COLS = {"s": ["District", "Target"]}


def issue_codes(specs, src_cols=SRC_COLS, severity="error"):
    return [i.code for i in validate_setup(specs, HEAD, src_cols) if i.severity == severity]


def test_valid_setup_has_no_errors():
    assert issue_codes([BASE]) == []


@pytest.mark.parametrize("change, expected", [
    ({"src_keys": ()}, "NO_KEYS"),
    ({"src_keys": ("District", "Target")}, "KEY_COUNT_MISMATCH"),
    ({"src_keys": ("Nope",)}, "UNKNOWN_KEY_COLUMN"),
    ({"dst_keys": ("Nope",)}, "UNKNOWN_KEY_COLUMN"),
    ({"mapping": {}}, "NO_MAPPING"),
    ({"mapping": {"Nope": "Target"}}, "UNKNOWN_SOURCE_COLUMN"),
    ({"mapping": {"Target": "Nope"}}, "UNKNOWN_DEST_COLUMN"),
    ({"mapping": {"Target": "_xl_row"}}, "UNKNOWN_DEST_COLUMN"),
    ({"mapping": {"District": "Population", "Target": "Population"}}, "DUPLICATE_TARGET"),
    ({"mapping": {"Target": "District"}}, "TARGET_IS_KEY"),
    ({"strategy": "merge"}, "INVALID_STRATEGY"),
])
def test_setup_errors(change, expected):
    assert expected in issue_codes([replace(BASE, **change)])


def test_key_protection_covers_keys_of_other_sources():
    a = replace(BASE, mapping={"Target": "Literacy"})
    b = replace(BASE, source_id="b", label="B.xlsx", dst_keys=("Literacy",))
    issues = validate_setup([a, b], HEAD, {"s": ["District", "Target"], "b": ["District", "Target"]})
    hit = [i for i in issues if i.code == "TARGET_IS_KEY"]
    assert [i.source_id for i in hit] == ["s"] and "B.xlsx" in hit[0].message


def test_excluded_sources_are_not_validated_and_do_not_protect_columns():
    a = replace(BASE, mapping={"Target": "Literacy"})
    b = replace(BASE, source_id="b", label="B.xlsx", dst_keys=("Literacy",), mapping={}, excluded=True)
    assert issue_codes([a, b], {"s": ["District", "Target"], "b": []}) == []


def test_shared_target_is_a_warning_not_an_error():
    a = replace(BASE, mapping={"Target": "Population"})
    b = replace(BASE, source_id="b", label="B.xlsx", mapping={"Target": "Population"})
    cols = {"s": ["District", "Target"], "b": ["District", "Target"]}
    assert issue_codes([a, b], cols) == []
    assert issue_codes([a, b], cols, severity="warning") == ["SHARED_TARGET"]


def test_duplicate_ids_and_unreadable_source_columns_are_errors():
    assert "DUPLICATE_SOURCE_ID" in issue_codes([BASE, BASE])
    assert "SOURCE_COLUMNS_UNAVAILABLE" in issue_codes([BASE], {})


# ---------------------------------------------------------------- build_source_plan
def test_build_source_plan_matches_updater_output_and_stamps_destination():
    dest = destination()
    e = entry(dest, "p", "Population.xlsx", ["District", "Population"], POP, {"Population": "Population"}, build=False)
    before = hashlib.sha256(dest.file_bytes).hexdigest()
    plan = build_source_plan(dest, e)
    direct, skips = build_plan(dest.file_bytes, "Data", e.match_result, e.table, dest.table, {"Population": "Population"})
    assert list(plan.updates) == direct and list(plan.skips) == skips
    assert plan.dest_signature == dest.signature and plan.source_id == "p"
    assert hashlib.sha256(dest.file_bytes).hexdigest() == before


def test_build_source_plan_refuses_excluded_or_unmatched_sources():
    dest = destination()
    e = entry(dest, "p", "Population.xlsx", ["District", "Population"], POP, {"Population": "Population"}, build=False)
    with pytest.raises(ValueError):
        build_source_plan(dest, replace(e, spec=replace(e.spec, excluded=True)))
    with pytest.raises(ValueError):
        build_source_plan(dest, replace(e, match_result=None))


def test_build_source_plan_rejects_unexpected_backend_structure(monkeypatch):
    dest = destination()
    e = entry(dest, "p", "Population.xlsx", ["District", "Population"], POP, {"Population": "Population"}, build=False)
    monkeypatch.setattr(cons, "build_plan", lambda *a, **k: [])
    with pytest.raises(TypeError):
        build_source_plan(dest, e)


def test_possible_matches_need_explicit_approval_per_source():
    data = plain_book(["District", "Population"], [["Kolar", None], ["Mysoor", None]])
    dest = make_destination(data, "Data", 1, load(data))
    table = load(plain_book(["District", "Population"], [["Kolar", 1], ["Mysuru", 2]]))
    res = fake_result(("EXACT", 0, None, None), ("FUZZY", None, 1, 91.4))
    s = spec("p", "P.xlsx", {"Population": "Population"})
    rows = lambda approve: {u["row"] for u in build_source_plan(dest, SourceEntry(s, table, res, frozenset(approve))).updates}
    assert rows(()) == {2}            # not approved: the possible match copies nothing
    assert rows({1}) == {2, 3}        # approved: it is copied


# ---------------------------------------------------------------- consolidate: merging
def test_three_sources_produce_the_expected_consolidated_updates():
    dest = destination()
    c = consolidate(dest, three(dest))
    assert updates_by_cell(c) == {
        (2, 2): 750000, (2, 3): 650000, (2, 4): 0.82,
        (3, 2): 400000, (3, 3): 350000, (3, 4): 0.85,
        (4, 2): 100,                    # Hassan: Population missing; Literacy cell is a formula
    }                                   # Tumkur: Population already filled; not in Target or Literacy
    assert c.included == ("t", "p", "l") and not c.blockers and not c.conflicts and c.can_apply


def test_single_source_consolidation_equals_the_updater_plan():
    dest = destination()
    e = three(dest)[1]
    c = consolidate(dest, [e])
    assert [(u.row, u.column, u.column_name, u.old_value, u.new_value, u.action) for u in c.updates] == \
           [(u["row"], u["column"], u["column_name"], u["old_value"], u["new_value"], u["action"]) for u in e.plan.updates]


def test_same_column_with_different_rows_is_not_a_conflict():
    dest = destination()
    a = entry(dest, "p", "P.xlsx", ["District", "Population"], POP, {"Population": "Population"})
    b = entry(dest, "q", "Q.xlsx", ["District", "Population"], [["Hassan", 77]], {"Population": "Population"})
    c = consolidate(dest, [a, b])
    assert not c.conflicts and updates_by_cell(c)[(4, 3)] == 77 and c.can_apply


def test_identical_values_for_the_same_cell_merge_into_one_update():
    dest = destination()
    a = entry(dest, "p", "P.xlsx", ["District", "Population"], POP, {"Population": "Population"})
    b = entry(dest, "q", "Q.xlsx", ["District", "Population"], [["Kolar", 650000.0]], {"Population": "Population"})
    c = consolidate(dest, [a, b])
    merged = [u for u in c.updates if (u.row, u.column) == (2, 3)]
    assert not c.conflicts and len(merged) == 1 and merged[0].source_ids == ("p", "q")


def test_different_values_for_the_same_cell_block_and_are_reported():
    dest = destination()
    a = entry(dest, "p", "P.xlsx", ["District", "Population"], POP, {"Population": "Population"})
    b = entry(dest, "q", "Q.xlsx", ["District", "Population"], [["Kolar", 111]], {"Population": "Population"})
    c = consolidate(dest, [a, b])
    assert not c.can_apply and len(c.conflicts) == 1
    assert (2, 3) not in updates_by_cell(c) and (3, 3) in updates_by_cell(c)   # only the clashing cell is held back
    t = conflict_table(c)
    assert list(t.columns) == list(cons.CONFLICT_COLUMNS)
    assert sorted(t["Proposed value"]) == [111, 650000] and set(t["Current value"]) == {"(empty)"}
    assert set(t["Source file"]) == {"P.xlsx", "Q.xlsx"}


@pytest.mark.parametrize("a, b, same", [
    (5, 5.0, True), ("x", "x", True), (5, "5", False), (True, 1, False), ("a", "A", False), (1.5, 1.50, True),
])
def test_value_equality_rules(a, b, same):
    assert cons._same_value(a, b) is same


def test_result_does_not_depend_on_source_order():
    dest = destination()
    es = three(dest)
    es.append(entry(dest, "q", "Q.xlsx", ["District", "Population"], [["Kolar", 111]], {"Population": "Population"}))

    def view(c):
        return ({(u.row, u.column, u.new_value, tuple(sorted(u.source_ids))) for u in c.updates},
                {(x.row, x.column, tuple(sorted((k.source_id, k.new_value) for k in x.candidates))) for x in c.conflicts})

    assert view(consolidate(dest, es)) == view(consolidate(dest, list(reversed(es))))


def test_sources_with_different_matching_columns_work_together():
    data = plain_book(["District", "Code", "Target", "Population"],
                      [["Kolar", "K01", None, None], ["Mysuru", "M02", None, None]])
    dest = make_destination(data, "Data", 1, load(data))
    a = entry(dest, "a", "A.xlsx", ["District", "Target"], [["Kolar", 1], ["Mysuru", 2]], {"Target": "Target"})
    b = entry(dest, "b", "B.xlsx", ["Code", "Population"], [["K01", 10], ["M02", 20]], {"Population": "Population"},
              src_keys=("Code",), dst_keys=("Code",))
    c = consolidate(dest, [a, b])
    assert updates_by_cell(c) == {(2, 3): 1, (3, 3): 2, (2, 4): 10, (3, 4): 20}


def test_existing_value_strategy_is_independent_per_source():
    dest = destination()
    keep = entry(dest, "p", "P.xlsx", ["District", "Population"], POP, {"Population": "Population"})
    over = entry(dest, "o", "O.xlsx", ["District", "Population"], POP, {"Population": "Population"},
                 strategy="overwrite")
    assert (5, 3) not in updates_by_cell(consolidate(dest, [keep]))
    assert updates_by_cell(consolidate(dest, [over]))[(5, 3)] == 5


# ---------------------------------------------------------------- merging identical proposals
POP_H = ["District", "Population"]
POP_MAP = {"Population": "Population"}


def both_strategies(dest, value=200):
    """The same proposal for Kolar from an overwrite source and from a fill-only source."""
    a = entry(dest, "a", "A.xlsx", POP_H, [["Kolar", value]], POP_MAP, strategy="overwrite")
    b = entry(dest, "b", "B.xlsx", POP_H, [["Kolar", value]], POP_MAP, strategy="fill_only_blank")
    return a, b


def merged_view(c):
    return [(u.row, u.column, u.old_value, u.new_value, u.action, tuple(sorted(u.source_ids))) for u in c.updates]


def forge(e, **changes):
    """Copy of an entry whose plan breaks the 'same original destination' invariant (a buggy caller)."""
    ups = tuple({**u, **changes} for u in e.plan.updates)
    return replace(e, plan=SourcePlan(e.plan.source_id, e.plan.dest_signature, ups, e.plan.skips))


@pytest.mark.parametrize("cell", [None, "   "])
def test_overwrite_and_fill_only_sources_merge_on_a_blank_cell_in_any_order(cell):
    data = plain_book(POP_H, [["Kolar", cell]])
    dest = make_destination(data, "Data", 1, load(data))
    a, b = both_strategies(dest)
    forward, backward = consolidate(dest, [a, b]), consolidate(dest, [b, a])
    assert merged_view(forward) == merged_view(backward) == [(2, 2, cell, 200, "UPDATE", ("a", "b"))]
    assert forward.can_apply and not forward.conflicts and not forward.blockers


def test_on_a_filled_cell_only_the_overwrite_source_proposes_a_change():
    dest = destination()                                    # Tumkur's Population (C5) already holds 999
    a = entry(dest, "a", "A.xlsx", POP_H, [["Tumkur", 200]], POP_MAP, strategy="overwrite")
    b = entry(dest, "b", "B.xlsx", POP_H, [["Tumkur", 200]], POP_MAP, strategy="fill_only_blank")
    assert b.plan.updates == ()                             # fill-only never proposes to change a filled cell
    c = consolidate(dest, [a, b])
    assert merged_view(c) == [(5, 3, 999, 200, "UPDATE", ("a",))] and not c.conflicts


def test_sources_proposing_the_same_cell_always_agree_on_its_current_value():
    rows = [["Kolar", None], ["Mysuru", "   "], ["Hassan", 100], ["Tumkur", 150]]
    data = plain_book(POP_H, rows)
    dest = make_destination(data, "Data", 1, load(data))
    proposals = [[r[0], 200] for r in rows]
    a = entry(dest, "a", "A.xlsx", POP_H, proposals, POP_MAP, strategy="overwrite")
    b = entry(dest, "b", "B.xlsx", POP_H, proposals, POP_MAP, strategy="fill_only_blank")
    ua = {(u["row"], u["column"]): u for u in a.plan.updates}
    ub = {(u["row"], u["column"]): u for u in b.plan.updates}
    assert set(ua) == {(2, 2), (3, 2), (4, 2), (5, 2)} and set(ub) == {(2, 2), (3, 2)}
    original = cell_map(data)
    for cell in set(ua) & set(ub):
        assert ua[cell]["old_value"] == ub[cell]["old_value"] == original[cell]
        assert ua[cell]["action"] == ub[cell]["action"]
    c = consolidate(dest, [a, b])
    assert not c.conflicts and not c.blockers and c.can_apply
    assert {u.row: u.old_value for u in c.updates} == {2: None, 3: "   ", 4: 100, 5: 150}


@pytest.mark.parametrize("changes", [{"old_value": 150}, {"action": "REPLACE"}, {"old_value": 150, "new_value": 300}])
def test_plans_that_disagree_about_a_cells_current_value_or_action_are_held_back(changes):
    dest = destination()
    a, b = both_strategies(dest)
    other = entry(dest, "c", "C.xlsx", ["District", "Target"], [["Mysuru", 400000]], {"Target": "Target"})
    bad = forge(b, **changes)
    results = [consolidate(dest, order) for order in ([a, bad, other], [other, bad, a], [bad, a, other])]
    for c in results:
        assert not c.can_apply and not c.conflicts                      # reported as an integrity problem
        assert sorted(x.source_id for x in c.blockers) == ["a", "b"]
        assert all("Population" in x.reason for x in c.blockers)
        assert (2, 3) not in updates_by_cell(c)                         # neither source's value is kept
        assert updates_by_cell(c)[(3, 2)] == 400000                     # unrelated cells still merge
        with pytest.raises(ConsolidationBlocked):
            apply_consolidated(dest, c)
    assert len({frozenset((x.source_id, x.reason) for x in c.blockers) for c in results}) == 1   # order independent


# ---------------------------------------------------------------- consolidate: per-source problems
def test_one_source_without_matches_does_not_affect_the_others():
    dest = destination()
    es = three(dest)
    es.append(entry(dest, "n", "None.xlsx", ["District", "Population"], [["Nowhere", 1]], {"Population": "Population"}))
    c = consolidate(dest, es)
    assert c.can_apply and "n" in c.included and len(c.updates) == 7
    assert next(s for s in c.summaries if s.source_id == "n").results["Not found"] == 4


def test_duplicate_keys_in_one_source_only_hold_back_that_source():
    dest = destination()
    t = entry(dest, "t", "Target.xlsx", ["District", "Target"], TARGET, {"Target": "Target"})
    p = entry(dest, "p", "P.xlsx", ["District", "Population"],
              [["Kolar", 1], ["Kolar", 2], ["Mysuru", 350000]], {"Population": "Population"})
    u = updates_by_cell(consolidate(dest, [t, p]))
    assert (2, 2) in u and (2, 3) not in u and u[(3, 3)] == 350000
    assert any(s["reason"] == "Match status: DUPLICATE_SOURCE" for s in p.plan.skips)


def test_blank_source_values_are_skipped_without_creating_conflicts():
    dest = destination()
    a = entry(dest, "p", "P.xlsx", ["District", "Population"], [["Kolar", None]], {"Population": "Population"})
    b = entry(dest, "q", "Q.xlsx", ["District", "Population"], [["Kolar", 650000]], {"Population": "Population"})
    c = consolidate(dest, [a, b])
    assert not c.conflicts and updates_by_cell(c) == {(2, 3): 650000}
    assert any(s.get("reason") == "Blank source value" for s in a.plan.skips)


def test_summaries_count_results_confirmations_updates_and_skips():
    dest = destination()
    s = {x.source_id: x for x in consolidate(dest, three(dest)).summaries}["p"]
    assert s.results["Matched"] == 3 and s.results["Not found"] == 1
    assert (s.updates, s.skips, s.needs_review) == (2, 2, 0)


def test_explicit_exclusion_is_listed_and_not_a_blocker():
    dest = destination()
    es = three(dest)
    es[2] = replace(es[2], spec=replace(es[2].spec, excluded=True))
    c = consolidate(dest, es)
    assert [x.source_id for x in c.excluded] == ["l"] and c.included == ("t", "p")
    assert not c.blockers and c.can_apply and all(col != 4 for _, col in updates_by_cell(c))


def test_excluding_every_source_leaves_nothing_to_apply():
    dest = destination()
    es = [replace(e, spec=replace(e.spec, excluded=True)) for e in three(dest)]
    c = consolidate(dest, es)
    assert len(c.excluded) == 3 and not c.updates and not c.can_apply
    assert not consolidate(dest, []).can_apply


def blocked_by(e_changed, expected_text):
    dest = destination()
    es = three(dest)
    es[1] = e_changed(es[1], dest)
    c = consolidate(dest, es)
    assert not c.can_apply
    assert [b.label for b in c.blockers] == ["Population.xlsx"] and expected_text in c.blockers[0].reason
    assert c.included == ("t", "l") and len(c.updates) == 5      # the others are still merged for the preview


def test_failed_source_blocks_with_its_reason():
    blocked_by(lambda e, d: replace(e, problem="We couldn't read this file."), "couldn't read")


def test_source_without_reviewed_changes_blocks():
    blocked_by(lambda e, d: replace(e, plan=None), "not been reviewed")


def test_source_planned_against_another_destination_blocks():
    def stale(e, d):
        other = destination(dst_book([["Extra", None, None, None]]))
        return entry(other, "p", "Population.xlsx", ["District", "Population"], POP, {"Population": "Population"})
    blocked_by(stale, "different version")


def test_mapping_into_a_key_column_blocks_even_if_validation_was_skipped():
    dest = destination()
    e = entry(dest, "p", "P.xlsx", ["District", "Population"], POP, {"Population": "District"}, build=True)
    c = consolidate(dest, [e])
    assert not c.can_apply and "identify records" in c.blockers[0].reason


def test_duplicate_source_ids_are_rejected():
    dest = destination()
    e = three(dest)[0]
    with pytest.raises(ValueError):
        consolidate(dest, [e, e])


# ---------------------------------------------------------------- apply_consolidated
def test_apply_consolidated_calls_apply_plan_exactly_once(monkeypatch):
    dest = destination()
    c = consolidate(dest, three(dest))
    calls = []
    monkeypatch.setattr(cons, "apply_plan", lambda *a: calls.append(a) or b"workbook")
    assert apply_consolidated(dest, c) == b"workbook"
    assert len(calls) == 1
    data, sheet, plain = calls[0]
    assert data is dest.file_bytes and sheet == "Data" and len(plain) == 7
    assert all(set(u) == {"row", "column", "column_name", "old_value", "new_value", "action"} for u in plain)


def test_apply_consolidated_refuses_unsafe_states_without_writing(monkeypatch):
    dest = destination()
    calls = []
    monkeypatch.setattr(cons, "apply_plan", lambda *a: calls.append(a) or b"x")
    good = three(dest)
    conflicting = good + [entry(dest, "q", "Q.xlsx", ["District", "Population"], [["Kolar", 1]], {"Population": "Population"})]
    cases = {
        "conflict": consolidate(dest, conflicting),
        "blocker": consolidate(dest, [replace(good[0], problem="Broken")] + good[1:]),
        "nothing to do": consolidate(dest, [replace(e, spec=replace(e.spec, excluded=True)) for e in good]),
        "other destination": replace(consolidate(dest, good), dest_signature="different"),
    }
    for name, c in cases.items():
        with pytest.raises(ConsolidationBlocked):
            apply_consolidated(dest, c)
    assert calls == []


def test_end_to_end_values_formulas_formatting_and_original_bytes():
    dest = destination()
    original = bytes(dest.file_bytes)
    out = apply_consolidated(dest, consolidate(dest, three(dest)))
    ws = load_workbook(io.BytesIO(out))["Data"]
    assert [ws.cell(2, c).value for c in (1, 2, 3, 4)] == ["Kolar", 750000, 650000, 0.82]
    assert [ws.cell(3, c).value for c in (1, 2, 3, 4)] == ["Mysuru", 400000, 350000, 0.85]
    assert ws["B4"].value == 100 and ws["C4"].value is None
    assert ws["D4"].value == "=1+1"                     # formula untouched
    assert ws["C5"].value == 999                        # existing value kept
    assert ws["A1"].font.bold and ws["B2"].number_format == "#,##0"
    assert ws["B2"].fill.fgColor.rgb.endswith("FFF2CC") and ws.max_row == 5
    assert dest.file_bytes == original                  # the uploaded destination is never modified


def test_single_pass_equals_applying_each_source_in_turn():
    dest = destination()
    es = three(dest)
    data = dest.file_bytes
    for e in es:
        data = apply_plan(data, "Data", list(e.plan.updates))
    assert cell_map(apply_consolidated(dest, consolidate(dest, es))) == cell_map(data)


def test_single_source_result_equals_the_existing_workflow():
    dest = destination()
    e = three(dest)[1]
    direct = apply_plan(dest.file_bytes, "Data", list(e.plan.updates))
    assert cell_map(apply_consolidated(dest, consolidate(dest, [e]))) == cell_map(direct)


# ---------------------------------------------------------------- matching report
STATUS_REF = {
    "EXACT": ("Matched", "Exact match"), "NORMALIZED": ("Matched", "Match after ignoring formatting differences"),
    "FUZZY": ("Needs review", "Possible match"), "AMBIGUOUS": ("Unclear", "Several similar records found"),
    "DUPLICATE_SOURCE": ("Duplicate", "Duplicate in source file"),
    "DUPLICATE_DEST": ("Duplicate", "Duplicate in destination file"),
    "INVALID_KEY": ("Invalid information", "Identifying information is missing"),
    "UNMATCHED": ("Not found", "No match found"),
}


def legacy_report_csv(approved, res, src_df, dst_df, src_keys, dst_keys):
    """Verbatim logic of the single-source report that shipped in app.py (reference for regression)."""
    accepted, rows = set(approved), []
    for _, r in res.iterrows():
        status, dst_index = r["status"], r["_dst_index"]
        src_index = r["_src_index"] if status in ("EXACT", "NORMALIZED") else r["_suggested_src_index"]
        score = r["match_score"]
        decision = ("Confirmed" if dst_index in accepted else "Not confirmed") if status == "FUZZY" else "Not needed"
        rows.append({
            "Row in destination file": cons._xl_row(dst_df, dst_index),
            "Record in destination file": cons._text(dst_df, dst_index, dst_keys),
            "Result": STATUS_REF[status][0], "Details": STATUS_REF[status][1],
            "Row in source file": cons._xl_row(src_df, src_index),
            "Record in source file": cons._text(src_df, src_index, src_keys),
            "Similarity (0-100)": "" if pd.isna(score) else round(float(score), 1),
            "Your decision": decision,
        })
    return pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig")


def mixed_status_entry(approve, label="S.xlsx", sid="s"):
    data = plain_book(["District"], [["Kolar"], ["Mysoor"], ["Hasan"], [None], ["Mand"], ["Tumkur"]])
    dest = make_destination(data, "Data", 1, load(data))
    table = load(plain_book(["District", "Population"], [["Kolar", 1], ["Mysuru", 2], ["Hassan", 3]]))
    res = fake_result(("EXACT", 0, None, None), ("FUZZY", None, 1, 91.4), ("FUZZY", None, 2, 93.0),
                      ("INVALID_KEY", None, None, None), ("AMBIGUOUS", None, None, 88.0), ("UNMATCHED", None, None, None))
    return dest, SourceEntry(spec(sid, label, {"Population": "District"}), table, res, frozenset(approve))


def test_single_source_report_is_identical_to_the_existing_report():
    dest, e = mixed_status_entry({1})
    got = matching_report_csv(dest, [e])
    assert got == legacy_report_csv({1}, e.match_result, e.table, dest.table, ["District"], ["District"])
    assert got.startswith(b"\xef\xbb\xbf") and b"Source file" not in got


def test_report_distinguishes_confirmed_from_unconfirmed_possible_matches():
    dest, e = mixed_status_entry({1})
    df = pd.read_csv(io.BytesIO(matching_report_csv(dest, [e])))
    fuzzy = df[df["Result"] == "Needs review"]
    assert list(fuzzy["Your decision"]) == ["Confirmed", "Not confirmed"]
    assert set(df[df["Result"] != "Needs review"]["Your decision"]) == {"Not needed"}


def test_multi_source_report_adds_source_file_and_keeps_decisions_separate():
    dest, a = mixed_status_entry({1}, "A.xlsx", "a")
    _, b = mixed_status_entry((), "B.xlsx", "b")
    df = pd.read_csv(io.BytesIO(matching_report_csv(dest, [a, b])))
    assert list(df.columns) == ["Source file", *cons.REPORT_COLUMNS]
    assert len(df) == 12 and list(df["Source file"].unique()) == ["A.xlsx", "B.xlsx"]
    decisions = df[df["Result"] == "Needs review"].groupby("Source file")["Your decision"].apply(list)
    assert decisions["A.xlsx"] == ["Confirmed", "Not confirmed"] and decisions["B.xlsx"] == ["Not confirmed"] * 2


def test_excluded_sources_are_left_out_but_the_source_column_stays_with_several_slots():
    dest, a = mixed_status_entry({1}, "A.xlsx", "a")
    _, b = mixed_status_entry((), "B.xlsx", "b")
    b = replace(b, spec=replace(b.spec, excluded=True))
    df = pd.read_csv(io.BytesIO(matching_report_csv(dest, [a, b])))
    assert list(df["Source file"].unique()) == ["A.xlsx"]


def test_report_uses_each_sources_own_destination_matching_columns():
    data = plain_book(["District", "Code"], [["Kolar", "K01"], ["Mysuru", "M02"]])
    dest = make_destination(data, "Data", 1, load(data))
    a = entry(dest, "a", "A.xlsx", ["District", "X"], [["Kolar", 1]], {"X": "Code"}, build=False)
    b = entry(dest, "b", "B.xlsx", ["Code", "X"], [["M02", 1]], {"X": "District"}, build=False,
              src_keys=("Code",), dst_keys=("Code",))
    df = pd.read_csv(io.BytesIO(matching_report_csv(dest, [a, b])))
    assert list(df[df["Source file"] == "A.xlsx"]["Record in destination file"]) == ["Kolar", "Mysuru"]
    assert list(df[df["Source file"] == "B.xlsx"]["Record in destination file"]) == ["K01", "M02"]


def test_report_is_available_with_zero_updates_and_without_a_plan():
    dest = destination()
    e = entry(dest, "n", "None.xlsx", ["District", "Population"], [["Nowhere", 1]], {"Population": "Population"}, build=False)
    assert not consolidate(dest, [replace(e, plan=build_source_plan(dest, e))]).can_apply
    df = pd.read_csv(io.BytesIO(matching_report_csv(dest, [e])))
    assert len(df) == 4 and set(df["Result"]) == {"Not found"}


def test_report_with_nothing_matched_is_header_only():
    out = matching_report_csv(destination(), [])
    assert out.decode("utf-8-sig").strip() == ",".join(cons.REPORT_COLUMNS)


# ---------------------------------------------------------------- unused source records
def test_unused_records_lists_source_rows_that_matched_nothing():
    dest = destination()
    e = three(dest)[1]
    df = unused_source_records(e)
    assert list(df.columns) == list(cons.UNUSED_COLUMNS)
    assert df.to_dict("records") == [{"Row in source file": 5, "Record in source file": "Mandya",
                                      "Reason": "No matching destination record"}]


def test_unused_records_treats_confirmed_possible_matches_as_used():
    dest, e = mixed_status_entry(())
    reasons = dict(zip(unused_source_records(e)["Record in source file"], unused_source_records(e)["Reason"]))
    assert reasons["Mysuru"] == "Possible match not confirmed" and "Kolar" not in reasons
    confirmed = unused_source_records(replace(e, approved_rows=frozenset({1})))
    assert "Mysuru" not in set(confirmed["Record in source file"])


def test_duplicate_source_records_are_labelled_as_duplicates_not_unmatched():
    dest = destination()
    e = entry(dest, "p", "P.xlsx", ["District", "Population"],
              [["Kolar", 1], ["Kolar", 2], ["Mysuru", 350000]], {"Population": "Population"})
    assert e.match_result.loc[0, "status"] == "DUPLICATE_SOURCE"       # consistent with the real matcher
    df = unused_source_records(e)
    assert list(df["Reason"]) == ["Duplicate in source file"] * 2 and list(df["Record in source file"]) == ["Kolar"] * 2


def test_unused_records_flags_missing_identifying_information():
    dest = destination()
    e = entry(dest, "p", "P.xlsx", ["District", "Population"], [["Kolar", 1], [None, 2]], {"Population": "Population"})
    df = unused_source_records(e)
    assert list(df["Reason"]) == ["Identifying information is missing"] and list(df["Row in source file"]) == [3]


def test_unused_records_needs_a_match_result():
    dest = destination()
    e = entry(dest, "p", "P.xlsx", ["District", "Population"], POP, {"Population": "Population"}, build=False)
    with pytest.raises(ValueError):
        unused_source_records(replace(e, match_result=None))