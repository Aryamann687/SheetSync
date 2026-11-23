"""Combine several source files into one update plan for one destination workbook.

Design rules (see the architecture review):

* Every source is matched against the destination exactly as uploaded. Sources never see
  values that another source would fill, so the order of sources never changes a result.
* Sources interact in one place only: `consolidate()` merges their cell updates and detects
  collisions. Nothing here writes a cell; updates come from `core.updater.build_plan` and are
  applied by `core.updater.apply_plan`, called exactly once, from `apply_consolidated()`.
* A source is either in the output or explicitly excluded by the user (`SourceSpec.excluded`).
  Anything else (error, incomplete, stale, mismatched destination) is a blocker, never a silent
  omission.

This module has no Streamlit imports and does not modify core.reader, core.matcher or
core.updater. The updater contracts it relies on (verified against the implementation):

    build_plan(...) -> (plan: list[dict], skips: list[dict])
        plan item  : row, column, column_name, old_value, new_value, action
        skip record: destination_index, reason, and optionally column
    apply_plan(dest_bytes, sheet, plan) -> bytes   (reads row, column, new_value from each item)
"""

from __future__ import annotations

import hashlib
from collections import Counter
from dataclasses import dataclass
from typing import Iterable, Mapping, Sequence

import pandas as pd

from core.updater import apply_plan, build_plan

HIDDEN = "_xl_row"  # internal column added by core.reader.load_table
STRATEGIES = ("fill_only_blank", "overwrite")
MATCHED_STATUSES = ("EXACT", "NORMALIZED")

# matcher status code -> (result label, detail wording); shared by the report and the app
STATUS_LABELS = {
    "EXACT": ("Matched", "Exact match"),
    "NORMALIZED": ("Matched", "Match after ignoring formatting differences"),
    "FUZZY": ("Needs review", "Possible match"),
    "AMBIGUOUS": ("Unclear", "Several similar records found"),
    "DUPLICATE_SOURCE": ("Duplicate", "Duplicate in source file"),
    "DUPLICATE_DEST": ("Duplicate", "Duplicate in destination file"),
    "INVALID_KEY": ("Invalid information", "Identifying information is missing"),
    "UNMATCHED": ("Not found", "No match found"),
}
RESULT_ORDER = ("Matched", "Needs review", "Unclear", "Not found", "Duplicate", "Invalid information")

REPORT_COLUMNS = (
    "Row in destination file",
    "Record in destination file",
    "Result",
    "Details",
    "Row in source file",
    "Record in source file",
    "Similarity (0-100)",
    "Your decision",
)
UNUSED_COLUMNS = ("Row in source file", "Record in source file", "Reason")
CONFLICT_COLUMNS = ("Row", "Column", "Current value", "Source file", "Proposed value")


# ---------------------------------------------------------------- data types
@dataclass(frozen=True)
class MatchSettings:
    """Opaque matching numbers passed to core.matcher.run_match by the caller."""
    threshold: int
    margin: int


@dataclass(frozen=True)
class SourceSpec:
    """What the user configured for one source file."""
    source_id: str
    label: str                      # filename shown to users and written to the report
    sheet: str
    header_row: int
    src_keys: tuple
    dst_keys: tuple
    settings: MatchSettings
    mapping: Mapping                # source column -> destination column
    strategy: str = "fill_only_blank"
    excluded: bool = False          # an explicit user decision to leave this source out

    def __post_init__(self):
        object.__setattr__(self, "src_keys", tuple(self.src_keys))
        object.__setattr__(self, "dst_keys", tuple(self.dst_keys))
        object.__setattr__(self, "mapping", dict(self.mapping))


@dataclass(frozen=True, eq=False)
class Destination:
    """The destination workbook exactly as uploaded, plus its loaded table."""
    file_bytes: bytes
    sheet: str
    header_row: int
    table: pd.DataFrame
    signature: str


@dataclass(frozen=True)
class SourcePlan:
    """One source's updates and skips, as returned by build_plan, stamped with the destination."""
    source_id: str
    dest_signature: str
    updates: tuple
    skips: tuple


@dataclass(frozen=True, eq=False)
class SourceEntry:
    """Everything known about one source at this moment."""
    spec: SourceSpec
    table: "pd.DataFrame | None" = None         # loaded source table (with _xl_row)
    match_result: "pd.DataFrame | None" = None  # result_df from run_match
    approved_rows: frozenset = frozenset()      # destination indexes of confirmed possible matches
    plan: "SourcePlan | None" = None
    problem: "str | None" = None                # plain-language reason this source cannot be used


@dataclass(frozen=True)
class Issue:
    code: str
    severity: str                   # "error" blocks processing; "warning" does not
    source_id: "str | None"
    message: str
    columns: tuple = ()


@dataclass(frozen=True)
class Blocker:
    source_id: str
    label: str
    reason: str


@dataclass(frozen=True)
class ExcludedSource:
    source_id: str
    label: str


@dataclass(frozen=True)
class MergedUpdate:
    row: int
    column: int
    column_name: str
    old_value: object
    new_value: object
    action: str
    source_ids: tuple               # sources that propose this value, in slot order


@dataclass(frozen=True)
class ConflictCandidate:
    source_id: str
    label: str
    new_value: object


@dataclass(frozen=True)
class Conflict:
    row: int
    column: int
    column_name: str
    old_value: object
    candidates: tuple


@dataclass(frozen=True)
class SourceSummary:
    source_id: str
    label: str
    results: dict                   # result label -> number of destination records
    needs_review: int
    confirmed: int
    unconfirmed: int
    updates: int
    skips: int


@dataclass(frozen=True)
class Consolidation:
    dest_signature: str
    included: tuple                 # source ids with a current plan, in slot order
    excluded: tuple                 # ExcludedSource, explicit user choice
    blockers: tuple
    updates: tuple                  # MergedUpdate, sorted by (row, column)
    conflicts: tuple
    summaries: tuple                # SourceSummary for every non-excluded source

    @property
    def can_apply(self) -> bool:
        return not self.blockers and not self.conflicts and bool(self.updates)


class ConsolidationBlocked(Exception):
    """Raised when applying is refused. The message is plain language."""

    def __init__(self, message, blockers=(), conflicts=()):
        super().__init__(message)
        self.blockers = tuple(blockers)
        self.conflicts = tuple(conflicts)


# ---------------------------------------------------------------- small helpers
def _missing(idx) -> bool:
    return idx is None or bool(pd.isna(idx))


def _xl_row(df, idx):
    """Excel row number for a DataFrame index label; blank when there is none."""
    if _missing(idx):
        return ""
    try:
        return int(df.loc[idx, HIDDEN])
    except (KeyError, TypeError, ValueError):
        pass
    try:
        return int(df.loc[int(idx), HIDDEN])
    except (KeyError, TypeError, ValueError):
        return ""


def _text(df, idx, cols):
    """Readable key text for one row; blank if the index is missing."""
    if _missing(idx):
        return ""
    try:
        row = df.loc[idx, list(cols)]
    except (KeyError, TypeError):
        try:
            row = df.loc[int(idx), list(cols)]
        except Exception:
            return ""
    return " / ".join("" if pd.isna(v) else str(v) for v in row)


def _kind(value):
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, (int, float)):
        return "number"
    return type(value)


def _same_value(a, b) -> bool:
    """Equal values: 5 and 5.0 are the same, but 5 and "5", or True and 1, are not."""
    return _kind(a) == _kind(b) and a == b


def _plain_update(u: "MergedUpdate") -> dict:
    """The exact six-key dict shape core.updater uses."""
    return {"row": u.row, "column": u.column, "column_name": u.column_name,
            "old_value": u.old_value, "new_value": u.new_value, "action": u.action}


def _result_counts(match_result):
    counts = {label: 0 for label in RESULT_ORDER}
    if match_result is not None:
        for status in match_result["status"]:
            counts[STATUS_LABELS[status][0]] += 1
    return counts


def _columns(table):
    return [c for c in table.columns if c != HIDDEN]


# ---------------------------------------------------------------- destination
def destination_signature(file_bytes, sheet, header_row) -> str:
    digest = hashlib.sha256(file_bytes).hexdigest()
    return hashlib.sha256(f"{digest}|{sheet}|{header_row}".encode("utf-8")).hexdigest()


def make_destination(file_bytes, sheet, header_row, table) -> Destination:
    """Bundle the destination bytes, sheet, header row and loaded table with a signature."""
    if not isinstance(file_bytes, (bytes, bytearray)):
        raise TypeError("file_bytes must be bytes.")
    if HIDDEN not in table.columns:
        raise ValueError("The destination table must come from core.reader.load_table.")
    file_bytes = bytes(file_bytes)
    return Destination(file_bytes, sheet, int(header_row), table,
                       destination_signature(file_bytes, sheet, int(header_row)))


# ---------------------------------------------------------------- validation
def _spec_issues(spec, dest_cols, src_cols):
    """Checks for one source. `src_cols` is None when its columns are unknown (skips those checks)."""
    out, label = [], spec.label

    def add(code, message, columns=()):
        out.append(Issue(code, "error", spec.source_id, message, tuple(columns)))

    if not spec.src_keys or not spec.dst_keys:
        add("NO_KEYS", f"{label}: choose the columns that identify a record, on both sides.")
    elif len(spec.src_keys) != len(spec.dst_keys):
        add("KEY_COUNT_MISMATCH", f"{label}: select the same number of matching columns on both sides.")

    if src_cols is not None:
        bad = [k for k in spec.src_keys if k not in src_cols]
        if bad:
            add("UNKNOWN_KEY_COLUMN", f"{label}: the matching column {', '.join(bad)} is not in this file.", bad)
    bad = [k for k in spec.dst_keys if k not in dest_cols]
    if bad:
        add("UNKNOWN_KEY_COLUMN",
            f"{label}: the matching column {', '.join(bad)} was not found in the destination workbook.", bad)

    if not spec.mapping:
        add("NO_MAPPING", f"{label}: choose at least one column to copy.")
    if src_cols is not None:
        bad = [c for c in spec.mapping if c not in src_cols]
        if bad:
            add("UNKNOWN_SOURCE_COLUMN", f"{label}: the column {', '.join(bad)} is not in this file.", bad)
    bad = [c for c in spec.mapping.values() if c not in dest_cols]
    if bad:
        add("UNKNOWN_DEST_COLUMN",
            f"{label}: the destination column {', '.join(bad)} was not found in the destination workbook.", bad)

    targets = list(spec.mapping.values())
    dupes = sorted({t for t in targets if targets.count(t) > 1}, key=str)
    if dupes:
        add("DUPLICATE_TARGET",
            f"{label}: more than one column is being copied into {', '.join(map(str, dupes))}. "
            "Each destination column can receive only one column.", dupes)

    if spec.strategy not in STRATEGIES:
        add("INVALID_STRATEGY", f"{label}: choose whether to fill empty cells only or replace existing values.")
    return out


def _cross_issues(active):
    """Checks that need every active source: key-column protection and shared targets."""
    out = []
    key_users = {}  # destination column -> labels of sources that use it to identify records
    for s in active:
        for k in s.dst_keys:
            key_users.setdefault(k, []).append(s.label)

    for s in active:
        hit = [c for c in dict.fromkeys(s.mapping.values()) if c in key_users]
        for c in hit:
            out.append(Issue(
                "TARGET_IS_KEY", "error", s.source_id,
                f"{s.label}: {c} can't be filled because it is used to identify records "
                f"({', '.join(dict.fromkeys(key_users[c]))}). Choose a different destination column.",
                (c,)))

    users = {}
    for s in active:
        for c in dict.fromkeys(s.mapping.values()):
            users.setdefault(c, []).append(s.label)
    for c, labels in users.items():
        if len(labels) > 1:
            out.append(Issue(
                "SHARED_TARGET", "warning", None,
                f"{', '.join(labels)} are all set to fill {c}. That is fine if they cover different "
                "records; if two files give a different value for the same cell you will be asked to resolve it.",
                (c,)))
    return out


def validate_setup(specs: Sequence[SourceSpec], dest_columns: Iterable[str],
                   source_columns: Mapping[str, Sequence[str]]) -> list:
    """Check every source's setup before any matching. Excluded sources are skipped.

    `source_columns` maps source_id -> that source's column names (without _xl_row).
    Returns a list of Issue; any with severity "error" must block matching and review.
    """
    dest_cols = [c for c in dest_columns if c != HIDDEN]
    issues = []
    ids = [s.source_id for s in specs]
    for sid in sorted({i for i in ids if ids.count(i) > 1}):
        issues.append(Issue("DUPLICATE_SOURCE_ID", "error", sid, f"Two files share the internal id {sid}."))
    active = [s for s in specs if not s.excluded]
    for s in active:
        cols = source_columns.get(s.source_id)
        if cols is None:
            issues.append(Issue("SOURCE_COLUMNS_UNAVAILABLE", "error", s.source_id,
                                f"{s.label}: we couldn't read the columns of this file."))
        issues.extend(_spec_issues(s, dest_cols, None if cols is None else [c for c in cols if c != HIDDEN]))
    issues.extend(_cross_issues(active))
    return issues


# ---------------------------------------------------------------- per-source plan
def build_source_plan(destination: Destination, entry: SourceEntry) -> SourcePlan:
    """Call core.updater.build_plan for one source against the original destination.

    Raises whatever build_plan raises (the caller records it as that source's `problem`), plus
    ValueError for an excluded source or one without a table or match result.
    """
    spec = entry.spec
    if spec.excluded:
        raise ValueError(f"{spec.label} is excluded; no plan can be built for it.")
    if entry.table is None or entry.match_result is None:
        raise ValueError(f"{spec.label} has not been matched yet.")

    out = build_plan(
        destination.file_bytes, destination.sheet, entry.match_result, entry.table,
        destination.table, dict(spec.mapping), strategy=spec.strategy,
        accepted_rows=set(entry.approved_rows),
    )
    if not (isinstance(out, tuple) and len(out) == 2
            and isinstance(out[0], list) and isinstance(out[1], list)):
        raise TypeError("build_plan returned an unexpected structure; expected (plan, skips) lists.")
    plan, skips = out
    return SourcePlan(spec.source_id, destination.signature,
                      tuple(dict(u) for u in plan), tuple(dict(s) for s in skips))


# ---------------------------------------------------------------- consolidation
def consolidate(destination: Destination, entries: Sequence[SourceEntry]) -> Consolidation:
    """Merge the plans of all included sources and report everything that blocks applying."""
    ids = [e.spec.source_id for e in entries]
    if len(set(ids)) != len(ids):
        raise ValueError("Source ids must be unique.")

    dest_cols = _columns(destination.table)
    active = [e for e in entries if not e.spec.excluded]
    excluded = tuple(ExcludedSource(e.spec.source_id, e.spec.label) for e in entries if e.spec.excluded)

    # setup errors are re-checked here so a forgotten validate_setup() cannot let them through
    setup_errors = {}
    for e in active:
        src_cols = None if e.table is None else _columns(e.table)
        for issue in _spec_issues(e.spec, dest_cols, src_cols):
            setup_errors.setdefault(e.spec.source_id, []).append(issue.message)
    for issue in _cross_issues([e.spec for e in active]):
        if issue.severity == "error":
            setup_errors.setdefault(issue.source_id, []).append(issue.message)

    blockers, usable = [], []
    for e in active:
        sid, label = e.spec.source_id, e.spec.label
        reasons = list(setup_errors.get(sid, []))
        if e.problem:
            reasons.append(e.problem)
        elif e.plan is None:
            reasons.append("Changes for this file have not been reviewed yet.")
        elif e.plan.source_id != sid:
            reasons.append("The reviewed changes belong to a different file. Review your changes again.")
        elif e.plan.dest_signature != destination.signature:
            reasons.append("This file was matched against a different version of the destination "
                           "workbook. Find matching records again.")
        if reasons:
            blockers.extend(Blocker(sid, label, r) for r in reasons)
        else:
            usable.append(e)

    cells = {}
    for e in usable:
        for u in e.plan.updates:
            cells.setdefault((u["row"], u["column"]), []).append((e.spec, u))

    merged, conflicts = [], []
    for key in sorted(cells):
        group = cells[key]
        first = group[0][1]
        if not all(_same_value(first["old_value"], u["old_value"]) and first["action"] == u["action"]
                   for _, u in group[1:]):
            # Every plan reads the same original workbook, so plans for one cell always agree on its
            # current value and action (even when the sources use different existing-value strategies).
            # If they do not, a plan was not built from this destination: hold the cell back instead of
            # silently keeping whichever source happens to come first.
            blockers.extend(Blocker(
                s.source_id, s.label,
                f"The change for row {first['row']}, column {first['column_name']} was prepared from a "
                "different version of the destination workbook than another file's. "
                "Find matching records and review your changes again.") for s, _ in group)
        elif all(_same_value(first["new_value"], u["new_value"]) for _, u in group[1:]):
            merged.append(MergedUpdate(first["row"], first["column"], first["column_name"],
                                       first["old_value"], first["new_value"], first["action"],
                                       tuple(s.source_id for s, _ in group)))
        else:
            conflicts.append(Conflict(
                first["row"], first["column"], first["column_name"], first["old_value"],
                tuple(ConflictCandidate(s.source_id, s.label, u["new_value"]) for s, u in group)))

    summaries = []
    for e in active:
        res = e.match_result
        fuzzy = set(res.loc[res["status"] == "FUZZY", "_dst_index"]) if res is not None else set()
        confirmed = len(fuzzy & set(e.approved_rows))
        summaries.append(SourceSummary(
            e.spec.source_id, e.spec.label, _result_counts(res), len(fuzzy), confirmed,
            len(fuzzy) - confirmed,
            len(e.plan.updates) if e.plan is not None else 0,
            len(e.plan.skips) if e.plan is not None else 0))

    return Consolidation(destination.signature, tuple(e.spec.source_id for e in usable), excluded,
                         tuple(blockers), tuple(merged), tuple(conflicts), tuple(summaries))


def conflict_table(consolidation: Consolidation) -> pd.DataFrame:
    """One row per source per conflicting cell, for the conflict report."""
    rows = []
    for c in consolidation.conflicts:
        current = "(empty)" if c.old_value is None else c.old_value
        for cand in c.candidates:
            rows.append({"Row": c.row, "Column": c.column_name, "Current value": current,
                         "Source file": cand.label, "Proposed value": cand.new_value})
    return pd.DataFrame(rows, columns=list(CONFLICT_COLUMNS))


def _blocked_message(consolidation):
    if consolidation.blockers:
        names = ", ".join(dict.fromkeys(b.label for b in consolidation.blockers))
        return f"Some files still need attention before the workbook can be created: {names}."
    if consolidation.conflicts:
        n = len(consolidation.conflicts)
        return (f"{n} cell{'s' if n != 1 else ''} would receive different values from different files. "
                "Resolve the conflicts first.")
    return "There are no changes to apply."


def apply_consolidated(destination: Destination, consolidation: Consolidation) -> bytes:
    """Apply the merged updates to the original destination with ONE call to apply_plan.

    Refuses (ConsolidationBlocked) unless the consolidation was built for this exact destination
    and has no blockers, no conflicts and at least one update.
    """
    if consolidation.dest_signature != destination.signature:
        raise ConsolidationBlocked("These changes were built for a different version of the destination "
                                   "workbook. Review your changes again.")
    if not consolidation.can_apply:
        raise ConsolidationBlocked(_blocked_message(consolidation),
                                   consolidation.blockers, consolidation.conflicts)
    plain = [_plain_update(u) for u in consolidation.updates]
    out = apply_plan(destination.file_bytes, destination.sheet, plain)
    if not isinstance(out, (bytes, bytearray)):
        raise TypeError("apply_plan returned an unexpected type; expected bytes.")
    return bytes(out)


# ---------------------------------------------------------------- reports
def _report_rows(destination, entry):
    spec, res = entry.spec, entry.match_result
    accepted = set(entry.approved_rows)
    for _, r in res.iterrows():
        status = r["status"]
        dst_index = r["_dst_index"]
        src_index = r["_src_index"] if status in MATCHED_STATUSES else r["_suggested_src_index"]
        score = r["match_score"]
        if status == "FUZZY":
            decision = "Confirmed" if dst_index in accepted else "Not confirmed"
        else:
            decision = "Not needed"
        yield {
            "Row in destination file": _xl_row(destination.table, dst_index),
            "Record in destination file": _text(destination.table, dst_index, spec.dst_keys),
            "Result": STATUS_LABELS[status][0],
            "Details": STATUS_LABELS[status][1],
            "Row in source file": _xl_row(entry.table, src_index),
            "Record in source file": _text(entry.table, src_index, spec.src_keys),
            "Similarity (0-100)": "" if pd.isna(score) else round(float(score), 1),
            "Your decision": decision,
        }


def matching_report_csv(destination: Destination, entries: Sequence[SourceEntry]) -> bytes:
    """One consolidated CSV: every destination record for every non-excluded matched source.

    The "Source file" column is included only when there is more than one source slot. With a
    single source the output has exactly the single-source report's columns and values.
    """
    with_source = len(entries) > 1
    columns = (["Source file"] if with_source else []) + list(REPORT_COLUMNS)
    rows = []
    for e in entries:
        if e.spec.excluded or e.match_result is None or e.table is None:
            continue
        for row in _report_rows(destination, e):
            rows.append({"Source file": e.spec.label, **row} if with_source else row)
    return pd.DataFrame(rows, columns=columns).to_csv(index=False).encode("utf-8-sig")


def unused_source_records(entry: SourceEntry) -> pd.DataFrame:
    """Source records that did not end up in any match, with the reason.

    Reasons: identifying information missing; duplicate in source file; possible match not
    confirmed; no matching destination record.

    Known limitation: "duplicate" means another source record has the same identifying
    information, either exactly or after core.matcher's formatting normalisation. This is
    detected with core.matcher.make_key (read-only), the same function the matcher uses, but it
    is a source-wide view: it does not say which destination record the duplicate collided with.
    """
    if entry.table is None or entry.match_result is None:
        raise ValueError("The file must be matched before unused records can be listed.")
    from core.matcher import make_key  # read-only reuse of the matcher's own key definition

    spec, table, res = entry.spec, entry.table, entry.match_result
    approved = set(entry.approved_rows)
    used, suggested = set(), set()
    for _, r in res.iterrows():
        if r["status"] in MATCHED_STATUSES and not _missing(r["_src_index"]):
            used.add(r["_src_index"])
        elif r["status"] == "FUZZY" and not _missing(r["_suggested_src_index"]):
            (used if r["_dst_index"] in approved else suggested).add(r["_suggested_src_index"])

    raw, norm = {}, {}
    for idx, row in table.iterrows():
        raw[idx] = make_key(row, spec.src_keys)
        norm[idx] = make_key(row, spec.src_keys, normalized=True)
    raw_n = Counter(k for k in raw.values() if k is not None)
    norm_n = Counter(k for k in norm.values() if k is not None)

    rows = []
    for idx in table.index:
        if idx in used:
            continue
        if raw[idx] is None:
            reason = "Identifying information is missing"
        elif raw_n.get(raw[idx], 0) > 1 or (norm[idx] is not None and norm_n.get(norm[idx], 0) > 1):
            reason = "Duplicate in source file"
        elif idx in suggested:
            reason = "Possible match not confirmed"
        else:
            reason = "No matching destination record"
        rows.append({"Row in source file": _xl_row(table, idx),
                     "Record in source file": _text(table, idx, spec.src_keys),
                     "Reason": reason})
    return pd.DataFrame(rows, columns=list(UNUSED_COLUMNS))