
import re
import unicodedata

import pandas as pd
from rapidfuzz import fuzz


def normalize_value(value):
    """Normalize text without removing meaningful leading zeros."""

    if pd.isna(value):
        return ""

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if float(value).is_integer():
            value = str(int(value))
        else:
            value = str(value)
    else:
        value = str(value)

    value = unicodedata.normalize("NFKC", value).casefold()

    if re.fullmatch(r"[+-]?\d+\.0+", value):
        value = value.split(".")[0]

    value = re.sub(r"[\u00a0\t\r\n]", " ", value)
    value = re.sub(r"[-_/(),]", " ", value)
    value = re.sub(r"[.']", "", value)
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def make_key(row, columns, normalized=False):
    """Create a composite key from selected columns."""

    values = []

    for column in columns:
        value = row[column]

        if pd.isna(value) or str(value).strip() == "":
            return None

        if normalized:
            value = normalize_value(value)

            if not value:
                return None
        else:
            value = str(value)

        values.append(value)

    return tuple(values)


def is_fuzzy_eligible(value):
    """Short and numeric-like values are exact-match only."""

    value = str(value).strip()

    if len(value) < 4:
        return False

    if re.fullmatch(r"[+-]?\d+(\.\d+)?", value):
        return False

    return True


def fuzzy_score(source_key, destination_key):
    """Use the weakest component score for composite keys."""

    scores = []

    for source_value, destination_value in zip(
        source_key, destination_key
    ):
        source_value = normalize_value(source_value)
        destination_value = normalize_value(destination_value)

        if (
            not is_fuzzy_eligible(source_value)
            or not is_fuzzy_eligible(destination_value)
        ):
            return None

        scores.append(
            fuzz.ratio(source_value, destination_value)
        )

    return min(scores) if scores else None


def run_match(
    src_df,
    dst_df,
    src_keys,
    dst_keys,
    threshold=90,
    margin=5
):
    """Match destination rows against source rows."""

    if not src_keys or not dst_keys:
        raise ValueError("At least one matching column is required.")

    if len(src_keys) != len(dst_keys):
        raise ValueError(
            "Source and destination key columns must have equal counts."
        )

    if not set(src_keys).issubset(src_df.columns):
        raise ValueError("Invalid source key column.")

    if not set(dst_keys).issubset(dst_df.columns):
        raise ValueError("Invalid destination key column.")

    results = []

    for dst_index, row in dst_df.iterrows():
        key = make_key(row, dst_keys)

        results.append({
            "_dst_index": dst_index,
            "_src_index": None,
            "_suggested_src_index": None,
            "match_score": None,
            "status": "INVALID_KEY" if key is None else "UNMATCHED"
        })

    # Build exact source lookup
    exact_lookup = {}

    for src_index, row in src_df.iterrows():
        key = make_key(row, src_keys)

        if key is not None:
            exact_lookup.setdefault(key, []).append(src_index)

    claimed_sources = set()

    # PASS 1: Exact matching
    for result in results:
        if result["status"] != "UNMATCHED":
            continue

        dst_row = dst_df.loc[result["_dst_index"]]
        key = make_key(dst_row, dst_keys)

        if key in exact_lookup:
            matches = exact_lookup[key]

            if len(matches) > 1:
                result["status"] = "DUPLICATE_SOURCE"

            elif matches[0] in claimed_sources:
                result["status"] = "DUPLICATE_DEST"

            else:
                result["_src_index"] = matches[0]
                result["status"] = "EXACT"
                claimed_sources.add(matches[0])

    # Build normalized source lookup
    normalized_lookup = {}

    for src_index, row in src_df.iterrows():
        key = make_key(row, src_keys, normalized=True)

        if key is not None:
            normalized_lookup.setdefault(key, []).append(src_index)

    # PASS 2: Normalized matching
    for result in results:
        if result["status"] != "UNMATCHED":
            continue

        dst_row = dst_df.loc[result["_dst_index"]]
        key = make_key(dst_row, dst_keys, normalized=True)

        matches = normalized_lookup.get(key, [])

        if len(matches) > 1:
            result["status"] = "DUPLICATE_SOURCE"

        elif len(matches) == 1:
            src_index = matches[0]

            if src_index in claimed_sources:
                result["status"] = "DUPLICATE_DEST"
            else:
                result["_src_index"] = src_index
                result["status"] = "NORMALIZED"
                claimed_sources.add(src_index)

    # PASS 3: Fuzzy suggestions
    # Include claimed sources so we can identify conflicts.
    source_candidates = {}

    for src_index, row in src_df.iterrows():
        key = make_key(row, src_keys, normalized=True)

        if key is not None:
            source_candidates[src_index] = key

    fuzzy_proposals = {}

    for result in results:
        if result["status"] != "UNMATCHED":
            continue

        dst_row = dst_df.loc[result["_dst_index"]]
        destination_key = make_key(
            dst_row, dst_keys, normalized=True
        )

        scored_candidates = []

        for src_index, source_key in source_candidates.items():
            score = fuzzy_score(source_key, destination_key)

            if score is not None:
                scored_candidates.append((score, src_index))

        scored_candidates.sort(reverse=True)

        if not scored_candidates:
            continue

        best_score, best_source = scored_candidates[0]

        second_score = (
            scored_candidates[1][0]
            if len(scored_candidates) > 1
            else 0
        )

        # Detect candidates that are too close to distinguish
        if best_score < threshold:
            if (
                len(scored_candidates) > 1
                and best_score >= 70
                and best_score - second_score < margin
            ):
                result["status"] = "AMBIGUOUS"
                result["match_score"] = best_score

            continue

        if best_score - second_score < margin:
            result["status"] = "AMBIGUOUS"
            result["match_score"] = best_score
            continue

        # A claimed source cannot be suggested again
        if best_source in claimed_sources:
            result["status"] = "AMBIGUOUS"
            result["match_score"] = best_score
            continue

        result["status"] = "FUZZY"
        result["_suggested_src_index"] = best_source
        result["match_score"] = best_score

        fuzzy_proposals.setdefault(best_source, []).append(result)

    # Detect multiple fuzzy proposals for the same source
    for source_index, proposals in fuzzy_proposals.items():
        if len(proposals) > 1:
            for result in proposals:
                result["status"] = "AMBIGUOUS"
                result["_suggested_src_index"] = None

    result_df = pd.DataFrame(results)

    summary = result_df["status"].value_counts().to_dict()

    return result_df, summary
