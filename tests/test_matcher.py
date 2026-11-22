
import pandas as pd
from core.matcher import run_match


def test_exact_matching():
    src = pd.DataFrame({
        "District": ["Kolar", "Bengaluru"],
        "Block": ["Kolar", "Anekal"]
    })

    dst = pd.DataFrame({
        "District": ["Kolar", "Bengaluru", "Mysuru"],
        "Block": ["Kolar", "Anekal", "Mysuru"]
    })

    result, summary = run_match(
        src, dst,
        ["District", "Block"],
        ["District", "Block"]
    )

    assert result["status"].tolist() == [
        "EXACT", "EXACT", "UNMATCHED"
    ]

    assert summary["EXACT"] == 2
    assert summary["UNMATCHED"] == 1


def test_invalid_key():
    src = pd.DataFrame({
        "District": ["Kolar"],
        "Block": ["Kolar"]
    })

    dst = pd.DataFrame({
        "District": ["Kolar"],
        "Block": [None]
    })

    result, summary = run_match(
        src, dst,
        ["District", "Block"],
        ["District", "Block"]
    )

    assert result.iloc[0]["status"] == "INVALID_KEY"


def test_duplicate_source_key():
    src = pd.DataFrame({
        "District": ["Kolar", "Kolar"],
        "Block": ["Kolar", "Kolar"]
    })

    dst = pd.DataFrame({
        "District": ["Kolar"],
        "Block": ["Kolar"]
    })

    result, summary = run_match(
        src, dst,
        ["District", "Block"],
        ["District", "Block"]
    )

    assert result.iloc[0]["status"] == "DUPLICATE_SOURCE"

def test_normalized_matching():
    src = pd.DataFrame({
        "District": ["Kolar", "New-York", "Block_1"],
        "Block": ["Main", "Central", "North"]
    })

    dst = pd.DataFrame({
        "District": ["KOLAR", "New York", "Block 1"],
        "Block": ["main", "CENTRAL", "north"]
    })

    result, summary = run_match(
        src, dst,
        ["District", "Block"],
        ["District", "Block"]
    )

    assert result["status"].tolist() == [
        "NORMALIZED", "NORMALIZED", "NORMALIZED"
    ]


def test_leading_zeros_preserved():
    src = pd.DataFrame({"ID": ["00101"]})
    dst = pd.DataFrame({"ID": ["101"]})

    result, summary = run_match(
        src, dst, ["ID"], ["ID"]
    )

    assert result.iloc[0]["status"] == "UNMATCHED"


def test_numeric_float_normalization():
    src = pd.DataFrame({"ID": [101.0]})
    dst = pd.DataFrame({"ID": ["101"]})

    result, summary = run_match(
        src, dst, ["ID"], ["ID"]
    )

    assert result.iloc[0]["status"] == "NORMALIZED"

def test_fuzzy_matching():
    src = pd.DataFrame({
        "District": ["Kolar"]
    })

    dst = pd.DataFrame({
        "District": ["Kolaar"]
    })

    result, summary = run_match(
        src, dst, ["District"], ["District"]
    )

    assert result.iloc[0]["status"] == "FUZZY"
    assert result.iloc[0]["_suggested_src_index"] == 0
    assert result.iloc[0]["match_score"] >= 90

    # Fuzzy suggestions are NOT automatically accepted
    assert result.iloc[0]["_src_index"] is None


def test_fuzzy_ambiguous_candidates():
    src = pd.DataFrame({
        "District": ["Kolar", "Kolas"]
    })

    dst = pd.DataFrame({
        "District": ["Kola"]
    })

    result, summary = run_match(
        src, dst, ["District"], ["District"]
    )

    assert result.iloc[0]["status"] == "AMBIGUOUS"


def test_short_values_are_exact_only():
    src = pd.DataFrame({"ID": ["ABC"]})
    dst = pd.DataFrame({"ID": ["ABD"]})

    result, summary = run_match(
        src, dst, ["ID"], ["ID"]
    )

    assert result.iloc[0]["status"] == "UNMATCHED"


def test_fuzzy_source_collision():
    src = pd.DataFrame({
        "District": ["Kolar"]
    })

    dst = pd.DataFrame({
        "District": ["Kolaar", "Kolar "]
    })

    result, summary = run_match(
        src, dst, ["District"], ["District"]
    )

    assert result["status"].tolist() == [
        "AMBIGUOUS", "NORMALIZED"
    ]
