
from io import BytesIO

import pandas as pd
from openpyxl import Workbook, load_workbook

from core.updater import build_plan, apply_plan


def create_workbook():
    """Create a sample destination workbook."""

    wb = Workbook()
    ws = wb.active
    ws.title = "Data"

    ws.append(["District", "Population", "Growth"])
    ws.append(["Kolar", None, "=10+20"])
    ws.append(["Bengaluru", 500, None])

    output = BytesIO()
    wb.save(output)

    return output.getvalue()


def create_dataframes():
    src_df = pd.DataFrame({
        "District": ["Kolar", "Bengaluru"],
        "New Population": [1000, 2000]
    })

    dst_df = pd.DataFrame({
        "District": ["Kolar", "Bengaluru"],
        "Population": [None, 500],
        "Growth": [None, None],
        "_xl_row": [2, 3]
    })

    result_df = pd.DataFrame({
        "_dst_index": [0, 1],
        "_src_index": [0, 1],
        "_suggested_src_index": [None, None],
        "status": ["EXACT", "EXACT"]
    })

    return src_df, dst_df, result_df


def test_fill_only_blank():
    src, dst, result = create_dataframes()

    plan, skips = build_plan(
        create_workbook(),
        "Data",
        result,
        src,
        dst,
        {"New Population": "Population"},
        strategy="fill_only_blank"
    )

    assert len(plan) == 1
    assert plan[0]["new_value"] == 1000


def test_overwrite_strategy():
    src, dst, result = create_dataframes()

    plan, skips = build_plan(
        create_workbook(),
        "Data",
        result,
        src,
        dst,
        {"New Population": "Population"},
        strategy="overwrite"
    )

    assert len(plan) == 2


def test_formula_protection():
    src, dst, result = create_dataframes()

    plan, skips = build_plan(
        create_workbook(),
        "Data",
        result,
        src,
        dst,
        {"New Population": "Growth"},
        strategy="overwrite"
    )

    assert len(plan) == 1
    assert plan[0]["row"] == 3


def test_apply_plan():
    src, dst, result = create_dataframes()

    original = create_workbook()

    plan, skips = build_plan(
        original,
        "Data",
        result,
        src,
        dst,
        {"New Population": "Population"},
        strategy="fill_only_blank"
    )

    updated = apply_plan(original, "Data", plan)

    wb = load_workbook(BytesIO(updated), data_only=False)
    ws = wb["Data"]

    assert ws["B2"].value == 1000
    assert ws["B3"].value == 500
    assert ws["C2"].value == "=10+20"
