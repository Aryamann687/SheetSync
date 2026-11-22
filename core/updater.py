
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell


def is_blank(value):
    """Check whether a source or destination value is blank."""

    if value is None:
        return True

    if isinstance(value, str):
        return value.strip() == ""

    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def convert_value(value):
    """Convert Pandas and NumPy values into Excel-compatible values."""

    if is_blank(value):
        return None

    if hasattr(value, "item"):
        value = value.item()

    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()

    return value


def build_plan(
    dest_bytes,
    sheet,
    result_df,
    src_df,
    dst_df,
    mapping,
    strategy="fill_only_blank",
    accepted_rows=None
):
    """
    Build a preview of proposed Excel cell updates.

    mapping: {source_column: destination_column}
    accepted_rows: destination DataFrame indexes approved for fuzzy matching.

    Returns:
        plan: List of proposed cell updates.
        skips: List of skipped operations and reasons.
    """

    if strategy not in ("overwrite", "fill_only_blank"):
        raise ValueError("Invalid update strategy.")

    if not mapping:
        raise ValueError("At least one transfer column is required.")

    if accepted_rows is None:
        accepted_rows = set()

    if not set(mapping.keys()).issubset(src_df.columns):
        raise ValueError("Invalid source transfer column.")

    if not set(mapping.values()).issubset(dst_df.columns):
        raise ValueError("Invalid destination transfer column.")

    workbook = load_workbook(BytesIO(dest_bytes), data_only=False)

    if sheet not in workbook.sheetnames:
        raise ValueError("Destination worksheet not found.")

    worksheet = workbook[sheet]

    destination_columns = {
        column: index + 1
        for index, column in enumerate(dst_df.columns)
        if column != "_xl_row"
    }

    plan = []
    skips = []

    for _, match in result_df.iterrows():

        status = match["status"]
        dst_index = match["_dst_index"]

        # Exact and normalized matches are accepted automatically.
        if status in ("EXACT", "NORMALIZED"):
            src_index = match["_src_index"]

        # Fuzzy matches require explicit user approval.
        elif status == "FUZZY" and dst_index in accepted_rows:
            src_index = match["_suggested_src_index"]

        else:
            skips.append({
                "destination_index": dst_index,
                "reason": f"Match status: {status}"
            })
            continue

        if pd.isna(src_index):
            skips.append({
                "destination_index": dst_index,
                "reason": "Missing source row"
            })
            continue

        excel_row = int(dst_df.loc[dst_index, "_xl_row"])

        for source_column, destination_column in mapping.items():

            source_value = src_df.loc[src_index, source_column]

            if is_blank(source_value):
                skips.append({
                    "destination_index": dst_index,
                    "column": destination_column,
                    "reason": "Blank source value"
                })
                continue

            excel_column = destination_columns[destination_column]
            cell = worksheet.cell(
                row=excel_row,
                column=excel_column
            )

            if isinstance(cell, MergedCell):
                skips.append({
                    "destination_index": dst_index,
                    "column": destination_column,
                    "reason": "Merged cell is not writable"
                })
                continue

            if cell.data_type == "f":
                skips.append({
                    "destination_index": dst_index,
                    "column": destination_column,
                    "reason": "Destination contains a formula"
                })
                continue

            if strategy == "fill_only_blank" and not is_blank(cell.value):
                skips.append({
                    "destination_index": dst_index,
                    "column": destination_column,
                    "reason": "Destination already contains a value"
                })
                continue

            new_value = convert_value(source_value)

            if new_value is None:
                continue

            plan.append({
                "row": excel_row,
                "column": excel_column,
                "column_name": destination_column,
                "old_value": cell.value,
                "new_value": new_value,
                "action": "UPDATE"
            })

    return plan, skips


def apply_plan(dest_bytes, sheet, plan):
    """Apply an approved update plan and return the updated workbook bytes."""

    workbook = load_workbook(BytesIO(dest_bytes), data_only=False)

    if sheet not in workbook.sheetnames:
        raise ValueError("Destination worksheet not found.")

    worksheet = workbook[sheet]

    for item in plan:
        row = item["row"]
        column = item["column"]

        cell = worksheet.cell(row=row, column=column)

        if isinstance(cell, MergedCell):
            raise ValueError(
                f"Cannot write to merged cell at row {row}, column {column}."
            )

        if cell.data_type == "f":
            raise ValueError(
                f"Cannot overwrite formula at row {row}, column {column}."
            )

        cell.value = item["new_value"]

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return output.getvalue()
