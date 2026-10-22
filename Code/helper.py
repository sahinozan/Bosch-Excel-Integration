"""Generic schedule-to-component conversion for the public demo edition.

Replaces the original organization-specific workbook parser and rules.
"""
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.packaging.core import DocumentProperties
from openpyxl.styles import Alignment, Font, PatternFill

SCHEDULE_COLUMNS = ["date", "line", "shift", "product_id", "quantity"]
BOM_COLUMNS = ["product_id", "component_id", "units_per_product"]


def _validate(frame, columns, numeric_column):
    missing = set(columns) - set(frame.columns)
    if missing:
        raise ValueError(f"Missing columns: {', '.join(sorted(missing))}")
    frame = frame.loc[:, columns].copy()
    if frame.empty or frame.isna().any().any():
        raise ValueError("Input must contain rows with no missing values")
    for column in set(columns) - {numeric_column, "date"}:
        frame[column] = frame[column].astype(str).str.strip()
        if (frame[column] == "").any():
            raise ValueError(f"Empty identifier in {column}")
    values = pd.to_numeric(frame[numeric_column], errors="raise")
    if not np.isfinite(values).all() or (values < 0).any():
        raise ValueError(f"{numeric_column} must be finite and nonnegative")
    frame[numeric_column] = values
    return frame


def convert_schedule(schedule, bill_of_materials):
    """Expand product quantities into component demand, preserving missing mappings."""
    schedule = _validate(schedule, SCHEDULE_COLUMNS, "quantity")
    bom = _validate(bill_of_materials, BOM_COLUMNS, "units_per_product")
    schedule["date"] = pd.to_datetime(schedule["date"], errors="raise")
    if bom.duplicated(["product_id", "component_id"]).any():
        raise ValueError("Duplicate product/component mapping")
    if (bom["units_per_product"] <= 0).any():
        raise ValueError("units_per_product must be positive")

    unmatched = schedule.loc[~schedule["product_id"].isin(bom["product_id"])].copy()
    detail = schedule.merge(bom, on="product_id", how="inner")
    detail["component_quantity"] = detail["quantity"] * detail["units_per_product"]
    totals = detail.groupby(
        ["date", "line", "shift", "component_id"], as_index=False
    )["component_quantity"].sum()
    return detail, totals, unmatched


def write_to_excel(output, detail, totals, unmatched):
    """Write newly constructed sheets; never copy source workbook metadata."""
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, frame in [("Detail", detail), ("Totals", totals), ("Unmapped", unmatched)]:
            frame.to_excel(writer, sheet_name=name, index=False)
    workbook = load_workbook(output)
    workbook.properties = DocumentProperties(
        creator="Public demo", lastModifiedBy="Public demo", title="Planning export"
    )
    for sheet in workbook:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="245D63")
        for column in sheet.columns:
            sheet.column_dimensions[column[0].column_letter].width = 22
            for cell in column:
                cell.alignment = Alignment(vertical="center")
                if cell.data_type == "f":
                    # Inputs are values, not spreadsheet instructions.
                    cell.data_type = "s"
                if cell.column == 1 and cell.row > 1:
                    cell.number_format = "yyyy-mm-dd"
    workbook.save(output)


def transform(input_dir, output):
    input_dir = Path(input_dir)
    schedule = pd.read_csv(input_dir / "schedule.csv")
    bom = pd.read_csv(input_dir / "bill_of_materials.csv")
    frames = convert_schedule(schedule, bom)
    write_to_excel(output, *frames)
    return frames
