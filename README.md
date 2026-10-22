# Bosch-Excel-Integration

A public demonstration of converting product schedules and a bill of materials into component demand in Excel. All committed examples are fictional and generated from scratch. The repository name is retained as a historical project label; the examples do not describe a company's operations.

**Sanitized public demo:** This project has been adapted for public sharing through synthetic-data replacement and generalization of company-specific behavior. [Read what changed and why](#sanitization-for-public-sharing).

## Run

Use Python 3.11 or newer in a virtual environment:

```sh
python -m pip install -r requirements.txt
python Code/excel_transformer.py
```

This creates `outputs/plan.xlsx`, with **Detail**, **Totals**, and **Unmapped** sheets. The optional desktop interface is `python Code/custom_ui.py`; it requires a Python installation with Tk support.

The example CSVs are already included. Recreate them with `python generate_demo.py`. The generator reads no external files and uses a fixed random seed, invented identifiers, and an arbitrary calendar week.

## Input format

`schedule.csv`: `date`, `line`, `shift`, `product_id`, `quantity`.

`bill_of_materials.csv`: `product_id`, `component_id`, `units_per_product`.

The converter multiplies scheduled quantities by component usage and aggregates by date, line, shift, and component. Unmapped products appear separately instead of being silently dropped. Negative quantities, missing values, and duplicate mappings are rejected.

```sh
python Code/excel_transformer.py --input-dir examples --output outputs/plan.xlsx
```

## Sanitization for public sharing

This edition demonstrates the project's technical ideas while excluding operational records and personal information. The sample data is generated from scratch by `generate_demo.py`, without reading original records or deriving values from their statistics. Identifiers, quantities, dates, and relationships in the examples are invented.

The organization-specific workbook formats, shift policies, department labels, original master data, and screenshot manual have been replaced with this generic example. This edition does **not** support the original corporate workbook format or reproduce its operational rules. Generated workbooks and local inputs are excluded from Git.

Personal filesystem paths and identifying document metadata have been removed from the published files. The public demo starts with a fresh initial commit containing the sanitized edition. These changes are described as **sanitization and synthetic-data replacement**; the sample workbooks and results should be understood as educational examples.

Run the checks with `python -m pip install pytest` and `python -m pytest tests`.
