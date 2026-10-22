# Bosch Senior Project — Excel Integration

This repository contains the Excel-processing work from a bachelor's senior project completed as a team in collaboration with Bosch. The aim was to turn production-planning spreadsheets into component requirements and reduce repetitive spreadsheet work.

The original implementation was tailored to the company's data and workbook formats. This repository is shared as an academic portfolio project, with a simplified example that can be explored independently.

## What the project covered

- Reading production schedules and product-to-component mappings.
- Calculating component quantities and organizing them by production line and shift.
- Exporting formatted Excel reports and providing a desktop interface for selecting files.

**Technologies:** Python, pandas, NumPy, openpyxl, and a desktop GUI.

## Sanitization for public sharing

The public version uses independently generated synthetic data. Original business data, company-specific workbook layouts, and operational rules have been replaced with fictional examples and generic logic. Personal paths, identifying document metadata, and the original screenshot manual are omitted from this version.

The current demo reads a CSV schedule and bill of materials, then exports an Excel workbook. It does not reproduce the original corporate workflow. The examples can be regenerated with `python generate_demo.py`; the generator does not read original records or use their statistics.

## Try the example

With Python 3.11 or newer, run these commands from the repository directory, preferably in a virtual environment:

```sh
python -m pip install -r requirements.txt
python Code/excel_transformer.py
```

The result is `outputs/plan.xlsx`, containing detail, aggregated totals, and any products without a component mapping. The optional desktop interface is available through `python Code/custom_ui.py` and requires Tk support.

## Related senior-project repositories

- [Forecasting](https://github.com/sahinozan/Bosch-Forecasting) — time-series analysis and forecasting experiments.
- [Reinforcement Learning](https://github.com/sahinozan/Bosch-Reinforcement-Learning) — a custom inventory environment and reinforcement learning experiments.
