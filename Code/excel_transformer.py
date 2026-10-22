"""Command-line entry point for the generic public demonstration."""
import argparse
from pathlib import Path

from helper import transform

ROOT = Path(__file__).resolve().parents[1]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-dir", type=Path, default=ROOT / "examples")
    parser.add_argument("--output", type=Path, default=ROOT / "outputs" / "plan.xlsx")
    args = parser.parse_args()
    detail, totals, unmatched = transform(args.input_dir, args.output)
    print(f"Created workbook: {len(detail)} detail rows, {len(totals)} totals, "
          f"{len(unmatched)} unmapped schedule rows.")


if __name__ == "__main__":
    main()
