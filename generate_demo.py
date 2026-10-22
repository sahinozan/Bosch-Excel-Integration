"""Generate independent fictional examples; no external data is read."""
import csv
from datetime import date, timedelta
from pathlib import Path
import random


def generate(destination=None):
    destination = Path(destination or Path(__file__).parent / "examples")
    destination.mkdir(parents=True, exist_ok=True)
    rng = random.Random(42)
    with (destination / "schedule.csv").open("w", newline="") as stream:
        writer = csv.writer(stream, lineterminator="\n")
        writer.writerow(["date", "line", "shift", "product_id", "quantity"])
        for day in range(5):
            for product in range(1, 4):
                for shift in ["Morning", "Evening"]:
                    writer.writerow([(date(2020, 1, 6) + timedelta(days=day)).isoformat(),
                                     f"SYN-LINE-{product}", shift,
                                     f"SYN-PRODUCT-{product:02}", rng.randint(5, 20)])
    with (destination / "bill_of_materials.csv").open("w", newline="") as stream:
        writer = csv.writer(stream, lineterminator="\n")
        writer.writerow(["product_id", "component_id", "units_per_product"])
        for product in range(1, 4):
            writer.writerow([f"SYN-PRODUCT-{product:02}", f"SYN-PART-{product:02}", 2])
            writer.writerow([f"SYN-PRODUCT-{product:02}", "SYN-PART-04", 1])


if __name__ == "__main__":
    generate()
