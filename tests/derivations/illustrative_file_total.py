"""Illustrative derivation: the total notional of a file-based variant.

Format example for docs/methodology/TEST_CASES.md (Derivation scripts). It
reads positions.csv in the variant directory named on the command line, sums
the notional column, and prints the result in the derivation output format.
It implements no regulatory rule.
"""

import csv
import json
import sys
from decimal import Decimal
from pathlib import Path


def main(variant: str) -> None:
    # newline="" lets the csv module read LF and CRLF files alike.
    with open(Path(variant) / "positions.csv", encoding="utf-8", newline="") as handle:
        total = sum((Decimal(row["notional"]) for row in csv.DictReader(handle)), Decimal(0))
    print(json.dumps({"outputs": [{"quantity": "illustrative_total_notional_amount",
                                   "scope": None, "value": float(total)}]}))


if __name__ == "__main__":
    main(sys.argv[1])
