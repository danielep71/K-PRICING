"""Illustrative derivation: the total notional amount of a fixture's records.

Format example for docs/methodology/TEST_CASES.md (Derivation scripts). It
reads the fixture named on the command line, sums notional_amount over every
record of every dataset, and prints the result in the derivation output
format. It implements no regulatory rule.
"""

import json
import sys
from decimal import Decimal


def main(path: str) -> None:
    with open(path, encoding="utf-8") as handle:
        fixture = json.load(handle, parse_float=Decimal, parse_int=Decimal)
    total = sum((record["inputs"]["notional_amount"]
                 for records in fixture["datasets"].values() for record in records), Decimal(0))
    print(json.dumps({"outputs": [{"quantity": "illustrative_total_notional_amount",
                                   "scope": None, "value": float(total)}]}))


if __name__ == "__main__":
    main(sys.argv[1])
