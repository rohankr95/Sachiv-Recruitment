"""
Converts an Excel export of the application sheet into the JSON format used by
local-test/dev-server.js and local-test/sheet-data-test.js.

    pip install openpyxl
    python local-test/xlsx-to-json.py  path/to/export.xlsx  path/outside/repo/sheet.json

Keep the output OUTSIDE the repository – it contains applicant data.
"""
import datetime
import json
import sys

import openpyxl


def main(src, dst):
    values_ws = openpyxl.load_workbook(src, data_only=True).worksheets[0]
    formula_ws = openpyxl.load_workbook(src).worksheets[0]

    def conv(x):
        if isinstance(x, datetime.datetime):
            return {"d": x.strftime("%Y-%m-%dT%H:%M:%S") + "+05:30"}
        if isinstance(x, datetime.date):
            return {"d": x.strftime("%Y-%m-%dT00:00:00") + "+05:30"}
        return "" if x is None else x

    values = [[conv(c) for c in row] for row in values_ws.iter_rows(values_only=True)]
    formulas = [[c if isinstance(c, str) and c.startswith("=") else "" for c in row]
                for row in formula_ws.iter_rows(values_only=True)]
    while values and all(c == "" for c in values[-1]):
        values.pop()
        formulas.pop()
    with open(dst, "w", encoding="utf-8") as f:
        json.dump({"sheetName": values_ws.title, "values": values, "formulas": formulas[:len(values)]}, f, ensure_ascii=False)
    print("Wrote %s: %d rows x %d columns (sheet '%s')" % (dst, len(values) - 1, len(values[0]), values_ws.title))


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])
