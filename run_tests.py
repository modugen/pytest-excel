import os

import pandas as pd
import subprocess

DATA_PATH = "test/data"
EXPECTED_RESULT_FILE = os.path.join(DATA_PATH, "expected_report.xls")
RESULT_FILE = os.path.join(DATA_PATH, "report.xls")

proc = subprocess.call(
    ["pytest", f"--excelreport={RESULT_FILE}", "test"],
)
print(proc)

result = pd.read_excel(RESULT_FILE)
expected = pd.read_excel(EXPECTED_RESULT_FILE)

assert (result == expected).all().all()
