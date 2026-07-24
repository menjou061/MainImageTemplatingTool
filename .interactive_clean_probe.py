from pathlib import Path
import sys

sys.path.insert(0, r"C:\Users\Public\L0_BatchTool\_internal")
from clean_data import build_data
from openpyxl import load_workbook


downloads = Path(r"C:\Users\123\Downloads")
workbooks = [
    path
    for path in downloads.rglob("*.xlsx")
    if not path.name.startswith(("._", ".~", "~"))
]
workbook_path = max(workbooks, key=lambda path: path.stat().st_size)
workbook = load_workbook(workbook_path, read_only=True, data_only=False)
sheet_name = workbook.worksheets[0].title
output = Path(r"C:\Users\Public\interactive_clean_probe")
output.mkdir(parents=True, exist_ok=True)
valid, errors, _, _ = build_data(workbook_path, sheet_name, output, None)
Path(r"C:\Users\Public\interactive_clean_probe.txt").write_text(
    f"valid={valid}\nerrors={errors}\nsheet={sheet_name}\n",
    encoding="utf-8",
)
