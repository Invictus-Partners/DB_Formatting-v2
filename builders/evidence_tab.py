import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# ── you can import this constant from settings.py later ──────────────
CANON_COLS = (
    "device_name", "database_name", "db_version", "option_name",
    "feature_name", "file_name", "result", "note", "dbid", "name",
    "version", "detected_usages", "total_samples", "currently_used",
    "first_usage_date", "last_usage_date", "feature_info",
    "last_sample_date", "last_sample_period", "sample_interval",
    "description", "host_name", "instance_name", "evidence"
)

def build_evidence_tab(json_path: str, workbook: Workbook | None = None) -> Workbook:
    """
    Read raw Options‑Usage JSON and write a fresh 'Evidence' sheet.
    Returns the (possibly newly‑created) openpyxl Workbook.
    """
    df = pd.read_json(json_path, orient="records")

    # ensure stable column order and presence
    df = df.reindex(columns=CANON_COLS)

    # tidy up obvious placeholders
    df.replace({"nan": pd.NA}, inplace=True)

    # pick or create workbook
    wb = workbook or Workbook()
    if "Evidence" in wb.sheetnames:
        wb.remove(wb["Evidence"])
    ws = wb.create_sheet("Evidence", 0)  # make it the first sheet

    # write the DataFrame
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    # optional: freeze header row, auto‑width later if you like
    ws.freeze_panes = ws["A2"]

    return wb