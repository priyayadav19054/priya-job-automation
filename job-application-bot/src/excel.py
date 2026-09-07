from pathlib import Path
import openpyxl


EXCEL_FILE = Path("input/job-tracker.xlsx")
SHEET_NAME = "Jobs"


def get_next_pending_job():
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    # data_only=False is important because we need hyperlink information
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=False)
    ws = wb[SHEET_NAME]

    headers = [cell.value for cell in ws[1]]

    for row in ws.iter_rows(min_row=2):

        data = {}

        for i, cell in enumerate(row):
            header = headers[i]

            # Normal cell value
            value = cell.value

            # If this cell contains an Excel hyperlink,
            # use the actual URL instead of the displayed text.
            if header == "URL" and cell.hyperlink:
                value = cell.hyperlink.target

            data[header] = value

        status = str(
            data.get("Application Status") or "PENDING"
        ).strip().upper()

        if status == "PENDING":
            return data

    return None