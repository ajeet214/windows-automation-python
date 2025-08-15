from __future__ import annotations

import logging
from pathlib import Path
import win32com.client as win32

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


def _get_worksheet_by_name(wb, sheet_name: str):
    """
    Return a Worksheet object matched case-insensitively by name.
    Returns None if not found or if the matching sheet isn't a real worksheet.
    """
    # Worksheets collection excludes chart sheets, which is what we want for printing ranges.
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        if ws.Name.strip().lower() == sheet_name.strip().lower():
            return ws
    return None


def export_sheet_to_pdf(xlsx: Path, pdf_out: Path, sheet_name: str | None = None) -> None:
    """
    Export a worksheet (by name) or the active sheet to PDF, with robust sheet resolution.
    Raises a ValueError with available sheet names if the requested sheet doesn't exist.
    """
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(str(xlsx))

    try:
        if sheet_name:
            ws = _get_worksheet_by_name(wb, sheet_name)
            # print(ws.Name)
            if ws is None:
                # Build a helpful error message listing available worksheet names
                available = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
                raise ValueError(
                    f"Worksheet '{sheet_name}' not found. Available worksheets: {available}"
                )
        else:
            ws = wb.ActiveSheet  # could be a chart sheet; if so, switch to first worksheet
            # Ensure we export a real worksheet, not a chart sheet.
            if ws.Name not in [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]:
                ws = wb.Worksheets(1)

        # Fit to one page
        ws.PageSetup.Orientation = 2      # xlLandscape
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = 1

        pdf_out.parent.mkdir(parents=True, exist_ok=True)
        # 0 = xlTypePDF
        wb.ExportAsFixedFormat(Type=0, Filename=str(pdf_out))
        logging.info("Exported PDF -> %s", pdf_out)

    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()


if __name__ == "__main__":
    export_sheet_to_pdf(
        xlsx=Path(r"D:\LifeLongLearning\windows-automation-python\data\reports\weekly.xlsx"),
        pdf_out=Path(r"D:\LifeLongLearning\windows-automation-python\data\reports\weekly_report.pdf"),
        sheet_name="Sheet1",
    )
