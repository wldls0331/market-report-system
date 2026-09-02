"""PDF export.

Windows local: Microsoft Excel COM (print ranges, charts).
Linux / Render: reportlab renderer from Google Sheets values.
"""

from __future__ import annotations

import re
import sys
from contextlib import contextmanager
from pathlib import Path

from config.cell_mapping import (
    EXPECTED_PDF_PAGES,
    PDF_FIT_TO_PAGES_TALL,
    PDF_FIT_TO_PAGES_WIDE,
    PDF_HELPER_COLUMNS,
    PDF_HELPER_ROWS,
    PDF_ORDER_OVER_THEN_DOWN,
    PDF_ORIENTATION_PORTRAIT,
    PDF_PAGE_HIDDEN_ROWS,
    PDF_PAGE_RANGES,
    PDF_PRINT_AREA,
)

XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0


class PdfExportError(RuntimeError):
    pass


def excel_com_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


def apply_pdf_print_layout(worksheet, print_area: str | None = None) -> None:
    """Keep original margins/paper. Fit the given range to one portrait page."""
    page_setup = worksheet.PageSetup
    page_setup.PrintArea = print_area or PDF_PRINT_AREA
    page_setup.Orientation = PDF_ORIENTATION_PORTRAIT
    page_setup.Zoom = False
    page_setup.FitToPagesWide = PDF_FIT_TO_PAGES_WIDE
    page_setup.FitToPagesTall = PDF_FIT_TO_PAGES_TALL
    page_setup.Order = PDF_ORDER_OVER_THEN_DOWN
    page_setup.CenterHorizontally = False
    page_setup.CenterVertically = False
    page_setup.PrintGridlines = False


def _set_charts_plot_hidden_cells(worksheet) -> None:
    """Keep chart series visible while helper source rows are hidden for print."""
    try:
        count = int(worksheet.ChartObjects().Count)
    except Exception:
        return
    for index in range(1, count + 1):
        try:
            worksheet.ChartObjects(index).Chart.PlotVisibleOnly = False
        except Exception:
            continue


def _refresh_charts(worksheet) -> None:
    try:
        count = int(worksheet.ChartObjects().Count)
    except Exception:
        return
    for index in range(1, count + 1):
        try:
            worksheet.ChartObjects(index).Chart.Refresh()
        except Exception:
            continue


def _set_rows_hidden(worksheet, start: int, end: int, hidden: bool) -> None:
    worksheet.Range(f"{start}:{end}").EntireRow.Hidden = hidden


@contextmanager
def _temporary_pdf_helper_view(worksheet):
    """Hide helper columns for PDF, then restore. Row hides are applied per page."""
    cover_start, cover_end = 22, 34
    previous_rows_hidden = bool(worksheet.Range(f"{cover_start}:{cover_end}").EntireRow.Hidden)
    previous_cols_hidden: dict[str, bool] = {}
    for letter in PDF_HELPER_COLUMNS:
        previous_cols_hidden[letter] = bool(worksheet.Range(f"{letter}:{letter}").EntireColumn.Hidden)

    for letter in PDF_HELPER_COLUMNS:
        worksheet.Range(f"{letter}:{letter}").EntireColumn.Hidden = True
    _set_charts_plot_hidden_cells(worksheet)
    try:
        yield
    finally:
        _set_rows_hidden(worksheet, cover_start, cover_end, previous_rows_hidden)
        for letter, hidden in previous_cols_hidden.items():
            worksheet.Range(f"{letter}:{letter}").EntireColumn.Hidden = hidden


def _apply_page_row_visibility(worksheet, hidden_rows: tuple[int, int]) -> None:
    _set_rows_hidden(worksheet, 22, 34, False)
    start, end = hidden_rows
    _set_rows_hidden(worksheet, start, end, True)


def count_pdf_pages(pdf_path: Path) -> int:
    path = Path(pdf_path)
    try:
        import fitz

        doc = fitz.open(path)
        count = int(doc.page_count)
        doc.close()
        return count
    except Exception:
        data = path.read_bytes()
        match = re.search(rb"/Type\s*/Pages[^>]*?/Count\s+(\d+)", data)
        if not match:
            raise PdfExportError("Could not determine PDF page count.")
        return int(match.group(1))


def _merge_pdfs(parts: list[Path], output: Path) -> Path:
    try:
        import fitz
    except ImportError as exc:
        raise PdfExportError("PyMuPDF is required to merge the report PDF.") from exc

    merged = fitz.open()
    try:
        for part in parts:
            doc = fitz.open(part)
            if doc.page_count != 1:
                count = doc.page_count
                doc.close()
                raise PdfExportError(f"Expected 1 PDF page for a report section, got {count}.")
            merged.insert_pdf(doc)
            doc.close()
        if merged.page_count != EXPECTED_PDF_PAGES:
            raise PdfExportError(
                f"Merged PDF has {merged.page_count} page(s); expected {EXPECTED_PDF_PAGES}."
            )
        output.parent.mkdir(parents=True, exist_ok=True)
        if output.exists():
            output.unlink()
        merged.save(output)
    finally:
        merged.close()
    return output


def prepare_sheet_and_export_pdf(excel_app, workbook, worksheet, pdf_path: Path) -> tuple[Path, int]:
    pdf_path = Path(pdf_path)
    temps: list[Path] = []
    try:
        with _temporary_pdf_helper_view(worksheet):
            try:
                excel_app.CalculateFullRebuild()
            except Exception:
                excel_app.Calculate()
            _refresh_charts(worksheet)
            for index, rng in enumerate(PDF_PAGE_RANGES, start=1):
                hidden_rows = PDF_PAGE_HIDDEN_ROWS[index - 1] if index <= len(PDF_PAGE_HIDDEN_ROWS) else PDF_HELPER_ROWS
                _apply_page_row_visibility(worksheet, hidden_rows)
                temp_path = pdf_path.parent / f"~{pdf_path.stem}_p{index}.pdf"
                temps.append(temp_path)
                apply_pdf_print_layout(worksheet, rng)
                export_sheet_with_com(worksheet, temp_path)
            exported = _merge_pdfs(temps, pdf_path)
    finally:
        for temp_path in temps:
            if temp_path.exists():
                temp_path.unlink()
    apply_pdf_print_layout(worksheet, PDF_PRINT_AREA)
    return exported, count_pdf_pages(exported)


def export_sheet_with_com(worksheet, pdf_path: Path) -> Path:
    pdf_path = Path(pdf_path)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    if pdf_path.exists():
        pdf_path.unlink()
    worksheet.ExportAsFixedFormat(
        Type=XL_TYPE_PDF,
        Filename=str(pdf_path.resolve()),
        Quality=XL_QUALITY_STANDARD,
        IncludeDocProperties=True,
        IgnorePrintAreas=False,
        OpenAfterPublish=False,
    )
    if not pdf_path.exists():
        raise PdfExportError("PDF file was not created.")
    return pdf_path


def pdf_provider() -> str:
    return "excel_com" if excel_com_available() else "cloud"


def export_report_pdf_cloud(
    *,
    pdf_path: Path,
    report_date,
    data_reference_date,
    pricing_month: str,
    sheet_name: str,
    inputs: dict,
    extra_cells: dict | None = None,
) -> tuple[Path, int]:
    from services.pdf_cloud import render_market_report_pdf

    exported = render_market_report_pdf(
        pdf_path=pdf_path,
        report_date=report_date,
        data_reference_date=data_reference_date,
        pricing_month=pricing_month,
        sheet_name=sheet_name,
        inputs=inputs,
        extra_cells=extra_cells,
    )
    return exported, count_pdf_pages(exported)


def export_sheet_to_pdf(excel_path: Path, sheet_name: str, pdf_path: Path) -> Path:
    if not excel_com_available():
        raise PdfExportError(
            "Windows Excel COM is unavailable. Microsoft Excel and pywin32 are required on this machine."
        )

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(Path(excel_path).resolve()), UpdateLinks=0, ReadOnly=False)
        worksheet = workbook.Worksheets(sheet_name)
        path, _pages = prepare_sheet_and_export_pdf(excel, workbook, worksheet, pdf_path)
        workbook.Save()
        return path
    except PdfExportError:
        raise
    except Exception as exc:
        raise PdfExportError(f"PDF export failed: {exc}") from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=True)
            except Exception:
                pass
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
