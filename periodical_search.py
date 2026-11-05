"""Generate periodical search links offline and export the original workbook."""
import re
from copy import copy
from urllib.parse import quote

from openpyxl.utils import get_column_letter
from spreadsheet_processor import ALL_BORDERS, insert_output_column, set_hyperlink, workbook_bytes

HEADER = "HOLLIS Periodical Search"
PREFIX = "https://hollis.harvard.edu/discovery/search?query=title,begins_with,"
SUFFIX = (",AND&tab=LibraryCatalog&search_scope=MyInstitution&sortby=title"
          "&vid=01HVD_INST:HVD2&facet=rtype,include,journals&mode=advanced&offset=0")


def periodical_url(value):
    title = "" if value is None else re.split(r"[:()\[\]/\\]", str(value), maxsplit=1)[0].strip()
    return PREFIX + quote(title, safe="") + SUFFIX if title else ""


def render_periodicals(workbook, sheet_name, title_col, header_row=1):
    """Include every row below the header, even completely blank rows.

    Mutates the in-memory workbook only. Original cell values are not edited.
    """
    sheet = workbook[sheet_name]
    rows = range(header_row + 1, sheet.max_row + 1)
    urls = {row: periodical_url(sheet.cell(row, title_col).value) for row in rows}
    if any(len(url) > 32767 for url in urls.values()):
        raise ValueError("A generated URL exceeds Excel's 32,767-character cell limit.")
    insert_output_column(sheet)
    sheet.cell(header_row, 1, HEADER)
    sheet.cell(header_row, 1)._style = copy(sheet.cell(header_row, 2)._style)
    sheet.column_dimensions['A'].width = 60
    # Column letters distinguish duplicate or blank headers.
    labels = {
        col: f"{get_column_letter(col)} — {sheet.cell(header_row, col).value or '(blank header)'}"
        for col in range(2, sheet.max_column + 1)
    }
    preview = []
    for row in rows:
        url = urls[row]
        if url:
            cell = sheet.cell(row, 1, url)
            set_hyperlink(cell, url)
        preview.append({HEADER: url, **{
            label:
                '' if sheet.cell(row, col).value is None else str(sheet.cell(row, col).value)
            for col, label in labels.items()}})
    # Border the entire used rectangle, including holes, blank link cells,
    # and whitespace-only cells, rather than only individually nonblank cells.
    for cells in sheet.iter_rows():
        for cell in cells:
            cell.border = copy(ALL_BORDERS)
    count = sum(bool(url) for url in urls.values())
    return workbook_bytes(workbook), len(rows), count, preview
