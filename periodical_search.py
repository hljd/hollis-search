"""Generate periodical search links offline and export the original workbook."""
import io
from copy import copy
from urllib.parse import quote

from openpyxl.utils import get_column_letter
from spreadsheet_processor import insert_output_column

HEADER = "HOLLIS Periodical Search"
PREFIX = "https://hollis.harvard.edu/discovery/search?query=title,begins_with,"
SUFFIX = (",AND&tab=LibraryCatalog&search_scope=MyInstitution&sortby=title"
          "&vid=01HVD_INST:HVD2&facet=rtype,include,journals&mode=advanced&offset=0")


def periodical_url(value):
    title = "" if value is None else str(value).split(":", 1)[0].strip()
    return PREFIX + quote(title, safe="") + SUFFIX if title else ""


def render_periodicals(workbook, sheet_name, title_col, header_row=1):
    """Include every row below the header, even completely blank rows.

    Mutates the in-memory workbook only. Original cell values are not edited.
    """
    sheet = workbook[sheet_name]
    rows = list(range(header_row + 1, sheet.max_row + 1))
    urls = {row: periodical_url(sheet.cell(row, title_col).value) for row in rows}
    if any(len(url) > 32767 for url in urls.values()):
        raise ValueError("A generated URL exceeds Excel's 32,767-character cell limit.")
    insert_output_column(sheet)
    sheet.cell(header_row, 1, HEADER)
    sheet.cell(header_row, 1)._style = copy(sheet.cell(header_row, 2)._style)
    sheet.column_dimensions['A'].width = 60
    preview = []
    for row in rows:
        url = urls[row]
        if url:
            cell = sheet.cell(row, 1, url)
            cell.hyperlink = url
            font = copy(cell.font)
            font.color, font.underline = '0563C1', 'single'
            cell.font = font
        # Column letters keep duplicate/blank headers distinct in the preview.
        preview.append({HEADER: url, **{
            f"{get_column_letter(col)} — {sheet.cell(header_row, col).value or '(blank header)'}":
                '' if sheet.cell(row, col).value is None else str(sheet.cell(row, col).value)
            for col in range(2, sheet.max_column + 1)}})
    buffer = io.BytesIO()
    workbook.save(buffer)
    count = sum(bool(url) for url in urls.values())
    return buffer.getvalue(), len(rows), count, preview
