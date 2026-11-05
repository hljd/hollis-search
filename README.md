# HOLLIS Book Search

A Python and Streamlit application developed for Harvard Library’s Judaica Division to help staff check spreadsheet book lists against the HOLLIS Library Catalog and identify potential purchase candidates.

> **Active development:** Local/testing deployments use Sandbox; the official Judaica deployment uses Production. Each deployment configures its own environment and API key. Periodical links need no credentials.

## Original development
The application was originally designed and developed by Chloe Qiu (HSPH ’27). The project is maintained by the Judaica Division of Harvard Library.

## What the app does

- Accepts Excel and delimited text files.
- Lets users choose a worksheet, a header row within the first five rows, and separate ISBN and original-title columns.
- Normalizes and validates ISBNs, then searches HOLLIS through Harvard APIgee’s Primo API.
- Confirms ISBN matches and checks for the displayed HOLLIS number.
- Adds a **Notes** column in column A and colors rows for purchase screening or review.
- Shows progress, elapsed search time, an estimated remaining time, and a results preview.
- Downloads a processed Excel workbook while retaining as much of the original workbook as practical.

**Book searches are ISBN-only.** The original-title selector is required by the book interface, but book titles are not searched or compared. The separate **Title Search (Periodicals)** mode generates links for manual HOLLIS searches; it does not query the catalog or classify holdings.

## Result colors and Notes

| Row color | Meaning | Notes |
| --- | --- | --- |
| Green | No confirmed ISBN match after a successful search with sufficient candidate metadata; a potential purchase candidate. | Blank. |
| Red | Exactly one confirmed ISBN match with a nonempty displayed HOLLIS number, with no unresolved candidates. | Matching title, linked to HOLLIS when a link can be constructed. |
| Yellow | Multiple confirmed matches, or a single confirmed match without a displayed HOLLIS number. | Matching title(s); records without a HOLLIS number receive a manual-review note. |
| Uncolored | Missing or invalid ISBN, search failure, insufficient metadata, or a row not searched because the run stopped. | Explanation of the issue and any available partial matches. |

The displayed HOLLIS number is read from **`pnx.display.lds01`**. This mapping was checked against three Sandbox API responses and their HOLLIS export on September 24, 2026. Generic identifiers such as `mms`, `control.recordid`, or an Alma ID in a URL are not used as substitutes.

This is a **screening rule**, not independent proof of physical holdings or permanent ebook ownership. Yellow and uncolored rows require review. “No HOLLIS number” does not itself prove that Harvard does not own a book.

A single yellow match without a HOLLIS number is formatted as:

```text
No HOLLIS number; manual review required: Matching title
```

The whole Notes cell links to the record when a URL is available. For multiple matches, each title and available URL appears on a separate line; the review prefix is added only to records missing a HOLLIS number. Multiple destinations are written as text because this implementation uses one Excel cell hyperlink per cell. The on-screen preview displays the Notes text; the inserted hyperlink is in the downloaded workbook.

The summary label **Search errors** includes all uncolored rows, including missing and invalid ISBNs. Uncolored means the app does not apply a new fill; any original fill remains.

## Project files

Keep these files together, using the exact names below. Remove download suffixes such as `(1)` or `(2)` from Python filenames.

| File | Responsibility |
| --- | --- |
| `app.py` | Streamlit interface, selectors, progress, timing display, results, and download. |
| `hollis_api.py` | API requests, response parsing, ISBN matching, HOLLIS-number extraction, links, and retries. |
| `isbn_utils.py` | Splitting multi-ISBN cells, normalization, checksum validation, and ISBN-10/ISBN-13 comparison. |
| `spreadsheet_processor.py` | File import, duplicate handling, recovery pass, Notes, and workbook export. |
| `periodical_search.py` | Offline periodical-title URL generation and hyperlink workbook export. |
| `time_estimator.py` | Per-run timing samples, pending searches, and scheduled-wait estimates. |
| `requirements.txt` | Python dependencies. |
| `README.md` | Setup and usage documentation. |

## Run locally

The source uses Python 3.10+ syntax. Use a Python version compatible with the installed dependencies and, when deploying, the selected hosting runtime.

From the project folder, create a virtual environment:

```bash
python3 -m venv .venv
```

Activate it on macOS or Linux:

```bash
source .venv/bin/activate
```

On Windows PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

Install the dependencies:

```bash
python -m pip install -r requirements.txt
```

The current dependency ranges are:

```text
streamlit>=1.40,<2
requests>=2.32,<3
openpyxl>=3.1,<4
```

Create a `.streamlit` folder in the project directory, then create `.streamlit/secrets.toml` with:

```toml
HOLLIS_API_KEY = "REPLACE_WITH_YOUR_API_KEY"
HOLLIS_ENVIRONMENT = "Sandbox"
```

Keep both settings at the top level, outside any TOML section. The environment is a visible label, not a dropdown. Missing/invalid settings block ISBN searches rather than choosing a default.

| Deployment | Secrets location | Environment | Key |
| --- | --- | --- | --- |
| Local | `.streamlit/secrets.toml` | `HOLLIS_ENVIRONMENT = "Sandbox"` | Sandbox key |
| Testing (`hollissearchtest.streamlit.app`) | App Secrets | `HOLLIS_ENVIRONMENT = "Sandbox"` | Sandbox key |
| Official (`hollissearch.streamlit.app`) | App Secrets | `HOLLIS_ENVIRONMENT = "Production"` | Production key |

Restart the local app after editing secrets. The environment selects the API endpoint and record-link host. The key is sent in `X-Api-Key`; the API secret is unused. Periodical links need neither setting and always use `hollis.harvard.edu`.

Add these entries to your repository’s `.gitignore` before committing:

```gitignore
.streamlit/secrets.toml
.env
.venv/
venv/
__pycache__/
*.pyc
.DS_Store
```

An optional `.streamlit/secrets.toml.example` may be committed with placeholder values only. Never commit the real secrets file or paste credentials into logs, screenshots, or issue reports. A `.gitignore` entry does not remove a secret already committed to Git history or prevent a manual upload through the GitHub website.

Start the app:

```bash
streamlit run app.py
```

If the command is not found, use `python -m streamlit run app.py` in the activated environment.

## Using the app

1. Select **ISBN Search (Books)** or **Title Search (Periodicals)**, then upload a supported file.
2. For text files, check the preview and adjust **Text file options** if the characters or columns look incorrect.
3. Select the worksheet and its header row. The default is row 1; available choices are within rows 1–5.
4. Select the ISBN column and a different original-title column. Check these selections carefully.
5. Check the environment label configured by the deployment owner.
6. Click **Search HOLLIS** and follow the progress and timing messages.
7. Review the summary and Notes, then download the checked workbook.

Only nonempty rows below the selected header are processed. The first-five-rows preview shows columns A–L; all eligible columns remain available in the selectors. Changing the file or search settings clears the previous downloadable result. Completed results and elapsed time are retained in the current Streamlit session across ordinary reruns, but are not a permanent run history.

### Periodical title search

Select **Title Search (Periodicals)** before uploading. Choose the worksheet, header row (rows 1–5), and **Periodical Title Column**, then click **Generate Search Links**. This mode accepts the same Excel and text formats as book searches, including single-column files, and requires no API key or network requests.

Titles are converted to text, cut before the first colon, and trimmed only at the ends. Internal spaces, capitalization, punctuation, and Unicode characters are preserved. Only the title portion is URL-encoded (spaces become `%20`). Links use `title,begins_with` and the fixed `facet=rtype,include,journals` filter, along with the specified LibraryCatalog/MyInstitution advanced-search parameters.

The output is `Periodical_Search_<original_filename>.xlsx`:

- Column A contains **HOLLIS Periodical Search** with full clickable URLs.
- All original columns shift right; values and other worksheets are preserved.
- Thin black borders cover the selected sheet's entire used rectangle, including blank cells. Merged cells retain outer borders.
- All rows below the header through the last used row count toward the summary. Blank titles and empty prefixes before a colon produce blank links.
- The preview shows original columns with column letters to distinguish duplicate headers.

Click **Download Excel** to save a new file. Switching modes or updating source code clears previous results. Links do not verify catalog holdings. Formula references are not rewritten; URLs exceeding Excel's 32,767-character limit prevent export.

### Supported input formats

- `.xlsx` workbooks.
- `.csv`, `.tsv`, and delimited `.txt` files.
- Comma, tab, semicolon, and pipe separators, with automatic detection or manual selection.

Text encoding options include UTF-8, UTF-16, UTF-32, Windows-1252, Latin-1, Windows-1255, Windows-1251, Shift JIS, GB18030, and Mac Roman. Automatic mode recognizes UTF-16/UTF-32 byte-order marks, otherwise tries UTF-8 and falls back to Windows-1252 with a warning. It is not a universal encoding detector: use an explicit setting if text looks wrong. UTF-16 without a byte-order mark requires the appropriate LE or BE option.

Trailing null padding is removed with a warning. Unsupported embedded control characters are rejected. Text imports create a worksheet named `Books` and preserve fields as text, including leading zeros and values beginning with `=`.

Legacy `.xls`, macro-enabled `.xlsm`, and `.ods` files are not supported. Convert them to `.xlsx` or a supported text format first.

### ISBN preparation

Use one or more ISBNs per cell, preferably stored as text. Multiple ISBNs can be separated by commas, semicolons, slashes, pipes, or line breaks; full-width commas/semicolons/slashes (`，；／`) and the ideographic comma (`、`) are also supported. For example, `978-0-306-40615-7; 0140328726` supplies two ISBNs. In CSV files, quote a field containing commas so it remains one column.

Each entry is validated separately. The app removes whitespace and hyphens within an ISBN, handles integer-valued numbers and digit strings ending in `.0`, and accepts a final `X` in ISBN-10 values. It validates length, character format, and ISBN-10/ISBN-13 checksums. Spaces and hyphens alone are not list separators; repeated separators and empty entries are ignored.

Valid ISBNs are considered in their listed order, reusing results already obtained in the run. The app stops at the first ISBN with confirmed matching records and applies the existing HOLLIS-number, color, and Notes rules to that search. It does not combine matches across ISBNs. A search containing a confirmed match alongside unverifiable candidates still stops the row and retains its uncolored manual-review result.

Invalid entries are skipped while valid ISBNs are searched. When at least one valid ISBN is present, invalid entries do not prevent green, red, or yellow coloring. If every valid ISBN search reliably reports no match, the row becomes green with blank Notes. Matches follow the existing red/yellow rules. Search failures or incomplete HOLLIS metadata still leave a row uncolored when a reliable classification cannot be established.

Blank cells and cells with no valid ISBN remain uncolored. The app cannot restore leading zeros already lost by Excel. ISBN formulas are not evaluated; use their calculated values as input.

## Search and API behavior

Requests use `q=any,contains,<ISBN>` with `vid=01HVD_INST:HVD2`, `tab=LibraryCatalog`, and `scope=MyInstitution`. Returned candidates are compared against **`pnx.addata.isbn`**, allowing equivalent ISBN-10 and ISBN-13 values. Titles come from **`pnx.display.title`**, and records are deduplicated using their returned `@id`.

The client retrieves pages of up to 50 candidates and stops for manual review when a query exceeds its configured 500-result limit or a response is incomplete or inconsistent. If candidate metadata cannot establish the result reliably, the app does not classify the row green. Two or more confirmed matches already qualify for yellow review.

Identical normalized ISBN strings share a result within each run, including across cells containing different ISBN lists. ISBN-10/ISBN-13 equivalence is used for matching, but equivalent inputs of different lengths may still generate separate searches. No shared cross-user result cache is implemented.

| Setting | Current behavior |
| --- | --- |
| Request pacing | Sequential requests, with a normal 2-second gap after a response. |
| Timeouts | 5-second connection timeout and 30-second read timeout. |
| Temporary failures | Retries HTTP 429, 500, 502, 503, 504, timeouts, and connection failures. |
| Attempts | Up to four attempts per page per pass. |
| Retry delays | 5, 15, and 30 seconds, plus jitter; honors a longer valid `Retry-After`. |
| Long server wait | A `Retry-After` over 120 seconds stops the run rather than retrying early. |
| Recovery | One recovery pass for transient failures and any searches deferred by a stopped first pass, after a wait of at least 30 seconds. |
| Persistent rate limit | Exhausted HTTP 429 retries stop the first pass; a persistent 429 during recovery ends the run. |
| Authentication failure | HTTP 401 or 403 stops the run without recovery. |

These are application defaults, not documented Harvard API quota limits. Pacing applies per client/run, so simultaneous users can collectively exceed a shared API key’s limit.

## Elapsed time and estimated remaining time

The elapsed timer includes requests, scheduled pauses, and recovery searches. It excludes workbook export. The final duration appears with the results.

The remaining-time estimate begins after five completed query attempts have supplied timing samples. The initial estimate counts unique valid ISBNs, including alternatives that may be skipped after a match. Pending work decreases as those alternatives are skipped. Normalized duplicates and invalid entries do not add searches; a query repeated in recovery contributes another timing sample. The estimate uses average observed query duration and pending work, including known scheduled retry/recovery waits. Completed retry delays remain in the historical average for that run.

Scheduled wait countdowns refresh roughly once per second. During a blocking HTTP request, the display may not update until the request returns or times out. Estimates can rise after new delays and may overestimate when historical retry costs are combined with a known wait. Small batches can finish before an estimate appears. No historical timing database is required or maintained.

Row progress measures rows with a result, including provisional failures. It can reach 100% before recovery finishes; wait for the completed results and download button.

## Excel output and preservation limits

ISBN output is `Searched_<original_filename>.xlsx`; periodical output is `Periodical_Search_<original_filename>.xlsx`. Neither overwrites the upload.

The ISBN workflow inserts **Notes** in column A of the selected worksheet, with its header on the chosen header row. Original columns shift right. Original worksheets, cell values, formula text, fonts, row heights, and basic styles are retained where supported by openpyxl. Existing column widths and merged-cell ranges are explicitly shifted with the inserted column.

Processed rows receive light green, red, or yellow fills. Thin black borders are applied to the header and processed rows across the original worksheet column span plus Notes. Other worksheets are retained and are not searched in that run. Text files become new workbooks, so they have no original Excel formatting to preserve.

Column insertion does **not** repair formula references, table ranges, named ranges, filters, or freeze-pane positions. Complex workbook features may not survive an openpyxl round trip unchanged. Ordinary tabular book lists work best; inspect output from complex workbooks. Notes beyond Excel’s 32,767-character cell limit are shortened in the workbook, with the full text retained in the preview.

## Deploy through GitHub and Streamlit Community Cloud

1. Commit all six application Python modules, `requirements.txt`, this README, and `.gitignore` to the deployed repository branch. Exclude the real secrets file.
2. Connect the repository to Streamlit Community Cloud and select `app.py` as the entry point.
3. Set `HOLLIS_API_KEY` and `HOLLIS_ENVIRONMENT` in the app’s private **Secrets**, following the deployment table above.
4. Deploy with a compatible Python version, then test a small book list in the configured environment.
5. Configure the app’s sharing settings for the intended staff testers and share its URL.

The key stays out of GitHub and the app interface, but is stored on the hosting server so the app can access HOLLIS. Staff upload their spreadsheets to that server for processing. The application builds output in memory and keeps completed output in session state; it does not implement a permanent upload archive.

Deploy related modules together. Mixing old and new files can produce errors such as a missing `hollis_number` attribute or an unexpected callback argument. After an update, start a fresh search; reboot the hosted app if it is still running an older version.

## Troubleshooting

| Problem | What to check |
| --- | --- |
| Missing/invalid environment | Set top-level `HOLLIS_ENVIRONMENT` to `"Sandbox"` or `"Production"` in secrets. |
| Missing API key | Local file must be `.streamlit/secrets.toml`, with `HOLLIS_API_KEY` at the top level; hosted deployments need the same key in Secrets settings. |
| All ISBNs appear invalid | Confirm the ISBN and title selectors were not swapped, choose the correct header, and check for lost leading zeros or unsupported separators between values. |
| Garbled text or incorrect columns | Set the text encoding and separator explicitly. |
| Frequent HTTP 429 messages | Allow scheduled waits and recovery; avoid simultaneous large runs using the same key. |
| Missing module or attribute after deployment | Check exact filenames and upload all related modules from the same version. |
| Workbook cannot be exported | Retain the original file, review available results, and check the server logs; unusual workbook structures or mismatched scripts may be involved. |
| Timer appears stationary | Scheduled waits update regularly, but an in-flight HTTP request can temporarily block display updates. |

Logging defaults to `WARNING`. To enable request/status diagnostics locally:

```bash
HOLLIS_LOG_LEVEL=INFO streamlit run app.py
```

Application logging can include normalized ISBNs, HTTP statuses, retry waits, and match counts. The application’s logging statements do not include the API key or API secret. Return to `WARNING` or `ERROR` to reduce log output.

## References

- [Harvard: Using the Primo API through Harvard APIgee](https://harvardwiki.atlassian.net/wiki/spaces/LibraryStaffDoc/pages/896663553)
- [Harvard: Primo VE API query and result guidance](https://harvardwiki.atlassian.net/wiki/spaces/LibraryStaffDoc/pages/43422622)
- [Ex Libris: Primo search response documentation](https://developers.exlibrisgroup.com/primo/apis/search-output/)
- [Streamlit Community Cloud secrets management](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/secrets-management)

## Feedback and ongoing development

This app will continue to be updated based on Sandbox testing and staff feedback. When reporting an issue, include the selected environment, input format, header/column selections, expected behavior, and the error message or a small non-sensitive example. Do not include API credentials. Further work includes book title-based searching and validation of the workflow against the Production environment.
