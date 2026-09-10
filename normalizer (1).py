"""
normalizer.py

Reads EVERY runsheet in the current folder -- Excel (.xlsx/.xls) or PDF -- whatever
header layout each one uses, and produces three files:
  - Normalized_Runsheet_CD.xlsx            (Conveyance / Assignment / Lease / Release)
  - Normalized_Runsheet_NCD.xlsx           (Deed of Trust, Easement, Agreement, Lien, etc.)
  - Normalized_Runsheet_MissingFields.txt  (rows with any blank or uncertain field)

USAGE
    python3 normalizer.py          # processes everything in the folder

WHAT IT HANDLES ON ITS OWN
  * Headers that differ from runsheet to runsheet: GRANTOR/GRANTEE, PARTY OF THE
    FIRST/SECOND PART, INST./INSTR./INSTRUMENT TYPE/INST. NAME,
    VOL-PG/Vol/Pg/Book/Page/Document Vol_Pg, DOD/DOR/DOF/EXEC-EFF DATE/Date Executed...
    (see HEADER_ALIASES).
  * The header row does not have to be the first row: the first 30 rows are searched.
  * Book and Page split apart, whether they arrive combined ("375-266", "104/287",
    "330_25", "Vol 104 Pg 287") or in separate columns. A book-type prefix
    ("DR 104/287", "OR 165") is split off and used for Booktype, which is more reliable
    than inferring it from the year.
  * Rows a runsheet split across several physical lines (overflowed text) are rejoined.
  * PDFs: every table on every page is read on its own -- these runsheets repeat the
    header on each page, and the number of columns the extractor detects can vary from
    page to page, so mapping page by page is safer. Requires pdfplumber.
  * Duplicates: the same instrument repeated ("Main"+"Index" sheets of one file, or
    overlapping supplemental runsheets) is kept only once (see DROP_DUPLICATES).
  * A blank Grantor/Grantee is never left empty: it is marked "Unclear, see instrument"
    and listed in the report along with the file/sheet/page the row came from.

FIX (2026-09-02): 'merge_split_rows' converted EVERY column to text (via str()) before
grouping, even when a group held a single value and there was nothing to merge. For date
columns that destroyed the original datetime/Timestamp object ("1907-04-12 00:00:00"
instead of the object), and 'clean_specific_dates' then re-parsed that text with an
ambiguous MM-DD-YY regex, producing nonsense dates (e.g. April 12, 1907 came out as
07/04/2012). 'join_text' now preserves the original value when a group holds a single
item -- it only stringifies and concatenates when split rows really do need merging.
"""

import glob
import re
import warnings
from datetime import date, datetime

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.dataframe import dataframe_to_rows

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ---------------------------------------------------------------------------------------
# Output schema
# ---------------------------------------------------------------------------------------

TARGET_COLS = [
    "Class", "Type", "Subtype", "Grantor", "Grantee", "Date", "Recorded",
    "Inst.Date", "Acknowledged", "Filed", "Booktype", "Book", "Page",
    "Inst.No.", "County", "State", "Notes", "Requirements", "Essences",
    "Flags", "Restrictions", "Warnings", "Reviews", "Files",
]

OUTPUT_FILES = {
    "Normalized_Runsheet_CD.xlsx",
    "Normalized_Runsheet_NCD.xlsx",
    "Normalized_Runsheet_Master.xlsx",
}

# Defense-in-depth against re-ingesting a normalized file under a different name (e.g. the
# "<prefix>_CD.xlsx" / "<prefix>_NCD.xlsx" produced by history_of_title_txt_to_xlsx.py):
# skip by filename pattern first (cheap), and 'looks_already_normalized' below skips by
# actual column content too, in case a file with any other name slips through.
OUTPUT_FILE_SUFFIXES = ("_CD.XLSX", "_NCD.XLSX", "_MASTER.XLSX")

# what a blank Grantor/Grantee becomes instead of being left empty -- matches the
# convention used by history_of_title_txt_to_xlsx.py so a reviewer sees the same flag
# regardless of which pipeline produced the row
UNCLEAR_MARKER = "Unclear, see instrument"

# The same instrument often arrives more than once ("Main" + "Index" sheets of one file,
# or several overlapping supplemental runsheets). By default ONE row per instrument is
# kept. Set to False to import every row exactly as it comes in.
DROP_DUPLICATES = True

# internal helper column holding the row number in the source file
SOURCE_ROW_COL = "__source_row__"

STATE_MAP = {
    "TEXAS": "TX", "NEW MEXICO": "NM", "OKLAHOMA": "OK",
    "LOUISIANA": "LA", "COLORADO": "CO",
}

# column names, normalized to letters-only uppercase, that identify each field -- used
# both to find the real header row and to map each runsheet's columns
HEADER_ALIASES = {
    "instrument": {"INSTRUMENT", "IMGINSTRUMENT", "INSTRUMENTTITLE", "INSTRUMENTTYPE",
                   "INSTR", "INST", "INSTNAME", "INSTRUMENTNAME", "DOCUMENT", "DOCUMENTTYPE",
                   "INSTRUMENTKIND", "TYPEOFINSTRUMENT", "DOCTYPE"},
    "book": {"BK", "BOOK", "VOL", "VOLUME"},
    "page": {"PG", "PAGE", "PGE"},
    "book_page": {"BOOKPAGE", "BKPG", "VOLPG", "VOLPAGE", "BOOKPG", "BKPAGE", "VOLUMEPAGE",
                  "DOCUMENTVOLPG", "DOCUMENTVOLPAGE", "DOCUMENTBOOKPAGE", "RECORDING",
                  "RECORDINGINFO", "VOLPGDOCUMENT"},
    "inst_date": {"EXECEFFDATE", "INSTDATE", "INSTRUMENTDATE", "DATEEXECUTED", "DATE", "DOD",
                  "EXECUTIONDATE", "DATEOFINSTRUMENT", "EXECDATE", "DATEEXEC"},
    "filed_date": {"FILINGDATE", "FILEDDATE", "FILEDATE", "FILED", "DOF", "DATEFILED"},
    "recorded_date": {"DOR", "DATERECORDED", "RECORDEDDATE", "RECORDINGDATE", "DATEOFRECORD"},
    "grantor": {"PARTYOFTHEFIRSTPART", "GRANTOR", "GRANTER", "GRANTORS", "FROM", "PARTYOFFIRSTPART"},
    "grantee": {"PARTYOFTHESECONDPART", "GRANTEE", "GRANTEES", "TO", "PARTYOFSECONDPART"},
    "county": {"COUNTYESTATE", "COUNTY"},
    "remarks": {"REMARKS", "COMMENTS", "COMMENT"},
    "inst_no": {"IMAGENO", "INSTNO", "IMG", "IMGNO", "IMAGE", "DOCUMENTNO", "DOCNO",
                "INSTRUMENTNO", "FILENO", "CLERKSFILENO", "CCFILENO"},
    "records_type": {"RECORDSTYPE", "RECORDTYPE", "BOOKTYPE"},
}

# candidate keys for grouping split rows in 'merge_split_rows' (order = priority)
MERGE_KEY_ALIASES = [
    "IMAGENO", "IMGINSTRUMENT", "BOOKPAGE", "BK", "BKPG", "VOLPG", "VOL", "VOLUME",
    "INST", "INSTR", "DOCUMENT", "FILINGDATE", "FILEDDATE", "FILEDATE", "ITEM", "ITEMNO",
    "DOCUMENTVOLPG", "VOLPAGE", "NO", "NUM",
]

HEADER_ROW_MARKERS = (
    "PARTYOFTHEFIRSTPART", "INSTRUMENTTITLE", "IMGINSTRUMENT", "INSTNAME",
)

# Columns that only ever appear in THIS script's own output schema (never in a raw source
# runsheet). Any .xlsx sitting in the folder whose header contains at least two of these is
# an already-normalized CD/NCD file -- not a runsheet to (re)normalize -- and gets skipped.
# This is what protects against the CD/NCD files silently feeding back into themselves and
# duplicating every record (with garbage "Class: CD" / "Subtype: ..." notes) on a re-run.
NORMALIZED_OUTPUT_MARKERS = {"CLASS", "SUBTYPE", "BOOKTYPE", "INSTDATE"}


def looks_already_normalized(df_raw: pd.DataFrame, header_idx: int) -> bool:
    header_clean = {clean_column_name(c) for c in df_raw.iloc[header_idx].dropna().values}
    return len(NORMALIZED_OUTPUT_MARKERS & header_clean) >= 2


def clean_column_name(name: object) -> str:
    """Normalize a column name to letters-only uppercase, so headers that differ only in
    punctuation/spacing compare equal across runsheets ("Book/Page" == "BOOK PAGE")."""
    return re.sub(r"[^A-Z]", "", str(name).upper())


def clean_specific_dates(date_val: object) -> str:
    """Return the date as MM/DD/YYYY. A real datetime/Timestamp/date is formatted directly
    (never re-parsed as text); only dates that already arrive as text (e.g. "04/12/1907
    EFF", "DOD 6-1-1965") go through the regex cleanup."""
    if pd.isna(date_val) or date_val == "":
        return ""

    if isinstance(date_val, (datetime, pd.Timestamp, date)):
        return date_val.strftime("%m/%d/%Y")

    d_str = str(date_val).replace("\n", " ").strip()

    if "_" in d_str and re.search(r"\d", d_str):
        return d_str

    # "EFF", "EFF.", "Eff:", "Effective", "effec." -- any abbreviation of "effective"
    eff_match_post = re.search(r"(?i)(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})\s*\bEFF[A-Z]*\.?:?", d_str)
    eff_match_pre = re.search(r"(?i)\bEFF[A-Z]*\.?:?\s*(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})", d_str)

    if eff_match_post:
        d_str = eff_match_post.group(1)
    elif eff_match_pre:
        d_str = eff_match_pre.group(1)
    else:
        if "DOD" in d_str.upper():
            d_str = re.sub(r"(?i)\s*DOD\s*", "", d_str).strip()
        date_pattern = re.search(r"(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})", d_str)
        if date_pattern:
            d_str = date_pattern.group(1)

    try:
        parsed_date = pd.to_datetime(d_str, errors="coerce")
        if pd.notna(parsed_date):
            return parsed_date.strftime("%m/%d/%Y")
    except Exception:
        pass
    return d_str


# Many runsheets write the book type attached to the volume: "DR 104/287", "OPR 82/261",
# "OR 165". That prefix is the MOST reliable source for Booktype (better than guessing
# from the year), so it is extracted and kept instead of discarded.
BOOKTYPE_PREFIX_RE = re.compile(r"^\s*(OPRR|OPR|ORR|0R|OR|DR|PR|MR|RPR|DRR)\b[\s.:_-]*", re.I)
BOOKTYPE_FROM_PREFIX = {
    "DR": "DR", "PR": "DR", "MR": "DR", "DRR": "DR",
    "OPR": "OPR", "OPRR": "OPR", "OR": "OPR", "0R": "OPR", "ORR": "OPR", "RPR": "OPR",
}
# separators used between book and page: 104/287, 375-266, 330_25, 104 287
BOOK_PAGE_SPLIT_RE = re.compile(r"^(\d+)\s*[/\-_]\s*(\d+)\s*$")
BOOK_PAGE_SPACE_RE = re.compile(r"^(\d+)\s+(\d+)$")
# "Cause No. 407", "Doc #20181038", "Case 12-345" -- not a book/page, but a reference
NON_BOOK_REF_RE = re.compile(r"(?i)\b(cause|case|doc|document|inst|instrument|file|clerk|no)\b")


def _first_line(val: object) -> str:
    if pd.isna(val) or str(val).strip() == "":
        return ""
    return str(val).split("\n")[0].strip()


def parse_book_page(val: object) -> tuple[str, str, str, str]:
    """Read a book/page cell in any of the formats these runsheets use and return
    (booktype, book, page, reference).

    "DR 104/287" -> ("DR", "104", "287", "")      "OR 165"     -> ("OPR", "165", "", "")
    "375_266"    -> ("",   "375", "266", "")      "Cause 407"  -> ("", "", "", "Cause 407")

    The 'reference' is whatever is not a real book/page (a cause number, a document
    number, etc.): it goes to Inst.No. instead of polluting Book/Page, and the row is
    flagged in the missing-fields report."""
    raw = _first_line(val)
    if not raw:
        return "", "", "", ""

    booktype = ""
    m = BOOKTYPE_PREFIX_RE.match(raw)
    if m:
        booktype = BOOKTYPE_FROM_PREFIX.get(m.group(1).upper(), "")
        raw = raw[m.end():].strip()

    if not raw:
        return booktype, "", "", ""

    m = BOOK_PAGE_SPLIT_RE.match(raw) or BOOK_PAGE_SPACE_RE.match(raw)
    if m:
        return booktype, m.group(1), m.group(2), ""

    if re.fullmatch(r"\d+", raw):
        return booktype, raw, "", ""

    # something like "Cause No. 407" or "Doc 20181038": not a book/page
    if NON_BOOK_REF_RE.search(raw) or not re.search(r"\d", raw):
        return booktype, "", "", raw

    # last resort: two numbers anywhere in the text ("Vol 104 Pg 287")
    nums = re.findall(r"\d+", raw)
    if len(nums) >= 2:
        return booktype, nums[0], nums[1], raw
    if len(nums) == 1:
        return booktype, nums[0], "", raw
    return booktype, "", "", raw


def clean_book(val: object) -> tuple[str, str]:
    """Return (booktype, book) -- splitting off the book-type prefix when one is attached."""
    raw = _first_line(val)
    if not raw:
        return "", ""
    booktype = ""
    m = BOOKTYPE_PREFIX_RE.match(raw)
    if m:
        booktype = BOOKTYPE_FROM_PREFIX.get(m.group(1).upper(), "")
        raw = raw[m.end():].strip()
    return booktype, raw


def clean_page(val: object) -> str:
    raw = _first_line(val)
    if not raw:
        return ""
    # a page is usually clean, but sometimes carries junk alongside it ("287 (OPR)")
    m = re.match(r"^\d+", raw)
    return m.group(0) if m else raw


def merge_split_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Rejoin rows a runsheet split across several physical lines (text that overflowed a
    cell). The key column (Item #, Volume, etc.) marks where each real record starts;
    anything below it without a key of its own belongs to the previous record and is
    concatenated into the same cell."""
    df = df.dropna(how="all").reset_index(drop=True)
    if len(df) == 0 or len(df.columns) == 0:
        return df

    col_0_name = df.columns[0]
    c0_clean = df[col_0_name].astype(str).str.replace(r"[^A-Z]", "", regex=True).str.upper()
    mask_headers = c0_clean.isin(["INSTRUMENTTITLE", "IMGINSTRUMENT", "INSTRUMENT", "INSTNAME"])
    df = df[~mask_headers].reset_index(drop=True)

    key_col = None
    for c in df.columns:
        if clean_column_name(c) in MERGE_KEY_ALIASES:
            key_col = c
            break
    if not key_col:
        return df

    df[key_col] = df[key_col].replace(r"^\s*$", np.nan, regex=True)
    group_ids = df[key_col].notna().cumsum()

    def join_text(x: pd.Series):
        valid = [v for v in x.dropna() if str(v).strip() != ""]
        if not valid:
            return ""
        if len(valid) == 1:
            # single-row group -> nothing to merge; keep the original value AS IS
            # (datetime, number, text...) instead of forcing it to text.
            return valid[0]
        return " ".join(str(v).strip() for v in valid)

    agg_dict = {col: join_text for col in df.columns}
    agg_dict[key_col] = "first"
    return df.groupby(group_ids).agg(agg_dict).reset_index(drop=True)


def find_header_row(df_raw: pd.DataFrame) -> int:
    """Find the real header row within the first 30 rows -- supports the several runsheet
    layouts in use (Grantor/Grantee, Party of the First Part, Instrument Title...)."""
    for i in range(min(30, len(df_raw))):
        row_vals = [str(x).upper() for x in df_raw.iloc[i].dropna().values]
        row_clean = re.sub(r"[^A-Z]", "", "".join(row_vals))
        if (
            ("GRANTOR" in row_clean and "GRANTEE" in row_clean)
            or any(marker in row_clean for marker in HEADER_ROW_MARKERS)
            or ("INSTRUMENT" in row_clean and "BK" in row_clean)
        ):
            return i
    return -1


def extract_county_state_from_preamble(df_raw: pd.DataFrame, header_idx: int) -> tuple[str, str]:
    """Look for 'XXX County, TX' (or variants) in the rows above the header -- many
    runsheets carry the county/state as loose text above the table."""
    preamble_cells = df_raw.iloc[:header_idx].fillna("").astype(str).values.flatten()
    for cell in preamble_cells:
        for line in str(cell).split("\n"):
            match = re.search(
                r"(?:^|[,|-])\s*([A-Za-z\s]+?)\s+(?:CO\.|COUNTY)"
                r"(?:[\s,]+(TX|TEXAS|NM|NEW\s*MEXICO|OK|OKLAHOMA|LA|LOUISIANA|CO|COLORADO|[A-Z]{2}))?\b",
                line, re.IGNORECASE,
            )
            if not match:
                continue
            c_name = match.group(1).strip().title()
            c_words = c_name.split()
            if len(c_words) > 2:
                c_name = " ".join(c_words[-2:])
            state = ""
            if match.group(2):
                raw_state = re.sub(r"\s+", " ", match.group(2).strip().upper())
                state = STATE_MAP.get(raw_state, raw_state)
            return c_name, state
    return "", ""


def frame_from_raw_table(df_raw: pd.DataFrame, label: str) -> tuple[pd.DataFrame, str, str, str] | None:
    """Take a raw table (an Excel sheet or one page's table from a PDF), find its real
    header row, and return (table with header, county, state, label). Returns None when it
    does not look like a runsheet, or when it is already a normalized output file."""
    if df_raw.empty or len(df_raw.columns) == 0:
        return None

    header_idx = find_header_row(df_raw)
    if header_idx == -1:
        return None

    if looks_already_normalized(df_raw, header_idx):
        print(f"Skipping {label} (already a normalized CD/NCD file, not a source runsheet)")
        return None

    ext_county, ext_state = extract_county_state_from_preamble(df_raw, header_idx)

    df = df_raw.iloc[header_idx + 1:].reset_index(drop=True)
    cols = []
    for j, c in enumerate(df_raw.iloc[header_idx].values):
        name = str(c).strip() if pd.notna(c) else ""
        # PDF headers arrive split across several lines ("Instrument\nType")
        name = re.sub(r"\s+", " ", name.replace("\n", " ")).strip()
        cols.append(name if name else f"Unnamed_{j}")
    df.columns = cols

    # row number AS SEEN in the source file (in Excel, the sheet's row number; in a PDF,
    # the row within that page's table), so a reviewer can go straight to the record. It is
    # carried through to the report and never reaches the final .xlsx files.
    df[SOURCE_ROW_COL] = [str(header_idx + 2 + i) for i in range(len(df))]
    return df, ext_county, ext_state, label


def extract_valid_sheets(filepath: str) -> list[tuple[pd.DataFrame, str, str, str]]:
    """Open an .xlsx/.xls and return, for every sheet that looks like a real runsheet, its
    table (header already applied) plus the county/state detected in the preamble."""
    try:
        xls = pd.ExcelFile(filepath)
    except Exception as e:
        print(f"Error opening {filepath}: {e}")
        return []

    sheets_data = []
    for sheet in xls.sheet_names:
        df_raw = pd.read_excel(xls, sheet_name=sheet, header=None)
        result = frame_from_raw_table(df_raw, f"{filepath} [{sheet}]")
        if result:
            sheets_data.append(result)

    return sheets_data


def extract_valid_tables_from_pdf(filepath: str) -> list[tuple[pd.DataFrame, str, str, str]]:
    """Read a runsheet that arrived as a PDF. Every table on every page is treated as its
    own sheet: these runsheets repeat the header on each page, and the number of columns
    the extractor detects can vary from page to page, so mapping headers page by page is
    safer than carrying the first page's mapping forward.

    Requires pdfplumber (pip install pdfplumber)."""
    try:
        import pdfplumber
    except ImportError:
        print(f"Skipping {filepath}: reading PDFs requires 'pdfplumber' (pip install pdfplumber)")
        return []

    tables_data = []
    carried_county, carried_state = "", ""
    try:
        with pdfplumber.open(filepath) as pdf:
            for page_no, page in enumerate(pdf.pages, start=1):
                for t_no, table in enumerate(page.extract_tables()):
                    if not table or len(table) < 2:
                        continue
                    # inside a PDF cell, line breaks are just the table's text wrapping,
                    # not data separators: collapse them to spaces
                    table = [
                        [re.sub(r"\s+", " ", c).strip() if isinstance(c, str) else c for c in row]
                        for row in table
                    ]
                    df_raw = pd.DataFrame(table)
                    label = f"{filepath} [page {page_no}]"
                    result = frame_from_raw_table(df_raw, label)
                    if not result:
                        continue
                    df, county, state, lbl = result
                    # the county is usually only in the first page's preamble
                    if county:
                        carried_county, carried_state = county, state
                    tables_data.append((df, county or carried_county, state or carried_state, lbl))
    except Exception as e:
        print(f"Error opening {filepath}: {e}")
        return []

    # if the county only turned up on a later page, apply it backwards as well
    if carried_county:
        tables_data = [(df, c or carried_county, s or carried_state, lbl) for df, c, s, lbl in tables_data]
    return tables_data


def map_sheet_columns(df: pd.DataFrame) -> dict[str, str]:
    """Match each runsheet column to its logical field (grantor, book, inst_date...) using
    HEADER_ALIASES, so runsheets with different headers can all be read."""
    mapped: dict[str, str] = {}
    for c in df.columns:
        c_clean = clean_column_name(c)
        for field, aliases in HEADER_ALIASES.items():
            # the FIRST matching column wins: if a runsheet carries both "Date Executed"
            # and "Effective Date", the field keeps the first one and a later column
            # cannot steal its meaning
            if c_clean in aliases and field not in mapped:
                mapped[field] = c
                break
    return mapped


def build_notes(df: pd.DataFrame, main_cols: set, temp_df: pd.DataFrame, remarks: pd.Series) -> None:
    """Fill 'Notes' (every runsheet field without a column of its own, as "Field: value")
    and 'Essences', plus the automatic Judgment/Agreement notes, row by row."""
    extra_cols = [c for c in df.columns if c not in main_cols]
    for idx in range(len(temp_df)):
        notes_parts = []
        for c in extra_cols:
            val = df[c].iloc[idx]
            # a date sitting in an unmapped column is written as MM/DD/YYYY, not as
            # pandas' raw "1992-12-31 00:00:00"
            val_str = (clean_specific_dates(val) if isinstance(val, (datetime, pd.Timestamp, date))
                       else str(val).strip())
            if pd.notna(val) and val_str and val_str.lower() != "nan" and val_str != "-":
                notes_parts.append(f"{c}: {val_str}")

        instr_val = temp_df["Original_Instrument"].iloc[idx]
        instr_val = str(instr_val) if pd.notna(instr_val) else ""
        if instr_val:
            notes_parts.append(f"Original Instrument: {instr_val}")
            instr_upper = instr_val.upper()

            if len(instr_val) >= 75 or "CONSOLIDATED DT" in instr_upper or "SUPPLEMENT TO DT" in instr_upper:
                notes_parts.append("Check instrument couldnt determine what it is form the runsheet import")
            if any(j in instr_upper for j in ("JUDG", "JUDGMENT")):
                notes_parts.append("Note: This document is a Judgment.")
            if any(a in instr_upper for a in ("AGREE", "TRUST AGREE", "STIP", "JOA", "CONTRACT")):
                notes_parts.append("Note: This document is an Agreement.")

        temp_df.at[idx, "Notes"] = "\n\n".join(notes_parts)
        rv = remarks.iloc[idx]
        temp_df.at[idx, "Essences"] = str(rv) if pd.notna(rv) and str(rv) != "-" else ""


def load_and_merge_excels() -> pd.DataFrame:
    """Read every runsheet in the current folder (Excel or PDF, skipping the output files
    of a previous run) and build a single DataFrame in the intermediate schema, ready to
    be classified."""
    all_files = [
        f for f in glob.glob("*.xlsx") + glob.glob("*.xls") + glob.glob("*.pdf")
        if f not in OUTPUT_FILES
    ]

    df_list = []
    for file in sorted(all_files):
        if file.startswith("~$"):
            continue
        if file.upper().endswith(OUTPUT_FILE_SUFFIXES):
            print(f"Skipping {file} (output filename of a normalized file -- not reprocessed)")
            continue
        try:
            if file.lower().endswith(".pdf"):
                valid_sheets = extract_valid_tables_from_pdf(file)
            else:
                valid_sheets = extract_valid_sheets(file)
            if not valid_sheets:
                print(f"Skipping {file} (no runsheet tables detected)")
                continue

            for df, fallback_county, fallback_state, source_label in valid_sheets:
                df = merge_split_rows(df)
                cols = map_sheet_columns(df)
                main_cols = set(cols.values()) | {SOURCE_ROW_COL}

                temp_df = pd.DataFrame(columns=TARGET_COLS)
                temp_df["Original_Instrument"] = ""

                if "instrument" in cols:
                    temp_df["Original_Instrument"] = df[cols["instrument"]]
                if "grantor" in cols:
                    temp_df["Grantor"] = df[cols["grantor"]]
                if "grantee" in cols:
                    temp_df["Grantee"] = df[cols["grantee"]]
                if "inst_date" in cols:
                    temp_df["Inst.Date"] = df[cols["inst_date"]].apply(clean_specific_dates)
                if "filed_date" in cols:
                    temp_df["Filed"] = df[cols["filed_date"]].apply(clean_specific_dates)
                if "recorded_date" in cols:
                    temp_df["Recorded"] = df[cols["recorded_date"]].apply(clean_specific_dates)
                if "inst_no" in cols:
                    temp_df["Inst.No."] = df[cols["inst_no"]]
                if "records_type" in cols:
                    # helper column: used for Booktype and dropped at the end
                    temp_df["Records_Type"] = df[cols["records_type"]]

                if "book" in cols and "page" in cols:
                    parsed = df[cols["book"]].apply(clean_book)
                    temp_df["Booktype_Hint"] = parsed.apply(lambda x: x[0])
                    temp_df["Book"] = parsed.apply(lambda x: x[1])
                    temp_df["Page"] = df[cols["page"]].apply(clean_page)
                elif "book_page" in cols:
                    parsed = df[cols["book_page"]].apply(parse_book_page)
                    temp_df["Booktype_Hint"] = parsed.apply(lambda x: x[0])
                    temp_df["Book"] = parsed.apply(lambda x: x[1])
                    temp_df["Page"] = parsed.apply(lambda x: x[2])
                    # whatever was not a book/page (a cause number, a document number...)
                    # is kept in Inst.No. instead of being lost or polluting Book/Page
                    refs = parsed.apply(lambda x: x[3])
                    if "inst_no" in cols:
                        temp_df["Inst.No."] = [
                            existing if str(existing).strip() not in ("", "nan", "None") else ref
                            for existing, ref in zip(temp_df["Inst.No."], refs)
                        ]
                    else:
                        temp_df["Inst.No."] = refs

                if "county" in cols:
                    county_col = df[cols["county"]]
                    temp_df["County"] = county_col.apply(lambda x: str(x).split(",")[0].strip() if pd.notna(x) else "")
                    temp_df["State"] = county_col.apply(lambda x: str(x).split(",")[1].strip() if pd.notna(x) and "," in str(x) else "")
                else:
                    temp_df["County"] = fallback_county
                    temp_df["State"] = fallback_state

                remarks = df[cols["remarks"]] if "remarks" in cols else pd.Series([""] * len(df))
                build_notes(df, main_cols, temp_df, remarks)

                # never leave Grantor/Grantee blank -- flag it instead, so a reviewer can
                # go straight to the source row (same convention as the PDF/txt pipeline)
                for col in ("Grantor", "Grantee"):
                    temp_df[col] = temp_df[col].apply(
                        lambda v: v if pd.notna(v) and str(v).strip() not in ("", "nan")
                        else UNCLEAR_MARKER
                    )

                # helper column: which file/sheet/page each row came from. It never
                # reaches the final files (the software expects a fixed schema), but the
                # missing-fields report uses it to trace a row back to its source.
                temp_df["Source_File"] = source_label
                if SOURCE_ROW_COL in df.columns:
                    # after merge_split_rows a merged row carries several numbers ("6 7")
                    temp_df["Source_Row"] = [
                        re.sub(r"\s+", "-", str(v).strip()) for v in df[SOURCE_ROW_COL]
                    ]

                df_list.append(temp_df)
            print(f"File processed successfully: {file}")

        except Exception as e:
            print(f"Error processing file {file}: {e}")

    if not df_list:
        df_empty = pd.DataFrame(columns=TARGET_COLS)
        df_empty["Original_Instrument"] = ""
        return df_empty

    final_df = pd.concat(df_list, ignore_index=True)

    # rows that are really repeated headers (every PDF page carries its own)
    final_df = final_df[
        ~(
            final_df["Grantor"].astype(str).str.strip().str.upper().isin(["GRANTOR", "GRANTORS"])
            | final_df["Grantee"].astype(str).str.strip().str.upper().isin(["GRANTEE", "GRANTEES"])
            | final_df["Original_Instrument"].astype(str).str.replace(r"[^A-Z]", "", regex=True).isin(
                ["INSTRUMENTTITLE", "INSTRUMENT", "IMGINSTRUMENT", "INSTNAME", "INSTRUMENTTYPE"]
            )
        )
    ].reset_index(drop=True)

    final_df = drop_empty_rows(final_df)
    if DROP_DUPLICATES:
        final_df = drop_duplicate_records(final_df)

    valid_county = final_df["County"].replace("", np.nan).dropna()
    valid_state = final_df["State"].replace("", np.nan).dropna()
    final_df["County"] = valid_county.iloc[0] if not valid_county.empty else ""
    final_df["State"] = valid_state.iloc[0] if not valid_state.empty else ""

    return final_df


def _is_blank(value: object) -> bool:
    return pd.isna(value) or str(value).strip().lower() in ("", "nan", "none", "-", "n/a")


def drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Drop rows with no usable data at all (separator and filler rows, which show up
    mostly in tables extracted from PDFs)."""
    if df.empty:
        return df
    key_fields = ["Grantor", "Grantee", "Original_Instrument", "Book", "Page", "Inst.Date",
                  "Notes", "Essences"]
    present = [c for c in key_fields if c in df.columns]

    def has_content(row: pd.Series) -> bool:
        for c in present:
            v = row[c]
            if c in ("Grantor", "Grantee") and str(v).strip() == UNCLEAR_MARKER:
                continue
            if not _is_blank(v):
                return True
        return False

    keep = df.apply(has_content, axis=1)
    dropped = int((~keep).sum())
    if dropped:
        print(f"Empty rows dropped: {dropped}")
    return df[keep].reset_index(drop=True)


def _note_entries(value: object) -> frozenset:
    """The "Field: value" entries of the Notes column, as a set -- used to compare how much
    information one row carries against another."""
    if _is_blank(value):
        return frozenset()
    parts = re.split(r"\n\s*\n", str(value))
    return frozenset(re.sub(r"\s+", " ", p).strip().upper() for p in parts if p.strip())


def drop_duplicate_records(df: pd.DataFrame) -> pd.DataFrame:
    """The same instrument often shows up more than once: 'MainRunsheet' + 'IndexSummary'
    sheets inside one file, or several overlapping supplemental runsheets.

    Only the copy that adds NOTHING new is dropped: same parties, same instrument, same
    book/page and same date, AND with its notes contained in those of the row being kept.
    That removes the poorer copies (the 'IndexSummary' repeating the 'MainRunsheet') while
    respecting rows that share an instrument and still say something different -- for
    example the same deed listed under Tract A and Tract B of a title opinion, where each
    row describes a different tract."""
    if df.empty:
        return df

    def norm(value: object) -> str:
        return re.sub(r"\s+", " ", str(value)).strip().upper() if not _is_blank(value) else ""

    key_cols = ["Grantor", "Grantee", "Original_Instrument", "Book", "Page", "Inst.Date"]
    present = [c for c in key_cols if c in df.columns]
    keys = df[present].map(norm).agg("|".join, axis=1)
    notes = df["Notes"].map(_note_entries) if "Notes" in df.columns else pd.Series(
        [frozenset()] * len(df), index=df.index
    )
    filled = df.map(lambda v: 0 if _is_blank(v) else 1).sum(axis=1)

    kept_by_key: dict[str, list[int]] = {}
    keep_idx: list[int] = []
    for i in filled.sort_values(ascending=False).index:       # most complete rows first
        k = keys.loc[i]
        mine = notes.loc[i]
        # dropped only if its notes are already contained in a row that was kept
        if any(mine <= notes.loc[j] for j in kept_by_key.get(k, [])):
            continue
        kept_by_key.setdefault(k, []).append(i)
        keep_idx.append(i)

    removed = len(df) - len(keep_idx)
    if removed:
        print(f"Duplicate records dropped: {removed}")
    return df.loc[sorted(keep_idx)].reset_index(drop=True)


def get_year(date_val: object) -> int | None:
    if pd.isna(date_val) or date_val == "":
        return None
    match = re.search(r"\d{4}", str(date_val))
    return int(match.group()) if match else None


def assign_booktype(row: pd.Series) -> str:
    """Priority: 1) an explicit 'Records Type' column; 2) the prefix attached to the volume
    ("DR 104/287", "OR 165"); 3) the year as a last resort (before 1990 -> DR)."""
    if "Records_Type" in row and pd.notna(row["Records_Type"]):
        rt_val = str(row["Records_Type"]).strip().upper()
        if "DEED" in rt_val:
            return "DR"
        if "OFFICIAL" in rt_val or "OPR" in rt_val or "PUBLIC" in rt_val:
            return "OPR"

    if "Booktype_Hint" in row and pd.notna(row["Booktype_Hint"]):
        hint = str(row["Booktype_Hint"]).strip().upper()
        if hint in ("DR", "OPR"):
            return hint

    year = get_year(row["Inst.Date"]) or get_year(row["Filed"])
    if year:
        return "DR" if year < 1990 else "OPR"
    return ""


SUBSTRING_RULES = {
    "ASSIGNMENT OF OIL AND GAS LEASE": "AOGL",
    "PARTIAL ASSIGNMENT OF OIL AND GAS LEASE": "PART AOGL",
    "ASSIGNMENT OF OIL, GAS, AND MINERAL LEASE": "AOGML",
    "PARTIAL ASSIGNMENT OF OIL, GAS, AND MINERAL LEASE": "PART AOGML",
    "ASSIGNMENT OF OVERRIDING ROYALTY INTEREST": "AORRI",
    "ASSIGNMENT OF OVERRIDING ROAYLTY INTEREST": "AORRI",
    "ASSIGNMENT OF DEED OF TRUST": "ASSIGN DT",
    "ASSIGNMENT OF PRODUCTION": "DT",
    "ASSIGNMENT OF AS-EXTRACTED COLLATERAL": "DT",
    "BILL OF SALE AND ASSIGNMENT": "ABOS",
    "ASSIGNMENT AND BILL OF SALE": "ABOS",
    "PARTIAL ASSIGNMENT": "PART ASSIGN",
    "ASSIGNMENT": "ASSIGN",
    "BLANKET MINERAL AND ROYALTY DEED": "MD",
    "MINERAL AND ROYALTY DEED": "MRD",
    "MINERAL/ROYALTY DEED": "MRD",
    "MINERAL AND ROAYLTY INTEREST DEED": "MRD",
    "CORRECTION MINERAL AND ROYALTY DEED": "MRD",
    "MINERAL DEED": "MD",
    "INDEPENDENT CO-EXECUTOR'S DISTRIBUTION DEED": "DIST DEED",
    "DISTRIBUTION DEED": "DIST DEED",
    "TRUSTEE'S DISTRIBUTION DEED": "DIST DEED",
    "DEED OF DISTRIBUTION": "DIST DEED",
    "AFFIDAVIT OF HEIRSHIP": "AFFT HEIR",
    "AFF OF HEIRSHIP": "AFFT HEIR",
    "AFFIDAVIT OF SUCCESSOR TRUSTEE": "AFFT SUCCESSOR TRST",
    "AFFIDAVIT": "AFFT",
    "SURFACE USE AND EASEMENT AGREEMENT": "AGREE",
    "EASEMENT": "EAS",
    "DEED OF TRUST": "DT",
    "QUIT CLAIM": "QCD",
    "GENERAL WARRANTY DEED": "GWD",
    "SPECIAL WARRANTY DEED": "SWD",
    "WARRANTY DEED WITH VENDOR'S LIEN": "WDVL",
    "WARRANTY DEED": "WD",
    "GIFT DEED": "GIFT DEED",
    "ADMINISTRATOR'S DEED": "ADMIN DEED",
    "EXECUTOR'S DEED": "EXEC DEED",
    "PARTITION DEED": "PART DEED",
    "PERSONAL REPRESENTATIVE'S DEED": "PR DEED",
    "RIGHT OF WAY DEED": "ROW DEED",
    "SHERIFF'S DEED": "SHRF DEED",
    "SUCCESSOR TRUSTEE'S DEED": "SUCCESSOR TRST DEED",
    "TRUSTEE'S DEED": "TRST DEED",
    "ROYALTY DEED": "RD",
    "DEED": "DEED",
    "BILL OF SALE": "BOS",
    "PARTIAL RELEASE OF LIEN": "RELL OF LIEN",
    "RELEASE OF LIEN": "RELL OF LIEN",
    "LIEN": "RELL OF LIEN",
    "RELEASE": "REL",
    "AMENDED AND RESTATED": "AMEND",
    "AMEND": "AMEND",
    "MEMORANDUM OF OIL AND GAS LEASE": "MEMO OGL",
    "MEMORANDUM OF OIL, GAS, AND MINERAL LEASE": "MEMO OGML",
    "OIL AND GAS LEASE": "OGL",
    "OIL & GAS LEASE": "OGL",
    "OIL, GAS, AND MINERAL LEASE": "OGML",
    "CC PROBATE": "PROBATE",
    "CC WILL": "PROBATE",
    "CC WILL / ORDER": "PROBATE",
    "WILL & PROBATE": "PROBATE",
    "CC/PROBATE": "PROBATE",
    "PROBATE": "PROBATE",
    "ESTATE": "PROBATE",
    "AGREEMENT": "AGREE",
    "CONVEYANCE": "CONV",
    # --- instruments seen in real runsheets that used to fall through to MISC ---
    "AFFIDAVIT OF IDENTITY": "AFFT IDENT",
    "AFFIDIVIT OF IDENTITY": "AFFT IDENT",
    "AFFIDAVIT OF NAME CHANGE": "AFFT NAME",
    "AFFIDAVIT OF NON-PRODUCTION": "AFFT NONPROD",
    "AFF OF NON-PRODUCTION": "AFFT NONPROD",
    "AFFIDAVIT OF FACT": "AFFT FACT",
    "AFFIDAVIT OF DEATH": "AFFT DEATH",
    "LIMITED POWER OF ATTORNEY": "LTD POA",
    "SPECIAL POWER OF ATTORNEY": "SPEC POA",
    "POWER OF ATTORNEY": "POA",
    "CERTIFICATE OF TRUST": "CERT TRUST",
    "CERTIFICATE OF INCUMBENCY": "CERT INCUMB",
    "MEMORANDUM OF MERGER": "MERGER",
    "ARTICLES OF MERGER": "MERGER",
    "MERGER": "MERGER",
    "ARTICLES OF CONVERSION": "CONVERSION",
    "STATEMENT OF AUTHORITY": "STMT AUTH",
    "RESIGNATION OF TRUSTEE": "RESIG TRST",
    "APPOINTMENT OF SUCCESSOR TRUSTEE": "APPT SUCCESSOR TRST",
    "RATIFICATION": "RATIF",
    "LIS PENDENS": "LIS PENDENS",
    "MARRIAGE LICENSE": "MARRIAGE LIC",
    "ORDER OF DISMISSAL": "ORDER",
    "ORDER": "ORDER",
    "DECREE OF DISTRIBUTION": "DIST DEED",
    "FINAL DECREE OF DIVORCE": "DIVORCE DECREE",
    "DIVORCE": "DIVORCE DECREE",
    "DECREE": "DECREE",
    "MEMORANDUM OF OPERATING AGREEMENT": "MEMO OA",
    "OPERATING AGREEMENT": "JOA",
    "MEMORANDUM OF PRODUCTION": "MEMO PROD",
    "SUBORDINATION": "SUB",
    "NOTICE": "NOTICE",
    "CORRECTION": "CORR",
    "SUPPLEMENT": "SUPP",
    "PARTITION": "PART DEED",
    "STIPULATION": "STIP",
    "EASEMENT AND RIGHT OF WAY": "EAS",
    "RIGHT OF WAY": "ROW",
    "PATENT": "PATENT",
}
SORTED_SUBSTRING_RULES = sorted(SUBSTRING_RULES.items(), key=lambda x: len(x[0]), reverse=True)

EXACT_RULES = {
    r"\bAOGL\b": "AOGL", r"\bAOGML\b": "AOGML", r"\bAORRI\b": "AORRI", r"\bASORI\b": "AORRI",
    r"\bASSIGN DT\b": "ASSIGN DT", r"\bABOS\b": "ABOS", r"\bASSN\b": "ASSIGN", r"\bASGN\b": "ASSIGN",
    r"\bASN\b": "ASSIGN", r"^AS$": "ASSIGN", r"\bPART AOGL\b": "PART AOGL", r"\bPART AOGML\b": "PART AOGML",
    r"\bPART ASSIGN\b": "PART ASSIGN", r"\bPART ASGN\b": "PART ASSIGN", r"\bMD\b": "MD", r"\bMRD\b": "MRD",
    r"\bDIST DEED\b": "DIST DEED", r"\bAFFT HEIR\b": "AFFT HEIR", r"\bAFF H\b": "AFFT HEIR", r"\bAFFT\b": "AFFT",
    r"\bAGREE\b": "AGREE", r"\bAGMT\b": "AGREE", r"\bDOTO\b": "DT", r"\bD OF TR\b": "DT", r"\bDT\b": "DT",
    r"\bQCD\b": "QCD", r"\bWD\b": "WD", r"\bCOR WD\b": "WD", r"\bCORRECTION WD\b": "WD", r"\bWD RECORD\b": "WD", r"\bWD/ VL\b": "WD",
    r"\bGWD\b": "GWD", r"\bSWD\b": "SWD", r"\bWDVL\b": "WDVL", r"\bBOS\b": "BOS", r"\bCONV\b": "CONV",
    r"\bREL\b": "REL", r"\bREL LN\b": "RELL OF LIEN", r"\bREL. OL\b": "REL",
    r"\bAMD\b": "AMEND", r"\bOGL\b": "OGL", r"\bOGML\b": "OGML", r"\bPROB\b": "PROBATE", r"\bCC PROB\b": "PROBATE",
    r"\bWILL\b": "PROBATE", r"\bPATENT\b": "PATENT", r"\bHEIRSHIP\b": "AFFT HEIR", r"\bO&GL\b": "OGL", r"\bOG&ML\b": "OGML", r"\bMOGL\b": "OGL",
    r"\bJDG\b": "JUDGMENT",
    # abbreviations seen in real runsheets
    r"\bAOH\b": "AFFT HEIR", r"\bAFF IDENT\b": "AFFT IDENT", r"\bAFF ID\b": "AFFT IDENT",
    r"\bPOA\b": "POA", r"\bLPOA\b": "LTD POA", r"\bSPOA\b": "SPEC POA",
    r"\bDOT\b": "DT", r"\bMEMO OA\b": "MEMO OA", r"\bJOA\b": "JOA",
    r"\bPAOGL[S]?\b": "PART AOGL", r"\bMEMO\b": "MEMO", r"\bSUB\b": "SUB",
    r"\bROW\b": "ROW", r"\bEAS\b": "EAS", r"\bLIS\b": "LIS PENDENS",
    # common typos made while typing the runsheet
    r"\bASSGIN\b": "ASSIGN", r"\bASIGN\b": "ASSIGN", r"\bASSIGMENT\b": "ASSIGN",
    r"\bCONVEY\b": "CONV", r"\bAFFIDIVIT\b": "AFFT", r"\bAFFADAVIT\b": "AFFT",
}

CONVEYANCES = {"MD", "WD", "GWD", "SWD", "QCD", "DEED", "CONV", "BOS", "PROBATE", "AFFT HEIR", "MRD", "DIST DEED", "GIFT DEED", "ADMIN DEED", "EXEC DEED", "PART DEED", "PR DEED", "SHRF DEED", "SUCCESSOR TRST DEED", "TRST DEED", "RD", "ROW DEED", "WDVL", "PATENT"}
ASSIGNMENTS = {"ASSIGN", "ABOS", "AOGL", "AOGML", "PART ASSIGN", "PART AOGL", "PART AOGML", "AORRI", "ASSIGN DT", "BLM ASSIGN", "BLM TRANSFER", "TOD"}
LEASES = {"OGL", "OGML"}
RELEASES = {"REL"}
EASEMENTS = {"EAS", "ROW", "ROW/EAS", "EASEMENT"}
D_OF_TRUST = {"DT", "MTG", "SUB", "DEED OF TRUST", "D OF TR"}
LIENS = {"MML", "FTL", "STL", "HL", "AJ", "RELL OF LIEN", "LIEN"}
AGREEMENTS = {"AGREE", "TRUST AGREE", "STIP", "JOA", "RDO", "AGREEMENT", "CONTRACT"}
JUDGEMENTS = {"JUDG", "JUDGMENT"}


def _subtype_matches(subtype: str, bucket: set) -> bool:
    return any(v == subtype or f" {v} " in f" {subtype} " for v in bucket)


def assign_subtype_and_type(instrument: object) -> tuple[str, str, str]:
    if pd.isna(instrument) or not str(instrument).strip():
        return "MISC", "MISC.", "NCD"

    instr = str(instrument).upper().strip()
    subtype = ""

    for phrase, sub_val in SORTED_SUBSTRING_RULES:
        if phrase in instr:
            subtype = sub_val
            break

    if not subtype:
        for pattern, sub_val in EXACT_RULES.items():
            if re.search(pattern, instr):
                subtype = sub_val
                break

    matched_by_rule = bool(subtype)
    if not subtype:
        subtype = instr.split("\n")[0].strip()

    if _subtype_matches(subtype, CONVEYANCES) or "PATENT" in instr:
        return subtype, "Conveyance", "CD"
    if _subtype_matches(subtype, ASSIGNMENTS):
        return subtype, "Assignment", "CD"
    if _subtype_matches(subtype, LEASES):
        return subtype, "Lease", "CD"
    if _subtype_matches(subtype, RELEASES):
        return subtype, "Release", "CD"

    doc_class = "NCD"
    if _subtype_matches(subtype, D_OF_TRUST) or "DEED OF TRUST" in instr or "D OF TR" in instr or re.search(r"\bDT\b", instr):
        return "DT", "DEED OF TRUST", doc_class
    if _subtype_matches(subtype, EASEMENTS) or "EASEMENT" in instr:
        return "Easement", "Easement", doc_class
    if _subtype_matches(subtype, JUDGEMENTS) or "JUDG" in instr:
        return "JUD", "MISC.", doc_class
    if "AMEND" in subtype or "AMEND" in instr:
        return "AMENDM", "MISC.", doc_class
    if _subtype_matches(subtype, AGREEMENTS) or "AGREEMENT" in instr or "CONTRACT" in instr:
        return "AGREMNT", "AGRE", doc_class
    if _subtype_matches(subtype, LIENS) or "LIEN" in instr:
        return "RELL OF LIEN", "Lien", doc_class

    # No software "Type" fits, but that is NO reason to lose the Subtype: if a rule did
    # recognize the instrument (POA, AFFT, CERT TRUST, MERGER...), that abbreviation is
    # kept and only the Type falls back to MISC. Subtype becomes "MISC" only when no rule
    # recognized anything -- and the report then flags it for manual review.
    if matched_by_rule:
        return subtype, "MISC.", doc_class
    return "MISC", "MISC.", doc_class


def classify(df_final: pd.DataFrame) -> pd.DataFrame:
    if df_final.empty:
        df_final["Subtype"] = ""
        df_final["Type"] = "MISC."
        df_final["Class"] = "NCD"
        return df_final

    df_final["Booktype"] = df_final.apply(assign_booktype, axis=1)
    mapped_values = df_final["Original_Instrument"].apply(assign_subtype_and_type)
    df_final["Subtype"] = mapped_values.apply(lambda x: x[0])
    df_final["Type"] = mapped_values.apply(lambda x: x[1])
    df_final["Class"] = mapped_values.apply(lambda x: x[2])
    return df_final


def save_formatted_excel(df: pd.DataFrame, out_file: str) -> None:
    if df.empty:
        print(f"No records to write to {out_file}.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Normalized Runsheet"

    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(
        left=Side(style="thin", color="BFBFBF"), right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"), bottom=Side(style="thin", color="BFBFBF"),
    )

    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if r_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                cell.border = border
                cell.alignment = Alignment(vertical="top", wrap_text=True)

    widths = {"A": 10, "B": 15, "C": 20, "D": 25, "E": 25, "H": 15, "J": 15, "N": 15, "O": 15, "Q": 50, "S": 50}
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w

    wb.save(out_file)
    print(f"Done. File generated: {out_file} ({len(df)} records)")


def write_missing_fields_report(df: pd.DataFrame, out_path: str) -> None:
    """Plain-text report of every row with a blank/uncertain Grantor, Grantee, Book or
    Page, or whose instrument could not be classified (Subtype=MISC).

    It is GROUPED BY SOURCE FILE, and each record states which row of the original file it
    came from, so that runsheet can be opened and the data checked directly."""
    by_source: dict[str, list[list[str]]] = {}
    flagged = 0

    for idx, row in df.reset_index(drop=True).iterrows():
        issues = []
        grantor = str(row.get("Grantor", "")).strip()
        grantee = str(row.get("Grantee", "")).strip()
        book = str(row.get("Book", "")).strip()
        page = str(row.get("Page", "")).strip()
        subtype = str(row.get("Subtype", "")).strip()

        if not grantor or grantor.lower() == "nan" or grantor == UNCLEAR_MARKER:
            issues.append(f"Grantor: {grantor or '(empty)'}")
        if not grantee or grantee.lower() == "nan" or grantee == UNCLEAR_MARKER:
            issues.append(f"Grantee: {grantee or '(empty)'}")
        if not book or book.lower() == "nan":
            issues.append("Book/Volume: (empty)")
        if not page or page.lower() == "nan":
            issues.append("Page: (empty)")
        if subtype == "MISC":
            issues.append("Subtype: MISC (instrument not recognized by the classification rules)")

        if not issues:
            continue

        flagged += 1
        source = str(row.get("Source_File", "")).strip()
        if not source or source.lower() == "nan":
            source = "(unknown source)"
        src_row = str(row.get("Source_Row", "")).strip()
        pos = f"row {src_row}" if src_row and src_row.lower() != "nan" else "row ?"

        entry = [
            f"  [{pos} of the file]   Class: {row.get('Class', '')}"
            f"   Type: {row.get('Type', '')}   Subtype: {row.get('Subtype', '')}"
            f"   (record {idx + 1} of the output)",
            f"     Grantor: {grantor or '(empty)'}  ->  Grantee: {grantee or '(empty)'}",
            f"     Book/Page: {book or '(empty)'}/{page or '(empty)'}"
            f"   |  Inst.Date: {str(row.get('Inst.Date', '')).strip() or '(empty)'}",
        ]
        for issue in issues:
            entry.append(f"     - MISSING -> {issue}")
        notes = str(row.get("Notes", "")).strip()
        if notes and notes.lower() != "nan":
            entry.append(f"     Notes: {notes}")
        entry.append("")
        by_source.setdefault(source, []).append(entry)

    lines: list[str] = [
        "MISSING / UNCERTAIN FIELDS REPORT",
        f"Records processed: {len(df)}   |   Records flagged: {flagged}",
        "",
        "Records are grouped by the file they came from, and each one states which row of",
        "THAT file it sits in, so it can be checked directly against the source.",
        "=" * 78,
        "",
    ]

    if "Source_File" in df.columns:
        lines.append("SUMMARY BY SOURCE FILE (total records / flagged):")
        counts = df["Source_File"].value_counts()
        for src in counts.index:
            marked = len(by_source.get(str(src), []))
            lines.append(f"   {counts[src]:5d} records / {marked:4d} flagged   {src}")
        lines.append("")
        lines.append("=" * 78)
        lines.append("")

    if not by_source:
        lines.append("(No record was left incomplete.)")
    else:
        for source in sorted(by_source):
            entries = by_source[source]
            lines.append("#" * 78)
            lines.append(f"# FILE: {source}")
            lines.append(f"# {len(entries)} record(s) to review")
            lines.append("#" * 78)
            lines.append("")
            for entry in entries:
                lines.extend(entry)

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"Report generated: {out_path} ({flagged} records flagged)")


def main() -> None:
    df_full = classify(load_and_merge_excels())
    # the report uses the helper columns (Source_File/Source_Row) to trace each row; the
    # .xlsx files imported into the software carry only the fixed TARGET_COLS schema
    df_final = df_full[TARGET_COLS]

    df_cd = df_final[df_final["Class"] == "CD"].copy()
    df_ncd = df_final[df_final["Class"] == "NCD"].copy()

    print("\n--- GENERATING SEPARATE FILES ---")
    save_formatted_excel(df_cd, "Normalized_Runsheet_CD.xlsx")
    save_formatted_excel(df_ncd, "Normalized_Runsheet_NCD.xlsx")
    write_missing_fields_report(df_full, "Normalized_Runsheet_MissingFields.txt")


if __name__ == "__main__":
    main()
