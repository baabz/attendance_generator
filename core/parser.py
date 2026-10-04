"""
Parsers for extracting student records from Excel, CSV, Word (.docx), and text inputs.
Includes smart format auto-detection for HTML-exported Excel sheets, framesets, legacy .xls files,
and automatic column correction for portal score sheets.
"""
import io
import os
import re
import urllib.parse
from typing import List, Optional, Tuple, Union
import pandas as pd
import docx
from .models import Student


def clean_header_text(val) -> str:
    """Normalize header text for robust keyword matching, removing common school banner words."""
    text = str(val).lower()
    text = text.replace('regular', '').replace('register', '').replace('programme', '').replace('scoresheet', '').replace('session', '')
    return re.sub(r'[^a-z0-9]', '', text)


def detect_columns(columns: List[str]) -> Tuple[Optional[str], Optional[str]]:
    """
    Attempt to auto-detect REG NO and NAME columns from a list of dataframe headers.
    """
    reg_col = None
    name_col = None

    reg_keywords = ['regno', 'regnumber', 'regnum', 'reg', 'matric', 'adminno', 'admissionno', 'studentid', 'idno']
    name_keywords = ['name', 'fullname', 'studentname', 'candidatesname', 'names']

    for orig_col in columns:
        clean = clean_header_text(orig_col)
        # Avoid column if it is pure S/N or #
        if clean in ['sn', 'no', 'sno', 'number']:
            continue
        if reg_col is None and any(k in clean for k in reg_keywords):
            reg_col = orig_col
            continue
        if name_col is None and any(k in clean for k in name_keywords):
            name_col = orig_col
            continue

    return reg_col, name_col


def normalize_dataframe(df: pd.DataFrame) -> Tuple[pd.DataFrame, Optional[str], Optional[str]]:
    """
    Auto-detect if headers are in row 0, or buried below institutional title banners
    (common in school mark sheets and portal exports).
    """
    df = df.dropna(how='all').reset_index(drop=True)
    if df.empty:
        return df, None, None

    has_unnamed = any('unnamed' in str(c).lower() for c in df.columns)

    # If all columns are properly named (no Unnamed: x), check them first
    if not has_unnamed:
        reg_col, name_col = detect_columns(list(df.columns))
        if reg_col and name_col:
            return df, reg_col, name_col

    # Scan the first 15 rows for the true header row
    for r_idx in range(min(15, len(df))):
        row_vals = [str(x) for x in df.iloc[r_idx]]
        r_reg, r_name = detect_columns(row_vals)
        if r_reg and r_name:
            header_series = df.iloc[r_idx]
            new_df = df.iloc[r_idx + 1:].copy().reset_index(drop=True)
            new_df.columns = [str(val).strip() for val in header_series]
            reg_col, name_col = detect_columns(list(new_df.columns))
            return new_df, reg_col, name_col

    # Fallback to column detection on existing df
    reg_col, name_col = detect_columns(list(df.columns))
    return df, reg_col, name_col


def smart_parse_spreadsheet(content_bytes: bytes, original_filename: str = "") -> pd.DataFrame:
    """
    Robust spreadsheet parser handling:
    1. HTML Excel framesets (linking to companion folder sheet001.htm)
    2. HTML tables disguised as .xls/.xlsx
    3. Modern Excel (.xlsx via openpyxl)
    4. Legacy binary Excel (.xls via xlrd)
    5. CSV / TSV
    """
    text_preview = content_bytes[:4000].decode("utf-8", errors="ignore")

    # Check for HTML / XML signatures
    if "<html" in text_preview.lower() or "<table" in text_preview.lower() or "<body" in text_preview.lower():
        # Case A: Table embedded directly in HTML
        try:
            tables = pd.read_html(io.BytesIO(content_bytes), flavor="lxml")
            if tables:
                return tables[0]
        except Exception:
            pass

        # Case B: HTML Frameset referencing sheet001.htm
        sheet_match = re.search(r'href=["\']([^"\']*sheet00\d\.htm)["\']', text_preview, re.IGNORECASE)
        if sheet_match:
            rel_href = urllib.parse.unquote(sheet_match.group(1))
            folder_name = os.path.dirname(rel_href)
            file_name = os.path.basename(rel_href)

            search_roots = [
                os.getcwd(),
                r"C:\Users\user\Desktop",
                os.path.expanduser("~/Desktop"),
                os.path.expanduser("~/Downloads")
            ]
            for root_dir in search_roots:
                if not os.path.exists(root_dir):
                    continue
                direct_path = os.path.join(root_dir, rel_href)
                if os.path.exists(direct_path):
                    tables = pd.read_html(direct_path, flavor="lxml")
                    if tables:
                        return tables[0]

                for root, dirs, _ in os.walk(root_dir):
                    if folder_name and folder_name in dirs:
                        full_sheet = os.path.join(root, folder_name, file_name)
                        if os.path.exists(full_sheet):
                            tables = pd.read_html(full_sheet, flavor="lxml")
                            if tables:
                                return tables[0]

    # Try openpyxl for modern .xlsx
    try:
        return pd.read_excel(io.BytesIO(content_bytes), engine="openpyxl")
    except Exception:
        pass

    # Try xlrd for legacy .xls
    try:
        return pd.read_excel(io.BytesIO(content_bytes), engine="xlrd")
    except Exception:
        pass

    # Try HTML parser fallback
    try:
        tables = pd.read_html(io.BytesIO(content_bytes), flavor="html5lib")
        if tables:
            return tables[0]
    except Exception:
        pass

    # Try CSV fallback
    try:
        return pd.read_csv(io.BytesIO(content_bytes))
    except Exception:
        pass

    raise ValueError("Excel file format could not be parsed automatically. Please ensure the file is a valid .xlsx, .xls, .csv, or .htm file.")


def is_likely_reg_no(val: str) -> bool:
    """Check if value has registration number characteristics (alphanumeric or session slash)."""
    v = str(val).strip().upper()
    if re.search(r'[A-Z]', v) and re.search(r'\d', v):
        return True
    if re.search(r'^\d{2}/\d{3,5}$', v):  # e.g. 22/1395
        return True
    return False


def is_pure_int_series(series: pd.Series) -> bool:
    """Check if series consists almost entirely of pure integers (serial numbers)."""
    sample = series.dropna().astype(str).str.strip()[:20]
    if len(sample) == 0:
        return False
    int_count = sum(1 for x in sample if x.isdigit())
    return (int_count / len(sample)) >= 0.75


def parse_dataframe(df: pd.DataFrame, reg_col: Optional[str] = None, name_col: Optional[str] = None) -> List[Student]:
    """
    Convert a pandas DataFrame into a list of Student objects.
    Includes smart auto-correction if S/N was mistakenly mapped as Reg No.
    """
    df, detected_reg, detected_name = normalize_dataframe(df)
    if df.empty:
        return []

    cols = list(df.columns)
    reg_col = reg_col or detected_reg
    name_col = name_col or detected_name

    if not reg_col:
        reg_col = cols[1] if len(cols) > 1 else cols[0]
    if not name_col:
        name_col = cols[2] if len(cols) > 2 else (cols[1] if len(cols) > 1 else cols[0])

    # Auto-correction check: If reg_col is pure integers (S/N) and name_col contains Reg Nos
    if reg_col in df.columns and name_col in df.columns:
        if is_pure_int_series(df[reg_col]):
            sample_names = df[name_col].dropna().astype(str).str.strip()[:10]
            if any(is_likely_reg_no(x) for x in sample_names):
                # reg_col was actually S/N, name_col is actual REG NO!
                real_reg_idx = cols.index(name_col)
                reg_col = name_col
                # The actual student name is the next column
                if real_reg_idx + 1 < len(cols):
                    name_col = cols[real_reg_idx + 1]

    students = []
    idx = 1
    for _, row in df.iterrows():
        reg_val = str(row.get(reg_col, '')).strip()
        name_val = str(row.get(name_col, '')).strip()

        if not reg_val or reg_val.lower() == 'nan' or not name_val or name_val.lower() == 'nan':
            continue
        if 'reg' in reg_val.lower() and ('no' in reg_val.lower() or 'number' in reg_val.lower()):
            continue
        if reg_val.startswith('#') or reg_val.lower() == 's/n':
            continue

        students.append(Student(reg_no=reg_val, name=name_val, original_sn=idx))
        idx += 1

    return students


def parse_excel(file_bytes_or_path, reg_col: Optional[str] = None, name_col: Optional[str] = None) -> List[Student]:
    """Parse an Excel (.xlsx, .xls) file or HTML spreadsheet."""
    if isinstance(file_bytes_or_path, str):
        with open(file_bytes_or_path, 'rb') as f:
            content_bytes = f.read()
            filename = os.path.basename(file_bytes_or_path)
    elif hasattr(file_bytes_or_path, 'read'):
        content_bytes = file_bytes_or_path.read()
        filename = getattr(file_bytes_or_path, 'name', '')
    else:
        content_bytes = file_bytes_or_path
        filename = ''

    df = smart_parse_spreadsheet(content_bytes, filename)
    return parse_dataframe(df, reg_col, name_col)


def parse_csv(file_bytes_or_path, reg_col: Optional[str] = None, name_col: Optional[str] = None) -> List[Student]:
    """Parse a CSV or TSV file."""
    if isinstance(file_bytes_or_path, str):
        with open(file_bytes_or_path, 'rb') as f:
            content_bytes = f.read()
    elif hasattr(file_bytes_or_path, 'read'):
        content_bytes = file_bytes_or_path.read()
    else:
        content_bytes = file_bytes_or_path

    try:
        df = pd.read_csv(io.BytesIO(content_bytes))
    except Exception:
        df = pd.read_csv(io.BytesIO(content_bytes), sep=None, engine='python')
    return parse_dataframe(df, reg_col, name_col)


def docx_to_dataframe(file_bytes_or_path) -> pd.DataFrame:
    """Extract table rows from a docx file into a pandas DataFrame."""
    doc = docx.Document(file_bytes_or_path)
    all_rows = []
    for table in doc.tables:
        for row in table.rows:
            row_data = [c.text.strip() for c in row.cells]
            if any(row_data):
                all_rows.append(row_data)
    if not all_rows:
        return pd.DataFrame()
    max_cols = max(len(r) for r in all_rows)
    padded = [r + [''] * (max_cols - len(r)) for r in all_rows]
    return pd.DataFrame(padded)


def parse_docx(file_bytes_or_path, reg_col: Optional[str] = None, name_col: Optional[str] = None) -> List[Student]:
    """Parse students from tables in an existing Word document with header detection & auto-correction."""
    df = docx_to_dataframe(file_bytes_or_path)
    if df.empty:
        return []
    return parse_dataframe(df, reg_col, name_col)


def extract_file_to_dataframe(file_bytes_or_path, filename: str = "") -> Tuple[pd.DataFrame, Optional[str], Optional[str]]:
    """
    Extract any supported file format (.xlsx, .xls, .htm, .html, .csv, .docx) into a normalized DataFrame
    and return (normalized_df, detected_reg_col, detected_name_col).
    """
    if isinstance(file_bytes_or_path, str):
        filename = filename or os.path.basename(file_bytes_or_path)
        with open(file_bytes_or_path, 'rb') as f:
            content_bytes = f.read()
    elif hasattr(file_bytes_or_path, 'read'):
        filename = filename or getattr(file_bytes_or_path, 'name', '')
        content_bytes = file_bytes_or_path.read()
        if hasattr(file_bytes_or_path, 'seek'):
            file_bytes_or_path.seek(0)
    else:
        content_bytes = file_bytes_or_path

    fname_lower = filename.lower()
    if fname_lower.endswith('.docx'):
        raw_df = docx_to_dataframe(io.BytesIO(content_bytes))
    elif fname_lower.endswith(('.csv', '.tsv', '.txt')):
        try:
            raw_df = pd.read_csv(io.BytesIO(content_bytes))
        except Exception:
            raw_df = pd.read_csv(io.BytesIO(content_bytes), sep=None, engine='python')
    else:
        raw_df = smart_parse_spreadsheet(content_bytes, filename)

    norm_df, reg_col, name_col = normalize_dataframe(raw_df)

    cols = list(norm_df.columns)
    if reg_col in cols and name_col in cols:
        if is_pure_int_series(norm_df[reg_col]):
            sample_names = norm_df[name_col].dropna().astype(str).str.strip()[:10]
            if any(is_likely_reg_no(x) for x in sample_names):
                real_reg_idx = cols.index(name_col)
                reg_col = name_col
                if real_reg_idx + 1 < len(cols):
                    name_col = cols[real_reg_idx + 1]

    return norm_df, reg_col, name_col


def parse_raw_text(text: str) -> List[Student]:
    """Parse pasted student text."""
    lines = [line.strip() for line in text.strip().splitlines() if line.strip()]
    students = []
    idx = 1

    for line in lines:
        if '\t' in line:
            parts = [p.strip() for p in line.split('\t') if p.strip()]
        elif ',' in line:
            parts = [p.strip() for p in line.split(',') if p.strip()]
        else:
            parts = [p.strip() for p in re.split(r'\s{2,}', line) if p.strip()]

        if not parts:
            continue

        if len(parts) >= 3 and parts[0].isdigit():
            reg_val = parts[1]
            name_val = " ".join(parts[2:])
        elif len(parts) >= 2:
            reg_val = parts[0]
            name_val = " ".join(parts[1:])
        else:
            continue

        if 'reg' in reg_val.lower() and ('no' in reg_val.lower() or 'number' in reg_val.lower()):
            continue

        students.append(Student(reg_no=reg_val, name=name_val, original_sn=idx))
        idx += 1

    return students


def generate_blank_excel_template() -> bytes:
    """Generate a clean sample Excel template for departments to fill out."""
    output = io.BytesIO()
    sample_data = {
        'S/N': [1, 2, 3],
        'REG NO': ['COE23HND0120', 'COE23HND0775', 'COE24HND0002'],
        'NAMES': ['KINSMAN ADAVIZE OMANANYI', 'EMMANUEL ORDUE LUMUNGA', 'ENOCH TUKURAH ILIYA']
    }
    df = pd.DataFrame(sample_data)
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Students')
    return output.getvalue()
