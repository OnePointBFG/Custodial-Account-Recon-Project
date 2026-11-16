"""
OnePoint BFG Wealth Partners - Account Cancellation Reconciliation Engine
Compares an Orion account export against a custodian account file.

All matching is exact on a standardized account number. Nothing is matched
automatically on similarity; near-matches are reported for human review only.
"""

from __future__ import annotations

import hashlib
import io
import json
import re
import unicodedata
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path

import warnings

import pandas as pd

# openpyxl prints a harmless warning for workbooks exported without a default style; hide it.
warnings.filterwarnings("ignore", message="Workbook contains no default style", category=UserWarning)

TOOL_VERSION = "1.0.0"
CONFIG_PATH = Path(__file__).with_name("custodians.json")

# ---------------------------------------------------------------------------
# Result categories (each row lands in exactly one)
# ---------------------------------------------------------------------------
# Orion-side results (every in-scope Orion row gets exactly one)
CANCEL = "Cancel in Orion"
MATCHED = "Matched"
ALREADY = "Already Cancelling"
NO_ACTION = "Already Inactive"
EXCLUDED_ZERO = "Excluded - Zero Value"
INACTIVE_OPEN = "Inactive in Orion - Open at Custodian"
CANCEL_PENDING_OPEN = "Cancel Pending - Open at Custodian"
CUST_UNCLEAR = "Custodian Status Unclear"
# Custodian-side results (every custodian row gets exactly one)
MISSING = "Not in Orion"
CUST_CLOSED = "Ignored - Closed"
CUST_NOT_ACTIVE = "Ignored - Not Active at Custodian"
CUST_IGNORED_ZERO = "Ignored - Zero Value"
CUST_BLANK = "Ignored - Blank Account Number"
EXCLUDED_LIST = "Excluded - On Exclusion List"   # used on both sides

ORION_CATEGORIES = [CANCEL, MATCHED, ALREADY, NO_ACTION, EXCLUDED_ZERO, INACTIVE_OPEN,
                    CANCEL_PENDING_OPEN, CUST_UNCLEAR, EXCLUDED_LIST]
CUST_CATEGORIES = [MATCHED, MISSING, CUST_CLOSED, CUST_NOT_ACTIVE, CUST_IGNORED_ZERO, CUST_BLANK, EXCLUDED_LIST]


class ReconError(Exception):
    """A problem that must stop the reconciliation."""


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
def load_config(path: Path = CONFIG_PATH) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------
def _norm_header(text) -> str:
    """Header comparison: case-insensitive, collapse whitespace."""
    if text is None:
        return ""
    s = unicodedata.normalize("NFKC", str(text))
    return re.sub(r"\s+", " ", s).strip().lower()


def _clean_text(value) -> str:
    """Generic cell to text; blanks become ''."""
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    if value is pd.NaT:
        return ""
    s = unicodedata.normalize("NFKC", str(value)).strip()
    return "" if s.lower() in ("nan", "nat", "none") else s


def account_raw_text(value) -> tuple[str, str]:
    """
    Convert an account-number cell to text exactly as stored.
    Returns (text, warning). A warning is raised when Excel stored the
    value as a number, because leading zeros may already have been lost.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "", ""
    if isinstance(value, bool):
        return str(value), "Account number stored as TRUE/FALSE"
    if isinstance(value, int):
        return str(value), "Stored as a number in Excel; leading zeros may be lost"
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value)), "Stored as a number in Excel; leading zeros may be lost"
        return repr(value), "Account number contains a decimal"
    text = _clean_text(value)
    if re.fullmatch(r"\d+(\.\d+)?[eE][+-]?\d+", text):
        return text, "Scientific notation; the account number is damaged"
    return text, ""


def standardize_account(text: str) -> str:
    """Uppercase and remove everything that is not a letter or digit."""
    s = unicodedata.normalize("NFKC", text).upper()
    return re.sub(r"[^A-Z0-9]", "", s)


def parse_money(value) -> tuple[float | None, bool]:
    """Returns (amount, parse_ok). Blank is (None, True)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None, True
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value), True
    s = _clean_text(value)
    if s == "" or s == "-":
        return None, True
    neg = s.startswith("(") and s.endswith(")")
    s2 = re.sub(r"[,$\s()]", "", s)
    try:
        amt = float(s2)
    except ValueError:
        return None, False
    return (-amt if neg else amt), True


def file_sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


# ---------------------------------------------------------------------------
# File loading
# ---------------------------------------------------------------------------
@dataclass
class LoadedFile:
    name: str
    sha256: str
    sheet: str
    header_row: int            # 1-based row number in the file
    df: pd.DataFrame           # only the mapped columns, renamed to field keys
    column_map: dict           # field key -> original header text
    total_rows: int
    warnings: list = field(default_factory=list)
    notes: list = field(default_factory=list)       # informational (e.g. footer ignored)
    preamble: dict = field(default_factory=dict)    # "Label: value" lines above the header


def _read_xlsx_cells(data: bytes, sheet_hint: str | None = None) -> tuple[pd.DataFrame, str]:
    """
    Fallback .xlsx reader: reads cell VALUES straight from the workbook XML and ignores formatting.
    Used only when the normal reader fails (e.g. a custodian export with malformed styles).
    Text stays text; numbers become int (if whole) or float, exactly as stored.
    """
    import posixpath
    import xml.etree.ElementTree as ET

    def strip(tag):
        return tag.rsplit("}", 1)[-1]

    def find_all(el, name):
        return [c for c in el.iter() if strip(c.tag) == name]

    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        raise ReconError("the file is not a valid .xlsx workbook")
    names = set(z.namelist())
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    sheets = [(s_.get("name"), next((v for k, v in s_.attrib.items() if strip(k) == "id"), None))
              for s_ in find_all(wb, "sheet")]
    if not sheets:
        raise ReconError("the workbook has no sheets")
    rels = {}
    if "xl/_rels/workbook.xml.rels" in names:
        for r in find_all(ET.fromstring(z.read("xl/_rels/workbook.xml.rels")), "Relationship"):
            rels[r.get("Id")] = r.get("Target")
    chosen = next((s_ for s_ in sheets if s_[0] == sheet_hint), sheets[0])
    target = rels.get(chosen[1], "worksheets/sheet1.xml")
    path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
    shared = []
    if "xl/sharedStrings.xml" in names:
        for si in find_all(ET.fromstring(z.read("xl/sharedStrings.xml")), "si"):
            shared.append("".join(t.text or "" for t in si.iter() if strip(t.tag) == "t"))

    # Date formats (so date cells read as dates, as the normal reader does). Formatting is otherwise ignored.
    date_styles = set()
    if "xl/styles.xml" in names:
        try:
            st = ET.fromstring(z.read("xl/styles.xml"))
            custom = {}
            for nf in find_all(st, "numFmt"):
                code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", nf.get("formatCode", "")).lower()
                custom[int(nf.get("numFmtId"))] = any(ch in code for ch in "dy") or ("m" in code and "h" not in code
                                                                                     and "s" not in code)
            xfs = next((x for x in st.iter() if strip(x.tag) == "cellXfs"), None)
            if xfs is not None:
                for i, xf in enumerate(x for x in xfs if strip(x.tag) == "xf"):
                    fid = int(xf.get("numFmtId", "0"))
                    if 14 <= fid <= 22 or 45 <= fid <= 47 or custom.get(fid, False):
                        date_styles.add(i)
        except Exception:  # noqa: BLE001 - styles unreadable: dates stay as serial numbers
            date_styles = set()

    def col_index(ref):
        letters = "".join(ch for ch in ref if ch.isalpha())
        n = 0
        for ch in letters.upper():
            n = n * 26 + (ord(ch) - 64)
        return n - 1

    rows = {}
    for row in find_all(ET.fromstring(z.read(path)), "row"):
        r_idx = int(row.get("r")) - 1 if row.get("r") else (max(rows) + 1 if rows else 0)
        cells, next_col = {}, 0
        for c in row:
            if strip(c.tag) != "c":
                continue
            ref = c.get("r")
            ci = col_index(ref) if ref else next_col
            next_col = ci + 1
            t = c.get("t", "n")
            v = next((x.text for x in c if strip(x.tag) == "v"), None)
            if t == "s" and v is not None:
                val = shared[int(v)]
            elif t == "inlineStr":
                val = "".join(x.text or "" for x in c.iter() if strip(x.tag) == "t")
            elif t in ("str", "e"):
                val = v
            elif t == "b":
                val = (v == "1") if v is not None else None
            elif v is None or v == "":
                val = None
            else:
                try:
                    f = float(v)
                    val = int(f) if f.is_integer() and "e" not in v.lower() and "." not in v else f
                    if c.get("s") is not None and int(c.get("s")) in date_styles:
                        val = datetime(1899, 12, 30) + timedelta(days=f)
                except ValueError:
                    val = v
            cells[ci] = val
        rows[r_idx] = cells
    if not rows:
        return pd.DataFrame(), chosen[0]
    width = max((max(c) + 1 for c in rows.values() if c), default=0)
    table = [[rows.get(i, {}).get(j, "") for j in range(width)] for i in range(max(rows) + 1)]
    table = [["" if v is None else v for v in r] for r in table]
    return pd.DataFrame(table, dtype=object), chosen[0]


def _read_raw(name: str, data: bytes, sheet_hint: str | None) -> tuple[pd.DataFrame, str, list]:
    warnings = []
    lower = name.lower()
    if lower.endswith(".csv"):
        import csv
        text = None
        for enc in ("utf-8-sig", "cp1252"):
            try:
                text = data.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise ReconError(f"{name}: the CSV file could not be read (unknown text encoding).")
        rows = list(csv.reader(io.StringIO(text, newline="")))
        width = max((len(r) for r in rows), default=0)
        rows = [r + [""] * (width - len(r)) for r in rows]  # title lines may be shorter
        return pd.DataFrame(rows, dtype=object), "(CSV)", warnings
    if lower.endswith((".xlsx", ".xlsm")):
        try:
            xls = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
        except Exception as e:  # noqa: BLE001 - e.g. malformed formatting (styles) in the export
            raw, sheet = _read_xlsx_cells(data, sheet_hint)
            warnings.append(f"Read with the fallback reader because the file's formatting could not be read "
                            f"({type(e).__name__}); cell values are unaffected.")
            return raw, sheet, warnings
        sheets = xls.sheet_names
        if sheet_hint and sheet_hint in sheets:
            sheet = sheet_hint
        else:
            sheet = sheets[0]
            if sheet_hint:
                warnings.append(f"Expected sheet '{sheet_hint}' not found; used first sheet '{sheet}'.")
        if len(sheets) > 1:
            warnings.append(f"Workbook has {len(sheets)} sheets; only '{sheet}' was read.")
        raw = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=object,
                            keep_default_na=False, na_values=[])  # never turn "N/A", "NA", "NULL" into blanks
        return raw, sheet, warnings
    raise ReconError(f"{name}: unsupported file type. Use .xlsx or .csv.")


def load_file(name: str, data: bytes, profile: dict, label: str) -> LoadedFile:
    """
    profile keys:
      columns          {field: [header aliases]}   only these columns are ever read
      required         [fields]
      sheet            optional sheet name
      footer_markers   optional; reading stops at the first row whose first non-empty
                       cell starts with one of these (e.g. "Disclosures")
    """
    raw, sheet, warnings = _read_raw(name, data, profile.get("sheet"))
    if raw.empty:
        raise ReconError(f"{label} file '{name}' is empty.")
    notes = []

    columns_cfg: dict = profile["columns"]
    required: list = profile["required"]
    key_aliases = {_norm_header(a) for a in columns_cfg["account"]}

    # Header row: first row (within the first 50) containing the account header.
    scan = min(50, len(raw))
    header_idx = None
    for i in range(scan):
        row_vals = {_norm_header(v) for v in raw.iloc[i].tolist()}
        if row_vals & key_aliases:
            header_idx = i
            break
    if header_idx is None:
        raise ReconError(
            f"{label} file '{name}': could not find an account number column "
            f"(looked for: {', '.join(columns_cfg['account'])}) in the first {scan} rows."
        )

    # Title lines above the header in "Label: value" form (e.g. "As of: Oct 22, 2026").
    preamble = {}
    for i in range(header_idx):
        cells = [_clean_text(v) for v in raw.iloc[i].tolist() if _clean_text(v) != ""]
        if len(cells) == 1 and ":" in cells[0]:
            k, v = cells[0].split(":", 1)
            preamble[k.strip()] = v.strip()

    headers = [_clean_text(v) for v in raw.iloc[header_idx].tolist()]
    norm_headers = [_norm_header(h) for h in headers]

    column_map, col_index = {}, {}
    for fld, aliases in columns_cfg.items():
        alias_norm = [_norm_header(a) for a in aliases]
        hits = [i for i, h in enumerate(norm_headers) if h in alias_norm]
        if len(hits) > 1:
            raise ReconError(
                f"{label} file '{name}': column for '{fld}' appears more than once "
                f"({', '.join(headers[i] for i in hits)}). Remove the duplicate column."
            )
        if hits:
            column_map[fld] = headers[hits[0]]
            col_index[fld] = hits[0]

    missing = [f for f in required if f not in column_map]
    if missing:
        names = [" / ".join(columns_cfg[m]) for m in missing]
        raise ReconError(f"{label} file '{name}' is missing required column(s): {'; '.join(names)}.")

    body = raw.iloc[header_idx + 1:]
    body_rows = [header_idx + 2 + k for k in range(len(body))]  # 1-based file row numbers

    # Stop at the footer (disclosures/disclaimers) so report text is never read as an account.
    markers = [m.lower() for m in profile.get("footer_markers", [])]
    if markers:
        cut = None
        for k in range(len(body)):
            cells = [_clean_text(v) for v in body.iloc[k].tolist() if _clean_text(v) != ""]
            if cells and any(cells[0].lower().startswith(m) for m in markers):
                cut = k
                break
        if cut is not None:
            # Safety: nothing after the footer may look like account data (footer lines are single text cells).
            for k in range(cut, len(body)):
                cells = [_clean_text(v) for v in body.iloc[k].tolist() if _clean_text(v) != ""]
                if len(cells) > 2:
                    raise ReconError(
                        f"{label} file '{name}': row {body_rows[k]} looks like account data but appears after the "
                        f"report footer that starts on row {body_rows[cut]}. The file layout is unexpected; "
                        "check the export before running again.")
            first = [_clean_text(v) for v in body.iloc[cut].tolist() if _clean_text(v) != ""][0]
            notes.append(f"Footer starting at row {body_rows[cut]} ('{first[:40]}') ignored "
                         f"({len(body) - cut} rows).")
            body = body.iloc[:cut]
            body_rows = body_rows[:cut]
        elif profile.get("footer_required", True):
            raise ReconError(f"{label} file '{name}': the expected footer ({profile['footer_markers'][0]}) "
                             "was not found. The export may be incomplete or its layout has changed. "
                             "Re-export the file and run again.")
        else:
            notes.append("No report footer found (none expected for every export).")

    # Drop completely empty rows, keeping true file row numbers.
    keep = [k for k in range(len(body)) if any(_clean_text(v) != "" for v in body.iloc[k].tolist())]
    blank_rows = len(body) - len(keep)
    if blank_rows:
        notes.append(f"{blank_rows} blank row(s) skipped.")
    if not keep:
        raise ReconError(f"{label} file '{name}' has headers but no data rows.")
    body = body.iloc[keep]
    file_rows = [body_rows[k] for k in keep]

    df = pd.DataFrame({fld: body.iloc[:, idx].tolist() for fld, idx in col_index.items()})
    df.insert(0, "source_row", file_rows)

    return LoadedFile(
        name=name, sha256=file_sha256(data), sheet=sheet, header_row=header_idx + 1,
        df=df, column_map=column_map, total_rows=len(df), warnings=warnings,
        notes=notes, preamble=preamble,
    )


# ---------------------------------------------------------------------------
# Value interpretation
# ---------------------------------------------------------------------------
def interpret_orion_active(value, cfg: dict) -> str:
    """Returns 'Active', 'Inactive', or 'Unrecognized'."""
    if isinstance(value, bool):
        return "Active" if value else "Inactive"
    s = _clean_text(value).lower()
    if s in [v.lower() for v in cfg["active_values"]]:
        return "Active"
    if s in [v.lower() for v in cfg["inactive_values"]]:
        return "Inactive"
    return "Unrecognized"


def interpret_custodian_state(status, close_date, prof: dict, account_raw: str = "") -> tuple[str, str]:
    """
    Returns (state, reason). state is 'Open', 'Closed', or 'Unrecognized'.
    Closed evidence: a closed status value, a close date, or a closed marker in the
    account number (e.g. Fidelity appends '&' to closed accounts). Evidence of closure
    combined with an explicit open status is a conflict and goes to Review.
    """
    s = _clean_text(status).lower()
    open_vals = [v.lower() for v in prof.get("open_status_values", [])]
    closed_vals = [v.lower() for v in prof.get("closed_status_values", [])]
    has_status_col = "status" in prof["columns"]

    evidence = []
    if _clean_text(close_date) != "":
        evidence.append("close date present")
    for m in prof.get("closed_if_account_contains", []):
        if m and m in (account_raw or ""):
            evidence.append(f"account number contains '{m}'")

    if s in closed_vals:
        return "Closed", f"Status '{_clean_text(status)}'"
    if s in open_vals:
        if evidence:
            return "Unrecognized", f"Status shows open but {' and '.join(evidence)}"
        return "Open", ""
    if evidence:
        if s == "":
            return "Closed", "; ".join(e[0].upper() + e[1:] for e in evidence)
        return "Unrecognized", f"Unrecognized status '{_clean_text(status)}' with {' and '.join(evidence)}"
    if s == "":
        if has_status_col:
            return "Unrecognized", "Status is blank"
        return "Open", ""  # custodian file has no status column; being listed means open
    if prof.get("only_open_status_is_active"):
        # e.g. Pontera: only 'Account ready' means active; any other status means not active (to be verified)
        return "NotActive", f"Status '{_clean_text(status)}', not '{prof['open_status_values'][0]}'"
    return "Unrecognized", f"Unrecognized status '{_clean_text(status)}'"


# ---------------------------------------------------------------------------
# File type detection (Orion or which custodian), by column signature
# ---------------------------------------------------------------------------
def detect_file(name: str, data: bytes, config: dict) -> tuple[list[str], str]:
    """
    Returns (matches, detail). matches lists every profile ('Orion' or a custodian name)
    whose signature columns are ALL present in one header row within the first 50 rows.
    Exactly one match is required for the file to be used.
    """
    try:
        raw, _, _ = _read_raw(name, data, None)
    except ReconError as e:
        return [], str(e)
    except Exception as e:  # noqa: BLE001
        return [], f"could not be read ({e})"
    profiles = {"Orion": config["orion"]}
    profiles.update(config["custodians"])
    rows = [{_norm_header(v) for v in raw.iloc[i].tolist()} for i in range(min(50, len(raw)))]
    matches = []
    for label, prof in profiles.items():
        sig = {_norm_header(h) for h in prof.get("signature", [])}
        if sig and any(sig <= r for r in rows):
            matches.append(label)
    return matches, ""


# ---------------------------------------------------------------------------
# Near-match detection (one edit or one adjacent swap)
# ---------------------------------------------------------------------------
def _within_one_edit(a: str, b: str) -> str | None:
    if a == b:
        return None
    la, lb = len(a), len(b)
    if abs(la - lb) > 1:
        return None
    if la == lb:
        diffs = [i for i in range(la) if a[i] != b[i]]
        if len(diffs) == 1:
            return "One character differs"
        if len(diffs) == 2 and diffs[1] == diffs[0] + 1 \
                and a[diffs[0]] == b[diffs[1]] and a[diffs[1]] == b[diffs[0]]:
            return "Two characters swapped"
        return None
    short, long_ = (a, b) if la < lb else (b, a)
    i = j = 0
    skipped = False
    while i < len(short) and j < len(long_):
        if short[i] != long_[j]:
            if skipped:
                return None
            skipped = True
            j += 1
            continue
        i += 1
        j += 1
    return "One character extra or missing"


def find_near_matches(orion_keys: list[str], cust_keys: list[str]) -> list[tuple[str, str, str]]:
    by_len: dict[int, list[str]] = {}
    for k in set(cust_keys):
        by_len.setdefault(len(k), []).append(k)
    out = []
    for ok in sorted(set(orion_keys)):
        for L in (len(ok) - 1, len(ok), len(ok) + 1):
            for ck in by_len.get(L, []):
                why = _within_one_edit(ok, ck)
                if why:
                    out.append((ok, ck, why))
    return out


# ---------------------------------------------------------------------------
# Reconciliation
# ---------------------------------------------------------------------------
@dataclass
class ReconResult:
    custodian: str
    orion_selected_values: list
    orion: pd.DataFrame          # Orion rows in scope, with Result
    custodian_df: pd.DataFrame   # custodian rows, with Result
    near_matches: pd.DataFrame
    excluded_orion: pd.DataFrame # Orion rows for other custodians (counts only)
    controls: list               # (description, expected, actual, passed)
    passed: bool
    orion_file: LoadedFile
    cust_file: LoadedFile
    run_time: datetime
    config_snapshot: dict
    completeness: list = field(default_factory=list)  # (description, report total, sum of rows, passed)
    exclusion_info: dict = field(default_factory=dict)  # entries, applied (orion/custodian rows), unused
    status_summary: list = field(default_factory=list)  # (custodian status as written, interpreted state, rows)
    match_info: dict = field(default_factory=dict)      # found, considered, rate, low (bool)


def parse_account(text: str, sleeve_sep: str | None = None, markers=()) -> tuple[str, str, str]:
    """
    Split an account number into (key, sleeve name, unusual-character note).
      key     the parent account number, standardized (uppercase, letters and digits only)
      sleeve  text after the sleeve separator (UMA sleeves in Orion, e.g. "X12345678_Growth")
      note    description of anything unexpected; '' when clean
    Letters, digits, spaces and dashes are normal. Configured markers (e.g. Fidelity's '&') are
    expected. Anything else is reported so the account is never silently matched.
    """
    base, sleeve, issues = text, "", []
    if sleeve_sep and sleeve_sep in text:
        base, sleeve = text.split(sleeve_sep, 1)
        sleeve = sleeve.strip()
        if not sleeve:
            issues.append(f"'{sleeve_sep}' with no sleeve name")
        if not standardize_account(base):
            issues.append("nothing before the sleeve separator")
    odd = sorted({ch for ch in unicodedata.normalize("NFKC", base)
                  if not ((ch.isascii() and ch.isalnum()) or ch in " -") and ch not in markers})
    if odd:
        issues.append("unusual character(s) " + " ".join(f"'{ch}'" for ch in odd))
    return standardize_account(base), sleeve, "; ".join(issues)


def account_shape(text: str, sleeve_sep: str | None = None) -> str:
    """Pattern of an account number: letters -> A, digits -> 9, other characters kept."""
    base, has_sleeve = text, False
    if sleeve_sep and sleeve_sep in text:
        base, rest = text.split(sleeve_sep, 1)
        has_sleeve = bool(rest.strip())
        if not has_sleeve:
            base = base + sleeve_sep  # keep the bare separator visible in the pattern
    shape = "".join("A" if (c.isascii() and c.isalpha()) else "9" if c.isdigit() else c for c in base.upper())
    return (shape or "(blank)") + (f"{sleeve_sep}<sleeve>" if has_sleeve else "")


def _prepare_accounts(df: pd.DataFrame, pattern: str | None, sleeve_sep: str | None = None,
                      markers=()) -> pd.DataFrame:
    df = df.copy()
    raws, warns, keys, fmt, sleeves, unusual, shapes = [], [], [], [], [], [], []
    rx = re.compile(pattern) if pattern else None
    for v in df["account"].tolist():
        t, w = account_raw_text(v)
        k, sl, odd = parse_account(t, sleeve_sep, markers)
        raws.append(t)
        warns.append(w)
        keys.append(k)
        sleeves.append(sl)
        unusual.append(odd)
        shapes.append(account_shape(t, sleeve_sep))
        fmt.append(bool(k) and rx is not None and not rx.fullmatch(k))
    df["Account Number (as in file)"] = raws
    df["Account Key"] = keys
    df["Sleeve Name"] = sleeves
    df["_acct_warning"] = warns
    df["_format_exception"] = fmt
    df["_unusual"] = unusual
    df["_shape"] = shapes
    return df


# ---------------------------------------------------------------------------
# Exclusion list (accounts to ignore), kept locally: Custodian | Account Number | Reason
# ---------------------------------------------------------------------------
EXCLUSION_HEADERS = ["Custodian", "Account Number", "Reason"]


def _exclusion_account(value, pattern_list) -> tuple[str, str, str]:
    """
    Validate one exclusion value. Returns (key, text as written, problem). problem is '' when valid.
    An Excel number is accepted only if, as written, it already has the full expected length
    (so no leading zero can have been lost).
    """
    if isinstance(value, bool):
        return "", str(value), "not an account number"
    numeric = isinstance(value, (int, float)) and not (isinstance(value, float) and pd.isna(value))
    if numeric:
        if isinstance(value, float) and not value.is_integer():
            return "", str(value), "a decimal number is not an account number"
        text = str(int(value))
        # Excel drops leading zeros from numbers. Accept a number only if it exactly fits a custodian whose
        # account numbers have a FIXED length - then no digit can be missing.
        fixed = [p for _, p in pattern_list if p and re.search(r"\{\d+\}\$$", p)]
        if not any(re.fullmatch(p, text) for p in fixed):
            return "", text, ("stored as a number in Excel, so leading zeros may be lost; "
                              "format the cell as Text and re-enter it")
    else:
        text = account_raw_text(value)[0]
    if text == "":
        return "", "", "blank"
    key, sleeve, odd = parse_account(text, "_")
    if sleeve or "_" in text:
        return "", text, "enter the account number only (no '_' sleeve)"
    if odd:
        return "", text, f"account number has {odd}"
    fits = [c for c, pat in pattern_list if not pat or re.fullmatch(pat, key)]
    if not key or not fits:
        return "", text, "not a valid account number for any custodian"
    return key, text, ""


def load_exclusions(name: str, data: bytes, config: dict) -> dict:
    """
    Returns {custodian: {account_key: (account number as written, reason, file row)}}.

    Two layouts are accepted (first sheet):
      1. Simple list   - account numbers in column A; an optional heading in row 1. Each number applies
                         to every custodian whose account format it fits.
      2. Custodian list - headers 'Custodian', 'Account Number' (optional 'Reason'); each entry applies
                         to the named custodian only.
    Any invalid entry stops the run: an exclusion must never be applied to the wrong account.
    """
    raw, sheet, _ = _read_raw(name, data, None)
    if raw.empty:
        return {}
    customs = config["custodians"]
    patterns = [(c, p.get("account_pattern")) for c, p in customs.items()]
    heads = [_norm_header(v) for v in raw.iloc[0].tolist()]
    need = [_norm_header(h) for h in EXCLUSION_HEADERS]
    out, problems = {}, []

    if need[0] in heads and need[1] in heads:          # ---- custodian list
        ic, ia = heads.index(need[0]), heads.index(need[1])
        ir = heads.index(need[2]) if need[2] in heads else None
        known = {c.lower(): c for c in customs}
        for i in range(1, len(raw)):
            row = raw.iloc[i].tolist()
            cust_txt = _clean_text(row[ic])
            reason = _clean_text(row[ir]) if ir is not None else ""
            if cust_txt == "" and _clean_text(row[ia]) == "" and reason == "":
                continue
            fr = i + 1
            if cust_txt == "":
                # Custodian left blank: applies to every custodian whose account format the number fits.
                key, text, why = _exclusion_account(row[ia], patterns)
                if why:
                    problems.append(f"row {fr}: '{text}' - {why}" + (
                        " (is the account number in column A? It belongs in the Account Number column)"
                        if why == "blank" else ""))
                    continue
                for c_, pat in patterns:
                    if not pat or re.fullmatch(pat, key):
                        out.setdefault(c_, {})[key] = (text, reason, fr)
                continue
            cust = known.get(cust_txt.lower())
            if cust is None:
                hint = (" - this looks like an account number; put it in the Account Number column"
                        if not _exclusion_account(row[ic], patterns)[2] else "")
                problems.append(f"row {fr}: custodian '{cust_txt}' is not one of {', '.join(customs)}{hint}")
                continue
            key, text, why = _exclusion_account(row[ia], [(cust, customs[cust].get("account_pattern"))])
            if why:
                problems.append(f"row {fr}: '{text}' - " + (f"not a valid {cust} account number"
                                                             if why.startswith("not a valid") else why))
            else:
                out.setdefault(cust, {})[key] = (text, reason, fr)
    else:                                              # ---- simple list in column A
        for i in range(len(raw)):
            row = raw.iloc[i].tolist()
            v = row[0] if row else ""
            key, text, why = _exclusion_account(v, patterns)
            if why == "blank":
                continue
            if why:
                if i == 0 and not isinstance(v, (int, float)):
                    continue  # heading
                problems.append(f"row {i + 1}: '{text}' - {why}")
                continue
            for c, pat in patterns:
                if not pat or re.fullmatch(pat, key):
                    out.setdefault(c, {})[key] = (text, "", i + 1)
    if problems:
        raise ReconError(f"Exclusion list '{name}' (sheet '{sheet}') has problems - fix them, then run again:\n    "
                         + "\n    ".join(problems))
    return out


LOW_MATCH_RATE = 0.5      # warn when fewer than half of a custodian's accounts are found in Orion
LOW_MATCH_MIN_ROWS = 20   # ...and the custodian file has at least this many accounts


def _status_summary(c: pd.DataFrame) -> list:
    if "status" not in c.columns:
        return []
    st = c["status"].map(_clean_text)
    g = pd.DataFrame({"s": st, "state": c["Custodian State"]}).value_counts().reset_index()
    label = {"Open": "active", "Closed": "closed", "NotActive": "not active", "Unrecognized": "unclear"}
    return [(r["s"] or "(blank)", label.get(r["state"], r["state"]), int(r["count"])) for _, r in g.iterrows()]


def _match_info(c: pd.DataFrame) -> dict:
    considered = c[(c["Account Key"] != "") & (c["Result"] != EXCLUDED_LIST)]
    found = int((considered["Result"] == MATCHED).sum())
    n = len(considered)
    rate = found / n if n else 1.0
    return {"found": found, "considered": n, "rate": rate,
            "low": n >= LOW_MATCH_MIN_ROWS and rate < LOW_MATCH_RATE}


def reconcile(orion_file: LoadedFile, cust_file: LoadedFile, custodian: str,
              orion_custodian_values: list, config: dict, exclusions: dict | None = None) -> ReconResult:
    run_time = datetime.now()
    ocfg = config["orion"]
    prof = config["custodians"][custodian]
    pattern = prof.get("account_pattern") or None

    # ---------------- Orion: scope to selected custodian values ----------------
    o_all = orion_file.df.copy()
    o_all["_cust_text"] = o_all["custodian"].map(_clean_text)
    selected = set(orion_custodian_values)
    in_scope = o_all["_cust_text"].isin(selected)
    o = o_all[in_scope].copy()
    excluded = o_all[~in_scope].copy()
    if len(o) == 0:
        raise ReconError("No Orion rows match the selected custodian value(s).")

    sleeve_sep = ocfg.get("sleeve_separator") or None
    markers = tuple(prof.get("closed_if_account_contains", []))
    o = _prepare_accounts(o, pattern, sleeve_sep)
    uma_parents = set(o.loc[(o["Sleeve Name"] != "") & (o["Account Key"] != ""), "Account Key"])
    o["_uma"] = o["Account Key"].isin(uma_parents) & (o["Account Key"] != "")
    o["UMA Status"] = ["UMA Sleeve" if sl else "UMA Parent" if u else "Standard"
                         for sl, u in zip(o["Sleeve Name"], o["_uma"])]
    o["Orion Active (interpreted)"] = o["active"].map(lambda v: interpret_orion_active(v, ocfg))
    o["_cancel_date_set"] = o["cancel_date"].map(lambda v: _clean_text(v) != "") \
        if "cancel_date" in o else False

    # Orion value: decides between Cancel (value > 0) and Excluded (value = 0).
    def _amount(v):
        if _clean_text(v) == "":
            return "blank", None
        amt, ok = parse_money(v)
        if not ok or amt is None:
            return "unreadable", None
        return "ok", amt

    value_class, value_note = [], []
    for _, r in o.iterrows():
        cv_state, cv = _amount(r["current_value"]) if "current_value" in o.columns else ("blank", None)
        cash_state, cash = _amount(r["cash_balance"]) if "cash_balance" in o.columns else ("blank", None)
        if cv_state != "ok":
            value_class.append("unclear"); value_note.append(f"Orion Current Value {cv_state}")
        elif cv < -0.005:
            value_class.append("unclear"); value_note.append(f"Orion Current Value negative ({cv:,.2f})")
        elif cv >= 0.005:
            value_class.append("value"); value_note.append(f"Orion value {cv:,.2f}")
        elif cash_state == "unreadable" or (cash_state == "ok" and abs(cash) >= 0.005):
            note = "unreadable" if cash_state == "unreadable" else f"{cash:,.2f}"
            value_class.append("unclear"); value_note.append(f"Orion Current Value 0 but Cash Balance {note}")
        else:
            value_class.append("zero"); value_note.append("Orion value 0")
    o["_value_class"] = value_class
    o["_value_note"] = value_note

    # ---------------- Custodian ----------------
    c = _prepare_accounts(cust_file.df, pattern, None, markers)  # custodians have no sleeves
    states = [interpret_custodian_state(r.get("status"), r.get("close_date"), prof,
                                        r["Account Number (as in file)"])
              for _, r in c.iterrows()]
    c["Custodian State"] = [s for s, _ in states]
    c["_state_reason"] = [why for _, why in states]

    # ---------------- Duplicates ----------------
    o_nonblank = o[o["Account Key"] != ""]
    c_nonblank = c[c["Account Key"] != ""]
    # Orion: several UMA sleeves under one parent are expected; the same parent AND sleeve twice is a duplicate.
    o_pair = o_nonblank["Account Key"] + "\x00" + o_nonblank["Sleeve Name"].str.upper()
    o_dups = set(o_nonblank["Account Key"][o_pair.duplicated(keep=False)])
    c_dups = set(c_nonblank["Account Key"][c_nonblank["Account Key"].duplicated(keep=False)])

    # Custodian lookup: key -> list of states (duplicates kept visible)
    c_states: dict[str, list[str]] = {}
    c_closed_why: dict[str, str] = {}
    c_unclear_why: dict[str, str] = {}
    for k, s, why in zip(c["Account Key"], c["Custodian State"], c["_state_reason"]):
        if k:
            c_states.setdefault(k, []).append(s)
            if s in ("Closed", "NotActive"):
                c_closed_why[k] = why
            if s == "Unrecognized":
                c_unclear_why[k] = why
    o_keys_scope = set(o_nonblank["Account Key"])
    o_keys_other = {parse_account(account_raw_text(v)[0], sleeve_sep)[0] for v in excluded["account"].tolist()}
    c_unusual = {k: why for k, why in zip(c["Account Key"], c["_unusual"]) if k and why}
    o_keys_other.discard("")

    # ---------------- Near matches ----------------
    o_unmatched = [k for k in o_keys_scope if k not in c_states]
    c_unmatched = [k for k in c_states if k not in o_keys_scope]
    nm = find_near_matches(o_unmatched, c_unmatched)
    nm_orion: dict[str, list[str]] = {}
    nm_cust: dict[str, list[str]] = {}
    for a, b, _ in nm:
        nm_orion.setdefault(a, []).append(b)
        nm_cust.setdefault(b, []).append(a)

    # ---------------- Categorize Orion rows ----------------
    # Every row is decided by rule. Anything uncertain is still decided (never left unassigned)
    # and marked Verify First with the reason, on the tab where it was placed.
    excl = (exclusions or {})
    results, reasons, flags, cust_state_col = [], [], [], []
    for _, r in o.iterrows():
        k = r["Account Key"]
        active = r["Orion Active (interpreted)"]
        cancel_set = bool(r["_cancel_date_set"])
        st_list = c_states.get(k, []) if k else []
        if not st_list:
            cstate = "Not in custodian file"
        elif len(set(st_list)) == 1:
            cstate = st_list[0]
        else:
            cstate = "Conflicting listings"
        cust_state_col.append(cstate)

        if k and k in excl:  # cust_state_col already appended above
            results.append(EXCLUDED_LIST)
            reasons.append(f"On exclusion list (row {excl[k][2]})" + (f": {excl[k][1]}" if excl[k][1] else ""))
            flags.append("")
            continue
        why, chk = [], []   # chk = reasons to verify first
        if k == "":
            chk.append("Blank account number in Orion")
        if r["_acct_warning"]:
            chk.append(r["_acct_warning"])
        if active == "Unrecognized":
            chk.append(f"Orion Active value '{_clean_text(r['active'])}' not recognized; treated as active")
            active = "Active"
        if r["_unusual"]:
            chk.append(f"Account number has {r['_unusual']}")
        if r["_format_exception"]:
            chk.append("Account number does not fit the expected format")
        if k in o_dups:
            chk.append("Same account and sleeve listed more than once in Orion" if r["Sleeve Name"]
                       else "Account number appears more than once in Orion")
        if r["Sleeve Name"]:
            why.append(f"UMA sleeve '{r['Sleeve Name']}' of {k}")
        elif r["_uma"]:
            why.append("UMA parent account (has sleeves)")

        if cstate in ("Open", "Unrecognized", "Conflicting listings"):
            # The custodian lists this account: never cancel.
            if c_unusual.get(k):
                chk.append(f"{custodian} account number has {c_unusual[k]}")
            if cstate == "Unrecognized":
                chk.append(f"{custodian} status unclear ({c_unclear_why.get(k, 'unrecognized')})")
                res = CUST_UNCLEAR
            elif cstate == "Conflicting listings":
                chk.append(f"{custodian} lists this account more than once with different statuses")
                res = CUST_UNCLEAR
            elif active == "Inactive":
                res = INACTIVE_OPEN
                why.append(f"Inactive in Orion but open at {custodian}; listed on Not in Orion")
            elif cancel_set:
                res = CANCEL_PENDING_OPEN
                chk.append(f"Orion cancel date is set but the account is open at {custodian}")
            else:
                res = MATCHED
                why.append(f"Open at {custodian}")
        else:
            base = (f"Not in {custodian} file" if cstate == "Not in custodian file"
                    else f"Not active at {custodian} ({c_closed_why.get(k, '')})" if cstate == "NotActive"
                    else f"Closed at {custodian} ({c_closed_why.get(k, 'closed')})")
            why.insert(0, base)
            not_active = cstate == "NotActive"
            vclass, vnote = r["_value_class"], r["_value_note"]
            if active == "Inactive":
                res = NO_ACTION
                why.append("already inactive in Orion")
            elif cancel_set:
                res = ALREADY
                why.append("Orion cancel date already set")
            elif vclass == "zero":
                res = EXCLUDED_ZERO
                why.append("Orion value 0; excluded from cancellation")
            else:
                res = CANCEL
                why.append(vnote)
                if not_active:
                    chk.append(f"{custodian} status is not '{prof['open_status_values'][0]}'; confirm the account is "
                               "really closed (not, for example, a login/access issue) before cancelling")
                if vclass == "unclear":
                    chk.append(vnote)
                if r["_uma"]:
                    chk.append("UMA account; confirm with the UMA team before cancelling")
                if k in nm_orion:
                    chk.append("Similar account number at custodian: " + ", ".join(sorted(nm_orion[k])))
                if k in o_keys_other:
                    chk.append("Same account number also listed in Orion under another custodian")
        results.append(res)
        reasons.append("; ".join(why + chk))
        flags.append("Yes" if chk else "")
    o["Custodian State"] = cust_state_col
    o["Result"] = results
    o["Reason"] = reasons
    o["Verify First"] = flags

    # ---------------- Categorize custodian rows ----------------
    active_keys = set(o.loc[(o["Account Key"] != "") & (~o["Result"].isin([NO_ACTION, INACTIVE_OPEN, EXCLUDED_LIST]))
                            & (o["Orion Active (interpreted)"] != "Inactive"), "Account Key"])
    inactive_keys = set(o.loc[o["Account Key"] != "", "Account Key"]) - active_keys
    c_results, c_reasons, c_flags = [], [], []
    for _, r in c.iterrows():
        k = r["Account Key"]
        if k == "":
            c_results.append(CUST_BLANK); c_reasons.append("Blank account number"); c_flags.append(""); continue
        if k in excl:
            c_results.append(EXCLUDED_LIST)
            c_reasons.append(f"On exclusion list (row {excl[k][2]})" + (f": {excl[k][1]}" if excl[k][1] else ""))
            c_flags.append(""); continue
        if k in active_keys:
            c_results.append(MATCHED); c_reasons.append("Found in Orion"); c_flags.append(""); continue
        why, chk = [], []
        if k in inactive_keys:
            why.append("In Orion but inactive")
        else:
            why.append("Not found in Orion")
        if r["_acct_warning"]:
            chk.append(r["_acct_warning"])
        if r["_unusual"]:
            chk.append(f"Account number has {r['_unusual']}")
        if r["_format_exception"]:
            chk.append("Account number does not fit the expected format")
        if k in c_dups:
            chk.append(f"Account number appears more than once in the {custodian} file")
        if k in nm_cust:
            chk.append("Similar account number in Orion: " + ", ".join(sorted(nm_cust[k])))
        if k in o_keys_other:
            chk.append("Listed in Orion under a different custodian")
        vclass = "n/a"
        if "account_value" in c.columns:
            raw_v = r["account_value"]
            amt, ok = parse_money(raw_v)
            if _clean_text(raw_v) == "" or not ok or amt is None:
                vclass = "unclear"
                vnote = f"{custodian} value {'blank' if _clean_text(raw_v) == '' else 'unreadable'}"
            elif abs(amt) < 0.005:
                vclass, vnote = "zero", f"{custodian} value 0"
            else:
                vclass, vnote = "value", f"{custodian} value {amt:,.2f}"
            why.append(vnote)
        state = r["Custodian State"]
        if state == "Closed":
            c_results.append(CUST_CLOSED)
            c_reasons.append("; ".join([f"Closed at {custodian} ({r['_state_reason']})"] + why + chk))
            c_flags.append(""); continue
        if state == "NotActive":
            c_results.append(CUST_NOT_ACTIVE)
            c_reasons.append("; ".join([f"Not active at {custodian} ({r['_state_reason']})"] + why + chk))
            c_flags.append(""); continue
        if state == "Unrecognized":
            chk.append(f"{custodian} status unclear ({r['_state_reason']}); treated as open")
        if vclass == "zero":
            c_results.append(CUST_IGNORED_ZERO); c_reasons.append("; ".join(why + chk)); c_flags.append(""); continue
        if vclass == "unclear":
            chk.append(vnote)
        c_results.append(MISSING)
        c_reasons.append("; ".join(why + chk))
        c_flags.append("Yes" if chk else "")
    c["Result"] = c_results
    c["Reason"] = c_reasons
    c["Verify First"] = c_flags

    near_df = pd.DataFrame(nm, columns=["Orion Account Key", "Custodian Account Key", "Difference"])

    # ---------------- Control totals ----------------
    controls = []

    def ctl(desc, expected, actual):
        controls.append((desc, int(expected), int(actual), int(expected) == int(actual)))

    ctl("Orion rows read = in scope + other custodians", orion_file.total_rows, len(o) + len(excluded))
    ctl("Orion in-scope rows = sum of Orion results", len(o),
        sum((o["Result"] == cat).sum() for cat in ORION_CATEGORIES))
    ctl("Custodian rows read = sum of custodian results", cust_file.total_rows,
        sum((c["Result"] == cat).sum() for cat in CUST_CATEGORIES))
    ctl("Every Orion row has exactly one result", len(o), o["Result"].notna().sum())
    ctl("Every custodian row has exactly one result", len(c), c["Result"].notna().sum())
    matched_keys_o = {k for k in active_keys if k in c_states}
    matched_keys_c = set(c.loc[c["Result"] == MATCHED, "Account Key"])
    ctl("Distinct matched account numbers agree (Orion side vs custodian side)",
        len(matched_keys_o), len(matched_keys_c))
    passed = all(p for *_, p in controls)

    # Completeness: totals printed in the report title block vs. the sum of the rows read.
    completeness = []
    for chk in prof.get("completeness_checks", []):
        lbl, fld = chk["preamble_label"], chk["column"]
        desc = f"{custodian} report '{lbl}' total equals the sum of rows read"
        rep_amt, rep_ok = parse_money(cust_file.preamble.get(lbl))
        if lbl not in cust_file.preamble or fld not in c.columns or not rep_ok or rep_amt is None:
            completeness.append((desc, "not available", "", None))
            continue
        total, bad = 0.0, 0
        for v in c[fld].tolist():
            amt, ok = parse_money(v)
            if not ok:
                bad += 1
            elif amt is not None:
                total += amt
        ok = bad == 0 and abs(round(total, 2) - round(rep_amt, 2)) < 0.005
        completeness.append((desc, round(rep_amt, 2), round(total, 2), ok))

    return ReconResult(
        custodian=custodian, orion_selected_values=sorted(selected), orion=o, custodian_df=c,
        near_matches=near_df, excluded_orion=excluded, controls=controls, passed=passed,
        orion_file=orion_file, cust_file=cust_file, run_time=run_time,
        config_snapshot={"orion": ocfg, "custodian": prof}, completeness=completeness,
        status_summary=_status_summary(c),
        match_info=_match_info(c),
        exclusion_info={
            "entries": excl,
            "orion_rows": int((o["Result"] == EXCLUDED_LIST).sum()),
            "custodian_rows": int((c["Result"] == EXCLUDED_LIST).sum()),
            "unused": sorted(v[0] for k_, v in excl.items()
                             if k_ not in set(o["Account Key"]) and k_ not in set(c["Account Key"])),
        },
    )


# ---------------------------------------------------------------------------
# Output: ONE workbook for all custodians in the run
#   Summary
#   Cancel in Orion      Orion rows: CANCEL
#   Not in Orion         Custodian rows: MISSING (open at custodian, value not 0, not active in Orion)
#   Do Not Cancel        Orion rows: every other result
#   Ignored              Custodian rows: closed, zero value, or blank account number (not in Orion)
#   UMA Sleeves          information (each row also appears on one of the tabs above)
#   Account Number Check information
#   Audit Log
# Custodian rows found in Orion are represented by their Orion row.
# Uncertain items carry Verify First = "Yes" and a reason; there is no separate review tab.
# ---------------------------------------------------------------------------
KIND_CANCEL = "Cancel"
KIND_KEEP = "Do Not Cancel"
KIND_NOT_IN = "Not in Orion"
TAB_IGNORED = "Ignored"


def tab_name(custodian: str, kind: str) -> str:
    """Per-custodian tab, e.g. 'Pershing - Cancel'. Excel limits sheet names to 31 characters."""
    return f"{custodian} - {kind}"[:31]


# Result -> kind of tab
def orion_tab(result: str) -> str | None:
    if result == EXCLUDED_LIST:
        return None  # excluded accounts are left out of the results entirely
    return KIND_CANCEL if result == CANCEL else KIND_KEEP


def cust_tab(result: str) -> str | None:
    if result == MISSING:
        return KIND_NOT_IN
    if result == EXCLUDED_LIST:
        return None  # excluded accounts are left out of the results entirely
    if result in (CUST_CLOSED, CUST_NOT_ACTIVE, CUST_IGNORED_ZERO, CUST_BLANK):
        return TAB_IGNORED
    return None  # MATCHED: represented by the Orion row


def _dv(v):
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.strftime("%m/%d/%Y")
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return v


def _col(df, name):
    return df[name].map(_dv).tolist() if name in df.columns else [""] * len(df)


def _orion_view(o: pd.DataFrame, custodian: str, with_result: bool) -> pd.DataFrame:
    d = {"Custodian": [custodian] * len(o)}
    if with_result:
        d["Orion Result"] = _col(o, "Result")
    d.update({
        "Verify First": _col(o, "Verify First"),
        "Reason": _col(o, "Reason"),
        "Account Number": _col(o, "Account Number (as in file)"),
        "Account Key": _col(o, "Account Key"),
        "UMA Status": _col(o, "UMA Status"),
        "Sleeve Name": _col(o, "Sleeve Name"),
        "Orion Account ID": _col(o, "account_id"),
        "Name": _col(o, "name"),
        "Household": _col(o, "household"),
        "Orion Custodian": _col(o, "custodian"),
        "Orion Active": _col(o, "Orion Active (interpreted)"),
        "Custodian State": _col(o, "Custodian State"),
        "Current Value": _col(o, "current_value"),
        "Cash Balance": _col(o, "cash_balance"),
        "Cancel Date": _col(o, "cancel_date"),
        "Representative": _col(o, "representative"),
        "Orion Account Type": _col(o, "account_type"),
        "Orion File Row": _col(o, "source_row"),
    })
    return pd.DataFrame(d)


CUST_COLUMNS = [  # (field, title) shown for every custodian; blank where a custodian has no such column
    ("name", "Name"), ("registration", "Registration Type"), ("account_value", "Custodian Account Value"),
    ("status", "Custodian Status"), ("STATE", "Custodian State"), ("open_date", "Open Date"),
    ("close_date", "Close Date"), ("last_activity", "Date of Last Activity"),
    ("last_maintenance", "Last Maintenance Date"), ("institution", "Financial Institution"),
    ("ip_name", "IP Name / Rep Code"), ("advisor", "Advisor"),
]


def _cust_view(c: pd.DataFrame, custodian: str, first: dict | None = None) -> pd.DataFrame:
    d = {"Custodian": [custodian] * len(c)}
    if first is not None:
        d["Verify First"] = _col(c, "Verify First")
        d["Tracking"] = [first.get(k, ("", ""))[0] for k in c["Account Key"]]
        d["First Seen"] = [first.get(k, ("", ""))[1] for k in c["Account Key"]]
    else:
        d["Ignored Because"] = [x.replace("Ignored - ", "").replace("Excluded - ", "") for x in c["Result"]]
    d["Reason"] = _col(c, "Reason")
    d["Account Number"] = _col(c, "Account Number (as in file)")
    d["Account Key"] = _col(c, "Account Key")
    for fld, title in CUST_COLUMNS:
        d[title] = _col(c, "Custodian State" if fld == "STATE" else fld)
    d["Custodian File Row"] = _col(c, "source_row")
    return pd.DataFrame(d)


ACCT_CATEGORY_ACTION = {
    "Standard": "Matched normally",
    "UMA sleeve": "Matched on the parent account; Verify First if it would be cancelled",
    "UMA parent (has sleeves)": "Matched normally; Verify First if it would be cancelled",
    "Closed marker": "Treated as closed at the custodian",
    "Unusual characters": "Never accepted silently; marked Verify First",
    "Does not fit expected format": "Marked Verify First if it would be cancelled or added",
    "Stored as a number in Excel": "Leading zeros may be lost; marked Verify First",
    "Damaged (scientific notation)": "Marked Verify First; re-export the file",
    "Blank": "Orion: Verify First; custodian: Ignored",
}


def _acct_category(r, side_markers=()) -> str:
    if r["Account Key"] == "":
        return "Blank"
    w = r["_acct_warning"]
    if w.startswith("Scientific"):
        return "Damaged (scientific notation)"
    if r["_unusual"]:
        return "Unusual characters"
    if w:
        return "Stored as a number in Excel"
    if r["_format_exception"]:
        return "Does not fit expected format"
    if r.get("Sleeve Name"):
        return "UMA sleeve"
    if r.get("_uma"):
        return "UMA parent (has sleeves)"
    if any(m in r["Account Number (as in file)"] for m in side_markers):
        return "Closed marker"
    return "Standard"


def account_number_check(res: ReconResult) -> pd.DataFrame:
    cust = res.custodian
    markers = tuple(res.config_snapshot["custodian"].get("closed_if_account_contains", []))
    rows = []
    for side, df, mk in ((f"Orion ({cust} accounts)", res.orion, ()), (f"{cust} file", res.custodian_df, markers)):
        cats = [_acct_category(r, mk) for _, r in df.iterrows()]
        counts = pd.Series(cats).value_counts() if cats else pd.Series(dtype=int)
        for cat in ACCT_CATEGORY_ACTION:
            if cat not in counts.index:
                continue
            ex = [a for a, c_ in zip(df["Account Number (as in file)"], cats) if c_ == cat][:3]
            label = cat + (" " + ", ".join(f"'{m}'" for m in markers) if cat == "Closed marker" else "")
            rows.append((cust, "Category", side, label, int(counts[cat]), "; ".join(e or "(blank)" for e in ex),
                         ACCT_CATEGORY_ACTION[cat]))
        shapes = df["_shape"].value_counts()
        total = int(shapes.sum()) if len(shapes) else 0
        for shp, n in shapes.items():
            ex = df.loc[df["_shape"] == shp, "Account Number (as in file)"].head(3).tolist()
            note = "Most common pattern" if shp == shapes.index[0] else (
                "Rare pattern - check" if total and n / total < 0.01 else "")
            rows.append((cust, "Pattern", side, shp, int(n), "; ".join(e or "(blank)" for e in ex), note))
    return pd.DataFrame(rows, columns=["Custodian", "Section", "File", "Category / Pattern", "Accounts",
                                       "Examples", "What happens"])


def _sort(df: pd.DataFrame) -> pd.DataFrame:
    """Clean rows before Verify First rows, then by account number (sleeves after their parent)."""
    if df.empty:
        return df.reset_index(drop=True)
    keys = [k for k in ("Verify First", "Account Key", "Sleeve Name") if k in df.columns]
    return df.sort_values(keys, kind="stable").reset_index(drop=True)


def build_tabs(results: list, tracking: dict | None = None) -> dict[str, pd.DataFrame]:
    """
    Ordered tabs: for each custodian -> '<C> - Cancel', '<C> - Do Not Cancel', '<C> - Not in Orion';
    then 'Ignored' (custodian accounts not in Orion that are closed, zero value, blank or excluded).
    tracking: {custodian: {account_key: (status text, first seen)}}.
    """
    out, ignored = {}, []
    for res in results:
        cust, o, c = res.custodian, res.orion, res.custodian_df
        ok = o["Result"].map(orion_tab)
        ck = c["Result"].map(cust_tab)
        first = (tracking or {}).get(cust, {})
        out[tab_name(cust, KIND_CANCEL)] = _sort(_orion_view(o[ok == KIND_CANCEL], cust, False))
        out[tab_name(cust, KIND_KEEP)] = _sort(_orion_view(o[ok == KIND_KEEP], cust, True))
        out[tab_name(cust, KIND_NOT_IN)] = _sort(_cust_view(c[ck == KIND_NOT_IN], cust, first))
        ignored.append(_cust_view(c[ck == TAB_IGNORED], cust))
    out[TAB_IGNORED] = pd.concat(ignored, ignore_index=True) if ignored else pd.DataFrame()
    return out


def tab_controls(res: ReconResult, tabs: dict[str, pd.DataFrame]) -> list:
    """Every Orion row and every custodian row of this custodian lands in exactly one tab,
    and no tab contains another custodian's rows."""
    cust = res.custodian
    t_c, t_k, t_n = (tabs[tab_name(cust, k)] for k in (KIND_CANCEL, KIND_KEEP, KIND_NOT_IN))
    ig = tabs[TAB_IGNORED]
    n_ig = int((ig["Custodian"] == cust).sum()) if len(ig) else 0
    c_found = int((res.custodian_df["Result"] == MATCHED).sum())
    c_excl = int((res.custodian_df["Result"] == EXCLUDED_LIST).sum())
    foreign = sum(int((t["Custodian"] != cust).sum()) for t in (t_c, t_k, t_n) if len(t))
    out = [
        (f"{cust}: Orion rows in scope = Cancel + Do Not Cancel + excluded", len(res.orion),
         len(t_c) + len(t_k) + int((res.orion["Result"] == EXCLUDED_LIST).sum())),
        (f"{cust}: custodian rows = found in Orion + Not in Orion + Ignored + excluded",
         res.cust_file.total_rows, c_found + len(t_n) + n_ig + c_excl),
        (f"{cust}: rows from another custodian on {cust} tabs", 0, foreign),
    ]
    return [(d, int(e), int(a), int(e) == int(a)) for d, e, a in out]


def audit_account_summary(res: ReconResult) -> list:
    """Account-number analysis and UMA summary, for the Audit Log."""
    chk = account_number_check(res)
    out = []
    for side in chk["File"].unique():
        cat = chk[(chk["File"] == side) & (chk["Section"] == "Category")]
        out.append((f"[{res.custodian}] Account numbers - {side}",
                    "; ".join(f"{r['Category / Pattern']}: {r['Accounts']}" for _, r in cat.iterrows()) or "none"))
        pat = chk[(chk["File"] == side) & (chk["Section"] == "Pattern")]
        out.append((f"[{res.custodian}] Account number patterns - {side}",
                    "; ".join(f"{r['Category / Pattern']} x{r['Accounts']}"
                              + (" (rare - check)" if r["What happens"].startswith("Rare") else "")
                              for _, r in pat.iterrows()) or "none"))
    o = res.orion
    out.append((f"[{res.custodian}] UMA sleeves",
                f"{int((o['Sleeve Name'] != '').sum())} sleeve row(s) under "
                f"{o.loc[o['_uma'], 'Account Key'].nunique()} parent account(s)"))
    return out


# ---------------------------------------------------------------------------
# Not in Orion tracking across runs ("added back over time")
# ---------------------------------------------------------------------------
HISTORY_COLUMNS = ["custodian", "account_key", "account_number", "first_seen", "last_seen", "status",
                   "resolved_on", "resolution"]


def update_tracking(history: pd.DataFrame, results: list, run_time: datetime):
    """
    history: previous tracking table (may be empty). Returns (new_history, tracking, resolved_df).
    Only custodians present in this run are updated. An account leaves the open list when it is
    no longer on Not in Orion; the reason is recorded (found in Orion, closed/zero, or no longer listed).
    """
    stamp = run_time.strftime("%m/%d/%Y")
    hist = history.copy() if len(history) else pd.DataFrame(columns=HISTORY_COLUMNS)
    for col in HISTORY_COLUMNS:
        if col not in hist.columns:
            hist[col] = ""
    hist = hist.fillna("").astype(str)
    tracking, resolved = {}, []
    for res in results:
        cust = res.custodian
        c = res.custodian_df
        now = c[c["Result"] == MISSING]
        now_keys = dict(zip(now["Account Key"], now["Account Number (as in file)"]))
        status_by_key = {}
        for k, r in zip(c["Account Key"], c["Result"]):
            if k:
                status_by_key.setdefault(k, set()).add(r)
        open_mask = (hist["custodian"] == cust) & (hist["status"] == "Open")
        prev_open = dict(zip(hist.loc[open_mask, "account_key"], hist.loc[open_mask].index))
        tr = {}
        for k, acct in now_keys.items():
            if k in prev_open:
                i = prev_open[k]
                hist.loc[i, "last_seen"] = stamp
                tr[k] = (f"Outstanding since {hist.loc[i, 'first_seen']}", hist.loc[i, "first_seen"])
            else:
                hist.loc[len(hist)] = [cust, k, acct, stamp, stamp, "Open", "", ""]
                tr[k] = ("New this run", stamp)
        for k, i in prev_open.items():
            if k in now_keys:
                continue
            st = status_by_key.get(k, set())
            pat = res.config_snapshot["custodian"].get("account_pattern")
            if pat and not re.fullmatch(pat, k) and not st:
                hist.loc[i, ["status", "resolved_on", "resolution"]] = [
                    "Resolved", stamp, "Removed - not an account (report text)"]
                continue  # not reported as a resolution: it was never a real account
            how = ("Now on the exclusion list" if EXCLUDED_LIST in st else
                   "Now found in Orion" if MATCHED in st else
                   "Now closed or zero value at custodian" if st & {CUST_CLOSED, CUST_IGNORED_ZERO} else
                   "Now not active at custodian" if CUST_NOT_ACTIVE in st else
                   "No longer in the custodian file")
            hist.loc[i, ["status", "resolved_on", "resolution"]] = ["Resolved", stamp, how]
            resolved.append((cust, k, hist.loc[i, "account_number"], hist.loc[i, "first_seen"], stamp, how))
        tracking[cust] = tr
    resolved_df = pd.DataFrame(resolved, columns=["Custodian", "Account Key", "Account Number", "First Seen",
                                                  "Resolved On", "How Resolved"])
    return hist.reset_index(drop=True), tracking, resolved_df


# ---------------------------------------------------------------------------
# Excel export (OnePoint brand standards)
# ---------------------------------------------------------------------------
STEEL_BLUE = "01465A"
MANDARIN = "FF8500"
LINEN = "F5F4ED"
GRAPHITE = "171A1F"
FOOTER_TEXT = "INTERNAL USE ONLY | OnePoint BFG Wealth Partners"
AI_DISCLAIMER = (
    'OnePoint BFG Wealth Partners ("OnePoint BFG") often uses Artificial Intelligence ("AI") in the '
    "generation of reports such as the above. OnePoint BFG and its employees are bound by all applicable "
    "Firm policies and procedures when using AI. AI is subject to risks and limitations. OnePoint BFG has "
    "established policies and procedures to ensure all AI generated material goes through human review "
    "prior to dissemination. For additional information regarding AI, please refer to OnePoint BFG's ADV 2A."
)
TEXT_COLUMNS = {"Account Number", "Account Key", "Parent Account"}


def export_workbook(results: list, preparer: str, run_time: datetime, tabs: dict,
                    resolved: pd.DataFrame | None = None, run_audit: list | None = None,
                    show_keep_tabs: bool = False) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    head_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    head_fill = PatternFill("solid", fgColor=STEEL_BLUE)
    linen_fill = PatternFill("solid", fgColor=LINEN)
    body_font = Font(name="Arial", size=10, color=GRAPHITE)
    bold_font = Font(name="Arial", size=10, bold=True, color=GRAPHITE)
    title_font = Font(name="Georgia", bold=True, size=16, color=STEEL_BLUE)
    section_font = Font(name="Georgia", bold=True, size=12, color=STEEL_BLUE)
    accent = Border(bottom=Side(style="medium", color=MANDARIN))
    wrap = Alignment(wrap_text=True, vertical="top")
    flag_fill = PatternFill("solid", fgColor="FFE8CC")  # Verify First rows

    controls_all = [x for res in results for x in list(res.controls) + tab_controls(res, tabs)]
    passed = all(x[-1] for x in controls_all)

    def page_setup(ws, title):
        ws.oddHeader.left.text = "OnePoint BFG Wealth Partners"
        ws.oddHeader.right.text = title
        ws.oddFooter.left.text = FOOTER_TEXT
        ws.oddFooter.right.text = "Page &P of &N"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.page_margins.left = ws.page_margins.right = 1.0
        ws.page_margins.top = ws.page_margins.bottom = 1.0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

    def header_row(ws, row, names):
        for j, name in enumerate(names, start=1):
            cell = ws.cell(row=row, column=j, value=name)
            cell.font, cell.fill = head_font, head_fill
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    def write_table(ws, df: pd.DataFrame):
        cols = list(df.columns)
        header_row(ws, 1, cols)
        vf = cols.index("Verify First") if "Verify First" in cols else None
        for i, row in enumerate(df.itertuples(index=False), start=2):
            flagged = vf is not None and row[vf] == "Yes"
            fill = flag_fill if flagged else (linen_fill if i % 2 == 1 else None)
            for j, v in enumerate(row, start=1):
                if v is None:
                    v = ""
                if isinstance(v, str) and v.startswith(("=", "+", "-", "@")):
                    v = "'" + v  # source text must never act as a formula
                cell = ws.cell(row=i, column=j, value=v)
                cell.font = body_font
                if cols[j - 1] in TEXT_COLUMNS:
                    cell.number_format = "@"
                if fill:
                    cell.fill = fill
        for j, name in enumerate(cols, start=1):
            sample = [len(str(x)) for x in df.iloc[:, j - 1].tolist()[:500]] if len(df) else []
            ws.column_dimensions[get_column_letter(j)].width = min(max(10, max([len(str(name))] + sample) + 2), 60)
        ws.freeze_panes = "A2"
        if cols:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(1, len(df) + 1)}"

    def col_of(tab, name):
        cols = list(tabs[tab].columns)
        return get_column_letter(cols.index(name) + 1)

    q = lambda t: "'" + t[:31] + "'"
    custs = [r.custodian for r in results]

    # ---------------- Summary ----------------
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    page_setup(ws, "Reconciliation Summary")
    ws["A1"] = "OnePoint BFG Wealth Partners"
    ws["A1"].font = Font(name="Arial", size=10, color=STEEL_BLUE)
    ws["A2"] = "Account Cancellation Reconciliation: Orion vs " + ", ".join(custs)
    ws["A2"].font = title_font
    for col in range(1, 4):
        ws.cell(row=2, column=col).border = accent
    ws["A3"] = f"Run {run_time.strftime('%m/%d/%Y %I:%M %p')}  |  Prepared by {preparer}"
    ws["A3"].font = body_font
    r0 = 6
    header_row(ws, r0, ["Custodian", "Cancel in Orion", "Not in Orion (Add to Orion)"])
    for i, cust in enumerate(custs, start=r0 + 1):
        vals = [cust] + [f"=COUNTA({q(tab_name(cust, kk))}!A:A)-1" for kk in (KIND_CANCEL, KIND_NOT_IN)]
        for j, v in enumerate(vals, start=1):
            cell = ws.cell(row=i, column=j, value=v)
            cell.font = Font(name="Arial", size=12, color=GRAPHITE, bold=(j == 1))
            if j > 1:
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="center")
            if (i - r0) % 2 == 0:
                cell.fill = linen_fill
    tot = r0 + len(custs) + 1
    ws.cell(row=tot, column=1, value="Total").font = Font(name="Arial", size=12, bold=True, color=STEEL_BLUE)
    for j in range(2, 4):
        L = get_column_letter(j)
        c_ = ws.cell(row=tot, column=j, value=f"=SUM({L}{r0 + 1}:{L}{tot - 1})")
        c_.font = Font(name="Arial", size=12, bold=True, color=STEEL_BLUE)
        c_.number_format = "#,##0"
        c_.alignment = Alignment(horizontal="center")
        c_.border = Border(top=Side(style="thin", color=STEEL_BLUE))
    ws.cell(row=tot, column=1).border = Border(top=Side(style="thin", color=STEEL_BLUE))
    status = ("All accuracy checks passed. Details are on the Audit Log tab." if passed
              else "ACCURACY CHECKS FAILED - DO NOT USE THESE RESULTS. See the Audit Log tab.")
    row_ = tot + 2
    ws.cell(row=row_, column=1, value=status).font = Font(name="Arial", size=10, bold=True,
                                                          color=STEEL_BLUE if passed else "C00000")
    for r in results:
        mi = r.match_info
        if mi.get("low"):
            row_ += 1
            ws.cell(row=row_, column=1, value=(
                f"WARNING: only {mi['found']:,} of {mi['considered']:,} {r.custodian} accounts ({mi['rate']:.0%}) "
                f"were found in Orion. Confirm the account numbers are recorded the same way in Orion and at "
                f"{r.custodian} before acting on the {r.custodian} tabs.")).font = \
                Font(name="Arial", size=10, bold=True, color="C00000")
    row_ += 1
    ws.cell(row=row_, column=1, value=(
        "Rows shaded orange on the custodian tabs are marked Verify First: check the Reason column before acting."
    )).font = Font(name="Arial", size=9, italic=True, color=GRAPHITE)
    row_ += 2
    ws.cell(row=row_, column=1, value=FOOTER_TEXT).font = Font(name="Arial", size=9, bold=True, color=STEEL_BLUE)
    row_ += 1
    ws.cell(row=row_, column=1, value=AI_DISCLAIMER).font = Font(name="Arial", size=8, color=GRAPHITE)
    ws.merge_cells(start_row=row_, start_column=1, end_row=row_, end_column=3)
    ws.cell(row=row_, column=1).alignment = wrap
    ws.row_dimensions[row_].height = 70
    ws.row_dimensions[r0].height = 30
    for j, w in enumerate([18, 24, 32], start=1):
        ws.column_dimensions[get_column_letter(j)].width = w

    # ---------------- Data tabs (in build order) ----------------
    for name, df in tabs.items():
        if name.endswith(" - " + KIND_KEEP) and not show_keep_tabs:
            continue  # still built and checked in memory; summarised on the Audit Log
        wt = wb.create_sheet(name[:31])
        page_setup(wt, name)
        if df.empty and not len(df.columns):
            df = pd.DataFrame(columns=["Custodian", "Note"])
        write_table(wt, df)

    # ---------------- Audit log ----------------
    wa = wb.create_sheet("Audit Log")
    page_setup(wa, "Audit Log")
    audit = list(run_audit or []) + [
        ("Run date and time", run_time.strftime("%m/%d/%Y %I:%M:%S %p")),
        ("Prepared by", preparer),
        ("Tool version", TOOL_VERSION),
        ("Control totals (all custodians)", "PASSED" if passed else "FAILED"),
        ("Cancel rule", "Active in Orion, no cancel date, absent from the custodian file or closed there, and Orion "
                        "Current Value not 0. Value 0 (cash 0 or blank) is excluded and listed on Do Not Cancel."),
        ("Not in Orion rule", "Open at the custodian (or status unclear), not active in Orion, and custodian value "
                              "not 0. Closed, zero-value, blank-number and excluded custodian accounts go to Ignored."),
        ("Verify First", "Every account is placed on a tab by rule; anything uncertain is marked Verify First = Yes "
                         "with the reason (UMA accounts, duplicates, similar numbers, unusual characters, unclear "
                         "values or statuses, other-custodian listings)."),
        ("UMA sleeve rule", "An Orion account number '<account>_<name>' is a UMA sleeve, matched on the parent "
                            "account. A sleeve or parent that would be cancelled is marked Verify First."),
        ("Exclusion rule", "Accounts on the exclusion list (see 'Exclusion list file') are left out of the results "
                           "entirely: never cancelled, never reported as Not in Orion, not listed as Ignored. "
                           "Sleeves of an excluded account are excluded too. Counts are recorded below."),
        ("Matching rule", "Exact match on the account number after removing spaces and dashes and converting to "
                          "uppercase. Similar numbers are never matched; they are reported in the Reason column."),
        ("Columns never read", "Only mapped columns are read; SSN/TIN, phone, email, address and similar columns "
                               "are ignored."),
    ]
    for res in results:
        cust = res.custodian
        for d, e, a_, ok in list(res.controls) + tab_controls(res, tabs):
            audit.append((f"[{cust}] Check: {d.replace(cust + ': ', '')}",
                          f"Expected {e:,}; actual {a_:,} -> {'Pass' if ok else 'FAIL'}"))
        tn = tabs[tab_name(cust, KIND_NOT_IN)]
        trk = tn["Tracking"] if len(tn) else pd.Series(dtype=str)
        n_res = int((resolved["Custodian"] == cust).sum()) if resolved is not None and len(resolved) else 0
        n_ig = int((tabs[TAB_IGNORED]["Custodian"] == cust).sum()) if len(tabs[TAB_IGNORED]) else 0
        mi = res.match_info
        audit += [
            (f"[{cust}] Custodian accounts found in Orion",
             f"{mi.get('found', 0):,} of {mi.get('considered', 0):,} ({mi.get('rate', 1):.0%})"
             + (" - LOW: confirm account numbers are recorded the same way" if mi.get("low") else "")),
            (f"[{cust}] Custodian status values",
             "; ".join(f"'{a_}' x{n} = {st}" for a_, st, n in res.status_summary) or "(no status column)"),
        ]
        audit += [
            (f"[{cust}] Not in Orion tracking", f"New this run {int((trk == 'New this run').sum()):,}; still "
                                                f"outstanding {int(trk.str.startswith('Outstanding').sum()):,}; "
                                                f"resolved since last run {n_res:,}"),
            (f"[{cust}] Verify First rows", "Cancel {:,}; Do Not Cancel {:,}; Not in Orion {:,}".format(
                *[int((tabs[tab_name(cust, kk)]["Verify First"] == "Yes").sum()) if len(tabs[tab_name(cust, kk)]) else 0
                  for kk in (KIND_CANCEL, KIND_KEEP, KIND_NOT_IN)])),
            (f"[{cust}] Ignored rows", f"{n_ig:,}"),
        ]
        keep = tabs[tab_name(cust, KIND_KEEP)]
        if len(keep):
            by_res = keep["Orion Result"].value_counts()
            audit.append((f"[{cust}] Do Not Cancel rows (not shown as a tab)",
                          f"{len(keep):,}: " + "; ".join(f"{k} {v:,}" for k, v in by_res.items())))
            vf = keep[keep["Verify First"] == "Yes"]
            audit.append((f"[{cust}] Do Not Cancel - Verify First",
                          " | ".join(f"{r['Account Number']}: {r['Reason']}" for _, r in vf.iterrows()) or "None"))
    for res in results:
        cust = res.custodian
        prof = res.config_snapshot["custodian"]
        o_vals = sorted(set(res.excluded_orion["custodian"].map(_clean_text))) if len(res.excluded_orion) else []
        xi = res.exclusion_info
        audit += [
            (f"[{cust}] Orion file", f"{res.orion_file.name} (SHA-256 {res.orion_file.sha256})"),
            (f"[{cust}] Orion rows read / in scope", f"{res.orion_file.total_rows:,} / {len(res.orion):,}"),
            (f"[{cust}] Orion Custodian values in scope", "; ".join(res.orion_selected_values)),
            (f"[{cust}] Other Orion Custodian values (not compared)", "; ".join(v or "(blank)" for v in o_vals) or "None"),
            (f"[{cust}] Custodian file", f"{res.cust_file.name} (SHA-256 {res.cust_file.sha256})"),
            (f"[{cust}] Custodian sheet / header row / rows read",
             f"{res.cust_file.sheet} / row {res.cust_file.header_row} / {res.cust_file.total_rows:,}"),
            (f"[{cust}] Custodian report title lines", "; ".join(f"{a}: {b}" for a, b in res.cust_file.preamble.items()
                                                                 if a in ("IBD/OFF/IP", "Quote Type", "As of")) or "None"),
            (f"[{cust}] Column mapping", "; ".join(f"{a} = {b}" for a, b in res.cust_file.column_map.items())),
            (f"[{cust}] Open / closed status values", f"{', '.join(prof.get('open_status_values', [])) or '(none)'} / "
                                                      f"{', '.join(prof.get('closed_status_values', [])) or '(none)'}"),
            (f"[{cust}] Closed if account number contains",
             ", ".join(f"'{m}'" for m in prof.get("closed_if_account_contains", [])) or "(not used)"),
            (f"[{cust}] Expected account pattern", prof.get("account_pattern") or "(none)"),
            (f"[{cust}] Completeness", "; ".join(f"report {r_}, rows {t_} -> "
                                                 f"{'Pass' if ok else 'Not available' if ok is None else 'FAIL'}"
                                                 for _, r_, t_, ok in res.completeness) or "Not configured"),
            (f"[{cust}] Exclusion list", (
                f"applied to {xi.get('orion_rows', 0)} Orion row(s) and {xi.get('custodian_rows', 0)} custodian "
                "row(s); these accounts are left out of every results tab")),
        ] + audit_account_summary(res) + [
            (f"[{cust}] Not in Orion resolved since last run",
             "; ".join(f"{r['Account Number']} ({r['How Resolved']}, first seen {r['First Seen']})"
                       for _, r in resolved[resolved["Custodian"] == cust].iterrows())
             if resolved is not None and len(resolved) and (resolved["Custodian"] == cust).any() else "None"),
            (f"[{cust}] Notes / warnings", " | ".join(res.orion_file.notes + res.cust_file.notes +
                                                      res.orion_file.warnings + res.cust_file.warnings) or "None"),
        ]
    write_table(wa, pd.DataFrame(audit, columns=["Item", "Value"]))
    wa.column_dimensions["A"].width = 46
    wa.column_dimensions["B"].width = 100
    for row in wa.iter_rows(min_row=2, min_col=2, max_col=2):
        for cell in row:
            cell.alignment = wrap

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
