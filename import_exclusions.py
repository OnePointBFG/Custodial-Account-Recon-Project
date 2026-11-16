r"""
Add account numbers to the exclusion list from a simple list file, on this computer.
OnePoint BFG Wealth Partners. INTERNAL USE ONLY.

Usage:
  .\.venv\Scripts\python.exe import_exclusions.py <Custodian> "<path to list file>" "<reason>"

Example:
  .\.venv\Scripts\python.exe import_exclusions.py Pershing "C:\Users\me\Downloads\Ignore_pershing.xlsx" "Do not reconcile"

The list file: account numbers in the first column of the first sheet. A heading in the first
row is skipped. Every other non-empty value must be a valid account number for the custodian,
otherwise nothing is imported. Accounts already on the list are not added twice.
A backup of the exclusion list is kept before any change.
"""

from __future__ import annotations

import shutil
import sys
from datetime import datetime
from pathlib import Path

import recon_engine as eng
import run_recon


def main(argv) -> int:
    if len(argv) != 4:
        print(__doc__)
        return 1
    _, cust_in, src, reason = argv
    config = eng.load_config()
    known = {c.lower(): c for c in config["custodians"]}
    cust = known.get(cust_in.strip().lower())
    if cust is None:
        print(f"STOPPED: custodian '{cust_in}' is not one of {', '.join(config['custodians'])}.")
        return 1
    src = Path(src)
    if not src.exists():
        print(f"STOPPED: file not found: {src}")
        return 1

    raw, _, _ = eng._read_raw(src.name, src.read_bytes(), None)
    pattern = config["custodians"][cust].get("account_pattern")
    import re
    found, problems = [], []
    for i in range(len(raw)):
        v = raw.iloc[i, 0]
        if isinstance(v, (int, float)) and not isinstance(v, bool) and not eng.pd.isna(v):
            problems.append(f"row {i + 1}: stored as a number in Excel (leading zeros may be lost); format as Text")
            continue
        text = eng.account_raw_text(v)[0]
        if text == "":
            continue
        key, sleeve, odd = eng.parse_account(text, "_")
        valid = key and not sleeve and "_" not in text and not odd and (not pattern or re.fullmatch(pattern, key))
        if valid:
            found.append((key, text, i + 1))
        elif i == 0:
            print(f"Row 1 treated as a heading: '{text[:60]}'")
        else:
            problems.append(f"row {i + 1}: '{text}' is not a valid {cust} account number")
    if problems:
        print("STOPPED: nothing was imported because of these rows:\n  " + "\n  ".join(problems))
        return 1
    if not found:
        print("STOPPED: no account numbers found in the first column.")
        return 1

    data = Path(config.get("data_folder") or run_recon.HERE.parent / "Recon_Data")
    run_recon.ensure_folders(data)
    path = data / run_recon.EXCLUSIONS / run_recon.EXCLUSION_FILE
    if run_recon.is_locked(path):
        print(f"STOPPED: '{path.name}' is open in Excel. Close it and run again.")
        return 1
    current = eng.load_exclusions(path.name, path.read_bytes(), config)  # validates the existing list
    existing = set(current.get(cust, {}))

    from openpyxl import load_workbook
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup = path.with_name(f"Exclusions_backup_{stamp}.xlsx")
    shutil.copy2(path, backup)
    wb = load_workbook(path)
    ws = wb.active
    last = max([r for r in range(1, ws.max_row + 1)
                if any(ws.cell(r, c).value not in (None, "") for c in (1, 2, 3))] or [1])
    added, skipped = [], []
    for key, text, _ in found:
        if key in existing or key in {k for k, _ in added}:
            skipped.append(text)
            continue
        last += 1
        ws.cell(last, 1, cust)
        cell = ws.cell(last, 2, key)
        cell.number_format = "@"
        ws.cell(last, 3, reason)
        added.append((key, text))
    wb.save(path)

    after = eng.load_exclusions(path.name, path.read_bytes(), config)  # re-validate what was written
    ok = all(k in after.get(cust, {}) for k, _ in added) and \
        len(after.get(cust, {})) == len(existing) + len(added)
    if not ok:
        shutil.copy2(backup, path)
        print("STOPPED: verification of the updated list failed; the previous list was restored.")
        return 1
    print(f"Exclusion list: {path}")
    print(f"  Added for {cust}: {len(added)}")
    for _, text in added:
        print(f"    {text}")
    if skipped:
        print(f"  Already on the list (not added again): {len(skipped)}")
    print(f"  {cust} entries now on the list: {len(after.get(cust, {}))}")
    print(f"  Backup of the previous list: {backup.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
