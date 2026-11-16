r"""
OnePoint BFG Wealth Partners - Account Cancellation Reconciliation
INTERNAL USE ONLY.

Run with:  .\.venv\Scripts\python.exe run_recon.py

1. Drop the Orion export and every custodian file into  <data folder>\1_Drop_Files_Here
   (file names do not matter; each file is identified by its column headers).
2. Run this script. All custodians found are reconciled in one run.
3. ONE results workbook is written to  <data folder>\2_Results\Recon_<date_time>.xlsx
   Tabs: Summary | for each custodian: <C> - Cancel, <C> - Not in Orion | Ignored | Audit Log
   (Do Not Cancel is checked in memory and summarised on the Audit Log; set
    "show_do_not_cancel_tabs": true in custodians.json to show those tabs.)
4. Inputs, results and a backup of the tracking file are archived to
   <data folder>\3_Archive\<date_time>\  and the inputs are moved out of the drop folder.
5. Not in Orion accounts are tracked across runs in  <data folder>\4_Tracking  (first seen, resolved).
6. Accounts to ignore are read from the exclusion list ("exclusion_list_path" in custodians.json):
   account numbers in column A (heading in row 1 allowed). They are left out of the results entirely.

If anything fails (unrecognized file, duplicate custodian file, control totals, etc.) NOTHING is
written and the files stay in the drop folder.
"""

from __future__ import annotations

import getpass
import io
import shutil
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

import recon_engine as eng

HERE = Path(__file__).resolve().parent
ALLOWED = {".xlsx", ".xlsm", ".csv"}
DROP, RESULTS, ARCHIVE, TRACKING = "1_Drop_Files_Here", "2_Results", "3_Archive", "4_Tracking"
EXCLUSIONS, EXCLUSION_FILE = "5_Exclusions", "Exclusions.xlsx"
TRACKING_FILE = "not_in_orion_tracking.csv"


def say(msg: str = "") -> None:
    print(msg, flush=True)


def exclusion_path(config: dict, data: Path) -> Path:
    """The exclusion list: 'exclusion_list_path' in custodians.json, else Recon_Data\\5_Exclusions\\Exclusions.xlsx."""
    p = config.get("exclusion_list_path")
    return Path(p) if p else data / EXCLUSIONS / EXCLUSION_FILE


def ensure_folders(data: Path) -> bool:
    created = False
    for p in (data / DROP, data / RESULTS, data / ARCHIVE, data / TRACKING, data / EXCLUSIONS):
        if not p.exists():
            p.mkdir(parents=True)
            created = True
    if not eng.load_config().get("exclusion_list_path"):
        ensure_exclusion_template(data / EXCLUSIONS / EXCLUSION_FILE)
    return created


def ensure_exclusion_template(path: Path) -> None:
    """Create an empty exclusion list (headers only, Account Number column formatted as Text)."""
    if path.exists():
        return
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "Exclusions"
    ws.append(eng.EXCLUSION_HEADERS)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=eng.STEEL_BLUE)
    for row in range(2, 501):
        ws.cell(row=row, column=2).number_format = "@"
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 50
    wb.save(path)


def find_input_files(drop: Path) -> tuple[list[Path], list[Path]]:
    usable, ignored = [], []
    for p in sorted(drop.rglob("*")):
        if p.is_dir() or p.name.startswith("~$") or p.name.lower() == "desktop.ini":
            continue
        (usable if p.suffix.lower() in ALLOWED else ignored).append(p)
    return usable, ignored


def is_locked(path: Path) -> bool:
    try:
        with open(path, "r+b"):
            return False
    except PermissionError:
        return True
    except OSError:
        return False


def orion_values_for(orion_file: eng.LoadedFile, prof: dict, custodian: str) -> list[str]:
    vals = sorted(set(orion_file.df["custodian"].map(eng._clean_text)))
    matches = [m.lower() for m in prof.get("orion_custodian_match", [custodian])]
    return [v for v in vals if v and any(m in v.lower() for m in matches)]


def unique_stamp(data: Path) -> str:
    base = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    stamp, n = base, 1
    while (data / RESULTS / f"Recon_{stamp}.xlsx").exists() or (data / ARCHIVE / stamp).exists():
        n += 1
        stamp = f"{base}_{n}"
    return stamp


def copy_verified(src: Path, dst: Path) -> None:
    shutil.copy2(src, dst)
    if eng.file_sha256(src.read_bytes()) != eng.file_sha256(dst.read_bytes()):
        raise eng.ReconError(f"Archive copy of '{src.name}' does not match the original.")


def load_tracking(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=eng.HISTORY_COLUMNS)
    return pd.read_csv(path, dtype=str, keep_default_na=False, na_values=[])


def run(data: Path | None = None, pause: bool = True) -> int:
    config = eng.load_config()
    data = Path(data or config.get("data_folder") or HERE.parent / "Recon_Data")
    drop = data / DROP
    preparer = getpass.getuser()
    move_inputs = bool(config.get("move_inputs_to_archive", True))

    say("=" * 70)
    say("ACCOUNT CANCELLATION RECONCILIATION")
    say("OnePoint BFG Wealth Partners | INTERNAL USE ONLY")
    say("=" * 70)
    say(f"Drop folder: {drop}")

    if ensure_folders(data):
        say("\nFolders were created. Drop the Orion export and the custodian file(s) here, then run again:")
        say(f"  {drop}")
        return finish(0, pause)

    try:
        # ---------------- Identify every file ----------------
        files, ignored = find_input_files(drop)
        for p in ignored:
            say(f"  Note: ignored '{p.relative_to(drop)}' (only .xlsx, .xlsm or .csv are read).")
        if not files:
            raise eng.ReconError(f"No files found. Drop the Orion export and custodian file(s) into:\n  {drop}")
        for p in files:
            if is_locked(p):
                raise eng.ReconError(f"'{p.name}' is open in another program. Close it and run again.")

        say("\nFiles found:")
        orion_paths, cust_paths, problems = [], {}, []
        for p in files:
            matches, detail = eng.detect_file(p.name, p.read_bytes(), config)
            rel = p.relative_to(drop)
            if len(matches) == 1:
                say(f"  {matches[0]:<10} {rel}")
                (orion_paths.append(p) if matches[0] == "Orion" else cust_paths.setdefault(matches[0], []).append(p))
            elif not matches:
                problems.append(f"'{rel}' is not recognized as Orion or any custodian"
                                + (f" ({detail})" if detail else "") + ".")
            else:
                problems.append(f"'{rel}' matches more than one type ({', '.join(matches)}).")
        for cust, ps in cust_paths.items():
            if len(ps) > 1:
                problems.append(f"{len(ps)} {cust} files found; leave only one: "
                                + ", ".join(str(x.relative_to(drop)) for x in ps))
        if not orion_paths:
            problems.append("No Orion file found.")
        if not cust_paths:
            problems.append("No custodian file found.")
        if problems:
            raise eng.ReconError("\n  ".join(["The drop folder needs attention:"] + problems))

        orion_loaded = {}
        for p in orion_paths:
            of = eng.load_file(p.name, p.read_bytes(), config["orion"], "Orion")
            orion_loaded[p] = of
            say(f"\nOrion file '{p.name}': {of.total_rows:,} rows read")
            for w in of.warnings:
                say(f"  Warning: {w}")

        jobs = []
        for cust in sorted(cust_paths):
            prof = config["custodians"][cust]
            covering = [(p, of, orion_values_for(of, prof, cust)) for p, of in orion_loaded.items()]
            covering = [x for x in covering if x[2]]
            if len(orion_loaded) == 1:
                p, of = next(iter(orion_loaded.items()))
                jobs.append((cust, cust_paths[cust][0], p, of, orion_values_for(of, prof, cust)))
            elif len(covering) == 1:
                jobs.append((cust, cust_paths[cust][0], *covering[0]))
            elif not covering:
                raise eng.ReconError(f"None of the Orion files contain {cust} accounts.")
            else:
                raise eng.ReconError(f"{cust} accounts appear in more than one Orion file ("
                                     + ", ".join(p.name for p, _, _ in covering) + "). Leave only one.")

        # ---------------- Exclusion list ----------------
        excl_path = exclusion_path(config, data)
        if not excl_path.exists():
            raise eng.ReconError(f"The exclusion list was not found:\n  {excl_path}\n"
                                 "Check the file name and location (exclusion_list_path in custodians.json).")
        if is_locked(excl_path):
            raise eng.ReconError(f"'{excl_path.name}' is open in another program. Close it and run again.")
        excl_bytes = excl_path.read_bytes()
        exclusions = eng.load_exclusions(excl_path.name, excl_bytes, config)
        excl_keys = {k for v in exclusions.values() for k in v}
        n_ex = len(excl_keys)
        say(f"\nExclusion list: {excl_path}")
        say(f"  {n_ex} account number{'' if n_ex == 1 else 's'} to ignore")

        # ---------------- Reconcile every custodian ----------------
        results, failures = [], []
        for custodian, cust_path, orion_path, orion_file, selected in jobs:
            say("\n" + "-" * 70)
            say(f"{custodian}: {cust_path.name}   (Orion: {orion_path.name})")
            try:
                prof = config["custodians"][custodian]
                cust_file = eng.load_file(cust_path.name, cust_path.read_bytes(), prof, custodian)
                say(f"  {cust_file.total_rows:,} rows read")
                for n in cust_file.notes:
                    say(f"  {n}")
                for w in cust_file.warnings:
                    say(f"  Warning: {w}")
                if not selected:
                    vals = sorted(set(orion_file.df["custodian"].map(eng._clean_text)))
                    raise eng.ReconError(
                        f"No Orion Custodian value matches {prof.get('orion_custodian_match')}. "
                        f"Values seen: {', '.join(v or '(blank)' for v in vals)}. "
                        "Update orion_custodian_match in custodians.json.")
                say(f"  Orion Custodian values treated as {custodian}: {'; '.join(selected)}")
                res = eng.reconcile(orion_file, cust_file, custodian, selected, config,
                                    exclusions.get(custodian, {}))
                if res.status_summary:
                    say("  Status values: " + "; ".join(f"'{a}' x{n} = {st}" for a, st, n in res.status_summary))
                mi = res.match_info
                say(f"  {custodian} accounts found in Orion: {mi['found']:,} of {mi['considered']:,} ({mi['rate']:.0%})")
                if mi["low"]:
                    say(f"  WARNING: fewer than half of the {custodian} accounts were found in Orion. Confirm the "
                        f"account numbers are recorded the same way in Orion and in the {custodian} file.")
                xi = res.exclusion_info
                if xi["orion_rows"] or xi["custodian_rows"]:
                    say(f"  Ignored per exclusion list: {xi['orion_rows']} Orion row(s), {xi['custodian_rows']} "
                        f"{custodian} row(s) - left out of the results")
                res._paths = (orion_path, cust_path)
                if not res.passed:
                    for d, ex, ac, ok in res.controls:
                        say(f"    {'Pass' if ok else 'FAIL'}  {d}: expected {ex}, actual {ac}")
                    raise eng.ReconError("control totals do not balance")
                for desc, rep, tot, ok in res.completeness:
                    if ok is False:
                        raise eng.ReconError(f"completeness check failed: report total {rep:,.2f}, sum of rows "
                                             f"{tot:,.2f}. The file may be incomplete. Re-export it.")
                    say(f"  Completeness check: {'PASSED' if ok else 'not available'}"
                        + (f" (report total {rep:,.2f} = sum of rows {tot:,.2f})" if ok else ""))
                results.append(res)
            except eng.ReconError as e:
                say(f"  STOPPED: {e}")
                failures.append(custodian)
        if not failures and excl_keys:
            seen = set()
            for res in results:
                seen |= set(res.orion["Account Key"]) | set(res.custodian_df["Account Key"])
            unused = sorted(t for k in excl_keys - seen
                            for t in {v[k][0] for v in exclusions.values() if k in v})
            if unused:
                say(f"\nNote: {len(unused)} exclusion entr{'y was' if len(unused) == 1 else 'ies were'} not found in "
                    f"any file this run (check for typos or closed accounts): {', '.join(unused)}")
        else:
            unused = []
        if failures:
            raise eng.ReconError("No results were written because these custodians could not be completed: "
                                 + ", ".join(failures) + ". Fix the messages above and run again.")

        # ---------------- Tracking, tabs, final controls ----------------
        run_time = datetime.now()
        stamp = unique_stamp(data)
        track_path = data / TRACKING / TRACKING_FILE
        history = load_tracking(track_path)
        new_history, tracking, resolved = eng.update_tracking(history, results, run_time)
        tabs = eng.build_tabs(results, tracking)
        all_ctl = []
        for res in results:
            all_ctl += eng.tab_controls(res, tabs)
        if not all(ok for *_, ok in all_ctl):
            for d, ex, ac, ok in all_ctl:
                say(f"    {'Pass' if ok else 'FAIL'}  {d}: expected {ex}, actual {ac}")
            raise eng.ReconError("Output tab totals do not balance. No results were written.")

        out_path = data / RESULTS / f"Recon_{stamp}.xlsx"
        arch_dir = data / ARCHIVE / stamp
        run_audit = [("Results file", str(out_path)), ("Archive folder", str(arch_dir)),
                     ("Tracking file", str(track_path)),
                     ("Exclusion list file", f"{excl_path} (SHA-256 {eng.file_sha256(excl_bytes)}; "
                                             f"{n_ex} account number{'' if n_ex == 1 else 's'})"),
                     ("Exclusion entries not found in any file this run", ", ".join(unused) or "None"),
                     ("File identification", "; ".join(f"{p.name} = Orion" for p in orion_paths) + "; " +
                      "; ".join(f"{ps[0].name} = {c}" for c, ps in sorted(cust_paths.items())))]
        xlsx = eng.export_workbook(results, preparer, run_time, tabs, resolved, run_audit,
                                   show_keep_tabs=bool(config.get("show_do_not_cancel_tabs", False)))
        if out_path.exists():
            raise eng.ReconError(f"Results file already exists and will not be overwritten: {out_path}")
        out_path.write_bytes(xlsx)

        # ---------------- Archive, tracking, move inputs ----------------
        arch_dir.mkdir(parents=True)
        used = list(dict.fromkeys([p for r in results for p in r._paths]))
        arch_names = {}
        for p in used:
            name, n = p.name, 1
            while (arch_dir / name).exists():  # two inputs with the same name in different subfolders
                n += 1
                name = f"{p.stem}_{n}{p.suffix}"
            copy_verified(p, arch_dir / name)
            arch_names[p] = name
        copy_verified(out_path, arch_dir / out_path.name)
        if n_ex:
            copy_verified(excl_path, arch_dir / f"exclusions_used_{stamp}.xlsx")
        if track_path.exists():
            copy_verified(track_path, arch_dir / f"tracking_before_{stamp}.csv")
        buf = io.StringIO()
        new_history.to_csv(buf, index=False)
        track_path.write_text(buf.getvalue(), encoding="utf-8")
    except eng.ReconError as e:
        say(f"\nSTOPPED: {e}")
        return finish(1, pause)

    # ---------------- Report ----------------
    say("\n" + "=" * 70)
    say("RESULTS (one workbook for all custodians)")
    say("=" * 70)
    for res in results:
        c = res.custodian
        t_c, t_k, t_n = (tabs[eng.tab_name(c, k)] for k in (eng.KIND_CANCEL, eng.KIND_KEEP, eng.KIND_NOT_IN))
        ig = tabs[eng.TAB_IGNORED]
        vf = lambda df: int((df["Verify First"] == "Yes").sum()) if len(df) else 0
        n_zero = int((res.orion["Result"] == eng.EXCLUDED_ZERO).sum())
        n_excl = int((res.orion["Result"] == eng.EXCLUDED_LIST).sum())
        n_new = int((t_n["Tracking"] == "New this run").sum()) if len(t_n) else 0
        n_res = int((resolved["Custodian"] == c).sum()) if len(resolved) else 0
        n_ig = int((ig["Custodian"] == c).sum()) if len(ig) else 0
        say(f"{c}")
        say(f"  {c} - Cancel ............ {len(t_c):>6,}   (Verify First: {vf(t_c):,})")
        say(f"  Do Not Cancel (not a tab)  {len(t_k):>6,}   (Verify First: {vf(t_k):,} - listed on the Audit Log;"
            f" zero value: {n_zero:,}; on exclusion list: {n_excl:,})")
        say(f"  {c} - Not in Orion ...... {len(t_n):>6,}   (Verify First: {vf(t_n):,}; new this run: {n_new:,};"
            f" resolved since last run: {n_res:,})")
        say(f"  Ignored ({c}) ......... {n_ig:>6,}")
    say("\nControl totals: PASSED for every custodian")
    say(f"Results:  {out_path}")
    say(f"Archive:  {arch_dir}")

    if move_inputs:
        for p in used:
            try:
                if eng.file_sha256((arch_dir / arch_names[p]).read_bytes()) == eng.file_sha256(p.read_bytes()):
                    p.unlink()
                else:
                    say(f"  Warning: '{p.name}' was not removed; its archived copy could not be verified.")
            except OSError as e:
                say(f"  Warning: '{p.name}' could not be removed from the drop folder ({e}).")
        say("Processed files moved out of the drop folder.")
    say("\nDONE")
    return finish(0, pause)


def finish(code: int, pause: bool) -> int:
    if pause and sys.stdin and sys.stdin.isatty():
        try:
            input("\nPress Enter to close.")
        except EOFError:
            pass
    return code


if __name__ == "__main__":
    sys.exit(run(pause="--no-pause" not in sys.argv))
