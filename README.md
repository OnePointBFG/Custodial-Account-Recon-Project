# Account Cancellation Reconciliation

OnePoint BFG Wealth Partners. INTERNAL USE ONLY.

## Folders (C:\Recon\Recon_Data; set by "data_folder" in custodians.json)
- 1_Drop_Files_Here   Drop the Orion export and every custodian file here. Files are identified by their
                      column headers (Orion, Pershing, Fidelity, LPL, Goldman Sachs, Pontera); names do not matter.
- 2_Results           ONE workbook per run: Recon_<date_time>.xlsx, covering every custodian dropped.
- 3_Archive\<date_time>  Inputs, results and a tracking backup. Inputs are moved here after a successful run.
- 4_Tracking          not_in_orion_tracking.csv: Not in Orion accounts across runs (first seen, resolved).
- Exclusion list      C:\Recon\Recon_Data\5_Exclusions\Exclusions.xlsx ("exclusion_list_path" in custodians.json)
                      Columns: Custodian | Account Number | Reason. Leave Custodian blank to apply the number to
                      whichever custodian it belongs to. Listed accounts are left out of the results entirely.
                      Missing file or any invalid entry stops the run.

## Commands (VS Code terminal, in C:\Recon\account-recon)
Run:        .\.venv\Scripts\python.exe run_recon.py
Self-test:  .\.venv\Scripts\python.exe test_recon.py


## Workbook tabs (one workbook per run)
Summary | for each custodian: "<C> - Cancel", "<C> - Not in Orion" | Ignored | Audit Log
Do Not Cancel accounts are still checked on every run and summarised on the Audit Log (with any marked
Verify First listed there). Set "show_do_not_cancel_tabs": true in custodians.json to show those tabs.
Every row on a custodian's tabs belongs to that custodian (checked on every run). There is no review tab. Every account is placed by rule; uncertain ones are marked Verify First = Yes
(shaded orange) with the reason.

## Rules
- Exact match on account number after removing spaces and dashes and converting to uppercase.
- Cancel in Orion: active in Orion, no cancel date, absent from or closed at the custodian, value not 0.
- Orion value 0: not cancelled (Do Not Cancel, "Excluded - Zero Value").
- Not in Orion: open at the custodian (or status unclear), not active in Orion, custodian value not 0.
  Closed, zero-value and blank-number custodian accounts go to Ignored.
- Inactive in Orion but open at the custodian: listed on Not in Orion ("In Orion but inactive").
- UMA sleeves "<account>_<name>": matched on the parent; marked Verify First if they would be cancelled.
- Fidelity: "&" in the Account # means closed (names are never checked).
- Pershing: the Disclosures footer and the Long Market Value total must be present and agree, or the run stops.
- Fidelity: reading stops at the Wealthscape disclaimer ("For investment professional use only", etc.).
- Workbooks whose formatting is malformed (seen in some LPL exports) are read with a fallback reader that
  takes cell values directly; this is recorded on the Audit Log. Orion exports may be .xlsx or .csv.
- Any account-like row found after a report footer stops the run (no account is ever silently dropped).
- LPL: open/closed from Account Status and the Closed date; value from Account Value; Birthdate is never read.
- Goldman Sachs: open/closed from Status; value from Previous Day Total Assets; Primary Owner is never read.
- Pontera: key = Account ID; only Account status "Account ready" is active. Any other status = not active:
  in Orion -> Cancel with Verify First; not in Orion -> Ignored. Value = Balance. Client names never read.
- Every custodian: if fewer than half of its accounts (20+) are found in Orion, the Summary shows a warning
  to confirm account numbers are recorded the same way before acting.
- Ignore list: a value Excel stored as a number is accepted only if it exactly fits a fixed-length account
  format (Pershing/Fidelity 9, LPL 8). Every account number must contain at least one digit.; otherwise format the cell as Text.
- All-or-nothing: if any file or custodian fails, nothing is written and files stay in the drop folder.
- Run the self-test after any change to custodians.json or the code.
- Never place client files inside the program folder.

Brand fonts (Mixta Sharp, Graphie) are substituted with Georgia and Arial.
