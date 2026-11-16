r"""
Self-test for the reconciliation. Uses FICTITIOUS data only.
Run with:  .\.venv\Scripts\python.exe test_recon.py

Builds dummy Orion, Pershing and Fidelity files that cover every rule, runs the reconciliation,
and compares EVERY row against a written answer key: the result, the output tab, and whether the
row is marked Verify First. Then runs the full folder process (single workbook, archive,
tracking across runs, all-or-nothing on failure) in a temporary folder.
The test passes only if every check passes.
"""

import io
import sys
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

import recon_engine as eng

HERE = Path(__file__).parent
TEST_DIR = HERE / "test_files"
Y, N = True, False  # Verify First expected

ORION_HEADERS = ["Account ID", "Active", "Name", "Registration ID", "Household", "Account Number",
                 "Household ID", "Custodian", "Account Type", "Model", "Current Value", "Cash Balance",
                 "Representative Name", "Fee", "Created Date", "Qualified", "Download Source", "Advisory",
                 "Cancel Date", "Edited By", "Edited Date", "Created By"]
PERSHING_HEADERS = [  # exact layout of the Pershing "Account List" export (row 17)
    "Account #", "Short Name", "Custodian", "Account Open Date", "IP Number", "Account Registration Type",
    "Account Value", "Base Currency", "Funds Due", "Equity %", "Debit Interest Rate", "Fed Call",
    "House Call/Surplus", "Buying Power", "Cash Available", "Full Name", "Trade Date Balance",
    "Settlement Date Balance", "Long Market Value", "Short Market Value", "Total Account Value",
    "Net Balance", "Total Balance", "Equity", "Cash Management Balance", "IP Name", "Account Mnemonic",
    "Master Mnemonic", "Annuity Value", "MMF Purchase/Redeem", "Linked Account Number",
    "Linked Account Cash", "As of Date", "Account Description", "Multi-CCY", "SSN/TIN", "Title",
    "Day Phone", "Evening Phone", "Primary IP", "Date Of Last Activity", "FDID", "Status",
    "Alert Access Code"]

FIDELITY_HEADERS = ['Account #', 'Primary Account Holder', 'FBSI Short Name', 'Custom Short Name', 'Reg Type', 'Reg Type Code', 'Established Date', 'Institution Name', 'Last Maintenance Date', 'Total Account Value', 'Corporate Tax Status', '3210 Affiliation Type', 'Program', 'International', 'Primary Core Symbol', 'Primary Core/Sweep', 'Div/Int Credit/Debit (TD)', 'Performance Reporting', 'Agency', 'MMKT Acct Designation', 'Asset Movement Authorization', 'Purge Exempt', 'Registration and Address', 'RR2', 'Total Portfolio Value', 'Dividend/Capital Gains', 'Tax Certification', 'W-8 Certification Date', 'W-8 Form Type', 'W8/W9 Certification Required', 'Country of Tax Residence', 'Control Relationship Ind', 'Managed Account', 'Age', 'Client Email Address', 'Primary Phone', 'Prime Broker Agreement', 'SAM Manager ID', 'Primary G#', 'Primary G# Advisor', 'Managed Account Type', 'Primary Email Address', 'Secondary G# 1', 'Secondary G# 10', 'Secondary G# 2', 'Secondary G# 3', 'Secondary G# 4', 'Secondary G# 5', 'Secondary G# 6', 'Secondary G# 7', 'Secondary G# 8', 'Secondary G# 9', 'SAM Institution Name', 'SAM Program ID', 'SAM Program Name', 'SAM FICIS ID', 'Class Code', 'Multi-Margin Relationship ID', 'Client Name', 'Client ID', 'Fractional Shares Indicator', 'Allow Consolidation', 'Commission Schedule']  # exact layout of the Fidelity account export (row 1)

# ---------------------------------------------------------------------------
# Answer keys
# Orion: (account, active, custodian, value, cash, cancel_date, expected_result, verify_first, scenario)
# ---------------------------------------------------------------------------
ORION_ROWS = [
    ("TST000001", True, "Pershing", 1000, 0, None, eng.MATCHED, N, "Active in both"),
    ("TST000002", True, "Pershing", 0, 0, None, eng.EXCLUDED_ZERO, N, "Not at custodian, value 0"),
    ("TST000003", "True", "Pershing", 0, 0, None, eng.EXCLUDED_ZERO, N, "Closed at custodian, value 0"),
    ("TST000004", True, "Pershing", 1500, 0, None, eng.CANCEL, N, "Not at custodian, has value"),
    ("TST000005", True, "Pershing", 0, 0, "01/15/2026", eng.ALREADY, N, "Cancel date already set"),
    ("TST000006", False, "Pershing", 0, 0, None, eng.NO_ACTION, N, "Already inactive in Orion"),
    ("TST000007", False, "Pershing", 0, 0, None, eng.INACTIVE_OPEN, N, "Inactive in Orion, open at custodian"),
    ("TST000008", True, "Pershing", 0, 0, None, eng.CUST_UNCLEAR, Y, "Custodian status 'Pending'"),
    (" tst-000009 ", True, "Pershing", 10, 0, None, eng.MATCHED, N, "Dash/space/lowercase standardized"),
    (None, True, "Pershing", 0, 0, None, eng.EXCLUDED_ZERO, Y, "Blank account number, value 0"),
    ("TST000011", "Maybe", "Pershing", 0, 0, None, eng.EXCLUDED_ZERO, Y, "Unrecognized Active value, value 0"),
    ("TST000012", True, "Pershing", 500, 0, None, eng.CANCEL, Y, "Duplicate in Orion (1 of 2)"),
    ("TST000012", True, "Pershing", 500, 0, None, eng.CANCEL, Y, "Duplicate in Orion (2 of 2)"),
    ("TST000013", True, "Pershing", 200, 0, None, eng.CANCEL, Y, "Similar number at custodian"),
    ("TST000014", True, "Pershing", 300, 0, None, eng.CANCEL, N, "Closed at custodian, has value"),
    ("TST000015", True, "Pershing", 0, 0, None, eng.CUST_UNCLEAR, Y, "Custodian status blank"),
    ("TST000016", True, "Fidelity", 0, 0, None, None, N, "Other custodian - out of Pershing scope"),
    ("TST000077", True, "Pershing", "$0.00", "0", None, eng.EXCLUDED_ZERO, N, "Zero stored as text"),
    ("TST000018", True, "Pershing", 0, 0, "02/01/2026", eng.CANCEL_PENDING_OPEN, Y, "Cancel date set, open at custodian"),
    (12345678, True, "Pershing", 50, 0, None, eng.CANCEL, Y, "Stored as number, wrong length"),
    ("TST000020", True, "Schwab", 0, 0, None, None, N, "Other custodian - out of scope"),
    ("TST000021", True, "Pershing", "(25.00)", 0, None, eng.CANCEL, Y, "Negative value"),
    ("TST000022", True, "Pershing", "N/A", 0, None, eng.CANCEL, Y, "Unreadable value"),
    ("TST000023", True, "Pershing", 0, 100, None, eng.CANCEL, Y, "Value 0 but cash 100"),
    ("TST000024", True, "Pershing", None, 0, None, eng.CANCEL, Y, "Value blank"),
    ("TST000025", True, "Pershing", "$2,500.00", 0, None, eng.CANCEL, N, "Value stored as text"),
]

# Pershing: (account, status, long_market_value, total_account_value, expected_result, verify_first, scenario)
PERSHING_ROWS = [
    ("TST000001", "Open", 1000.00, 1000.00, eng.MATCHED, N, "Active in both"),
    ("TST000003", "Closed", 0, 0, eng.MATCHED, N, "Closed; active in Orion"),
    ("TST000007", "Open", 250.50, 250.50, eng.MISSING, N, "In Orion but inactive -> Not in Orion"),
    ("TST000008", "Pending", 0, 0, eng.MATCHED, N, "Active in Orion"),
    ("TST000009", "Open", 10.00, 10.00, eng.MATCHED, N, "Matches standardized Orion number"),
    ("TST000014", "Closed", 0, 0, eng.MATCHED, N, "Closed; active in Orion"),
    ("TST000015", None, 0, 0, eng.MATCHED, N, "Blank status; active in Orion"),
    ("TST000016", "Open", 5000.00, 5000.00, eng.MISSING, Y, "In Orion under another custodian"),
    ("TST000018", "Open", 75.25, 75.25, eng.MATCHED, N, "Active in Orion"),
    ("TST000031", "Open", 50.00, 50.00, eng.MISSING, Y, "Similar to Orion TST000013"),
    ("PSH000501", "Open", 12345.67, 12345.67, eng.MISSING, N, "Open, has value, not in Orion"),
    ("PSH000503", "Closed", 0, 900.00, eng.CUST_CLOSED, N, "Closed, not in Orion"),
    (None, "Open", 0, 0, eng.CUST_BLANK, N, "Blank account number"),
    ("PSH000504", "Frozen", 0, 0, eng.CUST_IGNORED_ZERO, N, "Unrecognized status, value 0"),
    ("PSH000505", "Open", 1.00, 1.00, eng.MISSING, Y, "Duplicate at custodian (1 of 2)"),
    ("PSH000505", "Open", 1.00, 1.00, eng.MISSING, Y, "Duplicate at custodian (2 of 2)"),
    ("PSH000506", "Open", 0, 0, eng.CUST_IGNORED_ZERO, N, "Open, value 0"),
    ("PSH000507", "Open", 0, None, eng.MISSING, Y, "Open, value blank"),
    ("PSH000508", "Open", 0, -250.00, eng.MISSING, N, "Open, debit balance"),
    ("PSH000509", "Frozen", 0, 300.00, eng.MISSING, Y, "Unrecognized status, has value"),
]

ORION_FID_ROWS = [
    ("FID000001", True, "Fidelity", 1000, 0, None, eng.MATCHED, N, "Active in both"),
    ("FID000002", True, "Fidelity", 500, 0, None, eng.CANCEL, N, "Closed at Fidelity ('&'), has value"),
    ("FID000003", True, "Fidelity", 0, 0, None, eng.EXCLUDED_ZERO, N, "Closed at Fidelity ('&'), value 0"),
    ("FID000004", True, "Fidelity", 750, 0, None, eng.CANCEL, N, "Not at Fidelity, has value"),
    ("FID000005", True, "Fidelity", 0, 0, None, eng.EXCLUDED_ZERO, N, "Not at Fidelity, value 0"),
    ("FID000006", True, "National Financial Services", 100, 0, None, eng.MATCHED, N, "NFS counts as Fidelity"),
    ("FID000077", True, "Pershing", 10, 0, None, None, N, "Pershing account - out of Fidelity scope"),
    ("fid-000008", True, "Fidelity", 20, 0, None, eng.MATCHED, N, "Dash/lowercase standardized"),
    ("FID000009", True, "Fidelity", 300, 0, None, eng.CUST_UNCLEAR, Y, "Fidelity lists it open AND closed"),
    ("FID000010", False, "Fidelity", 0, 0, None, eng.INACTIVE_OPEN, N, "Inactive in Orion, open at Fidelity"),
    ("FID000020_Growth", True, "Fidelity", 1000, 0, None, eng.MATCHED, N, "UMA sleeve; parent open"),
    ("FID000020_Income", True, "Fidelity", 500, 0, None, eng.MATCHED, N, "Second sleeve, same parent"),
    ("FID000021_Growth", True, "Fidelity", 800, 0, None, eng.CANCEL, Y, "UMA sleeve; parent not at Fidelity"),
    ("FID000022", True, "Fidelity", 300, 0, None, eng.CANCEL, Y, "UMA parent row; not at Fidelity"),
    ("FID000022_Value", True, "Fidelity", 200, 0, None, eng.CANCEL, Y, "Its sleeve; not at Fidelity"),
    ("FID000023_Growth", True, "Fidelity", 0, 0, None, eng.EXCLUDED_ZERO, N, "UMA sleeve, value 0"),
    ("FID000025_Growth", True, "Fidelity", 100, 0, None, eng.CANCEL, Y, "Same sleeve twice (1 of 2)"),
    ("FID000025_Growth", True, "Fidelity", 100, 0, None, eng.CANCEL, Y, "Same sleeve twice (2 of 2)"),
    ("FID000030_Growth", True, "Fidelity", 400, 0, None, eng.CANCEL, Y, "UMA sleeve; parent closed ('&')"),
    ("FID000026/", True, "Fidelity", 100, 0, None, eng.MATCHED, Y, "Matches only after removing '/'"),
    ("FID00.0027", True, "Fidelity", 100, 0, None, eng.CANCEL, Y, "'.' in number, not at Fidelity"),
    ("FID000028_", True, "Fidelity", 100, 0, None, eng.MATCHED, Y, "Underscore with no sleeve name"),
    ("FID000029", True, "Fidelity", 100, 0, None, eng.MATCHED, Y, "Fidelity side has '#'"),
]

# Fidelity: (account, primary account holder, total account value, expected_result, verify_first, scenario)
FIDELITY_ROWS = [
    ("FID000001", "TEST CLIENT A", 1000.00, eng.MATCHED, N, "Active in both"),
    ("FID000002&", "TEST CLIENT B", 0, eng.MATCHED, N, "Closed ('&'); active in Orion"),
    ("FID-000003&", "TEST CLIENT C", 0, eng.MATCHED, N, "Closed ('&' and dash)"),
    ("FID000006", "TEST CLIENT D", 100.00, eng.MATCHED, N, "Active in Orion"),
    ("FID000008", "TEST CLIENT E", 20.00, eng.MATCHED, N, "Active in Orion"),
    ("FID000009", "TEST CLIENT F", 300.00, eng.MATCHED, N, "Open listing"),
    ("FID000009&", "TEST CLIENT F", 0, eng.MATCHED, N, "Closed listing of same account"),
    ("FID000010", "TEST CLIENT G", 50.00, eng.MISSING, N, "In Orion but inactive -> Not in Orion"),
    ("FXX770101", "TEST & CLIENT JOINT", 2000.00, eng.MISSING, N, "'&' in NAME only -> still open"),
    ("FXX770102", "TEST CLIENT H", 0, eng.CUST_IGNORED_ZERO, N, "Open, value 0"),
    ("FXX770103&", "TEST CLIENT I", 500.00, eng.CUST_CLOSED, N, "Closed ('&'), not in Orion"),
    ("FXX770104", "TEST CLIENT J", None, eng.MISSING, Y, "Value blank"),
    ("FID000077", "TEST CLIENT K", 10.00, eng.MISSING, Y, "In Orion under Pershing"),
    ("FID000020", "TEST CLIENT L", 1500.00, eng.MATCHED, N, "Parent of two UMA sleeves"),
    ("FID000030&", "TEST CLIENT M", 0, eng.MATCHED, N, "Closed parent of a UMA sleeve"),
    ("FID000026", "TEST CLIENT N", 100.00, eng.MATCHED, N, "Active in Orion"),
    ("FID000028", "TEST CLIENT O", 100.00, eng.MATCHED, N, "Active in Orion"),
    ("FID#000029", "TEST CLIENT P", 100.00, eng.MATCHED, N, "'#' in number; active in Orion"),
    ("FXX7701_05", "TEST CLIENT Q", 100.00, eng.MISSING, Y, "Underscore in a custodian number"),
    ("FXX770106.", "TEST CLIENT R", 100.00, eng.MISSING, Y, "'.' in a custodian number"),
]


FIDELITY_DISCLAIMER = [  # the disclaimer block printed at the bottom of Fidelity Wealthscape exports
    "For investment professional use only. Not authorized for distribution to the public as sales material in any form.",
    "WealthscapeSM is a service mark of FMR LLC.",
    "Fidelity is providing this content/report as service to your firm. You are responsible for evaluating your own practice.",
    "Please note cash activity is reported as of settlement date while any market values are reported as of trade date.",
    "The content provided herein is general in nature and is for informational purposes only.",
    "Third party marks are the property of their respective owners; all other marks are the property of FMR LLC.",
    "AI (Accrued Interest) - Represents interest accumulated since the last coupon date.",
    "Assets Held-Away",
    "The information for assets not held or custodied by National Financial Services LLC (NFS) ...",
    "Gain/loss information, Total Closing Value, Total Client Value, and/or Total Portfolio Value may include Assets Held-Away.",
    "Fidelity Institutional\u00ae (FI) provides clearing, custody, or other brokerage services through NFS or FBS.",
    "\u00a92022 FMR LLC. All rights reserved.",
    "888384.2.0",
]

LPL_HEADERS = ["Notification", "Rep ID", "Client Name", "Account Name", "LPL Account Number", "Sponsor Account",
               "Account Class", "Retirement Plan Type", "Account Value", "As of", "LPL Value", "Linked Value",
               "Combined Value", "Funds Available", "Funds Available - Previous Close",
               "Cash Balance & Margin Balance", "Cash Balance & Margin Balance - Previous Close",
               "Cash and Equivalents", "Cash and Equivalents Percent", "Objective", "Family", "Product",
               "Advisory Fee", "Birthdate", "Opened", "Closed", "Client Fee", "Account Status"]

ORION_LPL_ROWS = [
    ("1000-0001", True, "LPL Financial", 500, 0, None, eng.MATCHED, N, "Active in both (dash in Orion)"),
    ("10000002", True, "LPL Financial", 800, 0, None, eng.CANCEL, N, "Closed at LPL, has value"),
    ("10000003", True, "LPL Financial", 300, 0, None, eng.CUST_UNCLEAR, Y, "LPL status Open but Closed date set"),
    ("10000004", True, "LPL Financial", 0, 0, None, eng.EXCLUDED_ZERO, N, "Not at LPL, value 0"),
    ("10000005", True, "LPL Financial", 900, 0, None, eng.CANCEL, N, "Not at LPL, has value"),
    ("10000010", True, "Pershing", 10, 0, None, None, N, "Pershing account - out of LPL scope"),
]
# (account, status, value, closed date, expected, verify_first, scenario)
LPL_ROWS = [
    ("10000001", "Open", 500.00, None, eng.MATCHED, N, "Active in both"),
    ("10000002", "Closed", 0, "01/02/2026", eng.MATCHED, N, "Closed; active in Orion"),
    ("10000003", "Open", 300.00, "01/02/2026", eng.MATCHED, N, "Conflict; active in Orion"),
    ("2000-0006", "Open", 1200.00, None, eng.MISSING, N, "Open, has value, not in Orion"),
    ("30000007", "Open", 0, None, eng.CUST_IGNORED_ZERO, N, "Open, value 0"),
    ("40000008", "Closed", 75.00, "03/03/2026", eng.CUST_CLOSED, N, "Closed, not in Orion"),
    ("50000009", "Restricted", 50.00, None, eng.MISSING, Y, "Unrecognized status, has value"),
]


GS_HEADERS = ["Status", "Primary Owner", "Account Name", "Account Number", "Account Type",
              "Previous Day Available Cash", "Previous Day Total Assets", "Account Attribute"]

ORION_GS_ROWS = [
    ("123-45678-1", True, "Goldman Sachs", 1000, 0, None, eng.MATCHED, N, "Active in both (dashes in Orion)"),
    ("123456782", True, "Goldman Sachs", 700, 0, None, eng.CANCEL, N, "Closed at Goldman, has value"),
    ("123456783", True, "Goldman Sachs", 0, 0, None, eng.EXCLUDED_ZERO, N, "Not at Goldman, value 0"),
    ("987654321", True, "Goldman Sachs", 400, 0, None, eng.CANCEL, N, "Not at Goldman, has value"),
    ("123456785_Growth", True, "Goldman Sachs", 250, 0, None, eng.MATCHED, N, "UMA sleeve; parent open at Goldman"),
    ("123456786", False, "Goldman Sachs", 0, 0, None, eng.INACTIVE_OPEN, N, "Inactive in Orion, open at Goldman"),
    ("TST900001", True, "Pershing", 10, 0, None, None, N, "Pershing account - out of Goldman scope"),
]
# (account, status, total assets, expected, verify_first, scenario)
GS_ROWS = [
    ("123456781", "Open", 1000.00, eng.MATCHED, N, "Active in both"),
    ("123456782", "Closed", 0, eng.MATCHED, N, "Closed; active in Orion"),
    ("123-45678-5", "Open", 500.00, eng.MATCHED, N, "Parent of a UMA sleeve (dashes at Goldman)"),
    ("123456786", "Open", 300.00, eng.MISSING, N, "In Orion but inactive -> Not in Orion"),
    ("555000111", "Open", 2500.00, eng.MISSING, N, "Open, has value, not in Orion"),
    ("555000222", "Open", 0, eng.CUST_IGNORED_ZERO, N, "Open, value 0"),
    ("555000333", "Closed", 100.00, eng.CUST_CLOSED, N, "Closed, not in Orion"),
    ("555000444", "Pending", 50.00, eng.MISSING, Y, "Unrecognized status, has value"),
]


PNT_HEADERS = ["Client ID", "Client first name", "Client last name", "Risk tolerance", "Account name",
               "Financial institution", "Account type", "Account ID", "Balance", "Balance date", "Advisor", "Group",
               "Rep code", "Firm", "Last action date", "Auto-rebalance", "Auto-rebalance date", "Account status",
               "Access issue", "Account information", "Status date", "Added on", "MFA enabled", "Forwarding values",
               "MFA app client status"]

ORION_PNT_ROWS = [
    ("PNT100001", True, "Pontera", 5000, 0, None, eng.MATCHED, N, "Active in both ('Account ready')"),
    ("PNT100002", True, "Pontera", 3000, 0, None, eng.CANCEL, Y, "Pontera status 'Access issue', has value"),
    ("PNT100003", True, "Pontera", 0, 0, None, eng.EXCLUDED_ZERO, N, "Pontera status 'Disconnected', value 0"),
    ("PNT100004", True, "Pontera", 800, 0, None, eng.CANCEL, N, "Not at Pontera, has value"),
    ("PNT100005", True, "Pontera", 900, 0, None, eng.CUST_UNCLEAR, Y, "Pontera lists it ready AND not ready"),
    ("PNT100006", True, "Pontera", 100, 0, None, eng.CUST_UNCLEAR, Y, "Pontera status blank"),
    ("PNT100007", True, "Fidelity", 10, 0, None, None, N, "Fidelity account - out of Pontera scope"),
]
# (account, status, balance, expected, verify_first, scenario)
PNT_ROWS = [
    ("PNT100001", "Account ready", 5000.00, eng.MATCHED, N, "Active in both"),
    ("PNT100002", "Access issue", 3000.00, eng.MATCHED, N, "Not ready; active in Orion"),
    ("PNT100003", "Disconnected", 0, eng.MATCHED, N, "Not ready; active in Orion"),
    ("PNT100005", "Account ready", 900.00, eng.MATCHED, N, "Ready listing"),
    ("PNT100005", "Access issue", 900.00, eng.MATCHED, N, "Not-ready listing of the same account"),
    ("PNT100006", None, 100.00, eng.MATCHED, N, "Blank status; active in Orion"),
    ("ZPN770001", "Account ready", 1500.00, eng.MISSING, N, "Ready, has value, not in Orion"),
    ("ZPN770002", "Access issue", 700.00, eng.CUST_NOT_ACTIVE, N, "Not ready, not in Orion -> Ignored"),
    ("ZPN770003", "Account ready", 0, eng.CUST_IGNORED_ZERO, N, "Ready, value 0"),
    ("ZPN770004", "ACCOUNT READY", 200.00, eng.MISSING, N, "Status in capitals still counts as ready"),
]


# ---------------------------------------------------------------------------
# File builders
# ---------------------------------------------------------------------------
def write_orion(path: Path, rows) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet"
    ws.append(ORION_HEADERS)
    for i, (acct, active, cust, val, cash, cdate, *_rest) in enumerate(rows, start=1):
        row = {h: None for h in ORION_HEADERS}
        row.update({"Account ID": 800000 + i, "Active": active, "Name": f"Test Account {i}",
                    "Household": f"Test Household {i}", "Account Number": acct, "Custodian": cust,
                    "Account Type": "Individual", "Current Value": val, "Cash Balance": cash,
                    "Representative Name": "Test Rep", "Cancel Date": cdate, "Download Source": cust})
        ws.append([row[h] for h in ORION_HEADERS])
    wb.save(path)


def write_pershing(path: Path, rows, footer: bool = True) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "ExportExcel"
    lmv_total = sum(r[2] for r in rows)
    for line in ["Account List", "IBD/OFF/IP: TST/ALL/ALL", "Quote Type: Prior Day", "As of: Oct 22, 2026",
                 f"Long Market Value: {lmv_total:,.2f}", "Annuities: -", "Cash Available: 0.00",
                 "Buying Power: 0.00", "Cash Management Balance: 0.00", "Account Value: 0.00",
                 "Total Account Value: 0.00", "Equity: 0.00", "Funds Due: 0.00", "Short Market Value: 0.00",
                 "Linked Account Cash: 0.00"]:
        ws.append([line])
    ws.append([])
    ws.append(PERSHING_HEADERS)
    for i, (acct, status, lmv, tav, *_rest) in enumerate(rows, start=1):
        row = {h: None for h in PERSHING_HEADERS}
        row.update({"Account #": acct, "Short Name": f"TEST{i}", "Custodian": "PERSHING",
                    "Account Registration Type": "Individual", "Full Name": f"TEST CLIENT {i}",
                    "Long Market Value": f"{lmv:,.2f}", "Total Account Value": tav,
                    "SSN/TIN": "XXX-XX-0000", "Day Phone": "000-000-0000", "IP Name": "TEST REP",
                    "Status": status})
        ws.append([row[h] for h in PERSHING_HEADERS])
    ws.append([])
    ws.append([])
    if footer:
        ws.append(["Disclosures"])
    for line in ["This information displayed above is generated as a courtesy (test).",
                 "* All calculations exclude open/unsettled Fx Trades.", "* Please Note: test footnote."]:
        ws.append([line])
    wb.save(path)


def write_fidelity(path: Path, rows, disclaimer: bool = True) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(FIDELITY_HEADERS)
    for acct, holder, tav, *_rest in rows:
        row = {h: None for h in FIDELITY_HEADERS}
        row.update({"Account #": acct, "Primary Account Holder": holder, "FBSI Short Name": "TEST",
                    "Reg Type": "Individual", "Reg Type Code": "IND", "Total Account Value": tav,
                    "Client Email Address": "test@example.com", "Primary Phone": "000-000-0000",
                    "Registration and Address": "TEST CLIENT, 1 TEST ST", "Age": 99,
                    "Primary G# Advisor": "TEST ADVISOR"})
        ws.append([row[h] for h in FIDELITY_HEADERS])
    if disclaimer:
        ws.append([])
        for line in FIDELITY_DISCLAIMER:
            ws.append([line])
    wb.save(path)


def write_lpl(path: Path, rows) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(LPL_HEADERS)
    for acct, status, val, closed, *_rest in rows:
        row = {h: None for h in LPL_HEADERS}
        row.update({"LPL Account Number": acct, "Account Status": status, "Account Value": val,
                    "Closed": closed, "Opened": eng.datetime(2020, 1, 15), "Rep ID": "TREP", "Client Name": "TEST CLIENT",
                    "Account Class": "Brokerage", "Birthdate": "01/01/1900", "Sponsor Account": None})
        ws.append([row[h] for h in LPL_HEADERS])
    wb.save(path)


def write_gs(path: Path, rows) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(GS_HEADERS)
    for acct, status, val, *_rest in rows:
        row = {h: None for h in GS_HEADERS}
        row.update({"Status": status, "Primary Owner": "TEST OWNER PRIVATE", "Account Name": "TEST ACCOUNT",
                    "Account Number": acct, "Account Type": "Individual", "Previous Day Available Cash": 1.0,
                    "Previous Day Total Assets": val, "Account Attribute": "TEST ATTRIBUTE"})
        ws.append([row[h] for h in GS_HEADERS])
    wb.save(path)


def write_pontera(path: Path, rows) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Pontera demo"
    ws.append(PNT_HEADERS)
    for acct, status, bal, *_rest in rows:
        row = {h: None for h in PNT_HEADERS}
        row.update({"Client ID": "C1", "Client first name": "PRIVATEFIRST", "Client last name": "PRIVATELAST",
                    "Risk tolerance": "Moderate", "Account name": "TEST 401K", "Financial institution": "Test Plan Co",
                    "Account type": "401(k)", "Account ID": acct, "Balance": bal, "Advisor": "TEST ADVISOR",
                    "Rep code": "R01", "Account status": status, "MFA enabled": "Yes"})
        ws.append([row[h] for h in PNT_HEADERS])
    wb.save(path)


P_ORION = TEST_DIR / "TEST_Orion_All_Active_Accounts.xlsx"
P_ORION_PNT = TEST_DIR / "TEST_Orion_Pontera.xlsx"
P_PNT = TEST_DIR / "TEST_Pontera.xlsx"
P_ORION_GS = TEST_DIR / "TEST_Orion_Goldman.xlsx"
P_GS = TEST_DIR / "TEST_Goldman_Active.xlsx"
P_ORION_LPL = TEST_DIR / "TEST_Orion_LPL.xlsx"
P_LPL = TEST_DIR / "TEST_LPL_Active.xlsx"
P_PERSHING = TEST_DIR / "TEST_Pershing_Active.xlsx"
P_ORION_FID = TEST_DIR / "TEST_Orion_Fidelity.xlsx"
P_FIDELITY = TEST_DIR / "TEST_Fidelity_Active.xlsx"


def build_files():
    TEST_DIR.mkdir(exist_ok=True)
    write_orion(P_ORION, ORION_ROWS)
    write_pershing(P_PERSHING, PERSHING_ROWS)
    write_orion(P_ORION_FID, ORION_FID_ROWS)
    write_fidelity(P_FIDELITY, FIDELITY_ROWS)
    write_orion(P_ORION_LPL, ORION_LPL_ROWS)
    write_lpl(P_LPL, LPL_ROWS)
    write_orion(P_ORION_GS, ORION_GS_ROWS)
    write_gs(P_GS, GS_ROWS)
    write_orion(P_ORION_PNT, ORION_PNT_ROWS)
    write_pontera(P_PNT, PNT_ROWS)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
class Checker:
    def __init__(self):
        self.fails = 0
        self.passes = 0

    def check(self, ok: bool, text: str):
        ok = bool(ok)
        self.fails += not ok
        self.passes += ok
        print(f"  {'PASS' if ok else 'FAIL'}  {text}")


def open_wb(path: Path):
    """Read a workbook into memory first so the file is never left open (Windows locks open files)."""
    return load_workbook(io.BytesIO(Path(path).read_bytes()), read_only=True)


def workbook_text(xlsx: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(xlsx)) as z:
        return "".join(z.read(n).decode("utf-8", "ignore") for n in z.namelist())


def run_one(cfg, orion_path, cust_path, custodian):
    import run_recon
    of = eng.load_file(orion_path.name, orion_path.read_bytes(), cfg["orion"], "Orion")
    cf = eng.load_file(cust_path.name, cust_path.read_bytes(), cfg["custodians"][custodian], custodian)
    sel = run_recon.orion_values_for(of, cfg["custodians"][custodian], custodian)
    return eng.reconcile(of, cf, custodian, sel, cfg), sel


def check_rows(ck, res, orion_rows, cust_rows, cust_first_row):
    o_res = dict(zip(res.orion["source_row"], zip(res.orion["Result"], res.orion["Verify First"])))
    for i, row in enumerate(orion_rows):
        exp, flag, scen = row[6], row[7], row[8]
        got = o_res.get(i + 2)
        want = None if exp is None else (exp, "Yes" if flag else "")
        ck.check(got == want, f"Orion row {i + 2:>2}  {scen:<44} expected={want}  got={got}")
    c_res = dict(zip(res.custodian_df["source_row"], zip(res.custodian_df["Result"],
                                                         res.custodian_df["Verify First"])))
    for i, row in enumerate(cust_rows):
        exp, flag, scen = row[-3], row[-2], row[-1]
        got = c_res.get(i + cust_first_row)
        want = (exp, "Yes" if flag else "")
        ck.check(got == want, f"{res.custodian} row {i + cust_first_row:>2}  {scen:<41} expected={want}  got={got}")


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------
def test_rules(ck, cfg):
    print("FILE IDENTIFICATION")
    for path, want in ((P_ORION, "Orion"), (P_ORION_FID, "Orion"), (P_PERSHING, "Pershing"), (P_FIDELITY, "Fidelity")):
        got, _ = eng.detect_file(path.name, path.read_bytes(), cfg)
        ck.check(got == [want], f"{path.name:<38} identified as {got}")

    print("PERSHING")
    res_p, sel = run_one(cfg, P_ORION, P_PERSHING, "Pershing")
    ck.check(sel == ["Pershing"], f"Orion custodian values treated as Pershing: {sel}")
    check_rows(ck, res_p, ORION_ROWS, PERSHING_ROWS, 18)
    for d, e, a, ok in res_p.controls:
        ck.check(ok, f"{d}: expected {e}, actual {a}")
    for d, r, t, ok in res_p.completeness:
        ck.check(ok is True, f"{d}: report {r}, rows {t}")
    ck.check(any("Footer starting at row" in n for n in res_p.cust_file.notes), "Disclosures footer detected and ignored")
    nm = set(zip(res_p.near_matches["Orion Account Key"], res_p.near_matches["Custodian Account Key"]))
    ck.check(("TST000013", "TST000031") in nm, "Similar numbers TST000013 ~ TST000031 reported")

    print("FIDELITY")
    res_f, sel = run_one(cfg, P_ORION_FID, P_FIDELITY, "Fidelity")
    ck.check(sel == ["Fidelity", "National Financial Services"], f"Orion custodian values treated as Fidelity: {sel}")
    check_rows(ck, res_f, ORION_FID_ROWS, FIDELITY_ROWS, 2)
    for d, e, a, ok in res_f.controls:
        ck.check(ok, f"{d}: expected {e}, actual {a}")
    nm = set(zip(res_f.near_matches["Orion Account Key"], res_f.near_matches["Custodian Account Key"]))
    ck.check(nm == {("FID000027", "FID000077")}, f"Similar-number pairs exactly FID000027 ~ FID000077: {nm}")

    print("COMBINED WORKBOOK (one file, tabs per custodian)")
    hist, tracking, resolved = eng.update_tracking(eng.pd.DataFrame(), [res_p, res_f], eng.datetime.now())
    tabs = eng.build_tabs([res_p, res_f], tracking)
    for r in (res_p, res_f):
        for d, e, a, ok in eng.tab_controls(r, tabs):
            ck.check(ok, f"{d}: expected {e}, actual {a}")
    T = lambda c, k: tabs[eng.tab_name(c, k)]
    keys = lambda c, k: set(T(c, k)["Account Key"])
    for c in ("Pershing", "Fidelity"):
        for k in (eng.KIND_CANCEL, eng.KIND_KEEP, eng.KIND_NOT_IN):
            df = T(c, k)
            ck.check(len(df) == 0 or set(df["Custodian"]) == {c}, f"'{c} - {k}' contains only {c} rows")
    ck.check(keys("Pershing", eng.KIND_CANCEL) == {"TST000004", "TST000012", "TST000013", "TST000014", "12345678",
                                                    "TST000021", "TST000022", "TST000023", "TST000024", "TST000025"},
             "Pershing - Cancel holds exactly the expected accounts")
    ck.check(keys("Fidelity", eng.KIND_CANCEL) == {"FID000002", "FID000004", "FID000021", "FID000022", "FID000025",
                                                    "FID000030", "FID000027"},
             "Fidelity - Cancel holds exactly the expected accounts")
    ck.check(keys("Pershing", eng.KIND_NOT_IN) == {"TST000007", "TST000016", "TST000031", "PSH000501", "PSH000505",
                                                    "PSH000507", "PSH000508", "PSH000509"},
             "Pershing - Not in Orion holds exactly the expected accounts")
    ck.check(keys("Fidelity", eng.KIND_NOT_IN) == {"FID000010", "FXX770101", "FXX770104", "FID000077", "FXX770105",
                                                    "FXX770106"},
             "Fidelity - Not in Orion holds exactly the expected accounts")
    exp_p = sum(1 for r in ORION_ROWS if r[6] not in (None, eng.CANCEL))
    exp_f = sum(1 for r in ORION_FID_ROWS if r[6] not in (None, eng.CANCEL))
    ck.check(len(T("Pershing", eng.KIND_KEEP)) == exp_p and len(T("Fidelity", eng.KIND_KEEP)) == exp_f,
             f"Do Not Cancel row counts match the answer key: Pershing {exp_p}, Fidelity {exp_f}")
    ck.check(not (keys("Pershing", eng.KIND_CANCEL) & keys("Pershing", eng.KIND_KEEP) - {""}) and
             not (keys("Fidelity", eng.KIND_CANCEL) & keys("Fidelity", eng.KIND_KEEP) - {""}),
             "No account is on both Cancel and Do Not Cancel")
    clean = {k for c in ("Pershing", "Fidelity") for k in
             T(c, eng.KIND_CANCEL).loc[T(c, eng.KIND_CANCEL)["Verify First"] == "", "Account Key"]}
    ck.check(clean == {"TST000004", "TST000014", "TST000025", "FID000002", "FID000004"},
             "Only the 5 fully clean accounts are on Cancel without Verify First")
    fc = T("Fidelity", eng.KIND_CANCEL)
    ck.check(set(fc.loc[fc["UMA Status"] != "Standard", "Verify First"]) == {"Yes"},
             "Every UMA account on Fidelity - Cancel is marked Verify First")
    ck.check(all(set(T(c, eng.KIND_NOT_IN)["Tracking"]) == {"New this run"} for c in ("Pershing", "Fidelity")),
             "First run: every Not in Orion account is New")

    xlsx = eng.export_workbook([res_p, res_f], "Self-test", eng.datetime.now(), tabs, resolved, [])
    (TEST_DIR / "TEST_Recon_Output.xlsx").write_bytes(xlsx)
    names = load_workbook(io.BytesIO(xlsx), read_only=True).sheetnames
    ck.check(names == ["Summary", "Pershing - Cancel", "Pershing - Not in Orion",
                       "Fidelity - Cancel", "Fidelity - Not in Orion", "Ignored", "Audit Log"],
             f"Workbook tabs (no Do Not Cancel): {names}")
    au = {r[0]: r[1] for r in load_workbook(io.BytesIO(xlsx))["Audit Log"].iter_rows(min_row=2, values_only=True)}
    ck.check(str(au.get("[Pershing] Do Not Cancel rows (not shown as a tab)", "")).startswith(f"{exp_p}:"),
             "Audit Log records the Pershing Do Not Cancel count")
    vf_txt = str(au.get("[Pershing] Do Not Cancel - Verify First", ""))
    ck.check(all(k in vf_txt for k in ("TST000008", "TST000015", "TST000018")),
             "Audit Log lists the Do Not Cancel accounts marked Verify First")
    names_full = load_workbook(io.BytesIO(eng.export_workbook([res_p, res_f], "x", eng.datetime.now(), tabs,
                                                              resolved, [], show_keep_tabs=True)),
                               read_only=True).sheetnames
    ck.check("Pershing - Do Not Cancel" in names_full, "Setting show_do_not_cancel_tabs = true restores the tabs")
    ck.check(not any("review" in n.lower() for n in names), "There is no Needs Review tab")
    sm = load_workbook(io.BytesIO(xlsx))["Summary"]
    rows = [[c.value for c in r] for r in sm.iter_rows(min_row=6, max_row=9, max_col=4)]
    ck.check(rows[0] == ["Custodian", "Cancel in Orion", "Not in Orion (Add to Orion)", None],
             f"Summary headings (no Do Not Cancel): {rows[0]}")
    ck.check(not any("Do Not Cancel" in str(c.value) for r in sm.iter_rows() for c in r if c.value),
             "Do Not Cancel does not appear anywhere on the Summary")
    ok = True
    for r in rows[1:3]:
        c = r[0]
        ok &= r[1:] == [f"=COUNTA('{c} - Cancel'!A:A)-1", f"=COUNTA('{c} - Not in Orion'!A:A)-1", None]
    ck.check(ok and rows[1][0] == "Pershing" and rows[2][0] == "Fidelity",
             "Each Summary row counts only its own custodian's tabs")
    ck.check(rows[3][0] == "Total" and rows[3][1] == "=SUM(B7:B8)", "Summary Total row sums the custodian rows")
    audit_items = [r[0] for r in load_workbook(io.BytesIO(xlsx))["Audit Log"].iter_rows(min_row=2, values_only=True)]
    ck.check(sum(1 for a in audit_items if "] Check: " in a) == 18,
             "All 18 control checks (9 per custodian) are recorded on the Audit Log")
    text = workbook_text(xlsx)
    ck.check(not any(x in text for x in ("test@example.com", "000-000-0000", "XXX-XX-0000", "1 TEST ST")),
             "SSN, phone, email and address never appear in the workbook")
    ck.check(workbook_text(eng.export_workbook([res_p], "planted test@example.com", eng.datetime.now(),
                                               eng.build_tabs([res_p]), None, [])).count("test@example.com") > 0,
             "The sensitive-data search is able to detect a planted value")

    print("MISSING FOOTER")
    tmp = TEST_DIR / "TEST_Pershing_NoFooter.xlsx"
    write_pershing(tmp, PERSHING_ROWS, footer=False)
    try:
        eng.load_file(tmp.name, tmp.read_bytes(), cfg["custodians"]["Pershing"], "Pershing")
        ck.check(False, "Pershing file without Disclosures footer is rejected")
    except eng.ReconError:
        ck.check(True, "Pershing file without Disclosures footer is rejected (export may be incomplete)")
    tmp.unlink()


def test_new_rules(ck, cfg):
    print("FIDELITY DISCLAIMER")
    res_f, _ = run_one(cfg, P_ORION_FID, P_FIDELITY, "Fidelity")
    ck.check(res_f.cust_file.total_rows == len(FIDELITY_ROWS),
             f"Fidelity rows read = {len(FIDELITY_ROWS)} accounts; the {len(FIDELITY_DISCLAIMER)} disclaimer lines are not read")
    ck.check(any("Footer starting at row" in n for n in res_f.cust_file.notes), "Disclaimer detected and ignored")
    txt = " ".join(res_f.custodian_df["Account Number (as in file)"].astype(str))
    ck.check(not any(w in txt for w in ("investment professional", "Wealthscape", "FMR", "888384")),
             "No disclaimer text appears as an account")
    p = TEST_DIR / "TEST_Fidelity_NoDisclaimer.xlsx"
    write_fidelity(p, FIDELITY_ROWS, disclaimer=False)
    try:
        f2 = eng.load_file(p.name, p.read_bytes(), cfg["custodians"]["Fidelity"], "Fidelity")
        ck.check(f2.total_rows == len(FIDELITY_ROWS), "Fidelity export without the disclaimer still loads")
    except eng.ReconError:
        ck.check(False, "Fidelity export without the disclaimer still loads")
    wb = load_workbook(P_FIDELITY); ws = wb.active
    ws.append(["FID000999", "TEST", "TEST", None, None, None])  # account-like row AFTER the disclaimer
    for c_, v in zip(range(1, 8), ["FID000999", "TEST", "X", "Y", "Z", "W", "V"]):
        ws.cell(ws.max_row, c_, v)
    wb.save(p)
    try:
        eng.load_file(p.name, p.read_bytes(), cfg["custodians"]["Fidelity"], "Fidelity")
        ck.check(False, "Account data after the disclaimer stops the run (no account silently dropped)")
    except eng.ReconError:
        ck.check(True, "Account data after the disclaimer stops the run (no account silently dropped)")
    p.unlink()

    print("TRACKING CLEAN-UP OF DISCLAIMER TEXT RECORDED BY EARLIER VERSIONS")
    junk = eng.standardize_account("For investment professional use only. Not authorized")
    hist = eng.pd.DataFrame([["Fidelity", junk, "For investment professional use only...", "10/28/2026",
                              "10/28/2026", "Open", "", ""],
                             ["Fidelity", "FXX770101", "FXX770101", "10/28/2026", "10/28/2026", "Open", "", ""]],
                            columns=eng.HISTORY_COLUMNS)
    h2, tr, resolved = eng.update_tracking(hist, [res_f], eng.datetime.now())
    row = h2[h2["account_key"] == junk].iloc[0]
    ck.check(row["status"] == "Resolved" and row["resolution"] == "Removed - not an account (report text)",
             "Disclaimer text recorded earlier is closed as 'Removed - not an account'")
    ck.check(junk not in set(resolved["Account Key"]), "...and is not reported as a real resolution")
    ck.check(tr["Fidelity"].get("FXX770101", ("",))[0].startswith("Outstanding since"),
             "A real outstanding account keeps its first-seen date")

    print("LPL")
    got, _ = eng.detect_file(P_LPL.name, P_LPL.read_bytes(), cfg)
    ck.check(got == ["LPL"], f"LPL file identified as {got}")
    res_l, sel = run_one(cfg, P_ORION_LPL, P_LPL, "LPL")
    ck.check(sel == ["LPL Financial"], f"Orion custodian values treated as LPL: {sel}")
    check_rows(ck, res_l, ORION_LPL_ROWS, LPL_ROWS, 2)
    for d, e, a, ok in res_l.controls:
        ck.check(ok, f"{d}: expected {e}, actual {a}")
    ck.check(len(res_l.near_matches) == 0, "No similar-number pairs in the LPL data")
    ck.check("Birthdate" not in res_l.cust_file.column_map.values(), "LPL Birthdate column is never read")
    res_p, _ = run_one(cfg, P_ORION, P_PERSHING, "Pershing")
    tabs = eng.build_tabs([res_l, res_p, res_f])
    xlsx = eng.export_workbook([res_l, res_p, res_f], "Self-test", eng.datetime.now(), tabs, None, [])
    names = load_workbook(io.BytesIO(xlsx), read_only=True).sheetnames
    ck.check(names == ["Summary", "LPL - Cancel", "LPL - Not in Orion", "Pershing - Cancel", "Pershing - Not in Orion",
                       "Fidelity - Cancel", "Fidelity - Not in Orion", "Ignored", "Audit Log"],
             f"Three custodians in one workbook: {names}")
    for r in (res_l, res_p, res_f):
        for d, e, a, ok in eng.tab_controls(r, tabs):
            ck.check(ok, f"{d}: expected {e}, actual {a}")
    ck.check(set(tabs["LPL - Cancel"]["Account Key"]) == {"10000002", "10000005"}, "LPL - Cancel = 10000002, 10000005")
    ck.check(set(tabs["LPL - Not in Orion"]["Account Key"]) == {"20000006", "50000009"},
             "LPL - Not in Orion = 20000006, 50000009")
    ck.check("01/01/1900" not in workbook_text(xlsx), "LPL birthdates never appear in the workbook")


def break_styles(src: Path, dst: Path) -> None:
    """Copy a workbook and corrupt its formatting (empty <fill/>), like the real LPL export."""
    import re as _re
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            d = zin.read(n)
            if n == "xl/styles.xml":
                d = _re.sub(r"<fill>.*?</fill>", "<fill/>", d.decode(), count=1).encode()
            zout.writestr(n, d)


def test_readers(ck, cfg):
    print("MALFORMED FORMATTING (fallback reader)")
    import pandas as pd
    for path in (P_ORION, P_PERSHING, P_FIDELITY, P_LPL, P_ORION_LPL):
        bad = TEST_DIR / ("BROKEN_" + path.name)
        break_styles(path, bad)
        try:
            load_workbook(bad)
            ck.check(False, f"{path.name}: corrupted copy really fails in the normal reader")
        except Exception:
            ck.check(True, f"{path.name}: corrupted copy fails in the normal reader (reproduces the LPL error)")
        good_raw, _, _ = eng._read_raw(path.name, path.read_bytes(), None)
        bad_raw, _, warn = eng._read_raw(bad.name, bad.read_bytes(), None)
        norm = lambda df: [[(v.to_pydatetime() if isinstance(v, pd.Timestamp) else v) for v in r] for r in df.values.tolist()]
        ck.check(norm(good_raw) == norm(bad_raw), f"{path.name}: fallback reader returns identical cell values "
                                                  f"({good_raw.shape[0]} rows x {good_raw.shape[1]} columns)")
        ck.check(any("fallback reader" in w for w in warn), f"{path.name}: use of the fallback reader is recorded")
    bad = TEST_DIR / ("BROKEN_" + P_LPL.name)
    raw, _, _ = eng._read_raw(bad.name, bad.read_bytes(), None)
    ck.check(raw.iloc[1, LPL_HEADERS.index("Opened")] == eng.datetime(2020, 1, 15),
             "Fallback reader returns date cells as dates")
    got, _ = eng.detect_file(bad.name, bad.read_bytes(), cfg)
    ck.check(got == ["LPL"], f"Corrupted LPL export still identified as {got}")
    res_good, sel = run_one(cfg, P_ORION_LPL, P_LPL, "LPL")
    res_bad, _ = run_one(cfg, P_ORION_LPL, bad, "LPL")
    sig = lambda r: (list(zip(r.orion["source_row"], r.orion["Result"], r.orion["Verify First"])),
                     list(zip(r.custodian_df["source_row"], r.custodian_df["Result"], r.custodian_df["Verify First"])))
    ck.check(sig(res_good) == sig(res_bad), "Reconciliation results identical for the corrupted LPL export")

    print("ORION AS CSV (as exported from Orion)")
    import csv
    csv_path = TEST_DIR / "TEST_Orion_All_Active_Accounts.csv"
    wb = load_workbook(P_ORION)
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        for row in wb.active.iter_rows(values_only=True):
            w.writerow(["" if v is None else v for v in row])
    got, _ = eng.detect_file(csv_path.name, csv_path.read_bytes(), cfg)
    ck.check(got == ["Orion"], f"Orion CSV identified as {got}")
    r_x, _ = run_one(cfg, P_ORION, P_PERSHING, "Pershing")
    r_c, _ = run_one(cfg, csv_path, P_PERSHING, "Pershing")
    ck.check(sig(r_x) == sig(r_c), "Pershing results identical whether Orion is supplied as .xlsx or .csv")


def test_goldman(ck, cfg):
    print("GOLDMAN SACHS")
    for path, want in ((P_GS, "Goldman Sachs"), (P_ORION_GS, "Orion")):
        got, _ = eng.detect_file(path.name, path.read_bytes(), cfg)
        ck.check(got == [want], f"{path.name} identified as {got}")
    res_g, sel = run_one(cfg, P_ORION_GS, P_GS, "Goldman Sachs")
    ck.check(sel == ["Goldman Sachs"], f"Orion custodian values treated as Goldman Sachs: {sel}")
    check_rows(ck, res_g, ORION_GS_ROWS, GS_ROWS, 2)
    for d, e, a, ok in res_g.controls:
        ck.check(ok, f"{d}: expected {e}, actual {a}")
    ck.check(len(res_g.near_matches) == 0, "No similar-number pairs in the Goldman data")
    ck.check(not {"Primary Owner", "Account Attribute", "Previous Day Available Cash"}
             & set(res_g.cust_file.column_map.values()), "Primary Owner, Account Attribute and Available Cash never read")
    res_p, _ = run_one(cfg, P_ORION, P_PERSHING, "Pershing")
    res_f, _ = run_one(cfg, P_ORION_FID, P_FIDELITY, "Fidelity")
    res_l, _ = run_one(cfg, P_ORION_LPL, P_LPL, "LPL")
    allr = [res_f, res_g, res_l, res_p]
    tabs = eng.build_tabs(allr)
    for r in allr:
        for d, e, a, ok in eng.tab_controls(r, tabs):
            ck.check(ok, f"{d}: expected {e}, actual {a}")
    ck.check(set(tabs["Goldman Sachs - Cancel"]["Account Key"]) == {"123456782", "987654321"},
             "Goldman Sachs - Cancel = 123456782, 987654321")
    ck.check(set(tabs["Goldman Sachs - Not in Orion"]["Account Key"]) == {"123456786", "555000111", "555000444"},
             "Goldman Sachs - Not in Orion = 123456786, 555000111, 555000444")
    xlsx = eng.export_workbook(allr, "Self-test", eng.datetime.now(), tabs, None, [])
    names = load_workbook(io.BytesIO(xlsx), read_only=True).sheetnames
    ck.check(names == ["Summary", "Fidelity - Cancel", "Fidelity - Not in Orion", "Goldman Sachs - Cancel",
                       "Goldman Sachs - Not in Orion", "LPL - Cancel", "LPL - Not in Orion", "Pershing - Cancel",
                       "Pershing - Not in Orion", "Ignored", "Audit Log"], f"Four custodians in one workbook: {names}")
    sm = load_workbook(io.BytesIO(xlsx))["Summary"]
    rows = [[c.value for c in r] for r in sm.iter_rows(min_row=7, max_row=11, max_col=3)]
    ok = [r[0] for r in rows[:4]] == ["Fidelity", "Goldman Sachs", "LPL", "Pershing"] and all(
        r[1] == f"=COUNTA('{r[0]} - Cancel'!A:A)-1" and r[2] == f"=COUNTA('{r[0]} - Not in Orion'!A:A)-1"
        for r in rows[:4]) and rows[4][0] == "Total"
    ck.check(ok, "Summary has one row per custodian (incl. Goldman Sachs), each counting only its own tabs")
    ck.check("TEST OWNER PRIVATE" not in workbook_text(xlsx), "Goldman Primary Owner never appears in the workbook")


def test_pontera(ck, cfg):
    print("PONTERA")
    for path, want in ((P_PNT, "Pontera"), (P_ORION_PNT, "Orion")):
        got, _ = eng.detect_file(path.name, path.read_bytes(), cfg)
        ck.check(got == [want], f"{path.name} identified as {got}")
    res, sel = run_one(cfg, P_ORION_PNT, P_PNT, "Pontera")
    ck.check(sel == ["Pontera"], f"Orion custodian values treated as Pontera: {sel}")
    check_rows(ck, res, ORION_PNT_ROWS, PNT_ROWS, 2)
    for d, e, a, ok in res.controls:
        ck.check(ok, f"{d}: expected {e}, actual {a}")
    ss = {(a, st): n for a, st, n in res.status_summary}
    ck.check(ss.get(("Account ready", "active")) == 4 and ss.get(("Access issue", "not active")) == 3 and
             ss.get(("Disconnected", "not active")) == 1 and ss.get(("ACCOUNT READY", "active")) == 1,
             f"Status summary: only 'Account ready' counts as active; every other status is not active: {ss}")
    reason = res.orion.loc[res.orion["Account Key"] == "PNT100002", "Reason"].iloc[0]
    ck.check("not 'Account ready'" in reason and "confirm" in reason,
             "Cancel reason tells you the Pontera status and to confirm before cancelling")
    ck.check(not {"Client first name", "Client last name", "Risk tolerance", "MFA enabled"}
             & set(res.cust_file.column_map.values()), "Client names, risk tolerance and MFA fields never read")
    tabs = eng.build_tabs([res])
    for d, e, a, ok in eng.tab_controls(res, tabs):
        ck.check(ok, f"{d}: expected {e}, actual {a}")
    ck.check(set(tabs["Pontera - Cancel"]["Account Key"]) == {"PNT100002", "PNT100004"},
             "Pontera - Cancel = PNT100002 (Verify First), PNT100004")
    ck.check(set(tabs["Pontera - Not in Orion"]["Account Key"]) == {"ZPN770001", "ZPN770004"},
             "Pontera - Not in Orion = ZPN770001, ZPN770004")
    ig = tabs[eng.TAB_IGNORED]
    ck.check(set(ig.loc[ig["Account Key"] == "ZPN770002", "Ignored Because"]) == {"Not Active at Custodian"},
             "Not-ready Pontera account missing from Orion goes to Ignored as 'Not Active at Custodian'")
    xlsx = eng.export_workbook([res], "Self-test", eng.datetime.now(), tabs, None, [])
    txt = workbook_text(xlsx)
    ck.check("PRIVATEFIRST" not in txt and "PRIVATELAST" not in txt, "Client names never appear in the workbook")
    ck.check(res.match_info["low"] is False, "Normal file: no low match-rate warning")

    print("LOW MATCH-RATE WARNING (any custodian)")
    big = TEST_DIR / "TEST_Pontera_NoOverlap.xlsx"
    write_pontera(big, [(f"PNT8{i:05d}", "Account ready", 100.0) for i in range(25)])
    res2, _ = run_one(cfg, P_ORION_PNT, big, "Pontera")
    ck.check(res2.match_info["low"] and res2.match_info["found"] == 0 and res2.match_info["considered"] == 25,
             "25 Pontera accounts, none found in Orion -> low match-rate flagged")
    t2 = eng.build_tabs([res2])
    x2 = eng.export_workbook([res2], "Self-test", eng.datetime.now(), t2, None, [])
    sm = load_workbook(io.BytesIO(x2))["Summary"]
    vals = [c.value for r in sm.iter_rows() for c in r if isinstance(c.value, str)]
    ck.check(any(v.startswith("WARNING: only 0 of 25 Pontera accounts") for v in vals) and
             any(v.startswith("Rows shaded orange") for v in vals) and any(v.startswith("INTERNAL USE ONLY") for v in vals),
             "Summary shows the low match-rate warning, and no other Summary line is lost")
    big.unlink()


def test_folder(ck):
    import tempfile as _t
    _d = _t.mkdtemp()
    write_simple_list(Path(_d) / "Ignore Accounts.xlsx", [])
    with patched_config(exclusion_list_path=str(Path(_d) / "Ignore Accounts.xlsx")):
        _test_folder(ck)


def _test_folder(ck):
    import shutil
    import tempfile
    import run_recon
    print("FOLDER PROCESS (temporary test folder)")

    def drop_files(drop):
        return sorted(p.name for p in drop.rglob("*") if p.is_file())

    def results_files(data):
        return sorted((data / run_recon.RESULTS).glob("*.xlsx"))

    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        data = Path(tmp) / "Recon_Data"
        run_recon.run(data, pause=False)
        drop, arch = data / run_recon.DROP, data / run_recon.ARCHIVE

        # A. Run 1: Orion (all custodians) + Pershing + Fidelity, renamed.
        combined = [r for r in ORION_ROWS] + [r for r in ORION_FID_ROWS if r[2] != "Pershing"]
        write_orion(drop / "export_a.xlsx", combined)
        shutil.copy2(P_PERSHING, drop / "export_b.xlsx")
        shutil.copy2(P_FIDELITY, drop / "export_c.xlsx")
        code = run_recon.run(data, pause=False)
        rf = results_files(data)
        ck.check(code == 0, "A. Run 1 finished without errors")
        ck.check(len(rf) == 1 and rf[0].name.startswith("Recon_"), "A. Exactly ONE results workbook for both custodians")
        wbA = open_wb(rf[0])
        ck.check(wbA.sheetnames == ["Summary", "Fidelity - Cancel", "Fidelity - Not in Orion",
                                    "Pershing - Cancel", "Pershing - Not in Orion", "Ignored", "Audit Log"],
                 f"A. Workbook tabs per custodian: {wbA.sheetnames}")
        ck.check(drop_files(drop) == [], "A. Drop folder empty after the run")
        runs = sorted(arch.iterdir()) if arch.exists() else []
        ck.check(len(runs) == 1 and sorted(p.name for p in runs[0].iterdir()) ==
                 sorted(["export_a.xlsx", "export_b.xlsx", "export_c.xlsx", rf[0].name]) if runs else False,
                 "A. Archive\\<run> holds the 3 inputs + the results workbook")
        for c in ("Pershing", "Fidelity"):
            col = {r[0] for r in wbA[f"{c} - Cancel"].iter_rows(min_row=2, values_only=True)}
            ck.check(col == {c}, f"A. '{c} - Cancel' contains only {c} rows")
        track = run_recon.load_tracking(data / run_recon.TRACKING / run_recon.TRACKING_FILE)
        ck.check(len(track) == 14 and set(track["status"]) == {"Open"},
                 f"A. Tracking file records 14 open Not in Orion accounts (got {len(track)})")

        # B. Run 2: TST000007 reactivated in Orion, PSH000501 now added to Orion, PSH000508 now closed.
        orion2 = [(("TST000007", True) + r[2:]) if r[0] == "TST000007" else r for r in combined]
        orion2.append(("PSH000501", True, "Pershing", 12345.67, 0, None, eng.MATCHED, N, "added back"))
        pers2 = [(r[0], "Closed") + r[2:] if r[0] == "PSH000508" else r for r in PERSHING_ROWS]
        write_orion(drop / "orion2.xlsx", orion2)
        write_pershing(drop / "pershing2.xlsx", pers2)
        shutil.copy2(P_FIDELITY, drop / "fidelity2.xlsx")
        code = run_recon.run(data, pause=False)
        rf = results_files(data)
        ck.check(code == 0 and len(rf) == 2, "B. Run 2 finished; a second workbook was written (first kept)")
        trk = run_recon.load_tracking(data / run_recon.TRACKING / run_recon.TRACKING_FILE)
        resolved = set(zip(trk.loc[trk["status"] == "Resolved", "account_key"],
                           trk.loc[trk["status"] == "Resolved", "resolution"]))
        ck.check(resolved == {("TST000007", "Now found in Orion"), ("PSH000501", "Now found in Orion"),
                              ("PSH000508", "Now closed or zero value at custodian")},
                 f"B. Resolved since last run: {sorted(resolved)}")
        wb = open_wb(rf[-1])
        trk_col = [r[2] for r in wb["Pershing - Not in Orion"].iter_rows(min_row=2, values_only=True)]
        ck.check(trk_col and all(str(v).startswith("Outstanding since") for v in trk_col),
                 "B. Accounts still missing show 'Outstanding since <first seen>'")
        audit = {r[0]: r[1] for r in wb["Audit Log"].iter_rows(min_row=2, values_only=True)}
        ck.check("PSH000501" in str(audit.get("[Pershing] Not in Orion resolved since last run")),
                 "B. Audit Log lists the resolved accounts")
        track = run_recon.load_tracking(data / run_recon.TRACKING / run_recon.TRACKING_FILE)
        ck.check((track["status"] == "Resolved").sum() == 3 and (track["status"] == "Open").sum() == 11,
                 "B. Tracking file: 3 resolved, 11 still open")
        ck.check(any(p.name.startswith("tracking_before_") for p in sorted(arch.iterdir())[-1].iterdir()),
                 "B. Previous tracking file backed up in the run's archive folder")

        # C. All-or-nothing: Fidelity fine, Pershing file incomplete (no footer) -> nothing written.
        write_orion(drop / "orion3.xlsx", combined)
        write_pershing(drop / "pershing3.xlsx", PERSHING_ROWS, footer=False)
        shutil.copy2(P_FIDELITY, drop / "fidelity3.xlsx")
        before, files_before = results_files(data), drop_files(drop)
        code = run_recon.run(data, pause=False)
        ck.check(code == 1 and results_files(data) == before and drop_files(drop) == files_before,
                 "C. One custodian fails -> no workbook written, all files left in place")
        for p in drop.rglob("*"):
            if p.is_file():
                p.unlink()

        # D. Two Pershing files -> stop.
        write_orion(drop / "orion4.xlsx", combined)
        shutil.copy2(P_PERSHING, drop / "p1.xlsx")
        shutil.copy2(P_PERSHING, drop / "p2.xlsx")
        before, files_before = results_files(data), drop_files(drop)
        code = run_recon.run(data, pause=False)
        ck.check(code == 1 and results_files(data) == before and drop_files(drop) == files_before,
                 "D. Two files for the same custodian -> stop, nothing moved")
        (drop / "p2.xlsx").unlink()

        # E. Unrecognized file -> stop.
        wbx = Workbook(); wbx.active.append(["Some", "Other", "Report"]); wbx.active.append([1, 2, 3])
        wbx.save(drop / "unknown.xlsx")
        before, files_before = results_files(data), drop_files(drop)
        code = run_recon.run(data, pause=False)
        ck.check(code == 1 and results_files(data) == before and drop_files(drop) == files_before,
                 "E. Unrecognized file -> stop, nothing moved")
        (drop / "unknown.xlsx").unlink()

        # F. Overlapping Orion files -> stop.
        write_orion(drop / "orion_extra.xlsx", ORION_ROWS)
        before, files_before = results_files(data), drop_files(drop)
        code = run_recon.run(data, pause=False)
        ck.check(code == 1 and results_files(data) == before and drop_files(drop) == files_before,
                 "F. Same custodian in two Orion files -> stop, nothing moved")


def write_exclusions(path: Path, rows) -> None:
    """Custodian | Account Number | Reason layout."""
    wb = Workbook()
    ws = wb.active
    ws.append(["Custodian", "Account Number", "Reason"])
    for r in rows:
        ws.append(list(r))
    wb.save(path)


def write_simple_list(path: Path, values, heading="Accounts to ignore") -> None:
    """Simple layout, like 'Ignore Accounts.xlsx': account numbers in column A, heading in row 1."""
    wb = Workbook()
    ws = wb.active
    if heading is not None:
        ws.append([heading])
    for v in values:
        ws.append([v])
    wb.save(path)


class patched_config:
    """Temporarily override settings in custodians.json (used so tests never touch real files)."""
    def __init__(self, **overrides):
        self.overrides = overrides

    def __enter__(self):
        self.orig = eng.load_config
        base = self.orig()
        eng.load_config = lambda path=eng.CONFIG_PATH: {**base, **self.overrides}
        return self

    def __exit__(self, *exc):
        eng.load_config = self.orig


def expect_error(ck, fn, label):
    try:
        fn()
        ck.check(False, label)
    except eng.ReconError:
        ck.check(True, label)


def test_exclusions(ck, cfg):
    print("EXCLUSION LIST - SIMPLE LIST (account numbers in column A)")
    p = TEST_DIR / "excl.xlsx"
    load = lambda: eng.load_exclusions(p.name, p.read_bytes(), cfg)
    write_simple_list(p, ["PSH000501", None, " tst-000004 ", "1234-5678", 12345678, "FID000021"])
    ex = load()
    ck.check(set(ex["Pershing"]) == {"PSH000501", "TST000004", "FID000021"} and
             set(ex["Fidelity"]) == {"PSH000501", "TST000004", "FID000021"} and
             set(ex["LPL"]) == {"12345678"},
             "Heading skipped, blank row skipped; each number applied to every custodian whose format it fits")
    write_simple_list(p, ["1234567890"])
    ex = load()
    ck.check(set(ex) == {"Goldman Sachs", "Pontera"} and set(ex["Goldman Sachs"]) == {"1234567890"},
             "A 10-character number entered as Text applies only to the flexible-format custodians (Goldman, Pontera)")
    write_simple_list(p, ["PSH000501"], heading="Accounts")
    ck.check(set(load()["Pershing"]) == {"PSH000501"} and "ACCOUNTS" not in str(load()),
             "A one-word heading such as 'Accounts' is never taken as an account number")
    write_simple_list(p, [], heading="Accounts to ignore")
    ck.check(load() == {}, "A list with only a heading means nothing is excluded")
    write_simple_list(p, ["PSH000501"], heading=None)
    ck.check(set(load()["Pershing"]) == {"PSH000501"}, "A list without a heading also works")
    for vals, label in ((["PSH000501", "Fidelity accounts"], "text in a later row"),
                        (["PSH000501", 1234567], "Excel number with a lost leading zero (7 digits)"),
                        (["PSH000501", "PSH000501_Growth"], "a UMA sleeve instead of the account"),
                        (["PSH000501", "PSH0005/01"], "an unusual character"),
                        (["PSH000501", 123.45], "a decimal number"),
                        (["PSH000501", 1234567890], "Excel number of a variable-length (Goldman-style) account")):
        write_simple_list(p, vals)
        expect_error(ck, load, f"Simple list rejected (whole run stops): {label}")

    print("EXCLUSION LIST - CUSTODIAN LAYOUT (Custodian | Account Number | Reason)")
    for rows, label in (([("Schwab", "TST000004", "x")], "unknown custodian"),
                        ([("Pershing", "TST000004_Growth", "x")], "sleeve entered"),
                        ([("Pershing", "", "x")], "blank account number"),
                        ([("Pershing", "TST00/004", "x")], "unusual character"),
                        ([("Pershing", "TST0004", "x")], "wrong format for Pershing"),
                        ([("LPL", "TST000004", "x")], "Pershing-format number entered for LPL")):
        write_exclusions(p, rows)
        expect_error(ck, load, f"Invalid entry rejected: {label}")
    write_exclusions(p, [("Pershing", "PSH000501", "Test reason A"), ("pershing", " tst-000004 ", "Test reason B"),
                         ("Pershing", "TST000007", ""), ("Pershing", "TST000099", "Not in any file"),
                         (None, None, None), ("Fidelity", "FID000021", "Exclude UMA parent"),
                         ("LPL", 10000005, "number with full length")])
    ex = load()
    ck.check(set(ex) == {"Pershing", "Fidelity", "LPL"} and
             set(ex["Pershing"]) == {"PSH000501", "TST000004", "TST000007", "TST000099"},
             "Custodian layout parsed; names in any case; dashes/spaces handled; blank row skipped")

    write_exclusions(p, [(None, "PSH000501", "Custodian left blank"), (None, "1234-5678", "")])
    ex2 = load()
    ck.check(set(ex2["Pershing"]) == {"PSH000501"} and set(ex2["Fidelity"]) == {"PSH000501"} and
             set(ex2["LPL"]) == {"12345678"},
             "Custodian left blank: entry applies to every custodian whose format it fits")
    write_exclusions(p, [(None, "Some note", "x")])
    expect_error(ck, load, "Custodian blank and an invalid account number -> rejected")
    write_exclusions(p, [("PSH000501", None, "x")])
    try:
        load()
        ck.check(False, "Account number typed in the Custodian column -> rejected with a hint")
    except eng.ReconError as e:
        ck.check("Account Number column" in str(e), "Account number typed in the Custodian column -> rejected with a hint")
    write_exclusions(p, [("Pershing", "PSH000501", "x")])
    ck.check(set(load()) == {"Pershing"}, "Custodian filled in: entry applies to that custodian only")

    print("EXCLUSION LIST - RESULTS")
    res_p, sel = run_one(cfg, P_ORION, P_PERSHING, "Pershing")
    res = eng.reconcile(res_p.orion_file, res_p.cust_file, "Pershing", sel, cfg, ex["Pershing"])
    o = dict(zip(res.orion["Account Key"], res.orion["Result"]))
    c = dict(zip(res.custodian_df["Account Key"], res.custodian_df["Result"]))
    ck.check(o["TST000004"] == eng.EXCLUDED_LIST and c["PSH000501"] == eng.EXCLUDED_LIST and
             o["TST000007"] == eng.EXCLUDED_LIST and c["TST000007"] == eng.EXCLUDED_LIST,
             "Listed accounts excluded on both the Orion and custodian side")
    ck.check(o["TST000014"] == eng.CANCEL and c["PSH000508"] == eng.MISSING, "Accounts not on the list are unaffected")
    ck.check(res.exclusion_info["orion_rows"] == 2 and res.exclusion_info["custodian_rows"] == 2,
             "Applied to exactly 2 Orion rows and 2 Pershing rows")
    tabs = eng.build_tabs([res])
    for d, e, a, ok in list(res.controls) + eng.tab_controls(res, tabs):
        ck.check(ok, f"With exclusions - {d}: expected {e}, actual {a}")
    everywhere = set()
    for t in tabs.values():
        if len(t) and "Account Key" in t.columns:
            everywhere |= set(t["Account Key"])
    ck.check(not ({"TST000004", "PSH000501", "TST000007"} & everywhere),
             "Excluded accounts appear on NO tab (not Cancel, Not in Orion, Do Not Cancel or Ignored)")
    xlsx = eng.export_workbook([res], "x", eng.datetime.now(), tabs, None, [], show_keep_tabs=True)
    txt = workbook_text(xlsx)
    ck.check(not any(k in txt for k in ("TST000004", "PSH000501", "TST000007")),
             "Excluded account numbers do not appear anywhere in the results workbook")
    res_f0, self_ = run_one(cfg, P_ORION_FID, P_FIDELITY, "Fidelity")
    res_f = eng.reconcile(res_f0.orion_file, res_f0.cust_file, "Fidelity", self_, cfg, ex["Fidelity"])
    fo = res_f.orion
    ck.check(set(fo.loc[fo["Account Key"] == "FID000021", "Result"]) == {eng.EXCLUDED_LIST},
             "Excluding a UMA parent also excludes its sleeve")
    ck.check(set(fo.loc[fo["Account Key"] == "FID000022", "Result"]) == {eng.CANCEL},
             "Custodian-layout Pershing entries do not affect Fidelity accounts")
    p.unlink()


def test_exclusion_folder(ck):
    import shutil
    import tempfile
    import run_recon
    print("EXCLUSION LIST - FOLDER PROCESS (list kept outside Recon_Data, like the OneDrive file)")
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        data = Path(tmp) / "Recon_Data"
        excl = Path(tmp) / "OneDrive - Test" / "Ignore Accounts.xlsx"
        excl.parent.mkdir()
        with patched_config(exclusion_list_path=str(excl)):
            run_recon.run(data, pause=False)
            drop = data / run_recon.DROP
            shutil.copy2(P_ORION, drop / "o.xlsx")
            shutil.copy2(P_PERSHING, drop / "p.xlsx")
            code = run_recon.run(data, pause=False)
            ck.check(code == 1 and not list((data / run_recon.RESULTS).glob("*.xlsx")) and len(list(drop.iterdir())) == 2,
                     "Exclusion list file missing -> run stops, nothing written, files kept")
            write_simple_list(excl, ["PSH000501", "Not an account"])
            code = run_recon.run(data, pause=False)
            ck.check(code == 1 and not list((data / run_recon.RESULTS).glob("*.xlsx")),
                     "Invalid entry in the exclusion list -> run stops, nothing written")
            write_simple_list(excl, ["PSH000501", "TST000004", "ZZZ999999"])
            code = run_recon.run(data, pause=False)
            rf = list((data / run_recon.RESULTS).glob("*.xlsx"))
            ok = code == 0 and len(rf) == 1
            txt = workbook_text(rf[0].read_bytes()) if ok else ""
            ck.check(ok and "PSH000501" not in txt and "TST000004" not in txt,
                     "Run with the list: excluded accounts are absent from the entire results workbook")
            ck.check("ZZZ999999" in txt, "Entry not found in any file is reported on the Audit Log")
            arch = sorted((data / run_recon.ARCHIVE).iterdir())
            ck.check(any(p.name.startswith("exclusions_used_") for p in arch[-1].iterdir()) if arch else False,
                     "A copy of the exclusion list used is kept in the run's archive folder")
            ck.check(excl.exists(), "The exclusion list itself is never moved or changed")


def main() -> int:
    build_files()
    cfg = eng.load_config()
    ck = Checker()
    test_rules(ck, cfg)
    test_new_rules(ck, cfg)
    test_readers(ck, cfg)
    test_goldman(ck, cfg)
    test_pontera(ck, cfg)
    test_exclusions(ck, cfg)
    test_folder(ck)
    test_exclusion_folder(ck)
    print(f"\nTest Excel output written to {TEST_DIR / 'TEST_Recon_Output.xlsx'}")
    print(f"\nRESULT: {'ALL TESTS PASSED' if ck.fails == 0 else f'{ck.fails} TEST(S) FAILED'} "
          f"({ck.passes} checks passed)")
    return 0 if ck.fails == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
