#!/usr/bin/env python3
"""Refresh the Actuals tab of Renaissance_Setter_PnL.xlsx with booking-form data.

Sources (pick one with --source):
  csv         A CSV with the columns in ACTUAL_COLUMNS (default: data/bookings_webform.csv).
              Export it from the Setter Scorecard's 'Bookings Live' / 'Bookings Oct 2026 FINAL'
              tabs, or from the booking portal.
  clickhouse  Runs BOOKINGS_SQL against ClickHouse. The table and column names in BOOKINGS_SQL
              are NOT verified yet (ClickHouse was unreachable when this was written), so check
              them with DESCRIBE before the first run.

Credentials: the ClickHouse password is read from ~/.config/renaissance/clickhouse.env
(CLICKHOUSE_PASSWORD=...) or the CLICKHOUSE_PASSWORD environment variable. It is never printed.

Writes: only to the workbook, and (with --write-ch) to amir_lab.setter_pnl. The shared
schemas main, core, derived, dash and mirror are read-only and the script refuses to write
anywhere except amir_lab.

Usage:
  python refresh_actuals.py --source csv --csv data/bookings_webform.csv
  python refresh_actuals.py --source clickhouse --days 60
  python refresh_actuals.py --source clickhouse --days 60 --write-ch   # also writes amir_lab.setter_pnl
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEFAULT_XLSX = HERE / "Renaissance_Setter_PnL.xlsx"
DEFAULT_CSV = HERE / "data" / "bookings_webform.csv"

CH_HOST = "uryfmx4y96.us-east-1.aws.clickhouse.cloud"
CH_PORT = 8443
CH_USER = "amir"
CH_WRITE_DB = "amir_lab"
READ_ONLY_DBS = {"main", "core", "derived", "dash", "mirror"}

# Booking-form rows. Columns must come back in this order and with these names.
ACTUAL_COLUMNS = ["booking_id", "booked_date", "setter", "partner", "offer",
                  "verdict", "min_qual", "full_qual", "source"]

# UNVERIFIED: adjust table/column names after running DESCRIBE in ClickHouse.
# The partner show-label sheet says the portal table is called im_bookings.
BOOKINGS_SQL = """
SELECT
    toString(booking_id)                       AS booking_id,
    toDate(booked_at)                          AS booked_date,
    lower(trim(setter))                        AS setter,
    partner                                    AS partner,
    offer                                      AS offer,
    coalesce(recording_verdict, '')            AS verdict,
    if(min_qualified, 'Yes', 'No')             AS min_qual,
    if(fully_qualified, 'Yes', 'No')           AS full_qual,
    'clickhouse'                               AS source
FROM mirror.im_bookings
WHERE channel = 'Call'
  AND toDate(booked_at) >= today() - {days:UInt32}
ORDER BY booked_date, booking_id
"""

MAX_ROW = 20000  # Actuals named ranges run to this row; keep in sync with build_model.py
DEFAULT_PAY_TYPE = "LATAM full-time"


# --------------------------------------------------------------------------- helpers
def _password() -> str:
    pw = os.environ.get("CLICKHOUSE_PASSWORD")
    if pw:
        return pw
    env = Path.home() / ".config" / "renaissance" / "clickhouse.env"
    if env.exists():
        for line in env.read_text().splitlines():
            line = line.strip()
            if line.startswith(("CLICKHOUSE_PASSWORD=", "PASSWORD=", "CH_PASSWORD=")):
                return line.split("=", 1)[1].strip().strip('"').strip("'")
    sys.exit("No ClickHouse password: set CLICKHOUSE_PASSWORD or create "
             "~/.config/renaissance/clickhouse.env")


def _client():
    import clickhouse_connect  # pip install clickhouse-connect
    return clickhouse_connect.get_client(host=CH_HOST, port=CH_PORT, username=CH_USER,
                                         password=_password(), secure=True,
                                         database=CH_WRITE_DB)


def _to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return dt.date.fromisoformat(str(v)[:10])


def load_csv(path: Path) -> list[dict]:
    rows = []
    with open(path, newline="") as f:
        for r in csv.DictReader(f):
            if not r.get("booked_date") or not r.get("setter"):
                continue
            r["booked_date"] = _to_date(r["booked_date"])
            r["setter"] = r["setter"].strip().lower()
            r.setdefault("source", "csv")
            rows.append(r)
    return rows


def load_clickhouse(days: int) -> list[dict]:
    res = _client().query(BOOKINGS_SQL, parameters={"days": days})
    rows = [dict(zip(res.column_names, r)) for r in res.result_rows]
    for r in rows:
        r["booked_date"] = _to_date(r["booked_date"])
    return rows


# --------------------------------------------------------------------------- workbook
def write_actuals(wb, rows: list[dict]) -> None:
    """Rewrite Actuals data + its formula columns. Shared by build_model.py."""
    ws = wb["Actuals"]
    if len(rows) + 1 > MAX_ROW:
        sys.exit(f"{len(rows)} rows exceed MAX_ROW={MAX_ROW}; raise it in both scripts")
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row)
    for i, r in enumerate(sorted(rows, key=lambda x: (x["booked_date"], str(x["booking_id"]))), start=2):
        ws.cell(i, 1, str(r["booking_id"]))
        c = ws.cell(i, 2, r["booked_date"]); c.number_format = "yyyy-mm-dd"
        ws.cell(i, 3, r["setter"])
        ws.cell(i, 4, r["partner"])
        ws.cell(i, 5, r["offer"])
        ws.cell(i, 6, r["verdict"])
        ws.cell(i, 7, r["min_qual"])
        ws.cell(i, 8, r["full_qual"])
        ws.cell(i, 9, r.get("source", ""))
        ws.cell(i, 10, f"=IFERROR(VLOOKUP(D{i},PartnerTbl,5,0),in_OtherPrice)")
        ws.cell(i, 11, f'=IF(AND(G{i}="Yes",F{i}<>"FALSE BOOKING"),1,0)')
        ws.cell(i, 12, f'=IF(F{i}="FALSE BOOKING",0,IF(G{i}="Yes",1,in_UnqualBill))')
        ws.cell(i, 13, f"=J{i}*L{i}").number_format = '"$"#,##0'
        ws.cell(i, 14, f"=IF(AND(B{i}>=in_WinStart,B{i}<=in_WinEnd),1,0)")
        ws.cell(i, 15, f'=IF(F{i}="FALSE BOOKING",1,0)')
        ws.cell(i, 16, f"=J{i}*(1-O{i})").number_format = '"$"#,##0'  # revenue if every booking bills


def sync_roster(wb, rows: list[dict]) -> list[str]:
    """Append setters that appear in the data but not on the Roster. Returns new names."""
    ws = wb["Roster"]
    known = {str(ws.cell(r, 1).value).strip().lower()
             for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value}
    new = sorted({r["setter"] for r in rows} - known)
    if new:
        print(f"WARNING: {len(new)} setter(s) missing from Roster, added as "
              f"'{DEFAULT_PAY_TYPE}' (GUESS): {', '.join(new)}. Add rows to "
              "Actual_vs_Model too, or rebuild with build_model.py.")
    for name in new:
        r = ws.max_row + 1
        ws.cell(r, 1, name); ws.cell(r, 2, "?"); ws.cell(r, 3, DEFAULT_PAY_TYPE)
        ws.cell(r, 6, "Y"); ws.cell(r, 7, "Added by refresh script - confirm pay type")
        ws.cell(r, 8, f'=IF(D{r}<>"",D{r},VLOOKUP(C{r},PayTbl,2,0))')
        ws.cell(r, 9, f'=IF(E{r}<>"",E{r},in_Commission)')
    return new


def set_window(wb, rows: list[dict], days: int) -> None:
    end = max(r["booked_date"] for r in rows)
    start = end - dt.timedelta(days=days - 1)
    names = wb.defined_names
    for key, val in (("in_WinStart", start), ("in_WinEnd", end)):
        sheet, ref = names[key].attr_text.split("!")
        wb[sheet.strip("'")][ref.replace("$", "")].value = val
    print(f"Actuals window set to {start} .. {end}")


def recalc(xlsx: Path) -> Path | None:
    """Recalculate with LibreOffice (if installed) so cached values exist. Returns copy path."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    out = Path(tempfile.mkdtemp())
    subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir",
                    str(out), str(xlsx)], check=True, capture_output=True)
    return out / xlsx.name


def write_summary_to_clickhouse(xlsx: Path) -> None:
    import openpyxl
    calc = recalc(xlsx)
    if calc is None:
        sys.exit("--write-ch needs LibreOffice to compute formula values")
    ws = openpyxl.load_workbook(calc, data_only=True)["Actual_vs_Model"]
    header = [c.value for c in ws[3]]
    out = []
    for r in ws.iter_rows(min_row=4, values_only=True):
        if not r[0]:
            continue
        d = dict(zip(header, r))
        out.append([str(d["Setter"]), str(d["Pay type"]), int(d["Bookings"] or 0),
                    int(d["Qualified (billable)"] or 0), int(d["False bookings"] or 0),
                    int(d["Calling days"] or 0), float(d["Revenue (actual)"] or 0),
                    float(d["Total cost"] or 0), float(d["Contribution margin (actual)"] or 0),
                    float(d["Contribution margin (modeled)"] or 0), dt.datetime.utcnow()])
    table = f"{CH_WRITE_DB}.setter_pnl"
    assert table.split(".")[0] == CH_WRITE_DB and CH_WRITE_DB not in READ_ONLY_DBS
    cl = _client()
    cl.command(f"""CREATE TABLE IF NOT EXISTS {table} (
        setter String, pay_type String, bookings UInt32, qualified UInt32,
        false_bookings UInt32, active_days UInt32, revenue Float64, total_cost Float64,
        cm_actual Float64, cm_modeled Float64, refreshed_at DateTime
    ) ENGINE = MergeTree ORDER BY (refreshed_at, setter)""")
    cl.insert(table, out, column_names=["setter", "pay_type", "bookings", "qualified",
              "false_bookings", "active_days", "revenue", "total_cost", "cm_actual",
              "cm_modeled", "refreshed_at"])
    print(f"Wrote {len(out)} rows to {table}")


def main() -> None:
    import openpyxl
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--source", choices=["csv", "clickhouse"], default="csv")
    ap.add_argument("--csv", type=Path, default=DEFAULT_CSV)
    ap.add_argument("--xlsx", type=Path, default=DEFAULT_XLSX)
    ap.add_argument("--days", type=int, default=60, help="actuals window length (days)")
    ap.add_argument("--write-ch", action="store_true", help="also write amir_lab.setter_pnl")
    a = ap.parse_args()

    rows = load_csv(a.csv) if a.source == "csv" else load_clickhouse(a.days)
    if not rows:
        sys.exit("No rows loaded")
    wb = openpyxl.load_workbook(a.xlsx)
    write_actuals(wb, rows)
    sync_roster(wb, rows)
    set_window(wb, rows, a.days)
    wb.save(a.xlsx)
    calc = recalc(a.xlsx)
    if calc and calc.exists():
        shutil.copy(calc, a.xlsx)
    print(f"Refreshed {len(rows)} bookings into {a.xlsx}")
    if a.write_ch:
        write_summary_to_clickhouse(a.xlsx)


if __name__ == "__main__":
    main()
