#!/usr/bin/env python3
"""Build Renaissance_Setter_PnL.xlsx: per-setter and team P&L for the setter floor.

Every calculation cell is a live formula that reads from the Inputs tab (named cells) or
the Actuals tab. Re-run this script only to change the structure; to refresh data use
refresh_actuals.py.
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName

import shutil

from refresh_actuals import DEFAULT_CSV, MAX_ROW, load_csv, recalc, write_actuals

OUT = Path(__file__).resolve().parent / "Renaissance_Setter_PnL.xlsx"

# ----------------------------------------------------------------------------- styles
NAVY = "1F3A5F"
F_TITLE = Font(bold=True, size=14, color=NAVY)
F_H = Font(bold=True, color="FFFFFF")
F_B = Font(bold=True)
F_IN = Font(color="0000FF")            # inputs: blue
F_NOTE = Font(italic=True, color="666666", size=9)
FILL_H = PatternFill("solid", fgColor=NAVY)
FILL_SEC = PatternFill("solid", fgColor="DCE6F1")
FILL_GUESS = PatternFill("solid", fgColor="FFF2CC")   # guessed input: yellow
FILL_ACT = PatternFill("solid", fgColor="E2EFDA")     # derived from actual data: green
FILL_KPI = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
USD = '"$"#,##0;[Red]-"$"#,##0'
USD2 = '"$"#,##0.00;[Red]-"$"#,##0.00'
PCT = "0.0%"
NUM1 = "0.0"
NUM0 = "#,##0"
STATUS_FILL = {"GUESS": FILL_GUESS, "ACTUAL": FILL_ACT, "DERIVED": FILL_ACT}

wb = Workbook()


def name(n: str, sheet: str, ref: str) -> None:
    wb.defined_names[n] = DefinedName(n, attr_text=f"'{sheet}'!{ref}")


def header(ws, row: int, labels: list[str], col: int = 1) -> None:
    for i, t in enumerate(labels):
        c = ws.cell(row, col + i, t)
        c.font, c.fill, c.border = F_H, FILL_H, BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")


def section(ws, row: int, text: str, width: int = 5) -> None:
    for i in range(1, width + 1):
        ws.cell(row, i).fill = FILL_SEC
    ws.cell(row, 1, text).font = F_B


def widths(ws, w: dict[str, float]) -> None:
    for k, v in w.items():
        ws.column_dimensions[k].width = v


# ============================================================================ INPUTS
wsI = wb.active
wsI.title = "Inputs"
wsI["A1"] = "INPUTS - every assumption in the model lives here"
wsI["A1"].font = F_TITLE
wsI["A2"] = ("Blue = editable input. Yellow = GUESSED (no source found; confirm). "
             "Green = computed from booking-form actuals. Change a value and every tab updates.")
wsI["A2"].font = F_NOTE
header(wsI, 4, ["Assumption", "Value", "Unit", "Status", "Source / note"])

r = 5
SCALARS: list[tuple] = [
    ("SEC", "Productivity & revenue"),
    ("in_Target", "Target qualified meetings per setter per day (OTE)", 5, "mtgs/day", "CONFIRMED", "CEO brief: 5 qualified meetings/day", NUM1),
    ("in_Days", "Working (selling) days per month", 21, "days", "GUESS", "Mon-Fri; ~21.7 avg. Holidays not modelled", NUM0),
    ("in_Ramp1", "Ramp: month 1 productivity (% of target)", 0.5, "%", "GUESS", "Nov 13 class booked ~3 qualified/day in week 1 (leon 9 in 3 days)", PCT),
    ("in_Ramp2", "Ramp: month 2 productivity (% of target)", 0.8, "%", "GUESS", "", PCT),
    ("in_Ramp3", "Ramp: month 3+ productivity (% of target)", 1.0, "%", "CONFIRMED", "Brief", PCT),
    ("in_QualStart", "Qualified-rate basis: bookings dated on/after", dt.date(2026, 11, 1), "date", "INPUT", "Nov 1 = start of the CEO 3-question booking mandate", "yyyy-mm-dd"),
    ("in_QualRate", "Qualified rate (booked -> qualified, billable)", "=IFERROR(COUNTIFS(act_q,1,act_date,\">=\"&in_QualStart)/COUNTIFS(act_date,\">=\"&in_QualStart),0)", "%", "ACTUAL", "Caller QA 'Qualified' (Nov 1-15), else Scorecard Min qualified; false bookings excluded", PCT),
    ("in_UnqualBill", "Share of unqualified bookings that still bill", 0.0, "%", "CONFIRMED", "Brief: false/unqualified = unbillable. Set >0 to test partners paying per booking regardless", PCT),
    ("in_WinStart", "Actuals window start (last 60 days)", dt.date(2026, 9, 19), "date", "INPUT", "refresh_actuals.py resets this", "yyyy-mm-dd"),
    ("in_WinEnd", "Actuals window end (data through)", dt.date(2026, 11, 17), "date", "INPUT", "Scorecard 'Bookings Live' data through 2026-11-17; booking data starts 2026-10-01", "yyyy-mm-dd"),
    ("in_KPIRevPerMtg", "Cross-check only: company revenue per booked meeting", 111.69, "$/mtg", "ACTUAL", "W45 KPI pack (#leaders, 2026-11-09), all channels. Not used in formulas", USD2),
    ("SEC", "Cost per setter"),
    ("in_Fee", "Hiring fee (one-time per hire)", 1500, "$", "CONFIRMED", "Brief", USD),
    ("in_Tenure", "Expected tenure (months) for amortizing the hiring fee", 6, "months", "GUESS", "No tenure data found", NUM0),
    ("in_GoFunnel", "GoFunnel.ai licence", 150, "$/mo", "CONFIRMED", "Brief", USD),
    ("in_Trainual", "Trainual seat", 8, "$/mo", "GUESS", "", USD),
    ("in_Phone", "Phone numbers / caller ID", 10, "$/mo", "GUESS", "May already be inside GoFunnel", USD),
    ("in_Mgmt", "Management overhead allocation (Julian, QA, payroll)", 350, "$/mo", "GUESS", "Allocation per seat; no cost data found", USD),
    ("in_SeatCost", "Total per-seat cost (excl. base, commission, hiring)", "=in_GoFunnel+in_Trainual+in_Phone+in_Mgmt", "$/mo", "CALC", "", USD),
    ("in_Commission", "Commission per qualified meeting", 5, "$/mtg", "CONFIRMED", "Julian, #ops-amir-julian 2026-11-15: from Nov, $5 per qualified meeting (held no longer matters). Hekmate has an $8 tier (not modelled)", USD2),
    ("SEC", "Live transfers (to GreenBridge)"),
    ("in_XferPrice", "Price per live transfer", 170, "$/transfer", "CONFIRMED", "Amir 2026-11-19: $170 per live transfer", USD),
    ("in_XferComm", "Setter commission per live transfer", 5, "$/transfer", "GUESS", "Assumed same as a qualified meeting", USD2),
    ("in_XferShare", "Live transfers as a share of all outcomes (transfers + meetings)", 0.20, "%", "CONFIRMED", "Amir 2026-11-19: 20% transfers / 80% meetings. Transfers come on top of meetings, they don't replace them. Nov 18-19 actual: 19%", PCT),
    ("in_XferPerMtg", "Live transfers per booked meeting", "=in_XferShare/(1-in_XferShare)", "per mtg", "CALC", "20/80 mix = 1 transfer for every 4 meetings", "0.00"),
    ("in_RevBookedQ", "Revenue per booked meeting incl. its transfers, qualified meetings billed", "=in_QualRate*in_RevQEff+in_XferPerMtg*in_XferPrice", "$/mtg", "CALC", "A transfer always bills; a meeting bills only if qualified", USD2),
    ("in_RevBookedAll", "Revenue per booked meeting incl. its transfers, every booking billed", "=in_RevQ+in_XferPerMtg*in_XferPrice", "$/mtg", "CALC", "", USD2),
    ("in_CommBooked", "Commission per booked meeting incl. its transfers", "=in_QualRate*in_Commission+in_XferPerMtg*in_XferComm", "$/mtg", "CALC", "", USD2),
    ("in_PosWeek", "Funding positive replies with a phone entering the dialer per week", 3700, "per week", "ACTUAL", "ClickHouse: 3,699 leads added to the positive-reply list Nov 10-14", NUM0),
    ("SEC", "Positive-reply supply ladder (P0 to P4): see ladder table below"),
    ("in_Handled", "Live conversations handled per setter per day", 38, "convs/day", "DERIVED", "ClickHouse: 5,127 agent-handled funding calls Nov 10-18 / ~133 setter-days (19 setters x 7 weekdays)", NUM0),
    ("in_MeasDays", "Selling days in the ladder measurement window", 7, "days", "ACTUAL", "Nov 10-18 weekdays", NUM0),
    ("in_P0CapQ", "P0-P4 supply: qualified meetings per month available", "=SUM(Lad_Avail)*in_QualRate", "mtgs/mo", "CALC", "", NUM0),
    ("SEC", "Purchased cold data (new P4/P6 numbers) - only used once P0-P4 is fully worked"),
    ("in_ColdPerNum", "Cold data cost per phone number", 0.07, "$/number", "CONFIRMED", "Brief", USD2),
    ("in_ColdNums", "Phone numbers bought per booked cold meeting", 450, "numbers", "GUESS", "Nov 17: 31,585 dials -> 70 meetings (~450 dials/meeting) as proxy", NUM0),
    ("in_ColdBookDay", "Booked meetings per setter per day on purchased cold data", 2, "mtgs/day", "GUESS", "Scenario: 'if we book 2 a day cold'. No cold dialling yet under GoFunnel", NUM1),
]
for item in SCALARS:
    if item[0] == "SEC":
        section(wsI, r, item[1]); r += 1; continue
    n, label, val, unit, status, src, fmt = item
    wsI.cell(r, 1, label)
    c = wsI.cell(r, 2, val); c.number_format = fmt; c.border = BOX
    if status not in ("CALC", "ACTUAL") or not str(val).startswith("="):
        c.font = F_IN
    if status in STATUS_FILL:
        c.fill = STATUS_FILL[status]
    wsI.cell(r, 3, unit); wsI.cell(r, 4, status); wsI.cell(r, 5, src).font = F_NOTE
    name(n, "Inputs", f"$B${r}")
    r += 1

# ---- pay types
r += 1
section(wsI, r, "Base pay by setter type (monthly, USD)", 6); r += 1
header(wsI, r, ["Pay type", "Base $/mo", "Status", "Source / note", "Active setters", "Team mix %"]); r += 1
PAY = [
    ("PH full-time", 1213, "DERIVED", "Upwork agency $7/hr x 173 h (comp map, Oct 13). Katherine, Maria"),
    ("PH part-time", 606, "GUESS", "Half of PH full-time; no PT setters today"),
    ("LATAM full-time", 2500, "CONFIRMED", "Amir 2026-10-14: $2,500 + $5/meeting; LATAM JD"),
    ("LATAM junior", 1500, "GUESS", "No junior rate found"),
    ("Egypt (NAC)", 693, "DERIVED", "No Accent Callers $4/hr x 173 h (comp map). Elias, Hekmate"),
    ("Pipeline manager", 4000, "CONFIRMED", "Pablo base $4,000 (Amir, 2026-11-16)"),
]
pay_first = r
for p, base, status, src in PAY:
    wsI.cell(r, 1, p)
    c = wsI.cell(r, 2, base); c.font = F_IN; c.number_format = USD; c.border = BOX
    if status in STATUS_FILL:
        c.fill = STATUS_FILL[status]
    wsI.cell(r, 3, status); wsI.cell(r, 4, src).font = F_NOTE
    wsI.cell(r, 5, f'=COUNTIFS(Roster!$C$2:$C$500,A{r},Roster!$F$2:$F$500,"Y")')
    wsI.cell(r, 6, f"=IFERROR(E{r}/SUM($E${pay_first}:$E${pay_first+len(PAY)-1}),0)").number_format = PCT
    r += 1
pay_last = r - 1
name("PayTbl", "Inputs", f"$A${pay_first}:$B${pay_last}")
name("PayNames", "Inputs", f"$A${pay_first}:$A${pay_last}")
name("PayBase", "Inputs", f"$B${pay_first}:$B${pay_last}")
name("PayMix", "Inputs", f"$F${pay_first}:$F${pay_last}")
wsI.cell(r, 1, "Blended base (current roster mix)").font = F_B
c = wsI.cell(r, 2, "=SUMPRODUCT(PayBase,PayMix)"); c.number_format = USD; c.font = F_B
name("in_BlendBase", "Inputs", f"$B${r}")
wsI.cell(r, 4, "Team P&L uses this mix. Change a setter's pay type on Roster to change it.").font = F_NOTE
r += 2

# ---- partners
section(wsI, r, "Partner pricing (revenue per billable meeting)", 9); r += 1
header(wsI, r, ["Partner", "Billing basis (Booked/Held)", "Price $", "Show rate", "Eff. $ per billable booking",
                "Qualified bookings in window", "Mix share", "Status", "Source / note"]); r += 1
PARTNERS = [
    ("GoQualifi", "Booked", 155, 0.395, "CONFIRMED", "Amir 2026-11-18: we charge $155 per meeting"),
    ("Llama", "Booked", 155, 0.649, "CONFIRMED", "Amir 2026-11-18: we charge $155 per meeting"),
    ("Big Think Capital", "Booked", 155, 0.668, "CONFIRMED", "Amir 2026-11-18: we charge $155 per meeting"),
    ("GreenBridge Capital", "Booked", 155, 0.553, "CONFIRMED", "Amir 2026-11-18: we charge $155 per meeting"),
    ("Essential Capital", "Booked", 155, 0.685, "CONFIRMED", "Amir 2026-11-18: we charge $155 per meeting"),
    ("Direct Funding Now", "Booked", 155, 0.867, "CONFIRMED", "Amir 2026-11-18: we charge $155 per meeting"),
    ("Sunlight Capital", "Booked", 155, 0.824, "CONFIRMED", "Amir 2026-11-18: we charge $155 per meeting"),
    ("Fresh Funding", "Booked", 155, 0.778, "CONFIRMED", "Amir 2026-11-18: we charge $155 per meeting"),
    ("Mode", "Booked", 155, 0.978, "CONFIRMED", "Amir 2026-11-18: we charge $155 per meeting"),
    ("License Your Data", "Booked", 155, 0.60, "CONFIRMED", "Amir 2026-11-18: we charge $155 per meeting"),
    ("Collins", "Held", 260, 0.619, "CONFIRMED", "Oscar 2026-10-27: 'we only bill $260 per held meeting'. Show 61.9%"),
    ("SBG Lending", "Booked", 155, 0.60, "CONFIRMED", "Amir 2026-11-18: we charge $155 per meeting"),
]
part_first = r
for p, basis, price, show, status, src in PARTNERS:
    wsI.cell(r, 1, p)
    for col, v, fmt in ((2, basis, None), (3, price, USD), (4, show, PCT)):
        c = wsI.cell(r, col, v); c.font = F_IN; c.border = BOX
        if fmt: c.number_format = fmt
        if status == "GUESS" and col in (3,): c.fill = FILL_GUESS
    wsI.cell(r, 5, f'=C{r}*IF(B{r}="Held",D{r},1)').number_format = USD
    wsI.cell(r, 6, f"=COUNTIFS(act_partner,A{r},act_q,1,act_inwin,1)")
    r += 1
# catch-all row
wsI.cell(r, 1, "Other (unlisted partners)")
wsI.cell(r, 2, "Booked").font = F_IN
c = wsI.cell(r, 3, 155); c.font = F_IN; c.number_format = USD
c = wsI.cell(r, 4, 0.6); c.font = F_IN; c.number_format = PCT
wsI.cell(r, 5, f'=C{r}*IF(B{r}="Held",D{r},1)').number_format = USD
wsI.cell(r, 6, f"=COUNTIFS(act_q,1,act_inwin,1)-SUM(F{part_first}:F{r-1})")
name("in_OtherPrice", "Inputs", f"$E${r}")
part_last = r
for rr in range(part_first, part_last + 1):
    wsI.cell(rr, 7, f"=IFERROR(F{rr}/SUM($F${part_first}:$F${part_last}),0)").number_format = PCT
    st = next((x[4] for x in PARTNERS if x[0] == wsI.cell(rr, 1).value), "GUESS")
    src = next((x[5] for x in PARTNERS if x[0] == wsI.cell(rr, 1).value), "Used for any partner not listed above; $155 standard rate")
    wsI.cell(rr, 8, st); wsI.cell(rr, 9, src).font = F_NOTE
name("PartnerTbl", "Inputs", f"$A${part_first}:$E${part_last}")
r += 1
wsI.cell(r, 1, "Blended revenue per qualified meeting (window partner mix)").font = F_B
c = wsI.cell(r, 2, f"=SUMPRODUCT(E{part_first}:E{part_last},G{part_first}:G{part_last})")
c.number_format = USD2; c.font = F_B; c.fill = FILL_ACT
name("in_RevQ", "Inputs", f"$B${r}")
r += 1
wsI.cell(r, 1, "Effective revenue per qualified meeting (incl. billable unqualified share)").font = F_B
c = wsI.cell(r, 2, "=in_RevQ*(1+IFERROR(1/in_QualRate-1,0)*in_UnqualBill)")
c.number_format = USD2; c.font = F_B
name("in_RevQEff", "Inputs", f"$B${r}")
wsI.cell(r, 5, "Revenue used throughout the forward model").font = F_NOTE
r += 2
section(wsI, r, "Supply ladder: how positive replies are worked over 30+ days (funding calls, measured Nov 10-18 in ClickHouse)", 9); r += 1
header(wsI, r, ["Tier", "Bookings in window", "Handled conversations in window", "Bookings per conversation",
                "Booked meetings available / month", "Booked per setter-day on this tier", "Setter-months to work it",
                "Setter-months before this tier", "What it is"]); r += 1
LADDER = [
    ("P0  (first hour)", 196, 1700, "4 calls in 30 min, double-dialled (speed to lead)"),
    ("P1  (1-24 hours)", 84, 1058, "4 calls over 6 hours"),
    ("P2  (1-3 days)", 24, 328, "9 calls over 2 days"),
    ("P3  (3+ days)", 33, 475, "Parked; occasional re-dials"),
    ("P4  recent carry-over", 31, 645, "Unreached recent positives re-dialled (12.7 calls/person)"),
    ("P4  old carry-over", 24, 921, "Aged positives from before GoFunnel; ~13,800 never dialled yet"),
]
lad_first = r
for t_, bk_, hd_, what in LADDER:
    wsI.cell(r, 1, t_)
    for col, v in ((2, bk_), (3, hd_)):
        c = wsI.cell(r, col, v); c.font = F_IN; c.fill = FILL_ACT; c.border = BOX
    wsI.cell(r, 4, f"=IFERROR(B{r}/C{r},0)").number_format = PCT
    wsI.cell(r, 5, f"=B{r}*in_Days/in_MeasDays").number_format = NUM0
    wsI.cell(r, 6, f"=D{r}*in_Handled").number_format = "0.00"
    wsI.cell(r, 7, f"=IFERROR(E{r}/(F{r}*in_Days),0)").number_format = "0.0"
    wsI.cell(r, 8, "=0" if r == lad_first else f"=H{r-1}+G{r-1}").number_format = "0.0"
    wsI.cell(r, 9, what).font = F_NOTE
    r += 1
lad_last = r - 1
for n_, col in (("Lad_Avail", "E"), ("Lad_Prod", "F"), ("Lad_Need", "G"), ("Lad_Cum", "H")):
    name(n_, "Inputs", f"${col}${lad_first}:${col}${lad_last}")
wsI.cell(r, 1, "Total: setter-months the P0-P4 ladder can absorb each month").font = F_B
c = wsI.cell(r, 7, f"=SUM(G{lad_first}:G{lad_last})"); c.number_format = "0.0"; c.font = F_B
c = wsI.cell(r, 5, f"=SUM(E{lad_first}:E{lad_last})"); c.number_format = NUM0; c.font = F_B
r += 1
wsI.cell(r, 1, ("Source: GoFunnel call dispositions ('gfbook' = booked) joined to positive-reply list membership by lead; tier = hours "
                "between the reply entering the list and the call. Setters work the best tier first; once a tier is used up, "
                "extra setters move down to the next one. The 30+ day tail of P3 is not fully observed yet.")).font = F_NOTE
r += 2
wsI.cell(r, 1, "Rule: 3 confirmed false bookings = termination. False bookings never bill (Actuals col L) "
               "and are flagged on Actual_vs_Model.").font = F_NOTE
widths(wsI, {"A": 58, "B": 16, "C": 14, "D": 12, "E": 22, "F": 14, "G": 10, "H": 12, "I": 70})
wsI.freeze_panes = "A5"

# ============================================================================ ACTUALS
wsA = wb.create_sheet("Actuals")
header(wsA, 1, ["Booking ID", "Booked date", "Setter", "Partner", "Offer", "Verdict (recording check)",
                "Min qualified", "Fully qualified", "Source", "Eff. $ per billable booking",
                "Qualified (1/0)", "Billable units", "Revenue $", "In window (1/0)", "False booking (1/0)", "Revenue if all billed $", "First booking of setter-day (1/0)"])
rows = load_csv(DEFAULT_CSV)
write_actuals(wb, rows)
for n, col in (("act_date", "B"), ("act_setter", "C"), ("act_partner", "D"), ("act_q", "K"),
               ("act_rev", "M"), ("act_inwin", "N"), ("act_false", "O"), ("act_rev_all", "P"), ("act_dayfirst", "Q")):
    name(n, "Actuals", f"${col}$2:${col}${MAX_ROW}")
widths(wsA, {"A": 11, "B": 11, "C": 12, "D": 20, "E": 11, "F": 24, "G": 9, "H": 9, "I": 10,
             "J": 11, "K": 9, "L": 9, "M": 10, "N": 9, "O": 9})
wsA.freeze_panes = "A2"

# ============================================================================ TRANSFERS
wsTr = wb.create_sheet("Transfers")
header(wsTr, 1, ["Date", "GoFunnel agent", "Setter", "Destination", "Campaign"])
import csv as _csv
_tr = list(_csv.DictReader(open(Path(__file__).resolve().parent / "data" / "live_transfers.csv")))
for i, t_ in enumerate(_tr, start=2):
    c = wsTr.cell(i, 1, dt.date.fromisoformat(t_["date"])); c.number_format = "yyyy-mm-dd"
    for col, k in ((2, "agent"), (3, "setter"), (4, "destination"), (5, "campaign")):
        wsTr.cell(i, col, t_[k])
name("tr_date", "Transfers", "$A$2:$A$2000"); name("tr_setter", "Transfers", "$C$2:$C$2000")
wsTr["G1"] = "Successful setter live transfers from GoFunnel transfers(), deduplicated by lead; excludes test and agent-to-agent transfers."
wsTr["G1"].font = F_NOTE
widths(wsTr, {"A": 12, "B": 22, "C": 12, "D": 20, "E": 14})

# ============================================================================ ROSTER
wsR = wb.create_sheet("Roster")
header(wsR, 1, ["Setter", "Offer", "Pay type", "Base override $/mo", "Commission override $/mtg",
                "Active (Y/N)", "Notes", "Effective base", "Effective commission",
                "Paid from (start date)", "Calling from"])
ROSTER = [
    ("katherine", "Funding", "PH full-time", None, None, "Y", "Upwork agency $7/hr"),
    ("maria", "Funding", "PH full-time", None, None, "Y", "Upwork agency $7/hr"),
    ("erwell", "Funding", "PH full-time", 1386, None, "Y", "Upwork $8/hr x 173 h"),
    ("hekmate", "Funding", "Egypt (NAC)", None, None, "Y", "$8/mtg tier above threshold not modelled"),
    ("elias", "Funding", "Egypt (NAC)", None, None, "Y", ""),
    ("jose", "Funding", "LATAM full-time", None, None, "Y", "Contract Oct 6"),
    ("brian", "Funding", "LATAM full-time", None, None, "Y", "Contract Oct 6"),
    ("steve", "Funding", "LATAM full-time", None, None, "Y", ""),
    ("jacobo", "Funding", "LATAM full-time", None, None, "Y", "Moved Pre-IPO -> Funding Nov 4"),
    ("alejandro", "Data Broker", "LATAM full-time", 2800, 10, "N", "Data Broker campaign: separate P&L, excluded from the funding setter floor (Amir 11/19)"),
    ("oscarduran", "Funding", "LATAM full-time", None, None, "Y", "In Caller QA only (not in Scorecard data); pay type GUESSED"),
    ("pablo", "Data Broker", "Pipeline manager", None, 0, "N", "Runs the Data Broker team (manager, not a setter): excluded from setter P&L and headcount"),
    ("marwan", "Pre-IPO", "Egypt (NAC)", None, None, "N", "Terminated"),
] + [(n, "Funding", "LATAM full-time", None, None, "Y", "Nov 13 class; pay type GUESSED")
     for n in ("leon", "edgardo", "eddie", "angela", "edwin", "mateo")] \
  + [(n, "Funding", "LATAM full-time", None, None, "Y", "Training from Nov 9, calling from Nov 14; pay type GUESSED")
     for n in ("stephy", "manuel", "christopher")]
STARTS = {n: (dt.date(2026, 11, 9), dt.date(2026, 11, 14)) for n in ("stephy", "manuel", "christopher")}
seen = {x[0] for x in ROSTER}
for s in sorted({x["setter"] for x in rows} - seen):
    ROSTER.append((s, "?", "LATAM full-time", None, None, "Y", "Pay type GUESSED"))
for i, (s, off, pt, bo, co, act, note) in enumerate(ROSTER, start=2):
    wsR.cell(i, 1, s); wsR.cell(i, 2, off)
    for col, v in ((3, pt), (4, bo), (5, co), (6, act)):
        c = wsR.cell(i, col, v); c.font = F_IN; c.border = BOX
    if "GUESS" in note:
        wsR.cell(i, 3).fill = FILL_GUESS
    wsR.cell(i, 7, note).font = F_NOTE
    wsR.cell(i, 8, f'=IF(D{i}<>"",D{i},VLOOKUP(C{i},PayTbl,2,0))').number_format = USD
    wsR.cell(i, 9, f'=IF(E{i}<>"",E{i},in_Commission)').number_format = USD2
    for col, v in zip((10, 11), STARTS.get(s, (None, None))):
        c = wsR.cell(i, col, v); c.font = F_IN; c.border = BOX; c.number_format = "yyyy-mm-dd"
N_ROSTER = len(ROSTER)
name("in_ActiveN", "Roster", "$N$2")
wsR["M2"] = "Active setters"; wsR["N2"] = '=COUNTIF($F$2:$F$500,"Y")'; wsR["M2"].font = F_B
wsR["M3"] = "Blank start dates: cost runs from the first booking in the window."; wsR["M3"].font = F_NOTE
widths(wsR, {"A": 13, "B": 12, "C": 18, "D": 12, "E": 12, "F": 9, "G": 60, "H": 12, "I": 12, "J": 13, "K": 13, "M": 16})
wsR.freeze_panes = "A2"

# ============================================================================ SETTER P&L
wsS = wb.create_sheet("Setter_PnL")
wsS["A1"] = "Monthly P&L per setter at target (steady state, month 3+)"; wsS["A1"].font = F_TITLE
wsS["A2"] = "One column per pay type. P0 supply is not capped here; see Team_PnL and Supply."
wsS["A2"].font = F_NOTE
header(wsS, 4, ["Line ($/month)"] + [p[0] for p in PAY])
LINES = [
    ("Qualified meetings / day", "=in_Target*in_Ramp3", NUM1),
    ("Working days", "=in_Days", NUM0),
    ("Qualified meetings / month", "={c}5*{c}6", NUM0),
    ("Booked meetings / month (at qualified rate)", "=IFERROR({c}7/in_QualRate,0)", NUM0),
    ("Revenue per qualified meeting", "=in_RevQEff", USD2),
    ("REVENUE (meetings + live transfers on top)", "={c}8*in_RevBookedQ", USD),
    ("Base pay", "=INDEX(PayBase,MATCH({c}$4,PayNames,0))", USD),
    ("Commission (meetings + transfers)", "={c}8*in_CommBooked", USD),
    ("GoFunnel.ai", "=in_GoFunnel", USD),
    ("Trainual", "=in_Trainual", USD),
    ("Phone numbers", "=in_Phone", USD),
    ("Management overhead", "=in_Mgmt", USD),
    ("Hiring fee (amortized over tenure)", "=in_Fee/in_Tenure", USD),
    ("TOTAL COST", "=SUM({c}11:{c}17)", USD),
    ("CONTRIBUTION MARGIN", "={c}10-{c}18", USD),
    ("Contribution margin %", "=IFERROR({c}19/{c}10,0)", PCT),
    ("Cost per qualified meeting", "=IFERROR({c}18/{c}7,0)", USD2),
    ("", None, None),
    ("IF EVERY BOOKING BILLS (commission still on qualified only)", None, None),
    ("Revenue: booked meetings + live transfers, all billed", "={c}8*in_RevBookedAll", USD),
    ("Contribution margin (all billed)", "={c}24-{c}18", USD),
    ("Contribution margin % (all billed)", "=IFERROR({c}25/{c}24,0)", PCT),
    ("", None, None),
    ("Hiring fee shown two ways - month of hire (ramp month 1)", None, None),
    ("(a) Fee EXPENSED in month 1: contribution margin",
     "={c}8*in_Ramp1*(in_RevBookedQ-in_CommBooked)-({c}11+in_SeatCost+in_Fee)", USD),
    ("(b) Fee AMORTIZED: contribution margin month 1",
     "={c}8*in_Ramp1*(in_RevBookedQ-in_CommBooked)-({c}11+in_SeatCost+in_Fee/in_Tenure)", USD),
]
for i, (lab, f, fmt) in enumerate(LINES, start=5):
    wsS.cell(i, 1, lab)
    if lab.isupper() or lab.startswith("Hiring fee shown"):
        wsS.cell(i, 1).font = F_B
    if f is None:
        continue
    for j in range(len(PAY)):
        col = L(2 + j)
        c = wsS.cell(i, 2 + j, f.format(c=col)); c.number_format = fmt
        if lab.isupper(): c.font = F_B
for j in range(len(PAY)):
    wsS.cell(19, 2 + j).fill = FILL_KPI
widths(wsS, {"A": 52, **{L(2 + j): 17 for j in range(len(PAY))}})
wsS.conditional_formatting.add(f"B19:{L(1+len(PAY))}19", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="F8CBAD")))

# ============================================================================ BREAK-EVEN
wsB = wb.create_sheet("Breakeven")
wsB["A1"] = "Break-even meetings per setter per day"; wsB["A1"].font = F_TITLE
wsB["A2"] = "Qualified meetings/day at which contribution margin = 0 (hiring fee amortized)."; wsB["A2"].font = F_NOTE
header(wsB, 4, ["Line"] + [p[0] for p in PAY] + ["Blended (roster mix)"])
BE = [
    ("Fixed cost / month (base + seat + hiring amortized)", "=INDEX(PayBase,MATCH({c}$4,PayNames,0))+in_SeatCost+in_Fee/in_Tenure", USD),
    ("Margin per booked meeting incl. its transfers (after commission)", "=in_RevBookedQ-in_CommBooked", USD2),
    ("Break-even BOOKED meetings / month (plus transfers on top)", "=IFERROR({c}5/{c}6,0)", NUM1),
    ("Break-even QUALIFIED meetings / day (plus transfers on top)", "=IFERROR({c}9*in_QualRate,0)", "0.00"),
    ("Break-even BOOKED meetings / day (plus transfers on top)", "=IFERROR({c}7/in_Days,0)", "0.00"),
    ("Target qualified / day", "=in_Target", NUM1),
    ("Headroom: target / break-even", "=IFERROR({c}10/{c}8,0)", '0.0"x"'),
    ("IF EVERY BOOKING BILLS: break-even BOOKED meetings / day", "=IFERROR({c}5/(in_Days*(in_RevBookedAll-in_CommBooked)),0)", "0.00"),
    ("IF EVERY BOOKING BILLS: same, in QUALIFIED meetings / day", "=IFERROR({c}12*in_QualRate,0)", "0.00"),
]
for i, (lab, f, fmt) in enumerate(BE, start=5):
    wsB.cell(i, 1, lab)
    for j in range(len(PAY)):
        col = L(2 + j)
        wsB.cell(i, 2 + j, f.format(c=col)).number_format = fmt
    bc = L(2 + len(PAY))
    fb = f.replace("INDEX(PayBase,MATCH({c}$4,PayNames,0))", "in_BlendBase")
    wsB.cell(i, 2 + len(PAY), fb.format(c=bc)).number_format = fmt
for c in range(2, 3 + len(PAY)):
    wsB.cell(8, c).fill = FILL_KPI; wsB.cell(8, c).font = F_B
name("out_BE_Blend", "Breakeven", f"${L(2+len(PAY))}$8")
name("out_BE_BlendBooked", "Breakeven", f"${L(2+len(PAY))}$9")
name("out_BE_AllBooked", "Breakeven", f"${L(2+len(PAY))}$12")
widths(wsB, {"A": 52, **{L(2 + j): 17 for j in range(len(PAY) + 1)}})

# ============================================================================ PAYBACK
wsP = wb.create_sheet("Payback")
wsP["A1"] = "Hiring-fee payback: months until a new setter covers the $1,500 fee + ramp losses"; wsP["A1"].font = F_TITLE
wsP["A2"] = ("Cumulative contribution after expensing the hiring fee in month 1. "
             "Payback month = first month the cumulative turns >= 0. Assumes P0 supply is available.")
wsP["A2"].font = F_NOTE
header(wsP, 4, ["Month", "Ramp %"] + [f"{p[0]}: monthly CM" for p in PAY] + [f"{p[0]}: cumulative" for p in PAY])
NP = len(PAY)
for m in range(1, 25):
    i = 4 + m
    wsP.cell(i, 1, m)
    wsP.cell(i, 2, f"=IF(A{i}=1,in_Ramp1,IF(A{i}=2,in_Ramp2,in_Ramp3))").number_format = PCT
    for j in range(NP):
        cm_col = 3 + j; cum_col = 3 + NP + j
        q = f"in_Target*$B{i}*in_Days"
        wsP.cell(i, cm_col, f"={q}/in_QualRate*(in_RevBookedQ-in_CommBooked)-INDEX(PayBase,{j+1})-in_SeatCost").number_format = USD
        prev = f"+{L(cum_col)}{i-1}" if m > 1 else "-in_Fee"
        wsP.cell(i, cum_col, f"={L(cm_col)}{i}{prev}").number_format = USD
wsP.cell(30, 1, "PAYBACK MONTH").font = F_B
for j in range(NP):
    cc = L(3 + NP + j)
    c = wsP.cell(30, 3 + NP + j, f'=IF(COUNTIF({cc}5:{cc}28,"<0")>=24,"Not within 24m",COUNTIF({cc}5:{cc}28,"<0")+1)')
    c.font = F_B; c.fill = FILL_KPI
    wsP.cell(29, 3 + NP + j, f"={cc}$4").font = F_NOTE
lat = PAY.index(next(p for p in PAY if p[0] == "LATAM full-time"))
name("out_Payback_LATAM", "Payback", f"${L(3+NP+lat)}$30")
# LATAM full-time at ACTUAL window productivity instead of target
ca, cu = 3 + 2 * NP, 4 + 2 * NP
wsP.cell(4, ca, "LATAM FT @ actual productivity: monthly CM"); wsP.cell(4, cu, "LATAM FT @ actual: cumulative")
for k in (ca, cu):
    wsP.cell(4, k).font, wsP.cell(4, k).fill = F_H, PatternFill("solid", fgColor="7F6000")
for m in range(1, 25):
    i = 4 + m
    wsP.cell(i, ca, f"=MIN(in_Target/in_QualRate,out_ActBPD)*$B{i}*in_Days*(in_RevBookedQ-in_CommBooked)"
                    f"-INDEX(PayBase,{lat+1})-in_SeatCost").number_format = USD
    prev = f"+{L(cu)}{i-1}" if m > 1 else "-in_Fee"
    wsP.cell(i, cu, f"={L(ca)}{i}{prev}").number_format = USD
c = wsP.cell(30, cu, f'=IF(COUNTIF({L(cu)}5:{L(cu)}28,"<0")>=24,"Not within 24m",COUNTIF({L(cu)}5:{L(cu)}28,"<0")+1)')
c.font = F_B; c.fill = FILL_KPI
name("out_Payback_LATAM_Act", "Payback", f"${L(cu)}$30")
wsP.cell(31, 1, "Ramp loss vs steady state (months 1-2), LATAM full-time").font = F_NOTE
wsP.cell(31, 3 + NP + lat, f"=2*{L(3+lat)}28-({L(3+lat)}5+{L(3+lat)}6)").number_format = USD
widths(wsP, {"A": 10, "B": 9, **{L(3 + k): 15 for k in range(2 * NP)}})
wsP.freeze_panes = "C5"

# ============================================================================ TEAM P&L
wsT = wb.create_sheet("Team_PnL")
wsT["A1"] = "Team P&L by headcount (monthly, steady state, current roster pay mix)"; wsT["A1"].font = F_TITLE
wsT["A2"] = ("Setters work the positive-reply ladder best tier first (P0, then P1 ... P4 carry-over), at each tier's measured "
             "productivity. Block B sends setters left over after P4 onto purchased cold data.")
wsT["A2"].font = F_NOTE
cols = [6, 12, 20, 30, 50, 100, "=in_ActiveN", "=in_ActiveN+12"]
header(wsT, 4, ["Line ($/month)", "6", "12", "20", "30", "50", "100", "Current team", "Current + 12", "Change from +12"])
NT = len(cols)
def tier_row(k):
    return (f"Booked from {LADDER[k][0].split()[0]} {' '.join(LADDER[k][0].split()[1:])}",
            f"=INDEX(Lad_Prod,{k+1})*in_Days*MAX(0,MIN(INDEX(Lad_Need,{k+1}),{{c}}7-INDEX(Lad_Cum,{k+1})))", NUM0)
TL = [
    ("Setters", None, NUM0),
    ("Blended base / setter", "=in_BlendBase", USD),
    ("Setter-months available", "={c}5", NUM1),
] + [tier_row(k) for k in range(6)] + [
    ("Booked meetings from P0-P4", "=SUM({c}8:{c}13)", NUM0),
    ("Qualified meetings from P0-P4 (live transfers come on top)", "={c}14*in_QualRate", NUM0),
    ("Setter-months left after P0-P4 is fully worked", "=MAX(0,{c}7-SUM(Lad_Need))", NUM1),
    ("A. POSITIVE-REPLY LADDER ONLY, QUALIFIED MEETINGS BILLED", None, None),
    ("Revenue (meetings + live transfers)", "={c}14*in_RevBookedQ", USD),
    ("Base pay", "={c}5*{c}6", USD),
    ("Commission (meetings + transfers)", "={c}14*in_CommBooked", USD),
    ("GoFunnel + Trainual + phones + mgmt", "={c}5*in_SeatCost", USD),
    ("Hiring fee (amortized; = annual churn re-hiring)", "={c}5*in_Fee/in_Tenure", USD),
    ("Total cost", "=SUM({c}19:{c}22)", USD),
    ("CONTRIBUTION MARGIN (ladder only)", "={c}18-{c}23", USD),
    ("CM per setter", "=IFERROR({c}24/{c}5,0)", USD),
    ("CM %", "=IFERROR({c}24/{c}18,0)", PCT),
    ("B. ADD PURCHASED COLD DATA FOR SETTERS LEFT OVER", None, None),
    ("Cold booked meetings", "={c}16*in_ColdBookDay*in_Days", NUM0),
    ("Cold qualified meetings", "={c}28*in_QualRate", NUM0),
    ("Cold data cost", "={c}28*in_ColdNums*in_ColdPerNum", USD),
    ("Revenue (meetings + live transfers)", "={c}14*in_RevBookedQ+{c}29*in_RevQEff", USD),
    ("Total cost (ladder costs + extra commission + cold data)", "={c}23+{c}29*in_Commission+{c}30", USD),
    ("CONTRIBUTION MARGIN (with cold data)", "={c}31-{c}32", USD),
    ("CM per setter", "=IFERROR({c}33/{c}5,0)", USD),
    ("CM %", "=IFERROR({c}33/{c}31,0)", PCT),
    ("C. IF EVERY BOOKING BILLS (commission still on qualified only)", None, None),
    ("Revenue, ladder only (all bookings and transfers billed)", "={c}14*in_RevBookedAll", USD),
    ("CONTRIBUTION MARGIN, ladder only (all billed)", "={c}37-{c}23", USD),
    ("CM %, ladder only (all billed)", "=IFERROR({c}38/{c}37,0)", PCT),
    ("Revenue with cold data (all bookings and transfers billed)", "={c}14*in_RevBookedAll+{c}28*in_RevQ", USD),
    ("CONTRIBUTION MARGIN with cold data (all billed)", "={c}40-{c}32", USD),
    ("CM % with cold data (all billed)", "=IFERROR({c}41/{c}40,0)", PCT),
    ("", None, None),
    ("Hiring fee if EXPENSED: one-off cost to hire this many from zero", "={c}5*in_Fee", USD),
]
for i, (lab, f, fmt) in enumerate(TL, start=5):
    wsT.cell(i, 1, lab)
    if lab.isupper() or lab.startswith(("A.", "B.", "C.")) or lab.startswith("Booked meetings from"):
        wsT.cell(i, 1).font = F_B
    if lab.startswith(("A.", "B.", "C.")):
        for k in range(1, NT + 3): wsT.cell(i, k).fill = FILL_SEC
        continue
    for k in range(NT):
        col = L(2 + k)
        v = cols[k] if i == 5 else (f.format(c=col) if f else None)
        if v is None: continue
        c = wsT.cell(i, 2 + k, v); c.number_format = fmt
        if i == 5 and k < 6: c.font = F_IN
        if lab.isupper(): c.font = F_B; c.fill = FILL_KPI
    if f or i == 5:
        d = L(2 + NT)
        c = wsT.cell(i, 2 + NT, f"={L(1+NT)}{i}-{L(NT)}{i}" if fmt != PCT else None)
        if fmt != PCT: c.number_format = fmt
for rr in (24, 33, 38, 41):
    wsT.conditional_formatting.add(f"B{rr}:{L(2+NT)}{rr}", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="F8CBAD")))
name("out_TeamCM_Now", "Team_PnL", f"${L(NT)}$24")
name("out_TeamCM_Plus12", "Team_PnL", f"${L(1+NT)}$24")
name("out_TeamCM_Delta", "Team_PnL", f"${L(2+NT)}$24")
name("out_TeamCMc_Delta", "Team_PnL", f"${L(2+NT)}$33")
name("out_TeamCMAll_Now", "Team_PnL", f"${L(NT)}$38")
name("out_TeamCMAll_Delta", "Team_PnL", f"${L(2+NT)}$38")
name("out_TeamCMcAll_Delta", "Team_PnL", f"${L(2+NT)}$41")
widths(wsT, {"A": 58, **{L(2 + k): 14 for k in range(NT + 1)}})
wsT.freeze_panes = "B5"

# ============================================================================ SUPPLY
wsU = wb.create_sheet("Supply")
wsU["A1"] = "How deep into the positive-reply ladder each headcount reaches, and when cold data starts"; wsU["A1"].font = F_TITLE
wsU["A3"] = "P0-P4 supply: qualified meetings / month available"; wsU["B3"] = "=in_P0CapQ"; wsU["B3"].number_format = NUM0
wsU["A4"] = "Booked meetings / month available on P0-P4"; wsU["B4"] = "=SUM(Lad_Avail)"; wsU["B4"].number_format = NUM0
wsU["A5"] = "P0-P4 FULLY WORKED AT (setters)"; wsU["A5"].font = F_B
wsU["B5"] = "=SUM(Lad_Need)"; wsU["B5"].number_format = NUM1
wsU["B5"].font = F_B; wsU["B5"].fill = FILL_KPI
name("out_P0RunOut", "Supply", "$B$5")
wsU["A6"] = "Current active setters"; wsU["B6"] = "=in_ActiveN"
wsU["A7"] = "Cold data cost per cold qualified meeting"; wsU["B7"] = "=IFERROR(in_ColdNums*in_ColdPerNum/in_QualRate,0)"; wsU["B7"].number_format = USD2
wsU["A8"] = "Margin per cold qualified meeting (rev - commission - data)"; wsU["B8"] = "=in_RevQEff-in_Commission-B7"; wsU["B8"].number_format = USD2
header(wsU, 10, ["Setters", "Booked from P0-P4", "Qualified from P0-P4", "Cold qualified", "Cold data cost",
                 "CM: ladder only", "CM: with cold data", "CM / setter (ladder only)", "Cold data on?"])
NS = list(range(1, 41)) + [45, 50, 60, 70, 80, 90, 100]
for i, n in enumerate(NS, start=11):
    wsU.cell(i, 1, n).font = F_IN
    x = f"($A{i}-Lad_Cum)"
    a_ = f"((ABS({x})+{x})/2)"
    wsU.cell(i, 2, f"=SUMPRODUCT(Lad_Prod*in_Days*(({a_}+Lad_Need-ABS({a_}-Lad_Need))/2))").number_format = NUM0
    wsU.cell(i, 3, f"=B{i}*in_QualRate").number_format = NUM0
    left = f"MAX(0,A{i}-SUM(Lad_Need))"
    wsU.cell(i, 4, f"={left}*in_ColdBookDay*in_Days*in_QualRate").number_format = NUM0
    wsU.cell(i, 5, f"={left}*in_ColdBookDay*in_Days*in_ColdNums*in_ColdPerNum").number_format = USD
    fixed = f"A{i}*(in_BlendBase+in_SeatCost+in_Fee/in_Tenure)"
    wsU.cell(i, 6, f"=B{i}*(in_RevBookedQ-in_CommBooked)-{fixed}").number_format = USD
    wsU.cell(i, 7, f"=F{i}+D{i}*(in_RevQEff-in_Commission)-E{i}").number_format = USD
    wsU.cell(i, 8, f"=F{i}/A{i}").number_format = USD
    wsU.cell(i, 9, f'=IF(E{i}>0,"COLD DATA ON","P0-P4 only")')
end = 10 + len(NS)
wsU.conditional_formatting.add(f"A11:I{end}", FormulaRule(formula=["$E11>0"], fill=PatternFill("solid", fgColor="FCE4D6")))
wsU.conditional_formatting.add(f"F11:H{end}", CellIsRule(operator="lessThan", formula=["0"], font=Font(color="C00000")))
widths(wsU, {"A": 54, "B": 14, "C": 13, "D": 13, "E": 14, "F": 14, "G": 18, "H": 18, "I": 15})
wsU.freeze_panes = "A11"

# ============================================================================ GROWTH
wsG = wb.create_sheet("Growth")
wsG["A1"] = "What has to be true to grow to 25, 50, 75 and 100 setters"; wsG["A1"].font = F_TITLE
wsG["A2"] = ("Path A: grow positive replies so every setter works the ladder at today's productivity. "
             "Path B: keep today's replies and put every setter beyond the ladder on purchased cold data.")
wsG["A2"].font = F_NOTE
GN = ["=SUM(Lad_Need)", 25, 50, 75, 100]
header(wsG, 4, ["Line (monthly)", "Ladder full (today)", "25", "50", "75", "100"])
FIX = "(in_BlendBase+in_SeatCost+in_Fee/in_Tenure)"
def ladder_booked(c):
    x = f"({c}5-Lad_Cum)"; a_ = f"((ABS({x})+{x})/2)"
    return f"SUMPRODUCT(Lad_Prod*in_Days*(({a_}+Lad_Need-ABS({a_}-Lad_Need))/2))"
GL = [
    ("Setters", None, NUM1),
    ("Setters today's positive replies can keep busy (ladder)", "=SUM(Lad_Need)", NUM1),
    ("Setters beyond today's ladder", "=MAX(0,{c}5-{c}6)", NUM1),
    ("PATH A. GROW POSITIVE REPLIES (same tier mix and productivity)", None, None),
    ("Positive-reply volume needed vs today", "={c}5/{c}6", '0.0"x"'),
    ("Funding positive replies (with phone) needed per week", "=in_PosWeek*{c}9", NUM0),
    ("Team contribution margin, qualified meetings billed", "={c}5*(SUM(Lad_Avail)/{c}6*(in_RevBookedQ-in_CommBooked)-" + FIX + ")", USD),
    ("Team contribution margin, every booking billed", "={c}5*(SUM(Lad_Avail)/{c}6*(in_RevBookedAll-in_CommBooked)-" + FIX + ")", USD),
    ("PATH B. KEEP TODAY'S REPLIES, BUY COLD DATA FOR THE REST", None, None),
    ("Break-even cold booked meetings per setter-day, qualified billed",
     "=IFERROR(" + FIX + "/(in_Days*(in_QualRate*(in_RevQEff-in_Commission)-in_ColdNums*in_ColdPerNum)),0)", "0.00"),
    ("Break-even cold booked meetings per setter-day, every booking billed",
     "=IFERROR(" + FIX + "/(in_Days*(in_RevQ-in_QualRate*in_Commission-in_ColdNums*in_ColdPerNum)),0)", "0.00"),
    ("Cold numbers to buy per month (at the assumed cold booking rate)", "={c}7*in_ColdBookDay*in_Days*in_ColdNums", NUM0),
    ("Cold data cost per month", "={c}16*in_ColdPerNum", USD),
    ("Team contribution margin, qualified billed (cold at assumed rate)",
     "=__LB__*(in_RevBookedQ-in_CommBooked)+{c}7*in_ColdBookDay*in_Days*(in_QualRate*(in_RevQEff-in_Commission)-in_ColdNums*in_ColdPerNum)-{c}5*" + FIX, USD),
    ("Team contribution margin, every booking billed (cold at assumed rate)",
     "=__LB__*(in_RevBookedAll-in_CommBooked)+{c}7*in_ColdBookDay*in_Days*(in_RevQ-in_QualRate*in_Commission-in_ColdNums*in_ColdPerNum)-{c}5*" + FIX, USD),
]
for i, (lab, f, fmt) in enumerate(GL, start=5):
    wsG.cell(i, 1, lab)
    if lab.startswith("PATH"):
        wsG.cell(i, 1).font = F_B
        for k in range(1, 7): wsG.cell(i, k).fill = FILL_SEC
        continue
    for k in range(5):
        col = L(2 + k)
        v = GN[k] if i == 5 else f.format(c=col).replace("__LB__", "(" + ladder_booked(col) + ")")
        c = wsG.cell(i, 2 + k, v); c.number_format = fmt
        if i == 5 and k > 0: c.font = F_IN
        if "Team contribution" in lab: c.font = F_B; c.fill = FILL_KPI
for rr in (11, 12, 18, 19):
    wsG.conditional_formatting.add(f"B{rr}:F{rr}", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="F8CBAD")))
    wsG.conditional_formatting.add(f"B{rr}:F{rr}", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor="C6EFCE")))
name("out_BEColdQ", "Growth", "$B$14"); name("out_BEColdAll", "Growth", "$B$15")
widths(wsG, {"A": 66, "B": 18, "C": 14, "D": 14, "E": 14, "F": 14})

# ============================================================================ SENSITIVITY
wsX = wb.create_sheet("Sensitivity")
wsX["A1"] = "Sensitivity: show rate x revenue per shown meeting"; wsX["A1"].font = F_TITLE
wsX["A2"] = ("Monthly contribution margin per setter at target (blended roster base, hiring amortized), "
             "if partners paid per SHOWN qualified meeting. Edit the row/column headers to change the grid.")
wsX["A2"].font = F_NOTE
wsX["A4"] = "Revenue per shown meeting  ↓  /  Show rate  →"; wsX["A4"].font = F_B
shows = [0.3, 0.4, 0.5, 0.6, 0.7, 0.8]
prices = [100, 125, 150, 175, 200, 225, 250, 275, 300]
for j, s in enumerate(shows):
    c = wsX.cell(4, 2 + j, s); c.number_format = "0%"; c.font = F_IN; c.fill = FILL_SEC
for i, p in enumerate(prices):
    rr = 5 + i
    c = wsX.cell(rr, 1, p); c.number_format = USD; c.font = F_IN; c.fill = FILL_SEC
    for j in range(len(shows)):
        col = L(2 + j)
        f = (f"=in_Target*in_Days*({col}$4*$A{rr}-in_Commission)"
             f"-(in_BlendBase+in_SeatCost+in_Fee/in_Tenure)")
        wsX.cell(rr, 2 + j, f).number_format = USD
rng = f"B5:{L(1+len(shows))}{4+len(prices)}"
wsX.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="F8CBAD")))
wsX.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=PatternFill("solid", fgColor="C6EFCE")))
wsX["A16"] = "Break-even revenue per shown meeting at each show rate"; wsX["A16"].font = F_B
for j in range(len(shows)):
    col = L(2 + j)
    wsX.cell(16, 2 + j, f"=IFERROR(((in_BlendBase+in_SeatCost+in_Fee/in_Tenure)/(in_Target*in_Days)+in_Commission)/{col}$4,0)").number_format = USD
widths(wsX, {"A": 44, **{L(2 + j): 13 for j in range(len(shows))}})

# ============================================================================ ACTUAL vs MODEL
wsV = wb.create_sheet("Actual_vs_Model")
wsV["A1"] = "Actual vs modeled, current setters (actuals window on Inputs; booking form)"; wsV["A1"].font = F_TITLE
wsV["A2"] = ("Current setters only (Roster Active = Y). Revenue = billable bookings x partner price. Costs prorated to paid "
             "days (start date or first booking, to window end); productivity and the modeled column use calling days.")
wsV["A2"].font = F_NOTE
VH = ["Setter", "Pay type", "Bookings", "Qualified (billable)", "False bookings", "Qualified rate",
      "Calling days", "Qualified / day", "Revenue (actual)", "Base (prorated)", "Commission",
      "Seat + overhead (prorated)", "Hiring fee (amortized, prorated)", "Total cost",
      "Contribution margin (actual)", "CM per selling day", "Revenue (modeled)", "Cost (modeled)",
      "Contribution margin (modeled)", "Actual - modeled", "False-booking status", "rank key", "Rank", "Paid days",
      "Revenue (all billed)", "CM (all billed)", "CM % (qualified only)", "CM % (all billed)", "Days with a booking", "Booked / day worked", "Live transfers"]
header(wsV, 3, VH)
ACTIVE_RR = [2 + k for k, x in enumerate(ROSTER) if x[5] == "Y"]  # roster rows of current setters
N_ACT = len(ACTIVE_RR)
first, last = 4, 3 + N_ACT
for k, rr in enumerate(ACTIVE_RR):
    i = first + k
    s = f"Roster!$A${rr}"
    F = {
        1: f"={s}", 2: f"=Roster!$C${rr}",
        3: f"=COUNTIFS(act_setter,A{i},act_inwin,1)",
        4: f"=COUNTIFS(act_setter,A{i},act_inwin,1,act_q,1)",
        5: f"=COUNTIFS(act_setter,A{i},act_inwin,1,act_false,1)",
        6: f'=IF(C{i}=0,"",D{i}/C{i})',
        7: (f"=IF(C{i}=0,0,NETWORKDAYS(IF(Roster!$K${rr}<>\"\",MAX(Roster!$K${rr},in_WinStart),"
            f"_xlfn.MINIFS(act_date,act_setter,A{i},act_inwin,1)),in_WinEnd))"),
        25: f"=SUMIFS(act_rev_all,act_setter,A{i},act_inwin,1)+AE{i}*in_XferPrice",
        26: f"=Y{i}-N{i}",
        27: f'=IF(I{i}=0,"",O{i}/I{i})',
        28: f'=IF(Y{i}=0,"",Z{i}/Y{i})',
        24: (f"=IF(AND(C{i}=0,Roster!$J${rr}=\"\"),0,NETWORKDAYS(IF(Roster!$J${rr}<>\"\",MAX(Roster!$J${rr},in_WinStart),"
             f"_xlfn.MINIFS(act_date,act_setter,A{i},act_inwin,1)),in_WinEnd))"),
        8: f"=IF(AC{i}=0,0,D{i}/AC{i})",
        29: f"=COUNTIFS(act_setter,A{i},act_inwin,1,act_dayfirst,1)",
        30: f"=IF(AC{i}=0,0,C{i}/AC{i})",
        31: f'=COUNTIFS(tr_setter,A{i},tr_date,">="&in_WinStart,tr_date,"<="&in_WinEnd)',
        9: f"=SUMIFS(act_rev,act_setter,A{i},act_inwin,1)+AE{i}*in_XferPrice",
        10: f"=Roster!$H${rr}*X{i}/in_Days",
        11: f"=D{i}*Roster!$I${rr}+AE{i}*in_XferComm",
        12: f"=in_SeatCost*X{i}/in_Days",
        13: f"=in_Fee/in_Tenure*X{i}/in_Days",
        14: f"=SUM(J{i}:M{i})",
        15: f"=I{i}-N{i}",
        16: f"=IF(G{i}=0,0,O{i}/G{i})",
        17: f"=in_Target*G{i}*in_RevQEff",
        18: f"=J{i}+L{i}+M{i}+in_Target*G{i}*Roster!$I${rr}",
        19: f"=Q{i}-R{i}",
        20: f"=O{i}-S{i}",
        21: f'=IF(E{i}>=3,"TERMINATE (3+)",IF(E{i}>0,"WARNING ("&E{i}&")",""))',
        22: f"=IF(G{i}=0,-1E9,O{i})+ROW()/1E6",
        23: f"=RANK(V{i},$V${first}:$V${last})",
    }
    fm = {30: "0.00", 6: PCT, 8: "0.00", 25: USD, 26: USD, 27: PCT, 28: PCT, 9: USD, 10: USD, 11: USD, 12: USD, 13: USD, 14: USD, 15: USD,
          16: USD, 17: USD, 18: USD, 19: USD, 20: USD}
    for col, f in F.items():
        c = wsV.cell(i, col, f)
        if col in fm: c.number_format = fm[col]
    wsV.cell(i, 15).font = F_B
wsV.conditional_formatting.add(f"O{first}:O{last}", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="F8CBAD")))
wsV.conditional_formatting.add(f"O{first}:O{last}", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor="C6EFCE")))
tr = last + 1
wsV.cell(tr, 1, "TEAM TOTAL").font = F_B
for col in (3, 4, 5, 7, 9, 10, 11, 12, 13, 14, 15, 17, 18, 19, 20, 25, 26, 29, 31):
    c = wsV.cell(tr, col, f"=SUM({L(col)}{first}:{L(col)}{last})"); c.font = F_B
    if col >= 9: c.number_format = USD
name("out_ActTeamCM", "Actual_vs_Model", f"$O${tr}")
name("out_ActTeamRev", "Actual_vs_Model", f"$I${tr}")
name("out_ActTeamCost", "Actual_vs_Model", f"$N${tr}")
name("out_ActProfitable", "Actual_vs_Model", f"$C${tr+1}")
wsV.cell(tr + 1, 1, "Setters with positive CM (active in window)")
wsV.cell(tr + 1, 3, f'=COUNTIFS(O{first}:O{last},">0",G{first}:G{last},">0")').font = F_B
wsV.cell(tr + 2, 1, "Setters active in window")
wsV.cell(tr + 2, 3, f'=COUNTIF(G{first}:G{last},">0")')
name("out_ActActive", "Actual_vs_Model", f"$C${tr+2}")
for col, f_ in ((27, f"=IFERROR(O{tr}/I{tr},0)"), (28, f"=IFERROR(Z{tr}/Y{tr},0)")):
    c = wsV.cell(tr, col, f_); c.number_format = PCT; c.font = F_B
name("out_ActTeamRevAll", "Actual_vs_Model", f"$Y${tr}")
name("out_ActTeamCMAll", "Actual_vs_Model", f"$Z${tr}")
name("out_ActTeamPctQ", "Actual_vs_Model", f"$AA${tr}")
name("out_ActTeamPctAll", "Actual_vs_Model", f"$AB${tr}")
wsV.cell(tr + 1, 25, f'=COUNTIFS(Z{first}:Z{last},">0",G{first}:G{last},">0")').font = F_B
name("out_ActProfitableAll", "Actual_vs_Model", f"$Y${tr+1}")
wsV.cell(tr + 1, 24, "All billed:")
wsV.cell(tr + 3, 1, "Average qualified meetings per day worked (days with a booking)")
c = wsV.cell(tr + 3, 3, f"=IFERROR(D{tr}/AC{tr},0)"); c.number_format = "0.00"; c.font = F_B
wsV.cell(tr + 4, 1, "Average booked meetings per day worked")
c = wsV.cell(tr + 4, 3, f"=IFERROR(C{tr}/AC{tr},0)"); c.number_format = "0.00"; c.font = F_B
name("out_ActBPD", "Actual_vs_Model", f"$C${tr+4}")
name("out_ActXfers", "Actual_vs_Model", f"$AE${tr}")
name("in_ActQPD", "Actual_vs_Model", f"$C${tr+3}")

# ranked view
rk = tr + 6
wsV.cell(rk - 1, 1, "RANKED BY CONTRIBUTION MARGIN (actual, window)").font = F_TITLE
RH = ["Rank", "Setter", "Pay type", "Active days", "Qualified / day", "Qualified rate", "Revenue",
      "Total cost", "CM (actual)", "CM (modeled)", "Actual - modeled", "False-booking status",
      "CM % (qualified only)", "Revenue (all billed)", "CM (all billed)", "CM % (all billed)", "Booked / day worked", "Live transfers"]
header(wsV, rk, RH)
src_cols = {2: "A", 3: "B", 4: "G", 5: "H", 6: "F", 7: "I", 8: "N", 9: "O", 10: "S", 11: "T", 12: "U",
            13: "AA", 14: "Y", 15: "Z", 16: "AB", 17: "AD", 18: "AE"}
for k in range(N_ACT):
    i = rk + 1 + k
    wsV.cell(i, 1, k + 1)
    for col, sc in src_cols.items():
        c = wsV.cell(i, col, f"=INDEX(${sc}${first}:${sc}${last},MATCH($A{i},$W${first}:$W${last},0))")
        c.number_format = {4: NUM0, 5: "0.00", 6: PCT, 13: PCT, 16: PCT, 14: USD, 15: USD, 17: "0.00", 18: NUM0}.get(col, USD if 7 <= col <= 11 else "General")
rk_end = rk + N_ACT
wsV.conditional_formatting.add(f"I{rk+1}:I{rk_end}", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="F8CBAD")))
wsV.conditional_formatting.add(f"I{rk+1}:I{rk_end}", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor="C6EFCE")))
for nm, rr_ in (("out_Top1", rk + 1), ("out_Top2", rk + 2), ("out_Top3", rk + 3)):
    name(nm, "Actual_vs_Model", f"$B${rr_}")
widths(wsV, {"A": 13, "B": 17, **{L(k): 13 for k in range(3, 24)}})
wsV.column_dimensions["U"].width = 18
wsV.column_dimensions["L"].width = 16
wsV.freeze_panes = "C4"

# ============================================================================ SUMMARY
wsM = wb.create_sheet("Summary", 0)
wsM["A1"] = "Renaissance setter floor - P&L summary"; wsM["A1"].font = Font(bold=True, size=16, color=NAVY)
wsM["A2"] = ("All figures are live formulas. Yellow inputs on the Inputs tab are guesses: "
             "ramp, tenure, overhead and cold-lane conversion matter most. Price = $155/meeting (Collins $260/held).")
wsM["A2"].font = F_NOTE
wsM["A3"] = ("Contribution margin = partner revenue minus setter-floor costs. It does NOT include the email "
             "infrastructure cost of generating P0 replies (company KPI: ~$366 modelled cost per meeting, all channels).")
wsM["A3"].font = Font(italic=True, color="C00000", size=9)
KP = [
    ("SEC", "Unit economics (steady state, per setter per month)"),
    ("Blended revenue per qualified meeting", "=in_RevQEff", USD2),
    ("Qualified rate (booked -> qualified, since mandate)", "=in_QualRate", PCT),
    ("Contribution margin / setter at 5 qualified/day: LATAM full-time", "=Setter_PnL!D19", USD),
    ("Contribution margin / setter at 5 qualified/day: PH full-time", "=Setter_PnL!B19", USD),
    ("Contribution margin / setter at 5 qualified/day: Egypt (NAC)", "=Setter_PnL!F19", USD),
    ("Break-even QUALIFIED meetings / day (roster mix)", "=out_BE_Blend", "0.00"),
    ("Break-even BOOKED meetings / day (roster mix)", "=out_BE_BlendBooked", "0.00"),
    ("Hiring-fee payback, LATAM full-time, at target (month #)", "=out_Payback_LATAM", "General"),
    ("Actual booked meetings per day worked (window)", "=out_ActBPD", "0.00"),
    ("Actual qualified meetings per day worked (window)", "=in_ActQPD", "0.00"),
    ("Live transfers in window (setters, to GreenBridge)", "=out_ActXfers", NUM0),
    ("Hiring-fee payback, LATAM full-time, at ACTUAL productivity (month #)", "=out_Payback_LATAM_Act", "General"),
    ("SEC", "Supply and scale"),
    ("P0-P4 supply: qualified meetings / month available", "=in_P0CapQ", NUM0),
    ("P0-P4 ladder fully worked at (setters)", "=out_P0RunOut", NUM1),
    ("Current active setters", "=in_ActiveN", NUM0),
    ("Team CM / month, current team, P0-P4 ladder only", "=out_TeamCM_Now", USD),
    ("Team CM / month, current + 12, P0-P4 ladder only", "=out_TeamCM_Plus12", USD),
    ("Change in CM from adding 12 setters, P0-P4 ladder only", "=out_TeamCM_Delta", USD),
    ("Change in CM from adding 12 setters, with purchased cold data", "=out_TeamCMc_Delta", USD),
    ("SEC", "Actuals (booking form, window on Inputs)"),
    ("Setters active in window", "=out_ActActive", NUM0),
    ("Setters with positive contribution margin", "=out_ActProfitable", NUM0),
    ("Team revenue in window (billable bookings x price)", "=out_ActTeamRev", USD),
    ("Team cost in window", "=out_ActTeamCost", USD),
    ("Team contribution margin in window", "=out_ActTeamCM", USD),
    ("Contribution margin % in window (qualified only billed)", "=out_ActTeamPctQ", PCT),
    ("Top 3 by CM", '=out_Top1&", "&out_Top2&", "&out_Top3', "General"),
    ("SEC", "If every booking bills (commission still on qualified only)"),
    ("Team revenue in window (all billed)", "=out_ActTeamRevAll", USD),
    ("Team contribution margin in window (all billed)", "=out_ActTeamCMAll", USD),
    ("Contribution margin % in window (all billed)", "=out_ActTeamPctAll", PCT),
    ("Setters with positive CM (all billed)", "=out_ActProfitableAll", NUM0),
    ("Break-even BOOKED meetings / day, roster mix (all billed)", "=out_BE_AllBooked", "0.00"),
    ("Team CM / month, current team, P0-P4 ladder only (all billed)", "=out_TeamCMAll_Now", USD),
    ("Break-even cold booked meetings per setter-day (qualified billed)", "=out_BEColdQ", "0.00"),
    ("Break-even cold booked meetings per setter-day (every booking billed)", "=out_BEColdAll", "0.00"),
    ("Change in CM from adding 12 setters, P0-P4 ladder only (all billed)", "=out_TeamCMAll_Delta", USD),
    ("Change in CM from adding 12 setters, with purchased cold data (all billed)", "=out_TeamCMcAll_Delta", USD),
]
r = 4
for item in KP:
    if item[0] == "SEC":
        section(wsM, r, item[1], 2); r += 1; continue
    lab, f, fmt = item
    wsM.cell(r, 1, lab)
    c = wsM.cell(r, 2, f); c.number_format = fmt; c.font = F_B; c.fill = FILL_KPI; c.border = BOX
    r += 1
r += 1
section(wsM, r, "Inputs that are GUESSED (confirm these first)", 2); r += 1
for g in ["Ramp: month 1 = 50%, month 2 = 80%", "Expected tenure: 6 months",
          "Management overhead $350/seat; Trainual $8; phone numbers $10",
          "PH part-time and LATAM junior base pay; pay type of the Nov 13 class (set to LATAM full-time)",
          "Purchased cold data: 2 booked meetings a day per setter, 450 numbers per booked meeting"]:
    wsM.cell(r, 1, "• " + g); r += 1
r += 1
section(wsM, r, "Data sources", 2); r += 1
for s_ in ["Bookings: booking web form (renaissance-booking.com) via Setter_Scorecard_Baseline_Oct2026 "
           "('Bookings Oct 2026 FINAL' + 'Bookings Live'), Oct 1 - Nov 17 2026, recording-verified",
           "Billable rule: for Nov 1-15 the Caller QA grade decides (only 'Qualified' bills and earns commission; "
           "Review is unbilled until re-listened). Other dates use the Scorecard's Min qualified flag",
           "Show rates: partner-no-show-labels-summary (Jul 1 - Nov 5 2026)",
           "Pay: 'Setter comp - current state' sheet + Slack (#ops-amir-julian); P0 volume: weekly KPI packs (#leaders)",
           "ClickHouse was not reachable from the build environment; refresh_actuals.py can pull from it once access is set up"]:
    wsM.cell(r, 1, "• " + s_).font = F_NOTE; r += 1
widths(wsM, {"A": 70, "B": 22})

for ws in wb.worksheets:
    ws.sheet_view.showGridLines = ws.title in ("Actuals",)
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
calc = recalc(OUT)  # store computed values so previews/viewers show numbers; formulas are kept
if calc and calc.exists():
    shutil.copy(calc, OUT)
print(f"Wrote {OUT} ({len(rows)} bookings, {N_ROSTER} roster rows)")
