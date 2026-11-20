#!/usr/bin/env python3
"""Export the computed results of Renaissance_Setter_PnL.xlsx as JSON files for the dashboard.

Reads the workbook's stored (recalculated) values and writes one {rows, meta} JSON per
dataset into dashboard_data/. Run build_model.py or refresh_actuals.py first so the
values are current.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
XLSX = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "Renaissance_Setter_PnL.xlsx"
OUT = HERE / "dashboard_data"
OUT.mkdir(exist_ok=True)
CODE = Path(__file__).read_text()
wb = openpyxl.load_workbook(XLSX, data_only=True)


def num(v, nd=2):
    return round(float(v), nd) if isinstance(v, (int, float)) else v


def dump(name: str, rows: list[dict], source: str) -> None:
    meta = [{"label": "export_dashboard_data.py", "value": CODE, "type": "code", "lang": "python"},
            {"label": "Input", "value": f"Renaissance_Setter_PnL.xlsx › {source}"}]
    (OUT / f"{name}.json").write_text(json.dumps({"rows": rows, "meta": meta}, default=str))
    print(f"{name}: {len(rows)} rows")


# ---- KPIs (Summary tab, matched by label)
S = {r[0]: r[1] for r in wb["Summary"].iter_rows(min_row=4, max_row=120, values_only=True) if r[0]}
inp = {r[0]: r[1] for r in wb["Inputs"].iter_rows(min_row=5, max_row=60, values_only=True) if r[0]}
k = {
    "rev_per_q": S["Blended revenue per qualified meeting"],
    "qual_rate": S["Qualified rate (booked -> qualified, since mandate)"],
    "cm_latam": S["Contribution margin / setter at 5 qualified/day: LATAM full-time"],
    "cm_ph": S["Contribution margin / setter at 5 qualified/day: PH full-time"],
    "cm_egypt": S["Contribution margin / setter at 5 qualified/day: Egypt (NAC)"],
    "be_q_day": S["Break-even QUALIFIED meetings / day (roster mix)"],
    "be_booked_day": S["Break-even BOOKED meetings / day (roster mix)"],
    "payback_target": S["Hiring-fee payback, LATAM full-time, at target (month #)"],
    "act_q_day": S["Actual qualified meetings per day worked (window)"],
    "act_b_day": S["Actual booked meetings per day worked (window)"],
    "win_xfers": S["Live transfers in window (setters, to GreenBridge)"],
    "payback_actual": S["Hiring-fee payback, LATAM full-time, at ACTUAL productivity (month #)"],
    "p0_cap": S["P0-P4 supply: qualified meetings / month available"],
    "p0_runout": S["P0-P4 ladder fully worked at (setters)"],
    "active_n": S["Current active setters"],
    "team_cm_now": S["Team CM / month, current team, P0-P4 ladder only"],
    "team_cm_plus12": S["Team CM / month, current + 12, P0-P4 ladder only"],
    "delta12_p0": S["Change in CM from adding 12 setters, P0-P4 ladder only"],
    "delta12_cold": S["Change in CM from adding 12 setters, with purchased cold data"],
    "win_active": S["Setters active in window"],
    "win_profitable": S["Setters with positive contribution margin"],
    "win_rev": S["Team revenue in window (billable bookings x price)"],
    "win_cost": S["Team cost in window"],
    "win_cm": S["Team contribution margin in window"],
    "win_pct": S["Contribution margin % in window (qualified only billed)"],
    "win_rev_all": S["Team revenue in window (all billed)"],
    "win_cm_all": S["Team contribution margin in window (all billed)"],
    "win_pct_all": S["Contribution margin % in window (all billed)"],
    "win_profitable_all": S["Setters with positive CM (all billed)"],
    "be_booked_day_all": S["Break-even BOOKED meetings / day, roster mix (all billed)"],
    "team_cm_now_all": S["Team CM / month, current team, P0-P4 ladder only (all billed)"],
    "delta12_p0_all": S["Change in CM from adding 12 setters, P0-P4 ladder only (all billed)"],
    "delta12_cold_all": S["Change in CM from adding 12 setters, with purchased cold data (all billed)"],
    "target_q_day": inp["Target qualified meetings per setter per day (OTE)"],
    "win_start": str(inp["Actuals window start (last 60 days)"])[:10],
    "win_end": str(inp["Actuals window end (data through)"])[:10],
    "data_start": "2026-10-01",
    "xfer_share": inp["Live transfers as a share of all outcomes (transfers + meetings)"],
    "rev_outcome_q": inp["Revenue per booked meeting incl. its transfers, qualified meetings billed"],
    "rev_outcome_all": inp["Revenue per booked meeting incl. its transfers, every booking billed"],
    "pos_week": inp["Funding positive replies with a phone entering the dialer per week"],
    "be_cold_q": S["Break-even cold booked meetings per setter-day (qualified billed)"],
    "opp_cost": S["Email cost per positive reply"],
    "cost_per_pos": S["Email cost per positive reply"],
    "outcomes_month": S["Booked meetings + transfers per month, current team (model)"],
    "setter_cost_outcome": S["Setter cost per booked meeting or transfer (model)"],
    "true_cost": S["True cost per booked meeting or transfer: setter + positive reply"],
    "rev_outcome_q2": S["Revenue per booked meeting or transfer (qualified billed)"],
    "rev_outcome_all2": S["Revenue per booked meeting or transfer (all billed)"],
    "pos_cost_booked": S["Email cost of the positive replies setters book, per month"],
    "be_cpp_q": S["Break-even cost per positive reply (qualified billed)"],
    "be_cpp_all": S["Break-even cost per positive reply (all billed)"],
    "fm_oct_mtgs": S["October: funding meetings booked, all channels"],
    "fm_oct_call": S["October: of which booked by setter calls"],
    "fm_oct_rev": S["October: Renny revenue (all funding partner lines)"],
    "fm_oct_rpm": S["October: revenue per booked funding meeting"],
    "fm_nov_mtgs": S["November to date: funding meetings booked, all channels"],
    "fm_nov_rev": S["November to date: Renny revenue"],
    "cost_per_mtg": S["Setter cost per booked meeting or transfer (window)"],
    "full_cost_per_q": S["Fully loaded cost per billable outcome (window, qualified billed)"],
    "win_cm_full": S["Fully loaded margin in window (qualified billed)"],
    "team_cm_full_q": S["Team CM / month, current team, fully loaded (qualified billed)"],
    "team_cm_full_all": S["Team CM / month, current team, fully loaded (all billed)"],
    "incremental": S["Share of setter meetings that are net new"],
    "cm_incr_q": S["Team CM / month, current team, incremental only (qualified billed)"],
    "cm_incr_all": S["Team CM / month, current team, incremental only (all billed)"],
    "be_incr_q": S["Break-even net-new share (qualified billed)"],
    "be_incr_all": S["Break-even net-new share (all billed)"],
    "be_cold_all": S["Break-even cold booked meetings per setter-day (every booking billed)"],
}
dump("kpis", [{kk: num(v) for kk, v in k.items()}], "Summary + Inputs tabs")

# ---- funding meetings sold by month (Funding_Monthly tab)
fmw = wb["Funding_Monthly"]
rows = []
for c, lab in ((2, "Sep 2026"), (3, "Oct 2026"), (4, "Nov 1-18")):
    v = lambda r: fmw.cell(r, c).value
    rows.append({"month": lab, "email": num(v(5), 0), "call": num(v(6), 0), "other": num(v(7), 0), "total": num(v(8), 0),
                 "setter_share": num(v(9), 3), "revenue": num(v(10), 0), "rev_per_mtg": num(v(11)), "sends": num(v(13), 0)})
dump("funding_monthly", rows, "Funding_Monthly tab (ClickHouse: booking portal + Renny ledger)")

# ---- unit P&L per pay type (Setter_PnL rows 10-19, Breakeven row 8)
sp, be = wb["Setter_PnL"], wb["Breakeven"]
rows = []
for c in range(2, 8):
    v = lambda r: sp.cell(r, c).value
    rows.append({"pay_type": v(4), "revenue": num(v(10)), "base": num(v(11)), "commission": num(v(12)),
                 "gofunnel": num(v(13)), "other_seat": num(v(14) + v(15)), "management": num(v(16)), "hiring": num(v(17)),
                 "total_cost": num(v(18)), "cm": num(v(19)), "be_q_day": num(be.cell(8, c).value)})
dump("unit_pnl", rows, "Setter_PnL + Breakeven tabs")

# ---- team P&L scenarios (Team_PnL columns B..I)
t = wb["Team_PnL"]
labels = ["6", "12", "20", "30", "50", "100", "Current team", "Current + 12"]
rows = []
for i, c in enumerate(range(2, 10)):
    v = lambda r: t.cell(r, c).value
    rows.append({"scenario": labels[i], "setters": int(v(5)), "booked_ladder": num(v(14), 0), "p0_q": num(v(15), 0),
                 "revenue_p0": num(v(18), 0), "cost_p0": num(v(23), 0), "cm_p0": num(v(24), 0),
                 "cold_q": num(v(29), 0), "cold_data": num(v(30), 0), "revenue_cold": num(v(31), 0),
                 "cm_cold": num(v(33), 0), "revenue_p0_all": num(v(37), 0), "cm_p0_all": num(v(38), 0),
                 "revenue_cold_all": num(v(40), 0), "cm_cold_all": num(v(41), 0),
                 **{f"tier{k}": num(v(8 + k), 0) for k in range(6)}})
dump("team", rows, "Team_PnL tab")

# ---- growth (Growth tab)
g_ = wb["Growth"]
labels_g = ["Ladder full (today)", "25", "50", "75", "100"]
rows = []
for k_, c in enumerate(range(2, 7)):
    v = lambda r: g_.cell(r, c).value
    rows.append({"scenario": labels_g[k_], "setters": num(v(5), 1), "beyond": num(v(7), 1), "pos_multiple": num(v(9), 2),
                 "pos_week": num(v(10), 0), "cm_a_q": num(v(11), 0), "cm_a_all": num(v(12), 0),
                 "cold_cost": num(v(17), 0), "cm_b_q": num(v(18), 0), "cm_b_all": num(v(19), 0)})
dump("growth", rows, "Growth tab")
rows = []
for j in range(4):
    ra, rb, rc = 22 + j, 26 + j, 30 + j
    rows.append({"share": g_.cell(ra, 7).value, "a_today": num(g_.cell(ra, 2).value, 0), "a_50": num(g_.cell(ra, 4).value, 0),
                 "a_100": num(g_.cell(ra, 6).value, 0), "b_50": num(g_.cell(rb, 4).value, 0), "b_100": num(g_.cell(rb, 6).value, 0),
                 "be_cold": num(g_.cell(rc, 2).value, 2)})
dump("xfer_sens", rows, "Growth tab › transfer-share sensitivity")

# ---- supply ladder (Inputs tab ladder table)
wi = wb["Inputs"]
hdr_row = next(r for r in range(1, wi.max_row + 1) if wi.cell(r, 1).value == "Tier")
rows = []
for r in range(hdr_row + 1, hdr_row + 7):
    g = lambda c: wi.cell(r, c).value
    rows.append({"tier": g(1).replace("  ", " "), "bookings_window": g(2), "handled_window": g(3), "book_rate": num(g(4), 4),
                 "avail_month": num(g(5), 0), "booked_per_setter_day": num(g(6), 2), "setter_months": num(g(7), 1),
                 "what": g(9)})
dump("ladder", rows, "Inputs tab › Supply ladder")

# ---- supply curve (Supply rows 11+)
rows = []
for r in wb["Supply"].iter_rows(min_row=11, values_only=True):
    if r[0] is None:
        break
    rows.append({"setters": int(r[0]), "p0_q": num(r[2], 0), "cold_q": num(r[3], 0),
                 "cold_data": num(r[4], 0), "cm_p0": num(r[5], 0), "cm_cold": num(r[6], 0)})
dump("supply", rows, "Supply tab")

# ---- sensitivity grid
x = wb["Sensitivity"]
rows = []
for r in range(5, 14):
    for c in range(2, 8):
        rows.append({"price": x.cell(r, 1).value, "show_rate": x.cell(4, c).value,
                     "cm": num(x.cell(r, c).value, 0)})
dump("sensitivity", rows, "Sensitivity tab")

# ---- setters, ranked (Actual_vs_Model ranked view)
v = wb["Actual_vs_Model"]
hdr = next(r for r in range(1, v.max_row + 1) if v.cell(r, 1).value == "Rank")
rows = []
for r in range(hdr + 1, v.max_row + 1):
    if not isinstance(v.cell(r, 1).value, int):
        break
    g = lambda c: v.cell(r, c).value
    rows.append({"rank": g(1), "setter": g(2), "pay_type": g(3), "active_days": g(4),
                 "q_per_day": num(g(5)), "qual_rate": num(g(6), 3) if g(6) != "" else None,
                 "revenue": num(g(7), 0), "cost": num(g(8), 0), "cm": num(g(9), 0),
                 "cm_modeled": num(g(10), 0), "gap": num(g(11), 0), "status": g(12) or "",
                 "margin_q": num(g(13), 3) if isinstance(g(13), (int, float)) else None,
                 "revenue_all": num(g(14), 0), "cm_all": num(g(15), 0),
                 "margin_all": num(g(16), 3) if isinstance(g(16), (int, float)) else None,
                 "booked_per_day": num(g(17)), "transfers": g(18) or 0, "cm_full": num(g(19), 0),
                 "cost_per_mtg": num(g(20)), "full_cost_per_q": num(g(21)) if isinstance(g(21), (int, float)) else None})
dump("setters", rows, "Actual_vs_Model tab (ranked view)")

# ---- assumptions (Inputs: scalars + pay types)
rows = []
for r in wb["Inputs"].iter_rows(min_row=5, values_only=True):
    if r[0] and r[3] in ("GUESS", "CONFIRMED", "ACTUAL", "DERIVED", "INPUT"):
        val = r[1]
        rows.append({"assumption": r[0], "value": str(val)[:10] if hasattr(val, "year") else num(val, 3),
                     "unit": r[2] if isinstance(r[2], str) and len(r[2]) < 15 else "", "status": r[3],
                     "source": r[4] or ""})
    elif r[0] and r[2] in ("GUESS", "CONFIRMED", "DERIVED") and isinstance(r[1], (int, float)):
        rows.append({"assumption": f"Base pay: {r[0]}", "value": r[1], "unit": "$/mo", "status": r[2],
                     "source": r[3] or ""})
dump("assumptions", rows, "Inputs tab")
