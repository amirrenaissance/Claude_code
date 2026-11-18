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
S = {r[0]: r[1] for r in wb["Summary"].iter_rows(min_row=4, max_row=40, values_only=True) if r[0]}
inp = {r[0]: r[1] for r in wb["Inputs"].iter_rows(min_row=5, max_row=40, values_only=True) if r[0]}
k = {
    "rev_per_q": S["Blended revenue per qualified meeting"],
    "qual_rate": S["Qualified rate (booked -> qualified, since mandate)"],
    "cm_latam": S["Contribution margin / setter at 5 qualified/day: LATAM full-time"],
    "cm_ph": S["Contribution margin / setter at 5 qualified/day: PH full-time"],
    "cm_egypt": S["Contribution margin / setter at 5 qualified/day: Egypt (NAC)"],
    "be_q_day": S["Break-even QUALIFIED meetings / day (roster mix)"],
    "be_booked_day": S["Break-even BOOKED meetings / day (roster mix)"],
    "payback_target": S["Hiring-fee payback, LATAM full-time, at target (month #)"],
    "act_q_day": S["Actual qualified meetings per active setter-day (window)"],
    "payback_actual": S["Hiring-fee payback, LATAM full-time, at ACTUAL productivity (month #)"],
    "p0_cap": S["P0 supply cap (qualified meetings / month)"],
    "p0_runout": S["P0 runs out at (setters at target)"],
    "active_n": S["Current active setters"],
    "team_cm_now": S["Team CM / month, current team, P0 only"],
    "team_cm_plus12": S["Team CM / month, current + 12, P0 only"],
    "delta12_p0": S["Change in CM from adding 12 setters, P0 only"],
    "delta12_cold": S["Change in CM from adding 12 setters, with cold lane"],
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
    "team_cm_now_all": S["Team CM / month, current team, P0 only (all billed)"],
    "delta12_p0_all": S["Change in CM from adding 12 setters, P0 only (all billed)"],
    "delta12_cold_all": S["Change in CM from adding 12 setters, with cold lane (all billed)"],
    "target_q_day": inp["Target qualified meetings per setter per day (OTE)"],
    "win_start": str(inp["Actuals window start (last 60 days)"])[:10],
    "win_end": str(inp["Actuals window end (data through)"])[:10],
    "data_start": "2026-10-01",
}
dump("kpis", [{kk: num(v) for kk, v in k.items()}], "Summary + Inputs tabs")

# ---- unit P&L per pay type (Setter_PnL rows 10-19, Breakeven row 8)
sp, be = wb["Setter_PnL"], wb["Breakeven"]
rows = []
for c in range(2, 8):
    v = lambda r: sp.cell(r, c).value
    rows.append({"pay_type": v(4), "revenue": num(v(10)), "base": num(v(11)), "commission": num(v(12)),
                 "gofunnel": num(v(13)), "other_seat": num(v(14) + v(15) + v(16)), "hiring": num(v(17)),
                 "total_cost": num(v(18)), "cm": num(v(19)), "be_q_day": num(be.cell(8, c).value)})
dump("unit_pnl", rows, "Setter_PnL + Breakeven tabs")

# ---- team P&L scenarios (Team_PnL columns B..I)
t = wb["Team_PnL"]
labels = ["6", "12", "20", "30", "50", "100", "Current team", "Current + 12"]
rows = []
for i, c in enumerate(range(2, 10)):
    v = lambda r: t.cell(r, c).value
    rows.append({"scenario": labels[i], "setters": int(v(5)), "capacity_q": num(v(7), 0), "p0_q": num(v(9), 0),
                 "revenue_p0": num(v(12), 0), "cost_p0": num(v(17), 0), "cm_p0": num(v(18), 0),
                 "cold_q": num(v(23), 0), "cold_data": num(v(25), 0), "revenue_cold": num(v(26), 0),
                 "cm_cold": num(v(28), 0), "revenue_p0_all": num(v(35), 0), "cm_p0_all": num(v(36), 0),
                 "revenue_cold_all": num(v(38), 0), "cm_cold_all": num(v(39), 0)})
dump("team", rows, "Team_PnL tab")

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
                 "margin_all": num(g(16), 3) if isinstance(g(16), (int, float)) else None})
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
