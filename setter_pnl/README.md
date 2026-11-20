# Setter floor P&L model

`Renaissance_Setter_PnL.xlsx`: a per-setter and team P&L for the setter floor. Every calculation cell is a live formula driven by the **Inputs** tab.

| Tab | What it shows |
|---|---|
| Summary | Key results for the CEO |
| Inputs | Every assumption. Yellow cells are guesses; green cells are computed from actual data |
| Setter_PnL | Monthly P&L per setter, by pay type, with the hiring fee expensed and amortized |
| Breakeven | Break-even qualified and booked meetings per day |
| Payback | Months to recover the $1,500 hiring fee plus ramp losses, at target and at actual productivity |
| Team_PnL | Team P&L at 6, 12, 20, 30, 50 and 100 setters, the current team, and the current team + 12 |
| Supply | The headcount where P0 replies run out and cold-data cost starts |
| Sensitivity | Show rate × revenue per shown meeting |
| Actual_vs_Model | Current setters over the last 60 days, ranked by contribution margin |
| Actuals / Roster | Booking-form rows, and each setter's pay type |

## Data
- Bookings: booking web form, via the Setter Scorecard sheet (`Bookings Oct 2026 FINAL` + `Bookings Live`), saved in `data/bookings_webform.csv`.
- Revenue only counts bookings that are recording-verified, have **Min qualified = Yes**, and are not a false booking. Use `Inputs › Share of unqualified bookings that still bill` to test other billing assumptions.

## Refresh
```bash
python refresh_actuals.py --source csv --csv data/bookings_webform.csv
python refresh_actuals.py --source clickhouse --days 60            # check BOOKINGS_SQL table/columns first
python refresh_actuals.py --source clickhouse --days 60 --write-ch # also writes amir_lab.setter_pnl
```
The ClickHouse password is read from `~/.config/renaissance/clickhouse.env` (`CLICKHOUSE_PASSWORD=`). The script writes only to `amir_lab`. You need `pip install openpyxl clickhouse-connect`. LibreOffice is optional and is used to store calculated values in the file.

`python build_model.py` rebuilds the workbook from scratch. Use it only for structural changes.
