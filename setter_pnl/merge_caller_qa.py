#!/usr/bin/env python3
"""Merge Caller QA grades into the booking data the P&L reads.

Billable rule (Amir, 2026-11-18): where a Caller QA run graded a setter's bookings, the QA
grade decides. Only 'Qualified' bills and pays commission; Review, Not qualified and No
recording do not (Review stays unbilled until someone re-listens). Outside the QA run's
dates, or for setters it did not grade, the Scorecard's 'Min qualified' flag is used.

Inputs are the raw Google Sheets exports (JSON {values: [[...]]}) of:
  --scorecard  Setter Scorecard tabs 'Bookings Oct 2026 FINAL' and 'Bookings Live' (A:Q)
  --qa-funding 'Caller QA, Funding bookings Nov 1-15 2026' › Detail (A:P)
  --qa-data    'Caller QA, Data offer bookings Oct 30-Nov 15 2026' › Detail (A:T)
Output: data/bookings_webform.csv (no prospect names, phones or company names).

Bookings are matched on setter + booked date + company name. QA bookings with no Scorecard
match are kept (partner 'Other (unlisted partners)', priced at the standard rate).
"""
from __future__ import annotations

import argparse
import csv
import json
from collections import Counter
from pathlib import Path

HERE = Path(__file__).resolve().parent
QA_START, QA_END = "2026-11-01", "2026-11-15"
QA_TEAMS = {"Alejandro", "Pablo"}  # in-house data-offer callers (VirtuaVEL is an outside team)
FIELDS = ["source", "booking_id", "booked_date", "setter", "partner", "offer", "verdict",
          "meeting_agreed", "min_qual", "full_qual", "qa_grade"]


def values(path):
    return json.load(open(path))["values"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--scorecard", nargs="+", required=True)
    ap.add_argument("--qa-funding", required=True)
    ap.add_argument("--qa-data", required=True)
    ap.add_argument("--out", default=str(HERE / "data" / "bookings_webform.csv"))
    a = ap.parse_args()

    sc = []
    for p in a.scorecard:
        src = "oct_final" if "oct" in Path(p).name.lower() else "live"
        for r in values(p)[1:]:
            r = r + [""] * 17
            if not r[0] and not r[2]:
                continue
            sc.append(dict(source=src, booking_id=r[0], booked_date=r[1], setter=r[2].strip().lower(),
                           company=r[3].strip().lower(), partner=r[5], offer=r[6], verdict=r[7],
                           meeting_agreed=r[13], min_qual=r[14], full_qual=r[15], qa_grade=""))

    qa = []
    for r in values(a.qa_funding)[1:]:
        if r[0] == "AI phone line" or not r[0]:
            continue
        qa.append(dict(setter=r[0].lower(), date=r[1][:10], company=r[3].strip().lower(), grade=r[5], offer="Funding"))
    for r in values(a.qa_data)[1:]:
        if r[1] in QA_TEAMS and QA_START <= r[2] <= QA_END:
            qa.append(dict(setter=r[0].lower(), date=r[2], company=r[3].strip().lower(), grade=r[6],
                           offer="Data Broker", partner=r[4]))
    graded = {q["setter"] for q in qa}

    index = {}
    for i, s in enumerate(sc):
        index.setdefault((s["setter"], s["booked_date"], s["company"]), []).append(i)
    used, out, stats = set(), [], Counter()
    for n, q in enumerate(qa):
        hits = [i for i in index.get((q["setter"], q["date"], q["company"]), []) if i not in used]
        base = sc[hits[0]] if hits else None
        if hits:
            used.add(hits[0]); stats["matched"] += 1
        else:
            stats["qa_only"] += 1
        verdict = "NO CALL FOUND" if q["grade"] == "No recording" else f"QA {q['grade'].upper()}"
        if base and base["verdict"] == "FALSE BOOKING":
            verdict = "FALSE BOOKING"  # a confirmed false booking stays unbillable
        out.append(dict(source="caller_qa", booking_id=base["booking_id"] if base else f"qa-{n+1}",
                        booked_date=q["date"], setter=q["setter"],
                        partner=(base["partner"] if base else q.get("partner", "Other (unlisted partners)")),
                        offer=q["offer"], verdict=verdict,
                        meeting_agreed=base["meeting_agreed"] if base else "",
                        min_qual="Yes" if q["grade"] == "Qualified" else "No",
                        full_qual=base["full_qual"] if base else "", qa_grade=q["grade"]))
    for i, s in enumerate(sc):
        if i in used:
            continue
        if s["setter"] in graded and QA_START <= s["booked_date"] <= QA_END:
            stats["scorecard_dropped_in_qa_window"] += 1  # graded set is authoritative here
            continue
        out.append({k: s.get(k, "") for k in FIELDS})
    out.sort(key=lambda r: (r["booked_date"], r["setter"], str(r["booking_id"])))
    with open(a.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS); w.writeheader(); w.writerows(out)
    print(f"wrote {len(out)} rows to {a.out}: {dict(stats)}")


if __name__ == "__main__":
    main()
