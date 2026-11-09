# Business Lending TAM workbook

`build_tam.py` builds `output/TAM_Business_Lending_<date>.xlsx`: every Business Lending company (MCA funders, direct lenders, ISOs, brokers) we hold, cleaned, de-duplicated, classified (Funder / Broker/ISO / Hybrid), sized, and ranked by an editable, formula-driven Fit Score.

## Run weekly

```bash
pip install clickhouse-connect rapidfuzz openpyxl pandas
python tam/build_tam.py                     # live: reads CLICKHOUSE_PASSWORD from ~/.config/renaissance/clickhouse.env
python tam/build_tam.py --source snapshot   # rebuild from tam/data/raw/*.json without DB access
```

Live mode stages `amir_lab.tam_universe`, `tam_rollup`, `tam_site`, `tam_verify`, `tam_contacts`, `tam_engagement` (it only writes to `amir_lab`), then saves JSON snapshots to `tam/data/raw/`.

## Inputs
- `tam/input/attio_companies.xlsx`, `tam/input/attio_people.csv`: Attio exports (git-ignored)
- `tam/data/partners.csv`: Attio Partners list (update when Partners change)
- `tam/data/partner_contacts.csv`: Partner main contacts (git-ignored, contains emails)
- `tam/data/extra_domains.txt`: Attio keyword-hit + Partner domains fed into the universe

## Files
- `sql/00_universe.sql` … `06_partner_meetings.sql`: one query per stage
- `classify.py`: every keyword rule (scope, segment, sub-type, size bands, seniority). Edit rules here.
- `build_tam.py`: loading, dedupe, scoring inputs, workbook writer, recalculation, console summary
- `last_build_summary.txt`: console summary from the last build

Contact data, raw snapshots, and the workbook are git-ignored (PII).
