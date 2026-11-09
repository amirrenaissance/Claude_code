#!/usr/bin/env python3
"""Build the Renaissance Business Lending TAM workbook.

    python tam/build_tam.py                 # live: query ClickHouse, stage amir_lab.tam_*, build
    python tam/build_tam.py --source snapshot   # rebuild from tam/data/raw/*.json (no DB access)

Live mode reads CLICKHOUSE_PASSWORD from ~/.config/renaissance/clickhouse.env.
The password is never printed, logged, or written to any output.
Writes only to amir_lab (tables prefixed tam_). Never writes to shared schemas.

Inputs that are not in the warehouse:
  tam/input/attio_companies.xlsx   Attio "Companies - All Companies" export
  tam/input/attio_people.csv       Attio "People - Recently Contacted" export
  tam/data/partners.csv            Attio Partners list (signed customers), curated
  tam/data/partner_contacts.csv    Partner main contacts from Attio (git-ignored: contains emails)
  tam/data/partner_meetings.json   meetings per partner key (snapshot mode only)
"""
from __future__ import annotations

import argparse
import datetime as dt
import glob
import json
import math
import os
import re
import subprocess
import sys
from collections import Counter
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
import classify as C  # noqa: E402

SQL = ROOT / "sql"
RAW = ROOT / "data" / "raw"
INPUT = ROOT / "input"
OUT = ROOT / "output"
TODAY = dt.date.today()

PERSONAL = set("""gmail.com yahoo.com hotmail.com aol.com comcast.net msn.com sbcglobal.net att.net bellsouth.net
verizon.net outlook.com live.com icloud.com cox.net me.com mac.com ymail.com protonmail.com charter.net earthlink.net
optonline.net rocketmail.com googlemail.com mail.com gmx.com""".split())
LENDING_NAME_RX = r"fund|capital|financ|lend|loan|credit|advance|factor|leas|\bcap\b|money|merchant"


# ============================================================ data loading
def _read_pages(pattern: str) -> pd.DataFrame:
    rows = []
    for f in sorted(glob.glob(str(RAW / pattern))):
        rows += json.load(open(f))["result"]["data"]
    return pd.DataFrame(rows)


def load_snapshot() -> dict[str, pd.DataFrame]:
    d = {
        "rollup": _read_pages("rollup_p*.json"),
        "site": _read_pages("site_p*.json").rename(columns={"d": "domain"}),
        "verify": _read_pages("verify_p*.json").rename(columns={"U.d": "domain", "d": "domain"}),
        "contacts": pd.concat([_read_pages("contacts_p*.json"), _read_pages("contacts_lo_p*.json")], ignore_index=True),
        "engagement": _read_pages("engagement_p*.json").rename(columns={"d": "domain"}),
    }
    pm = json.load(open(ROOT / "data" / "partner_meetings.json"))
    d["partner_meetings"] = pd.DataFrame([{"pk": k, "meetings": v} for k, v in pm.items()])
    return d


def _ch_client():
    import clickhouse_connect  # pip install clickhouse-connect

    env = Path.home() / ".config" / "renaissance" / "clickhouse.env"
    pw = None
    for line in env.read_text().splitlines():
        if "=" in line and line.split("=", 1)[0].strip() in ("CLICKHOUSE_PASSWORD", "PASSWORD"):
            pw = line.split("=", 1)[1].strip().strip('"').strip("'")
    if not pw:
        raise SystemExit(f"CLICKHOUSE_PASSWORD not found in {env}")
    return clickhouse_connect.get_client(host="uryfmx4y96.us-east-1.aws.clickhouse.cloud", port=8443, secure=True,
                                         username="amir", password=pw, database="amir_lab",
                                         send_receive_timeout=1800)


def _sql(name: str, **params) -> str:
    text = (SQL / name).read_text()
    for k, v in params.items():
        text = text.replace("{" + k + "}", str(v))
    return text


def load_live(extra_domains: list[str], per_domain: int) -> dict[str, pd.DataFrame]:
    client = _ch_client()
    arr = "[" + ",".join("'" + x.replace("'", "") + "'" for x in extra_domains) + "]::Array(String)"
    uni = "amir_lab.tam_universe"

    def stage(table: str, sql: str):
        assert table.startswith("tam_"), "staging tables must be prefixed tam_"
        client.command(f"CREATE OR REPLACE TABLE amir_lab.{table} ENGINE = MergeTree ORDER BY tuple() AS {sql}")
        df = client.query_df(f"SELECT * FROM amir_lab.{table}")
        RAW.mkdir(parents=True, exist_ok=True)
        df.to_json(RAW / f"live_{table}.json", orient="records", date_format="iso")
        print(f"  staged amir_lab.{table}: {len(df):,} rows")
        return df

    print("Querying ClickHouse (writes only to amir_lab.tam_*) ...")
    stage("tam_universe", _sql("00_universe.sql", extra_domains=arr))
    d = {
        "rollup": stage("tam_rollup", _sql("01_rollup.sql", universe=uni)),
        "site": stage("tam_site", _sql("02_site.sql", universe=uni)).rename(columns={"d": "domain"}),
        "verify": stage("tam_verify", _sql("03_verify.sql", universe=uni)).rename(columns={"U.d": "domain", "d": "domain"}),
        "contacts": stage("tam_contacts", _sql("04_contacts.sql", universe=uni, per_domain=per_domain)),
        "engagement": stage("tam_engagement", _sql("05_engagement.sql", universe=uni)).rename(columns={"d": "domain"}),
        "partner_meetings": client.query_df(_sql("06_partner_meetings.sql")),
    }
    return d


def load_attio():
    comp = pd.read_excel(INPUT / "attio_companies.xlsx")
    people = pd.read_csv(INPUT / "attio_people.csv", low_memory=False)
    partners = pd.read_csv(ROOT / "data" / "partners.csv", dtype=str).fillna("")
    return comp, people, partners


def norm_domain(x) -> str:
    x = str(x or "").strip().lower()
    x = re.sub(r"^https?://", "", x)
    x = re.sub(r"^www\.", "", x).split("/")[0].strip(".")
    return x


NAME_SUFFIX = r"\b(llc|l\.l\.c|inc|incorporated|corp|corporation|co|company|ltd|lp|llp|pllc|group|the|usa|us|of america)\b"


def norm_name(x) -> str:
    x = str(x or "").lower().replace("&", " and ")
    x = re.sub(r"[^a-z0-9 ]", " ", x)
    x = re.sub(NAME_SUFFIX, " ", x)
    return re.sub(r"\s+", " ", x).strip()


def num(x):
    try:
        v = float(x)
        return v if v == v else None
    except (TypeError, ValueError):
        return None


# ============================================================ universe + cleaning
def build_universe(src, attio_comp, partners, extra_domains):
    dq = {"merged": [], "excluded": [], "notes": []}
    rollup, site, verify = src["rollup"].copy(), src["site"].copy(), src["verify"].copy()
    for df in (rollup, site, verify):
        df["domain"] = df["domain"].map(norm_domain)
    rollup = rollup.drop_duplicates("domain")
    site = site.drop_duplicates("domain")
    verify = verify.drop_duplicates("domain")

    # Attio companies: explode multi-domain records
    a = attio_comp.copy()
    a["dom_list"] = a["Domains"].fillna("").astype(str).str.split(",")
    a = a.explode("dom_list")
    a["domain"] = a["dom_list"].map(norm_domain)
    a = a[a.domain != ""]
    attio_by_domain = a.drop_duplicates("domain").set_index("domain")

    partner_domains = {}
    for _, p in partners.iterrows():
        for dmn in p["domains"].split("|"):
            if dmn:
                partner_domains[norm_domain(dmn)] = p

    doms = set(rollup.domain) | set(site.domain) | set(verify.domain) | set(map(norm_domain, extra_domains)) \
        | {d for d, p in partner_domains.items() if p["business_lending"] == "Y"}
    doms = {d for d in doms if d and "." in d}
    u = pd.DataFrame({"domain": sorted(doms)})
    u = u.merge(rollup, on="domain", how="left").merge(site, on="domain", how="left", suffixes=("_lc", "_site")) \
         .merge(verify, on="domain", how="left")
    u["attio_id"] = u.domain.map(lambda d: attio_by_domain["Record ID"].get(d, ""))
    u["attio_name"] = u.domain.map(lambda d: attio_by_domain["record"].get(d, ""))
    u["attio_conn"] = u.domain.map(lambda d: attio_by_domain["Connection strength"].get(d, ""))
    u["attio_last"] = u.domain.map(lambda d: attio_by_domain["Last interaction > When"].get(d, ""))
    for c in ("attio_id", "attio_name", "attio_conn", "attio_last"):
        u[c] = u[c].fillna("").astype(str).replace({"nan": "", "NaT": ""})
    u["is_partner"] = u.domain.isin(partner_domains)
    u["partner_name"] = u.domain.map(lambda d: partner_domains[d]["partner_name"] if d in partner_domains else "")

    u["in_vertical"] = u.pv.fillna("").eq("mca_alt_lending")
    u["in_site_profile"] = u.si.notna() & u.si.fillna("").ne("")
    u["in_leads_tag"] = u.il.astype(str).isin(["1", "True", "true"])
    u["in_attio"] = u.attio_id.ne("")
    u["n"] = pd.to_numeric(u.n, errors="coerce").fillna(0).astype(int)
    for c in ("nv", "nx", "ns"):
        u[c] = pd.to_numeric(u[c], errors="coerce").fillna(0).astype(int)

    def best_name(r):
        for v in (r.partner_name, r.attio_name, r.cn_site, r.cn_lc):
            if isinstance(v, str) and v.strip() and v.strip().lower() not in ("nan", "self-employed"):
                return v.strip()
        return r.domain
    u["company"] = u.apply(best_name, axis=1)

    def sources(r):
        s = []
        if r.in_vertical: s.append("WH:vertical_tag")
        if r.in_site_profile: s.append("WH:website_profile")
        if r.in_leads_tag: s.append("WH:contact_industry")
        if r.in_attio: s.append("Attio")
        if r.is_partner: s.append("Attio:Partners")
        return "; ".join(s) or "Attio keyword list"
    u["source"] = u.apply(sources, axis=1)
    u["text"] = (u.company + " | " + u.cn_lc.fillna("").astype(str) + " | " + u.si.fillna("").astype(str) + " | "
                 + u.ol.fillna("").astype(str) + " | " + u.domain)

    # ---------- scope / exclusion
    keep = []
    for _, r in u.iterrows():
        reason, hit = (None, None)
        if r.is_partner:
            keep.append((None, None)); continue
        if r.domain in PERSONAL:
            reason, hit = "Personal email domain", r.domain
        elif r.n > 5000:
            reason, hit = "Generic / very large domain (>5,000 contacts)", f"{r.n:,} contacts"
        elif re.search(r"\.(edu|gov|mil)$|\.(edu|gov|ac)\.[a-z]{2}$", r.domain):
            reason, hit = "Education / government domain", r.domain
        elif re.search(r"\.(co\.za|co\.uk|com\.au|ca|in|ng|ph|pk|ae|co\.in|com\.ng|com\.ph|de|fr|es|it|nl|br|mx)$", r.domain):
            reason, hit = "Non-US domain", r.domain
        else:
            reason, hit = C.scope(r.text)
            if not reason:
                name_ev = re.search(LENDING_NAME_RX, f"{r.company} {r.domain}", re.I)
                site_txt = f"{r.si if isinstance(r.si, str) else ''} {r.ol if isinstance(r.ol, str) else ''}"
                if r.in_site_profile:
                    # the website profile overrides tags: it must describe business lending/funding
                    if not re.search(C.FUNDER_RX + "|" + C.BROKER_RX + r"|merchant|funding|capital", site_txt, re.I):
                        reason, hit = "Website profile does not describe business lending", str(r.si)
                elif not name_ev and not r.in_attio:
                    reason, hit = "Weak signal: tag only, no lending name or website evidence", ""
        keep.append((reason, hit))
    u["excl_reason"] = [k[0] for k in keep]
    u["excl_hit"] = [k[1] for k in keep]
    excluded = u[u.excl_reason.notna()][["domain", "company", "source", "excl_reason", "excl_hit", "si", "n"]].copy()
    u = u[u.excl_reason.isna()].copy()

    # ---------- dedupe: domain is already unique; fuzzy name second
    u["name_key"] = u.company.map(norm_name)
    u["rank"] = u.is_partner.astype(int) * 10**7 + u.in_attio.astype(int) * 10**6 + u.n
    u = u.sort_values("rank", ascending=False)
    groups = {}
    try:
        from rapidfuzz import fuzz
    except ImportError:
        fuzz = None
    generic = {"capital", "funding", "financial", "finance", "business funding", "capital funding", "lending",
               "business capital", "financial services", "advance", "credit"}
    # domains listed for one Partner are the same company (Attio multi-domain records are not trusted: they mix in sending/unrelated domains)
    grp_of = {}
    for _, p in partners.iterrows():
        for dmn in p["domains"].split("|"):
            if dmn:
                grp_of[norm_domain(dmn)] = "partner:" + p["partner_name"]
    grp_seen = {}
    canon = {}
    for idx, r in u.iterrows():
        g = grp_of.get(r.domain)
        if g:
            if g in grp_seen:
                canon[idx] = grp_seen[g]; continue
            grp_seen[g] = idx
        k = r.name_key
        if len(k) < 6 or k in generic:
            canon[idx] = idx; continue
        target = groups.get(k)
        if target is None and fuzz is not None:
            block = [kk for kk in groups if kk[:4] == k[:4]]
            for kk in block:
                if fuzz.token_sort_ratio(k, kk) >= 95:
                    target = groups[kk]; break
        if target is None:
            groups[k] = idx; canon[idx] = idx
        else:
            canon[idx] = target
    u["canon"] = pd.Series(canon)
    merged = u[u.index != u.canon]
    alt = merged.groupby("canon").domain.apply(lambda s: ", ".join(sorted(s)))
    for _, r in merged.iterrows():
        keep_row = u.loc[r.canon]
        dq["merged"].append({"kept_domain": keep_row.domain, "kept_company": keep_row.company,
                             "merged_domain": r.domain, "merged_company": r.company,
                             "match": ("same Partner / Attio record" if grp_of.get(r.domain) and grp_of.get(r.domain) == grp_of.get(keep_row.domain)
                                       else "exact normalized name" if norm_name(r.company) == keep_row.name_key else "fuzzy name ≥95")})
    # roll merged contact counts into the kept row
    for col in ("n", "nv", "nx", "ns"):
        u[col] = u.groupby("canon")[col].transform("sum").where(u.index == u.canon, u[col])
    u["alt_domains"] = u.index.map(lambda i: alt.get(i, ""))
    u["merged_into"] = u.canon.map(lambda c: u.loc[c, "domain"])
    domain_to_primary = dict(zip(u.domain, u.merged_into))
    u = u[u.index == u.canon].copy()
    return u, excluded, dq, domain_to_primary


# ============================================================ classification, sizing, signals
def classify_companies(u, partners, src):
    seg, segr, sub, subr, conf = [], [], [], [], []
    for _, r in u.iterrows():
        s, sr = C.segment(r.text, r.company)
        st, sh = C.subtype(r.text)
        if s == "Hybrid" and st == "Marketplace / broker":
            st2, sh2 = C.subtype(re.sub(r"marketplace|brokerage|\bbroker\b|\biso\b", "", r.text, flags=re.I))
            if st2 != "Unspecified":
                st, sh = st2, sh2
        if s == "Unknown" and r.in_vertical:
            s, sr = "Funder", "vertical tag mca_alt_lending (no segment keywords)"
            if st == "Unspecified":
                st, sh = "MCA funder", "vertical tag"
        has_text = bool(str(r.si or "").strip() or str(r.ol or "").strip())
        c = "High" if (has_text and s != "Unknown" and st != "Unspecified") else ("Medium" if s != "Unknown" else "Low")
        if "name kw" in sr or "vertical tag" in sr:
            c = "Low" if not has_text else "Medium"
        seg.append(s); segr.append(sr); sub.append(st); subr.append(f'sub-type kw "{sh}"' if sh else "no sub-type kw")
        conf.append(c)
    u["segment"], u["sub_type"], u["class_conf"] = seg, sub, conf
    u["classification_reason"] = [f"{a}; {b}" for a, b in zip(segr, subr)]

    # partners: segment overrides from what we know first-hand
    p_seg = {  # from the Partners' Attio descriptions (first-hand)
        "capfrontfunding.com": ("Broker/ISO", "Marketplace / broker", "Attio: 'marketplace connecting merchants and lenders'"),
        "llamaloan.com": ("Broker/ISO", "Marketplace / broker", "Attio: 'compare multiple lenders ... through a single application'"),
        "capitalgurus.com": ("Broker/ISO", "LOC / working capital", "website: small business lending brokerage"),
        "clarifycapital.com": ("Funder", "Term lender", "Attio: 'business loans ... up to $5 million'"),
        "goswyft.com": ("Funder", "MCA funder", "Attio: 'capital solutions ... manage MCA obligations'"),
        "greenbridgecapital.com": ("Funder", "Commercial lending (general)", "Partner: anchor funding partner (rev share)"),
        "1westfinance.com": ("Funder", "Commercial lending (general)", "Attio: 'wide array of small business financing options'"),
        "1west.com": ("Funder", "Commercial lending (general)", "Attio: 'wide array of small business financing options'"),
        "goqualifi.com": ("Funder", "LOC / working capital", "Attio: 'lines of credit up to $75 million'"),
    }
    for d, (s, st, why) in p_seg.items():
        u.loc[u.domain == d, ["segment", "sub_type", "classification_reason", "class_conf"]] = [s, st, f"Partner override: {why}", "High"]

    # ---------- size
    p_emp = {}
    for _, p in partners.iterrows():
        for d in p["domains"].split("|"):
            if p["employee_range"]:
                p_emp[norm_domain(d)] = p["employee_range"]
    emps, bands, basis = [], [], []
    for _, r in u.iterrows():
        e_site = C.parse_employees(r.es)
        e_partner = C.parse_employees(p_emp.get(r.domain))
        e_lc = C.parse_employees(r.sb) if str(r.sb or "") not in ("", "nan") else None
        emp, b = None, ""
        for v, label in ((e_partner, "Attio employee range"), (e_site, "website stated staff"), (e_lc, "contact DB size bucket")):
            if v:
                emp, b = v, label; break
        rev = num(r.ar)
        band, why = C.size_band(emp, rev)
        emps.append(emp); bands.append(band); basis.append(f"{b}: {why}" if emp else why)
    u["employees_est"], u["size_band"], u["size_basis"] = emps, bands, basis
    u["years_in_business"] = u.apply(
        lambda r: (TODAY.year - int(num(r.yf))) if num(r.yf) and 1900 < num(r.yf) <= TODAY.year
        else (round(num(r.age) / 12, 1) if num(r.age) else None), axis=1)

    # ---------- verification
    def verify(r):
        ss = r.ss if isinstance(r.ss, str) else ""
        mx = r.mx if isinstance(r.mx, str) else ""
        if str(r.dead) in ("1", "1.0") or ss in ("parked", "dns_dead") or mx in ("nxdomain", "null_mx"):
            return "Dead / parked"
        if ss == "" and mx == "":
            return "Not checked"
        ok_site, ok_mx = ss == "loads", mx in ("resolved", "spf_only")
        if ok_site and ok_mx: return "Verified active (site + mail)"
        if ok_site or ok_mx: return "Likely active (site or mail)"
        return "Unverified (site/mail errors)"
    u["verification"] = u.apply(verify, axis=1)

    # ---------- engagement
    eng = src["engagement"].copy()
    eng["domain"] = eng["domain"].map(norm_domain)
    eng = eng.groupby("domain").agg(replies=("n", lambda s: pd.to_numeric(s).sum()),
                                    pos=("pos", lambda s: pd.to_numeric(s).sum()), last=("last", "max"))
    u = u.merge(eng, left_on="domain", right_index=True, how="left")
    u["replies"] = u.replies.fillna(0).astype(int)
    u["pos"] = u.pos.fillna(0).astype(int)
    return u


def build_contacts(src, u, domain_to_primary, partners, attio_people):
    c = src["contacts"].copy()
    c["d"] = c["d"].map(norm_domain)
    c["domain"] = c.d.map(lambda d: domain_to_primary.get(d, d))
    c = c[c.domain.isin(set(u.domain))].copy()
    c = c.drop_duplicates("e")
    c["source"] = "Warehouse (mirror.leads_current)"
    # Partner main contacts recorded in Attio (exact text; nothing inferred)
    extra = []
    pc_file = ROOT / "data" / "partner_contacts.csv"  # main contacts copied from the Attio Partners list (git-ignored)
    pcs = pd.read_csv(pc_file, dtype=str).fillna("") if pc_file.exists() else pd.DataFrame(columns=["partner_name", "name", "title", "email"])
    for _, pc in pcs.iterrows():
        name, title, email = pc["name"], pc["title"], pc["email"]
        dom = domain_to_primary.get(norm_domain(email.split("@")[1]), norm_domain(email.split("@")[1]))
        if email.lower() in set(c.e):
            continue
        fn, _, ln = name.partition(" ")
        extra.append({"d": dom, "domain": dom, "e": email.lower(), "fn": fn, "ln": ln, "t": title or "Main contact (Attio)",
                      "sr": "", "vs": "", "ph": "", "li": "", "source": "Attio Partners list (main_contact)"})
    if extra:
        c = pd.concat([c, pd.DataFrame(extra)], ignore_index=True)
    for col in ("fn", "ln", "t", "sr", "vs", "ph", "li"):
        c[col] = c[col].fillna("").astype(str).replace({"nan": "", "None": ""})
    c["seniority"] = [C.seniority(t, s) for t, s in zip(c.t, c.sr)]
    c["is_sales_role"] = c.t.str.contains(C.SALES_DM_RX, case=False, regex=True)
    vs_map = {"valid": "Verified valid", "catch_all_valid": "Catch-all (accepts mail)", "catch_all": "Catch-all (unconfirmed)",
              "catch_all_invalid": "Catch-all (likely invalid)", "not_valid": "Invalid - do not email", "unknown": "Unknown", "": "Not verified"}
    c["email_status"] = c.vs.map(lambda v: vs_map.get(v, v))
    c["linkedin"] = c.li.map(lambda s: "" if not s else (s if s.startswith("http") else "https://www.linkedin.com/in/" + s))
    # Attio people (name only in the export) -> flag possible match, never used to fill fields
    ap = attio_people.copy()
    ap["k"] = ap["Record"].fillna("").astype(str).str.lower().str.strip()
    ap = ap[ap.k.str.contains(" ")].drop_duplicates("k").set_index("k")
    c["full"] = (c.fn + " " + c.ln).str.strip()
    k = c.full.str.lower()
    c["attio_match"] = k.map(lambda x: "Name match: " + str(ap["Connection strength"].get(x, "")) if x in ap.index else "")

    srank = {"Owner/C-level": 1, "VP/Director": 2, "Manager": 3, "IC": 4}
    c["_r"] = c.seniority.map(srank)
    c["_v"] = c.vs.isin(["valid", "catch_all_valid"]).astype(int)
    c["_p"] = c.ph.ne("").astype(int)
    c = c.sort_values(["domain", "_r", "_v", "_p", "is_sales_role"], ascending=[True, True, False, False, False])
    c["dm_rank"] = ""
    dm = c[c.seniority.isin(["Owner/C-level", "VP/Director"]) & c.vs.ne("not_valid")]
    ranks = dm.groupby("domain").cumcount() + 1
    c.loc[ranks.index, "dm_rank"] = ranks.where(ranks <= 3).dropna().astype(int).astype(str).reindex(ranks.index).fillna("")
    c["best_dm"] = c.dm_rank.map(lambda x: f"Yes (#{x})" if x else "")
    return c


def score_inputs(u, contacts, partners):
    # contact coverage per company
    g = contacts.groupby("domain")
    has_dm_verified = g.apply(lambda df: ((df.dm_rank != "") & (df._v == 1)).any())
    has_dm = g.apply(lambda df: (df.dm_rank != "").any())
    has_vp = g.apply(lambda df: df.seniority.eq("VP/Director").any())
    n_c = g.size()
    u["contacts_on_file"] = u.domain.map(n_c).fillna(0).astype(int)
    u["dm_on_file"] = u.domain.map(has_dm).fillna(False).astype(bool)
    u["dm_verified"] = u.domain.map(has_dm_verified).fillna(False).astype(bool)

    def coverage(r):
        if r.dm_verified: return 1.0, "verified decision-maker on file"
        if r.dm_on_file: return 0.7, "decision-maker on file (email unverified)"
        if has_vp.get(r.domain, False): return 0.5, "VP/Director on file"
        if r.contacts_on_file > 0: return 0.25, "contacts on file, no decision-maker"
        if r.nx > 0: return 0.25, "exec contacts in warehouse (not pulled)"
        return 0.0, "no contacts"
    cov = u.apply(coverage, axis=1)
    u["s_coverage"], u["why_coverage"] = [x[0] for x in cov], [x[1] for x in cov]

    # outbound / lead-buying signal
    def outbound(r):
        s, why = 0.0, []
        if r.ns > 0: s += 0.4; why.append(f"{r.ns} sales/ISO/BD staff")
        if str(r.hi) in ("1", "True", "true"): s += 0.2; why.append("hiring")
        if r.segment in ("Broker/ISO", "Hybrid"): s += 0.2; why.append("broker/ISO model buys leads")
        if r.sub_type in ("MCA funder", "LOC / working capital", "Marketplace / broker"): s += 0.2; why.append(f"{r.sub_type} relies on volume")
        if r.n >= 10: s += 0.2; why.append(f"{r.n} staff in contact DB")
        return min(s, 1.0), ", ".join(why) or "no outbound signal"
    ob = u.apply(outbound, axis=1)
    u["s_outbound"], u["why_outbound"] = [x[0] for x in ob], [x[1] for x in ob]

    # engagement with us
    cs = {"Very strong": 1.0, "Strong": 0.9, "Good": 0.7, "Weak": 0.4, "Very weak": 0.2}
    def engagement(r):
        best, why = 0.0, []
        if r.attio_conn in cs: best = max(best, cs[r.attio_conn]); why.append(f"Attio connection {r.attio_conn}")
        if r.pos > 0: best = max(best, 0.6); why.append(f"{r.pos} positive replies")
        elif r.replies > 0: best = max(best, 0.3); why.append(f"{r.replies} replies")
        if r.attio_last and r.attio_last[:4].isdigit():
            try:
                days = (TODAY - pd.to_datetime(r.attio_last).date()).days
                if days <= 90: best = min(1.0, best + 0.2); why.append(f"Attio touch {days}d ago")
            except Exception:
                pass
        return best, ", ".join(why) or "no engagement"
    en = u.apply(engagement, axis=1)
    u["s_engagement"], u["why_engagement"] = [x[0] for x in en], [x[1] for x in en]

    # lookalike vs lending Partners
    u["yib_bucket"] = pd.cut(u.years_in_business.astype(float), [-1, 3, 7, 12, 200], labels=["0-3", "4-7", "8-12", "13+"]).astype(str)
    P = u[u.is_partner]
    def share(col, val):
        vals = P[col].dropna()
        vals = vals[vals.astype(str).ne("Unknown") & vals.astype(str).ne("Unspecified")]
        if len(vals) == 0: return 0.0
        cnt = Counter(vals.astype(str)); top = max(cnt.values())
        return cnt.get(str(val), 0) / top
    st_share = {k: v for k, v in Counter(P.st.dropna().astype(str)).items() if k not in ("", "nan")}
    st_top = max(st_share.values()) if st_share else 1
    def lookalike(r):
        parts = [share("segment", r.segment), share("sub_type", r.sub_type), share("size_band", r.size_band),
                 share("yib_bucket", r.yib_bucket), st_share.get(str(r.st), 0) / st_top]
        return round(sum(parts) / len(parts), 3)
    u["s_lookalike"] = u.apply(lookalike, axis=1)

    # disqualification (formula tier uses this column)
    def disq(r):
        if r.is_partner: return "Partner"
        if re.search(C.COMPETITOR_RX, r.text, re.I): return "Disqualified - Competitor"
        if r.verification == "Dead / parked": return "Disqualified - Out of business"
        if r.contacts_on_file == 0 and r.n == 0: return "Disqualified - No contacts"
        return ""
    u["disq"] = u.apply(disq, axis=1)
    return u


# ============================================================ workbook
WEIGHTS = [  # key, label, default weight, description
    ("s_lookalike", "Lookalike to Partners", 25, "Mean share of lending Partners matching this company's segment, sub-type, size band, years-in-business bucket, and HQ state (0-1)."),
    ("seg_fit", "Segment fit", 20, "Looked up from the Segment fit table below."),
    ("sub_fit", "Sub-type fit", 10, "Looked up from the Sub-type fit table below."),
    ("size_fit", "Size band fit", 15, "Looked up from the Size band fit table below."),
    ("s_outbound", "Buys leads / runs outbound", 10, "Sales/ISO/BD staff on file (+0.4), hiring (+0.2), broker/ISO model (+0.2), volume sub-type (+0.2), 10+ staff (+0.2); capped at 1."),
    ("s_engagement", "Engagement with us", 10, "Attio connection strength (Very strong 1.0 … Very weak 0.2), positive replies (0.6) or any reply (0.3), +0.2 if Attio touch in last 90 days; max, capped at 1."),
    ("s_coverage", "Decision-maker coverage", 10, "Verified decision-maker 1.0; unverified DM 0.7; VP/Director 0.5; any contact 0.25; none 0."),
]
SEG_FIT = [("Funder", 1.0), ("Hybrid", 0.9), ("Broker/ISO", 0.75), ("Unknown", 0.3)]
SUB_FIT = [("MCA funder", 1.0), ("LOC / working capital", 1.0), ("Marketplace / broker", 0.85), ("Commercial lending (general)", 0.8),
           ("Term lender", 0.75), ("Revenue-based", 0.75), ("SBA lender", 0.6), ("Equipment finance", 0.55), ("Factoring", 0.55),
           ("Asset-based / ABL", 0.5), ("CDFI / community lender", 0.3), ("Unspecified", 0.4)]
SIZE_FIT = [("Micro", 0.6), ("Small", 1.0), ("Mid", 0.9), ("Large", 0.6), ("Enterprise", 0.3), ("Unknown", 0.5)]
TIER_CUTS = [("Tier 1 minimum Fit Score", 70), ("Tier 2 minimum Fit Score", 55)]


def compute_python_score(u):
    """Mirror of the Excel formula, used only to sort rows and print the summary."""
    seg = dict(SEG_FIT); sub = dict(SUB_FIT); size = dict(SIZE_FIT)
    u["seg_fit"] = u.segment.map(seg).fillna(0)
    u["sub_fit"] = u.sub_type.map(sub).fillna(0)
    u["size_fit"] = u.size_band.map(size).fillna(0)
    tot = sum(w for _, _, w, _ in WEIGHTS)
    raw = 100 * sum(u[k] * w for k, _, w, _ in WEIGHTS) / tot
    u["fit_py"] = (raw * 10 + 0.5 + 1e-9).apply(math.floor) / 10  # Excel ROUND is half-up
    t1, t2 = TIER_CUTS[0][1], TIER_CUTS[1][1]
    u["tier_py"] = u.apply(lambda r: r.disq if r.disq else ("Tier 1" if r.fit_py >= t1 else "Tier 2" if r.fit_py >= t2 else "Tier 3"), axis=1)

    def reason(r):
        contrib = sorted([(r[k] * w, lab, k) for k, lab, w, _ in WEIGHTS], reverse=True)
        bits = []
        for v, lab, k in contrib[:3]:
            if v <= 0: continue
            extra = {"s_outbound": r.why_outbound, "s_engagement": r.why_engagement, "s_coverage": r.why_coverage,
                     "seg_fit": r.segment, "sub_fit": r.sub_type, "size_fit": r.size_band,
                     "s_lookalike": f"{r.s_lookalike:.2f} match"}[k]
            bits.append(f"{lab}: {extra}")
        weak = [lab for v, lab, k in contrib if v == 0][:2]
        s = "; ".join(bits)
        if weak: s += " | gaps: " + ", ".join(weak)
        return s
    u["score_reason"] = u.apply(reason, axis=1)
    return u


def write_workbook(u, contacts, excluded, dq, partners, src, out_path, mode):
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    F = "Arial"
    H = Font(name=F, bold=True, color="FFFFFF"); HF = PatternFill("solid", fgColor="1F3864")
    B = Font(name=F); BB = Font(name=F, bold=True); TITLE = Font(name=F, bold=True, size=14)
    INP = Font(name=F, color="0000FF"); YEL = PatternFill("solid", fgColor="FFFF00")
    LINK = Font(name=F, color="008000")
    thin = Side(style="thin", color="BFBFBF")

    wb = Workbook()
    ws_readme = wb.active; ws_readme.title = "README"
    names = ["Summary", "Companies", "Contacts", "Partners_Profile", "Scoring_Weights", "Assumptions", "Data_Quality", "Excluded"]
    S = {n: wb.create_sheet(n) for n in names}

    def header(ws, row, cols, widths=None):
        for i, c in enumerate(cols, 1):
            cell = ws.cell(row=row, column=i, value=c); cell.font = H; cell.fill = HF
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        if widths:
            for i, w in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(i)].width = w

    def clean(v):
        """Data values must never be read as formulas or error codes (source has names stored as '=#NAME?')."""
        if isinstance(v, str) and re.fullmatch(r"=?#(NAME\?|N/A|VALUE!|REF!|DIV/0!|NULL!|NUM!)", v.strip()):
            return None
        if isinstance(v, str) and v[:1] in ("=", "+", "@", "-") and len(v) > 0 and not v[1:2].isdigit():
            return v.lstrip("=+@- ") or None
        return v

    def put(ws, r, c, v, font=None):
        if isinstance(v, float) and v != v: v = None
        if isinstance(v, str) and v in ("nan", "None", "NaT"): v = None
        cell = ws.cell(row=r, column=c, value=v); cell.font = font or B
        return cell

    # ---------------- Scoring_Weights
    sw = S["Scoring_Weights"]
    sw["A1"] = "Scoring weights and tier cutoffs (edit the blue cells; the Companies tab recalculates)"; sw["A1"].font = TITLE
    sw["A2"] = "Legend: blue text on yellow = editable input. Fit Score = 100 × Σ(component × weight) ÷ Σ(weights). Components are 0-1."
    sw["A2"].font = B
    header(sw, 4, ["Key", "Component", "Weight", "How the component is computed"], [16, 30, 10, 110])
    wref = {}
    for i, (k, lab, w, desc) in enumerate(WEIGHTS):
        r = 5 + i
        put(sw, r, 1, k); put(sw, r, 2, lab)
        c = put(sw, r, 3, w, INP); c.fill = YEL
        put(sw, r, 4, desc)
        wref[k] = f"Scoring_Weights!$C${r}"
    wsum_row = 5 + len(WEIGHTS)
    put(sw, wsum_row, 2, "Total weight", BB); put(sw, wsum_row, 3, f"=SUM(C5:C{wsum_row - 1})", BB)
    WSUM = f"Scoring_Weights!$C${wsum_row}"

    r0 = wsum_row + 2
    put(sw, r0, 1, "Tier cutoffs", BB)
    cut_ref = []
    for i, (lab, v) in enumerate(TIER_CUTS):
        put(sw, r0 + 1 + i, 2, lab); c = put(sw, r0 + 1 + i, 3, v, INP); c.fill = YEL
        cut_ref.append(f"Scoring_Weights!$C${r0 + 1 + i}")
    put(sw, r0 + 3, 2, "Disqualified / Partner rows keep their label regardless of score.")

    def lookup_table(start, title, rows):
        put(sw, start, 1, title, BB)
        header(sw, start + 1, ["", "Value", "Fit (0-1)"])
        for j, (k, v) in enumerate(rows):
            put(sw, start + 2 + j, 2, k); c = put(sw, start + 2 + j, 3, v, INP); c.fill = YEL
        a, b = start + 2, start + 1 + len(rows)
        return f"Scoring_Weights!$B${a}:$B${b}", f"Scoring_Weights!$C${a}:$C${b}", b
    seg_k, seg_v, e1 = lookup_table(r0 + 5, "Segment fit", SEG_FIT)
    sub_k, sub_v, e2 = lookup_table(e1 + 2, "Sub-type fit", SUB_FIT)
    siz_k, siz_v, e3 = lookup_table(e2 + 2, "Size band fit", SIZE_FIT)
    put(sw, e3 + 2, 1, "Defaults set from the Partners profile: lending Partners are mostly Funders/working-capital/MCA lenders with 11-50 staff. Change any blue cell to re-rank.")

    # ---------------- Companies
    u = u.sort_values(["fit_py", "contacts_on_file"], ascending=[False, False]).reset_index(drop=True)
    u["company_id"] = ["C" + str(i + 1).zfill(5) for i in range(len(u))]
    cols = [("Company_ID", "company_id", 10), ("Company", "company", 30), ("Domain", "domain", 26), ("Alt_Domains (merged)", "alt_domains", 22),
            ("Segment", "segment", 12), ("Sub_Type", "sub_type", 22), ("Classification_Confidence", "class_conf", 12),
            ("Classification_Reason", "classification_reason", 46), ("Size_Band", "size_band", 10), ("Employees_Est", "employees_est", 10),
            ("Size_Basis", "size_basis", 30), ("Years_In_Business", "years_in_business", 9), ("HQ_State", "st", 8),
            ("Website_Sub_Industry", "si", 30), ("Website_One_Liner", "ol", 50), ("Verification_Status", "verification", 22),
            ("Website_Check", "ss", 12), ("Mail_Server_Check", "mx", 12), ("Contacts_In_Warehouse", "n", 10),
            ("Contacts_On_Tab", "contacts_on_file", 9), ("Decision_Maker_On_File", None, 10), ("Replies_To_Us", "replies", 8),
            ("Positive_Replies", "pos", 8), ("Attio_Record_ID", "attio_id", 22), ("Attio_Connection", "attio_conn", 12),
            ("Attio_Last_Interaction", "attio_last", 18),
            ("C_Lookalike", "s_lookalike", 9), ("C_Segment_Fit", None, 9), ("C_SubType_Fit", None, 9), ("C_Size_Fit", None, 9),
            ("C_Outbound", "s_outbound", 9), ("C_Engagement", "s_engagement", 9), ("C_DM_Coverage", "s_coverage", 9),
            ("Fit_Score", None, 9), ("Disqualifier", "disq", 22), ("Tier", None, 14), ("Reason_For_Score", "score_reason", 70),
            ("Source", "source", 40)]
    cs = S["Companies"]
    header(cs, 1, [c[0] for c in cols], [c[2] for c in cols])
    col_idx = {c[0]: i + 1 for i, c in enumerate(cols)}
    L = {k: get_column_letter(v) for k, v in col_idx.items()}
    n_rows = len(u)
    for i, r in enumerate(u.itertuples(index=False), start=2):
        rd = r._asdict()
        for name, key, _ in cols:
            if key:
                v = rd.get(key)
                if key == "years_in_business" and v is not None and v == v: v = float(v)
                put(cs, i, col_idx[name], clean(v))
        put(cs, i, col_idx["Decision_Maker_On_File"], "Yes" if rd["dm_on_file"] else "No")
        put(cs, i, col_idx["C_Segment_Fit"], f"=IFERROR(INDEX({seg_v},MATCH({L['Segment']}{i},{seg_k},0)),0)")
        put(cs, i, col_idx["C_SubType_Fit"], f"=IFERROR(INDEX({sub_v},MATCH({L['Sub_Type']}{i},{sub_k},0)),0)")
        put(cs, i, col_idx["C_Size_Fit"], f"=IFERROR(INDEX({siz_v},MATCH({L['Size_Band']}{i},{siz_k},0)),0)")
        terms = "+".join([f"{L['C_Lookalike']}{i}*{wref['s_lookalike']}", f"{L['C_Segment_Fit']}{i}*{wref['seg_fit']}",
                          f"{L['C_SubType_Fit']}{i}*{wref['sub_fit']}", f"{L['C_Size_Fit']}{i}*{wref['size_fit']}",
                          f"{L['C_Outbound']}{i}*{wref['s_outbound']}", f"{L['C_Engagement']}{i}*{wref['s_engagement']}",
                          f"{L['C_DM_Coverage']}{i}*{wref['s_coverage']}"])
        put(cs, i, col_idx["Fit_Score"], f"=IF({WSUM}=0,0,ROUND(100*({terms})/{WSUM},1))")
        fs, dq_ = f"{L['Fit_Score']}{i}", f"{L['Disqualifier']}{i}"
        put(cs, i, col_idx["Tier"], f'=IF({dq_}<>"",{dq_},IF({fs}>={cut_ref[0]},"Tier 1",IF({fs}>={cut_ref[1]},"Tier 2","Tier 3")))')
    cs.freeze_panes = "C2"
    cs.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{n_rows + 1}"
    cs.cell(row=1, column=col_idx["Fit_Score"]).comment = Comment("Formula: 100 × Σ(component × weight) ÷ Σ weights. Weights on Scoring_Weights.", "TAM")
    cs.cell(row=1, column=col_idx["Disqualifier"]).comment = Comment("Partner, Competitor, Out of business, or No contacts. Overrides the score-based tier.", "TAM")
    rng = lambda name: f"Companies!${L[name]}$2:${L[name]}${n_rows + 1}"

    # ---------------- Contacts
    ct = S["Contacts"]
    cid = dict(zip(u.domain, u.company_id)); cname = dict(zip(u.domain, u.company))
    contacts = contacts[contacts.domain.isin(cid)].copy()
    contacts["company_id"] = contacts.domain.map(cid)
    contacts["_ord"] = contacts.company_id
    contacts = contacts.sort_values(["_ord", "dm_rank"], key=lambda s: s if s.name != "dm_rank" else s.replace("", "9"))
    ccols = [("Contact_ID", 10), ("Company_ID", 10), ("Company", 28), ("Domain", 24), ("First_Name", 14), ("Last_Name", 16),
             ("Title", 32), ("Seniority", 14), ("Best_Decision_Maker", 12), ("Email", 30), ("Email_Status", 22), ("Phone", 16),
             ("LinkedIn", 36), ("Source", 28), ("Attio_Match", 18)]
    header(ct, 1, [c[0] for c in ccols], [c[1] for c in ccols])
    for i, r in enumerate(contacts.itertuples(index=False), start=2):
        vals = ["P" + str(i - 1).zfill(6), r.company_id, cname[r.domain], r.domain, r.fn, r.ln, r.t, r.seniority, r.best_dm,
                r.e, r.email_status, r.ph, r.linkedin, r.source, r.attio_match]
        for j, v in enumerate(vals, 1):
            put(ct, i, j, clean(v) if v != "" else None)
    ct.freeze_panes = "C2"
    ct.auto_filter.ref = f"A1:{get_column_letter(len(ccols))}{len(contacts) + 1}"

    # ---------------- Partners_Profile
    pp = S["Partners_Profile"]
    pp["A1"] = "Current Partners (Attio 'Partners' list) and the profile derived from them"; pp["A1"].font = TITLE
    pcols = ["Partner", "Domains", "Offer", "Business_Lending", "Account_Stage", "Start_Date", "Last_Booking", "Meetings_Delivered",
             "Segment", "Sub_Type", "Size_Band", "Employee_Range (Attio)", "Est_ARR (Attio)", "Founded", "HQ", "In_TAM_Company_ID", "Notes"]
    header(pp, 3, pcols, [24, 30, 12, 10, 20, 11, 11, 10, 12, 22, 10, 12, 12, 8, 16, 12, 50])
    pm = src["partner_meetings"]
    pmd = dict(zip(pm.pk.astype(str), pd.to_numeric(pm.meetings)))
    udom = u.set_index("domain")
    for i, p in enumerate(partners.itertuples(index=False), start=4):
        doms = [norm_domain(x) for x in p.domains.split("|") if x]
        row = next((udom.loc[d] for d in doms if d in udom.index), None)
        meets = sum(pmd.get(k, 0) for k in p.meeting_partner_keys.split("|") if k) or None
        vals = [p.partner_name, ", ".join(doms), p.offer, p.business_lending, p.account_stage, p.start_date, p.last_booking, meets,
                row.segment if row is not None else "n/a (non-lending)", row.sub_type if row is not None else "",
                row.size_band if row is not None else "", p.employee_range, p.estimated_arr, p.foundation_year, p.hq,
                row.company_id if row is not None else "", p.notes]
        for j, v in enumerate(vals, 1):
            put(pp, i, j, v if v != "" else None)
    p_last = 3 + len(partners)
    pr = lambda col: f"${col}$4:${col}${p_last}"
    r = p_last + 3
    put(pp, r, 1, "Profile of Business Lending Partners (formulas over the table above)", BB); r += 1
    for title, col, vals in [("Segment mix", "I", ["Funder", "Hybrid", "Broker/ISO", "Unknown"]),
                             ("Size band mix", "K", ["Micro", "Small", "Mid", "Large", "Enterprise", "Unknown"]),
                             ("Sub-type mix", "J", [s for s, _ in SUB_FIT]),
                             ("Account stage", "E", ["Scaling", "Steady State", "First Batch Delivering", "Onboarding", "Paused", "Churned"])]:
        put(pp, r, 1, title, BB); put(pp, r, 2, "Partners", BB); put(pp, r, 3, "TAM companies", BB); put(pp, r, 4, "Partner share", BB); r += 1
        start = r
        for v in vals:
            put(pp, r, 1, v)
            put(pp, r, 2, f'=COUNTIFS({pr(col)},A{r},{pr("D")},"Y")')
            tam_col = {"I": "Segment", "K": "Size_Band", "J": "Sub_Type"}.get(col)
            if tam_col:
                put(pp, r, 3, f"=COUNTIF({rng(tam_col)},A{r})")
            put(pp, r, 4, f"=IF(SUM($B${start}:$B${start + len(vals) - 1})=0,0,B{r}/SUM($B${start}:$B${start + len(vals) - 1}))").number_format = "0%"
            r += 1
        r += 1
    put(pp, r, 1, "What separates Partners from the rest of the universe", BB); r += 1
    for line in partner_insights(u):
        put(pp, r, 1, line); r += 1

    # ---------------- Summary
    sm = S["Summary"]
    sm["A1"] = f"Business Lending TAM: summary (refreshed {TODAY.isoformat()})"; sm["A1"].font = TITLE
    tiers = ["Tier 1", "Tier 2", "Tier 3", "Partner", "Disqualified - Competitor", "Disqualified - Out of business", "Disqualified - No contacts"]
    segs = ["Funder", "Hybrid", "Broker/ISO", "Unknown"]
    bands = ["Micro", "Small", "Mid", "Large", "Enterprise", "Unknown"]
    r = 3
    put(sm, r, 1, "Headline", BB); r += 1
    heads = [("Companies in TAM", f"=COUNTA({rng('Company_ID')})"),
             ("Funders", f'=COUNTIF({rng("Segment")},"Funder")'),
             ("Brokers/ISOs", f'=COUNTIF({rng("Segment")},"Broker/ISO")'),
             ("Hybrid", f'=COUNTIF({rng("Segment")},"Hybrid")'),
             ("Unknown segment", f'=COUNTIF({rng("Segment")},"Unknown")'),
             ("Tier 1 accounts", f'=COUNTIF({rng("Tier")},"Tier 1")'),
             ("Tier 1 with a decision-maker on file", f'=COUNTIFS({rng("Tier")},"Tier 1",{rng("Decision_Maker_On_File")},"Yes")'),
             ("Tier 1 with a verified decision-maker", f'=COUNTIFS({rng("Tier")},"Tier 1",{rng("C_DM_Coverage")},1)'),
             ("Companies with any contact on Contacts tab", f'=COUNTIF({rng("Contacts_On_Tab")},">0")'),
             ("Companies with a decision-maker on file", f'=COUNTIF({rng("Decision_Maker_On_File")},"Yes")'),
             ("Contacts on Contacts tab", f"=COUNTA(Contacts!$A$2:$A${len(contacts) + 1})"),
             ("Contacts with verified-valid email", f'=COUNTIF(Contacts!$K$2:$K${len(contacts) + 1},"Verified valid")'),
             ("Contacts flagged as best decision-makers", f'=COUNTIF(Contacts!$I$2:$I${len(contacts) + 1},"Yes*")'),
             ("Companies excluded as not Business Lending (see Excluded)", len(excluded))]
    for lab, f in heads:
        put(sm, r, 1, lab); put(sm, r, 2, f, BB); r += 1
    r += 1
    put(sm, r, 1, "Companies by Segment × Tier", BB); r += 1
    tier_hdr = r
    header(sm, r, ["Segment"] + tiers + ["Total"]); r += 1
    seg_start = r
    for s in segs:
        put(sm, r, 1, s)
        for j, t in enumerate(tiers, 2):
            put(sm, r, j, f'=COUNTIFS({rng("Segment")},$A{r},{rng("Tier")},{get_column_letter(j)}${tier_hdr})')
        put(sm, r, len(tiers) + 2, f"=SUM(B{r}:{get_column_letter(len(tiers) + 1)}{r})", BB); r += 1
    seg_end = r - 1
    put(sm, r, 1, "Total", BB)
    for j in range(2, len(tiers) + 3):
        put(sm, r, j, f"=SUM({get_column_letter(j)}{seg_start}:{get_column_letter(j)}{seg_end})", BB)
    r += 2
    put(sm, r, 1, "Companies by Segment × Size Band × Tier (pivot-style; Tier 1 / Tier 2 / Tier 3)", BB); r += 1
    hdr_row = r
    hdr = ["Segment", "Size Band"] + ["Tier 1", "Tier 2", "Tier 3", "Partner", "Disqualified (all)", "Total", "Tier 1 w/ DM"]
    header(sm, r, hdr); r += 1
    for s in segs:
        for b in bands:
            put(sm, r, 1, s); put(sm, r, 2, b)
            base = f'{rng("Segment")},$A{r},{rng("Size_Band")},$B{r}'
            for j, t in enumerate(["Tier 1", "Tier 2", "Tier 3", "Partner"], 3):
                put(sm, r, j, f'=COUNTIFS({base},{rng("Tier")},"{t}")')
            put(sm, r, 7, f'=COUNTIFS({base},{rng("Tier")},"Disqualified*")')
            put(sm, r, 8, f"=SUM(C{r}:G{r})", BB)
            put(sm, r, 9, f'=COUNTIFS({base},{rng("Tier")},"Tier 1",{rng("Decision_Maker_On_File")},"Yes")')
            r += 1
    sm.column_dimensions["A"].width = 52
    for col in "BCDEFGHIJ":
        sm.column_dimensions[col].width = 14
    chart = BarChart(); chart.type = "col"; chart.grouping = "stacked"; chart.overlap = 100
    chart.title = "TAM by segment and tier"; chart.y_axis.title = "Companies"
    data = Reference(sm, min_col=2, max_col=4, min_row=tier_hdr, max_row=seg_end)
    cats = Reference(sm, min_col=1, min_row=seg_start, max_row=seg_end)
    chart.add_data(data, titles_from_data=True); chart.set_categories(cats)
    chart.height, chart.width = 8, 16
    sm.add_chart(chart, "L3")

    # ---------------- Assumptions
    asn = S["Assumptions"]
    asn["A1"] = "Assumptions, thresholds, and rules"; asn["A1"].font = TITLE
    r = 3
    put(asn, r, 1, "Size bands (employees; first available of: Attio employee range → website stated staff → contact-DB size bucket; revenue as fallback)", BB); r += 1
    header(asn, r, ["Band", "Min employees", "Max employees", "Revenue fallback (USD)"]); r += 1
    for (b, lo, hi), (_, rlo, rhi) in zip(C.SIZE_BANDS, C.REVENUE_BANDS):
        put(asn, r, 1, b); put(asn, r, 2, lo); put(asn, r, 3, hi if hi < 10**8 else "no max"); put(asn, r, 4, f"${rlo:,.0f} – {'$' + format(rhi, ',.0f') if rhi < 10**14 else 'no max'}"); r += 1
    r += 1
    put(asn, r, 1, "Scope: a company is excluded (Excluded tab) when the first matching rule fires", BB); r += 1
    for reason, rx in C.EXCLUDE_RULES:
        put(asn, r, 1, reason); put(asn, r, 2, rx); r += 1
    for line in ["Personal email domains, .edu/.gov, non-US country domains, and generic domains with >5,000 contacts are excluded.",
                 "Domains found only via the contact-database industry tag are kept only if the company name or domain contains a lending word (fund, capital, financ, lend, loan, credit, advance, factor, leas, money, merchant).",
                 "A website profile that exists but describes no lending activity excludes the company (unless it carries the mca_alt_lending vertical tag).",
                 "Real-estate/mortgage, consumer lenders, and banks/credit unions are excluded as out of scope (not buyers of business-funding appointments). Change EXCLUDE_RULES in tam/classify.py to include them."]:
        put(asn, r, 1, line); r += 1
    r += 1
    put(asn, r, 1, "Segment rules (keywords searched in company name, website sub-industry, website one-liner, contact-DB company name, domain)", BB); r += 1
    put(asn, r, 1, "Broker/ISO keywords"); put(asn, r, 2, C.BROKER_RX); r += 1
    put(asn, r, 1, "Funder keywords"); put(asn, r, 2, C.FUNDER_RX); r += 1
    put(asn, r, 1, "Both → Hybrid. Neither → Unknown (Funder if it carries the mca_alt_lending vertical tag). Known Partner business models override."); r += 1
    put(asn, r, 1, "Confidence: High = website text + segment + sub-type keywords; Medium = segment from keywords or tag; Low = name-only or tag-only."); r += 2
    put(asn, r, 1, "Sub-type rules (first match wins)", BB); r += 1
    for st, rx in C.SUBTYPE_RULES:
        put(asn, r, 1, st); put(asn, r, 2, rx); r += 1
    r += 1
    put(asn, r, 1, "Disqualifiers", BB); r += 1
    for line in [f"Competitor: text matches {C.COMPETITOR_RX}", "Out of business: on defunct list, website parked/DNS dead, or mail domain NXDOMAIN/null MX.",
                 "No contacts: no contact in the warehouse and none on the Contacts tab.", "Partner: domain on the Attio Partners list (flagged, not ranked)."]:
        put(asn, r, 1, line); r += 1
    r += 1
    put(asn, r, 1, "Contacts", BB); r += 1
    for line in ["Seniority from title + vendor seniority: Owner/C-level, VP/Director, Manager, IC.",
                 "Best decision-makers: up to 3 per company among Owner/C-level and VP/Director with a non-invalid email, ranked by seniority, verified email, phone on file, sales/ISO-relevant title.",
                 "Emails and phones are exactly as stored in the warehouse or Attio. Nothing was guessed or generated. Invalid emails are kept but flagged 'Invalid - do not email'.",
                 "Attio_Match: the Attio people export has names only (no email/company), so matches are by full name and must be confirmed."]:
        put(asn, r, 1, line); r += 1
    asn.column_dimensions["A"].width = 70; asn.column_dimensions["B"].width = 110

    # ---------------- Data_Quality
    dqs = S["Data_Quality"]
    dqs["A1"] = "Data quality: what was cleaned, what is missing"; dqs["A1"].font = TITLE
    r = 3
    gaps = [("Companies in TAM", len(u)),
            ("Duplicate companies merged (fuzzy/exact name across domains)", len(dq["merged"])),
            ("Excluded as not Business Lending / out of scope", len(excluded)),
            ("Segment = Unknown (unclassified)", int((u.segment == "Unknown").sum())),
            ("Low classification confidence", int((u.class_conf == "Low").sum())),
            ("Size band = Unknown", int((u.size_band == "Unknown").sum())),
            ("No website profile in warehouse", int((~u.in_site_profile).sum())),
            ("Verification = Not checked (no site/MX check retrieved)", int((u.verification == "Not checked").sum())),
            ("Verification = Dead / parked", int((u.verification == "Dead / parked").sum())),
            ("No contacts on Contacts tab", int((u.contacts_on_file == 0).sum())),
            ("No decision-maker on file", int((~u.dm_on_file).sum())),
            ("No HQ state", int(u.st.isna().sum() + (u.st.astype(str) == "").sum())),
            ("No years-in-business signal", int(u.years_in_business.isna().sum())),
            ("Not in Attio", int((~u.in_attio).sum())),
            ("Contacts with invalid email (kept, flagged)", int((contacts.vs == "not_valid").sum())),
            ("Contacts with no phone", int((contacts.ph == "").sum())),
            ("Contacts with no LinkedIn", int((contacts.linkedin == "").sum()))]
    header(dqs, r, ["Check", "Count"], [70, 12]); r += 1
    for lab, v in gaps:
        put(dqs, r, 1, lab); put(dqs, r, 2, v); r += 1
    r += 1
    put(dqs, r, 1, "Known gaps and caveats", BB); r += 1
    for n in dq["notes"]:
        put(dqs, r, 1, n); r += 1
    r += 1
    put(dqs, r, 1, "Duplicates merged", BB); r += 1
    header(dqs, r, ["Kept domain", "Kept company", "Merged domain", "Merged company", "Match"]); r += 1
    for m in dq["merged"]:
        for j, k in enumerate(["kept_domain", "kept_company", "merged_domain", "merged_company", "match"], 1):
            put(dqs, r, j, m[k])
        r += 1
    r += 1
    put(dqs, r, 1, "Unclassified (Segment = Unknown), top 200 by contacts", BB); r += 1
    header(dqs, r, ["Company_ID", "Company", "Domain", "Reason"]); r += 1
    for x in u[u.segment == "Unknown"].sort_values("n", ascending=False).head(200).itertuples():
        for j, v in enumerate([x.company_id, x.company, x.domain, x.classification_reason], 1):
            put(dqs, r, j, v)
        r += 1
    for col, w in zip("BCDE", [30, 26, 30, 30]):
        dqs.column_dimensions[col].width = w

    # ---------------- Excluded
    ex = S["Excluded"]
    header(ex, 1, ["Domain", "Company", "Exclusion_Reason", "Rule_Hit", "Website_Sub_Industry", "Contacts_In_Warehouse", "Source"],
           [28, 32, 40, 30, 34, 12, 40])
    excluded = excluded.sort_values(["excl_reason", "n"], ascending=[True, False])
    for i, x in enumerate(excluded.itertuples(index=False), start=2):
        for j, v in enumerate([x.domain, x.company, x.excl_reason, x.excl_hit, x.si, x.n, x.source], 1):
            put(ex, i, j, clean(v))
    ex.freeze_panes = "B2"; ex.auto_filter.ref = f"A1:G{len(excluded) + 1}"

    # ---------------- README
    rd = ws_readme
    rd["A1"] = "Renaissance Business Lending TAM"; rd["A1"].font = TITLE
    lines = [
        f"Refresh date: {TODAY.isoformat()}  |  Build mode: {mode}",
        "",
        "What this is",
        "Every Business Lending company (MCA funders, direct lenders, ISOs, brokers) we hold in the warehouse or Attio, cleaned, de-duplicated,",
        "classified as Funder / Broker/ISO / Hybrid, sized, and ranked by fit to our current Partners. Use it to decide who the setter floor calls first.",
        "",
        "How to use it",
        "1. Companies tab: sorted by Fit_Score. Filter Tier = 'Tier 1' for the call-now list. Partners are flagged, not ranked.",
        "2. Contacts tab: filter Company_ID; 'Best_Decision_Maker' marks the top 1-3 people to ask for. Never email 'Invalid - do not email'.",
        "3. Scoring_Weights tab: change any blue cell (weights, fit tables, tier cutoffs) and every Fit_Score and Tier recalculates.",
        "4. Reason_For_Score explains the top three drivers and the biggest gaps for each company.",
        "",
        "Data sources",
        "ClickHouse warehouse (read-only): mirror.domain_vertical (lending vertical tag), mirror.company_site_profile (website profile),",
        "  mirror.leads_current (contacts, titles, email verification, size), mirror.site_check_domains + mirror.email_domain_mx (website / mail-server checks),",
        "  mirror.tib_domain_mv (domain age), mirror.defunct_domain, derived.v_reply_canonical + derived.reply_is_positive_strict (replies),",
        "  core.meeting (partner meeting volume).",
        "Attio: Partners list (signed customers), Companies export (record IDs, connection strength, last interaction), People export (names only).",
        "No paid enrichment was run. LeadMagic/enrichment needs approval first (see Data_Quality for the gaps it would close).",
        "",
        "Verification performed (free signals only)",
        "Website crawl status, mail-server (MX) status, defunct-domain list, email verification status per contact, Attio record match.",
        "This container could not browse the web, so no live website visits were made in this run.",
        "",
        "Refresh weekly",
        "python tam/build_tam.py            (live: queries ClickHouse, stages amir_lab.tam_* tables, rebuilds this workbook)",
        "python tam/build_tam.py --source snapshot   (rebuild from the saved JSON snapshots in tam/data/raw)",
        "Put new Attio exports in tam/input/ (attio_companies.xlsx, attio_people.csv) and update tam/data/partners.csv when Partners change.",
    ]
    for i, line in enumerate(lines, 3):
        c = put(rd, i, 1, line, BB if line in ("What this is", "How to use it", "Data sources", "Refresh weekly", "Verification performed (free signals only)") else B)
    rd.column_dimensions["A"].width = 150

    for ws in wb.worksheets:
        ws.sheet_view.zoomScale = 100
    wb.save(out_path)
    return u, contacts


def partner_insights(u):
    P, R = u[u.is_partner], u[~u.is_partner & (u.disq == "")]
    out = []
    def pct(df, mask): return f"{(mask.sum() / max(len(df), 1)):.0%}"
    out.append(f"Lending Partners in TAM: {len(P)}. Rest of universe (rankable): {len(R):,}.")
    out.append(f"Funder or Hybrid: Partners {pct(P, P.segment.isin(['Funder', 'Hybrid']))} vs universe {pct(R, R.segment.isin(['Funder', 'Hybrid']))}.")
    out.append(f"Small (11-50 staff): Partners {pct(P, P.size_band.eq('Small'))} vs universe {pct(R, R.size_band.eq('Small'))}.")
    out.append(f"Working-capital/MCA/LOC/general business lending sub-type: Partners {pct(P, P.sub_type.isin(['MCA funder', 'LOC / working capital', 'Commercial lending (general)', 'Marketplace / broker']))} vs universe {pct(R, R.sub_type.isin(['MCA funder', 'LOC / working capital', 'Commercial lending (general)', 'Marketplace / broker']))}.")
    out.append(f"HQ in NY: Partners {pct(P, P.st.astype(str).str.upper().isin(['NY', 'NEW YORK']))} vs universe {pct(R, R.st.astype(str).str.upper().isin(['NY', 'NEW YORK']))} (Attio shows 9 of 16 lending Partners in the NY metro).")
    out.append(f"Has sales/ISO/BD staff in contact DB: Partners {pct(P, P.ns > 0)} vs universe {pct(R, R.ns > 0)}.")
    out.append("Equipment finance, factoring, ABL, SBA, and CDFI lenders are absent from the Partners list, so they get lower sub-type fit by default.")
    return out


def recalc(path: Path):
    script = Path("/root/.claude/skills/synced")
    cands = list(script.glob("*/xlsx/scripts/recalc.py"))
    if not cands:
        print("  (recalc.py not found; open in Excel to calculate)")
        return None
    res = subprocess.run([sys.executable, str(cands[0]), str(path), "180"], capture_output=True, text=True)
    try:
        return json.loads(res.stdout.strip().splitlines()[-1] if res.stdout.strip().startswith("{") is False else res.stdout)
    except Exception:
        return {"raw": res.stdout[-2000:], "err": res.stderr[-2000:]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", choices=["live", "snapshot"], default="live")
    ap.add_argument("--per-domain", type=int, default=25, help="contacts per company in live mode")
    args = ap.parse_args()

    attio_comp, attio_people, partners = load_attio()
    extra = [d.strip() for d in (ROOT / "data" / "extra_domains.txt").read_text().split() if d.strip()]
    src = load_snapshot() if args.source == "snapshot" else load_live(extra, args.per_domain)
    mode = "snapshot (read-only connector extraction)" if args.source == "snapshot" else "live (ClickHouse)"

    u, excluded, dq, d2p = build_universe(src, attio_comp, partners, extra)
    u = classify_companies(u, partners, src)
    contacts = build_contacts(src, u, d2p, partners, attio_people)
    u = score_inputs(u, contacts, partners)
    u = compute_python_score(u)

    if args.source == "snapshot":
        dq["notes"] += [
            "This edition was extracted through the read-only ClickHouse connector (2 MB / 60 s limits), so it differs from a live run:",
            " - Contacts tab holds the best 3 contacts per company (ranked by seniority, verified email, phone). Contacts_In_Warehouse shows the full count. A live run pulls 25 per company.",
            " - Website/MX verification was retrieved for about two-thirds of domains; the rest show 'Not checked'. A live run checks all.",
            " - 750 Attio-only keyword domains were not queried in the warehouse; they appear with Attio data only.",
        ]
    dq["notes"] += [
        "Replies_To_Us counts replies to any campaign; most lender replies came from our borrower-side 'Business Funding' campaigns, not a pitch for leads.",
        "The Attio people export contains names only, so people cannot be tied to companies without the Attio API.",
        "Attio Partners list: 1West exists as two company records; GreenBridge has a typo domain (greenbrudgecapital.com).",
        "Paid enrichment (LeadMagic) not run. It would close: size for 'Unknown' size bands, decision-makers for Tier 1/2 accounts without one, and email verification for unverified decision-makers.",
    ]

    OUT.mkdir(exist_ok=True)
    out_path = OUT / f"TAM_Business_Lending_{TODAY.isoformat()}.xlsx"
    u, contacts = write_workbook(u, contacts, excluded, dq, partners, src, out_path, mode)
    rc = recalc(out_path)
    print(f"\nWorkbook: {out_path}")
    print(f"Recalc: {rc}")
    print_summary(u, contacts, excluded, dq)


def print_summary(u, contacts, excluded, dq):
    seg = u.segment.value_counts()
    t = u.tier_py.value_counts()
    print("\n================ TAM SUMMARY ================")
    print(f"Total companies in TAM: {len(u):,}  (excluded as out of scope: {len(excluded):,}; duplicates merged: {len(dq['merged'])})")
    print("Segments: " + ", ".join(f"{k} {v:,}" for k, v in seg.items()))
    print("Tiers:    " + ", ".join(f"{k} {v:,}" for k, v in t.items()))
    t1 = u[u.tier_py == "Tier 1"]
    print(f"Tier 1: {len(t1):,} accounts, {int(t1.dm_on_file.sum()):,} with a decision-maker on file ({int(t1.dm_verified.sum()):,} verified)")
    print(f"Contacts on tab: {len(contacts):,}  | best decision-makers flagged: {(contacts.best_dm != '').sum():,}")
    print("\nTop 25 prospects:")
    top = u[u.disq == ""].head(25)
    for i, r in enumerate(top.itertuples(), 1):
        print(f"{i:>2}. {r.company} ({r.domain}) {r.fit_py:.1f} [{r.segment}/{r.sub_type}/{r.size_band}] - {r.score_reason}")
    print("\nBiggest data gaps:")
    gaps = [("Size band Unknown", (u.size_band == "Unknown").sum()), ("No decision-maker on file", (~u.dm_on_file).sum()),
            ("Segment Unknown", (u.segment == "Unknown").sum()), ("No website profile", (~u.in_site_profile).sum()),
            ("Verification not checked", (u.verification == "Not checked").sum()), ("Not in Attio", (~u.in_attio).sum())]
    for lab, v in sorted(gaps, key=lambda x: -x[1]):
        print(f" - {lab}: {v:,} of {len(u):,}")


if __name__ == "__main__":
    main()
