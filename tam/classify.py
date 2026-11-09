"""Explicit, explainable rules for scope, segment, sub-type, size band, and
score inputs. Every rule that fires is recorded so the workbook can show why.

Edit the keyword lists here (and mirror any change on the Assumptions tab text,
which is generated from these same lists).
"""
from __future__ import annotations

import re

# ---------------------------------------------------------------- scope rules
# Checked in order; the first match removes the company from the TAM and
# sends it to the Excluded tab with this reason.
EXCLUDE_RULES: list[tuple[str, str]] = [
    ("Consumer lender", r"\bpayday\b|consumer (lend|loan|financ)|personal loans?\b|installment loans?|title loans?|pawn|check cashing|cash store|auto (loan|financ)|student loan|tax refund"),
    ("Mortgage / real-estate lender", r"mortgage|home loan|real estate (lend|loan|invest)|hard money|private money|fix and flip|bridge loans? for real|commercial real estate lend|\bdscr\b"),
    ("Bank / credit union", r"\bbank(ing)?\b(?! ?statement)|credit union|savings (and|&) loan|\bfdic\b|bancorp|\bbancshares\b"),
    ("Payments / processing (not a lender)", r"payment processing|merchant services|credit card processing|\bpos\b systems?|payment gateway|payments? (platform|company)"),
    ("Accounting / tax / advisory (not a lender)", r"accounting|bookkeeping|\bcpa\b|tax (prep|advis|credit)|wealth management|financial (planning|advis)|investment (advis|management|banking)|insurance agency|\binsurance\b(?!.*fund)"),
    ("Auto dealer / BHPH", r"car dealership|auto dealer|buy here pay here|\bbhph\b|used cars?"),
    ("Litigation / medical / pre-settlement funding", r"pre.?settlement|lawsuit (fund|loan)|litigation (fund|financ)|personal injury|medical (funding|lien)"),
    ("Debt relief / credit repair", r"debt (relief|settlement|consolidation)|credit repair|credit counsel"),
]

# Lead-gen / appointment-setting firms selling into lenders compete with us.
COMPETITOR_RX = r"lead generation|appointment setting|\bleads? for (lenders|funders|brokers)|live transfers?|\baged leads\b|mca leads|pay per (call|lead|appointment)"

# ------------------------------------------------------------- segment rules
BROKER_RX = r"brokerage|\bbroker(s)?\b|\biso\b|independent sales|marketplace|we (match|connect)|connects? (small )?business(es)? (owners )?with|compare (multiple )?(lenders|offers)|network of (lenders|funders)|loan (consultant|advisor)|funding (consultant|advisor|advisory)|commercial finance (broker|advisor)"
FUNDER_RX = r"direct (lender|funder)|we fund|\bfunder\b|merchant cash advance|\bmca\b|factoring|equipment (financ|leas)|revenue.based|asset.based|working capital|business (lending|loans?|financing|funding)|lines? of credit|term loans?|\bsba\b|commercial (lending|financ)|small business (lending|financ|funding)|alternative (lending|financ)|capital provider|lender"

SUBTYPE_RULES: list[tuple[str, str]] = [
    ("MCA funder", r"merchant cash advance|\bmca\b|cash advance|daily (payments?|remittance)|split funding"),
    ("Factoring", r"factoring|invoice financ|receivables? financ|accounts receivable"),
    ("Equipment finance", r"equipment (financ|leas|loan)|\bleasing\b"),
    ("SBA lender", r"\bsba\b|7\(a\)|\b504\b"),
    ("Revenue-based", r"revenue.based"),
    ("Asset-based / ABL", r"asset.based|\babl\b"),
    ("CDFI / community lender", r"community development|\bcdfi\b|microloan|micro.?lend|nonprofit lend"),
    ("LOC / working capital", r"lines? of credit|working capital|business credit line"),
    ("Term lender", r"term loans?"),
    ("Marketplace / broker", r"marketplace|brokerage|\bbroker\b|\biso\b"),
    ("Commercial lending (general)", r"commercial (lending|financ|loan)|business (lending|loans?|financing|funding)|small business (lending|financ)"),
]

# --------------------------------------------------------------- size bands
SIZE_BANDS = [  # (band, min employees inclusive, max inclusive)
    ("Micro", 1, 10),
    ("Small", 11, 50),
    ("Mid", 51, 200),
    ("Large", 201, 1000),
    ("Enterprise", 1001, 10**9),
]
REVENUE_BANDS = [  # fallback when no employee signal: annual revenue USD
    ("Micro", 0, 1_000_000),
    ("Small", 1_000_000, 10_000_000),
    ("Mid", 10_000_000, 50_000_000),
    ("Large", 50_000_000, 500_000_000),
    ("Enterprise", 500_000_000, 10**15),
]

DM_TITLE_RX = r"\bowner\b|founder|\bceo\b|chief|president|\bcoo\b|\bcfo\b|\bcro\b|managing (member|director|partner)|\bprincipal\b|\bpartner\b(?! (relations|manager|success))"
VP_TITLE_RX = r"\bvp\b|vice president|director|head of|\bhead\b"
MGR_TITLE_RX = r"manager|supervisor|lead\b"
SALES_DM_RX = r"\biso\b|partner (?:relations|ship)|broker relations|channel|business development|sales|revenue|growth|marketing"


def _search(rx: str, text: str):
    m = re.search(rx, text, re.I)
    return m.group(0) if m else None


def scope(text: str):
    for reason, rx in EXCLUDE_RULES:
        hit = _search(rx, text)
        if hit:
            return reason, hit
    return None, None


def segment(text: str, name: str):
    b = _search(BROKER_RX, text)
    f = _search(FUNDER_RX, text)
    if b and f:
        # "broker" language wins only when it is explicit; lenders often say "lender network".
        return "Hybrid", f'broker kw "{b}" + funder kw "{f}"'
    if b:
        return "Broker/ISO", f'broker kw "{b}"'
    if f:
        return "Funder", f'funder kw "{f}"'
    nb = _search(r"funding|capital|financ|lend|advance", name)
    if nb:
        return "Unknown", f'name kw "{nb}" only'
    return "Unknown", "no segment keywords"


def subtype(text: str):
    for st, rx in SUBTYPE_RULES:
        hit = _search(rx, text)
        if hit:
            return st, hit
    return "Unspecified", None


def parse_employees(*vals):
    """Return a representative employee count from strings like '11-50', '25', '51-200 employees'."""
    for v in vals:
        if v is None:
            continue
        s = str(v).replace(",", "")
        nums = [int(x) for x in re.findall(r"\d+", s)]
        if not nums:
            continue
        if "+" in s and len(nums) == 1:
            return nums[0]
        if len(nums) >= 2:
            return int((nums[0] + nums[1]) / 2)
        if nums[0] > 0:
            return nums[0]
    return None


def size_band(emp, revenue):
    if emp:
        for band, lo, hi in SIZE_BANDS:
            if lo <= emp <= hi:
                return band, f"employees≈{emp}"
    if revenue and revenue > 0:
        for band, lo, hi in REVENUE_BANDS:
            if lo <= revenue < hi:
                return band, f"revenue≈${revenue:,.0f}"
    return "Unknown", "no size signal"


def seniority(title: str, raw: str = ""):
    t = f"{title or ''} {raw or ''}".lower()
    if re.search(DM_TITLE_RX, t) or re.search(r"c_suite|c-suite|c-team|founder|owner", t):
        return "Owner/C-level"
    if re.search(VP_TITLE_RX, t) or raw.lower() in ("vp", "director", "head"):
        return "VP/Director"
    if re.search(MGR_TITLE_RX, t) or raw.lower() == "manager":
        return "Manager"
    return "IC"
