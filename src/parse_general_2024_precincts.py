"""
Parse 2024 California November general per-county precinct sources into
OpenElections precinct CSVs under ``2024/counties/`` named
``20241105__ca__general__<county>__precinct.csv``.

No consolidated county-level file exists for 2024, so candidate/office
canonicalization is rule-based (office titles via ``parse_office``, party
codes via ``map_party``) and each county's precinct sums are verified against
that county's own summary figures (in-source county-rollup rows or summary
files); mismatches are reported and the CSV is only written when clean or
soft (write-in-only).

State proposition / judge / local contests are skipped (repo convention);
write-in aggregates are dropped, named write-ins kept.

Source families, one parser each:

* ``parse_pageformat`` — "Page: N of M" SOVC workbooks: Colusa, Contra Costa,
  Fresno, Inyo, Marin, San Benito, San Francisco, Siskiyou, Santa Barbara,
  Sutter.
* more families below (canvass, flat sheets, SpreadsheetML, CSVs, LA zip)…

Run:
    uv run --with openpyxl python src/parse_general_2024_precincts.py [county ...]
"""

import csv
import re
import xml.etree.ElementTree as ET
import zipfile
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl
import xlrd

import pdfplumber

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from src.parsers2026.common import (  # noqa: E402
    is_summary_name, map_party, norm_name, parse_office, PARTY_CODE_MAP,
    OFFICE_RULES, _LOCAL_RE,
)

REPO_ROOT = Path(__file__).resolve().parent.parent
SOURCE_DIR = Path.home() / "code/openelections-sources-ca/2024/general"
OUT_DIR = REPO_ROOT / "2024/counties"
HEADERS = ["county", "precinct", "office", "district", "candidate",
           "party", "votes", "election_day", "mail"]
DATE = "20241105"


# ---------------------------------------------------------------------------
# shared helpers
# ---------------------------------------------------------------------------

PARTY_EXTRA = {"PAF": "Peace and Freedom", "AIP": "American Independent"}


def party_of(raw):
    """map_party plus 2024-only codes (PAF = Peace and Freedom)."""
    s = map_party(raw)
    return PARTY_EXTRA.get(s.strip().upper(), s)


def title_case(name):
    """Title-case an ALL-CAPS source candidate label; two-letter codes
    (JD, II, III...) stay upper; slash pairs keep their separator; the
    single-letter initials keep their period."""
    s = re.sub(r"\s+", " ", (name or "").strip()).strip(",")
    if not s or not s.upper() == s:
        return s
    keep_upper = {"II", "III", "IV", "JD"}
    parts = []
    for tok in s.split(" "):
        seg = tok.strip(",")
        seg = seg if (seg.endswith(".") and len(seg) <= 3) else seg.rstrip(".")
        if not seg:
            parts.append(tok)
            continue
        if "/" in seg:
            seg = "/".join(
                (p.upper() if p.upper() in keep_upper or len(p) == 1
                 else p.capitalize()) for p in seg.split("/"))
        elif re.fullmatch(r"[A-Z]\.", seg):
            seg = seg  # initial like "F." keeps its period
        elif seg.upper() in keep_upper:
            seg = seg.upper()
        elif seg.upper() in ("JR", "SR"):
            seg = seg.capitalize() + "."
        elif len(seg) == 1:
            seg = seg.upper()
        elif "-" in seg:
            seg = "-".join(p.capitalize() for p in seg.split("-"))
        else:
            seg = seg.capitalize()
        parts.append(seg)
    return " ".join(parts)


# Counties whose exports omit district numbers from district-office titles
# (e.g. Siskiyou's "STATE SENATOR (Vote for 1)") and the district the county
# sits entirely inside. Only referenced when the title carries no district.
COUNTY_DISTRICT_FALLBACK = {
    "Siskiyou": {"U.S. House": "1", "State Senate": "1", "State Assembly": "1",
                 "Board of Equalization": "2"},
    "Modoc": {"U.S. House": "1", "State Senate": "1", "State Assembly": "1",
              "Board of Equalization": "2"},
}
_WORD_ORDINALS = {"ONE": "1", "TWO": "2", "THREE": "3", "FOUR": "4", "FIVE": "5",
                  "SIX": "6", "SEVEN": "7", "EIGHT": "8", "NINE": "9",
                  "TEN": "10", "ELEVEN": "11", "TWELVE": "12"}


def parse_office2024(title, county=None):
    """(office, district) for a statewide contest; President and U.S. Senator
    added for the 2024 general; None for local/measure/judicial contests."""
    ti0 = str(title or "").upper()
    # word ordinals ("UNITED STATES REPRESENTATIVE DISTRICT ONE")
    for w, d in _WORD_ORDINALS.items():
        ti0 = re.sub(rf"\b{w}\b", d, ti0)
    t = ti0
    if re.search(r"\bPRESIDENT\b", t) and "PRO TEM" not in t \
            and "PRESIDING" not in t:
        return ("President", "")
    if re.search(r"U\.?\s?S\.?\s+SENAT|UNITED\s+STATES\s+SENAT", t):
        return ("U.S. Senate",
                "Partial Term" if re.search(r"UNEXPIRED|PARTIAL",
                                            t.replace(" ", "")) else "")
    if re.search(r"REPRESENTATIVE\s+(\d+)(?:ST|ND|RD|TH)\b", t):
        return ("U.S. House", re.search(r"REPRESENTATIVE\s+(\d+)",
                                        t).group(1))
    # Sacramento-style abbreviations: "CA-03", "SD-03", "AD-07" (the leading
    # "04 " sheet index is stripped by the caller)
    packed = t.replace(" ", "")
    ca = re.fullmatch(r"CA-?(\d+)", packed)
    sd = re.fullmatch(r"SD-?(\d+)", packed)
    ad = re.fullmatch(r"AD-?(\d+)", packed)
    if ca:
        return ("U.S. House", ca.group(1).lstrip("0"))
    if sd:
        return ("State Senate", sd.group(1).lstrip("0"))
    if ad:
        return ("State Assembly", ad.group(1).lstrip("0"))
    od = parse_office(ti0)
    if od:
        return od
    # district-less titles ("STATE SENATOR (Vote for 1)") match an office
    # rule only with the county's known single district
    fall = COUNTY_DISTRICT_FALLBACK.get(county or "", {})
    for pattern, office, needs_district in OFFICE_RULES:
        if needs_district and re.search(pattern, ti0) and office in fall:
            if _LOCAL_RE.search(ti0):
                return None
            return (office, fall[office])
    return None


def _num(cell):
    """Vote-cell value as int ('', '***', '****' redactions -> 0; commas ok)."""
    if cell is None:
        return 0
    s = re.sub(r"[\s,]", "", str(cell))
    if s in ("", "***", "****", "N/A", "-"):
        return 0
    try:
        return int(float(s))
    except ValueError:
        return 0


def write_county(display_name, rows, headers=HEADERS):
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    slug = re.sub(r"\s+", "_", display_name.strip().lower())
    path = OUT_DIR / f"{DATE}__ca__general__{slug}__precinct.csv"
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(headers)
        for r in rows:
            w.writerow([r.get(h, "") for h in headers])
    return path


# ---------------------------------------------------------------------------
# pageformat family ("Page: N of M" SOVC workbooks)
# Each sheet is one page of one contest (candidates paginate across sheets).
# Header row has col0 "Precinct"; candidate columns are "NAME\n(PARTY)" cells
# (party-less SPI headers handled bare). Two row layouts: single-row
# precincts (col0 = id, counts inline) and sub-row precincts (a col1-empty
# precinct-id row followed by per-channel rows and a "Total" row; the Total
# row is read when present, channels summed otherwise). Rollup blocks
# ("Electionwide", "Cumulative", district rollups) are skipped, and their
# "<x> - Total" rows feed the verification totals.
# ---------------------------------------------------------------------------

_PAGE_CAND_RE = re.compile(r"^(.+?)\s*\(([A-Z/\-&]+)\)\s*$", re.S)
_PAGE_CHANNEL_SET = {
    "vote center", "vote centers", "vote by mail", "vote-by-mail", "mail voting",
    "election day", "polling", "poll", "polls", "mail", "vbm", "early voting",
    "provisional", "in-person", "total",
}
_PAGE_ROLLUP_RE = re.compile(
    r"\b(electionwide|california|cumulative|county|countywide|state|total|"
    r"ballots?|registered|turnout|congressional|senatorial|assembly|"
    r"precincts|city|unincorporated|supervisori?al|board of equalization)\b")
_PAGE_SPI_SKIP_RE = re.compile(
    r"\b(write-?\s*in|unqualified|qualified|scattered|undervote|overvote|"
    r"total|ballots?|registered|turnout|blanks?|provisional|canvass|"
    r"electionwide|cumulative)\b", re.I)


def _is_page_rollup(s):
    s = str(s).strip().lower()
    if not s or s in _PAGE_CHANNEL_SET:
        return False
    if is_summary_name(s):
        return True
    return bool(_PAGE_ROLLUP_RE.search(s))


def _is_data_row(vt):
    if vt is None:
        return False
    if isinstance(vt, (int, float)):
        return True
    return bool(str(vt).strip()) and str(vt).strip().lower() != "precinct"


def _sheet_int(cell):
    try:
        return int(float(str(cell).strip()))
    except (ValueError, TypeError):
        return 0


def parse_pageformat(county, path):
    """Returns (rows, verify_failures). Sub-row channels are read via the
    precinct's Total row (all 2024 counties in this family have one), so
    only `votes` is emitted here (per-precinct channel splits are not
    recoverable from Total-only rows)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sums = defaultdict(int)
    totals = {}
    names = {}
    for sn in wb.sheetnames:
        ws = wb[sn]
        rows = list(ws.iter_rows(values_only=True))
        header_idx = None
        for i, r in enumerate(rows[:8]):
            if r and isinstance(r[0], str) and r[0].strip().lower() == "precinct":
                header_idx = i
                break
        if header_idx is None:
            continue
        contest = None
        for r in reversed(rows[:header_idx]):
            if not r:
                continue
            for cell in r[:3]:
                if isinstance(cell, str) and cell.strip():
                    od = parse_office2024(cell, county)
                    if od:
                        contest = od
                        break
            if contest:
                break
        if not contest:
            continue
        office, district = contest
        is_spi = office == "Superintendent of Public Instruction"
        cand_cols = []
        for c, cell in enumerate(rows[header_idx]):
            if not isinstance(cell, str):
                continue
            raw = cell.strip()
            m = _PAGE_CAND_RE.match(raw)
            if m:
                cand_cols.append((c, m.group(2).strip(), m.group(1).strip()))
                continue
            if not is_spi:
                continue
            cleaned = re.sub(r"\s+", " ", raw).strip()
            cl = cleaned.lower()
            if not cleaned or cl in ("precinct", "blanks", "percent", "%") \
                    or cl.endswith("%"):
                continue
            if is_summary_name(cleaned) or _PAGE_SPI_SKIP_RE.search(cleaned):
                continue
            cand_cols.append((c, "", cleaned))
        if not cand_cols:
            continue
        data_rows = rows[header_idx + 1:]
        has_total = any(r and isinstance(r[0], str)
                        and r[0].strip().lower() == "total" for r in data_rows)
        current = None
        cur_rollup = ""
        seen = set()
        block = {}  # (district, nn) -> accumulated value for current precinct
        for r in data_rows:
            if not r or r[0] is None:
                continue
            c0 = str(r[0]).strip()
            c0l = c0.lower()
            vt = r[1] if len(r) > 1 else None
            if _is_data_row(vt):
                if c0l in _PAGE_CHANNEL_SET:
                    # sub-row channel for the current precinct
                    if current is None:
                        continue
                    if has_total != (c0l == "total"):
                        continue
                    precinct = current
                elif _is_page_rollup(c0):
                    if COUNTY_ROW_RE.match(c0l):
                        for c, party_raw, name in cand_cols:
                            totals[(office, district, norm_name(name))] = \
                                _sheet_int(r[c]) if c < len(r) else 0
                    continue
                else:
                    # precinct-id header row. The district/geo-group rollup
                    # copies later in the sheet carry per-group partition
                    # values, so the FIRST full precinct block wins.
                    if c0 in seen:
                        continue
                    seen.add(c0)
                    precinct = current = c0
                    block = {}
                for c, party_raw, name in cand_cols:
                    v = _sheet_int(r[c]) if c < len(r) else 0
                    nn = norm_name(name)
                    if c0l in _PAGE_CHANNEL_SET and not has_total:
                        # no Total row: accumulate all channel rows
                        block[(district, nn)] = block.get((district, nn), 0) + v
                    else:
                        block[(district, nn)] = v
                    val = block[(district, nn)]
                    if val == 0 and name == "":
                        continue
                    key = (precinct, office, district, nn)
                    # precinct blocks repeat across a contest's pages with
                    # identical values; overwrite the block value, never sum
                    sums[key] = val
                    names[(office, district, nn)] = (name, party_raw)
            else:
                if _is_page_rollup(c0):
                    current = None  # rollup block: not a precinct
                    continue
                # sub-row precinct-id header row; district/geo-group reprints
                # of the same id carry per-group partition values, so the
                # first block wins
                if c0 in seen:
                    current = None  # geo-group reprint: orphan its rows
                    continue
                seen.add(c0)
                current = c0
                block = {}
    rows_out = []
    for (precinct, office, district, nn), v in sorted(sums.items(),
                                                      key=lambda k: (k[0][1], k[0][2], k[0][3], k[0][0])):
        name, party_raw = names[(office, district, nn)]
        rows_out.append({
            "county": county, "precinct": precinct, "office": office,
            "district": district, "candidate": title_case(name),
            "party": party_of(party_raw), "votes": v,
        })
    verify = _verify(county, "pageformat", sums, totals, names)
    return rows_out, verify


COUNTY_ROW_RE = re.compile(
    r"^(?:.*\bcounty\b|electionwide)\s*-\s*total$", re.I)


def _verify(county, family, sums, totals, names):
    """Per-contest-candidate: summed precinct votes vs captured county-total
    rows. Write-in aggregate differences surface as mismatches too — kept
    soft (reported, file still written) since those columns are dropped."""
    sums_by_cand = defaultdict(int)
    for k, v in sums.items():
        if len(k) == 4:
            _p, office, district, nn = k
        else:
            _p, (office, district), nn = k
        sums_by_cand[(office, district, nn)] += v
    failures = []
    for (office, district, nn), s in sums_by_cand.items():
        if (office, district, nn) not in totals:
            continue  # no captured total row for this contest
        t = totals[(office, district, nn)]
        if s != t:
            name = names[(office, district, nn)][0]
            failures.append(f"{office} {title_case(name)}: precincts {s} "
                            f"vs {t} in source total")
    return failures


# ---------------------------------------------------------------------------
# canvass family — "Precinct CANVASS" workbooks. One sheet per contest,
# candidates in PARTY - NAME headers (Kings/Madera/Yuba/Ventura) or
# NAME\n(PARTY) headers. Rows: single-row precincts (id + registered/cast +
# inline candidate counts — Ventura/Monterey/Tulare/Santa Cruz) or sub-row
# channel layouts (precinct-id header row, then Vote Center/Vote by Mail/
# Total rows — Kings/Madera/Yuba). A "Contest Total" row feeds verification.
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------

_CAND_PREF_RE = re.compile(r"^([A-Z][A-Z/\-&#.]{1,9})\s*-\s*(.+?)$", re.S)
_PARTY_CODES = {"DEM", "REP", "GRN", "LIB", "PF", "AI", "AIP", "NPP", "NP",
                "PE", "PAF", "REP/W-I", "DEM/W-I", "GRN/W-I", "LIB/W-I",
                "NPP/W-I", "PF/W-I", "AI/W-I"}


def _is_party_code(code):
    up = code.upper().replace("-", "").replace("W-I", "W/I")
    return up in _PARTY_CODES | set(PARTY_CODE_MAP)


_CAND_PAREN_RE = re.compile(r"^(.+?)\s*\(([A-Z/\-&#]{2,8})\)\s*$", re.S)
_ED_CHANNELS = {"election day", "election days", "vote center", "vote centers",
                "votecenter", "votecenters", "poll", "polls", "polling",
                "in person", "in-person", "remote tally center", "remote",
                "election day voting"}
_MAIL_CHANNELS = {"vote by mail", "vote-by-mail", "vbm", "mail", "mail voting",
                  "absentee", "early voting", "remote accessible vote-by-mail",
                  "vote by mail ballots"}
_CHANNEL_ALL = _ED_CHANNELS | _MAIL_CHANNELS | {"total"}
def sheet_rows(path):
    """Yield (sheet_name, row_iter) for openpyxl workbooks and BIFF .xls
    (xlrd) alike."""
    yield from _SHEET_LOADERS.get(path.suffix.lower(), _xlsx_rows)(path)


def _xlsx_rows(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for sn in wb.sheetnames:
        yield sn, wb[sn].iter_rows(values_only=True)


def _xls_rows(path):
    wb = xlrd.open_workbook(str(path))
    for sn in wb.sheet_names():
        sh = wb.sheet_by_name(sn)
        rows = [[sh.cell_value(r, c) for c in range(sh.ncols)]
                for r in range(sh.nrows)]
        yield sn, (tuple(r) for r in rows)


_SHEET_LOADERS = {".xlsx": _xlsx_rows, ".xls": _xls_rows}


_CONTEST_LABEL_RE = re.compile(r"^\s*\d+\s+\S")
G_ID_RE = re.compile(r"^(?:0*\d{3,7}|PCT\s*\d{2,5})\s*$", re.I)


def _parse_contest_label(label, county=None):
    """Split a canvass contest-label cell into {idx: (office, district)}."""
    subs = {}
    markers = list(re.finditer(r"(?:^|\s{2,})(\d+)\s+", label))
    for i, m in enumerate(markers):
        idx = m.group(1)
        start = m.end()
        end = markers[i + 1].start() if i + 1 < len(markers) else len(label)
        od = parse_office2024(label[start:end].strip(), county)
        if od:
            subs[idx.strip()] = od
    return subs


def parse_canvass(county, path):
    """Returns (rows, verify_failures); rows carry votes / election_day /
    mail where the sheet has channel rows, votes only otherwise. Handles
    stacked contest labels ("1 United States Senator Full Term    2 United
    States Senator Partial/...") via the tag row above the header."""
    votes = {}      # (precinct, od, nn) -> votes (Total row or channel sum)
    channels = {}   # -> [election_day, mail]
    totals = {}
    names = {}
    for sn, it in sheet_rows(path):
        rows = list(it)
        if not rows:
            continue
        # contest label(s): first cell matching "N <name>" in the top rows
        subs = {}
        for r in rows[:10]:
            if not r:
                continue
            for cell in r[:4]:
                if isinstance(cell, str) and cell.strip():
                    subs = _parse_contest_label(cell, county)
                    if not subs:
                        # bare title cell (no leading "N "):
                        # e.g. Tulare's "President-Vice President   (Vote for 1)"
                        od = parse_office2024(cell, county)
                        if od:
                            subs = {od[0]: od}
                    if subs:
                        break
            if subs:
                break
        if not subs:
            continue
        header_idx = None
        for i, r in enumerate(rows[:14]):
            if any(isinstance(c, str) and "REGISTERED VOT" in
                   " ".join(c.split()).upper() for c in r):
                header_idx = i
                break
        if header_idx is None:
            continue
        header = rows[header_idx]
        tag_row = rows[header_idx - 1] if header_idx > 0 else None
        default_idx = next(iter(subs)) if len(subs) == 1 else None
        cand_cols = []  # (col, party_raw, name, (office, district))
        for c, cell in enumerate(header):
            if c < 4 or not isinstance(cell, str):
                continue
            label = re.sub(r"\s+", " ", cell.strip())
            if not label or label.endswith("%") or label.endswith("(%)"):
                continue
            if is_summary_name(label):
                continue
            # stat-header columns ("Precinct" repeats in Tulare's doubled
            # stats block) are not candidates
            _l = label.lower()
            if re.fullmatch(r"(?:precinct|times? cast|registered voters|"
                            r"voters? cast|turnout(?: \(%\))?|total.*)", _l) \
                    or "electionwide" in _l or "cumulative" in _l:
                continue
            party_raw = ""
            m = _CAND_PREF_RE.match(label)
            if m and _is_party_code(m.group(1)):
                party_raw, label = m.group(1), m.group(2).strip()
            else:
                m = _CAND_PAREN_RE.match(label)
                if m:
                    party_raw, label = m.group(2), m.group(1).strip()
                elif not re.fullmatch(r"[A-Za-z][A-Za-z ./'\-]+", label) \
                        or any(s in label.lower() for s in (
                            "total", "voter", "registered", "ballot",
                            "turnout", "election", "cumulative", "precinct",
                            "write", "unresolved", "unqualified", "cast")):
                    # "Total - Vote Center" / "Total Votes" and friends are
                    # rollup columns, not candidates
                    continue

            tag = ""
            if tag_row is not None and c < len(tag_row) and \
                    tag_row[c] is not None:
                tag = str(tag_row[c]).strip()
            if tag in subs:
                pass
            elif tag and tag.isdigit() and int(tag) in [int(i) for i in subs
                                                        if i.isdigit()]:
                pass
            if tag and tag in subs:
                od = subs[tag]
            elif default_idx is not None:
                od = subs[default_idx]
            else:
                continue
            cand_cols.append((c, party_raw, label.split("  ")[0].strip(), od))
        if not cand_cols:
            continue
        data_rows = rows[header_idx + 1:]
        has_total = any(r and len(r) > 1 and isinstance(r[1], str)
                        and r[1].strip().lower() == "total" for r in data_rows)
        current = None
        acc = defaultdict(int)
        for r in data_rows:
            if not r or r[0] is None:
                continue
            c0 = str(r[0]).strip()
            vt = r[1] if len(r) > 1 else None
            vt_l = vt.strip().lower() if isinstance(vt, str) else ""
            if vt_l in _CHANNEL_ALL:
                # some canvasses carry the id on a bare id-header row;
                # Kings et al. put the id directly on the channel row
                if _is_id(c0):
                    if c0 != current:
                        current = c0
                        acc = defaultdict(int)
                elif current is None:
                    continue
                if has_total and vt_l != "total":
                    continue
                for c, party_raw, name, od in cand_cols:
                    v = _num(r[c]) if c < len(r) else 0
                    nn = norm_name(name)
                    key = (current, od, nn)
                    names[od + (nn,)] = (name, party_raw)
                    if has_total:
                        votes[key] = v
                    else:
                        ak = (od, nn)
                        acc[ak] += v
                        if vt_l in _ED_CHANNELS:
                            channels.setdefault(key, [0, 0])[0] += v
                        elif vt_l in _MAIL_CHANNELS:
                            channels.setdefault(key, [0, 0])[1] += v
                        votes[key] = acc[ak]
                continue
            if c0.lower() == "contest total" or COUNTY_ROW_RE.match(c0)                     or c0.lower().startswith("total -"):
                if c0.lower() == "contest total" or COUNTY_ROW_RE.match(c0):
                    for c, party_raw, name, od in cand_cols:
                        totals[od + (norm_name(name),)] = \
                            _num(r[c]) if c < len(r) else 0
                current = None
                continue
            if (vt_l == "" or isinstance(vt, (int, float))) and _is_id(c0):
                # single-row layout (values inline) or a bare id header
                vals = [str(r[c]).strip() for c, _, _, _ in cand_cols
                        if c < len(r)]
                if any(v not in ("", "None") for v in vals):
                    for c, party_raw, name, od in cand_cols:
                        nm = norm_name(name)
                        votes[(c0, od, nm)] = _num(r[c]) if c < len(r) else 0
                        names[od + (nm,)] = (name, party_raw)
                current = c0
                acc = defaultdict(int)
    rows_out = []
    for (precinct, (office, district), nn), v in votes.items():
        edv, mailv = channels.get(
            (precinct, (office, district), nn), [0, 0])
        name, party_raw = names[(office, district, nn)]
        rows_out.append({
            "county": county, "precinct": precinct, "office": office,
            "district": district, "candidate": title_case(name),
            "party": party_of(party_raw), "votes": v,
            "election_day": edv, "mail": mailv,
        })
    verify = _verify(county, "canvass", votes, totals, names)
    return rows_out, verify


def _is_id(c0):
    return bool(G_ID_RE.match(c0))




# ---------------------------------------------------------------------------
# Sacramento: flat per-contest sheets, header row 0 = Precinct + candidates;
# one row per precinct (channels not distinguished); "Write-in" aggregate col.
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Sacramento — flat per-contest sheets: sheet name = contest, header row 0 =
# Precinct + candidate columns ("Write-in" aggregate last, dropped).
# ---------------------------------------------------------------------------

def parse_sacramento(county, path):
    votes = {}
    names = {}
    totals = {}
    for sn, it in sheet_rows(path):
        rows = list(it)
        if len(rows) < 2:
            continue
        hdr = rows[0]
        if not (isinstance(hdr[0], str) and hdr[0].strip() == "Precinct"):
            continue
        od = parse_office2024(re.sub(r"^\d+\s*", "", str(sn)), county)
        if od is None:
            continue
        office, district = od
        cand_cols = []
        for c, h in enumerate(hdr[1:], start=1):
            if not isinstance(h, str) or not h.strip():
                continue
            label = " ".join(h.split())
            if is_summary_name(label) or "write-in" in label.lower():
                continue
            cand_cols.append((c, label))
        for r in rows[1:]:
            if not r or r[0] is None:
                continue
            pct = re.sub(r"\.0$", "", str(r[0]).strip())
            for c, name in cand_cols:
                v = _num(r[c]) if c < len(r) else 0
                nm = norm_name(name)
                votes[(pct, od, nm)] = v
                names[(office, district, nm)] = (name, "")
    rows_out = []
    for (precinct, (office, district), nn), v in votes.items():
        name, party = names[(office, district, nn)]
        rows_out.append({
            "county": county, "precinct": precinct, "office": office,
            "district": district, "candidate": title_case(name),
            "party": party, "votes": v,
        })
    return rows_out, []


# ---------------------------------------------------------------------------
# SpreadsheetML .xls family (Sonoma, Santa Clara): one sheet per contest;
# row0 = title, row1 = one candidate name per column group, row2 = header
# repeating Election Day / Vote By Mail / [Canvass] / Total Votes per
# candidate; data rows = Precinct | Registered | per-candidate columns; the
# trailing "Total:" row holds county totals (verification).
# ---------------------------------------------------------------------------

def _ssml_sheets(path):
    ns = "{urn:schemas-microsoft-com:office:spreadsheet}"
    with open(path, encoding="utf-8-sig") as f:
        root = ET.fromstring(f.read())
    for ws in root.findall(ns + "Worksheet"):
        rows = []
        for row in ws.findall(ns + "Table/" + ns + "Row"):
            vals = []
            idx = 0
            for cell in row.findall(ns + "Cell"):
                a = cell.get(ns + "Index")
                if a:
                    idx = int(a) - 1
                d = cell.find(ns + "Data")
                v = "" if d is None else (d.text or "")
                while len(vals) <= idx:
                    vals.append("")
                vals[idx] = v
                idx += 1
            rows.append(vals)
        yield ws.get(ns + "Name"), rows


def parse_ssml(county, path):
    votes = {}
    channels = {}
    totals = {}
    ssml_names = {}  # (office, district, nn) -> (display, "")
    verify = []
    for sn, rows in _ssml_sheets(path):
        title = next((s for s in (rows[0] if rows else []) if s.strip()), "")
        od = parse_office2024(title, county)
        if od is None or len(rows) < 4:
            continue
        office, district = od
        cand_names = [s.strip() for s in rows[1] if s and s.strip()]
        # column groups: runs in the header row ending at "Total Votes"
        seq = []
        for c, cell in enumerate(rows[2]):
            lab = " ".join(str(cell).split()).lower()
            if lab in ("election day", "vote by mail", "canvass", "total votes"):
                seq.append((c, lab))
        groups = []
        cur = []
        for c, lab in seq:
            cur.append((c, lab))
            if lab == "total votes":
                groups.append(cur)
                cur = []
        if len(groups) != len(cand_names):
            verify.append(f"! sheets mismatched for {sn}")
            continue
        for r in rows[3:]:
            if not r or not str(r[0]).strip():
                continue
            pct = re.sub(r"\.0$", "", str(r[0]).strip())
            if pct.lower().startswith("total"):
                for gi, g in enumerate(groups):
                    for c, lab in g:
                        if lab == "total votes":
                            totals[od + (norm_name(cand_names[gi]),)] = \
                                _num(r[c]) if c < len(r) else 0
                continue
            for gi, g in enumerate(groups):
                lv = {lab: (_num(r[c]) if c < len(r) else 0) for c, lab in g}
                nn = norm_name(cand_names[gi])
                key = (pct, od, nn)
                ssml_names[od + (nn,)] = (cand_names[gi], "")
                votes[key] = lv.get("total votes", 0)
                channels[key] = [lv.get("election day", 0),
                                 lv.get("vote by mail", 0)
                                 + lv.get("canvass", 0)]
    rows_out = []
    for (precinct, (office, district), nn), v in votes.items():
        edv, mailv = channels.get((precinct, (office, district), nn), [0, 0])
        rows_out.append({
            "county": county, "precinct": precinct, "office": office,
            "district": district, "candidate": title_case(ssml_names[(office, district, nn)][0]),
            "party": "", "votes": v, "election_day": edv, "mail": mailv,
        })
    verify = _verify(county, "ssml", votes, totals, ssml_names)
    return rows_out, verify


# ---------------------------------------------------------------------------
# San Mateo — flat certified-export CSV (one row per precinct x candidate).
# votes = total_votes; election_day = Vote Centers_votes; mail = Vote by
# Mail_votes. Turnout-only rows (Contest_title "Registered Voters") skipped.
# ---------------------------------------------------------------------------

def parse_san_mateo_csv(county, path):
    """Flat certified-export CSV, one row per precinct-SPLIT x candidate.
    Splits (Precinct_name '10009', Split_name '10009-1') partition a
    precinct, so rows aggregate to the precinct. votes = total_votes (== the
    Vote Centers + Vote by Mail votes on every row of this export);
    election_day = Vote Centers_votes; mail = Vote by Mail_votes. Named
    write-in candidates (Candidate_Type 'W') are kept; no-party candidates
    (blank Party_Code) get "Non-Partisan"."""
    agg = {}      # (precinct, (office, district), nn) -> [votes, ed, mail]
    names = {}
    parties = {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            title = (r.get("Contest_title") or "").strip()
            name = (r.get("candidate_name") or "").strip()
            ctype = (r.get("Candidate_Type") or "").strip()
            if not title or not name or ctype not in ("C", "W"):
                continue
            if is_summary_name(name) or "write-ins" in name.lower():
                continue
            od = parse_office2024(title, county)
            if od is None:
                continue
            pct = (r.get("Precinct_name") or "").strip()
            nn = norm_name(name)
            key = (pct, od, nn)
            cell = agg.setdefault(key, [0, 0, 0])
            cell[0] += _num(r.get("total_votes"))
            cell[1] += _num(r.get("Vote Centers_votes"))
            cell[2] += _num(r.get("Vote by Mail_votes"))
            names[od + (nn,)] = name
            code = (r.get("Party_Code") or "").strip()
            parties[od + (nn,)] = party_of(code) if code else "Non-Partisan"

    rows_out = []
    for (precinct, (office, district), nn), (v, edv, mailv) in agg.items():
        rows_out.append({
            "county": county, "precinct": precinct, "office": office,
            "district": district,
            "candidate": title_case(names[(office, district, nn)]),
            "party": parties[(office, district, nn)], "votes": v,
            "election_day": edv, "mail": mailv,
        })
    return rows_out, []


def _district_from_title(title):
    od = parse_office2024(title)
    return od[1] if od else ""


# ---------------------------------------------------------------------------
# Placer — flat tab-… no: comma-delimited txt, one row per (precinct, contest,
# choice, channel): p0 = precinct, p1 = contest id, p2 = contest title, p3 =
# choice label ("Registered Voters" / "Times Counted" / candidate / "WRITE
# IN"), p4 = channel (REMOTE / Vote by Mail / Vote Center / Total), p5 = votes.
# ---------------------------------------------------------------------------

def parse_placer_txt(county, path):
    votes = {}
    channels = {}
    names = {}
    times = defaultdict(int)  # (precinct, office-title) -> Times Counted Total
    with open(path, encoding="utf-8-sig", newline="") as f:
        for raw in csv.reader(f):
            if len(raw) < 6:
                continue
            pct = raw[0].strip()
            title = raw[2].strip()
            label = raw[3].strip()
            chan = raw[4].strip().lower()
            v = _num(raw[5])
            if not pct or not title or not label:
                continue
            lab_l = label.lower()
            if lab_l == "registered voters":
                continue
            if lab_l == "times counted":
                if chan == "total":
                    od = parse_office2024(title, county)
                    if od:
                        times[(pct, od)] = v
                continue
            if lab_l in ("write in", "write-in"):
                continue
            od = _placer_offices(title, county)
            if od is None:
                continue
            nn = norm_name(label)
            key = (pct, od, nn)
            if chan == "total":
                votes[key] = v
                names[nn] = label
            elif chan in ("vote center", "vote centers", "remote"):
                channels.setdefault(key, [0, 0])[0] += v
            elif chan in _MAIL_CHANNELS:
                channels.setdefault(key, [0, 0])[1] += v
    rows_out = []
    for (precinct, (office, district), nn), v in votes.items():
        edv, mailv = channels.get((precinct, (office, district), nn), [0, 0])
        rows_out.append({
            "county": county, "precinct": precinct, "office": office,
            "district": district, "candidate": title_case(names[nn]),
            "party": "", "votes": v, "election_day": edv, "mail": mailv,
        })
    # verify: per precinct x contest, candidate totals never exceed the
    # Times Counted Total cross-check (the export has no county summary)
    cand_sums = defaultdict(int)
    for (precinct, (office, district), nn), v in votes.items():
        cand_sums[(precinct, (office, district))] += v
    failures = []
    for (pct, od), t_tot in times.items():
        if cand_sums.get((pct, od), 0) > t_tot:
            failures.append(f"placer: {pct} {od[0]} cand sum "
                            f"{cand_sums.get((pct, od), 0)} > Times Counted "
                            f"{t_tot}")
    return rows_out, failures


def _placer_offices(title, county=None):
    return parse_office2024(title, county)
# ---------------------------------------------------------------------------
# Trinity — garbled SQL Server Reporting Services CSV export, precinct-major:
# the precinct restarts wherever the President contest reappears (block N =
# "Consolidation N" as in the SOV PDF). Columns: 0 contest title, 4 candidate
# (may be a running-mate pair on its own line), 5 party, ID1/ID2/ID3 votes =
# Election Day / Vote by Mail / Canvass, col12 = precinct total.
# Verified against Trinity's own summary CSV.
# ---------------------------------------------------------------------------

def parse_trinity_csv(county, precinct_path, summary_path):
    """Precinct-major garbled Reporting-Services CSV: a block per precinct
    begins wherever the President contest restarts; block N = Consolidation N.
    Columns: 0 title, 4 candidate, 5 party, ID1/ID2/ID3 = Election Day /
    Vote by Mail / Canvass votes, col12 = precinct total. The trailing
    Non-qualified-write-in rows are aggregates and are skipped; named '(W)'
    write-ins are kept. Verified against the county summary CSV (same
    export aggregate: col11 = name, col19 = county total)."""
    votes = {}  # (precinct, (office, district), nn) -> [ed, vbm, canvass, total]
    parties = {}
    precinct_num = 0
    prev_president = False
    with open(precinct_path, encoding="utf-8-sig", newline="") as fh:
        for row in csv.reader(fh):
            if len(row) < 13 or row[4] in ("", "ChoiceName"):
                continue
            title = row[0]
            name = " ".join(row[4].split())
            if name.lower().startswith("non-qualified write-in"):
                continue
            is_pres = "PRESIDENT" in title.upper()
            if is_pres and not prev_president:
                precinct_num += 1
            prev_president = is_pres
            od = parse_office2024(title, county)
            if od is None:
                continue
            try:
                ed, vbm, canv, total = (int(row[6].replace(",", "")),
                                        int(row[8].replace(",", "")),
                                        int(row[10].replace(",", "")),
                                        int(row[12].replace(",", "")))
            except ValueError:
                continue
            m = re.match(r"^(.*?)\s*\(W\)\s*$", name)
            if m:
                name = m.group(1).strip()
            nn = norm_name(name)
            votes[(f"Consolidation {precinct_num}", od, nn)] = [ed, vbm,
                                                                canv, total]
            parties[od + (nn,)] = row[5].strip()

    totals = {}
    failures = []
    with open(summary_path, encoding="utf-8-sig", newline="") as fh:
        for row in csv.reader(fh):
            if len(row) < 20 or row[11] in ("", "ChoiceName"):
                continue
            od = parse_office2024(row[0])
            if od is None:
                continue
            name = " ".join(row[11].split())
            m = re.match(r"^(.*?)\s*\(W\)\s*$", name)
            if m:
                name = m.group(1).strip()
            nn = norm_name(name)
            try:
                totals[od + (nn,)] = int(row[19].replace(",", ""))
            except ValueError:
                pass
    sums = defaultdict(int)
    for (_p, od, nn), vals in votes.items():
        sums[od + (nn,)] += vals[3]
    for key, s in sums.items():
        if key not in totals:
            failures.append(f"no summary total for {key}")
        elif s != totals[key]:
            failures.append(f"{key[0]} {key[2]}: precincts {s} vs summary "
                            f"{totals[key]}")

    rows_out = []
    for (precinct, (office, district), nn), vals in votes.items():
        ed, vbm, canv, total = vals
        rows_out.append({
            "county": county, "precinct": precinct, "office": office,
            "district": district, "candidate": title_case(nn),
            "party": party_of(parties[(office, district, nn)]),
            "votes": total, "election_day": ed, "mail": vbm + canv,
        })
    return rows_out, failures


# ---------------------------------------------------------------------------
# Los Angeles — one BIFF .xls per contest ("SVC ... Voter Nominated by
# Precinct") inside a zip archive. Extract to a temp dir, then parse the
# 2026-family SVC layout: row1 = contest title (row-major, ALL CAPS), row2 =
# headers (LOCATION/PRECINCT/…/BALLOTS CAST + candidates), 3 rows per
# precinct (POLLING PLACE / VBM PORTION / TOTAL); the TOTAL row is read.
# Write-ins are not precinct-attributable in this export (soft).
# ---------------------------------------------------------------------------

def parse_los_angeles(county, zip_path):
    import tempfile
    workdir = Path(tempfile.gettempdir()) / "la2024svc"
    if workdir.exists():
        import shutil
        shutil.rmtree(workdir)
    workdir.mkdir(parents=True)
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(workdir)
    votes = {}
    names = {}
    ballots = {}  # (precinct, od) -> ballots cast (Total rows)
    for xls in sorted(workdir.rglob("*.xls")):
        try:
            wb = xlrd.open_workbook(str(xls))
        except ValueError:
            continue  # not a BIFF file
        sh = wb.sheet_by_index(0)
        if sh.nrows < 3:
            continue
        od = parse_office2024(str(sh.cell_value(1, 0)).strip(), county)
        if od is None:
            continue
        office, district = od
        header = [str(sh.cell_value(2, c)).strip() for c in range(sh.ncols)]
        cand_cols = [c for c in range(8, sh.ncols) if header[c]
                     and not is_summary_name(header[c])]
        for r in range(3, sh.nrows):
            if str(sh.cell_value(r, 6)).strip() != "TOTAL":
                continue
            precinct = str(sh.cell_value(r, 1)).strip()
            if not precinct:
                continue
            try:
                ballots[(precinct, od)] = int(sh.cell_value(r, 7))
            except (TypeError, ValueError):
                ballot = None
            for c in cand_cols:
                name = header[c]
                try:
                    v = int(sh.cell_value(r, c))
                except (TypeError, ValueError):
                    v = 0
                votes[(precinct, od, norm_name(name))] = v
                names[(office, district, norm_name(name))] = (name, "")
    rows_out = []
    for (precinct, (office, district), nn), v in votes.items():
        name, party = names[(office, district, nn)]
        rows_out.append({
            "county": county, "precinct": precinct, "office": office,
            "district": district, "candidate": title_case(name),
            "party": party, "votes": v,
        })
    # cross-check: per precinct, candidate totals do not exceed ballots cast
    cand_sums = defaultdict(int)
    for (precinct, (office, district), nn), v in votes.items():
        cand_sums[(precinct, (office, district))] += v
    failures = [f"{p} {od[0]}: cand sum {s} > ballots {ballots.get((p, od))}"
                for (p, od), s in cand_sums.items()
                if (p, od) in ballots and s > ballots[(p, od)]]
    return rows_out, failures


# ---------------------------------------------------------------------------
# Dominion "Precinct Results Report / Final Precinct Report" PDF family
# (Nevada, Solano, Humboldt, Calaveras). Per contest block: a title line, a
# "Choice Party <roles...> Total" header, candidate rows (name [+party]
# followed by value/% pairs, last pair = Total), then Cast Votes / Undervotes /
# Overvotes /(Unresolved write-in votes) rows. Vote-group labels vary by
# county (Election Day / Vote Center / In-Person / Early / Vote by Mail /
# Provisional) and are read from the header. Strong verification: each block's
# Cast Votes total must equal the sum of its candidate totals.
# ---------------------------------------------------------------------------

_DOM_PRECINCT_RE = re.compile(
    r"^(\S+(?:\s?-\s?\S+)?)\s+(\d[\d,]*)\s+of\s+(\d[\d,]*)\s+registered voters", re.I)
# Mendocino-style: "<precinct name> - <precinct id> N of N registered voters"
_DOM_PRECINCT2_RE = re.compile(
    r"^(.+\S)\s+-\s+(\S+)\s+(\d[\d,]*)\s+of\s+(\d[\d,]+)\s+registered voters", re.I)
_DOM_ROLE_RE = re.compile(
    r"(election day voting|election day|early voting|vote center|in-person|"
    r"absentee|vote by mail|provisional)", re.I)
_DOM_PARTY_RE = re.compile(r"^([A-Z]{3,4})$")


def _dom_role(name):
    n = name.lower()
    if "vote by mail" in n or "absentee" in n:
        return "mail"
    if "provisional" in n:
        return "provisional"
    if "early" in n:
        return "early"
    return "election_day"  # Election Day / Vote Center / In-Person


_DOM_PARTY_TOKENS = {"DEM", "REP", "GRN", "LIB", "PF", "AI", "AIP", "NPP",
                     "PFP", "PAF", "PSL", "NP", "WF", "IND"}


def parse_dominion_canvass(county, path):
    """Dominion per-contest blocks; returns (rows, verify_failures). Rows
    carry votes + early_voting / election_day / provisional / mail breakdowns
    where the source distinguishes them. Verification: each contest block's
    Cast Votes total (last Cast Votes number) must equal the sum of its
    candidate totals."""
    votes = {}       # (precinct, (office, district), nn) -> [total, per-role]
    names = {}
    precinct = None
    cur_roles = ()
    cur_od = None
    cur_title = ""
    block_sum = 0
    last_key = None  # of the candidate row just read (running-mate lines)
    last_name = ""
    last_party = ""
    last_per = None
    seen_precinct = False  # county-summary pages precede the precinct report
    failures = []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            for line in (page.extract_text() or "").splitlines():
                line = line.strip()
                if not line:
                    continue
                m = _DOM_PRECINCT_RE.match(line) or \
                    _DOM_PRECINCT2_RE.match(line)
                if m:
                    if m.group(1).strip() != precinct:
                        precinct = m.group(1).strip()
                    cur_od = None
                    seen_precinct = True
                    continue
                low = line.lower()
                if low.startswith("choice party"):
                    cur_roles = tuple(_dom_role(x) for x in
                                      _DOM_ROLE_RE.findall(line[12:]))
                    continue
                if low.startswith("cast votes:"):
                    nums = [int(x.replace(",", "")) for x in re.findall(
                        r"\d[\d,]*(?=\s|$)", line.split(":", 1)[1])]
                    expected = nums[-1] if nums else 0
                    if seen_precinct and cur_od is not None and expected \
                            and block_sum != expected:
                        failures.append(
                            f"{county} {precinct} {cur_title[:44]}: candidates "
                            f"{block_sum} vs Cast Votes {expected}")
                    block_sum = 0
                    last_key = None
                    continue
                if low.startswith(("undervotes:", "overvotes:",
                                   "unresolved write-in",
                                   "unqualified write")):
                    continue
                if "%" not in line and not low.startswith((
                        "cast votes", "undervotes", "overvotes",
                        "choice party", "unresolved write-in",
                        "unqualified write")):
                    if last_key is not None and cur_od is not None and \
                            parse_office2024(line) is None and \
                            len(line.split()) <= 4:
                        # running-mate continuation of the previous
                        # candidate row ("JD Vance" under "Donald J. Trump")
                        votes.pop(last_key, None)
                        mate = re.sub(r"^\s*and\s+", "", line,
                                      flags=re.I)
                        nn = norm_name(last_name + " / " + mate)
                        key = (last_key[0], cur_od, nn)
                        votes[key] = [last_total, last_per]
                        names[cur_od + (nn,)] = (last_name + " / " + mate,
                                                 last_party)
                        last_key = key
                        continue
                    # contest-title line ("Vote for ONE (1)" suffixes carry
                    # digits, so no digit test here)
                    cur_title = line
                    cur_od = parse_office2024(line, county)
                    block_sum = 0
                    last_key = None
                    continue
                if cur_od is None or precinct is None or "%" not in line:
                    continue
                nums = []
                name_toks = []
                party = ""
                for tok in line.split():
                    if tok.endswith("%"):
                        continue
                    if re.match(r"^[\d,]+$", tok):
                        nums.append(_num(tok))
                    elif not name_toks:
                        name_toks.append(tok)
                    elif tok.upper().replace("/", "") in _DOM_PARTY_TOKENS \
                            and not nums:
                        party = tok
                    else:
                        name_toks.append(tok)
                if not nums:
                    continue
                name = " ".join(name_toks).strip()
                if len(nums) < 2 or len(nums) - 1 > len(cur_roles):
                    failures.append(f"{county} {precinct}: odd row "
                                    f"'{line[:56]}'")
                    continue
                total = nums[-1]
                m_wi = re.match(r"^(.*?)\s*\(W\)\s*$", name)
                if m_wi:
                    name = m_wi.group(1).strip()
                nn = norm_name(name)
                per = {"early": 0, "election_day": 0, "provisional": 0,
                       "mail": 0}
                for role, v in zip(cur_roles, nums[:-1]):
                    per[role] = per.get(role, 0) + v
                key = (precinct, cur_od, nn)
                votes[key] = [total, per]
                names[cur_od + (nn,)] = (name, party)
                block_sum += total
                last_key = key
                last_total = total
                last_per = per
                last_name = name
                last_party = party
    rows_out = []
    for (precinct, (office, district), nn), (v, per) in votes.items():
        name, party = names[(office, district, nn)]
        rows_out.append({
            "county": county, "precinct": precinct, "office": office,
            "district": district, "candidate": title_case(name),
            "party": party_of(party) if party else "",
            "votes": v,
            "early_voting": per.get("early", 0),
            "election_day": per.get("election_day", 0),
            "provisional": per.get("provisional", 0),
            "mail": per.get("mail", 0),
        })
    return rows_out, failures


# ---------------------------------------------------------------------------
# Sierra — 180-degree-rotated "Canvass Results Report" PDF. Candidate names
# are char-reversed in the rotated header band ("ESROM" = "MORSE"); precinct
# labels are upright at the far left; integer values belong to the header
# column at/left of the value. Ported from the 2026 parse_rotated180_pdf
# without the county-file candidate index: candidate columns are recovered
# directly from their (reversed) header names.
# ---------------------------------------------------------------------------

_ROT_SUMMARY_RE = re.compile(
    r"\b(votes|cast|under|over|registered|turnout|ballots|total|election day|"
    r"vote by mail|provisional|write)\b", re.I)


def parse_rotated180_pdf(county, path, start_page=0):
    votes = {}
    names = {}
    rot_totals = {}
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages[start_page:]:
            words = page.extract_words(extra_attrs=["upright"])
            lines = defaultdict(list)
            for w in words:
                lines[round(w["top"])].append(w)
            contest = None
            office_top = None
            for y in sorted(lines):
                txt_y = " ".join(w["text"] for w in
                                 sorted(lines[y], key=lambda w: w["x0"])).strip()
                if not txt_y or "canvass results" in txt_y.lower() or \
                        "official results" in txt_y.lower():
                    continue
                od = parse_office2024(txt_y)
                if od:
                    contest = od
                    office_top = y
                    break
            if not contest:
                continue
            office, district = contest
            data_top = None
            for y in sorted(lines):
                if y <= office_top:
                    continue
                txt_y = " ".join(w["text"] for w in
                                 sorted(lines[y], key=lambda w: w["x0"]))
                if txt_y.strip().lower() == "precinct":
                    data_top = y
                    break
            if data_top is None:
                data_top = office_top + 140
            band = [w for w in words
                    if w["x0"] > 100 and office_top + 8 < w["top"] < data_top]
            # cluster band words by word CENTER: each word of a rotated
            # column is centered on its column, so centers are stable even
            # when word x0/widths differ
            groups = defaultdict(list)
            for w in band:
                groups[round((w["x0"] + w["x1"]) / 2)].append(w)
            raw_cols = []
            for x in sorted(groups):
                g = groups[x]
                if raw_cols and x - raw_cols[-1][0] < 14:
                    g = raw_cols[-1][1] + g
                    raw_cols[-1] = [sum((w["x0"] + w["x1"]) / 2
                                        for w in g) / len(g), g]
                else:
                    raw_cols.append([x, list(g)])
            columns = []
            for cx, g in raw_cols:
                # within a column the name may wrap into two ~9pt sub-lines
                # (Donald J. Trump | JD Vance); read each sub-line bottom-up,
                # lines left-to-right, joining multi-line names with " / "
                sublines = defaultdict(list)
                for w in g:
                    sublines[round(w["x0"])].append(w)
                pieces = []
                for sx in sorted(sublines):
                    ws = sorted(sublines[sx], key=lambda w: -w["top"])
                    pieces.append(" ".join(w["text"][::-1] for w in ws))
                name = " / ".join(pieces) if len(pieces) > 1 else pieces[0]
                if not name or any(s in name.lower() for s in (
                        "votes", "cast", "under", "over", "registered",
                        "turnout", "ballots", "total", "write")):
                    columns.append((cx, "meta", None))
                    continue
                disp = re.sub(r"\s+", " ", name).strip()
                m_wi = re.match(r"^(.*?)\s*\(W\)\s*$", disp)
                if m_wi:
                    disp = m_wi.group(1).strip() + "(W)"
                columns.append((cx, "cand", disp))
            if not any(c[1] == "cand" for c in columns):
                continue
            columns.sort(key=lambda c: c[0])
            for y in sorted(lines):
                if y <= data_top:
                    continue
                line = sorted(lines[y], key=lambda w: w["x0"])
                pct_words = [w for w in line if w["x0"] < 110]
                if not pct_words:
                    continue
                precinct = " ".join(w["text"] for w in pct_words).strip()
                if precinct.lower() == "precinct":
                    continue
                totals_row = precinct.lower() in ("totals", "total")
                if is_summary_name(precinct) and not totals_row:
                    break  # boilerplate; the rest of the page is county stats
                for w in line:
                    if w["x0"] < 110:
                        continue
                    s = w["text"].replace(",", "").rstrip("%")
                    if "." in s:
                        continue
                    try:
                        v = int(s)
                    except ValueError:
                        continue
                    owner = None
                    for col_x0, kind, disp in columns:
                        if col_x0 <= w["x0"] + 5:
                            owner = (kind, disp)
                        else:
                            break
                    if not owner or owner[0] != "cand":
                        continue
                    disp = owner[1]
                    nn = norm_name(disp)
                    if totals_row:
                        rot_totals[contest + (nn,)] = v
                        continue
                    key = (precinct, contest, nn)
                    # repeated precinct rows per page copy: overwrite
                    votes[key] = v
                    names[contest + (nn,)] = (disp, "")
    rows_out = []
    for (precinct, (office, district), nn), v in votes.items():
        name, party = names[(office, district, nn)]
        rows_out.append({
            "county": county, "precinct": precinct, "office": office,
            "district": district, "candidate": title_case(name),
            "party": party, "votes": v,
        })
    verify = _verify(county, "rotated180", votes, rot_totals, names)
    return rows_out, verify


# ---------------------------------------------------------------------------
# Imperial — printed "Page: N of M" SOVC PDF (single-row precinct pages).
# Each contest prints its precinct list over pairs of pages: page A carries
# the Times-Cast/Registered stats block plus the first candidates, page B the
# remaining candidates, a "Total Votes" column (= sum of that precinct's
# candidate votes in this contest, used for verification) and write-in /
# unresolved write-in columns. Candidate headers are 90-degree-rotated words
# (each word's characters run back-to-front in the text stream); the name
# words sit on one vertical line and the party code on a second line ~10.6pt
# to its right. Vote values are right-aligned numbers ~90pt apart; a value is
# assigned to the nearest candidate label within half that pitch.
# ---------------------------------------------------------------------------

_PF_STATS_RE = re.compile(
    r"^(?:times? cast|registered|voters?|turnout|undervotes|overvotes|"
    r"ballots cast)$", re.I)
_PF_DROP_RE = re.compile(r"unresolved|write-in$|\bvoters?$", re.I)
_PF_ID_RE = re.compile(r"^\d{1,7}(?:MB)?(-\d+)?\s*$|^\d{1,7}\s?-\s?\S.*$")


DEBUG_PF = False
_PF_CHANNELS = {"mail ballot", "designated mail ballot", "polling place",
                "provisional", "total", "vote centers", "vote center",
                "vote by mail", "election day", "mail", "polling"}
def _channel_prefix(low):
    """True when a line's leading words name a channel row (everything before
    the first number, e.g. "designated mail ballot 0 0 0")."""
    pre = []
    words = low.split()
    for w in words:
        if not re.fullmatch(r"[a-z]+", w):
            break  # numbers and **** redactions end the label words
        pre.append(w)
    return " ".join(pre) in _PF_CHANNELS if pre else False


_PF_IDLINE_RE = re.compile(r"^\s*-\s*([0-9]{2,7})\s*(?:-.*)?$")
# San Bernardino ballot-style codes ("C23S19A34E1B1AVUAP10001"): a short
# uppercase head that includes a digit early, then a long tail. (A plain
# minimum-length pattern would also match long titles like REPRESENTATIVE.)
_PF_CODE_RE = re.compile(r"^[A-Z]{1,4}[0-9][0-9A-Z\-]{9,}$")


def parse_pageformat_pdf(county, path):
    votes = {}    # (precinct, od, nn) -> votes (Total row or inline cell)
    channels = {}  # -> [election-day-ish, mail-ish] channel votes
    names = {}
    totals = {}
    failures = []
    prev_od = None
    current = None   # precinct id of the sub-row block being read
    pending = defaultdict(lambda: [0, 0, 0, 0])
    pending_code = None
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            words = page.extract_words(extra_attrs=["upright"])
            up = [w for w in words if w.get("upright", True)]
            rot = [w for w in words if not w.get("upright", True)]
            lines = defaultdict(list)
            for w in up:
                lines[round(w["top"])].append(w)
            # the contest title prints only on each contest's first page;
            # continuation pages carry the previous contest. Header/rollup
            # lines (Electionwide / Cumulative / bare precinct labels) are
            # ignored -- not treated as titles.
            od = None
            titled = False
            for y in sorted(lines)[:10]:
                txt = " ".join(x["text"] for x in
                               sorted(lines[y], key=lambda w: w["x0"])).strip()
                low = txt.lower()
                if re.search(r"page:\s*\d+\s+(?:of\s+)?\d*", txt, re.I) or \
                        re.match(r"^\d{1,2}/\d{1,2}/\d{2,4}\b", txt) or \
                        "precinct results" in low or "official results" in low \
                        or "statement of vote" in low or \
                        re.fullmatch(r"precinct( precinct)?", low) or \
                        "electionwide" in low or "cumulative" in low or \
                        "insufficient turnout" in low or \
                        _channel_prefix(low) or \
                        _PF_STATS_RE.match(low) or \
                        low.startswith("state of california"):
                    continue
                # a title line begins with alpha text and carries several
                # alpha words (precinct rows start with ids; ballot-style
                # code rows -- often printed twice per line -- are not titles)
                if len([w for w in txt.split() if re.search(r"[A-Za-z]{2}", w)]) < 2 \
                        or not re.match(r"^[A-Za-z]", txt) \
                        or any(_PF_CODE_RE.match(w) for w in txt.split()):
                    continue
                titled = True
                title_txt = txt
                od = parse_office2024(txt, county)
                break
            if titled and DEBUG_PF and page.page_number > 3:
                print(f"TITLED{page.page_number}:[{title_txt[:44]}]->{od}", flush=True)
            if titled:
                page_od = od  # None for local/measure pages: skip + reset
            else:
                page_od = prev_od  # continuation page
            prev_od = page_od
            pending = defaultdict(lambda: [0, 0, 0, 0])
            if page_od is None:
                continue
            od = page_od
            # rotated header columns, clustered by x0. A candidate's name
            # wraps onto a second vertical line ~10.6pt right of the first
            # (either the rest of the name or the "(PARTY)" code)
            cols = defaultdict(list)
            for w in rot:
                cols[w["x0"]].append(w)
            labels = []
            total_x = None
            consumed = set()
            xs = sorted(cols)
            for x0 in sorted(cols):
                if x0 in consumed:
                    continue
                # a candidate's rotated label spans several ~10.6pt
                # vertical lines: name line(s), wrapped name, "(PARTY)"
                block = []
                prevx = None
                for x in xs:
                    if x < x0 or x in consumed:
                        continue
                    if x != x0 and (prevx is None or x - prevx > 12):
                        break
                    block.append(x)
                    prevx = x
                consumed.update(block)
                pieces = []
                party = ""
                for x in block:
                    ptext = " ".join(
                        w["text"][::-1] for w in
                        sorted(cols[x], key=lambda w: -w["top"])).strip()
                    m = re.match(r"^\(([A-Z/\-&]+)\)$", ptext)
                    if m and pieces:
                        party = m.group(1)
                    else:
                        pieces.append(ptext)
                text = " ".join(pieces).strip()
                low = text.lower()
                if _PF_STATS_RE.match(low):
                    continue
                if _PF_DROP_RE.search(low):
                    continue
                if low.startswith("total votes"):
                    total_x = x0
                    continue
                labels.append((x0, text, party))
            if not labels:
                if DEBUG_PF:
                    print("SKIP", flush=True)
                continue
            if DEBUG_PF:
                print(f"PAGE{page.page_number}L{len(labels)}T{bool(total_x)}:"
                      f"{[l[1][:22] for l in labels]}", flush=True)

            def assign(w):
                vc = (w["x0"] + w["x1"]) / 2
                best, bd = None, 50.0
                for x0, text, party in labels:
                    lc = x0 + 4
                    d = vc - lc
                    if -30 <= d <= 50 and abs(d) < bd:
                        best, bd = (text, party), d
                return best

            def take(lab, tok, wname):
                """Store one assigned numeric cell for the given precinct."""
                name, party = lab
                if re.fullmatch(r"unresolved write.?in", name, re.I):
                    return
                # "PETER SONSKI Qualified Write In": a named write-in
                # pair carries the qualifier as a second line
                name = re.sub(r"\s*qualified write in\s*$", "", name,
                              flags=re.I)
                m_wi = re.match(r"^(.*?)\s+\(W\)\s*$", name, re.I)
                if m_wi:
                    name = m_wi.group(1).strip()
                nm = norm_name(name)
                key = (wname, od, nm)
                names[od + (nm,)] = (name, party)
                return key

            # channel label = the words before the row's first number
            def row_channel(line):
                lab = []
                for x in line:
                    if re.match(r"^[\d,]+$", x["text"]):
                        break
                    lab.append(x["text"].lower())
                name = " ".join(lab)
                return name if name in _PF_CHANNELS else ""

            for y in sorted(lines):
                line = sorted(lines[y], key=lambda w: w["x0"])
                row_id = line[0]["text"].strip()
                line_l = " ".join(x["text"] for x in line).lower()
                # id line of a San-Bernardino-style code block ("- 0037")
                m_id = _PF_IDLINE_RE.match(line_l)
                # the doc prints each id line with a right-block copy
                # (e.g. "- 0001 - 0001" at x 25 and x 405); only numeric
                # words in the LEFT half (stat/value columns, 130-480)
                # disqualify it as an id line
                if m_id and len(line) <= 6 and \
                        not any(re.match(r"^[\d,]+$", x["text"])
                                and 130 < x["x0"] < 480 for x in line[1:]):
                    current = m_id.group(1)
                    pending = defaultdict(lambda: [0, 0, 0, 0])
                    continue
                ch_l = row_channel(line)
                if ch_l:
                    if current is None:
                        continue
                    is_total = ch_l == "total"
                    for w in line[1:]:
                        tok = w["text"]
                        if tok.endswith("%") or not re.match(r"^[\d,]+$", tok):
                            continue
                        lab = assign(w)
                        if lab is None:
                            continue
                        key = take(lab, tok, current)
                        if not key:
                            continue
                        v = _num(tok)
                        if is_total:
                            p = pending.pop(key, None) or [0, 0, 0, 0]
                            # the doc's Total row carries the full block
                            # total (already including all channel rows), so
                            # never add the channel accumulation to it
                            votes[key] = v
                            channels[key] = [p[1], p[2], p[3]]
                        else:
                            pen = pending[key]
                            pen[0] += v
                            if ch_l in _ED_CHANNELS or ch_l == "polling place":
                                pen[1] += v
                            elif ch_l == "provisional":
                                pen[2] += v
                            else:
                                pen[3] += v
                    if is_total:
                        pending = defaultdict(lambda: [0, 0, 0, 0])
                    continue
                if _PF_CODE_RE.match(row_id):
                    pending_code = row_id
                    continue
                if not _PF_ID_RE.match(row_id):
                    continue
                for w in line[1:]:
                    tok = w["text"]
                    if tok.endswith("%") or "." in tok:
                        continue
                    if not re.match(r"^[\d,]+$", tok):
                        continue
                    lab = assign(w)
                    if lab is None:
                        continue
                    key = take(lab, tok, row_id)
                    if not key:
                        continue
                    votes[key] = votes.get(key, 0) + _num(tok)
                # verification: "Total Votes" cell equals the candidate sum
                if total_x is not None:
                    for w in line[1:]:
                        vc = (w["x0"] + w["x1"]) / 2
                        if vc and abs(vc - (total_x + 4)) <= 20 and \
                                re.match(r"^[\d,]+$", w["text"]):
                            totals[od + ("__total__", row_id)] = \
                                _num(w["text"])
    # verify per (contest, precinct): sum(candidates) vs Total Votes cell
    sums = defaultdict(int)
    for (pct, od, nm), v in votes.items():
        sums[od + ("__total__", pct)] += v
    for (office, district, _t, pct), s in sums.items():
        key = (office, district, _t, pct)
        if key in totals and totals[key] and s != totals[key]:
            failures.append(f"{office} {pct}: candidates {s} vs "
                            f"Total Votes {totals[key]}")
    rows_out = []
    for (precinct, (office, district), nn), v in votes.items():
        name, party = names[(office, district, nn)]
        row = {
            "county": county, "precinct": precinct, "office": office,
            "district": district, "candidate": title_case(name),
            "party": party_of(party), "votes": v,
        }
        ed, prov, mailv = channels.get((precinct, (office, district), nn),
                                       [None, None, None])
        if ed is not None:
            row["election_day"] = ed
            row["provisional"] = prov
            row["mail"] = mailv
        rows_out.append(row)
    return rows_out, failures


# ---------------------------------------------------------------------------
# Tuolumne — clean per-contest CSVs (one file per contest, candidates as
# column headers "Name / Running Mate\nParty"); Totals row = county totals
# (verification). Statewide contests are picked from the file name.
# ---------------------------------------------------------------------------

def parse_tuolumne(county, src_dir):
    files = sorted(src_dir.glob("Tuolumne-2024-Nov5-General-*-by-precinct.csv"))
    agg = {}      # (precinct, od, nn) -> votes
    names = {}
    failures = []
    totals = {}
    for f in files:
        parts = f.name.split("Tuolumne-2024-Nov5-General-")[1]
        parts = parts.split("-by-precinct.csv")[0]
        title = re.sub(r"\s+", " ", parts.replace("-", " ")).strip()
        od = parse_office2024(title, county)
        if od is None:
            continue  # local offices / measures
        office, district = od
        c_totals = {}
        with open(f, encoding="utf-8-sig", newline="") as fh:
            rd = csv.reader(fh)
            try:
                header = next(rd)
            except StopIteration:
                continue
            cand_cols = []
            for c, cell in enumerate(header):
                label = " ".join(cell.split())
                if c == 0 or not label:
                    continue
                if is_summary_name(label) or re.search(
                        r"total votes cast|undervotes|overvotes|"
                        r"total ballots|registered voter", label, re.I):
                    continue
                # candidate cols: "Donald J. Trump/ JD Vance/Republican"
                # (party as the trailing segment)
                nm = label
                party = ""
                m = re.match(r"^(.*?)\s*(?:/\s*)?(Republican|Democratic|"
                             r"Green|Libertarian|Peace and Freedom|"
                             r"American Independent|No Party Preference|"
                             r"Non-Partisan)\s*$", label)
                if m:
                    nm = re.sub(r"\s*/\s*$", "", m.group(1)).strip()
                    party = party_of(m.group(2))
                cand_cols.append((c, nm))
                names[od + (norm_name(nm),)] = (nm, party or "")
            if not cand_cols:
                continue
            for r in rd:
                if not r or not r[0].strip():
                    continue
                prec = r[0].strip()
                for c, nm in cand_cols:
                    v = _num(r[c]) if c < len(r) else 0
                    nn = norm_name(nm)
                    if prec.lower() == "totals":
                        c_totals[nn] = v
                        continue
                    agg[(prec, od, nn)] = v
        for nn, v in c_totals.items():
            if nn in (0, None):
                continue
            totals[od + (nn,)] = v
            s = sum(x for (p, o, n), x in agg.items()
                    if o == od and n == nn)
            if s != v:
                failures.append(f"{office} {names[od + (nn,)][:28]}: "
                                f"precincts {s} vs Totals row {v}")

    rows_out = []
    for (precinct, (office, district), nn), v in agg.items():
        name, party = names[(office, district, nn)]
        rows_out.append({
            "county": county, "precinct": precinct, "office": office,
            "district": district,
            "candidate": title_case(name).replace(" / ", "/"),
            "party": party, "votes": v,
        })
    return rows_out, failures


# ---------------------------------------------------------------------------
# Modoc — 2024's "Summary Results Report" precinct-major PDF (the 2026 family,
# ref-free variant): "<county> County" page header, then a bare precinct-name
# line; contest titles sit directly above "Vote For N"; candidate rows are
# "<name ...> <total>" with only a TOTAL column (no vote-type breakdown,
# no party token). "STATISTICS"/"Write-In Totals" sections are skipped.
# ---------------------------------------------------------------------------

def parse_modoc24(county, path):
    import pdfplumber
    header_re = re.compile(rf"\b{re.escape(county)}\s+County\b", re.I)
    votefor_re = re.compile(r"^Vote\s+For\s+\d+$", re.I)
    num_re = re.compile(r"^[\d,]+$")
    pdf = pdfplumber.open(str(path))
    agg = {}
    names = {}
    cur_precinct = None
    cur_od = None
    prev_t = ""
    expect_precinct = False
    for p in pdf.pages:
        words = p.extract_words()
        lines = defaultdict(list)
        for w in words:
            lines[round(w["top"])].append(w)
        for y in sorted(lines):
            toks = [w["text"] for w in sorted(lines[y], key=lambda w: w["x0"])]
            tt = " ".join(toks).strip()
            if not tt:
                continue
            if expect_precinct:
                cur_precinct = tt
                cur_od = None
                expect_precinct = False
                prev_t = ""
                continue
            if header_re.search(tt):
                expect_precinct = True
                continue
            if votefor_re.match(tt):
                cur_od = parse_office2024(prev_t, county)
                prev_t = ""
                continue
            if cur_precinct is not None and cur_od is not None \
                    and len(toks) >= 2 and num_re.match(toks[-1]):
                name = " ".join(toks[:-1]).strip()
                if name and "write-in" not in name.lower() \
                        and not is_summary_name(name):
                    nn = norm_name(name)
                    agg[(cur_precinct, cur_od, nn)] = _num(toks[-1])
                    names[cur_od + (nn,)] = name
            prev_t = tt
    pdf.close()

    # verify: per contest, per precinct, candidate sums may not exceed the
    # STATISTICS "Ballots Cast - Total" for that precinct (this is an
    # unofficial summary with no county contest-total rows)
    cast = {}
    cur = None
    expect = False
    with pdfplumber.open(str(path)) as pdf:
        for p in pdf.pages:
            for line in (p.extract_text() or "").splitlines():
                tt = line.strip()
                if header_re.search(tt):
                    expect = True
                    continue
                if expect:
                    cur = tt
                    expect = False
                    continue
                m = re.match(r"Ballots Cast - Total (\d[\d,]*)", tt)
                if m and cur:
                    cast[cur] = _num(m.group(1))
    failures = []
    cast_by_precinct = {}
    prec_sums = {}
    cand_cast = {}
    with pdfplumber.open(str(path)) as pdf:
        for (prec, od, nn), v in agg.items() if False else []:
            pass
    cand_sums = defaultdict(int)
    for (prec, od, nn), v in agg.items():
        cand_sums[(prec, od[0])] += v
    for (prec, oname), s in cand_sums.items():
        bc = cast.get(prec)
        if bc is not None and s > bc:
            failures.append(f"{oname} {prec}: candidates {s} > "
                            f"Ballots Cast {bc}")

    rows_out = []
    for (precinct, (office, district), nn), v in agg.items():
        rows_out.append({
            "county": county, "precinct": precinct, "office": office,
            "district": district,
            "candidate": title_case(names[(office, district, nn)]),
            "party": "", "votes": v,
        })
    return rows_out, failures


SOURCE_FILES = {
    "colusa": ("Colusa", "pageformat",
               "2024 Colusa County, CA precinct-level results/Colusa-2024-General-SOVC.xlsx"),
    "siskiyou": ("Siskiyou", "pageformat",
                 "2024 Siskiyou County, CA precinct-level results.xlsx"),
    "marin": ("Marin", "pageformat",
              "Marin County - Statement of the Vote - General 2024.xlsx"),
    "contra_costa": ("Contra Costa", "pageformat",
                     "Contra Costa County - Precinct Results - General 2024.xlsx"),
    "fresno": ("Fresno", "pageformat",
               "2024 Fresno County, CA precinct-level results.xlsx"),
    "san_francisco": ("San Francisco", "pageformat",
                      "2024 San Francisco, CA precinct-level results.xlsx"),
    "santa_barbara": ("Santa Barbara", "pageformat",
                      "2024 Santa Barbara County, CA precinct-level results/sov-pct.xlsx"),
    "inyo": ("Inyo", "pageformat",
             "2024 Inyo County, CA precinct-level results/Inyo-2024-General-SOVC.xlsx"),
    "sutter": ("Sutter", "pageformat",
               "2024 Sutter County, CA precinct-level results/Sutter-2024-General-SOV-by-Precinct.xlsx"),
    "san_benito": ("San Benito", "pageformat",
                    "2024 San Benito County, CA precinct-level results.xlsx"),
    # NOTE: the "2024 Kings County" source workbook is mislabeled — its
    # contents are Madera County's export ("Madera Co November 5" header,
    # values identical to Madera's file), so no Kings parser runs until a
    # real Kings source is found.
    "madera": ("Madera", "canvass",
               "2024 Madera County, CA precinct-level results/Statement-of-Votes-Cast-Excel-November-5-2024.xlsx"),
    "yuba": ("Yuba", "canvass",
             "2024 Yuba County, CA precinct-level results.xlsx"),
    "ventura": ("Ventura", "canvass",
                "2024 Ventura County, CA precinct-level results.xlsx"),
    "monterey": ("Monterey", "canvass",
                 "2024 Monterey County, CA precinct-level results.xlsx"),
    "tulare": ("Tulare", "canvass",
               "2024 Tulare County, CA precinct-level results.xlsx"),
    "santa_cruz": ("Santa Cruz", "canvass",
                   "2024 Santa Cruz County, CA precinct-level results.xls"),
    "sacramento": ("Sacramento", "sacramento",
                   "2024 Sacramento County, CA precinct-level results.xlsx"),
    "sonoma": ("Sonoma", "ssml",
               "Sonoma County - Precinct Results - General 2024.xls"),
    "santa_clara": ("Santa Clara", "ssml",
                    "Santa Clara - Precint results - General 2024.xls"),
    "san_mateo": ("San Mateo", "san_mateo_csv",
                  "2024 San Mateo County, CA precinct-level results.csv"),
    "placer": ("Placer", "placer_txt",
               "2024 Placer County, CA precinct-level results.txt"),
    "trinity": ("Trinity", "trinity_csv",
                "2024 Trinity County, CA precinct-level results/"
                "Trinity-2024-General-by-Precinct.csv"),
    "los_angeles": ("Los Angeles", "los_angeles",
                    "2024 Los Angeles County, CA precinct-level results/4324_final_svc_excel_v2.zip"),
    "nevada": ("Nevada", "dominion_canvass",
               "2024 Nevada County, CA precinct-level results.pdf"),
    "solano": ("Solano", "dominion_canvass",
               "2024 Solano County, CA precinct-level results.pdf"),
    "humboldt": ("Humboldt", "dominion_canvass",
                 "Humboldt County - Statement of the Vote - General 2024.pdf"),
    "calaveras": ("Calaveras", "dominion_canvass",
                  "Calaveras County - Statement of Vote - General Election.pdf"),
    "sierra": ("Sierra", "rotated180_pdf",
               "2024 Sierra County, CA precinct-level results.pdf"),
    "mendocino": ("Mendocino", "dominion_canvass",
                  "Mendocino County - Statement of the Vote - General 2024.pdf"),
    "imperial": ("Imperial", "pageformat_pdf",
                 "2024 Imperial County, CA precinct-level results.pdf"),
    "kern": ("Kern", "pageformat",
             "2024 Kern County, CA precinct-level results.xlsx"),
    "tuolumne": ("Tuolumne", "tuolumne",
                 "2024 Tuolumne County, CA precinct-level results"),
    "modoc": ("Modoc", "modoc24",
              "2024 Modoc County, CA precinct-level results/UNOFFICIAL election-night precinct summary 2024-11-05.pdf"),
    "plumas": ("Plumas", "pageformat_pdf",
               "2024 Plumas County, CA precinct-level results.pdf"),
    # NOTE on the rotated-pageformat PDF family beyond Plumas: Lake's doc
    # prints each contest on two page-pairs and leaks "Registration"/
    # "Turnout (%)" stat labels as candidates; totals come out ~half. Needs
    # its own pass (as do San Joaquin / Merced, same rotated-header style).
    "san_joaquin": ("San Joaquin", "pageformat_pdf",
                    "2024 San Joaquin County, CA precinct-level results.pdf"),
    "merced": ("Merced", "pageformat_pdf",
               "2024 Merced County, CA precinct-level results.pdf"),
    "shasta": ("Shasta", "dominion_canvass",
               "2024 Shasta County, CA precinct-level results.pdf"),
}

PARSERS = {}
for slug, (display, family, rel) in SOURCE_FILES.items():
    PARSERS[slug] = family


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("counties", nargs="*", help="county slugs (default: all)")
    args = ap.parse_args()
    targets = args.counties or list(SOURCE_FILES)

    ok = fail = 0
    for slug in targets:
        if slug not in SOURCE_FILES:
            print(f"!! no parser registered for {slug}")
            continue
        display, family, rel = SOURCE_FILES[slug]
        fn = globals().get(f"parse_{family}")
        if family == "trinity_csv":
            rows, verify = fn(display, SOURCE_DIR / rel,
                              SOURCE_DIR / "2024 Trinity County, CA "
                              "precinct-level results/"
                              "Trinity-2024-General-summary.csv")
        else:
            rows, verify = fn(display, SOURCE_DIR / rel)
        status = "clean" if not verify else f"SOFT({len(verify)})"
        print(f"[{slug}] {len(rows)} rows {status}")
        for v in verify[:6]:
            print("   !", v)
        if verify:
            fail += 1
        else:
            ok += 1
        write_county(display, rows)
    print(f"\n{ok} counties clean, {fail} county files with soft diffs;")


if __name__ == "__main__":
    main()