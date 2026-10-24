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
* ``parse_rotated_sov`` — the rotated-header canvass PDFs: Lake, San
  Joaquin, Merced.
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
        # a quoted nickname ("MIKE") keeps its quotes and is titled inside
        m_q = re.fullmatch(r'"([^"]+)"', seg)
        if m_q:
            core = m_q.group(1)
            core = core.upper() if core.upper() in keep_upper or len(core) == 1 \
                else core.capitalize()
            parts.append('"' + core + '"')
            continue
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
        # "United States Senator (Short Term)" = Feinstein's unexpired seat;
        # 2024 CA ran two senate races with identical candidates, so a Short
        # Term file needs its own contest key or the county aggregate
        # double-counts both races into one
        return ("U.S. Senate",
                "Partial Term" if re.search(r"UNEXPIRED|PARTIAL|SHORTTERM",
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
# Lake / San Joaquin / Merced — the three 2024 rotated-header canvass PDFs
# ("District Results Report" for Lake/San Joaquin, "Custom Table Report" for
# Merced). Candidate and stat labels print sideways at ~90 degrees with
# their glyphs mirrored in the character stream; a vertical label line is a
# run of characters sharing one x0, drawn bottom-up for Lake/San Joaquin
# (chars ascend in top, line text = the char stream reversed) and top-down
# for Merced (chars ascend in top, no reversal). Label lines ~9.4-10.6pt
# apart merge into one label block. Lake/San Joaquin blocks hold a [name
# line, running-mate line] pair (the bands carry no party codes, so pairs
# are stored as "Name / Mate" with a blank party); Merced blocks hold the
# wrapped name chunks, the party code leading the right-most chunk line.
#
# Value cells are right-aligned: every cell in a column shares one exact
# right edge, so each page's label blocks map 1:1 (same count, same order)
# onto its sorted value columns and no guessed column offset is involved.
# Column typing is by label words: the registration / ballots-cast /
# turnout stat columns, "Write-in Totals" aggregate columns and overvotes /
# undervotes / contest-total columns are skipped as candidates, and a
# "Total Votes" column (Merced only — Lake/San Joaquin print none) is
# captured for a per-precinct cross-check.
#
# Lake/San Joaquin print each contest as a page group: Cumulative Totals
# (the per-precinct rows; the group's first page omits the section line),
# Vote by Mail / Election Day / Grand Totals rollup pages (district and
# county rows only — no per-precinct channel detail, so rows are votes
# only), and, for contests with named write-ins, a Write-in Tally group at
# the end of the document whose rows ARE precinct-attributable. The county
# row on each contest's Grand Totals page ("COUNTY OF LAKE" / "SAN JOAQUIN
# COUNTY") verifies the precinct sums. San Joaquin's write-in pair has no
# county row there, so it cross-checks against the "(W)" row of the
# countywide report (Lake's Write-in Tally page carries its own bottom
# "Grand Totals" row).
#
# Merced prints one block of one or two contests per page-set side by side
# ("DISTRICT 13 DISTRICT 27"; a precinct outside the second contest's
# district simply omits its six columns), each block ending in a county
# Totals row that verifies it; the two sub-contests split after each
# "Contest Total" column. Merced's titles shorten to the district number
# alone (the office word is never printed), resolved via the county's 2024
# district table below.
# ---------------------------------------------------------------------------

_ROT_CONFIG = {
    "Lake": {"family": "district"},
    "San Joaquin": {"family": "district"},
    # Stanislaus prints the same "District Results Report" blocks (title on
    # line 48, section on line 69, per-precinct rows at y>=194). Its opening
    # "Cumulative Results Report" countywide section (pages 7-53) renders
    # the statewide contests' text in a double-printed, part-invisible
    # layer (garbled extract), so unlike San Joaquin's it yields no
    # countywide totals; verification rests on the district reports' own
    # Grand-Totals "STANISLAUS COUNTY," rows.
    "Stanislaus": {
        "family": "stanislaus",
        # The rotated bands print "NAME MATE(, PARTY)" wrapped across one
        # or two vertical lines — the name/mate split is only recoverable
        # from the known 2024 statewide running-mate names.
        "mate_names": ("JD VANCE", "KARINA GARCIA", "TIM WALZ",
                       "NICOLE SHANAHAN", "MIKE TER MAAT", "RUDOLPH WARE",
                       "LAUREN ONAK"),
    },
    "Merced": {
        "family": "custom",
        # Merced's canvass strips the office word out of the contest
        # titles: district contests print "DISTRICT N" and the two U.S.
        # Senate seats print "FULL TERM" / "PARTIAL / UNEXPIRED TERM".
        # Offices for the county's 2024 contests, checked against the
        # candidates the bands carry (Gray/Duarte for CD 13, Alanis/Self
        # for AD 22, Garcia Rose/Soria for AD 27 — the same race as in the
        # Madera precinct file in this repo; Garvey/Schiff for both
        # Senate seats).
        "district_offices": {
            "13": ("U.S. House", "13"),
            "22": ("State Assembly", "22"),
            "27": ("State Assembly", "27"),
        },
        "title_offices": {
            "FULL TERM": ("U.S. Senate", ""),
            "PARTIAL / UNEXPIRED TERM": ("U.S. Senate", "Partial Term"),
        },
    },
}

_ROT_STATS_RE = re.compile(r"registration|ballots?\s*cast|turnout", re.I)
_ROT_TOTAL_RE = re.compile(r"total votes", re.I)
_ROT_SKIP_RE = re.compile(
    r"registration|ballots?\s*cast|turnout", re.I)
_ROT_PARTY_TOKEN_RE = re.compile(
    r"^(?:DEM|REP|LIB|GRN|PAF|AIP|AI|NPP|NP|PF|PSL|PEF|PFP)$", re.I)
_ROT_GRAND_RE = re.compile(r"grand totals", re.I)


def _rot_line_text(chars, reverse):
    """Text of one vertical rotated label line (chars share one x0)."""
    cs = sorted(chars, key=lambda c: c["top"])
    s = "".join(c["text"] for c in cs).strip()
    if reverse:
        s = s[::-1]
    return re.sub(r"\s+", " ", s).strip()


def _rot_label(block, reverse, mate_lines, label_mode="district",
               mate_names=()):
    """Assemble one rotated label block into (name, mate, party) — either
    the [name, running-mate] pair lines (Lake/San Joaquin), the wrapped
    name chunks with the party code leading the right-most line (Merced),
    or Stanislaus's "[name mate(, party)]" wrapped lines (the trailing
    comma only marks the printed line break; the party code trails the
    mate, on that same line or the next)."""
    if label_mode == "stanislaus":
        words, party = [], ""
        for _x, cs in block:   # lines read left to right
            words.extend(_rot_line_text(cs, reverse).split())
        # a trailing comma only marks the printed line break ("...VANCE," /
        # "GARCIA, PARTY"); name tokens keep their periods (" J. ", "JR.")
        clean = [t.rstrip(",") for t in words]
        clean = [t for t in clean if t]
        if clean and _ROT_PARTY_TOKEN_RE.match(clean[-1].rstrip(".,")):
            party = clean[-1].rstrip(".,")
            clean = clean[:-1]
        if clean and clean[-1].endswith("."):
            # stray terminal period after the name (not "J."/"JR." of a
            # longer token)
            clean[-1] = clean[-1].rstrip(".")
        name = " ".join(clean)
        mate = ""
        low = name.lower()
        for mn in mate_names:
            if low.endswith(" " + mn.lower()):
                cut = len(name) - len(mn)
                name, mate = name[:cut].strip(), name[cut:].strip()
                break
        return name, mate, party
    texts = [_rot_line_text(cs, reverse) for _x, cs in block]
    if mate_lines:
        return (texts[0] if texts else "",
                texts[1] if len(texts) > 1 else "", "")
    ordered = sorted(block, key=lambda b: -b[0])
    words, party = [], ""
    toks0 = _rot_line_text(ordered[0][1], reverse).split()
    for i, tok in enumerate(toks0):
        if i == 0 and _ROT_PARTY_TOKEN_RE.match(tok):
            party = tok
        else:
            words.append(tok)
    for _x, cs in ordered[1:]:
        words.extend(_rot_line_text(cs, reverse).split())
    return " ".join(words), "", party


def _rot_column_kind(label):
    """total (per-precinct Total Votes column) / writein (aggregate, dropped)
    / under / over / cast (contest-total column) / stat -> not candidates."""
    lo = label.lower()
    if _ROT_TOTAL_RE.search(lo):
        return "total"
    if re.search(r"write.?in", lo):
        return "writein"
    if re.search(r"undervot", lo):
        return "under"
    if re.search(r"overvot", lo):
        return "over"
    if re.search(r"contest\s*total|votes?\s*cast", lo):
        return "cast"
    if _ROT_SKIP_RE.search(lo):
        return "skip"
    return "cand"


def _rot_pair_split(name):
    """The presidential write-in pair prints on one compressed line
    ("PETER SONSKI LAUREN ONAK", the only 4-word label line); split it."""
    toks = name.split()
    return " ".join(toks[:len(toks) // 2]), " ".join(toks[len(toks) // 2:])


def _rot_countywide_totals(county, path):
    """San Joaquin's opening "Cumulative Results Report" (one block per
    contest; each candidate row carries VBM/Early/Election-Day/Total value
    percent pairs — the last number pair is the county total) — returns
    {(od, nn): county total}; an empty dict for documents that have no
    countywide report (Lake). The "(W)" write-in rows it carries are the
    only county-level figure for the write-in pair the per-contest Grand
    Totals pages have no column for."""
    totals = {}
    num_re = re.compile(r"^[\d,]+$")
    with pdfplumber.open(str(path)) as pdf:
        # the countywide report precedes the District Results Report; the
        # first 60 pages over it amply (SJ's spans pages 10-41)
        for page in pdf.pages[:60]:
            if "Cumulative Results Report" not in (page.extract_text() or ""):
                continue
            byline = defaultdict(list)
            for w in page.extract_words(extra_attrs=["upright"]):
                if w.get("upright", True):
                    byline[round(w["top"])].append(w)
            od = None
            state = None
            last_key = None
            for y in sorted(byline):
                ws = sorted(byline[y], key=lambda w: w["x0"])
                toks = [w["text"].strip() for w in ws]
                txt = " ".join(toks)
                low = txt.lower()
                nums = [t for t in toks if num_re.fullmatch(t)]
                alpha_n = len([t for t in toks if re.search(r"[A-Za-z]", t)])
                if re.match(r"^choice party", low):
                    state = "cands"
                    continue
                if state == "cands" and re.match(
                        r"^(?:cast votes|undervotes|overvotes|rejected "
                        r"write.?in|unresolved write.?in)", low):
                    state = None
                    last_key = None
                    continue
                txt_low = re.sub(r"\s*[-–]\s*vote\s+for\s+.*$", "", txt,
                                 flags=re.I)
                w_flag = bool(re.search(r"\(W\)", txt))
                if nums and len(nums) > 2 and state == "cands" and od:
                    # candidate row: name words + value/percent pairs
                    name_toks = [t for t in toks
                                 if not num_re.fullmatch(t)
                                 and not t.endswith("%")]
                    name = " ".join(name_toks)
                    name = re.sub(r"\s*\(W\)\s*$", "", name, flags=re.I)
                    mate = ""
                    if w_flag and len(name.split()) == 4:
                        # the write-in pair row prints both names together
                        name, mate = _rot_pair_split(name)
                    nm = norm_name(name + (" / " + mate if mate else ""))
                    key = od + (nm,)
                    last_key = key
                    totals[key] = _num(nums[-1])
                    continue
                if state == "cands" and od and last_key is not None and \
                        not nums and alpha_n <= 4:
                    # running mate on its own line (e.g. "JD VANCE")
                    old = totals.get(last_key)
                    if old is None:
                        last_key = None
                        continue
                    display = last_key[2].split(" / ")[0]
                    nm = norm_name(f"{display} / {' '.join(toks)}")
                    totals.pop(last_key)
                    last_key = od + (nm,)
                    totals[last_key] = old
                    continue
                if alpha_n >= 2 and not nums:
                    cand = parse_office2024(txt_low, county)
                    if cand:
                        od, state, last_key = cand, None, None
    return totals


def parse_rotated_sov(county, path):
    """Rotated-label canvass PDFs; returns (rows, verify_failures). See the
    family notes above for the two layouts."""
    cfg = _ROT_CONFIG.get(county)
    if cfg is None:
        raise ValueError(f"no rotated_sov config for {county}")
    family = cfg["family"]
    label_mode = family
    mate_names = cfg.get("mate_names", ())
    if family == "stanislaus":
        # Stanislaus's page geometry is the district family's (title on
        # line 48, section on line 69, per-precinct rows from y=195);
        # only the label assembly differs.
        family = "district"
    reverse = mate_lines = family == "district"
    num_re = re.compile(r"^[\d,]+$")
    votes = {}     # (precinct, od, nn) -> votes
    ptotal = {}    # (precinct, od) -> Total Votes cell value (Merced)
    rowaudit = {}  # (precinct, od) -> {under/over/cast/writein: value}
    totals = {}    # (od, nn) -> county total from the rollup rows
    names = {}     # (od, nn) -> (display, party raw)
    failures = []

    def county_label(label):
        u = label.upper().replace(",", "").strip()
        return u in (f"COUNTY OF {county.upper()}",
                     f"{county.upper()} COUNTY")

    def title_od(text):
        key = " ".join(str(text or "").split()).upper()
        od = parse_office2024(key, county)
        if od is not None:
            return od
        if family == "custom":
            m = re.fullmatch(r"DISTRICT\s+(\d+)", key)
            if m:
                od = cfg.get("district_offices", {}).get(m.group(1))
                if od:
                    return od
            title_offices = cfg.get("title_offices", {})
            if key in title_offices:
                return title_offices[key]
        return None

    def store(precinct, od, name, mate, party, value):
        if not name or _ROT_PARTY_TOKEN_RE.match(name):
            return
        if mate:
            name = f"{name} / {mate}"
        nm = norm_name(name)
        key = (precinct, od, nm)
        old = votes.get(key)
        if old is not None and old != value:
            failures.append(f"{od[0]} {precinct} {name[:26]}: block "
                            f"repeated, {old} vs {value}")
        votes[key] = value
        names[od + (nm,)] = (name, party)

    def assign(x1, cols):
        best, bd = None, 3.0
        for i, cx in enumerate(cols):
            d = abs(cx - x1)
            if d < bd:
                best, bd = i, d
        return best

    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            words = page.extract_words(extra_attrs=["upright"])
            byline = defaultdict(list)
            for w in words:
                if w.get("upright", True):
                    byline[round(w["top"])].append(w)

            def line(y):
                ws = byline.get(y)
                return " ".join(w["text"] for w in
                                sorted(ws, key=lambda x: x["x0"])) \
                    if ws else ""

            # ---- contest title(s) for the page --------------------------
            if family == "district":
                title = line(48).strip()
                if not title:
                    continue        # countywide report / certification forms
                od0 = title_od(title)
                if od0 is None:
                    continue        # local contest / state measure
                ods = [od0]
                sec = (line(69) + " " + line(83)).strip().lower()
                want = "rollup" if _ROT_GRAND_RE.search(sec) else \
                    "precinct" if (re.search(r"write.?in", sec)
                                   or "cumulative" in sec or not sec) \
                    else None
                if want is None:
                    continue        # Vote by Mail / Election Day rollups
            else:
                # data rows begin with a bare digit id line (or a "Totals"
                # county-row id line); the title zone = the upright text
                # lines between the page header and the first data row
                data_ys = [y for y in sorted(byline) if y > 130 and
                           len(byline[y]) == 1 and byline[y][0]["x0"] < 30
                           and (byline[y][0]["text"].isdigit() or
                                byline[y][0]["text"].lower() == "totals")]
                zone_end = data_ys[0] if data_ys else 790
                zone = [y for y in sorted(byline)
                        if 50 < y < zone_end - 6 and
                        any(re.search(r"[A-Za-z]", w["text"])
                            for w in byline[y])]
                if not zone:
                    continue        # title-less local continuation page
                ts = sorted(byline[zone[0]], key=lambda w: w["x0"])
                subs, cur = [], [ts[0]]     # side-by-side contests split
                for w in ts[1:]:            # on the title-line word gap
                    if w["x0"] - cur[-1]["x1"] > 40:
                        subs.append(cur)
                        cur = [w]
                    else:
                        cur.append(w)
                subs.append(cur)
                ods = [title_od(" ".join(w["text"] for w in s).strip())
                       for s in subs]
                want = "precinct"
                if not any(ods):
                    continue        # measures / local contest pairs

            # ---- rotated label blocks (chars grouped by shared x0) ------
            xcols = defaultdict(list)
            for c in page.chars:
                if not c.get("upright", True):
                    xcols[round(c["x0"], 1)].append(c)
            blocks, cur = [], []
            for x0 in sorted(xcols):
                if cur and x0 - cur[-1][0] > 13.0:
                    blocks.append(cur)
                    cur = []
                cur.append((x0, xcols[x0]))
            if cur:
                blocks.append(cur)
            if not blocks:
                continue
            labels = [_rot_label(b, reverse, mate_lines, label_mode,
                                 mate_names) for b in blocks]
            kinds = [_rot_column_kind(n + " " + m + " " + p)
                     for n, m, p in labels]

            # ---- one od per value column: side-by-side contests end at
            #      their contest-total column (custom family) -------------
            col_od = []
            if family == "custom":
                gi = 0
                for lab, k in zip(labels, kinds):
                    col_od.append(ods[gi] if gi < len(ods) else None)
                    if k == "cast":
                        gi += 1
                if sum(1 for k in kinds if k == "cast") != len(ods):
                    failures.append(f"{county} p{page.page_number}: "
                                    f"{sum(1 for k in kinds if k == 'cast')} "
                                    f"contest-total columns for "
                                    f"{len(ods)} titles")
                    continue
            else:
                col_od = [ods[0]] * len(blocks)

            # ---- rows ----------------------------------------------------
            recs = []   # ("precinct", pct, cells) / ("rollup", label, cells)
            if family == "district":
                for y in sorted(byline):
                    if y < 185:
                        continue
                    ws = sorted(byline[y], key=lambda w: w["x0"])
                    head = ws[0]
                    tok0 = head["text"].strip()
                    cellws = [w for w in ws[1:]
                              if num_re.match(w["text"].strip()) or
                              re.fullmatch(r"[\d,]+\.\d+",
                                           w["text"].strip())]
                    if head["x0"] < 26 and re.fullmatch(r"[0-9]{1,7}", tok0):
                        if cellws:
                            recs.append(("precinct", tok0, cellws))
                        continue
                    if cellws:
                        label = " ".join(w["text"] for w in ws
                                         if w not in cellws).strip()
                        if _ROT_GRAND_RE.search(label) or county_label(label):
                            recs.append(("rollup", label, cellws))
            else:
                # one id-token line, then the row's value cells on their
                # own line; the last id line of each block is "Totals"
                pct, pending_totals = None, False
                for y in sorted(byline):
                    if y < 130:
                        continue
                    ws = sorted(byline[y], key=lambda w: w["x0"])
                    tok0 = ws[0]["text"].strip()
                    if len(ws) == 1 and ws[0]["x0"] < 30 and \
                            (tok0.isdigit() or tok0.lower() == "totals"):
                        if tok0.isdigit():
                            pct, pending_totals = tok0, False
                        else:
                            pct, pending_totals = None, True
                        continue
                    cellws = [w for w in ws if num_re.match(w["text"].strip())]
                    if cellws and len(cellws) == len(ws):
                        if len(cellws) >= 6 and (pct is not None or
                                                 pending_totals):
                            if pct is not None:
                                recs.append(("precinct", pct, cellws))
                            else:
                                recs.append(("rollup", "Totals", cellws))
                        pct, pending_totals = None, False
                    else:
                        pct, pending_totals = None, False

            if want != "precinct" and not any(r[0] == "rollup" for r in recs):
                continue    # rollup page without the county row: nothing

            # ---- value columns: cluster every row cell's right edge ------
            edges = sorted(w["x1"] for r in recs for w in r[2])
            buckets = []
            for x in edges:
                if not buckets or x - buckets[-1][-1] > 3.0:
                    buckets.append([x])
                else:
                    buckets[-1].append(x)
            cols = [sum(b) / len(b) for b in buckets]
            if len(cols) != len(blocks):
                failures.append(f"{county} p{page.page_number}: "
                                f"{len(cols)} value columns vs "
                                f"{len(blocks)} label blocks "
                                f"({[l[0][:16] for l in labels]})")
                continue

            # ---- store ----------------------------------------------------
            for kind, key0, cellws in recs:
                for w in cellws:
                    i = assign(w["x1"], cols)
                    if i is None:
                        continue
                    k = kinds[i]
                    if k == "cand" and col_od[i] is None:
                        continue
                    if kind == "precinct":
                        if k == "total":
                            ptotal[(key0, col_od[i])] = _num(w["text"])
                        elif k in ("under", "over", "cast", "writein"):
                            rowaudit.setdefault((key0, col_od[i]), {})[k] = \
                                _num(w["text"])
                        elif k == "cand":
                            name, mate, party = labels[i]
                            if family == "district" and label_mode \
                                    != "stanislaus" and not mate and \
                                    len(name.split()) == 4:
                                # the presidential write-in pair prints on
                                # one compressed line
                                name, mate = _rot_pair_split(name)
                            store(key0, col_od[i], name, mate, party,
                                  _num(w["text"]))
                    elif k == "cand":   # county / Grand Totals / Totals row
                        name, mate, party = labels[i]
                        if family == "district" and label_mode \
                                != "stanislaus" and not mate and \
                                len(name.split()) == 4:
                            # the write-in pair's county-total row carries
                            # both names on one line, like the band
                            name, mate = _rot_pair_split(name)
                        nm = norm_name(name + (" / " + mate if mate else ""))
                        totals[col_od[i] + (nm,)] = _num(w["text"])

    # ---- verification ------------------------------------------------------
    # 1. county totals rollup rows vs the precinct sums
    sums = defaultdict(int)
    for (pct, od, nn), v in votes.items():
        sums[od + (nn,)] += v
    for key, s in sums.items():
        if key in totals and s != totals[key]:
            failures.append(f"{key[0]} {key[2]}: precincts {s} vs "
                            f"{totals[key]} in source total")
    # candidates present in the rollups but missing from any precinct page
    no_rows = [f"{k[0]} {names.get(k, ('?', ''))[0][:40]}"
               for k in totals if k not in sums]
    for msg in no_rows[:4]:
        failures.append(f"{msg}: in source totals, absent from precincts")
    # 2. Merced per-precinct Total Votes column: candidate cells (with the
    #    write-in aggregate) must total it, and cast = total + over + under
    per = defaultdict(int)
    for (pct, od, nm), v in votes.items():
        per[(pct, od)] += v
    tdiffs = [(p, o, per[(p, o)] + rowaudit[(p, o)].get("writein", 0),
               ptotal[(p, o)])
              for (p, o) in ptotal
              if per.get((p, o), 0) + rowaudit.get(
                  (p, o), {}).get("writein", 0) != ptotal[(p, o)]]
    cdiffs = [(p, o)
              for (p, o), d in rowaudit.items()
              if d.get("cast") is not None and
              d.get("cast") != ptotal.get((p, o), 0) + d.get("over", 0)
              + d.get("under", 0)]
    if tdiffs:
        failures.append(f"{county}: Total Votes column mismatches on "
                        f"{len(tdiffs)} of {len(ptotal)} rows; first: "
                        f"{tdiffs[0][0]} {tdiffs[0][1][0]} "
                        f"{tdiffs[0][2]} vs {tdiffs[0][3]}")
    if cdiffs:
        failures.append(f"{county}: Contest Total != Total Votes + "
                        f"over + under on {len(cdiffs)} of "
                        f"{len(rowaudit)} rows")
    # 3. San Joaquin countywide report: a second county-total source;
    #    the write-in pair's totals only exist there
    cw = _rot_countywide_totals(county, path)
    if cw:
        for key, t in cw.items():
            s = sums.get(key)
            if s is None:
                continue    # candidate not in the precinct pages (skipped)
            if key not in totals and s != t:
                failures.append(f"{key[0]} {names[key][0][:40]}: precincts "
                                f"{s} vs countywide {t} (no Grand Totals "
                                f"row)")
    rows_out = []
    for (precinct, (office, district), nn), v in sorted(
            votes.items(), key=lambda k: (k[0][1][0], k[0][1][1], k[0][2],
                                          k[0][0])):
        name, party = names[(office, district, nn)]
        rows_out.append({
            "county": county, "precinct": precinct, "office": office,
            "district": district,
            "candidate": _mc_fix(title_case(name)),
            "party": party_of(party), "votes": v,
        })
    return rows_out, failures


def _mc_fix(name):
    """title_case renders McNerney as "Mcnerney"; re-capitalize the Mc stem."""
    return " ".join(
        f"Mc{t[2:].capitalize()}"
        if len(t) > 3 and t[:2].lower() == "mc" and t[2].islower() else t
        for t in name.split())


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
                "vote by mail", "election day", "mail", "polling", "poll"}
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


# San Bernardino prints some ids with a stray short token in front of the
# block id ("6 - 0206 6 - 0206"; the digit belongs to the ballot-style code
# line above) -- without accepting it the id line doesn't match and `current`
# goes stale, so every following block on the page stores under the previous
# precinct's id and total-row overwrites clobber its results.
_PF_IDLINE_RE = re.compile(r"^\s*(?:(?:[a-z%]{1,3}|\d{1,3})\s+)?-\s*([0-9]{2,7})\s*.*$")
# San Bernardino ballot-style codes ("C23S19A34E1B1AVUAP10001"): a short
# uppercase head that includes a digit early, then a long tail. (A plain
# minimum-length pattern would also match long titles like REPRESENTATIVE.)
_PF_CODE_RE = re.compile(r"^[A-Z]{1,4}[0-9][0-9A-Z\-]{9,}$")
# Glenn/SLO sub-row block id lines: bare reporting-precinct ids ("10040",
# "114", "MB01"), optionally printed twice per page ("10040 10040" = the
# left stat band's copy and the right candidate band's copy).
_PF_PCT_ID_RE = re.compile(r"^(?:[A-Z]{2})?\d{1,7}$")
# Glenn/SLO county-total rollup rows ("Electionwide - Total ...", SLO's
# "County - Total ..."): per-candidate county totals for verification.
_PF_COUNTY_RE = re.compile(r"^(?:electionwide|county|county of \w+|"
                           r"\S+ county)\s*-\s*total\b", re.I)

# 2024 running-mate pairs: the printed president labels mix the pair into one
# string with or without slashes and line wraps ("DONALD J. TRUMP/JD VANCE"),
# so the split runs on the known pairs.
_PRES_PAIRS = [
    (r"DONALD\s+J\.?\s*TRUMP", r"JD\s+VANCE"),
    (r"KAMALA\s+D\.?\s*HARRIS", r"TIM\s+WALZ"),
    (r"ROBERT\s+F\.?\s*KENNEDY\s+JR\.?", r"NICOLE\s+SHANAHAN"),
    (r"CHASE\s+OLIVER", r"MIKE\s+TER\s+MAAT"),
    (r"JILL\s+STEIN", r"RUDOLPH\s+WARE"),
    (r"CLAUDIA\s+DE\s+LA\s+CRUZ", r"KARINA\s+GARCIA"),
    (r"PETER\s+SONSKI", r"LAUREN\s+ONAK"),
]


def _split_mate(text, office):
    """(name, mate) for a rotated president label; the 2024 pairs are the
    only multi-name candidates in the 2024 general's statewide contests."""
    if (office or "") != "President":
        return (text, "")
    for pn, pm in _PRES_PAIRS:
        m = re.search(rf"({pn})\s*/?\s*({pm})", text)
        if m:
            name = text[:m.start(1)].strip() or m.group(1)
            return (name.strip(), m.group(2))
    return (text, "")


def parse_pageformat_pdf(county, path, dedup=False, subrow=False,
                         rollups=False):
    """``subrow`` (Glenn/SLO): each contest's table runs as sub-row blocks —
    a bare reporting-precinct id line followed by Election Day / Vote by Mail
    / Total rows per candidate column, bands printed two-up (left band =
    Times Cast / Registered Voters stats, right band = one candidate slice).
    ``rollups``: capture "Electionwide - Total" / "County - Total" rows as
    per-candidate county totals and verify the precinct sums against them."""
    votes = {}    # (precinct, od, nn) -> votes (Total row or inline cell)
    channels = {}  # -> [election-day-ish, mail-ish] channel votes
    ctotals = {}  # (od, nn) -> county rollup total (Butte dedup mode)
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
                        _PF_STATS_RE.match(low):
                    continue
                # a standalone "State of California" banner line is a header;
                # a contest title may embed the same words ("State of
                # California State Senator District 19") -- strip the prefix
                # and evaluate the rest as the title
                if low.startswith("state of california"):
                    txt = txt[len("State of California"):].strip()
                    low = txt.lower()
                    if not re.search(r"[A-Za-z]{2}", low):
                        continue
                # a title line begins with alpha text and carries several
                # alpha words (precinct rows start with ids; ballot-style
                # code rows -- often printed twice per line -- are not titles;
                # MB-id lines ("MB01 MB01") are blocks' id lines, not titles)
                if _PF_PCT_ID_RE.match(txt.split()[0]) or \
                        len([w for w in txt.split() if re.search(r"[A-Za-z]{2}", w)]) < 2 \
                        or not re.match(r"^[A-Za-z]", txt) \
                        or any(_PF_CODE_RE.match(w) for w in txt.split()):
                    continue
                if dedup and ("%" in txt or \
                        re.search(r"\b\d{1,3}(?:,\d{3})+\b", txt)) \
                        and "(vote for" not in low:
                    # Butte: the doc's district-group rollup value rows
                    # ("State Board of Equalization ... 5,006 69.36% ...")
                    # open a page and would otherwise read as contest
                    # titles
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
                pn, pm = _split_mate(name, od[0])
                if pm:
                    name = f"{pn} / {pm}"
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
                # id lines print a right-block copy at x~399-411 ("- 0001
                # - 0001"); only numeric words in the left value columns
                # (x 130-390) disqualify the line
                if m_id and len(line) <= 6 and \
                        not any(re.match(r"^[\d,]+$", x["text"])
                                and 130 < x["x0"] < 390 for x in line[1:]):
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
                        if subrow and total_x is not None and \
                                abs((w["x0"] + w["x1"]) / 2 - (total_x + 4)) \
                                <= 20:
                            # the band's Total Votes cell is captured by its
                            # own column below; without this it is the
                            # unassigned value nearest the next label block
                            # (Glenn/SLO: the write-in pair) and pollutes
                            # that candidate's channel rows
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
                    if total_x is not None:
                        # the spread's "Total Votes" column (all candidates,
                        # all column-passes) prints on the candidate pages;
                        # capture it on every channel row
                        for w in line[1:]:
                            vc = (w["x0"] + w["x1"]) / 2
                            if abs(vc - (total_x + 4)) <= 20 and \
                                    re.match(r"^[\d,]+$", w["text"]):
                                totals[od + ("__total__", current)] = \
                                    _num(w["text"])
                    continue
                if _PF_CODE_RE.match(row_id):
                    pending_code = row_id
                    continue
                if rollups and _PF_COUNTY_RE.match(line_l) and od is not None:
                    # per-candidate county-total rollup row; candidate values
                    # assign to their columns, the stat/"Total Votes" cells
                    # assign elsewhere or to none
                    for w in line:
                        tok = w["text"]
                        if tok.endswith("%") or "." in tok or \
                                not re.match(r"^[\d,]+$", tok):
                            continue
                        lab = assign(w)
                        if lab is None:
                            continue
                        cname = lab[0]
                        if re.fullmatch(r"unresolved write.?in", cname,
                                        re.I) or \
                                "total votes" in cname.lower():
                            continue
                        cname = re.sub(r"\s*qualified write in\s*$", "",
                                       cname, flags=re.I)
                        pn2, pm2 = _split_mate(cname, od[0])
                        nm = norm_name(f"{pn2} / {pm2}" if pm2 else pn2)
                        v = _num(tok)
                        old = totals.get(od + ("__cty__", nm))
                        if old is not None and old != v:
                            failures.append(f"{od[0]} {cname[:28]}: county "
                                            f"rollup copies {old} vs {v}")
                        totals[od + ("__cty__", nm)] = v
                    current = None
                    continue
                if dedup and (line_l.startswith("electionwide") or
                              line_l.startswith("county supervisor "
                                               "district")):
                    # rollup value row ("Electionwide - Total" / "County
                    # Supervisor District - <counts>"): the contest's
                    # county totals; the diff against the precinct sums
                    # is exactly the '****'-redacted precincts' votes
                    for w in line:
                        tok = w["text"]
                        if tok.endswith("%") or "." in tok or \
                                not re.match(r"^[\d,]+$", tok):
                            continue
                        vc = (w["x0"] + w["x1"]) / 2
                        if total_x is not None and \
                                abs(vc - (total_x + 4)) <= 20:
                            continue    # county "Total Votes" cell
                        lab = assign(w)
                        if lab is None:
                            continue
                        cname = lab[0]
                        if re.fullmatch(r"unresolved write.?in", cname,
                                        re.I):
                            continue
                        cname = re.sub(r"\s*qualified write in\s*$", "",
                                       cname, flags=re.I)
                        nm = norm_name(cname)
                        v = _num(tok)
                        old = ctotals.get((od, nm))
                        if old is None:
                            ctotals[(od, nm)] = v
                        elif old != v:
                            failures.append(f"{od[0]} {cname[:28]}: county "
                                            f"rollup copies {old} vs {v}")
                    continue
                if subrow and len(line) <= 4 and \
                        all(_PF_PCT_ID_RE.match(x["text"].strip())
                            for x in line):
                    # bare id line starting a sub-row block (single or
                    # left/right doubled copy)
                    current = line[0]["text"].strip()
                    pending = defaultdict(lambda: [0, 0, 0, 0])
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
                    v = _num(tok)
                    if dedup:
                        # identical reprint (Butte repeats each contest's
                        # whole precinct table back-to-back): dedup equal
                        # values, flag unequal ones and keep the first
                        old = votes.get(key)
                        if old is not None:
                            if old != v:
                                name = names[od + (key[2],)][0]
                                failures.append(
                                    f"{office} {row_id} {name[:28]}: "
                                    f"reprint {old} vs {v}")
                            continue
                        votes[key] = v
                    else:
                        votes[key] = votes.get(key, 0) + v
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
    if rollups:
        # county-total rollups: diffs are what the doc redacts ('****' value
        # cells under the Insufficient-Turnout privacy rule) or write-in
        # aggregates — reported as soft, the file still writes
        psums = defaultdict(int)
        for (pct, (o, d), nm), v in votes.items():
            psums[(o, d, nm)] += v
        cty = {k[:2] + (k[3],): v for k, v in totals.items()
               if k[2] == "__cty__"}
        for (office, district, nm), t in sorted(
                cty.items(), key=lambda kv: (kv[0][0], kv[0][2])):
            disp = names.get((office, district, nm), (nm, ""))[0]
            s = psums.get((office, district, nm))
            if s is None:
                failures.append(f"{office} {disp[:36]}: precinct rows "
                                f"missing vs county total {t}")
            elif s != t:
                failures.append(f"{office} {disp[:36]}: precincts {s} vs "
                                f"county total {t} (redacted/write-in)")
    if dedup and ctotals:
        # county rollup rows: soft diffs only — a candidate's precinct
        # sum minus its county total is exactly what the doc hides in
        # '****' redactions (<10-voter precincts)
        psums = defaultdict(int)
        for (pct, (office, district), nm), v in votes.items():
            psums[(office, district, nm)] += v
        for (od, nm), t in sorted(ctotals.items()):
            s = psums.get(od + (nm,))
            disp = names.get(od + (nm,), (nm, ""))[0]
            if s is None:
                failures.append(f"{od[0]} {disp[:36]}: precinct rows "
                                f"missing (all redacted) vs county "
                                f"total {t}")
            elif s != t:
                failures.append(f"{od[0]} {disp[:36]}: precincts {s} "
                                f"vs county total {t} (redacted "
                                f"precincts)")
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


def parse_pageformat_pdf_dedup(county, path):
    """Butte's "Page: N of M" SOVC: parse_pageformat_pdf over reprinted
    precinct tables (dedup, don't sum)."""
    return parse_pageformat_pdf(county, path, dedup=True)


def parse_san_bernardino(county, src_dir):
    """San Bernardino — one "Page: N of M" SOVC PDF per contest (187 files,
    ~33k printed pages). Each contest's statewide files parse via
    parse_pageformat_pdf; measures/judges/local offices yield nothing there
    (parse_office2024 returns None) and are skipped. Per-file verify failures
    come back prefixed with the file name."""
    rows = []
    failures = []
    files = sorted(src_dir.glob("*.pdf"))
    for f in files:
        f_rows, f_fails = parse_pageformat_pdf(county, f)
        rows.extend(f_rows)
        for msg in f_fails:
            failures.append(f"{f.name}: {msg}")
        print(f"[san_bernardino] {f.name}: {len(f_rows)} rows "
              f"{'clean' if not f_fails else f'{len(f_fails)} diffs'}", flush=True)
    return rows, failures


# ---------------------------------------------------------------------------
# Yolo — the "District Results Report" PDF: county-level "Cumulative Report"
# pages (landscape), then one block per contest of "District Results Report"
# pages, each page carrying the contest title + a section label:
#   * "Cumulative Totals" pages — the per-precinct rows (id, Registration,
#     Ballots Cast, Turnout %, one total-votes column per candidate) plus
#     trailing channel/Grand-Totals rollup rows (district level, not
#     per-precinct).
#   * "Vote Center Totals" / "Vote by Mail Totals" / "Conditional Totals" /
#     "Election Day Totals" / "Grand Totals" pages — district-group rollups
#     only, ending in the contest's county Grand-Totals row.
# Candidate/stat column headers are rotated; the label columns sit ~49pt
# apart and each candidate's value column is right-aligned ~21pt right of
# its label line. Per-precinct rows carry totals only (no channel split),
# so rows are votes-only; the county Grand-Totals row verifies the sums,
# and the county-level Cumulative Report pages supply party codes.
# ---------------------------------------------------------------------------

_YOLO_TOTALS_RE = re.compile(r"^(?:cumulative|vote center|vote by mail|conditional|"
                             r"election day|grand)\s+totals?$", re.I)
_YOLO_PRECINCT_RE = re.compile(r"^\d{2,7}$")


def _yolo_label_lines(page):
    """Rotated vertical label lines -> [(x0, text)] (words read bottom-up,
    each word's characters reversed)."""
    cols = defaultdict(list)
    for w in page.extract_words(extra_attrs=["upright"]):
        if not w.get("upright", True):
            cols[round(w["x0"])].append(w)
    out = []
    for x in sorted(cols):
        ws = sorted(cols[x], key=lambda w: -w["top"])
        out.append((x, " ".join(w["text"][::-1] for w in ws)))
    return out


def _yolo_uptight_lines(page):
    """Upright lines -> {top: [words sorted by x0]}."""
    lines = defaultdict(list)
    for w in page.extract_words():
        if w.get("upright", True):
            lines[round(w["top"])].append(w)
    for y in lines:
        lines[y].sort(key=lambda w: w["x0"])
    return lines


def _yolo_assign(labels, words):
    """Map candidate value columns onto a row's numeric words: a value
    belongs to the label line right of which it is right-aligned (the
    value column ends ~21pt right of its label line; label columns are
    ~49pt apart)."""
    out = []
    for x, text in labels:
        for w in words:
            tok = w["text"].strip()
            if "%" in tok or "." in tok or not re.fullmatch(r"[\d,]+", tok):
                continue
            vc = (w["x0"] + w["x1"]) / 2
            if x - 20 <= vc <= x + 36:
                out.append((x, text, _num(tok)))
    return out


def parse_yolo_pdf(county, path):
    """District Results Report: per-precinct totals from the "Cumulative
    Totals" pages; verified against each contest's county "Grand Totals" row
    and the county-level Cumulative Report pages (which also carry the
    candidate party codes). Rows are votes-only (no per-precinct channels)."""
    votes = {}      # (precinct, od, nn) -> votes
    names = {}      # (od, nn) -> display
    ref_party = {}  # (od, nn) -> party token (county-level pages)
    ref_totals = {} # (od, nn) -> county-level page total
    gt_totals = {}  # (od, nn) -> county Grand-Totals row value
    failures = []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            header = page.extract_text() or ""
            lines = _yolo_uptight_lines(page)
            if "Cumulative Report" in header and \
                    "District Results Report" not in header:
                # county-level pages: "CANDIDATE - PARTY v p ... v p" rows
                od = None
                for y in sorted(lines):
                    txt = " ".join(w["text"] for w in lines[y]).strip()
                    low = txt.lower()
                    if "vote for" in low or re.match(r"^president\b", low):
                        od = parse_office2024(txt, county)
                        continue
                    if od is None or not re.search(r"[A-Za-z]", txt):
                        continue
                    toks = [w["text"].strip() for w in lines[y]]
                    m = re.match(r"^(.+?)\s*-\s*([A-Z]{2,4})\s+(.+)$",
                                 " ".join(toks))
                    if not m or not re.search(r"[A-Za-z]{2}", m.group(1)):
                        continue
                    nums = [t for t in toks if re.fullmatch(r"[\d,]+", t)]
                    if len(nums) < 3 or is_summary_name(m.group(1)):
                        continue
                    nn = norm_name(title_case(m.group(1)))
                    ref_party[od + (nn,)] = m.group(2)
                    ref_totals[od + (nn,)] = _num(nums[-1])
                continue
            if "District Results Report" not in header:
                continue
            # title + section lines under the run header
            titles = []
            for y in sorted(lines)[:8]:
                txt = " ".join(w["text"] for w in lines[y]).strip()
                low = txt.lower()
                if "district results report" in low or \
                        re.match(r"^\d{2}/\d{2}/\d{4}\b", low) or \
                        low.startswith(("run ", "page")) or \
                        re.fullmatch(r"\d+", txt) or not txt:
                    continue
                titles.append((y, txt))
                if len(titles) == 2:
                    break
            if len(titles) < 2 or not _YOLO_TOTALS_RE.match(titles[1][1]):
                continue
            od = parse_office2024(titles[0][1], county)
            if od is None:
                continue
            office, district = od
            is_cumulative = titles[1][1].lower().startswith("cumulative")
            labels = [(x, text) for x, text in _yolo_label_lines(page)
                      if not re.search(r"registration|cast|ballots?|turnout|%|"
                                       r"total|write|under|over|precinct",
                                       text.lower())]
            for y in sorted(lines):
                if y <= titles[1][0]:
                    continue
                ws = lines[y]
                head = ws[0]["text"].strip()
                if _YOLO_PRECINCT_RE.match(head.replace(",", "")) and \
                        len(head) >= 2 and is_cumulative:
                    # per-precinct row: id, Registration, Ballots Cast,
                    # Turnout %, then one total column per candidate
                    nums = [w for w in ws[1:] if
                            re.fullmatch(r"[\d,]+", w["text"].strip())]
                    if len(nums) < 2:
                        continue
                    cast = _num(nums[1]["text"])
                    csum = 0
                    for _x, text, v in _yolo_assign(labels, ws):
                        nn = norm_name(title_case(text))
                        votes[(head, od, nn)] = v
                        names[od + (nn,)] = text
                        csum += v
                    if cast and csum > cast:
                        failures.append(f"{office} {head}: candidates {csum}"
                                        f" > ballots cast {cast}")
                    continue
                if not is_cumulative:
                    continue
                txt = " ".join(w["text"] for w in ws).strip()
                if re.match(r"^grand\s+totals\b", txt.lower()):
                    for _x, text, v in _yolo_assign(labels, ws):
                        nn = norm_name(title_case(text))
                        gt_totals[od + (nn,)] = v
    # verification: precinct sums vs the in-doc county totals
    psums = defaultdict(int)
    for (pct, (o, d), nn), v in votes.items():
        psums[(o, d, nn)] += v
    for (o, d, nn), t in sorted(gt_totals.items()):
        s = psums.get((o, d, nn))
        disp = names.get((o, d, nn), nn)
        if s is None:
            failures.append(f"{o} {disp[:36]}: precinct rows missing vs "
                            f"Grand Totals {t}")
        elif s != t:
            failures.append(f"{o} {disp[:36]}: precincts {s} vs Grand "
                            f"Totals {t}")
        rt = ref_totals.get((o, d, nn))
        if rt is not None and rt != t:
            failures.append(f"{o} {disp[:36]}: {county} page {rt} vs "
                            f"Grand Totals {t}")
    rows_out = []
    for (precinct, (office, district), nn), v in votes.items():
        rows_out.append({
            "county": county, "precinct": precinct, "office": office,
            "district": district, "candidate": title_case(names[(office, district, nn)]),
            "party": party_of(ref_party.get((office, district, nn), "")),
            "votes": v,
        })
    return rows_out, failures


# ---------------------------------------------------------------------------
# Napa — one "Statement of Votes Cast" PDF per contest inside a directory;
# each file: the contest title line, rotated single-line column headers
# (Registered Voters / Ballots Cast Total / Turnout / Total Votes / one
# "NAME / RUNNING MATE (PARTY)" column pair per candidate, each candidate's
# label block = a name line 15pt left of a "(PARTY)" or continuation line),
# then ONE row per precinct (right-aligned values; the Turnout cell and the
# Total Votes value sometimes wrap onto a second line under the row), and a
# trailing rollup section (district-group rows, all equal to the county
# totals) which verifies the per-candidate sums.
# ---------------------------------------------------------------------------

_NAPA_ID_RE = re.compile(r"^\d{4,7}$")


def _napa_labels(page):
    """Candidate + Total Votes column anchors from the rotated header band:
    vertical lines grouped into blocks closed by a '(PARTY)' line; the
    anchor = the block's last (rightmost) line. Stat columns are skipped."""
    cols = defaultdict(list)
    for w in page.extract_words(extra_attrs=["upright"]):
        if not w.get("upright", True):
            cols[round(w["x0"])].append(w)
    blocks, cur = [], []
    for x in sorted(cols):
        ws = sorted(cols[x], key=lambda w: -w["top"])
        text = " ".join(w["text"][::-1] for w in ws)
        low = text.lower()
        party_line = any(re.fullmatch(r"\(([A-Z]{2,4})\)", t)
                         for t in text.split())
        tv_line = re.search(r"total votes", low) is not None
        stat_line = "qualified write in" not in low and re.search(
            r"registration|registered|ballots|turnout|voters?|cast", low) \
            is not None
        if cur and (x - cur[-1][0] > 16 or tv_line or stat_line):
            blocks.append(cur)
            cur = []
        if tv_line:
            blocks.append([(x, "Total Votes")])
            continue
        if stat_line:
            continue
        cur.append((x, text))
        if party_line:
            blocks.append(cur)
            cur = []
    if cur:
        blocks.append(cur)
    out = []
    for b in blocks:
        anchor = b[-1][0] + 4
        toks = []
        for _x, text in b:
            toks.extend(text.split())
        toks = [t.strip("/") for t in toks if t and t != "/"]
        disp = " ".join(toks)
        if re.fullmatch(r"total votes", disp, re.I):
            out.append((anchor, "Total Votes"))
            continue
        m = re.search(r"\(([A-Z]{2,4})\)", disp)
        party = m.group(1) if m else ""
        disp = re.sub(r"\s*\([A-Z]{2,4}\)\s*", " ", disp)
        disp = _strip_qual(disp)
        out.append((anchor, disp, party))
    return out


def _strip_qual(text):
    """Drop a trailing '(Qualified Write In)' qualifier."""
    return re.sub(r"\s*qualified write in\s*$", "", text, flags=re.I)


def parse_napa_pdf(county, src_dir):
    """One contest per PDF; returns (rows, verify_failures). Rows are
    votes-only (the source's Registered / Ballots Cast columns are stats,
    not channels). Verified per precinct against the Total Votes cell and
    per candidate against the trailing "Countywide Total" row."""
    rows = []
    failures = []
    for f in sorted(src_dir.glob("*.pdf")):
        with pdfplumber.open(str(f)) as pdf:
            txt0 = pdf.pages[0].extract_text() or ""
            title = next((" ".join(ln.split()) for ln in txt0.splitlines()
                          if re.search(r"\(vote for[^)]*\)", ln, re.I)),
                         None)
            if title is None:
                continue
            od = parse_office2024(title, county)
            if od is None:
                continue    # local contests / state propositions
            office, district = od
            labels = None
            tvs = {}    # precinct -> Total Votes cell value
            votes = {}  # (precinct, nn) -> votes
            disp = {}   # nn -> display name
            parti = {}  # nn -> party token
            counties = {}  # (od, nn) -> "Countywide Total" row value
            for page in pdf.pages:
                if labels is None:
                    labels = _napa_labels(page)
                lines = defaultdict(list)
                for w in page.extract_words():
                    lines[round(w["top"])].append(w)
                pending = None
                for y in sorted(lines):
                    ws = sorted(lines[y], key=lambda w: w["x0"])
                    is_id = ws[0]["x0"] < 60 and _NAPA_ID_RE.match(
                        ws[0]["text"].strip())
                    if is_id:
                        pending = [ws[0]["text"].strip(), ws[1:]]
                    elif pending is not None and len(ws) <= 6 and \
                            ws[0]["x0"] > 200 and \
                            not any("%" in w["text"]
                                    for w in pending[1]):
                        # wrapped continuation row (Turnout cell and the
                        # Total Votes + candidate cells often print on a
                        # second line under the id line)
                        pending[1].extend(ws)
                    else:
                        if ws[0]["text"].strip().lower() == \
                                "countywide total":
                            nums = [w for w in ws[1:]
                                    if re.fullmatch(r"[\d,]+",
                                                    w["text"].strip())]
                            for w in nums:
                                tok = w["text"].strip()
                                vc = (w["x0"] + w["x1"]) / 2
                                best, bd = None, 12.0
                                for lab in labels:
                                    d = abs(vc - lab[0])
                                    if d < bd:
                                        best, bd = lab, d
                                if best is None or best[1] == "Total Votes":
                                    continue
                                nm = norm_name(_strip_qual(best[1]))
                                if (od, nm) in counties and \
                                        counties[(od, nm)] != _num(tok):
                                    failures.append(
                                        f"countywide copies {od[0]} "
                                        f"{disp.get(nm, nm)[:26]}: "
                                        f"{counties[(od, nm)]} vs "
                                        f"{_num(tok)}")
                                counties[(od, nm)] = _num(tok)
                        pending = None
                        continue
                    for w in pending[1]:
                        tok = w["text"].strip()
                        if "%" in tok or not re.fullmatch(r"[\d,]+", tok):
                            continue
                        vc = (w["x0"] + w["x1"]) / 2
                        best, bd = None, 12.0
                        for lab in labels:
                            d = abs(vc - lab[0])
                            if d < bd:
                                best, bd = lab, d
                        if best is None:
                            continue
                        if best[1] == "Total Votes":
                            tvs[pending[0]] = _num(tok)
                            continue
                        name = _strip_qual(best[1])
                        nm = norm_name(name)
                        votes[(pending[0], nm)] = _num(tok)
                        disp[nm] = name
                        parti[nm] = best[2]
        # per-precinct verification: candidate sum vs Total Votes cell
        rowsum = defaultdict(int)
        for (pct, nm), v in votes.items():
            rowsum[pct] += v
        for pct, s in rowsum.items():
            if pct in tvs and s != tvs[pct]:
                failures.append(f"{f.name}: {pct}: candidates {s} vs "
                                f"Total Votes {tvs[pct]}")
        # county verification: precinct sums vs "Countywide Total" row
        if counties:
            cand_sums = defaultdict(int)
            for (pct, nm), v in votes.items():
                cand_sums[nm] += v
            for (od, nm), t in sorted(counties.items()):
                s = cand_sums.get(nm)
                if s is None:
                    failures.append(f"{f.name}: {disp.get(nm, nm)[:36]}: "
                                    f"precinct rows missing vs rollup {t}")
                elif s != t:
                    failures.append(f"{f.name}: {disp.get(nm, nm)[:36]}: "
                                    f"precincts {s} vs rollup {t}")
        for (pct, nm), v in votes.items():
            pn, pm = _split_mate(" ".join(disp[nm].split()).upper(), office)
            cand = f"{pn} / {pm}" if pm else disp[nm]
            rows.append({
                "county": county, "precinct": pct, "office": office,
                "district": district, "candidate": title_case(cand),
                "party": party_of(parti[nm]), "votes": v,
            })
    return rows, failures


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


# ---- Glenn / San Luis Obispo wrappers for the sub-row pageformat mode ----

def parse_glenn_sovc(county, path):
    return parse_pageformat_pdf(county, path, subrow=True, rollups=True)


def parse_slo_sovc(county, path):
    return parse_pageformat_pdf(county, path, subrow=True, rollups=True)


def parse_yolo_drr(county, path):
    return parse_yolo_pdf(county, path)


def parse_napa_svc(county, src_dir):
    return parse_napa_pdf(county, src_dir)


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
    "lake": ("Lake", "rotated_sov",
             "2024 Lake County, CA precinct-level results.pdf"),
    "san_joaquin": ("San Joaquin", "rotated_sov",
                    "2024 San Joaquin County, CA precinct-level results.pdf"),
    "merced": ("Merced", "rotated_sov",
               "2024 Merced County, CA precinct-level results.pdf"),
    "shasta": ("Shasta", "dominion_canvass",
               "2024 Shasta County, CA precinct-level results.pdf"),
    "san_bernardino": ("San Bernardino", "san_bernardino",
                       "2024 San Bernardino County, CA precinct-level results"),
    "butte": ("Butte", "pageformat_pdf_dedup",
              "Butte County - Statement of Vote - General Election.pdf"),
    "stanislaus": ("Stanislaus", "rotated_sov",
                   "2024 Stanislaus County, CA precinct-level results.pdf"),
    # Glenn/SLO: the "Page: N of M" SOVC PDF family with sub-row channel
    # blocks (Election Day / Vote by Mail / Total per candidate column
    # band) — parse_pageformat_pdf's subrow/rollups mode
    "glenn": ("Glenn", "glenn_sovc",
              "Glenn County - Statement of the Vote - General 2024.pdf"),
    "san_luis_obispo": ("San Luis Obispo", "slo_sovc",
                        "2024 San Luis Obispo County, CA precinct-level results.pdf"),
    "yolo": ("Yolo", "yolo_drr",
             "2024 Yolo County, CA precinct-level results.pdf"),
    "napa": ("Napa", "napa_svc",
             "2024 Napa County, CA precinct-level results"),
    # NOTE: the "2024 Orange County, CA precinct-level results.pdf" is a
    # district-level rollup report only — every page of its 1,915-page
    # "District Results Report" partitions the county by city / council /
    # school / water district, and the document's only four occurrences of
    # the word "precinct" are in the city-code list and the page footer.
    # There are no per-precinct rows to parse; no Orange parser until a
    # precinct-level source is found.
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