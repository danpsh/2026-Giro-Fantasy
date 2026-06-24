#!/usr/bin/env python3
"""
Scrape Tour de France results from ProCyclingStats -> tdf-results.xlsx.

The procyclingstats library (0.2.8) can't parse PCS's current HTML, so we fetch
the pages directly (HTTP 200, browser UA) and parse the results table with
selectolax. Riders are read in table (finishing) order, which IS the ranking.

Per classification PCS uses its own page:
    finish : race/<race>/<year>/stage-<n>
    GC     : .../stage-<n>-gc      points : .../stage-<n>-points
    KOM    : .../stage-<n>-kom     youth  : .../stage-<n>-youth

    pip install procyclingstats openpyxl   # (brings selectolax)
    YEAR=2025 OUT=tdf-results-2025-test.xlsx python scripts/scrape_tdf.py
"""
import os
import re
import sys
import urllib.request

from selectolax.parser import HTMLParser
from openpyxl import Workbook

YEAR = os.environ.get("YEAR", "2026")
OUT = os.environ.get("OUT", "tdf-results.xlsx")
RACE = os.environ.get("RACE_SLUG", "tour-de-france")
MAX_STAGES = int(os.environ.get("MAX_STAGES", "21"))
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

# 2026 stage -> ISO date (used when a date can't be parsed off the page)
SCHED_2026 = {1: "2026-07-04", 2: "2026-07-05", 3: "2026-07-06", 4: "2026-07-07",
              5: "2026-07-08", 6: "2026-07-09", 7: "2026-07-10", 8: "2026-07-11",
              9: "2026-07-12", 10: "2026-07-14", 11: "2026-07-15", 12: "2026-07-16",
              13: "2026-07-17", 14: "2026-07-18", 15: "2026-07-19", 16: "2026-07-21",
              17: "2026-07-22", 18: "2026-07-23", 19: "2026-07-24", 20: "2026-07-25",
              21: "2026-07-26"}
MONTHS = {m: i for i, m in enumerate(
    ["january", "february", "march", "april", "may", "june", "july",
     "august", "september", "october", "november", "december"], start=1)}

print(f"=== ENV === YEAR={YEAR} OUT={OUT} RACE={RACE}")


def fetch(url):
    full = "https://www.procyclingstats.com/" + url
    req = urllib.request.Request(full, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode("utf-8", "replace")


def fmt_name(pcs_name):
    if not pcs_name:
        return ""
    toks = str(pcs_name).split()
    surname = [t for t in toks if t == t.upper()]
    given = [t for t in toks if t != t.upper()]
    return (" ".join(given) + " " + " ".join(w.capitalize() for w in surname)).strip()


def parse_results(html):
    """Return rider names in finishing order from the main results table."""
    tree = HTMLParser(html)
    table = tree.css_first("table.results") or tree.css_first("table")
    if not table:
        return []
    body = table.css_first("tbody") or table
    names = []
    for tr in body.css("tr"):
        a = tr.css_first('a[href*="rider/"]')
        if a:
            nm = fmt_name(a.text(strip=True))
            if nm:
                names.append(nm)
    return names


def parse_date(html, n):
    # Try "05 July 2025" / "5 july 2025" anywhere on the page
    m = re.search(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", html)
    if m and m.group(2).lower() in MONTHS:
        return f"{int(m.group(3)):04d}-{MONTHS[m.group(2).lower()]:02d}-{int(m.group(1)):02d}"
    if YEAR == "2026" and n in SCHED_2026:
        return SCHED_2026[n]
    return ""


def page_names(url, debug=False):
    try:
        names = parse_results(fetch(url))
        if debug:
            print(f"  {url} -> {len(names)} riders; top3: {names[:3]}")
        return names
    except Exception as e:
        if debug:
            print(f"  {url} -> ERROR {e}")
        return []


def row_of(names, n):
    out = names[:n] + [""] * max(0, n - len(names))
    return out[:n]


HEADER = (["Date", "Stage"] + ["1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th", "9th", "10th"]
          + [f"GC #{i}" for i in range(1, 11)] + [f"Points #{i}" for i in range(1, 4)]
          + [f"Mountain #{i}" for i in range(1, 4)] + [f"Youth #{i}" for i in range(1, 4)])


def main():
    print("\n=== STAGE 1 DETAIL ===")
    b1 = f"race/{RACE}/{YEAR}/stage-1"
    page_names(b1, debug=True)
    page_names(b1 + "-gc", debug=True)
    page_names(b1 + "-points", debug=True)
    page_names(b1 + "-kom", debug=True)
    page_names(b1 + "-youth", debug=True)

    rows = []
    for n in range(1, MAX_STAGES + 1):
        b = f"race/{RACE}/{YEAR}/stage-{n}"
        try:
            html = fetch(b)
        except Exception as e:
            print(f"stage {n}: fetch error {e}")
            continue
        finish = parse_results(html)
        if not finish:
            continue
        row = ([parse_date(html, n), n]
               + row_of(finish, 10)
               + row_of(page_names(b + "-gc"), 10)
               + row_of(page_names(b + "-points"), 3)
               + row_of(page_names(b + "-kom"), 3)
               + row_of(page_names(b + "-youth"), 3))
        rows.append(row)
        print(f"stage {n}: {row[0]} winner {row[2]}")

    if not rows:
        print("\nNo completed stages found; nothing written.")
        return
    wb = Workbook(); ws = wb.active; ws.title = "Results"; ws.append(HEADER)
    for r in rows:
        ws.append(r)
    wb.save(OUT)
    print(f"\nWrote {len(rows)} stage(s) to {OUT}")


if __name__ == "__main__":
    main()
