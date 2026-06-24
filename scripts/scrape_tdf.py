#!/usr/bin/env python3
"""
DIAGNOSTIC build. Figures out why scraping returns nothing:
  1) prints the installed procyclingstats version
  2) does a raw HTTP request to PCS (with a browser User-Agent) and reports the
     status code + a snippet  -> tells us if PCS is blocking the runner
  3) tries the library and prints the FULL error to stdout (not stderr)
If the library works it still writes the results file.

    pip install procyclingstats openpyxl
    YEAR=2025 OUT=tdf-results-2025-test.xlsx python scripts/scrape_tdf.py
"""
import os
import sys
import traceback
import urllib.request

from openpyxl import Workbook

YEAR = os.environ.get("YEAR", "2026")
OUT = os.environ.get("OUT", "tdf-results.xlsx")
RACE = os.environ.get("RACE_SLUG", "tour-de-france")
MAX_STAGES = int(os.environ.get("MAX_STAGES", "21"))
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

print("=== ENV ===")
print(f"YEAR={YEAR} OUT={OUT} RACE={RACE}")
try:
    from importlib.metadata import version
    print("procyclingstats version:", version("procyclingstats"))
except Exception as e:
    print("version lookup failed:", e)

# ---- raw probe ----
def probe(url):
    full = "https://www.procyclingstats.com/" + url
    print(f"\n=== RAW PROBE {full} ===")
    try:
        req = urllib.request.Request(full, headers={"User-Agent": UA})
        with urllib.request.urlopen(req, timeout=25) as r:
            body = r.read().decode("utf-8", "replace")
            print(f"HTTP {r.status}, {len(body)} bytes")
            low = body.lower()
            print("looks like results table:", ("resulttable" in low or 'class="results' in low or "<table" in low))
            print("looks blocked (cf/captcha):", ("cloudflare" in low or "captcha" in low or "just a moment" in low))
            print("snippet:", " ".join(body[:300].split()))
    except Exception as e:
        print("RAW PROBE ERROR:", e)

probe(f"race/{RACE}/{YEAR}/stage-1")

# ---- library attempt ----
print("\n=== LIBRARY ATTEMPT (stage 1) ===")
Stage = None
try:
    from procyclingstats import Stage as _Stage
    Stage = _Stage
except Exception:
    print("import procyclingstats FAILED:")
    traceback.print_exc(file=sys.stdout)

if Stage:
    try:
        s = Stage(f"race/{RACE}/{YEAR}/stage-1")
        try:
            p = s.parse()
            print("parse() keys:", list(p.keys()) if isinstance(p, dict) else type(p))
        except Exception:
            print("parse() traceback:")
            traceback.print_exc(file=sys.stdout)
        try:
            res = s.results()
            print("results() type:", type(res), "len:", len(res) if res else 0)
            if res:
                print("first result row:", res[0])
        except Exception:
            print("results() traceback:")
            traceback.print_exc(file=sys.stdout)
    except Exception:
        print("Stage() construction traceback:")
        traceback.print_exc(file=sys.stdout)

# ---- name helpers ----
def fmt_name(pcs_name):
    if not pcs_name:
        return ""
    toks = str(pcs_name).split()
    surname = [t for t in toks if t == t.upper()]
    given = [t for t in toks if t != t.upper()]
    return (" ".join(given) + " " + " ".join(w.capitalize() for w in surname)).strip()

def name_of(row):
    if isinstance(row, dict):
        for k in ("rider_name", "rider", "name"):
            if row.get(k):
                return row[k]
    return ""

def rank_of(row, fb):
    if isinstance(row, dict):
        for k in ("rank", "position", "place"):
            if row.get(k) not in (None, ""):
                try:
                    return int(row[k])
                except (TypeError, ValueError):
                    pass
    return fb

def ranked(rows, n):
    out = [""] * n
    for i, row in enumerate(rows or []):
        r = rank_of(row, i + 1)
        if 1 <= r <= n:
            out[r - 1] = fmt_name(name_of(row))
    return out

def cls(stage, m):
    try:
        return getattr(stage, m)() or []
    except Exception:
        return []

HEADER = (["Date", "Stage"] + ["1st","2nd","3rd","4th","5th","6th","7th","8th","9th","10th"]
          + [f"GC #{i}" for i in range(1,11)] + [f"Points #{i}" for i in range(1,4)]
          + [f"Mountain #{i}" for i in range(1,4)] + [f"Youth #{i}" for i in range(1,4)])

def main():
    if not Stage:
        return
    rows = []
    for n in range(1, MAX_STAGES + 1):
        try:
            st = Stage(f"race/{RACE}/{YEAR}/stage-{n}")
            results = cls(st, "results")
            if not results:
                continue
            try:
                date = st.date()
            except Exception:
                date = ""
            row = [date, n] + ranked(results, 10) + ranked(cls(st, "gc"), 10) \
                + ranked(cls(st, "points"), 3) + ranked(cls(st, "kom"), 3) + ranked(cls(st, "youth"), 3)
            rows.append(row)
            print(f"stage {n}: {date} winner {row[2]}")
        except Exception as e:
            print(f"stage {n}: error {e}")
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
