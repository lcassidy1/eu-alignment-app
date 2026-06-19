"""
EU Delegation UN New York – Alignment Statement Scraper — GA79 edition
Covers Sept 10, 2024 – Sept 9, 2025.

Discovers URLs from the EEAS sitemap (lastmod >= GA79 start), scrapes each page,
filters by actual publication date, detects alignment clauses, writes Excel.

Resumable: results are cached in scraped_results_ga79.csv, so if the run is
interrupted (rate limiting), just run it again and it continues where it left off.

Usage:
    python3 eu_alignment_scraper_ga79.py
"""

import re
import time
import urllib.error
import urllib.request
from datetime import date, datetime
from pathlib import Path

import pandas as pd

# ── Configuration ─────────────────────────────────────────────────────────────

GA79_START  = date(2024, 9, 10)
GA79_END    = date(2025, 9, 9)
OUTPUT_FILE = Path(__file__).parent / "eu_alignment_stats_GA79.xlsx"
RESULTS_CACHE = Path(__file__).parent / "scraped_results_ga79.csv"
SKIP_CACHE    = Path(__file__).parent / "skipped_urls_ga79.txt"
URL_LIST      = Path(__file__).parent / "statement_urls_ga79.txt"
SITEMAP_PAGES = 21

TRACKED_COUNTRIES = {
    "Turkey":                 r"\bTurkey\b|Türkiye",
    "North Macedonia":        r"North Macedonia",
    "Montenegro":             r"\bMontenegro\b",
    "Serbia":                 r"\bSerbia\b",
    "Albania":                r"\bAlbania\b",
    "Ukraine":                r"\bUkraine\b",
    "Moldova":                r"Republic of Moldova|\bMoldova\b",
    "Bosnia and Herzegovina": r"Bosnia and Herzegovina",
    "Georgia":                r"\bGeorgia\b",
    "Iceland":                r"\bIceland\b",
    "Liechtenstein":          r"\bLiechtenstein\b",
    "Norway":                 r"\bNorway\b",
    "Armenia":                r"\bArmenia\b",
    "Azerbaijan":             r"\bAzerbaijan\b",
    "Andorra":                r"\bAndorra\b",
    "Monaco":                 r"\bMonaco\b",
    "San Marino":             r"San Marino",
    "UK":                     r"United Kingdom|\bUK\b",
}

ALIGNMENT_PATTERN = re.compile(
    r"\b(align|associate)\b[^.!?]{0,60}\bthemselves\b",
    re.IGNORECASE,
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

SKIP_PATTERNS = [
    "/global-gateway", "/about-ambassador", "/who-we-are",
    "/european-union-and-united-nations", "/meet-eu-youth",
    "/newsletter", "/vacancy", "/job-",
]


# ── URL discovery from sitemap ────────────────────────────────────────────────

def collect_urls():
    """Scan sitemap for UN NY pages with lastmod >= GA79 start.

    No upper bound on lastmod: a GA79-era page edited recently still belongs
    to GA79 — the real filter is the publication date read during scraping.
    Cached in statement_urls_ga79.txt so we only scan the sitemap once.
    """
    if URL_LIST.exists():
        urls = URL_LIST.read_text().splitlines()
        print(f"Loaded {len(urls)} URLs from {URL_LIST.name} (delete it to re-scan)")
        return urls

    entry_pat = re.compile(r"<url>(.*?)</url>", re.DOTALL)
    loc_pat = re.compile(r"<loc>([^<]+)</loc>")
    lastmod_pat = re.compile(r"<lastmod>([^<]+)</lastmod>")

    seen, urls = set(), []
    for pg in range(1, SITEMAP_PAGES + 1):
        url = f"https://www.eeas.europa.eu/sitemap.xml?page={pg}"
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            text = urllib.request.urlopen(req, timeout=15).read().decode("utf-8", errors="replace")
        except Exception as e:
            print(f"  Sitemap page {pg}: ERROR {e}")
            time.sleep(3)
            continue

        count = 0
        for entry in entry_pat.findall(text):
            loc_m = loc_pat.search(entry)
            if not loc_m:
                continue
            loc = loc_m.group(1).split("?")[0]
            if "/delegations/un-new-york/" not in loc:
                continue
            if not loc.endswith("_en"):
                continue
            if any(s in loc for s in SKIP_PATTERNS):
                continue
            lastmod_m = lastmod_pat.search(entry)
            try:
                lastmod = datetime.strptime(lastmod_m.group(1)[:10], "%Y-%m-%d").date() if lastmod_m else None
            except Exception:
                lastmod = None
            if lastmod and lastmod >= GA79_START and loc not in seen:
                seen.add(loc)
                urls.append(loc)
                count += 1

        print(f"  Sitemap page {pg:2d}: +{count} (total {len(urls)})")
        time.sleep(1.0)

    URL_LIST.write_text("\n".join(urls))
    print(f"\nSaved {len(urls)} candidate URLs to {URL_LIST.name}")
    return urls


# ── Page scraping ─────────────────────────────────────────────────────────────

def fetch_page(url):
    """Returns (html, gone). gone=True means permanent 404/410 — never retry."""
    req = urllib.request.Request(url, headers=HEADERS)
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.read().decode("utf-8", errors="replace"), False
    except urllib.error.HTTPError as e:
        return None, e.code in (404, 410)
    except Exception:
        return None, False


def parse_date(html):
    m = re.search(r'article:published_time"\s+content="(\d{4}-\d{2}-\d{2})', html)
    if m:
        try:
            return datetime.strptime(m.group(1), "%Y-%m-%d").date()
        except ValueError:
            pass
    m = re.search(r"\b(\d{2}\.\d{2}\.(202[0-9]))\b", html)
    if m:
        try:
            return datetime.strptime(m.group(1), "%d.%m.%Y").date()
        except ValueError:
            pass
    return None


def extract_title(html):
    m = re.search(r"<h1[^>]*>(.*?)</h1>", html, re.DOTALL | re.IGNORECASE)
    if m:
        return re.sub(r"<[^>]+>", "", m.group(1)).strip()
    return ""


def extract_body_text(html):
    html = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", html, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", html)
    return re.sub(r"\s+", " ", text).strip()


def extract_alignment_info(text):
    sentences = re.split(r"(?<=[.!?])\s+", text)
    alignment_sentences = [s for s in sentences if ALIGNMENT_PATTERN.search(s)]
    if not alignment_sentences:
        return False, {c: False for c in TRACKED_COUNTRIES}
    combined = " ".join(alignment_sentences)
    flags = {
        country: bool(re.search(pattern, combined, re.IGNORECASE))
        for country, pattern in TRACKED_COUNTRIES.items()
    }
    return True, flags


# ── Cache ─────────────────────────────────────────────────────────────────────

def load_cache():
    rows, done = [], set()
    if RESULTS_CACHE.exists():
        cached = pd.read_csv(RESULTS_CACHE)
        rows = cached.to_dict("records")
        done = set(cached["URL"])
    if SKIP_CACHE.exists():
        done |= set(SKIP_CACHE.read_text().splitlines())
    return rows, done


def save_cache(rows):
    pd.DataFrame(rows).to_csv(RESULTS_CACHE, index=False)


def mark_skipped(url):
    with open(SKIP_CACHE, "a") as f:
        f.write(url + "\n")


# ── Main loop ─────────────────────────────────────────────────────────────────

def scrape_all(urls):
    urls = list(dict.fromkeys(urls))
    rows, done = load_cache()
    todo = [u for u in urls if u not in done]
    total = len(todo)
    print(f"  Cached: {len(rows)} statements ({len(done)} URLs done, {total} to fetch)\n")

    failed = 0
    for i, url in enumerate(todo, 1):
        slug = url.rstrip("/").split("/")[-1][:55]
        print(f"  [{i:3d}/{total}] {slug}", end=" ... ", flush=True)

        html, gone = fetch_page(url)
        if html is None:
            if gone:
                print("404 (skipping permanently)")
                mark_skipped(url)
                continue
            print("FAILED")
            failed += 1
            if failed >= 10:
                print("\n  Too many consecutive failures — likely rate-limited.")
                print("  Progress saved; re-run later to continue.")
                break
            time.sleep(2)
            continue
        failed = 0

        stmt_date = parse_date(html)
        if stmt_date is None:
            print("no date")
            mark_skipped(url)
            continue
        if not (GA79_START <= stmt_date <= GA79_END):
            print(f"skip ({stmt_date})")
            mark_skipped(url)
            continue

        title = extract_title(html)
        text = extract_body_text(html)
        has_alignment, flags = extract_alignment_info(text)

        row = {
            "Date": str(stmt_date),
            "Title": title or slug,
            "URL": url,
            "Has Alignment": has_alignment,
        }
        row.update(flags)
        rows.append(row)
        save_cache(rows)

        print(f"OK [{'ALIGNMENT' if has_alignment else 'no alignment'}] ({stmt_date})")
        time.sleep(0.5)

    print(f"\nTotal statements (cached + new): {len(rows)}")
    return rows


# ── Excel output ──────────────────────────────────────────────────────────────

def write_excel(rows):
    df = pd.DataFrame(rows)
    df["_sort"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.sort_values("_sort", ascending=False).reset_index(drop=True)
    df["Date"] = df["_sort"].dt.strftime("%d/%m/%Y")
    df = df.drop(columns=["_sort"])

    country_cols = list(TRACKED_COUNTRIES.keys())
    df_aligned = df[df["Has Alignment"] == True].copy()
    n = len(df_aligned)

    summary = []
    for c in country_cols:
        cnt = int(df_aligned[c].sum()) if c in df_aligned.columns else 0
        summary.append({
            "Country": c,
            "Times Aligned": cnt,
            "Statements with Alignment Clause": n,
            "Alignment %": round(cnt / n * 100, 1) if n else 0.0,
        })
    df_summary = pd.DataFrame(summary).sort_values("Alignment %", ascending=False)

    raw_cols = ["Date", "Title", "URL", "Has Alignment"] + [c for c in country_cols if c in df.columns]

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="Summary", index=False)
        ws = writer.sheets["Summary"]
        for col, w in [("A", 22), ("B", 16), ("C", 38), ("D", 14)]:
            ws.column_dimensions[col].width = w

        df[raw_cols].to_excel(writer, sheet_name="All Statements", index=False)
        ws2 = writer.sheets["All Statements"]
        for col, w in [("A", 13), ("B", 65), ("C", 75), ("D", 15)]:
            ws2.column_dimensions[col].width = w

        if n:
            df_aligned[raw_cols].to_excel(writer, sheet_name="Alignment Statements Only", index=False)
            ws3 = writer.sheets["Alignment Statements Only"]
            for col, w in [("A", 13), ("B", 65), ("C", 75)]:
                ws3.column_dimensions[col].width = w

    print(f"\nSaved: {OUTPUT_FILE}")
    print(f"\n{'='*60}")
    print(f"Total GA79 statements:             {len(df)}")
    print(f"Statements WITH alignment clause:  {n}")
    if n:
        print(f"\nTop aligners:")
        print(df_summary[df_summary['Times Aligned'] > 0].to_string(index=False))


if __name__ == "__main__":
    print(f"GA79 window: {GA79_START} to {GA79_END}\n")
    print("Step 1/2: Collecting candidate URLs from EEAS sitemap...")
    urls = collect_urls()
    print(f"\nStep 2/2: Scraping {len(urls)} pages...\n")
    rows = scrape_all(urls)
    if rows:
        write_excel(rows)
    else:
        print("Nothing scraped yet — re-run to continue.")
