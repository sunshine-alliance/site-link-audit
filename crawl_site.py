#!/usr/bin/env python3
"""
Sunshine Alliance - Website Link Audit Crawler
-----------------------------------------------
Crawls a website starting from a seed URL, follows every internal page,
and checks every link it finds: internal pages, internal resources
(images, CSS, JS, PDFs), and external links to other sites. It records
the HTTP status of each and flags anything broken.

Output: a single branded Excel workbook with four tabs:
  1. Summary        - scan metadata and the headline counts
  2. Broken Links   - only the problems, with the page each one sits on
  3. All Links      - the full audit trail of every link checked
  4. Pages Crawled  - every internal page visited and its status

No external API. Pure requests + BeautifulSoup + openpyxl.
Designed to run on GitHub Actions (which has the web access needed to crawl).
"""

import sys
import time
import argparse
from collections import deque
from datetime import datetime, timezone
from urllib.parse import urljoin, urlparse, urldefrag

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Brand
# --------------------------------------------------------------------------
NAVY = "1F3864"
GOLD = "C9A44A"
CRIMSON = "8B1E3F"
WHITE = "FFFFFF"
LIGHT_RED = "F8D7DA"
LIGHT_GREY = "F2F2F2"

# --------------------------------------------------------------------------
# Crawl settings
# --------------------------------------------------------------------------
USER_AGENT = "Mozilla/5.0 (compatible; SunshineAllianceLinkAudit/1.0; +https://www.alliancewoodworking.com)"
TIMEOUT = 20          # seconds per request
DELAY = 0.3           # politeness pause between page fetches
DEFAULT_MAX_PAGES = 500

SKIP_SCHEMES = ("mailto:", "tel:", "javascript:", "data:", "sms:", "ftp:", "file:")
RESOURCE_EXTS = (
    ".css", ".js", ".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".ico",
    ".pdf", ".woff", ".woff2", ".ttf", ".eot", ".otf", ".mp4", ".webm", ".mov",
    ".zip", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
)

# Caches / collectors
status_cache = {}     # url -> dict(code, reason, ok, final_url)
link_records = []     # list of dicts, one per discovered link occurrence
pages_crawled = []    # list of dicts, one per internal page visited
queued_pages = set()  # internal page urls already queued or crawled


def reg_host(netloc):
    """Normalize a host for same-site comparison (drop leading www.)."""
    return netloc.lower().split(":")[0].replace("www.", "", 1) if netloc else ""


def normalize(url):
    """Strip fragment and collapse a bare-domain empty path to '/'.

    This makes https://site.com and https://site.com/ resolve to one url so
    the homepage is not crawled or reported twice.
    """
    url, _frag = urldefrag(url)
    url = url.strip()
    p = urlparse(url)
    if p.scheme in ("http", "https") and p.path == "":
        url = url + "/"
    return url


def classify(url, base_host):
    """Return one of: internal-page, internal-resource, external."""
    p = urlparse(url)
    host = reg_host(p.netloc)
    if host and host != base_host:
        return "external"
    path = p.path.lower()
    if path.endswith(RESOURCE_EXTS):
        return "internal-resource"
    return "internal-page"


def check_url(url, session):
    """HEAD-first, GET-fallback status check. Cached per unique url."""
    if url in status_cache:
        return status_cache[url]
    code, reason, ok, final_url = None, "", False, url
    try:
        r = session.head(url, allow_redirects=True, timeout=TIMEOUT)
        # Many servers do not support HEAD (405) or lie about it; fall back.
        if r.status_code >= 400 or r.status_code in (403, 405, 999):
            r = session.get(url, allow_redirects=True, timeout=TIMEOUT, stream=True)
            r.close()
        code = r.status_code
        final_url = r.url
        reason = r.reason or ""
        ok = code < 400
    except requests.exceptions.SSLError as e:
        reason = "SSL error: " + str(e)[:160]
    except requests.exceptions.ConnectionError as e:
        reason = "Connection error: " + str(e)[:160]
    except requests.exceptions.Timeout:
        reason = "Timeout after %ss" % TIMEOUT
    except requests.exceptions.TooManyRedirects:
        reason = "Too many redirects (loop)"
    except requests.exceptions.RequestException as e:
        reason = "Request error: " + str(e)[:160]
    result = {"code": code, "reason": reason, "ok": ok, "final_url": final_url}
    status_cache[url] = result
    return result


def extract_links(soup, page_url):
    """Yield (raw_value, anchor_text, source_tag) for every linkish element."""
    for a in soup.find_all("a", href=True):
        text = (a.get_text() or "").strip()
        yield a["href"], text[:120], "a"
    for img in soup.find_all("img", src=True):
        yield img["src"], (img.get("alt") or "")[:120], "img"
    for link in soup.find_all("link", href=True):
        rel = " ".join(link.get("rel", []))
        yield link["href"], rel[:120], "link"
    for s in soup.find_all("script", src=True):
        yield s["src"], "", "script"


def crawl(seed, max_pages):
    base_host = reg_host(urlparse(seed).netloc)
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})

    to_crawl = deque([seed])
    queued_pages.add(seed)

    while to_crawl and len(pages_crawled) < max_pages:
        page = to_crawl.popleft()
        page_status = check_url(page, session)

        # Fetch the page body so we can parse it for more links.
        html, content_type = None, ""
        try:
            r = session.get(page, allow_redirects=True, timeout=TIMEOUT)
            content_type = r.headers.get("Content-Type", "")
            if r.status_code < 400 and "html" in content_type.lower():
                html = r.text
            page_status = {
                "code": r.status_code,
                "reason": r.reason or "",
                "ok": r.status_code < 400,
                "final_url": r.url,
            }
            status_cache[page] = page_status
        except requests.exceptions.RequestException as e:
            page_status = {"code": None, "reason": "Fetch error: " + str(e)[:160],
                           "ok": False, "final_url": page}
            status_cache[page] = page_status

        pages_crawled.append({
            "url": page,
            "code": page_status["code"],
            "ok": page_status["ok"],
            "reason": page_status["reason"],
            "content_type": content_type,
        })

        sys.stdout.write("[crawl] %3d  %s  %s\n" % (
            len(pages_crawled),
            page_status["code"] if page_status["code"] is not None else "ERR",
            page))
        sys.stdout.flush()

        if not html:
            time.sleep(DELAY)
            continue

        soup = BeautifulSoup(html, "lxml")
        for raw, anchor, tag in extract_links(soup, page):
            raw = (raw or "").strip()
            if not raw or raw.startswith("#"):
                continue
            if any(raw.lower().startswith(s) for s in SKIP_SCHEMES):
                continue

            target = normalize(urljoin(page, raw))
            if not target.lower().startswith(("http://", "https://")):
                continue

            ltype = classify(target, base_host)
            st = check_url(target, session)

            link_records.append({
                "source": page,
                "target": target,
                "anchor": anchor,
                "tag": tag,
                "type": ltype,
                "code": st["code"],
                "ok": st["ok"],
                "reason": st["reason"],
                "final_url": st["final_url"] if st["final_url"] != target else "",
            })

            if (ltype == "internal-page" and target not in queued_pages
                    and len(queued_pages) < max_pages):
                queued_pages.add(target)
                to_crawl.append(target)

        time.sleep(DELAY)

    return base_host


# --------------------------------------------------------------------------
# Excel report
# --------------------------------------------------------------------------
THIN = Side(style="thin", color="D0D0D0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name="Arial", bold=True, color=WHITE, size=10)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 24


def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_report(out_path, seed, base_host, max_pages):
    broken = [r for r in link_records if not r["ok"]]
    # Dedup broken by (source, target) so the same nav link on every page
    # is not counted dozens of times in the headline figure.
    uniq_broken_pairs = {(r["source"], r["target"]) for r in broken}

    n_pages = len(pages_crawled)
    n_pages_broken = sum(1 for p in pages_crawled if not p["ok"])
    n_links = len(link_records)
    n_unique_targets = len(status_cache)
    b_internal = sum(1 for r in broken if r["type"] == "internal-page")
    b_resource = sum(1 for r in broken if r["type"] == "internal-resource")
    b_external = sum(1 for r in broken if r["type"] == "external")

    wb = Workbook()

    # ---- Tab 1: Summary ----
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False
    ws["A1"] = "SUNSHINE ALLIANCE  |  Website Link Audit"
    ws["A1"].font = Font(name="Arial", bold=True, size=16, color=NAVY)
    ws["A2"] = "Luxury Cabinetry"
    ws["A2"].font = Font(name="Arial", italic=True, size=10, color=GOLD)

    rows = [
        ("", ""),
        ("Site crawled", seed),
        ("Scan date (UTC)", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M")),
        ("Max pages setting", str(max_pages)),
        ("", ""),
        ("Pages crawled", n_pages),
        ("Pages returning an error", n_pages_broken),
        ("Unique URLs checked", n_unique_targets),
        ("Total link occurrences", n_links),
        ("", ""),
        ("BROKEN - internal pages", b_internal),
        ("BROKEN - internal resources (img/css/js/pdf)", b_resource),
        ("BROKEN - external links", b_external),
        ("Unique broken links (page + target)", len(uniq_broken_pairs)),
    ]
    r0 = 4
    for i, (label, val) in enumerate(rows):
        rr = r0 + i
        ws.cell(row=rr, column=1, value=label).font = Font(
            name="Arial", bold=label.startswith("BROKEN") or label in (
                "Pages crawled", "Site crawled"), size=10,
            color=CRIMSON if label.startswith("BROKEN") and val else NAVY)
        vcell = ws.cell(row=rr, column=2, value=val)
        vcell.font = Font(name="Arial", size=10)
        if label.startswith("BROKEN") and isinstance(val, int) and val > 0:
            vcell.font = Font(name="Arial", bold=True, size=10, color=CRIMSON)
    autosize(ws, [46, 70])

    # ---- Tab 2: Broken Links ----
    ws = wb.create_sheet("Broken Links")
    ws.sheet_view.showGridLines = False
    headers = ["Source page (where the link lives)", "Broken link URL",
               "Link text / alt", "Type", "Status", "Reason", "Redirects to"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)
    style_header(ws, len(headers))
    ws.freeze_panes = "A2"
    if broken:
        broken_sorted = sorted(broken, key=lambda r: (r["source"], r["type"], r["target"]))
        for i, r in enumerate(broken_sorted, start=2):
            vals = [r["source"], r["target"], r["anchor"], r["type"],
                    r["code"] if r["code"] is not None else "ERR",
                    r["reason"], r["final_url"]]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=i, column=c, value=v)
                cell.font = Font(name="Arial", size=9)
                cell.alignment = Alignment(vertical="top", wrap_text=(c in (1, 2, 6)))
                cell.fill = PatternFill("solid", fgColor=LIGHT_RED)
                cell.border = BORDER
    else:
        ws.cell(row=2, column=1, value="No broken links found. Clean scan.").font = Font(
            name="Arial", bold=True, size=11, color="2E7D32")
    autosize(ws, [42, 46, 26, 16, 8, 34, 36])

    # ---- Tab 3: All Links ----
    ws = wb.create_sheet("All Links")
    ws.sheet_view.showGridLines = False
    headers = ["Source page", "Link URL", "Link text / alt", "Tag", "Type",
               "Status", "OK", "Reason", "Redirects to"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)
    style_header(ws, len(headers))
    ws.freeze_panes = "A2"
    for i, r in enumerate(sorted(link_records, key=lambda r: (r["source"], r["target"])), start=2):
        vals = [r["source"], r["target"], r["anchor"], r["tag"], r["type"],
                r["code"] if r["code"] is not None else "ERR",
                "yes" if r["ok"] else "NO", r["reason"], r["final_url"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.font = Font(name="Arial", size=9)
            cell.alignment = Alignment(vertical="top", wrap_text=(c in (1, 2)))
            cell.border = BORDER
            if not r["ok"]:
                cell.fill = PatternFill("solid", fgColor=LIGHT_RED)
            elif i % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=LIGHT_GREY)
    autosize(ws, [40, 44, 24, 8, 16, 8, 6, 30, 32])

    # ---- Tab 4: Pages Crawled ----
    ws = wb.create_sheet("Pages Crawled")
    ws.sheet_view.showGridLines = False
    headers = ["Page URL", "Status", "OK", "Reason", "Content-Type"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)
    style_header(ws, len(headers))
    ws.freeze_panes = "A2"
    for i, p in enumerate(sorted(pages_crawled, key=lambda p: p["url"]), start=2):
        vals = [p["url"], p["code"] if p["code"] is not None else "ERR",
                "yes" if p["ok"] else "NO", p["reason"], p["content_type"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.font = Font(name="Arial", size=9)
            cell.alignment = Alignment(vertical="top", wrap_text=(c == 1))
            cell.border = BORDER
            if not p["ok"]:
                cell.fill = PatternFill("solid", fgColor=LIGHT_RED)
    autosize(ws, [56, 8, 6, 30, 28])

    wb.save(out_path)
    return {
        "pages": n_pages, "links": n_links,
        "broken_unique": len(uniq_broken_pairs),
        "b_internal": b_internal, "b_resource": b_resource, "b_external": b_external,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", required=True, help="Seed URL, include https://")
    ap.add_argument("--max-pages", type=int, default=DEFAULT_MAX_PAGES)
    ap.add_argument("--out", default="link_audit.xlsx")
    args = ap.parse_args()

    seed = args.url.strip()
    if not seed.lower().startswith(("http://", "https://")):
        seed = "https://" + seed
    seed = normalize(seed)

    print("=" * 64)
    print("Sunshine Alliance - Website Link Audit")
    print("Seed:", seed, "| Max pages:", args.max_pages)
    print("=" * 64)

    base_host = crawl(seed, args.max_pages)
    summary = write_report(args.out, seed, base_host, args.max_pages)

    print("-" * 64)
    print("Pages crawled        :", summary["pages"])
    print("Link occurrences     :", summary["links"])
    print("Broken (unique)      :", summary["broken_unique"])
    print("  internal pages     :", summary["b_internal"])
    print("  internal resources :", summary["b_resource"])
    print("  external links     :", summary["b_external"])
    print("Report written       :", args.out)
    print("-" * 64)


if __name__ == "__main__":
    main()
