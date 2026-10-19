# site-link-audit: Claude Code handoff

> **Status:** TOOL. Runs by hand only (no schedule). **Where:** GitHub Actions, `ubuntu-latest`.
> **Trigger:** `workflow_dispatch` from the Actions tab ("Site Link Audit").
> **GitHub:** `allianceinstallation-bot/site-link-audit`. It is **PUBLIC**, the only public repo in the fleet.
> **Local clone:** `C:\SunshineCode\agents\site-link-audit` on branch `main` (`1a7a8b2`, last commit 2026-06-28). The tree is clean.
> **Last verified:** 2026-09-24. Last runs: 2026-06-23 to 06-29, 4 successes and 1 cancelled. It has not run since.

## What it does

It crawls a website from a seed URL (default `https://www.alliancewoodworking.com`) and follows every internal page. It checks every link it finds: internal pages, internal resources (images, CSS, JS, PDFs) and external links.

Output is one branded Excel workbook, uploaded as the run artifact `link-audit-report` (kept 30 days). It has four tabs:

1. **Summary**
2. **Broken Links**, with the page each one sits on
3. **All Links**
4. **Pages Crawled**

There is no external API, no secrets and no API cost. The crawler uses requests, BeautifulSoup and openpyxl.

## How it runs

- The workflow is `.github/workflows/site-link-audit.yml`. It checks out the repo, sets up Python 3.12, runs `pip install -r requirements.txt`, then `python crawl_site.py --url <input> --max-pages <input> --out link_audit.xlsx`, and uploads the xlsx.
- Inputs:
  - `url` (required, default is the company site)
  - `max_pages` (default 500, a safety cap on internal pages)
- A full run of the company site takes about 1 to 2.5 minutes.

## Repo layout

```
.github/workflows/site-link-audit.yml   the only workflow (manual)
crawl_site.py                           the crawler + Excel writer (~476 lines)
requirements.txt                        requests, beautifulsoup4, lxml, openpyxl (>= pins, not exact)
.gitignore                              __pycache__, *.pyc, *.xlsx, .venv, .pytest_cache
```

The local folder also holds a git-ignored `.venv\` and `__pycache__\`. There are no tests.

## Commands

```
# on GitHub
gh workflow run site-link-audit.yml -R allianceinstallation-bot/site-link-audit -f url=https://www.alliancewoodworking.com -f max_pages=500
gh run list -R allianceinstallation-bot/site-link-audit -L 5
gh run download <run-id> -R allianceinstallation-bot/site-link-audit -n link-audit-report

# locally
python -m venv .venv && .venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python crawl_site.py --url https://www.alliancewoodworking.com --max-pages 50 --out link_audit.xlsx
```

## Configuration and secrets

None. Settings are constants at the top of `crawl_site.py`:

| Constant | Value / meaning |
|---|---|
| `USER_AGENT` | Identifies as `SunshineAllianceLinkAudit/1.0` |
| `TIMEOUT` | 20 s |
| `DELAY` | 0.3 s between page fetches |
| `DEFAULT_MAX_PAGES` | 500 |
| `BLOCKED_CODES` | 401/403/405/429/503/999. An external link with one of these codes is reported as Blocked, not broken |
| `HINT_RELS` | `<link rel>` resource hints that are skipped |
| `NAVY` / `GOLD` / `CRIMSON` | Excel brand colours |

## External systems and data flow

The flow is: the public website, then the GitHub runner (the crawler), then the Actions artifact. It writes nothing anywhere else and sends no mail.

## Safety gates

- It is read-only by nature (HTTP GET/HEAD only).
- The politeness delay and the page cap protect the WordPress host.
- The site uses captcha forms and caching is OFF. Crawling at 500 pages is fine, but do not raise `DELAY` to 0 against production.

## Current state (2026-09-24)

- `gh pr list --state open`: none.
- `git status --short`: clean.
- Branches: `main` only (local, and `origin/main`).
- PR #1 (dedup workflow, double fetch) and PR #2 (false positives: resource hints, anti-bot codes) are merged.

## Traps and lessons

- **False positives were the big problem, and both are fixed in PR #2:**
  - `dns-prefetch`/`preconnect` hints pointing at a bare origin (e.g. `https://fonts.googleapis.com/`) returned 404s.
  - Yelp, LinkedIn, Instagram and Cloudflare sites return 403/429/999 to bots.
  - Do not "simplify" those two rules away.
- An old pre-repo copy of `crawl_site.py` sits in SharePoint `_000 CLAUDE\99_ARCHIVE\site-link-audit-prerepo\`. The repo is canonical.
- `*.xlsx` is ignored, so local test outputs never get committed.
- Site edits themselves happen in wp-admin, not here. See memory `website-live-fixes-2026-09-17` and skill `sunshine-website`. Run this crawler after a site change as the regression check (the 09-17 fix pass used 253/253 URLs = 200).

## Moving to the org

- **Keep it public** if possible. It holds no secrets and no business data, and public repos do not consume the `sunshine-alliance` Free-plan pool of 2,000 Actions minutes/month.
- If it must be private, each run costs about 1 to 3 billed minutes. That is fine for a manual tool.
- Transfer command: `gh api -X POST repos/allianceinstallation-bot/site-link-audit/transfer -f new_owner=sunshine-alliance`. Then run `git remote set-url origin https://github.com/sunshine-alliance/site-link-audit.git` locally.
- There is no cron, so the schedule actor trap does not apply.
- Optional: pin exact versions in `requirements.txt` (the org template convention).

## Related docs

- `sunshine-design-system.md` covers the web brand tokens. This crawler's Excel uses the report palette (navy/gold/crimson), not the web palette.
- `00_START_HERE.md` and `01_FLEET_MAP.md` in this pack.
