---
name: wos-paper-crawler
description: Crawl Web of Science for research papers from journals, extracting metadata including title, authors, affiliations, citations, keywords, WoS categories from 1990 to present. Use when user mentions WoS crawling, 期刊论文爬取, paper metadata extraction, or wants to collect paper-level data for Stata analysis.
---

# WoS Paper Crawler

Crawls Web of Science for all research papers from specified journals (1990-present), extracting full metadata. Designed for the RUC Core Journal Directory project.

## Architecture

Three-stage pipeline:

1. **Crawl** (`scripts/crawl_wos.py`): Selenium browser automation → raw CSV
2. **Postprocess** (`scripts/postprocess.py`): Fix cities/countries, convert to long-format Excel
3. **Verify**: Compare with reference data, flag anomalies

## Prerequisites

- Python 3 with: `selenium`, `undetected-chromedriver`, `openpyxl`, `pandas`
- Chrome browser installed
- Access to `webofscience.clarivate.cn` via campus network (institutional login required)
- Files: `D:\Opencode\Scholar-city\期刊目录.xlsx` (journal list)

## Cost Estimate

| Item | Per journal (AER scale) |
|------|--------------------------|
| Articles since 1990 | ~3,000-5,000 |
| Result pages (50/page) | ~60-100 |
| Detail page loads | ~3,000-5,000 |
| Crawl time | ~2-4 hours per journal |
| Risk | Anti-bot blocking after ~500-1000 pages |

For ~100 journals: ~300,000-500,000 total pages. Budget ~1-2 weeks continuous crawling with anti-bot pauses.

---

## Phase 1: Crawl

### Run the crawler

```bash
python "scripts/crawl_wos.py" --journal "American Economic Review" --start-year 1990 --output "aer_raw.csv"
```

Or batch from Excel:
```bash
python "scripts/crawl_wos.py" --from-excel "D:\Opencode\Scholar-city\期刊目录.xlsx" --grade A+ --output-dir "raw_data/"
```

### Key improvements over legacy crawler

1. **Checkpoint/resume**: Saves progress every 50 papers. On restart, skips already-crawled papers.
2. **Anti-bot evasion**: Random delays (3-8s between pages), human-like scrolling, periodic browser restarts.
3. **Early Access handling**: Detects "Early Access" date and uses it when published date is missing.
4. **Citation 0 vs missing**: Distinguishes "0 citations" (real) from "citation count not found" (error).
5. **Dual citation counts**: Collects both "Times Cited, WoS Core" and "Times Cited, All Databases".
6. **All authors**: No 10-author limit; captures all authors listed on WoS.
7. **Document type filter**: Only collects "Article" type (excludes editorial, review, proceedings, etc.).

### Anti-bot strategy

When WoS starts returning empty search results (sign of blocking):
1. Pause crawling for 5-15 minutes (random)
2. Close and restart the browser
3. Resume from the last checkpoint
4. If repeated blocking: increase delay to 10-30s per page

### Monitoring

The crawler outputs a log line per paper:
```
[Page 12/45] Paper 23/50 - WOS:001234567800001 - Title: ... - OK
```

Watch for:
- `EMPTY` → WoS returned no results (likely blocked) → trigger anti-bot recovery
- `ERROR` → Page load failed → retry 3 times, then skip
- `STALE` → Browser session expired → restart browser

---

## Phase 2: Postprocess

### Fix cities and countries

The raw WoS data has inconsistent city/country naming (e.g., "England" vs "GBR", city missing/suburb listed). Run:

```bash
python "scripts/postprocess.py" --input "aer_raw.csv" --output "aer_clean.xlsx" --mode fix-locations
```

This applies:
1. Country name → ISO 3166-1 alpha-3 code (e.g., "United States" → "USA", "England" → "GBR")
2. City extraction: from address string → canonical city name
3. Multi-campus disambiguation: checks author affiliation/email on personal webpages
4. Flagging: marks unparseable addresses in 备注 column

### City determination rules

| Scenario | Approach |
|----------|----------|
| Standard "City, State, Country" format | Extract city before state |
| Suburb listed instead of city | Check institution homepage for main campus city |
| Multi-campus university (e.g. UC) | Match to specific campus: UC Berkeley → Berkeley, UC San Diego → San Diego |
| China: "中山大学深圳" vs "中山大学广州" | Check department/school name for campus hint; if ambiguous, search author's personal page |
| Singapore, Monaco, etc. (city-state) | Use country name as city |
| Tokyo (special ward vs city) | Use "Tokyo" for all 23 special wards |
| UK: "London" distributed across boroughs | Use "London" |

### Convert to long format

```bash
python "scripts/postprocess.py" --input "aer_clean.xlsx" --output "aer_long.xlsx" --mode to-long
```

Output: two-sheet Excel with Paper table + Author table (Stata-ready).

---

## Phase 3: Verify

Compare a random sample of papers against `期刊文章output - 人工核对版.xlsx`:

```bash
python "scripts/postprocess.py" --verify --new "aer_long.xlsx" --reference "期刊文章output - 人工核对版.xlsx" --sample 20
```

Reports discrepancies in citation counts, author names, affiliations.

---

## Output Format

### Paper sheet (论文元数据)

| Col | Field | Source |
|-----|-------|--------|
| A | paper_id | WoS accession number (WOS:...) |
| B | journal_name | From Excel or WoS |
| C | article_title | WoS title |
| D | source_journal | WoS source (may differ from journal_name for name changes) |
| E | published_date | YYYY/MM format |
| F | keywords | WoS author keywords + Keywords Plus |
| G | wos_categories | WoS subject categories |
| H | document_type | "Article" (filtered) |
| I | citations_wos_core | Times Cited, WoS Core Collection |
| J | citations_all_db | Times Cited, All Databases |
| K | early_access | TRUE/FALSE |
| L | 信息来源 | "WoS (webofscience.clarivate.cn)" |

### Author sheet (长数据)

| Col | Field |
|-----|-------|
| A | paper_id (link to Paper sheet) |
| B | author_seq | 1-based order |
| C | author_name | Full name |
| D | is_corresponding | TRUE/FALSE |
| E | institution | Parsed from address |
| F | city | Canonical city |
| G | country | ISO 3166-1 alpha-3 |
| H | 备注 | Any issues with this record |

## Special Cases

1. **Author with multiple affiliations**: Create one row per affiliation, same author_seq.
2. **Missing affiliation**: Leave institution/city/country blank, note in 备注.
3. **Journal name changes**: Use current name; note historical name in 备注.
4. **Pre-1990 papers**: Skip (target is 1990+).
5. **Early Access (no volume/issue)**: Mark early_access=TRUE, use available date.
