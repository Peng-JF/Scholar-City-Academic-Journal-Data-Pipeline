# Scholar-City: Academic Journal Data Pipeline

A toolkit for constructing large-scale bibliometric datasets from Web of Science, tracking journal editorial leadership, and building author-editor crosswalks for empirical research in the economics of science.

## Overview

This project provides three integrated modules that together form a complete pipeline from raw journal metadata to structured, analysis-ready data:

| Module | Function |
|--------|----------|
| **journal-eic-search** | Parse journal directories, search for past editors-in-chief, and record their institutions, cities, and tenure periods |
| **wos-paper-crawler** | Crawl Web of Science for all papers published in target journals (1995–present), extract metadata, and clean city/country/institution fields |
| **editor-name-crosswalk** | Build a name-variant mapping from editor canonical names to all author-name forms found in the paper database |

The pipeline is designed for research on editor turnover effects, same-city publication premiums, and academic network analysis. All output is in Stata-friendly long-format Excel.

## Prerequisites

- **Python 3.9+** with packages: `openpyxl`, `selenium`, `pandas`
- **Microsoft Edge** browser + [Edge WebDriver](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/) (for WoS crawling)
- Campus network access to Web of Science (required for the crawler)
- **Exa** API access (for web search in the EIC module; used via the opencode agent)

```bash
pip install openpyxl selenium pandas
```

## Installation

```bash
git clone https://github.com/YOUR_USERNAME/scholar-city.git
cd scholar-city
```

Copy the skill directories into your opencode skills folder, or use them as standalone Python scripts:

```bash
cp -r journal-eic-search wos-paper-crawler editor-name-crosswalk ~/.config/opencode/skills/
```

Alternatively, run the scripts directly:

```bash
python journal-eic-search/scripts/parse_pdf.py input.pdf --output journals.csv
python wos-paper-crawler/scripts/crawl_wos.py
python editor-name-crosswalk/scripts/build_crosswalk.py --eic-xlsx 期刊目录.xlsx --paper-xlsx WoS_Papers_All.xlsx --output 主编名称对照.xlsx
```

## Workflow

### Step 1: Build Journal Directory

Parse the journal classification PDF and create the journal metadata sheet.

```bash
python journal-eic-search/scripts/parse_pdf.py "期刊目录.pdf" --output "journals_temp.csv"
python journal-eic-search/scripts/write_excel.py "期刊目录.xlsx" --add-journals "journals_temp.csv"
```

### Step 2: Search for Editors-in-Chief

Use the `journal-eic-search` skill to find each journal's past editors-in-chief, their institutions, cities, and tenure start/end years. Results are written to `期刊目录.xlsx` Sheet2 (EIC_Long format).

### Step 3: Crawl Web of Science

Edit `wos-paper-crawler/scripts/crawl_wos.py` to set the target journal, output path, and stop time:

```python
OUT = r"D:\Scholar-city\WoS_JournalName.csv"
STOP_TIME = datetime(2026, 6, 1, 22, 0, 0)
JOURNAL = "American Economic Review"
START_YEAR = 1995
```

Then run:

```bash
python wos-paper-crawler/scripts/crawl_wos.py
```

The crawler supports checkpoint/resume, automatic browser restart on errors, and stop-at-time.

### Step 4: Postprocess Crawler Output

```bash
python wos-paper-crawler/scripts/postprocess.py --input WoS_JournalName.csv --output WoS_JournalName_long.xlsx
```

This performs:
- Country name normalization (WoS names → ISO 3166-1 alpha-3)
- City extraction from addresses (handles postal codes, abbreviations, multi-campus)
- Institution name cleaning (strips WoS number prefixes)
- City/country inference from institution names when missing
- Conversion to long-format Excel (Papers + Authors sheets)

### Step 5: Build Editor Name Crosswalk

```bash
python editor-name-crosswalk/scripts/build_crosswalk.py \
    --eic-xlsx 期刊目录.xlsx \
    --paper-xlsx WoS_Papers_All.xlsx \
    --output 主编名称对照.xlsx
```

This maps each editor's canonical name to all variant forms found in the paper database (e.g., "Acemoglu, Daron" ↔ "Acemoglu, D"), enabling identification of editor-authored papers.

## Output Format

### WoS_Papers_All.xlsx

**Papers Sheet:**
| Column | Content |
|--------|---------|
| paper_id | WoS unique identifier |
| journal_name | Journal name |
| article_title | Paper title |
| published_date | YYYY/MM |
| citations_wos_core | WoS Core citations |
| citations_all_db | All-database citations |
| cited_references | Number of references |
| document_type | Article, Review, etc. |

**Authors Sheet:**
| Column | Content |
|--------|---------|
| paper_id | Links to Papers sheet |
| author_seq | Author order (1, 2, ...) |
| author_name | "Lastname, Firstname" |
| institution | Primary affiliation |
| city | Extracted city |
| country | ISO 3166-1 alpha-3 |
| inst_canonical | Normalized institution name |

### 期刊目录.xlsx

**Sheet1 (期刊元数据):** Journal-level metadata (ISSN, tier, source).

**Sheet2 (EIC_Long):** One row per editor-term:
| Column | Content |
|--------|---------|
| 期刊名 | Journal name |
| 姓名 | Editor's name |
| 单位 | Primary institution |
| 单位所在城市 | Institution city |
| 上任时间 | Start year |
| 卸任时间 | End year |

### 主编名称对照.xlsx

| Column | Content |
|--------|---------|
| 主编规范名 | Canonical EIC name |
| 数据库中匹配到的名称变体 | Pipe-separated name variants found in database |
| 匹配数量 | Number of variants matched |
| 匹配状态 | 已匹配 / 未找到匹配 |

## File Structure

```
scholar-city/
├── README.md
├── journal-eic-search/
│   ├── SKILL.md
│   └── scripts/
│       ├── parse_pdf.py          # Parse journal directory PDF
│       └── write_excel.py        # Write journal metadata to xlsx
├── wos-paper-crawler/
│   ├── SKILL.md
│   └── scripts/
│       ├── crawl_wos.py          # Edge + Selenium WoS crawler
│       └── postprocess.py        # City/country cleanup + long-format export
└── editor-name-crosswalk/
    ├── SKILL.md
    └── scripts/
        └── build_crosswalk.py    # Editor name variant matching
```

## Technical Notes

### WoS Crawling

- Uses **Microsoft Edge** (not Chrome) due to campus network compatibility at the authors' institution. The Edge driver path is configurable in `crawl_wos.py`.
- The crawler operates on `webofscience.clarivate.cn` (Chinese WoS portal) but can be adapted to any WoS endpoint.
- Implements gradual page scrolling (50 results per page), autocomplete click, checkpoint/resume, automatic browser restart on empty pages, and stop-at-time.
- Raw output uses `utf-8-sig` encoding with GBK-safe fallback for Windows environments.

### City/Country Cleaning

The postprocessor handles several known data quality issues:
- Canadian postal codes (FSA → city mapping)
- European postal code prefixes embedded in city names ("CH-8006 Zurich" → "Zurich")
- Institution → city inference for known universities
- City-as-country-name cases ("U Arab Emirates" → ARE / Abu Dhabi)
- Institution name normalization (WoS "1 \nHarvard Univ" → "Harvard Univ")

### Editor Name Matching

The crosswalk builder generates name variants (surname-only, initials-only, full-first-name) and matches them against the paper author database using strict surname + initial rules to minimize false positives.

## License

MIT

## Citation

If you use this toolkit in your research, please cite:

```
[Your Name]. (2026). Scholar-City: Academic Journal Data Pipeline.
GitHub repository: https://github.com/YOUR_USERNAME/scholar-city
```
