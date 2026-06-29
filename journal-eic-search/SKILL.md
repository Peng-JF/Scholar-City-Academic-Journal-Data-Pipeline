---
name: journal-eic-search
description: Search for past editors-in-chief of academic journals listed in the RUC Core Journal Directory. Use when user mentions journal editors, editor-in-chief, 期刊主编, 历任主编, 期刊目录, journal EIC history, or wants to populate 期刊目录.xlsx with editorial leadership data from 中国人民大学核心期刊目录（2017年修订试行）.pdf.
---

# Journal Editor-in-Chief Search

This skill automates the process of finding and recording past editors-in-chief for academic journals. It parses the RUC Core Journal Directory PDF to extract international A+/A/A- journals in ECONOMICS/BUSINESS/BUSINESS,FINANCE/MANAGEMENT categories, writes them to `期刊目录.xlsx`, then searches the web for each journal's past editors-in-chief.

## Files

Only these files are used:
- `D:\Opencode\Scholar-city\中国人民大学核心期刊目录（2017年修订试行）.pdf`
- `D:\Opencode\Scholar-city\期刊目录.xlsx`

## Prerequisites

- Python 3 with: `pdfplumber`, `openpyxl`
- Use Exa web search (`exa_web_search_exa`) and web fetch (`exa_web_fetch_exa`)
- Scripts: `scripts/parse_pdf.py`, `scripts/write_excel.py`

---

## Excel Layout (Two-Sheet)

### Sheet1: 期刊元数据

One row per journal, metadata only. This is the benchmark table for Stata `merge`.

| Col | Header | Content |
|-----|--------|---------|
| A | 期刊编号 | Sequential ID |
| B | 期刊名 | Full journal name |
| C | 人大分级 | A+ / A / A- |
| D | 中科院分区 | Leave blank |
| E | JCR分区 | Leave blank |
| F | ISSN | Print ISSN |
| G | 信息来源 | Data source (e.g. "AEA官方 (aeaweb.org)") |
| H | 备注 | Special notes |

### Sheet2: EIC_Long (长数据)

One row per editor-in-chief, unlimited count. Stata-friendly: use `merge m:1 期刊编号 using Sheet1`.

| Col | Header | Content |
|-----|--------|---------|
| A | 期刊编号 | Matches Sheet1 col A |
| B | 期刊名 | Journal name |
| C | 人大分级 | A+ / A / A- |
| D | ISSN | Print ISSN |
| E | EIC序号 | 1 = current, 2 = previous, etc. |
| F | 姓名 | Editor's full name |
| G | 单位 | Institution |
| H | 单位所在城市 | City |
| I | 国籍 | ISO 3166-1 alpha-3 |
| J | 性别 | 男 / 女 |
| K | 上任时间 | Start year |
| L | 离任时间 | End year (empty if current) |
| M | 信息来源 | Data source URL/name |

---

## Workflow

### Phase 1: Parse PDF

```bash
python "scripts/parse_pdf.py" "D:\Opencode\Scholar-city\中国人民大学核心期刊目录（2017年修订试行）.pdf" --output "journals_temp.csv"
```

Verify the extracted list against the PDF. Check that only target categories (ECONOMICS, BUSINESS, BUSINESS,FINANCE, MANAGEMENT) are included.

### Phase 2: Populate Sheet1

```bash
python "scripts/write_excel.py" "D:\Opencode\Scholar-city\期刊目录.xlsx" --add-journals "journals_temp.csv"
```

### Phase 3: Search for EICs

For each journal, run parallel web searches:
1. `"[Journal Name]" wikipedia editor-in-chief`
2. `"[Journal Name]" editorial board past editors`
3. Publisher pages: `site:wiley.com`, `site:springer.com`, `site:elsevier.com`, `site:tandfonline.com`
4. Chinese sources: `"[Journal Name]" 主编 历任`
5. By ISSN: `"[ISSN]" journal editor-in-chief`

#### Coverage rule

Record editors going back to **2000 at minimum**. The number of rows varies per journal — there is no fixed count. A long-tenured editor covering 1985-2001 may mean only 2 editors suffice; a journal with 3-year terms may need 8+.

#### Accuracy rules

1. **NEVER fabricate data.** If a field is unknown, leave it blank or write "信息缺失".
2. **Verify from official sources**: journal website > publisher page > Wikipedia.
3. **Record the source** in `sheet1.G` and `sheet2.M`: e.g. "AEA官方 (aeaweb.org)".
4. **Check timelines are current**: always fetch the journal's official editorial board page.
5. **Gaps are normal.** If no EIC served for a period, note it in 备注.

#### Special cases (use 备注 col H)

- **Multiple co-editors**: Record the first-listed/primary editor. Note in 备注: "多位共同Editor(s)，此处列第一位"
- **Non-standard titles**: If the journal uses "Chair", "Managing Editor", "Executive Editor" instead of "Editor-in-Chief", note it: "采用Board Chair角色（非Managing Editor）"
- **Data gaps**: "2000-2012年间Lead Editor信息缺失"
- **Known pre-2000 editors not included**: "Oliver(1997-2003)覆盖2000，此处完整列出"

#### Determining locations

- US: Cambridge (MIT/Harvard), Palo Alto (Stanford), Berkeley (UCB), New Haven (Yale), Chicago (UChicago), Princeton, Evanston (Northwestern), Ann Arbor (Michigan), etc.
- UK: Cambridge, Oxford, London, Coventry (Warwick)
- Europe: Frankfurt, Barcelona, Zurich, Stockholm, Milan, Paris, Lausanne, Fontainebleau
- Do NOT use state/province/country as city.

### Phase 4: Write to Sheet2

```python
import sys
sys.path.insert(0, r"C:\Users\24882\.config\opencode\skills\journal-eic-search\scripts")
from write_excel import write_eics_long

write_eics_long("期刊目录.xlsx", "Journal Name", [
    {"name":"John Smith","institution":"Harvard University","city":"Cambridge",
     "nationality":"USA","gender":"男","start_year":"2020","end_year":""},
    {"name":"Jane Doe","institution":"Yale University","city":"New Haven",
     "nationality":"USA","gender":"女","start_year":"2015","end_year":"2020"},
], source="Wikipedia (en.wikipedia.org)")
```

EICs are sorted by start_year descending automatically. The function replaces all existing rows for that journal in Sheet2.

Save after each batch of 5-10 journals.

### Phase 5: Final Report

At end, run `--read` to verify:
```bash
python "scripts/write_excel.py" "期刊目录.xlsx" --read
```

Report: total journals, total EIC rows, journals with full 2000+ coverage, journals with gaps.

---

## Stata usage

```stata
import excel "期刊目录.xlsx", sheet("EIC_Long") firstrow clear
save "eic_long.dta", replace
import excel "期刊目录.xlsx", sheet("期刊元数据") firstrow clear
save "journal_meta.dta", replace
use "eic_long.dta", clear
merge m:1 期刊编号 using "journal_meta.dta"
```

## Reference: A+ Journals State (as of 2025-07)

| Journal | EICs | Covers 2000 | Notes |
|---------|------|-------------|-------|
| AER | 6 | 1985- ✓ | |
| JPE | 4 | 2012- ✗ | 2000-2012 gap |
| REStud | 2 | 2020- ✗ | Board Chair role; 2000-2020 gap |
| Econometrica | 7 | 2000- ✓ | |
| Management Science | 6 | 1997- ✓ | |
| AMR | 9 | 2000- ✓ | |
| Journal of Finance | 6 | 2000- ✓ | |
| QJE | 6 | 1989- ✓ | Multiple co-editors |
| AMJ | 9 | 2000- ✓ | |
| ASQ | 7 | 1997- ✓ | |
| Accounting Review | 8 | 1997- ✓ | |
