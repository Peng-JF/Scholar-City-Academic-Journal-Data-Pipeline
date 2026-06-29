---
name: editor-name-crosswalk
description: Build a name crosswalk mapping journal editors-in-chief to all name variants found in the paper database, enabling identification of editor-authored papers for later exclusion or robustness checks. Use when user mentions 主编名称对照, editor name matching, 主编论文识别, or wants to flag papers authored by editors.
---

# Editor Name Crosswalk

Builds a mapping from editor-in-chief canonical names to all name variants found in the paper database. This enables Stata-based identification of editor-authored papers for robustness checks.

## Purpose

Research question: do editor-authored papers or papers from the editor's institution/city affect journal outcomes differently? To answer this, we need to:

1. Identify which papers have editors as authors → use this crosswalk
2. Identify papers from editor's institution → use affiliation data
3. Identify papers from editor's city → use city data
4. Run regressions with/without these papers as robustness

## Files

Input:
- `D:\Opencode\Scholar-city\期刊目录.xlsx` (Sheet2 EIC_Long: canonical EIC names)
- `D:\Opencode\Scholar-city\期刊文章output.xlsx` (Author 1-10 + Corresponding Author columns)

Output:
- `D:\Opencode\Scholar-city\主编名称对照.xlsx`

## Workflow

### Step 1: Build crosswalk

```bash
python "scripts/build_crosswalk.py" \
    --eic-xlsx "D:\Opencode\Scholar-city\期刊目录.xlsx" \
    --paper-xlsx "D:\Opencode\Scholar-city\期刊文章output.xlsx" \
    --output "D:\Opencode\Scholar-city\主编名称对照.xlsx"
```

Output columns:
- Col A: 主编规范名 (canonical name from 期刊目录)
- Cols B+: 数据库中匹配到的名称变体 (up to 20 variants)
- Col 3: 匹配数量
- Col 4: 匹配状态 (已匹配/未找到匹配)

### Step 2: Verify matches

Review the output for false positives. Common issues:
- Same surname + first initial, different person (e.g., "Luttmer, EFP" vs "Luttmer, EGJ")
- Chinese names with identical romanized surnames
- Authors who share a surname with an editor but are different people

The matcher uses surname + initial matching with safeguards:
- If both sides have 2+ initials, ALL must match
- Single initials match only when consistent with full name

### Step 3: Use in Stata

```stata
* Import crosswalk
import excel "主编名称对照.xlsx", sheet("Name Crosswalk") firstrow clear
* Reshape to long format for merge
reshape long 数据库中匹配到的名称变体, i(主编规范名) j(variant_num)
rename 数据库中匹配到的名称变体 author_name
drop if missing(author_name)
save "eic_name_xwalk.dta", replace

* Merge with paper-author data
use "authors_long.dta", clear
merge m:1 author_name using "eic_name_xwalk.dta", keepusing(主编规范名)
gen is_editor_paper = (_merge == 3)
```

## Matching Algorithm

1. For each EIC canonical name, generate variants:
   - "Lastname, Firstname" (e.g., "Luttmer, Erzo F. P.")
   - "Lastname, Initials" (e.g., "Luttmer, EFP")
   - "Lastname, Initial" (e.g., "Luttmer, E")
   - Hyphenated surname handling (e.g., "Rossi-Hansberg")

2. Normalize all names (lowercase, strip punctuation, extract surname)

3. Match surname + initals against paper author list

4. Filter out false positives using strict initial matching rules

## Limitations

- Does not handle Chinese name disambiguation
- Middle names may cause false negatives (e.g., "Smith, J" matching "John Smith" but not "John Michael Smith")
- Journal coverage mismatch: editors from journals not yet crawled will show as "未找到匹配"
