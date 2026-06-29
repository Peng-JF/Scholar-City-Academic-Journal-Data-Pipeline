"""
Parse the RUC Core Journal Directory PDF (2017 revised trial).
Extracts international journals with categories ECONOMICS, BUSINESS, BUSINESS,FINANCE, MANAGEMENT
at A+, A, and A- levels.

Usage:
    python parse_pdf.py <pdf_path> [--output <csv_path>]
    
Output: CSV with columns: section, category, journal_title, issn, page
"""
import sys
import re
import pdfplumber
from pathlib import Path
from collections import Counter

TARGET_CATS = {"ECONOMICS", "BUSINESS", "BUSINESS,FINANCE", "MANAGEMENT"}

def extract_journals(pdf_path):
    """Extract journals from the PDF. Returns list of dicts."""
    pdf = pdfplumber.open(pdf_path)
    
    # First pass: extract all text, identify which pages have which section
    pages_with_tables = {}
    current_section = ""
    for pg_num, page in enumerate(pdf.pages):
        text = page.extract_text()
        if not text:
            continue
        if "A+类期刊" in text:
            current_section = "A+"
        elif "A 类期刊" in text and "A+类" not in text and "A-类" not in text:
            current_section = "A"
        elif "A-类期刊" in text:
            current_section = "A-"
        pages_with_tables[pg_num] = current_section
    
    # Second pass: extract from tables where category is available
    results = []
    for pg_num in pages_with_tables:
        if pages_with_tables[pg_num] == "":
            continue
        
        section = pages_with_tables[pg_num]
        page = pdf.pages[pg_num]
        tables = page.extract_tables()
        
        for table in tables:
            if not table or len(table) < 2:
                continue
            
            # Find columns from header
            cat_col = title_col = issn_col = None
            for row in table:
                if not row:
                    continue
                for ci, cell in enumerate(row):
                    if cell is None:
                        continue
                    s = str(cell).upper().replace("\n", " ").strip()
                    if "CATEGORIES" in s or (s in TARGET_CATS):
                        if cat_col is None:
                            cat_col = ci
                    if "FULL JOURNAL TITLE" in s:
                        title_col = ci
                    if s == "ISSN":
                        issn_col = ci
            
            if title_col is None or issn_col is None:
                continue
            
            current_cat = ""
            for row in table:
                if not row or len(row) <= max(title_col, issn_col):
                    continue
                
                # Skip header-like rows
                skip = any(
                    cell and ("SERIAL" in str(cell).upper() or 
                             "CATEGORIES" in str(cell).upper() or
                             "FULL JOURNAL TITLE" in str(cell).upper() or
                             str(cell).strip() == "ISSN" or
                             str(cell).strip() == "NUMBER")
                    for cell in row if cell
                )
                if skip:
                    continue

                # Update category: check all columns for a category value
                for ci in range(len(row)):
                    cell = row[ci]
                    if cell and str(cell).strip() in TARGET_CATS:
                        current_cat = str(cell).strip()
                        break
                    elif cell and str(cell).strip() not in ("", None) and ci == (cat_col if cat_col is not None else 0):
                        # A non-empty value in the category column that's NOT a target cat
                        # means we should stop tracking this category
                        cat_val = str(cell).strip().upper().replace("\n", " ")
                        # Check if it might be a non-target category
                        if cat_val and not any(tc in cat_val for tc in TARGET_CATS):
                            current_cat = ""

                if current_cat not in TARGET_CATS:
                    continue

                # Extract title
                title = ""
                if row[title_col]:
                    title = str(row[title_col]).replace("\n", " ").strip()
                    title = " ".join(title.split())
                
                # Extract ISSN
                issn = ""
                if row[issn_col]:
                    issn = str(row[issn_col]).replace("\n", " ").strip()
                    issn = " ".join(issn.split())

                # Validate: good title + valid ISSN format
                if title and len(title) > 3 and issn and re.match(r'^\d{4}[-]?\d{3}[\dXx]$', issn):
                    results.append({
                        "section": section,
                        "category": current_cat,
                        "title": title,
                        "issn": issn,
                        "page": pg_num + 1
                    })

    # Deduplicate by (title, issn)
    seen = set()
    unique = []
    for r in results:
        key = (r["title"].upper(), r["issn"])
        if key not in seen:
            seen.add(key)
            unique.append(r)
    
    return unique


def main():
    if len(sys.argv) < 2:
        print("Usage: python parse_pdf.py <pdf_path> [--output <csv_path>]")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    output_path = None
    if len(sys.argv) >= 4 and sys.argv[2] == "--output":
        output_path = sys.argv[3]
    
    journals = extract_journals(pdf_path)
    
    if output_path:
        with open(output_path, "w", encoding="utf-8-sig") as f:
            f.write("section,category,journal_title,issn,page\n")
            for j in journals:
                f.write('{},"{}","{}",{},{}\n'.format(
                    j["section"], j["category"], 
                    j["title"].replace('"', '""'), 
                    j["issn"], j["page"]))
        print(f"Written {len(journals)} journals to {output_path}")
    else:
        for j in journals:
            print(f'{j["section"]}\t{j["category"]}\t{j["title"]}\t{j["issn"]}\t{j["page"]}')
        print(f"\nTotal: {len(journals)} journals")
    
    sections = Counter(j["section"] for j in journals)
    print("By section:", dict(sections))
    cats = Counter(j["category"] for j in journals)
    print("By category:", dict(cats))


if __name__ == "__main__":
    main()
