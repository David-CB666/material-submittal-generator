# Material Submittal Generator — Full Workflow Guide

---

## Complete Workflow

```
┌─────────────────────────────────────────────────────────┐
│ Step 1: Prepare two files                                 │
│   A. Submittal template .xlsx (with {placeholders})       │
│   B. Source data .xlsx (material list)                    │
└────────────────────────┬────────────────────────────────┘
                         ↓
┌─────────────────────────────────────────────────────────┐
│ Step 2: Inspect template structure                        │
│   python inspect.py "submittal_template.xlsx"             │
│   → Output: placeholder list, cell positions, style IDs  │
└────────────────────────┬────────────────────────────────┘
                         ↓
┌─────────────────────────────────────────────────────────┐
│ Step 3: Write config.json                                 │
│   → fields: placeholder → field name mapping             │
│   → data: source data as list of objects                 │
└────────────────────────┬────────────────────────────────┘
                         ↓
┌─────────────────────────────────────────────────────────┐
│ Step 4: Generate submittal sheets                         │
│   python gen.py config.json                               │
└────────────────────────┬────────────────────────────────┘
                         ↓
┌─────────────────────────────────────────────────────────┐
│ Step 5: Verify in Excel                                   │
│   → Images visible?                                       │
│   → Data in correct cells?                                │
│   → Text centered? (check style s= attribute)             │
└─────────────────────────────────────────────────────────┘
```

---

## config.json Specification

```json
{
  "name": "Project-Example",
  "template": "./templates/submittal_template.xlsx",
  "output": "./output.xlsx",
  "description": "MEP material submittal master list (EL + ELV), 15 items",
  "fields": {
    "{编号}": "ref_no",
    "{BQ编号}": "bq_ref",
    "{材料}": "material",
    "{品牌}": "brand",
    "{数量}": "quantity",
    "{供货期}": "lead_time",
    "{日期}": "date"
  },
  "data": [
    {
      "ref_no": "MAT-001",
      "bq_ref": "1.1",
      "material": "Weatherproof Distribution Board",
      "brand": "Schneider",
      "quantity": "1 set",
      "lead_time": "30 days",
      "date": "2026-06-26"
    }
  ]
}
```

---

## Placeholder → Field Mapping

From `inspect.py` output:
```
{placeholder} → A1 (s=20)
```

In `config.json` `fields`:
- **key**: `{placeholder}` (exact match from template)
- **value**: column header in your source data spreadsheet

### Common placeholders

| Placeholder | Meaning |
|:---|:---|
| `{编号}` | Reference number (e.g., MAT-001) |
| `{BQ编号}` | BQ (Bill of Quantities) reference |
| `{材料}` | Material name |
| `{品牌}` | Brand / model |
| `{数量}` | Quantity |
| `{供货期}` | Lead time (days) |
| `{日期}` | Submission date |

---

## Source Data Format

Your source data `.xlsx`:
- **Row 1**: Column headers (field names)
- **Data rows**: Starting from row 2
- **Termination**: Empty row signals end of data

---

## Multi-Project Management

One config JSON per project:
```
configs/
├── Project-A.json
├── Project-B.json
└── Project-C.json
```

---

## Style (Centering) Troubleshooting

If text is not centered in the generated output, run `inspect.py`:
```
{编号} → H7 (s=95)
```

`s=95` is the style index defined in `styles.xml`.

**If style is lost** (text not centered):
→ Check generated sheet XML for `s="95"` attribute on `<c r="H7" ...>`
→ `gen.py` v2+ preserves original cell `s` attribute during sharedStrings → inlineStr conversion

---

## Common Errors

| Error | Cause | Fix |
|:---|:---|:---|
| File corrupted / won't open | Template has printerSettings; openpyxl drops them | Use ZIP engine (gen.py) |
| Images don't appear | copy_worksheet breaks DrawingML links | Use ZIP engine (gen.py) |
| Text not centered | sharedStrings→inlineStr dropped s attribute | gen.py v2+ preserves s |
| rId conflict | New sheets overwrite theme/styles rIds | rId5~rId18 reserved for new sheets |
| XML corrupted | append_before_close inserted at wrong position | Uses rfind for last closing tag |
| Placeholder not found | sharedStrings uses inlineStr not t="s" | Inspect template first |

---

## Full Output Pipeline (PDF + BQ Merge)

### Overview

```
Excel master list (N sheets)
       ↓  inspect.py — verify content
       ↓  Excel — manual check
       ↓  export_pdf.py   or   Excel 'Save As'
Submittal PDFs (N files, one per sheet)
       ↓  merge_bq.py
Final PDFs (N files, each = submittal form + BQ pages)
```

### Step A: Export to PDF

**Method 1 (Recommended): Manual**
1. Open generated Excel
2. File → Save As → PDF → "Entire Workbook"

**Method 2 (Auto, requires Excel)**
```bash
python export_pdf.py "output.xlsx" "./pdfs/"
```

> ⚠️ `export_pdf.py` requires Windows + Microsoft Excel installed

### Step B: Merge BQ Pages

```bash
python merge_bq.py \
  --zongbiao "master_tracker.xlsx" \
  --bq      "BQ_tender.pdf" \
  --input   "./pdfs/" \
  --output  "./final_output/"
```

**Parameters:**

| Parameter | Description | Default |
|:---|:---|:---|
| `--zongbiao` | Master tracker Excel (ref_no, BQ_ref, material name columns) | Required |
| `--bq` | BQ tender PDF (must have text layer) | Required |
| `--input` | Directory of submittal PDFs (from export_pdf step) | Required |
| `--output` | Output directory for merged PDFs | Required |
| `--src-sheet` | Master tracker sheet index | 0 |
| `--src-start` | Data start row in master tracker | 7 |

**Output:**
- File naming: `MAT-001 Material Name.pdf`
- Each PDF: page 1 = submittal form, page 2+ = BQ pages
- Generates `merge_summary.txt` with matching results

### BQ PDF Requirements

1. **Text layer required** — `fitz.get_text()` must extract BQ references
2. **BQ refs regex-matchable** — e.g., `1.1`, `2.3.1`, `5.10-a`

**If BQ PDF is pure scanned image:**

| Option | Approach |
|:---|:---|
| A. OCR | Use pytesseract or OCR API to extract text per page |
| B. Visual match | Render page screenshots, manually map against BQ list |
| C. Hard-coded map | `bq_page_map = {'1.1': 1, '2.1': 2, ...}` |

**Master tracker column layout (merge_bq.py reads):**

| Column | Content | Example |
|:---:|:---|:---|
| A | BQ reference | 1.1, 2.3.1 |
| B | Submittal ref | MAT-001, MAT-002 |
| C | Material name | Distribution Board |

---

## Lessons Learned

### Excel → PDF: Three approaches

| Approach | Tool | Pros | Cons |
|:---|:---|:---|:---|
| Manual Save As | Excel UI | 100% fidelity | Slow for many sheets |
| COM export | export_pdf.py | Batch automated | Requires Windows + Excel |
| Python lib | pdfkit/weasyprint | No Excel dependency | Layout differs from original |

**Conclusion: Manual Save As is the most reliable.**

### BQ PDF text layer sources

Different projects have different BQ PDF sources:
- **Word/Excel export** → has text layer ✅ → merge_bq.py works directly
- **Scanned image** → no text layer ❌ → use OCR or manual page mapping
