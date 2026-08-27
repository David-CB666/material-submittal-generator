<div align="center">

# Material Submittal Generator

### One-click batch generation of material submittal sheets + automatic BQ page merging — for MEP construction projects.

Automate the most tedious part of MEP construction docs: every material delivered to site needs a formatted submittal sheet. This tool batch-generates 50+ sheets from one template, preserving all images and styles via raw ZIP manipulation. Then merge each sheet with its corresponding BQ tender pages. Built from real construction workflows.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?logo=python&logoColor=white)](https://python.org)
[![PyMuPDF](https://img.shields.io/badge/PyMuPDF-1.23+-00A000?logo=python&logoColor=white)](https://pymupdf.readthedocs.io)
[![Stars](https://img.shields.io/github/stars/David-CB666/material-submittal-generator?style=social)](https://github.com/David-CB666/material-submittal-generator/stargazers)
[![Forks](https://img.shields.io/github/forks/David-CB666/material-submittal-generator?style=social)](https://github.com/David-CB666/material-submittal-generator/network/members)
[![Last Commit](https://img.shields.io/github/last-commit/David-CB666/material-submittal-generator)](https://github.com/David-CB666/material-submittal-generator/commits)

[Quick Start](#-quick-start) · [Features](#-features) · [Documentation](#-documentation) · [中文介绍](#-中文介绍)

</div>

---

## 📸 Demo

![Material Submittal Workflow](demo/submittal_workflow_demo.jpg)

*4-step pipeline: Inspect template → Generate 150+ sheets (images preserved) → Export PDFs → Merge BQ pages*

## 🎯 What Problem Does This Solve?

In MEP (Mechanical, Electrical, Plumbing) construction projects, every material delivered to site requires a **material submittal form** — a formatted Excel sheet with product photos, brand info, quantities, and approval fields. Engineers typically spend **hours manually copy-pasting** data into templates, one item at a time.

This tool automates the entire pipeline:

```
Template .xlsx + Source Data → Multi-Sheet Excel → Per-Sheet PDFs → Merge with BQ pages
```

## ✨ Features

| Feature | Description |
|---------|-------------|
| 📦 **Batch generate** | 50+ submittal sheets from one template and one data source |
| 🖼️ **Preserves everything** | Images, print settings, cell styles (centering, fonts, fills) |
| ⚡ **Pure ZIP engine** | No openpyxl `copy_worksheet()` quirks, no broken images |
| 📄 **Auto BQ merging** | Automatically matches and appends Bill of Quantities pages to each PDF |
| 🔍 **Template inspector** | Diagnose any .xlsx template: find placeholders, cell positions, style indexes |

## 🚀 Quick Start

```bash
git clone https://github.com/David-CB666/material-submittal-generator.git
cd material-submittal-generator
pip install -r requirements.txt
```

### 1. Inspect your template

```bash
python inspect.py "submittal_template.xlsx"
```

This shows all `{placeholders}`, their cell locations, and style indexes.

### 2. Create a config file

```json
{
  "template": "submittal_template.xlsx",
  "output": "output.xlsx",
  "fields": {
    "{编号}": "ref_no",
    "{材料}": "material",
    "{品牌}": "brand",
    "{数量}": "quantity"
  },
  "data": [
    {"ref_no": "MAT-001", "material": "Distribution Board", "brand": "Schneider", "quantity": "5"},
    {"ref_no": "MAT-002", "material": "Cable Tray", "brand": "OBO Bettermann", "quantity": "120m"}
  ]
}
```

### 3. Generate

```bash
python gen.py config.json
```

Output: a multi-sheet `.xlsx` with one sheet per data row — all images and styles intact.

### 4. Export to PDF (optional)

```bash
python export_pdf.py output.xlsx ./pdfs/
```

> Requires: Windows + Microsoft Excel

### 5. Merge with BQ pages (optional)

```bash
python merge_bq.py \
  --zongbiao "master_list.xlsx" \
  --bq "BQ_tender.pdf" \
  --input "./pdfs/" \
  --output "./final_output/"
```

Output: `MAT-001 Cable Tray.pdf` (page 1 = submittal form, page 2+ = corresponding BQ pages).

## 📋 Scripts Overview

| Script | Purpose | Dependencies |
|--------|---------|-------------|
| `inspect.py` | Diagnose template structure (placeholders, styles, images) | Standard library |
| `gen.py` | Batch-generate multi-sheet Excel from template + data | Standard library |
| `export_pdf.py` | Export each sheet to a separate PDF | pywin32, Windows + Excel |
| `merge_bq.py` | Merge submittal PDFs with corresponding BQ pages | openpyxl, PyMuPDF |
| `material_tools.py` | **Unified CLI** — all-in-one: `gen`, `inspect`, `merge-bq`, `export-pdf` | All above |

## 📖 Documentation

| Document | Description |
|----------|-------------|
| [Workflow Guide](docs/WORKFLOW.md) | Detailed step-by-step guide and troubleshooting |

## 📊 Real-World Impact

> *"每个工程百几二百份材料要报批。以前一份份人手填：Copy template → Paste data → 加相 → 较打印设定 → 出 PDF → 找对应 BQ 页。一份起码 3~5 分钟。现在 gen.py 一条 command，5 分钟出晒全部。"*
> — Mike, MEP Project Manager

| Metric | Before (Manual) | After (Generator) |
|--------|----------------|-----------------|
| Time per batch (150 items) | 8-12 hours | **5 minutes** |
| Error rate | ~20% (manual copy-paste) | **<1%** |
| Image preservation | Often broken | **100% preserved** |

## 🔧 Why a ZIP Engine Instead of openpyxl?

openpyxl's `copy_worksheet()` and `save()` operations:

- Drop `printerSettings.bin` (print settings)
- Break DrawingML links (images disappear)
- Corrupt `workbook.xml.rels` rId mappings

**Our ZIP engine** reads the `.xlsx` as a raw ZIP, manipulates the XML directly, and writes back — preserving every binary resource and style definition.

### BQ Page Merging

The `merge_bq.py` script:

1. Reads the master tracking Excel to get `(ref_no, BQ_ref, material_name)` mappings
2. Scans the BQ tender PDF with regex to build `{BQ_ref → page_number}` index
3. Merges each submittal PDF with its corresponding BQ page(s)

**Requirements for BQ PDF:**
- Must have a **text layer** (generated from Word/Excel, not scanned images)
- BQ reference numbers must be regex-matchable (e.g., `1.1`, `2.3.1`, `5.10-a`)

> For scanned BQ PDFs, use OCR (pytesseract) or provide a manual page mapping.

## 📦 Installation

```bash
# Core (gen.py + inspect.py): zero dependencies — Python standard library only
python gen.py config.json

# export_pdf.py
pip install pywin32

# merge_bq.py
pip install openpyxl PyMuPDF
```

## 🇨🇳 中文介绍

材料报批表一键批量生成 + BQ 标书页自动合并。纯 ZIP 引擎保留图片及打印设置，支持模板诊断、批量生成、COM 导出 PDF、自动匹配 BQ 页。基于真实 MEP 工程实战。

**核心问题：** 每个工程上百份材料需要报批，工程师手动 Copy-Paste 每份表格需 3~5 分钟。本工具一键生成全部，5 分钟完成 150+ 份。

**技术亮点：**
- 纯标准库实现核心功能（gen.py + inspect.py 零依赖）
- ZIP 引擎直接操作 XML，100% 保留图片和打印设置
- 自动匹配 BQ 页面并合并到对应 PDF

## 📖 Complete Guide

This tool is the **`material-approval-workflow` skill pack** from the **AI Agent Cultivation Field Manual v2.0** — a 13-chapter + 8-appendix handbook that teaches you how to train a generic AI into your domain engineering expert.

**The manual's Excel chapter (Ch.7) documents the complete methodology behind this tool:**

- **5 Iron Rules** for Excel automation (why pure ZIP, why not openpyxl save, why backup first)
- **4 Painful Lessons** — real failures that led to each rule
- **Complete case study**: 28 submittal sheets, 14 hrs → 30 min (**28x faster**)
- **Material Submittal 5-Step Pipeline**: master list → sub-items → PDF with BQ → model column → brand fill

**Full manual covers 4 software domains:**

| Software | Chapter | Efficiency Gain |
|----------|---------|-----------------|
| Word | Ch.6 | 2-3 days → 30 min (10x+) |
| **Excel** ← this tool's chapter | Ch.7 | 14 hrs → 30 min (28x) |
| PDF | Ch.8 | 200 pages → 40 min (25x) |
| CAD | Ch.9 | Auto-generate arrangement drawings (10x+) |

**You get with the full version:**
- 13 chapters + 8 appendices (HTML + PDF + EPUB)
- 7 installable skill packs (this one + 6 more)
- Complete Python case studies with real code
- Battle-tested rules + pitfalls from 128 days of AI cultivation
- CAD chapter: DWG/DXF processing, SLD → arrangement drawing automation

👉 **[Get the full manual + all 7 skill packs →](https://david-cb666.github.io/ai-agent-manual)**

> Free preview: Chapter 1 (5-min AI level self-assessment) available on the landing page.

---

## 🔗 My Other Tools

| Tool | Description |
|------|-------------|
| [**Excel Template Filler**](https://github.com/David-CB666/excel-template-filler) | Dual-engine batch template filling — images & print settings preserved |
| [**GanttChart Pro**](https://github.com/David-CB666/gantt-chart-pro) | Professional Gantt charts in Excel — no MS Project |
| [**VBA Macro Reader**](https://github.com/David-CB666/VBA-Macro-Reader-v2.0.0) | Read, modify & execute VBA macros from .xlsm files |

## 🤝 Contributing

Contributions are welcome! Please read the [Contributing Guide](CONTRIBUTING.md) before submitting a pull request.

## 📄 License

MIT — see [LICENSE](LICENSE).

## 👤 Author

**Mike** — MEP Project Manager.

---

<div align="center">

### ⭐ If this tool saved you time, give it a star!

[![Star History Chart](https://api.star-history.com/svg?repos=David-CB666/material-submittal-generator&type=Date)](https://star-history.com/#David-CB666/material-submittal-generator&Date)

</div>
