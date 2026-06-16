# Service Assurance — FRACC Reporting — TIS Accreditation Status & Report

**Owner:** Carl Solomon, Rail Delivery Group  
**Team:** Service Assurance  
**Last updated:** June 2026

---

## Overview

This tool automates the generation of the monthly **FRACC (Finance, Risk, Assurance & Compliance Committee) TIS Accreditation Status paper** from a collated weekly TOC TIS report.

Given two inputs — a collated Excel report and a Word template — it produces:

| Output | Description |
|--------|-------------|
| `FRACC - TIS Accreditation Status - Latest - <timestamp>.xlsx` | Structured Excel with colour-coded accreditation data grouped by Owning Group |
| `FRACC Paper - Accreditation Status Update - <date>.docx` | Fully populated Word paper ready for review and distribution |
| `pie_chart_fracc.png` | Accreditation status summary chart (embedded in the Word document) |

---

## Quick Start

### 1. Install dependencies

```bash
pip install pandas openpyxl python-docx matplotlib lxml Pillow
```

### 2. Run the generator

```bash
python generate_fracc_paper.py \
    --collated  "TOC TIS Accreditation Status - Collated Report - v3.0.xlsx" \
    --template  "FRACC Paper Template.docx" \
    --paper-date "16 June 2026" \
    --output-dir "./output"
```

### 3. Optional — embed a live ASSIST News screenshot

If you have a screenshot of the ASSIST News widget (captured from the portal), pass it via `--news-image`:

```bash
python generate_fracc_paper.py \
    --collated  "..." \
    --template  "..." \
    --paper-date "16 June 2026" \
    --news-image "assist_news.png" \
    --output-dir "./output"
```

---

## Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `--collated` | ✅ | Path to the collated master TOC TIS Excel report |
| `--template` | ✅ | Path to the FRACC Paper Word template (`.docx`) |
| `--paper-date` | ✅ | Paper date in `DD Month YYYY` format, e.g. `"16 June 2026"` |
| `--output-dir` | ❌ | Output folder. Defaults to current directory |
| `--news-image` | ❌ | Path to optional ASSIST News widget screenshot (`.png`) |

---

## What the script does

### Step 1 — Load & normalise the collated report

Reads the collated Excel file. Supports two formats:
- **Pivot export** — file already has an `Owning Group` column (standard output from the TIS Accreditation Status Agent)
- **Raw weekly format** — header at row 4, data from row 5 (direct portal export)

### Step 2 — Peak-count logic

For every unique combination of `Owning Group + TOC/LPC + System + Machine Type ID + Version`, the **highest device count ever recorded** across all historical weekly blocks is selected. This avoids dips caused by missing portal shifts and gives an accurate representation of true fleet size.

### Step 3 — Build the FRACC Excel file

Three-sheet workbook:

| Sheet | Content |
|-------|---------|
| `TIS Accreditation Status` | Grouped by Owning Group → TOC/LPC. Colour-coded State cells (green = Accredited, red = Expired). Appendix header rows in RDG navy. |
| `Pivot Data` | Flat table of all peak-count rows. Used for validation and mail merge. |
| `Accreditation Status Chart` | Summary counts by State for chart generation. |

### Step 4 — Generate the pie chart

Matplotlib pie chart with:
- Colour-coded wedges (green / light green / red)
- Percentage and device count labels outside each wedge
- Legend placed to the right (no overlap)
- Saved at 180 DPI for crisp embedding

### Step 5 — Patch the Word document

The following changes are applied to the template:

| # | What changes | Detail |
|---|--------------|--------|
| 1 | **Meeting Date** | Set to `XX Month` with yellow highlight — reviewer fills in before sending |
| 2 | **Paper Date** | Set to the `--paper-date` value |
| 3 | **Update paragraph** | The bold "this report was run on **[date]**" date is updated to match Paper Date |
| 4 | **Trailer table** (Governance grid) | Date cell updated to `XX Month` with yellow highlight |
| 5 | **Pie chart image** | Replaced with the freshly generated chart |
| 6 | **ASSIST News image** | Replaced with supplied screenshot (if `--news-image` provided) |
| 7 | **Appendix list** | All 16 entries (A–P) converted to internal hyperlinks pointing to their sections |

---

## Template requirements

The Word template must contain:

- A paragraph beginning `Meeting Date:` with a date run
- A paragraph beginning `Paper Date:` with a date run
- A paragraph containing the phrase `this report was run on` with a **bold** date
- A final table (trailer/governance grid) with the meeting date in column 2, row 1
- Two embedded images: `image1.png` (pie chart placeholder), `image2.png` (news screenshot placeholder)
- Appendix section headings with Word bookmarks named `Appendix_A` through `Appendix_L`, `AppendixM` through `AppendixP`
- An appendix contents list (paragraphs beginning `Appendix A` through `Appendix P`) before the `END OF PAPER` marker
- A `Hyperlink` character style defined (standard in all Word documents)

---

## Appendix structure

| Appendix | Owning Group |
|----------|--------------|
| A | Go Ahead |
| B | Transport UK Group (formerly Abellio Group) |
| C | Serco / Transport UK Group (formerly Abellio Group) |
| D | First Group |
| E | Directly Operated Railway |
| F | Arriva |
| G | Transport for Wales |
| H | London Overground |
| I | GTS Elizabeth Line |
| K | Heathrow Express Operating Company |
| L | Scottish Rail Holdings |
| M | Accreditation Status (summary + pie chart) |
| N | RDG Compliance Standards — Recent Updates |
| O | Third Party Retailer Systems |
| P | Governance Lifecycle Tracking Grid |

---

## Accreditation state colour coding

| State | Fill | Font |
|-------|------|------|
| Accredited | Light green `#C6EFCE` | Dark green `#276221` |
| Pilot phase | Very light green `#E2EFDA` | Forest green `#375623` |
| Accreditation expired | Light red `#FFC7CE` | Dark red `#9C0006` |
| Application acknowledged | Light amber `#FFEB9C` | Dark amber `#9C6500` |

---

## File naming conventions

| File | Convention |
|------|------------|
| FRACC Excel | `FRACC - TIS Accreditation Status - Latest - YYYYMMDD_HHMM.xlsx` |
| FRACC Word paper | `FRACC Paper - Accreditation Status Update - DD Month YYYY.docx` |
| Pie chart | `pie_chart_fracc.png` (intermediate, kept in output dir) |

---

## Related tools

| Tool | Location | Purpose |
|------|----------|---------|
| Portal downloader | `.agents/skills/download_toc_tis_report/scripts/run.py` | Downloads the weekly TOC TIS Report (2) from rspaccreditation.org |
| Collator | `.agents/skills/download_toc_tis_report/scripts/collate.py` | Merges all weekly files into the master collated report |
| **This script** | `generate_fracc_paper.py` | Generates the FRACC paper from the collated report |

---

## Support

Maintained by the TIS Accreditation Status Agent (Base44 Superagent).  
For issues, contact Carl Solomon at Rail Delivery Group.
