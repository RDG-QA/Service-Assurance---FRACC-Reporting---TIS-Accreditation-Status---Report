#!/usr/bin/env python3
"""
====================================================================
FRACC TIS Accreditation Status Report Generator
====================================================================
Rail Delivery Group — Service Assurance
Generates the monthly FRACC paper and supporting FRACC Excel data
file from a collated TOC TIS weekly report.

Usage:
    python generate_fracc_paper.py \\
        --collated  "TOC TIS Accreditation Status - Collated Report - v3.0.xlsx" \\
        --template  "FRACC Paper Template.docx" \\
        --paper-date "16 June 2026" \\
        [--output-dir "./output"]

Parameters:
    --collated      Path to the collated master Excel report.
    --template      Path to the FRACC Word template (.docx).
    --paper-date    Date string for the Paper Date field, e.g. "16 June 2026".
    --output-dir    Optional output folder (default: current directory).

Outputs (written to --output-dir):
    FRACC - TIS Accreditation Status - Latest - <YYYYMMDD_HHMM>.xlsx
    FRACC Paper - Accreditation Status Update - <DD Month YYYY>.docx

Dependencies:
    pip install pandas openpyxl python-docx matplotlib lxml Pillow
====================================================================
"""

import argparse
import copy
import io
import os
import re
import sys
import zipfile
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import pandas as pd
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side
)
from openpyxl.utils import get_column_letter


# ════════════════════════════════════════════════════════════════════
# CONSTANTS
# ════════════════════════════════════════════════════════════════════

# Appendix → Owning Group mapping (display order in the paper)
APPENDIX_MAP = [
    ("A",  "Go Ahead"),
    ("B",  "Transport UK Group (formerly Abellio Group)"),
    ("C",  "Serco / Transport UK Group (formerly Abellio Group)"),
    ("D",  "First Group"),
    ("E",  "Directly Operated Railway"),
    ("F",  "Arriva"),
    ("G",  "Transport for Wales"),
    ("H",  "London Overground"),
    ("I",  "GTS Elizabeth Line"),
    ("K",  "Heathrow Express Operating Company"),
    ("L",  "Scottish Rail Holdings"),
]

# Owning group name aliases (raw data value → canonical display name)
# Handles variations between source portal data and display names
OG_ALIASES = {
    "serco/transport uk group (formerly abellio group)": "Serco / Transport UK Group (formerly Abellio Group)",
    "transport uk group (formerly abellio group)":       "Transport UK Group (formerly Abellio Group)",
    "gts elizabeth line":                                "GTS Elizabeth Line",
    "arriva":                                            "Arriva",
    "go ahead":                                          "Go Ahead",
    "first group":                                       "First Group",
    "directly operated railway":                         "Directly Operated Railway",
    "transport for wales":                               "Transport for Wales",
    "london overground":                                 "London Overground",
    "heathrow express operating company":                "Heathrow Express Operating Company",
    "scottish rail holdings":                            "Scottish Rail Holdings",
}

# TOC-level owning group overrides
# These take precedence over whatever Owning Group the source data carries.
# Format: { "TOC / LPC value (exact)": "Canonical Owning Group name" }
TOC_OG_OVERRIDES = {
    "GTR - SOUTHERN & GATWICK EXPRESS":  "Directly Operated Railway",
    "GTR-THAMESLINK & GREAT NORTHERN":   "Directly Operated Railway",
    "WEST MIDLANDS TRAINS LTD":          "Directly Operated Railway",
}

# Bookmark names for each appendix (must match bookmarks in template)
BOOKMARK_MAP = {
    "A": "Appendix_A",
    "B": "Appendix_B",
    "C": "Appendix_C",
    "D": "Appendix_D",
    "E": "Appendix_E",
    "F": "Appendix_F",
    "G": "Appendix_G",
    "H": "Appendix_H",
    "I": "Appendix_I",
    "K": "Appendix_K",
    "L": "Appendix_L",
    "M": "AppendixM",
    "N": "AppendixN",
    "O": "AppendixO",
    "P": "AppendixP",
}

# Colour palette
COLOURS = {
    "Accredited":             "#00B050",
    "Pilot phase":            "#92D050",
    "Accreditation expired":  "#EE0000",
    "Application acknowledged": "#FFC000",
}

# Header fill colour (dark navy, matching RDG template)
HDR_FILL_HEX  = "1F3864"
HDR_FONT_HEX  = "FFFFFF"
SUBHDR_FILL   = "D6E4F0"
SUBHDR_FONT   = "1F3864"

STATE_FILL = {
    "Accredited":             "C6EFCE",
    "Pilot phase":            "E2EFDA",
    "Accreditation expired":  "FFC7CE",
    "Application acknowledged": "FFEB9C",
}
STATE_FONT = {
    "Accredited":             "276221",
    "Pilot phase":            "375623",
    "Accreditation expired":  "9C0006",
    "Application acknowledged": "9C6500",
}


# ════════════════════════════════════════════════════════════════════
# STEP 1 — Load & normalise collated report
# ════════════════════════════════════════════════════════════════════

def load_collated_report(path: str) -> pd.DataFrame:
    """
    Load the collated TOC TIS master report.

    The file may be either:
    - A raw weekly-format file (header at row 4, data from row 5)
    - A pre-collated pivot export with a direct header row

    Returns a normalised DataFrame with columns:
        Owning Group, TOC / LPC, System, Machine Type ID,
        Version, State, Count, CoA Expiry
    """
    print(f"\n[1] Loading collated report: {os.path.basename(path)}")

    # Try reading as a pivot export first (scan all sheets for 'Owning Group' column)
    try:
        xl = pd.ExcelFile(path, engine="openpyxl")
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            if "Owning Group" in df.columns and "Count" in df.columns:
                df = _normalise_pivot(df)
                print(f"    Loaded as pivot export from sheet '{sheet}' — {len(df)} rows")
                return df
    except Exception:
        pass

    # Fall back to raw weekly format
    raw = pd.read_excel(path, header=None, engine="openpyxl")
    date_range_str = str(raw.iloc[2, 0]) if pd.notna(raw.iloc[2, 0]) else ""
    data = raw.iloc[5:].copy()
    data.columns = [str(c).replace("\n", " ").strip() for c in raw.iloc[4].tolist()]
    data = data.dropna(how="all")
    m = re.search(r"to\s+(\d{2}-\w+-\d{4})", date_range_str)
    report_date = pd.to_datetime(m.group(1), format="%d-%b-%Y") if m else pd.NaT
    data["_report_date"] = report_date
    data["Count"] = pd.to_numeric(data["Count"], errors="coerce").fillna(0).astype(int)
    print(f"    Loaded as raw weekly format — {len(data)} rows (report date: {report_date})")
    return _normalise_raw(data)


def _normalise_pivot(df: pd.DataFrame) -> pd.DataFrame:
    """Normalise a pivot-style export."""
    df = df.rename(columns={
        "Owning Group":  "Owning Group",
        "TOC / LPC":     "TOC / LPC",
        "System":        "System",
        "Machine Type ID": "Machine Type ID",
        "Version":       "Version",
        "State":         "State",
        "Count":         "Count",
        "CoA Expiry":    "CoA Expiry",
    })
    df["Count"] = pd.to_numeric(df["Count"], errors="coerce").fillna(0).astype(int)
    df["Owning Group"] = df["Owning Group"].apply(_canonicalise_og)
    # Apply per-TOC owning group overrides
    df["Owning Group"] = df.apply(
        lambda r: TOC_OG_OVERRIDES.get(str(r["TOC / LPC"]).strip(), r["Owning Group"]),
        axis=1
    )
    return df[[c for c in ["Owning Group","TOC / LPC","System","Machine Type ID",
                            "Version","State","Count","CoA Expiry"] if c in df.columns]]


def _normalise_raw(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalise a raw weekly-format DataFrame.
    The raw format has 'Lennon profit centre' and 'Company' but no
    Owning Group — we derive it from the Company column via the alias map.
    """
    df = df.rename(columns={
        "Lennon profit centre": "TOC / LPC",
        "Company":              "System",       # System vendor
        "Machine Type ID":      "Machine Type ID",
        "Version":              "Version",
        "State":                "State",
        "Count":                "Count",
        "CoA expiry date":      "CoA Expiry",
        "Latest shift":         "_latest_shift",
    })
    # Raw format does not carry Owning Group — set as Unknown for peak logic
    df["Owning Group"] = "Unknown"
    # Apply per-TOC owning group overrides
    df["Owning Group"] = df.apply(
        lambda r: TOC_OG_OVERRIDES.get(str(r.get("TOC / LPC", "")).strip(), r["Owning Group"]),
        axis=1
    )
    return df[["Owning Group","TOC / LPC","System","Machine Type ID",
               "Version","State","Count","CoA Expiry"]]


def _canonicalise_og(name) -> str:
    if pd.isna(name):
        return str(name)
    key = str(name).strip().lower()
    return OG_ALIASES.get(key, str(name).strip())


# ════════════════════════════════════════════════════════════════════
# STEP 2 — Apply peak-count logic
# ════════════════════════════════════════════════════════════════════

def apply_peak_counts(df: pd.DataFrame) -> pd.DataFrame:
    """
    For every combination of (Owning Group, TOC/LPC, System, Machine Type ID, Version)
    take the highest Count ever recorded across all historical weekly blocks.
    The State and CoA Expiry are taken from the row with the peak count.
    """
    print("\n[2] Applying peak-count logic…")
    key_cols = ["Owning Group", "TOC / LPC", "System", "Machine Type ID", "Version"]

    # Sort so the highest count is first, then deduplicate on key columns
    df_sorted = df.sort_values("Count", ascending=False)
    peak = df_sorted.drop_duplicates(subset=key_cols, keep="first").copy()
    peak = peak.sort_values(["Owning Group", "TOC / LPC", "System", "Machine Type ID", "Version"])
    print(f"    {len(df)} rows → {len(peak)} unique combinations after peak-count")
    return peak.reset_index(drop=True)


# ════════════════════════════════════════════════════════════════════
# STEP 3 — Build FRACC Excel output
# ════════════════════════════════════════════════════════════════════

def _thin_border():
    side = Side(style="thin", color="BFBFBF")
    return Border(left=side, right=side, top=side, bottom=side)


def _header_font(bold=True, colour=HDR_FONT_HEX, size=11):
    return Font(bold=bold, color=colour, name="Calibri", size=size)


def _header_fill(hex_colour=HDR_FILL_HEX):
    return PatternFill("solid", fgColor=hex_colour)


def _centre():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)


def _left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)


def build_fracc_excel(peak_df: pd.DataFrame, paper_date_str: str,
                      output_dir: str, timestamp: str) -> str:
    """
    Build the FRACC TIS Accreditation Status Excel file.

    Sheet 1 — TIS Accreditation Status
        Grouped by Owning Group > TOC/LPC with coloured State cells.

    Sheet 2 — Pivot Data
        Flat table of all records (used for validation / mail merge).

    Sheet 3 — Accreditation Status Chart
        Summary counts by State + a pie chart.
    """
    print("\n[3] Building FRACC Excel file…")

    wb = Workbook()

    # ── Sheet 1: TIS Accreditation Status ──────────────────────────
    ws1 = wb.active
    ws1.title = "TIS Accreditation Status"

    col_headers = ["System", "Machine Type ID", "Version", "State", "Count", "CoA Expiry"]
    col_widths   = [38, 18, 16, 26, 10, 16]

    row_num = 1
    owning_groups = [og for og, _ in APPENDIX_MAP]

    for app_letter, og_display in APPENDIX_MAP:
        og_data = peak_df[peak_df["Owning Group"] == og_display].copy()
        if og_data.empty:
            # Try case-insensitive match
            og_data = peak_df[
                peak_df["Owning Group"].str.lower() == og_display.lower()
            ].copy()
        if og_data.empty:
            continue

        # ── Owning Group header row ──────────────────────────────
        ws1.merge_cells(start_row=row_num, start_column=1,
                        end_row=row_num, end_column=6)
        cell = ws1.cell(row=row_num, column=1,
                        value=f"Appendix {app_letter} — {og_display}")
        cell.font   = _header_font(size=12)
        cell.fill   = _header_fill()
        cell.alignment = _centre()
        cell.border = _thin_border()
        row_num += 1

        # ── TOC sub-groups ────────────────────────────────────────
        for toc in sorted(og_data["TOC / LPC"].dropna().unique()):
            toc_data = og_data[og_data["TOC / LPC"] == toc]

            # TOC sub-header
            ws1.merge_cells(start_row=row_num, start_column=1,
                            end_row=row_num, end_column=6)
            toc_cell = ws1.cell(row=row_num, column=1, value=toc)
            toc_cell.font      = Font(bold=True, color=SUBHDR_FONT, name="Calibri", size=10)
            toc_cell.fill      = PatternFill("solid", fgColor=SUBHDR_FILL)
            toc_cell.alignment = _left()
            toc_cell.border    = _thin_border()
            row_num += 1

            # Column headers
            for ci, hdr in enumerate(col_headers, start=1):
                c = ws1.cell(row=row_num, column=ci, value=hdr)
                c.font      = Font(bold=True, color="1F3864", name="Calibri", size=9)
                c.fill      = PatternFill("solid", fgColor="EBF3FB")
                c.alignment = _centre()
                c.border    = _thin_border()
            row_num += 1

            # Data rows
            for _, drow in toc_data.sort_values(
                    ["System", "Machine Type ID", "Version"]).iterrows():
                state = str(drow.get("State", ""))
                fill_hex = STATE_FILL.get(state, "FFFFFF")
                font_hex = STATE_FONT.get(state, "000000")

                vals = [
                    drow.get("System", ""),
                    drow.get("Machine Type ID", ""),
                    drow.get("Version", ""),
                    state,
                    drow.get("Count", 0),
                    drow.get("CoA Expiry", ""),
                ]
                for ci, val in enumerate(vals, start=1):
                    c = ws1.cell(row=row_num, column=ci, value=val)
                    c.border    = _thin_border()
                    c.alignment = _centre() if ci != 1 else _left()
                    c.font      = Font(name="Calibri", size=9)
                    if ci == 4:  # State column — colour coded
                        c.fill = PatternFill("solid", fgColor=fill_hex)
                        c.font = Font(bold=True, color=font_hex,
                                      name="Calibri", size=9)
                row_num += 1

        row_num += 1  # blank gap between owning groups

    # Set column widths
    for ci, w in enumerate(col_widths, start=1):
        ws1.column_dimensions[get_column_letter(ci)].width = w

    # Freeze header-ish row
    ws1.freeze_panes = "A1"

    # ── Sheet 2: Pivot Data ─────────────────────────────────────────
    ws2 = wb.create_sheet("Pivot Data")
    pivot_cols = ["Owning Group", "TOC / LPC", "System",
                  "Machine Type ID", "Version", "State", "Count", "CoA Expiry"]
    pivot_widths = [30, 38, 38, 18, 16, 26, 10, 16]

    for ci, hdr in enumerate(pivot_cols, start=1):
        c = ws2.cell(row=1, column=ci, value=hdr)
        c.font      = _header_font()
        c.fill      = _header_fill()
        c.alignment = _centre()
        c.border    = _thin_border()

    export_df = peak_df[[c for c in pivot_cols if c in peak_df.columns]]
    for ri, (_, row) in enumerate(export_df.iterrows(), start=2):
        for ci, col in enumerate(pivot_cols, start=1):
            val = row.get(col, "")
            c = ws2.cell(row=ri, column=ci, value=val)
            c.border    = _thin_border()
            c.alignment = _left() if ci <= 3 else _centre()
            c.font      = Font(name="Calibri", size=9)
            if col == "State":
                state = str(val)
                c.fill = PatternFill("solid", fgColor=STATE_FILL.get(state, "FFFFFF"))
                c.font = Font(bold=True, color=STATE_FONT.get(state, "000000"),
                              name="Calibri", size=9)

    for ci, w in enumerate(pivot_widths, start=1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
    ws2.freeze_panes = "A2"

    # ── Sheet 3: Accreditation Status Chart ────────────────────────
    ws3 = wb.create_sheet("Accreditation Status Chart")
    state_order = ["Accredited", "Pilot phase", "Accreditation expired",
                   "Application acknowledged"]
    counts = {}
    for state in state_order:
        total = int(peak_df[peak_df["State"] == state]["Count"].sum())
        if total > 0:
            counts[state] = total
    grand_total = sum(counts.values())

    ws3.cell(row=1, column=1, value="Accreditation Status").font = _header_font()
    ws3.cell(row=1, column=1).fill = _header_fill()
    ws3.cell(row=1, column=1).alignment = _centre()
    ws3.cell(row=1, column=2, value="Count").font = _header_font()
    ws3.cell(row=1, column=2).fill = _header_fill()
    ws3.cell(row=1, column=2).alignment = _centre()

    for ri, (state, cnt) in enumerate(counts.items(), start=2):
        ws3.cell(row=ri, column=1, value=state).border = _thin_border()
        ws3.cell(row=ri, column=2, value=cnt).border = _thin_border()

    last_data_row = 1 + len(counts)
    ws3.cell(row=last_data_row + 1, column=1, value="Total Devices").font = Font(bold=True, name="Calibri")
    ws3.cell(row=last_data_row + 1, column=2, value=grand_total).font = Font(bold=True, name="Calibri")

    ws3.column_dimensions["A"].width = 30
    ws3.column_dimensions["B"].width = 12

    # ── Save ──────────────────────────────────────────────────────
    fname = f"FRACC - TIS Accreditation Status - Latest - {timestamp}.xlsx"
    fpath = os.path.join(output_dir, fname)
    wb.save(fpath)
    print(f"    Saved FRACC Excel → {fname}")
    return fpath, counts, grand_total


# ════════════════════════════════════════════════════════════════════
# STEP 4 — Generate pie chart PNG
# ════════════════════════════════════════════════════════════════════

def generate_pie_chart(counts: dict, total: int, output_dir: str) -> str:
    """
    Render an accreditation status pie chart.
    Returns path to the saved PNG.
    """
    print("\n[4] Generating pie chart…")

    labels = list(counts.keys())
    values = list(counts.values())
    colors = [
        {"Accredited": "#00B050",
         "Pilot phase": "#92D050",
         "Accreditation expired": "#EE0000",
         "Application acknowledged": "#FFC000"}.get(l, "#AAAAAA")
        for l in labels
    ]

    fig, ax = plt.subplots(figsize=(9, 5.5))
    fig.patch.set_facecolor("white")

    wedges, _ = ax.pie(
        values,
        colors=colors,
        startangle=140,
        wedgeprops=dict(linewidth=1.2, edgecolor="white"),
        radius=0.88,
    )

    # Percentage + count labels outside wedges
    for wedge, val in zip(wedges, values):
        angle = (wedge.theta2 + wedge.theta1) / 2
        pct   = val / total * 100
        x  = 0.62 * np.cos(np.radians(angle))
        y  = 0.62 * np.sin(np.radians(angle))
        lx = 1.18 * np.cos(np.radians(angle))
        ly = 1.18 * np.sin(np.radians(angle))
        ax.annotate(
            f"{pct:.1f}%\n({val:,})",
            xy=(x, y), xytext=(lx, ly),
            ha="center", va="center",
            fontsize=9.5, fontweight="bold", color="#333333",
            arrowprops=dict(arrowstyle="-", color="#888888", lw=0.8),
            bbox=dict(boxstyle="round,pad=0.2", facecolor="white",
                      edgecolor="none", alpha=0.85),
        )

    # Legend — right side, no overlap
    legend_patches = [
        mpatches.Patch(facecolor=c, edgecolor="#cccccc", linewidth=0.8,
                       label=f"{l}  ({v:,})")
        for l, v, c in zip(labels, values, colors)
    ]
    legend = ax.legend(
        handles=legend_patches,
        loc="center left",
        bbox_to_anchor=(1.02, 0.5),
        frameon=True, framealpha=1.0, edgecolor="#cccccc",
        fontsize=10,
        title="Accreditation Status", title_fontsize=10.5,
        labelspacing=1.0, handlelength=1.5, handleheight=1.2,
    )
    legend.get_frame().set_linewidth(0.8)

    ax.set_title(
        f"TIS Accreditation Status Summary\n(Total: {total:,} devices)",
        fontsize=13, fontweight="bold", color="#1F3864", pad=12,
    )

    plt.tight_layout(rect=[0, 0, 0.78, 1])
    pie_path = os.path.join(output_dir, "pie_chart_fracc.png")
    plt.savefig(pie_path, dpi=180, bbox_inches="tight",
                facecolor="white", edgecolor="none")
    plt.close()
    print(f"    Saved pie chart → {os.path.basename(pie_path)}")
    return pie_path


# ════════════════════════════════════════════════════════════════════
# STEP 5 — Patch the Word document
# ════════════════════════════════════════════════════════════════════

def _set_highlight(run, colour="yellow"):
    rpr = run._r.get_or_add_rPr()
    hl  = rpr.find(qn("w:highlight"))
    if hl is None:
        hl = OxmlElement("w:highlight")
        rpr.append(hl)
    hl.set(qn("w:val"), colour)


def _add_bookmark(para, name, bm_id="99"):
    """Add a w:bookmarkStart / w:bookmarkEnd pair to a paragraph."""
    existing = [b.get(qn("w:name"))
                for b in para._p.findall(".//" + qn("w:bookmarkStart"))]
    if name in existing:
        return
    bm_start = OxmlElement("w:bookmarkStart")
    bm_start.set(qn("w:id"), str(bm_id))
    bm_start.set(qn("w:name"), name)
    bm_end = OxmlElement("w:bookmarkEnd")
    bm_end.set(qn("w:id"), str(bm_id))
    para._p.insert(0, bm_start)
    para._p.append(bm_end)


def _make_hyperlink(para, text, anchor):
    """
    Replace the content of *para* with a single internal hyperlink
    pointing to *anchor*, preserving any existing run formatting.
    """
    # Capture existing rPr from first run
    existing_runs = para._p.findall(qn("w:r"))
    if existing_runs:
        existing_rpr = existing_runs[0].find(qn("w:rPr"))
        rpr = copy.deepcopy(existing_rpr) if existing_rpr is not None \
              else OxmlElement("w:rPr")
    else:
        rpr = OxmlElement("w:rPr")

    # Apply Hyperlink character style
    rStyle = OxmlElement("w:rStyle")
    rStyle.set(qn("w:val"), "Hyperlink")
    rpr.insert(0, rStyle)

    # Build hyperlink element
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor)
    hyperlink.set(qn("w:history"), "1")

    r = OxmlElement("w:r")
    r.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    hyperlink.append(r)

    # Remove all existing runs/hyperlinks from the paragraph
    for child in list(para._p):
        if child.tag in (qn("w:r"), qn("w:hyperlink")):
            para._p.remove(child)

    para._p.append(hyperlink)


def patch_word_document(template_path: str, paper_date_str: str,
                        pie_chart_path: str, news_image_path: str,
                        output_dir: str) -> str:
    """
    Apply all patches to the Word template and save the final paper.

    Patches applied:
      1. Meeting Date → 'XX Month' with yellow highlight
      2. Paper Date   → paper_date_str
      3. Update para  → bold date updated to paper_date_str
      4. Trailer table date cell → 'XX Month' with yellow highlight
      5. Pie chart image (image1) replaced
      6. ASSIST News image (image2) replaced  [optional — skipped if no path]
      7. Appendix list entries → internal hyperlinks
    """
    print(f"\n[5] Patching Word document…")
    doc = Document(template_path)

    # ── Parse paper date ────────────────────────────────────────
    try:
        paper_dt = datetime.strptime(paper_date_str.strip(), "%d %B %Y")
    except ValueError:
        try:
            paper_dt = datetime.strptime(paper_date_str.strip(), "%d/%m/%Y")
        except ValueError:
            paper_dt = datetime.now()
            print(f"    WARNING: Could not parse paper date '{paper_date_str}', "
                  f"using today ({paper_dt.strftime('%d %B %Y')})")
    paper_date_display = paper_dt.strftime("%-d %B %Y")  # e.g. "16 June 2026"
    paper_date_bold_part1 = paper_dt.strftime("%-d %B")  # e.g. "16 June"
    paper_date_bold_part2 = f" {paper_dt.year}"           # e.g. " 2026"

    # ── Find key paragraphs ─────────────────────────────────────
    paras = doc.paragraphs

    # Para indices (based on template structure — robust search fallback)
    meeting_date_para = None
    paper_date_para   = None
    update_para       = None
    toc_list_start    = None
    toc_list_end      = None

    for i, p in enumerate(paras):
        txt = p.text.strip()
        if txt.startswith("Meeting Date:") and meeting_date_para is None:
            meeting_date_para = i
        if txt.startswith("Paper Date:") and paper_date_para is None:
            paper_date_para = i
        if ("this report was run on" in txt or "report was run on" in txt) \
                and update_para is None:
            update_para = i
        if txt.startswith("Appendix A") and toc_list_start is None:
            toc_list_start = i
        if toc_list_start and txt.startswith("Appendix P"):
            toc_list_end = i

    print(f"    meeting_date_para={meeting_date_para}, "
          f"paper_date_para={paper_date_para}, "
          f"update_para={update_para}, "
          f"toc_list=[{toc_list_start}–{toc_list_end}]")

    # ── Patch 1: Meeting Date → 'XX Month' (yellow highlight) ──
    if meeting_date_para is not None:
        p = paras[meeting_date_para]
        for run in p.runs:
            if any(m in run.text for m in [
                    "January","February","March","April","May","June",
                    "July","August","September","October","November","December",
                    "XX Month", "Month"]):
                run.text = "XX Month"
                _set_highlight(run, "yellow")
                break
        print("    ✓ Meeting Date → 'XX Month' (highlighted)")

    # ── Patch 2: Paper Date ─────────────────────────────────────
    if paper_date_para is not None:
        p = paras[paper_date_para]
        # Replace any date-like run with the new paper date
        for run in p.runs:
            if any(m in run.text for m in [
                    "January","February","March","April","May","June","July",
                    "August","September","October","November","December",
                    "2025","2026","2027"]):
                run.text = paper_date_display
                break
        print(f"    ✓ Paper Date → '{paper_date_display}'")

    # ── Patch 3: Update paragraph date (bold) ──────────────────
    if update_para is not None:
        p = paras[update_para]
        runs = p.runs
        for ri, run in enumerate(runs):
            if run.bold and any(m in run.text for m in [
                    "January","February","March","April","May","June","July",
                    "August","September","October","November","December"]):
                run.text = paper_date_bold_part1
                if ri + 1 < len(runs) and runs[ri + 1].bold:
                    runs[ri + 1].text = paper_date_bold_part2
                break
        print(f"    ✓ Update para date → '{paper_date_bold_part1}{paper_date_bold_part2}'")

    # ── Patch 4: Trailer table date cell ───────────────────────
    if doc.tables:
        trailer_tbl  = doc.tables[-1]
        date_cell    = trailer_tbl.rows[0].cells[1]
        date_cell_p  = date_cell.paragraphs[0]
        # Clear existing runs
        for run in date_cell_p.runs:
            run.clear()
        if date_cell_p.runs:
            r = date_cell_p.runs[0]
            r.text = "XX Month"
            r.bold = True
            _set_highlight(r, "yellow")
            for extra in date_cell_p.runs[1:]:
                extra._r.getparent().remove(extra._r)
        print("    ✓ Trailer table date → 'XX Month' (highlighted)")

    # ── Patch 5 & 6: Replace images in the zip ─────────────────
    doc.save(template_path + ".tmp.docx")

    with zipfile.ZipFile(template_path + ".tmp.docx", "r") as zin:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/media/image1.png" \
                        and pie_chart_path and os.path.exists(pie_chart_path):
                    with open(pie_chart_path, "rb") as f:
                        zout.writestr(item, f.read())
                elif item.filename == "word/media/image2.png" \
                        and news_image_path and os.path.exists(news_image_path):
                    with open(news_image_path, "rb") as f:
                        zout.writestr(item, f.read())
                else:
                    zout.writestr(item, zin.read(item.filename))
        buf.seek(0)

    os.remove(template_path + ".tmp.docx")

    # Reload the patched zip as a Document for hyperlink step
    doc2 = Document(buf)
    paras2 = doc2.paragraphs

    # ── Patch 7: Appendix list → hyperlinks ─────────────────────
    # Ensure Appendix I bookmark exists in the body
    for i, p in enumerate(paras2):
        if "Appendix I" in p.text and "GTS Elizabeth Line" in p.text and \
                "Appendix_I" not in [b.get(qn("w:name"))
                                     for b in p._p.findall(".//" + qn("w:bookmarkStart"))]:
            _add_bookmark(p, "Appendix_I", bm_id="199")

    # Build full appendix list — A-L plus M-P
    all_appendix_letters = [a for a, _ in APPENDIX_MAP] + ["M", "N", "O", "P"]
    all_appendix_names = {a: n for a, n in APPENDIX_MAP}
    all_appendix_names.update({
        "M": "Accreditation Status",
        "N": "RDG Compliance Standards - Recent Updates",
        "O": "Third Party Retailer Systems",
        "P": "Governance Lifecycle Tracking Grid",
    })

    # Find the TOC list paragraphs and convert to hyperlinks
    toc_start = None
    converted = 0
    for i, p in enumerate(paras2):
        txt = p.text.strip()
        # Detect start of the appendix list (first "Appendix A" before END OF PAPER)
        if txt.startswith("Appendix A") and toc_start is None:
            # Make sure we're in the TOC section (before body appendices)
            toc_start = i
        if toc_start is not None and i >= toc_start:
            for letter in all_appendix_letters:
                prefix = f"Appendix {letter}"
                if txt.startswith(prefix) and BOOKMARK_MAP.get(letter):
                    _make_hyperlink(p, txt, BOOKMARK_MAP[letter])
                    converted += 1
                    break
        # Stop after Appendix P in the TOC list
        if toc_start and txt.startswith("Appendix P") and i > toc_start + 2:
            break

    print(f"    ✓ Converted {converted} appendix entries to internal hyperlinks")

    # ── Save final document ─────────────────────────────────────
    out_name = f"FRACC Paper - Accreditation Status Update - {paper_date_display}.docx"
    out_path = os.path.join(output_dir, out_name)

    final_buf = io.BytesIO()
    doc2.save(final_buf)
    final_buf.seek(0)
    with open(out_path, "wb") as f:
        f.write(final_buf.read())

    print(f"    Saved Word document → {out_name}")
    return out_path


# ════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Generate the monthly FRACC TIS Accreditation Status paper."
    )
    parser.add_argument(
        "--collated", required=True,
        help="Path to the collated TOC TIS master Excel report."
    )
    parser.add_argument(
        "--template", required=True,
        help="Path to the FRACC Paper Word template (.docx)."
    )
    parser.add_argument(
        "--paper-date", required=True,
        help='Paper date string, e.g. "16 June 2026" or "16/06/2026".'
    )
    parser.add_argument(
        "--output-dir", default=".",
        help="Output directory for generated files (default: current directory)."
    )
    parser.add_argument(
        "--news-image", default=None,
        help="(Optional) Path to a screenshot of the ASSIST News widget to embed."
    )

    args = parser.parse_args()

    # Validate inputs
    if not os.path.exists(args.collated):
        print(f"ERROR: Collated report not found: {args.collated}")
        sys.exit(1)
    if not os.path.exists(args.template):
        print(f"ERROR: Word template not found: {args.template}")
        sys.exit(1)

    os.makedirs(args.output_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")

    print("=" * 60)
    print("  FRACC TIS Accreditation Status Report Generator")
    print("=" * 60)
    print(f"  Collated report : {args.collated}")
    print(f"  Template        : {args.template}")
    print(f"  Paper date      : {args.paper_date}")
    print(f"  Output dir      : {args.output_dir}")
    print("=" * 60)

    # Step 1 — Load
    df = load_collated_report(args.collated)

    # Step 2 — Peak counts
    peak_df = apply_peak_counts(df)

    # Step 3 — Excel output
    excel_path, state_counts, grand_total = build_fracc_excel(
        peak_df, args.paper_date, args.output_dir, timestamp
    )

    # Step 4 — Pie chart
    pie_path = generate_pie_chart(state_counts, grand_total, args.output_dir)

    # Step 5 — Word document
    docx_path = patch_word_document(
        template_path  = args.template,
        paper_date_str = args.paper_date,
        pie_chart_path = pie_path,
        news_image_path= args.news_image,
        output_dir     = args.output_dir,
    )

    print("\n" + "=" * 60)
    print("  ✓  Generation complete")
    print("=" * 60)
    print(f"  Excel  → {os.path.basename(excel_path)}")
    print(f"  Word   → {os.path.basename(docx_path)}")
    print(f"  Chart  → {os.path.basename(pie_path)}")
    print("=" * 60)


if __name__ == "__main__":
    main()
