#!/usr/bin/env python3
"""
Generates the PDF User Manual for the FRACC TIS Accreditation Status Report Generator.
Run once to produce: FRACC Report Generator - User Manual.pdf

Dependencies:
    pip install reportlab
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, ListFlowable, ListItem, KeepTogether
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.platypus.flowables import Flowable
import os

# ── Brand colours ────────────────────────────────────────────────
RDG_NAVY   = colors.HexColor("#1F3864")
RDG_BLUE   = colors.HexColor("#2E74B5")
RDG_LIGHT  = colors.HexColor("#D6E4F0")
RDG_GREEN  = colors.HexColor("#00B050")
RDG_RED    = colors.HexColor("#C00000")
RDG_AMBER  = colors.HexColor("#FFC000")
LIGHT_GREY = colors.HexColor("#F5F5F5")
MID_GREY   = colors.HexColor("#BFBFBF")
WHITE      = colors.white
BLACK      = colors.black

OUTPUT_FILE = "FRACC Report Generator - User Manual.pdf"


def build_styles():
    base = getSampleStyleSheet()

    styles = {}

    styles["cover_title"] = ParagraphStyle(
        "cover_title",
        fontName="Helvetica-Bold",
        fontSize=26,
        textColor=WHITE,
        alignment=TA_CENTER,
        spaceAfter=10,
        leading=32,
    )
    styles["cover_sub"] = ParagraphStyle(
        "cover_sub",
        fontName="Helvetica",
        fontSize=13,
        textColor=colors.HexColor("#D6E4F0"),
        alignment=TA_CENTER,
        spaceAfter=6,
    )
    styles["cover_meta"] = ParagraphStyle(
        "cover_meta",
        fontName="Helvetica",
        fontSize=10,
        textColor=colors.HexColor("#D6E4F0"),
        alignment=TA_CENTER,
    )
    styles["h1"] = ParagraphStyle(
        "h1",
        fontName="Helvetica-Bold",
        fontSize=16,
        textColor=RDG_NAVY,
        spaceBefore=18,
        spaceAfter=6,
        borderPad=4,
    )
    styles["h2"] = ParagraphStyle(
        "h2",
        fontName="Helvetica-Bold",
        fontSize=12,
        textColor=RDG_BLUE,
        spaceBefore=12,
        spaceAfter=4,
    )
    styles["h3"] = ParagraphStyle(
        "h3",
        fontName="Helvetica-BoldOblique",
        fontSize=10,
        textColor=RDG_NAVY,
        spaceBefore=8,
        spaceAfter=2,
    )
    styles["body"] = ParagraphStyle(
        "body",
        fontName="Helvetica",
        fontSize=9.5,
        textColor=BLACK,
        leading=14,
        spaceBefore=3,
        spaceAfter=3,
        alignment=TA_JUSTIFY,
    )
    styles["body_left"] = ParagraphStyle(
        "body_left",
        parent=styles["body"],
        alignment=TA_LEFT,
    )
    styles["code"] = ParagraphStyle(
        "code",
        fontName="Courier",
        fontSize=8.5,
        textColor=colors.HexColor("#1A1A2E"),
        backColor=colors.HexColor("#F0F4F8"),
        leading=13,
        leftIndent=12,
        rightIndent=12,
        spaceBefore=4,
        spaceAfter=4,
        borderPad=6,
        borderRadius=3,
    )
    styles["note"] = ParagraphStyle(
        "note",
        fontName="Helvetica-Oblique",
        fontSize=9,
        textColor=colors.HexColor("#555555"),
        leftIndent=12,
        spaceBefore=3,
        spaceAfter=3,
    )
    styles["table_hdr"] = ParagraphStyle(
        "table_hdr",
        fontName="Helvetica-Bold",
        fontSize=9,
        textColor=WHITE,
        alignment=TA_CENTER,
    )
    styles["table_cell"] = ParagraphStyle(
        "table_cell",
        fontName="Helvetica",
        fontSize=9,
        textColor=BLACK,
        alignment=TA_LEFT,
        leading=12,
    )
    styles["table_cell_c"] = ParagraphStyle(
        "table_cell_c",
        parent=styles["table_cell"],
        alignment=TA_CENTER,
    )
    styles["footer"] = ParagraphStyle(
        "footer",
        fontName="Helvetica",
        fontSize=8,
        textColor=MID_GREY,
        alignment=TA_CENTER,
    )
    return styles


# ── Cover page background rectangle ─────────────────────────────
class ColourRect(Flowable):
    def __init__(self, width, height, fill_colour):
        super().__init__()
        self.width  = width
        self.height = height
        self.fill   = fill_colour

    def draw(self):
        self.canv.setFillColor(self.fill)
        self.canv.rect(0, 0, self.width, self.height, fill=1, stroke=0)


def table_style(hdr_colour=RDG_NAVY, row_alt=LIGHT_GREY):
    return TableStyle([
        ("BACKGROUND",  (0, 0), (-1, 0),  hdr_colour),
        ("TEXTCOLOR",   (0, 0), (-1, 0),  WHITE),
        ("FONTNAME",    (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",    (0, 0), (-1, 0),  9),
        ("ALIGN",       (0, 0), (-1, 0),  "CENTER"),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME",    (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE",    (0, 1), (-1, -1), 9),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, row_alt]),
        ("GRID",        (0, 0), (-1, -1), 0.4, MID_GREY),
        ("TOPPADDING",  (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 0), (-1, 0), [hdr_colour]),
    ])


def build_manual():
    S = build_styles()
    PAGE_W, PAGE_H = A4
    MARGIN = 2 * cm
    DOC_W  = PAGE_W - 2 * MARGIN

    doc = SimpleDocTemplate(
        OUTPUT_FILE,
        pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN,  bottomMargin=2.2 * cm,
        title="FRACC Report Generator - User Manual",
        author="Rail Delivery Group — Service Assurance",
    )

    story = []

    # ════════════════════════════════════════════════════════════
    # COVER PAGE
    # ════════════════════════════════════════════════════════════
    story.append(ColourRect(DOC_W, 3.5 * cm, RDG_NAVY))
    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph("FRACC TIS Accreditation Status", S["cover_title"]))
    story.append(Paragraph("Report Generator", S["cover_title"]))
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph("User Manual", S["cover_sub"]))
    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph("Rail Delivery Group — Service Assurance", S["cover_meta"]))
    story.append(Paragraph("Version 1.0 &nbsp;|&nbsp; June 2026", S["cover_meta"]))
    story.append(Spacer(1, 0.6 * cm))
    story.append(HRFlowable(width=DOC_W, thickness=1.5, color=RDG_BLUE))
    story.append(Spacer(1, 0.4 * cm))

    # ════════════════════════════════════════════════════════════
    # SECTION 1 — INTRODUCTION
    # ════════════════════════════════════════════════════════════
    story.append(Paragraph("1.  Introduction", S["h1"]))
    story.append(HRFlowable(width=DOC_W, thickness=1, color=RDG_LIGHT))
    story.append(Spacer(1, 0.2 * cm))
    story.append(Paragraph(
        "The <b>FRACC TIS Accreditation Status Report Generator</b> is a command-line Python "
        "tool that produces the monthly FRACC paper and supporting Excel data file automatically. "
        "It takes two user-supplied inputs — a collated Excel report and a Word template — and "
        "applies all required data, formatting, and document patches in a single run.",
        S["body"]
    ))
    story.append(Paragraph(
        "The tool was designed and refined by the Service Assurance team at Rail Delivery Group "
        "to replicate and replace the manual steps previously performed via Excel macros and "
        "manual copy-paste into Word.",
        S["body"]
    ))

    story.append(Paragraph("1.1  Outputs", S["h2"]))
    out_data = [
        ["File", "Description"],
        ["FRACC - TIS Accreditation Status - Latest - <timestamp>.xlsx",
         "Colour-coded Excel with data grouped by Owning Group, TOC, System, and accreditation state"],
        ["FRACC Paper - Accreditation Status Update - <date>.docx",
         "Fully populated Word paper with updated dates, embedded chart, hyperlinked appendix list"],
        ["pie_chart_fracc.png",
         "Accreditation status pie chart (intermediate file, embedded in the Word doc)"],
    ]
    out_tbl = Table(out_data, colWidths=[7 * cm, DOC_W - 7 * cm])
    out_tbl.setStyle(table_style())
    story.append(out_tbl)

    # ════════════════════════════════════════════════════════════
    # SECTION 2 — PREREQUISITES
    # ════════════════════════════════════════════════════════════
    story.append(Paragraph("2.  Prerequisites", S["h1"]))
    story.append(HRFlowable(width=DOC_W, thickness=1, color=RDG_LIGHT))
    story.append(Spacer(1, 0.2 * cm))

    story.append(Paragraph("2.1  Python version", S["h2"]))
    story.append(Paragraph("Python 3.9 or later is required.", S["body"]))

    story.append(Paragraph("2.2  Install dependencies", S["h2"]))
    story.append(Paragraph("Run the following command once to install all required packages:", S["body"]))
    story.append(Paragraph(
        "pip install pandas openpyxl python-docx matplotlib lxml Pillow",
        S["code"]
    ))

    story.append(Paragraph("2.3  Required input files", S["h2"]))
    req_data = [
        ["File", "Format", "Description"],
        ["Collated TOC TIS Report", ".xlsx",
         "Master collated report produced by the TIS Accreditation Status Agent, "
         "or a direct export from the portal in pivot format"],
        ["FRACC Paper Template", ".docx",
         "The Word template from the previous month's paper. Must contain bookmarks, "
         "placeholder images, and the standard paragraph structure (see Section 5)"],
    ]
    req_tbl = Table(req_data, colWidths=[4.5 * cm, 2 * cm, DOC_W - 6.5 * cm])
    req_tbl.setStyle(table_style())
    story.append(req_tbl)

    # ════════════════════════════════════════════════════════════
    # SECTION 3 — RUNNING THE TOOL
    # ════════════════════════════════════════════════════════════
    story.append(Paragraph("3.  Running the Tool", S["h1"]))
    story.append(HRFlowable(width=DOC_W, thickness=1, color=RDG_LIGHT))
    story.append(Spacer(1, 0.2 * cm))

    story.append(Paragraph("3.1  Basic usage", S["h2"]))
    story.append(Paragraph(
        "Open a terminal, navigate to the folder containing <code>generate_fracc_paper.py</code>, "
        "and run:",
        S["body"]
    ))
    story.append(Paragraph(
        'python generate_fracc_paper.py \\\n'
        '    --collated  "TOC TIS Accreditation Status - Collated Report - v3.0.xlsx" \\\n'
        '    --template  "FRACC Paper Template.docx" \\\n'
        '    --paper-date "16 June 2026" \\\n'
        '    --output-dir "./output"',
        S["code"]
    ))

    story.append(Paragraph("3.2  All parameters", S["h2"]))
    param_data = [
        ["Parameter", "Required", "Description / Example"],
        ["--collated",    "Yes", 'Path to the collated Excel report.\n"TOC TIS Accreditation Status - Collated Report - v3.0.xlsx"'],
        ["--template",    "Yes", 'Path to the Word template.\n"FRACC Paper Template.docx"'],
        ["--paper-date",  "Yes", 'Paper date for the document header.\n"16 June 2026"  or  "16/06/2026"'],
        ["--output-dir",  "No",  'Folder to write output files into.\nDefault: current directory'],
        ["--news-image",  "No",  'Path to an optional ASSIST News widget screenshot (.png).\nIf supplied, replaces the news image placeholder in the Word doc'],
    ]
    param_tbl = Table(param_data, colWidths=[3.5 * cm, 2 * cm, DOC_W - 5.5 * cm])
    param_tbl.setStyle(table_style())
    story.append(param_tbl)

    story.append(Paragraph("3.3  Expected console output", S["h2"]))
    story.append(Paragraph(
        "When the tool runs successfully you will see output similar to the following:",
        S["body"]
    ))
    story.append(Paragraph(
        "============================================================\n"
        "  FRACC TIS Accreditation Status Report Generator\n"
        "============================================================\n"
        "  Collated report : TOC TIS Accreditation Status...xlsx\n"
        "  Template        : FRACC Paper Template.docx\n"
        "  Paper date      : 16 June 2026\n"
        "  Output dir      : ./output\n"
        "============================================================\n\n"
        "[1] Loading collated report: ...\n"
        "[2] Applying peak-count logic...\n"
        "    182 rows → 121 unique combinations after peak-count\n"
        "[3] Building FRACC Excel file...\n"
        "    Saved FRACC Excel → FRACC - TIS Accreditation Status - Latest - 20260616_0930.xlsx\n"
        "[4] Generating pie chart...\n"
        "    Saved pie chart → pie_chart_fracc.png\n"
        "[5] Patching Word document...\n"
        "    ✓ Meeting Date → 'XX Month' (highlighted)\n"
        "    ✓ Paper Date → '16 June 2026'\n"
        "    ✓ Update para date → '16 June 2026'\n"
        "    ✓ Trailer table date → 'XX Month' (highlighted)\n"
        "    ✓ Converted 16 appendix entries to internal hyperlinks\n"
        "    Saved Word document → FRACC Paper - Accreditation Status Update - 16 June 2026.docx\n\n"
        "============================================================\n"
        "  ✓  Generation complete\n"
        "============================================================",
        S["code"]
    ))

    # ════════════════════════════════════════════════════════════
    # SECTION 4 — HOW IT WORKS
    # ════════════════════════════════════════════════════════════
    story.append(PageBreak())
    story.append(Paragraph("4.  How It Works", S["h1"]))
    story.append(HRFlowable(width=DOC_W, thickness=1, color=RDG_LIGHT))
    story.append(Spacer(1, 0.2 * cm))

    story.append(Paragraph("4.1  Step 1 — Load & normalise", S["h2"]))
    story.append(Paragraph(
        "The tool reads the collated report and auto-detects its format. Two formats are supported:",
        S["body"]
    ))
    step1_data = [
        ["Format", "Detection", "Notes"],
        ["Pivot export",
         "Has an 'Owning Group' column",
         "Standard output from the TIS Accreditation Status Agent. Used directly."],
        ["Raw weekly format",
         "Header at row 4, data from row 5",
         "Direct export from the RSP portal. Owning Group is inferred where possible."],
    ]
    step1_tbl = Table(step1_data, colWidths=[3.5 * cm, 4 * cm, DOC_W - 7.5 * cm])
    step1_tbl.setStyle(table_style())
    story.append(step1_tbl)

    story.append(Paragraph("4.2  Step 2 — Peak-count logic", S["h2"]))
    story.append(Paragraph(
        "For every unique combination of <b>Owning Group + TOC/LPC + System + Machine Type ID + "
        "Version</b>, the tool selects the <b>highest device count ever recorded</b> across all "
        "historical weekly blocks in the collated file. The State and CoA Expiry date are taken "
        "from the same row as the peak count.",
        S["body"]
    ))
    story.append(Paragraph(
        "This approach avoids spurious dips caused by portal shifts being missing or incomplete "
        "on any given week, giving a stable and accurate representation of the true accredited "
        "fleet size.",
        S["body"]
    ))

    story.append(Paragraph("4.3  Step 3 — FRACC Excel", S["h2"]))
    story.append(Paragraph(
        "A three-sheet Excel workbook is built:",
        S["body"]
    ))
    xl_data = [
        ["Sheet", "Content"],
        ["TIS Accreditation Status",
         "Data grouped by Owning Group header rows (RDG navy) → TOC sub-headers (light blue) → "
         "column headers → data rows. State cells are colour-coded."],
        ["Pivot Data",
         "Flat table of all peak-count records. Suitable for pivot tables, mail merge, or "
         "further analysis."],
        ["Accreditation Status Chart",
         "Summary counts by State used to drive the pie chart. Shows total device count."],
    ]
    xl_tbl = Table(xl_data, colWidths=[4.5 * cm, DOC_W - 4.5 * cm])
    xl_tbl.setStyle(table_style())
    story.append(xl_tbl)

    story.append(Paragraph("4.4  Step 4 — Pie chart", S["h2"]))
    story.append(Paragraph(
        "A Matplotlib pie chart is generated at 180 DPI showing the proportion of devices in "
        "each accreditation state. Labels showing percentage and device count are placed outside "
        "the wedges with leader lines. The legend is positioned to the right with no overlap.",
        S["body"]
    ))

    story.append(Paragraph("4.5  Step 5 — Word document patches", S["h2"]))
    patch_data = [
        ["#", "What changes", "Detail"],
        ["1", "Meeting Date",
         "Set to 'XX Month' with yellow highlight. The reviewer fills this in before distribution."],
        ["2", "Paper Date",
         "Set to the value supplied via --paper-date."],
        ["3", "Update paragraph",
         "The bold date in 'this report was run on [date]' is updated to match the Paper Date."],
        ["4", "Trailer table",
         "The date cell in the Governance Lifecycle Tracking Grid (last table in the doc) is set "
         "to 'XX Month' with yellow highlight."],
        ["5", "Pie chart image",
         "The existing chart placeholder (image1.png in the docx zip) is replaced with the newly "
         "generated chart."],
        ["6", "News image",
         "If --news-image is supplied, the ASSIST News screenshot placeholder (image2.png) is "
         "replaced with the provided image."],
        ["7", "Appendix hyperlinks",
         "All 16 entries in the appendix contents list (A–P) are converted to internal "
         "hyperlinks (w:hyperlink with w:anchor) pointing to named bookmarks in the document body."],
    ]
    patch_tbl = Table(patch_data, colWidths=[0.8 * cm, 3.8 * cm, DOC_W - 4.6 * cm])
    patch_tbl.setStyle(table_style())
    story.append(patch_tbl)

    # ════════════════════════════════════════════════════════════
    # SECTION 5 — TEMPLATE REQUIREMENTS
    # ════════════════════════════════════════════════════════════
    story.append(Paragraph("5.  Word Template Requirements", S["h1"]))
    story.append(HRFlowable(width=DOC_W, thickness=1, color=RDG_LIGHT))
    story.append(Spacer(1, 0.2 * cm))
    story.append(Paragraph(
        "The Word template must meet the following structural requirements for the patches to "
        "apply correctly. The easiest way to satisfy all requirements is to use the previous "
        "month's paper as the template.",
        S["body"]
    ))

    tmpl_data = [
        ["Requirement", "Details"],
        ["Meeting Date paragraph",
         "A paragraph beginning 'Meeting Date:' with a tab and the previous date as a run."],
        ["Paper Date paragraph",
         "A paragraph beginning 'Paper Date:' with the previous date as a run."],
        ["Update paragraph",
         "A paragraph containing the phrase 'this report was run on' with the date in two "
         "consecutive bold runs (e.g. '28 April' and ' 2026')."],
        ["Trailer/governance table",
         "The last table in the document must have the meeting date in column 2 (index 1), "
         "row 1 (index 0), highlighted yellow."],
        ["Image placeholders",
         "Two embedded PNG images: image1.png (pie chart) and image2.png (news screenshot). "
         "These are in word/media/ inside the docx zip."],
        ["Appendix bookmarks",
         "Section heading paragraphs for each appendix must carry a Word bookmarkStart with "
         "names: Appendix_A … Appendix_L, AppendixM … AppendixP."],
        ["Appendix contents list",
         "Paragraphs before END OF PAPER reading 'Appendix A - …' through 'Appendix P - …'. "
         "These are converted to hyperlinks."],
        ["Hyperlink character style",
         "A character style named 'Hyperlink' must be defined (present in all Word documents "
         "by default)."],
    ]
    tmpl_tbl = Table(tmpl_data, colWidths=[4.5 * cm, DOC_W - 4.5 * cm])
    tmpl_tbl.setStyle(table_style())
    story.append(tmpl_tbl)

    # ════════════════════════════════════════════════════════════
    # SECTION 6 — COLOUR REFERENCE
    # ════════════════════════════════════════════════════════════
    story.append(Paragraph("6.  Colour Reference", S["h1"]))
    story.append(HRFlowable(width=DOC_W, thickness=1, color=RDG_LIGHT))
    story.append(Spacer(1, 0.2 * cm))

    story.append(Paragraph("6.1  Accreditation state colours (Excel)", S["h2"]))
    colour_data = [
        ["State", "Cell Fill", "Font Colour"],
        ["Accredited",             "#C6EFCE (light green)",   "#276221 (dark green)"],
        ["Pilot phase",            "#E2EFDA (very light green)", "#375623 (forest green)"],
        ["Accreditation expired",  "#FFC7CE (light red)",     "#9C0006 (dark red)"],
        ["Application acknowledged", "#FFEB9C (light amber)", "#9C6500 (dark amber)"],
    ]
    colour_tbl = Table(colour_data, colWidths=[5 * cm, 5 * cm, DOC_W - 10 * cm])
    colour_tbl.setStyle(TableStyle([
        ("BACKGROUND",  (0, 0), (-1, 0),  RDG_NAVY),
        ("TEXTCOLOR",   (0, 0), (-1, 0),  WHITE),
        ("FONTNAME",    (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",    (0, 0), (-1, -1), 9),
        ("ALIGN",       (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME",    (0, 1), (-1, -1), "Helvetica"),
        # Row 1 — Accredited (green)
        ("BACKGROUND",  (0, 1), (-1, 1), colors.HexColor("#C6EFCE")),
        ("TEXTCOLOR",   (0, 1), (-1, 1), colors.HexColor("#276221")),
        ("FONTNAME",    (0, 1), (-1, 1), "Helvetica-Bold"),
        # Row 2 — Pilot
        ("BACKGROUND",  (0, 2), (-1, 2), colors.HexColor("#E2EFDA")),
        ("TEXTCOLOR",   (0, 2), (-1, 2), colors.HexColor("#375623")),
        ("FONTNAME",    (0, 2), (-1, 2), "Helvetica-Bold"),
        # Row 3 — Expired
        ("BACKGROUND",  (0, 3), (-1, 3), colors.HexColor("#FFC7CE")),
        ("TEXTCOLOR",   (0, 3), (-1, 3), colors.HexColor("#9C0006")),
        ("FONTNAME",    (0, 3), (-1, 3), "Helvetica-Bold"),
        # Row 4 — Acknowledged
        ("BACKGROUND",  (0, 4), (-1, 4), colors.HexColor("#FFEB9C")),
        ("TEXTCOLOR",   (0, 4), (-1, 4), colors.HexColor("#9C6500")),
        ("FONTNAME",    (0, 4), (-1, 4), "Helvetica-Bold"),
        ("GRID",        (0, 0), (-1, -1), 0.4, MID_GREY),
        ("TOPPADDING",  (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(colour_tbl)

    story.append(Paragraph("6.2  Excel structure colours", S["h2"]))
    struct_data = [
        ["Element", "Fill", "Font"],
        ["Owning Group header row",  "#1F3864 (RDG navy)",   "White, bold, 12pt"],
        ["TOC sub-header row",       "#D6E4F0 (light blue)", "#1F3864, bold, 10pt"],
        ["Column headers",           "#EBF3FB (very light blue)", "#1F3864, bold, 9pt"],
        ["Data rows",                "White / alternating light grey", "Black, 9pt"],
    ]
    struct_tbl = Table(struct_data, colWidths=[4.5 * cm, 4.5 * cm, DOC_W - 9 * cm])
    struct_tbl.setStyle(table_style())
    story.append(struct_tbl)

    # ════════════════════════════════════════════════════════════
    # SECTION 7 — APPENDIX STRUCTURE
    # ════════════════════════════════════════════════════════════
    story.append(Paragraph("7.  Appendix Structure", S["h1"]))
    story.append(HRFlowable(width=DOC_W, thickness=1, color=RDG_LIGHT))
    story.append(Spacer(1, 0.2 * cm))

    app_data = [
        ["Appendix", "Owning Group / Content", "Bookmark Name"],
        ["A",  "Go Ahead",                                          "Appendix_A"],
        ["B",  "Transport UK Group (formerly Abellio Group)",        "Appendix_B"],
        ["C",  "Serco / Transport UK Group (formerly Abellio Group)","Appendix_C"],
        ["D",  "First Group",                                        "Appendix_D"],
        ["E",  "Directly Operated Railway",                          "Appendix_E"],
        ["F",  "Arriva",                                             "Appendix_F"],
        ["G",  "Transport for Wales",                                "Appendix_G"],
        ["H",  "London Overground",                                  "Appendix_H"],
        ["I",  "GTS Elizabeth Line",                                 "Appendix_I"],
        ["K",  "Heathrow Express Operating Company",                 "Appendix_K"],
        ["L",  "Scottish Rail Holdings",                             "Appendix_L"],
        ["M",  "Accreditation Status (summary + pie chart)",         "AppendixM"],
        ["N",  "RDG Compliance Standards — Recent Updates",          "AppendixN"],
        ["O",  "Third Party Retailer Systems",                       "AppendixO"],
        ["P",  "Governance Lifecycle Tracking Grid",                 "AppendixP"],
    ]
    app_tbl = Table(app_data, colWidths=[2 * cm, 8 * cm, DOC_W - 10 * cm])
    app_tbl.setStyle(table_style())
    story.append(app_tbl)

    story.append(Paragraph(
        "Note: Appendix J is intentionally omitted. This is consistent with the original "
        "paper structure.",
        S["note"]
    ))

    # ════════════════════════════════════════════════════════════
    # SECTION 8 — TROUBLESHOOTING
    # ════════════════════════════════════════════════════════════
    story.append(PageBreak())
    story.append(Paragraph("8.  Troubleshooting", S["h1"]))
    story.append(HRFlowable(width=DOC_W, thickness=1, color=RDG_LIGHT))
    story.append(Spacer(1, 0.2 * cm))

    ts_data = [
        ["Symptom", "Likely Cause", "Fix"],
        ["'Owning Group' column missing",
         "Collated file is in raw weekly format, not pivot export",
         "This is handled automatically. Check that the file has a header row at row 4 "
         "with 'Lennon profit centre', 'Company', etc."],
        ["Meeting Date not updated",
         "The template paragraph text doesn't start with 'Meeting Date:'",
         "Check the template paragraph text. The script searches for month names "
         "(January–December) within that paragraph's runs."],
        ["Update para date not updated",
         "The phrase 'this report was run on' is not present in the template",
         "Check the Update section paragraph. Ensure the phrase is exact and the date "
         "runs are bold."],
        ["Hyperlinks not appearing",
         "Appendix list paragraphs don't start with 'Appendix A', 'Appendix B' etc.",
         "Check that the TOC paragraphs use that exact prefix. The script matches on "
         "the start of the stripped paragraph text."],
        ["'Appendix_I' bookmark missing",
         "Template was created before Appendix I had a bookmark",
         "The script automatically adds the missing bookmark. If still failing, open "
         "the template in Word and manually add a bookmark named 'Appendix_I' to the "
         "Appendix I heading."],
        ["Images not replaced",
         "Template docx contains more than two images, or images are named differently",
         "The script replaces word/media/image1.png (pie chart) and word/media/image2.png "
         "(news screenshot). Open the docx as a zip to verify image names."],
        ["ValueError: Cannot parse paper date",
         "Date string format not recognised",
         "Use 'DD Month YYYY' (e.g. '16 June 2026') or 'DD/MM/YYYY'. "
         "The script falls back to today's date with a warning if parsing fails."],
        ["ModuleNotFoundError",
         "Dependencies not installed",
         "Run: pip install pandas openpyxl python-docx matplotlib lxml Pillow"],
    ]
    ts_tbl = Table(ts_data, colWidths=[3.8 * cm, 4 * cm, DOC_W - 7.8 * cm])
    ts_tbl.setStyle(table_style())
    story.append(ts_tbl)

    # ════════════════════════════════════════════════════════════
    # SECTION 9 — FILE NAMING
    # ════════════════════════════════════════════════════════════
    story.append(Paragraph("9.  File Naming Conventions", S["h1"]))
    story.append(HRFlowable(width=DOC_W, thickness=1, color=RDG_LIGHT))
    story.append(Spacer(1, 0.2 * cm))

    fn_data = [
        ["File", "Name pattern", "Example"],
        ["FRACC Excel",
         "FRACC - TIS Accreditation Status - Latest - YYYYMMDD_HHMM.xlsx",
         "FRACC - TIS Accreditation Status - Latest - 20260616_0930.xlsx"],
        ["FRACC Word paper",
         "FRACC Paper - Accreditation Status Update - DD Month YYYY.docx",
         "FRACC Paper - Accreditation Status Update - 16 June 2026.docx"],
        ["Pie chart",
         "pie_chart_fracc.png",
         "pie_chart_fracc.png (kept in output dir for reference)"],
    ]
    fn_tbl = Table(fn_data, colWidths=[3.5 * cm, 6.5 * cm, DOC_W - 10 * cm])
    fn_tbl.setStyle(table_style())
    story.append(fn_tbl)

    # ════════════════════════════════════════════════════════════
    # SECTION 10 — VERSION HISTORY
    # ════════════════════════════════════════════════════════════
    story.append(Paragraph("10.  Version History", S["h1"]))
    story.append(HRFlowable(width=DOC_W, thickness=1, color=RDG_LIGHT))
    story.append(Spacer(1, 0.2 * cm))

    vh_data = [
        ["Version", "Date", "Changes"],
        ["1.0", "June 2026",
         "Initial release. Full pipeline: collated report → peak counts → Excel → "
         "pie chart → Word patches → hyperlinked appendix list."],
    ]
    vh_tbl = Table(vh_data, colWidths=[2 * cm, 2.5 * cm, DOC_W - 4.5 * cm])
    vh_tbl.setStyle(table_style())
    story.append(vh_tbl)

    # ════════════════════════════════════════════════════════════
    # FOOTER / BUILD
    # ════════════════════════════════════════════════════════════
    def add_page_number(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(MID_GREY)
        canvas.drawString(MARGIN, 1.2 * cm, "Rail Delivery Group — Service Assurance")
        canvas.drawRightString(PAGE_W - MARGIN, 1.2 * cm,
                               f"Page {doc.page}")
        canvas.setStrokeColor(RDG_LIGHT)
        canvas.setLineWidth(0.5)
        canvas.line(MARGIN, 1.5 * cm, PAGE_W - MARGIN, 1.5 * cm)
        canvas.restoreState()

    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    print(f"✓ PDF manual saved: {OUTPUT_FILE}")


if __name__ == "__main__":
    build_manual()
