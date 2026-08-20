#!/usr/bin/env python3
"""
DMIC Commissioning Contract Management Workbook — generator.

Builds the Excel workbook described in
DMIC_COMMISSIONING_CONTRACT_MANAGEMENT_WORKBOOK_v3.docx.

Usage:
    python3 build_workbook.py --demo  -o DMIC_..._DEMO.xlsx
    python3 build_workbook.py --blank -o DMIC_..._TEMPLATE.xlsx

Both variants are produced by the same code path so their structure,
formulas, validation and formatting are guaranteed identical; only the
data rows differ.
"""

import argparse
import datetime as dt

from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------------------
# Capacity
# ---------------------------------------------------------------------------
FIRST = 4                 # first data row on every tracker
N_SCOPE = 100
N_EQUIP = 100
N_DOC = 150               # documents get extra headroom: 1 row per deliverable per scope
N_ACTION = 100
N_MILESTONE = 30
N_TREND = 12

SCOPE_LAST = FIRST + N_SCOPE - 1        # 103
EQUIP_LAST = FIRST + N_EQUIP - 1        # 103
DOC_LAST = FIRST + N_DOC - 1            # 153
ACT_LAST = FIRST + N_ACTION - 1         # 103

MS_SECT, MS_HDR = 25, 26
MS_FIRST, MS_LAST = 27, 27 + N_MILESTONE - 1      # 27..56
TREND_FIRST, TREND_LAST = 27, 27 + N_TREND - 1    # 27..38
CHART_SECT = 59

# ---------------------------------------------------------------------------
# Reference data (spec section: SHEET 7 – CONFIGURATION)
# ---------------------------------------------------------------------------
CONTRACT_STATUS = ["Not Started", "Preparation", "In Progress", "Commissioning",
                   "Handover", "Complete", "On Hold"]

SCOPE_STATUS = ["Not Started", "Preparation", "Review", "Ready", "In Progress",
                "Complete", "On Hold"]

# (phase, progress fraction) — spec gives the 11 phases, the % ramp is our assumption
PHASES = [
    ("Pre-Commissioning", 0.00),
    ("Commissioning Preparation", 0.10),
    ("Documentation Review", 0.20),
    ("Testing Preparation", 0.30),
    ("Testing In Progress", 0.45),
    ("Results Review", 0.60),
    ("Ready For Handover", 0.70),
    ("Beneficial Use", 0.80),
    ("Further Testing", 0.85),
    ("Final Handover", 0.95),
    ("Complete", 1.00),
]

# (workflow status, progress fraction, counts as approved 1/0)
WORKFLOW = [
    ("Not Started", 0.00, 0),
    ("Draft", 0.15, 0),
    ("Internal Review", 0.30, 0),
    ("Submitted", 0.45, 0),
    ("Under Review", 0.55, 0),
    ("Comments Received", 0.50, 0),
    ("Rework", 0.40, 0),
    ("Resubmitted", 0.70, 0),
    ("Approved", 0.90, 1),
    ("Released", 0.95, 1),
    ("Implemented", 0.98, 1),
    ("Closed", 1.00, 1),
]

GATES = ["Preparation", "Readiness", "Authority To Start", "Commissioning",
         "Results", "Handover", "Final Takeover"]

PRIORITY = ["Critical", "High", "Medium", "Low"]
ACTION_STATUS = ["Open", "In Progress", "Waiting", "Closed"]
CRITICALITY = ["Critical", "High", "Medium", "Low"]
YES_NO = ["Yes", "No"]
MILESTONE_STATUS = ["Not Started", "In Progress", "At Risk", "Delayed", "Achieved"]

ACTION_CATEGORY = ["Document", "Technical", "Quality", "Safety", "Commercial",
                   "Resource", "Interface", "Client"]

# Fictional placeholder names for the demo file.
PERSONNEL = ["A. Mercer", "S. Okonkwo", "L. Fontaine", "D. Petrov", "M. Haddad",
             "J. Whitfield", "K. Rasmussen", "P. Nakamura", "R. Bellamy",
             "T. Oyelaran", "N. Bergstrom", "C. Vasquez", "H. Lindqvist",
             "F. Moreau", "G. Ashworth", "E. Devereux"]

# (document type, spec lead-time label, lead days before Planned Commissioning Start)
# Negative lead days = the target date falls AFTER commissioning start.
DOC_TYPES = [
    ("Basis of Commissioning", "T-12 Months", 360),
    ("Surveillance Strategy", "T-12 Months", 360),
    ("Surveillance Programme", "T-9 Months", 270),
    ("Competency Evidence", "T-9 Months", 270),
    ("Safe System Of Work", "T-6 Months", 180),
    ("Risk Assessment And Method Statement", "T-6 Months", 180),
    ("System Commissioning Programme", "T-6 Months", 180),
    ("Commissioning Procedure", "T-3 Months", 90),
    ("Inspection and Test Plan", "T-3 Months", 90),
    ("Initial End Of Erection Status Report", "T-3 Months", 90),
    ("Planned End Of Erection Status Report", "T-1 Month", 30),
    ("Authority To Start Evidence", "T-2 Weeks", 14),
    ("Commissioning Results Sheet", "During Commissioning", -30),
    ("Commissioning Phase Report", "Commissioning Completion", -60),
    ("Handover Certificate", "Before FTC", -75),
    ("Functional Takeover Package", "Before FTC", -75),
    ("Care And Maintenance Plan", "Before TFM (assumed)", -90),
    ("Maintenance Transfer Package", "Before TFM", -90),
    ("Temporary Operations Transfer Package", "Before TFTO", -105),
    ("Final End Of Erection Status Report", "Before Final Takeover", -120),
]
DOC_TYPE_NAMES = [d[0] for d in DOC_TYPES]

MILESTONES = [
    ("Kick Off Meeting", 360),
    ("Commissioning Readiness Review", 60),
    ("Authority To Start", 14),
    ("Commissioning Start", 0),
    ("Functional Takeover", -75),
    ("Beneficial Use", -85),
    ("Maintenance Transfer", -100),
    ("Temporary Operations Transfer", -115),
    ("Final Takeover", -130),
]
MILESTONE_NAMES = [m[0] for m in MILESTONES]

# Spec: ATS Readiness Gate — minimum required approved deliverables
ATS_REQUIRED = [
    "Planned End Of Erection Status Report",
    "System Commissioning Programme",
    "Commissioning Procedure",
    "Inspection and Test Plan",
    "Surveillance Strategy",
    "Surveillance Programme",
    "Competency Evidence",
    "Safe System Of Work",
    "Risk Assessment And Method Statement",
    "Authority To Start Evidence",
]

HEALTH_COMPONENTS = [
    ("Documents", 0.40),
    ("Readiness", 0.20),
    ("Actions", 0.15),
    ("Milestones", 0.10),
    ("Equipment Progress", 0.10),
    ("Handovers", 0.05),
]

# ---------------------------------------------------------------------------
# Configuration sheet geometry
# ---------------------------------------------------------------------------
CFG = "Configuration"
LIST_HDR, LIST_TOP = 4, 5

MAP_HDR, MAP_TOP = 29, 30          # phase / workflow maps and parameter blocks
PHASE_TOP, PHASE_BOT = 30, 30 + len(PHASES) - 1        # 30..40
WF_TOP, WF_BOT = 30, 30 + len(WORKFLOW) - 1            # 30..41
HW_TOP, HW_BOT = 30, 30 + len(HEALTH_COMPONENTS) - 1   # 30..35
HW_TOTAL = HW_BOT + 1                                  # 36
THR_TOP = 30                                           # 30..38

ATS_TITLE, ATS_HDR = 44, 45
ATS_TOP, ATS_BOT = 46, 46 + len(ATS_REQUIRED) - 1      # 46..55
ATS_SCORE, ATS_STATUS = 57, 58

CD_TITLE, CD_HDR, CD_TOP = 200, 202, 203

# ---------------------------------------------------------------------------
# Control Panel geometry
# ---------------------------------------------------------------------------
CP = "Control Panel"
CP_INFO_TOP = 5                     # 5..11  contract information (7 fields)
CP_DATE_TOP = 14                    # 14..19 commissioning dates (6 fields)
CP_LEAD_HDR = 31
CP_LEAD_TOP = 32
CP_LEAD_BOT = CP_LEAD_TOP + len(DOC_TYPES) - 1          # 32..51
CP_MS_HDR = 54
CP_MS_TOP = 55
CP_MS_BOT = CP_MS_TOP + len(MILESTONES) - 1             # 55..63

DASH = "Dashboard"

# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------
FONT = "Arial"
NAVY = "1F3864"
BLUE = "2F5597"
LIGHT = "D9E2F3"
GREY_FILL = "F2F2F2"
INPUT_FILL = "FFFFFF"
KEY_FILL = "FFF2CC"

f_title = Font(name=FONT, size=14, bold=True, color=NAVY)
f_sub = Font(name=FONT, size=9, italic=True, color="595959")
f_sect = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_hdr = Font(name=FONT, size=9, bold=True, color="FFFFFF")
f_lbl = Font(name=FONT, size=10, bold=True, color="000000")
f_in = Font(name=FONT, size=10, color="0000FF")          # user input
f_calc = Font(name=FONT, size=10, color="000000")        # formula
f_body = Font(name=FONT, size=9)
f_calcs = Font(name=FONT, size=9, color="404040")
f_note = Font(name=FONT, size=8, italic=True, color="808080")
f_big = Font(name=FONT, size=36, bold=True, color=NAVY)

fill_sect = PatternFill("solid", fgColor=BLUE)
fill_hdr = PatternFill("solid", fgColor=NAVY)
fill_light = PatternFill("solid", fgColor=LIGHT)
fill_calc = PatternFill("solid", fgColor=GREY_FILL)
fill_key = PatternFill("solid", fgColor=KEY_FILL)
fill_example = PatternFill("solid", fgColor="FFF9E6")

thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)

c_center = Alignment(horizontal="center", vertical="center")
c_left = Alignment(horizontal="left", vertical="center")
c_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
c_hdr = Alignment(horizontal="center", vertical="center", wrap_text=True)

FMT_DATE = "DD-MMM-YYYY"
FMT_PCT = "0.0%"
FMT_INT = "#,##0"
FMT_DAYS = "#,##0;[Red]-#,##0;-"

# Spec: Document Colour Coding
DOC_COLOURS = {
    "grey": "D9D9D9", "blue": "BDD7EE", "yellow": "FFE699",
    "orange": "F8CBAD", "red": "FFC7CE", "green": "C6EFCE", "dgreen": "A9D08E",
}
# Spec: Action Colour Rules
ACT_COLOURS = {
    "darkred": "C00000", "red": "FFC7CE", "amber": "FFE699",
    "blue": "BDD7EE", "green": "C6EFCE", "grey": "D9D9D9",
}


def section(ws, row, first_col, last_col, text):
    """Full-width coloured section banner."""
    ws.merge_cells(start_row=row, start_column=first_col,
                   end_row=row, end_column=last_col)
    c = ws.cell(row=row, column=first_col, value=text)
    c.font = f_sect
    c.fill = fill_sect
    c.alignment = c_left
    for col in range(first_col, last_col + 1):
        ws.cell(row=row, column=col).fill = fill_sect
    ws.row_dimensions[row].height = 18


def header_row(ws, row, first_col, headers):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=first_col + i, value=h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = c_hdr
        c.border = box
    ws.row_dimensions[row].height = 30


def label(ws, row, col, text, note=None):
    c = ws.cell(row=row, column=col, value=text)
    c.font = f_lbl
    c.alignment = c_left
    if note:
        c.comment = Comment(note, "Workbook")
    return c


def input_cell(ws, row, col, value=None, fmt=None, key=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = f_in
    c.fill = fill_key if key else PatternFill("solid", fgColor=INPUT_FILL)
    c.border = box
    c.alignment = c_left
    if fmt:
        c.number_format = fmt
    return c


def calc_cell(ws, row, col, formula, fmt=None, note=None):
    c = ws.cell(row=row, column=col, value=formula)
    c.font = f_calc
    c.fill = fill_calc
    c.border = box
    c.alignment = c_left
    if fmt:
        c.number_format = fmt
    if note:
        c.comment = Comment(note, "Workbook")
    return c


def fill_down(ws, col, first, last, template, fmt=None, calc=True):
    """Write a per-row formula down a column. {r} is replaced by the row number."""
    for r in range(first, last + 1):
        c = ws.cell(row=r, column=col, value=template.format(r=r))
        c.font = f_calcs if calc else f_body
        c.border = box
        c.alignment = c_left
        if calc:
            c.fill = fill_calc
        if fmt:
            c.number_format = fmt


def style_input_range(ws, first_col, last_col, first, last, fmt_map=None):
    for r in range(first, last + 1):
        for col in range(first_col, last_col + 1):
            c = ws.cell(row=r, column=col)
            c.font = f_body
            c.border = box
            c.alignment = c_left
            if fmt_map and col in fmt_map:
                c.number_format = fmt_map[col]


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


# ===========================================================================
# CONFIGURATION (hidden)
# ===========================================================================
def build_configuration(ws):
    ws.sheet_properties.tabColor = "808080"
    ws["A1"] = "CONFIGURATION — WORKBOOK LOGIC"
    ws["A1"].font = f_title
    ws["A2"] = ("Administrative sheet. Dropdown lists, progress maps, Health Index weights and "
                "RAG thresholds live here. Lead times are edited on the Control Panel. "
                "Do not delete rows or columns — every other sheet references these ranges by name.")
    ws["A2"].font = f_sub

    lists = [
        ("A", "Contract Status", CONTRACT_STATUS),
        ("B", "Scope Status", SCOPE_STATUS),
        ("C", "Commissioning Phase", [p[0] for p in PHASES]),
        ("D", "Workflow Status", [w[0] for w in WORKFLOW]),
        ("E", "Readiness Gate", GATES),
        ("F", "Priority", PRIORITY),
        ("G", "Action Status", ACTION_STATUS),
        ("H", "Criticality", CRITICALITY),
        ("I", "Yes / No", YES_NO),
        ("J", "Milestone Status", MILESTONE_STATUS),
        ("K", "Document Type", DOC_TYPE_NAMES),
        ("L", "Personnel", PERSONNEL),
        ("M", "Action Category", ACTION_CATEGORY),
        ("N", "Milestone Name", MILESTONE_NAMES),
    ]
    for col, title, values in lists:
        c = ws[f"{col}{LIST_HDR}"]
        c.value = title
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = c_hdr
        c.border = box
        for i, v in enumerate(values):
            cc = ws.cell(row=LIST_TOP + i, column=c.column, value=v)
            cc.font = f_body
            cc.border = box
        ws.column_dimensions[col].width = 30

    # --- Phase progress map -------------------------------------------------
    ws["A28"] = "PHASE PROGRESS MAP"
    ws["A28"].font = f_lbl
    header_row(ws, MAP_HDR, 1, ["Commissioning Phase", "Progress %"])
    for i, (name, pct) in enumerate(PHASES):
        r = PHASE_TOP + i
        ws.cell(row=r, column=1, value=name).font = f_body
        c = ws.cell(row=r, column=2, value=pct)
        c.font = f_in
        c.number_format = FMT_PCT
        for col in (1, 2):
            ws.cell(row=r, column=col).border = box
    ws.cell(row=PHASE_BOT + 2, column=1,
            value="Assumption: the % ramp across the 11 spec phases is not defined in the "
                  "source document. Edit the blue cells to retune.").font = f_note

    # --- Workflow status map ------------------------------------------------
    ws["D28"] = "WORKFLOW STATUS MAP"
    ws["D28"].font = f_lbl
    header_row(ws, MAP_HDR, 4, ["Workflow Status", "Progress %", "Counts As Approved"])
    for i, (name, pct, appr) in enumerate(WORKFLOW):
        r = WF_TOP + i
        ws.cell(row=r, column=4, value=name).font = f_body
        c = ws.cell(row=r, column=5, value=pct)
        c.font = f_in
        c.number_format = FMT_PCT
        c2 = ws.cell(row=r, column=6, value=appr)
        c2.font = f_in
        c2.alignment = c_center
        for col in (4, 5, 6):
            ws.cell(row=r, column=col).border = box

    # --- Readiness weights --------------------------------------------------
    ws["H28"] = "READINESS & TIMING PARAMETERS"
    ws["H28"].font = f_lbl
    header_row(ws, MAP_HDR, 8, ["Parameter", "Value"])
    params = [
        ("Equipment readiness — phase weight", 0.50, FMT_PCT),
        ("Equipment readiness — document weight", 0.50, FMT_PCT),
        ("Review days (submission → approval)", 21, FMT_INT),
    ]
    for i, (name, val, fmt) in enumerate(params):
        r = MAP_TOP + i
        ws.cell(row=r, column=8, value=name).font = f_body
        c = ws.cell(row=r, column=9, value=val)
        c.font = f_in
        c.number_format = fmt
        for col in (8, 9):
            ws.cell(row=r, column=col).border = box
    ws.cell(row=MAP_TOP + 4, column=8,
            value="Equipment with no linked documents falls back to 100% phase weight.").font = f_note

    # --- Health index weights ----------------------------------------------
    ws["K28"] = "CONTRACT HEALTH INDEX WEIGHTS"
    ws["K28"].font = f_lbl
    header_row(ws, MAP_HDR, 11, ["Component", "Weight"])
    for i, (name, w) in enumerate(HEALTH_COMPONENTS):
        r = HW_TOP + i
        ws.cell(row=r, column=11, value=name).font = f_body
        c = ws.cell(row=r, column=12, value=w)
        c.font = f_in
        c.number_format = FMT_PCT
        for col in (11, 12):
            ws.cell(row=r, column=col).border = box
    ws.cell(row=HW_TOTAL, column=11, value="TOTAL").font = f_lbl
    tc = ws.cell(row=HW_TOTAL, column=12, value=f"=SUM(L{HW_TOP}:L{HW_BOT})")
    tc.font = f_lbl
    tc.number_format = FMT_PCT
    tc.border = box

    # --- Thresholds ---------------------------------------------------------
    ws["N28"] = "THRESHOLDS"
    ws["N28"].font = f_lbl
    header_row(ws, MAP_HDR, 14, ["Threshold", "Value"])
    thresholds = [
        ("Health — Excellent (min)", 0.90, FMT_PCT),
        ("Health — Good (min)", 0.75, FMT_PCT),
        ("Health — Attention Required (min)", 0.50, FMT_PCT),
        ("RAG — Green (min)", 0.75, FMT_PCT),
        ("RAG — Amber (min)", 0.50, FMT_PCT),
        ("Document due-soon window (days)", 30, FMT_INT),
        ("Action due-soon window (days)", 14, FMT_INT),
        ("Scope 'Ready' threshold", 0.80, FMT_PCT),
        ("Upcoming milestone window (days)", 30, FMT_INT),
    ]
    for i, (name, val, fmt) in enumerate(thresholds):
        r = THR_TOP + i
        ws.cell(row=r, column=14, value=name).font = f_body
        c = ws.cell(row=r, column=15, value=val)
        c.font = f_in
        c.number_format = fmt
        for col in (14, 15):
            ws.cell(row=r, column=col).border = box

    # --- ATS readiness gate -------------------------------------------------
    ws.cell(row=ATS_TITLE, column=1,
            value="ATS READINESS GATE — MINIMUM REQUIRED APPROVED DELIVERABLES").font = f_lbl
    header_row(ws, ATS_HDR, 1,
               ["Required Document Type", "Approved", "Outstanding", "Met (1/0)"])
    ws.cell(row=ATS_HDR, column=4).comment = Comment(
        "A required type counts as met only when at least one document of that type exists AND "
        "none of that type is still outstanding. Counting merely 'one approved somewhere' would "
        "turn the gate green while other scopes still had the deliverable open.", "Workbook")
    for i, name in enumerate(ATS_REQUIRED):
        r = ATS_TOP + i
        ws.cell(row=r, column=1, value=name).font = f_body
        ws.cell(row=r, column=2, value=f"=COUNTIFS(Doc_Type,$A{r},Doc_Approved,1)").font = f_calcs
        ws.cell(row=r, column=3, value=f"=COUNTIFS(Doc_Type,$A{r},Doc_Approved,0)").font = f_calcs
        ws.cell(row=r, column=4, value=f"=IF(AND($B{r}>0,$C{r}=0),1,0)").font = f_calcs
        for col in (1, 2, 3, 4):
            ws.cell(row=r, column=col).border = box
    ws.cell(row=ATS_SCORE, column=1, value="ATS Readiness Score").font = f_lbl
    # Share of all ATS-required deliverables that are approved. A type-by-type binary
    # score reads 0% until the very last document lands, which is useless as a KPI;
    # the strict all-types-met test drives the Status below instead.
    sc = ws.cell(row=ATS_SCORE, column=2, value=(
        f"=IFERROR(SUM(B{ATS_TOP}:B{ATS_BOT})/"
        f"(SUM(B{ATS_TOP}:B{ATS_BOT})+SUM(C{ATS_TOP}:C{ATS_BOT})),0)"))
    sc.font = f_lbl
    sc.number_format = FMT_PCT
    sc.border = box
    ws.cell(row=ATS_SCORE, column=1).comment = Comment(
        "Approved ATS-required deliverables as a share of all ATS-required deliverables "
        "in the tracker.", "Workbook")

    ws.cell(row=ATS_STATUS, column=1, value="ATS Status").font = f_lbl
    st = ws.cell(row=ATS_STATUS, column=2, value=(
        f'=IF(SUM(D{ATS_TOP}:D{ATS_BOT})=COUNTA(A{ATS_TOP}:A{ATS_BOT}),"Ready",'
        f'IF($B${ATS_SCORE}>=0.8,"At Risk","Not Ready"))'))
    st.font = f_lbl
    st.border = box
    ws.cell(row=ATS_STATUS, column=1).comment = Comment(
        "Ready only when every one of the 10 required types is fully satisfied. "
        "At Risk when 80%+ of the required deliverables are approved but at least one "
        "type still has an outstanding document.", "Workbook")
    ws.cell(row=ATS_STATUS + 2, column=1, value=(
        "Types Met") ).font = f_lbl
    tm = ws.cell(row=ATS_STATUS + 2, column=2,
                 value=f'=SUM(D{ATS_TOP}:D{ATS_BOT})&" of "&COUNTA(A{ATS_TOP}:A{ATS_BOT})')
    tm.font = f_lbl
    tm.border = box

    # --- Chart data ---------------------------------------------------------
    ws.cell(row=CD_TITLE, column=1,
            value="CHART DATA — generated automatically, do not edit").font = f_lbl

    # 1. Readiness by Scope (top 20 scopes)
    header_row(ws, CD_HDR, 1, ["Scope", "Readiness %"])
    for i in range(20):
        r, sr = CD_TOP + i, FIRST + i
        ws.cell(row=r, column=1,
                value=f"=IF('Scope Tracker'!$A{sr}=\"\",\"\",'Scope Tracker'!$B{sr})").font = f_calcs
        c = ws.cell(row=r, column=2,
                    value=f"=IF('Scope Tracker'!$A{sr}=\"\",\"\",'Scope Tracker'!$J{sr})")
        c.font = f_calcs
        c.number_format = FMT_PCT

    # 2. Equipment status distribution
    header_row(ws, CD_HDR, 4, ["Equipment Status", "Count"])
    for i, s in enumerate(["Not Started", "In Progress", "Behind Schedule", "Complete"]):
        r = CD_TOP + i
        ws.cell(row=r, column=4, value=s).font = f_calcs
        ws.cell(row=r, column=5, value=f"=COUNTIF(Eq_Status,$D{r})").font = f_calcs

    # 3. Document status distribution
    header_row(ws, CD_HDR, 7, ["Document Status", "Count"])
    for i, (name, _, _) in enumerate(WORKFLOW):
        r = CD_TOP + i
        ws.cell(row=r, column=7, value=name).font = f_calcs
        ws.cell(row=r, column=8, value=f"=COUNTIF(Doc_Workflow,$G{r})").font = f_calcs

    # 4. Top 10 overdue documents
    header_row(ws, CD_HDR, 10, ["Document", "Days Late", "key"])
    for i in range(10):
        r = CD_TOP + i
        ws.cell(row=r, column=12,
                value=f"=IFERROR(LARGE(Doc_RankKey,{i + 1}),\"\")").font = f_calcs
        ws.cell(row=r, column=10, value=(
            f'=IF($L{r}="","",IFERROR(INDEX(Doc_Number,MATCH($L{r},Doc_RankKey,0))'
            f'&" — "&INDEX(Doc_Type,MATCH($L{r},Doc_RankKey,0)),""))')).font = f_calcs
        ws.cell(row=r, column=11,
                value=f'=IF($L{r}="","",ROUNDDOWN($L{r},0))').font = f_calcs

    # 5. Top 10 overdue actions
    header_row(ws, CD_HDR, 14, ["Action", "Days Overdue", "key"])
    for i in range(10):
        r = CD_TOP + i
        ws.cell(row=r, column=16,
                value=f"=IFERROR(LARGE(Act_RankKey,{i + 1}),\"\")").font = f_calcs
        ws.cell(row=r, column=14, value=(
            f'=IF($P{r}="","",IFERROR(INDEX(Act_ID,MATCH($P{r},Act_RankKey,0))'
            f'&" — "&LEFT(INDEX(Act_Desc,MATCH($P{r},Act_RankKey,0)),40),""))')).font = f_calcs
        ws.cell(row=r, column=15,
                value=f'=IF($P{r}="","",ROUNDDOWN($P{r},0))').font = f_calcs

    # 6. Workload by owner
    header_row(ws, CD_HDR, 18, ["Owner", "Open Items", "Open Blockers"])
    for i, p in enumerate(PERSONNEL):
        r = CD_TOP + i
        ws.cell(row=r, column=18, value=f"=Configuration!$L${LIST_TOP + i}").font = f_calcs
        ws.cell(row=r, column=19, value=(
            f'=COUNTIFS(Act_Owner,$R{r},Act_Status,"<>Closed")'
            f'+COUNTIFS(Doc_Owner,$R{r},Doc_Approved,0)')).font = f_calcs
        ws.cell(row=r, column=20, value=(
            f'=COUNTIFS(Act_Owner,$R{r},Act_Status,"<>Closed",Act_Blocker,"Yes")')).font = f_calcs

    # 7. Milestone timeline
    header_row(ws, CD_HDR, 22, ["Milestone", "Days From Today"])
    for i in range(len(MILESTONES)):
        r, mr = CD_TOP + i, MS_FIRST + i
        ws.cell(row=r, column=22, value=f"=IF(Dashboard!$B{mr}=\"\",\"\",Dashboard!$B{mr})").font = f_calcs
        ws.cell(row=r, column=23, value=f"=IF(Dashboard!$I{mr}=\"\",0,Dashboard!$I{mr})").font = f_calcs

    ws.sheet_state = "hidden"


# ===========================================================================
# CONTROL PANEL
# ===========================================================================
def build_control_panel(ws, demo, planned_start):
    ws.sheet_properties.tabColor = NAVY
    ws.sheet_view.showGridLines = False
    # Column E doubles as the spacer between the two top blocks and as the
    # "Target Submission" date column of the lead-time tables below, so it has
    # to be wide enough for a DD-MMM-YYYY date.
    widths(ws, {"A": 2, "B": 36, "C": 22, "D": 20, "E": 17, "F": 34,
                "G": 20, "H": 18, "I": 3})

    ws["B1"] = "DMIC COMMISSIONING CONTRACT — CONTROL PANEL"
    ws["B1"].font = f_title
    ws["B2"] = ("Master page. Blue cells are for you to type in; grey cells calculate themselves. "
                "The amber cell (Planned Commissioning Start) drives every target date in the workbook.")
    ws["B2"].font = f_sub

    # --- Contract information ----------------------------------------------
    section(ws, 4, 2, 4, "CONTRACT INFORMATION")
    info_fields = ["Contract Name", "Contract Manager", "Project Manager",
                   "Commissioning Manager", "Lead Commissioning Engineer",
                   "Quality Lead", "Contract Status"]
    demo_info = ["DMIC Package 4 — Systems Commissioning", "A. Mercer", "S. Okonkwo",
                 "L. Fontaine", "D. Petrov", "M. Haddad", "In Progress"]
    for i, name in enumerate(info_fields):
        r = CP_INFO_TOP + i
        label(ws, r, 2, name)
        input_cell(ws, r, 3, demo_info[i] if demo else None)
    ws.merge_cells(start_row=CP_INFO_TOP, start_column=3,
                   end_row=CP_INFO_TOP, end_column=4)

    # --- Commissioning dates ------------------------------------------------
    section(ws, 13, 2, 4, "COMMISSIONING DATES")
    date_fields = ["Planned Commissioning Start", "Planned Commissioning Finish",
                   "Forecast Commissioning Start", "Forecast Commissioning Finish",
                   "Actual Commissioning Start", "Actual Commissioning Finish"]
    if demo:
        demo_dates = [planned_start,
                      planned_start + dt.timedelta(days=180),
                      planned_start + dt.timedelta(days=21),
                      planned_start + dt.timedelta(days=205),
                      None, None]
    else:
        demo_dates = [planned_start, None, None, None, None, None]
    for i, name in enumerate(date_fields):
        r = CP_DATE_TOP + i
        label(ws, r, 2, name)
        input_cell(ws, r, 3, demo_dates[i], FMT_DATE, key=(i == 0))
    ws[f"D{CP_DATE_TOP}"] = "◀ drives all target dates"
    ws[f"D{CP_DATE_TOP}"].font = f_note

    section(ws, 21, 2, 4, "SCHEDULE PERFORMANCE")
    label(ws, 22, 2, "Delay vs Baseline (days)")
    calc_cell(ws, 22, 3,
              f'=IF(OR($C${CP_DATE_TOP}="",$C${CP_DATE_TOP + 2}=""),"",'
              f'$C${CP_DATE_TOP + 2}-$C${CP_DATE_TOP})', FMT_DAYS,
              note="Forecast Commissioning Start minus Planned Commissioning Start. "
                   "Positive = late against baseline.")
    label(ws, 23, 2, "Days To Commissioning Start")
    calc_cell(ws, 23, 3,
              f'=IF($C${CP_DATE_TOP}="","",$C${CP_DATE_TOP}-TODAY())', FMT_DAYS)

    # --- Contract readiness summary ----------------------------------------
    section(ws, 4, 6, 8, "CONTRACT READINESS SUMMARY")
    summary = [
        ("Overall Contract Readiness %", "=IFERROR(AVERAGE(Eq_Readiness),0)", FMT_PCT,
         "Average blended readiness across every equipment item "
         "(phase progress + approved documents)."),
        ("Number of Scopes", "=COUNTA(Sc_ID)", FMT_INT, None),
        ("Number of Equipment Items", "=COUNTA(Eq_Name)", FMT_INT, None),
        ("Open Actions", '=COUNTA(Act_ID)-COUNTIF(Act_Status,"Closed")', FMT_INT, None),
        ("Overdue Actions", "=SUM(Act_Overdue)", FMT_INT,
         "Actions past their Due Date with no Completion Date and status not Closed."),
        ("Missing Documents", '=COUNTIF(Doc_Workflow,"Not Started")', FMT_INT,
         "Assumption: the spec does not define 'missing'. Counted here as deliverables that "
         "exist as a row but whose Workflow Status is still 'Not Started'."),
        ("Open Risks", '=COUNTIFS(Act_Blocker,"Yes",Act_Status,"<>Closed")', FMT_INT,
         "Assumption: the spec has no separate Risk sheet. Counted here as open actions "
         "flagged as a Commissioning Blocker."),
        ("Upcoming Milestones", '=COUNTIFS(Ms_Days,">=0",Ms_Days,"<="&Cfg_MsWindow)', FMT_INT,
         "Milestones whose forecast (or planned) date falls within the window set on the "
         "Configuration sheet, and which are not yet Achieved."),
    ]
    for i, (name, formula, fmt, note) in enumerate(summary):
        r = 5 + i
        label(ws, r, 6, name)
        calc_cell(ws, r, 7, formula, fmt, note=note)

    # --- RAG status ---------------------------------------------------------
    section(ws, 14, 6, 8, "RAG STATUS")
    label(ws, 15, 6, "Contract Health Index")
    calc_cell(ws, 15, 7, "=Dashboard!$L$12", FMT_PCT)
    label(ws, 16, 6, "Health Category")
    calc_cell(ws, 16, 7,
              '=IF($G$15>=Cfg_ThrExcellent,"Excellent",IF($G$15>=Cfg_ThrGood,"Good",'
              'IF($G$15>=Cfg_ThrAttention,"Attention Required","Critical")))')
    label(ws, 17, 6, "RAG")
    rag = calc_cell(ws, 17, 7,
                    '=IF($G$15>=Cfg_RagGreen,"GREEN",IF($G$15>=Cfg_RagAmber,"AMBER","RED"))')
    rag.font = Font(name=FONT, size=12, bold=True)
    rag.alignment = c_center

    section(ws, 19, 6, 8, "AUTHORITY TO START (ATS) READINESS GATE")
    label(ws, 20, 6, "ATS Readiness Score")
    calc_cell(ws, 20, 7, f"=Configuration!$B${ATS_SCORE}", FMT_PCT)
    label(ws, 21, 6, "ATS Status")
    ats = calc_cell(ws, 21, 7, f"=Configuration!$B${ATS_STATUS}")
    ats.font = Font(name=FONT, size=12, bold=True)
    ats.alignment = c_center
    label(ws, 22, 6, "Required Types Met")
    calc_cell(ws, 22, 7, f"=Configuration!$B${ATS_STATUS + 2}")
    ws["F23"] = "'Ready' requires all 10 minimum deliverable types fully approved."
    ws["F23"].font = f_note

    # --- Legend -------------------------------------------------------------
    section(ws, 24, 2, 8, "LEGEND")
    legend = [
        ("Amber cell", "The single key driver — Planned Commissioning Start.", KEY_FILL),
        ("Blue text on white", "Type here. These are your inputs.", "FFFFFF"),
        ("Grey cell", "Calculated automatically — do not type over it.", GREY_FILL),
    ]
    for i, (what, why, colour) in enumerate(legend):
        r = 25 + i
        c = ws.cell(row=r, column=2, value=what)
        c.fill = PatternFill("solid", fgColor=colour)
        c.font = f_body
        c.border = box
        d = ws.cell(row=r, column=3, value=why)
        d.font = f_body
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)

    # --- Lead time configuration: documents ---------------------------------
    section(ws, 29, 2, 6, "LEAD TIME CONFIGURATION — DOCUMENTS")
    ws["B30"] = ("Target dates are generated as: Planned Commissioning Start − Lead Days. "
                 "Edit the blue Lead Days column and every target date in the Document Tracker updates. "
                 "A negative Lead Day value places the target AFTER commissioning start.")
    ws["B30"].font = f_sub
    header_row(ws, CP_LEAD_HDR, 2,
               ["Document Type", "Spec Lead Time", "Lead Days (T minus)",
                "Target Submission", "Target Approval"])
    for i, (name, spec_label, days) in enumerate(DOC_TYPES):
        r = CP_LEAD_TOP + i
        c = ws.cell(row=r, column=2, value=name)
        c.font = f_body
        c.border = box
        c2 = ws.cell(row=r, column=3, value=spec_label)
        c2.font = f_note
        c2.border = box
        input_cell(ws, r, 4, days, FMT_INT)
        # Target Submission (E) = Target Approval (F) - review days
        calc_cell(ws, r, 5, f'=IF($F{r}="","",$F{r}-Cfg_ReviewDays)', FMT_DATE)
        calc_cell(ws, r, 6,
                  f'=IF(OR($C${CP_DATE_TOP}="",$D{r}=""),"",$C${CP_DATE_TOP}-$D{r})', FMT_DATE)
    ws.cell(row=CP_LEAD_BOT + 1, column=2, value=(
        "Months are converted at 30 days each: T-12 Months = 360, T-9 = 270, T-6 = 180, "
        "T-3 = 90, T-1 Month = 30, T-2 Weeks = 14. This reproduces the worked example in the "
        "specification exactly (Planned Start 01-Jan-2028, Commissioning Procedure lead 90 days "
        "→ Target Approval 03-Oct-2027). The spec's post-commissioning deliverables "
        "(During Commissioning, Before FTC/TFM/TFTO, Before Final Takeover) have no stated "
        "duration and are seeded as T+30 to T+120; adjust to your programme. "
        "Source: DMIC Commissioning Contract Management Workbook v3, 'Default Lead Times' table "
        "and 'Automatic Document Deadline Calculation' example."
    )).font = f_note

    # --- Lead time configuration: milestones --------------------------------
    section(ws, 53, 2, 6, "LEAD TIME CONFIGURATION — MILESTONES")
    header_row(ws, CP_MS_HDR, 2,
               ["Milestone", "Spec Reference", "Lead Days (T minus)", "Target Date"])
    for i, (name, days) in enumerate(MILESTONES):
        r = CP_MS_TOP + i
        c = ws.cell(row=r, column=2, value=name)
        c.font = f_body
        c.border = box
        c2 = ws.cell(row=r, column=3, value="Milestone Tracker")
        c2.font = f_note
        c2.border = box
        input_cell(ws, r, 4, days, FMT_INT)
        calc_cell(ws, r, 5,
                  f'=IF(OR($C${CP_DATE_TOP}="",$D{r}=""),"",$C${CP_DATE_TOP}-$D{r})', FMT_DATE)
    ws.cell(row=CP_MS_BOT + 1, column=2, value=(
        "Assumption: only Kick Off (T-12M), Authority To Start (T-2W) and Commissioning Start (T) "
        "are given in the spec. The remaining milestone offsets are seeded estimates — adjust them."
    )).font = f_note

    # Data validation
    dv = DataValidation(type="list", formula1="Cfg_ContractStatus", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"C{CP_INFO_TOP + 6}")

    dvp = DataValidation(type="list", formula1="Cfg_Personnel", allow_blank=True)
    ws.add_data_validation(dvp)
    dvp.add(f"C{CP_INFO_TOP + 1}:C{CP_INFO_TOP + 5}")

    ws.freeze_panes = "A4"


# ===========================================================================
# SCOPE TRACKER
# ===========================================================================
SCOPE_COLS = ["Scope ID", "Scope Name", "Scope Owner", "Planned Start", "Planned Finish",
              "Forecast Start", "Forecast Finish", "Actual Start", "Actual Finish",
              "Readiness %", "Status", "Number of Equipment", "Equipment Complete",
              "Equipment In Progress", "Equipment Not Started", "Open Actions",
              "Outstanding Documents", "Comments"]


def build_scope_tracker(ws):
    ws.sheet_properties.tabColor = BLUE
    ws["A1"] = "SCOPE TRACKER"
    ws["A1"].font = f_title
    ws["A2"] = ("One row per Commissioning Scope. White columns are inputs; grey columns "
                "roll up automatically from the Equipment, Document and Action trackers.")
    ws["A2"].font = f_sub
    header_row(ws, 3, 1, SCOPE_COLS)

    style_input_range(ws, 1, 9, FIRST, SCOPE_LAST,
                      fmt_map={c: FMT_DATE for c in range(4, 10)})
    style_input_range(ws, 11, 11, FIRST, SCOPE_LAST)
    style_input_range(ws, 18, 18, FIRST, SCOPE_LAST)

    fill_down(ws, 10, FIRST, SCOPE_LAST,
              '=IF($A{r}="","",IFERROR(AVERAGEIF(Eq_Scope,$A{r},Eq_Readiness),0))', FMT_PCT)
    fill_down(ws, 12, FIRST, SCOPE_LAST,
              '=IF($A{r}="","",COUNTIF(Eq_Scope,$A{r}))', FMT_INT)
    fill_down(ws, 13, FIRST, SCOPE_LAST,
              '=IF($A{r}="","",COUNTIFS(Eq_Scope,$A{r},Eq_Phase,"Complete"))', FMT_INT)
    fill_down(ws, 14, FIRST, SCOPE_LAST,
              '=IF($A{r}="","",$L{r}-$M{r}-$O{r})', FMT_INT)
    fill_down(ws, 15, FIRST, SCOPE_LAST,
              '=IF($A{r}="","",COUNTIFS(Eq_Scope,$A{r},Eq_Phase,"Pre-Commissioning"))', FMT_INT)
    fill_down(ws, 16, FIRST, SCOPE_LAST,
              '=IF($A{r}="","",COUNTIFS(Act_Scope,$A{r},Act_Status,"<>Closed"))', FMT_INT)
    fill_down(ws, 17, FIRST, SCOPE_LAST,
              '=IF($A{r}="","",COUNTIFS(Doc_Scope,$A{r},Doc_Approved,0))', FMT_INT)

    widths(ws, {"A": 12, "B": 34, "C": 16, "D": 14, "E": 14, "F": 14, "G": 14,
                "H": 14, "I": 14, "J": 12, "K": 15, "L": 11, "M": 11, "N": 11,
                "O": 11, "P": 11, "Q": 13, "R": 40})

    dv = DataValidation(type="list", formula1="Cfg_ScopeStatus", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"K{FIRST}:K{SCOPE_LAST}")
    dvo = DataValidation(type="list", formula1="Cfg_Personnel", allow_blank=True)
    ws.add_data_validation(dvo)
    dvo.add(f"C{FIRST}:C{SCOPE_LAST}")

    ws.conditional_formatting.add(
        f"J{FIRST}:J{SCOPE_LAST}",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                    color="638EC6", showValue=True))
    ws.auto_filter.ref = f"A3:R{SCOPE_LAST}"
    ws.freeze_panes = "C4"


# ===========================================================================
# EQUIPMENT TRACKER
# ===========================================================================
EQUIP_COLS = ["Equipment Name", "Parent Scope", "Responsible Engineer",
              "Commissioning Phase", "Planned Start", "Forecast Start", "Actual Start",
              "Planned Finish", "Forecast Finish", "Actual Finish", "Readiness %",
              "Status", "Comments"]


def build_equipment_tracker(ws):
    ws.sheet_properties.tabColor = BLUE
    ws["A1"] = "EQUIPMENT TRACKER"
    ws["A1"].font = f_title
    ws["A2"] = ("One row per Equipment Item. Each item progresses independently; Scope and "
                "Contract readiness roll up from here.")
    ws["A2"].font = f_sub
    header_row(ws, 3, 1, EQUIP_COLS)
    header_row(ws, 3, 14, ["Linked Docs", "Phase %", "Doc %"])

    style_input_range(ws, 1, 10, FIRST, EQUIP_LAST,
                      fmt_map={c: FMT_DATE for c in range(5, 11)})
    style_input_range(ws, 13, 13, FIRST, EQUIP_LAST)

    fill_down(ws, 14, FIRST, EQUIP_LAST,
              '=IF($A{r}="","",COUNTIF(Doc_Equipment,$A{r}))', FMT_INT)
    fill_down(ws, 15, FIRST, EQUIP_LAST,
              '=IF($A{r}="","",IFERROR(INDEX(Cfg_PhasePct,MATCH($D{r},Cfg_PhaseList,0)),0))',
              FMT_PCT)
    fill_down(ws, 16, FIRST, EQUIP_LAST,
              '=IF($A{r}="","",IF($N{r}=0,"",'
              'COUNTIFS(Doc_Equipment,$A{r},Doc_Approved,1)/$N{r}))', FMT_PCT)
    fill_down(ws, 11, FIRST, EQUIP_LAST,
              '=IF($A{r}="","",IF($N{r}=0,$O{r},Cfg_WPhase*$O{r}+Cfg_WDoc*$P{r}))', FMT_PCT)
    fill_down(ws, 12, FIRST, EQUIP_LAST,
              '=IF($A{r}="","",IF($K{r}>=1,"Complete",IF($K{r}<=0,"Not Started",'
              'IF(AND($H{r}<>"",$H{r}<TODAY(),$J{r}=""),"Behind Schedule","In Progress"))))')

    ws.cell(row=3, column=11).comment = Comment(
        "Blended readiness: (phase weight x phase progress) + (document weight x share of "
        "linked documents approved). Weights are on the Configuration sheet. Equipment with "
        "no linked documents falls back to phase progress alone.", "Workbook")

    widths(ws, {"A": 30, "B": 12, "C": 18, "D": 24, "E": 14, "F": 14, "G": 14,
                "H": 14, "I": 14, "J": 14, "K": 12, "L": 17, "M": 36,
                "N": 11, "O": 10, "P": 10})
    for col in ("N", "O", "P"):
        ws.column_dimensions[col].hidden = True

    dvp = DataValidation(type="list", formula1="Cfg_Phase", allow_blank=True)
    ws.add_data_validation(dvp)
    dvp.add(f"D{FIRST}:D{EQUIP_LAST}")
    dvs = DataValidation(type="list", formula1="Sc_IDList", allow_blank=True)
    ws.add_data_validation(dvs)
    dvs.add(f"B{FIRST}:B{EQUIP_LAST}")
    dvo = DataValidation(type="list", formula1="Cfg_Personnel", allow_blank=True)
    ws.add_data_validation(dvo)
    dvo.add(f"C{FIRST}:C{EQUIP_LAST}")

    ws.conditional_formatting.add(
        f"K{FIRST}:K{EQUIP_LAST}",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                    color="63BE7B", showValue=True))
    ws.conditional_formatting.add(
        f"A{FIRST}:M{EQUIP_LAST}",
        FormulaRule(formula=[f'AND($A{FIRST}<>"",$L{FIRST}="Behind Schedule")'],
                    fill=PatternFill("solid", fgColor=DOC_COLOURS["red"]), stopIfTrue=False))
    ws.auto_filter.ref = f"A3:M{EQUIP_LAST}"
    ws.freeze_panes = "C4"


# ===========================================================================
# DOCUMENT TRACKER
# ===========================================================================
DOC_COLS = ["Document Type", "Scope", "Equipment (optional)", "Document Number",
            "Document Title", "Revision", "Owner", "Reviewer", "Workflow Status",
            "Readiness Gate", "Target Submission Date", "Target Approval Date",
            "Actual Submission Date", "Actual Approval Date", "Days Late",
            "Criticality", "Comments"]


def build_document_tracker(ws):
    ws.sheet_properties.tabColor = "C00000"
    ws["A1"] = "DOCUMENT TRACKER"
    ws["A1"].font = f_title
    ws["A2"] = ("Primary operational sheet — one row per commissioning deliverable. Target dates "
                "are generated from the Planned Commissioning Start Date and the lead times on "
                "the Control Panel.")
    ws["A2"].font = f_sub
    header_row(ws, 3, 1, DOC_COLS)
    header_row(ws, 3, 19, ["Approved", "Progress %", "Days Remaining", "Flag", "Rank Key"])

    style_input_range(ws, 1, 10, FIRST, DOC_LAST)
    style_input_range(ws, 13, 14, FIRST, DOC_LAST, fmt_map={13: FMT_DATE, 14: FMT_DATE})
    style_input_range(ws, 16, 17, FIRST, DOC_LAST)

    fill_down(ws, 12, FIRST, DOC_LAST,
              '=IF($A{r}="","",IFERROR(Ctrl_PlannedStart-'
              'INDEX(Cfg_LeadDays,MATCH($A{r},Cfg_LeadType,0)),""))', FMT_DATE)
    fill_down(ws, 11, FIRST, DOC_LAST,
              '=IF($L{r}="","",$L{r}-Cfg_ReviewDays)', FMT_DATE)
    fill_down(ws, 15, FIRST, DOC_LAST,
              '=IF($A{r}="","",IF($N{r}<>"",MAX(0,$N{r}-$L{r}),'
              'IF(AND($L{r}<>"",TODAY()>$L{r}),TODAY()-$L{r},0)))', FMT_DAYS)
    fill_down(ws, 19, FIRST, DOC_LAST,
              '=IF($A{r}="","",IFERROR(INDEX(Cfg_WfApproved,MATCH($I{r},Cfg_WfList,0)),0))',
              FMT_INT)
    fill_down(ws, 20, FIRST, DOC_LAST,
              '=IF($A{r}="","",IFERROR(INDEX(Cfg_WfPct,MATCH($I{r},Cfg_WfList,0)),0))', FMT_PCT)
    fill_down(ws, 21, FIRST, DOC_LAST,
              '=IF(OR($A{r}="",$L{r}=""),"",$L{r}-TODAY())', FMT_DAYS)
    # Overdue outranks Not Started: a deliverable nobody has started is still late once
    # its target approval date has passed. "Missing" is counted from Workflow Status.
    fill_down(ws, 22, FIRST, DOC_LAST,
              '=IF($A{r}="","",IF($S{r}=1,"Approved",'
              'IF(AND($L{r}<>"",TODAY()>$L{r}),"Overdue",'
              'IF(AND($L{r}<>"",$L{r}-TODAY()<=Cfg_DocWindow),"Due Soon",'
              'IF($I{r}="Not Started","Not Started","On Track")))))')
    fill_down(ws, 23, FIRST, DOC_LAST,
              '=IF($A{r}="","",IF(AND($S{r}=0,$O{r}>0),$O{r}+ROW()/1000000,""))')

    ws.cell(row=3, column=15).comment = Comment(
        "Days late against the Target Approval Date. Once an Actual Approval Date is entered "
        "it freezes at how late the approval actually was.", "Workbook")

    widths(ws, {"A": 34, "B": 10, "C": 26, "D": 22, "E": 34, "F": 9, "G": 15, "H": 15,
                "I": 18, "J": 19, "K": 16, "L": 16, "M": 16, "N": 16, "O": 11,
                "P": 12, "Q": 34, "S": 10, "T": 11, "U": 13, "V": 13, "W": 12})
    for col in ("S", "T", "U", "V", "W"):
        ws.column_dimensions[col].hidden = True

    for formula, target in [
        ("Cfg_DocType", "A"), ("Sc_IDList", "B"), ("Eq_NameList", "C"),
        ("Cfg_WorkflowStatus", "I"), ("Cfg_Gate", "J"), ("Cfg_Criticality", "P"),
    ]:
        dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{target}{FIRST}:{target}{DOC_LAST}")
    for target in ("G", "H"):
        dv = DataValidation(type="list", formula1="Cfg_Personnel", allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{target}{FIRST}:{target}{DOC_LAST}")

    # Spec: Document Colour Coding, applied to the Workflow Status column
    rng = f"I{FIRST}:I{DOC_LAST}"
    colour_rules = [
        (f'$I{FIRST}="Not Started"', "grey"),
        (f'$I{FIRST}="Draft"', "blue"),
        (f'OR($I{FIRST}="Internal Review",$I{FIRST}="Submitted",'
         f'$I{FIRST}="Under Review",$I{FIRST}="Resubmitted")', "yellow"),
        (f'$I{FIRST}="Comments Received"', "orange"),
        (f'$I{FIRST}="Rework"', "red"),
        (f'OR($I{FIRST}="Approved",$I{FIRST}="Released",$I{FIRST}="Implemented")', "green"),
        (f'$I{FIRST}="Closed"', "dgreen"),
    ]
    for formula, colour in colour_rules:
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[formula], fill=PatternFill("solid", fgColor=DOC_COLOURS[colour]),
            stopIfTrue=True))

    # Overdue / due-soon row highlighting
    ws.conditional_formatting.add(f"A{FIRST}:Q{DOC_LAST}", FormulaRule(
        formula=[f'AND($A{FIRST}<>"",$V{FIRST}="Overdue")'],
        fill=PatternFill("solid", fgColor=DOC_COLOURS["red"]), stopIfTrue=True))
    ws.conditional_formatting.add(f"A{FIRST}:Q{DOC_LAST}", FormulaRule(
        formula=[f'AND($A{FIRST}<>"",$V{FIRST}="Due Soon")'],
        fill=PatternFill("solid", fgColor=DOC_COLOURS["yellow"]), stopIfTrue=True))

    ws.auto_filter.ref = f"A3:Q{DOC_LAST}"
    ws.freeze_panes = "E4"


# ===========================================================================
# ACTION TRACKER
# ===========================================================================
ACT_COLS = ["Action ID", "Action Description", "Category", "Linked Scope",
            "Linked Equipment", "Linked Document", "Owner", "Support Owner",
            "Date Raised", "Due Date", "Completion Date", "Priority",
            "Commissioning Blocker", "Status", "Comments"]


def build_action_tracker(ws):
    ws.sheet_properties.tabColor = "BF8F00"
    ws["A1"] = "ACTION TRACKER"
    ws["A1"].font = f_title
    ws["A2"] = ("One row per Action. Rows colour themselves by priority and lateness — "
                "see the Read Me sheet for the colour key.")
    ws["A2"].font = f_sub
    header_row(ws, 3, 1, ACT_COLS)
    header_row(ws, 3, 17, ["Days To Due", "Overdue", "Rank Key"])

    style_input_range(ws, 1, 15, FIRST, ACT_LAST,
                      fmt_map={9: FMT_DATE, 10: FMT_DATE, 11: FMT_DATE})

    fill_down(ws, 17, FIRST, ACT_LAST,
              '=IF($A{r}="","",IF(OR($K{r}<>"",$J{r}=""),"",$J{r}-TODAY()))', FMT_DAYS)
    fill_down(ws, 18, FIRST, ACT_LAST,
              '=IF($A{r}="",0,IF(AND($N{r}<>"Closed",$K{r}="",$J{r}<>"",$J{r}<TODAY()),1,0))',
              FMT_INT)
    fill_down(ws, 19, FIRST, ACT_LAST,
              '=IF($A{r}="","",IF($R{r}=1,(TODAY()-$J{r})+ROW()/1000000,""))')

    widths(ws, {"A": 11, "B": 46, "C": 14, "D": 12, "E": 26, "F": 22, "G": 15,
                "H": 15, "I": 13, "J": 13, "K": 14, "L": 11, "M": 20, "N": 13,
                "O": 34, "Q": 12, "R": 10, "S": 12})
    for col in ("Q", "R", "S"):
        ws.column_dimensions[col].hidden = True

    for formula, target in [
        ("Cfg_ActionCategory", "C"), ("Sc_IDList", "D"), ("Eq_NameList", "E"),
        ("Doc_NumberList", "F"), ("Cfg_Priority", "L"), ("Cfg_YesNo", "M"),
        ("Cfg_ActionStatus", "N"),
    ]:
        dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{target}{FIRST}:{target}{ACT_LAST}")
    for target in ("G", "H"):
        dv = DataValidation(type="list", formula1="Cfg_Personnel", allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{target}{FIRST}:{target}{ACT_LAST}")

    # Spec: Action Colour Rules, most severe first
    rng = f"A{FIRST}:O{ACT_LAST}"
    rules = [
        (f'AND($A{FIRST}<>"",$L{FIRST}="Critical",$R{FIRST}=1)', "darkred", True),
        (f'AND($A{FIRST}<>"",$R{FIRST}=1)', "red", False),
        (f'AND($A{FIRST}<>"",$N{FIRST}<>"Closed",$Q{FIRST}>=0,$Q{FIRST}<=Cfg_ActWindow)',
         "amber", False),
        (f'AND($A{FIRST}<>"",$N{FIRST}="In Progress")', "blue", False),
        (f'AND($A{FIRST}<>"",$N{FIRST}="Closed")', "green", False),
        (f'AND($A{FIRST}<>"",$N{FIRST}="Waiting")', "grey", False),
    ]
    for formula, colour, white in rules:
        font = Font(name=FONT, size=9, bold=True, color="FFFFFF") if white else None
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[formula], fill=PatternFill("solid", fgColor=ACT_COLOURS[colour]),
            font=font, stopIfTrue=True))

    ws.auto_filter.ref = f"A3:O{ACT_LAST}"
    ws.freeze_panes = "C4"


# ===========================================================================
# DASHBOARD & MILESTONES
# ===========================================================================
QUESTIONS = [
    ("What is late?",
     '=COUNTIF(Doc_Flag,"Overdue")&" documents overdue, "&SUM(Act_Overdue)&" actions overdue"'),
    ("What is missing?",
     '=COUNTIF(Doc_Workflow,"Not Started")&" deliverables not started"'),
    ("What is blocking commissioning?",
     '=COUNTIFS(Act_Blocker,"Yes",Act_Status,"<>Closed")&" open commissioning blockers"'),
    ("Who owns the blocker?",
     '=IF(MAX(Cfg_OwnerBlockers)=0,"No open blockers",'
     'IFERROR(INDEX(Cfg_OwnerList,MATCH(MAX(Cfg_OwnerBlockers),Cfg_OwnerBlockers,0))'
     '&" ("&MAX(Cfg_OwnerBlockers)&")","—"))'),
    ("Which scopes are not ready?",
     '=COUNTIFS(Sc_ID,"<>",Sc_Readiness,"<"&Cfg_ScopeReady)&" of "&COUNTA(Sc_ID)'
     '&" scopes below the ready threshold"'),
    ("Which equipment items are behind schedule?",
     '=COUNTIF(Eq_Status,"Behind Schedule")&" equipment items past planned finish"'),
    ("Which documents are not approved?",
     '=COUNTIF(Doc_Approved,0)&" of "&COUNTA(Doc_Type)&" deliverables not yet approved"'),
    ("Which milestones are at risk?",
     '=COUNTIF(Ms_Status,"At Risk")+COUNTIF(Ms_Status,"Delayed")&" milestones at risk or delayed"'),
    ("Which actions are due this week?",
     '=COUNTIFS(Act_DaysToDue,">=0",Act_DaysToDue,"<=7")&" actions due within 7 days"'),
]

HEALTH_FORMULAS = [
    "=IFERROR(AVERAGE(Doc_Progress),0)",
    "=IFERROR(AVERAGE(Eq_Readiness),0)",
    '=IF(COUNTA(Act_ID)=0,1,1-SUM(Act_Overdue)/COUNTA(Act_ID))',
    '=IF(COUNTA(Ms_Name)=0,1,1-(COUNTIF(Ms_Status,"At Risk")+COUNTIF(Ms_Status,"Delayed"))'
    '/COUNTA(Ms_Name))',
    "=IFERROR(AVERAGE(Eq_PhasePct),0)",
    '=IF((COUNTIF(Doc_Gate,"Handover")+COUNTIF(Doc_Gate,"Final Takeover"))=0,0,'
    '(COUNTIFS(Doc_Gate,"Handover",Doc_Approved,1)'
    '+COUNTIFS(Doc_Gate,"Final Takeover",Doc_Approved,1))'
    '/(COUNTIF(Doc_Gate,"Handover")+COUNTIF(Doc_Gate,"Final Takeover")))',
]


def build_dashboard(ws, demo):
    ws.sheet_properties.tabColor = "375623"
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 2, "B": 38, "C": 34, "D": 15, "E": 15, "F": 15, "G": 17,
                "H": 16, "I": 12, "J": 17, "K": 20, "L": 12, "M": 4,
                "N": 12, "O": 12, "P": 12})

    ws["B1"] = "EXECUTIVE DASHBOARD & MILESTONES"
    ws["B1"].font = f_title
    ws["B2"] = "Everything on this sheet is calculated. Nothing here needs typing except the Milestone Tracker and the Readiness Trend log."
    ws["B2"].font = f_sub

    # --- Contract health ----------------------------------------------------
    section(ws, 3, 2, 12, "CONTRACT HEALTH INDEX")
    ws.merge_cells("B5:D8")
    big = ws["B5"]
    big.value = "=$L$12"
    big.font = f_big
    big.number_format = FMT_PCT
    big.alignment = c_center
    big.fill = fill_light
    for r in range(5, 9):
        for c in range(2, 5):
            ws.cell(row=r, column=c).border = box

    label(ws, 5, 6, "Health Category")
    calc_cell(ws, 5, 7,
              '=IF($L$12>=Cfg_ThrExcellent,"Excellent",IF($L$12>=Cfg_ThrGood,"Good",'
              'IF($L$12>=Cfg_ThrAttention,"Attention Required","Critical")))')
    label(ws, 6, 6, "RAG Status")
    calc_cell(ws, 6, 7,
              '=IF($L$12>=Cfg_RagGreen,"GREEN",IF($L$12>=Cfg_RagAmber,"AMBER","RED"))')
    label(ws, 7, 6, "Contract Readiness")
    calc_cell(ws, 7, 7, "=IFERROR(AVERAGE(Eq_Readiness),0)", FMT_PCT)
    label(ws, 8, 6, "ATS Readiness")
    calc_cell(ws, 8, 7, f"=Configuration!$B${ATS_SCORE}", FMT_PCT)

    header_row(ws, 5, 9, ["Component", "Score", "Weight", "Weighted"])
    for i, (name, _) in enumerate(HEALTH_COMPONENTS):
        r = 6 + i
        c = ws.cell(row=r, column=9, value=name)
        c.font = f_body
        c.border = box
        calc_cell(ws, r, 10, HEALTH_FORMULAS[i], FMT_PCT)
        calc_cell(ws, r, 11, f"=Configuration!$L${HW_TOP + i}", FMT_PCT)
        calc_cell(ws, r, 12, f"=$J{r}*$K{r}", FMT_PCT)
    ws.cell(row=12, column=9, value="CONTRACT HEALTH INDEX").font = f_lbl
    tot = ws.cell(row=12, column=12, value="=SUM(L6:L11)")
    tot.font = f_lbl
    tot.number_format = FMT_PCT
    tot.fill = fill_light
    tot.border = box

    # --- Key questions ------------------------------------------------------
    section(ws, 14, 2, 7, "THE DASHBOARD ANSWERS")
    for i, (q, formula) in enumerate(QUESTIONS):
        r = 15 + i
        c = ws.cell(row=r, column=2, value=q)
        c.font = f_lbl
        c.border = box
        a = ws.cell(row=r, column=3, value=formula)
        a.font = f_calc
        a.fill = fill_calc
        a.border = box
        a.alignment = c_left
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)

    # --- Milestone tracker --------------------------------------------------
    section(ws, MS_SECT, 2, 8, "MILESTONE TRACKER")
    header_row(ws, MS_HDR, 2, ["Milestone", "Scope", "Planned Date", "Forecast Date",
                               "Actual Date", "Status", "Owner"])
    ws.cell(row=MS_HDR, column=9, value="Days").font = f_hdr
    ws.cell(row=MS_HDR, column=9).fill = fill_hdr

    for i in range(N_MILESTONE):
        r = MS_FIRST + i
        if i < len(MILESTONES):
            c = ws.cell(row=r, column=2, value=MILESTONES[i][0])
            c.font = f_body
            calc_cell(ws, r, 4,
                      f'=IFERROR(Ctrl_PlannedStart-INDEX(Cfg_MsLeadDays,'
                      f'MATCH($B{r},Cfg_MsList,0)),"")', FMT_DATE)
        else:
            input_cell(ws, r, 4, None, FMT_DATE)
        for col in (2, 3, 5, 6, 7, 8):
            cc = ws.cell(row=r, column=col)
            cc.font = f_body
            cc.border = box
            if col in (5, 6):
                cc.number_format = FMT_DATE
        ws.cell(row=r, column=9, value=(
            f'=IF($B{r}="","",IF($G{r}="Achieved","",'
            f'IF($E{r}<>"",$E{r}-TODAY(),IF($D{r}<>"",$D{r}-TODAY(),""))))')).font = f_calcs
        ws.cell(row=r, column=9).number_format = FMT_DAYS
    ws.column_dimensions["I"].hidden = True

    dv = DataValidation(type="list", formula1="Cfg_MilestoneStatus", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"G{MS_FIRST}:G{MS_LAST}")
    dvs = DataValidation(type="list", formula1="Sc_IDList", allow_blank=True)
    ws.add_data_validation(dvs)
    dvs.add(f"C{MS_FIRST}:C{MS_LAST}")
    dvo = DataValidation(type="list", formula1="Cfg_Personnel", allow_blank=True)
    ws.add_data_validation(dvo)
    dvo.add(f"H{MS_FIRST}:H{MS_LAST}")

    ws.conditional_formatting.add(f"B{MS_FIRST}:H{MS_LAST}", FormulaRule(
        formula=[f'AND($B{MS_FIRST}<>"",OR($G{MS_FIRST}="Delayed",$G{MS_FIRST}="At Risk"))'],
        fill=PatternFill("solid", fgColor=DOC_COLOURS["red"]), stopIfTrue=True))
    ws.conditional_formatting.add(f"B{MS_FIRST}:H{MS_LAST}", FormulaRule(
        formula=[f'AND($B{MS_FIRST}<>"",$G{MS_FIRST}="Achieved")'],
        fill=PatternFill("solid", fgColor=DOC_COLOURS["green"]), stopIfTrue=True))

    # --- Readiness trend log ------------------------------------------------
    section(ws, MS_SECT, 10, 11, "READINESS TREND LOG")
    header_row(ws, MS_HDR, 10, ["Snapshot Date", "Contract Readiness %"])
    ws.cell(row=MS_SECT, column=10).comment = Comment(
        "The workbook holds no history of its own. Record one snapshot per month: type the date "
        "and paste the Contract Readiness value from the Control Panel as a number. "
        "The Readiness Trend chart plots this table.", "Workbook")
    for i in range(N_TREND):
        r = TREND_FIRST + i
        input_cell(ws, r, 10, None, FMT_DATE)
        input_cell(ws, r, 11, None, FMT_PCT)

    section(ws, CHART_SECT, 2, 12, "REQUIRED CHARTS")
    ws.freeze_panes = "A4"


def add_charts(ws, cfg):
    def style(ch, title, h=7.2, w=13.5):
        ch.title = title
        ch.height = h
        ch.width = w
        ch.style = 2
        ch.plotVisOnly = False
        return ch

    top = CD_TOP

    # Readiness by Scope
    c1 = style(BarChart(), "Readiness by Scope")
    c1.type = "col"
    c1.y_axis.numFmt = "0%"
    c1.add_data(Reference(cfg, min_col=2, min_row=CD_HDR, max_row=top + 19),
                titles_from_data=True)
    c1.set_categories(Reference(cfg, min_col=1, min_row=top, max_row=top + 19))
    c1.legend = None
    ws.add_chart(c1, "B60")

    # Equipment status distribution
    c2 = style(DoughnutChart(), "Equipment Status Distribution")
    c2.add_data(Reference(cfg, min_col=5, min_row=CD_HDR, max_row=top + 3),
                titles_from_data=True)
    c2.set_categories(Reference(cfg, min_col=4, min_row=top, max_row=top + 3))
    c2.dataLabels = DataLabelList()
    c2.dataLabels.showVal = True
    ws.add_chart(c2, "K60")

    # Document status distribution
    c3 = style(DoughnutChart(), "Document Status Distribution")
    c3.add_data(Reference(cfg, min_col=8, min_row=CD_HDR, max_row=top + 11),
                titles_from_data=True)
    c3.set_categories(Reference(cfg, min_col=7, min_row=top, max_row=top + 11))
    c3.dataLabels = DataLabelList()
    c3.dataLabels.showVal = True
    ws.add_chart(c3, "B76")

    # Workload by owner
    c4 = style(BarChart(), "Workload by Owner (open documents + open actions)")
    c4.type = "col"
    c4.add_data(Reference(cfg, min_col=19, min_row=CD_HDR, max_row=top + 15),
                titles_from_data=True)
    c4.set_categories(Reference(cfg, min_col=18, min_row=top, max_row=top + 15))
    c4.legend = None
    ws.add_chart(c4, "K76")

    # Top 10 overdue documents
    c5 = style(BarChart(), "Top 10 Overdue Documents (days late)")
    c5.type = "bar"
    c5.add_data(Reference(cfg, min_col=11, min_row=CD_HDR, max_row=top + 9),
                titles_from_data=True)
    c5.set_categories(Reference(cfg, min_col=10, min_row=top, max_row=top + 9))
    c5.legend = None
    ws.add_chart(c5, "B92")

    # Top 10 overdue actions
    c6 = style(BarChart(), "Top 10 Overdue Actions (days overdue)")
    c6.type = "bar"
    c6.add_data(Reference(cfg, min_col=15, min_row=CD_HDR, max_row=top + 9),
                titles_from_data=True)
    c6.set_categories(Reference(cfg, min_col=14, min_row=top, max_row=top + 9))
    c6.legend = None
    ws.add_chart(c6, "K92")

    # Milestone timeline
    c7 = style(BarChart(), "Milestone Timeline (days from today)")
    c7.type = "bar"
    c7.add_data(Reference(cfg, min_col=23, min_row=CD_HDR, max_row=top + len(MILESTONES) - 1),
                titles_from_data=True)
    c7.set_categories(Reference(cfg, min_col=22, min_row=top,
                                max_row=top + len(MILESTONES) - 1))
    c7.legend = None
    ws.add_chart(c7, "B108")

    # Readiness trend
    c8 = style(LineChart(), "Readiness Trend")
    c8.y_axis.numFmt = "0%"
    c8.add_data(Reference(ws, min_col=11, min_row=MS_HDR, max_row=TREND_LAST),
                titles_from_data=True)
    c8.set_categories(Reference(ws, min_col=10, min_row=TREND_FIRST, max_row=TREND_LAST))
    c8.legend = None
    ws.add_chart(c8, "K108")


# ===========================================================================
# READ ME
# ===========================================================================
def build_readme(ws, demo, planned_start):
    ws.sheet_properties.tabColor = "7F7F7F"
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 2, "B": 34, "C": 96})

    ws["B1"] = "DMIC COMMISSIONING CONTRACT MANAGEMENT WORKBOOK"
    ws["B1"].font = f_title
    ws["B2"] = ("Built to the specification 'DMIC Commissioning Contract Management Workbook v3'. "
                + ("DEMONSTRATION COPY — pre-filled with fictional data so every dashboard, KPI "
                   "and chart is alive." if demo else
                   "BLANK TEMPLATE — row 4 of each tracker holds an example; overwrite it with "
                   "your first real record."))
    ws["B2"].font = f_sub

    rows = [
        ("SECTION", "START HERE", None),
        ("1. Set the driver date", "Control Panel → 'Planned Commissioning Start' (amber cell). "
                                   "Every target date in the workbook is generated from it.", None),
        ("2. Tune the lead times", "Control Panel → 'Lead Time Configuration'. Edit the blue "
                                   "Lead Days column; all target dates update.", None),
        ("3. Enter your scopes", "Scope Tracker → one row per commissioning scope. The Scope ID "
                                 "is the key everything else links to.", None),
        ("4. Enter your equipment", "Equipment Tracker → one row per item, each with a Parent "
                                    "Scope and a Commissioning Phase.", None),
        ("5. Enter your deliverables", "Document Tracker → one row per deliverable. Pick a "
                                       "Document Type and the target dates appear.", None),
        ("6. Work the actions", "Action Tracker → rows colour themselves by priority and "
                                "lateness.", None),
        ("7. Read the Dashboard", "Dashboard → Contract Health Index, the nine executive "
                                  "questions, milestones and nine charts.", None),

        ("SECTION", "CELL COLOUR LEGEND", None),
        ("Amber cell", "The single key driver — Planned Commissioning Start.", KEY_FILL),
        ("Blue text, white cell", "Your input. Type here.", "FFFFFF"),
        ("Grey cell", "Calculated. Do not type over it — you would break the roll-up.", GREY_FILL),

        ("SECTION", "DOCUMENT COLOUR CODING (spec)", None),
        ("Grey", "Not Started", DOC_COLOURS["grey"]),
        ("Blue", "Draft", DOC_COLOURS["blue"]),
        ("Yellow", "Internal Review / Submitted / Under Review / Resubmitted", DOC_COLOURS["yellow"]),
        ("Orange", "Comments Received", DOC_COLOURS["orange"]),
        ("Red", "Rework (rejected) — and any overdue row", DOC_COLOURS["red"]),
        ("Green", "Approved / Released / Implemented", DOC_COLOURS["green"]),
        ("Dark Green", "Closed", DOC_COLOURS["dgreen"]),

        ("SECTION", "ACTION COLOUR RULES (spec)", None),
        ("Dark Red", "Critical and overdue", ACT_COLOURS["darkred"]),
        ("Red", "Overdue", ACT_COLOURS["red"]),
        ("Amber", "Due within 14 days", ACT_COLOURS["amber"]),
        ("Blue", "In Progress", ACT_COLOURS["blue"]),
        ("Green", "Closed", ACT_COLOURS["green"]),
        ("Grey", "Waiting", ACT_COLOURS["grey"]),

        ("SECTION", "HOW READINESS IS CALCULATED", None),
        ("Equipment Readiness", "50% x phase progress + 50% x share of that item's linked "
                                "documents approved. Equipment with no linked documents uses "
                                "phase progress alone. Weights: Configuration sheet.", None),
        ("Scope Readiness", "Average equipment readiness across the scope's equipment.", None),
        ("Contract Readiness", "Average equipment readiness across all equipment.", None),
        ("Contract Health Index", "40% Documents + 20% Readiness + 15% Actions + 10% Milestones "
                                  "+ 10% Equipment Progress + 5% Handovers. Weights: "
                                  "Configuration sheet.", None),
        ("ATS Readiness", "Share of the 10 spec-mandated deliverables that are fully satisfied — "
                          "the type exists in the tracker AND none of that type is still "
                          "outstanding. 100% = Ready, 80%+ = At Risk, below = Not Ready.", None),

        ("SECTION", "ASSUMPTIONS WE HAD TO MAKE", None),
        ("Phase progress ramp", "The spec names 11 commissioning phases but gives no % for each. "
                                "Seeded 0% → 100%; retune on the Configuration sheet.", None),
        ("Months → days", "Converted at 30 days per month (T-3 Months = 90), which reproduces "
                          "the spec's own worked example exactly.", None),
        ("Post-commissioning lead times", "'During Commissioning', 'Before FTC/TFM/TFTO' and "
                                          "'Before Final Takeover' have no stated duration. "
                                          "Seeded as T+30 to T+120 days.", None),
        ("'Missing Documents'", "Counted as deliverables that exist as a row but are still "
                                "'Not Started'. The workbook cannot count a row you never "
                                "created.", None),
        ("'Open Risks'", "The spec defines no Risk sheet. Counted as open actions flagged as a "
                         "Commissioning Blocker.", None),
        ("Authority To Start Evidence", "Added as a 20th document type. The spec lists 19 "
                                        "mandatory types but the ATS gate and the Control Panel "
                                        "lead-time list both require it.", None),
        ("Readiness Trend", "Excel keeps no history. Record a monthly snapshot in the Readiness "
                            "Trend log on the Dashboard; the trend chart plots it.", None),

        ("SECTION", "CAPACITY & EXTENDING", None),
        ("Row capacity", f"Scope {N_SCOPE} rows · Equipment {N_EQUIP} · Documents {N_DOC} · "
                         f"Actions {N_ACTION} · Milestones {N_MILESTONE}.", None),
        ("To add more rows", "Select the last data row, copy it, and paste into the rows below. "
                             "Then widen the named ranges (Formulas → Name Manager) so the "
                             "roll-ups see the new rows.", None),
        ("Hidden columns", "Each tracker has hidden helper columns to the right of the visible "
                           "ones. They feed the dashboard and charts. Leave them alone.", None),
        ("Hidden sheet", "'Configuration' holds the dropdown lists, progress maps, weights and "
                         "chart data. Right-click any tab → Unhide to reach it.", None),
    ]

    r = 4
    for a, b, colour in rows:
        if a == "SECTION":
            section(ws, r, 2, 3, b)
            r += 1
            continue
        c = ws.cell(row=r, column=2, value=a)
        c.font = f_lbl
        c.alignment = c_wrap
        c.border = box
        if colour:
            c.fill = PatternFill("solid", fgColor=colour)
        d = ws.cell(row=r, column=3, value=b)
        d.font = f_body
        d.alignment = c_wrap
        d.border = box
        ws.row_dimensions[r].height = 30
        r += 1

    ws.cell(row=r + 1, column=2, value="Generated").font = f_note
    ws.cell(row=r + 1, column=3,
            value=f"{dt.date.today():%d-%b-%Y} from DMIC_COMMISSIONING_CONTRACT_"
                  f"MANAGEMENT_WORKBOOK_v3.docx").font = f_note


# ===========================================================================
# Defined names
# ===========================================================================
def add_names(wb):
    q = "'Scope Tracker'"
    e = "'Equipment Tracker'"
    d = "'Document Tracker'"
    a = "'Action Tracker'"
    names = {
        # Configuration lists
        "Cfg_ContractStatus": f"{CFG}!$A${LIST_TOP}:$A${LIST_TOP + len(CONTRACT_STATUS) - 1}",
        "Cfg_ScopeStatus": f"{CFG}!$B${LIST_TOP}:$B${LIST_TOP + len(SCOPE_STATUS) - 1}",
        "Cfg_Phase": f"{CFG}!$C${LIST_TOP}:$C${LIST_TOP + len(PHASES) - 1}",
        "Cfg_WorkflowStatus": f"{CFG}!$D${LIST_TOP}:$D${LIST_TOP + len(WORKFLOW) - 1}",
        "Cfg_Gate": f"{CFG}!$E${LIST_TOP}:$E${LIST_TOP + len(GATES) - 1}",
        "Cfg_Priority": f"{CFG}!$F${LIST_TOP}:$F${LIST_TOP + len(PRIORITY) - 1}",
        "Cfg_ActionStatus": f"{CFG}!$G${LIST_TOP}:$G${LIST_TOP + len(ACTION_STATUS) - 1}",
        "Cfg_Criticality": f"{CFG}!$H${LIST_TOP}:$H${LIST_TOP + len(CRITICALITY) - 1}",
        "Cfg_YesNo": f"{CFG}!$I${LIST_TOP}:$I${LIST_TOP + len(YES_NO) - 1}",
        "Cfg_MilestoneStatus": f"{CFG}!$J${LIST_TOP}:$J${LIST_TOP + len(MILESTONE_STATUS) - 1}",
        "Cfg_DocType": f"{CFG}!$K${LIST_TOP}:$K${LIST_TOP + len(DOC_TYPES) - 1}",
        "Cfg_Personnel": f"{CFG}!$L${LIST_TOP}:$L${LIST_TOP + len(PERSONNEL) - 1}",
        "Cfg_ActionCategory": f"{CFG}!$M${LIST_TOP}:$M${LIST_TOP + len(ACTION_CATEGORY) - 1}",
        # Maps
        "Cfg_PhaseList": f"{CFG}!$A${PHASE_TOP}:$A${PHASE_BOT}",
        "Cfg_PhasePct": f"{CFG}!$B${PHASE_TOP}:$B${PHASE_BOT}",
        "Cfg_WfList": f"{CFG}!$D${WF_TOP}:$D${WF_BOT}",
        "Cfg_WfPct": f"{CFG}!$E${WF_TOP}:$E${WF_BOT}",
        "Cfg_WfApproved": f"{CFG}!$F${WF_TOP}:$F${WF_BOT}",
        # Parameters
        "Cfg_WPhase": f"{CFG}!$I${MAP_TOP}",
        "Cfg_WDoc": f"{CFG}!$I${MAP_TOP + 1}",
        "Cfg_ReviewDays": f"{CFG}!$I${MAP_TOP + 2}",
        "Cfg_ThrExcellent": f"{CFG}!$O${THR_TOP}",
        "Cfg_ThrGood": f"{CFG}!$O${THR_TOP + 1}",
        "Cfg_ThrAttention": f"{CFG}!$O${THR_TOP + 2}",
        "Cfg_RagGreen": f"{CFG}!$O${THR_TOP + 3}",
        "Cfg_RagAmber": f"{CFG}!$O${THR_TOP + 4}",
        "Cfg_DocWindow": f"{CFG}!$O${THR_TOP + 5}",
        "Cfg_ActWindow": f"{CFG}!$O${THR_TOP + 6}",
        "Cfg_ScopeReady": f"{CFG}!$O${THR_TOP + 7}",
        "Cfg_MsWindow": f"{CFG}!$O${THR_TOP + 8}",
        # Chart-data helper columns
        "Cfg_OwnerList": f"{CFG}!$R${CD_TOP}:$R${CD_TOP + len(PERSONNEL) - 1}",
        "Cfg_OwnerBlockers": f"{CFG}!$T${CD_TOP}:$T${CD_TOP + len(PERSONNEL) - 1}",
        # Lead times (master copy lives on the Control Panel)
        "Cfg_LeadType": f"'{CP}'!$B${CP_LEAD_TOP}:$B${CP_LEAD_BOT}",
        "Cfg_LeadDays": f"'{CP}'!$D${CP_LEAD_TOP}:$D${CP_LEAD_BOT}",
        "Cfg_MsList": f"'{CP}'!$B${CP_MS_TOP}:$B${CP_MS_BOT}",
        "Cfg_MsLeadDays": f"'{CP}'!$D${CP_MS_TOP}:$D${CP_MS_BOT}",
        "Ctrl_PlannedStart": f"'{CP}'!$C${CP_DATE_TOP}",
        # Scope tracker
        "Sc_ID": f"{q}!$A${FIRST}:$A${SCOPE_LAST}",
        "Sc_IDList": f"{q}!$A${FIRST}:$A${SCOPE_LAST}",
        "Sc_Readiness": f"{q}!$J${FIRST}:$J${SCOPE_LAST}",
        # Equipment tracker
        "Eq_Name": f"{e}!$A${FIRST}:$A${EQUIP_LAST}",
        "Eq_NameList": f"{e}!$A${FIRST}:$A${EQUIP_LAST}",
        "Eq_Scope": f"{e}!$B${FIRST}:$B${EQUIP_LAST}",
        "Eq_Phase": f"{e}!$D${FIRST}:$D${EQUIP_LAST}",
        "Eq_Readiness": f"{e}!$K${FIRST}:$K${EQUIP_LAST}",
        "Eq_Status": f"{e}!$L${FIRST}:$L${EQUIP_LAST}",
        "Eq_PhasePct": f"{e}!$O${FIRST}:$O${EQUIP_LAST}",
        # Document tracker
        "Doc_Type": f"{d}!$A${FIRST}:$A${DOC_LAST}",
        "Doc_Scope": f"{d}!$B${FIRST}:$B${DOC_LAST}",
        "Doc_Equipment": f"{d}!$C${FIRST}:$C${DOC_LAST}",
        "Doc_Number": f"{d}!$D${FIRST}:$D${DOC_LAST}",
        "Doc_NumberList": f"{d}!$D${FIRST}:$D${DOC_LAST}",
        "Doc_Owner": f"{d}!$G${FIRST}:$G${DOC_LAST}",
        "Doc_Workflow": f"{d}!$I${FIRST}:$I${DOC_LAST}",
        "Doc_Gate": f"{d}!$J${FIRST}:$J${DOC_LAST}",
        "Doc_Approved": f"{d}!$S${FIRST}:$S${DOC_LAST}",
        "Doc_Progress": f"{d}!$T${FIRST}:$T${DOC_LAST}",
        "Doc_Flag": f"{d}!$V${FIRST}:$V${DOC_LAST}",
        "Doc_RankKey": f"{d}!$W${FIRST}:$W${DOC_LAST}",
        # Action tracker
        "Act_ID": f"{a}!$A${FIRST}:$A${ACT_LAST}",
        "Act_Desc": f"{a}!$B${FIRST}:$B${ACT_LAST}",
        "Act_Scope": f"{a}!$D${FIRST}:$D${ACT_LAST}",
        "Act_Owner": f"{a}!$G${FIRST}:$G${ACT_LAST}",
        "Act_Blocker": f"{a}!$M${FIRST}:$M${ACT_LAST}",
        "Act_Status": f"{a}!$N${FIRST}:$N${ACT_LAST}",
        "Act_DaysToDue": f"{a}!$Q${FIRST}:$Q${ACT_LAST}",
        "Act_Overdue": f"{a}!$R${FIRST}:$R${ACT_LAST}",
        "Act_RankKey": f"{a}!$S${FIRST}:$S${ACT_LAST}",
        # Milestones
        "Ms_Name": f"{DASH}!$B${MS_FIRST}:$B${MS_LAST}",
        "Ms_Status": f"{DASH}!$G${MS_FIRST}:$G${MS_LAST}",
        "Ms_Days": f"{DASH}!$I${MS_FIRST}:$I${MS_LAST}",
    }
    for name, ref in names.items():
        wb.defined_names.add(DefinedName(name, attr_text=ref))


# ===========================================================================
# Demo / example data
# ===========================================================================
DEMO_SCOPES = [
    ("SC-01", "Traction Power & 33kV Distribution", "A. Mercer", "In Progress"),
    ("SC-02", "Signalling & Train Control", "S. Okonkwo", "Review"),
    ("SC-03", "Tunnel Ventilation & Environmental Control", "L. Fontaine", "Preparation"),
]

DEMO_EQUIPMENT = [
    ("TR-BSS-01 Bulk Supply Substation", "SC-01", "D. Petrov", "Complete"),
    ("TR-TSS-02 Traction Substation Alpha", "SC-01", "D. Petrov", "Ready For Handover"),
    ("TR-TSS-03 Traction Substation Bravo", "SC-01", "M. Haddad", "Testing In Progress"),
    ("TR-OHL-04 Overhead Line Section A", "SC-01", "M. Haddad", "Results Review"),
    ("SG-IXL-01 Interlocking Cabinet North", "SC-02", "J. Whitfield", "Testing In Progress"),
    ("SG-IXL-02 Interlocking Cabinet South", "SC-02", "J. Whitfield", "Testing Preparation"),
    ("SG-ATP-03 ATP Trackside Equipment", "SC-02", "K. Rasmussen", "Documentation Review"),
    ("SG-OCC-04 Control Centre Workstations", "SC-02", "K. Rasmussen", "Commissioning Preparation"),
    ("TV-JET-01 Jet Fan Array Tunnel 1", "SC-03", "P. Nakamura", "Commissioning Preparation"),
    ("TV-JET-02 Jet Fan Array Tunnel 2", "SC-03", "P. Nakamura", "Pre-Commissioning"),
    ("TV-SMK-03 Smoke Extraction Plant", "SC-03", "R. Bellamy", "Pre-Commissioning"),
    ("TV-BMS-04 Environmental Control BMS", "SC-03", "R. Bellamy", "Documentation Review"),
]

# Scope-level deliverables, and the workflow status per scope (SC-01, SC-02, SC-03)
SCOPE_DOCS = [
    ("Basis of Commissioning", "Preparation", "Critical", ["Closed", "Approved", "Under Review"]),
    ("Surveillance Strategy", "Preparation", "High", ["Closed", "Approved", "Draft"]),
    ("Surveillance Programme", "Readiness", "High", ["Approved", "Resubmitted", "Draft"]),
    ("Competency Evidence", "Readiness", "Critical", ["Approved", "Submitted", "Not Started"]),
    ("Safe System Of Work", "Readiness", "Critical", ["Approved", "Under Review", "Draft"]),
    ("Risk Assessment And Method Statement", "Readiness", "Critical",
     ["Approved", "Comments Received", "Not Started"]),
    ("System Commissioning Programme", "Readiness", "High",
     ["Approved", "Under Review", "Not Started"]),
    ("Initial End Of Erection Status Report", "Readiness", "Medium",
     ["Closed", "Draft", "Not Started"]),
    ("Planned End Of Erection Status Report", "Authority To Start", "High",
     ["Approved", "Draft", "Not Started"]),
    ("Authority To Start Evidence", "Authority To Start", "Critical",
     ["Approved", "Not Started", "Not Started"]),
]

# Equipment-level deliverables, status keyed by the equipment's index within its scope
EQUIP_DOCS = [
    ("Commissioning Procedure", "Commissioning", "Critical",
     ["Closed", "Approved", "Under Review", "Approved",
      "Approved", "Rework", "Draft", "Not Started",
      "Draft", "Not Started", "Not Started", "Internal Review"]),
    ("Inspection and Test Plan", "Commissioning", "High",
     ["Closed", "Approved", "Comments Received", "Approved",
      "Under Review", "Draft", "Not Started", "Not Started",
      "Not Started", "Not Started", "Not Started", "Draft"]),
]

HANDOVER_DOCS = [
    ("Handover Certificate", "SC-01", "Handover", "Medium", "Draft"),
    ("Functional Takeover Package", "SC-01", "Handover", "High", "Not Started"),
    ("Commissioning Results Sheet", "SC-01", "Results", "Medium", "Draft"),
    ("Final End Of Erection Status Report", "SC-01", "Final Takeover", "Medium", "Not Started"),
]

DEMO_ACTIONS = [
    ("ACT-001", "Close out RAMS comments raised by the Client safety review", "Safety",
     "SC-02", "", "Critical", "Yes", "Open", -18, "J. Whitfield", "M. Haddad"),
    ("ACT-002", "Obtain signed competency records for the ATP trackside team", "Quality",
     "SC-02", "SG-ATP-03 ATP Trackside Equipment", "Critical", "Yes", "In Progress", -9,
     "K. Rasmussen", "S. Okonkwo"),
    ("ACT-003", "Resolve interlocking cabinet earthing non-conformance", "Technical",
     "SC-02", "SG-IXL-01 Interlocking Cabinet North", "High", "Yes", "Open", -4,
     "J. Whitfield", "D. Petrov"),
    ("ACT-004", "Issue System Commissioning Programme revision C for approval", "Document",
     "SC-02", "", "High", "No", "In Progress", 6, "S. Okonkwo", "L. Fontaine"),
    ("ACT-005", "Agree tunnel ventilation test witnessing schedule with the Authority", "Interface",
     "SC-03", "", "High", "No", "Open", 11, "L. Fontaine", "P. Nakamura"),
    ("ACT-006", "Confirm 33kV protection settings against the final study", "Technical",
     "SC-01", "TR-TSS-03 Traction Substation Bravo", "Medium", "No", "In Progress", 19,
     "M. Haddad", "D. Petrov"),
    ("ACT-007", "Update Basis of Commissioning for the revised staging strategy", "Document",
     "SC-03", "", "Medium", "No", "Waiting", 27, "L. Fontaine", ""),
    ("ACT-008", "Mobilise second commissioning engineer for tunnel ventilation", "Resource",
     "SC-03", "", "High", "No", "Open", 34, "A. Mercer", "L. Fontaine"),
    ("ACT-009", "Close punch list items on Traction Substation Alpha", "Technical",
     "SC-01", "TR-TSS-02 Traction Substation Alpha", "Medium", "No", "In Progress", 13,
     "D. Petrov", "M. Haddad"),
    ("ACT-010", "Submit Initial EESR for signalling scope", "Document",
     "SC-02", "", "High", "No", "Open", -2, "K. Rasmussen", ""),
    ("ACT-011", "Agree beneficial use conditions with Operations", "Client",
     "SC-01", "", "Medium", "No", "Waiting", 45, "A. Mercer", "S. Okonkwo"),
    ("ACT-012", "Verify smoke extraction damper actuator supply", "Technical",
     "SC-03", "TV-SMK-03 Smoke Extraction Plant", "Low", "No", "Open", 52,
     "R. Bellamy", ""),
    ("ACT-013", "Complete traction power commissioning phase report", "Document",
     "SC-01", "", "Medium", "No", "Open", 61, "D. Petrov", "A. Mercer"),
    ("ACT-014", "Rectify jet fan vibration readings on Tunnel 1", "Technical",
     "SC-03", "TV-JET-01 Jet Fan Array Tunnel 1", "High", "Yes", "Open", -1,
     "P. Nakamura", "R. Bellamy"),
    ("ACT-015", "Archive approved bulk supply substation records", "Quality",
     "SC-01", "TR-BSS-01 Bulk Supply Substation", "Low", "No", "Closed", -40,
     "M. Haddad", ""),
]

DOC_ABBREV = {
    "Basis of Commissioning": "BOC", "Surveillance Strategy": "SVS",
    "Surveillance Programme": "SVP", "Competency Evidence": "CPE",
    "Safe System Of Work": "SSW", "Risk Assessment And Method Statement": "RAM",
    "System Commissioning Programme": "SCP", "Commissioning Procedure": "CPR",
    "Inspection and Test Plan": "ITP",
    "Initial End Of Erection Status Report": "IER",
    "Planned End Of Erection Status Report": "PER",
    "Final End Of Erection Status Report": "FER",
    "Authority To Start Evidence": "ATS", "Commissioning Results Sheet": "CRS",
    "Commissioning Phase Report": "CPT", "Handover Certificate": "HOC",
    "Functional Takeover Package": "FTP", "Care And Maintenance Plan": "CMP",
    "Maintenance Transfer Package": "MTP",
    "Temporary Operations Transfer Package": "TOP",
}

REV_BY_STATUS = {
    "Not Started": "-", "Draft": "A", "Internal Review": "A", "Submitted": "A",
    "Under Review": "A", "Comments Received": "A", "Rework": "B", "Resubmitted": "B",
    "Approved": "C", "Released": "C", "Implemented": "C", "Closed": "C",
}
APPROVED_SET = {"Approved", "Released", "Implemented", "Closed"}


def write_demo(wb, planned_start):
    today = dt.date.today()

    # --- Scopes -------------------------------------------------------------
    ws = wb["Scope Tracker"]
    offsets = [(-210, 120), (-150, 165), (-60, 210)]
    for i, (sid, name, owner, status) in enumerate(DEMO_SCOPES):
        r = FIRST + i
        ps = planned_start + dt.timedelta(days=offsets[i][0])
        pf = planned_start + dt.timedelta(days=offsets[i][1])
        ws.cell(row=r, column=1, value=sid)
        ws.cell(row=r, column=2, value=name)
        ws.cell(row=r, column=3, value=owner)
        ws.cell(row=r, column=4, value=ps)
        ws.cell(row=r, column=5, value=pf)
        ws.cell(row=r, column=6, value=ps + dt.timedelta(days=[7, 21, 35][i]))
        ws.cell(row=r, column=7, value=pf + dt.timedelta(days=[7, 28, 45][i]))
        if i == 0:
            ws.cell(row=r, column=8, value=ps + dt.timedelta(days=9))
        ws.cell(row=r, column=11, value=status)
        ws.cell(row=r, column=18,
                value=["Lead scope — first to reach handover.",
                       "RAMS comments outstanding with the Client.",
                       "Late start; resourcing under review."][i])
        for col in (4, 5, 6, 7, 8):
            ws.cell(row=r, column=col).number_format = FMT_DATE

    # --- Equipment ----------------------------------------------------------
    ws = wb["Equipment Tracker"]
    for i, (name, scope, eng, phase) in enumerate(DEMO_EQUIPMENT):
        r = FIRST + i
        base = planned_start + dt.timedelta(days=-180 + i * 14)
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=scope)
        ws.cell(row=r, column=3, value=eng)
        ws.cell(row=r, column=4, value=phase)
        ws.cell(row=r, column=5, value=base)
        ws.cell(row=r, column=6, value=base + dt.timedelta(days=5 + i))
        ws.cell(row=r, column=8, value=base + dt.timedelta(days=75))
        ws.cell(row=r, column=9, value=base + dt.timedelta(days=80 + i * 2))
        if phase in ("Complete", "Ready For Handover", "Results Review", "Testing In Progress"):
            ws.cell(row=r, column=7, value=base + dt.timedelta(days=6 + i))
        if phase == "Complete":
            ws.cell(row=r, column=10, value=base + dt.timedelta(days=72))
        for col in (5, 6, 7, 8, 9, 10):
            ws.cell(row=r, column=col).number_format = FMT_DATE

    # --- Documents ----------------------------------------------------------
    ws = wb["Document Tracker"]
    lead_lookup = {n: d for n, _, d in DOC_TYPES}
    review_days = 21
    row = FIRST
    counter = 0

    def put_doc(dtype, scope, equip, gate, crit, status, title):
        nonlocal row, counter
        counter += 1
        lead = lead_lookup[dtype]
        target_appr = planned_start - dt.timedelta(days=lead)
        ws.cell(row=row, column=1, value=dtype)
        ws.cell(row=row, column=2, value=scope)
        ws.cell(row=row, column=3, value=equip)
        ws.cell(row=row, column=4,
                value=f"DMIC-{scope}-{DOC_ABBREV.get(dtype, 'DOC')}-{counter:03d}")
        ws.cell(row=row, column=5, value=title)
        ws.cell(row=row, column=6, value=REV_BY_STATUS[status])
        ws.cell(row=row, column=7, value=PERSONNEL[counter % len(PERSONNEL)])
        ws.cell(row=row, column=8, value=PERSONNEL[(counter + 5) % len(PERSONNEL)])
        ws.cell(row=row, column=9, value=status)
        ws.cell(row=row, column=10, value=gate)
        if status != "Not Started":
            sub = target_appr - dt.timedelta(days=review_days)
            actual_sub = sub + dt.timedelta(days=(counter % 11) - 3)
            if actual_sub <= today:
                c = ws.cell(row=row, column=13, value=actual_sub)
                c.number_format = FMT_DATE
        if status in APPROVED_SET:
            actual_appr = target_appr + dt.timedelta(days=(counter % 9) - 4)
            if actual_appr <= today:
                c = ws.cell(row=row, column=14, value=actual_appr)
                c.number_format = FMT_DATE
        ws.cell(row=row, column=16, value=crit)
        row += 1

    for si, (sid, sname, _, _) in enumerate(DEMO_SCOPES):
        for dtype, gate, crit, statuses in SCOPE_DOCS:
            put_doc(dtype, sid, "", gate, crit, statuses[si], f"{sname} — {dtype}")

    for ei, (ename, scope, _, _) in enumerate(DEMO_EQUIPMENT):
        for dtype, gate, crit, statuses in EQUIP_DOCS:
            put_doc(dtype, scope, ename, gate, crit, statuses[ei],
                    f"{ename} — {dtype}")

    for dtype, scope, gate, crit, status in HANDOVER_DOCS:
        put_doc(dtype, scope, "", gate, crit, status, f"{scope} — {dtype}")

    # --- Actions ------------------------------------------------------------
    ws = wb["Action Tracker"]
    for i, (aid, desc, cat, scope, equip, prio, blocker, status,
            due_offset, owner, support) in enumerate(DEMO_ACTIONS):
        r = FIRST + i
        due = today + dt.timedelta(days=due_offset)
        raised = due - dt.timedelta(days=28)
        ws.cell(row=r, column=1, value=aid)
        ws.cell(row=r, column=2, value=desc)
        ws.cell(row=r, column=3, value=cat)
        ws.cell(row=r, column=4, value=scope)
        ws.cell(row=r, column=5, value=equip)
        ws.cell(row=r, column=7, value=owner)
        ws.cell(row=r, column=8, value=support)
        ws.cell(row=r, column=9, value=raised)
        ws.cell(row=r, column=10, value=due)
        if status == "Closed":
            ws.cell(row=r, column=11, value=due - dt.timedelta(days=3))
        ws.cell(row=r, column=12, value=prio)
        ws.cell(row=r, column=13, value=blocker)
        ws.cell(row=r, column=14, value=status)
        for col in (9, 10, 11):
            ws.cell(row=r, column=col).number_format = FMT_DATE

    # --- Milestones ---------------------------------------------------------
    ws = wb[DASH]
    ms_status = ["Achieved", "In Progress", "At Risk", "Not Started", "Not Started",
                 "Not Started", "Not Started", "Not Started", "Not Started"]
    slip = [0, 5, 12, 21, 21, 21, 21, 21, 21]
    for i, (name, lead) in enumerate(MILESTONES):
        r = MS_FIRST + i
        ws.cell(row=r, column=3, value="SC-01")
        # Forecast Date: the planned date plus a little slippage
        ws.cell(row=r, column=5,
                value=planned_start - dt.timedelta(days=lead - slip[i]))
        ws.cell(row=r, column=5).number_format = FMT_DATE
        if ms_status[i] == "Achieved":
            ws.cell(row=r, column=6, value=planned_start - dt.timedelta(days=lead - 2))
            ws.cell(row=r, column=6).number_format = FMT_DATE
        ws.cell(row=r, column=7, value=ms_status[i])
        ws.cell(row=r, column=8, value=PERSONNEL[i % len(PERSONNEL)])

    # --- Readiness trend ----------------------------------------------------
    trend = [0.08, 0.14, 0.19, 0.26, 0.31, 0.35, 0.39, 0.44]
    for i, v in enumerate(trend):
        r = TREND_FIRST + i
        d = (today.replace(day=1) - dt.timedelta(days=30 * (len(trend) - 1 - i)))
        c = ws.cell(row=r, column=10, value=d)
        c.number_format = FMT_DATE
        c.font = f_in
        c2 = ws.cell(row=r, column=11, value=v)
        c2.number_format = FMT_PCT
        c2.font = f_in


def write_examples(wb, planned_start):
    """Blank template: one realistic example row per tracker, in row 4."""
    tag = "◀ EXAMPLE ROW — overwrite with your own data"

    ws = wb["Scope Tracker"]
    vals = {1: "SC-01", 2: "Traction Power & 33kV Distribution", 3: "A. Mercer",
            4: planned_start - dt.timedelta(days=210),
            5: planned_start + dt.timedelta(days=120),
            6: planned_start - dt.timedelta(days=203),
            7: planned_start + dt.timedelta(days=127),
            11: "In Progress", 18: tag}
    for col, v in vals.items():
        c = ws.cell(row=FIRST, column=col, value=v)
        c.fill = fill_example
        if col in (4, 5, 6, 7):
            c.number_format = FMT_DATE

    ws = wb["Equipment Tracker"]
    vals = {1: "TR-TSS-02 Traction Substation Alpha", 2: "SC-01", 3: "D. Petrov",
            4: "Testing In Progress",
            5: planned_start - dt.timedelta(days=180),
            6: planned_start - dt.timedelta(days=175),
            8: planned_start - dt.timedelta(days=105),
            9: planned_start - dt.timedelta(days=98), 13: tag}
    for col, v in vals.items():
        c = ws.cell(row=FIRST, column=col, value=v)
        c.fill = fill_example
        if col in (5, 6, 8, 9):
            c.number_format = FMT_DATE

    ws = wb["Document Tracker"]
    vals = {1: "System Commissioning Programme", 2: "SC-01",
            3: "TR-TSS-02 Traction Substation Alpha",
            4: "DMIC-SC-01-SCP-001",
            5: "Traction Power — System Commissioning Programme", 6: "B",
            7: "D. Petrov", 8: "M. Haddad", 9: "Under Review", 10: "Readiness",
            16: "Critical", 17: tag}
    for col, v in vals.items():
        ws.cell(row=FIRST, column=col, value=v).fill = fill_example

    ws = wb["Action Tracker"]
    today = dt.date.today()
    vals = {1: "ACT-001",
            2: "Close out RAMS comments raised by the Client safety review",
            3: "Safety", 4: "SC-01", 5: "TR-TSS-02 Traction Substation Alpha",
            6: "DMIC-SC-01-SCP-001", 7: "A. Mercer", 8: "M. Haddad",
            9: today - dt.timedelta(days=14), 10: today + dt.timedelta(days=7),
            12: "High", 13: "Yes", 14: "Open", 15: tag}
    for col, v in vals.items():
        c = ws.cell(row=FIRST, column=col, value=v)
        c.fill = fill_example
        if col in (9, 10):
            c.number_format = FMT_DATE


# ===========================================================================
# Build
# ===========================================================================
def build(demo, path, planned_start):
    wb = Workbook()
    wb.remove(wb.active)

    ws_readme = wb.create_sheet("Read Me")
    ws_cp = wb.create_sheet(CP)
    ws_scope = wb.create_sheet("Scope Tracker")
    ws_equip = wb.create_sheet("Equipment Tracker")
    ws_doc = wb.create_sheet("Document Tracker")
    ws_act = wb.create_sheet("Action Tracker")
    ws_dash = wb.create_sheet(DASH)
    ws_cfg = wb.create_sheet(CFG)

    build_readme(ws_readme, demo, planned_start)
    build_control_panel(ws_cp, demo, planned_start)
    build_scope_tracker(ws_scope)
    build_equipment_tracker(ws_equip)
    build_document_tracker(ws_doc)
    build_action_tracker(ws_act)
    build_dashboard(ws_dash, demo)
    build_configuration(ws_cfg)
    add_charts(ws_dash, ws_cfg)
    add_names(wb)

    if demo:
        write_demo(wb, planned_start)
    else:
        write_examples(wb, planned_start)

    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(path)
    print(f"wrote {path}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--demo", action="store_true")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("--start", default="2027-06-01",
                    help="Planned Commissioning Start Date (YYYY-MM-DD)")
    args = ap.parse_args()
    build(args.demo, args.output, dt.date.fromisoformat(args.start))


if __name__ == "__main__":
    main()
