#!/usr/bin/env python3
"""
DMIC Commissioning Control - workbook generator (v2).

Merges the v50 calculation engine (per-coverage target dates, milestone
anchors, 21 workflow codes, core-gap detection, data-quality Check column)
with a flat TRACKER table and the visual treatment of v1.

    python3 build_v2.py -o DMIC_Commissioning_Control.xlsx
    python3 build_v2.py --demo -o DMIC_Commissioning_Control_DEMO.xlsx
"""

import argparse
import datetime as dt

from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import (Alignment, Border, Font, PatternFill, Protection,
                             Side)
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------------------
# Capacity
# ---------------------------------------------------------------------------
TRK_TOP, N_TRK = 4, 500
TRK_BOT = TRK_TOP + N_TRK - 1                 # 503
ACT_TOP, N_ACT = 4, 150
ACT_BOT = ACT_TOP + N_ACT - 1                 # 153
N_SYS, N_EQP, N_PPL = 50, 80, 30
N_MS_SLOT = 15
N_CORE = 12                                   # core-type slots tracked for gaps
LIST_ROWS = 10                                # rows per dashboard list

SU = "SETUP"

# ---------------------------------------------------------------------------
# LISTS geometry
# ---------------------------------------------------------------------------
LS = "LISTS"
CORE_TOP = 2
CORE_BOT = CORE_TOP + N_CORE - 1              # 2..13
GAP_TOP = 2
N_GAP = N_CORE + N_SYS * N_CORE + N_EQP * N_CORE
GAP_BOT = GAP_TOP + N_GAP - 1                 # 2..1273
CMB_TOP = 2
N_CMB = N_TRK + N_ACT
CMB_BOT = CMB_TOP + N_CMB - 1                 # 2..651
DST_TOP, DST_BOT = 2, 13                      # document states
AST2_TOP, AST2_BOT = 2, 5                     # action statuses
BYS_TOP = 2
BYS_BOT = BYS_TOP + N_SYS - 1
MON_TOP, MON_BOT = 2, 13                      # 12 month buckets
EQW_TOP = 2
EQW_BOT = EQW_TOP + N_EQP - 1

DB = "DASHBOARD"
TR = "TRACKER"
AC = "ACTIONS"

# ---------------------------------------------------------------------------
# Reference data
# ---------------------------------------------------------------------------
# (label shown in SETUP, default wording, defined-name)
WORDING = [
    ("Document approved (S4A reached) or published", "Done", "S_DONE"),
    ("Target approval date passed without approval", "Late", "S_LATE"),
    ("Actual start filled in, not late", "In progress", "S_WIP"),
    ("Due inside the document alert window", "At risk", "S_RISK"),
    ("Due inside the coming horizon, not started", "Coming", "S_COMING"),
    ("Due beyond the horizon", "Not yet due", "S_NOTYET"),
    ("Cancelled, revoked, superseded, transferred, destroyed", "Inactive", "S_INACT"),
    ("Actions only: due beyond the action alert window", "On track", "S_ONTRACK"),
    ("No commissioning start date behind this row", "Scope not resolved", "S_NOSCOPE"),
    ("Anchor milestone has no date", "No target date", "S_NOTARGET"),
    ("Type missing from the library", "Type not configured", "S_TYPEKO"),
    ("Actions only: due date not filled in", "No due date", "S_NODUE"),
]

RAG_WORDING = [
    ("At least one document or action is late", "DELAYED", "RAG_RED"),
    ("Nothing late, but something is at risk or missing", "AT RISK OF DELAY", "RAG_AMB"),
    ("Nothing late, nothing at risk", "ON SCHEDULE", "RAG_GRN"),
]

ACTION_WORDING = [
    ("Action status 1", "Open", "A_OPEN"),
    ("Action status 2", "In progress", "A_WIP"),
    ("Action status 3", "On hold", "A_HOLD"),
    ("Action status 4", "Done", "A_DONE"),
]

# (code, family, counts as)
WF_CODES = [
    ("W0 - Work In Progress", "Preparation", "In progress"),
    ("W1 - Draft", "Preparation", "In progress"),
    ("W2 - Revoked", "Preparation", "Inactive"),
    ("S2 - For Information", "Review", "In progress"),
    ("S3 - For Review", "Review", "In progress"),
    ("S4 - For Internal Approval", "Review", "In progress"),
    ("S4A - Approved", "Approval result", "Done"),
    ("S4B - Comments", "Approval result", "In progress"),
    ("S4C - Rejected", "Approval result", "In progress"),
    ("S6 - For Acceptance", "Acceptance", "In progress"),
    ("S6Q - Quality Checked", "Acceptance", "In progress"),
    ("S6B - Comments", "Acceptance", "In progress"),
    ("S6C - Rejected", "Acceptance", "In progress"),
    ("P1 - For Implementation", "Published", "Done"),
    ("P2 - For Costing", "Published", "Done"),
    ("P3 - For Tender", "Published", "Done"),
    ("P4 - For Contractor Design", "Published", "Done"),
    ("P5 - For Manufacture / Procurement", "Published", "Done"),
    ("P6 - For Construction", "Published", "Done"),
    ("P7 - As-Built", "Published", "Done"),
    ("A1 - Cancelled", "Closed", "Inactive"),
    ("A2 - Superseded", "Closed", "Inactive"),
    ("A3 - Transferred", "Closed", "Inactive"),
    ("A4 - Destroyed", "Closed", "Inactive"),
]

# (milestone, offset in days from earliest commissioning start)
MILESTONES = [
    ("Kick Off Meeting", -360),
    ("Commissioning Readiness Review", -60),
    ("Authority To Start", -14),
    ("NTP", -7),
    ("Commissioning Start", 0),
    ("Commissioning Finish", 180),
    ("Functional Takeover", 195),
    ("Beneficial Use", 205),
    ("Maintenance Transfer", 220),
    ("Temporary Operations Transfer", 235),
    ("Final Takeover", 250),
]

# (type, phase, core required (ATS), anchor milestone, lead days, default scope, note)
TYPES = [
    ("Basis of Commissioning", "Pre-commissioning", "Yes", "Commissioning Start", 360, "All", "T-12 months"),
    ("Surveillance Strategy", "Pre-commissioning", "Yes", "Commissioning Start", 360, "All", "T-12 months, ATS"),
    ("Surveillance Programme", "Pre-commissioning", "Yes", "Commissioning Start", 270, "System", "T-9 months, ATS"),
    ("Competency Evidence", "Pre-commissioning", "Yes", "Commissioning Start", 270, "All", "T-9 months, ATS"),
    ("Safe System Of Work", "Pre-commissioning", "Yes", "Commissioning Start", 180, "System", "T-6 months, ATS"),
    ("Risk Assessment And Method Statement", "Pre-commissioning", "Yes", "Commissioning Start", 180, "System", "T-6 months, ATS"),
    ("System Commissioning Programme", "Pre-commissioning", "Yes", "Commissioning Start", 180, "System", "T-6 months, ATS"),
    ("Commissioning Procedure", "Pre-commissioning", "Yes", "Commissioning Start", 90, "Equipment", "T-3 months, ATS"),
    ("Inspection and Test Plan", "Pre-commissioning", "Yes", "Commissioning Start", 90, "Equipment", "T-3 months, ATS"),
    ("Initial End Of Erection Status Report", "Pre-commissioning", "No", "Commissioning Start", 90, "Equipment", "T-3 months"),
    ("Planned End Of Erection Status Report", "Pre-commissioning", "Yes", "Commissioning Start", 30, "Equipment", "T-1 month, ATS"),
    ("Authority To Start Evidence", "Pre-commissioning", "Yes", "Commissioning Start", 14, "System", "T-2 weeks, ATS"),
    ("Commissioning Results Sheet", "Commissioning", "No", "Commissioning Finish", 0, "Equipment", "During commissioning"),
    ("Commissioning Phase Report", "Post-commissioning", "No", "Functional Takeover", 15, "System", "Anchor estimated"),
    ("Handover Certificate", "Post-commissioning", "No", "Functional Takeover", 0, "Equipment", "Anchor estimated"),
    ("Functional Takeover Package", "Post-commissioning", "No", "Functional Takeover", 15, "System", "Before FTC"),
    ("Care And Maintenance Plan", "Post-commissioning", "No", "Maintenance Transfer", 30, "System", "Not in the spec lead table"),
    ("Maintenance Transfer Package", "Post-commissioning", "No", "Maintenance Transfer", 15, "System", "Before TFM"),
    ("Temporary Operations Transfer Package", "Post-commissioning", "No", "Temporary Operations Transfer", 15, "System", "Before TFTO"),
    ("Final End Of Erection Status Report", "Post-commissioning", "No", "Final Takeover", 15, "System", "Before Final Takeover"),
]

PHASES = [("Pre-commissioning", 1), ("Commissioning", 2), ("Post-commissioning", 3)]
SCOPES = ["All", "System", "Equipment"]

# ---------------------------------------------------------------------------
# SETUP geometry - each section's rows are derived from the row count of the
# section above it, so bumping a capacity constant (N_SYS, N_EQP, N_MS_SLOT,
# adding a WF_CODES or TYPES entry, ...) can never leave one section's data
# overlapping the next section's banner the way a hand-maintained row number
# eventually will.
# ---------------------------------------------------------------------------
def _section(top, n, header_row=False):
    """(top row of this section) x (row count) -> (header row or None, data
    top, data bottom, banner row of the NEXT section). One blank row always
    separates a section's data from the next banner."""
    hdr = top + 1 if header_row else None
    data_top = top + 2 if header_row else top + 1
    data_bot = data_top + n - 1
    return hdr, data_top, data_bot, data_bot + 2


BAN = {}
b = 4
BAN["contract"] = b
_, CT_TOP, CT_BOT, b = _section(b, 5)
BAN["thresholds"] = b
_, TH_TOP, TH_BOT, b = _section(b, 4)
BAN["wording"] = b
_, WD_TOP, WD_BOT, b = _section(b, len(WORDING))
BAN["rag"] = b
_, RAG_TOP, RAG_BOT, b = _section(b, len(RAG_WORDING))
BAN["actstat"] = b
_, AST_TOP, AST_BOT, b = _section(b, len(ACTION_WORDING))
BAN["wf"] = b
N_WF = len(WF_CODES)
WF_HDR, WF_TOP, WF_BOT, b = _section(b, N_WF, header_row=True)
BAN["ms"] = b
MS_HDR, MS_TOP, MS_BOT, b = _section(b, N_MS_SLOT, header_row=True)
BAN["sys"] = b
SYS_HDR, SYS_TOP, SYS_BOT, b = _section(b, N_SYS, header_row=True)
BAN["eqp"] = b
EQP_HDR, EQP_TOP, EQP_BOT, b = _section(b, N_EQP, header_row=True)
BAN["typ"] = b
TYP_HDR, TYP_TOP, TYP_BOT, b = _section(b, len(TYPES), header_row=True)
BAN["phase"] = b
PH_HDR, PH_TOP, PH_BOT, b = _section(b, len(PHASES), header_row=True)
BAN["ppl"] = b
PPL_HDR, PPL_TOP, PPL_BOT, b = _section(b, N_PPL, header_row=True)

# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------
FONT = "Arial"
NAVY = "1F3864"
BLUE = "2F5597"
STEEL = "44618B"
LIGHT = "D9E2F3"
PALE = "EEF3FB"
GREY = "F2F2F2"

C_RED = "C00000"
C_RED_F = "FFC7CE"
C_AMB_F = "FFE699"
C_GRN_F = "C6EFCE"
C_BLU_F = "BDD7EE"
C_GRY_F = "D9D9D9"
C_ORA_F = "F8CBAD"

f_title = Font(name=FONT, size=16, bold=True, color=NAVY)
f_sub = Font(name=FONT, size=9, italic=True, color="595959")
f_sect = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_hdr = Font(name=FONT, size=9, bold=True, color="FFFFFF")
f_lbl = Font(name=FONT, size=10, bold=True)
f_lbl9 = Font(name=FONT, size=9, bold=True)
f_in = Font(name=FONT, size=10, color="0000FF")
f_in9 = Font(name=FONT, size=9, color="0000FF")
f_body = Font(name=FONT, size=9)
f_calc = Font(name=FONT, size=9, color="404040")
f_calc10 = Font(name=FONT, size=10)
f_note = Font(name=FONT, size=8, italic=True, color="808080")
f_rag = Font(name=FONT, size=26, bold=True, color="FFFFFF")
f_kpi = Font(name=FONT, size=20, bold=True, color=NAVY)

fill_sect = PatternFill("solid", fgColor=BLUE)
fill_sect2 = PatternFill("solid", fgColor=STEEL)
fill_hdr = PatternFill("solid", fgColor=NAVY)
fill_light = PatternFill("solid", fgColor=LIGHT)
fill_pale = PatternFill("solid", fgColor=PALE)
fill_calc = PatternFill("solid", fgColor=GREY)
fill_white = PatternFill("solid", fgColor="FFFFFF")
fill_grey = PatternFill("solid", fgColor=C_GRY_F)

thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)

a_c = Alignment(horizontal="center", vertical="center")
a_l = Alignment(horizontal="left", vertical="center")
a_w = Alignment(horizontal="left", vertical="center", wrap_text=True)
a_h = Alignment(horizontal="center", vertical="center", wrap_text=True)

FMT_DATE = "DD-MMM-YYYY"
# same, but a zero or negative serial shows as blank instead of ###### / 00-Jan-1900
FMT_DATE0 = "DD-MMM-YYYY;;"
FMT_PCT = "0%"
FMT_PCT1 = "0.0%"
FMT_INT = "#,##0;-#,##0;-"
FMT_DAYS = "#,##0;[Red]-#,##0;-"


def section(ws, row, c1, c2, text, alt=False):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill_sect2 if alt else fill_sect
    cell = ws.cell(row=row, column=c1)
    cell.value = text
    cell.font = f_sect
    cell.alignment = a_l
    ws.row_dimensions[row].height = 18
    return cell


def header(ws, row, c1, labels, height=28):
    for i, h in enumerate(labels):
        c = ws.cell(row=row, column=c1 + i, value=h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = a_h
        c.border = box
    ws.row_dimensions[row].height = height


def lab(ws, row, col, text, note=None, small=False):
    c = ws.cell(row=row, column=col, value=text)
    c.font = f_lbl9 if small else f_lbl
    c.alignment = a_l
    if note:
        c.comment = Comment(note, "Workbook")
    return c


UNLOCKED = Protection(locked=False)


def inp(ws, row, col, value=None, fmt=None, small=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = f_in9 if small else f_in
    c.fill = fill_white
    c.border = box
    c.alignment = a_l
    c.protection = UNLOCKED
    if fmt:
        c.number_format = fmt
    return c


def calc(ws, row, col, formula, fmt=None, note=None, big=False):
    c = ws.cell(row=row, column=col, value=formula)
    c.font = f_calc10 if big else f_calc
    c.fill = fill_calc
    c.border = box
    c.alignment = a_l
    if fmt:
        c.number_format = fmt
    if note:
        c.comment = Comment(note, "Workbook")
    return c


def down(ws, col, r1, r2, tpl, fmt=None, style="calc"):
    """Write a per-row formula down a column. {r} -> row number."""
    for r in range(r1, r2 + 1):
        c = ws.cell(row=r, column=col, value=tpl.format(r=r))
        c.border = box
        c.alignment = a_l
        if style == "calc":
            c.font = f_calc
            c.fill = fill_calc
        else:
            c.font = f_body
        if fmt:
            c.number_format = fmt


def inputs(ws, c1, c2, r1, r2, fmts=None):
    for r in range(r1, r2 + 1):
        for col in range(c1, c2 + 1):
            c = ws.cell(row=r, column=col)
            c.font = f_body
            c.border = box
            c.alignment = a_l
            c.protection = UNLOCKED
            if fmts and col in fmts:
                c.number_format = fmts[col]


def widths(ws, m):
    for k, v in m.items():
        ws.column_dimensions[k].width = v


def databar(ws, rng, colour="638EC6"):
    ws.conditional_formatting.add(rng, DataBarRule(
        start_type="num", start_value=0, end_type="num", end_value=1,
        color=colour, showValue=True))


def dv(ws, formula, rng):
    v = DataValidation(type="list", formula1=formula, allow_blank=True)
    ws.add_data_validation(v)
    v.add(rng)
    return v


# ===========================================================================
# SETUP
# ===========================================================================
def build_setup(ws, demo):
    ws.sheet_properties.tabColor = NAVY
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 2, "B": 44, "C": 22, "D": 20, "E": 20, "F": 18, "G": 16,
                "H": 14, "I": 14, "J": 14, "K": 12, "L": 18, "M": 30, "N": 12,
                "O": 12})

    ws["B1"] = "SETUP - CONTROL PANEL"
    ws["B1"].font = f_title
    ws["B2"] = ("Everything in the workbook is driven from this sheet. Blue cells are yours to "
                "edit; grey cells calculate themselves. To drop a status code, a milestone or "
                "a document type, CLEAR the cells - never delete the row: the rest of the "
                "workbook is addressed by row and deleting one breaks it.")
    ws["B2"].font = f_sub

    # 1 - CONTRACT ----------------------------------------------------------
    section(ws, BAN["contract"], 2, 7, "1 - CONTRACT")
    fields = ["Contract number", "Contract name", "Contract Manager",
              "Commissioning Manager", "Quality Lead"]
    demo_vals = ["DMIC-HPC-04", "DMIC Commissioning - Hinkley Point C",
                 "A. Mercer", "S. Okonkwo", "L. Fontaine"]
    for i, f in enumerate(fields):
        lab(ws, CT_TOP + i, 2, f)
        inp(ws, CT_TOP + i, 3, demo_vals[i] if demo else None)
        ws.merge_cells(start_row=CT_TOP + i, start_column=3,
                       end_row=CT_TOP + i, end_column=4)

    # 2 - THRESHOLDS --------------------------------------------------------
    section(ws, BAN["thresholds"], 2, 7, "2 - THRESHOLDS")
    thr = [("Document alert window (days)", 30,
            "An unapproved document turns At risk inside this many days before its target."),
           ("Action alert window (days)", 3, "Same idea, for actions."),
           ("Coming horizon (days)", 90, "How far ahead the COMING block looks.")]
    for i, (t, v, note) in enumerate(thr):
        lab(ws, TH_TOP + i, 2, t)
        inp(ws, TH_TOP + i, 3, v, FMT_INT)
        n = ws.cell(row=TH_TOP + i, column=4, value=note)
        n.font = f_note
        ws.merge_cells(start_row=TH_TOP + i, start_column=4,
                       end_row=TH_TOP + i, end_column=7)
    lab(ws, TH_TOP + 3, 2, "Earliest commissioning start (calculated)")
    # smallest strictly positive value: EQP_EFF holds 0 for equipment with no date,
    # and zeros sort first, so skip exactly as many of them as there are
    calc(ws, TH_TOP + 3, 3,
         '=IF(COUNTIF(EQP_EFF,">0")=0,"",SMALL(EQP_EFF,COUNTIF(EQP_EFF,"<=0")+1))',
         FMT_DATE,
         note="Earliest commissioning start across all equipment in section 7 - the actual "
              "start where one is filled in, the planned start otherwise. "
              "Every 'Commissioning Start' anchored target date ultimately derives from "
              "the start date of whatever the document covers, not from this single figure.")
    n = ws.cell(row=TH_TOP + 3, column=4,
                value="Drives the milestone dates in section 5.")
    n.font = f_note

    # 3 - WORDING -----------------------------------------------------------
    section(ws, BAN["wording"], 2, 7,
            "3 - WORDING      rename anything here: the whole workbook follows")
    for i, (desc, default, _) in enumerate(WORDING):
        lab(ws, WD_TOP + i, 2, desc, small=True)
        inp(ws, WD_TOP + i, 3, default, small=True)

    section(ws, BAN["rag"], 2, 7, "3b - CONTRACT STATUS WORDING", alt=True)
    for i, (desc, default, _) in enumerate(RAG_WORDING):
        lab(ws, RAG_TOP + i, 2, desc, small=True)
        inp(ws, RAG_TOP + i, 3, default, small=True)

    section(ws, BAN["actstat"], 2, 7, "3c - ACTION STATUS WORDING", alt=True)
    for i, (desc, default, _) in enumerate(ACTION_WORDING):
        lab(ws, AST_TOP + i, 2, desc, small=True)
        inp(ws, AST_TOP + i, 3, default, small=True)

    # 4 - DOCUMENT STATUS CODES --------------------------------------------
    section(ws, BAN["wf"], 2, 7,
            "4 - DOCUMENT STATUS CODES      'Counts as' is how each code is read everywhere else")
    header(ws, WF_HDR, 2, ["Status code", "Family", "Counts as"])
    for i, (code, fam, cnt) in enumerate(WF_CODES):
        r = WF_TOP + i
        c = ws.cell(row=r, column=2, value=code); c.font = f_body; c.border = box
        c = ws.cell(row=r, column=3, value=fam); c.font = f_body; c.border = box
        inp(ws, r, 4, cnt, small=True)
    ws.cell(row=WF_BOT + 1, column=2, value=(
        "S4A and the P-codes stop the delay counter. A-codes and W2 are Inactive: they leave "
        "every active counter, stay visible in their own column, and no longer count as "
        "coverage for a required deliverable.")).font = f_note

    # 5 - MILESTONES --------------------------------------------------------
    section(ws, BAN["ms"], 2, 8,
            "5 - MILESTONES      dates only; anything to organise goes in ACTIONS")
    header(ws, MS_HDR, 2, ["Milestone", "Offset from earliest start (days)",
                           "Date typed (optional)", "Planned date", "Actual date", "Note"])
    for i in range(N_MS_SLOT):
        r = MS_TOP + i
        if i < len(MILESTONES):
            c = ws.cell(row=r, column=2, value=MILESTONES[i][0])
            c.font = f_body; c.border = box
            inp(ws, r, 3, MILESTONES[i][1], FMT_INT)
        else:
            inp(ws, r, 2, None)
            inp(ws, r, 3, None, FMT_INT)
        inp(ws, r, 4, None, FMT_DATE)
        calc(ws, r, 5,
             f'=IF($D{r}<>"",$D{r},IF(OR($B{r}="",$C{r}="",CTR_EARLIEST=""),"",'
             f'CTR_EARLIEST+$C{r}))', FMT_DATE)
        inp(ws, r, 6, None, FMT_DATE)
        inp(ws, r, 7, None)
    ws.cell(row=MS_TOP, column=4).comment = Comment(
        "Type a date here to override the offset calculation for that milestone.", "Workbook")

    # 6 - SYSTEMS -----------------------------------------------------------
    section(ws, BAN["sys"], 2, 13, "6 - SYSTEMS")
    header(ws, SYS_HDR, 2, ["System ID", "System name", "Commissioning Engineer",
                            "Earliest equipment start", "Equip.", "Docs", "Done",
                            "Late", "At risk", "Core missing", "% done", "Status"])
    inputs(ws, 2, 4, SYS_TOP, SYS_BOT)
    for r in range(SYS_TOP, SYS_BOT + 1):
        # earliest equipment start, without MINIFS (Excel 2016 compatible)
        calc(ws, r, 5,
             f'=IF($B{r}="","",IF(COUNTIFS(EQP_SYS,$B{r},EQP_EFF,">0")=0,"",'
             f'SUMPRODUCT(MIN((EQP_SYS=$B{r})*(EQP_EFF>0)*EQP_EFF'
             f'+((EQP_SYS=$B{r})*(EQP_EFF>0)=0)*2958465))))', FMT_DATE)
        calc(ws, r, 6, f'=IF($B{r}="","",COUNTIF(EQP_SYS,$B{r}))', FMT_INT)
        calc(ws, r, 7, f'=IF($B{r}="","",COUNTIFS(DOC_SYS,$B{r},DOC_ACTIVE,1))', FMT_INT)
        calc(ws, r, 8, f'=IF($B{r}="","",COUNTIFS(DOC_SYS,$B{r},DOC_DONE,1))', FMT_INT)
        calc(ws, r, 9, f'=IF($B{r}="","",COUNTIFS(DOC_SYS,$B{r},DOC_STATE,S_LATE))', FMT_INT)
        calc(ws, r, 10, f'=IF($B{r}="","",COUNTIFS(DOC_SYS,$B{r},DOC_STATE,S_RISK))', FMT_INT)
        calc(ws, r, 11,
             f'=IF($B{r}="","",COUNTIFS(GAP_WHERE,$B{r},GAP_KIND,"System",GAP_KEY,">0"))',
             FMT_INT,
             note="Core (ATS) deliverables that do not exist as a row for this system. "
                  "Only rows carrying a gap key count: a covered slot leaves the key empty.")
        calc(ws, r, 12, f'=IF($B{r}="","",IF($G{r}=0,"",$H{r}/$G{r}))', FMT_PCT)
        calc(ws, r, 13,
             f'=IF($B{r}="","",IF($I{r}>0,S_LATE,IF(OR($J{r}>0,$K{r}>0),S_RISK,'
             f'IF(AND($G{r}>0,$H{r}=$G{r}),S_DONE,S_WIP))))')
    databar(ws, f"L{SYS_TOP}:L{SYS_BOT}", "63BE7B")

    # 7 - EQUIPMENT ---------------------------------------------------------
    section(ws, BAN["eqp"], 2, 14,
            "7 - EQUIPMENT      the planned commissioning start typed here drives "
            "every target date behind that equipment")
    header(ws, EQP_HDR, 2, ["Parent system", "Equipment", "Planned comm. start",
                            "Planned comm. finish", "Actual start", "Actual finish",
                            "Docs", "Done", "Late", "Core missing", "% done",
                            "Status", "Key", "Start used"])
    inputs(ws, 2, 7, EQP_TOP, EQP_BOT,
           fmts={4: FMT_DATE, 5: FMT_DATE, 6: FMT_DATE, 7: FMT_DATE})
    for r in range(EQP_TOP, EQP_BOT + 1):
        # the single date every target date behind this equipment is measured from:
        # the actual start once it is known, the planned start until then, 0 if neither
        calc(ws, r, 15,
             f'=IF($C{r}="",0,IF($F{r}<>"",$F{r},IF($D{r}<>"",$D{r},0)))', FMT_DATE0)
        calc(ws, r, 8, f'=IF($C{r}="","",COUNTIFS(DOC_EQP,$C{r},DOC_ACTIVE,1))', FMT_INT)
        calc(ws, r, 9, f'=IF($C{r}="","",COUNTIFS(DOC_EQP,$C{r},DOC_DONE,1))', FMT_INT)
        calc(ws, r, 10, f'=IF($C{r}="","",COUNTIFS(DOC_EQP,$C{r},DOC_STATE,S_LATE))', FMT_INT)
        calc(ws, r, 11,
             f'=IF($C{r}="","",COUNTIFS(GAP_WHERE,$C{r},GAP_KIND,"Equipment",GAP_KEY,">0"))',
             FMT_INT)
        calc(ws, r, 12, f'=IF($C{r}="","",IF($H{r}=0,"",$I{r}/$H{r}))', FMT_PCT)
        calc(ws, r, 13,
             f'=IF($C{r}="","",IF($J{r}>0,S_LATE,IF($K{r}>0,S_RISK,'
             f'IF(AND($H{r}>0,$I{r}=$H{r}),S_DONE,S_WIP))))')
        # severity key: late count dominates, then core gaps, then lack of progress
        calc(ws, r, 14,
             f'=IF($C{r}="","",$J{r}*1000+$K{r}*10+IF($H{r}=0,0,(1-$I{r}/$H{r}))'
             f'+ROW()/1000000)')
    databar(ws, f"L{EQP_TOP}:L{EQP_BOT}", "63BE7B")
    ws.column_dimensions["N"].hidden = True
    ws.column_dimensions["O"].hidden = True
    ws.cell(row=EQP_HDR, column=6).comment = Comment(
        "Fill this in when the equipment actually starts commissioning. From that moment "
        "every target date behind this equipment is measured from the actual start "
        "instead of the planned one.", "Workbook")

    # 8 - DOCUMENT TYPE LIBRARY ---------------------------------------------
    section(ws, BAN["typ"], 2, 9, "8 - DOCUMENT TYPE LIBRARY")
    header(ws, TYP_HDR, 2, ["Document type", "Phase", "Core required (ATS)",
                            "Anchor milestone", "Lead before anchor (days)",
                            "Default scope", "Note", "Core key"])
    for i, (t, ph, core, anch, lead, scope, note) in enumerate(TYPES):
        r = TYP_TOP + i
        inp(ws, r, 2, t, small=True)
        inp(ws, r, 3, ph, small=True)
        inp(ws, r, 4, core, small=True)
        inp(ws, r, 5, anch, small=True)
        inp(ws, r, 6, lead, FMT_INT, small=True)
        inp(ws, r, 7, scope, small=True)
        c = ws.cell(row=r, column=8, value=note); c.font = f_note; c.border = box
        calc(ws, r, 9, f'=IF(AND($B{r}<>"",$D{r}="Yes"),ROW()-{TYP_TOP - 1},"")')
    ws.column_dimensions["I"].hidden = True
    ws.cell(row=TYP_BOT + 1, column=2, value=(
        "Months are converted at 30 days each (T-3 months = 90), which reproduces the worked "
        "example in the source specification exactly. Post-commissioning types are anchored to "
        "a named milestone rather than to a guessed number of days.")).font = f_note

    # 9 - PHASES ------------------------------------------------------------
    section(ws, BAN["phase"], 2, 5, "9 - PHASES", alt=True)
    header(ws, PH_HDR, 2, ["Phase", "Order"])
    for i, (p, o) in enumerate(PHASES):
        r = PH_TOP + i
        c = ws.cell(row=r, column=2, value=p); c.font = f_body; c.border = box
        c = ws.cell(row=r, column=3, value=o); c.font = f_body; c.border = box

    # 10 - PEOPLE -----------------------------------------------------------
    section(ws, BAN["ppl"], 2, 5,
            "10 - PEOPLE      feeds the Responsible, Approver and Owner dropdowns", alt=True)
    header(ws, PPL_HDR, 2, ["Name", "Role"])
    inputs(ws, 2, 3, PPL_TOP, PPL_BOT)

    dv(ws, "SCOPE_LIST", f"G{TYP_TOP}:G{TYP_BOT}")
    dv(ws, "PHASE_LIST", f"C{TYP_TOP}:C{TYP_BOT}")
    dv(ws, "YESNO", f"D{TYP_TOP}:D{TYP_BOT}")
    dv(ws, "MS_NOM", f"E{TYP_TOP}:E{TYP_BOT}")
    dv(ws, "PPL_NAME", f"D{SYS_TOP}:D{SYS_BOT}")
    dv(ws, "SYS_ID", f"B{EQP_TOP}:B{EQP_BOT}")
    dv(ws, "COUNTS_AS", f"D{WF_TOP}:D{WF_BOT}")
    ws.freeze_panes = "A4"

    # Every blue cell stays editable; what is blocked is inserting or deleting
    # rows and columns, which is what silently breaks the named ranges the rest
    # of the workbook is built on. No password - Review > Unprotect Sheet lifts
    # it for anyone who really means to restructure the sheet.
    ws.protection.sheet = True
    ws.protection.insertRows = True
    ws.protection.deleteRows = True
    ws.protection.insertColumns = True
    ws.protection.deleteColumns = True
    ws.protection.formatCells = False
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.sort = False
    ws.protection.autoFilter = False
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False


# ===========================================================================
# TRACKER  (flat table - one row per document x what it covers)
# ===========================================================================
# One ordered table drives every column letter and index below it - add, remove
# or reorder a column here and every formula, width and hidden-column list in
# build_tracker() follows without hand renumbering.
TRK_FIELDS = [
    ("TYPE", "Type"), ("SCOPE", "Scope"), ("SYS", "System"), ("EQP", "Equipment"),
    ("TCN", "Teamcenter no."), ("REV", "Rev."),
    ("RESP", "Responsible (drafting)"), ("APPR", "Approver"), ("STAT", "Status"),
    ("ASTART", "Actual start"), ("OVR", "Target override"),
    ("TARGET", "Target approval"), ("SRC", "Source"), ("AAPPR", "Actual approval"),
    ("DAYS", "Days late"), ("STATE", "State"), ("CHECK", "Check"), ("CMT", "Comment"),
    # helper columns - not shown, but addressed by name like the rest
    ("SDATE", "Scope date"), ("AUTO", "Auto target"), ("ACTIVE", "Active"),
    ("DONE", "Done"), ("INACT", "Inactive"), ("PHASE", "Phase"),
    ("KLATE", "k late"), ("KWIP", "k wip"), ("KCOME", "k coming"),
]
N_TRK_VIS = 18                                # A..R are shown; the rest are helpers
TRK_COLS = [label for _, label in TRK_FIELDS[:N_TRK_VIS]]
TRK_HELP = [label for label in (l for _, l in TRK_FIELDS[N_TRK_VIS:])]
TRK_IDX = {key: i + 1 for i, (key, _) in enumerate(TRK_FIELDS)}
TRK_COL = {key: get_column_letter(i) for key, i in TRK_IDX.items()}
globals().update({f"K_{key}": letter for key, letter in TRK_COL.items()})
TRK_HELP_C = TRK_IDX["SDATE"]                 # first helper column


def build_tracker(ws):
    ws.sheet_properties.tabColor = C_RED
    ws["A1"] = "DELIVERABLES TRACKER"
    ws["A1"].font = f_title
    ws["A2"] = ("One row per document per thing it covers: the whole contract (Scope = All), a "
                "system, or a piece of equipment. A document covering three systems is three "
                "rows. Fill Actual start when work begins. Days late: positive = late, "
                "negative = float.")
    ws["A2"].font = f_sub
    header(ws, 3, 1, TRK_COLS)
    header(ws, 3, TRK_HELP_C, TRK_HELP)

    I = TRK_IDX
    inputs(ws, I["TYPE"], I["TARGET"] - 1, TRK_TOP, TRK_BOT,
           fmts={I["ASTART"]: FMT_DATE, I["OVR"]: FMT_DATE})
    inputs(ws, I["AAPPR"], I["AAPPR"], TRK_TOP, TRK_BOT, fmts={I["AAPPR"]: FMT_DATE})
    inputs(ws, I["CMT"], I["CMT"], TRK_TOP, TRK_BOT)

    T = TRK_TOP, TRK_BOT
    # --- helper columns ---------------------------------------------------
    # A row "exists" once it has a Type - nothing else is required to compute a
    # target date or a colour. Scope date: always numeric. 0 means "no start
    # date known for what this row covers" - never "" and never a bare INDEX,
    # because INDEX on a blank cell returns 0 and 0-lead is a negative serial,
    # which Excel renders as ###### regardless of column width.
    down(ws, I["SDATE"], *T,
         f'=IF(OR(${K_TYPE}{{r}}="",${K_SCOPE}{{r}}=""),"",'
         f'IF(${K_SCOPE}{{r}}="Equipment",'
         f'N(IFERROR(INDEX(EQP_EFF,MATCH(${K_EQP}{{r}},EQP_NOM,0)),0)),'
         f'IF(${K_SCOPE}{{r}}="System",'
         f'N(IFERROR(INDEX(SYS_EARLIEST,MATCH(${K_SYS}{{r}},SYS_ID,0)),0)),'
         f'IF(${K_SCOPE}{{r}}="All",N(CTR_EARLIEST),0))))', FMT_DATE0)
    down(ws, I["AUTO"], *T,
         f'=IF(${K_TYPE}{{r}}="","",IFERROR('
         f'IF(INDEX(TYP_ANCHOR,MATCH(${K_TYPE}{{r}},TYP_TYPE,0))="Commissioning Start",'
         f'IF(N(${K_SDATE}{{r}})<=0,"",'
         f'${K_SDATE}{{r}}-INDEX(TYP_LEAD,MATCH(${K_TYPE}{{r}},TYP_TYPE,0))),'
         f'IF(N(INDEX(MS_PLAN,MATCH(INDEX(TYP_ANCHOR,MATCH(${K_TYPE}{{r}},TYP_TYPE,0)),'
         f'MS_NOM,0)))<=0,"",'
         f'INDEX(MS_PLAN,MATCH(INDEX(TYP_ANCHOR,MATCH(${K_TYPE}{{r}},TYP_TYPE,0)),MS_NOM,0))'
         f'-INDEX(TYP_LEAD,MATCH(${K_TYPE}{{r}},TYP_TYPE,0)))),""))', FMT_DATE0)
    down(ws, I["INACT"], *T,
         f'=IF(${K_TYPE}{{r}}="",0,'
         f'IF(IFERROR(INDEX(WF_CLASS,MATCH(${K_STAT}{{r}},WF_CODE,0)),"")="Inactive",1,0))',
         FMT_INT)
    down(ws, I["ACTIVE"], *T,
         f'=IF(${K_TYPE}{{r}}="",0,IF(${K_INACT}{{r}}=1,0,1))', FMT_INT)
    down(ws, I["DONE"], *T,
         f'=IF(${K_ACTIVE}{{r}}=0,0,IF(${K_AAPPR}{{r}}<>"",1,'
         f'IF(IFERROR(INDEX(WF_CLASS,MATCH(${K_STAT}{{r}},WF_CODE,0)),"")="Done",1,0)))',
         FMT_INT)
    down(ws, I["PHASE"], *T,
         f'=IF(${K_TYPE}{{r}}="","",'
         f'IFERROR(INDEX(TYP_PHASE,MATCH(${K_TYPE}{{r}},TYP_TYPE,0)),"-"))')
    down(ws, I["KLATE"], *T,
         f'=IF(AND(${K_ACTIVE}{{r}}=1,${K_DONE}{{r}}=0,N(${K_DAYS}{{r}})>0),'
         f'${K_DAYS}{{r}}+ROW()/1000000,"")')
    down(ws, I["KWIP"], *T,
         f'=IF(AND(${K_ACTIVE}{{r}}=1,${K_DONE}{{r}}=0,'
         f'OR(${K_ASTART}{{r}}<>"",N(${K_DAYS}{{r}})>0)),N(${K_DAYS}{{r}})+ROW()/1000000,"")')
    down(ws, I["KCOME"], *T,
         f'=IF(AND(${K_ACTIVE}{{r}}=1,${K_DONE}{{r}}=0,${K_ASTART}{{r}}="",'
         f'N(${K_DAYS}{{r}})<=0,-N(${K_DAYS}{{r}})<=HORIZON),'
         f'N(${K_DAYS}{{r}})+ROW()/1000000,"")')

    # --- visible calculated columns ---------------------------------------
    # the N()<=0 guard is what keeps a nonsense date out of a visible date cell
    down(ws, I["TARGET"], *T,
         f'=IF(${K_TYPE}{{r}}="","",IF(${K_OVR}{{r}}<>"",${K_OVR}{{r}},'
         f'IF(N(${K_AUTO}{{r}})<=0,"",${K_AUTO}{{r}})))', FMT_DATE)
    down(ws, I["SRC"], *T,
         f'=IF(${K_TYPE}{{r}}="","",'
         f'IF(${K_OVR}{{r}}<>"","Manual",IF(N(${K_AUTO}{{r}})<=0,"-","Auto")))')
    down(ws, I["DAYS"], *T,
         f'=IF(${K_ACTIVE}{{r}}=0,"",IF(${K_DONE}{{r}}=1,'
         f'IF(${K_AAPPR}{{r}}="","",${K_AAPPR}{{r}}-${K_TARGET}{{r}}),'
         f'IF(${K_TARGET}{{r}}="","",TODAY()-${K_TARGET}{{r}})))', FMT_DAYS)
    down(ws, I["STATE"], *T,
         f'=IF(${K_TYPE}{{r}}="","",IF(${K_INACT}{{r}}=1,S_INACT,'
         f'IF(ISNA(MATCH(${K_TYPE}{{r}},TYP_TYPE,0)),S_TYPEKO,'
         f'IF(${K_DONE}{{r}}=1,S_DONE,'
         f'IF(AND(${K_TARGET}{{r}}="",N(${K_SDATE}{{r}})<=0),S_NOSCOPE,'
         f'IF(${K_TARGET}{{r}}="",S_NOTARGET,'
         f'IF(${K_DAYS}{{r}}>0,S_LATE,'
         f'IF(${K_ASTART}{{r}}<>"",S_WIP,'
         f'IF(-${K_DAYS}{{r}}<=FEN_DOC,S_RISK,'
         f'IF(-${K_DAYS}{{r}}<=HORIZON,S_COMING,S_NOTYET))))))))))')
    down(ws, I["CHECK"], *T,
         f'=IF(${K_TYPE}{{r}}="","",IF(${K_INACT}{{r}}=1,"",'
         f'IF(ISNA(MATCH(${K_TYPE}{{r}},TYP_TYPE,0)),"Type not in the library",'
         f'IF(${K_SCOPE}{{r}}="","Scope not set",'
         f'IF(AND(${K_SCOPE}{{r}}="System",${K_SYS}{{r}}=""),"System not set",'
         f'IF(AND(${K_SCOPE}{{r}}="Equipment",${K_EQP}{{r}}=""),"Equipment not set",'
         f'IF(${K_RESP}{{r}}="","Responsible not set",'
         f'IF(${K_TARGET}{{r}}="","No commissioning start behind this row",'
         f'IF(AND(${K_SCOPE}{{r}}="Equipment",${K_SYS}{{r}}<>"",'
         f'IFERROR(INDEX(EQP_SYS,MATCH(${K_EQP}{{r}},EQP_NOM,0)),"")<>${K_SYS}{{r}}),'
         f'"Equipment not in that system",'
         f'IF(AND(${K_DONE}{{r}}=1,${K_AAPPR}{{r}}=""),'
         f'"Approved with no approval date",""))))))))))')

    ws.cell(row=3, column=I["TCN"]).comment = Comment(
        "Teamcenter identifier of the document. Optional - a row lights up and counts "
        "towards every total as soon as Type is filled in, whether or not this is set yet.",
        "Workbook")
    ws.cell(row=3, column=I["REV"]).comment = Comment(
        "Revision of the document as it stands today - A, B, 01, 02, whatever your "
        "numbering uses. Kept in its own column so it never has to be retyped into the "
        "Teamcenter number when the revision moves on.", "Workbook")
    ws.cell(row=3, column=I["OVR"]).comment = Comment(
        "Type a date here only when the automatic calculation cannot know something - a "
        "cross-system dependency, a site event. The Source column then reads Manual. "
        "Clear this cell to fall back to the automatic date.", "Workbook")
    ws.cell(row=3, column=I["TARGET"]).comment = Comment(
        "Automatic = the commissioning start of whatever this row covers (actual start if "
        "one is filled in on SETUP, otherwise the planned start), minus the lead time of "
        "the type. Post-commissioning types are measured back from their anchor milestone "
        "instead. Blank means no start date is set for what this row covers.", "Workbook")

    widths(ws, {"A": 34, "B": 11, "C": 12, "D": 26, "E": 18, "F": 7,
                "G": 18, "H": 16, "I": 24, "J": 13, "K": 16, "L": 16, "M": 9,
                "N": 16, "O": 10, "P": 15, "Q": 24, "R": 30,
                "S": 12, "T": 12, "U": 8, "V": 8, "W": 9, "X": 16,
                "Y": 11, "Z": 11, "AA": 11})
    for key in ("SDATE", "AUTO", "ACTIVE", "DONE", "INACT", "PHASE",
                "KLATE", "KWIP", "KCOME"):
        ws.column_dimensions[TRK_COL[key]].hidden = True

    dv(ws, "TYP_TYPE", f"{K_TYPE}{TRK_TOP}:{K_TYPE}{TRK_BOT}")
    dv(ws, "SCOPE_LIST", f"{K_SCOPE}{TRK_TOP}:{K_SCOPE}{TRK_BOT}")
    dv(ws, "SYS_ID", f"{K_SYS}{TRK_TOP}:{K_SYS}{TRK_BOT}")
    dv(ws, "EQP_NOM", f"{K_EQP}{TRK_TOP}:{K_EQP}{TRK_BOT}")
    dv(ws, "PPL_NAME", f"{K_RESP}{TRK_TOP}:{K_RESP}{TRK_BOT}")
    dv(ws, "PPL_NAME", f"{K_APPR}{TRK_TOP}:{K_APPR}{TRK_BOT}")
    dv(ws, "WF_CODE", f"{K_STAT}{TRK_TOP}:{K_STAT}{TRK_BOT}")

    t = TRK_TOP
    rng_all = f"{K_TYPE}{TRK_TOP}:{K_CMT}{TRK_BOT}"
    # greying of cells that do not apply to this row's scope
    ws.conditional_formatting.add(f"{K_SYS}{TRK_TOP}:{K_SYS}{TRK_BOT}", FormulaRule(
        formula=[f'OR(${K_SCOPE}{t}="All",${K_SCOPE}{t}="")'],
        fill=fill_grey, stopIfTrue=True))
    ws.conditional_formatting.add(f"{K_EQP}{TRK_TOP}:{K_EQP}{TRK_BOT}", FormulaRule(
        formula=[f'${K_SCOPE}{t}<>"Equipment"'], fill=fill_grey, stopIfTrue=True))
    # empty rows
    ws.conditional_formatting.add(rng_all, FormulaRule(
        formula=[f'${K_TYPE}{t}=""'], fill=PatternFill("solid", fgColor="FAFAFA"),
        stopIfTrue=True))
    # state colours
    for expr, colour, white in [
        (f'${K_STATE}{t}=S_LATE', C_RED_F, False),
        (f'${K_STATE}{t}=S_RISK', C_AMB_F, False),
        (f'${K_STATE}{t}=S_WIP', C_BLU_F, False),
        (f'${K_STATE}{t}=S_COMING', "FFF2CC", False),
        (f'${K_STATE}{t}=S_DONE', C_GRN_F, False),
        (f'${K_STATE}{t}=S_INACT', C_GRY_F, False),
        (f'OR(${K_STATE}{t}=S_NOSCOPE,${K_STATE}{t}=S_NOTARGET,'
         f'${K_STATE}{t}=S_TYPEKO)', C_ORA_F, False),
    ]:
        ws.conditional_formatting.add(rng_all, FormulaRule(
            formula=[f'AND(${K_TYPE}{t}<>"",{expr})'],
            fill=PatternFill("solid", fgColor=colour), stopIfTrue=True))
    # data-quality flag
    ws.conditional_formatting.add(f"{K_CHECK}{TRK_TOP}:{K_CHECK}{TRK_BOT}", FormulaRule(
        formula=[f'${K_CHECK}{t}<>""'],
        fill=PatternFill("solid", fgColor=C_ORA_F),
        font=Font(name=FONT, size=9, bold=True, color="843C0C")))

    ws.auto_filter.ref = f"{K_TYPE}3:{K_CMT}{TRK_BOT}"
    ws.freeze_panes = "E4"


# ===========================================================================
# ACTIONS
# ===========================================================================
ACT_COLS = ["Action ID", "Description", "Phase", "System", "Equipment",
            "Teamcenter no.", "Owner", "Actual start", "Due date", "Done date",
            "Status", "Days late", "State", "Comment"]


def build_actions(ws):
    ws.sheet_properties.tabColor = "BF8F00"
    ws["A1"] = "ACTIONS"
    ws["A1"].font = f_title
    ws["A2"] = ("One filterable table. Link an action to a system, an equipment item, a "
                "document, any combination, or nothing at all - every link is optional.")
    ws["A2"].font = f_sub
    header(ws, 3, 1, ACT_COLS)
    header(ws, 3, 16, ["k late", "k wip", "k coming"])

    inputs(ws, 1, 11, ACT_TOP, ACT_BOT, fmts={8: FMT_DATE, 9: FMT_DATE, 10: FMT_DATE})
    inputs(ws, 14, 14, ACT_TOP, ACT_BOT)

    A = ACT_TOP, ACT_BOT
    down(ws, 12, *A,
         '=IF($A{r}="","",IF($K{r}=A_DONE,IF($J{r}="","",$J{r}-$I{r}),'
         'IF($I{r}="","",TODAY()-$I{r})))', FMT_DAYS)
    down(ws, 13, *A,
         '=IF($A{r}="","",IF($K{r}=A_DONE,S_DONE,IF($I{r}="",S_NODUE,'
         'IF($L{r}>0,S_LATE,IF($H{r}<>"",S_WIP,'
         'IF(-$L{r}<=FEN_ACT,S_RISK,IF(-$L{r}<=HORIZON,S_COMING,S_ONTRACK)))))))')
    down(ws, 16, *A,
         '=IF(AND($A{r}<>"",$K{r}<>A_DONE,N($L{r})>0),$L{r}+ROW()/1000000,"")')
    down(ws, 17, *A,
         '=IF(AND($A{r}<>"",$K{r}<>A_DONE,OR($H{r}<>"",N($L{r})>0)),'
         'N($L{r})+ROW()/1000000,"")')
    down(ws, 18, *A,
         '=IF(AND($A{r}<>"",$K{r}<>A_DONE,$H{r}="",$I{r}<>"",N($L{r})<=0,'
         '-N($L{r})<=HORIZON),N($L{r})+ROW()/1000000,"")')

    widths(ws, {"A": 11, "B": 52, "C": 18, "D": 12, "E": 26, "F": 20, "G": 16,
                "H": 13, "I": 13, "J": 13, "K": 14, "L": 11, "M": 15, "N": 30,
                "P": 11, "Q": 11, "R": 11})
    for col in ("P", "Q", "R"):
        ws.column_dimensions[col].hidden = True

    dv(ws, "PHASE_LIST", f"C{ACT_TOP}:C{ACT_BOT}")
    dv(ws, "SYS_ID", f"D{ACT_TOP}:D{ACT_BOT}")
    dv(ws, "EQP_NOM", f"E{ACT_TOP}:E{ACT_BOT}")
    dv(ws, "DOC_NUM", f"F{ACT_TOP}:F{ACT_BOT}")
    dv(ws, "PPL_NAME", f"G{ACT_TOP}:G{ACT_BOT}")
    dv(ws, "ACT_STATUS_LIST", f"K{ACT_TOP}:K{ACT_BOT}")

    a = ACT_TOP
    rng = f"A{ACT_TOP}:N{ACT_BOT}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'$A{a}=""'], fill=PatternFill("solid", fgColor="FAFAFA"), stopIfTrue=True))
    for expr, colour in [
        (f'AND($M{a}=S_LATE,$K{a}="Critical")', C_RED),
        (f'$M{a}=S_LATE', C_RED_F),
        (f'$M{a}=S_RISK', C_AMB_F),
        (f'$M{a}=S_WIP', C_BLU_F),
        (f'$M{a}=S_COMING', "FFF2CC"),
        (f'$M{a}=S_DONE', C_GRN_F),
        (f'$K{a}=A_HOLD', C_GRY_F),
        (f'$M{a}=S_NODUE', C_ORA_F),
    ]:
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'AND($A{a}<>"",{expr})'],
            fill=PatternFill("solid", fgColor=colour), stopIfTrue=True))

    ws.auto_filter.ref = f"A3:N{ACT_BOT}"
    ws.freeze_panes = "C4"


# ===========================================================================
# LISTS  (hidden staging: core slots, gap matrix, merged list, chart data)
# ===========================================================================
DOC_STATE_KEYS = ["S_DONE", "S_WIP", "S_LATE", "S_RISK", "S_COMING",
                  "S_NOTYET", "S_INACT", "S_NOTARGET"]


def build_lists(ws):
    ws.sheet_properties.tabColor = "808080"
    ws["A1"] = "LISTS - staging. Generated automatically; do not edit."
    ws["A1"].font = f_lbl

    # --- core type slots ---------------------------------------------------
    for k in range(N_CORE):
        r = CORE_TOP + k
        ws.cell(row=r, column=1,
                value=f'=IFERROR(INDEX(TYP_TYPE,SMALL(TYP_COREKEY,{k + 1})),"")')
        ws.cell(row=r, column=2,
                value=f'=IF($A{r}="","",INDEX(TYP_SCOPE,MATCH($A{r},TYP_TYPE,0)))')
        ws.cell(row=r, column=3,
                value=f'=IF($A{r}="","",INDEX(TYP_LEAD,MATCH($A{r},TYP_TYPE,0)))')
        ws.cell(row=r, column=4,
                value=f'=IF($A{r}="","",INDEX(TYP_ANCHOR,MATCH($A{r},TYP_TYPE,0)))')
        ws.cell(row=r, column=5,
                value=f'=IF($A{r}="","",IF($D{r}="Commissioning Start","",'
                      f'IFERROR(INDEX(MS_PLAN,MATCH($D{r},MS_NOM,0)),"")))')

    # --- gap matrix --------------------------------------------------------
    r = GAP_TOP
    for k in range(N_CORE):                                   # contract level
        cr = CORE_TOP + k
        ws.cell(row=r, column=7,
                value=f'=IF(OR($A${cr}="",$B${cr}<>"All"),"","Contract")')
        ws.cell(row=r, column=8, value=f'=IF($G{r}="","","Whole contract")')
        ws.cell(row=r, column=9, value=f'=IF($G{r}="","",$A${cr})')
        ws.cell(row=r, column=10, value=(
            f'=IF($G{r}="","",IF(COUNTIFS(DOC_TYPE,$I{r},DOC_ACTIVE,1)>0,"",'
            f'IFERROR(IF($E${cr}<>"",$E${cr}-$C${cr},CTR_EARLIEST-$C${cr}),60000)'
            f'+ROW()/1000000))'))
        r += 1
    for i in range(N_SYS):                                    # system level
        sr = SYS_TOP + i
        for k in range(N_CORE):
            cr = CORE_TOP + k
            ws.cell(row=r, column=7, value=(
                f'=IF(OR(SETUP!$B${sr}="",$A${cr}="",$B${cr}<>"System"),"","System")'))
            ws.cell(row=r, column=8, value=f'=IF($G{r}="","",SETUP!$B${sr})')
            ws.cell(row=r, column=9, value=f'=IF($G{r}="","",$A${cr})')
            ws.cell(row=r, column=10, value=(
                f'=IF($G{r}="","",IF(COUNTIFS(DOC_TYPE,$I{r},DOC_SYS,$H{r},'
                f'DOC_SCOPE,"System",DOC_ACTIVE,1)'
                f'+COUNTIFS(DOC_TYPE,$I{r},DOC_SCOPE,"All",DOC_ACTIVE,1)>0,"",'
                f'IFERROR(IF($E${cr}<>"",$E${cr}-$C${cr},SETUP!$E${sr}-$C${cr}),60000)'
                f'+ROW()/1000000))'))
            r += 1
    for j in range(N_EQP):                                    # equipment level
        er = EQP_TOP + j
        for k in range(N_CORE):
            cr = CORE_TOP + k
            ws.cell(row=r, column=7, value=(
                f'=IF(OR(SETUP!$C${er}="",$A${cr}="",$B${cr}<>"Equipment"),"","Equipment")'))
            ws.cell(row=r, column=8, value=f'=IF($G{r}="","",SETUP!$C${er})')
            ws.cell(row=r, column=9, value=f'=IF($G{r}="","",$A${cr})')
            ws.cell(row=r, column=10, value=(
                f'=IF($G{r}="","",IF(COUNTIFS(DOC_TYPE,$I{r},DOC_EQP,$H{r},'
                f'DOC_SCOPE,"Equipment",DOC_ACTIVE,1)'
                f'+COUNTIFS(DOC_TYPE,$I{r},DOC_SYS,SETUP!$B${er},DOC_SCOPE,"System",'
                f'DOC_ACTIVE,1)'
                f'+COUNTIFS(DOC_TYPE,$I{r},DOC_SCOPE,"All",DOC_ACTIVE,1)>0,"",'
                f'IFERROR(IF($E${cr}<>"",$E${cr}-$C${cr},SETUP!$D${er}-$C${cr}),60000)'
                f'+ROW()/1000000))'))
            r += 1

    # --- merged document + action list ------------------------------------
    for i in range(N_TRK):
        r = CMB_TOP + i
        tr = TRK_TOP + i
        ws.cell(row=r, column=12, value=f'=IF({TR}!${K_TYPE}{tr}="","","Document")')
        ws.cell(row=r, column=13, value=(
            f'=IF($L{r}="","",IF({TR}!${K_TCN}{tr}<>"",{TR}!${K_TCN}{tr}'
            f'&IF({TR}!${K_REV}{tr}="",""," rev "&{TR}!${K_REV}{tr})&"  -  ","")'
            f'&{TR}!${K_TYPE}{tr})'))
        ws.cell(row=r, column=14, value=(
            f'=IF($L{r}="","",IF({TR}!${K_SCOPE}{tr}="Equipment",{TR}!${K_EQP}{tr},'
            f'IF({TR}!${K_SCOPE}{tr}="System",{TR}!${K_SYS}{tr},"Whole contract")))'))
        ws.cell(row=r, column=15, value=f'=IF($L{r}="","",{TR}!${K_RESP}{tr})')
        ws.cell(row=r, column=16, value=f'=IF($L{r}="","",{TR}!${K_TARGET}{tr})')
        ws.cell(row=r, column=17, value=f'=IF($L{r}="","",{TR}!${K_DAYS}{tr})')
        ws.cell(row=r, column=18, value=f'={TR}!${K_KWIP}{tr}')
        ws.cell(row=r, column=19, value=f'={TR}!${K_KCOME}{tr}')
        ws.cell(row=r, column=20, value=f'={TR}!${K_KLATE}{tr}')
    for j in range(N_ACT):
        r = CMB_TOP + N_TRK + j
        ar = ACT_TOP + j
        ws.cell(row=r, column=12, value=f'=IF(ACTIONS!$A{ar}="","","Action")')
        ws.cell(row=r, column=13, value=(
            f'=IF($L{r}="","",ACTIONS!$A{ar}&"  -  "&LEFT(ACTIONS!$B{ar},60))'))
        ws.cell(row=r, column=14, value=(
            f'=IF($L{r}="","",IF(ACTIONS!$E{ar}<>"",ACTIONS!$E{ar},'
            f'IF(ACTIONS!$D{ar}<>"",ACTIONS!$D{ar},"-")))'))
        ws.cell(row=r, column=15, value=f'=IF($L{r}="","",ACTIONS!$G{ar})')
        ws.cell(row=r, column=16, value=f'=IF($L{r}="","",ACTIONS!$I{ar})')
        ws.cell(row=r, column=17, value=f'=IF($L{r}="","",ACTIONS!$L{ar})')
        ws.cell(row=r, column=18, value=f'=ACTIONS!$Q{ar}')
        ws.cell(row=r, column=19, value=f'=ACTIONS!$R{ar}')
        ws.cell(row=r, column=20, value=f'=ACTIONS!$P{ar}')

    # --- chart data --------------------------------------------------------
    for i, key in enumerate(DOC_STATE_KEYS):                  # V / W
        r = DST_TOP + i
        ws.cell(row=r, column=22, value=f'={key}')
        ws.cell(row=r, column=23, value=f'=COUNTIF(DOC_STATE,$V{r})')
    for i, key in enumerate(["A_OPEN", "A_WIP", "A_HOLD", "A_DONE"]):   # Y / Z
        r = AST2_TOP + i
        ws.cell(row=r, column=25, value=f'={key}')
        ws.cell(row=r, column=26, value=f'=COUNTIF(ACT_STATUS,$Y{r})')
    for i in range(N_SYS):                                    # AB / AC / AD
        r = BYS_TOP + i
        sr = SYS_TOP + i
        ws.cell(row=r, column=28, value=f'=IF(SETUP!$B${sr}="","",SETUP!$B${sr})')
        ws.cell(row=r, column=29, value=f'=IF($AB{r}="",0,SETUP!$I${sr})')
        ws.cell(row=r, column=30, value=f'=IF($AB{r}="",0,SETUP!$J${sr})')
    for i in range(12):                                       # AF / AG
        r = MON_TOP + i
        ws.cell(row=r, column=32, value=f'=TEXT(EOMONTH(TODAY(),{i})," MMM YY")')
        ws.cell(row=r, column=33, value=(
            f'=COUNTIFS(DOC_TARGET,">="&EOMONTH(TODAY(),{i - 1})+1,'
            f'DOC_TARGET,"<="&EOMONTH(TODAY(),{i}),DOC_DONE,0)'))

    ws.sheet_state = "hidden"


# ===========================================================================
# DASHBOARD
# ===========================================================================
D_RAG = 6
D_PH, D_PD = 13, 14
L1_B, L1_H, L1_T = 17, 18, 19
L1_E = L1_T + LIST_ROWS - 1                    # 28
L2_B, L2_H, L2_T = 30, 31, 32
L2_E = L2_T + LIST_ROWS - 1                    # 41
L3_B, L3_H, L3_T = 43, 44, 45
L3_E = L3_T + LIST_ROWS - 1                    # 54
SY_B, SY_H, SY_T = 56, 57, 58
SY_E = SY_T + N_SYS - 1                        # 82
EQ_B, EQ_H, EQ_T = 84, 85, 86
N_EQ_WATCH = 15
EQ_E = EQ_T + N_EQ_WATCH - 1                   # 100
MS_B, MS_H, MS_T = 102, 103, 104
MS_E = MS_T + N_MS_SLOT - 1                    # 118
CH_B = 120
HELP = 18                                      # column R holds the ranking keys


def _pick(ws, r, rng_name, key_col, cols):
    """INDEX/MATCH a ranked row across several source ranges."""
    for col, src in cols:
        ws.cell(row=r, column=col, value=(
            f'=IF(${key_col}{r}="","",IFERROR(INDEX({src},'
            f'MATCH(${key_col}{r},{rng_name},0)),""))'))


def build_dashboard(ws):
    ws.sheet_properties.tabColor = "375623"
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 2, "B": 26, "C": 46, "D": 24, "E": 16, "F": 14, "G": 13,
                "H": 13, "I": 13, "J": 13, "K": 16, "L": 12, "M": 12, "R": 12})
    ws.column_dimensions["R"].hidden = True

    ws["B1"] = "COMMISSIONING DASHBOARD"
    ws["B1"].font = f_title
    ws["B2"] = ('=IF(CTR_NAME="","Set the contract name in SETUP",CTR_NAME&"   -   "&CTR_NUM)'
                '&"     |     what is late, what has never been created, what is coming."')
    ws["B2"].font = f_sub

    # Integrity banner. Every list on this sheet is wrapped in IFERROR, so a
    # broken name would otherwise show as an empty list rather than as a fault.
    ws.merge_cells("B3:M3")
    warn = ws["B3"]
    warn.value = (
        '=IF(ISERROR(S_DONE&S_LATE&S_WIP&S_RISK&S_COMING&S_NOTYET&S_INACT'
        '&S_NOSCOPE&S_NOTARGET&S_TYPEKO&S_NODUE&S_ONTRACK'
        '&A_OPEN&A_WIP&A_HOLD&A_DONE&RAG_RED&RAG_AMB&RAG_GRN'
        '&N(FEN_DOC)&N(FEN_ACT)&N(HORIZON)),'
        '"SETUP IS DAMAGED - a row was deleted on the SETUP sheet, so the wording or '
        'threshold cells the whole workbook reads have been lost. The lists below are '
        'blank because of it, not because there is nothing to show. Undo the deletion '
        '(Ctrl+Z), or reopen the last saved copy.","")')
    warn.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
    warn.alignment = a_l
    ws.conditional_formatting.add("B3:M3", FormulaRule(
        formula=['$B$3<>""'], fill=PatternFill("solid", fgColor=C_RED),
        stopIfTrue=True))

    # --- contract status ---------------------------------------------------
    section(ws, 4, 2, 13, "CONTRACT STATUS")
    ws.merge_cells(f"B{D_RAG}:D{D_RAG + 4}")
    tile = ws[f"B{D_RAG}"]
    tile.value = ('=IF(OR(COUNTIF(DOC_STATE,S_LATE)>0,COUNTIF(ACT_STATE,S_LATE)>0),RAG_RED,'
                  'IF(OR(COUNTIF(DOC_STATE,S_RISK)>0,COUNTIF(ACT_STATE,S_RISK)>0,'
                  'COUNT(GAP_KEY)>0),RAG_AMB,RAG_GRN))')
    tile.font = f_rag
    tile.alignment = a_c
    for rr in range(D_RAG, D_RAG + 5):
        for cc in range(2, 5):
            ws.cell(row=rr, column=cc).border = box
    ws.conditional_formatting.add(f"B{D_RAG}:D{D_RAG + 4}", FormulaRule(
        formula=[f'$B${D_RAG}=RAG_RED'],
        fill=PatternFill("solid", fgColor=C_RED), stopIfTrue=True))
    ws.conditional_formatting.add(f"B{D_RAG}:D{D_RAG + 4}", FormulaRule(
        formula=[f'$B${D_RAG}=RAG_AMB'],
        fill=PatternFill("solid", fgColor="BF8F00"), stopIfTrue=True))
    ws.conditional_formatting.add(f"B{D_RAG}:D{D_RAG + 4}", FormulaRule(
        formula=[f'$B${D_RAG}=RAG_GRN'],
        fill=PatternFill("solid", fgColor="375623"), stopIfTrue=True))

    kpis = [
        ("Worst item", '=IF(COUNT(CMB_KLATE)=0,"nothing late",IFERROR('
                       'INDEX(CMB_SUBJ,MATCH(MAX(CMB_KLATE),CMB_KLATE,0))&"     ("'
                       '&INT(MAX(CMB_KLATE))&" d late  -  "'
                       '&INDEX(CMB_SCOPE,MATCH(MAX(CMB_KLATE),CMB_KLATE,0))&")","-"))', None),
        ("Share late", '=IF(SUM(DOC_ACTIVE)+COUNTIF(ACT_ID,"?*")=0,"-",'
                       '(COUNTIF(DOC_STATE,S_LATE)+COUNTIF(ACT_STATE,S_LATE))'
                       '/(SUM(DOC_ACTIVE)+COUNTIF(ACT_ID,"?*")))', FMT_PCT1),
        ("Earliest commissioning start",
         '=IF(CTR_EARLIEST="","not set yet",TEXT(CTR_EARLIEST,"DD-MMM-YYYY")'
         '&IF(CTR_EARLIEST<TODAY(),"     started "&TODAY()-CTR_EARLIEST&" d ago",'
         '"     in "&CTR_EARLIEST-TODAY()&" d"))', None),
        ("Core deliverables missing", "=COUNT(GAP_KEY)", FMT_INT),
        ("Data quality flags", '=COUNTIF(DOC_CHECK,"?*")', FMT_INT),
    ]
    for i, (name, formula, fmt) in enumerate(kpis):
        r = D_RAG + i
        # label spans E:F - column F alone is sized for the lists' date column
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
        lab(ws, r, 5, name)
        calc(ws, r, 7, formula, fmt, big=True)
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=13)

    # --- progress ----------------------------------------------------------
    section(ws, 12, 2, 9, "PROGRESS")
    header(ws, D_PH, 2, ["", "Done", "In progress", "Late", "Inactive",
                         "Missing", "Total", "% done"])
    rows = [
        ("DOCUMENTS", ['=COUNTIF(DOC_STATE,S_DONE)', '=COUNTIF(DOC_STATE,S_WIP)',
                       '=COUNTIF(DOC_STATE,S_LATE)', '=COUNTIF(DOC_STATE,S_INACT)',
                       '=COUNT(GAP_KEY)', '=SUM(DOC_ACTIVE)']),
        ("ACTIONS", ['=COUNTIF(ACT_STATE,S_DONE)', '=COUNTIF(ACT_STATE,S_WIP)',
                     '=COUNTIF(ACT_STATE,S_LATE)', '=COUNTIF(ACT_STATUS,A_HOLD)',
                     '=COUNTIF(ACT_STATE,S_NODUE)', '=COUNTIF(ACT_ID,"?*")']),
    ]
    for i, (name, fs) in enumerate(rows):
        r = D_PD + i
        c = ws.cell(row=r, column=2, value=name)
        c.font = f_lbl
        c.fill = fill_light
        c.border = box
        for j, f in enumerate(fs):
            calc(ws, r, 3 + j, f, FMT_INT, big=True)
        calc(ws, r, 9, f'=IF($H{r}=0,"",$C{r}/$H{r})', FMT_PCT, big=True)
    databar(ws, f"I{D_PD}:I{D_PD + 1}", "63BE7B")

    # --- 1. IN PROGRESS ----------------------------------------------------
    section(ws, L1_B, 2, 13, "x").value = (
        '="1. IN PROGRESS      work started, or already late.  Most late first.'
        '      Showing "&MIN(' + str(LIST_ROWS) + ',COUNT(CMB_KWIP))&" of "&COUNT(CMB_KWIP)')
    header(ws, L1_H, 2, ["Kind", "Subject", "Scope", "Owner", "Target date", "Days late"])
    for k in range(LIST_ROWS):
        r = L1_T + k
        ws.cell(row=r, column=HELP, value=f'=IFERROR(LARGE(CMB_KWIP,{k + 1}),"")')
        _pick(ws, r, "CMB_KWIP", "R", [(2, "CMB_KIND"), (3, "CMB_SUBJ"),
                                       (4, "CMB_SCOPE"), (5, "CMB_OWNER"),
                                       (6, "CMB_TARGET"), (7, "CMB_DAYS")])
        for cc in range(2, 8):
            cell = ws.cell(row=r, column=cc)
            cell.font = f_calc
            cell.border = box
            cell.alignment = a_l
        ws.cell(row=r, column=6).number_format = FMT_DATE
        ws.cell(row=r, column=7).number_format = FMT_DAYS

    # --- 2. MISSING --------------------------------------------------------
    section(ws, L2_B, 2, 13, "x").value = (
        '="2. MISSING      core deliverables that do not exist yet.  Soonest needed first.'
        '      Showing "&MIN(' + str(LIST_ROWS) + ',COUNT(GAP_KEY))&" of "&COUNT(GAP_KEY)')
    header(ws, L2_H, 2, ["Level", "Deliverable", "Where", "Needed by"])
    for k in range(LIST_ROWS):
        r = L2_T + k
        ws.cell(row=r, column=HELP, value=f'=IFERROR(SMALL(GAP_KEY,{k + 1}),"")')
        _pick(ws, r, "GAP_KEY", "R", [(2, "GAP_KIND"), (3, "GAP_WHAT"), (4, "GAP_WHERE")])
        ws.cell(row=r, column=5, value=(
            f'=IF($R{r}="","",IF(INT($R{r})>=60000,"",INT($R{r})))'))
        for cc in range(2, 6):
            cell = ws.cell(row=r, column=cc)
            cell.font = f_calc
            cell.border = box
            cell.alignment = a_l
        ws.cell(row=r, column=5).number_format = FMT_DATE

    # --- 3. COMING ---------------------------------------------------------
    section(ws, L3_B, 2, 13, "x").value = (
        '="3. COMING      not started, due inside the horizon set in SETUP.  Soonest first.'
        '      Showing "&MIN(' + str(LIST_ROWS) + ',COUNT(CMB_KCOME))&" of "&COUNT(CMB_KCOME)')
    header(ws, L3_H, 2, ["Kind", "Subject", "Scope", "Owner", "Target date", "Days left"])
    for k in range(LIST_ROWS):
        r = L3_T + k
        ws.cell(row=r, column=HELP, value=f'=IFERROR(LARGE(CMB_KCOME,{k + 1}),"")')
        _pick(ws, r, "CMB_KCOME", "R", [(2, "CMB_KIND"), (3, "CMB_SUBJ"),
                                        (4, "CMB_SCOPE"), (5, "CMB_OWNER"),
                                        (6, "CMB_TARGET")])
        ws.cell(row=r, column=7, value=(
            f'=IF($R{r}="","",-INDEX(CMB_DAYS,MATCH($R{r},CMB_KCOME,0)))'))
        for cc in range(2, 8):
            cell = ws.cell(row=r, column=cc)
            cell.font = f_calc
            cell.border = box
            cell.alignment = a_l
        ws.cell(row=r, column=6).number_format = FMT_DATE
        ws.cell(row=r, column=7).number_format = FMT_INT

    # --- 4. BY SYSTEM ------------------------------------------------------
    section(ws, SY_B, 2, 13, "4. BY SYSTEM")
    header(ws, SY_H, 2, ["System", "System name", "Equip.", "Docs", "Done", "Late",
                         "At risk", "Core missing", "% done", "Status"])
    for i in range(N_SYS):
        r, sr = SY_T + i, SYS_TOP + i
        src = [(2, "B"), (3, "C"), (4, "F"), (5, "G"), (6, "H"), (7, "I"),
               (8, "J"), (9, "K"), (10, "L"), (11, "M")]
        for col, sc in src:
            ws.cell(row=r, column=col, value=(
                f'=IF(OR(SETUP!$B${sr}="",SETUP!${sc}${sr}=""),"",SETUP!${sc}${sr})'))
        for cc in range(2, 12):
            cell = ws.cell(row=r, column=cc)
            cell.font = f_calc
            cell.border = box
            cell.alignment = a_l
        ws.cell(row=r, column=10).number_format = FMT_PCT
        for cc in range(4, 10):
            ws.cell(row=r, column=cc).number_format = FMT_INT
    databar(ws, f"J{SY_T}:J{SY_E}", "63BE7B")
    ws.conditional_formatting.add(f"B{SY_T}:K{SY_E}", FormulaRule(
        formula=[f'AND($B{SY_T}<>"",$K{SY_T}=S_LATE)'],
        fill=PatternFill("solid", fgColor=C_RED_F), stopIfTrue=False))

    # --- 5. EQUIPMENT TO WATCH --------------------------------------------
    section(ws, EQ_B, 2, 13, "5. EQUIPMENT TO WATCH      most severe first")
    header(ws, EQ_H, 2, ["System", "Equipment", "Docs", "Done", "Late",
                         "Core missing", "% done", "Status"])
    for k in range(N_EQ_WATCH):
        r = EQ_T + k
        ws.cell(row=r, column=HELP, value=f'=IFERROR(LARGE(EQW_KEY,{k + 1}),"")')
        _pick(ws, r, "EQW_KEY", "R",
              [(2, "EQP_SYS"), (3, "EQP_NOM"), (4, "EQP_DOCS"), (5, "EQP_DONE"),
               (6, "EQP_LATE"), (7, "EQP_CORE"), (8, "EQP_PCT"), (9, "EQP_STATUS")])
        for cc in range(2, 10):
            cell = ws.cell(row=r, column=cc)
            cell.font = f_calc
            cell.border = box
            cell.alignment = a_l
        ws.cell(row=r, column=8).number_format = FMT_PCT
        for cc in range(4, 8):
            ws.cell(row=r, column=cc).number_format = FMT_INT
    databar(ws, f"H{EQ_T}:H{EQ_E}", "63BE7B")

    # --- 6. MILESTONES -----------------------------------------------------
    section(ws, MS_B, 2, 13, "6. MILESTONES")
    header(ws, MS_H, 2, ["Milestone", "Note", "Planned date", "Actual date",
                         "Days from today"])
    for i in range(N_MS_SLOT):
        r, mr = MS_T + i, MS_TOP + i
        for col, sc in [(2, "B"), (3, "G"), (4, "E"), (5, "F")]:
            ws.cell(row=r, column=col, value=(
                f'=IF(OR(SETUP!$B${mr}="",SETUP!${sc}${mr}=""),"",SETUP!${sc}${mr})'))
        ws.cell(row=r, column=6, value=(
            f'=IF(SETUP!$B${mr}="","",IF(SETUP!$F${mr}<>"","",'
            f'IF(SETUP!$E${mr}="","",SETUP!$E${mr}-TODAY())))'))
        for cc in range(2, 7):
            cell = ws.cell(row=r, column=cc)
            cell.font = f_calc
            cell.border = box
            cell.alignment = a_l
        ws.cell(row=r, column=4).number_format = FMT_DATE
        ws.cell(row=r, column=5).number_format = FMT_DATE
        ws.cell(row=r, column=6).number_format = FMT_DAYS
    ws.conditional_formatting.add(f"B{MS_T}:F{MS_E}", FormulaRule(
        formula=[f'AND($B{MS_T}<>"",$E{MS_T}<>"")'],
        fill=PatternFill("solid", fgColor=C_GRN_F), stopIfTrue=True))
    ws.conditional_formatting.add(f"B{MS_T}:F{MS_E}", FormulaRule(
        formula=[f'AND($B{MS_T}<>"",$E{MS_T}="",N($F{MS_T})<0)'],
        fill=PatternFill("solid", fgColor=C_RED_F), stopIfTrue=True))

    section(ws, CH_B, 2, 13, "CHARTS")
    ws.freeze_panes = "A5"


def add_charts(ws, ls):
    def style(ch, title, h=7.0, w=13.0):
        ch.title = title
        ch.height = h
        ch.width = w
        ch.style = 2
        ch.plotVisOnly = False
        return ch

    c1 = style(DoughnutChart(), "Documents by state")
    c1.add_data(Reference(ls, min_col=23, min_row=DST_TOP, max_row=DST_TOP + 7))
    c1.set_categories(Reference(ls, min_col=22, min_row=DST_TOP, max_row=DST_TOP + 7))
    c1.dataLabels = DataLabelList()
    c1.dataLabels.showVal = True
    ws.add_chart(c1, f"B{CH_B + 1}")

    c2 = style(DoughnutChart(), "Actions by status")
    c2.add_data(Reference(ls, min_col=26, min_row=AST2_TOP, max_row=AST2_BOT))
    c2.set_categories(Reference(ls, min_col=25, min_row=AST2_TOP, max_row=AST2_BOT))
    c2.dataLabels = DataLabelList()
    c2.dataLabels.showVal = True
    ws.add_chart(c2, f"H{CH_B + 1}")

    c3 = style(BarChart(), "Late and at risk, by system", h=7.5)
    c3.type = "col"
    c3.grouping = "stacked"
    c3.overlap = 100
    c3.add_data(Reference(ls, min_col=29, max_col=30, min_row=BYS_TOP - 1,
                          max_row=BYS_BOT), titles_from_data=True)
    c3.set_categories(Reference(ls, min_col=28, min_row=BYS_TOP, max_row=BYS_BOT))
    ws.add_chart(c3, f"B{CH_B + 16}")

    c4 = style(BarChart(), "Approval workload - documents due per month", h=7.5)
    c4.type = "col"
    c4.add_data(Reference(ls, min_col=33, min_row=MON_TOP, max_row=MON_BOT))
    c4.set_categories(Reference(ls, min_col=32, min_row=MON_TOP, max_row=MON_BOT))
    c4.legend = None
    ws.add_chart(c4, f"H{CH_B + 16}")


# ===========================================================================
# HOW IT WORKS
# ===========================================================================
def build_how(ws):
    ws.sheet_properties.tabColor = "7F7F7F"
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 2, "B": 34, "C": 104})
    ws["B1"] = "HOW IT WORKS"
    ws["B1"].font = f_title
    ws["B2"] = ("This workbook does not measure progress. It measures whether deliverables "
                "are on time against the commissioning start of whatever they cover.")
    ws["B2"].font = f_sub

    rows = [
        ("SECTION", "THE ONE RULE", None),
        ("Target approval date",
         "The earliest commissioning start among what the document covers "
         "(actual if known, otherwise planned), minus the lead time of its type. "
         "Post-commissioning types are measured back from their anchor milestone instead.", None),
        ("Days late", "Positive = late. Negative = float. One convention everywhere.", None),
        ("Scope", "All = the whole contract. System = one system. Equipment = one item. "
                  "A document covering three systems is three rows.", None),

        ("SECTION", "START HERE", None),
        ("1. Contract and thresholds", "SETUP sections 1 and 2.", None),
        ("2. Systems and equipment", "SETUP sections 6 and 7. The commissioning start you type "
                                     "against each equipment item is what drives every target "
                                     "date behind it. Fill Actual start when the item really "
                                     "starts: from then on every target behind it is measured "
                                     "from the actual date instead of the planned one, and the "
                                     "hidden 'Start used' column shows which of the two is in "
                                     "force. Equipment with neither date leaves its documents "
                                     "with a blank target and the Check column explains why.",
         None),
        ("3. Check the type library", "SETUP section 8. Phase, anchor milestone, lead time, "
                                      "core (ATS) flag and default scope for each type.", None),
        ("4. Enter deliverables", "TRACKER. Pick a Type and a Scope; the target date appears.", None),
        ("5. Work the lists", "DASHBOARD. In progress, Missing, Coming - documents and actions "
                              "together.", None),

        ("SECTION", "READING THE TRACKER", None),
        ("Teamcenter no. / Rev.",
         "Two separate columns on purpose: the revision moves on without the Teamcenter "
         "number ever being retyped. Both are optional metadata - a row is real and "
         "coloured the moment Type is filled in, whether or not either is set yet.", None),
        ("Target override", "Type a date to force a target the calculation cannot know - a "
                            "cross-system dependency, a site event. The Source column then reads "
                            "Manual. Clear the cell to fall back to Auto.", None),
        ("Source", "Auto = calculated. Manual = you overrode it. Nothing is ever silently "
                   "overwritten.", None),
        ("Check", "Data-quality flag: scope not set, equipment not in that system, approved "
                  "with no approval date. Empty means the row is clean.", None),
        ("Greyed cells", "Not relevant for that row's scope. System is greyed for Scope = All; "
                         "Equipment is greyed unless Scope = Equipment.", None),

        ("SECTION", "STATE COLOURS", None),
        ("Red", "Late - the target approval date has passed.", C_RED_F),
        ("Amber", "At risk - due inside the document alert window.", C_AMB_F),
        ("Blue", "In progress - an actual start date is filled in.", C_BLU_F),
        ("Pale amber", "Coming - not started, due inside the horizon.", "FFF2CC"),
        ("Green", "Done - S4A reached, a P-code published, or an approval date entered.", C_GRN_F),
        ("Grey", "Inactive - cancelled, revoked, superseded, transferred, destroyed.", C_GRY_F),
        ("Orange", "Needs attention: scope not resolved, no target date, type not configured, "
                   "or a data-quality flag.", C_ORA_F),

        ("SECTION", "WHAT 'MISSING' MEANS", None),
        ("Missing", "A deliverable flagged Core required (ATS) in the library that does not "
                    "exist as a row at all for a system or an equipment item. The workbook "
                    "builds the expected matrix from the library and reports the gaps, ordered "
                    "by when each would have been needed.", None),
        ("Covered by", "A system-wide or contract-wide document counts as coverage for the "
                       "equipment beneath it. Inactive documents do not count as coverage.", None),

        ("SECTION", "NEVER DELETE A ROW ON SETUP", None),
        ("Why", "The rest of the workbook reads SETUP by position. Deleting a row there - a "
                "status code, a milestone, a wording line - destroys the reference and the "
                "dashboard lists go blank without saying why. SETUP is therefore protected "
                "against inserting and deleting rows and columns. Every blue cell stays "
                "editable as normal.", None),
        ("To drop a status code", "Clear the cells on that line instead. An empty line is "
                                  "simply ignored everywhere.", None),
        ("If you really must", "Review > Unprotect Sheet - there is no password. If a deletion "
                               "does break something, a red banner appears at the top of the "
                               "DASHBOARD instead of the lists silently emptying.", None),

        ("SECTION", "CAPACITY AND LIMITS", None),
        ("Rows", f"TRACKER {N_TRK} rows  -  ACTIONS {N_ACT}  -  systems {N_SYS}  -  "
                 f"equipment {N_EQP}  -  milestones {N_MS_SLOT}  -  "
                 f"core types tracked for gaps {N_CORE}.", None),
        ("One flat table", "The TRACKER is a single table, not one block per type. Capacity is "
                           "shared: 500 rows go wherever they are needed, and you can filter or "
                           "sort across every type at once.", None),
        ("To add rows", "Copy the last data row and paste down, then widen the named ranges in "
                        "Formulas > Name Manager.", None),
        ("No macros", "Formulas only. Nothing to enable, nothing a corporate policy can block.", None),
        ("Excel version", "Uses no function newer than Excel 2007, so it runs on any current "
                          "corporate build.", None),

        ("SECTION", "ASSUMPTIONS", None),
        ("Months to days", "30 days per month (T-3 months = 90), which reproduces the worked "
                           "example in the source specification exactly.", None),
        ("Post-commissioning anchors", "Anchored to named milestones rather than to guessed "
                                       "day counts. Adjust the offsets in SETUP section 5.", None),
        ("Equipment with no date", "If nothing behind a row has a commissioning start, the row "
                                   "reads 'Scope not resolved' rather than inventing a target.", None),
    ]

    r = 4
    for a, b, colour in rows:
        if a == "SECTION":
            section(ws, r, 2, 3, b)
            r += 1
            continue
        c = ws.cell(row=r, column=2, value=a)
        c.font = f_lbl9
        c.alignment = a_w
        c.border = box
        if colour:
            c.fill = PatternFill("solid", fgColor=colour)
        d = ws.cell(row=r, column=3, value=b)
        d.font = f_body
        d.alignment = a_w
        d.border = box
        ws.row_dimensions[r].height = 30
        r += 1


def build_static_lists(ws):
    """Small fixed dropdown sources that do not belong in SETUP."""
    for i, v in enumerate(["Done", "In progress", "Inactive"]):
        ws.cell(row=2 + i, column=35, value=v)
    for i, v in enumerate(SCOPES):
        ws.cell(row=2 + i, column=37, value=v)
    for i, v in enumerate(["Yes", "No"]):
        ws.cell(row=2 + i, column=39, value=v)


# ===========================================================================
# Defined names
# ===========================================================================
def add_names(wb):
    n = {}
    for i, (_, _, key) in enumerate(WORDING):
        n[key] = f"{SU}!$C${WD_TOP + i}"
    for i, (_, _, key) in enumerate(RAG_WORDING):
        n[key] = f"{SU}!$C${RAG_TOP + i}"
    for i, (_, _, key) in enumerate(ACTION_WORDING):
        n[key] = f"{SU}!$C${AST_TOP + i}"

    n.update({
        "CTR_NUM": f"{SU}!$C${CT_TOP}",
        "CTR_NAME": f"{SU}!$C${CT_TOP + 1}",
        "FEN_DOC": f"{SU}!$C${TH_TOP}",
        "FEN_ACT": f"{SU}!$C${TH_TOP + 1}",
        "HORIZON": f"{SU}!$C${TH_TOP + 2}",
        "CTR_EARLIEST": f"{SU}!$C${TH_TOP + 3}",
        "WF_CODE": f"{SU}!$B${WF_TOP}:$B${WF_BOT}",
        "WF_CLASS": f"{SU}!$D${WF_TOP}:$D${WF_BOT}",
        "MS_NOM": f"{SU}!$B${MS_TOP}:$B${MS_BOT}",
        "MS_PLAN": f"{SU}!$E${MS_TOP}:$E${MS_BOT}",
        "SYS_ID": f"{SU}!$B${SYS_TOP}:$B${SYS_BOT}",
        "SYS_EARLIEST": f"{SU}!$E${SYS_TOP}:$E${SYS_BOT}",
        "SYS_LATE": f"{SU}!$I${SYS_TOP}:$I${SYS_BOT}",
        "EQP_SYS": f"{SU}!$B${EQP_TOP}:$B${EQP_BOT}",
        "EQP_NOM": f"{SU}!$C${EQP_TOP}:$C${EQP_BOT}",
        "EQP_START": f"{SU}!$D${EQP_TOP}:$D${EQP_BOT}",
        "EQP_ACTUAL": f"{SU}!$F${EQP_TOP}:$F${EQP_BOT}",
        "EQP_EFF": f"{SU}!$O${EQP_TOP}:$O${EQP_BOT}",
        "EQP_DOCS": f"{SU}!$H${EQP_TOP}:$H${EQP_BOT}",
        "EQP_DONE": f"{SU}!$I${EQP_TOP}:$I${EQP_BOT}",
        "EQP_LATE": f"{SU}!$J${EQP_TOP}:$J${EQP_BOT}",
        "EQP_CORE": f"{SU}!$K${EQP_TOP}:$K${EQP_BOT}",
        "EQP_PCT": f"{SU}!$L${EQP_TOP}:$L${EQP_BOT}",
        "EQP_STATUS": f"{SU}!$M${EQP_TOP}:$M${EQP_BOT}",
        "EQW_KEY": f"{SU}!$N${EQP_TOP}:$N${EQP_BOT}",
        "TYP_TYPE": f"{SU}!$B${TYP_TOP}:$B${TYP_BOT}",
        "TYP_PHASE": f"{SU}!$C${TYP_TOP}:$C${TYP_BOT}",
        "TYP_CORE": f"{SU}!$D${TYP_TOP}:$D${TYP_BOT}",
        "TYP_ANCHOR": f"{SU}!$E${TYP_TOP}:$E${TYP_BOT}",
        "TYP_LEAD": f"{SU}!$F${TYP_TOP}:$F${TYP_BOT}",
        "TYP_SCOPE": f"{SU}!$G${TYP_TOP}:$G${TYP_BOT}",
        "TYP_COREKEY": f"{SU}!$I${TYP_TOP}:$I${TYP_BOT}",
        "PHASE_LIST": f"{SU}!$B${PH_TOP}:$B${PH_BOT}",
        "ACT_STATUS_LIST": f"{SU}!$C${AST_TOP}:$C${AST_TOP + 3}",
        "PPL_NAME": f"{SU}!$B${PPL_TOP}:$B${PPL_BOT}",
        # TRACKER
        "DOC_TYPE": f"{TR}!${K_TYPE}${TRK_TOP}:${K_TYPE}${TRK_BOT}",
        "DOC_SCOPE": f"{TR}!${K_SCOPE}${TRK_TOP}:${K_SCOPE}${TRK_BOT}",
        "DOC_SYS": f"{TR}!${K_SYS}${TRK_TOP}:${K_SYS}${TRK_BOT}",
        "DOC_EQP": f"{TR}!${K_EQP}${TRK_TOP}:${K_EQP}${TRK_BOT}",
        "DOC_NUM": f"{TR}!${K_TCN}${TRK_TOP}:${K_TCN}${TRK_BOT}",
        "DOC_TCN": f"{TR}!${K_TCN}${TRK_TOP}:${K_TCN}${TRK_BOT}",
        "DOC_REV": f"{TR}!${K_REV}${TRK_TOP}:${K_REV}${TRK_BOT}",
        "DOC_TARGET": f"{TR}!${K_TARGET}${TRK_TOP}:${K_TARGET}${TRK_BOT}",
        "DOC_STATE": f"{TR}!${K_STATE}${TRK_TOP}:${K_STATE}${TRK_BOT}",
        "DOC_CHECK": f"{TR}!${K_CHECK}${TRK_TOP}:${K_CHECK}${TRK_BOT}",
        "DOC_ACTIVE": f"{TR}!${K_ACTIVE}${TRK_TOP}:${K_ACTIVE}${TRK_BOT}",
        "DOC_DONE": f"{TR}!${K_DONE}${TRK_TOP}:${K_DONE}${TRK_BOT}",
        # ACTIONS
        "ACT_ID": f"{AC}!$A${ACT_TOP}:$A${ACT_BOT}",
        "ACT_STATUS": f"{AC}!$K${ACT_TOP}:$K${ACT_BOT}",
        "ACT_STATE": f"{AC}!$M${ACT_TOP}:$M${ACT_BOT}",
        # LISTS
        "GAP_KIND": f"{LS}!$G${GAP_TOP}:$G${GAP_BOT}",
        "GAP_WHERE": f"{LS}!$H${GAP_TOP}:$H${GAP_BOT}",
        "GAP_WHAT": f"{LS}!$I${GAP_TOP}:$I${GAP_BOT}",
        "GAP_KEY": f"{LS}!$J${GAP_TOP}:$J${GAP_BOT}",
        "CMB_KIND": f"{LS}!$L${CMB_TOP}:$L${CMB_BOT}",
        "CMB_SUBJ": f"{LS}!$M${CMB_TOP}:$M${CMB_BOT}",
        "CMB_SCOPE": f"{LS}!$N${CMB_TOP}:$N${CMB_BOT}",
        "CMB_OWNER": f"{LS}!$O${CMB_TOP}:$O${CMB_BOT}",
        "CMB_TARGET": f"{LS}!$P${CMB_TOP}:$P${CMB_BOT}",
        "CMB_DAYS": f"{LS}!$Q${CMB_TOP}:$Q${CMB_BOT}",
        "CMB_KWIP": f"{LS}!$R${CMB_TOP}:$R${CMB_BOT}",
        "CMB_KCOME": f"{LS}!$S${CMB_TOP}:$S${CMB_BOT}",
        "CMB_KLATE": f"{LS}!$T${CMB_TOP}:$T${CMB_BOT}",
        "COUNTS_AS": f"{LS}!$AI$2:$AI$4",
        "SCOPE_LIST": f"{LS}!$AK$2:$AK$4",
        "YESNO": f"{LS}!$AM$2:$AM$3",
    })
    for name, ref in n.items():
        wb.defined_names.add(DefinedName(name, attr_text=ref))


# ===========================================================================
# Build
# ===========================================================================
def build(demo, path):
    wb = Workbook()
    wb.remove(wb.active)
    ws_dash = wb.create_sheet(DB)
    ws_trk = wb.create_sheet(TR)
    ws_act = wb.create_sheet(AC)
    ws_set = wb.create_sheet(SU)
    ws_how = wb.create_sheet("HOW IT WORKS")
    ws_ls = wb.create_sheet(LS)

    build_dashboard(ws_dash)
    build_tracker(ws_trk)
    build_actions(ws_act)
    build_setup(ws_set, demo)
    build_how(ws_how)
    build_lists(ws_ls)
    build_static_lists(ws_ls)
    add_charts(ws_dash, ws_ls)
    add_names(wb)

    if demo:
        write_demo(wb)
    else:
        write_example(wb)

    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(path)
    print(f"wrote {path}")


# ===========================================================================
# Demo / example data
# ===========================================================================
DEMO_SYS = [
    ("SYS-TP", "Traction Power & 33kV Distribution", "D. Petrov", 60),
    ("SYS-SG", "Signalling & Train Control", "J. Whitfield", 150),
    ("SYS-TV", "Tunnel Ventilation", "P. Nakamura", 240),
    ("SYS-FA", "Fire Alarm & Detection", "K. Rasmussen", 330),
    ("SYS-LV", "LV Distribution & UPS", "R. Bellamy", 420),
    ("SYS-CC", "Control Centre SCADA", "N. Bergstrom", 520),
]

DEMO_EQP = [
    ("SYS-TP", "TP-BSS-01 Bulk Supply Substation", 0),
    ("SYS-TP", "TP-TSS-02 Traction Substation Alpha", 20),
    ("SYS-TP", "TP-TSS-03 Traction Substation Bravo", 40),
    ("SYS-TP", "TP-OHL-04 Overhead Line Section A", 55),
    ("SYS-SG", "SG-IXL-01 Interlocking Cabinet North", 0),
    ("SYS-SG", "SG-IXL-02 Interlocking Cabinet South", 25),
    ("SYS-SG", "SG-ATP-03 ATP Trackside Equipment", 45),
    ("SYS-SG", "SG-OCC-04 Signalling Workstations", 70),
    ("SYS-TV", "TV-JET-01 Jet Fan Array Tunnel 1", 0),
    ("SYS-TV", "TV-JET-02 Jet Fan Array Tunnel 2", 30),
    ("SYS-TV", "TV-SMK-03 Smoke Extraction Plant", 60),
    ("SYS-FA", "FA-PAN-01 Fire Panel Main", 0),
    ("SYS-FA", "FA-DET-02 Detection Loop Zone A", 30),
    ("SYS-FA", "FA-SUP-03 Suppression Skid", 55),
    ("SYS-LV", "LV-SWB-01 Main LV Switchboard", 0),
    ("SYS-LV", "LV-UPS-02 UPS Room A", 35),
    ("SYS-LV", "LV-GEN-03 Standby Generator", 65),
    ("SYS-CC", "CC-SRV-01 SCADA Server Cluster", 0),
    ("SYS-CC", "CC-HMI-02 Operator Workstations", 30),
    ("SYS-CC", "CC-NET-03 Control Network Core", 55),
]

DEMO_PPL = [
    ("A. Mercer", "Contract Manager"), ("S. Okonkwo", "Commissioning Manager"),
    ("L. Fontaine", "Quality Lead"), ("D. Petrov", "Lead Commissioning Engineer"),
    ("J. Whitfield", "Commissioning Engineer"), ("P. Nakamura", "Commissioning Engineer"),
    ("K. Rasmussen", "Commissioning Engineer"), ("R. Bellamy", "Commissioning Engineer"),
    ("N. Bergstrom", "Commissioning Engineer"), ("M. Haddad", "Document Controller"),
    ("C. Vasquez", "Approver"), ("H. Lindqvist", "Approver"),
]

# maturity per system: 0 = most advanced, 5 = barely started
MATURITY = {"SYS-TP": 0, "SYS-SG": 1, "SYS-TV": 2,
            "SYS-FA": 3, "SYS-LV": 4, "SYS-CC": 5}

DONE_CODES = ["S4A - Approved", "P1 - For Implementation", "P6 - For Construction"]
WIP_CODES = ["S3 - For Review", "S4 - For Internal Approval", "S6 - For Acceptance",
             "S4B - Comments", "W1 - Draft"]

ABBR = {
    "Basis of Commissioning": "BOC", "Surveillance Strategy": "SVS",
    "Surveillance Programme": "SVP", "Competency Evidence": "CPE",
    "Safe System Of Work": "SSW", "Risk Assessment And Method Statement": "RAM",
    "System Commissioning Programme": "SCP", "Commissioning Procedure": "CPR",
    "Inspection and Test Plan": "ITP",
    "Initial End Of Erection Status Report": "IER",
    "Planned End Of Erection Status Report": "PER",
    "Authority To Start Evidence": "ATS",
    "Commissioning Results Sheet": "CRS",
}

LEAD = {t[0]: t[4] for t in TYPES}


def _status(maturity, idx, days_to_target):
    """Pick a plausible workflow code from maturity and how close the target is."""
    if maturity <= 1 and idx % 7 != 3:
        return DONE_CODES[idx % len(DONE_CODES)], True
    if maturity == 2:
        return (DONE_CODES[idx % 3], True) if idx % 3 == 0 else \
               (WIP_CODES[idx % len(WIP_CODES)], False)
    if days_to_target < 0 and maturity >= 3:
        return WIP_CODES[idx % len(WIP_CODES)], False
    if maturity >= 4 and idx % 5 == 0:
        return "W0 - Work In Progress", False
    if maturity >= 4:
        return "", False
    return WIP_CODES[idx % len(WIP_CODES)], False


def write_demo(wb):
    today = dt.date.today()
    su, tr, ac = wb[SU], wb[TR], wb[AC]

    for i, (sid, name, eng, _) in enumerate(DEMO_SYS):
        r = SYS_TOP + i
        su.cell(row=r, column=2, value=sid)
        su.cell(row=r, column=3, value=name)
        su.cell(row=r, column=4, value=eng)

    eq_start = {}
    for i, (sysid, eq, off) in enumerate(DEMO_EQP):
        r = EQP_TOP + i
        base = today + dt.timedelta(days=dict((s[0], s[3]) for s in DEMO_SYS)[sysid] + off)
        su.cell(row=r, column=2, value=sysid)
        su.cell(row=r, column=3, value=eq)
        c = su.cell(row=r, column=4, value=base); c.number_format = FMT_DATE
        c = su.cell(row=r, column=5, value=base + dt.timedelta(days=45))
        c.number_format = FMT_DATE
        actual = base + dt.timedelta(days=3) if base < today else None
        if actual:
            c = su.cell(row=r, column=6, value=actual)
            c.number_format = FMT_DATE
        # the workbook measures from the actual start once there is one, so seed
        # the demo's own target dates from the same date it will use
        eq_start[eq] = actual or base

    for i, (nm, role) in enumerate(DEMO_PPL):
        r = PPL_TOP + i
        su.cell(row=r, column=2, value=nm)
        su.cell(row=r, column=3, value=role)

    sys_start = {s[0]: min(eq_start[e] for sy, e, _ in DEMO_EQP if sy == s[0])
                 for s in DEMO_SYS}
    contract_start = min(sys_start.values())

    # deliberate gaps so the MISSING block has content
    skip_sys = {("SYS-LV", "Authority To Start Evidence"),
                ("SYS-CC", "Authority To Start Evidence"),
                ("SYS-CC", "System Commissioning Programme")}
    skip_eqp = {("CC-HMI-02 Operator Workstations", "Commissioning Procedure"),
                ("CC-NET-03 Control Network Core", "Commissioning Procedure"),
                ("LV-GEN-03 Standby Generator", "Inspection and Test Plan"),
                ("FA-SUP-03 Suppression Skid", "Planned End Of Erection Status Report")}

    row = TRK_TOP
    n = 0

    def put(dtype, scope, sysid, eq, anchor_date, maturity):
        nonlocal row, n
        n += 1
        target = anchor_date - dt.timedelta(days=LEAD[dtype])
        days_to = (target - today).days
        code, done = _status(maturity, n, days_to)
        J = TRK_IDX
        tr.cell(row=row, column=J["TYPE"], value=dtype)
        tr.cell(row=row, column=J["SCOPE"], value=scope)
        if scope in ("System", "Equipment"):
            tr.cell(row=row, column=J["SYS"], value=sysid)
        if scope == "Equipment":
            tr.cell(row=row, column=J["EQP"], value=eq)
        tr.cell(row=row, column=J["TCN"], value=f"TC-{6100000 + n * 37:07d}")
        tr.cell(row=row, column=J["REV"], value="ABCDE"[min(4, (n % 7) // 2)])
        tr.cell(row=row, column=J["RESP"], value=DEMO_PPL[n % len(DEMO_PPL)][0])
        tr.cell(row=row, column=J["APPR"], value=DEMO_PPL[(n + 4) % len(DEMO_PPL)][0])
        tr.cell(row=row, column=J["STAT"], value=code)
        if code and not done:
            st = target - dt.timedelta(days=45 + (n % 20))
            if st <= today:
                c = tr.cell(row=row, column=J["ASTART"], value=st)
                c.number_format = FMT_DATE
        if done:
            appr = target + dt.timedelta(days=(n % 11) - 5)
            if appr > today:
                appr = today - dt.timedelta(days=n % 7)
            c = tr.cell(row=row, column=J["AAPPR"], value=appr)
            c.number_format = FMT_DATE
        if n % 37 == 0:                       # a couple of cancelled documents
            tr.cell(row=row, column=J["STAT"], value="A2 - Superseded")
            tr.cell(row=row, column=J["AAPPR"], value=None)
        row += 1

    for t in TYPES:
        if t[2] == "Yes" and t[5] == "All":
            put(t[0], "All", None, None, contract_start, 0)
    for sid, _, _, _ in DEMO_SYS:
        for t in TYPES:
            if t[2] == "Yes" and t[5] == "System" and (sid, t[0]) not in skip_sys:
                put(t[0], "System", sid, None, sys_start[sid], MATURITY[sid])
    for sysid, eq, _ in DEMO_EQP:
        for t in TYPES:
            if t[2] == "Yes" and t[5] == "Equipment" and (eq, t[0]) not in skip_eqp:
                put(t[0], "Equipment", sysid, eq, eq_start[eq], MATURITY[sysid])
    for sysid, eq, _ in DEMO_EQP[:10]:        # a few non-core deliverables
        put("Initial End Of Erection Status Report", "Equipment", sysid, eq,
            eq_start[eq], MATURITY[sysid])

    # one deliberate override, to show the Source column working
    tr.cell(row=TRK_TOP + 4, column=TRK_IDX["OVR"],
            value=today + dt.timedelta(days=21)).number_format = FMT_DATE
    tr.cell(row=TRK_TOP + 4, column=TRK_IDX["CMT"],
            value="Target moved by agreement - waiting on the SYS-SG interface package.")

    acts = [
        ("ACT-001", "Close out RAMS comments raised by the Client safety review",
         "SYS-SG", "", "Critical", -22, "In progress"),
        ("ACT-002", "Obtain signed competency records for the ATP trackside team",
         "SYS-SG", "SG-ATP-03 ATP Trackside Equipment", "", -12, "Open"),
        ("ACT-003", "Resolve interlocking cabinet earthing non-conformance",
         "SYS-SG", "SG-IXL-01 Interlocking Cabinet North", "", -5, "In progress"),
        ("ACT-004", "Issue System Commissioning Programme revision C for approval",
         "SYS-TV", "", "", 2, "Open"),
        ("ACT-005", "Agree tunnel ventilation witnessing schedule with the Authority",
         "SYS-TV", "", "", 9, "Open"),
        ("ACT-006", "Confirm 33kV protection settings against the final study",
         "SYS-TP", "TP-TSS-03 Traction Substation Bravo", "", 16, "In progress"),
        ("ACT-007", "Update Basis of Commissioning for the revised staging strategy",
         "", "", "", 24, "On hold"),
        ("ACT-008", "Mobilise second commissioning engineer for tunnel ventilation",
         "SYS-TV", "", "", 31, "Open"),
        ("ACT-009", "Close punch list items on Traction Substation Alpha",
         "SYS-TP", "TP-TSS-02 Traction Substation Alpha", "", 12, "In progress"),
        ("ACT-010", "Submit Initial EESR for the signalling scope",
         "SYS-SG", "", "", -3, "Open"),
        ("ACT-011", "Agree beneficial use conditions with Operations",
         "SYS-TP", "", "", 44, "On hold"),
        ("ACT-012", "Verify smoke extraction damper actuator supply",
         "SYS-TV", "TV-SMK-03 Smoke Extraction Plant", "", 51, "Open"),
        ("ACT-013", "Complete fire detection loop cause and effect matrix",
         "SYS-FA", "FA-DET-02 Detection Loop Zone A", "", 60, "Open"),
        ("ACT-014", "Rectify jet fan vibration readings on Tunnel 1",
         "SYS-TV", "TV-JET-01 Jet Fan Array Tunnel 1", "", -1, "In progress"),
        ("ACT-015", "Archive approved bulk supply substation records",
         "SYS-TP", "TP-BSS-01 Bulk Supply Substation", "", -30, "Done"),
        ("ACT-016", "Confirm SCADA point list against the signalling interface",
         "SYS-CC", "CC-SRV-01 SCADA Server Cluster", "", 75, "Open"),
        ("ACT-017", "Book UPS discharge test window with Operations",
         "SYS-LV", "LV-UPS-02 UPS Room A", "", 68, "Open"),
        ("ACT-018", "Issue Authority To Start evidence pack for traction power",
         "SYS-TP", "", "", 6, "In progress"),
    ]
    for i, (aid, desc, sysid, eq, _, due_off, status) in enumerate(acts):
        r = ACT_TOP + i
        due = today + dt.timedelta(days=due_off)
        ac.cell(row=r, column=1, value=aid)
        ac.cell(row=r, column=2, value=desc)
        ac.cell(row=r, column=3, value="Pre-commissioning")
        if sysid:
            ac.cell(row=r, column=4, value=sysid)
        if eq:
            ac.cell(row=r, column=5, value=eq)
        ac.cell(row=r, column=7, value=DEMO_PPL[(i + 3) % len(DEMO_PPL)][0])
        if status in ("In progress", "Done"):
            c = ac.cell(row=r, column=8, value=due - dt.timedelta(days=20))
            c.number_format = FMT_DATE
        c = ac.cell(row=r, column=9, value=due); c.number_format = FMT_DATE
        if status == "Done":
            c = ac.cell(row=r, column=10, value=due - dt.timedelta(days=2))
            c.number_format = FMT_DATE
        ac.cell(row=r, column=11, value=status)


def write_example(wb):
    """Blank production file: one worked example per sheet, to overwrite."""
    today = dt.date.today()
    su, tr, ac = wb[SU], wb[TR], wb[AC]
    tag = PatternFill("solid", fgColor="FFF9E6")

    for i, (nm, role) in enumerate([("A. Example", "Commissioning Manager"),
                                    ("B. Example", "Commissioning Engineer"),
                                    ("C. Example", "Approver")]):
        su.cell(row=PPL_TOP + i, column=2, value=nm).fill = tag
        su.cell(row=PPL_TOP + i, column=3, value=role).fill = tag

    su.cell(row=SYS_TOP, column=2, value="SYS-01").fill = tag
    su.cell(row=SYS_TOP, column=3, value="EXAMPLE - overwrite with your first system").fill = tag
    su.cell(row=SYS_TOP, column=4, value="B. Example").fill = tag

    start = today + dt.timedelta(days=240)
    su.cell(row=EQP_TOP, column=2, value="SYS-01").fill = tag
    su.cell(row=EQP_TOP, column=3, value="EQ-001 Example equipment item").fill = tag
    c = su.cell(row=EQP_TOP, column=4, value=start)
    c.number_format = FMT_DATE
    c.fill = tag
    c = su.cell(row=EQP_TOP, column=5, value=start + dt.timedelta(days=45))
    c.number_format = FMT_DATE
    c.fill = tag

    K = TRK_IDX
    vals = {K["TYPE"]: "Commissioning Procedure", K["SCOPE"]: "Equipment",
            K["SYS"]: "SYS-01", K["EQP"]: "EQ-001 Example equipment item",
            K["TCN"]: "TC-6100001", K["REV"]: "A",
            K["RESP"]: "B. Example", K["APPR"]: "C. Example", K["STAT"]: "W1 - Draft",
            K["CMT"]: "EXAMPLE ROW - overwrite with your first deliverable"}
    for col, v in vals.items():
        tr.cell(row=TRK_TOP, column=col, value=v).fill = tag

    vals = {1: "ACT-001", 2: "EXAMPLE - overwrite with your first action",
            3: "Pre-commissioning", 4: "SYS-01", 7: "B. Example",
            9: today + dt.timedelta(days=14), 11: "Open"}
    for col, v in vals.items():
        c = ac.cell(row=ACT_TOP, column=col, value=v)
        c.fill = tag
        if col == 9:
            c.number_format = FMT_DATE


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--demo", action="store_true")
    ap.add_argument("-o", "--output", required=True)
    args = ap.parse_args()
    build(args.demo, args.output)


if __name__ == "__main__":
    main()
