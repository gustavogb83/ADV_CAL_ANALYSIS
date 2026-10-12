"""Builds Advent_Calendar_Sales.xlsx

Inputs live on the 'Data' sheet (one row per market x product x year).
'Calc' holds every formula/text; the five slide sheets only display Calc.
Re-run this script only if you want to regenerate the workbook layout;
day-to-day updates are done directly in Excel.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.chart.label import DataLabelList, DataLabel
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.text import RichText
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font as DFont
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
from datetime import date

OUT = "Advent_Calendar_Sales.xlsx"

# ── design tokens (styles.css) ────────────────────────────────────────────
BG, TEXT = "F2F2F3", "1D1F20"
A100, A200, A700, A900 = "EEF6FF", "D6EBFF", "416180", "1D2D3D"
MUTED = "686969"          # text @ 65% on bg
HAIR = "C3C3C5"           # hairline border
ROWLINE = "E0E0E1"
HEAD, BODY = "Barlow Condensed", "Barlow"

EUR0, EUR2, INT = '"€"#,##0', '"€"#,##0.00', '#,##0'
PCT = '+0.0%;"−"0.0%;0.0%'

NCOL, NROW = 60, 36        # slide grid: 60 cols x 36 rows (col 21px, row 20px ≈ 0.65 × 1920×1080)
FIRST, LAST = 6, 505       # Data rows


def F(name=BODY, size=12, bold=False, color=TEXT, italic=False):
    return Font(name=name, size=size, bold=bold, color=color, italic=italic)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


wb = Workbook()

# ═════════════════════════════════════════════════════════════════════════
# CONTROL
# ═════════════════════════════════════════════════════════════════════════
ctl = wb.active
ctl.title = "Control"
ctl.sheet_view.showGridLines = False
ctl.sheet_properties.tabColor = A700
for col, w in zip("ABCDEFGH", (3, 30, 22, 4, 26, 4, 70, 4)):
    ctl.column_dimensions[col].width = w

ctl["B2"] = "Advent Calendar Sales — control panel"
ctl["B2"].font = F(HEAD, 24, True, A900)
ctl["B3"] = "Change the yellow-blue input cells; every slide sheet updates automatically."
ctl["B3"].font = F(BODY, 11, color=MUTED)

INPUT_FILL, INPUT_BORDER = fill(A100), Border(*(Side("thin", color=A700),) * 4)


def label(cell, text, bold=False):
    ctl[cell] = text
    ctl[cell].font = F(BODY, 11, bold)


def inp(cell, value, fmt=None):
    c = ctl[cell]
    c.value = value
    c.font = F(BODY, 12, True, A900)
    c.fill, c.border = INPUT_FILL, INPUT_BORDER
    c.alignment = Alignment(horizontal="center")
    if fmt:
        c.number_format = fmt


ctl["B5"] = "SETTINGS"
ctl["B5"].font = F(HEAD, 14, True, A700)
label("B6", "Current year")
inp("C6", 2026)
label("B7", "Comparison year")
ctl["C7"] = "=C6-1"
ctl["C7"].alignment = Alignment(horizontal="center")
ctl["C7"].font = F(BODY, 12)
label("B8", "Market")
inp("C8", "All markets")

ctl["B10"] = "PRODUCTS IN SCOPE"
ctl["B10"].font = F(HEAD, 14, True, A700)
for i, (p, v) in enumerate((("24-day", "Yes"), ("12-day", "Yes"), ("Refill", "No"))):
    r = 11 + i
    ctl[f"B{r}"] = p
    ctl[f"B{r}"].font = F(BODY, 11, True)
    inp(f"C{r}", v)
ctl["B14"] = "(product names are fixed — do not rename)"
ctl["B14"].font = F(BODY, 9, color=MUTED, italic=True)

ctl["B16"] = "STATUS"
ctl["B16"].font = F(HEAD, 14, True, A700)
label("B17", "Data check")
label("B18", "Report date (latest cut-off)")
label("B19", "Rows used in comparison")

ctl["E5"] = "MARKETS"
ctl["E5"].font = F(HEAD, 14, True, A700)
ctl["E6"] = "All markets"
ctl["E6"].font = F(BODY, 11, color=MUTED)
ctl["E7"] = "Market A"
for r in range(7, 46):
    c = ctl[f"E{r}"]
    c.fill, c.border = INPUT_FILL, INPUT_BORDER
    c.font = F(BODY, 11, True, A900)
ctl["E4"] = "Type your markets here (one per row)"
ctl["E4"].font = F(BODY, 9, color=MUTED, italic=True)

guide = [
    ("HOW TO UPDATE", None),
    ("1", "Go to the Data sheet. For each market × product × year add one row (or overwrite the numbers on an existing row)."),
    ("2", "Fill Launch date (blank row = product not sold in that market that year), Data through (the cut-off date of the numbers), Units and Revenue (€, cumulative from launch to the cut-off)."),
    ("3", "Enter the comparison-year numbers to the same calendar date (e.g. 9 Oct 2025 next to 9 Oct 2026)."),
    ("4", "New market? Add its name in the MARKETS list (left), then pick it in Data column A."),
    ("5", "Pick the Market and the products in scope on this sheet. 'All markets' adds up only markets/products that have rows for BOTH years."),
    ("6", "Look at the sheets 1 Cover … 5 By size. Check the 'Data check' status above; fix any row flagged in Data column S."),
    ("NOTE", "Equal-days estimate = comparison-year × (days selling this year ÷ days selling last year), done market by market, assuming an even daily rate. Head-start value = comparison-year × (extra days ÷ days selling last year)."),
    ("NOTE", "Fonts: install 'Barlow' and 'Barlow Condensed' (free, Google Fonts) for the intended look; otherwise Excel substitutes."),
    ("NOTE", "Needs Excel 2019 / Microsoft 365 (uses MAXIFS / MINIFS)."),
]
for i, (k, t) in enumerate(guide):
    r = 5 + i
    if t is None:
        ctl[f"G{r}"] = k
        ctl[f"G{r}"].font = F(HEAD, 14, True, A700)
        continue
    ctl[f"G{r}"] = f"{k}.  {t}" if k.isdigit() else f"{k}:  {t}"
    ctl[f"G{r}"].font = F(BODY, 10)
    ctl[f"G{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    ctl.row_dimensions[r].height = 42 if len(t) > 110 else 28

# validations
dv_market = DataValidation(type="list", formula1="=Market_List", allow_blank=False)
dv_yn = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
dv_year = DataValidation(type="whole", operator="between", formula1="2000", formula2="2100")
ctl.add_data_validation(dv_market)
ctl.add_data_validation(dv_yn)
ctl.add_data_validation(dv_year)
dv_market.add("C8")
dv_yn.add("C11:C13")
dv_year.add("C6")

# names
names = {
    "Yr_Cur": "Control!$C$6",
    "Yr_Prev": "Control!$C$7",
    "Sel_Market": "Control!$C$8",
    "Market_List": "OFFSET(Control!$E$6,0,0,COUNTA(Control!$E$6:$E$45),1)",
    "Market_Entry": "OFFSET(Control!$E$7,0,0,MAX(1,COUNTA(Control!$E$7:$E$45)),1)",
}
for col in "ABCDEFGHIJKLMNOPQRS":
    names[f"Dat_{col}"] = f"Data!${col}${FIRST}:${col}${LAST}"
for k, v in names.items():
    wb.defined_names[k] = DefinedName(k, attr_text=v)

# ═════════════════════════════════════════════════════════════════════════
# DATA
# ═════════════════════════════════════════════════════════════════════════
d = wb.create_sheet("Data")
d.sheet_properties.tabColor = A900
d["A1"] = "Sales data — one row per market × product × year"
d["A1"].font = F(HEAD, 22, True, A900)
d["A2"] = ("Blue columns = you type. Grey columns = calculated. Numbers are cumulative from launch to 'Data through'. "
           "Leave a product out (or the launch date blank) if it was not on sale in that market that year. Rows 6–505 are ready to use.")
d["A2"].font = F(BODY, 10, color=MUTED)
d["A3"] = "Years other than Current/Comparison year are kept as history and ignored."
d["A3"].font = F(BODY, 10, color=MUTED)

heads = ["Market", "Product", "Year", "Launch date", "Data through", "Units", "Revenue (€)",
         "Selling days", "Avg price (€)", "Role", "Days other yr", "In comparison",
         "Days CY", "Days PY", "Units eq-days", "Rev eq-days", "Head-start units", "Head-start rev", "Row check"]
widths = [16, 11, 7, 13, 13, 10, 13, 10, 11, 7, 9, 10, 8, 8, 11, 11, 11, 11, 38]
for i, (h, w) in enumerate(zip(heads, widths), start=1):
    c = d.cell(5, i, h)
    c.font = F(BODY, 10, True, "FFFFFF")
    c.fill = fill(A900 if i <= 7 else "7A7A7D")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    d.column_dimensions[c.column_letter].width = w
d.row_dimensions[5].height = 32
d.freeze_panes = "A6"

thin = Side("thin", color="D4D4D7")
box_b = Border(left=thin, right=thin, top=thin, bottom=thin)
calc_fill = fill("F0F0F1")
S = lambda col: f"Dat_{col}"

for r in range(FIRST, LAST + 1):
    for col in range(1, 8):
        c = d.cell(r, col)
        c.border = box_b
        c.font = F(BODY, 10)
    d.cell(r, 4).number_format = "d mmm yyyy"
    d.cell(r, 5).number_format = "d mmm yyyy"
    d.cell(r, 6).number_format = INT
    d.cell(r, 7).number_format = "#,##0.00"
    f = {
        8: f'=IF(AND(ISNUMBER(D{r}),ISNUMBER(E{r})),E{r}-D{r}+1,"")',
        9: f'=IF(AND(ISNUMBER(F{r}),N(F{r})>0,ISNUMBER(G{r})),G{r}/F{r},"")',
        10: f'=IF(OR(A{r}="",B{r}=""),"",IF(C{r}=Yr_Cur,"CY",IF(C{r}=Yr_Prev,"PY","")))',
        11: f'=IF(J{r}="","",IF(J{r}="CY",SUMIFS(Dat_H,Dat_A,A{r},Dat_B,B{r},Dat_C,Yr_Prev),SUMIFS(Dat_H,Dat_A,A{r},Dat_B,B{r},Dat_C,Yr_Cur)))',
        12: (f'=IF(J{r}="",0,IF(AND(ISNUMBER(H{r}),N(K{r})>0,ISNUMBER(F{r}),ISNUMBER(G{r}),'
             f'OR(Sel_Market="All markets",A{r}=Sel_Market),'
             f'COUNTIFS(Control!$B$11:$B$13,B{r},Control!$C$11:$C$13,"Yes")>0),1,0))'),
        13: f'=IF(L{r}=1,IF(J{r}="CY",H{r},K{r}),"")',
        14: f'=IF(L{r}=1,IF(J{r}="PY",H{r},K{r}),"")',
        15: f'=IF(L{r}=1,IF(J{r}="PY",F{r}*M{r}/N{r},F{r}),0)',
        16: f'=IF(L{r}=1,IF(J{r}="PY",G{r}*M{r}/N{r},G{r}),0)',
        17: f'=IF(AND(L{r}=1,J{r}="PY"),F{r}*(N{r}-M{r})/N{r},0)',
        18: f'=IF(AND(L{r}=1,J{r}="PY"),G{r}*(N{r}-M{r})/N{r},0)',
        19: (f'=IF(COUNTA(A{r}:G{r})=0,"",IF(OR(A{r}="",B{r}="",NOT(ISNUMBER(C{r}))),"Fill market, product and year",'
             f'IF(COUNTIF(Control!$B$11:$B$13,B{r})=0,"Unknown product",'
             f'IF(COUNTIFS(Dat_A,A{r},Dat_B,B{r},Dat_C,C{r})>1,"Duplicate market/product/year",'
             f'IF(OR(NOT(ISNUMBER(D{r})),NOT(ISNUMBER(E{r}))),"Add launch date and Data through date",'
             f'IF(E{r}<D{r},"Data through is before launch",'
             f'IF(OR(NOT(ISNUMBER(F{r})),NOT(ISNUMBER(G{r}))),"Add units and revenue","OK")))))))'),
    }
    for col, formula in f.items():
        c = d.cell(r, col, formula)
        c.fill = calc_fill
        c.font = F(BODY, 10, color="5D5D60")
        c.border = box_b
    d.cell(r, 9).number_format = "#,##0.00"
    for col in (15, 16, 17, 18):
        d.cell(r, col).number_format = "#,##0.0"

# sample data (from the handoff README)
sample = [
    ("Market A", "24-day", 2025, date(2025, 9, 6), date(2025, 10, 9), 711, 76876),
    ("Market A", "12-day", 2025, date(2025, 9, 6), date(2025, 10, 9), 52, 3572),
    ("Market A", "24-day", 2026, date(2026, 9, 11), date(2026, 10, 9), 398, 53563),
    ("Market A", "12-day", 2026, date(2026, 9, 11), date(2026, 10, 9), 54, 3896),
]
for i, row in enumerate(sample):
    for j, v in enumerate(row, start=1):
        d.cell(FIRST + i, j, v)

dvm = DataValidation(type="list", formula1="=Market_Entry", allow_blank=True, showErrorMessage=True,
                     errorTitle="Unknown market", error="Add the market to the MARKETS list on the Control sheet first.")
dvp = DataValidation(type="list", formula1="=Control!$B$11:$B$13", allow_blank=True)
dvy = DataValidation(type="whole", operator="between", formula1="2000", formula2="2100", allow_blank=True)
dvd = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                     errorTitle="Date needed", error="Enter a valid date.")
for dv, rng in ((dvm, f"A{FIRST}:A{LAST}"), (dvp, f"B{FIRST}:B{LAST}"), (dvy, f"C{FIRST}:C{LAST}"),
                (dvd, f"D{FIRST}:E{LAST}")):
    d.add_data_validation(dv)
    dv.add(rng)
d.conditional_formatting.add(f"S{FIRST}:S{LAST}", FormulaRule(formula=[f'AND(S{FIRST}<>"",S{FIRST}<>"OK")'],
                             font=Font(color="B00020", bold=True)))
d.conditional_formatting.add(f"S{FIRST}:S{LAST}", FormulaRule(formula=[f'S{FIRST}="OK"'], font=Font(color="2E7D32")))
d.column_dimensions.group("J", "R", hidden=True)
d.print_area = "A1:S60"

# Control status formulas
ctl["C17"] = ('=IF(COUNTIF(Dat_S,"?*")-COUNTIF(Dat_S,"OK")=0,"All rows OK",'
              '(COUNTIF(Dat_S,"?*")-COUNTIF(Dat_S,"OK"))&" row(s) need attention (Data col S)")')
ctl["C18"] = '=IF(COUNTIFS(Dat_J,"CY",Dat_L,1)=0,"no data",_xlfn.MAXIFS(Dat_E,Dat_J,"CY",Dat_L,1))'
ctl["C18"].number_format = "d mmm yyyy"
ctl["C19"] = '=COUNTIF(Dat_L,1)'
for c in ("C17", "C18", "C19"):
    ctl[c].font = F(BODY, 11, True)
    ctl[c].alignment = Alignment(horizontal="left")
ctl.conditional_formatting.add("C17", FormulaRule(formula=['$C$17<>"All rows OK"'], font=Font(color="B00020", bold=True)))

# ═════════════════════════════════════════════════════════════════════════
# CALC
# ═════════════════════════════════════════════════════════════════════════
cs = wb.create_sheet("Calc")
cs.sheet_properties.tabColor = "7A7A7D"
cs["A1"] = "Calculation sheet — no inputs here"
cs["A1"].font = F(HEAD, 18, True, A900)
cs["A2"] = "Everything below is computed from Data (only rows 'In comparison' = 1)."
cs["A2"].font = F(BODY, 10, color=MUTED)
hdr = ["Product", "In scope", "Units CY", "Rev CY", "Units PY", "Rev PY", "Units PY eq-days", "Rev PY eq-days",
       "Head-start units", "Head-start rev", "Units Δ same date", "Rev Δ same date", "Units Δ equal days",
       "Rev Δ equal days", "Avg price CY", "Avg price PY", "Price Δ"]
for i, h in enumerate(hdr, start=1):
    c = cs.cell(5, i, h)
    c.font = F(BODY, 10, True, "FFFFFF")
    c.fill = fill(A900)
    c.alignment = Alignment(horizontal="center", wrap_text=True)
    cs.column_dimensions[c.column_letter].width = 14
cs.column_dimensions["A"].width = 34
cs.column_dimensions["B"].width = 60
cs.row_dimensions[5].height = 30

PR = {"24-day": 6, "12-day": 7, "Refill": 8}
TOT = 9
srcs = {"C": ("F", "CY"), "D": ("G", "CY"), "E": ("F", "PY"), "F": ("G", "PY"),
        "G": ("O", "PY"), "H": ("P", "PY"), "I": ("Q", "PY"), "J": ("R", "PY")}
for i, (p, r) in enumerate(PR.items()):
    cs[f"A{r}"] = f"=Control!B{11 + i}"
    cs[f"B{r}"] = f'=IF(Control!C{11 + i}="Yes",1,0)'
    for col, (dc, role) in srcs.items():
        cs[f"{col}{r}"] = f'=SUMIFS(Dat_{dc},Dat_J,"{role}",Dat_L,1,Dat_B,$A{r})'
cs[f"A{TOT}"] = "Total"
cs[f"B{TOT}"] = 1
for col, (dc, role) in srcs.items():
    cs[f"{col}{TOT}"] = f'=SUMIFS(Dat_{dc},Dat_J,"{role}",Dat_L,1)'
for r in (*PR.values(), TOT):
    cs[f"K{r}"] = f'=IF(E{r}>0,C{r}/E{r}-1,"n/a")'
    cs[f"L{r}"] = f'=IF(F{r}>0,D{r}/F{r}-1,"n/a")'
    cs[f"M{r}"] = f'=IF(G{r}>0,C{r}/G{r}-1,"n/a")'
    cs[f"N{r}"] = f'=IF(H{r}>0,D{r}/H{r}-1,"n/a")'
    cs[f"O{r}"] = f'=IF(C{r}>0,D{r}/C{r},"n/a")'
    cs[f"P{r}"] = f'=IF(E{r}>0,F{r}/E{r},"n/a")'
    cs[f"Q{r}"] = f'=IF(AND(ISNUMBER(O{r}),ISNUMBER(P{r})),O{r}/P{r}-1,"n/a")'
    for col in "CEGI":
        cs[f"{col}{r}"].number_format = "#,##0.0"
    for col in "DFHJ":
        cs[f"{col}{r}"].number_format = "#,##0"
    for col in "KLMNQ":
        cs[f"{col}{r}"].number_format = PCT
    for col in "OP":
        cs[f"{col}{r}"].number_format = "#,##0.00"
    for col in "ABCDEFGHIJKLMNOPQ":
        cs[f"{col}{r}"].font = F(BODY, 10, r == TOT)

# text / helper items
item_row = [12]
cs["A12"] = "Item"
cs["B12"] = "Value"
for c in ("A12", "B12"):
    cs[c].font = F(BODY, 10, True, "FFFFFF")
    cs[c].fill = fill(A900)
item_row[0] = 13
REF = {}


def item(key, formula, fmt=None):
    r = item_row[0]
    item_row[0] += 1
    cs[f"A{r}"] = key
    cs[f"B{r}"] = formula
    cs[f"A{r}"].font = F(BODY, 10, True)
    cs[f"B{r}"].font = F(BODY, 10)
    cs[f"B{r}"].alignment = Alignment(wrap_text=False)
    if fmt:
        cs[f"B{r}"].number_format = fmt
    REF[key] = f"Calc!$B${r}"
    return REF[key]


MON = '"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"'
MONF = ('"January","February","March","April","May","June","July","August","September","October",'
        '"November","December"')


def dtxt(c):
    return f'IF({c}=0,"n/a",DAY({c})&" "&CHOOSE(MONTH({c}),{MON}))'


R = item("report_date", "=IF(ISNUMBER(Control!C18),Control!C18,0)", "d mmm yyyy")
d_cy = item("d_cy", '=IFERROR(_xlfn.MAXIFS(Dat_M,Dat_L,1),0)')
d_cy_min = item("d_cy_min", '=IFERROR(_xlfn.MINIFS(Dat_M,Dat_L,1),0)')
d_py = item("d_py", '=IFERROR(_xlfn.MAXIFS(Dat_N,Dat_L,1),0)')
d_py_min = item("d_py_min", '=IFERROR(_xlfn.MINIFS(Dat_N,Dat_L,1),0)')
hs = item("hs_days", f"={d_py}-{d_cy}")
lcmin = item("launch_cy_min", '=IFERROR(_xlfn.MINIFS(Dat_D,Dat_J,"CY",Dat_L,1),0)', "d mmm yyyy")
lcmax = item("launch_cy_max", '=IFERROR(_xlfn.MAXIFS(Dat_D,Dat_J,"CY",Dat_L,1),0)', "d mmm yyyy")
lpmin = item("launch_py_min", '=IFERROR(_xlfn.MINIFS(Dat_D,Dat_J,"PY",Dat_L,1),0)', "d mmm yyyy")
lpmax = item("launch_py_max", '=IFERROR(_xlfn.MAXIFS(Dat_D,Dat_J,"PY",Dat_L,1),0)', "d mmm yyyy")
market = item("market", '=IF(Sel_Market="All markets","All markets",Sel_Market)')
cs_rng = lambda c: f"Calc!${c}$"  # noqa

t = lambda col: f"Calc!${col}${TOT}"
gap_rev = item("gap_rev", f"={t('F')}-{t('D')}", "#,##0")
gap_units = item("gap_units", f"={t('E')}-{t('C')}", "#,##0")
share_rev = item("share_rev", f'=IF({gap_rev}>0,{t("J")}/{gap_rev},"n/a")', "0.0%")
share_units = item("share_units", f'=IF({gap_units}>0,{t("I")}/{gap_units},"n/a")', "0.0%")
gap_rev_eq = item("gap_rev_eq", f"={gap_rev}-{t('J')}", "#,##0")
gap_units_eq = item("gap_units_eq", f"={gap_units}-{t('I')}", "#,##0")
phrase = item("share_phrase", f'=IF(ISNUMBER({share_rev}),IF(ABS({share_rev}-0.5)<=0.1,"half the revenue gap",'
                              f'ROUND({share_rev}*100,0)&"% of the revenue gap"),"")')
multi = item("multi_note", f'=IF(OR({d_cy}<>{d_cy_min},{d_py}<>{d_py_min})," Days shown are for the longest-running '
                           f'market; estimates are computed market by market.","")')

sp = '&'.join(f'IF($B${r}=1,$A${r}&" ","")' for r in PR.values())
ex = '&'.join(f'IF($B${r}=0,$A${r}&" ","")' for r in PR.values())
sc_raw = item("scope_raw", f'=SUBSTITUTE(TRIM({sp})," ","|")')
scope_and = item("scope_and", f'=IF(LEN({sc_raw})-LEN(SUBSTITUTE({sc_raw},"|",""))=2,'
                              f'SUBSTITUTE(SUBSTITUTE({sc_raw},"|",", ",1),"|"," and "),SUBSTITUTE({sc_raw},"|"," and "))')
scope_plus = item("scope_plus", f'=SUBSTITUTE({sc_raw},"|"," + ")')
ex_raw = item("excl_raw", f'=SUBSTITUTE(TRIM({ex})," ","|")')
excluded = item("excluded", f'=IF({ex_raw}="","All products included",SUBSTITUTE({ex_raw},"|"," and ")&" excluded")')

launch_cy = item("t_launch_cy", f'=IF({lcmin}=0,"n/a",IF({lcmin}={lcmax},{dtxt(lcmin)},{dtxt(lcmin)}&" – "&{dtxt(lcmax)}))')
launch_py = item("t_launch_py", f'=IF({lpmin}=0,"n/a",IF({lpmin}={lpmax},{dtxt(lpmin)},{dtxt(lpmin)}&" – "&{dtxt(lpmax)}))')
rep_s = item("t_report_short", f"={dtxt(R)}")
rep_up = item("t_report_long", f'=IF({R}=0,"",UPPER(DAY({R})&" "&CHOOSE(MONTH({R}),{MONF})&" "&YEAR({R})))')

item("k_cover_left", f'=UPPER("Sales update · "&{market})')
item("k_cover_sub", '="First results "&Yr_Cur&" vs "&Yr_Prev')
item("k_cover_scope", f'="Scope: "&{scope_and}')
item("k_cover_l1", f'="Launch "&Yr_Cur&": "&{launch_cy}')
item("k_cover_l2", f'="Launch "&Yr_Prev&": "&{launch_py}')
item("k_summary", f'=UPPER("Summary · launch to "&{rep_s}&" "&Yr_Cur&" vs "&Yr_Prev)')
item("h_summary", f'=IF(ISNUMBER(Calc!$L${TOT}),IF(Calc!$L${TOT}<0,IF(AND({hs}>0,{phrase}<>""),'
                  f'"Behind "&Yr_Prev&", with "&{phrase}&" from a later launch","Behind "&Yr_Prev),"Ahead of "&Yr_Prev),'
                  f'"No comparison data yet — add rows on the Data sheet")')
item("h_launch", f'=IF({hs}>0,Yr_Prev&" had a "&{hs}&"-day head start"&IF({phrase}<>"",", worth about "&{phrase},""),'
                 f'IF({hs}<0,Yr_Cur&" launched "&-{hs}&" days earlier than "&Yr_Prev,"Both years launched on the same day"))')
item("h_total", f'=IF(ISNUMBER(Calc!$L${TOT}),IF(Calc!$L${TOT}<0,IF({hs}>0,"Total sales are behind "&Yr_Prev&", partly due to the later launch",'
                f'"Total sales are behind "&Yr_Prev),"Total sales are ahead of "&Yr_Prev),"No comparison data yet")')
item("h_size", f'=IF(AND(ISNUMBER(Calc!$L$6),ISNUMBER(Calc!$L$7)),IF(AND(Calc!$L$6<0,Calc!$L$7>0),"The 24-day drives the decline; the 12-day is growing",'
               f'IF(AND(Calc!$L$6>0,Calc!$L$7<0),"The 12-day drives the decline; the 24-day is growing",IF(AND(Calc!$L$6>=0,Calc!$L$7>=0),'
               f'"Both sizes are growing","Both sizes are behind "&Yr_Prev))),"Sales by size")')
item("k_launch", '="LAUNCH TIMING"')
item("k_total", f'="TOTAL · "&{scope_plus}')
item("k_size", f'="BY SIZE · LAUNCH TO "&UPPER({rep_s})')
item("strip_title", f'=IF({hs}>0,"Launch +"&{hs}&" days later",IF({hs}<0,"Launch "&-{hs}&" days earlier","Same launch timing"))')
item("strip_detail", f'={launch_cy}&" vs "&{launch_py}&" · "&{d_cy}&" vs "&{d_py}&" selling days · "&IF({hs}>0,Yr_Prev&" head start worth","Launch-timing effect")'
                     f'&" ~€"&FIXED(ROUND({t("J")},-2),0)&" and ~"&FIXED(ROUND({t("I")},0),0)&" units (est.)"')
item("fn_launch", f'=IF({hs}<=0,"Estimate: "&Yr_Prev&" sales pro rata for the difference in selling days. "&{scope_plus}&"; "&{excluded}&"."&{multi},"Estimate: "&Yr_Prev&" sales to "&{rep_s}&" spread evenly over "&{d_py}&" days (× "&{hs}&"/"&{d_py}&"). '
                  f'Early days are likely slower, so this is an upper bound. "&{scope_plus}&"; "&{excluded}&"."&{multi})')
item("fn_total", f'=Yr_Prev&" same-days figures estimated pro rata ("&Yr_Prev&" to "&{rep_s}&" × "&{d_cy}&"/"&{d_py}&"). "&{excluded}&"."&{multi}')
item("fn_size", f'="* First "&{d_cy}&" days after launch; "&Yr_Prev&" estimated pro rata. Gains shown in blue. "&{excluded}&"."&{multi}')
item("sub_total_a", f'="Launch to "&{rep_s}&" · "&{d_cy}&" vs "&{d_py}&" selling days"')
item("sub_total_b", f'="First "&{d_cy}&" days · "&Yr_Prev&" estimated"')
item("vs_rev", f'="vs €"&FIXED({t("F")},0)&" in "&Yr_Prev')
item("vs_units", f'="vs "&FIXED({t("E")},0)&" in "&Yr_Prev')
item("vs_rev_eq", f'="vs ~€"&FIXED(ROUND({t("H")},-2),0)&" in "&Yr_Prev')
item("vs_units_eq", f'="vs ~"&FIXED(ROUND({t("G")},0),0)&" in "&Yr_Prev')
item("st_l1", '=UPPER("Head-start revenue "&Yr_Prev)')
item("st_l2", '=UPPER("Head-start units "&Yr_Prev)')
item("st_v1", f'="~€"&FIXED(ROUND({t("J")},-2),0)')
item("st_v2", f'="~"&FIXED(ROUND({t("I")},0),0)')
item("st_v3", f'="~€"&FIXED(ROUND({gap_rev_eq},-2),0)')
item("st_s1", f'=IF(ISNUMBER({share_rev}),ROUND({share_rev}*100,0)&"% of the €"&FIXED({gap_rev},0)&" revenue gap","no revenue gap")')
item("st_s2", f'=IF(ISNUMBER({share_units}),ROUND({share_units}*100,0)&"% of the "&FIXED({gap_units},0)&"-unit gap","no unit gap")')
item("st_s3", f'="and ~"&FIXED(ROUND({gap_units_eq},0),0)&" units, on equal days"')
item("cut_note", f'="Both periods end on "&{rep_s}')

# chart block
cb = item_row[0] + 2
cs[f"A{cb - 1}"] = "Launch-timing chart data"
cs[f"A{cb - 1}"].font = F(BODY, 10, True)
cs[f"A{cb}"] = f'=Yr_Prev&"  ·  launch "&{launch_py}&"  ·  "&{d_py}&" selling days"'
cs[f"A{cb + 1}"] = f'=Yr_Cur&"  ·  launch "&{launch_cy}&"  ·  "&{d_cy}&" selling days"'
cs[f"B{cb - 1}"] = f'="+"&MAX({hs},0)&" days"'
cs[f"C{cb - 1}"] = "Selling days"
cs[f"B{cb}"] = f"=MAX({hs},0)"
cs[f"B{cb + 1}"] = f"=MAX({hs},0)"
cs[f"C{cb}"] = f"=IF({hs}>0,{d_cy},{d_py})"
cs[f"C{cb + 1}"] = f"={d_cy}"
cs.column_dimensions["B"].width = 60

# ═════════════════════════════════════════════════════════════════════════
# SLIDE SHEETS
# ═════════════════════════════════════════════════════════════════════════
def new_slide(title, bg=BG):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = A700
    for c in range(1, NCOL + 1):
        ws.column_dimensions[ws.cell(1, c).column_letter].width = 2.3
    for r in range(1, NROW + 1):
        ws.row_dimensions[r].height = 15
        for c in range(1, NCOL + 1):
            ws.cell(r, c).fill = fill(bg)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:{ws.cell(1, NCOL).column_letter}{NROW}"
    return ws


def box(ws, r1, c1, r2, c2, v=None, font=None, h="left", va="center", wrap=False, fmt=None, shrink=False, bg=None):
    if (r1, c1) != (r2, c2):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    if bg:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(r, c).fill = fill(bg)
    cell = ws.cell(r1, c1)
    cell.value = v
    cell.font = font or F()
    cell.alignment = Alignment(horizontal=h, vertical=va, wrap_text=wrap, shrink_to_fit=shrink and not wrap)
    if fmt:
        cell.number_format = fmt
    return cell


def edge(ws, r1, c1, r2, c2, color=HAIR, sides="tblr"):
    s = Side("thin", color=color)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            b = cell.border
            cell.border = Border(
                left=s if (c == c1 and "l" in sides) else b.left,
                right=s if (c == c2 and "r" in sides) else b.right,
                top=s if (r == r1 and "t" in sides) else b.top,
                bottom=s if (r == r2 and "b" in sides) else b.bottom,
            )


def blueprint(ws, r1, c1, r2, c2):
    edge(ws, r1, c1, r2, c2)
    mk = F(BODY, 9, False, "7A7A7D")
    for (r, c, hh, vv) in ((r1 - 1, c1 - 1, "right", "bottom"), (r1 - 1, c2 + 1, "left", "bottom"),
                           (r2 + 1, c1 - 1, "right", "top"), (r2 + 1, c2 + 1, "left", "top")):
        cell = ws.cell(r, c)
        cell.value = "+"
        cell.font = mk
        cell.alignment = Alignment(horizontal=hh, vertical=vv)


def header(ws, kicker_key, head_key, kick_color=A700):
    box(ws, 3, 5, 3, 38, f"={REF[kicker_key]}", F(HEAD, 14, True, kick_color))
    box(ws, 3, 39, 3, 56, f'="MARKET: "&UPPER({REF["market"]})', F(HEAD, 14, True, MUTED), h="right")
    box(ws, 4, 5, 9, 56, f"={REF[head_key]}", F(HEAD, 32, True), va="top", wrap=True)


def val(r, col):
    return f"Calc!${col}${r}"


def guard(r, col):
    return f'=IF(Calc!$B${r}=1,{val(r, col)},"–")'


# ── 1 Cover ──────────────────────────────────────────────────────────────
s1 = new_slide("1 Cover", A900)
box(s1, 4, 5, 4, 38, f"={REF['k_cover_left']}", F(HEAD, 14, True, A200))
box(s1, 4, 39, 4, 56, f"={REF['t_report_long']}", F(HEAD, 14, True, A200), h="right")
box(s1, 12, 5, 19, 56, "Advent Calendar Sales", F(HEAD, 72, True, BG), shrink=True, va="bottom")
box(s1, 21, 5, 22, 56, f"={REF['k_cover_sub']}", F(BODY, 21, False, A100), va="top")
for (ca, cb_), key in (((5, 21), "k_cover_scope"), ((22, 38), "k_cover_l1"), ((39, 56), "k_cover_l2")):
    box(s1, 30, ca, 31, cb_, f"={REF[key]}", F(BODY, 14, False, A100), va="center")
edge(s1, 30, 5, 30, 56, "5D6874", "t")

# ── 2 Summary ────────────────────────────────────────────────────────────
s2 = new_slide("2 Summary")
header(s2, "k_summary", "h_summary")
r0 = 11
cards = [("Total", TOT, (5, 20)), (None, PR["24-day"], (23, 38)), (None, PR["12-day"], (41, 56))]
for name, row, (c1, c2) in cards:
    blueprint(s2, r0, c1, r0 + 17, c2)
    flag = f"Calc!$B${row}"
    title = "Total" if name else f"=Calc!$A${row}"
    box(s2, r0 + 1, c1 + 1, r0 + 2, c2 - 1, title, F(HEAD, 20, True))
    box(s2, r0 + 3, c1 + 1, r0 + 6, c2 - 1, guard(row, "D"), F(HEAD, 44, True), fmt=EUR0, shrink=True, va="center")
    box(s2, r0 + 10, c1 + 1, r0 + 12, c2 - 1,
        f'=IF({flag}=1,{val(row, "C")},"–")', F(HEAD, 30, True), fmt=INT + '" units"', shrink=True)

    def gline(rr, v, txt):
        box(s2, rr, c1 + 1, rr, c1 + 5, v, F(BODY, 12, True, A700), fmt=PCT, shrink=True)
        box(s2, rr, c1 + 6, rr, c2 - 1, txt, F(BODY, 12))

    gline(r0 + 7, guard(row, "L"), "same date")
    gline(r0 + 8, guard(row, "N"), "equal days")
    if name:
        gline(r0 + 13, guard(row, "K"), "same date")
        gline(r0 + 14, guard(row, "M"), "equal days")
    else:
        gline(r0 + 13, guard(row, "K"), '="units vs "&Yr_Prev')
        gline(r0 + 14, guard(row, "Q"), "avg price")
r1 = 31
for _r in range(r1, r1 + 4):
    for _c in range(5, 57):
        s2.cell(_r, _c).fill = fill(A900)
box(s2, r1, 6, r1 + 3, 22, f"={REF['strip_title']}", F(HEAD, 18, True, BG), bg=A900, shrink=True)
box(s2, r1, 24, r1 + 3, 55, f"={REF['strip_detail']}", F(BODY, 12, False, BG), wrap=True, bg=A900)

# ── 3 Launch timing ──────────────────────────────────────────────────────
s3 = new_slide("3 Launch timing")
header(s3, "k_launch", "h_launch")
blueprint(s3, 11, 5, 22, 56)
box(s3, 22, 6, 22, 55, f"={REF['cut_note']}", F(BODY, 12, False, MUTED), h="right")
for (ca, cb_), (lk, vk, sk) in zip(((5, 20), (23, 38), (41, 56)),
                                  (("st_l1", "st_v1", "st_s1"), ("st_l2", "st_v2", "st_s2"), (None, "st_v3", "st_s3"))):
    box(s3, 25, ca, 25, cb_, f"={REF[lk]}" if lk else "GAP WITHOUT HEAD START", F(BODY, 14, False, A700))
    box(s3, 26, ca, 28, cb_, f"={REF[vk]}", F(HEAD, 40, True), shrink=True)
    box(s3, 29, ca, 31, cb_, f"={REF[sk]}", F(BODY, 14, False, MUTED), va="top", wrap=True)
box(s3, 33, 5, 34, 56, f"={REF['fn_launch']}", F(BODY, 11, False, MUTED), wrap=True, va="top")

ch = BarChart()
ch.type = "bar"
ch.grouping = "stacked"
ch.overlap = 100
ch.gapWidth = 45
ch.legend = None
ch.width, ch.height = 27.6, 4.9
s_hs = Series(Reference(cs, min_col=2, min_row=cb - 1, max_row=cb + 1), title_from_data=True)
s_dy = Series(Reference(cs, min_col=3, min_row=cb - 1, max_row=cb + 1), title_from_data=True)
ch.series.append(s_hs)
ch.series.append(s_dy)
ch.set_categories(Reference(cs, min_col=1, min_row=cb, max_row=cb + 1))
s_hs.graphicalProperties = GraphicalProperties(solidFill=A900)
s_hs.graphicalProperties.line = LineProperties(noFill=True)
s_dy.graphicalProperties = GraphicalProperties(solidFill=A200)
s_dy.graphicalProperties.line = LineProperties(noFill=True)
pt = DataPoint(idx=1)
pt.graphicalProperties = GraphicalProperties(noFill=True)
pt.graphicalProperties.line = LineProperties(solidFill=A700, prstDash="dash", w=12700)
s_hs.dPt.append(pt)
cp_w = CharacterProperties(sz=1200, b=True, solidFill="FFFFFF", latin=DFont(typeface=BODY))
s_hs.dLbls = DataLabelList()
s_hs.dLbls.showSerName = True
s_hs.dLbls.showVal = s_hs.dLbls.showCatName = s_hs.dLbls.showLegendKey = s_hs.dLbls.showPercent = False
s_hs.dLbls.dLbl = [DataLabel(idx=1, showSerName=False, showVal=False, showCatName=False, showLegendKey=False, showPercent=False)]
s_hs.dLbls.position = "ctr"
s_hs.dLbls.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp_w), endParaRPr=cp_w)])
cp_d = CharacterProperties(sz=1200, b=True, solidFill=TEXT, latin=DFont(typeface=HEAD))
ch.x_axis.scaling.orientation = "maxMin"
ch.x_axis.delete = False
ch.x_axis.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp_d), endParaRPr=cp_d)])
ch.x_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
ch.y_axis.delete = True
ch.y_axis.scaling.min = 0
ch.y_axis.majorGridlines = None
ch.graphical_properties = GraphicalProperties(noFill=True)
ch.graphical_properties.line = LineProperties(noFill=True)
ch.plot_area.graphicalProperties = GraphicalProperties(noFill=True)
ch.anchor = TwoCellAnchor(editAs="twoCell", _from=AnchorMarker(col=5, row=11, colOff=0, rowOff=0),
                          to=AnchorMarker(col=55, row=21, colOff=0, rowOff=0))
s3.add_chart(ch)

# ── 4 Total ──────────────────────────────────────────────────────────────
s4 = new_slide("4 Total")
header(s4, "k_total", "h_total")
for (c1, c2), (tt, sk, rk, uk, gR, gU) in zip(
        ((5, 28), (33, 56)),
        (("Same calendar date", "sub_total_a", "vs_rev", "vs_units", "L", "K"),
         ("Same days since launch", "sub_total_b", "vs_rev_eq", "vs_units_eq", "N", "M"))):
    blueprint(s4, 11, c1, 26, c2)
    box(s4, 12, c1 + 1, 13, c2 - 1, tt, F(HEAD, 20, True))
    box(s4, 14, c1 + 1, 14, c2 - 1, f"={REF[sk]}", F(BODY, 14, False, MUTED))
    for k, (m0, m1, lab, bigcol, gcol, txt, fm) in enumerate((
            (c1 + 1, c1 + 11, "REVENUE", "D", gR, rk, EUR0),
            (c1 + 13, c1 + 22, "UNITS", "C", gU, uk, INT))):
        box(s4, 17, m0, 17, m1, lab, F(BODY, 14, False, A700))
        box(s4, 18, m0, 21, m1, guard(TOT, bigcol), F(HEAD, 42, True), fmt=fm, shrink=True)
        box(s4, 22, m0, 23, m0 + 5, guard(TOT, gcol), F(HEAD, 20, True, A700), fmt=PCT, shrink=True)
        box(s4, 25, m0, 25, m1, f"={REF[txt]}", F(BODY, 12, False, MUTED), shrink=True)
box(s4, 29, 5, 30, 56, f"={REF['fn_total']}", F(BODY, 11, False, MUTED), wrap=True, va="top")

# ── 5 By size ────────────────────────────────────────────────────────────
s5 = new_slide("5 By size")
header(s5, "k_size", "h_size")
blueprint(s5, 11, 5, 26, 56)
cols = [(6, 13), (14, 19), (20, 25), (26, 31), (32, 37), (38, 43), (44, 49), (50, 55)]
hd = ["SIZE", '="UNITS "&Yr_Cur', '="UNITS VS "&Yr_Prev', '="REVENUE "&Yr_Cur', "REV. SAME DATE", "REV. SAME DAYS*", "AVG PRICE",
      '="PRICE VS "&Yr_Prev']
for i, ((ca, cb_), h) in enumerate(zip(cols, hd)):
    box(s5, 12, ca, 13, cb_, h, F(BODY, 11, False, MUTED), h="left" if i == 0 else "right", wrap=True, va="center")
edge(s5, 13, 6, 13, 55, HAIR, "b")
rows5 = [(PR["24-day"], 14), (PR["12-day"], 17), (PR["Refill"], 20), (TOT, 23)]
colmap = [None, "C", "K", "D", "L", "N", "O", "Q"]
fmts = [None, INT, PCT, EUR0, PCT, PCT, EUR2, PCT]
for (crow, sr) in rows5:
    is_tot = crow == TOT
    for i, ((ca, cb_), cm) in enumerate(zip(cols, colmap)):
        if i == 0:
            box(s5, sr, ca, sr + 2, cb_, "Total" if is_tot else f"=Calc!$A${crow}", F(HEAD, 20, True))
        else:
            box(s5, sr, ca, sr + 2, cb_, guard(crow, cm), F(BODY, 16, is_tot), h="right", fmt=fmts[i], shrink=True)
    if not is_tot:
        edge(s5, sr + 2, 6, sr + 2, 55, ROWLINE, "b")
green = FormulaRule
for ci in (2, 4, 5, 7):
    ca, cb_ = cols[ci]
    L1 = s5.cell(14, ca).coordinate
    rng = f"{L1}:{s5.cell(25, ca).coordinate}"
    s5.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({L1}),{L1}>0)"],
                                                   font=Font(color=A700, bold=True)))
for ci in range(1, 8):
    ca, _ = cols[ci]
    L1 = s5.cell(14, ca).coordinate
    s5.conditional_formatting.add(f"{L1}:{s5.cell(25, ca).coordinate}",
                                  FormulaRule(formula=[f'{L1}="–"'], font=Font(color="98989B")))
box(s5, 28, 5, 29, 56, f"={REF['fn_size']}", F(BODY, 11, False, MUTED), wrap=True, va="top")

# sheet order: Control, Data, slides, Calc
order = ["Control", "Data", "1 Cover", "2 Summary", "3 Launch timing", "4 Total", "5 By size", "Calc"]
wb._sheets = [wb[n] for n in order]
wb.active = 0
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved", OUT)
