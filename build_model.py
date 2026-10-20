"""Builds nike_forecast_model.xlsx: a five-year, three-statement forecast for NIKE, Inc.

History (FY2024 to FY2026) comes from Nike's Form 10-K filings and the fiscal 2026 results
release (revenue by region). The forecast (FY2027 to FY2031) is driven by one Assumptions
sheet with base, upside and downside scenarios: revenue is built up by region, then costs,
working capital, capital spending, dividends and buybacks flow through a linked income
statement, balance sheet and cash flow statement. Every forecast cell is a formula.

Run:  python build_model.py
"""
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

OUT = "nike_forecast_model.xlsx"
YEARS = ["FY2024A", "FY2025A", "FY2026A", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
COLS = ["B", "C", "D", "E", "F", "G", "H", "I"]
HIST, PROJ = COLS[:3], COLS[3:]
FIRSTP, LASTA = PROJ[0], HIST[-1]

# ---------- reported data ($ millions; FY2024 to FY2026, from the 10-K) ----------
H = {
    "rev": [51362, 46309, 46398], "cogs": [28475, 26519, 26487], "dc": [4285, 4689, 4754], "oh": [12291, 11399, 11360],
    "int": [-161, -107, -50], "oth": [-228, -76, -53], "tax": [1000, 666, 792], "shares": [1529.7, 1487.6, 1481.0],
    "cash": [9860, 7464, 7563], "sti": [1722, 1687, 1464], "ar": [4427, 4717, 5931], "inv": [7519, 7489, 7501],
    "prep": [1854, 2005, 2144], "ppe": [5000, 4828, 4796], "rou": [2718, 2712, 2838], "intang": [499, 499, 499],
    "oa": [4511, 5178, 5674],
    "cdebt": [1006, 5, 2000], "ap": [2851, 3479, 3600], "accr": [5725, 5911, 6092], "itp": [534, 669, 377],
    "cll": [477, 502, 478], "ltd": [7903, 7961, 5942], "ll": [2566, 2550, 2613], "ol": [2618, 2289, 2443],
    "capital": [13412, 14198, 15161], "aoci": [53, -258, -141], "re": [965, -727, -155],
    "da": [796, 775, 747], "sbc": [804, 709, 715], "cfo": [7429, 3698, 2868], "capex": [-812, -430, -684],
    "cfi": [894, -275, -488], "repay": [0, -1001, 0], "div": [-2169, -2300, -2407], "bb": [-4250, -2985, -146],
    "cff": [-5888, -5820, -2292], "fx": [-16, 1, 11], "cash_open": [7441, 9860, 7464],
}
# Revenue by region (FY2025, FY2026), from the fiscal 2026 results release
SEG = ["North America", "Europe, Middle East & Africa", "Greater China", "Asia Pacific & Latin America", "Converse"]
SEG_HIST = {"North America": [19572, 20511], "Europe, Middle East & Africa": [12257, 12572], "Greater China": [6586, 5847],
            "Asia Pacific & Latin America": [6251, 6243], "Converse": [1692, 1174]}
OTHER_HIST = [48 - 97, 49 + 2]   # Global Brand Divisions + Corporate
# Growth assumptions, FY2027 to FY2031, by scenario (base, upside, downside)
GROWTH = {
    "North America":                [[-.06, .03, .05, .05, .04], [-.03, .05, .07, .06, .05], [-.09, 0, .02, .03, .03]],
    "Europe, Middle East & Africa": [[-.08, .02, .04, .04, .04], [-.05, .04, .06, .05, .05], [-.11, -.02, .01, .02, .02]],
    "Greater China":                [[-.20, -.03, .03, .04, .04], [-.15, .03, .06, .06, .05], [-.26, -.10, -.03, 0, .01]],
    "Asia Pacific & Latin America": [[-.06, .03, .05, .05, .05], [-.03, .05, .07, .07, .06], [-.09, 0, .02, .03, .03]],
    "Converse":                     [[-.15, -.05, 0, .02, .02], [-.10, 0, .03, .04, .03], [-.20, -.10, -.03, 0, 0]],
}
GM = [[.425, .430, .435, .440, .445], [.430, .440, .450, .455, .460], [.420, .420, .425, .430, .430]]

# ---------- styles ----------
AR = "Arial"
F = Font(name=AR, size=10)
F_IN = Font(name=AR, size=10, color="0000FF")
F_LINK = Font(name=AR, size=10, color="008000")
F_B = Font(name=AR, size=10, bold=True)
F_HEAD = Font(name=AR, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=AR, size=14, bold=True, color="1F4E79")
F_SEC = Font(name=AR, size=10, bold=True, color="1F4E79")
F_NOTE = Font(name=AR, size=9, italic=True, color="595959")
FILL_HEAD = PatternFill("solid", fgColor="1F4E79")
FILL_SEC = PatternFill("solid", fgColor="DCE6F1")
FILL_KEY = PatternFill("solid", fgColor="FFFF00")
FILL_HIST = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="808080")
TOP = Border(top=THIN)
TOTAL = Border(top=THIN, bottom=Side(style="double", color="808080"))
M = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
DAYS = '0.0'
EPS = '$0.00'


def prev(c):
    return COLS[COLS.index(c) - 1]


def sheet_header(ws, title, sub):
    ws["A1"], ws["A2"] = title, sub
    ws["A1"].font, ws["A2"].font = F_TITLE, F_NOTE
    ws.cell(row=4, column=1, value="$ millions unless stated").font = F_HEAD
    ws.cell(row=4, column=1).fill = FILL_HEAD
    for i, y in enumerate(YEARS):
        c = ws.cell(row=4, column=2 + i, value=y)
        c.font, c.fill, c.alignment = F_HEAD, FILL_HEAD, Alignment(horizontal="right")
    ws.column_dimensions["A"].width = 46
    for col in COLS:
        ws.column_dimensions[col].width = 11
    ws.freeze_panes = "B5"
    ws.sheet_view.showGridLines = False


def section(ws, r, text, width=9):
    for c in range(1, width + 1):
        ws.cell(row=r, column=c).fill = FILL_SEC
    ws.cell(row=r, column=1, value=text).font = F_SEC


class Sheet:
    """Small helper that writes labelled rows and remembers where each key lives."""

    def __init__(self, ws, start=5):
        self.ws, self.r, self.rows = ws, start, {}

    def sec(self, text):
        section(self.ws, self.r, text)
        self.r += 1

    def gap(self):
        self.r += 1

    def line(self, key, label, hist=None, proj=None, fmt=M, bold=False, border=None, hist_formula=None, note=None):
        ws, r = self.ws, self.r
        self.rows[key] = r
        ws.cell(row=r, column=1, value=label).font = F_B if bold else F
        for i, col in enumerate(COLS):
            c = ws[f"{col}{r}"]
            if col in HIST:
                if hist_formula:
                    c.value, c.font = hist_formula(col, i), F_B if bold else F
                elif hist is not None:
                    c.value, c.font = hist[i], F_IN
                c.fill = FILL_HIST
            elif proj:
                v = proj(col)
                c.value = v
                c.font = F_LINK if ("!" in str(v) and not bold) else (F_B if bold else F)
            c.number_format = fmt
            if border:
                c.border = border
        if note:
            ws.cell(row=r, column=10, value=note).font = F_NOTE
        self.r += 1
        return r

    def ref(self, key, col, sheet=None):
        a = f"{col}{self.rows[key]}"
        return f"'{sheet}'!{a}" if sheet else a


def build():
    wb = Workbook()
    cover = wb.active
    cover.title = "Cover"
    ws_sum = wb.create_sheet("Summary")
    ws_a = wb.create_sheet("Assumptions")
    ws_rev = wb.create_sheet("Revenue Build")
    ws_is = wb.create_sheet("Income Statement")
    ws_bs = wb.create_sheet("Balance Sheet")
    ws_cf = wb.create_sheet("Cash Flow")
    ws_sc = wb.create_sheet("Scenarios")
    ws_ck = wb.create_sheet("Checks")

    # ================= Assumptions =================
    sheet_header(ws_a, "Assumptions", "Blue = input you can change. Yellow = key assumption. Forecast years only; history shown grey for reference.")
    ws_a["A6"], ws_a["B6"] = "Scenario (1 = Base, 2 = Upside, 3 = Downside)", 1
    ws_a["A6"].font = F_B
    ws_a["B6"].font, ws_a["B6"].fill = F_IN, FILL_KEY
    ws_a["C6"] = '=CHOOSE(B6,"Base case","Upside case","Downside case")'
    ws_a["C6"].font = F_B
    SCN = "Assumptions!$B$6"
    A = Sheet(ws_a, start=8)
    A.sec("Revenue growth by region: selected scenario")
    for s in SEG:
        A.line(("g", s), s, proj=lambda col, s=s: f"=CHOOSE({SCN},{A.ref(('gb', s), col) if False else '{b}'},{{u}},{{d}})", fmt=PCT)
    A.gap()
    # placeholders filled after the three scenario tables exist
    sel_rows = {s: A.rows[("g", s)] for s in SEG}
    tables = {}
    for si, sname in enumerate(["Base case", "Upside case", "Downside case"]):
        A.sec(f"Revenue growth by region: {sname}")
        for s in SEG:
            A.line((sname, s), s, proj=lambda col, s=s, si=si: GROWTH[s][si][PROJ.index(col)], fmt=PCT)
            for col in PROJ:
                ws_a[f"{col}{A.rows[(sname, s)]}"].font = F_IN
        tables[sname] = {s: A.rows[(sname, s)] for s in SEG}
        A.gap()
    for s in SEG:
        r = sel_rows[s]
        for col in PROJ:
            ws_a[f"{col}{r}"] = (f"=CHOOSE({SCN},{col}{tables['Base case'][s]},{col}{tables['Upside case'][s]},"
                                 f"{col}{tables['Downside case'][s]})")
            ws_a[f"{col}{r}"].font = F_B
    A.line("other", "Global Brand Divisions and Corporate ($m, held flat)", proj=lambda col: OTHER_HIST[1], fmt=M)
    for col in PROJ:
        ws_a[f"{col}{A.rows['other']}"].font = F_IN
    A.gap()

    A.sec("Margins and costs (% of revenue)")
    gm_rows = []
    for si, sname in enumerate(["Base", "Upside", "Downside"]):
        gm_rows.append(A.line(("gm", sname), f"Gross margin: {sname.lower()}", proj=lambda col, si=si: GM[si][PROJ.index(col)], fmt=PCT))
        for col in PROJ:
            ws_a[f"{col}{gm_rows[-1]}"].font = F_IN
    A.line("gm", "Gross margin: selected scenario",
           proj=lambda col: f"=CHOOSE({SCN},{col}{gm_rows[0]},{col}{gm_rows[1]},{col}{gm_rows[2]})", fmt=PCT, bold=True)
    for key, label, vals in [("dc", "Demand creation (marketing)", [.105, .102, .099, .097, .096]),
                             ("oh", "Operating overhead", [.273, .260, .250, .243, .238]),
                             ("da", "Depreciation and amortization", [.016] * 5),
                             ("sbc", "Stock-based compensation (non-cash, inside overhead)", [.015] * 5),
                             ("capex", "Capital expenditure", [.016] * 5)]:
        A.line(key, label, proj=lambda col, vals=vals: vals[PROJ.index(col)], fmt=PCT)
        for col in PROJ:
            ws_a[f"{col}{A.rows[key]}"].font = F_IN
    A.gap()
    A.sec("Tax and interest")
    for key, label, v, fmt in [("taxr", "Effective tax rate", .19, PCT), ("rdebt", "Interest rate on debt", .04, PCT),
                               ("rcash", "Interest earned on cash and investments", .035, PCT)]:
        A.line(key, label, proj=lambda col, v=v: v, fmt=fmt)
        for col in PROJ:
            ws_a[f"{col}{A.rows[key]}"].font = F_IN
    A.gap()
    A.sec("Working capital")
    for key, label, vals, fmt in [("dso", "Receivable days (on revenue)", [44, 42, 40, 40, 40], DAYS),
                                  ("dio", "Inventory days (on cost of sales)", [100, 98, 96, 95, 95], DAYS),
                                  ("dpo", "Payable days (on cost of sales)", [50] * 5, DAYS),
                                  ("prep", "Prepaid and other current assets (% of revenue)", [.046] * 5, PCT),
                                  ("accr", "Accrued liabilities (% of revenue)", [.131] * 5, PCT)]:
        A.line(key, label, proj=lambda col, vals=vals: vals[PROJ.index(col)], fmt=fmt)
        for col in PROJ:
            ws_a[f"{col}{A.rows[key]}"].font = F_IN
    A.gap()
    A.sec("Capital allocation")
    for key, label, vals, fmt in [("repay", "Debt repaid ($m)", [2000, 0, 0, 0, 0], M),
                                  ("dpsg", "Dividend per share growth", [0, 0, .03, .05, .05], PCT),
                                  ("bb", "Share buybacks ($m)", [250, 750, 1500, 2500, 3000], M),
                                  ("bbp", "Average buyback share price ($)", [40, 45, 50, 55, 60], EPS),
                                  ("issue", "New shares issued to employees (m)", [8] * 5, '#,##0.0')]:
        A.line(key, label, proj=lambda col, vals=vals: vals[PROJ.index(col)], fmt=fmt)
        for col in PROJ:
            ws_a[f"{col}{A.rows[key]}"].font = F_IN
    for key in ("gm", "dc", "oh"):
        for col in PROJ:
            ws_a[f"{col}{A.rows[key]}"].fill = FILL_KEY
    ws_a.column_dimensions["A"].width = 50
    a = lambda key, col: f"Assumptions!{col}{A.rows[key]}"
    ag = lambda s, col: f"Assumptions!{col}{sel_rows[s]}"

    # ================= Revenue build =================
    sheet_header(ws_rev, "Revenue Build by Region", "FY2025 and FY2026 actual revenue by region from Nike's fiscal 2026 results; forecast = prior year x (1 + growth).")
    R = Sheet(ws_rev)
    R.sec("Revenue ($m)")
    for s in SEG:
        R.line(("seg", s), s, proj=lambda col, s=s: f"={prev(col)}{R.rows[('seg', s)] if ('seg', s) in R.rows else 0}*(1+{ag(s, col)})")
    # fill properly now that rows exist (hist FY2024 not split by region in this model)
    for s in SEG:
        r = R.rows[("seg", s)]
        ws_rev[f"B{r}"] = None
        ws_rev[f"C{r}"], ws_rev[f"D{r}"] = SEG_HIST[s]
        ws_rev[f"C{r}"].font = ws_rev[f"D{r}"].font = F_IN
        for col in PROJ:
            ws_rev[f"{col}{r}"] = f"={prev(col)}{r}*(1+{ag(s, col)})"
            ws_rev[f"{col}{r}"].font = F
    R.line("other", "Global Brand Divisions and Corporate", proj=lambda col: f"={a('other', col)}")
    ro = R.rows["other"]
    ws_rev[f"B{ro}"] = None
    ws_rev[f"C{ro}"], ws_rev[f"D{ro}"] = OTHER_HIST
    ws_rev[f"C{ro}"].font = ws_rev[f"D{ro}"].font = F_IN
    first_seg = R.rows[("seg", SEG[0])]
    R.line("rev", "Total revenue", hist=H["rev"], proj=lambda col: f"=SUM({col}{first_seg}:{col}{ro})", bold=True, border=TOP)
    for col in ("C", "D"):
        ws_rev[f"{col}{R.rows['rev']}"] = f"=SUM({col}{first_seg}:{col}{ro})"
        ws_rev[f"{col}{R.rows['rev']}"].font = F_B
    R.gap()
    R.sec("Growth")
    for s in SEG:
        R.line(("gr", s), s, fmt=PCT, proj=lambda col, s=s: f"={col}{R.rows[('seg', s)]}/{prev(col)}{R.rows[('seg', s)]}-1")
        ws_rev[f"D{R.rows[('gr', s)]}"] = f"=D{R.rows[('seg', s)]}/C{R.rows[('seg', s)]}-1"
    R.line("grt", "Total revenue growth", fmt=PCT, bold=True,
           hist_formula=lambda col, i: None if i == 0 else f"={col}{R.rows['rev']}/{prev(col)}{R.rows['rev']}-1",
           proj=lambda col: f"={col}{R.rows['rev']}/{prev(col)}{R.rows['rev']}-1")
    R.gap()
    R.sec("Revenue mix")
    for s in SEG:
        R.line(("mix", s), s, fmt=PCT, proj=lambda col, s=s: f"={col}{R.rows[('seg', s)]}/{col}{R.rows['rev']}")
        for col in ("C", "D"):
            ws_rev[f"{col}{R.rows[('mix', s)]}"] = f"={col}{R.rows[('seg', s)]}/{col}{R.rows['rev']}"

    # ================= Income statement =================
    sheet_header(ws_is, "Income Statement", "FY2024 to FY2026 as reported (10-K). FY2027 to FY2031 forecast from the Assumptions sheet.")
    I = Sheet(ws_is)
    I.line("rev", "Revenues", hist=H["rev"], proj=lambda col: f"='Revenue Build'!{col}{R.rows['rev']}")
    I.line("cogs", "Cost of sales", hist=H["cogs"], proj=lambda col: f"={col}{I.rows['rev']}*(1-{a('gm', col)})")
    I.line("gp", "Gross profit", hist_formula=lambda col, i: f"={col}{I.rows['rev']}-{col}{I.rows['cogs']}",
           proj=lambda col: f"={col}{I.rows['rev']}-{col}{I.rows['cogs']}", bold=True, border=TOP)
    I.line("dc", "Demand creation expense", hist=H["dc"], proj=lambda col: f"={col}{I.rows['rev']}*{a('dc', col)}")
    I.line("oh", "Operating overhead expense", hist=H["oh"], proj=lambda col: f"={col}{I.rows['rev']}*{a('oh', col)}")
    I.line("ebit", "Operating income (EBIT)", hist_formula=lambda col, i: f"={col}{I.rows['gp']}-{col}{I.rows['dc']}-{col}{I.rows['oh']}",
           proj=lambda col: f"={col}{I.rows['gp']}-{col}{I.rows['dc']}-{col}{I.rows['oh']}", bold=True, border=TOP)
    I.line("int", "Interest expense (income), net", hist=H["int"],
           proj=lambda col: f"={a('rdebt', col)}*('Balance Sheet'!{prev(col)}{{cdebt}}+'Balance Sheet'!{prev(col)}{{ltd}})"
                            f"-{a('rcash', col)}*('Balance Sheet'!{prev(col)}{{cash}}+'Balance Sheet'!{prev(col)}{{sti}})",
           note="Forecast: opening debt x rate less opening cash x rate (avoids a circular reference)")
    I.line("oth", "Other (income) expense, net", hist=H["oth"], proj=lambda col: 0, note="Forecast: assumed nil")
    I.line("pbt", "Income before income taxes", hist_formula=lambda col, i: f"={col}{I.rows['ebit']}-{col}{I.rows['int']}-{col}{I.rows['oth']}",
           proj=lambda col: f"={col}{I.rows['ebit']}-{col}{I.rows['int']}-{col}{I.rows['oth']}", bold=True, border=TOP)
    I.line("tax", "Income tax expense", hist=H["tax"], proj=lambda col: f"={col}{I.rows['pbt']}*{a('taxr', col)}")
    I.line("ni", "Net income", hist_formula=lambda col, i: f"={col}{I.rows['pbt']}-{col}{I.rows['tax']}",
           proj=lambda col: f"={col}{I.rows['pbt']}-{col}{I.rows['tax']}", bold=True, border=TOTAL)
    I.gap()
    I.line("shares", "Diluted shares (m)", hist=H["shares"], fmt='#,##0.0',
           proj=lambda col: f"={prev(col)}{I.rows['shares']}-{a('bb', col)}/{a('bbp', col)}+{a('issue', col)}")
    I.line("eps", "Diluted EPS ($)", fmt=EPS, hist_formula=lambda col, i: f"={col}{I.rows['ni']}/{col}{I.rows['shares']}",
           proj=lambda col: f"={col}{I.rows['ni']}/{col}{I.rows['shares']}", bold=True)
    I.line("dps", "Dividends per share ($)", fmt=EPS, hist_formula=lambda col, i: f"=-'Cash Flow'!{col}{{div}}/{col}{I.rows['shares']}",
           proj=lambda col: f"={prev(col)}{I.rows['dps']}*(1+{a('dpsg', col)})")
    I.gap()
    I.sec("Margins")
    for key, label, num in [("gm_p", "Gross margin", "gp"), ("dc_p", "Demand creation % of revenue", "dc"),
                            ("oh_p", "Operating overhead % of revenue", "oh"), ("ebit_p", "EBIT margin", "ebit"), ("ni_p", "Net margin", "ni")]:
        I.line(key, label, fmt=PCT, hist_formula=lambda col, i, num=num: f"={col}{I.rows[num]}/{col}{I.rows['rev']}",
               proj=lambda col, num=num: f"={col}{I.rows[num]}/{col}{I.rows['rev']}")
    I.line("taxr_p", "Effective tax rate", fmt=PCT, hist_formula=lambda col, i: f"={col}{I.rows['tax']}/{col}{I.rows['pbt']}",
           proj=lambda col: f"={col}{I.rows['tax']}/{col}{I.rows['pbt']}")

    # ================= Balance sheet =================
    sheet_header(ws_bs, "Balance Sheet", "Balances at 31 May. History from the 10-K (intangibles and goodwill combined; notes payable in short-term debt). Cash is the balancing figure from the cash flow.")
    B = Sheet(ws_bs)
    B.sec("Assets")
    B.line("cash", "Cash and equivalents", hist=H["cash"], proj=lambda col: f"='Cash Flow'!{col}{{close}}")
    B.line("sti", "Short-term investments", hist=H["sti"], proj=lambda col: f"={prev(col)}{B.rows['sti']}")
    B.line("ar", "Accounts receivable, net", hist=H["ar"], proj=lambda col: f"='Income Statement'!{col}{I.rows['rev']}*{a('dso', col)}/365")
    B.line("inv", "Inventories", hist=H["inv"], proj=lambda col: f"='Income Statement'!{col}{I.rows['cogs']}*{a('dio', col)}/365")
    B.line("prep", "Prepaid expenses and other current assets", hist=H["prep"], proj=lambda col: f"='Income Statement'!{col}{I.rows['rev']}*{a('prep', col)}")
    B.line("ca", "Total current assets", hist_formula=lambda col, i: f"=SUM({col}{B.rows['cash']}:{col}{B.rows['prep']})",
           proj=lambda col: f"=SUM({col}{B.rows['cash']}:{col}{B.rows['prep']})", bold=True, border=TOP)
    B.line("ppe", "Property, plant and equipment, net", hist=H["ppe"],
           proj=lambda col: f"={prev(col)}{B.rows['ppe']}-'Cash Flow'!{col}{{capex}}-'Cash Flow'!{col}{{da}}")
    B.line("rou", "Operating lease right-of-use assets", hist=H["rou"], proj=lambda col: f"={prev(col)}{B.rows['rou']}")
    B.line("intang", "Intangible assets and goodwill", hist=H["intang"], proj=lambda col: f"={prev(col)}{B.rows['intang']}")
    B.line("oa", "Deferred income taxes and other assets", hist=H["oa"], proj=lambda col: f"={prev(col)}{B.rows['oa']}")
    B.line("ta", "Total assets", hist_formula=lambda col, i: f"={col}{B.rows['ca']}+SUM({col}{B.rows['ppe']}:{col}{B.rows['oa']})",
           proj=lambda col: f"={col}{B.rows['ca']}+SUM({col}{B.rows['ppe']}:{col}{B.rows['oa']})", bold=True, border=TOTAL)
    B.gap()
    B.sec("Liabilities")
    B.line("cdebt", "Short-term debt and current portion of long-term debt", hist=H["cdebt"],
           proj=lambda col: f"=MAX({prev(col)}{B.rows['cdebt']}-{a('repay', col)},0)")
    B.line("ap", "Accounts payable", hist=H["ap"], proj=lambda col: f"='Income Statement'!{col}{I.rows['cogs']}*{a('dpo', col)}/365")
    B.line("accr", "Accrued liabilities", hist=H["accr"], proj=lambda col: f"='Income Statement'!{col}{I.rows['rev']}*{a('accr', col)}")
    B.line("itp", "Income taxes payable", hist=H["itp"], proj=lambda col: f"={prev(col)}{B.rows['itp']}")
    B.line("cll", "Current operating lease liabilities", hist=H["cll"], proj=lambda col: f"={prev(col)}{B.rows['cll']}")
    B.line("cl", "Total current liabilities", hist_formula=lambda col, i: f"=SUM({col}{B.rows['cdebt']}:{col}{B.rows['cll']})",
           proj=lambda col: f"=SUM({col}{B.rows['cdebt']}:{col}{B.rows['cll']})", bold=True, border=TOP)
    B.line("ltd", "Long-term debt", hist=H["ltd"],
           proj=lambda col: f"={prev(col)}{B.rows['ltd']}-MAX({a('repay', col)}-{prev(col)}{B.rows['cdebt']},0)")
    B.line("ll", "Operating lease liabilities", hist=H["ll"], proj=lambda col: f"={prev(col)}{B.rows['ll']}")
    B.line("ol", "Deferred income taxes and other liabilities", hist=H["ol"], proj=lambda col: f"={prev(col)}{B.rows['ol']}")
    B.line("tl", "Total liabilities", hist_formula=lambda col, i: f"={col}{B.rows['cl']}+SUM({col}{B.rows['ltd']}:{col}{B.rows['ol']})",
           proj=lambda col: f"={col}{B.rows['cl']}+SUM({col}{B.rows['ltd']}:{col}{B.rows['ol']})", bold=True, border=TOP)
    B.gap()
    B.sec("Shareholders' equity")
    B.line("capital", "Common stock and capital in excess of stated value", hist=H["capital"],
           proj=lambda col: f"={prev(col)}{B.rows['capital']}+'Cash Flow'!{col}{{sbc}}")
    B.line("aoci", "Accumulated other comprehensive income (loss)", hist=H["aoci"], proj=lambda col: f"={prev(col)}{B.rows['aoci']}")
    B.line("re", "Retained earnings (deficit)", hist=H["re"],
           proj=lambda col: f"={prev(col)}{B.rows['re']}+'Income Statement'!{col}{I.rows['ni']}+'Cash Flow'!{col}{{div}}+'Cash Flow'!{col}{{bb}}")
    B.line("eq", "Total shareholders' equity", hist_formula=lambda col, i: f"=SUM({col}{B.rows['capital']}:{col}{B.rows['re']})",
           proj=lambda col: f"=SUM({col}{B.rows['capital']}:{col}{B.rows['re']})", bold=True, border=TOP)
    B.line("tle", "Total liabilities and shareholders' equity", hist_formula=lambda col, i: f"={col}{B.rows['tl']}+{col}{B.rows['eq']}",
           proj=lambda col: f"={col}{B.rows['tl']}+{col}{B.rows['eq']}", bold=True, border=TOTAL)
    B.gap()
    B.line("chk", "Check: assets - liabilities - equity (should be 0)", fmt=M,
           hist_formula=lambda col, i: f"=ROUND({col}{B.rows['ta']}-{col}{B.rows['tle']},0)",
           proj=lambda col: f"=ROUND({col}{B.rows['ta']}-{col}{B.rows['tle']},0)")

    # ================= Cash flow =================
    sheet_header(ws_cf, "Cash Flow Statement", "History from the 10-K (smaller items grouped). Forecast built from the income statement and balance sheet movements.")
    C = Sheet(ws_cf)
    C.sec("Operating activities")
    C.line("ni", "Net income", hist_formula=lambda col, i: f"='Income Statement'!{col}{I.rows['ni']}",
           proj=lambda col: f"='Income Statement'!{col}{I.rows['ni']}")
    C.line("da", "Depreciation and amortization", hist=H["da"], proj=lambda col: f"='Income Statement'!{col}{I.rows['rev']}*{a('da', col)}")
    C.line("sbc", "Stock-based compensation", hist=H["sbc"], proj=lambda col: f"='Income Statement'!{col}{I.rows['rev']}*{a('sbc', col)}")
    C.line("d_ar", "(Increase) decrease in receivables", proj=lambda col: f"='Balance Sheet'!{prev(col)}{B.rows['ar']}-'Balance Sheet'!{col}{B.rows['ar']}")
    C.line("d_inv", "(Increase) decrease in inventories", proj=lambda col: f"='Balance Sheet'!{prev(col)}{B.rows['inv']}-'Balance Sheet'!{col}{B.rows['inv']}")
    C.line("d_prep", "(Increase) decrease in prepaid and other", proj=lambda col: f"='Balance Sheet'!{prev(col)}{B.rows['prep']}-'Balance Sheet'!{col}{B.rows['prep']}")
    C.line("d_ap", "Increase (decrease) in payables", proj=lambda col: f"='Balance Sheet'!{col}{B.rows['ap']}-'Balance Sheet'!{prev(col)}{B.rows['ap']}")
    C.line("d_accr", "Increase (decrease) in accrued liabilities", proj=lambda col: f"='Balance Sheet'!{col}{B.rows['accr']}-'Balance Sheet'!{prev(col)}{B.rows['accr']}")
    C.line("wc_other", "Working capital and other items (history, total)", proj=lambda col: 0,
           hist_formula=lambda col, i: f"={H['cfo'][i]}-{col}{C.rows['ni']}-{col}{C.rows['da']}-{col}{C.rows['sbc']}")
    C.line("cfo", "Cash provided by operations", hist_formula=lambda col, i: f"=SUM({col}{C.rows['ni']}:{col}{C.rows['wc_other']})",
           proj=lambda col: f"=SUM({col}{C.rows['ni']}:{col}{C.rows['wc_other']})", bold=True, border=TOP)
    C.gap()
    C.sec("Investing activities")
    C.line("capex", "Additions to property, plant and equipment", hist=H["capex"],
           proj=lambda col: f"=-'Income Statement'!{col}{I.rows['rev']}*{a('capex', col)}")
    C.line("oinv", "Short-term investments and other", hist_formula=lambda col, i: f"={H['cfi'][i]}-{col}{C.rows['capex']}",
           proj=lambda col: 0)
    C.line("cfi", "Cash used by investing activities", hist_formula=lambda col, i: f"={col}{C.rows['capex']}+{col}{C.rows['oinv']}",
           proj=lambda col: f"={col}{C.rows['capex']}+{col}{C.rows['oinv']}", bold=True, border=TOP)
    C.gap()
    C.sec("Financing activities")
    C.line("repay", "Debt repaid", hist=H["repay"], proj=lambda col: f"=-{a('repay', col)}")
    C.line("div", "Dividends paid", hist=H["div"], proj=lambda col: f"=-'Income Statement'!{col}{I.rows['dps']}*'Income Statement'!{col}{I.rows['shares']}")
    C.line("bb", "Share buybacks", hist=H["bb"], proj=lambda col: f"=-{a('bb', col)}")
    C.line("ofin", "Share issuance and other", hist_formula=lambda col, i: f"={H['cff'][i]}-{col}{C.rows['repay']}-{col}{C.rows['div']}-{col}{C.rows['bb']}",
           proj=lambda col: 0)
    C.line("cff", "Cash used by financing activities", hist_formula=lambda col, i: f"=SUM({col}{C.rows['repay']}:{col}{C.rows['ofin']})",
           proj=lambda col: f"=SUM({col}{C.rows['repay']}:{col}{C.rows['ofin']})", bold=True, border=TOP)
    C.gap()
    C.line("fx", "Effect of exchange rates", hist=H["fx"], proj=lambda col: 0)
    C.line("net", "Net change in cash", hist_formula=lambda col, i: f"={col}{C.rows['cfo']}+{col}{C.rows['cfi']}+{col}{C.rows['cff']}+{col}{C.rows['fx']}",
           proj=lambda col: f"={col}{C.rows['cfo']}+{col}{C.rows['cfi']}+{col}{C.rows['cff']}+{col}{C.rows['fx']}", bold=True, border=TOP)
    C.line("open", "Cash at beginning of year", hist=H["cash_open"], proj=lambda col: f"={prev(col)}{C.rows['close']}" if False else "")
    C.line("close", "Cash at end of year", hist_formula=lambda col, i: f"={col}{C.rows['open']}+{col}{C.rows['net']}",
           proj=lambda col: f"={col}{C.rows['open']}+{col}{C.rows['net']}", bold=True, border=TOTAL)
    for col in PROJ:
        ws_cf[f"{col}{C.rows['open']}"] = f"={prev(col)}{C.rows['close']}"
        ws_cf[f"{col}{C.rows['open']}"].font = F
    C.gap()
    C.sec("Free cash flow")
    C.line("fcf", "Free cash flow (operations less capex)", hist_formula=lambda col, i: f"={col}{C.rows['cfo']}+{col}{C.rows['capex']}",
           proj=lambda col: f"={col}{C.rows['cfo']}+{col}{C.rows['capex']}", bold=True)
    C.line("ret", "Cash returned to shareholders", hist_formula=lambda col, i: f"=-{col}{C.rows['div']}-{col}{C.rows['bb']}",
           proj=lambda col: f"=-{col}{C.rows['div']}-{col}{C.rows['bb']}")

    # resolve {placeholders} that point to rows on later sheets
    lookup = {"cdebt": B.rows["cdebt"], "ltd": B.rows["ltd"], "cash": B.rows["cash"], "sti": B.rows["sti"],
              "close": C.rows["close"], "capex": C.rows["capex"], "da": C.rows["da"], "sbc": C.rows["sbc"],
              "div": C.rows["div"], "bb": C.rows["bb"]}
    for ws in (ws_is, ws_bs):
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and "{" in c.value:
                    v = c.value
                    for k, rr in lookup.items():
                        v = v.replace("{" + k + "}", str(rr))
                    c.value = v

    # ================= Scenarios (all three side by side) =================
    sheet_header(ws_sc, "Scenario Comparison", "Revenue and operating income under each scenario, calculated in parallel (independent of the scenario selector).")
    S = Sheet(ws_sc)
    sc_rows = {}
    for si, sname in enumerate(["Base case", "Upside case", "Downside case"]):
        S.sec(sname)
        seg_rows = []
        for s in SEG:
            rr = S.line((sname, s), s, proj=lambda col: "")
            ws_sc[f"C{rr}"], ws_sc[f"D{rr}"] = f"='Revenue Build'!C{R.rows[('seg', s)]}", f"='Revenue Build'!D{R.rows[('seg', s)]}"
            ws_sc[f"C{rr}"].font = ws_sc[f"D{rr}"].font = F_LINK
            for col in PROJ:
                ws_sc[f"{col}{rr}"] = f"={prev(col)}{rr}*(1+Assumptions!{col}{tables[sname][s]})"
                ws_sc[f"{col}{rr}"].font = F
            seg_rows.append(rr)
        rr = S.line((sname, "rev"), "Total revenue", bold=True, border=TOP,
                    proj=lambda col: f"=SUM({col}{seg_rows[0]}:{col}{seg_rows[-1]})+{a('other', col)}")
        ws_sc[f"D{rr}"] = f"='Revenue Build'!D{R.rows['rev']}"
        rev_r = rr
        ebit_r = S.line((sname, "ebit"), "Operating income (EBIT)", bold=True,
                        proj=lambda col: f"={col}{rev_r}*(Assumptions!{col}{gm_rows[si]}-{a('dc', col)}-{a('oh', col)})")
        ws_sc[f"D{ebit_r}"] = f"='Income Statement'!D{I.rows['ebit']}"
        m_r = S.line((sname, "m"), "EBIT margin", fmt=PCT, proj=lambda col: f"={col}{ebit_r}/{col}{rev_r}")
        ws_sc[f"D{m_r}"] = f"=D{ebit_r}/D{rev_r}"
        sc_rows[sname] = (rev_r, ebit_r, m_r)
        S.gap()
    S.sec("Summary: FY2031")
    head = S.r
    for j, t in enumerate(["Scenario", "FY2031 revenue", "Revenue CAGR FY26-31", "FY2031 EBIT", "FY2031 EBIT margin"]):
        c = ws_sc.cell(row=head, column=1 + j, value=t)
        c.font = F_B
        c.alignment = Alignment(wrap_text=True, horizontal="left" if j == 0 else "right", vertical="bottom")
    ws_sc.row_dimensions[head].height = 30
    for k, sname in enumerate(["Base case", "Upside case", "Downside case"]):
        r = head + 1 + k
        rev_r, ebit_r, m_r = sc_rows[sname]
        ws_sc.cell(row=r, column=1, value=sname).font = F
        ws_sc[f"B{r}"] = f"=I{rev_r}"
        ws_sc[f"C{r}"] = f"=(I{rev_r}/D{rev_r})^(1/5)-1"
        ws_sc[f"D{r}"] = f"=I{ebit_r}"
        ws_sc[f"E{r}"] = f"=I{m_r}"
        for col, fmt in zip("BCDE", [M, PCT, M, PCT]):
            ws_sc[f"{col}{r}"].number_format = fmt
    # chart: revenue by scenario
    ws_sc[f"K{head}"] = "Chart data"
    ws_sc[f"K{head}"].font = F_NOTE
    lc = LineChart()
    for sname, color in (("Upside case", "2A78D6"), ("Base case", "1F4E79"), ("Downside case", "EB6834")):
        rev_r = sc_rows[sname][0]
        ws_sc[f"K{rev_r}"] = sname
        data_ref = Reference(ws_sc, min_col=4, max_col=9, min_row=rev_r)
        lc.add_data(data_ref, from_rows=True, titles_from_data=False)
        s = lc.series[-1]
        from openpyxl.chart.series import SeriesLabel
        s.tx = SeriesLabel(v=sname)
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.width = 28575
        s.smooth = False
    hdr_ref = Reference(ws_sc, min_col=4, max_col=9, min_row=4)
    lc.set_categories(hdr_ref)
    lc.title, lc.y_axis.numFmt = "Revenue by scenario ($m)", "#,##0"
    lc.y_axis.scaling.min = 35000
    lc.legend.position = "b"
    lc.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E1E0D9"))
    lc.x_axis.delete = lc.y_axis.delete = False
    lc.height, lc.width = 8, 16
    ws_sc.add_chart(lc, f"A{head + 6}")

    # ================= Checks =================
    sheet_header(ws_ck, "Checks", "All should be 0 (or OK).")
    K = Sheet(ws_ck)
    K.line("c1", "Balance sheet balances", fmt=M, hist_formula=lambda col, i: f"='Balance Sheet'!{col}{B.rows['chk']}",
           proj=lambda col: f"='Balance Sheet'!{col}{B.rows['chk']}")
    K.line("c2", "Cash flow closing cash = balance sheet cash", fmt=M,
           hist_formula=lambda col, i: f"=ROUND('Cash Flow'!{col}{C.rows['close']}-'Balance Sheet'!{col}{B.rows['cash']},0)",
           proj=lambda col: f"=ROUND('Cash Flow'!{col}{C.rows['close']}-'Balance Sheet'!{col}{B.rows['cash']},0)")
    K.line("c3", "Revenue build = income statement revenue", fmt=M,
           hist_formula=lambda col, i: "=0" if i == 0 else f"=ROUND('Revenue Build'!{col}{R.rows['rev']}-'Income Statement'!{col}{I.rows['rev']},0)",
           proj=lambda col: f"=ROUND('Revenue Build'!{col}{R.rows['rev']}-'Income Statement'!{col}{I.rows['rev']},0)")
    K.line("c4", "Reported net income agrees (history)", fmt=M,
           hist_formula=lambda col, i: f"=ROUND('Income Statement'!{col}{I.rows['ni']}-{[5700, 3219, 3108][i]},0)", proj=lambda col: 0)
    K.line("c5", "Reported total assets agree (history)", fmt=M,
           hist_formula=lambda col, i: f"=ROUND('Balance Sheet'!{col}{B.rows['ta']}-{[38110, 36579, 38410][i]},0)", proj=lambda col: 0)
    K.line("c6", "Base case scenario sheet = model (when scenario 1 selected)", fmt=M,
           hist_formula=lambda col, i: 0,
           proj=lambda col: f"=IF({SCN}=1,ROUND(Scenarios!{col}{sc_rows['Base case'][0]}-'Income Statement'!{col}{I.rows['rev']},0),0)")
    K.line("c7", "Cash stays positive", fmt='General', hist_formula=lambda col, i: '="OK"',
           proj=lambda col: f"=IF('Balance Sheet'!{col}{B.rows['cash']}>0,\"OK\",\"LOW CASH\")")
    sr = K.r + 1
    ws_ck.cell(row=sr, column=1, value="Overall status").font = F_B
    ws_ck[f"B{sr}"] = (f'=IF(AND(SUMPRODUCT(ABS(B{K.rows["c1"]}:I{K.rows["c6"]}))=0,'
                       f'COUNTIF(B{K.rows["c7"]}:I{K.rows["c7"]},"OK")=8),"OK","ERROR")')
    ws_ck[f"B{sr}"].font, ws_ck[f"B{sr}"].fill = F_B, FILL_KEY
    status = f"Checks!B{sr}"

    # ================= Summary =================
    sheet_header(ws_sum, "Nike Five-Year Forecast: Summary", "Selected scenario shown. Change the scenario on the Assumptions sheet (cell B6).")
    ws_sum["A3"] = "=\"Scenario: \"&Assumptions!C6"
    ws_sum["A3"].font = F_B
    U = Sheet(ws_sum)
    U.line("rev", "Revenue", fmt=M, hist_formula=lambda col, i: f"='Income Statement'!{col}{I.rows['rev']}",
           proj=lambda col: f"='Income Statement'!{col}{I.rows['rev']}")
    U.line("g", "Revenue growth", fmt=PCT, hist_formula=lambda col, i: None if i == 0 else f"='Revenue Build'!{col}{R.rows['grt']}",
           proj=lambda col: f"='Revenue Build'!{col}{R.rows['grt']}")
    U.line("gm", "Gross margin", fmt=PCT, hist_formula=lambda col, i: f"='Income Statement'!{col}{I.rows['gm_p']}",
           proj=lambda col: f"='Income Statement'!{col}{I.rows['gm_p']}")
    U.line("ebit", "Operating income (EBIT)", fmt=M, hist_formula=lambda col, i: f"='Income Statement'!{col}{I.rows['ebit']}",
           proj=lambda col: f"='Income Statement'!{col}{I.rows['ebit']}")
    U.line("em", "EBIT margin", fmt=PCT, hist_formula=lambda col, i: f"='Income Statement'!{col}{I.rows['ebit_p']}",
           proj=lambda col: f"='Income Statement'!{col}{I.rows['ebit_p']}")
    U.line("ni", "Net income", fmt=M, hist_formula=lambda col, i: f"='Income Statement'!{col}{I.rows['ni']}",
           proj=lambda col: f"='Income Statement'!{col}{I.rows['ni']}")
    U.line("eps", "Diluted EPS ($)", fmt=EPS, hist_formula=lambda col, i: f"='Income Statement'!{col}{I.rows['eps']}",
           proj=lambda col: f"='Income Statement'!{col}{I.rows['eps']}")
    U.line("cfo", "Cash from operations", fmt=M, hist_formula=lambda col, i: f"='Cash Flow'!{col}{C.rows['cfo']}",
           proj=lambda col: f"='Cash Flow'!{col}{C.rows['cfo']}")
    U.line("fcf", "Free cash flow", fmt=M, hist_formula=lambda col, i: f"='Cash Flow'!{col}{C.rows['fcf']}",
           proj=lambda col: f"='Cash Flow'!{col}{C.rows['fcf']}")
    U.line("ret", "Cash returned to shareholders", fmt=M, hist_formula=lambda col, i: f"='Cash Flow'!{col}{C.rows['ret']}",
           proj=lambda col: f"='Cash Flow'!{col}{C.rows['ret']}")
    U.line("cash", "Cash and short-term investments", fmt=M,
           hist_formula=lambda col, i: f"='Balance Sheet'!{col}{B.rows['cash']}+'Balance Sheet'!{col}{B.rows['sti']}",
           proj=lambda col: f"='Balance Sheet'!{col}{B.rows['cash']}+'Balance Sheet'!{col}{B.rows['sti']}")
    U.line("debt", "Total debt", fmt=M,
           hist_formula=lambda col, i: f"='Balance Sheet'!{col}{B.rows['cdebt']}+'Balance Sheet'!{col}{B.rows['ltd']}",
           proj=lambda col: f"='Balance Sheet'!{col}{B.rows['cdebt']}+'Balance Sheet'!{col}{B.rows['ltd']}")
    for row in ws_sum.iter_rows(min_row=5, max_row=U.r, min_col=2, max_col=9):
        for c in row:
            c.font = F_LINK
    ws_sum.cell(row=U.r + 1, column=1, value="Checks status").font = F_B
    ws_sum.cell(row=U.r + 1, column=2, value=f"={status}").font = F_B
    # charts
    bc = BarChart()
    bc.add_data(Reference(ws_sum, min_col=1, max_col=9, min_row=U.rows["rev"]), from_rows=True, titles_from_data=True)
    bc.series[0].graphicalProperties.solidFill = "2A78D6"
    bc.series[0].graphicalProperties.line.solidFill = "2A78D6"
    bc.set_categories(Reference(ws_sum, min_col=2, max_col=9, min_row=4))
    bc.title, bc.y_axis.numFmt, bc.legend = "Revenue ($m): actual FY2024-26, forecast FY2027-31", "#,##0", None
    lc2 = LineChart()
    lc2.add_data(Reference(ws_sum, min_col=1, max_col=9, min_row=U.rows["em"]), from_rows=True, titles_from_data=True)
    lc2.series[0].graphicalProperties.line.solidFill = "EB6834"
    lc2.series[0].graphicalProperties.line.width = 28575
    lc2.set_categories(Reference(ws_sum, min_col=2, max_col=9, min_row=4))
    lc2.title, lc2.y_axis.numFmt, lc2.legend = "EBIT margin", "0%", None
    lc2.y_axis.scaling.min = 0
    lc2.series[0].smooth = False
    bc.y_axis.scaling.min = 0
    for ch, anchor in ((bc, f"A{U.r + 4}"), (lc2, f"F{U.r + 4}")):
        ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E1E0D9"))
        ch.x_axis.delete = ch.y_axis.delete = False
        ch.height, ch.width = 7.5, 13
        ws_sum.add_chart(ch, anchor)

    # ================= Cover =================
    cover["A1"] = "NIKE, Inc.: Five-Year Financial Forecast"
    cover["A1"].font = Font(name=AR, size=18, bold=True, color="1F4E79")
    cover["A2"] = "Three-statement model: FY2024 to FY2026 actual, FY2027 to FY2031 forecast (years ending 31 May). $ millions."
    cover["A2"].font = F_NOTE
    rows = [
        ("Purpose", "Forecast Nike's revenue, profit and cash flow for five years under base, upside and downside scenarios."),
        ("Selected scenario", "=Assumptions!C6"),
        ("Revenue", "Built up by region (North America, EMEA, Greater China, APLA, Converse), each with its own growth rate per scenario."),
        ("Costs", "Gross margin by scenario; marketing (demand creation) and overhead as a % of revenue."),
        ("Statements", "Linked income statement, balance sheet and cash flow. Cash is the balancing item; the balance sheet balances every year."),
        ("", ""),
        ("Sheets", ""),
        ("Summary", "Key outputs and charts for the selected scenario"),
        ("Assumptions", "Every driver, with base / upside / downside tables. Change cell B6 to switch scenario."),
        ("Revenue Build", "Revenue by region, growth and mix"),
        ("Income Statement / Balance Sheet / Cash Flow", "The three linked statements"),
        ("Scenarios", "All three scenarios side by side"),
        ("Checks", "Balance sheet, cash and source checks"),
        ("", ""),
        ("Sources", "NIKE, Inc. Form 10-K filings (FY2024 to FY2026) and fiscal 2026 fourth quarter results release (revenue by region)."),
        ("Note", "The FY2027 base case is set to match Nike's guidance of 1 Oct 2026 (revenue down a high single digit percentage, EPS of about $1.15 to $1.35 before restructuring). The regional split and FY2028 to FY2031 are my own assumptions."),
        ("Colour key", "Blue = input, black = formula, green = link to another sheet, yellow = key assumption, grey = actual (history)"),
        ("Checks status", f"={status}"),
    ]
    for i, (k, v) in enumerate(rows):
        cover.cell(row=4 + i, column=1, value=k).font = F_B
        c = cover.cell(row=4 + i, column=2, value=v)
        c.font = F
        if k in ("Checks status", "Selected scenario"):
            c.font, c.fill = F_B, FILL_KEY
    cover.column_dimensions["A"].width = 44
    cover.column_dimensions["B"].width = 110
    cover.sheet_view.showGridLines = False

    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(OUT)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    build()
