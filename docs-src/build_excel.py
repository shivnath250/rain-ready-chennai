import csv
import json
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment

import rain_ready_chennai as M

HERE = os.path.dirname(os.path.abspath(__file__))
R = json.load(open(os.path.join(HERE, "outputs", "results.json")))

wb = Workbook()
F = "Arial"
thin = Side(style="thin", color="000000")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
B = Font(name=F, bold=True)
N = Font(name=F)
T = Font(name=F, bold=True, size=13)
I = Font(name=F, italic=True)


def put(ws, ref, v, font=N, fmt=None, box=False, wrap=False):
    c = ws[ref]
    c.value = v
    c.font = font
    if fmt:
        c.number_format = fmt
    if box:
        c.border = BOX
    if wrap:
        c.alignment = Alignment(wrap_text=True, vertical="top")
    return c


def header(ws, row, col, labels):
    for k, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + k, value=lab)
        c.font = B
        c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")


def cell(ws, row, col, v, fmt=None, font=N):
    c = ws.cell(row=row, column=col, value=v)
    c.font = font
    c.border = BOX
    if fmt:
        c.number_format = fmt
    return c


def widths(ws, d):
    for k, v in d.items():
        ws.column_dimensions[k].width = v


def solver_names(ws, sheet_idx, objective, typ, adj, cons):
    """Store a Solver model in the sheet so Excel Solver opens pre-loaded."""
    names = {
        "solver_opt": objective, "solver_typ": str(typ), "solver_adj": adj,
        "solver_num": str(len(cons)), "solver_eng": "2", "solver_lin": "1",
        "solver_neg": "1", "solver_ver": "3", "solver_val": "0",
        "solver_cvg": "0.0001", "solver_mip": "2147483647", "solver_nod": "2147483647",
        "solver_itr": "2147483647", "solver_tim": "2147483647", "solver_tol": "0",
        "solver_pre": "0.000001", "solver_scl": "1", "solver_sho": "2", "solver_rlx": "2",
        "solver_est": "1", "solver_ssz": "100", "solver_rsd": "0", "solver_rbv": "1",
        "solver_drv": "1", "solver_msl": "2", "solver_nwt": "1", "solver_mni": "30",
        "solver_mrt": "0.075", "solver_hss": "1", "solver_seed": "0",
    }
    q = f"'{ws.title}'!"
    for k, (lhs, rel, rhs) in enumerate(cons, start=1):
        names[f"solver_lhs{k}"] = q + lhs
        names[f"solver_rel{k}"] = str(rel)
        names[f"solver_rhs{k}"] = rhs if not rhs.startswith("$") else q + rhs
    names["solver_opt"] = q + objective
    names["solver_adj"] = ",".join(q + p for p in adj.split(","))
    for k, v in names.items():
        dn = DefinedName(k, attr_text=v, localSheetId=sheet_idx, hidden=True)
        ws.defined_names[k] = dn


# ============================================================================
# README
# ============================================================================
ws = wb.active
ws.title = "README"
widths(ws, {"A": 26, "B": 100})
put(ws, "A1", "Rain-Ready Chennai: Markov-informed relief logistics and evacuation planner", T)
put(ws, "A2", "Decision Science, Creative AI project | Shivnath Rajendra Gujar (FT271017) | GLIM PGPM 2026-27", I)
rows = [
    ("Problem", "Every North-East monsoon (Oct to Dec), low-lying Chennai localities such as Velachery, "
                "Pallikaranai and Mudichur flood. Relief kits must flow from warehouses through hubs to camps "
                "on roads that may themselves be under water. The planner must decide, a day ahead and under "
                "uncertain rain, how many kits to pre-position where, how to route them, and which evacuation "
                "road to use."),
    ("Concepts used", "Markov chain (rain regimes), transshipment LP (relief routing), two-stage stochastic LP "
                      "(pre-positioning, VSS and EVPI), network models (most-reliable path, max-flow / min-cut)."),
    ("Data", "Rainfall series is SYNTHETIC (5 seasons x 92 days, generated in Python). Network costs, capacities "
             "and flood probabilities are ILLUSTRATIVE. Rain categories follow IMD daily rainfall classes. "
             "Paste real IMD daily rainfall into Rain_Data column C and edit the Inputs sheet to reuse the model."),
    ("Sheet: Inputs", "All parameters: supplies, hub limits, camp demand, regime effects, flood probabilities, "
                      "penalty, and the arc list. Edit these."),
    ("Sheet: Rain_Data", "Daily rainfall (column C, input). Columns D and E classify each day and link it to the next day."),
    ("Sheet: Markov", "Counts transitions, estimates the transition matrix, forecasts the regime N days ahead, "
                      "gives the steady state, the mean days until dry weather returns, and spell lengths."),
    ("Sheet: Transship_LP", "Solver-ready transshipment LP for the regime chosen in cell B3. "
                            "Open Data > Solver (model is pre-loaded) and click Solve. Re-solve after changing B3."),
    ("Sheet: Playbook", "Optimal routing for all four regimes, from the Python run. The 'regime playbook'."),
    ("Sheet: Prepositioning", "Two-stage stochastic results from Python: what to stock today, VSS, EVPI, policy by "
                              "today's regime, and penalty sensitivity."),
    ("Sheet: Evacuation", "Solver-ready most-reliable path LP from Mudichur to a shelter, plus max-flow / min-cut "
                          "and pump-placement results from Python."),
    ("How to read cells", "Bold headers. Cells marked 'INPUT' may be edited. Cells marked 'DECISION' are changed by "
                          "Solver. Everything else is a formula. Sheets Playbook and Prepositioning hold values pasted "
                          "from rain_ready_chennai.py (re-run the script to refresh them)."),
    ("Solver note", "If Solver does not appear: File > Options > Add-ins > Excel Add-ins > Go > tick Solver Add-in. "
                    "Use 'Simplex LP', tick 'Make unconstrained variables non-negative'. If the pre-loaded model is "
                    "missing, the constraints are written out on each LP sheet."),
]
for k, (a, b) in enumerate(rows, start=4):
    put(ws, f"A{k}", a, B, box=True, wrap=True)
    put(ws, f"B{k}", b, box=True, wrap=True)
    ws.row_dimensions[k].height = 48

# ============================================================================
# INPUTS
# ============================================================================
wi = wb.create_sheet("Inputs")
widths(wi, {"A": 18, "B": 16, "C": 32, "D": 14, "E": 14, "F": 14, "G": 40})
put(wi, "A1", "INPUTS - edit any value in this sheet", T)
put(wi, "A2", "All values are illustrative assumptions for teaching (not surveyed figures).", I)

put(wi, "A4", "Sources", B)
header(wi, 5, 1, ["Source", "Kits per day (INPUT)"])
for k, (s, v) in enumerate(M.SOURCES.items()):
    cell(wi, 6 + k, 1, s); cell(wi, 6 + k, 2, v, "#,##0")

put(wi, "A10", "Hubs", B)
header(wi, 11, 1, ["Hub", "Max kits per day (INPUT)"])
for k, (s, v) in enumerate(M.HUBS.items()):
    cell(wi, 12 + k, 1, s); cell(wi, 12 + k, 2, v, "#,##0")

put(wi, "A16", "Relief camps", B)
header(wi, 17, 1, ["Camp", "Kits per day at peak, regime V (INPUT)"])
for k, (s, v) in enumerate(M.CAMPS_PEAK.items()):
    cell(wi, 18 + k, 1, s); cell(wi, 18 + k, 2, v, "#,##0")
cell(wi, 24, 1, "Total"); cell(wi, 24, 2, "=SUM(B18:B23)", "#,##0", B)
cell(wi, 25, 1, "Camp storage"); cell(wi, 25, 2, M.CAMP_STORAGE, "#,##0")
put(wi, "C25", "kits a camp can hold if pre-positioned")

put(wi, "A27", "Rain regimes (IMD daily rainfall classes)", B)
header(wi, 28, 1, ["Regime", "Share of peak demand (INPUT)", "Description", "Cost multiplier (INPUT)"])
for k, s in enumerate(M.STATES):
    cell(wi, 29 + k, 1, s); cell(wi, 29 + k, 2, M.DEMAND_SHARE[s], "0%")
    cell(wi, 29 + k, 3, M.STATE_NAMES[s]); cell(wi, 29 + k, 4, M.COST_MULT[s], "0.00")

put(wi, "A34", "Probability a road is impassable (INPUT)", B)
header(wi, 35, 1, ["Vulnerability"] + M.STATES)
for k, v in enumerate(["low", "med", "high"]):
    cell(wi, 36 + k, 1, v)
    for j, s in enumerate(M.STATES):
        cell(wi, 36 + k, 2 + j, M.FLOOD_PROB[v][s], "0%")

put(wi, "A40", "Costs", B)
header(wi, 41, 1, ["Parameter", "Rs per kit (INPUT)", "Meaning"])
for k, (a, v, m) in enumerate([
        ("Penalty", M.PENALTY, "humanitarian cost of one kit of unmet demand"),
        ("Holding", M.HOLDING, "loading and guarding a pre-positioned kit"),
        ("Waste", M.WASTE, "pre-positioned kit left unused at a camp")]):
    cell(wi, 42 + k, 1, a); cell(wi, 42 + k, 2, v, "#,##0"); cell(wi, 42 + k, 3, m)

put(wi, "A46", "Road arcs of the relief network", B)
header(wi, 47, 1, ["From", "To", "Base cost Rs/kit (INPUT)", "Base capacity kits/day (INPUT)",
                   "Vulnerability (INPUT)"])
for k, (f, t, c, cap, v) in enumerate(M.ARCS):
    r = 48 + k
    cell(wi, r, 1, f); cell(wi, r, 2, t); cell(wi, r, 3, c, "#,##0")
    cell(wi, r, 4, cap, "#,##0"); cell(wi, r, 5, v)
NA = len(M.ARCS)
ARC_LAST = 47 + NA  # 66

# ============================================================================
# RAIN DATA
# ============================================================================
wr = wb.create_sheet("Rain_Data")
widths(wr, {"A": 10, "B": 8, "C": 16, "D": 10, "E": 14, "G": 70})
header(wr, 1, 1, ["Season", "Day", "Rain mm (INPUT)", "State", "Next-day state"])
rain = list(csv.reader(open(os.path.join(HERE, "outputs", "rain_data.csv"))))[1:]
nr = len(rain)
for k, (s, d, mm, _) in enumerate(rain):
    r = 2 + k
    wr.cell(row=r, column=1, value=int(s)).font = N
    wr.cell(row=r, column=2, value=int(d)).font = N
    c = wr.cell(row=r, column=3, value=float(mm)); c.font = N; c.number_format = "0.0"
    wr.cell(row=r, column=4,
            value=f'=IF(C{r}>=Markov!$B$8,"V",IF(C{r}>=Markov!$B$7,"H",IF(C{r}>=Markov!$B$6,"M","D")))').font = N
    nxt = f'=IF(A{r+1}=A{r},D{r+1},"")' if k < nr - 1 else '=""'
    wr.cell(row=r, column=5, value=nxt).font = N
LAST = 1 + nr
put(wr, "G1", "SYNTHETIC data generated in Python (seed 42) to demonstrate the template.", B)
put(wr, "G2", "To use real data: paste IMD daily rainfall (Chennai Nungambakkam, Oct-Dec) into columns A-C.", N)
put(wr, "G3", "Next-day state is blank at the end of each season so 31 Dec is not linked to 1 Oct.", N)
put(wr, "G4", f"Rows used: 2 to {LAST}. If you paste more rows, copy D and E down and extend the Markov ranges.", N)

# ============================================================================
# MARKOV
# ============================================================================
wm = wb.create_sheet("Markov")
widths(wm, {"A": 30, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 60})
put(wm, "A1", "Markov chain of daily rain regimes", T)
put(wm, "A3", "State thresholds", B)
header(wm, 4, 1, ["State", "Lower bound mm (INPUT)"])
for k, (s, v) in enumerate(zip(M.STATES, [0, 15.6, 64.5, 115.6])):
    cell(wm, 5 + k, 1, s); cell(wm, 5 + k, 2, v, "0.0")

put(wm, "A10", "Transition counts (from row state today to column state tomorrow)", B)
header(wm, 11, 1, ["From \\ To"] + M.STATES + ["Total"])
for i, s in enumerate(M.STATES):
    r = 12 + i
    cell(wm, r, 1, s, font=B)
    for j in range(4):
        col = "BCDE"[j]
        cell(wm, r, 2 + j,
             f"=COUNTIFS(Rain_Data!$D$2:$D${LAST},$A{r},Rain_Data!$E$2:$E${LAST},{col}$11)", "0")
    cell(wm, r, 6, f"=SUM(B{r}:E{r})", "0")

put(wm, "A17", "Estimated transition matrix P", B)
header(wm, 18, 1, ["From \\ To"] + M.STATES + ["Row sum"])
for i, s in enumerate(M.STATES):
    r = 19 + i
    cell(wm, r, 1, s, font=B)
    for j in range(4):
        col = "BCDE"[j]
        cell(wm, r, 2 + j, f"=IF($F{12+i}=0,0,{col}{12+i}/$F{12+i})", "0.000")
    cell(wm, r, 6, f"=SUM(B{r}:E{r})", "0.000")
put(wm, "G19", "P(i,j) = count(i to j) / count(i to anything)")

put(wm, "A25", "Today's regime (INPUT)", B)
c = cell(wm, 25, 2, "M")
put(wm, "C25", "type D, M, H or V")
put(wm, "A27", "N-day forecast  pi(n+1) = pi(n) x P", B)
header(wm, 28, 1, ["Days ahead"] + M.STATES + ["Sum"])
for n in range(31):
    r = 29 + n
    cell(wm, r, 1, n)
    for j in range(4):
        col = "BCDE"[j]
        if n == 0:
            f = f"=IF({col}$28=$B$25,1,0)"
        else:
            p = r - 1
            f = f"=$B{p}*{col}$19+$C{p}*{col}$20+$D{p}*{col}$21+$E{p}*{col}$22"
        cell(wm, r, 2 + j, f, "0.000")
    cell(wm, r, 6, f"=SUM(B{r}:E{r})", "0.000")
put(wm, "G30", "Row 'Days ahead = 1' is tomorrow's scenario probability (used in Prepositioning).")
put(wm, "G31", "The forecast settles by day 3 or 4: the monsoon 'forgets' today quickly.")
# rows 29..59
put(wm, "A61", "Steady state (long-run share of days)", B)
header(wm, 62, 1, [""] + M.STATES + ["Sum"])
cell(wm, 63, 1, "Approx. = day 30 row")
for j in range(4):
    col = "BCDE"[j]
    cell(wm, 63, 2 + j, f"={col}59", "0.000")
cell(wm, 63, 6, "=SUM(B63:E63)", "0.000")

put(wm, "A65", "Mean days until a Dry day, m = (I - Q)^-1 x 1", B)
header(wm, 66, 1, ["I - Q", "M", "H", "V", "", "Ones"])
wet = ["M", "H", "V"]
for i in range(3):
    r = 67 + i
    cell(wm, r, 1, wet[i], font=B)
    for j in range(3):
        src = f"{'CDE'[j]}{20+i}"
        cell(wm, r, 2 + j, f"={1 if i == j else 0}-{src}", "0.000")
    cell(wm, r, 6, 1)
header(wm, 71, 1, ["Starting regime", "Mean days to Dry"])
for i in range(3):
    r = 72 + i
    cell(wm, r, 1, wet[i], font=B)
    cell(wm, r, 2, f"=INDEX(MMULT(MINVERSE($B$67:$D$69),$F$67:$F$69),{i+1},1)", "0.00")
put(wm, "C72", "Q = rows/columns of P for the wet states M, H, V (Dry treated as the target state)")

put(wm, "A76", "Expected length of a spell, 1 / (1 - P(i,i))", B)
header(wm, 77, 1, ["Regime", "Days"])
for i, s in enumerate(M.STATES):
    cell(wm, 78 + i, 1, s, font=B)
    cell(wm, 78 + i, 2, f"=1/(1-{'BCDE'[i]}{19+i})", "0.00")

# ============================================================================
# TRANSSHIPMENT LP (Solver-ready)
# ============================================================================
wt = wb.create_sheet("Transship_LP")
widths(wt, {"A": 16, "B": 16, "C": 12, "D": 12, "E": 12, "F": 11, "G": 11, "H": 12, "I": 14,
            "J": 12, "K": 10, "M": 70})
put(wt, "A1", "Transshipment LP: route relief kits for one rain regime", T)
put(wt, "A3", "Regime (INPUT)", B)
cell(wt, 3, 2, "V")
put(wt, "C3", "D, M, H or V. Then run Solver.")
put(wt, "A4", "Penalty Rs/kit", B)
cell(wt, 4, 2, "=Inputs!$B$42", "#,##0")

header(wt, 6, 1, ["From", "To", "Vulnerability", "Base cost", "Base cap", "P(flooded)",
                  "Cost Rs/kit", "Capacity", "Flow (DECISION)", "Cost Rs", "Spare cap"])
flowsV = R["playbook"]["V"]["flows"]
for k in range(NA):
    r = 7 + k
    ir = 48 + k
    cell(wt, r, 1, f"=Inputs!A{ir}")
    cell(wt, r, 2, f"=Inputs!B{ir}")
    cell(wt, r, 3, f"=Inputs!E{ir}")
    cell(wt, r, 4, f"=Inputs!C{ir}", "#,##0")
    cell(wt, r, 5, f"=Inputs!D{ir}", "#,##0")
    cell(wt, r, 6, f"=INDEX(Inputs!$B$36:$E$38,MATCH(C{r},Inputs!$A$36:$A$38,0),MATCH($B$3,Inputs!$B$35:$E$35,0))", "0%")
    cell(wt, r, 7, f"=D{r}*INDEX(Inputs!$D$29:$D$32,MATCH($B$3,Inputs!$A$29:$A$32,0))", "#,##0.0")
    cell(wt, r, 8, f"=E{r}*(1-F{r})", "#,##0")
    f, t = M.ARCS[k][0], M.ARCS[k][1]
    cell(wt, r, 9, flowsV[f"{f} -> {t}"], "#,##0", B)
    cell(wt, r, 10, f"=G{r}*I{r}", "#,##0")
    cell(wt, r, 11, f"=H{r}-I{r}", "#,##0")
AL = 6 + NA  # 25

r0 = AL + 3  # 28
put(wt, f"A{r0}", "Sources: outflow <= supply", B)
header(wt, r0 + 1, 1, ["Source", "Outflow", "<=", "Supply"])
for k in range(3):
    r = r0 + 2 + k
    cell(wt, r, 1, f"=Inputs!A{6+k}")
    cell(wt, r, 2, f"=SUMIF($A$7:$A${AL},A{r},$I$7:$I${AL})", "#,##0")
    cell(wt, r, 3, "<=")
    cell(wt, r, 4, f"=Inputs!B{6+k}", "#,##0")
SRC = (r0 + 2, r0 + 4)

r1 = r0 + 6  # 34
put(wt, f"A{r1}", "Hubs: inflow = outflow, inflow <= throughput", B)
header(wt, r1 + 1, 1, ["Hub", "Inflow", "Outflow", "Net (= 0)", "Throughput"])
for k in range(3):
    r = r1 + 2 + k
    cell(wt, r, 1, f"=Inputs!A{12+k}")
    cell(wt, r, 2, f"=SUMIF($B$7:$B${AL},A{r},$I$7:$I${AL})", "#,##0")
    cell(wt, r, 3, f"=SUMIF($A$7:$A${AL},A{r},$I$7:$I${AL})", "#,##0")
    cell(wt, r, 4, f"=B{r}-C{r}", "#,##0")
    cell(wt, r, 5, f"=Inputs!B{12+k}", "#,##0")
HUB = (r1 + 2, r1 + 4)

r2 = r1 + 6  # 40
put(wt, f"A{r2}", "Camps: delivered + unmet = demand", B)
header(wt, r2 + 1, 1, ["Camp", "Delivered", "Unmet (DECISION)", "Delivered + unmet", "Demand"])
unV = R["playbook"]["V"]["unmet"]
camps = list(M.CAMPS_PEAK)
for k in range(6):
    r = r2 + 2 + k
    cell(wt, r, 1, f"=Inputs!A{18+k}")
    cell(wt, r, 2, f"=SUMIF($B$7:$B${AL},A{r},$I$7:$I${AL})", "#,##0")
    cell(wt, r, 3, unV[camps[k]], "#,##0", B)
    cell(wt, r, 4, f"=B{r}+C{r}", "#,##0")
    cell(wt, r, 5, f"=Inputs!B{18+k}*INDEX(Inputs!$B$29:$B$32,MATCH($B$3,Inputs!$A$29:$A$32,0))", "#,##0")
CMP = (r2 + 2, r2 + 7)

r3 = r2 + 10  # 50
put(wt, f"A{r3}", "Objective (minimise)", B)
cell(wt, r3 + 1, 1, "Transport cost"); cell(wt, r3 + 1, 2, f"=SUM(J7:J{AL})", "#,##0")
cell(wt, r3 + 2, 1, "Shortfall penalty"); cell(wt, r3 + 2, 2, f"=B4*SUM(C{CMP[0]}:C{CMP[1]})", "#,##0")
cell(wt, r3 + 3, 1, "TOTAL COST Rs", font=B); cell(wt, r3 + 3, 2, f"=B{r3+1}+B{r3+2}", "#,##0", B)
cell(wt, r3 + 4, 1, "Kits short"); cell(wt, r3 + 4, 2, f"=SUM(C{CMP[0]}:C{CMP[1]})", "#,##0")
OBJ = f"$B${r3+3}"

notes = [
    "SOLVER SET-UP (pre-loaded; check under Data > Solver)",
    f"Set objective: {OBJ}   To: Min",
    f"By changing: $I$7:$I${AL}, $C${CMP[0]}:$C${CMP[1]}",
    f"1. $I$7:$I${AL} <= $H$7:$H${AL}   (road capacity in this regime)",
    f"2. $B${SRC[0]}:$B${SRC[1]} <= $D${SRC[0]}:$D${SRC[1]}   (supply)",
    f"3. $D${HUB[0]}:$D${HUB[1]} = 0   (hub flow balance, the transshipment condition)",
    f"4. $B${HUB[0]}:$B${HUB[1]} <= $E${HUB[0]}:$E${HUB[1]}   (hub throughput)",
    f"5. $D${CMP[0]}:$D${CMP[1]} = $E${CMP[0]}:$E${CMP[1]}   (meet demand or record shortfall)",
    "Tick 'Make unconstrained variables non-negative'. Method: Simplex LP.",
    "Values shown are the optimum for regime V. Python gives Rs 13,14,180 total for V.",
    "Unmet demand is allowed at a penalty so the LP stays feasible when roads fail;",
    "the shortfall shows exactly where roads, not supply, are the bottleneck.",
]
for k, n in enumerate(notes):
    put(wt, f"M{6+k}", n, B if k == 0 else N)

solver_names(
    wt, wb.sheetnames.index("Transship_LP"), OBJ, 2,
    f"$I$7:$I${AL},$C${CMP[0]}:$C${CMP[1]}",
    [
        (f"$I$7:$I${AL}", 1, f"$H$7:$H${AL}"),
        (f"$B${SRC[0]}:$B${SRC[1]}", 1, f"$D${SRC[0]}:$D${SRC[1]}"),
        (f"$D${HUB[0]}:$D${HUB[1]}", 2, "0"),
        (f"$B${HUB[0]}:$B${HUB[1]}", 1, f"$E${HUB[0]}:$E${HUB[1]}"),
        (f"$D${CMP[0]}:$D${CMP[1]}", 2, f"$E${CMP[0]}:$E${CMP[1]}"),
    ],
)

# ============================================================================
# PLAYBOOK (Python output)
# ============================================================================
wp = wb.create_sheet("Playbook")
widths(wp, {"A": 18, "B": 18, "C": 12, "D": 12, "E": 12, "F": 12, "H": 60})
put(wp, "A1", "Regime playbook: optimal flows (kits/day) for each rain regime", T)
put(wp, "A2", "Values from rain_ready_chennai.py (transshipment LP solved for each regime).", I)
header(wp, 4, 1, ["From", "To"] + [f"Regime {s}" for s in M.STATES])
for k, (f, t, *_rest) in enumerate(M.ARCS):
    r = 5 + k
    cell(wp, r, 1, f); cell(wp, r, 2, t)
    for j, s in enumerate(M.STATES):
        cell(wp, r, 3 + j, R["playbook"][s]["flows"][f"{f} -> {t}"], "#,##0")
r = 5 + NA + 1
header(wp, r, 1, ["Camp", "Unmet kits"] + [f"Regime {s}" for s in M.STATES])
for k, c in enumerate(camps):
    cell(wp, r + 1 + k, 1, c); cell(wp, r + 1 + k, 2, "")
    for j, s in enumerate(M.STATES):
        cell(wp, r + 1 + k, 3 + j, R["playbook"][s]["unmet"][c], "#,##0")
r = r + 8
summary = [("Total demand", "demand_total"), ("Total unmet", "unmet_total"),
           ("Transport cost Rs", "transport_cost"), ("Total cost incl. penalty Rs", "total_cost")]
header(wp, r, 1, ["Summary", ""] + [f"Regime {s}" for s in M.STATES])
for k, (lab, key) in enumerate(summary):
    cell(wp, r + 1 + k, 1, lab); cell(wp, r + 1 + k, 2, "")
    for j, s in enumerate(M.STATES):
        cell(wp, r + 1 + k, 3 + j, R["playbook"][s][key], "#,##0")
r = r + 6
header(wp, r, 1, ["Shadow price, regime V", "Rs per extra kit"])
for k, c in enumerate(camps):
    cell(wp, r + 1 + k, 1, c)
    cell(wp, r + 1 + k, 2, R["playbook"]["V"]["duals_demand"][c], "#,##0.0")
put(wp, "H4", "How to read", B)
put(wp, "H5", "D, M and H: everything is delivered on the same routes, scaled up (only 60 kits re-routed in H).")
put(wp, "H6", "V: the network reorganises; Sholinganallur hub (unused until now) carries 2,175 kits from Chengalpattu.")
put(wp, "H7", "V: 2,225 kits short, all at Velachery (700) and Mudichur (1,525).")
put(wp, "H8", "Supply is not the problem: only the last-mile roads into those two camps fail.")
put(wp, "H9", "A shadow price of Rs 500 (= penalty) at Velachery and Mudichur confirms they are road-bound.")

# ============================================================================
# PREPOSITIONING (Python output)
# ============================================================================
wq = wb.create_sheet("Prepositioning")
widths(wq, {"A": 34, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "H": 70})
put(wq, "A1", "Two-stage stochastic LP: what to pre-position today", T)
put(wq, "A2", "Stage 1 (today): stock kits at hubs/camps. Stage 2 (tomorrow, 4 rain scenarios weighted by "
              "the Markov row of today's regime): route, serve, pay penalty or waste.", I)
st = R["stochastic"]
put(wq, "A4", "Case: today = M (Moderate rain, IMD alert)", B)
header(wq, 5, 1, ["Tomorrow's regime"] + M.STATES)
cell(wq, 6, 1, "Probability (Markov row M)")
for j, s in enumerate(M.STATES):
    cell(wq, 6, 2 + j, st["probs"][s], "0.000")

header(wq, 8, 1, ["Plan", "Expected cost Rs"])
plans = [("RP: stochastic plan", st["RP"]["objective"]),
         ("EEV: plan for the 'average' day", st["EEV"]["objective"]),
         (f"EML: plan for likeliest day ({st['likely']})", st["EML"]["objective"]),
         ("No pre-positioning", st["NONE"]["objective"]),
         ("WS: perfect forecast (wait and see)", st["WS"])]
for k, (a, v) in enumerate(plans):
    cell(wq, 9 + k, 1, a); cell(wq, 9 + k, 2, round(v, 0), "#,##0")
cell(wq, 15, 1, "VSS = EEV - RP", font=B); cell(wq, 15, 2, "=B10-B9", "#,##0", B)
cell(wq, 16, 1, "EVPI = RP - WS", font=B); cell(wq, 16, 2, "=B9-B13", "#,##0", B)
cell(wq, 17, 1, "VSS as % of EEV"); cell(wq, 17, 2, "=B15/B10", "0.0%")

header(wq, 19, 1, ["RP plan: stock at", "Kits"])
rr = 20
for c, v in st["RP"]["camp_stock"].items():
    cell(wq, rr, 1, f"Camp: {c}"); cell(wq, rr, 2, round(v), "#,##0"); rr += 1
for h, v in st["RP"]["hub_stock"].items():
    cell(wq, rr, 1, f"Hub: {h}"); cell(wq, rr, 2, round(v), "#,##0"); rr += 1

header(wq, 31, 1, ["Kits short tomorrow"] + M.STATES)
for k, (lab, key) in enumerate([("RP plan", "RP"), ("No pre-positioning", "NONE")]):
    cell(wq, 32 + k, 1, lab)
    for j, s in enumerate(M.STATES):
        cell(wq, 32 + k, 2 + j, round(st[key]["scenario"][s]["unmet"]), "#,##0")

put(wq, "A35", "Policy table: re-solve for each possible 'today'", B)
header(wq, 36, 1, ["Today's regime", "Kits pre-positioned", "Expected cost RP", "No pre-positioning",
                   "VSS", "EVPI"])
for k, td in enumerate(["M", "H", "V"]):
    p = R["policy"][td]
    tot = sum(p["camp_stock"].values()) + sum(p["hub_stock"].values())
    r = 37 + k
    cell(wq, r, 1, td); cell(wq, r, 2, round(tot), "#,##0"); cell(wq, r, 3, round(p["RP"]), "#,##0")
    cell(wq, r, 4, round(p["NONE"]), "#,##0"); cell(wq, r, 5, round(p["VSS"]), "#,##0")
    cell(wq, r, 6, round(p["EVPI"]), "#,##0")
header(wq, 41, 1, ["Where (today's regime)", "Velachery", "Mudichur", "Other camps", "Hubs"])
for k, td in enumerate(["M", "H", "V"]):
    p = R["policy"][td]
    cs = p["camp_stock"]
    cell(wq, 42 + k, 1, td)
    cell(wq, 42 + k, 2, round(cs["Velachery"]), "#,##0")
    cell(wq, 42 + k, 3, round(cs["Mudichur"]), "#,##0")
    cell(wq, 42 + k, 4, round(sum(v for c, v in cs.items() if c not in ("Velachery", "Mudichur"))), "#,##0")
    cell(wq, 42 + k, 5, round(sum(p["hub_stock"].values())), "#,##0")

put(wq, "A47", "Sensitivity to the shortfall penalty (today = M)", B)
header(wq, 48, 1, ["Penalty Rs/kit", "Kits pre-positioned", "Expected cost", "VSS", "Short if V"])
for k, srow in enumerate(R["sensitivity"]):
    r = 49 + k
    cell(wq, r, 1, srow["penalty"], "#,##0"); cell(wq, r, 2, round(srow["prepos_total"]), "#,##0")
    cell(wq, r, 3, round(srow["RP"]), "#,##0"); cell(wq, r, 4, round(srow["VSS"]), "#,##0")
    cell(wq, r, 5, round(srow["unmet_V"]), "#,##0")

notes = [
    "How to read",
    "VSS: what you lose by planning for an 'average' day instead of the full set of scenarios.",
    "Here the average-day plan, the likeliest-day plan and doing nothing are the same plan: stock nothing.",
    "EVPI: the most a better (perfect) one-day forecast is worth per decision day.",
    "Policy: pre-position the most when today is H. Once today is V, the Tambaram-Mudichur road",
    "is itself flooded (only 25% passable), so less can be moved: act one regime early.",
    "Every kit goes to a camp, none to hubs: hubs are reachable later, camps are not.",
    "Penalty sensitivity: below about Rs 100/kit nothing is worth stocking; above about Rs 1,000",
    "the plan saturates at 2,200 kits (Velachery 700, Mudichur 1,500 = its storage limit).",
]
for k, n in enumerate(notes):
    put(wq, f"H{4+k}", n, B if k == 0 else N)

# ============================================================================
# EVACUATION (Solver-ready most-reliable path)
# ============================================================================
we = wb.create_sheet("Evacuation")
widths(we, {"A": 18, "B": 18, "C": 8, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12, "I": 13,
            "J": 13, "L": 70})
put(we, "A1", "Evacuation from Mudichur: most-reliable route (Solver-ready) and road bottleneck", T)
put(we, "A3", "Regime (INPUT)", B)
cell(we, 3, 2, "V")
header(we, 5, 1, ["From", "To", "km", "Vulnerability", "Vehicles/h", "Reliability r",
                  "-ln(r)", "Effective veh/h", "Use (DECISION)", "Shortest-km path"])
evac_arcs = M.EVAC_ARCS + [(sh, "SAFE", 0, "low", 100000) for sh in M.SHELTERS]
rel_path = R["evac"]["V"]["reliable"]["path"] + ["SAFE"]
sh_path = R["evac"]["V"]["shortest"]["path"] + ["SAFE"]
rel_edges = set(zip(rel_path[:-1], rel_path[1:]))
sh_edges = set(zip(sh_path[:-1], sh_path[1:]))
for k, (f, t, km, v, cap) in enumerate(evac_arcs):
    r = 6 + k
    cell(we, r, 1, f); cell(we, r, 2, t); cell(we, r, 3, km, "0.0"); cell(we, r, 4, v)
    cell(we, r, 5, cap, "#,##0")
    if t == "SAFE":
        cell(we, r, 6, 1, "0.000")
    else:
        cell(we, r, 6, f"=1-INDEX(Inputs!$B$36:$E$38,MATCH(D{r},Inputs!$A$36:$A$38,0),MATCH($B$3,Inputs!$B$35:$E$35,0))", "0.000")
    cell(we, r, 7, f"=-LN(F{r})", "0.000")
    cell(we, r, 8, f"=E{r}*F{r}", "#,##0")
    cell(we, r, 9, 1 if (f, t) in rel_edges else 0, "0", B)
    cell(we, r, 10, 1 if (f, t) in sh_edges else 0, "0")
EL = 5 + len(evac_arcs)
put(we, f"A{EL+1}", "Rows to SAFE are a dummy super-sink so either shelter can be chosen.", I)

nodes = ["Mudichur", "Varadharajapuram", "Tambaram West", "Tambaram East", "Perungalathur",
         "Chromepet", "Vandalur", "Pallavaram", "SAFE"]
n0 = EL + 3
put(we, f"A{n0}", "Flow balance: out - in = +1 at Mudichur, -1 at SAFE, 0 elsewhere", B)
header(we, n0 + 1, 1, ["Node", "Out - in", "Required"])
for k, nd in enumerate(nodes):
    r = n0 + 2 + k
    cell(we, r, 1, nd)
    cell(we, r, 2, f"=SUMIF($A$6:$A${EL},A{r},$I$6:$I${EL})-SUMIF($B$6:$B${EL},A{r},$I$6:$I${EL})", "0")
    cell(we, r, 3, 1 if nd == "Mudichur" else (-1 if nd == "SAFE" else 0), "0")
NB = (n0 + 2, n0 + 1 + len(nodes))

o0 = NB[1] + 2
header(we, o0, 1, ["Route measure", "Chosen route", "Shortest-km route"])
cell(we, o0 + 1, 1, "Sum of -ln(r) (minimise)", font=B)
cell(we, o0 + 1, 2, f"=SUMPRODUCT(G6:G{EL},I6:I{EL})", "0.000", B)
cell(we, o0 + 1, 3, f"=SUMPRODUCT(G6:G{EL},J6:J{EL})", "0.000")
cell(we, o0 + 2, 1, "Route reliability = exp(-sum)")
cell(we, o0 + 2, 2, f"=EXP(-B{o0+1})", "0.0%")
cell(we, o0 + 2, 3, f"=EXP(-C{o0+1})", "0.0%")
cell(we, o0 + 3, 1, "Distance km")
cell(we, o0 + 3, 2, f"=SUMPRODUCT(C6:C{EL},I6:I{EL})", "0.0")
cell(we, o0 + 3, 3, f"=SUMPRODUCT(C6:C{EL},J6:J{EL})", "0.0")
OBJE = f"$B${o0+1}"

m0 = o0 + 6
put(we, f"A{m0}", "Max-flow / min-cut (from Python, networkx)", B)
header(we, m0 + 1, 1, ["Regime", "Max evacuation flow veh/h", "Min cut (bottleneck roads)"])
for k, s in enumerate(["H", "V"]):
    e = R["evac"][s]
    cell(we, m0 + 2 + k, 1, s); cell(we, m0 + 2 + k, 2, e["maxflow"], "#,##0")
    cell(we, m0 + 2 + k, 3, "; ".join(e["mincut"]))
p0 = m0 + 5
put(we, f"A{p0}", "Where to place one dewatering pump (makes that road low-risk), regime V", B)
header(we, p0 + 1, 1, ["Road", "Was", "New max flow", "Gain veh/h"])
for k, u in enumerate(R["evac"]["V"]["upgrades"]):
    r = p0 + 2 + k
    cell(we, r, 1, u["arc"]); cell(we, r, 2, u["from_vuln"]); cell(we, r, 3, u["new_maxflow"], "#,##0")
    cell(we, r, 4, u["gain"], "#,##0")
we.column_dimensions["A"].width = 36

enotes = [
    "SOLVER SET-UP (pre-loaded)",
    f"Set objective: {OBJE}  To: Min    By changing: $I$6:$I${EL}",
    f"Constraint: $B${NB[0]}:$B${NB[1]} = $C${NB[0]}:$C${NB[1]};  $I$6:$I${EL} <= 1;  non-negative; Simplex LP",
    "Why -ln(r)? Reliability of a route is the PRODUCT of road reliabilities.",
    "Maximising a product = minimising the sum of -ln(r), so it becomes an ordinary shortest-path LP.",
    "Result in V: the shortest route (8.0 km via Varadharajapuram) gets through only 13.5% of the time;",
    "the most-reliable route (8.5 km via Perungalathur) gets through 81%. 0.5 km more for 6x the odds.",
    "Min cut = the three roads leaving Mudichur: the bottleneck is the first link out, not the shelters.",
    "One pump on Mudichur - Varadharajapuram adds the most evacuation capacity (+180 veh/h in V).",
]
for k, n in enumerate(enotes):
    put(we, f"L{5+k}", n, B if k == 0 else N)

solver_names(
    we, wb.sheetnames.index("Evacuation"), OBJE, 2, f"$I$6:$I${EL}",
    [(f"$B${NB[0]}:$B${NB[1]}", 2, f"$C${NB[0]}:$C${NB[1]}"),
     (f"$I$6:$I${EL}", 1, "1")],
)

for wsx in wb.worksheets:
    wsx.sheet_view.showGridLines = True

out = os.path.join(HERE, "Rain_Ready_Chennai_Decision_Science.xlsx")
wb.save(out)
print("saved", out)
