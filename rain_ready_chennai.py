"""
Rain-Ready Chennai
A Markov-informed relief logistics and evacuation planner for the North-East Monsoon

Decision Science - Creative AI project
Concepts used: Markov chains, transshipment LP, two-stage stochastic LP,
network models (most-reliable path, max-flow / min-cut)

Run:  pip install "pulp>=2.7,<3" networkx numpy matplotlib
      python rain_ready_chennai.py
Outputs go to ./outputs (CSV tables, results.json, figures)

All network data below is ILLUSTRATIVE (built for teaching, not surveyed).
Replace the tables in the DATA section with real figures to reuse the template.
"""

import json
import math
import os
from itertools import product

import networkx as nx
import numpy as np
import pulp

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outputs")
os.makedirs(OUT, exist_ok=True)

# =============================================================================
# DATA
# =============================================================================

# Rain regimes, using IMD daily rainfall categories (mm/day)
STATES = ["D", "M", "H", "V"]
STATE_NAMES = {
    "D": "Dry / light (< 15.6 mm)",
    "M": "Moderate (15.6 - 64.4 mm)",
    "H": "Heavy (64.5 - 115.5 mm)",
    "V": "Very heavy / extreme (>= 115.6 mm)",
}
STATE_BOUNDS = [15.6, 64.5, 115.6]  # lower bounds of M, H, V

# Sources: kits available per day
SOURCES = {"Koyambedu": 9000, "Madhavaram": 6000, "Chengalpattu": 7000}

# Transshipment hubs: max kits handled per day (also storage limit for pre-positioning)
HUBS = {"Guindy": 10000, "Tambaram": 8000, "Sholinganallur": 6000}

# Relief camps: kits needed per day in the worst regime (V)
CAMPS_PEAK = {
    "Velachery": 3200,
    "Pallikaranai": 2600,
    "Mudichur": 2400,
    "Saidapet": 1800,
    "Perumbakkam": 2000,
    "Vyasarpadi": 1600,
}
CAMP_STORAGE = 1500  # kits a camp can hold if pre-positioned (school halls etc.)

# Share of peak demand that appears in each regime
DEMAND_SHARE = {"D": 0.05, "M": 0.25, "H": 0.60, "V": 1.00}

# Probability that a road arc is impassable, by vulnerability class and regime
FLOOD_PROB = {
    "low": {"D": 0.00, "M": 0.02, "H": 0.05, "V": 0.10},
    "med": {"D": 0.00, "M": 0.05, "H": 0.20, "V": 0.40},
    "high": {"D": 0.00, "M": 0.10, "H": 0.40, "V": 0.75},
}
# Cost multiplier for detours / slower trucks
COST_MULT = {"D": 1.00, "M": 1.10, "H": 1.30, "V": 1.60}

# Arcs: (from, to, base cost Rs/kit, base capacity kits/day, vulnerability)
ARCS = [
    ("Koyambedu", "Guindy", 6, 8000, "med"),
    ("Koyambedu", "Tambaram", 10, 5000, "med"),
    ("Koyambedu", "Saidapet", 7, 3000, "high"),
    ("Koyambedu", "Vyasarpadi", 8, 2500, "med"),
    ("Madhavaram", "Guindy", 11, 4000, "med"),
    ("Madhavaram", "Vyasarpadi", 4, 4000, "med"),
    ("Madhavaram", "Sholinganallur", 16, 3000, "low"),
    ("Chengalpattu", "Tambaram", 7, 7000, "low"),
    ("Chengalpattu", "Sholinganallur", 12, 4000, "low"),
    ("Guindy", "Velachery", 4, 4000, "high"),
    ("Guindy", "Saidapet", 3, 3000, "med"),
    ("Guindy", "Pallikaranai", 6, 2500, "high"),
    ("Guindy", "Vyasarpadi", 9, 2000, "med"),
    ("Tambaram", "Mudichur", 3, 3500, "high"),
    ("Tambaram", "Perumbakkam", 5, 3000, "med"),
    ("Tambaram", "Pallikaranai", 6, 2500, "med"),
    ("Sholinganallur", "Pallikaranai", 4, 3000, "med"),
    ("Sholinganallur", "Perumbakkam", 3, 3000, "med"),
    ("Sholinganallur", "Velachery", 6, 2500, "med"),
]

PENALTY = 500      # Rs per kit of unmet demand (humanitarian shortfall cost)
HOLDING = 3        # Rs per kit pre-positioned (loading, guarding)
WASTE = 30         # Rs per pre-positioned kit left unused at a camp (return / spoilage)

TODAY = "M"        # regime today, when the pre-positioning decision is taken

# Evacuation road network around Mudichur (km, vulnerability, vehicles/hour)
EVAC_ARCS = [
    ("Mudichur", "Varadharajapuram", 2.0, "high", 400),
    ("Mudichur", "Tambaram West", 3.0, "med", 600),
    ("Mudichur", "Perungalathur", 5.5, "low", 300),
    ("Varadharajapuram", "Tambaram East", 2.5, "high", 400),
    ("Varadharajapuram", "Perungalathur", 3.0, "med", 300),
    ("Tambaram West", "Tambaram East", 1.5, "high", 800),
    ("Tambaram West", "Chromepet", 4.0, "med", 500),
    ("Tambaram East", "Chromepet", 3.5, "low", 700),
    ("Perungalathur", "Vandalur", 3.0, "low", 600),
    ("Tambaram East", "Vandalur", 5.0, "low", 500),
    ("Chromepet", "Pallavaram", 2.5, "low", 700),
]
EVAC_ORIGIN = "Mudichur"
SHELTERS = ["Vandalur", "Pallavaram"]


# =============================================================================
# 1. RAINFALL DATA AND MARKOV CHAIN
# =============================================================================

def generate_synthetic_rain(seed=42, seasons=5, days=92):
    """Synthetic Oct-Dec daily rainfall. Replace with IMD Nungambakkam data."""
    true_P = np.array([
        [0.78, 0.15, 0.05, 0.02],
        [0.45, 0.35, 0.13, 0.07],
        [0.30, 0.35, 0.22, 0.13],
        [0.20, 0.30, 0.25, 0.25],
    ])
    rng = np.random.default_rng(seed)
    rows = []
    for s in range(seasons):
        st = 0
        for d in range(days):
            if d > 0:
                st = rng.choice(4, p=true_P[st])
            if st == 0:
                mm = 0.0 if rng.random() < 0.7 else rng.uniform(0.1, 15.5)
            elif st == 1:
                mm = rng.uniform(15.6, 64.4)
            elif st == 2:
                mm = rng.uniform(64.5, 115.5)
            else:
                mm = rng.uniform(115.6, 250.0)
            rows.append((2021 + s, d + 1, round(mm, 1)))
    return rows


def classify(mm):
    if mm < STATE_BOUNDS[0]:
        return "D"
    if mm < STATE_BOUNDS[1]:
        return "M"
    if mm < STATE_BOUNDS[2]:
        return "H"
    return "V"


def estimate_P(rain_rows):
    counts = np.zeros((4, 4))
    for (s1, _, mm1), (s2, _, mm2) in zip(rain_rows[:-1], rain_rows[1:]):
        if s1 != s2:
            continue  # do not link 31 Dec of one year to 1 Oct of the next
        counts[STATES.index(classify(mm1)), STATES.index(classify(mm2))] += 1
    P = counts / counts.sum(axis=1, keepdims=True)
    return counts, P


def markov_analysis(P, today):
    i = STATES.index(today)
    pi0 = np.zeros(4)
    pi0[i] = 1
    forecast = [pi0]
    for _ in range(7):
        forecast.append(forecast[-1] @ P)
    # steady state: solve pi (P - I) = 0, sum = 1
    A = np.vstack([(P.T - np.eye(4))[:-1], np.ones(4)])
    b = np.zeros(4)
    b[-1] = 1
    steady = np.linalg.solve(A, b)
    # mean first-passage time to Dry from each wet state: m = (I - Q)^-1 1
    Q = P[1:, 1:]
    m = np.linalg.solve(np.eye(3) - Q, np.ones(3))
    # expected length of a run of each state
    spell = 1 / (1 - np.diag(P))
    return np.array(forecast), steady, m, spell


# =============================================================================
# 2. REGIME-DEPENDENT NETWORK
# =============================================================================

def arc_params(regime):
    out = {}
    for f, t, c, cap, v in ARCS:
        p = FLOOD_PROB[v][regime]
        out[(f, t)] = {
            "cost": c * COST_MULT[regime],
            "cap": cap * (1 - p),
        }
    return out


def expected_arc_params(probs):
    """Probability-weighted network, for the expected-value (EV) problem."""
    out = {}
    for f, t, c, cap, v in ARCS:
        out[(f, t)] = {
            "cost": sum(probs[s] * c * COST_MULT[s] for s in STATES),
            "cap": sum(probs[s] * cap * (1 - FLOOD_PROB[v][s]) for s in STATES),
        }
    return out


def demand(regime):
    return {c: d * DEMAND_SHARE[regime] for c, d in CAMPS_PEAK.items()}


def solve(prob):
    prob.solve(pulp.PULP_CBC_CMD(msg=False))
    assert pulp.LpStatus[prob.status] == "Optimal", pulp.LpStatus[prob.status]


# =============================================================================
# 3. MODEL 1 - TRANSSHIPMENT LP FOR EACH REGIME (the "playbook")
# =============================================================================

def transshipment(regime):
    A = arc_params(regime)
    d = demand(regime)
    m = pulp.LpProblem(f"transship_{regime}", pulp.LpMinimize)
    x = {a: pulp.LpVariable(f"x_{a[0]}_{a[1]}", lowBound=0, upBound=A[a]["cap"]) for a in A}
    u = {c: pulp.LpVariable(f"unmet_{c}", lowBound=0) for c in CAMPS_PEAK}
    m += pulp.lpSum(A[a]["cost"] * x[a] for a in A) + PENALTY * pulp.lpSum(u.values())
    for s, cap in SOURCES.items():
        m += pulp.lpSum(x[a] for a in A if a[0] == s) <= cap, f"supply_{s}"
    for h, cap in HUBS.items():
        inflow = pulp.lpSum(x[a] for a in A if a[1] == h)
        outflow = pulp.lpSum(x[a] for a in A if a[0] == h)
        m += inflow == outflow, f"balance_{h}"
        m += inflow <= cap, f"throughput_{h}"
    for c in CAMPS_PEAK:
        m += pulp.lpSum(x[a] for a in A if a[1] == c) + u[c] == d[c], f"demand_{c}"
    solve(m)
    transport = sum(A[a]["cost"] * x[a].value() for a in A)
    return {
        "regime": regime,
        "flows": {f"{a[0]} -> {a[1]}": round(x[a].value(), 1) for a in A},
        "unmet": {c: round(u[c].value(), 1) for c in u},
        "transport_cost": round(transport, 0),
        "total_cost": round(pulp.value(m.objective), 0),
        "demand_total": round(sum(d.values()), 0),
        "unmet_total": round(sum(u[c].value() for c in u), 1),
        "duals_demand": {c: round(m.constraints[f"demand_{c}"].pi, 2) for c in CAMPS_PEAK},
    }


# =============================================================================
# 4. MODEL 2 - TWO-STAGE STOCHASTIC PRE-POSITIONING
# =============================================================================

def prepositioning(scen_probs, scen_arcs, scen_demand, fixed_stage1=None, today=None):
    """
    Stage 1 (today, regime TODAY): move kits from sources to hubs/camps and hold them.
    Stage 2 (tomorrow, each scenario): ship again on that scenario's network,
    meet demand, pay penalty for shortfall and waste for unused camp stock.
    scen_* are dicts keyed by scenario name.
    """
    today = today or TODAY
    A1 = arc_params(today)
    m = pulp.LpProblem("preposition", pulp.LpMinimize)
    y = {a: pulp.LpVariable(f"y_{a[0]}_{a[1]}", lowBound=0, upBound=A1[a]["cap"]) for a in A1}
    S = list(scen_probs)
    x = {(s, a): pulp.LpVariable(f"x_{s}_{a[0]}_{a[1]}", lowBound=0, upBound=scen_arcs[s][a]["cap"])
         for s in S for a in A1}
    u = {(s, c): pulp.LpVariable(f"u_{s}_{c}", lowBound=0) for s in S for c in CAMPS_PEAK}
    w = {(s, c): pulp.LpVariable(f"w_{s}_{c}", lowBound=0) for s in S for c in CAMPS_PEAK}

    # stage-1 stock at hubs and camps
    hub_stock = {h: pulp.lpSum(y[a] for a in A1 if a[1] == h) - pulp.lpSum(y[a] for a in A1 if a[0] == h)
                 for h in HUBS}
    camp_stock = {c: pulp.lpSum(y[a] for a in A1 if a[1] == c) for c in CAMPS_PEAK}

    stage1_cost = pulp.lpSum((A1[a]["cost"] + HOLDING) * y[a] for a in A1)
    stage2_cost = pulp.lpSum(
        scen_probs[s] * (
            pulp.lpSum(scen_arcs[s][a]["cost"] * x[s, a] for a in A1)
            + PENALTY * pulp.lpSum(u[s, c] for c in CAMPS_PEAK)
            + WASTE * pulp.lpSum(w[s, c] for c in CAMPS_PEAK)
        ) for s in S)
    m += stage1_cost + stage2_cost

    for src, cap in SOURCES.items():
        m += pulp.lpSum(y[a] for a in A1 if a[0] == src) <= cap
    for h, cap in HUBS.items():
        m += hub_stock[h] >= 0
        m += pulp.lpSum(y[a] for a in A1 if a[1] == h) <= cap
    for c in CAMPS_PEAK:
        m += camp_stock[c] <= CAMP_STORAGE

    for s in S:
        A = scen_arcs[s]
        d = scen_demand[s]
        for src, cap in SOURCES.items():
            m += pulp.lpSum(x[s, a] for a in A if a[0] == src) <= cap
        for h, cap in HUBS.items():
            inflow = pulp.lpSum(x[s, a] for a in A if a[1] == h)
            outflow = pulp.lpSum(x[s, a] for a in A if a[0] == h)
            m += hub_stock[h] + inflow - outflow >= 0
            m += inflow <= cap
        for c in CAMPS_PEAK:
            arrive = pulp.lpSum(x[s, a] for a in A if a[1] == c)
            # stock + deliveries + shortfall = demand + unused stock
            m += camp_stock[c] + arrive + u[s, c] == d[c] + w[s, c]
            m += w[s, c] <= camp_stock[c]  # only pre-positioned stock can be "wasted"

    if fixed_stage1 is not None:
        for a in A1:
            m += y[a] == fixed_stage1[a]
    solve(m)
    res = {
        "objective": pulp.value(m.objective),
        "stage1_cost": pulp.value(stage1_cost),
        "y": {a: y[a].value() for a in A1},
        "hub_stock": {h: pulp.value(hub_stock[h]) for h in HUBS},
        "camp_stock": {c: pulp.value(camp_stock[c]) for c in CAMPS_PEAK},
        "scenario": {},
    }
    for s in S:
        res["scenario"][s] = {
            "unmet": sum(u[s, c].value() for c in CAMPS_PEAK),
            "waste": sum(w[s, c].value() for c in CAMPS_PEAK),
            "cost": (sum(scen_arcs[s][a]["cost"] * x[s, a].value() for a in A1)
                     + PENALTY * sum(u[s, c].value() for c in CAMPS_PEAK)
                     + WASTE * sum(w[s, c].value() for c in CAMPS_PEAK)),
        }
    return res


def stochastic_study(probs, today=None):
    today = today or TODAY
    scen_arcs = {s: arc_params(s) for s in STATES}
    scen_dem = {s: demand(s) for s in STATES}
    # Recourse problem (RP): the stochastic solution
    rp = prepositioning(probs, scen_arcs, scen_dem, today=today)
    # Expected-value problem (EV): plan for the "average" day
    ev_arcs = {"EV": expected_arc_params(probs)}
    ev_dem = {"EV": {c: sum(probs[s] * scen_dem[s][c] for s in STATES) for c in CAMPS_PEAK}}
    ev = prepositioning({"EV": 1.0}, ev_arcs, ev_dem, today=today)
    eev = prepositioning(probs, scen_arcs, scen_dem, fixed_stage1=ev["y"], today=today)
    # Most-likely-scenario plan
    likely = max(probs, key=probs.get)
    ml = prepositioning({likely: 1.0}, {likely: scen_arcs[likely]}, {likely: scen_dem[likely]}, today=today)
    eml = prepositioning(probs, scen_arcs, scen_dem, fixed_stage1=ml["y"], today=today)
    # No pre-positioning
    zero = {a: 0.0 for a in arc_params(today)}
    enone = prepositioning(probs, scen_arcs, scen_dem, fixed_stage1=zero, today=today)
    # Wait-and-see (perfect information)
    ws = sum(probs[s] * prepositioning({s: 1.0}, {s: scen_arcs[s]}, {s: scen_dem[s]}, today=today)["objective"]
             for s in STATES)
    return {
        "probs": probs,
        "RP": rp, "EEV": eev, "EML": eml, "NONE": enone, "WS": ws,
        "likely": likely,
        "VSS": eev["objective"] - rp["objective"],
        "EVPI": rp["objective"] - ws,
    }


# =============================================================================
# 5. MODEL 3 - EVACUATION: MOST-RELIABLE PATH AND MAX-FLOW / MIN-CUT
# =============================================================================

def evac_graph(regime, upgraded=None):
    G = nx.DiGraph()
    for f, t, km, v, cap in EVAC_ARCS:
        vv = "low" if upgraded == (f, t) else v
        r = 1 - FLOOD_PROB[vv][regime]
        G.add_edge(f, t, km=km, rel=r, neglog=-math.log(r) if r > 0 else 1e9,
                   cap=cap * r, base_cap=cap, vuln=vv)
    for sh in SHELTERS:
        G.add_edge(sh, "SAFE", km=0, rel=1, neglog=0, cap=1e9, base_cap=1e9, vuln="low")
    return G


def path_stats(G, path):
    edges = list(zip(path[:-1], path[1:]))
    km = sum(G[a][b]["km"] for a, b in edges)
    rel = math.prod(G[a][b]["rel"] for a, b in edges)
    return km, rel


def evacuation(regime):
    G = evac_graph(regime)
    short = nx.shortest_path(G, EVAC_ORIGIN, "SAFE", weight="km")[:-1]
    reliable = nx.shortest_path(G, EVAC_ORIGIN, "SAFE", weight="neglog")[:-1]
    flow, _ = nx.maximum_flow(G, EVAC_ORIGIN, "SAFE", capacity="cap")
    cut_val, (S_side, T_side) = nx.minimum_cut(G, EVAC_ORIGIN, "SAFE", capacity="cap")
    cut = [(a, b) for a in S_side for b in G[a] if b in T_side]
    # Where should the single dewatering pump / NDRF team go?
    upgrades = []
    for f, t, km, v, cap in EVAC_ARCS:
        if v == "low":
            continue
        G2 = evac_graph(regime, upgraded=(f, t))
        f2, _ = nx.maximum_flow(G2, EVAC_ORIGIN, "SAFE", capacity="cap")
        upgrades.append({"arc": f"{f} -> {t}", "from_vuln": v,
                         "new_maxflow": round(f2, 0), "gain": round(f2 - flow, 0)})
    upgrades.sort(key=lambda r: -r["gain"])
    ks, rs = path_stats(G, short)
    kr, rr = path_stats(G, reliable)
    return {
        "regime": regime,
        "shortest": {"path": short, "km": ks, "reliability": rs},
        "reliable": {"path": reliable, "km": kr, "reliability": rr},
        "maxflow": round(flow, 0),
        "mincut": [f"{a} -> {b}" for a, b in cut],
        "mincut_value": round(cut_val, 0),
        "upgrades": upgrades,
    }


# =============================================================================
# MAIN
# =============================================================================

def main():
    rain = generate_synthetic_rain()
    with open(os.path.join(OUT, "rain_data.csv"), "w") as f:
        f.write("Season,Day,Rain_mm,State\n")
        for s, d, mm in rain:
            f.write(f"{s},{d},{mm},{classify(mm)}\n")

    counts, P = estimate_P(rain)
    forecast, steady, mfpt, spell = markov_analysis(P, TODAY)
    probs_tomorrow = dict(zip(STATES, P[STATES.index(TODAY)]))

    print("=" * 70)
    print("1. MARKOV CHAIN")
    print("Transition counts:\n", counts.astype(int))
    print("Estimated P:\n", np.round(P, 3))
    print("Steady state:", dict(zip(STATES, np.round(steady, 3))))
    print("Mean days until Dry from M, H, V:", np.round(mfpt, 2))
    print("Expected spell length (days):", dict(zip(STATES, np.round(spell, 2))))
    print(f"Forecast from today = {TODAY}:")
    for n, row in enumerate(forecast):
        print(f"  day +{n}:", dict(zip(STATES, np.round(row, 3))))

    print("=" * 70)
    print("2. TRANSSHIPMENT PLAYBOOK")
    playbook = {s: transshipment(s) for s in STATES}
    for s, r in playbook.items():
        print(f"  {s}: demand {r['demand_total']:.0f}, unmet {r['unmet_total']:.0f}, "
              f"transport Rs {r['transport_cost']:,.0f}, total Rs {r['total_cost']:,.0f}")
        print("     unmet by camp:", {k: v for k, v in r["unmet"].items() if v > 0.5})

    print("=" * 70)
    print("3. STOCHASTIC PRE-POSITIONING (today =", TODAY, ")")
    st = stochastic_study(probs_tomorrow)
    print("  P(tomorrow):", {k: round(v, 3) for k, v in st["probs"].items()})
    print(f"  RP  (stochastic plan)       : Rs {st['RP']['objective']:,.0f}")
    print(f"  EEV (plan for average day)  : Rs {st['EEV']['objective']:,.0f}")
    print(f"  EML (plan for likeliest = {st['likely']}): Rs {st['EML']['objective']:,.0f}")
    print(f"  No pre-positioning          : Rs {st['NONE']['objective']:,.0f}")
    print(f"  WS  (perfect forecast)      : Rs {st['WS']:,.0f}")
    print(f"  VSS = Rs {st['VSS']:,.0f}   EVPI = Rs {st['EVPI']:,.0f}")
    print("  RP camp stock:", {k: round(v) for k, v in st["RP"]["camp_stock"].items()})
    print("  RP hub stock :", {k: round(v) for k, v in st["RP"]["hub_stock"].items()})
    for name in ["RP", "EEV", "EML", "NONE"]:
        print(f"  {name} unmet by scenario:",
              {s: round(v["unmet"]) for s, v in st[name]["scenario"].items()})


    print("  Shadow price of one more kit at each camp in regime V (Rs):", playbook["V"]["duals_demand"])

    print("-" * 70)
    print("  Policy table: what to pre-position, by today's regime")
    policy = {}
    for td in ["M", "H", "V"]:
        pr = dict(zip(STATES, P[STATES.index(td)]))
        r = stochastic_study(pr, today=td)
        policy[td] = {
            "probs": pr,
            "camp_stock": r["RP"]["camp_stock"], "hub_stock": r["RP"]["hub_stock"],
            "RP": r["RP"]["objective"], "EEV": r["EEV"]["objective"],
            "NONE": r["NONE"]["objective"], "WS": r["WS"],
            "VSS": r["VSS"], "EVPI": r["EVPI"],
            "unmet_RP": {s: v["unmet"] for s, v in r["RP"]["scenario"].items()},
            "unmet_NONE": {s: v["unmet"] for s, v in r["NONE"]["scenario"].items()},
        }
        tot = sum(r["RP"]["camp_stock"].values()) + sum(r["RP"]["hub_stock"].values())
        print(f"  today={td}: pre-position {tot:,.0f} kits | RP Rs {r['RP']['objective']:,.0f} | "
              f"no-prepos Rs {r['NONE']['objective']:,.0f} | VSS Rs {r['VSS']:,.0f} | EVPI Rs {r['EVPI']:,.0f}")
        print("      camps:", {k: round(v) for k, v in r["RP"]["camp_stock"].items() if v > 0.5},
              " hubs:", {k: round(v) for k, v in r["RP"]["hub_stock"].items() if v > 0.5})

    print("-" * 70)
    print("  Sensitivity: penalty per unmet kit (today = M)")
    global PENALTY
    base_pen = PENALTY
    sens = []
    for pen in [100, 250, 500, 1000, 2000]:
        PENALTY = pen
        r = stochastic_study(probs_tomorrow)
        tot = sum(r["RP"]["camp_stock"].values()) + sum(r["RP"]["hub_stock"].values())
        sens.append({"penalty": pen, "prepos_total": tot, "RP": r["RP"]["objective"],
                     "VSS": r["VSS"], "unmet_V": r["RP"]["scenario"]["V"]["unmet"]})
        print(f"   penalty Rs {pen:>5}: pre-position {tot:,.0f} kits, expected cost Rs {r['RP']['objective']:,.0f}, "
              f"VSS Rs {r['VSS']:,.0f}, unmet if V = {r['RP']['scenario']['V']['unmet']:,.0f}")
    PENALTY = base_pen

    print("=" * 70)
    print("4. EVACUATION FROM MUDICHUR")
    evac = {s: evacuation(s) for s in ["H", "V"]}
    for s, e in evac.items():
        print(f"  Regime {s}")
        print(f"    shortest : {' > '.join(e['shortest']['path'])} | "
              f"{e['shortest']['km']} km | reliability {e['shortest']['reliability']:.3f}")
        print(f"    reliable : {' > '.join(e['reliable']['path'])} | "
              f"{e['reliable']['km']} km | reliability {e['reliable']['reliability']:.3f}")
        print(f"    max flow {e['maxflow']} vehicles/h; min cut {e['mincut']} = {e['mincut_value']}")
        print(f"    best pump placement: {e['upgrades'][:3]}")

    # ---------- save results for the Excel builder and the report ----------
    def clean(o):
        if isinstance(o, dict):
            return {(" -> ".join(k) if isinstance(k, tuple) else k): clean(v) for k, v in o.items()}
        if isinstance(o, (list, tuple)):
            return [clean(v) for v in o]
        if isinstance(o, (np.floating, float)):
            return float(o)
        return o

    results = {
        "counts": counts.tolist(), "P": P.tolist(), "steady": steady.tolist(),
        "mfpt": mfpt.tolist(), "spell": spell.tolist(), "forecast": forecast.tolist(),
        "playbook": playbook, "stochastic": clean(st), "evac": evac,
        "policy": clean(policy), "sensitivity": sens,
    }
    with open(os.path.join(OUT, "results.json"), "w") as f:
        json.dump(results, f, indent=1, default=float)

    make_figures(forecast, playbook, evac)
    print("\nSaved outputs to", OUT)


def make_figures(forecast, playbook, evac):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        return
    # Forecast chart
    fig, ax = plt.subplots(figsize=(7, 3.6))
    days = range(len(forecast))
    styles = ["-", "--", "-.", ":"]
    for k, s in enumerate(STATES):
        ax.plot(days, forecast[:, k], color="black", linestyle=styles[k], marker="o",
                markersize=3, label=STATE_NAMES[s])
    ax.set_xlabel("Days ahead")
    ax.set_ylabel("Probability")
    ax.set_title(f"Rain-regime forecast starting from today = {TODAY}")
    ax.legend(fontsize=7, frameon=False)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, "fig_forecast.png"), dpi=200)
    plt.close(fig)

    # Network diagram (relief network)
    web = {"Madhavaram": (400, 60), "Vyasarpadi": (620, 110), "Koyambedu": (260, 230), "Saidapet": (560, 250),
           "Guindy": (420, 330), "Velachery": (580, 390), "Pallikaranai": (640, 480), "Sholinganallur": (790, 470),
           "Perumbakkam": (700, 600), "Tambaram": (380, 560), "Mudichur": (190, 600), "Chengalpattu": (300, 720)}
    pos = {k: (x, -y) for k, (x, y) in web.items()}
    G = nx.DiGraph()
    styles = {"low": "solid", "med": "dashed", "high": "dotted"}
    fig, ax = plt.subplots(figsize=(8, 7.5))
    for f, t, c, cap, v in ARCS:
        G.add_edge(f, t)
        nx.draw_networkx_edges(G, pos, edgelist=[(f, t)], style=styles[v], ax=ax, arrows=True,
                               arrowsize=14, edge_color="black", width=1.3, node_size=700)
    shapes = {"s": list(SOURCES), "D": list(HUBS), "o": list(CAMPS_PEAK)}
    for shp, nodes in shapes.items():
        nx.draw_networkx_nodes(G, pos, nodelist=nodes, node_shape=shp, node_color="white",
                               edgecolors="black", linewidths=1.5, node_size=700, ax=ax)
    nx.draw_networkx_labels(G, {k: (x, y - 42) for k, (x, y) in pos.items()},
                            font_size=10, font_weight="bold", ax=ax,
                            bbox=dict(facecolor="white", edgecolor="none", pad=0.5))
    ax.set_title("Squares = warehouses, diamonds = hubs, circles = camps\n"
                 "Solid line = low flood risk, dashed = medium, dotted = high", fontsize=10)
    ax.axis("off")
    ax.margins(0.08)
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, "fig_network.png"), dpi=200)
    plt.close(fig)


if __name__ == "__main__":
    main()
