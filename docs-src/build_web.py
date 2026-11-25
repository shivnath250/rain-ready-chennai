import json, os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import rain_ready_chennai as M
HERE = os.path.dirname(os.path.abspath(__file__))
R = json.load(open(os.path.join(HERE, "..", "outputs", "results.json")))
REPO = sys.argv[1] if len(sys.argv) > 1 else "#"
data = {
  "P": R["P"], "steady": R["steady"],
  "net": {"sources": M.SOURCES, "hubs": M.HUBS, "camps": M.CAMPS_PEAK, "share": M.DEMAND_SHARE,
          "mult": M.COST_MULT, "flood": M.FLOOD_PROB, "arcs": [list(a) for a in M.ARCS]},
  "policy": {k: {kk: v[kk] for kk in ["camp_stock","hub_stock","RP","NONE","WS","VSS","EVPI","EEV"]} for k, v in R["policy"].items()},
  "sens": R["sensitivity"],
  "evac": {"arcs": [list(a) for a in M.EVAC_ARCS], "shelters": M.SHELTERS},
}
html = open(os.path.join(HERE, "template.html")).read().replace("__DATA__", json.dumps(data)).replace("__REPO__", REPO)
open(os.path.join(HERE, "index.html"), "w").write(html)
print("ok", len(html))
