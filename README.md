# Rain-Ready Chennai

Planning relief supplies and evacuation for Chennai's North-East Monsoon using Markov chains, linear programming and network models.

**Live interactive page:** https://shivnath250.github.io/rain-ready-chennai/

Decision Science: Creative AI Project · Shivnath Rajendra Gujar (FT271017) · PGPM 2026-27, Great Lakes Institute of Management, Chennai

## What the project does

| Step | Model | Question |
|---|---|---|
| 1 | Markov chain | How likely is each rain level tomorrow, and how long does a wet spell last? |
| 2 | Transshipment LP | What is the cheapest routing for each rain level, and where do flooded roads cause shortages? |
| 3 | Two-stage stochastic LP | What should be stocked at camps tonight, before tomorrow's rain is known? |
| 4 | Network models | Which evacuation road is safest, and where is the bottleneck? |

## Key results

- On a very heavy rain day, 2,225 kits cannot reach Velachery and Mudichur while 10,625 kits stay unused in warehouses: roads, not supply, cause the shortage.
- On a moderate day, stocking 700 kits at Velachery and 600 at Mudichur saves Rs 10,547 in expected cost (VSS); a perfect forecast would be worth up to Rs 60,244 a day (EVPI).
- Stock on the heavy day: once rain is very heavy, the road into Mudichur is mostly blocked.
- On a very heavy day the shortest route out of Mudichur gets through 13.5% of the time; the safest route (0.5 km longer) gets through 81%.

## Files

| File | Contents |
|---|---|
| `Rain_Ready_Chennai_Report.docx` | Full report |
| `Rain_Ready_Chennai_Slides.pptx` | Presentation |
| `Rain_Ready_Chennai_Decision_Science.xlsx` | Excel template with Solver models stored in it (transshipment LP and safest-route LP) |
| `rain_ready_chennai.py` | Python model that runs all four models |
| `index.html` | Interactive page (runs a min-cost flow optimiser in the browser) |
| `outputs/` | Results, synthetic rainfall series and figures |
| `docs-src/` | Scripts used to build the Excel file, report, slides and page |

## Run the Python model

```
pip install "pulp>=2.7,<3" networkx numpy matplotlib
python rain_ready_chennai.py
```

## Data note

Rain levels follow IMD daily rainfall classes. The rainfall series is synthetic and the network costs, capacities and flood probabilities are illustrative. Replace them in the Excel `Inputs` and `Rain_Data` sheets or the `DATA` section of the Python file to use real data.

Built with AI assistance (Claude); all results were cross-checked between Python, Excel and the web page.
