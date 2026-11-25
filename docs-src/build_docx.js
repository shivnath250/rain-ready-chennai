// Builds the Word report. Usage: node build_docx.js <repo_url> <pages_url>
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType,
  BorderStyle, AlignmentType, ImageRun, ExternalHyperlink, LevelFormat, Footer, PageNumber,
  PageBreak, TableOfContents,
} = require("docx");

const REPO = process.argv[2] || "https://github.com/USER/rain-ready-chennai";
const PAGE = process.argv[3] || "https://USER.github.io/rain-ready-chennai/";
const OUT = path.join(__dirname, "..", "outputs");

const FONT = "Arial";
const W = 9026; // A4 width minus 1 inch margins, in DXA
const line = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: line, bottom: line, left: line, right: line };

const t = (text, o = {}) => new TextRun({ text, font: FONT, ...o });
const P = (children, o = {}) =>
  new Paragraph({ children: typeof children === "string" ? [t(children)] : children, spacing: { after: 120, line: 288 }, ...o });
const B = (s) => t(s, { bold: true });
const H1 = (s) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [t(s)] });
const H2 = (s) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [t(s)] });
const bullet = (children) =>
  new Paragraph({ numbering: { reference: "bul", level: 0 }, spacing: { after: 80, line: 276 },
    children: typeof children === "string" ? [t(children)] : children });
const num = (children, ref = "num") =>
  new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 80, line: 276 },
    children: typeof children === "string" ? [t(children)] : children });
const formula = (s) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 160 },
  children: [t(s, { italics: true })] });
const link = (text, url) => new ExternalHyperlink({ link: url, children: [t(text, { underline: {}, color: "000000" })] });

function table(rows, widths, opts = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((r, i) => new TableRow({
      tableHeader: i === 0, cantSplit: true,
      children: r.map((c, j) => new TableCell({
        borders, width: { size: widths[j], type: WidthType.DXA },
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        children: [new Paragraph({
          alignment: j > 0 && opts.right && i > 0 && /^[-+\d,.%]+$/.test(String(c)) ? AlignmentType.RIGHT : AlignmentType.LEFT,
          children: [t(String(c), { bold: i === 0 || (opts.boldFirstCol && j === 0), size: 20 })],
        })],
      })),
    })),
  });
}
const thick = { style: BorderStyle.SINGLE, size: 18, color: "000000" };
function callout(lines) {
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ children: [new TableCell({ width: { size: W, type: WidthType.DXA },
      borders: { top: thick, bottom: thick, left: thick, right: thick },
      margins: { top: 160, bottom: 160, left: 220, right: 220 }, children: lines })] })] });
}
const hl = (text, o = {}) => new TextRun({ text, font: FONT, bold: true, highlight: "yellow", ...o });
const hlink = (url, size = 24) => new ExternalHyperlink({ link: url, children: [new TextRun({ text: url, font: FONT, bold: true, underline: {}, color: "0000EE", highlight: "yellow", size })] });
const pageBox = () => callout([
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [hl("REVIEWER: PLEASE OPEN THE INTERACTIVE PLANNER", { size: 26 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [hlink(PAGE, 26)] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [t("The page solves the routing LP live in your browser: change rain level, road flooding, demand and penalty, and click a road to place the pump. All project files: ", { size: 20 }), link(REPO, REPO)] }),
]);
const gap = () => new Paragraph({ spacing: { after: 120 }, children: [] });
const caption = (s) => new Paragraph({ spacing: { before: 60, after: 200 }, alignment: AlignmentType.CENTER,
  children: [t(s, { italics: true, size: 18 })] });
function image(file, wpx, hpx) {
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120 },
    children: [new ImageRun({ type: "png", data: fs.readFileSync(path.join(OUT, file)),
      transformation: { width: wpx, height: hpx },
      altText: { title: file, description: file, name: file } })] });
}

const body = [];
const add = (...x) => body.push(...x);

// ---------------- Title page ----------------
add(
  new Paragraph({ spacing: { before: 2400, after: 200 }, alignment: AlignmentType.CENTER,
    children: [t("Rain-Ready Chennai", { bold: true, size: 52 })] }),
  new Paragraph({ spacing: { after: 600 }, alignment: AlignmentType.CENTER,
    children: [t("Planning relief supplies and evacuation for the North-East Monsoon using Markov chains, linear programming and network models", { size: 26 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [t("Decision Science: Creative AI Project", { bold: true })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [t("Shivnath Rajendra Gujar | Roll No. FT271017 | Section 1")] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [t("PGPM 2026-27, Great Lakes Institute of Management, Chennai")] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 900 }, children: [t("November 2026")] }),
  pageBox(),
  new Paragraph({ children: [new PageBreak()] }),
);

// ---------------- 1. Summary ----------------
add(H1("1. Summary"),
  P("Every year between October and December, heavy rain floods low-lying parts of Chennai. Relief kits (food and water) must reach relief camps on roads that may themselves be under water. This project builds a set of models that help a relief control room decide, each evening, what to stock, how to move it and how to get people out."),
  P([B("Main results")]),
  bullet([B("Roads, not supply, cause shortages. "), t("On a very heavy rain day, 2,225 kits cannot reach Velachery and Mudichur even though the warehouses still have 10,625 kits to spare.")]),
  bullet([B("Stock the road-bound camps the night before. "), t("On a moderate-rain day, stocking 700 kits at Velachery and 600 at Mudichur lowers expected cost by Rs 10,547 (9.7%) and cuts the shortfall on a very heavy day from 2,225 to 925 kits.")]),
  bullet([B("Act on the heavy day, not the very heavy day. "), t("Once rain is very heavy, the road into Mudichur is mostly flooded, so fewer kits can be moved in.")]),
  bullet([B("Choose evacuation roads by safety, not distance. "), t("On a very heavy day the shortest road out of Mudichur gets through only 13.5% of the time; a road 0.5 km longer gets through 81% of the time.")]),
  bullet([B("Put the pump at the exit. "), t("The bottleneck is the three roads leaving Mudichur. One pump on the Mudichur to Varadharajapuram road adds 180 vehicles per hour of evacuation capacity.")]),
  P("All network figures are illustrative and the rainfall series is synthetic, so the rupee values show how the method works rather than real costs. The Excel and Python templates accept real data."),
  pageBox(),
  gap(),
);

// ---------------- 2. Problem ----------------
add(H1("2. The problem"),
  P("A relief control room in Chennai has to take three decisions every evening of the monsoon:"),
  num([B("Stock: "), t("how many kits to place at which hubs and camps tonight, before knowing tomorrow's rain.")]),
  num([B("Route: "), t("once tomorrow's rain is known, how to move kits from warehouses through hubs to camps at the lowest cost on the roads that are still open.")]),
  num([B("Evacuate: "), t("which road gives residents of a flooding area the best chance of reaching a shelter, and where a single dewatering pump helps most.")]),
  P("The project uses four Decision Science models. Each one answers one question and passes its result to the next."),
  table([
    ["Step", "Model", "Question it answers", "Used by"],
    ["1", "Markov chain", "How likely is each rain level tomorrow, and how long does a wet spell last?", "Step 3 (scenario probabilities)"],
    ["2", "Transshipment LP", "For each rain level, what is the cheapest routing, and where do roads cause shortages?", "Step 3 (shows which camps need stock)"],
    ["3", "Two-stage stochastic LP", "What should be stocked tonight when tomorrow's rain is uncertain?", "Daily decision rule"],
    ["4", "Network models", "Which evacuation road is safest, and where is the bottleneck?", "Pump and rescue team placement"],
  ], [700, 2000, 4026, 2300]),
  gap(),
);

// ---------------- 3. Data ----------------
add(H1("3. Data and assumptions"),
  P("The relief network has 3 warehouses (sources), 3 hubs, 6 relief camps and 19 roads. Rain levels follow the India Meteorological Department (IMD) daily rainfall classes."),
  table([
    ["Rain level", "Daily rainfall", "Share of peak demand", "Transport cost multiplier"],
    ["D: Dry or light", "below 15.6 mm", "5%", "1.00"],
    ["M: Moderate", "15.6 to 64.4 mm", "25%", "1.10"],
    ["H: Heavy", "64.5 to 115.5 mm", "60%", "1.30"],
    ["V: Very heavy or extreme", "115.6 mm and above", "100%", "1.60"],
  ], [2600, 2200, 2100, 2126]),
  gap(),
  table([
    ["Node type", "Nodes and daily limit (kits)"],
    ["Warehouses", "Koyambedu 9,000; Madhavaram 6,000; Chengalpattu 7,000"],
    ["Hubs", "Guindy 10,000; Tambaram 8,000; Sholinganallur 6,000"],
    ["Camps (need on a very heavy day)", "Velachery 3,200; Pallikaranai 2,600; Mudichur 2,400; Saidapet 1,800; Perumbakkam 2,000; Vyasarpadi 1,600 (total 13,600)"],
  ], [2600, 6426]),
  gap(),
  P("Each road has a flood risk class. On any day, the usable capacity of a road is its normal capacity multiplied by (1 minus the chance it is under water)."),
  table([
    ["Flood risk class", "D", "M", "H", "V"],
    ["Low", "0%", "2%", "5%", "10%"],
    ["Medium", "0%", "5%", "20%", "40%"],
    ["High", "0%", "10%", "40%", "75%"],
  ], [2626, 1600, 1600, 1600, 1600]),
  gap(),
  P([B("Costs: "), t("transport Rs 3 to 16 per kit depending on the road; Rs 500 penalty for each kit that cannot be delivered; Rs 3 to hold a stocked kit; Rs 30 for a stocked kit that is not used. A camp can store up to 1,500 kits.")]),
  P([B("Data note: "), t("road costs, capacities and flood chances are illustrative values chosen for this project. The rainfall series is synthetic: 5 seasons of 92 days (460 days), generated in Python. Both can be replaced with real data in the templates.")]),
  image("fig_network.png", 560, 350),
  caption("Figure 1: The relief network. Squares are warehouses, diamonds are hubs and circles are camps."),
);

// ---------------- 4. Markov ----------------
add(H1("4. Model 1: Markov chain of rain levels"),
  P([B("Result: "), t("the monsoon has a short memory. Two days ahead, the forecast is already close to the long-run average, so plans should look one day ahead and be revised every evening.")]),
  H2("Method"),
  P("Each day is put into one of the four rain levels. Counting how often one level is followed by another gives the transition matrix P. Each row adds up to 1."),
  table([
    ["Today \\ Tomorrow", "D", "M", "H", "V"],
    ["D", "0.761", "0.170", "0.045", "0.024"],
    ["M", "0.554", "0.293", "0.087", "0.065"],
    ["H", "0.286", "0.238", "0.262", "0.214"],
    ["V", "0.156", "0.219", "0.312", "0.312"],
  ], [2626, 1600, 1600, 1600, 1600], { right: true, boldFirstCol: true }),
  gap(),
  P("The forecast for n days ahead is found by multiplying today's state by P, n times:"),
  formula("Forecast(day n+1) = Forecast(day n) x P"),
  H2("Results"),
  table([
    ["Measure", "D", "M", "H", "V"],
    ["Tomorrow, if today is M", "0.554", "0.293", "0.087", "0.065"],
    ["Day after tomorrow, if today is M", "0.620", "0.215", "0.094", "0.072"],
    ["Long run (share of all days)", "0.632", "0.205", "0.093", "0.071"],
    ["Average length of a spell (days)", "4.19", "1.42", "1.35", "1.45"],
  ], [3426, 1400, 1400, 1400, 1400], { right: true }),
  gap(),
  image("fig_forecast.png", 520, 268),
  caption("Figure 2: Forecast starting from a moderate-rain day. It settles within about three days."),
  P("To estimate how long until roads clear, the model finds the average number of days until the next dry day. With Q as the part of P for the wet levels M, H and V:"),
  formula("Mean days to a dry day = (I - Q)^-1 x 1"),
  table([
    ["Starting rain level", "Average days until a dry day"],
    ["M: Moderate", "2.12"],
    ["H: Heavy", "3.06"],
    ["V: Very heavy", "3.52"],
  ], [4513, 4513], { right: true }),
  gap(),
  P("In practice, once a very heavy day arrives, camps should plan for about 3.5 days before roads can be expected to clear. A very heavy day is followed by another very heavy day 31% of the time, compared with 2.4% after a dry day."),
);

// ---------------- 5. Transshipment ----------------
add(H1("5. Model 2: Transshipment LP (routing playbook)"),
  P([B("Result: "), t("on a very heavy day, 2,225 kits (16% of need) cannot reach Velachery and Mudichur, while the warehouses still have 10,625 kits to spare. The shortage is caused by flooded roads, not by a lack of supply.")]),
  H2("Method"),
  P("For each rain level, the LP chooses how many kits to send on each road to keep total cost as low as possible:"),
  formula("Minimise: sum of (cost per kit x kits on each road) + 500 x (kits not delivered)"),
  P("Subject to:"),
  bullet("Each warehouse sends no more than its supply."),
  bullet("At each hub, kits in = kits out (the transshipment condition), and kits handled stay within the hub limit."),
  bullet("At each camp, kits delivered + kits not delivered = kits needed."),
  bullet("Kits on each road stay within that road's usable capacity for the rain level."),
  P("Allowing kits to go undelivered, at a penalty, keeps the model solvable when roads fail and shows exactly where the network breaks."),
  H2("Results"),
  table([
    ["Rain level", "Kits needed", "Kits short", "Transport cost (Rs)", "Total cost with penalty (Rs)"],
    ["D", "680", "0", "6,510", "6,510"],
    ["M", "3,400", "0", "35,805", "35,805"],
    ["H", "8,160", "0", "1,01,634", "1,01,634"],
    ["V", "13,600", "2,225", "2,01,680", "13,14,180"],
  ], [1600, 1600, 1600, 2000, 2226], { right: true }),
  gap(),
  bullet([B("D, M and H use the same routes, only with more kits. "), t("Koyambedu supplies Guindy and Saidapet, Madhavaram supplies Vyasarpadi and Chengalpattu supplies Tambaram. Only 60 kits change route in H.")]),
  bullet([B("On a very heavy day the network changes. "), t("The Sholinganallur hub, unused until now, carries 2,175 kits from Chengalpattu on low-risk roads. Chengalpattu to Tambaram rises from 2,700 to 4,175 kits.")]),
  bullet([B("Where the shortfall is: "), t("Velachery is 700 kits short (its two roads carry at most 1,000 + 1,500 = 2,500 kits against 3,200 needed). Mudichur is 1,525 kits short (its only road, from Tambaram, is 75% blocked and carries 875 kits against 2,400 needed).")]),
  bullet([B("Shadow prices confirm it. "), t("One more kit of need costs Rs 500 at Velachery and Mudichur (it simply goes undelivered) but only Rs 6 to 26 at the other camps.")]),
  P("No routing plan made on the day can fully serve Velachery and Mudichur. Kits have to be there before the rain peaks. Model 3 decides how many."),
);

// ---------------- 6. Stochastic ----------------
add(H1("6. Model 3: What to stock tonight (two-stage stochastic LP)"),
  P([B("Result: "), t("on a moderate-rain day, stock 700 kits at Velachery and 600 at Mudichur tonight. This lowers expected cost by Rs 10,547 (9.7%) compared with planning for an average day.")]),
  H2("Method"),
  P([B("Stage 1 (tonight, rain level known): "), t("send kits to hubs or camps and hold them there.")]),
  P([B("Stage 2 (tomorrow): "), t("one of four rain levels happens, with probabilities from today's row of the Markov matrix. For each one, the model routes kits again on that day's roads and counts penalties for kits not delivered and for stocked kits not used.")]),
  P("The model finds the stage 1 plan that gives the lowest total expected cost across all four possible tomorrows:"),
  formula("Minimise: stage 1 cost + sum over tomorrows of [probability x stage 2 cost]"),
  H2("Results when today is moderate"),
  P("Chances for tomorrow: D 55.4%, M 29.3%, H 8.7%, V 6.5%."),
  table([
    ["Plan", "Expected cost (Rs)", "Kits short if tomorrow is V"],
    ["Stochastic plan (stock 1,300 kits)", "98,115", "925"],
    ["Plan for the average day", "1,08,662", "2,225"],
    ["Plan for the most likely day (dry)", "1,08,662", "2,225"],
    ["Do not stock", "1,08,662", "2,225"],
    ["Perfect forecast (wait and see)", "37,871", "-"],
  ], [4426, 2300, 2300], { right: true }),
  gap(),
  bullet([B("Value of the stochastic solution (VSS) = Rs 1,08,662 - Rs 98,115 = Rs 10,547. "), t("Planning for the average day or the most likely day both say \"stock nothing\". Only by looking at all scenarios does the model see that a 6.5% chance of a very heavy day is worth preparing for.")]),
  bullet([B("Expected value of perfect information (EVPI) = Rs 98,115 - Rs 37,871 = Rs 60,244 per day. "), t("This is the most a perfect one-day forecast would be worth, and so the upper limit on what is worth spending on better local forecasting.")]),
  H2("Decision rule for each rain level today"),
  table([
    ["Today", "Kits to stock", "Where", "Expected cost (Rs)", "Without stocking (Rs)", "VSS (Rs)", "EVPI (Rs)"],
    ["M", "1,300", "Velachery 700, Mudichur 600", "98,115", "1,08,662", "10,547", "60,244"],
    ["H", "2,200", "Velachery 700, Mudichur 1,500", "1,36,565", "3,18,613", "1,82,048", "47,507"],
    ["V", "1,575", "Velachery 700, Mudichur 875", "2,37,759", "4,51,291", "2,13,533", "22,477"],
  ], [800, 1000, 2026, 1300, 1400, 1250, 1250], { right: true }),
  gap(),
  bullet([B("All stock goes to camps, none to hubs. "), t("Hubs can still be reached tomorrow; the two flood-prone camps cannot.")]),
  bullet([B("Act on the heavy day. "), t("The plan stocks less when today is already very heavy, because the road into Mudichur is 75% blocked today as well and only 875 kits can get in.")]),
  H2("Sensitivity to the penalty"),
  table([
    ["Penalty per kit not delivered (Rs)", "Kits stocked", "Expected cost (Rs)", "VSS (Rs)", "Kits short if V"],
    ["100", "0", "50,618", "0", "2,225"],
    ["250", "280", "69,807", "2,578", "1,945"],
    ["500", "1,300", "98,115", "10,547", "925"],
    ["1,000", "2,200", "1,06,981", "74,235", "25"],
    ["2,000", "2,200", "1,08,611", "2,17,714", "25"],
  ], [2626, 1500, 1700, 1600, 1600], { right: true }),
  gap(),
  P("The penalty is a judgement about what a missed meal is worth. From Rs 1,000 upwards the plan stops changing: Mudichur is filled to its 1,500-kit storage limit. At that point storage space, not cost, is the limit, which supports adding temporary storage at Mudichur."),
);

// ---------------- 7. Evacuation ----------------
add(H1("7. Model 4: Evacuating Mudichur (network models)"),
  P([B("Result: "), t("on a very heavy day, the shortest route out of Mudichur gets through only 13.5% of the time, while a route 0.5 km longer gets through 81% of the time. The bottleneck is the three roads leaving Mudichur, so the one pump should go on the Mudichur to Varadharajapuram road.")]),
  H2("Most reliable route"),
  P("The evacuation network has 8 locations, 11 roads and two shelters on higher ground (Vandalur and Pallavaram). Each road has a reliability r, the chance it is passable. The chance that a whole route works is the product of its roads' reliabilities. Taking logarithms turns this into a normal shortest-path problem:"),
  formula("Maximise product of r  =  Minimise sum of (-ln r)"),
  table([
    ["Rain level", "Route", "Distance (km)", "Chance of getting through"],
    ["H", "Shortest: Mudichur - Varadharajapuram - Perungalathur - Vandalur", "8.0", "45.6%"],
    ["H", "Safest: Mudichur - Perungalathur - Vandalur", "8.5", "90.2%"],
    ["V", "Shortest: Mudichur - Varadharajapuram - Perungalathur - Vandalur", "8.0", "13.5%"],
    ["V", "Safest: Mudichur - Perungalathur - Vandalur", "8.5", "81.0%"],
  ], [1200, 4626, 1400, 1800], { right: true }),
  gap(),
  P("A navigation app that picks the shortest distance would send people down the riskiest road."),
  H2("Maximum flow and bottleneck"),
  P("Using the expected capacity of each road (vehicles per hour x reliability), the maximum evacuation flow is 1,005 vehicles per hour on a heavy day and 730 on a very heavy day. In both cases the minimum cut is the three roads leaving Mudichur: roads further out and the shelters have spare capacity."),
  table([
    ["Road given the pump (very heavy day)", "Current risk", "New maximum flow (vehicles/hour)", "Gain"],
    ["Mudichur to Varadharajapuram", "High", "910", "+180"],
    ["Mudichur to Tambaram West", "Medium", "870", "+140"],
    ["Any of the 4 flood-prone roads further out", "Medium or high", "730", "0"],
  ], [3626, 1600, 2400, 1400], { right: true }),
  gap(),
  P("Pumping a road that is not on the minimum cut adds nothing, however flooded it looks."),
);

// ---------------- 8. Recommendations ----------------
add(H1("8. Recommendations"),
  num([B("Check today's rain level every evening "), t("and use tomorrow's probabilities from the Markov matrix. Revise the plan daily.")], "rec"),
  num([B("Stock the camps that roads cannot reach, not the hubs. "), t("Velachery and Mudichur are the only camps where the shadow price reaches the penalty.")], "rec"),
  num([B("Act on the heavy day. "), t("Stock 2,200 kits (Velachery 700, Mudichur 1,500) when today is heavy. Waiting until very heavy rain leaves room for only 875 kits at Mudichur.")], "rec"),
  num([B("Do not plan for the average day or the most likely day. "), t("Both say stock nothing, which costs Rs 10,547 more on a moderate day and Rs 1,82,048 more on a heavy day.")], "rec"),
  num([B("Add storage at Mudichur. "), t("At higher penalties, its 1,500-kit storage is the limiting factor.")], "rec"),
  num([B("Spend on better forecasting only up to the EVPI, "), t("about Rs 60,244 per moderate day in this network.")], "rec"),
  num([B("Guide evacuation by safety, not distance, and place the pump on the Mudichur to Varadharajapuram road.")], "rec"),
);

// ---------------- 9. Limitations ----------------
add(H1("9. Limitations"),
  bullet([B("Illustrative data. "), t("The rainfall is synthetic and the network values are assumptions. The method and the pattern of the results carry over; the rupee figures do not.")]),
  bullet([B("One rain level for the whole city. "), t("North and south Chennai can have different rain on the same day. A Markov chain per zone would handle this.")]),
  bullet([B("Tomorrow depends only on today. "), t("This first-order Markov assumption should be tested on real IMD data.")]),
  bullet([B("Two stages only. "), t("A model that plans several days ahead would capture 3 to 4 day wet spells better.")]),
  bullet([B("Average road capacity. "), t("Roads are treated as partly open. Simulating roads as fully open or fully closed would show the risk in more detail.")]),
  bullet([B("The penalty is a judgement. "), t("The sensitivity table shows its effect; the final value should be agreed with the relief authority.")]),
);

// ---------------- 10. AI ----------------
add(H1("10. How AI was used"),
  P("The project was built with Claude (Anthropic) as an AI assistant. All outputs were checked by hand and against each other."),
  table([
    ["Stage", "What the AI did", "What was checked"],
    ["Problem choice", "Suggested a live local problem (monsoon relief in Chennai) that links four models", "That each model supports a real decision"],
    ["Data", "Generated synthetic rainfall and an illustrative network; used IMD rainfall classes", "All assumed values labelled and kept editable"],
    ["Models", "Wrote the Python models (PuLP, networkx) and the in-browser optimiser", "Bottlenecks checked by hand: Velachery 2,500 of 3,200; Mudichur 3,500 x 25% = 875 of 2,400"],
    ["Excel", "Built the formula-based workbook with Solver models stored in it", "Every Excel result matched Python; the web page's optimiser gives the same results"],
    ["Report and slides", "Drafted text, tables and charts", "Figures re-checked; one error corrected (spare supply is 10,625 kits, not 8,625)"],
  ], [1700, 3826, 3500]),
  gap(),
);

// ---------------- 11. Files ----------------
add(H1("11. Project files and links"),
  P([hl("Interactive web page (live): "), hlink(PAGE, 22)]),
  P("The page runs the routing optimisation live in the browser. You can change the rain level, road flooding, demand and penalty and watch the flows and shortages update. You can also click a road on the evacuation map to place the pump and see the new routes and capacity."),
  P([B("GitHub repository: "), link(REPO, REPO)]),
  table([
    ["File", "What it contains"],
    ["Rain_Ready_Chennai_Report.docx", "This report"],
    ["Rain_Ready_Chennai_Decision_Science.xlsx", "Excel template: inputs, Markov chain, Solver-ready transshipment LP, results, Solver-ready safest-route LP"],
    ["rain_ready_chennai.py", "Python model that runs all four models and saves the results"],
    ["index.html", "The interactive web page"],
    ["Rain_Ready_Chennai_Slides.pptx", "Presentation"],
  ], [3800, 5226]),
  gap(),
);

// ---------------- Appendix ----------------
add(H1("Appendix: Using the templates"),
  H2("Excel"),
  table([
    ["Sheet", "Purpose", "What to change"],
    ["Inputs", "All parameters and the 19 roads", "Any value"],
    ["Rain_Data", "Classifies each day and links it to the next day", "Paste real daily rainfall into columns A to C"],
    ["Markov", "Transition counts, matrix P, 30-day forecast, long-run shares, days to a dry day", "Today's rain level in cell B25"],
    ["Transship_LP", "Routing LP for one rain level", "Rain level in B3, then Data > Solver > Solve"],
    ["Playbook", "Routes for all four rain levels (from Python)", "Re-run Python to refresh"],
    ["Prepositioning", "Stocking results, decision rule and sensitivity (from Python)", "Re-run Python to refresh"],
    ["Evacuation", "Safest-route LP and pump results", "Rain level in B3, then Solver"],
  ], [2000, 4026, 3000]),
  gap(),
  P("Solver settings for the routing LP (already stored in the file): minimise B53 by changing I7:I25 and C42:C47, subject to I7:I25 <= H7:H25, B30:B32 <= D30:D32, D36:D38 = 0, B36:B38 <= E36:E38 and D42:D47 = E42:E47. Tick \"Make unconstrained variables non-negative\" and choose Simplex LP."),
  H2("Python"),
  P("Install the libraries and run the script:"),
  P([t("pip install \"pulp>=2.7,<3\" networkx numpy matplotlib", { font: "Courier New", size: 20 })]),
  P([t("python rain_ready_chennai.py", { font: "Courier New", size: 20 })]),
  P("The script prints every result table in this report and saves results.json, rain_data.csv and the two figures in an outputs folder. To use real data, edit the DATA section at the top of the file."),
);

const doc = new Document({
  creator: "Shivnath Rajendra Gujar",
  title: "Rain-Ready Chennai",
  styles: {
    default: { document: { run: { font: FONT, size: 22, color: "000000" } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 30, bold: true, color: "000000" },
        paragraph: { spacing: { before: 360, after: 160 }, keepNext: true, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 24, bold: true, color: "000000" },
        paragraph: { spacing: { before: 220, after: 100 }, keepNext: true, outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
      { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
      { reference: "rec", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [t("Rain-Ready Chennai | Page ", { size: 18 }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18 })] })] }) },
    children: body,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(__dirname, "Rain_Ready_Chennai_Report.docx");
  fs.writeFileSync(out, buf);
  console.log("saved", out);
});
