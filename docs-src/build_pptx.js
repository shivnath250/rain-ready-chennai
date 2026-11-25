const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const md = require("react-icons/md");
const path = require("path");
const fs = require("fs");
const SK = "/root/.claude/skills/synced/0a3040a7-a559-47e4-8886-debde45bc76a_b0407df3-85d2-4a42-a4d2-ea17845e3c65/pptx";
const { applyTheme } = require(SK + "/scripts/apply_theme.js");

const REPO = process.argv[2] || "https://github.com/shivnath250/rain-ready-chennai";
const PAGE = process.argv[3] || "https://shivnath250.github.io/rain-ready-chennai/";
const OUT = path.join(__dirname, "..", "outputs");

const HEX = { navy: "0B2441", teal: "0D6E78", rain: "3F86B5", mist: "E6EEF3", ink: "10212F", red: "C23B2E", grey: "5B6B78", pale: "F4F7F9" };
const THEME = {
  name: "Rain Ready",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: { dk1: HEX.ink, lt1: "FFFFFF", dk2: HEX.navy, lt2: HEX.pale, accent1: HEX.teal, accent2: HEX.rain,
    accent3: HEX.red, accent4: "8FB3C9", accent5: "C48A10", accent6: "2F7D4F", hlink: HEX.teal, folHlink: HEX.grey },
};

async function icon(Comp, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.title = "Rain-Ready Chennai";
  pres.author = "Shivnath Rajendra Gujar";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  pres.defineSlideMaster({
    title: "Title Dark", background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.35, w: 8.8, h: 1.1, fontSize: 44, bold: true, color: C.background1, valign: "bottom", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 2.55, w: 8.4, h: 1.0, fontSize: 18, color: "CADCFC", valign: "top", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "Content", background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9.0, h: 0.65, fontSize: 24, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
      { text: { text: "Rain-Ready Chennai  |  Decision Science", options: { x: 0.5, y: 5.25, w: 5, h: 0.25, fontSize: 9, color: C.accent4, margin: 0 } } },
    ],
    slideNumber: { x: 9.0, y: 5.25, w: 0.5, h: 0.25, fontSize: 9, color: C.accent4, align: "right" },
  });
  pres.defineSlideMaster({
    title: "Closing Dark", background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9.0, h: 0.65, fontSize: 24, bold: true, color: C.background1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    ],
    slideNumber: { x: 9.0, y: 5.25, w: 0.5, h: 0.25, fontSize: 9, color: "8FB3C9", align: "right" },
  });

  const sub = (slide, text, y = 1.02) => slide.addText(text, { x: 0.5, y, w: 9, h: 0.4, fontSize: 14, color: C.text1, isTextBox: true, margin: 0, objectName: "Subtitle" });
  const circleIcon = async (slide, Comp, x, y, d, fill, name) => {
    slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill }, objectName: name + " circle" });
    slide.addImage({ data: await icon(Comp, "FFFFFF"), x: x + d * 0.22, y: y + d * 0.22, w: d * 0.56, h: d * 0.56, objectName: name + " icon" });
  };
  const stat = (slide, x, y, w, big, label, color, name, fs = 40) => {
    slide.addText(big, { x, y, w, h: 0.75, fontSize: fs, bold: true, color, fontFace: "Cambria", margin: 0, isTextBox: true, objectName: name + " value" });
    slide.addText(label, { x, y: y + 0.75, w, h: 0.6, fontSize: 12, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: name + " label" });
  };
  const card = (slide, x, y, w, h, name, fill = C.background2) =>
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill }, objectName: name });
  const linkButton = (slide, x, y, w, h, fill, txtColor, name, size = 14) => {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.1, fill: { color: fill }, line: { color: fill },
      hyperlink: { url: PAGE, tooltip: "Open the interactive planner" }, objectName: name + " button" });
    slide.addText([
      { text: "\u25B6  OPEN THE INTERACTIVE PLANNER", options: { bold: true, fontSize: size, color: txtColor, breakLine: true } },
      { text: PAGE.replace("https://", ""), options: { fontSize: size - 2, color: txtColor, underline: { style: "sng" }, hyperlink: { url: PAGE } } },
    ], { x: x + 0.2, y, w: w - 0.4, h, margin: 0, valign: "middle", isTextBox: true, objectName: name + " text" });
  };
  const chartText = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt",
    catAxisLabelFontSize: 11, valAxisLabelFontSize: 10, dataLabelFontSize: 10, legendFontSize: 11, catAxisLabelColor: HEX.ink, valAxisLabelColor: HEX.grey };

  // ---------- 1. Title ----------
  pres.addSection({ title: "Introduction" });
  let s = pres.addSlide({ masterName: "Title Dark", sectionTitle: "Introduction" });
  s.addText("Rain-Ready Chennai", { placeholder: "title" });
  s.addText("Planning relief supplies and evacuation for the North-East Monsoon with Markov chains, linear programming and network models", { placeholder: "body" });
  await circleIcon(s, md.MdWaterDrop, 0.6, 0.45, 0.7, C.accent1, "Drop");
  s.addText([
    { text: "Shivnath Rajendra Gujar  |  FT271017  |  PGPM 2026-27, Great Lakes Institute of Management", options: { breakLine: true } },
    { text: "Decision Science: Creative AI Project" },
  ], { x: 0.6, y: 4.55, w: 5.0, h: 0.6, fontSize: 12, color: "CADCFC", margin: 0, isTextBox: true, valign: "top", paraSpaceAfter: 3, objectName: "Byline" });
  linkButton(s, 5.75, 4.3, 3.75, 0.85, "FFD24A", HEX.navy, "Title link", 13);
  s.addNotes("Every monsoon, low-lying Chennai floods. This project builds four linked models to help a relief control room decide what to stock, how to route it, and how to evacuate.");

  // ---------- 2. Problem ----------
  pres.addSection({ title: "Problem and approach" });
  s = pres.addSlide({ masterName: "Content", sectionTitle: "Problem and approach" });
  s.addText("Three decisions every monsoon evening", { placeholder: "title" });
  sub(s, "Velachery, Pallikaranai, Mudichur and other low-lying areas flood. The roads into them flood first.");
  const dec = [
    [md.MdInventory2, "Stock", "How many kits to place at which camps tonight, before knowing tomorrow's rain?"],
    [md.MdLocalShipping, "Route", "How to move kits from warehouses through hubs to camps on the roads still open?"],
    [md.MdDirectionsRun, "Evacuate", "Which road gets residents to a shelter, and where does one pump help most?"],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.5 + i * 3.08;
    card(s, x, 1.75, 2.85, 3.1, "Decision card " + (i + 1));
    await circleIcon(s, dec[i][0], x + 0.3, 2.0, 0.75, C.accent1, dec[i][1]);
    s.addText(dec[i][1], { x: x + 0.3, y: 2.95, w: 2.3, h: 0.45, fontSize: 22, bold: true, color: C.text2, fontFace: "Cambria", margin: 0, isTextBox: true, objectName: dec[i][1] + " title" });
    s.addText(dec[i][2], { x: x + 0.3, y: 3.45, w: 2.3, h: 1.2, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: dec[i][1] + " text" });
  }

  // ---------- 3. Approach ----------
  s = pres.addSlide({ masterName: "Content", sectionTitle: "Problem and approach" });
  s.addText("Four models, each feeding the next", { placeholder: "title" });
  const steps = [
    [md.MdCloud, "Markov chain", "Chance of each rain level tomorrow; length of wet spells"],
    [md.MdAltRoute, "Transshipment LP", "Cheapest routing for each rain level; where roads fail"],
    [md.MdShield, "Stochastic LP", "What to stock tonight across all possible tomorrows"],
    [md.MdHub, "Network models", "Safest evacuation road; bottleneck and pump location"],
  ];
  for (let i = 0; i < 4; i++) {
    const x = 0.5 + i * 2.3;
    card(s, x, 1.45, 1.95, 3.2, "Step card " + (i + 1));
    s.addText(String(i + 1), { x: x + 0.2, y: 1.6, w: 0.6, h: 0.6, fontSize: 32, bold: true, color: C.accent2, fontFace: "Cambria", margin: 0, isTextBox: true, objectName: "Step number " + (i + 1) });
    await circleIcon(s, steps[i][0], x + 1.2, 1.65, 0.55, C.accent1, "Step " + (i + 1));
    s.addText(steps[i][1], { x: x + 0.2, y: 2.45, w: 1.6, h: 0.7, fontSize: 16, bold: true, color: C.text2, margin: 0, valign: "top", isTextBox: true, objectName: "Step title " + (i + 1) });
    s.addText(steps[i][2], { x: x + 0.2, y: 3.2, w: 1.6, h: 1.3, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "Step text " + (i + 1) });
    if (i < 3) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 2.0, y: 2.9, w: 0.25, h: 0.3, fill: { color: C.accent4 }, line: { color: C.accent4 }, objectName: "Arrow " + (i + 1) });
  }
  s.addText("Rain levels follow IMD daily classes: dry (<15.6 mm), moderate, heavy, very heavy (115.6 mm+). Network values are illustrative; rainfall is synthetic.",
    { x: 0.5, y: 4.8, w: 9, h: 0.35, fontSize: 11, italic: true, color: C.accent4, margin: 0, isTextBox: true, objectName: "Data note" });

  // ---------- 4. Network ----------
  s = pres.addSlide({ masterName: "Content", sectionTitle: "Problem and approach" });
  s.addText("The relief network: 3 warehouses, 3 hubs, 6 camps", { placeholder: "title" });
  s.addImage({ path: path.join(OUT, "fig_network.png"), x: 0.5, y: 1.15, w: 4.1, h: 3.85, sizing: { type: "contain", w: 4.1, h: 3.85 }, altText: "Relief network diagram", objectName: "Network diagram" });
  const opts = { fontSize: 12, color: HEX.ink, margin: 0.05 };
  const head = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: HEX.navy } } });
  s.addTable([
    [head("Rain level"), head("Need"), head("Low"), head("Med"), head("High")],
    ["Dry / light", "5%", "0%", "0%", "0%"],
    ["Moderate", "25%", "2%", "5%", "10%"],
    ["Heavy", "60%", "5%", "20%", "40%"],
    ["Very heavy", "100%", "10%", "40%", "75%"],
  ], { x: 5.0, y: 1.25, w: 4.5, colW: [1.5, 0.75, 0.75, 0.75, 0.75], ...opts, border: { type: "solid", pt: 0.5, color: "C9D3DB" }, align: "center", objectName: "Regime table" });
  s.addText([
    { text: "Need", options: { bold: true } }, { text: " = share of the 13,600-kit peak demand.", options: { breakLine: true } },
    { text: "Low / Med / High", options: { bold: true } }, { text: " = chance a road of that risk class is under water.", options: { breakLine: true } },
    { text: "Road capacity", options: { bold: true } }, { text: " = normal capacity x (1 - that chance)." },
  ], { x: 5.0, y: 3.2, w: 4.5, h: 1.2, fontSize: 13, color: C.text1, margin: 0, isTextBox: true, valign: "top", paraSpaceAfter: 6, objectName: "Table legend" });
  s.addText("Penalty Rs 500 per kit not delivered. Camp storage 1,500 kits.", { x: 5.0, y: 4.5, w: 4.5, h: 0.4, fontSize: 12, italic: true, color: C.accent1, margin: 0, isTextBox: true, objectName: "Cost note" });

  // ---------- 5. Markov ----------
  pres.addSection({ title: "Results" });
  s = pres.addSlide({ masterName: "Content", sectionTitle: "Results" });
  s.addText("The monsoon forgets fast: plan one day ahead", { placeholder: "title" });
  const R = JSON.parse(fs.readFileSync(path.join(OUT, "results.json")));
  const days = ["Today", "+1", "+2", "+3", "+4", "+5", "+6", "+7"];
  const names = ["Dry / light", "Moderate", "Heavy", "Very heavy"];
  s.addChart(pres.charts.LINE, names.map((n, j) => ({ name: n, labels: days, values: R.forecast.map((r) => +(r[j] * 100).toFixed(1)) })), {
    x: 0.4, y: 1.1, w: 5.8, h: 3.9, ...chartText, chartColors: ["8FB3C9", HEX.rain, HEX.teal, HEX.navy], lineSize: 2.5, lineDataSymbolSize: 6,
    showLegend: true, legendPos: "b", valAxisMaxVal: 100, valAxisMinVal: 0, valAxisMajorUnit: 25, valAxisLabelFormatCode: "0\"%\"",
    valGridLine: { color: "E2E8EE", size: 0.5 }, catGridLine: { style: "none" }, showTitle: true, title: "Forecast starting from a moderate day (% chance)", titleFontSize: 12, titleColor: HEX.ink,
    objectName: "Forecast chart",
  });
  stat(s, 6.6, 1.25, 3, "3.5 days", "average wait for a dry day after a very heavy day", C.accent1, "Days to dry");
  stat(s, 6.6, 2.55, 3, "31%", "chance a very heavy day is followed by another (2.4% after a dry day)", C.accent3, "Repeat chance");
  stat(s, 6.6, 3.85, 3, "~3 days", "until the forecast is no better than the long-run average", C.text2, "Memory");

  // ---------- 6. Transshipment ----------
  s = pres.addSlide({ masterName: "Content", sectionTitle: "Results" });
  s.addText("On a very heavy day, roads, not supply, cause the shortage", { placeholder: "title" });
  const camps = ["Velachery", "Pallikaranai", "Mudichur", "Saidapet", "Perumbakkam", "Vyasarpadi"];
  const peak = { Velachery: 3200, Pallikaranai: 2600, Mudichur: 2400, Saidapet: 1800, Perumbakkam: 2000, Vyasarpadi: 1600 };
  const un = R.playbook.V.unmet;
  s.addChart(pres.charts.BAR, [
    { name: "Delivered", labels: camps, values: camps.map((c) => Math.round(peak[c] - un[c])) },
    { name: "Short", labels: camps, values: camps.map((c) => Math.round(un[c])) },
  ], {
    x: 0.4, y: 1.1, w: 5.9, h: 3.9, barDir: "bar", barGrouping: "stacked", ...chartText, chartColors: [HEX.teal, HEX.red],
    showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", dataLabelFormatCode: "#,##0;;;", showLegend: true, legendPos: "b",
    valGridLine: { color: "E2E8EE", size: 0.5 }, catGridLine: { style: "none" }, catAxisOrientation: "maxMin",
    showTitle: true, title: "Kits per camp, very heavy day", titleFontSize: 12, titleColor: HEX.ink, objectName: "Shortfall chart",
  });
  stat(s, 6.6, 1.25, 3, "2,225", "kits cannot reach Velachery and Mudichur", C.accent3, "Short");
  stat(s, 6.6, 2.55, 3, "10,625", "kits still unused in the warehouses", C.accent1, "Spare");
  s.addText("Mudichur's only road is 75% blocked: it carries 875 kits against 2,400 needed. No routing plan on the day can fix this, so kits must be there before the rain peaks.",
    { x: 6.6, y: 3.85, w: 3, h: 1.2, fontSize: 12, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "Shortfall insight" });

  // ---------- 7. Stochastic ----------
  s = pres.addSlide({ masterName: "Content", sectionTitle: "Results" });
  s.addText("Stock the two road-bound camps the night before", { placeholder: "title" });
  sub(s, "Today moderate. Tomorrow: dry 55.4%, moderate 29.3%, heavy 8.7%, very heavy 6.5%.");
  const st = R.stochastic;
  s.addChart(pres.charts.BAR, [{ name: "Expected cost", labels: ["Stochastic plan", "Average-day plan", "Perfect forecast"],
    values: [Math.round(st.RP.objective), Math.round(st.EEV.objective), Math.round(st.WS)] }], {
    x: 0.4, y: 1.5, w: 5.4, h: 3.5, barDir: "col", ...chartText, chartColors: [HEX.teal, HEX.red, "9AA8B3"], invertedColors: [HEX.teal, HEX.red, "9AA8B3"],
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: HEX.ink, dataLabelFormatCode: "\"Rs \"#,##0", showLegend: false,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showTitle: true, title: "Expected cost of each plan (Rs)", titleFontSize: 12, titleColor: HEX.ink, objectName: "Plan cost chart",
  });
  card(s, 6.2, 1.55, 3.3, 1.35, "Stock card");
  s.addText([{ text: "Stock tonight", options: { bold: true, fontSize: 14, color: HEX.navy, breakLine: true } },
    { text: "Velachery 700  |  Mudichur 600", options: { fontSize: 18, bold: true, color: HEX.teal, breakLine: true } },
    { text: "Shortfall if very heavy: 2,225 to 925 kits", options: { fontSize: 12, color: HEX.ink } }],
    { x: 6.4, y: 1.65, w: 3.0, h: 1.15, margin: 0, isTextBox: true, valign: "top", paraSpaceAfter: 4, objectName: "Stock text" });
  stat(s, 6.2, 3.1, 1.6, "Rs 10.5k", "VSS: gain from planning with scenarios", C.accent1, "VSS", 26);
  stat(s, 7.9, 3.1, 1.6, "Rs 60k", "EVPI: most a perfect forecast is worth", C.text2, "EVPI", 26);
  s.addText("The average day and the most likely day both say \"stock nothing\".", { x: 6.2, y: 4.55, w: 3.3, h: 0.5, fontSize: 11, italic: true, color: C.text1, margin: 0, isTextBox: true, objectName: "Average-day note" });

  // ---------- 8. Policy ----------
  s = pres.addSlide({ masterName: "Content", sectionTitle: "Results" });
  s.addText("Act on the heavy day, not the very heavy day", { placeholder: "title" });
  sub(s, "Re-solving the stochastic LP for each possible 'today' gives a simple decision rule.");
  const pol = R.policy;
  const fm = (n) => Math.round(n).toLocaleString("en-IN");
  s.addTable([
    [head("Today"), head("Kits to stock"), head("Velachery"), head("Mudichur"), head("Expected cost (Rs)"), head("Without stocking (Rs)"), head("VSS (Rs)")],
    ...["M", "H", "V"].map((td) => {
      const p = pol[td]; const tot = Object.values(p.camp_stock).reduce((a, b) => a + b, 0);
      const hl = td === "H" ? { bold: true, fill: { color: "DDEFF0" } } : {};
      return [{ text: { M: "Moderate", H: "Heavy", V: "Very heavy" }[td], options: hl }, { text: fm(tot), options: hl }, { text: fm(p.camp_stock.Velachery), options: hl },
        { text: fm(p.camp_stock.Mudichur), options: hl }, { text: fm(p.RP), options: hl }, { text: fm(p.NONE), options: hl }, { text: fm(p.VSS), options: hl }];
    }),
  ], { x: 0.5, y: 1.6, w: 9, colW: [1.3, 1.1, 1.1, 1.1, 1.5, 1.6, 1.3], fontSize: 13, color: HEX.ink, align: "center", border: { type: "solid", pt: 0.5, color: "C9D3DB" }, rowH: 0.45, objectName: "Policy table" });
  const ins = [
    [md.MdTrendingUp, "Heavy day = peak stocking", "2,200 kits; Mudichur filled to its 1,500-kit storage limit"],
    [md.MdBlock, "Very heavy is too late", "Mudichur's road is already 75% blocked, so only 875 kits get in"],
    [md.MdStore, "Camps, never hubs", "Hubs stay reachable tomorrow; the flood-prone camps do not"],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.5 + i * 3.05;
    await circleIcon(s, ins[i][0], x, 3.65, 0.55, i === 1 ? C.accent3 : C.accent1, "Insight " + (i + 1));
    s.addText(ins[i][1], { x: x + 0.7, y: 3.6, w: 2.2, h: 0.35, fontSize: 14, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "Insight title " + (i + 1) });
    s.addText(ins[i][2], { x: x + 0.7, y: 3.97, w: 2.2, h: 0.9, fontSize: 12, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "Insight text " + (i + 1) });
  }

  // ---------- 9. Evacuation ----------
  s = pres.addSlide({ masterName: "Content", sectionTitle: "Results" });
  s.addText("Evacuate by safety, not distance; pump the exit road", { placeholder: "title" });
  sub(s, "Very heavy day, Mudichur to a shelter. Route reliability = product of road reliabilities, solved as min sum of -ln(r).");
  card(s, 0.5, 1.6, 2.9, 2.4, "Shortest card");
  s.addText([{ text: "Shortest route", options: { bold: true, fontSize: 14, color: HEX.navy, breakLine: true } },
    { text: "8.0 km via Varadharajapuram", options: { fontSize: 12, color: HEX.ink } }], { x: 0.7, y: 1.75, w: 2.5, h: 0.6, margin: 0, isTextBox: true, valign: "top", objectName: "Shortest text" });
  s.addText("13.5%", { x: 0.7, y: 2.5, w: 2.5, h: 0.9, fontSize: 48, bold: true, color: C.accent3, fontFace: "Cambria", margin: 0, isTextBox: true, objectName: "Shortest value" });
  s.addText("chance of getting through", { x: 0.7, y: 3.4, w: 2.5, h: 0.4, fontSize: 12, color: C.text1, margin: 0, isTextBox: true, objectName: "Shortest label" });
  card(s, 3.6, 1.6, 2.9, 2.4, "Safest card", "DDEFF0");
  s.addText([{ text: "Safest route", options: { bold: true, fontSize: 14, color: HEX.navy, breakLine: true } },
    { text: "8.5 km via Perungalathur", options: { fontSize: 12, color: HEX.ink } }], { x: 3.8, y: 1.75, w: 2.5, h: 0.6, margin: 0, isTextBox: true, valign: "top", objectName: "Safest text" });
  s.addText("81.0%", { x: 3.8, y: 2.5, w: 2.5, h: 0.9, fontSize: 48, bold: true, color: C.accent1, fontFace: "Cambria", margin: 0, isTextBox: true, objectName: "Safest value" });
  s.addText("chance of getting through", { x: 3.8, y: 3.4, w: 2.5, h: 0.4, fontSize: 12, color: C.text1, margin: 0, isTextBox: true, objectName: "Safest label" });
  await circleIcon(s, md.MdWaterDamage, 6.85, 1.6, 0.6, C.accent1, "Pump");
  s.addText("Max-flow / min-cut", { x: 7.6, y: 1.65, w: 2, h: 0.5, fontSize: 14, bold: true, color: C.text2, margin: 0, isTextBox: true, valign: "middle", objectName: "Maxflow title" });
  s.addText([
    { text: "730 vehicles/hour can leave on a very heavy day.", options: { bullet: true, breakLine: true } },
    { text: "The bottleneck (min cut) is the three roads leaving Mudichur.", options: { bullet: true, breakLine: true } },
    { text: "One pump on Mudichur to Varadharajapuram adds 180 vehicles/hour; a pump further out adds 0.", options: { bullet: true } },
  ], { x: 6.85, y: 2.35, w: 2.7, h: 2.0, fontSize: 12, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 6, isTextBox: true, objectName: "Maxflow bullets" });
  s.addText("A navigation app that minimises distance sends residents down the riskiest road: 0.5 km longer gives 6x the odds.", { x: 0.5, y: 4.3, w: 6, h: 0.6, fontSize: 13, italic: true, color: C.text2, margin: 0, isTextBox: true, objectName: "Evac insight" });

  // ---------- 10. Interactive page ----------
  pres.addSection({ title: "Tools" });
  s = pres.addSlide({ masterName: "Content", sectionTitle: "Tools" });
  s.addText("Try it: the interactive planner re-solves the LP live", { placeholder: "title" });
  s.addImage({ path: "/tmp/route.png", x: 0.5, y: 1.15, w: 5.6, h: 3.95, sizing: { type: "crop", w: 5.6, h: 3.95, x: 0, y: 0 }, altText: "Screenshot of the interactive page", objectName: "Page screenshot" });
  s.addText([
    { text: "Change rain level, road flooding, demand and penalty; flows and shortages update instantly", options: { bullet: true, breakLine: true } },
    { text: "Click a road on the evacuation map to place the pump", options: { bullet: true, breakLine: true } },
    { text: "Same optimum as the Python and Excel models", options: { bullet: true } },
  ], { x: 6.4, y: 1.2, w: 3.1, h: 2.1, fontSize: 13, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 8, isTextBox: true, objectName: "Page bullets" });
  linkButton(s, 6.4, 3.45, 3.1, 0.95, "FFD24A", HEX.navy, "Page link", 12);
  s.addText([{ text: "All files (report, Excel, Python, slides): ", options: { bold: true, color: HEX.navy } },
    { text: REPO.replace("https://", ""), options: { hyperlink: { url: REPO }, color: HEX.teal } }],
    { x: 6.4, y: 4.5, w: 3.1, h: 0.6, fontSize: 10.5, margin: 0, valign: "top", isTextBox: true, objectName: "Repo link" });

  // ---------- 11. Recommendations ----------
  pres.addSection({ title: "Close" });
  s = pres.addSlide({ masterName: "Closing Dark", sectionTitle: "Close" });
  s.addText("Recommendations for the relief control room", { placeholder: "title" });
  const recs = [
    ["Revise daily", "Use tomorrow's Markov probabilities every evening"],
    ["Stock camps, not hubs", "Velachery and Mudichur are road-bound"],
    ["Act on the heavy day", "2,200 kits before the road floods"],
    ["Add storage at Mudichur", "It limits the plan at high penalties"],
    ["Cap forecasting spend", "Up to the EVPI, about Rs 60k a day"],
    ["Route by safety; pump the exit", "81% vs 13.5%; +180 vehicles/hour"],
  ];
  recs.forEach((r, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.5 + col * 4.6, y = 1.3 + row * 1.25;
    s.addText(String(i + 1), { x, y, w: 0.6, h: 0.6, fontSize: 30, bold: true, color: "8FD3DA", fontFace: "Cambria", margin: 0, isTextBox: true, objectName: "Rec number " + (i + 1) });
    s.addText(r[0], { x: x + 0.7, y, w: 3.7, h: 0.4, fontSize: 17, bold: true, color: C.background1, margin: 0, isTextBox: true, objectName: "Rec title " + (i + 1) });
    s.addText(r[1], { x: x + 0.7, y: y + 0.42, w: 3.7, h: 0.4, fontSize: 13, color: "CADCFC", margin: 0, isTextBox: true, objectName: "Rec text " + (i + 1) });
  });
  linkButton(s, 2.75, 4.78, 4.5, 0.55, "FFD24A", HEX.navy, "Closing link", 11);

  // ---------- 12. Limits & AI ----------
  s = pres.addSlide({ masterName: "Content", sectionTitle: "Close" });
  s.addText("Limits, and how AI was used", { placeholder: "title" });
  card(s, 0.5, 1.2, 4.3, 3.8, "Limits card");
  s.addText("Limits", { x: 0.75, y: 1.35, w: 3.8, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "Limits title" });
  s.addText([
    { text: "Illustrative network and synthetic rainfall: the method carries over, the rupee values do not", options: { bullet: true, breakLine: true } },
    { text: "One rain level for the whole city", options: { bullet: true, breakLine: true } },
    { text: "First-order Markov; two stages only", options: { bullet: true, breakLine: true } },
    { text: "Penalty value is a judgement (tested from Rs 100 to Rs 2,000)", options: { bullet: true } },
  ], { x: 0.75, y: 1.85, w: 3.85, h: 3.0, fontSize: 13, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 8, isTextBox: true, objectName: "Limits bullets" });
  card(s, 5.2, 1.2, 4.3, 3.8, "AI card", "DDEFF0");
  await circleIcon(s, md.MdAutoAwesome, 5.45, 1.32, 0.45, C.accent1, "AI");
  s.addText("AI as co-builder", { x: 6.05, y: 1.35, w: 3.2, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "AI title" });
  s.addText([
    { text: "Claude helped frame the problem, generate data, write the Python, Excel and web optimiser, and draft the report", options: { bullet: true, breakLine: true } },
    { text: "Every number was cross-checked: Excel = Python = web page", options: { bullet: true, breakLine: true } },
    { text: "Bottlenecks hand-checked; one figure corrected in review", options: { bullet: true } },
  ], { x: 5.45, y: 1.95, w: 3.85, h: 2.9, fontSize: 13, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 8, isTextBox: true, objectName: "AI bullets" });

  const file = path.join(__dirname, "Rain_Ready_Chennai_Slides.pptx");
  await pres.writeFile({ fileName: file });
  await applyTheme(file, THEME);
  console.log("saved", file);
})();
