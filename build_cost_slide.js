// Build: NODE_PATH=<dir with node_modules> node build_cost_slide.js
// One-slide extract: 5-year annual operating & maintenance cost (source: Maintenance_Cost_Basis_v2_EN.xlsx, sheet "Annual Cost Schedule").
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const SKILL = "/root/.claude/skills/synced/740b1519-7df3-42c4-87b2-972ad2eb207a_a3d7b2ae-009d-4d22-98fe-59c0c04fc335/pptx";
const { applyTheme } = require(SKILL + "/scripts/apply_theme.js");

const OUT = "KTPH_Annual_OM_Cost_Slide.pptx";
const THEME = {
  name: "MONIT RFP", headFontFace: "Cambria", bodyFontFace: "Calibri",
  colors: { dk1: "1F2937", lt1: "FFFFFF", dk2: "0B2A4A", lt2: "F1F4F8", accent1: "00A3AD", accent2: "0B2A4A", accent3: "F2A900", accent4: "6B7A8F", accent5: "2E7D6B", accent6: "C0392B", hlink: "00A3AD", folHlink: "6B7A8F" },
};
const H = { propBg: "E6F6F7", propTx: "0B6F76", navy: "0B2A4A", teal: "00A3AD", grey: "6B7A8F", light: "F1F4F8", line: "D5DCE6", amber: "F2A900", amberBg: "FFF4D6", amberTx: "7A4F00", text: "1F2937", white: "FFFFFF", sub: "DDEFF1" };

// ---------- figures (inputs from the workbook; years 2-5 follow the 5% escalation) ----------
const ESC = 0.05;
const PENTEST_Y1 = 12000; // workbook schedule value; the Summary sheet implies 0 (see notes)
const esc = (y1, n) => y1 * Math.pow(1 + ESC, n);
const yrs = [0, 1, 2, 3, 4];
const azure = yrs.map((n) => esc(24000, n));
const fd = yrs.map((n) => esc(4800, n));
const saas = yrs.map((n) => esc(12000, n));
const pen = yrs.map((n) => (n === 0 ? PENTEST_Y1 : esc(7000, n - 1)));
// ISO/IEC 27001 + 27017 + 27018 certification (one certification body, integrated audit).
// Year 1 = initial certification S$8,000. Later years follow the standard 3-year cycle
// (ISO/IEC 17021-1, IAF MD 5): surveillance audit = 1/3 and recertification = 2/3 of the
// initial audit effort, then the same 5% escalation as the other lines.
const ISO_Y1 = 8000;
const ISO_CYCLE = ["initial", "surveillance", "surveillance", "recertification", "surveillance"];
const isoFactor = { initial: 1, surveillance: 1 / 3, recertification: 2 / 3 };
const iso = yrs.map((n) => ISO_Y1 * isoFactor[ISO_CYCLE[n]] * Math.pow(1 + ESC, n));
const labourTotal = yrs.map((n) => esc(36000, n));
const item6 = [19200, 19800, 20400, 21000, 21600];
const labourSub = yrs.map((n) => labourTotal[n] - item6[n]);
const subtotal = yrs.map((n) => azure[n] + fd[n] + saas[n] + pen[n] + iso[n] + labourSub[n]);
const opTotal = yrs.map((n) => subtotal[n] + item6[n]);
const setup = [14285.71, 0, 0, 0, 0];
const grand = yrs.map((n) => opTotal[n] + setup[n]);
const sum = (a) => a.reduce((x, y) => x + y, 0);
const f = (v) => (v ? Math.round(v).toLocaleString("en-US") : "–");
const k = (v) => "S$" + (v / 1000).toFixed(1) + "k";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "KTPH-RFP-26-165-MJ — Annual operating and maintenance cost";
pres.author = "MONIT";
const C = pres.SchemeColor;
pres.defineSlideMaster({
  title: "CONTENT", background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.7, w: 12.13, h: 0.95, fontSize: 22, bold: true, align: "left", color: C.text2, valign: "top", margin: 0 }, text: "" } },
    { text: { text: "MONIT  |  KTPH-RFP-26-165-MJ  |  Discussion document — Confidential", options: { x: 0.6, y: 7.05, w: 8, h: 0.25, fontSize: 9, color: C.accent4, margin: 0 } } },
  ],
  slideNumber: { x: 12.0, y: 7.05, w: 0.73, h: 0.25, fontSize: 9, color: C.accent4, align: "right" },
});
const T = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, fontSize: 12, color: C.text1, valign: "top", ...o });
const iconCache = {};
async function icon(name, color = "#FFFFFF") {
  const key = name + color;
  if (iconCache[key]) return iconCache[key];
  const svg = RDS.renderToStaticMarkup(React.createElement(fa[name] || fa.FaCircle, { color, size: "256" }));
  return (iconCache[key] = "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64"));
}

async function main() {
  pres.addSection({ title: "3 · Smart diaper system" });
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "3 · Smart diaper system" });
  T(s, "3 · SMART DIAPER SYSTEM", { x: 0.6, y: 0.32, w: 8, h: 0.26, fontSize: 10, bold: true, color: C.accent1, charSpacing: 2 });
  s.addText(`Five-year operating and maintenance cost is ${k(sum(opTotal))}, with 5% annual escalation built in so charges stay fixed`, { placeholder: "title" });
  T(s, "Source: MONIT cost basis (Rev. 2, 16 Sep 2026); Master Agreement Clause 13.1 (Charges fixed for the term). SGD, ex-GST, rounded; 5.0% a year from Year 2. ISO: Y1 initial certification; Y2, Y3, Y5 surveillance (1/3) and Y4 recertification (2/3) of the initial audit, per ISO/IEC 17021-1.", { x: 0.6, y: 6.6, w: 12.1, h: 0.4, fontSize: 9, color: C.accent4 });

  // KPI cards
  const kpis = [
    ["FaCalendarCheck", k(sum(opTotal) / 5), "Average operating cost / year"],
    ["FaChartLine", k(sum(opTotal)), "Five-year operating cost"],
    ["FaCloud", k(sum(subtotal)), `Annual subscription (${Math.round((sum(subtotal) / sum(opTotal)) * 100)}% of total)`],
    ["FaTools", k(sum(item6)), `Preventive maintenance (${Math.round((sum(item6) / sum(opTotal)) * 100)}% of total)`],
  ];
  const cw = 2.9, gap = (12.13 - 4 * cw) / 3;
  for (let i = 0; i < kpis.length; i++) {
    const [ic, big, lab] = kpis[i];
    const x = 0.6 + i * (cw + gap), y = 1.75;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: 1.0, fill: { color: C.background2 }, line: { type: "none" } });
    s.addShape(pres.shapes.OVAL, { x: x + 0.18, y: y + 0.22, w: 0.56, h: 0.56, fill: { color: C.text2 }, line: { type: "none" } });
    s.addImage({ data: await icon(ic), x: x + 0.18 + 0.15, y: y + 0.22 + 0.15, w: 0.26, h: 0.26 });
    T(s, big, { x: x + 0.9, y: y + 0.1, w: cw - 1.0, h: 0.5, fontSize: 26, bold: true, color: C.accent1, fontFace: "Cambria", valign: "middle" });
    T(s, lab, { x: x + 0.9, y: y + 0.6, w: cw - 1.0, h: 0.3, fontSize: 11, color: C.accent4, valign: "top" });
  }

  // table
  const fs = 10.5;
  const noB = [{ type: "none" }, { type: "none" }, { type: "none" }, { type: "none" }];
  const rule = [{ type: "none" }, { type: "none" }, { type: "solid", pt: 0.75, color: H.line }, { type: "none" }];
  const cell = (text, o = {}) => ({ text, options: { fontSize: fs, valign: "middle", color: H.text, border: rule, margin: [0.03, 0.08, 0.03, 0.08], ...o } });
  const num = (text, o = {}) => cell(text, { align: "right", ...o });
  const header = ["Line item", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5", "5-year total"].map((t, i) => cell(t, { bold: true, color: H.white, fill: { color: H.navy }, border: noB, align: i ? "right" : "left" }));
  const line = (label, arr, o = {}) => [cell(label, o), ...arr.map((v) => num(f(v), o)), num(f(sum(arr)), { bold: true, ...o })];
  const rows = [
    header,
    line("Azure server (Singapore, site-isolated)", azure, { fill: { color: H.white } }),
    line("Azure Front Door (WAF / CDN)", fd, { fill: { color: H.light } }),
    line("MONIT SaaS platform (Azure PaaS)", saas, { fill: { color: H.white } }),
    line("Penetration re-test (one per year)", pen, { fill: { color: H.light } }),
    line("ISO 27001/27017/27018 certification & audits", iso, { fill: { color: H.white } }),
    line("Remote support & 24×7 monitoring", labourSub, { fill: { color: H.light } }),
    line("Subscription (Section 3.1) subtotal", subtotal, { bold: true, fill: { color: H.sub } }),
    line("Preventive maintenance labour (Item 6)", item6, { fill: { color: H.white } }),
    line("Total annual operating cost", opTotal, { bold: true, color: H.white, fill: { color: H.navy } }),
    [cell("Initial set-up (one-off) – to confirm", { italic: true, color: H.amberTx, fill: { color: H.amberBg } }), num(f(setup[0]), { italic: true, color: H.amberTx, fill: { color: H.amberBg } }), ...[1, 2, 3, 4].map(() => num("–", { color: H.amberTx, fill: { color: H.amberBg } })), num(f(sum(setup)), { bold: true, italic: true, color: H.amberTx, fill: { color: H.amberBg } })],
    line("Total annual expenditure", grand, { bold: true, fill: { color: H.sub } }),
  ];
  s.addTable(rows, { x: 0.6, y: 2.98, w: 8.7, colW: [3.2, 0.83, 0.83, 0.83, 0.83, 0.83, 1.35], rowH: 0.29, autoPage: false });

  // right panel: what the price covers
  s.addShape(pres.shapes.RECTANGLE, { x: 9.55, y: 2.98, w: 3.18, h: 3.48, fill: { color: C.text2 }, line: { type: "none" } });
  T(s, "What the price covers", { x: 9.75, y: 3.12, w: 2.8, h: 0.35, fontSize: 14, bold: true, color: C.background1 });
  const pts = [
    "Hosting in Singapore with a separate tenancy for each hospital and n+1 redundancy",
    "Encrypted backup, replication and a web application firewall",
    "Annual penetration re-test by a licensed provider",
    "ISO 27001/27017/27018 certification, yearly surveillance audits and Year-4 recertification",
    "Remote support, incident response and 24×7 security monitoring",
    "Four preventive-maintenance rounds a year across all three hospitals",
  ].map((t, i, a) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < a.length - 1, paraSpaceAfter: 5 } }));
  T(s, pts, { x: 9.75, y: 3.5, w: 2.8, h: 2.9, fontSize: 10.5, color: C.background1 });

  s.addNotes([
    "Source: Maintenance_Cost_Basis_v2_EN.xlsx, sheet 'Annual Cost Schedule' (figures recalculated in build_cost_slide.js from the Year-1 inputs and the 5.0% escalation; they reproduce the sheet exactly).",
    "ISO/IEC 27001/27017/27018 (added 2026-10-07): Year 1 initial certification S$8,000; Years 2, 3, 5 surveillance audits at 1/3 and Year 4 recertification at 2/3 of the initial audit (ISO/IEC 17021-1 / IAF MD 5 three-year cycle), plus the 5% escalation. Note the public Data Protection Notice states certification is targeted for H1 2027.",
    "Labour is charged once: total O&M labour of 36,000 in Year 1 = Item 6 preventive maintenance 19,200 + subscription portion 16,800 (sheet 'Labour Split Check').",
    "Escalation of 5% is built in because MA Clause 13.1 does not permit the Charges to be adjusted during the term.",
    "OPEN POINTS BEFORE ISSUE (internal - do not show to KTPH): (1) The Summary sheet says Year-1 penetration testing was removed (Year-1 operating cost 76,800; five-year operating 454,539; total 468,825), but the Annual Cost Schedule still carries 12,000 in Year 1 (88,800; 466,539; 480,825) and its own remark says the pre-deployment test is absorbed in the System price. This slide follows the schedule; set PENTEST_Y1 = 0 in build_cost_slide.js to follow the Summary. (2) Initial set-up 14,285.71 has no derivation; shown in amber. (3) The Summary header says WH 40 sensors (140 total) while the contract draft total is 135 (KTPH 60 / TTSH 40 / WH 35). (4) The cost is 2.16x the Oct 2025 quotation (210,000 over five years); the Remarks Wording sheet attributes this to the Synapxe cloud security requirements. Keep the INTERNAL action-list sheet out of any external copy.",
  ].join("\n\n"));

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT, { iso: iso.map(Math.round), iso5: Math.round(sum(iso)), avg: Math.round(sum(opTotal) / 5), sub: subtotal.map(Math.round), grandY: grand.map(Math.round), opTotal: opTotal.map(Math.round), five: Math.round(sum(opTotal)), grand: Math.round(sum(grand)), subtotal5: Math.round(sum(subtotal)) });
}
main().catch((e) => { console.error(e); process.exit(1); });
