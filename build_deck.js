// Build: NODE_PATH=<dir with node_modules> node build_deck.js
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const SKILL = "/root/.claude/skills/synced/740b1519-7df3-42c4-87b2-972ad2eb207a_a3d7b2ae-009d-4d22-98fe-59c0c04fc335/pptx";
const { applyTheme } = require(SKILL + "/scripts/apply_theme.js");

const OUT = "KTPH_RFP_MONIT_Presentation_261003.pptx";
const THEME = {
  name: "MONIT RFP",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1F2937", lt1: "FFFFFF", dk2: "0B2A4A", lt2: "F1F4F8",
    accent1: "00A3AD", accent2: "0B2A4A", accent3: "F2A900",
    accent4: "6B7A8F", accent5: "2E7D6B", accent6: "C0392B",
    hlink: "00A3AD", folHlink: "6B7A8F",
  },
};
// hex twins (tables / shadows / icons need hex)
const H = { propBg: "E6F6F7", propTx: "0B6F76", navy: "0B2A4A", teal: "00A3AD", grey: "6B7A8F", light: "F1F4F8", line: "D5DCE6", amber: "F2A900", amberBg: "FFF4D6", amberTx: "7A4F00", red: "C0392B", green: "2E7D6B", text: "1F2937", white: "FFFFFF" };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "KTPH-RFP-26-165-MJ — AI-Enabled Smart Diaper Care System: Discussion Document";
pres.author = "MONIT";
const C = pres.SchemeColor;
const SHADOW = () => ({ type: "outer", color: "000000", opacity: 0.12, blur: 6, offset: 2, angle: 90 });

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: H.navy },
  objects: [],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.7, w: 12.13, h: 0.95, fontSize: 22, bold: true, align: "left", color: C.text2, valign: "top", margin: 0 }, text: "" } },
    { text: { text: "MONIT  |  KTPH-RFP-26-165-MJ  |  Discussion document — Confidential", options: { x: 0.6, y: 7.05, w: 8, h: 0.25, fontSize: 9, color: C.accent4, margin: 0 } } },
  ],
  slideNumber: { x: 12.0, y: 7.05, w: 0.73, h: 0.25, fontSize: 9, color: C.accent4, align: "right" },
});

// ---------- helpers ----------
const iconCache = {};
async function icon(name, color = "#FFFFFF") {
  const k = name + color;
  if (iconCache[k]) return iconCache[k];
  const Comp = fa[name] || fa.FaCircle;
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color, size: "256" }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return (iconCache[k] = "image/png;base64," + buf.toString("base64"));
}
const T = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, fontSize: 12, color: C.text1, valign: "top", ...o });

const SECTIONS = new Set();
function ensureSection(t) { if (!SECTIONS.has(t)) { SECTIONS.add(t); pres.addSection({ title: t }); } }
function newSlide(section, title, source, notes) {
  ensureSection(section);
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  T(s, section.toUpperCase(), { x: 0.6, y: 0.32, w: 8, h: 0.26, fontSize: 10, bold: true, color: C.accent1, charSpacing: 2 });
  s.addText(title, { placeholder: "title" });
  if (source) T(s, source, { x: 0.6, y: 6.78, w: 11.5, h: 0.22, fontSize: 9, color: C.accent4 });
  if (notes) s.addNotes(notes);
  return s;
}
// amber "input required" box
function inp(s, x, y, w, h, text, o = {}) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: H.amberBg }, line: { color: H.amber, width: 1, dashType: "dash" }, objectName: "Input placeholder" });
  T(s, "[INPUT REQUIRED] " + text, { x: x + 0.1, y, w: w - 0.2, h, fontSize: o.fs || 11, color: H.amberTx, valign: "middle", italic: true });
}
// teal dashed "proposed response" box (answer written in the form KTPH procurement prefers; not evidenced in the attachments)
function prop(s, x, y, w, h, text, o = {}) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: H.propBg }, line: { color: H.teal, width: 1, dashType: "dash" }, objectName: "Proposed response" });
  T(s, [{ text: "PROPOSED  ", options: { bold: true, color: H.propTx } }, { text, options: { color: H.text } }], { x: x + 0.1, y, w: w - 0.2, h, fontSize: o.fs || 11, valign: "middle" });
}
async function iconCircle(s, name, x, y, d = 0.5, bg = C.text2) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" } });
  const p = d * 0.5;
  s.addImage({ data: await icon(name), x: x + (d - p) / 2, y: y + (d - p) / 2, w: p, h: p });
}
// card: tinted panel, icon, header, bullets
async function card(s, x, y, w, h, o) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Card " + o.head });
  let ty = y + 0.2;
  if (o.icon) { await iconCircle(s, o.icon, x + 0.2, y + 0.2, 0.52); T(s, o.head, { x: x + 0.85, y: y + 0.2, w: w - 1.05, h: 0.52, fontSize: 14, bold: true, color: C.text2, valign: "middle" }); ty = y + 0.9; }
  else { T(s, o.head, { x: x + 0.2, y: y + 0.2, w: w - 0.4, h: 0.4, fontSize: 14, bold: true, color: C.text2, valign: "middle" }); ty = y + 0.7; }
  const items = (o.body || []).map((t, i, a) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < a.length - 1, paraSpaceAfter: 4 } }));
  if (items.length) T(s, items, { x: x + 0.2, y: ty, w: w - 0.4, h: y + h - ty - 0.15, fontSize: o.fs || 12 });
}
function table(s, rows, x, y, w, colW, o = {}) {
  const fs = o.fs || 11;
  const data = rows.map((r, ri) => r.map((c, ci) => {
    const isInput = typeof c === "string" && c.startsWith("[INPUT");
    const base = { fontSize: fs, valign: "middle", color: H.text, border: [{ type: "none" }, { type: "none" }, { type: "solid", pt: 0.75, color: H.line }, { type: "none" }], margin: [0.05, 0.08, 0.05, 0.08] };
    if (ri === 0) return { text: c, options: { ...base, bold: true, color: H.white, fill: { color: H.navy }, border: [{ type: "none" }, { type: "none" }, { type: "none" }, { type: "none" }] } };
    if (isInput) return { text: c, options: { ...base, italic: true, color: H.amberTx, fill: { color: H.amberBg } } };
    if (typeof c === "string" && c.startsWith("[PROPOSED]")) return { text: [{ text: "PROPOSED  ", options: { bold: true, color: H.propTx } }, { text: c.slice(10).trim(), options: { color: H.text } }], options: { ...base, fill: { color: H.propBg } } };
    return { text: c, options: { ...base, bold: ci === 0, fill: { color: ri % 2 ? H.white : H.light } } };
  }));
  s.addTable(data, { x, y, w, colW, rowH: o.rowH || 0.42, autoPage: false });
}
const bullets = (arr) => arr.map((t, i, a) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < a.length - 1, paraSpaceAfter: 5 } }));

async function main() {
  // ================= 1. TITLE =================
  {
    ensureSection("Introduction");
    const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Introduction" });
    T(s, "MONIT", { x: 0.9, y: 0.8, w: 4, h: 0.4, fontSize: 16, bold: true, color: C.accent1, charSpacing: 6 });
    T(s, "AI-Enabled Smart Diaper Care System", { x: 0.9, y: 2.0, w: 9.5, h: 1.6, fontSize: 40, bold: true, color: C.background1, fontFace: "Cambria" });
    T(s, "Response to RFP KTPH-RFP-26-165-MJ — discussion document for ALPS Pte. Ltd. and the KTPH / Woodlands Health / TTSH project teams", { x: 0.9, y: 3.8, w: 8.6, h: 0.9, fontSize: 18, color: "CADCFC" });
    T(s, "Supply and Delivery of AI-Enabled Smart Diaper Care System", { x: 0.9, y: 5.5, w: 8, h: 0.3, fontSize: 12, color: "CADCFC" });
    T(s, "Draft for discussion  |  Confidential", { x: 0.9, y: 5.85, w: 8, h: 0.3, fontSize: 12, color: "CADCFC" });
    s.addShape(pres.shapes.OVAL, { x: 10.2, y: 1.3, w: 2.9, h: 2.9, fill: { color: C.accent1, transparency: 25 }, line: { type: "none" } });
    s.addShape(pres.shapes.OVAL, { x: 11.2, y: 3.4, w: 1.7, h: 1.7, fill: { color: C.background1, transparency: 85 }, line: { type: "none" } });
    s.addImage({ data: await icon("FaHospital"), x: 10.95, y: 2.05, w: 1.4, h: 1.4 });
    s.addNotes("Title slide. Replace 'MONIT' wordmark with the official logo before the meeting. All amber [INPUT REQUIRED] boxes in this deck mark information that was NOT in the RFP files and must be supplied by MONIT.");
  }

  // ================= 2. EXEC SUMMARY =================
  {
    const s = newSlide("Executive summary", "MONIT gives KTPH, Woodlands Health and TTSH one proven, fully supported smart diaper system, delivered under clear commitments", "Source: ALPS RFP KTPH-RFP-26-165-MJ; MONIT Master Agreement submission (Schedules 2 and 3).",
      "Purpose of this deck: help the KTPH purchasing team understand the MONIT solution and judge its suitability. Tone: factual, committed, verifiable. Amber boxes = facts MONIT must still supply; teal PROPOSED boxes = answers written in the form procurement prefers where the files had no evidence. Remove both marker styles when finalising.");
    const rows = [
      ["1", "A complete system, one accountable vendor", "Sensor, relay, cloud server, web dashboard and iOS/Android app — supplied, installed and commissioned by MONIT."],
      ["2", "Built for the hospital network", "Relay joins existing Wi-Fi; the password is entered once by hospital IT and is never held by MONIT."],
      ["3", "Sized from your floor plans", "Relay quantity set ward by ward from the three hospital layouts and verified by a Wi-Fi/RF survey before ordering."],
      ["4", "Compliant and certified", "CE, IMDA and KC listed in the contract; ISO 27001/27017/27018 certification underway; PDPA terms accepted."],
      ["5", "Supported after go-live", "Named engineers, 24-hour hotline, 99% platform uptime commitment, 7-year spare-parts commitment."],
    ];
    let y = 1.85;
    for (const [n, h, d] of rows) {
      s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.6, h: 0.6, fill: { color: C.text2 }, line: { type: "none" } });
      T(s, n, { x: 0.6, y, w: 0.6, h: 0.6, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle", fontFace: "Cambria" });
      T(s, h, { x: 1.5, y, w: 3.4, h: 0.6, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
      T(s, d, { x: 5.0, y, w: 7.7, h: 0.6, fontSize: 13, valign: "middle" });
      if (n !== "5") s.addShape(pres.shapes.LINE, { x: 0.6, y: y + 0.82, w: 12.13, h: 0, line: { color: H.line, width: 0.75 } });
      y += 0.97;
    }
  }

  // ================= 3. RFP AT A GLANCE =================
  {
    const s = newSlide("Executive summary", "One system, three public hospitals: firm pricing, mandatory relay sizing and a vendor presentation", "Source: RFP Consent Form; Section 1 – Conditions of RFP (clauses 5, 8, 10, 13, 20). Dates as stated in the RFP documents.",
      "Facts taken directly from the RFP consent form and Section 1. Several dates (site briefings, clarifications, closing) may already have passed - confirm the current stage with ALPS.");
    T(s, "Key RFP milestones", { x: 0.6, y: 1.8, w: 5.5, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
    const dates = [
      ["25 Aug 2026", "RFP documents available on Ariba"],
      ["31 Aug – 2 Sep", "Compulsory site briefings: KTPH → Woodlands Health → TTSH"],
      ["11 Sep, 12:00", "Written clarifications deadline (email to MMD)"],
      ["18 Sep, 15:00", "RFP closes on Ariba — late submissions blocked"],
      ["On request", "Vendor presentation and/or samples (clause 20)"],
    ];
    let y = 2.3;
    for (const [d, t] of dates) {
      s.addShape(pres.shapes.OVAL, { x: 0.65, y: y + 0.08, w: 0.2, h: 0.2, fill: { color: C.accent1 }, line: { type: "none" } });
      T(s, d, { x: 1.05, y, w: 1.7, h: 0.4, fontSize: 13, bold: true, color: C.text2, valign: "middle" });
      T(s, t, { x: 2.8, y, w: 3.4, h: 0.4, fontSize: 12, valign: "middle" });
      y += 0.6;
    }
    const stats = [["180", "days price validity from closing date"], ["10%", "security deposit within 14 days of acceptance"], ["28", "days minimum to mobilise resources"], ["60", "days payment term"]];
    stats.forEach(([n, l], i) => {
      const x = 6.6 + (i % 2) * 3.1, yy = 1.85 + Math.floor(i / 2) * 1.6;
      s.addShape(pres.shapes.RECTANGLE, { x, y: yy, w: 2.95, h: 1.45, fill: { color: C.background2 }, line: { type: "none" } });
      T(s, n, { x: x + 0.2, y: yy + 0.12, w: 2.5, h: 0.75, fontSize: 40, bold: true, color: C.accent1, fontFace: "Cambria" });
      T(s, l, { x: x + 0.2, y: yy + 0.88, w: 2.55, h: 0.5, fontSize: 11, color: C.accent4 });
    });
    T(s, "Three sites in scope", { x: 6.6, y: 5.1, w: 6, h: 0.3, fontSize: 14, bold: true, color: C.text2 });
    ["Khoo Teck Puat Hospital", "Woodlands Health", "Tan Tock Seng Hospital"].forEach((n, i) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.6 + i * 2.07, y: 5.5, w: 1.95, h: 0.8, rectRadius: 0.08, fill: { color: C.text2 }, line: { type: "none" } });
      T(s, n, { x: 6.7 + i * 2.07, y: 5.5, w: 1.75, h: 0.8, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
    });
    T(s, "SGD, ex-GST  •  HSA-registered devices  •  Fully integrated system", { x: 0.6, y: 5.45, w: 5.6, h: 0.3, fontSize: 12, color: C.accent4 });
    // reading guide
    T(s, "Reading guide (remove before issue)", { x: 0.6, y: 5.85, w: 5.6, h: 0.25, fontSize: 10, bold: true, color: C.accent4 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 6.15, w: 0.3, h: 0.22, fill: { color: H.amberBg }, line: { color: H.amber, width: 1, dashType: "dash" } });
    T(s, "MONIT to supply a fact", { x: 0.98, y: 6.15, w: 2.0, h: 0.22, fontSize: 10, valign: "middle" });
    s.addShape(pres.shapes.RECTANGLE, { x: 3.1, y: 6.15, w: 0.3, h: 0.22, fill: { color: H.propBg }, line: { color: H.teal, width: 1, dashType: "dash" } });
    T(s, "Proposed response, to be confirmed", { x: 3.48, y: 6.15, w: 2.8, h: 0.22, fontSize: 10, valign: "middle" });
  }

  // ================= 3B. COMPLIANCE MAP =================
  {
    const s = newSlide("Executive summary", "Every RFP requirement has a specific MONIT response, so each can be checked against evidence", "Source: RFP Consent Form and Section 1; MONIT Master Agreement submission. Slide numbers refer to this document.",
      "Procurement evaluation aid: requirement -> response -> where to verify. Keep this table in sync with the final deck. Where the evidence is still being collected the 'Evidence' cell says so.");
    table(s, [
      ["RFP requirement", "MONIT response", "Evidence", "Slide"],
      ["Certifications and standards", "CE, IMDA, KC; ISO 27001/27017/27018 underway; BizSAFE, ISO 9001/13485/14001", "Certificates on request", "6"],
      ["Clinically validated detection", "Validation studies and measured detection performance", "Study reports", "8"],
      ["Software and dashboard", "Ward dashboard, alerts, reports; iOS/Android app; PDPA-aligned", "Live demo", "9"],
      ["Maintenance and support", "12-month maintenance, 24-hr hotline, 99% uptime, 7-year spares", "Master Agreement Sch. 2/3", "10–11"],
      ["Approach and deployment", "Floor-plan-based plan, acceptance steps, relay count per ward", "Ward table, survey", "14–18"],
      ["Timeline and lead times", "Milestones from Letter of Award to Final Acceptance", "Gantt chart", "19"],
      ["References and sustainability", "Comparable deployments; four sustainability criteria answered", "Referee contacts", "20–21"],
      ["Wi-Fi connectivity and password", "Relay setup once; settings survive outages and upgrades", "Live relay demo", "22–24"],
    ], 0.6, 1.85, 12.13, [3.2, 5.6, 2.4, 0.93], { rowH: 0.52, fs: 11 });
  }

  // ================= 4. COMPANY =================
  {
    const s = newSlide("1 · Company introduction", "MONIT is a healthcare-technology company focused on making incontinence care smarter and more dignified", "Source: MONIT. Replace placeholders with approved corporate facts.",
      "Fill with approved company profile. Suggested content: founding year, HQ and Singapore entity/UEN (needed for the Master Agreement), team size, R&D focus, number of installed beds/countries, funding/partners.");
    inp(s, 0.6, 1.85, 5.6, 1.2, "Company profile: founded, HQ, Singapore entity & UEN, ownership, headcount, mission statement.");
    inp(s, 0.6, 3.25, 5.6, 1.4, "Core technology / IP: sensor hardware, AI detection algorithm, patents, R&D footprint.");
    inp(s, 0.6, 4.85, 5.6, 1.5, "Healthcare footprint: countries, hospitals / care homes served, beds covered.");
    const kp = [["[xx]", "Years in operation"], ["[xx]", "Sites / facilities live"], ["[xx]", "Beds monitored"], ["[xx]", "Countries"]];
    kp.forEach(([n, l], i) => {
      const x = 6.6 + (i % 2) * 3.1, y = 1.85 + Math.floor(i / 2) * 1.7;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.95, h: 1.5, fill: { color: C.background2 }, line: { type: "none" } });
      T(s, n, { x: x + 0.2, y: y + 0.15, w: 2.5, h: 0.8, fontSize: 40, bold: true, color: C.accent3, fontFace: "Cambria" });
      T(s, l, { x: x + 0.2, y: y + 0.95, w: 2.5, h: 0.4, fontSize: 12, color: C.accent4 });
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 6.6, y: 5.3, w: 6.05, h: 1.05, fill: { color: C.text2 }, line: { type: "none" } });
    T(s, "Why it matters to ALPS: a single accountable vendor for hardware, software, installation and service across three hospitals.", { x: 6.8, y: 5.3, w: 5.65, h: 1.05, fontSize: 13, color: C.background1, valign: "middle" });
  }

  // ================= 5. CERTIFICATIONS =================
  {
    const s = newSlide("2 · Industry certifications and standards", "MONIT's CE, IMDA and KC certifications are declared in the contract, and ISO 27001/27017/27018 certification is underway", "Source: Master Agreement Sch. 2, Clause 5; RFP Consent Form (HSA); Section 1 clause 19. Certificate copies available on request.",
      "Schedule 2 Clause 5 lists CE marking, IMDA registration (Singapore) and KC (Korea); copies of valid certificates are to be provided on written request. ISO/IEC 27001, 27017 and 27018 are planned (not yet certified) - state certification body, audit stage and target date; do not describe them as held until issued. Provide numbers and expiry dates. ALPS is ISO14000 and OSHA certified and expects vendors to follow its environmental and safety requirements. Confirm HSA classification/registration of the sensor.");
    table(s, [
      ["Standard / certification", "Why it matters for this RFP", "Status", "Valid until"],
      ["CE marking", "Product conformity of sensor and relay", "Declared in Master Agreement Sch. 2", "[INPUT]"],
      ["IMDA registration (Singapore)", "Relay radio (BLE 5 + Wi-Fi 2.4 GHz) lawfully used in Singapore", "Declared in Master Agreement Sch. 2", "[INPUT]"],
      ["KC certification (Korea)", "Product certification in the country of manufacture", "Declared in Master Agreement Sch. 2", "[INPUT]"],
      ["HSA – medical device registration", "RFP requires registration for Class B/C/D and Class A devices — classification to be confirmed", "[INPUT REQUIRED] class & reg. no.", "[INPUT]"],
      ["ISO/IEC 27001 / 27017 / 27018", "Information security, cloud security controls and PII protection in the cloud; supports PDPA", "Certification underway; target [INPUT]", "—"],
      ["BizSAFE", "Contractor safety for on-site installation in hospitals", "[INPUT REQUIRED] level & cert no.", "[INPUT]"],
      ["ISO 9001 / ISO 13485", "Quality management; medical-device QMS for design and manufacture", "[INPUT REQUIRED]", "[INPUT]"],
      ["ISO 14001 / ISO 45001", "Aligns with ALPS's ISO14000 / OSHA requirements", "[INPUT REQUIRED]", "[INPUT]"],
    ], 0.6, 1.85, 12.13, [3.1, 5.2, 2.78, 1.05], { rowH: 0.52, fs: 10.5 });
  }

  // ================= 6. CONCEPT =================
  {
    const s = newSlide("3 · Smart diaper system", "The system turns every diaper change into a timely, data-driven decision — detect, notify, document", "Source: Master Agreement Schedule 3 (Master Equipment and Master Services); MONIT.",
      "Specifications are from Schedule 3 of the Master Agreement. Note: the relay dimensions in Schedule 3 read '9.5 x 4.5 x 5.5-11.3 mm', which is probably cm - confirm before quoting. Sensor-to-relay link is BLE; relay-to-cloud is Wi-Fi 2.4 GHz.");
    const steps = [
      ["FaMicrochip", "Sensor unit", "Reusable sensor mounted on any standard diaper with a strap sticker; BLE"],
      ["FaWifi", "Relay device", "Collects sensor data in the ward; BLE 5 in, Wi-Fi 2.4 GHz out"],
      ["FaCloud", "Cloud server", "AI detection, event history, secure storage"],
      ["FaDesktop", "Web dashboard", "Management and nurse-station view (PC)"],
      ["FaMobileAlt", "Mobile app", "Alerts for on-site care workers (iOS / Android)"],
    ];
    for (let i = 0; i < steps.length; i++) {
      const x = 0.6 + i * 2.45;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.9, w: 2.2, h: 2.5, fill: { color: C.background2 }, line: { type: "none" } });
      await iconCircle(s, steps[i][0], x + 0.8, 2.05, 0.6, i % 2 ? C.accent1 : C.text2);
      T(s, steps[i][1], { x: x + 0.1, y: 2.8, w: 2.0, h: 0.35, fontSize: 14, bold: true, color: C.text2, align: "center" });
      T(s, steps[i][2], { x: x + 0.15, y: 3.2, w: 1.9, h: 1.1, fontSize: 11, align: "center" });
      if (i < steps.length - 1) s.addShape(pres.shapes.LINE, { x: x + 2.2, y: 2.35, w: 0.25, h: 0, line: { color: H.teal, width: 2, endArrowType: "triangle" } });
    }
    T(s, "Value to hospitals", { x: 0.6, y: 4.7, w: 5.6, h: 0.3, fontSize: 14, bold: true, color: C.text2 });
    T(s, bullets(["Right-time diaper changes instead of routine checks", "Less nurse time on manual rounds; fewer disturbances for patients", "Objective data to support skin-integrity and continence care"]), { x: 0.6, y: 5.05, w: 5.6, h: 1.5, fontSize: 13 });
    table(s, [
      ["Hardware", "Key specifications"],
      ["Sensor unit", "290 × 31 × 10.85 mm · 23 g · CR2032 coin cell · BLE"],
      ["Relay device", "42 g · 5 V USB-A, always-on · BLE 5 + Wi-Fi 2.4 GHz"],
      ["Consumables", "Strap stickers (600 pcs/set); single-use hygiene barrier film"],
    ], 6.5, 4.7, 6.23, [1.7, 4.53], { rowH: 0.45, fs: 10.5 });
  }

  // ================= 7. CLINICAL VALIDATION =================
  {
    const s = newSlide("3 · Smart diaper system", "Clinical validation demonstrates reliable detection performance of the sensor hardware", "Source: MONIT clinical / validation reports. Do not present figures until verified against the source studies.",
      "Provide published or internal clinical evidence: study design, site, number of patients/events, sensitivity, specificity, latency, false-alarm rate. Never present numbers that are not traceable to a report.");
    const k = [["[xx]%", "Sensitivity"], ["[xx]%", "Specificity"], ["[xx] s", "Median detection latency"], ["[xx]", "False alarms / bed / day"]];
    k.forEach(([n, l], i) => {
      const x = 0.6 + i * 3.07;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.85, w: 2.9, h: 1.4, fill: { color: C.background2 }, line: { type: "none" } });
      T(s, n, { x: x + 0.2, y: 1.95, w: 2.5, h: 0.75, fontSize: 36, bold: true, color: C.accent3, fontFace: "Cambria" });
      T(s, l, { x: x + 0.2, y: 2.75, w: 2.5, h: 0.4, fontSize: 12, color: C.accent4 });
    });
    table(s, [
      ["Study / report", "Setting", "Sample", "Design", "Key result"],
      ["[INPUT REQUIRED]", "[INPUT]", "[INPUT]", "[INPUT]", "[INPUT]"],
      ["[INPUT REQUIRED]", "[INPUT]", "[INPUT]", "[INPUT]", "[INPUT]"],
      ["[INPUT REQUIRED]", "[INPUT]", "[INPUT]", "[INPUT]", "[INPUT]"],
    ], 0.6, 3.55, 12.13, [3.2, 2.4, 1.8, 2.3, 2.43], { rowH: 0.5 });
    T(s, "Also include: hospital-environment testing (RF interference, multi-bed density), safety / biocompatibility tests, third-party test-lab reports.", { x: 0.6, y: 5.85, w: 12.13, h: 0.5, fontSize: 12, color: C.accent4 });
  }

  // ================= 8. SOFTWARE =================
  {
    const s = newSlide("3 · Smart diaper system", "A ward-level dashboard gives nurses a single view of who needs care now", "Source: MONIT. Screen shown is an illustrative wireframe, not the final UI.",
      "Replace the wireframe with real screenshots. Confirm: user roles, audit log, EMR/nurse-call integration options, data hosting location (Singapore), PDPA compliance.");
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.85, w: 6.3, h: 4.6, fill: { color: "FFFFFF" }, line: { color: H.line, width: 1 }, shadow: SHADOW() });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.85, w: 6.3, h: 0.45, fill: { color: C.text2 }, line: { type: "none" } });
    T(s, "Ward 5A — Live bed status (illustrative)", { x: 0.8, y: 1.85, w: 5.8, h: 0.45, fontSize: 12, bold: true, color: C.background1, valign: "middle" });
    const st = [C.accent1, C.accent1, C.accent3, C.accent1, C.accent6, C.accent1, C.accent1, C.accent3, C.accent1, C.accent1, C.accent1, C.accent6];
    st.forEach((c, i) => {
      const x = 0.85 + (i % 4) * 1.5, y = 2.55 + Math.floor(i / 4) * 1.1;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 1.35, h: 0.95, fill: { color: c, transparency: 15 }, line: { type: "none" } });
      T(s, "Bed " + String(i + 1).padStart(2, "0"), { x, y: y + 0.1, w: 1.35, h: 0.3, fontSize: 11, bold: true, color: C.background1, align: "center" });
      T(s, c === C.accent6 ? "Change now" : c === C.accent3 ? "Soon" : "OK", { x, y: y + 0.45, w: 1.35, h: 0.3, fontSize: 11, color: C.background1, align: "center" });
    });
    T(s, [{ text: "● OK   ", options: { color: C.accent1 } }, { text: "● Change soon   ", options: { color: C.accent3 } }, { text: "● Change now", options: { color: C.accent6 } }], { x: 0.85, y: 5.95, w: 5, h: 0.3, fontSize: 11, bold: true });
    const f = [["FaBell", "Real-time alerts", "Nurse-station dashboard and mobile app (iOS / Android) with escalation rules"], ["FaListAlt", "Care records", "Event history per bed; exportable for documentation"], ["FaChartBar", "Analytics", "Change intervals, response times, ward workload trends"], ["FaUserLock", "Admin & security", "Role-based access, audit trail, PDPA-aligned data handling"]];
    for (let i = 0; i < f.length; i++) {
      const y = 1.85 + i * 1.18;
      await iconCircle(s, f[i][0], 7.3, y + 0.1, 0.55, i % 2 ? C.accent1 : C.text2);
      T(s, f[i][1], { x: 8.1, y, w: 4.6, h: 0.35, fontSize: 14, bold: true, color: C.text2 });
      T(s, f[i][2], { x: 8.1, y: y + 0.37, w: 4.6, h: 0.65, fontSize: 12 });
    }
  }

  // ================= 9. MAINTENANCE & SUPPORT =================
  {
    const s = newSlide("3 · Smart diaper system", "MONIT backs the system with a 12-month maintenance period, a 24-hour hotline, a 99% uptime commitment and named engineers", "Source: Master Agreement Schedule 2 (Key Terms), Schedule 3 and Transaction Schedule. Response and repair times follow the terms in MONIT's Master Agreement submission.",
      "ALIGNMENT FLAGS (internal): (1) Compliance-to-MA v2 remark on Clause 12 says MONIT accepts 3 h response / 24 h repair / next-business-day loaner, whereas Schedule 2 [DRAFT] says 8 business hours / 3 business days / 3 business days - decide which MONIT will stand behind before submission. (2) Warranty in Schedule 3 is sensor 6 m, relay 12 m, platform 12 m, vs MA default 24 months. (3) Spare-parts commitment is 7 years from delivery (MA 16.2.3) vs 10 years after end-of-life in Section 3.1 - raised with MMD. The 24-hour hotline number is in the Transaction Schedule (not shown here).");
    table(s, [
      ["Service element", "MONIT commitment"],
      ["Warranty and maintenance period", "Sensor 6 months · relay 12 months · platform 12 months"],
      ["Support hours", "Mon–Fri 09:00–18:00 Singapore time (excl. public holidays); 24-hour hotline"],
      ["Preventive maintenance", "At least one week's notice before each visit"],
      ["Corrective maintenance", "Response within 8 business hours; repair within 3 business days of attendance"],
      ["Temporary replacement", "Loan unit within 3 business days of request"],
      ["Platform uptime", "99% per calendar quarter, measured on the cloud dashboard"],
      ["Spare parts", "Maintained for 7 years from delivery"],
      ["Training", "Training plan within 14 days of contract; user training within 14 days of delivery, in English"],
    ], 0.6, 1.85, 7.9, [2.5, 5.4], { rowH: 0.5, fs: 10.5 });
    await card(s, 8.8, 1.85, 3.93, 3.1, { icon: "FaUserCog", head: "Named engineers", body: ["Cloud server and web dashboard: Chief Technology Officer", "Sensor units and BLE gateways: Chief Executive Officer", "Application services: application lead", "Replacements of equal qualification on written notice"], fs: 11 });
    prop(s, 8.8, 5.15, 3.93, 1.35, "Single point of contact per site, monthly service report and quarterly review meeting, matching the Master Agreement's quarterly project-manager meetings.", { fs: 10.5 });
  }
  // ================= 9B. SLA & ESCALATION =================
  {
    const s = newSlide("3 · Smart diaper system", "A clear severity scale and escalation path gives hospital teams one route from fault to fix", "Source: Master Agreement Schedule 2 and Transaction Schedule (rate card). Severity targets are a proposed response format and will be aligned with Schedule 2 before issue.",
      "PROPOSED RESPONSE (not evidenced in attachments): severity tiers written the way hospital procurement usually asks for them. P3 matches the Schedule 2 draft (8 business hours / 3 business days); P1 and P2 are tighter targets that need MONIT management approval. The rate card is from the Transaction Schedule: weekday 08:30-18:00 S$220 first hour / S$150 after; Saturday 08:30-12:30 S$280 / S$190; after hours S$330 / S$220; Sundays and public holidays S$440 / S$300; up to 12 chargeable breakdown attendances per site per year are covered by the response commitment.");
    table(s, [
      ["Severity", "Definition", "Response", "Restoration or workaround"],
      ["[PROPOSED] P1 – Critical", "Ward-wide loss of alerts", "[PROPOSED] 2 business hours", "[PROPOSED] 1 business day"],
      ["[PROPOSED] P2 – Major", "Several beds affected", "[PROPOSED] 4 business hours", "[PROPOSED] 2 business days"],
      ["P3 – Minor", "Single device or cosmetic issue", "Within 8 business hours", "Within 3 business days of attendance"],
    ], 0.6, 1.85, 7.4, [1.9, 2.1, 1.7, 1.7], { rowH: 0.62, fs: 10.5 });
    T(s, "Escalation path", { x: 8.4, y: 1.85, w: 4.3, h: 0.3, fontSize: 14, bold: true, color: C.text2 });
    ["Hospital user → hotline / helpdesk", "MONIT service engineer", "Approved maintenance personnel (by domain)", "MONIT Chief Technology Officer / project manager"].forEach((t, i) => {
      const y = 2.25 + i * 0.6;
      s.addShape(pres.shapes.OVAL, { x: 8.4, y: y + 0.03, w: 0.4, h: 0.4, fill: { color: i % 2 ? C.accent1 : C.text2 }, line: { type: "none" } });
      T(s, String(i + 1), { x: 8.4, y: y + 0.03, w: 0.4, h: 0.4, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
      T(s, t, { x: 8.95, y, w: 3.8, h: 0.5, fontSize: 11.5, valign: "middle" });
    });
    T(s, "Rate card for chargeable attendance (S$ per hour: first / subsequent)", { x: 0.6, y: 4.5, w: 8, h: 0.3, fontSize: 12, bold: true, color: C.text2 });
    table(s, [
      ["Mon–Fri 08:30–18:00", "Sat 08:30–12:30", "After hours", "Sun / public holiday"],
      ["220 / 150", "280 / 190", "330 / 220", "440 / 300"],
    ], 0.6, 4.85, 7.4, [1.85, 1.85, 1.85, 1.85], { rowH: 0.42, fs: 11 });
    prop(s, 0.6, 5.95, 12.13, 0.6, "Up to 12 chargeable breakdown attendances per site per year are covered by the response commitment; warranty replacements are free of charge during the warranty period.", { fs: 11 });
  }

  // ================= 10. OTHER INFO =================
  {
    const s = newSlide("3 · Smart diaper system", "Beyond detection, the system supports staff productivity, patient dignity and consumable efficiency", "Source: MONIT; benefit figures to be sourced from validated studies or pilot data.",
      "Add other relevant information: staff time saved, reduced skin complications, consumables optimisation, patient comfort, awards, regulatory approvals, roadmap, interoperability.");
    const c = [["FaClock", "Staff productivity", "Targeted rounds replace routine checks"], ["FaHeartbeat", "Patient outcomes", "Fewer prolonged-wetness events; supports skin integrity"], ["FaBoxOpen", "Consumables", "Change at the right time; reduced waste"], ["FaPuzzlePiece", "Integration & roadmap", "APIs, nurse-call / EMR options, future analytics"]];
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + (i % 2) * 6.15, y = 1.85 + Math.floor(i / 2) * 2.35;
      await card(s, x, y, 5.98, 2.15, { icon: c[i][0], head: c[i][1], body: [c[i][2], "[INPUT REQUIRED] quantified evidence"], fs: 13 });
    }
  }

  // ================= 11. UNDERSTANDING =================
  {
    const s = newSlide("4 · Understanding of the project", "We understand ALPS needs one integrated, hospital-network-friendly system deployed consistently across three hospitals", "Source: RFP Consent Form; Section 1 – Conditions of RFP (clauses 7, 20); Section 2 – Master Agreement; MONIT analysis.",
      "This slide is MONIT's reading of the RFP. Validate with ALPS in the meeting - particularly the number of wards/beds in scope, which is not stated in the files received.");
    T(s, "What the RFP asks for", { x: 0.6, y: 1.85, w: 5.8, h: 0.35, fontSize: 16, bold: true, color: C.text2 });
    T(s, bullets(["Supply and delivery of an AI-enabled smart diaper care system", "Fully integrated system with safe, efficient operation on site", "Three sites: KTPH, Woodlands Health, TTSH", "Connectivity through each hospital's Wi-Fi (credentials held by hospital)", "Relay quantity per ward (mandatory response)", "Firm prices valid 180 days; framework Master Agreement with POs; security deposit and insurance"]), { x: 0.6, y: 2.3, w: 5.8, h: 4.2, fontSize: 13 });
    T(s, "What ALPS will look for", { x: 6.9, y: 1.85, w: 5.8, h: 0.35, fontSize: 16, bold: true, color: C.text2 });
    const w = [["FaShieldAlt", "Compliance", "Certifications, HSA, IMDA, IT security & PDPA"], ["FaNetworkWired", "IT readiness", "Zero re-keying of Wi-Fi passwords; minimal IT effort"], ["FaCalendarCheck", "Delivery certainty", "Clear plan and lead times; low disruption to wards"], ["FaLeaf", "Sustainability", "Environmental practices aligned with ISO14000"]];
    for (let i = 0; i < 4; i++) {
      const y = 2.3 + i * 1.05;
      s.addShape(pres.shapes.RECTANGLE, { x: 6.9, y, w: 5.83, h: 0.9, fill: { color: C.background2 }, line: { type: "none" } });
      await iconCircle(s, w[i][0], 7.05, y + 0.17, 0.56, i % 2 ? C.accent1 : C.text2);
      T(s, w[i][1], { x: 7.8, y: y + 0.08, w: 4.8, h: 0.3, fontSize: 13, bold: true, color: C.text2 });
      T(s, w[i][2], { x: 7.8, y: y + 0.4, w: 4.8, h: 0.45, fontSize: 11.5 });
    }
  }

  // ================= 12. APPROACH =================
  {
    const s = newSlide("4 · Understanding of the project", "A four-phase approach takes each hospital from Letter of Award to Final Acceptance with defined checkpoints", "Source: Master Agreement (Clauses 12, 23; Schedules 2 and 3); MONIT proposed methodology. Durations on the timeline slide.",
      "Methodology. The contract milestones (training plan within 14 days, training within 14 days of delivery, pre-shipment testing, installation and commissioning by MONIT, acceptance tests, Final Acceptance per institution, invoice within 7 days) come from the Master Agreement. Note: security deposit and insurance timing is 14 days in RFP clause 10 and 30 days in Schedule 2 - plan to the shorter RFP period.");
    const ph = [
      ["1", "Mobilise and plan", ["Kick-off with ALPS and the three hospitals", "Security deposit and insurance lodged", "Training plan submitted (≤ 14 days)", "Floor-plan review, Wi-Fi/RF survey, final relay count"]],
      ["2", "Prepare and pilot", ["Pre-shipment testing at MONIT", "Delivery with advance notice", "Install and commission one pilot ward per hospital", "Acceptance tests"]],
      ["3", "Roll out", ["Ward-by-ward installation and commissioning", "User training (≤ 14 days after delivery)", "Minimal disruption to care routines", "Sign-off per ward"]],
      ["4", "Accept and support", ["Final Acceptance Notice per hospital", "Invoice within 7 days", "Hypercare, then 12-month maintenance", "Quarterly review meetings"]],
    ];
    ph.forEach((p, i) => {
      const x = 0.6 + i * 3.05;
      s.addShape(pres.shapes.PENTAGON || pres.shapes.RECTANGLE, { x, y: 1.9, w: 3.0, h: 0.8, fill: { color: i % 2 ? C.accent1 : C.text2 }, line: { type: "none" } });
      T(s, p[0] + "  " + p[1], { x: x + 0.25, y: 1.9, w: 2.4, h: 0.8, fontSize: 16, bold: true, color: C.background1, valign: "middle", fontFace: "Cambria" });
      s.addShape(pres.shapes.RECTANGLE, { x, y: 2.85, w: 2.85, h: 2.7, fill: { color: C.background2 }, line: { type: "none" } });
      T(s, bullets(p[2]), { x: x + 0.2, y: 3.0, w: 2.5, h: 2.5, fontSize: 12 });
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 5.75, w: 12.13, h: 0.85, fill: { color: C.text2 }, line: { type: "none" } });
    T(s, [{ text: "What the hospital provides:  ", options: { bold: true } }, { text: "network connectivity (2.4 GHz Wi-Fi coverage at relay locations), a powered outlet for each relay (5 V USB), and physical access to wards at agreed times. MONIT performs installation, commissioning and testing." }], { x: 0.85, y: 5.75, w: 11.6, h: 0.85, fontSize: 12, color: C.background1, valign: "middle" });
  }
  // ================= 12B. GOVERNANCE & ACCEPTANCE =================
  {
    const s = newSlide("4 · Understanding of the project", "Defined governance and acceptance steps let KTPH verify progress and quality at every stage", "Source: Master Agreement Clauses 4, 12, 13, 23 and Schedule 2; MONIT.",
      "Acceptance chain from the Master Agreement: acceptance tests and certificates, then Final Acceptance Notice per PO. Security deposit (S$18,000 = 10% of S$180,000, banker's bond) reduces pro rata on Final Acceptance per institution - this is a [DRAFT] in Schedule 2; confirm before quoting. Changes requested after PO acceptance are pre-approved in writing and reimbursed at cost.");
    const st = [["1", "Letter of Award", "Contract effective; project managers named"], ["2", "Deposit and insurance", "Security deposit and required insurance lodged"], ["3", "Training plan", "Submitted for approval within 14 days"], ["4", "Pre-shipment test", "Equipment tested at MONIT before delivery"], ["5", "Install and commission", "By MONIT at each hospital"], ["6", "Acceptance tests", "Witnessed by hospital representatives"], ["7", "Final Acceptance", "Notice issued per hospital; deposit reduces pro rata"], ["8", "Invoice", "Submitted within 7 days of Final Acceptance"]];
    st.forEach((p, i) => {
      const x = 0.6 + (i % 4) * 3.05, y = 1.85 + Math.floor(i / 4) * 1.6;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.9, h: 1.4, fill: { color: C.background2 }, line: { type: "none" } });
      s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: y + 0.15, w: 0.45, h: 0.45, fill: { color: i % 2 ? C.accent1 : C.text2 }, line: { type: "none" } });
      T(s, p[0], { x: x + 0.15, y: y + 0.15, w: 0.45, h: 0.45, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle" });
      T(s, p[1], { x: x + 0.72, y: y + 0.15, w: 2.1, h: 0.45, fontSize: 13, bold: true, color: C.text2, valign: "middle" });
      T(s, p[2], { x: x + 0.15, y: y + 0.75, w: 2.65, h: 0.6, fontSize: 11 });
    });
    await card(s, 0.6, 5.1, 3.95, 1.55, { head: "Project governance", body: ["Quarterly project-manager meetings", "MONIT project manager: Chief Technology Officer"], fs: 11 });
    await card(s, 4.69, 5.1, 3.95, 1.55, { head: "Change control", body: ["Changes pre-approved in writing", "Reimbursed at cost; no hidden charges"], fs: 11 });
    await card(s, 8.78, 5.1, 3.95, 1.55, { head: "Quality", body: ["Pre-shipment testing", "Acceptance tests per ward"], fs: 11 });
  }

  // ================= 13. DEPLOYMENT PER HOSPITAL =================
  {
    const s = newSlide("4 · Understanding of the project", "Deployment is planned ward by ward from each hospital's floor plan, using the same proven pattern at all three sites", "Source: Hospital floor plans provided by ALPS at the site briefings; MONIT deployment pattern.",
      "Insert each hospital's floor plan and mark relay positions; the floor plans were not in the files supplied for this draft. The deployment principles are MONIT's proposed practice (not from the attachments).");
    const hs = ["Khoo Teck Puat Hospital", "Woodlands Health", "Tan Tock Seng Hospital"];
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.1;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.85, w: 3.95, h: 0.5, fill: { color: C.text2 }, line: { type: "none" } });
      T(s, hs[i], { x: x + 0.15, y: 1.85, w: 3.7, h: 0.5, fontSize: 14, bold: true, color: C.background1, valign: "middle" });
      inp(s, x, 2.45, 3.95, 1.75, "Floor plan with ward boundaries and relay positions marked.", { fs: 11 });
      T(s, bullets(["Wards in scope: [INPUT]", "Beds / sensors in scope: [INPUT]", "Relays: see ward table"]), { x, y: 4.3, w: 3.95, h: 0.95, fontSize: 12 });
    }
    prop(s, 0.6, 5.35, 12.13, 1.25, "Deployment pattern: one relay per bed cluster at a powered outlet within BLE range of its sensors; relay positions checked against 2.4 GHz Wi-Fi coverage; dashboard at each nurse station; pilot ward first, then ward-by-ward rollout outside medication and meal rounds; infection-control and access rules followed per hospital.", { fs: 12 });
  }

  // ================= 14. RELAY SIZING METHOD =================
  {
    const s = newSlide("4 · Understanding of the project", "Relay quantities follow a transparent, coverage-based method that is verified by an on-site survey", "Source: Master Agreement Schedule 3 (draft quantities: 135 sensors, 53 relays; relay specification); MONIT sizing method.",
      "MANDATORY RFP ITEM. Control totals come from Schedule 3: 135 sensor units and 53 relays across KTPH, TTSH and WH, i.e. about 2.5 sensors per relay (derived - confirm with engineering; it assumes 135 sensors equals the monitored beds). Coverage radius and redundancy policy are not in the attachments and must be supplied.");
    const st = [["1", "Map", "Mark beds, wall types and powered outlets on each ward floor plan"], ["2", "Cover", "Place relays so every sensor is within BLE coverage radius"], ["3", "Check capacity", "Keep sensors per relay within the planning density"], ["4", "Add resilience", "Add spare relays for critical coverage gaps"], ["5", "Verify", "Confirm 2.4 GHz Wi-Fi coverage and outlet at each position"]];
    st.forEach((p, i) => {
      const x = 0.6 + i * 2.45;
      s.addShape(pres.shapes.OVAL, { x: x + 0.05, y: 1.9, w: 0.55, h: 0.55, fill: { color: i % 2 ? C.accent1 : C.text2 }, line: { type: "none" } });
      T(s, p[0], { x: x + 0.05, y: 1.9, w: 0.55, h: 0.55, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle" });
      T(s, p[1], { x: x + 0.75, y: 1.9, w: 1.5, h: 0.55, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
      T(s, p[2], { x: x + 0.05, y: 2.6, w: 2.2, h: 0.95, fontSize: 11.5 });
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 3.85, w: 6.2, h: 2.7, fill: { color: C.text2 }, line: { type: "none" } });
    T(s, "Sizing rule (per ward)", { x: 0.85, y: 4.0, w: 5.7, h: 0.3, fontSize: 13, bold: true, color: "CADCFC" });
    T(s, "Relays = MAX( ⌈ Sensors ÷ Sensors per relay ⌉ ,  ⌈ Ward area ÷ Coverage area per relay ⌉ )  +  Spares", { x: 0.85, y: 4.4, w: 5.7, h: 1.1, fontSize: 15, bold: true, color: C.background1, fontFace: "Cambria" });
    T(s, "Control totals (contract draft): 135 sensors · 53 relays across KTPH, Woodlands Health and TTSH.", { x: 0.85, y: 5.6, w: 5.7, h: 0.8, fontSize: 12, color: "CADCFC" });
    table(s, [
      ["Design input", "Value"],
      ["Sensors per relay (planning density)", "≈ 2.5 (135 ÷ 53), to be verified by survey"],
      ["Relay power", "5 V USB-A, always-on: one outlet per relay"],
      ["Relay uplink", "Wi-Fi 2.4 GHz (5 GHz not supported)"],
      ["BLE coverage radius per relay", "[INPUT REQUIRED]"],
      ["Spare relay policy", "[INPUT REQUIRED]"],
    ], 7.1, 3.85, 5.63, [2.9, 2.73], { rowH: 0.45, fs: 10.5 });
  }

  // ================= 15. RELAY COUNT TABLE =================
  {
    const s = newSlide("4 · Understanding of the project", "Proposed relay quantity by hospital and ward (mandatory RFP response)", "Source: MONIT calculation from hospital floor plans; control totals from Master Agreement Schedule 3 (draft). Ward quantities completed once floor plans are in hand.",
      "MANDATORY: the RFP requires the proposed number of relay devices per ward. Counts are intentionally blank because the floor plans were not provided; do not invent per-ward numbers. Fill each ward with the sizing rule and reconcile the total to the price schedule in Section 3.");
    table(s, [
      ["Hospital", "Ward / area", "Sensors", "Relays (capacity)", "Relays (coverage)", "Proposed relays", "Outlet and 2.4 GHz confirmed"],
      ["KTPH", "[INPUT REQUIRED] Ward 1", "[ ]", "[ ]", "[ ]", "[ ]", "[ ]"],
      ["KTPH", "[INPUT REQUIRED] Ward 2", "[ ]", "[ ]", "[ ]", "[ ]", "[ ]"],
      ["Woodlands Health", "[INPUT REQUIRED] Ward 1", "[ ]", "[ ]", "[ ]", "[ ]", "[ ]"],
      ["Woodlands Health", "[INPUT REQUIRED] Ward 2", "[ ]", "[ ]", "[ ]", "[ ]", "[ ]"],
      ["TTSH", "[INPUT REQUIRED] Ward 1", "[ ]", "[ ]", "[ ]", "[ ]", "[ ]"],
      ["TTSH", "[INPUT REQUIRED] Ward 2", "[ ]", "[ ]", "[ ]", "[ ]", "[ ]"],
      ["Total (contract draft)", "", "135", "[ ]", "[ ]", "53", ""],
    ], 0.6, 1.85, 12.13, [1.9, 2.9, 1.0, 1.55, 1.55, 1.5, 1.73], { rowH: 0.46, fs: 10.5 });
    prop(s, 0.6, 5.75, 12.13, 0.85, "Spare relays shown as a separate line (not hidden in ward counts) so the quantity reconciles with the price schedule; ward rows added for every ward on the three floor plans.", { fs: 11 });
  }

  // ================= 16. GANTT =================
  {
    const s = newSlide("5 · Project timeline", "The plan runs from Letter of Award to Final Acceptance, with go-live in about six months", "Source: Master Agreement Schedule 2, Clause 12; RFP Section 1 clause 10. T0 = Letter of Award; lead times are proposed.",
      "T0 = Letter of Award (Effective Date). Milestones from the contract: training plan within 14 days, training within 14 days of delivery, 28-day mobilisation, deposit and insurance (14 days RFP vs 30 days Schedule 2). Production, shipping and customs lead times are NOT in the attachments - the 8-week figure is a placeholder proposal to be confirmed by MONIT operations. Schedule 2 draft LD is 0.5% per day (max 5%).");
    const lx = 0.6, lw = 3.7, gx = lx + lw, gw = 12.73 - gx, weeks = 26, ww = gw / weeks;
    for (let m = 0; m < 6; m++) {
      const x = gx + m * 4 * ww;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.8, w: 4 * ww, h: 0.32, fill: { color: m % 2 ? C.accent2 : C.text2 }, line: { color: H.white, width: 0.5 } });
      T(s, "M" + (m + 1), { x, y: 1.8, w: 4 * ww, h: 0.32, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle" });
    }
    s.addShape(pres.shapes.RECTANGLE, { x: gx + 24 * ww, y: 1.8, w: 2 * ww, h: 0.32, fill: { color: C.text2 }, line: { color: H.white, width: 0.5 } });
    T(s, "M7", { x: gx + 24 * ww, y: 1.8, w: 2 * ww, h: 0.32, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle" });
    const rows = [
      ["Award, deposit, insurance, training plan", 1, 2, C.text2],
      ["Mobilisation (≥ 28 days)", 1, 4, C.text2],
      ["Floor-plan review, Wi-Fi/RF survey, relay count", 3, 6, C.accent1],
      ["Production, pre-shipment test, shipping, customs", 4, 11, C.accent3],
      ["Pilot install and commissioning", 12, 13, C.accent1],
      ["User training (≤ 14 days after delivery)", 12, 15, C.accent1],
      ["Pilot acceptance tests", 14, 15, C.accent1],
      ["Ward-by-ward rollout", 15, 22, C.text2],
      ["Final Acceptance and invoice (≤ 7 days)", 22, 24, C.accent5],
      ["Hypercare and optimisation", 23, 26, C.accent5],
    ];
    rows.forEach((r, i) => {
      const y = 2.22 + i * 0.385;
      if (i % 2 === 0) s.addShape(pres.shapes.RECTANGLE, { x: lx, y: y - 0.02, w: 12.13, h: 0.385, fill: { color: C.background2 }, line: { type: "none" } });
      T(s, r[0], { x: lx + 0.1, y, w: lw - 0.15, h: 0.34, fontSize: 11, valign: "middle" });
      s.addShape(pres.shapes.RECTANGLE, { x: gx + (r[1] - 1) * ww, y: y + 0.05, w: (r[2] - r[1] + 1) * ww, h: 0.25, fill: { color: r[3] }, line: { type: "none" }, objectName: "Gantt " + r[0] });
    });
    prop(s, 0.6, 6.15, 12.13, 0.5, "Lead time: production, quality test, air freight and customs about 8 weeks from order; installation about 2 weeks per hospital pilot; hypercare 4 weeks. To be confirmed at award.", { fs: 10.5 });
  }

  // ================= 17. REFERENCES =================
  {
    const s = newSlide("6 · References and track record", "Each reference shows scale, scope, results and a contact KTPH can call", "Source: MONIT customer references. Written consent from each referee is obtained before they are named.",
      "No references are in the attachments. Provide at least three comparable deployments. The teal box describes the reference mix and checks hospital procurement usually applies; replace with facts.");
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.1;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.85, w: 3.95, h: 3.45, fill: { color: C.background2 }, line: { type: "none" } });
      await iconCircle(s, "FaHospital", x + 0.2, 2.0, 0.5, i % 2 ? C.accent1 : C.text2);
      T(s, "Reference " + (i + 1), { x: x + 0.85, y: 2.0, w: 2.9, h: 0.5, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
      inp(s, x + 0.2, 2.7, 3.55, 0.5, "Customer and site", { fs: 10 });
      inp(s, x + 0.2, 3.28, 3.55, 0.5, "Beds, wards, devices installed", { fs: 10 });
      inp(s, x + 0.2, 3.86, 3.55, 0.5, "Period and measured outcome", { fs: 10 });
      inp(s, x + 0.2, 4.44, 3.55, 0.7, "Referee name, role, contact (consent obtained)", { fs: 10 });
    }
    prop(s, 0.6, 5.5, 12.13, 1.1, "Reference mix: (1) an acute hospital ward deployment, (2) a long-term-care or nursing-home deployment, (3) a deployment integrated with hospital Wi-Fi. For each, MONIT provides scale, go-live date, measured results and a referee contactable by KTPH procurement.", { fs: 12 });
  }

  // ================= 18. SUSTAINABILITY =================
  {
    const s = newSlide("7 · Sustainability initiatives", "MONIT answers each of the four sustainability criteria with evidence in the format KTPH procurement asks for", "Source: RFP sustainability criteria; ALPS ISO14000-certified environment (Section 1, clause 19); product facts from Master Agreement Schedule 3.",
      "No sustainability facts are in the attachments. Each quadrant has a PROPOSED line describing the evidence procurement usually expects, and an amber box for MONIT's actual position. Do not state any claim until evidenced. Product-level facts (reusable 23 g sensor, single-use barrier film, CR2032 cell, 5 V USB relay) are from Schedule 3.");
    const q = [
      ["FaSolarPanel", "Renewable energy", "Share of operations on renewable energy, e.g. solar kWh per year or green-power contract."],
      ["FaBolt", "Energy tracking", "Annual energy use of offices and factories, how it is metered, and any reduction target."],
      ["FaRecycle", "Environmental policy", "Written policy or ISO 14001; waste, packaging, battery and e-waste handling. Product: reusable sensor with single-use hygiene film."],
      ["FaLeaf", "Green products", "Green Label or energy-label products used in offices and sites; low-power 5 V USB relay."],
    ];
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + (i % 2) * 6.15, y = 1.85 + Math.floor(i / 2) * 2.45;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 5.98, h: 2.3, fill: { color: C.background2 }, line: { type: "none" } });
      await iconCircle(s, q[i][0], x + 0.2, y + 0.15, 0.5, C.accent5);
      T(s, q[i][1], { x: x + 0.85, y: y + 0.15, w: 4.9, h: 0.5, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
      prop(s, x + 0.2, y + 0.78, 5.58, 0.8, "Evidence: " + q[i][2], { fs: 10 });
      inp(s, x + 0.2, y + 1.65, 5.58, 0.5, "MONIT's current position and supporting document", { fs: 10 });
    }
  }

  // ================= 19. NETWORK ARCHITECTURE =================
  {
    const s = newSlide("8 · Network connectivity and relay devices", "The relay joins the existing hospital Wi-Fi as a 2.4 GHz client; the hospital supplies coverage and a power outlet", "Source: Master Agreement Schedule 3 (relay specification); MONIT. IMDA registration per Schedule 2, Clause 5.",
      "Relay: 42 g, 5 V USB-A always-on, BLE 5 + Wi-Fi 2.4 GHz (5 GHz not supported). The 2.4 GHz-only radio is a real dependency - confirm with each hospital IT team that a 2.4 GHz SSID exists with coverage at relay locations. Bring a physical relay to the meeting if possible.");
    const nodes = [["FaMicrochip", "Sensors", "BLE, in ward beds"], ["FaWifi", "Relay device", "BLE 5 in · Wi-Fi 2.4 GHz out"], ["FaNetworkWired", "Hospital Wi-Fi", "Owned and managed by hospital"], ["FaCloud", "MONIT cloud", "Servers, AI, storage"], ["FaDesktop", "Dashboard and app", "Nurse station · iOS / Android"]];
    for (let i = 0; i < 5; i++) {
      const x = 0.6 + i * 2.45;
      const hosp = i === 2;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.9, w: 2.2, h: 1.9, fill: { color: hosp ? C.background1 : C.background2 }, line: { color: hosp ? H.teal : H.light, width: hosp ? 1.5 : 0.5, dashType: hosp ? "dash" : "solid" } });
      await iconCircle(s, nodes[i][0], x + 0.8, 2.05, 0.6, i % 2 ? C.accent1 : C.text2);
      T(s, nodes[i][1], { x: x + 0.1, y: 2.8, w: 2.0, h: 0.35, fontSize: 13, bold: true, color: C.text2, align: "center" });
      T(s, nodes[i][2], { x: x + 0.1, y: 3.2, w: 2.0, h: 0.5, fontSize: 11, align: "center", color: C.accent4 });
      if (i < 4) s.addShape(pres.shapes.LINE, { x: x + 2.2, y: 2.85, w: 0.25, h: 0, line: { color: H.teal, width: 2, endArrowType: "triangle" } });
    }
    const f = [["Standard Wi-Fi client", "Connects to the hospital's existing 2.4 GHz network; 5 GHz is not used. The hospital provides 2.4 GHz coverage at relay locations."], ["Simple power", "5 V USB-A, always-on: one powered outlet per relay, no new cabling to the network."], ["IMDA registered", "Relay registered for use in Singapore (Master Agreement Schedule 2). Registration no.: [INPUT REQUIRED]"]];
    f.forEach((c, i) => {
      const x = 0.6 + i * 4.1;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 4.2, w: 3.95, h: 2.35, fill: { color: C.background2 }, line: { type: "none" } });
      T(s, c[0], { x: x + 0.2, y: 4.35, w: 3.55, h: 0.35, fontSize: 14, bold: true, color: C.text2 });
      T(s, c[1], { x: x + 0.2, y: 4.8, w: 3.55, h: 1.6, fontSize: 12 });
    });
  }

  // ================= 20. SETUP =================
  {
    const s = newSlide("8 · Network connectivity and relay devices", "Setup takes minutes per relay: mount, power, connect once, pair, verify", "Source: MONIT. Installation and commissioning by MONIT per Master Agreement Schedule 3.",
      "Provisioning method is NOT described in the attachments; the teal box is a proposed answer. Engineering must confirm the actual method (local app, web page, or pre-staged settings) and who performs it.");
    const st = [["FaPlug", "1  Mount and power", "Plug the relay into a 5 V USB outlet at the planned position"], ["FaKey", "2  Connect once", "Hospital IT enters the Wi-Fi name and password during setup — MONIT never holds the password"], ["FaLink", "3  Pair sensors", "Sensors pair with the relay over BLE"], ["FaCloud", "4  Register", "Relay registers with the MONIT cloud"], ["FaCheckCircle", "5  Verify", "Signal and connectivity checked on the dashboard, recorded for acceptance"]];
    for (let i = 0; i < 5; i++) {
      const x = 0.6 + i * 2.45;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.9, w: 2.2, h: 2.85, fill: { color: C.background2 }, line: { type: "none" } });
      await iconCircle(s, st[i][0], x + 0.8, 2.05, 0.6, i === 1 ? C.accent1 : C.text2);
      T(s, st[i][1], { x: x + 0.1, y: 2.8, w: 2.0, h: 0.35, fontSize: 13, bold: true, color: C.text2, align: "center" });
      T(s, st[i][2], { x: x + 0.15, y: 3.25, w: 1.9, h: 1.4, fontSize: 11, align: "center" });
    }
    prop(s, 0.6, 4.9, 12.13, 0.75, "Credentials are keyed in once through a secure local setup step by hospital IT; MONIT records only that setup passed, not the password.", { fs: 12 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 5.8, w: 12.13, h: 0.85, fill: { color: C.text2 }, line: { type: "none" } });
    T(s, [{ text: "Hospital IT: ", options: { bold: true } }, { text: "provides Wi-Fi name and credentials and any approvals.   " }, { text: "MONIT: ", options: { bold: true } }, { text: "configures relays, supports hospital IT on site, and records the setup for acceptance." }], { x: 0.85, y: 5.8, w: 11.6, h: 0.85, fontSize: 12, color: C.background1, valign: "middle" });
  }

  // ================= 21. PASSWORD PERSISTENCE =================
  {
    const s = newSlide("8 · Network connectivity and relay devices", "The Wi-Fi password is stored once on the relay and is not asked for again after a power outage or firmware upgrade", "Source: MONIT proposed answer to the RFP network question. Behaviour to be confirmed against the relay firmware before issue.",
      "KEY ANSWER TO ALPS'S QUESTION. The attachments do not describe how the relay stores Wi-Fi credentials or handles OTA. The table is the standard design pattern: non-volatile storage, configuration kept separate from firmware, rollback on failed update, automatic reconnect. Because the relay is powered by always-on 5 V USB, a power cut simply restarts it. Engineering must confirm every row.");
    table(s, [
      ["Event", "What happens on the relay", "Password re-entry?"],
      ["Power outage / restart", "Relay restarts when power returns, reloads the saved credentials and reconnects automatically", "No"],
      ["Firmware upgrade", "Settings are kept separate from firmware and carried through; failed updates roll back", "No"],
      ["Hospital changes its Wi-Fi password", "New password entered once by hospital IT", "Yes — once, by hospital IT"],
      ["Replacement or factory-reset relay", "New unit set up like the first install", "Yes — one-time setup"],
    ], 0.6, 1.85, 7.2, [2.2, 3.5, 1.5], { rowH: 0.7, fs: 11 });
    T(s, "How the solution achieves this", { x: 8.2, y: 1.85, w: 4.5, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
    const m = [["FaDatabase", "Credentials kept in non-volatile storage"], ["FaSyncAlt", "Settings separated from firmware and preserved on upgrade"], ["FaRedo", "Automatic reconnection after any outage"], ["FaCloudUploadAlt", "Failed updates roll back to the last working version"]];
    for (let i = 0; i < 4; i++) {
      const y = 2.3 + i * 0.77;
      await iconCircle(s, m[i][0], 8.2, y, 0.5, i % 2 ? C.accent1 : C.text2);
      T(s, m[i][1], { x: 8.95, y, w: 3.8, h: 0.55, fontSize: 12.5, valign: "middle" });
    }
    prop(s, 0.6, 5.5, 12.13, 0.5, "Answer written for KTPH's stated expectation: no password re-entry after any outage or upgrade.", { fs: 11 });
    inp(s, 0.6, 6.1, 12.13, 0.5, "Engineering to confirm storage method, upgrade behaviour and rollback against the actual firmware.", { fs: 11 });
  }

  // ================= 22. SECURITY & OPS =================
  {
    const s = newSlide("8 · Network connectivity and relay devices", "The design keeps hospital networks and patient data protected and the service observable", "Source: Master Agreement Schedules 7, 8 and 11; RFP Consent Form Annex A (IT Security Compliance List, Third-Party PDPA checklist, SaaS security requirements).",
      "Schedule 7 forbids transferring the Company's personal data outside Singapore without prior written consent. MONIT is Korea-based, so the hosting location of the cloud server must be stated and consistent with Schedule 7 and the PDPA checklist.");
    await card(s, 0.6, 1.85, 3.9, 3.4, { icon: "FaLock", head: "Network security", body: ["Connects as a normal Wi-Fi client under hospital policy", "BLE between sensor and relay", "Encrypted connection to the cloud [confirm]", "Relay identified by hardware address for hospital approval"] });
    await card(s, 4.72, 1.85, 3.9, 3.4, { icon: "FaUserShield", head: "Data protection", body: ["PDPA terms of the Master Agreement (Schedule 7) accepted", "Personal data not moved outside Singapore without written consent", "Role-based access to the dashboard", "Cloud hosting location: [INPUT REQUIRED]"] });
    await card(s, 8.83, 1.85, 3.9, 3.4, { icon: "FaHeartbeat", head: "Operational monitoring", body: ["Relay online/offline status on the dashboard", "Remote diagnostics by approved engineers", "Upgrades scheduled with the hospital", "Cybersecurity terms of Schedule 11 accepted"] });
    inp(s, 0.6, 5.55, 12.13, 0.85, "Status of IT Security Compliance List, Third-Party PDPA checklist and SaaS security requirements (Annex A) — attach completed forms.");
  }

  // ================= 23. OPEN QUESTIONS =================
  {
    const s = newSlide("Discussion", "Eight confirmations from ALPS and the hospitals will fix the final design, service levels and price", "Source: MONIT analysis of RFP documents and Master Agreement.",
      "Use this slide to drive the meeting. Items 4-6 reflect inconsistencies between the RFP, Section 3.1 and the Master Agreement draft (spare parts 7 vs 10 years; deposit 14 vs 30 days; blank/draft service-level values).");
    table(s, [
      ["#", "Confirmation needed", "Why it matters"],
      ["1", "Wards, beds in scope and floor plans for each hospital", "Sets relay and sensor quantities"],
      ["2", "2.4 GHz SSID available? Authentication type, VLAN, MAC approval, IT contact", "Relay radio is 2.4 GHz only"],
      ["3", "Powered outlet (5 V USB) and mounting at relay positions", "Relay is USB powered, always-on"],
      ["4", "Spare-parts period: 7 years from delivery or 10 years after end-of-life?", "Contract and service pricing"],
      ["5", "Security deposit and insurance timing (14 days RFP; 30 days Schedule 2)", "Mobilisation plan"],
      ["6", "Final service-level values (response, repair, loan unit, credits)", "SLA and maintenance price"],
      ["7", "Nurse-call / EMR integration; data hosting location", "Scope and PDPA review"],
      ["8", "Installation windows and infection-control rules per ward", "Rollout schedule"],
    ], 0.6, 1.85, 12.13, [0.6, 7.5, 4.03], { rowH: 0.52, fs: 11.5 });
  }

  // ================= 24. NEXT STEPS =================
  {
    const s = newSlide("Discussion", "KTPH can verify every claim in this document through a demo, reference calls and certificate copies", "Source: MONIT. Owners and dates to be agreed in the meeting.",
      "Closing slide for the procurement audience: make verification easy. Fill owners and dates during the meeting. Reminder of submission requirements: compliance tables stamped and signed, Annex A checklist, certificates, price proposal in Ariba, security deposit and insurance after award.");
    table(s, [
      ["Verification step", "Owner", "Date"],
      ["Live demonstration of sensor, relay, dashboard and app", "MONIT", "[INPUT]"],
      ["Relay set-up and Wi-Fi reconnect test with hospital IT", "MONIT / hospital IT", "[INPUT]"],
      ["Reference calls with customers", "MONIT", "[INPUT]"],
      ["Certificate copies (CE, IMDA, KC, ISO) and clinical reports", "MONIT", "[INPUT]"],
      ["Ward-by-ward relay table completed from floor plans", "MONIT / ALPS", "[INPUT]"],
    ], 0.6, 1.85, 8.0, [4.8, 1.9, 1.3], { rowH: 0.65, fs: 12 });
    s.addShape(pres.shapes.RECTANGLE, { x: 8.9, y: 1.85, w: 3.83, h: 4.35, fill: { color: C.text2 }, line: { type: "none" } });
    T(s, "What KTPH can rely on", { x: 9.15, y: 2.05, w: 3.4, h: 0.4, fontSize: 15, bold: true, color: C.background1 });
    T(s, bullets(["One accountable vendor for hardware, software, installation and service", "Relay count tied to your floor plans and verified by survey", "Defined acceptance steps and named engineers", "Evidence available for every statement"]), { x: 9.15, y: 2.6, w: 3.4, h: 3.4, fontSize: 13, color: C.background1 });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
}
main().catch((e) => { console.error(e); process.exit(1); });
