// Build: NODE_PATH=<dir with node_modules> node build_deck.js
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const SKILL = "/root/.claude/skills/synced/740b1519-7df3-42c4-87b2-972ad2eb207a_a3d7b2ae-009d-4d22-98fe-59c0c04fc335/pptx";
const { applyTheme } = require(SKILL + "/scripts/apply_theme.js");

const OUT = "KTPH_RFP_MONIT_Presentation.pptx";
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
const H = { navy: "0B2A4A", teal: "00A3AD", grey: "6B7A8F", light: "F1F4F8", line: "D5DCE6", amber: "F2A900", amberBg: "FFF4D6", amberTx: "7A4F00", red: "C0392B", green: "2E7D6B", text: "1F2937", white: "FFFFFF" };

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
    const s = newSlide("Executive summary", "MONIT offers a validated, Wi-Fi-ready smart diaper system deployable across all three hospitals", "Source: ALPS RFP KTPH-RFP-26-165-MJ; MONIT analysis. Items in amber require MONIT inputs.",
      "Key messages. Each row should be validated by MONIT product, clinical and operations leads before the meeting.");
    const rows = [
      ["1", "Proven technology", "Clinically validated sensor with measured detection performance; certified to the standards ALPS expects (BizSAFE, ISO, IMDA, HSA)."],
      ["2", "Complete, integrated system", "Sensor → relay → hospital Wi-Fi → cloud → nurse dashboard, delivered as one fully integrated system as the RFP requires."],
      ["3", "Sized to your wards", "Relay quantities derived per ward from the floor plans of KTPH, Woodlands Health and TTSH using a transparent coverage-based method."],
      ["4", "Zero-touch Wi-Fi operations", "Hospital Wi-Fi credentials are entered once; they survive power outages and firmware upgrades without re-keying."],
      ["5", "Low-risk delivery", "Phased plan aligned to the 28-day mobilisation period, pilot-first rollout, hypercare, and an SLA-backed service model."],
    ];
    let y = 1.85;
    for (const [n, h, d] of rows) {
      s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.6, h: 0.6, fill: { color: C.text2 }, line: { type: "none" } });
      T(s, n, { x: 0.6, y, w: 0.6, h: 0.6, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle", fontFace: "Cambria" });
      T(s, h, { x: 1.5, y, w: 3.2, h: 0.6, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
      T(s, d, { x: 4.8, y, w: 7.9, h: 0.6, fontSize: 13, valign: "middle" });
      if (n !== "5") s.addShape(pres.shapes.LINE, { x: 0.6, y: y + 0.82, w: 12.13, h: 0, line: { color: H.line, width: 0.75 } });
      y += 0.97;
    }
  }

  // ================= 3. RFP AT A GLANCE =================
  {
    const s = newSlide("Executive summary", "One system, three public hospitals: firm pricing, mandatory relay sizing and a vendor presentation", "Source: RFP Consent Form; Section 1 – Conditions of RFP (clauses 5, 8, 10, 13, 20). Dates as stated in the RFP documents.",
      "Facts taken directly from the RFP consent form and Section 1. Note that several dates (site briefings, clarifications, closing) may already have passed depending on the meeting date - confirm the current stage with ALPS.");
    // key dates
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
    // stat callouts
    const stats = [["180", "days price validity from closing date"], ["10%", "security deposit within 14 days of acceptance"], ["28", "days minimum to mobilise resources"], ["60", "days payment term"]];
    stats.forEach(([n, l], i) => {
      const x = 6.6 + (i % 2) * 3.1, yy = 1.85 + Math.floor(i / 2) * 1.6;
      s.addShape(pres.shapes.RECTANGLE, { x, y: yy, w: 2.95, h: 1.45, fill: { color: C.background2 }, line: { type: "none" } });
      T(s, n, { x: x + 0.2, y: yy + 0.12, w: 2.5, h: 0.75, fontSize: 40, bold: true, color: C.accent1, fontFace: "Cambria" });
      T(s, l, { x: x + 0.2, y: yy + 0.88, w: 2.55, h: 0.5, fontSize: 11, color: C.accent4 });
    });
    // sites
    T(s, "Three sites in scope", { x: 6.6, y: 5.1, w: 6, h: 0.3, fontSize: 14, bold: true, color: C.text2 });
    ["Khoo Teck Puat Hospital", "Woodlands Health", "Tan Tock Seng Hospital"].forEach((n, i) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.6 + i * 2.07, y: 5.5, w: 1.95, h: 0.8, rectRadius: 0.08, fill: { color: C.text2 }, line: { type: "none" } });
      T(s, n, { x: 6.7 + i * 2.07, y: 5.5, w: 1.75, h: 0.8, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
    });
    T(s, "SGD, ex-GST  •  HSA-registered devices  •  Fully integrated system", { x: 0.6, y: 5.5, w: 5.6, h: 0.5, fontSize: 12, color: C.accent4 });
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
    const s = newSlide("2 · Industry certifications and standards", "Our certifications cover safety, quality, security and the network-device rules for hospital Wi-Fi", "Source: MONIT certificates; RFP Consent Form (HSA registration requirement); Section 1 clause 19 (ISO14000/OSHA). Attach copies in the proposal.",
      "Provide certificate numbers, issuing bodies and expiry dates, and append copies. Note that ALPS is ISO14000 and OSHA certified and expects vendors to follow its environmental and safety requirements. The consent form requires medical devices to be HSA-registered - confirm the device class and registration for the sensor.");
    table(s, [
      ["Standard / certification", "Why it matters for this RFP", "Status / certificate no.", "Valid until"],
      ["BizSAFE (level 3 / Star)", "Contractor safety for on-site installation in hospitals", "[INPUT REQUIRED] level & cert no.", "[INPUT]"],
      ["ISO 9001 – Quality management", "Consistent delivery, QA and service processes", "[INPUT REQUIRED]", "[INPUT]"],
      ["ISO 13485 – Medical devices QMS", "Design & manufacture of the sensor as a medical device", "[INPUT REQUIRED]", "[INPUT]"],
      ["ISO 27001 – Information security", "Protects patient-related data in cloud and dashboard", "[INPUT REQUIRED]", "[INPUT]"],
      ["ISO 14001 / ISO 45001", "Aligns with ALPS's ISO14000 / OSHA requirements", "[INPUT REQUIRED]", "[INPUT]"],
      ["IMDA – network devices (relay)", "Relay / Wi-Fi radio equipment lawfully connected in Singapore", "[INPUT REQUIRED] registration / label no.", "[INPUT]"],
      ["HSA – medical device registration", "RFP requires registration for Class B/C/D (before 1 May 2010) and Class A (before 1 May 2011) — confirm applicability", "[INPUT REQUIRED] class & reg. no.", "[INPUT]"],
    ], 0.6, 1.85, 12.13, [3.3, 4.9, 2.7, 1.23], { rowH: 0.56, fs: 11 });
  }

  // ================= 6. CONCEPT =================
  {
    const s = newSlide("3 · Smart diaper system", "The system turns every diaper change into a timely, data-driven decision — detect, notify, document", "Source: MONIT. Architecture is indicative; confirm radio protocol and data path with the product team.",
      "Walk through the chain left-to-right. Confirm which radio the sensor uses to talk to the relay (e.g. BLE / sub-GHz) and whether the relay posts directly to cloud or an on-prem server.");
    const steps = [
      ["FaMicrochip", "Sensor", "Clips onto the diaper; detects urination / stool events"],
      ["FaWifi", "Relay device", "Collects sensor data in the ward and forwards via hospital Wi-Fi"],
      ["FaCloud", "Platform", "AI detection, event history, secure storage"],
      ["FaUserNurse", "Nurse alert", "Real-time alert to station dashboard / mobile"],
      ["FaChartLine", "Insights", "Reports on change intervals, workload, skin-care trends"],
    ];
    for (let i = 0; i < steps.length; i++) {
      const x = 0.6 + i * 2.45;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.9, w: 2.2, h: 2.5, fill: { color: C.background2 }, line: { type: "none" } });
      await iconCircle(s, steps[i][0], x + 0.8, 2.05, 0.6, i % 2 ? C.accent1 : C.text2);
      T(s, steps[i][1], { x: x + 0.1, y: 2.8, w: 2.0, h: 0.35, fontSize: 14, bold: true, color: C.text2, align: "center" });
      T(s, steps[i][2], { x: x + 0.15, y: 3.2, w: 1.9, h: 1.1, fontSize: 11, align: "center" });
      if (i < steps.length - 1) s.addShape(pres.shapes.LINE, { x: x + 2.2, y: 2.35, w: 0.25, h: 0, line: { color: H.teal, width: 2, endArrowType: "triangle" } });
    }
    T(s, "Value to hospitals", { x: 0.6, y: 4.7, w: 6, h: 0.3, fontSize: 14, bold: true, color: C.text2 });
    T(s, bullets(["Right-time diaper changes instead of routine checks", "Less nurse time on manual rounds; fewer disturbances for patients", "Objective data to support skin-integrity and continence care"]), { x: 0.6, y: 5.05, w: 6, h: 1.5, fontSize: 13 });
    inp(s, 7.0, 4.7, 5.73, 1.85, "Product photo / system diagram, sensor specs (size, weight, battery life, diaper compatibility), relay specs.");
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
      "Replace the wireframe with real screenshots. Confirm: user roles, audit log, mobile alerts, EMR/nurse-call integration options, data hosting location (Singapore), PDPA compliance.");
    // wireframe
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
    // functions
    const f = [["FaBell", "Real-time alerts", "Nurse station and mobile notification with escalation rules"], ["FaListAlt", "Care records", "Event history per bed; exportable for documentation"], ["FaChartBar", "Analytics", "Change intervals, response times, ward workload trends"], ["FaUserLock", "Admin & security", "Role-based access, audit trail, PDPA-aligned data handling"]];
    for (let i = 0; i < f.length; i++) {
      const y = 1.85 + i * 1.18;
      await iconCircle(s, f[i][0], 7.3, y + 0.1, 0.55, i % 2 ? C.accent1 : C.text2);
      T(s, f[i][1], { x: 8.1, y, w: 4.6, h: 0.35, fontSize: 14, bold: true, color: C.text2 });
      T(s, f[i][2], { x: 8.1, y: y + 0.37, w: 4.6, h: 0.65, fontSize: 12 });
    }
  }

  // ================= 9. MAINTENANCE & SUPPORT =================
  {
    const s = newSlide("3 · Smart diaper system", "A dedicated service model keeps the system running with defined response times, spares and training", "Source: MONIT. SLA values to be aligned with Section 2 Master Agreement (maintenance obligations) and Section 3 Technical Specifications.",
      "Align the SLA with the Master Agreement for Supply and Maintenance of Equipment. Provide response/resolution times, support hours, spares holding, escalation path, local Singapore support presence.");
    await card(s, 0.6, 1.85, 3.9, 2.35, { icon: "FaTools", head: "Maintenance", body: ["Preventive checks & remote health monitoring", "Replacement of faulty sensors / relays", "Spares pool held locally [confirm]"] });
    await card(s, 4.72, 1.85, 3.9, 2.35, { icon: "FaHeadset", head: "Technical support", body: ["Helpdesk hours: [INPUT]", "Remote diagnostics via platform", "Named account / service manager"] });
    await card(s, 8.83, 1.85, 3.9, 2.35, { icon: "FaChalkboardTeacher", head: "Training & onboarding", body: ["Nurse & IT admin training", "Quick-reference guides", "On-site support during go-live"] });
    table(s, [
      ["Severity", "Definition", "Response", "Resolution / workaround"],
      ["P1 – Critical", "Ward-wide outage / no alerts", "[INPUT]", "[INPUT]"],
      ["P2 – Major", "Multiple beds affected", "[INPUT]", "[INPUT]"],
      ["P3 – Minor", "Single device / cosmetic", "[INPUT]", "[INPUT]"],
    ], 0.6, 4.45, 12.13, [2.2, 4.4, 2.5, 3.03], { rowH: 0.45 });
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
    const s = newSlide("4 · Understanding of the project", "A four-phase approach de-risks deployment: plan, pilot, roll out, optimise", "Source: MONIT proposed methodology. Durations on the timeline slide.",
      "Methodology. Emphasise the pilot ward per hospital and the Wi-Fi site survey before ordering relays.");
    const ph = [
      ["1", "Plan", ["Kick-off & governance", "Wi-Fi / RF site survey with hospital IT", "Final relay count & placement", "Security & PDPA review"]],
      ["2", "Pilot", ["Install in 1 pilot ward / hospital", "Credential provisioning once", "Staff training", "UAT & acceptance"]],
      ["3", "Roll out", ["Ward-by-ward installation", "Minimise disruption to care", "Daily progress tracking", "Sign-off per ward"]],
      ["4", "Optimise", ["Hypercare period", "Alert-rule tuning", "Usage & benefit reporting", "Handover to service team"]],
    ];
    ph.forEach((p, i) => {
      const x = 0.6 + i * 3.05;
      s.addShape(pres.shapes.PENTAGON || pres.shapes.RECTANGLE, { x, y: 1.9, w: 3.0, h: 0.9, fill: { color: i % 2 ? C.accent1 : C.text2 }, line: { type: "none" } });
      T(s, p[0] + "  " + p[1], { x: x + 0.25, y: 1.9, w: 2.3, h: 0.9, fontSize: 18, bold: true, color: C.background1, valign: "middle", fontFace: "Cambria" });
      s.addShape(pres.shapes.RECTANGLE, { x, y: 3.0, w: 2.85, h: 2.6, fill: { color: C.background2 }, line: { type: "none" } });
      T(s, bullets(p[2]), { x: x + 0.2, y: 3.2, w: 2.5, h: 2.3, fontSize: 13 });
    });
  }

  // ================= 13. DEPLOYMENT PER HOSPITAL =================
  {
    const s = newSlide("4 · Understanding of the project", "Deployment is planned ward by ward from each hospital's floor plan, with the same design pattern at all three sites", "Source: Hospital floor plans (provided by ALPS at site briefings) — not included in the files received by MONIT for this document.",
      "Insert the floor plans for each hospital and mark relay positions. The floor plans were not part of the files supplied for this draft.");
    const hs = [["Khoo Teck Puat Hospital", "FaHospital"], ["Woodlands Health", "FaHospital"], ["Tan Tock Seng Hospital", "FaHospital"]];
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.1;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.85, w: 3.95, h: 0.5, fill: { color: C.text2 }, line: { type: "none" } });
      T(s, hs[i][0], { x: x + 0.15, y: 1.85, w: 3.7, h: 0.5, fontSize: 14, bold: true, color: C.background1, valign: "middle" });
      inp(s, x, 2.45, 3.95, 2.3, "Floor plan with ward boundaries and proposed relay positions marked.", { fs: 11 });
      T(s, bullets(["Wards in scope: [INPUT]", "Beds in scope: [INPUT]", "Relay positions: see next slides", "Install window per ward: [INPUT]"]), { x, y: 4.9, w: 3.95, h: 1.7, fontSize: 12 });
    }
  }

  // ================= 14. RELAY SIZING METHOD =================
  {
    const s = newSlide("4 · Understanding of the project", "Relay quantities follow a transparent, coverage-based method that is verified by an on-site survey", "Source: MONIT sizing method. Coverage radius and beds-per-relay are assumptions to be confirmed by the product / RF team.",
      "MANDATORY RFP ITEM. The method must be defensible. Confirm the real coverage radius per relay, max sensors per relay, wall-attenuation assumptions and redundancy policy with engineering.");
    const st = [["1", "Map", "Mark bed locations & wall types on each ward floor plan"], ["2", "Cover", "Place relays so every bed lies within the validated coverage radius"], ["3", "Check capacity", "Ensure sensors per relay stays below the maximum"], ["4", "Add resilience", "Add redundancy for critical coverage gaps (e.g. 10% spare)"], ["5", "Verify", "Confirm with Wi-Fi/RF survey and pilot ward"]];
    st.forEach((p, i) => {
      const x = 0.6 + i * 2.45;
      s.addShape(pres.shapes.OVAL, { x: x + 0.05, y: 1.9, w: 0.55, h: 0.55, fill: { color: i % 2 ? C.accent1 : C.text2 }, line: { type: "none" } });
      T(s, p[0], { x: x + 0.05, y: 1.9, w: 0.55, h: 0.55, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle" });
      T(s, p[1], { x: x + 0.75, y: 1.9, w: 1.5, h: 0.55, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
      T(s, p[2], { x: x + 0.05, y: 2.6, w: 2.2, h: 0.95, fontSize: 11.5 });
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 3.85, w: 6.2, h: 2.7, fill: { color: C.text2 }, line: { type: "none" } });
    T(s, "Sizing rule (per ward)", { x: 0.85, y: 4.0, w: 5.7, h: 0.3, fontSize: 13, bold: true, color: "CADCFC" });
    T(s, "Relays = MAX( ⌈ Beds ÷ Max sensors per relay ⌉ ,  ⌈ Ward area ÷ Coverage area per relay ⌉ )  +  Redundancy", { x: 0.85, y: 4.4, w: 5.7, h: 1.1, fontSize: 16, bold: true, color: C.background1, fontFace: "Cambria" });
    T(s, "Rounded up per ward; adjusted after the Wi-Fi/RF survey.", { x: 0.85, y: 5.6, w: 5.7, h: 0.4, fontSize: 12, color: "CADCFC" });
    table(s, [
      ["Design assumption", "Value"],
      ["Coverage radius per relay", "[INPUT REQUIRED]"],
      ["Max sensors per relay", "[INPUT REQUIRED]"],
      ["Wall / obstruction derating", "[INPUT REQUIRED]"],
      ["Redundancy policy", "[INPUT REQUIRED]"],
    ], 7.1, 3.85, 5.63, [3.4, 2.23], { rowH: 0.52 });
  }

  // ================= 15. RELAY COUNT TABLE =================
  {
    const s = newSlide("4 · Understanding of the project", "Proposed relay quantity by hospital and ward (mandatory RFP response)", "Source: MONIT calculation from hospital floor plans. Quantities to be completed once floor plans and design assumptions are confirmed.",
      "MANDATORY: the RFP requires the proposed number of relay devices per ward. Counts are intentionally blank because the floor plans were not provided with the files. Use the sizing rule on the previous slide, fill each ward row and sum the totals.");
    table(s, [
      ["Hospital", "Ward / area", "Beds", "Area (m²)", "Relays (coverage)", "Relays (capacity)", "Proposed relays (incl. spare)"],
      ["KTPH", "[INPUT REQUIRED] Ward 1", "[ ]", "[ ]", "[ ]", "[ ]", "[ ]"],
      ["KTPH", "[INPUT REQUIRED] Ward 2", "[ ]", "[ ]", "[ ]", "[ ]", "[ ]"],
      ["Woodlands Health", "[INPUT REQUIRED] Ward 1", "[ ]", "[ ]", "[ ]", "[ ]", "[ ]"],
      ["Woodlands Health", "[INPUT REQUIRED] Ward 2", "[ ]", "[ ]", "[ ]", "[ ]", "[ ]"],
      ["TTSH", "[INPUT REQUIRED] Ward 1", "[ ]", "[ ]", "[ ]", "[ ]", "[ ]"],
      ["TTSH", "[INPUT REQUIRED] Ward 2", "[ ]", "[ ]", "[ ]", "[ ]", "[ ]"],
      ["Total", "", "[ ]", "[ ]", "[ ]", "[ ]", "[ ]"],
    ], 0.6, 1.85, 12.13, [1.9, 3.1, 1.0, 1.2, 1.6, 1.6, 1.73], { rowH: 0.5 });
    inp(s, 0.6, 6.05, 12.13, 0.6, "Add one row per ward from the three floor plans; confirm quantity per ward with the sizing rule.");
  }

  // ================= 16. GANTT =================
  {
    const s = newSlide("5 · Project timeline", "Implementation can start on award and reach full go-live in about six months", "Source: MONIT indicative plan; RFP Section 1 clause 10 (10% deposit within 14 days; at least 28 days to mobilise). T0 = Letter of Acceptance. Lead times to be confirmed.",
      "Indicative plan. Replace lead times with real figures: sensor/relay manufacturing and shipment, IMDA/HSA paperwork, installation capacity. Week 1 = week of award.");
    const lx = 0.6, lw = 3.3, gx = lx + lw, gw = 12.73 - gx, weeks = 26, ww = gw / weeks;
    // header months
    for (let m = 0; m < 7; m++) {
      const x = gx + m * 4 * ww;
      if (m < 6) { s.addShape(pres.shapes.RECTANGLE, { x, y: 1.85, w: 4 * ww, h: 0.35, fill: { color: m % 2 ? C.accent2 : C.text2 }, line: { color: H.white, width: 0.5 } }); T(s, "M" + (m + 1), { x, y: 1.85, w: 4 * ww, h: 0.35, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle" }); }
    }
    s.addShape(pres.shapes.RECTANGLE, { x: gx + 24 * ww, y: 1.85, w: 2 * ww, h: 0.35, fill: { color: C.text2 }, line: { color: H.white, width: 0.5 } });
    T(s, "M7", { x: gx + 24 * ww, y: 1.85, w: 2 * ww, h: 0.35, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle" });
    const rows = [
      ["Contract, deposit & insurance", 1, 2, C.text2],
      ["Mobilisation (≥ 28 days)", 1, 4, C.text2],
      ["Wi-Fi / RF site survey, final relay count", 3, 6, C.accent1],
      ["Procurement & shipment (lead time)", 4, 10, C.accent3],
      ["Pilot ward install — KTPH, WH, TTSH", 9, 11, C.accent1],
      ["UAT & pilot acceptance", 11, 13, C.accent1],
      ["Ward-by-ward rollout", 13, 21, C.text2],
      ["Training (nurses, IT)", 10, 22, C.accent1],
      ["Hypercare & optimisation", 21, 26, C.accent5],
    ];
    rows.forEach((r, i) => {
      const y = 2.35 + i * 0.46;
      if (i % 2 === 0) s.addShape(pres.shapes.RECTANGLE, { x: lx, y: y - 0.03, w: 12.13, h: 0.46, fill: { color: C.background2 }, line: { type: "none" } });
      T(s, r[0], { x: lx + 0.1, y, w: lw - 0.15, h: 0.4, fontSize: 11.5, valign: "middle" });
      s.addShape(pres.shapes.RECTANGLE, { x: gx + (r[1] - 1) * ww, y: y + 0.06, w: (r[2] - r[1] + 1) * ww, h: 0.28, fill: { color: r[3] }, line: { type: "none" }, objectName: "Gantt " + r[0] });
    });
    T(s, "Go-live complete: indicative week 26", { x: gx + 14 * ww, y: 6.48, w: 12 * ww, h: 0.28, fontSize: 10, color: C.accent4, align: "right" });
    T(s, [{ text: "■ ", options: { color: C.accent3 } }, { text: "Lead-time item — [INPUT REQUIRED] confirm weeks", options: { color: H.amberTx } }], { x: lx, y: 6.48, w: 5, h: 0.28, fontSize: 10 });
  }

  // ================= 17. REFERENCES =================
  {
    const s = newSlide("6 · References and track record", "Our track record in comparable healthcare deployments gives ALPS evidence it can rely on", "Source: MONIT customer references. Obtain written consent from each reference before naming them.",
      "Provide 3+ comparable references (hospitals, nursing homes, public sector). For each: customer, scale, scope, dates, results and a contact person who agreed to be called.");
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.1;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.85, w: 3.95, h: 4.1, fill: { color: C.background2 }, line: { type: "none" } });
      await iconCircle(s, "FaHospital", x + 0.2, 2.05, 0.55, i % 2 ? C.accent1 : C.text2);
      T(s, "Reference " + (i + 1), { x: x + 0.9, y: 2.05, w: 2.9, h: 0.55, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
      inp(s, x + 0.2, 2.9, 3.55, 0.55, "Customer / site", { fs: 10 });
      inp(s, x + 0.2, 3.55, 3.55, 0.55, "Scale: beds, wards, devices", { fs: 10 });
      inp(s, x + 0.2, 4.2, 3.55, 0.55, "Scope & timeline", { fs: 10 });
      inp(s, x + 0.2, 4.85, 3.55, 0.9, "Outcome metrics + reference contact", { fs: 10 });
    }
    T(s, "Recommended: include at least one reference of a hospital Wi-Fi deployment to evidence the network integration.", { x: 0.6, y: 6.15, w: 12.13, h: 0.4, fontSize: 12, color: C.accent4 });
  }

  // ================= 18. SUSTAINABILITY =================
  {
    const s = newSlide("7 · Sustainability initiatives", "We can evidence sustainability practices across energy, tracking, policy and product choices", "Source: MONIT. Reflects ALPS's ISO14000-certified environment (Section 1, clause 19). Provide evidence for each item.",
      "Answer each of the four RFP sub-points with facts: renewable energy (e.g. solar), energy tracking in offices/factories, environmental policies, Green Label / energy-efficient products. If MONIT has no initiative in an area, say so honestly and describe the plan.");
    const q = [["FaSolarPanel", "Renewable energy", "Use of solar or other renewable energy in operations"], ["FaBolt", "Energy tracking", "Tracking of energy use in offices and factories"], ["FaRecycle", "Environmental policy", "Policies, practices and initiatives (waste, recycling, packaging)"], ["FaLeaf", "Green products", "Green Label products and energy-efficient equipment in offices / sites"]];
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + (i % 2) * 6.15, y = 1.85 + Math.floor(i / 2) * 2.4;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 5.98, h: 2.2, fill: { color: C.background2 }, line: { type: "none" } });
      await iconCircle(s, q[i][0], x + 0.2, y + 0.2, 0.55, C.accent5);
      T(s, q[i][1], { x: x + 0.95, y: y + 0.2, w: 4.8, h: 0.55, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
      T(s, q[i][2], { x: x + 0.2, y: y + 0.9, w: 5.6, h: 0.4, fontSize: 12 });
      inp(s, x + 0.2, y + 1.35, 5.58, 0.65, "Evidence: certificate, kWh data, policy document, product list", { fs: 10 });
    }
  }

  // ================= 19. NETWORK ARCHITECTURE =================
  {
    const s = newSlide("8 · Network connectivity and relay devices", "The relay joins the hospital Wi-Fi as a standard client device — no new network infrastructure is needed", "Source: MONIT. Protocols and security settings to be confirmed with the hospital IT teams and product engineering.",
      "Demonstrate the relay device(s). Bring a physical unit if the presentation allows. Confirm Wi-Fi standards supported (2.4/5 GHz, WPA2/WPA3-Personal/Enterprise), IMDA registration of the relay, power (PoE or adapter), mounting.");
    const nodes = [["FaMicrochip", "Sensors", "In ward beds"], ["FaWifi", "Relay device", "Joins hospital Wi-Fi"], ["FaNetworkWired", "Hospital Wi-Fi / AP", "Owned & managed by hospital"], ["FaCloud", "MONIT platform", "Secure cloud / server"], ["FaDesktop", "Dashboard", "Nurse station / mobile"]];
    for (let i = 0; i < 5; i++) {
      const x = 0.6 + i * 2.45;
      const hosp = i === 2;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.9, w: 2.2, h: 1.9, fill: { color: hosp ? C.background1 : C.background2 }, line: { color: hosp ? H.teal : H.light, width: hosp ? 1.5 : 0.5, dashType: hosp ? "dash" : "solid" } });
      await iconCircle(s, nodes[i][0], x + 0.8, 2.05, 0.6, i % 2 ? C.accent1 : C.text2);
      T(s, nodes[i][1], { x: x + 0.1, y: 2.8, w: 2.0, h: 0.35, fontSize: 13, bold: true, color: C.text2, align: "center" });
      T(s, nodes[i][2], { x: x + 0.1, y: 3.2, w: 2.0, h: 0.5, fontSize: 11, align: "center", color: C.accent4 });
      if (i < 4) s.addShape(pres.shapes.LINE, { x: x + 2.2, y: 2.85, w: 0.25, h: 0, line: { color: H.teal, width: 2, endArrowType: "triangle" } });
    }
    const f = [["Wi-Fi client", "Standard 802.11 client; works with existing SSID & security policy [confirm bands / WPA mode]"], ["Outbound-only", "Initiates secure connections out; no inbound ports opened on hospital network [confirm]"], ["IMDA compliant", "Relay registered / labelled for use in Singapore [INPUT REQUIRED: reg. no.]"]];
    f.forEach((c, i) => {
      const x = 0.6 + i * 4.1;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 4.2, w: 3.95, h: 2.1, fill: { color: C.background2 }, line: { type: "none" } });
      T(s, c[0], { x: x + 0.2, y: 4.35, w: 3.55, h: 0.35, fontSize: 14, bold: true, color: C.text2 });
      T(s, c[1], { x: x + 0.2, y: 4.8, w: 3.55, h: 1.4, fontSize: 12 });
    });
  }

  // ================= 20. SETUP =================
  {
    const s = newSlide("8 · Network connectivity and relay devices", "Setup takes minutes per relay: mount, power, connect once, verify in the dashboard", "Source: MONIT proposed setup flow — confirm with product engineering which provisioning method(s) the relay supports.",
      "Describe the real provisioning flow. Options typically are: (a) local web/app provisioning by hospital IT, (b) pre-staged configuration file / central push, (c) WPS/QR. State which is supported and who performs it (hospital IT vs MONIT engineer with hospital IT present).");
    const st = [["FaWrench", "1  Mount", "Relay mounted at the planned ward position"], ["FaPlug", "2  Power on", "Plug-in / PoE; relay starts in setup mode"], ["FaKey", "3  Connect once", "Hospital IT enters SSID & credentials during setup — MONIT never holds the password"], ["FaLink", "4  Register", "Relay registers with the platform and pairs with ward sensors"], ["FaCheckCircle", "5  Verify", "Signal & connectivity check shown in dashboard"]];
    for (let i = 0; i < 5; i++) {
      const x = 0.6 + i * 2.45;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.9, w: 2.2, h: 3.0, fill: { color: C.background2 }, line: { type: "none" } });
      await iconCircle(s, st[i][0], x + 0.8, 2.05, 0.6, i === 2 ? C.accent1 : C.text2);
      T(s, st[i][1], { x: x + 0.1, y: 2.8, w: 2.0, h: 0.35, fontSize: 14, bold: true, color: C.text2, align: "center" });
      T(s, st[i][2], { x: x + 0.15, y: 3.25, w: 1.9, h: 1.5, fontSize: 11.5, align: "center" });
    }
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 5.15, w: 12.13, h: 1.4, fill: { color: C.text2 }, line: { type: "none" } });
    T(s, "Who does what", { x: 0.85, y: 5.25, w: 4, h: 0.3, fontSize: 13, bold: true, color: "CADCFC" });
    T(s, [{ text: "Hospital IT: ", options: { bold: true } }, { text: "provides SSID/credentials and any whitelisting (MAC addresses supplied by MONIT).   ", options: {} }, { text: "MONIT: ", options: { bold: true } }, { text: "configures relays, provides configuration guide, supports on site during rollout." }], { x: 0.85, y: 5.6, w: 11.6, h: 0.85, fontSize: 13, color: C.background1 });
  }

  // ================= 21. PASSWORD PERSISTENCE =================
  {
    const s = newSlide("8 · Network connectivity and relay devices", "Wi-Fi credentials persist in protected memory — no re-keying after power outages or firmware upgrades", "Source: MONIT design intent. IMPORTANT — verify each statement against the actual firmware before submission.",
      "KEY ANSWER TO ALPS'S QUESTION. The mechanism described is the standard design pattern (non-volatile storage, configuration preserved across OTA, A/B partitions with rollback, auto-reconnect). MONIT engineering must confirm this matches the real firmware; edit anything that does not.");
    table(s, [
      ["Event", "What happens on the relay", "Wi-Fi password re-entry?"],
      ["Power outage / restart", "Credentials reloaded from non-volatile memory; relay auto-reconnects with retry", "No"],
      ["Firmware upgrade (OTA)", "Configuration partition is kept separate and migrated; rollback if the update fails", "No"],
      ["Hospital changes Wi-Fi password", "New credentials pushed by hospital IT (locally or centrally)", "Yes — once, by hospital IT"],
      ["Factory reset / replacement unit", "Relay returns to setup mode", "Yes — one-time setup"],
    ], 0.6, 1.85, 7.2, [2.2, 3.5, 1.5], { rowH: 0.78, fs: 11 });
    T(s, "How the solution achieves this", { x: 8.2, y: 1.85, w: 4.5, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
    const m = [["FaDatabase", "Persistent, encrypted storage of credentials"], ["FaSyncAlt", "A/B firmware slots; settings kept across upgrades"], ["FaRedo", "Automatic reconnection after any outage"], ["FaCloudUploadAlt", "Optional remote credential update by hospital IT"]];
    for (let i = 0; i < 4; i++) {
      const y = 2.35 + i * 0.95;
      await iconCircle(s, m[i][0], 8.2, y, 0.55, i % 2 ? C.accent1 : C.text2);
      T(s, m[i][1], { x: 8.95, y, w: 3.8, h: 0.55, fontSize: 12.5, valign: "middle" });
    }
    inp(s, 0.6, 6.0, 12.13, 0.6, "Engineering to confirm storage method, OTA behaviour and remote-update capability against the actual firmware.");
  }

  // ================= 22. SECURITY & OPS =================
  {
    const s = newSlide("8 · Network connectivity and relay devices", "The design keeps hospital networks and patient data protected and the service observable", "Source: MONIT; RFP Consent Form Annex A (IT Security Compliance List, Third-Party PDPA checklist, SaaS security requirements).",
      "The submission checklist in Annex A includes IT security, PDPA and SaaS security compliance forms. Make sure these statements are consistent with those forms.");
    await card(s, 0.6, 1.85, 3.9, 3.4, { icon: "FaLock", head: "Network security", body: ["Encrypted Wi-Fi (WPA2/3) per hospital policy", "TLS-encrypted traffic to platform", "Outbound-only connections", "Device identity & whitelisting by MAC"] });
    await card(s, 4.72, 1.85, 3.9, 3.4, { icon: "FaUserShield", head: "Data protection", body: ["PDPA-aligned handling of personal data", "Data minimisation and role-based access", "Hosting location: [INPUT REQUIRED]", "Compliance with Annex A security lists"] });
    await card(s, 8.83, 1.85, 3.9, 3.4, { icon: "FaHeartbeat", head: "Operational monitoring", body: ["Relay online / offline status in dashboard", "Alerts to support team on disconnects", "Remote diagnostics and firmware management", "Change-controlled upgrades scheduled with hospital"] });
    inp(s, 0.6, 5.55, 12.13, 0.85, "Status of IT Security Compliance List, Third-Party PDPA checklist and SaaS security requirements (Annex A) — attach completed forms.");
  }

  // ================= 23. OPEN QUESTIONS =================
  {
    const s = newSlide("Discussion", "Six points to confirm with ALPS and the three hospitals will fix the final design and price", "Source: MONIT analysis of RFP documents.",
      "Use this slide to drive the meeting. Record answers and owners.");
    table(s, [
      ["#", "Question for ALPS / hospitals", "Why it matters"],
      ["1", "Which wards and how many beds are in scope at each hospital?", "Drives relay count, quantity and price"],
      ["2", "Wi-Fi: SSID/band, authentication type, VLAN, MAC whitelisting process, IT contact?", "Determines relay configuration and approval time"],
      ["3", "Are relay mounting locations / power (PoE or sockets) available in wards?", "Installation scope and lead time"],
      ["4", "Preferred integration with nurse call / EMR, if any?", "Software scope and security review"],
      ["5", "Installation windows and infection-control requirements per ward?", "Rollout schedule"],
      ["6", "Required SLA, support hours and reporting for Section 2 Master Agreement?", "Service model and pricing"],
    ], 0.6, 1.85, 12.13, [0.6, 7.2, 4.33], { rowH: 0.62, fs: 12 });
  }

  // ================= 24. NEXT STEPS =================
  {
    const s = newSlide("Discussion", "Next steps: agree open points, complete the relay table and finalise the proposal package", "Source: MONIT. Owners and dates to be agreed in the meeting.",
      "Fill owners and dates during the meeting. Also remind the team of submission requirements: compliance tables stamped and signed, Annex A checklist, certificates, price proposal in Ariba, security deposit and insurance after award.");
    table(s, [
      ["Action", "Owner", "Date"],
      ["Confirm scope (wards / beds) and Wi-Fi requirements with hospital IT", "[INPUT]", "[INPUT]"],
      ["Complete relay sizing table from floor plans", "MONIT", "[INPUT]"],
      ["Compile certificates, clinical evidence, references, sustainability evidence", "MONIT", "[INPUT]"],
      ["Prepare Section 2/3 compliance, Annex A checklist and price proposal", "MONIT", "[INPUT]"],
      ["Schedule vendor presentation / demo and sample relay", "ALPS / MONIT", "[INPUT]"],
    ], 0.6, 1.85, 8.0, [5.2, 1.4, 1.4], { rowH: 0.65, fs: 12 });
    s.addShape(pres.shapes.RECTANGLE, { x: 8.9, y: 1.85, w: 3.83, h: 4.35, fill: { color: C.text2 }, line: { type: "none" } });
    T(s, "Decisions needed today", { x: 9.15, y: 2.05, w: 3.4, h: 0.4, fontSize: 15, bold: true, color: C.background1 });
    T(s, bullets(["Confirm in-scope wards and beds", "Agree provisioning method for Wi-Fi", "Agree pilot ward per hospital", "Confirm presentation / demo date"]), { x: 9.15, y: 2.6, w: 3.4, h: 3.4, fontSize: 13, color: C.background1 });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
}
main().catch((e) => { console.error(e); process.exit(1); });
