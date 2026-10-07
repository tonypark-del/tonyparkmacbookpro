// Build: NODE_PATH=<dir with node_modules> node build_deck_261008.js
// Rebuild of KTPH_RFP_MONIT_Presentation_261007 for the 8 Oct 2026 ALPS meeting:
// one theme (Cambria / Calibri, navy + teal), facts aligned with the vault notes
// (docs/KTPH_RFP_261007_정합성검토.md), images cropped from the 261007 PDF into assets/261008.
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const SKILL = "/root/.claude/skills/synced/740b1519-7df3-42c4-87b2-972ad2eb207a_a3d7b2ae-009d-4d22-98fe-59c0c04fc335/pptx";
const { applyTheme } = require(SKILL + "/scripts/apply_theme.js");

const OUT = "KTPH_RFP_MONIT_Presentation_261008.pptx";
const A = (n) => path.join(__dirname, "assets/261008", n + ".png");
const RFP = "KTPH-RFP-26-165-MJ";
const THEME = {
  name: "MONIT RFP", headFontFace: "Cambria", bodyFontFace: "Calibri",
  colors: { dk1: "1F2937", lt1: "FFFFFF", dk2: "0B2A4A", lt2: "F1F4F8", accent1: "00A3AD", accent2: "0B2A4A", accent3: "F2A900", accent4: "6B7A8F", accent5: "2E7D6B", accent6: "C0392B", hlink: "00A3AD", folHlink: "6B7A8F" },
};
const H = { propBg: "E6F6F7", propTx: "0B6F76", navy: "0B2A4A", teal: "00A3AD", grey: "6B7A8F", light: "F1F4F8", line: "D5DCE6", amber: "F2A900", amberBg: "FFF4D6", amberTx: "7A4F00", red: "C0392B", green: "2E7D6B", text: "1F2937", white: "FFFFFF", sub: "DDEFF1" };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = RFP + " — AI-Enabled Smart Diaper Care System";
pres.author = "MONIT";
const C = pres.SchemeColor;

pres.defineSlideMaster({ title: "TITLE", background: { color: "FFFFFF" }, objects: [] });
pres.defineSlideMaster({
  title: "CONTENT", background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.62, w: 12.13, h: 0.95, fontSize: 22, bold: true, align: "left", color: C.text2, valign: "top", margin: 0 }, text: "" } },
    { text: { text: `MONIT  |  ${RFP}  |  Discussion document — Confidential`, options: { x: 0.6, y: 7.05, w: 8, h: 0.25, fontSize: 9, color: C.accent4, margin: 0 } } },
  ],
  slideNumber: { x: 12.0, y: 7.05, w: 0.73, h: 0.25, fontSize: 9, color: C.accent4, align: "right" },
});

// ---------- helpers ----------
const T = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, fontSize: 12, color: C.text1, valign: "top", ...o });
const iconCache = {};
async function icon(name, color = "#FFFFFF") {
  const k = name + color;
  if (iconCache[k]) return iconCache[k];
  const svg = RDS.renderToStaticMarkup(React.createElement(fa[name] || fa.FaCircle, { color, size: "256" }));
  return (iconCache[k] = "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64"));
}
// image fitted inside a box (contain), centred
async function img(s, name, x, y, w, h, o = {}) {
  const m = await sharp(A(name)).metadata();
  const r = Math.min(w / m.width, h / m.height);
  const iw = m.width * r, ih = m.height * r;
  const ix = o.align === "left" ? x : x + (w - iw) / 2;
  const iy = o.valign === "top" ? y : y + (h - ih) / 2;
  s.addImage({ path: A(name), x: ix, y: iy, w: iw, h: ih });
  if (o.border) s.addShape(pres.shapes.RECTANGLE, { x: ix, y: iy, w: iw, h: ih, fill: { type: "none" }, line: { color: H.line, width: 0.75 } });
  return { x: ix, y: iy, w: iw, h: ih };
}
const SECTIONS = new Set();
function newSlide(section, title, source, notes) {
  if (!SECTIONS.has(section)) { SECTIONS.add(section); pres.addSection({ title: section }); }
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  T(s, section.toUpperCase(), { x: 0.6, y: 0.3, w: 10, h: 0.26, fontSize: 10, bold: true, color: C.accent1, charSpacing: 2 });
  s.addText(title, { placeholder: "title" });
  if (source) T(s, source, { x: 0.6, y: 6.72, w: 12.1, h: 0.3, fontSize: 9, color: C.accent4 });
  if (notes) s.addNotes(notes);
  return s;
}
function prop(s, x, y, w, h, text, fs = 11) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: H.propBg }, line: { color: H.teal, width: 1, dashType: "dash" } });
  T(s, [{ text: "PROPOSED  ", options: { bold: true, color: H.propTx } }, { text, options: { color: H.text } }], { x: x + 0.12, y, w: w - 0.24, h, fontSize: fs, valign: "middle" });
}
async function iconCircle(s, name, x, y, d = 0.5, bg = C.text2) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" } });
  const p = d * 0.5;
  s.addImage({ data: await icon(name), x: x + (d - p) / 2, y: y + (d - p) / 2, w: p, h: p });
}
// tinted card with optional icon, head and body (string or bullet list)
async function card(s, x, y, w, h, o) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: o.fill || C.background2 }, line: { type: "none" } });
  if (o.bar) s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.07, h, fill: { color: C.accent1 }, line: { type: "none" } });
  const px = x + (o.bar ? 0.27 : 0.2);
  let ty = y + 0.16;
  if (o.icon) {
    await iconCircle(s, o.icon, px, y + 0.18, 0.5, o.iconBg);
    T(s, o.head, { x: px + 0.65, y: y + 0.18, w: w - (px - x) - 0.8, h: 0.5, fontSize: o.hfs || 14, bold: true, color: C.text2, valign: "middle" });
    ty = y + 0.82;
  } else if (o.head) {
    T(s, o.head, { x: px, y: ty, w: w - (px - x) - 0.15, h: 0.34, fontSize: o.hfs || 14, bold: true, color: C.text2, valign: "middle" });
    ty += 0.42;
  }
  if (o.body) {
    const body = Array.isArray(o.body) ? bullets(o.body) : o.body;
    T(s, body, { x: px, y: ty, w: w - (px - x) - 0.15, h: y + h - ty - 0.1, fontSize: o.fs || 11.5 });
  }
}
const bullets = (arr, sp = 4) => arr.map((t, i, a) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < a.length - 1, paraSpaceAfter: sp } }));
function table(s, rows, x, y, w, colW, o = {}) {
  const fs = o.fs || 11;
  const data = rows.map((r, ri) => r.map((c, ci) => {
    const base = { fontSize: fs, valign: "middle", color: H.text, border: [{ type: "none" }, { type: "none" }, { type: "solid", pt: 0.75, color: H.line }, { type: "none" }], margin: [0.04, 0.08, 0.04, 0.08], align: (o.right || []).includes(ci) ? "right" : "left" };
    if (ri === 0) return { text: c, options: { ...base, bold: true, color: H.white, fill: { color: H.navy }, border: [{ type: "none" }, { type: "none" }, { type: "none" }, { type: "none" }] } };
    if (typeof c === "string" && c.startsWith("[PROPOSED]")) return { text: [{ text: "PROPOSED  ", options: { bold: true, color: H.propTx } }, { text: c.slice(10).trim(), options: { color: H.text } }], options: { ...base, fill: { color: H.propBg } } };
    if (c && typeof c === "object" && c.text !== undefined) return { text: c.text, options: { ...base, fill: { color: ri % 2 ? H.white : H.light }, ...c.options } };
    return { text: c, options: { ...base, bold: ci === 0 && !o.noBoldFirst, fill: { color: ri % 2 ? H.white : H.light } } };
  }));
  s.addTable(data, { x, y, w, colW, rowH: o.rowH || 0.4, autoPage: false });
}
const numCircle = (s, n, x, y, d = 0.5, bg = C.text2) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" } });
  T(s, String(n), { x, y, w: d, h: d, fontSize: d > 0.45 ? 16 : 12, bold: true, color: C.background1, align: "center", valign: "middle", fontFace: "Cambria" });
};

// ---------- cost model (same as build_cost_slide.js, incl. ISO) ----------
const ESC = 0.05, yrs = [0, 1, 2, 3, 4];
const esc = (y1, n) => y1 * Math.pow(1 + ESC, n);
const azure = yrs.map((n) => esc(24000, n)), fd = yrs.map((n) => esc(4800, n)), saas = yrs.map((n) => esc(12000, n));
const pen = yrs.map((n) => (n === 0 ? 12000 : esc(7000, n - 1)));
const ISO_CYCLE = ["initial", "surveillance", "surveillance", "recertification", "surveillance"];
const isoFactor = { initial: 1, surveillance: 1 / 3, recertification: 2 / 3 };
const iso = yrs.map((n) => 8000 * isoFactor[ISO_CYCLE[n]] * Math.pow(1 + ESC, n));
const item6 = [19200, 19800, 20400, 21000, 21600];
const labourSub = yrs.map((n) => esc(36000, n) - item6[n]);
const subtotal = yrs.map((n) => azure[n] + fd[n] + saas[n] + pen[n] + iso[n] + labourSub[n]);
const opTotal = yrs.map((n) => subtotal[n] + item6[n]);
const setup = [14285.71, 0, 0, 0, 0];
const grand = yrs.map((n) => opTotal[n] + setup[n]);
const sum = (a) => a.reduce((x, y) => x + y, 0);
const f = (v) => (v ? Math.round(v).toLocaleString("en-US") : "–");
const k = (v) => "S$" + (v / 1000).toFixed(1) + "k";

async function main() {
  // ================= 1. COVER =================
  {
    pres.addSection({ title: "Cover" });
    const s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Cover" });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.33, h: 5.35, fill: { color: H.navy }, line: { type: "none" } });
    T(s, "MONIT", { x: 0.9, y: 0.75, w: 4, h: 0.4, fontSize: 14, bold: true, color: C.accent1, charSpacing: 6 });
    T(s, "AI-Enabled Smart Diaper Care System", { x: 0.9, y: 1.55, w: 9, h: 1.5, fontSize: 40, bold: true, color: C.background1, fontFace: "Cambria" });
    T(s, `Response to ${RFP} — discussion with ALPS and the KTPH, Woodlands Health and TTSH project teams`, { x: 0.9, y: 3.2, w: 8.6, h: 0.8, fontSize: 17, color: "CADCFC" });
    T(s, "8 October 2026  |  Confidential", { x: 0.9, y: 4.35, w: 8, h: 0.3, fontSize: 12, color: "CADCFC" });
    s.addShape(pres.shapes.OVAL, { x: 10.15, y: 1.0, w: 2.9, h: 2.9, fill: { color: H.teal, transparency: 20 }, line: { type: "none" } });
    s.addShape(pres.shapes.OVAL, { x: 11.2, y: 3.1, w: 1.7, h: 1.7, fill: { color: "FFFFFF", transparency: 85 }, line: { type: "none" } });
    s.addImage({ data: await icon("FaHospital"), x: 11.05, y: 1.9, w: 1.1, h: 1.1 });
    await img(s, "logo_samsung", 4.65, 5.6, 4.0, 1.45);
  }

  // ================= 2. COMPANY =================
  {
    const s = newSlide("1 · Company introduction", "MONIT is a Samsung spin-off digital-health company focused on making incontinence care smarter and more dignified",
      "Source: MONIT company records; NHIS (National Health Insurance Service, Korea) innovative welfare device listing; IMDA, KC and CE registrations (Section 2).");
    table(s, [["Company", "Established", "CEO", "Head office (Korea)", "Singapore"],
      ["MONIT Corp.", "10 April 2017", "Tony (Dohyeong) Park", "Seocho AICT Center, 56 Yangjae-daero 12-gil, Seoul", "9 Straits View, Marina One West #05-07 · distributor Intega Healthcare"]],
      0.6, 1.72, 12.13, [1.6, 1.45, 1.9, 3.4, 3.78], { fs: 11, rowH: 0.42, noBoldFirst: true });
    // business band
    await card(s, 0.6, 2.75, 5.9, 1.55, { head: "Technology and business", bar: true, body: [
      "AIoT urine and faecal soiling sensing technology",
      "Senior Diaper Care System (sensor, relay, dashboard) and Senior Diaper Subscription",
      "B2B: general hospitals, elderly-care hospitals and care facilities",
    ], fs: 11 });
    await img(s, "elderly", 6.7, 2.75, 1.9, 1.55);
    await card(s, 8.8, 2.75, 3.93, 1.55, { head: "B2G (Korea)", bar: true, body: "Selected as an innovative welfare device by NHIS, Korea's National Health Insurance Service", fs: 11 });
    const rows = [
      ["Proven history", "Samsung spin-off, invested by Samsung Venture Investment · NHIS innovative welfare device · customers in Korea, Japan, Singapore, the Netherlands and Australia"],
      ["Proven technology", "Clinically trialled at KTPH Ward B76 (7 weeks, 15 patients) · IMDA, KC and CE registered · ISO 27001 / 27017 / 27018 certification in November 2026"],
      ["Complete, integrated system", "Sensor → relay → hospital Wi-Fi → cloud → nurse dashboard, delivered and supported by one team with Intega on site"],
    ];
    for (let i = 0; i < rows.length; i++) {
      const y = 4.55 + i * 0.68;
      numCircle(s, i + 1, 0.6, y, 0.5);
      T(s, rows[i][0], { x: 1.3, y, w: 3.0, h: 0.5, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
      T(s, rows[i][1], { x: 4.3, y, w: 8.43, h: 0.5, fontSize: 11.5, valign: "middle" });
      if (i < 2) s.addShape(pres.shapes.LINE, { x: 0.6, y: y + 0.6, w: 12.13, h: 0, line: { color: H.line, width: 0.75 } });
    }
    await img(s, "intega", 11.5, 0.2, 1.23, 0.42);
  }

  // ================= 3. CERTIFICATIONS =================
  {
    const s = newSlide("2 · Industry certifications and standards", "Product registrations are in place, and ISO 27001 / 27017 / 27018 certification follows in November 2026",
      "Source: IMDA, CE and KC registration records; MONIT ISMS programme. ISO certificates will be provided on issue.");
    table(s, [["Standard", "What it covers", "Status", "Timing"],
      ["ISO/IEC 27001 — information security (ISMS)", "Systematic protection of patient-related data across platform and dashboard", "Certification audit in progress", "November 2026"],
      ["ISO/IEC 27017 — cloud security controls", "Security controls for the cloud platform hosting hospital data", "Certification audit in progress", "November 2026"],
      ["ISO/IEC 27018 — PII protection in public cloud", "Safeguards for personal data processed in the cloud; supports PDPA compliance", "Certification audit in progress", "November 2026"],
    ], 0.6, 1.72, 12.13, [3.9, 4.6, 2.2, 1.43], { fs: 11, rowH: 0.44 });
    table(s, [["Registration", "Sensor", "Relay gateway", "Validity"],
      ["IMDA (Singapore)", "Registered — ESER/26/5288 · BLE diaper sensor for sale and use in Singapore", "Registered — ESER/26/5450 · relay gateway for sale and use in Singapore", "Sensor Aug 2031 · Gateway Sep 2031"],
      ["CE (EU)", "Certified — RED 2014/53/EU (radio, EMC and electrical safety)", "Certified — EMC 2014/30/EU (EU conformity assessment)", "No expiry"],
      ["KC (Korea)", "Registered — R-R-mNT-SEMSYS200KSEN · radio and EMC conformity", "Registered — R-R-mNT-SSG · radio and EMC conformity", "No expiry"],
    ], 0.6, 3.85, 12.13, [2.0, 4.1, 4.1, 1.93], { fs: 11, rowH: 0.56 });
    T(s, "ISO certification is scoped to the MECS PRO platform, dashboard and Azure Singapore hosting used for the three hospitals.", { x: 0.6, y: 6.25, w: 12.1, h: 0.3, fontSize: 11, italic: true, color: C.accent4 });
  }

  // ================= 4. PRODUCT =================
  {
    const s = newSlide("3 · Introduction to the smart diaper system", "MECS PRO: an AI diaper care system built for professional care settings",
      "Components: sensor (clips onto the diaper), Safeguard sticker, relay gateway (hospital Wi-Fi), app service and web dashboard.");
    await img(s, "product", 0.6, 1.55, 12.13, 5.1);
  }

  // ================= 5. VALUE =================
  {
    const s = newSlide("3 · Introduction to the smart diaper system", "Smart care for health, dignity and efficiency: changes at the right time, not on a fixed round",
      "Source: MONIT; KTPH Ward B76 trial (9 Jul – 24 Aug 2025): 4.73 wet diapers per patient-day on average. Diaper saving is a MONIT customer estimate.");
    await card(s, 0.6, 1.72, 5.3, 1.75, { head: "Value to hospitals", body: [
      "Right-time diaper changes instead of routine checks",
      "Less nurse time on manual rounds; fewer disturbances for patients",
      "Objective data to support skin-integrity and continence care",
    ] });
    const pillars = [
      ["FaHeartbeat", "Patients", "Skin integrity", "Prompt changes lower the risk of incontinence-associated dermatitis and pressure injury"],
      ["FaHandHoldingUsd", "Hospital", "Consumables", "Up to 25% fewer unnecessary diaper changes"],
      ["FaUserNurse", "Caregivers", "Workload", "Fewer manual checks; alerts go to the nurse who needs them"],
    ];
    for (let i = 0; i < pillars.length; i++) {
      const x = 0.6 + i * 1.8, y = 3.7;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 1.7, h: 2.85, fill: { color: C.background2 }, line: { type: "none" } });
      await iconCircle(s, pillars[i][0], x + 0.55, y + 0.2, 0.6, i === 1 ? C.accent1 : C.text2);
      T(s, pillars[i][1], { x: x + 0.1, y: y + 0.9, w: 1.5, h: 0.3, fontSize: 11, color: C.accent4, align: "center" });
      T(s, pillars[i][2], { x: x + 0.1, y: y + 1.2, w: 1.5, h: 0.35, fontSize: 14, bold: true, color: C.text2, align: "center" });
      T(s, pillars[i][3], { x: x + 0.12, y: y + 1.6, w: 1.46, h: 1.2, fontSize: 10.5, align: "center" });
    }
    await img(s, "sensor_band", 6.1, 1.72, 2.0, 2.6);
    await img(s, "tablet", 6.0, 4.5, 2.3, 2.0);
    await img(s, "diaper1", 8.5, 1.72, 4.23, 2.4);
    await img(s, "diaper2", 8.5, 4.2, 4.23, 2.4);
  }

  // ================= 6. AI + COMPARISON =================
  {
    const s = newSlide("3 · Introduction to the smart diaper system", "AI pattern learning and a multi-channel sensor give a graded soiling level, not just wet or dry",
      "Source: KTPH Smart Diaper Sensor Trial, Ward B76, Phase 1 (9 Jul – 24 Aug 2025), 7-week mean, dropouts excluded; MONIT patent register.");
    await img(s, "sensor_band", 0.6, 1.7, 2.6, 3.3);
    await img(s, "samsung", 0.75, 5.25, 2.3, 0.4);
    T(s, "Developed from Samsung Electronics sensing research", { x: 0.6, y: 5.75, w: 2.6, h: 0.5, fontSize: 10, color: C.accent4, align: "center" });
    T(s, "Global patents in diaper sensing", { x: 3.5, y: 1.7, w: 5, h: 0.3, fontSize: 13, bold: true, color: C.text2 });
    await img(s, "patents", 3.5, 2.05, 9.23, 0.95, { align: "left" });
    const rows = [
      ["Measured detection", "Accuracy 76.2% · specificity 82.6% · sensitivity 72.7%, measured by KTPH nurses on Ward B76", "Typically wet / dry threshold only"],
      ["Soiling display", "Colour spectrum shows the soiling level so staff can judge urgency", "Two states (wet / dry); staff cannot judge urgency"],
      ["Sensitivity tuning", "AI pattern analysis plus 5 presets and 15 personal levels per patient", "Fixed threshold or manual configuration"],
    ];
    T(s, "MONIT", { x: 6.1, y: 3.2, w: 3.7, h: 0.4, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle", fill: { color: H.green } });
    T(s, "Typical alternatives", { x: 9.95, y: 3.2, w: 2.78, h: 0.4, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle", fill: { color: H.grey } });
    for (let i = 0; i < rows.length; i++) {
      const y = 3.72 + i * 0.95;
      s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y, w: 9.23, h: 0.85, fill: { color: C.background2 }, line: { type: "none" } });
      s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y, w: 0.07, h: 0.85, fill: { color: C.accent1 }, line: { type: "none" } });
      T(s, rows[i][0], { x: 3.7, y, w: 2.3, h: 0.85, fontSize: 12.5, bold: true, color: C.text2, valign: "middle" });
      T(s, rows[i][1], { x: 6.1, y: y + 0.06, w: 3.7, h: 0.73, fontSize: 10.5, valign: "middle", fill: { color: "E8F3EF" } });
      T(s, rows[i][2], { x: 9.95, y: y + 0.06, w: 2.78, h: 0.73, fontSize: 10.5, valign: "middle", color: C.accent4, fill: { color: H.white } });
    }
  }

  // ================= 7. KTPH TRIAL =================
  {
    const s = newSlide("3 · Introduction to the smart diaper system", "The KTPH Ward B76 trial measured the sensor in real care, and staff feedback shaped the final system",
      "Source: Smart Diapers Sensor Trial — Phase 1 consolidated results, Week 1 to 7, KTPH (report dated 3 Dec 2025). Phase 2 continues on the B105 Smart Ward.");
    const kpi = [["76.2%", "Accuracy", "(TP+TN) / all readings"], ["82.6%", "Specificity", "dry diaper correctly left alone"], ["72.7%", "Sensitivity", "wet diaper correctly alerted"], ["818", "Readings", "108 patient-days, 15 patients"]];
    for (let i = 0; i < kpi.length; i++) {
      const x = 0.6 + i * 1.95;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.72, w: 1.82, h: 1.5, fill: { color: C.background2 }, line: { type: "none" } });
      T(s, kpi[i][0], { x: x + 0.12, y: 1.8, w: 1.6, h: 0.6, fontSize: 28, bold: true, color: C.accent1, fontFace: "Cambria" });
      T(s, kpi[i][1], { x: x + 0.12, y: 2.42, w: 1.6, h: 0.3, fontSize: 12, bold: true, color: C.text2 });
      T(s, kpi[i][2], { x: x + 0.12, y: 2.72, w: 1.65, h: 0.45, fontSize: 9.5, color: C.accent4 });
    }
    await card(s, 0.6, 3.4, 3.8, 3.2, { head: "Trial design", body: [
      "Khoo Teck Puat Hospital, Ward B76 (B105 Nesting Ward)",
      "Team 3 male and Team 2 female cubicles",
      "15 patients enrolled",
      "Phase 1: 9 July – 24 August 2025 (7 weeks)",
      "Wet-diaper rate: mean 4.73 per patient-day (median 4.56)",
    ], fs: 11 });
    await card(s, 4.55, 3.4, 3.65, 3.2, { head: "Why Singapore, why KTPH", body: [
      "KTPH recorded 33 patients with IAD in 2024",
      "Another restructured hospital reports IAD in ~20–30% of inpatients",
      "Tests local climate and humidity, diverse patients and staff acceptance",
      "Feedback (alarm handling, sensor placement) fed into the current release",
    ], fs: 11 });
    await img(s, "nurses1", 8.45, 1.72, 4.28, 2.4);
    await img(s, "nurses2", 8.45, 4.2, 4.28, 2.4);
  }

  // ================= 8. PLATFORM =================
  {
    const s = newSlide("3 · Introduction to the smart diaper system", "The MECS PRO web platform gives each ward a real-time view of every patient's sensor",
      "Screens show sample data from a test environment. Also: global search (ward, room, patient, sensor), language selector, manual refresh with last-updated time and secure sign-out.");
    await img(s, "dashboard", 0.6, 1.62, 6.6, 5.0, { border: true });
    const nav = [["Dashboard", "Real-time view of wards, rooms, patients and sensors; layout editor for wards and rooms"], ["Device info", "Information on registered devices"], ["Manage sensors", "Sensor registry with search and filters, patient assignment, sensitivity, logs and CSV export"], ["Alerts", "Alert list and notification handling"], ["Settings", "System and account configuration"]];
    T(s, "Main navigation", { x: 7.6, y: 1.62, w: 5, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
    for (let i = 0; i < nav.length; i++) await card(s, 7.6, 2.05 + i * 0.92, 5.13, 0.86, { bar: true, head: nav[i][0], hfs: 12.5, body: nav[i][1], fs: 10 });
  }

  // ================= 9. SENSITIVITY =================
  {
    const s = newSlide("3 · Introduction to the smart diaper system", "Real-time monitoring and alerts: caregivers see status, act on alerts and tune sensitivity per patient",
      "Sensitivity settings: 5 quick presets (Sensitive → Insensitive) or a personal level of 1–15, set per sensor to suit the patient and diaper.");
    const cards = [["Hierarchical drill-down", "Ward / room list → ward → room, with breadcrumb and back button; patient count per room"], ["Status summary", "Per-room counters: dirty, connected, disconnected and unconnected sensors"], ["Per-patient sensor card", "Room tag, sensor name, battery %, alert toggle, status bar and last-updated time"], ["Notifications", "Per-sensor alert on/off, room-alert shortcut and notification history"], ["Sensor action menu", "Sensitivity · move room · edit device info · view sensor info · unassign"], ["Patient workflow", "Register and assign patients from the room view; data refreshes automatically"]];
    for (let i = 0; i < cards.length; i++) {
      const x = 0.6 + (i % 2) * 3.75, y = 1.72 + Math.floor(i / 2) * 1.63;
      await card(s, x, y, 3.6, 1.5, { bar: true, head: cards[i][0], hfs: 13, body: cards[i][1], fs: 11 });
    }
    await img(s, "sensitivity", 8.3, 1.72, 4.43, 4.85, { border: true });
  }

  // ================= 10. ADMIN =================
  {
    const s = newSlide("3 · Introduction to the smart diaper system", "Management and administration: one sensor registry and simple layout tools for each hospital",
      "Screens show sample data from a test environment.");
    await img(s, "manage", 0.6, 1.62, 6.5, 3.2, { border: true, align: "left" });
    await img(s, "layout", 0.6, 5.0, 6.5, 1.6, { border: true, align: "left" });
    const c = [["Sensor registry", "Sensor ID, memo, status, nickname, location, battery and last seen; register sensor and export CSV"], ["Search and filters", "Filter by ward, room or nickname, plus status and battery level"], ["Per-sensor actions", "Sensitivity · assign patient · edit info · logs · delete"], ["Layout management", "Add, rename or delete wards and rooms; save all, discard or close"], ["Access and localisation", "Organisation-scoped sign-in, global search and language selector"]];
    for (let i = 0; i < c.length; i++) await card(s, 7.4, 1.62 + i * 1.0, 5.33, 0.9, { bar: true, head: c[i][0], hfs: 12.5, body: c[i][1], fs: 10.5 });
  }

  // ================= 11. MOBILE =================
  {
    const s = newSlide("3 · Introduction to the smart diaper system", "The MECS PRO mobile app brings the same core functions to the bedside",
      "iOS app. Screens show sample data from a test environment; the notification page and per-bed menu are not shown.");
    await img(s, "phones", 0.6, 1.62, 8.4, 4.95);
    const c = [["① Bed-by-bed monitoring", "Card per bed: bed and patient no., battery, connection, last update and status bar; empty beds offer Add Patient"], ["② Per-device sensitivity", "5 presets from Sensitive to Insensitive, or personal customisation; saved per sensor ID"], ["③ Settings", "User mode (e.g. Enterprise), app language, notification sound, alert time, app version"], ["④ Alert time", "Slider (1–10 s) sets how long the alert sound plays when a change is needed"]];
    for (let i = 0; i < c.length; i++) await card(s, 9.25, 1.62 + i * 1.25, 3.48, 1.15, { bar: true, head: c[i][0], hfs: 12, body: c[i][1], fs: 10 });
  }

  // ================= 12. WHY MAINTENANCE (new) =================
  {
    const s = newSlide("3 · Maintenance, service and technical support", "Maintenance keeps a hospital-grade service running: each charge maps to work that would otherwise stop",
      "Source: MONIT cost basis (Rev. 2) with ISO 27001/27017/27018 added; Master Agreement Clause 13.1 (Charges fixed for the term).",
      "New slide for the 8 Oct ALPS meeting (agenda: why a maintenance charge is needed). Each card maps to a line on the next slide.");
    const c = [
      ["FaServer", "Singapore hosting", `${k(sum(azure) + sum(fd) + sum(saas))} over 5 years`, "Azure Singapore servers, web application firewall and the MONIT SaaS platform run 24×7, with a separate tenancy per hospital and n+1 redundancy. Without it, alerts stop."],
      ["FaShieldAlt", "Security and compliance", `${k(sum(pen) + sum(iso))} over 5 years`, "Yearly penetration re-test, ISO 27001/27017/27018 surveillance and recertification audits, and certificate renewal (now every 200 days) — the evidence hospital IT asks for."],
      ["FaHeadset", "Remote support and monitoring", `${k(sum(labourSub))} over 5 years`, "24×7 security and service monitoring, incident response, firmware and app updates as phones, OS versions and security patches change."],
      ["FaTools", "On-site preventive maintenance", `${k(sum(item6))} over 5 years`, "Four rounds a year across the three hospitals with Intega: relay checks, battery and sensor health, coverage re-test after ward changes."],
    ];
    for (let i = 0; i < c.length; i++) {
      const x = 0.6 + (i % 2) * 6.13, y = 1.72 + Math.floor(i / 2) * 2.05;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 6.0, h: 1.9, fill: { color: C.background2 }, line: { type: "none" } });
      await iconCircle(s, c[i][0], x + 0.2, y + 0.2, 0.55);
      T(s, c[i][1], { x: x + 0.9, y: y + 0.18, w: 3.0, h: 0.35, fontSize: 14, bold: true, color: C.text2 });
      T(s, c[i][2], { x: x + 3.6, y: y + 0.18, w: 2.25, h: 0.35, fontSize: 13, bold: true, color: C.accent1, align: "right" });
      T(s, c[i][3], { x: x + 0.9, y: y + 0.6, w: 4.95, h: 1.2, fontSize: 11 });
    }
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 5.9, w: 12.13, h: 0.6, fill: { color: H.navy }, line: { type: "none" } });
    T(s, [{ text: "Why a fixed annual charge:  ", options: { bold: true } }, { text: "prices cannot change for the contract term, so costs that recur every year are priced up front with 5% escalation instead of re-quoted." }], { x: 0.85, y: 5.9, w: 11.7, h: 0.6, fontSize: 12, color: C.background1, valign: "middle" });
  }

  // ================= 13. COST =================
  {
    const s = newSlide("3 · Maintenance, service and technical support", `Five-year operating and maintenance cost is ${k(sum(opTotal))}, with 5% annual escalation built in so charges stay fixed`,
      "Source: MONIT cost basis (Rev. 2, 16 Sep 2026); Master Agreement Clause 13.1. SGD, ex-GST, rounded; 5.0% a year from Year 2. ISO: Y1 initial certification; Y2, Y3, Y5 surveillance (1/3) and Y4 recertification (2/3) of the initial audit, per ISO/IEC 17021-1.");
    const kpis = [["FaCalendarCheck", k(sum(opTotal) / 5), "Average operating cost / year"], ["FaChartLine", k(sum(opTotal)), "Five-year operating cost"], ["FaCloud", k(sum(subtotal)), `Annual subscription (${Math.round((sum(subtotal) / sum(opTotal)) * 100)}% of total)`], ["FaTools", k(sum(item6)), `Preventive maintenance (${Math.round((sum(item6) / sum(opTotal)) * 100)}% of total)`]];
    const cw = 2.9, gap = (12.13 - 4 * cw) / 3;
    for (let i = 0; i < kpis.length; i++) {
      const [ic, big, lab] = kpis[i];
      const x = 0.6 + i * (cw + gap), y = 1.62;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: 1.0, fill: { color: C.background2 }, line: { type: "none" } });
      await iconCircle(s, ic, x + 0.18, y + 0.22, 0.56);
      T(s, big, { x: x + 0.9, y: y + 0.1, w: cw - 1.0, h: 0.5, fontSize: 26, bold: true, color: C.accent1, fontFace: "Cambria", valign: "middle" });
      T(s, lab, { x: x + 0.9, y: y + 0.6, w: cw - 1.0, h: 0.3, fontSize: 11, color: C.accent4 });
    }
    const fs = 10.5, noB = [{ type: "none" }, { type: "none" }, { type: "none" }, { type: "none" }];
    const rule = [{ type: "none" }, { type: "none" }, { type: "solid", pt: 0.75, color: H.line }, { type: "none" }];
    const cell = (text, o = {}) => ({ text, options: { fontSize: fs, valign: "middle", color: H.text, border: rule, margin: [0.03, 0.08, 0.03, 0.08], ...o } });
    const num = (text, o = {}) => cell(text, { align: "right", ...o });
    const header = ["Line item", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5", "5-year total"].map((t, i) => cell(t, { bold: true, color: H.white, fill: { color: H.navy }, border: noB, align: i ? "right" : "left" }));
    const line = (label, arr, o = {}) => [cell(label, o), ...arr.map((v) => num(f(v), o)), num(f(sum(arr)), { bold: true, ...o })];
    const W = { fill: { color: H.white } }, L = { fill: { color: H.light } };
    const amb = { italic: true, color: H.amberTx, fill: { color: H.amberBg } };
    s.addTable([
      header,
      line("Azure server (Singapore, site-isolated)", azure, W), line("Azure Front Door (WAF / CDN)", fd, L), line("MONIT SaaS platform (Azure PaaS)", saas, W),
      line("Penetration re-test (one per year)", pen, L), line("ISO 27001/27017/27018 certification & audits", iso, W), line("Remote support & 24×7 monitoring", labourSub, L),
      line("Subscription (Section 3.1) subtotal", subtotal, { bold: true, fill: { color: H.sub } }),
      line("Preventive maintenance labour (Item 6)", item6, W),
      line("Total annual operating cost", opTotal, { bold: true, color: H.white, fill: { color: H.navy } }),
      [cell("Initial set-up (one-off) – to confirm", amb), num(f(setup[0]), amb), ...[1, 2, 3, 4].map(() => num("–", amb)), num(f(sum(setup)), { ...amb, bold: true })],
      line("Total annual expenditure", grand, { bold: true, fill: { color: H.sub } }),
    ], { x: 0.6, y: 2.82, w: 8.7, colW: [3.2, 0.83, 0.83, 0.83, 0.83, 0.83, 1.35], rowH: 0.29, autoPage: false });
    s.addShape(pres.shapes.RECTANGLE, { x: 9.55, y: 2.82, w: 3.18, h: 3.5, fill: { color: C.text2 }, line: { type: "none" } });
    T(s, "What the price covers", { x: 9.75, y: 2.96, w: 2.8, h: 0.35, fontSize: 14, bold: true, color: C.background1 });
    T(s, bullets(["Hosting in Singapore with a separate tenancy for each hospital and n+1 redundancy", "Encrypted backup, replication and a web application firewall", "Annual penetration re-test by a licensed provider", "ISO 27001/27017/27018 certification, yearly surveillance audits and Year-4 recertification", "Remote support, incident response and 24×7 security monitoring", "Four preventive-maintenance rounds a year across all three hospitals"], 5), { x: 9.75, y: 3.36, w: 2.8, h: 2.9, fontSize: 10.5, color: C.background1 });
  }

  // ================= 14. SIX PHASES =================
  {
    const s = newSlide("4 · Understanding of the project", "Six phases take each hospital from Letter of Award to hand-over, each closed by a hospital sign-off",
      "Source: RFP and Master Agreement (MA) timings as marked; teal cells are MONIT proposals to be confirmed.");
    table(s, [["Phase", "Key activities", "MONIT deliverable", "Hospital input", "Sign-off", "Timing"],
      ["1  Mobilise", "Kick-off with ALPS and the three hospitals; project managers named; insurance in place", "Project plan and contact list", "Leads named for IT, nursing and facilities", "Plan agreed at kick-off", "Within 28 days of award (RFP)"],
      ["2  Survey and design", "Review ward layouts; survey 2.4 GHz Wi-Fi and power points incl. day rooms and corridors; fix relay count per ward", "Site design per ward: relay positions and quantities", "Ward layouts, SSID and VLAN details, ward access", "Site design signed by hospital IT and ward", "[PROPOSED] Weeks 2–4"],
      ["3  Test and deliver", "Factory test of every sensor and relay; kits labelled by hospital and ward; advance delivery notice", "Test records and delivery note", "Receiving point and storage", "Delivery accepted", "Before installation (MA)"],
      ["4  Install and commission", "Pilot ward first at each hospital, then ward by ward; relays mounted to each hospital's rules; sensors registered; dashboard at the nurse station", "Commissioning checklist per ward", "Escorted access at agreed times; IT on call for network join", "Ward commissioning signed", "[PROPOSED] 1–2 wards per day"],
      ["5  Accept", "Acceptance tests witnessed by hospital staff: alerts, dashboard, connectivity, reconnection after power loss", "Test report and certificate", "Witnesses from nursing and IT", "Final Acceptance Notice per hospital", "Invoice within 7 days of acceptance (MA)"],
      ["6  Train and stabilise", "On-site training for nurses and IT administrators; quick-reference guides; daily check of connectivity and alerts during hypercare", "Training records; hypercare report", "Staff released for training; ward nursing champions", "Hand-over to maintenance", "Plan ≤ 14 days of contract; training ≤ 14 days of delivery (MA)"],
    ], 0.6, 1.62, 12.13, [1.6, 3.55, 1.95, 1.95, 1.65, 1.43], { fs: 10, rowH: 0.78 });
  }

  // ================= 15. RISKS =================
  {
    const s = newSlide("4 · Understanding of the project", "Site-readiness checks built on our Singapore deployments fix issues before go-live",
      "Source: MONIT deployment experience at KTPH and other Singapore sites.");
    table(s, [["Risk", "How we prevent it", "When checked"],
      ["Relays use 2.4 GHz Wi-Fi only", "Confirm SSID (English characters), authentication and VLAN with hospital IT; test the join at every relay position", "Survey; commissioning"],
      ["No power point where a relay is needed", "Map power points in the survey: ceiling power at KTPH, concealed plug-in at WH, wall plug-in without drilling at TTSH", "Survey"],
      ["Patients spend hours in day rooms and corridors", "Survey covers day rooms and common areas; add or move relays where patients spend time", "Survey; first week of hypercare"],
      ["Relay plug knocked or damaged", "Mount away from walkways and secure the plug; keep spare relays in Singapore for quick swap", "Installation; maintenance"],
      ["Dashboard and app show different status", "One registration per sensor; check dashboard, app and nurse-station view match before sign-off", "Acceptance test"],
      ["Patient or bed changes", "Simple re-assignment steps in training and the quick-reference guide", "Training"],
    ], 0.6, 1.62, 12.13, [3.3, 6.6, 2.23], { fs: 11, rowH: 0.6 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 5.95, w: 12.13, h: 0.55, fill: { color: H.navy }, line: { type: "none" } });
    T(s, [{ text: "Result:  ", options: { bold: true } }, { text: "each ward is accepted only after coverage, power, network and dashboard checks pass." }], { x: 0.85, y: 5.95, w: 11.7, h: 0.55, fontSize: 12, color: C.background1, valign: "middle" });
  }

  // ================= 16. GOVERNANCE =================
  {
    const s = newSlide("4 · Understanding of the project", "Clear roles, a fixed meeting rhythm and written change control keep all three hospitals on one plan",
      "Source: Master Agreement Clauses 4 and 23 (project managers, quarterly meetings, change control); MONIT. Teal boxes are MONIT proposals to be confirmed.");
    table(s, [["Role", "Who", "Responsibility"],
      ["Project manager", "MONIT", "Single point of contact; plan, progress, risks and reporting"],
      ["Engineering and cloud", "MONIT", "Platform set-up, security, hosting in Singapore, remote support"],
      ["On-site support", "Intega Healthcare (Singapore partner)", "Installation, on-site troubleshooting, part replacement"],
      ["Hospital IT", "KTPH / WH / TTSH", "2.4 GHz Wi-Fi access, network approvals, power points"],
      ["Nursing champions", "Each ward", "Workflow sign-off, user feedback, first-line help"],
      ["Procurement", "ALPS", "Contract, change approval, acceptance records"],
    ], 0.6, 1.62, 7.9, [2.0, 2.4, 3.5], { fs: 11, rowH: 0.62 });
    await card(s, 8.8, 1.62, 3.93, 1.3, { icon: "FaCalendarAlt", head: "Meeting rhythm", fs: 11, body: ["Quarterly project-manager meetings (Master Agreement)"] });
    prop(s, 8.8, 3.1, 3.93, 0.85, "Weekly progress call and one-page written report while deployment is under way", 10.5);
    prop(s, 8.8, 4.1, 3.93, 0.85, "Changes agreed in writing through ALPS before work starts (change log shared weekly)", 10.5);
  }

  // ================= 17. ROLLOUT SEQUENCE =================
  {
    const s = newSlide("4 · Understanding of the project", "Rollout follows the six phases in four stages, ward by ward, starting with one pilot ward per hospital",
      "Stages group the six phases: Plan = phases 1–2 · Pilot = phases 3–5 on the first ward · Roll out = phases 4–5 ward by ward · Optimise = phase 6. Ward order to be confirmed with each hospital.");
    const st = [["1  Plan", ["Kick-off and governance", "Wi-Fi / RF site survey with hospital IT", "Final relay count and placement", "Security and PDPA review"], H.navy],
      ["2  Pilot", ["Install in one pilot ward per hospital", "Credential provisioning once", "Staff training", "UAT and acceptance"], H.teal],
      ["3  Roll out", ["Ward-by-ward installation", "Minimise disruption to care", "Daily progress tracking", "Sign-off per ward"], H.navy],
      ["4  Optimise", ["Hypercare period", "Alert-rule tuning", "Usage and benefit reporting", "Hand-over to service team"], H.teal]];
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + i * 3.08;
      s.addShape(pres.shapes.PENTAGON, { x, y: 1.72, w: 2.95, h: 0.8, fill: { color: st[i][2] }, line: { type: "none" } });
      T(s, st[i][0], { x: x + 0.25, y: 1.72, w: 2.4, h: 0.8, fontSize: 16, bold: true, color: C.background1, valign: "middle", fontFace: "Cambria" });
      s.addShape(pres.shapes.RECTANGLE, { x, y: 2.65, w: 2.85, h: 2.0, fill: { color: C.background2 }, line: { type: "none" } });
      T(s, bullets(st[i][1], 5), { x: x + 0.2, y: 2.8, w: 2.5, h: 1.75, fontSize: 11.5 });
    }
    T(s, "Ward sequence by hospital", { x: 0.6, y: 4.9, w: 6, h: 0.35, fontSize: 14, bold: true, color: C.text2 });
    const seq = [["KTPH", "Tower D → Tower B"], ["Woodlands Health", "A41 → W62 → W74"], ["TTSH", "5A → 7C → 13B → 5H"]];
    for (let i = 0; i < seq.length; i++) {
      const x = 0.6 + i * 4.1;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 5.35, w: 3.95, h: 1.15, fill: { color: H.navy }, line: { type: "none" } });
      T(s, seq[i][0], { x: x + 0.25, y: 5.45, w: 3.5, h: 0.35, fontSize: 12, color: "CADCFC", bold: true });
      T(s, seq[i][1], { x: x + 0.25, y: 5.8, w: 3.5, h: 0.55, fontSize: 18, bold: true, color: C.background1, fontFace: "Cambria" });
    }
  }

  // ================= 18. RELAY QUANTITIES =================
  {
    const s = newSlide("4 · Understanding of the project", "Relay quantities follow a transparent, coverage-based method that is verified by an on-site survey",
      "Source: NHG ward requirement tables for KTPH, Woodlands Health and TTSH; relay reception about 20 m indoors. Final counts confirmed by the Wi-Fi / RF survey and pilot ward.");
    const steps = [["Map", "Mark bed locations and wall types on each ward floor plan"], ["Cover", "Place relays so every bed lies within the validated coverage radius"], ["Check capacity", "Keep sensors per relay below the maximum"], ["Add resilience", "Add redundancy for critical gaps (e.g. 10% spare)"], ["Verify", "Confirm with the Wi-Fi / RF survey and pilot ward"]];
    for (let i = 0; i < steps.length; i++) {
      const x = 0.6 + i * 2.45;
      numCircle(s, i + 1, x, 1.65, 0.5, i % 2 ? C.accent1 : C.text2);
      T(s, steps[i][0], { x: x + 0.62, y: 1.65, w: 1.7, h: 0.5, fontSize: 13, bold: true, color: C.text2, valign: "middle" });
      T(s, steps[i][1], { x, y: 2.22, w: 2.3, h: 0.6, fontSize: 10.5 });
    }
    const req = [["Khoo Teck Puat Hospital", "60 sensors · ceiling mount preferred · barcode scanner at monitoring station"], ["Woodlands Health", "Minimal, concealed relays (behind cabinet, not on ceiling) · dashboard at monitoring station"], ["Tan Tock Seng Hospital", "40 sensors · wall power-point plug-in · no drilling · dashboard at nursing station"]];
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.1;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 3.0, w: 3.95, h: 0.42, fill: { color: H.navy }, line: { type: "none" } });
      T(s, req[i][0], { x: x + 0.15, y: 3.0, w: 3.7, h: 0.42, fontSize: 12.5, bold: true, color: C.background1, valign: "middle" });
      T(s, req[i][1], { x: x + 0.15, y: 3.5, w: 3.7, h: 0.6, fontSize: 10.5, color: C.accent4 });
    }
    const R = { right: [1] };
    table(s, [["KTPH ward", "Relays"], ["Tower D (7-series)", "14"], ["Tower B (acute stroke unit)", "13"], ["Tower D (8-series)", "14"], ["", ""], [{ text: "Total", options: { bold: true, fill: { color: H.sub } } }, { text: "41", options: { bold: true, fill: { color: H.sub } } }]], 0.6, 4.2, 3.95, [2.75, 1.2], { fs: 11, rowH: 0.36, ...R });
    table(s, [["WH ward", "Relays"], ["A41", "14"], ["W62", "6"], ["W74", "10"], ["", ""], [{ text: "Total", options: { bold: true, fill: { color: H.sub } } }, { text: "30", options: { bold: true, fill: { color: H.sub } } }]], 4.7, 4.2, 3.95, [2.75, 1.2], { fs: 11, rowH: 0.36, ...R });
    table(s, [["TTSH ward", "Relays"], ["5A (Class C)", "6"], ["7C (Class B2)", "8"], ["13B (Class A)", "19"], ["5H (Class B2)", "19"], [{ text: "Total", options: { bold: true, fill: { color: H.sub } } }, { text: "52", options: { bold: true, fill: { color: H.sub } } }]], 8.8, 4.2, 3.93, [2.73, 1.2], { fs: 11, rowH: 0.36, ...R });
  }

  // ================= 19. INSTALLATION (new) =================
  {
    const s = newSlide("4 · Understanding of the project", "Installation is a short, repeatable routine per ward once the hospital has confirmed Wi-Fi and power",
      "Source: MECS PRO BLE Gateway installation guide; MONIT site-preparation checklist (6 Oct 2026). Detailed relay set-up in Section 8.",
      "New slide for the 8 Oct ALPS meeting (agenda: installation process).");
    await card(s, 0.6, 1.62, 4.2, 4.9, { icon: "FaClipboardCheck", head: "Hospital provides before installation", fs: 11.5, body: [
      "2.4 GHz Wi-Fi network (5 GHz not supported) with an SSID in English characters",
      "Wi-Fi name and password, or a dedicated IoT SSID / VLAN",
      "MAC-address whitelisting if required — MONIT supplies relay MAC addresses before delivery",
      "One power point per relay position (USB adapter supplied)",
      "Escorted ward access at agreed times; IT contact on call",
    ] });
    const steps = [["Survey", "Walk the ward with a test relay; confirm coverage and power points"], ["Mount", "Upper wall or corridor ceiling, away from metal cabinets; secure the plug"], ["Connect", "Power on; enter the hospital Wi-Fi once on the relay's set-up page (region code 1002)"], ["Register", "Scan each sensor QR code to the bed in the dashboard"], ["Verify", "Test alert on every bed; dashboard, app and nurse-station view must match"], ["Sign off", "Ward commissioning checklist signed by nursing and IT"]];
    for (let i = 0; i < steps.length; i++) {
      const y = 1.62 + i * 0.72;
      numCircle(s, i + 1, 5.1, y + 0.1, 0.5, i % 2 ? C.accent1 : C.text2);
      T(s, steps[i][0], { x: 5.75, y: y + 0.1, w: 1.5, h: 0.5, fontSize: 13, bold: true, color: C.text2, valign: "middle" });
      T(s, steps[i][1], { x: 7.25, y: y + 0.1, w: 5.48, h: 0.5, fontSize: 11, valign: "middle" });
      if (i < steps.length - 1) s.addShape(pres.shapes.LINE, { x: 5.1, y: y + 0.68, w: 7.63, h: 0, line: { color: H.line, width: 0.75 } });
    }
    prop(s, 5.1, 6.02, 7.63, 0.5, "1–2 wards per day per team; pilot ward first at each hospital", 11);
  }

  // ================= 20. GANTT =================
  {
    const s = newSlide("5 · Project timeline", "Implementation can start on award and reach full go-live in about five months",
      "Source: RFP (mobilisation within 28 days of award); Master Agreement (training plan ≤ 14 days of contract). Production lead time depends on the final relay count.");
    const M0 = 4.2, MW = 1.45;
    for (let m = 0; m < 5; m++) T(s, "M" + (m + 1), { x: M0 + m * MW, y: 1.62, w: MW - 0.03, h: 0.35, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle", fill: { color: H.navy } });
    const g = [["Contract and insurance", 0, 0.5, H.navy], ["Mobilisation (within 28 days)", 0, 1, H.navy], ["Wi-Fi / RF site survey, final relay count", 0.5, 1.5, H.teal], ["Procurement and shipment (lead time)", 0.75, 2.5, H.amber], ["Pilot ward install — KTPH, WH, TTSH", 2, 2.75, H.teal], ["UAT and pilot acceptance", 2.5, 3.25, H.teal], ["Ward-by-ward rollout", 3, 4, H.navy], ["Training (nurses, IT)", 2, 4, H.teal], ["Hypercare and optimisation", 4, 5, H.green]];
    for (let i = 0; i < g.length; i++) {
      const y = 2.1 + i * 0.5;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 12.13, h: 0.44, fill: { color: i % 2 ? H.white : H.light }, line: { type: "none" } });
      T(s, g[i][0], { x: 0.75, y, w: 3.4, h: 0.44, fontSize: 11.5, valign: "middle" });
      s.addShape(pres.shapes.RECTANGLE, { x: M0 + g[i][1] * MW, y: y + 0.09, w: (g[i][2] - g[i][1]) * MW, h: 0.26, fill: { color: g[i][3] }, line: { type: "none" } });
    }
  }

  // ================= 21-22. REFERENCES =================
  async function refCard(s, x, y, w, h, o) {
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.background2 }, line: { type: "none" } });
    await iconCircle(s, o.icon || "FaHospital", x + 0.18, y + 0.18, 0.5, o.hl ? C.accent1 : C.text2);
    T(s, o.name, { x: x + 0.8, y: y + 0.18, w: w - 0.95, h: 0.5, fontSize: 13, bold: true, color: C.text2, valign: "middle" });
    const lines = [["Where", o.where], ["Scale", o.scale], ["Timeline", o.when]];
    lines.forEach(([a, b], i) => T(s, [{ text: a + ":  ", options: { bold: true, color: C.text2 } }, { text: b }], { x: x + 0.2, y: y + 0.8 + i * 0.42, w: w - 0.35, h: 0.4, fontSize: 10.5, valign: "middle" }));
    if (o.img) await img(s, o.img, x + 0.12, y + 2.1, w - 0.24, h - 2.2);
    else if (o.note) T(s, o.note, { x: x + 0.2, y: y + 2.15, w: w - 0.4, h: h - 2.3, fontSize: 10.5, color: C.text1 });
  }
  {
    const s = newSlide("6 · References and track record", "Our Singapore track record starts at KTPH itself and extends to public eldercare and primary care",
      "Source: MONIT customer records; KTPH Smart Diapers Sensor Trial report (3 Dec 2025).");
    const w = 2.93, gap = (12.13 - 4 * w) / 3, y = 1.62, h = 4.95;
    await refCard(s, 0.6, y, w, h, { name: "Khoo Teck Puat Hospital", hl: true, where: "Ward B76, Singapore", scale: "15 patients, 818 readings", when: "9 Jul – 24 Aug 2025 · Phase 2 ongoing", note: "Accuracy 76.2%, specificity 82.6%, sensitivity 72.7% measured by KTPH nurses. Staff feedback on alarm handling and sensor placement went into the current release." });
    await refCard(s, 0.6 + (w + gap), y, w, h, { name: "Vanguard Healthcare", where: "Tampines Care Home", scale: "PoC with 2 residents", when: "20 May – 3 Jul 2026 · subscription roll-out under discussion", img: "ref_vanguard" });
    await refCard(s, 0.6 + 2 * (w + gap), y, w, h, { name: "NTUC Health", where: "Geylang East, Singapore", scale: "6 residents", when: "From 7 Sep 2026", img: "ref_ntuc" });
    await refCard(s, 0.6 + 3 * (w + gap), y, w, h, { name: "SingHealth Polyclinics", where: "Sengkang, Pasir Ris", scale: "30 participants", when: "From 10 Jul 2026", img: "ref_singhealth" });
  }
  {
    const s = newSlide("6 · References and track record", "Larger programmes in Japan and Korea show the system running at scale",
      "Source: MONIT customer records and NHIS listing. Heart Care figures as reported by the customer.");
    const w = 3.95, gap = (12.13 - 3 * w) / 2, y = 1.62, h = 4.95;
    await refCard(s, 0.6, y, w, h, { name: "Heart Care (Japan)", where: "14 elderly-care facilities, Osaka", scale: "About 2,000 sensors; 1 million diapers supplied", when: "From Dec 2025", img: "ref_heartcare" });
    await refCard(s, 0.6 + (w + gap), y, w, h, { name: "NHIS innovative welfare device (Korea)", where: "National Health Insurance Service", scale: "National long-term-care programme (1.23 million)", when: "Listed March 2024", img: "ref_nhis" });
    await refCard(s, 0.6 + 2 * (w + gap), y, w, h, { name: "Korean elderly hospitals", where: "Samsung Noble County and premium elderly hospitals", scale: "180 residents · Noble County PoC of 25 confirmed", when: "Ongoing; Noble County from late Oct 2026", img: "ref_korea" });
  }

  // ================= 23. SUSTAINABILITY =================
  {
    const s = newSlide("7 · Sustainability initiatives", "Environmental policy: MECS PRO cuts unnecessary diaper changes and the waste that goes with them",
      "Source: MONIT. Diaper reduction is a MONIT customer estimate.");
    await card(s, 0.6, 1.62, 7.0, 2.25, { icon: "FaRecycle", iconBg: C.accent5, head: "Environmental policy — in place", body: [
      "Policies and practices on waste, recycling and packaging",
      "MECS PRO can reduce unnecessary diaper consumption by up to 25%: changes happen when a diaper is soiled, not on a fixed round",
      "Reusable sensor — only the battery is replaced, every 2–3 months",
    ], fs: 11.5 });
    await img(s, "env1", 0.6, 4.05, 3.42, 2.5);
    await img(s, "env2", 4.18, 4.05, 3.42, 2.5);
    T(s, "Other criteria", { x: 8.0, y: 1.62, w: 4.7, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
    const o = [["FaSolarPanel", "Renewable energy", "Use of solar or other renewable energy in operations"], ["FaBolt", "Energy tracking", "Tracking of energy use in offices and factories"], ["FaLeaf", "Green products", "Green Label products and energy-efficient equipment in offices and sites"]];
    for (let i = 0; i < o.length; i++) {
      const y = 2.1 + i * 1.45;
      s.addShape(pres.shapes.RECTANGLE, { x: 8.0, y, w: 4.73, h: 1.3, fill: { color: C.background2 }, line: { type: "none" } });
      await iconCircle(s, o[i][0], 8.2, y + 0.2, 0.5, C.accent5);
      T(s, o[i][1], { x: 8.85, y: y + 0.2, w: 2.3, h: 0.5, fontSize: 13, bold: true, color: C.text2, valign: "middle" });
      T(s, "Not applicable", { x: 11.15, y: y + 0.27, w: 1.45, h: 0.36, fontSize: 10, bold: true, color: C.accent4, align: "center", valign: "middle", fill: { color: H.white } });
      T(s, o[i][2], { x: 8.85, y: y + 0.72, w: 3.75, h: 0.5, fontSize: 10.5, color: C.accent4 });
    }
  }

  // ================= 24. GATEWAY SPEC =================
  {
    const s = newSlide("8 · Network connectivity and relay devices", "The MECS PRO BLE gateway relays sensor data over the hospital Wi-Fi so wards are covered without new infrastructure",
      "Model SEM-SYS200K_GAT. Registrations in Section 2.");
    await img(s, "gw_labels", 0.6, 1.62, 4.4, 2.75);
    await img(s, "gw_views", 0.6, 4.45, 4.4, 2.15);
    table(s, [["Item", "Specification"],
      ["Function", "Receives MECS PRO sensor data over BLE and sends it to the cloud over Wi-Fi, for the nursing station and tablets"],
      ["Communication", "BLE 5 (beacon) receive · Wi-Fi 2.4 GHz transmit (5 GHz not supported)"],
      ["Reception range", "About 20 m indoors; varies with obstacles, interference and antenna orientation"],
      ["Data cycle", "Collects all nearby sensors every minute; stores and re-sends if Wi-Fi drops"],
      ["Updates", "Remote firmware update checked every hour — wireless, no cable"],
      ["Registrations", "IMDA ESER/26/5450 · CE · KC"],
      ["Dimensions / weight", "9.5 × 4.5 × 5.5–11.3 cm (antenna incl.) · 42 g"],
      ["Power", "5 V USB-A, always on; restarts automatically after power loss"],
      ["Installation", "Upper wall or corridor ceiling, away from metal cabinets, frames and pillars; exact number confirmed on site"],
      ["Data privacy", "No personal information stored — sensor identifiers and soiling data only"],
      ["Package / origin", "Gateway and quick-start guide (adapter not included) · made in China, developed and owned by MONIT (Korea) · HS 8517.62"],
    ], 5.3, 1.62, 7.43, [1.8, 5.63], { fs: 10, rowH: 0.4 });
  }

  // ================= 25. GATEWAY STEPS =================
  {
    const s = newSlide("8 · Network connectivity and relay devices", "Each relay is set up once from a phone or tablet in a few minutes — no new network infrastructure",
      "Set-up screens from the installation guide; network names and passwords are masked.");
    await img(s, "gw_steps", 0.6, 1.62, 12.13, 4.25);
    prop(s, 0.6, 6.0, 12.13, 0.55, "Relays supplied to the three hospitals carry a unique set-up password per device (strict-institution firmware); MONIT hands the password list to hospital IT.", 11);
  }

  // ================= 26. GATEWAY WIFI =================
  {
    const s = newSlide("8 · Network connectivity and relay devices", "The relay joins the hospital Wi-Fi as a standard client — only a 2.4 GHz Wi-Fi network is required",
      "Region code 1002 = Singapore. The set-up page closes once the relay restarts.");
    await img(s, "gw_portal", 0.6, 1.62, 2.9, 4.95, { border: true });
    await card(s, 3.7, 1.62, 2.6, 4.95, { head: "Enter the Wi-Fi details", fs: 12 });
    T(s, [
      { text: "①  Wi-Fi name (SSID)", options: { breakLine: true, paraSpaceAfter: 10 } },
      { text: "②  Wi-Fi password", options: { breakLine: true, paraSpaceAfter: 10 } },
      { text: "③  Region code 1002", options: { breakLine: true, paraSpaceAfter: 10 } },
      { text: "④  Save" },
    ], { x: 3.9, y: 2.4, w: 2.3, h: 2.5, fontSize: 13 });
    await img(s, "gw_verified", 6.6, 1.62, 3.4, 1.9, { border: true });
    await card(s, 10.2, 1.62, 2.53, 1.9, { head: "Verified", body: "When the page shows 'Verified', relay set-up is complete.", fs: 11 });
    await img(s, "gw_led", 6.6, 3.75, 3.4, 2.3, { border: true });
    await card(s, 10.2, 3.75, 2.53, 2.3, { head: "Solid white LED", body: "When the LED turns solid white, the relay is online.", fs: 11 });
  }

  // ================= 27. CREDENTIALS =================
  {
    const s = newSlide("8 · Network connectivity and relay devices", "Wi-Fi credentials are stored once on the relay — no re-entry after power outages or firmware updates",
      "Source: MECS PRO BLE Gateway installation guide and FAQ.");
    table(s, [["Event", "What happens on the relay", "Re-enter Wi-Fi password?"],
      ["Power outage / restart", "Credentials reloaded from non-volatile memory; relay restarts and reconnects automatically", "No"],
      ["Firmware update (over the air)", "Wi-Fi settings are kept across the update", "No"],
      ["Wi-Fi drops temporarily", "Sensor data is stored on the relay and re-sent when Wi-Fi returns", "No"],
      ["Hospital changes the Wi-Fi password", "Hospital IT enters the new credentials on each relay's set-up page", "Yes — once, by hospital IT"],
      ["Replacement unit", "New relay starts in set-up mode", "Yes — one-time set-up"],
    ], 0.6, 1.62, 7.9, [2.4, 3.9, 1.6], { fs: 11, rowH: 0.7 });
    T(s, "How the solution achieves this", { x: 8.8, y: 1.62, w: 4, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
    const h = [["FaDatabase", "Credentials kept in non-volatile memory"], ["FaSyncAlt", "Settings kept across firmware updates"], ["FaRedo", "Automatic reconnection after any outage"], ["FaHdd", "Local buffering while Wi-Fi is down"]];
    for (let i = 0; i < h.length; i++) {
      await iconCircle(s, h[i][0], 8.8, 2.2 + i * 1.05, 0.6, i % 2 ? C.accent1 : C.text2);
      T(s, h[i][1], { x: 9.6, y: 2.2 + i * 1.05, w: 3.13, h: 0.6, fontSize: 12, valign: "middle" });
    }
  }

  // ================= 28. DATA PROTECTION (new) =================
  {
    const s = newSlide("8 · Network connectivity and relay devices", "Patient data stays in Singapore, is kept only as long as needed, and is protected to ISO 27001/27017/27018",
      "Source: MONIT Corp Data Protection Notice v1.0 (10 Sep 2026); MONIT platform security design.",
      "New slide: summary of data-protection commitments asked for in the RFP security review.");
    const c = [
      ["FaMapMarkerAlt", "Data residency", ["Hosted on Microsoft Azure, Singapore region", "Separate tenancy for each hospital"]],
      ["FaUserShield", "Minimal personal data", ["No name, birth date or gender collected at device registration — device nickname only", "Relays store no personal information"]],
      ["FaLock", "Encryption and access", ["Azure SQL TDE and AES-256 encryption", "Web application firewall; organisation-scoped sign-in"]],
      ["FaHistory", "Retention", ["Care events kept 14 days", "Audit logs kept 12 months"]],
      ["FaBell", "Breach response", ["Notification within 1 day of a confirmed breach", "Named Data Protection Officer"]],
      ["FaCertificate", "Assurance", ["ISO 27001 / 27017 / 27018 certification in November 2026", "Annual penetration re-test by a licensed provider"]],
    ];
    for (let i = 0; i < c.length; i++) {
      const x = 0.6 + (i % 3) * 4.1, y = 1.62 + Math.floor(i / 3) * 2.5;
      await card(s, x, y, 3.95, 2.35, { icon: c[i][0], head: c[i][1], body: c[i][2], fs: 11.5 });
    }
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT, { five: Math.round(sum(opTotal)), avg: Math.round(sum(opTotal) / 5), grand: Math.round(sum(grand)) });
}
main().catch((e) => { console.error(e); process.exit(1); });
