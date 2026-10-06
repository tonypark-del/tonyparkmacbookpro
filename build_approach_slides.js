// Build: NODE_PATH=<dir with node_modules> node build_approach_slides.js
// Four-slide answer to the KTPH question "Proposed approach and methodology" (same tone and manner as the main deck).
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const SKILL = "/root/.claude/skills/synced/740b1519-7df3-42c4-87b2-972ad2eb207a_a3d7b2ae-009d-4d22-98fe-59c0c04fc335/pptx";
const { applyTheme } = require(SKILL + "/scripts/apply_theme.js");

const OUT = "KTPH_Approach_Methodology.pptx";
const THEME = {
  name: "MONIT RFP", headFontFace: "Cambria", bodyFontFace: "Calibri",
  colors: { dk1: "1F2937", lt1: "FFFFFF", dk2: "0B2A4A", lt2: "F1F4F8", accent1: "00A3AD", accent2: "0B2A4A", accent3: "F2A900", accent4: "6B7A8F", accent5: "2E7D6B", accent6: "C0392B", hlink: "00A3AD", folHlink: "6B7A8F" },
};
const H = { propBg: "E6F6F7", propTx: "0B6F76", navy: "0B2A4A", teal: "00A3AD", grey: "6B7A8F", light: "F1F4F8", line: "D5DCE6", amber: "F2A900", amberBg: "FFF4D6", amberTx: "7A4F00", text: "1F2937", white: "FFFFFF" };
const SECTION = "4 · Understanding of the project";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "KTPH-RFP-26-165-MJ — Proposed approach and methodology";
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

// ---------- helpers (same pattern as build_deck.js) ----------
const T = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, fontSize: 12, color: C.text1, valign: "top", ...o });
const iconCache = {};
async function icon(name, color = "#FFFFFF") {
  const k = name + color;
  if (iconCache[k]) return iconCache[k];
  const svg = RDS.renderToStaticMarkup(React.createElement(fa[name] || fa.FaCircle, { color, size: "256" }));
  return (iconCache[k] = "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64"));
}
let sectionAdded = false;
function newSlide(title, source, notes) {
  if (!sectionAdded) { pres.addSection({ title: SECTION }); sectionAdded = true; }
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: SECTION });
  T(s, SECTION.toUpperCase(), { x: 0.6, y: 0.32, w: 8, h: 0.26, fontSize: 10, bold: true, color: C.accent1, charSpacing: 2 });
  s.addText(title, { placeholder: "title" });
  if (source) T(s, source, { x: 0.6, y: 6.78, w: 12.1, h: 0.22, fontSize: 9, color: C.accent4 });
  if (notes) s.addNotes(notes);
  return s;
}
function prop(s, x, y, w, h, text, o = {}) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: H.propBg }, line: { color: H.teal, width: 1, dashType: "dash" }, objectName: "Proposed response" });
  T(s, [{ text: "PROPOSED  ", options: { bold: true, color: H.propTx } }, { text, options: { color: H.text } }], { x: x + 0.12, y, w: w - 0.24, h, fontSize: o.fs || 11, valign: "middle" });
}
async function iconCircle(s, name, x, y, d = 0.5) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: C.text2 }, line: { type: "none" } });
  const p = d * 0.5;
  s.addImage({ data: await icon(name), x: x + (d - p) / 2, y: y + (d - p) / 2, w: p, h: p });
}
const bullets = (arr, gap = 4) => arr.map((t, i, a) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < a.length - 1, paraSpaceAfter: gap } }));
function table(s, rows, x, y, w, colW, o = {}) {
  const fs = o.fs || 11;
  const data = rows.map((r, ri) => r.map((c, ci) => {
    const base = { fontSize: fs, valign: "middle", color: H.text, border: [{ type: "none" }, { type: "none" }, { type: "solid", pt: 0.75, color: H.line }, { type: "none" }], margin: [0.04, 0.07, 0.04, 0.07] };
    if (ri === 0) return { text: c, options: { ...base, bold: true, color: H.white, fill: { color: H.navy }, border: [{ type: "none" }, { type: "none" }, { type: "none" }, { type: "none" }] } };
    if (typeof c === "string" && c.startsWith("[PROPOSED]")) return { text: [{ text: "PROPOSED  ", options: { bold: true, color: H.propTx } }, { text: c.slice(10).trim(), options: { color: H.text } }], options: { ...base, fill: { color: H.propBg } } };
    return { text: c, options: { ...base, bold: ci === 0, color: ci === 0 ? H.navy : H.text, fill: { color: ri % 2 ? H.white : H.light } } };
  }));
  s.addTable(data, { x, y, w, colW, rowH: o.rowH || 0.42, autoPage: false });
}

async function main() {
  // ================= A. WHAT KTPH EVALUATES =================
  {
    const s = newSlide("Our approach is built to answer the four questions KTPH will use to judge delivery risk",
      "Source: RFP Section 1 (28-day mobilisation); Master Agreement Clauses 4, 12, 23 and Schedule 2; MONIT field experience in Singapore (2025–2026).",
      "Purpose: the 'Proposed approach and methodology' question asks HOW MONIT will execute the project, not only how the technology works. Each card answers one evaluation question and points to the slide that proves it.");
    const cards = [
      ["FaClipboardCheck", "A credible plan", "Can MONIT explain who does what, in which order, and to what standard?", ["Six phases from Letter of Award to hand-over, each with a deliverable and a sign-off", "One MONIT project manager as single point of contact"], "Next slide"],
      ["FaUserNurse", "Minimal disruption to care", "Will installation and go-live disturb wards, nurses and patients?", ["Pilot ward first at each hospital, then ward by ward at agreed times", "Mounting follows each hospital's rules: ceiling, concealed or plug-in with no drilling"], "Slides 2 and 4"],
      ["FaCheckDouble", "Acceptance-ready", "Can KTPH verify quality before accepting and paying?", ["Factory test before shipment, witnessed acceptance tests on site", "Final Acceptance Notice per hospital; invoice only after acceptance"], "Slides 2 and 3"],
      ["FaHospital", "Delivered across three hospitals", "Can one vendor handle KTPH, Woodlands Health and TTSH together?", ["Same proven pattern at all three sites, sequenced in one plan", "MONIT engineering plus local on-site support from our Singapore partner"], "Slide 3"],
    ];
    const w = 5.98, h = 2.32;
    for (let i = 0; i < cards.length; i++) {
      const [ic, head, ask, ans, ref] = cards[i];
      const x = 0.6 + (i % 2) * (w + 0.17), y = 1.8 + Math.floor(i / 2) * (h + 0.15);
      s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.background2 }, line: { type: "none" } });
      await iconCircle(s, ic, x + 0.2, y + 0.2, 0.52);
      T(s, head, { x: x + 0.85, y: y + 0.2, w: w - 2.3, h: 0.52, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
      T(s, ref, { x: x + w - 1.45, y: y + 0.2, w: 1.25, h: 0.52, fontSize: 10, color: C.accent1, bold: true, align: "right", valign: "middle" });
      T(s, [{ text: "KTPH asks:  ", options: { bold: true, color: H.grey } }, { text: ask, options: { italic: true, color: H.grey } }], { x: x + 0.2, y: y + 0.85, w: w - 0.4, h: 0.45, fontSize: 11 });
      T(s, bullets(ans), { x: x + 0.2, y: y + 1.33, w: w - 0.4, h: h - 1.43, fontSize: 12 });
    }
  }

  // ================= B. SIX-PHASE METHODOLOGY =================
  {
    const s = newSlide("Six phases take each hospital from Letter of Award to hand-over, each closed by a hospital sign-off",
      "Source: RFP Section 1 (mobilisation); Master Agreement Clauses 12, 23 and Schedule 2 (training, testing, acceptance, invoicing). Teal cells are MONIT proposals to be confirmed.",
      "Timings marked 'MA' come from the Master Agreement; the 28-day mobilisation comes from the RFP. Durations in teal (survey window, install rate, 4-week hypercare) are MONIT proposals, not contract terms - confirm before submission. Security deposit timing differs (RFP 14 days vs MA Schedule 2 30 days): plan to the shorter.");
    table(s, [
      ["Phase", "Key activities", "MONIT deliverable", "Hospital input", "Sign-off", "Timing"],
      ["1  Mobilise", "Kick-off with ALPS and the three hospitals; project managers named; security deposit and insurance lodged", "Project plan and contact list", "Leads named for IT, nursing and facilities", "Plan agreed at kick-off", "Within 28 days of award (RFP)"],
      ["2  Survey and design", "Review ward layouts; survey 2.4 GHz Wi-Fi and power points, including day rooms and corridors; fix the gateway count per ward", "Site design per ward: gateway positions and quantities", "Ward layouts, SSID and VLAN details, ward access", "Site design signed by hospital IT and ward", "[PROPOSED] Weeks 2–4"],
      ["3  Test and deliver", "Factory test of every sensor and gateway; kits labelled by hospital and ward; advance delivery notice", "Test records and delivery note", "Receiving point and storage", "Delivery accepted", "Before installation (MA)"],
      ["4  Install and commission", "Pilot ward first at each hospital, then ward by ward; gateways mounted to each hospital's rules; sensors registered; dashboard set up at the nurse station", "Commissioning checklist per ward", "Escorted access at agreed times; IT on call for network join", "Ward commissioning signed", "[PROPOSED] 1–2 wards per day"],
      ["5  Accept", "Acceptance tests witnessed by hospital staff: alerts, dashboard, connectivity, reconnection after power loss", "Test report and certificate", "Witnesses from nursing and IT", "Final Acceptance Notice per hospital", "Invoice within 7 days of acceptance (MA)"],
      ["6  Train and stabilise", "On-site training for nurses and IT administrators; quick-reference guides; daily check of connectivity and alerts during hypercare", "Training records; hypercare report", "Staff released for training; ward nursing champions", "Hand-over to maintenance", "Plan ≤14 days of contract; training ≤14 days of delivery (MA)"],
    ], 0.6, 1.8, 12.13, [1.55, 3.6, 2.05, 1.95, 1.48, 1.5], { fs: 10, rowH: [0.35, 0.74, 0.74, 0.74, 0.74, 0.74, 0.74] });
  }

  // ================= C. GOVERNANCE AND ROLES =================
  {
    const s = newSlide("Clear roles, a fixed meeting rhythm and written change control keep all three hospitals on one plan",
      "Source: Master Agreement Clauses 4 and 23 (project managers, quarterly meetings, change control); MONIT. Teal boxes are MONIT proposals to be confirmed.",
      "Role table shows responsibilities, not headcount. Named individuals to be added in the submission. Intega Pte Ltd is the Singapore partner providing on-site support.");
    table(s, [
      ["Role", "Who", "Responsibility"],
      ["Project manager", "MONIT", "Single point of contact; plan, progress, risks and reporting"],
      ["Engineering and cloud", "MONIT", "Platform set-up, security, hosting in Singapore, remote support"],
      ["On-site support", "Intega (Singapore partner)", "Installation, on-site troubleshooting, part replacement"],
      ["Hospital IT", "KTPH / WH / TTSH", "2.4 GHz Wi-Fi access, network approvals, power points"],
      ["Nursing champions", "Each ward", "Workflow sign-off, user feedback, first-line help"],
      ["Procurement", "ALPS / MMD", "Contract, change approval, acceptance records"],
    ], 0.6, 1.8, 7.55, [1.9, 2.05, 3.6], { fs: 11, rowH: 0.62 });
    const x = 8.4, w = 4.33;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.8, w, h: 1.25, fill: { color: C.background2 }, line: { type: "none" } });
    await iconCircle(s, "FaCalendarAlt", x + 0.2, 1.95, 0.45);
    T(s, "Meeting rhythm", { x: x + 0.78, y: 1.95, w: w - 1.0, h: 0.45, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
    T(s, bullets(["Quarterly project-manager meetings (Master Agreement)"]), { x: x + 0.2, y: 2.5, w: w - 0.4, h: 0.45, fontSize: 11 });
    prop(s, x, 3.15, w, 0.75, "Weekly progress call and one-page written report while deployment is under way", { fs: 10.5 });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 4.05, w, h: 0.55, fill: { color: C.background2 }, line: { type: "none" } });
    T(s, "Escalation", { x: x + 0.2, y: 4.05, w: w - 0.4, h: 0.55, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
    prop(s, x, 4.65, w, 0.75, "Project manager → MONIT CTO → MONIT CEO, within the agreed service-level times", { fs: 10.5 });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 5.55, w, h: 0.95, fill: { color: C.text2 }, line: { type: "none" } });
    T(s, [{ text: "Change control:  ", options: { bold: true } }, { text: "every change is approved in writing before work starts; no unapproved charges." }], { x: x + 0.2, y: 5.55, w: w - 0.4, h: 0.95, fontSize: 11.5, color: C.background1, valign: "middle" });
  }

  // ================= D. RISK-LED SITE READINESS =================
  {
    const s = newSlide("Site readiness checks are built on what we learned in Singapore deployments, so issues are fixed before go-live",
      "Source: MONIT deployments in Singapore 2025–2026 (hospital clinical trial, long-term-care PoC); NHG requirement table; Master Agreement Schedule 3 (gateway radio and power).",
      "Lessons are stated as prevention measures, not as past incidents. Behind them: residents spending 4-5 hours in day rooms outside gateway range, a gateway lost after its USB adapter was bent, and a sensor registered on the tablet but not visible on the web dashboard. All were resolved on site; the checks below stop them recurring.");
    table(s, [
      ["Risk", "How we prevent it", "When checked"],
      ["Gateways use 2.4 GHz Wi-Fi only", "Confirm SSID, authentication and VLAN with hospital IT; test the join at every gateway position", "Survey; commissioning"],
      ["No power point where a gateway is needed", "Map power points in the survey; ceiling power at KTPH, concealed plug-in at WH, wall plug-in without drilling at TTSH", "Survey"],
      ["Patients spend hours in day rooms and corridors", "Survey covers day rooms and common areas; add or move gateways where patients spend time", "Survey; first week of hypercare"],
      ["Gateway plug knocked or damaged", "Mount away from walkways and secure the plug; keep spare gateways in Singapore for quick swap", "Installation; maintenance"],
      ["Dashboard and app show different status", "One registration per sensor; check dashboard, app and nurse-station view match before sign-off", "Acceptance test"],
      ["Patient or bed changes", "Simple re-assignment steps in training and the quick-reference guide", "Training"],
    ], 0.6, 1.8, 12.13, [3.2, 6.6, 2.33], { fs: 11, rowH: 0.6 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 6.1, w: 12.13, h: 0.52, fill: { color: C.text2 }, line: { type: "none" } });
    T(s, [{ text: "Result:  ", options: { bold: true } }, { text: "each ward is accepted only after coverage, power, network and dashboard checks pass." }], { x: 0.85, y: 6.1, w: 11.6, h: 0.52, fontSize: 12, color: C.background1, valign: "middle" });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
}
main().catch((e) => { console.error(e); process.exit(1); });
