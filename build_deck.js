const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5

// ---- Palette: Teal Trust ----
const C = {
  primary: "028090",   // teal
  secondary: "00A896", // seafoam
  accent: "02C39A",    // mint
  ink: "0B2E33",        // deep teal-black for text
  slate: "5B7C80",      // muted teal-gray for secondary text
  bg: "FFFFFF",
  cardBg: "F3F9F8",
  cardBg2: "EAF5F3",
  warnBg: "FDEEEA",
  warnText: "B24A2E",
  white: "FFFFFF",
};

const FONT_HEAD = "Cambria";
const FONT_BODY = "Calibri";

function bgSlide(slide, color) {
  slide.background = { color };
}

function pageNum(slide, n, total) {
  slide.addText(`${n} / ${total}`, {
    x: 12.5, y: 7.15, w: 0.7, h: 0.3, fontFace: FONT_BODY, fontSize: 9, color: C.slate, align: "right",
  });
}

function kicker(slide, text, opts = {}) {
  slide.addText(text.toUpperCase(), {
    x: opts.x ?? 0.6, y: opts.y ?? 0.5, w: opts.w ?? 8, h: 0.35,
    fontFace: FONT_BODY, fontSize: 12, bold: true, color: opts.color ?? C.accent,
    charSpacing: 1.5,
  });
}

const TOTAL = 10;

// ============================================================
// SLIDE 1 — Title
// ============================================================
{
  const slide = pres.addSlide();
  bgSlide(slide, C.ink);

  slide.addText("WORKFORCE REPORTING AUTOMATION", {
    x: 0.8, y: 2.35, w: 11.7, h: 0.4, fontFace: FONT_BODY, fontSize: 14, bold: true,
    color: C.accent, charSpacing: 2,
  });
  slide.addText("From raw attendance sheets to\nboardroom-ready reports — automatically", {
    x: 0.8, y: 2.75, w: 11.7, h: 2.0, fontFace: FONT_HEAD, fontSize: 40, bold: true,
    color: C.white, lineSpacingMultiple: 1.08,
  });
  slide.addText("Two Google Apps Script projects that turn spreadsheet data into styled, KPI-driven\nemail reports — no manual compilation, no formatting drift, no missed cycles.", {
    x: 0.8, y: 4.75, w: 10.5, h: 0.9, fontFace: FONT_BODY, fontSize: 15, color: "BFDBD6",
    lineSpacingMultiple: 1.3,
  });

  // small credential row
  slide.addText("GMR MONTHLY COMPARISON DASHBOARD     •     APRAAVA WORKFORCE MANDATE & ATTENDANCE REPORT", {
    x: 0.8, y: 6.55, w: 11.5, h: 0.4, fontFace: FONT_BODY, fontSize: 11, color: C.secondary, charSpacing: 1,
  });

  // decorative circles motif (top right)
  slide.addShape("ellipse", { x: 10.6, y: 0.4, w: 2.4, h: 2.4, fill: { color: C.primary, transparency: 60 }, line: { type: "none" } });
  slide.addShape("ellipse", { x: 11.6, y: 1.6, w: 1.3, h: 1.3, fill: { color: C.accent, transparency: 40 }, line: { type: "none" } });

  pageNum(slide, 1, TOTAL);
}

// ============================================================
// SLIDE 2 — The Problem
// ============================================================
{
  const slide = pres.addSlide();
  bgSlide(slide, C.bg);
  kicker(slide, "The Problem");
  slide.addText("Field workforce reporting was a manual, error-prone chore", {
    x: 0.6, y: 0.85, w: 12.0, h: 0.9, fontFace: FONT_HEAD, fontSize: 30, bold: true, color: C.ink,
  });

  const painPoints = [
    { icon: "1", title: "Manual data pulls", body: "Someone opens raw attendance sheets and copies numbers out by hand, every reporting cycle." },
    { icon: "2", title: "Hand-calculated KPIs", body: "Attendance %, deltas, and deficits computed manually — a single formula slip changes the story." },
    { icon: "3", title: "Inconsistent formatting", body: "Reports built fresh each time in Excel or PowerPoint, so structure and quality vary cycle to cycle." },
    { icon: "4", title: "Slow, late delivery", body: "Compilation takes hours, so leadership often sees numbers days after the period actually closed." },
  ];

  const cardW = 2.85, gap = 0.28, startX = 0.6, y = 2.15, cardH = 3.7;
  painPoints.forEach((p, i) => {
    const x = startX + i * (cardW + gap);
    slide.addShape("roundRect", { x, y, w: cardW, h: cardH, rectRadius: 0.12, fill: { color: C.cardBg }, line: { type: "none" } });
    slide.addShape("ellipse", { x: x + 0.25, y: y + 0.3, w: 0.55, h: 0.55, fill: { color: C.primary }, line: { type: "none" } });
    slide.addText(p.icon, { x: x + 0.25, y: y + 0.3, w: 0.55, h: 0.55, fontFace: FONT_BODY, fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle" });
    slide.addText(p.title, { x: x + 0.25, y: y + 1.05, w: cardW - 0.5, h: 0.6, fontFace: FONT_HEAD, fontSize: 15, bold: true, color: C.ink, margin: 0 });
    slide.addText(p.body, { x: x + 0.25, y: y + 1.65, w: cardW - 0.5, h: 1.9, fontFace: FONT_BODY, fontSize: 12, color: C.slate, lineSpacingMultiple: 1.25, margin: 0 });
  });

  pageNum(slide, 2, TOTAL);
}

// ============================================================
// SLIDE 3 — The Solution / Architecture
// ============================================================
{
  const slide = pres.addSlide();
  bgSlide(slide, C.bg);
  kicker(slide, "The Solution");
  slide.addText("Reports build and send themselves — inside the sheet", {
    x: 0.6, y: 0.85, w: 12.0, h: 0.9, fontFace: FONT_HEAD, fontSize: 30, bold: true, color: C.ink,
  });
  slide.addText("Both projects follow the same 3-stage pipeline, running natively inside Google Apps Script — no servers, no external APIs.", {
    x: 0.6, y: 1.65, w: 11.5, h: 0.5, fontFace: FONT_BODY, fontSize: 14, color: C.slate,
  });

  const stages = [
    { label: "DATA LAYER", title: "Read & aggregate", body: "Reads raw attendance / mandate sheets. Computes attendance %, deltas, deficits, and breakdowns by project, circle, role, or location." },
    { label: "RENDER LAYER", title: "Build the report", body: "Assembles a self-contained, styled HTML email — inline CSS, KPI cards, tables, and charts — no external assets or dependencies." },
    { label: "DELIVERY LAYER", title: "Send & log", body: "Emails via MailApp/GmailApp — on demand from a custom menu, or automatically on a time-based trigger. Every send is logged." },
  ];

  const boxW = 3.55, gapX = 0.55, startX = 0.9, y = 2.55, boxH = 3.3;
  stages.forEach((s, i) => {
    const x = startX + i * (boxW + gapX);
    slide.addShape("roundRect", { x, y, w: boxW, h: boxH, rectRadius: 0.12, fill: { color: i === 1 ? C.primary : C.cardBg }, line: { type: "none" } });
    const textColor = i === 1 ? C.white : C.ink;
    const subColor = i === 1 ? "D7F0EC" : C.slate;
    slide.addText(s.label, { x: x + 0.3, y: y + 0.3, w: boxW - 0.6, h: 0.35, fontFace: FONT_BODY, fontSize: 11, bold: true, color: i === 1 ? C.accent : C.primary, charSpacing: 1.5 });
    slide.addText(s.title, { x: x + 0.3, y: y + 0.7, w: boxW - 0.6, h: 0.6, fontFace: FONT_HEAD, fontSize: 18, bold: true, color: textColor });
    slide.addText(s.body, { x: x + 0.3, y: y + 1.4, w: boxW - 0.6, h: 1.75, fontFace: FONT_BODY, fontSize: 12.5, color: subColor, lineSpacingMultiple: 1.3 });

    if (i < stages.length - 1) {
      slide.addText("→", { x: x + boxW + 0.08, y: y + boxH / 2 - 0.3, w: gapX - 0.16, h: 0.6, fontFace: FONT_BODY, fontSize: 26, bold: true, color: C.secondary, align: "center", valign: "middle" });
    }
  });

  pageNum(slide, 3, TOTAL);
}

// ============================================================
// SLIDE 4 — Project 1: GMR overview
// ============================================================
{
  const slide = pres.addSlide();
  bgSlide(slide, C.bg);
  kicker(slide, "Project 1 of 2");
  slide.addText("GMR Monthly Comparison Dashboard", {
    x: 0.6, y: 0.85, w: 12.0, h: 0.7, fontFace: FONT_HEAD, fontSize: 28, bold: true, color: C.ink,
  });
  slide.addText("Month-over-month attendance comparison for the Kashi and Triveni field projects.", {
    x: 0.6, y: 1.55, w: 11.5, h: 0.5, fontFace: FONT_BODY, fontSize: 14, color: C.slate,
  });

  // left: feature list
  const features = [
    "Side-by-side Kashi / Triveni header cards, each themed with its own accent color",
    "KPI rows: Active headcount, Present+Week-Off, Absent, Attendance %, Zero-present flags",
    "Gmail-safe HTML-table bar charts — deliberately not SVG, since Gmail/Outlook strip it",
    "Circle-wise and role-wise breakdowns per project",
    "Custom menu: Send Comparison Email · Preview Email · Write Summary Sheet",
  ];
  slide.addShape("roundRect", { x: 0.6, y: 2.3, w: 6.8, h: 4.35, rectRadius: 0.12, fill: { color: C.cardBg }, line: { type: "none" } });
  slide.addText("What it delivers", { x: 0.95, y: 2.55, w: 6.1, h: 0.4, fontFace: FONT_HEAD, fontSize: 16, bold: true, color: C.primary });
  slide.addText(
    features.map((f, i) => ({ text: f, options: { bullet: { code: "2022" }, breakLine: i < features.length - 1, color: C.ink, fontSize: 13 } })),
    { x: 0.95, y: 3.05, w: 6.1, h: 3.4, fontFace: FONT_BODY, valign: "top", paraSpaceAfter: 10, lineSpacingMultiple: 1.2 }
  );

  // right: fair comparison callout
  slide.addShape("roundRect", { x: 7.65, y: 2.3, w: 5.05, h: 4.35, rectRadius: 0.12, fill: { color: C.primary }, line: { type: "none" } });
  slide.addText("KEY DESIGN DECISION", { x: 8.0, y: 2.6, w: 4.4, h: 0.35, fontFace: FONT_BODY, fontSize: 11, bold: true, color: C.accent, charSpacing: 1.5 });
  slide.addText("Fair period\ncomparison", { x: 8.0, y: 3.0, w: 4.4, h: 1.1, fontFace: FONT_HEAD, fontSize: 22, bold: true, color: C.white, lineSpacingMultiple: 1.05 });
  slide.addText("If today is the 8th, comparing 8 days of this month against a full 31-day prior month always makes the current month look worse.", {
    x: 8.0, y: 4.15, w: 4.4, h: 1.15, fontFace: FONT_BODY, fontSize: 12.5, color: "D7F0EC", lineSpacingMultiple: 1.3,
  });
  slide.addText("The script auto-trims the previous month to the same day-range, so the comparison is genuinely apples-to-apples.", {
    x: 8.0, y: 5.35, w: 4.4, h: 1.15, fontFace: FONT_BODY, fontSize: 12.5, bold: true, color: C.white, lineSpacingMultiple: 1.3,
  });

  pageNum(slide, 4, TOTAL);
}

// ============================================================
// SLIDE 5 — GMR data flow / KPI health
// ============================================================
{
  const slide = pres.addSlide();
  bgSlide(slide, C.bg);
  kicker(slide, "Project 1 of 2");
  slide.addText("Attendance health, ranked automatically", {
    x: 0.6, y: 0.85, w: 12.0, h: 0.7, fontFace: FONT_HEAD, fontSize: 28, bold: true, color: C.ink,
  });
  slide.addText("Every circle and project is auto-classified by attendance %, so nothing needs manual flagging.", {
    x: 0.6, y: 1.55, w: 11.5, h: 0.5, fontFace: FONT_BODY, fontSize: 14, color: C.slate,
  });

  const tiers = [
    { label: "Healthy", range: "≥ 90%", color: "1E7A4F", bg: "E3F5EA" },
    { label: "Moderate", range: "≥ 80%", color: "9A7B0A", bg: "FBF3D9" },
    { label: "Needs Improvement", range: "≥ 70%", color: "B25A1E", bg: "FBE9DC" },
    { label: "Critical", range: "< 70%", color: "B23A2E", bg: "FBE1DE" },
  ];
  const cw = 2.85, gx = 0.28, sx = 0.6, y = 2.5;
  tiers.forEach((t, i) => {
    const x = sx + i * (cw + gx);
    slide.addShape("roundRect", { x, y, w: cw, h: 2.0, rectRadius: 0.12, fill: { color: t.bg }, line: { type: "none" } });
    slide.addText(t.range, { x, y: y + 0.25, w: cw, h: 0.7, fontFace: FONT_HEAD, fontSize: 26, bold: true, color: t.color, align: "center" });
    slide.addText(t.label, { x, y: y + 1.05, w: cw, h: 0.6, fontFace: FONT_BODY, fontSize: 14, bold: true, color: t.color, align: "center" });
  });

  slide.addShape("roundRect", { x: 0.6, y: 4.85, w: 12.1, h: 1.8, rectRadius: 0.12, fill: { color: C.cardBg2 }, line: { type: "none" } });
  slide.addText("Same source of truth, two views", { x: 0.95, y: 5.05, w: 11.4, h: 0.4, fontFace: FONT_HEAD, fontSize: 15, bold: true, color: C.primary });
  slide.addText("A single buildComparisonData() function feeds both the email report and the auto-generated \u201cMonthly Comparison\u201d sheet — so the two views can never drift out of sync, and the summary sheet is fully rebuilt (not appended) on every run, so it's always safe to re-run.", {
    x: 0.95, y: 5.45, w: 11.4, h: 1.1, fontFace: FONT_BODY, fontSize: 13, color: C.ink, lineSpacingMultiple: 1.3,
  });

  pageNum(slide, 5, TOTAL);
}

// ============================================================
// SLIDE 6 — Project 2: Apraava overview
// ============================================================
{
  const slide = pres.addSlide();
  bgSlide(slide, C.bg);
  kicker(slide, "Project 2 of 2");
  slide.addText("Apraava Workforce Mandate & Attendance Report", {
    x: 0.6, y: 0.85, w: 12.0, h: 0.7, fontFace: FONT_HEAD, fontSize: 27, bold: true, color: C.ink,
  });
  slide.addText("State-wise workforce mandate vs. achieved headcount across Gujarat, Rajasthan & Madhya Pradesh.", {
    x: 0.6, y: 1.55, w: 11.5, h: 0.5, fontFace: FONT_BODY, fontSize: 14, color: C.slate,
  });

  const items = [
    { n: "1", t: "Mandate vs. achieved", d: "Sanctioned headcount vs. current active technicians, per state, with fulfillment %." },
    { n: "2", t: "Deployment pipeline", d: "Tracks new hiring separately from sanctioned strength — requirement, deployed, pending, ranked by largest gap." },
    { n: "3", t: "Location-level detail", d: "Present / absent mandays and attendance % for every operating location, grouped by state." },
    { n: "4", t: "Daily trend chart", d: "A real embedded chart image (inline cid: attachment) — not HTML bars — for a true continuous trend line." },
    { n: "5", t: "Auto-ranked actions", d: "Weakest location, pipeline gaps, and trend direction are computed and surfaced as a prioritized follow-up list." },
  ];

  const rowH = 0.85, y0 = 2.35;
  items.forEach((it, i) => {
    const y = y0 + i * rowH;
    slide.addShape("ellipse", { x: 0.6, y: y, w: 0.5, h: 0.5, fill: { color: C.secondary }, line: { type: "none" } });
    slide.addText(it.n, { x: 0.6, y: y, w: 0.5, h: 0.5, fontFace: FONT_BODY, fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle" });
    slide.addText(it.t, { x: 1.3, y: y - 0.02, w: 3.1, h: 0.55, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: C.ink, valign: "middle" });
    slide.addText(it.d, { x: 4.55, y: y - 0.02, w: 8.15, h: 0.68, fontFace: FONT_BODY, fontSize: 12.5, color: C.slate, valign: "middle", lineSpacingMultiple: 1.15 });
  });

  pageNum(slide, 6, TOTAL);
}

// ============================================================
// SLIDE 7 — Apraava engineering notes
// ============================================================
{
  const slide = pres.addSlide();
  bgSlide(slide, C.bg);
  kicker(slide, "Project 2 of 2");
  slide.addText("Built to survive messy real-world data", {
    x: 0.6, y: 0.85, w: 12.0, h: 0.7, fontFace: FONT_HEAD, fontSize: 28, bold: true, color: C.ink,
  });

  const notes = [
    { t: "Config-driven", d: "Sheet names, recipients, status tokens (P / A / Wo / WO) all live in one CONFIG object — adapting to a new state or spreadsheet needs zero logic changes." },
    { t: "Name normalization", d: "\u201cGujarat\u201d / \u201cMP\u201d / \u201cMadhya Pradesh\u201d and mixed-case city names (\u201cRAJKOT\u201d, \u201cjunagadh\u201d) are normalized so lookups and display stay consistent." },
    { t: "Resilient totals", d: "Deployment pipeline auto-detects a dedicated totals row when present, and falls back to summing per-location rows when it isn't." },
    { t: "Week-offs excluded from trend", d: "The daily attendance line reflects true working-day performance, not diluted by rest days." },
  ];

  const cw = 5.85, gx = 0.4, sx = 0.6, y = 1.85, ch = 2.35;
  notes.forEach((n, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = sx + col * (cw + gx);
    const yy = y + row * (ch + 0.3);
    slide.addShape("roundRect", { x, y: yy, w: cw, h: ch, rectRadius: 0.12, fill: { color: C.cardBg }, line: { type: "none" } });
    slide.addText(n.t, { x: x + 0.3, y: yy + 0.25, w: cw - 0.6, h: 0.5, fontFace: FONT_HEAD, fontSize: 15, bold: true, color: C.primary });
    slide.addText(n.d, { x: x + 0.3, y: yy + 0.85, w: cw - 0.6, h: 1.35, fontFace: FONT_BODY, fontSize: 13, color: C.ink, lineSpacingMultiple: 1.3 });
  });

  pageNum(slide, 7, TOTAL);
}

// ============================================================
// SLIDE 8 — Business impact (stat callouts)
// ============================================================
{
  const slide = pres.addSlide();
  bgSlide(slide, C.ink);
  kicker(slide, "Business Impact", { color: C.accent });
  slide.addText("What automation actually buys the organization", {
    x: 0.6, y: 0.85, w: 12.0, h: 0.7, fontFace: FONT_HEAD, fontSize: 28, bold: true, color: C.white,
  });

  const stats = [
    { big: "0", label: "Manual report\nbuilding hours\nper cycle" },
    { big: "100%", label: "Consistent\nformatting,\nevery send" },
    { big: "Same-day", label: "Leadership\nvisibility instead\nof multi-day lag" },
    { big: "Full", label: "Send log audit\ntrail (GMR\nComparison Log)" },
  ];
  const cw = 2.85, gx = 0.28, sx = 0.6, y = 2.15;
  stats.forEach((s, i) => {
    const x = sx + i * (cw + gx);
    slide.addShape("roundRect", { x, y, w: cw, h: 3.2, rectRadius: 0.12, fill: { color: "0F3A3E" }, line: { type: "none" } });
    slide.addText(s.big, { x, y: y + 0.35, w: cw, h: 1.1, fontFace: FONT_HEAD, fontSize: 40, bold: true, color: C.accent, align: "center" });
    slide.addText(s.label, { x: x + 0.2, y: y + 1.55, w: cw - 0.4, h: 1.4, fontFace: FONT_BODY, fontSize: 13, color: "D7F0EC", align: "center", lineSpacingMultiple: 1.25 });
  });

  slide.addText("Both scripts also eliminate calculation error at the source: attendance %, deltas, and deficits are computed once, in code, and reused identically across the email and the summary sheet.", {
    x: 0.6, y: 5.7, w: 12.0, h: 0.9, fontFace: FONT_BODY, fontSize: 13.5, italic: true, color: "9DC7C2", lineSpacingMultiple: 1.3,
  });

  pageNum(slide, 8, TOTAL);
}

// ============================================================
// SLIDE 9 — Tech stack / how to adopt
// ============================================================
{
  const slide = pres.addSlide();
  bgSlide(slide, C.bg);
  kicker(slide, "Getting Started");
  slide.addText("Runs entirely inside Google Sheets — no infrastructure", {
    x: 0.6, y: 0.85, w: 12.0, h: 0.7, fontFace: FONT_HEAD, fontSize: 27, bold: true, color: C.ink,
  });

  // left: stack
  slide.addShape("roundRect", { x: 0.6, y: 1.9, w: 5.6, h: 4.7, rectRadius: 0.12, fill: { color: C.cardBg }, line: { type: "none" } });
  slide.addText("Tech stack", { x: 0.95, y: 2.15, w: 5.0, h: 0.4, fontFace: FONT_HEAD, fontSize: 16, bold: true, color: C.primary });
  const stack = [
    "Google Apps Script (V8 runtime) — the entire engine",
    "SpreadsheetApp — reads sheet ranges, writes summary tabs",
    "HtmlService — renders the dialog UI and email preview",
    "MailApp / GmailApp — sends the final HTML email",
    "ScriptApp triggers — optional scheduled, unattended sends",
    "Zero npm packages, zero external servers, zero API keys",
  ];
  slide.addText(
    stack.map((f, i) => ({ text: f, options: { bullet: { code: "2022" }, breakLine: i < stack.length - 1, color: C.ink, fontSize: 13.5 } })),
    { x: 0.95, y: 2.65, w: 5.0, h: 3.8, fontFace: FONT_BODY, valign: "top", paraSpaceAfter: 12, lineSpacingMultiple: 1.25 }
  );

  // right: setup steps
  slide.addShape("roundRect", { x: 6.5, y: 1.9, w: 6.2, h: 4.7, rectRadius: 0.12, fill: { color: C.primary }, line: { type: "none" } });
  slide.addText("5-step setup", { x: 6.85, y: 2.15, w: 5.5, h: 0.4, fontFace: FONT_HEAD, fontSize: 16, bold: true, color: C.white });
  const steps = [
    "Open the target Google Sheet",
    "Extensions → Apps Script",
    "Paste the project's .gs file into the editor",
    "Edit the CONFIG / PROJECT_THEME block to match your sheet",
    "Run once to authorize, then send, preview, or schedule",
  ];
  steps.forEach((s, i) => {
    const y = 2.75 + i * 0.72;
    slide.addShape("ellipse", { x: 6.85, y, w: 0.42, h: 0.42, fill: { color: C.accent }, line: { type: "none" } });
    slide.addText(String(i + 1), { x: 6.85, y, w: 0.42, h: 0.42, fontFace: FONT_BODY, fontSize: 13, bold: true, color: C.ink, align: "center", valign: "middle" });
    slide.addText(s, { x: 7.4, y: y - 0.06, w: 5.1, h: 0.55, fontFace: FONT_BODY, fontSize: 13, color: C.white, valign: "middle", lineSpacingMultiple: 1.2 });
  });

  pageNum(slide, 9, TOTAL);
}

// ============================================================
// SLIDE 10 — Closing / repo
// ============================================================
{
  const slide = pres.addSlide();
  bgSlide(slide, C.ink);

  slide.addText("Both projects are in the repo, ready to adapt", {
    x: 0.8, y: 2.5, w: 11.7, h: 1.1, fontFace: FONT_HEAD, fontSize: 34, bold: true, color: C.white, lineSpacingMultiple: 1.1,
  });
  slide.addText("gmr-monthly-comparison/   and   apraava-workforce-report/", {
    x: 0.8, y: 3.65, w: 11.7, h: 0.5, fontFace: FONT_BODY, fontSize: 16, color: C.accent, bold: true,
  });
  slide.addText("Each folder includes the full .gs source and a README covering setup, sheet structure, configuration, and the engineering decisions behind it.", {
    x: 0.8, y: 4.2, w: 10.5, h: 0.9, fontFace: FONT_BODY, fontSize: 14, color: "BFDBD6", lineSpacingMultiple: 1.35,
  });

  slide.addShape("ellipse", { x: 10.6, y: 0.4, w: 2.4, h: 2.4, fill: { color: C.primary, transparency: 60 }, line: { type: "none" } });
  slide.addShape("ellipse", { x: 11.6, y: 1.6, w: 1.3, h: 1.3, fill: { color: C.accent, transparency: 40 } , line: { type: "none" } });

  pageNum(slide, 10, TOTAL);
}

pres.writeFile({ fileName: "/home/claude/workforce-automation-presentation.pptx" }).then(() => {
  console.log("done");
});
