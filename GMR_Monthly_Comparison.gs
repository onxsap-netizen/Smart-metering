// ============================================================
//  GMR Monthly Comparison Dashboard  — v3 (Gmail-safe bar charts)
//  Layout : Kashi | Triveni side-by-side header cards
//           KPI rows with HTML table bar charts (no SVG — Gmail strips SVG)
//  Date   : Fair comparison – current month uses same day-range
//           as previous month (e.g. 01-08 Jun vs 01-08 May)
//  Style  : No emojis, enterprise-grade typography, muted palette
// ============================================================

const MCD_SHEETS = {
  CURRENT:  "Raw_Attendance",
  PREVIOUS: "Previous Month",
  REPORT:   "Monthly Comparison",
  LOG:      "Comparison Log",
};

// PROJECT_THEME declared in GMR Master script.
// Uncomment if running standalone:
/*
const PROJECT_THEME = {
  Kashi:   { color: "#1D4ED8", light: "#EFF6FF", grad: "linear-gradient(135deg,#1E3A8A,#1D4ED8)", label: "KASHI PROJECT" },
  Triveni: { color: "#7C3AED", light: "#F5F3FF", grad: "linear-gradient(135deg,#4C1D95,#7C3AED)", label: "TRIVENI PROJECT" },
};
*/

// ──────────────────────────────────────────────
//  MENU
// ──────────────────────────────────────────────
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Monthly Comparison")
      .addItem("Send Comparison Email",   "sendComparisonEmail")
      .addItem("Preview Email",            "previewComparisonEmail")
      .addSeparator()
      .addItem("Write Summary Sheet",     "writeComparisonSummarySheet")
      .addToUi();
  } catch (e) {}
}

// ──────────────────────────────────────────────
//  EMAIL DIALOG
// ──────────────────────────────────────────────
function sendComparisonEmail() {
  const html = `<!DOCTYPE html>
  <html><head><style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Segoe UI, Arial, sans-serif; background: #F1F5F9; }
    .header { background: linear-gradient(135deg,#0F172A,#1E3A8A); padding: 18px 24px 14px; color:#fff; }
    .header h2 { font-size:15px; font-weight:700; margin-bottom:3px; }
    .header p  { font-size:11px; color:#93C5FD; }
    .body { padding:20px 24px; background:#fff; }
    .field { margin-bottom:14px; }
    label { display:block; font-size:11px; font-weight:700; color:#475569; text-transform:uppercase; letter-spacing:.5px; margin-bottom:5px; }
    .req { color:#DC2626; }
    input[type=text], input[type=email] { width:100%; padding:8px 11px; border:1.5px solid #CBD5E1; border-radius:6px; font-size:13px; color:#1E293B; outline:none; }
    .hint { font-size:11px; color:#94A3B8; margin-top:4px; }
    #errBox { display:none; background:#FEF2F2; border:1px solid #FCA5A5; border-radius:6px; padding:8px 12px; font-size:12px; color:#991B1B; margin-bottom:12px; }
    #successBox { display:none; background:#F0FDF4; border:1.5px solid #86EFAC; border-radius:8px; padding:18px 16px; text-align:center; margin-bottom:12px; }
    .footer { background:#F8FAFC; border-top:1px solid #E2E8F0; padding:14px 24px; }
    .footer-inner { display:table; width:100%; }
    .fl { display:table-cell; vertical-align:middle; }
    .fr { display:table-cell; vertical-align:middle; text-align:right; }
    .btn { padding:8px 22px; border:none; border-radius:6px; font-size:13px; font-weight:700; cursor:pointer; margin-left:8px; }
    .btn-cancel { background:#E2E8F0; color:#475569; }
    .btn-send   { background:#1D4ED8; color:#fff; }
    .btn-send:disabled { background:#93C5FD; cursor:not-allowed; }
    .tag { display:inline-block; background:#EFF6FF; color:#1D4ED8; border:1px solid #BFDBFE; border-radius:12px; padding:2px 10px 2px 8px; font-size:11px; margin:2px 3px 2px 0; }
    .tag-x { cursor:pointer; color:#93C5FD; font-size:13px; margin-left:4px; }
  </style></head><body>
    <div class="header">
      <h2>Send Monthly Comparison Report</h2>
      <p>Current Month vs Previous Month — KPIs, Trends &amp; Improvements</p>
    </div>
    <div class="body">
      <div id="errBox"></div>
      <div id="successBox">
        <div style="font-size:28px;margin-bottom:6px;">✓</div>
        <div style="font-size:14px;font-weight:700;color:#166534;">Email sent successfully!</div>
        <div style="font-size:11px;color:#16a34a;margin-top:4px;">This window will close automatically...</div>
      </div>
      <div id="formFields">
        <div class="field">
          <label>Recipient Name <span class="req">*</span></label>
          <input type="text" id="rname" value="GMR Operations Head" placeholder="e.g. Operations Head">
        </div>
        <div class="field">
          <label>To Email <span class="req">*</span></label>
          <input type="email" id="toEmail" placeholder="e.g. manager@company.com">
          <div class="hint">Primary recipient — required</div>
        </div>
        <div class="field">
          <label>CC Emails <span style="font-weight:400;color:#94A3B8;">(optional)</span></label>
          <input type="email" id="ccInput" placeholder="Type email and press Enter or comma">
          <div class="hint">Press Enter or comma after each address</div>
          <div id="ccTags" style="margin-top:6px;"></div>
          <input type="hidden" id="ccHidden" value="">
        </div>
      </div>
    </div>
    <div class="footer"><div class="footer-inner">
      <div class="fl"><div id="spinner" style="display:none;font-size:12px;font-weight:600;color:#1D4ED8;">Sending, please wait...</div></div>
      <div class="fr">
        <button class="btn btn-cancel" id="cancelBtn" onclick="google.script.host.close()">Cancel</button>
        <button class="btn btn-send"   id="sendBtn"   onclick="doSend()">Send Email</button>
      </div>
    </div></div>
    <script>
      var ccList = [];
      function addTag(email) {
        email = email.trim();
        if (!email || email.indexOf("@") === -1) return;
        if (ccList.indexOf(email) !== -1) return;
        ccList.push(email); renderTags();
        document.getElementById("ccInput").value = ""; syncHidden();
      }
      function removeTag(email) {
        ccList = ccList.filter(function(e) { return e !== email; });
        renderTags(); syncHidden();
      }
      function renderTags() {
        var row = document.getElementById("ccTags"); row.innerHTML = "";
        ccList.forEach(function(email) {
          var d = document.createElement("div"); d.className = "tag";
          d.innerHTML = email + "<span class='tag-x' onclick='removeTag(\\"" + email + "\\")'>x</span>";
          row.appendChild(d);
        });
      }
      function syncHidden() { document.getElementById("ccHidden").value = ccList.join(","); }
      document.getElementById("ccInput").addEventListener("keydown", function(e) {
        if (e.key === "Enter" || e.key === ",") { e.preventDefault(); addTag(this.value.replace(/,$/, "")); }
      });
      document.getElementById("ccInput").addEventListener("blur", function() {
        var v = this.value.trim().replace(/,$/, ""); if (v) addTag(v);
      });
      function showErr(msg) {
        var box = document.getElementById("errBox"); box.textContent = msg; box.style.display = "block";
      }
      function doSend() {
        var name    = document.getElementById("rname").value.trim();
        var toEmail = document.getElementById("toEmail").value.trim();
        var pending = document.getElementById("ccInput").value.trim().replace(/,$/, "");
        if (pending && pending.indexOf("@") !== -1) addTag(pending);
        var ccEmails = document.getElementById("ccHidden").value.trim();
        if (!name) { showErr("Recipient Name is required."); return; }
        if (!toEmail || toEmail.indexOf("@") === -1) { showErr("Please enter a valid To email address."); return; }
        document.getElementById("errBox").style.display = "none";
        var btn = document.getElementById("sendBtn");
        var cancelBtn = document.getElementById("cancelBtn");
        var spinner = document.getElementById("spinner");
        btn.disabled = true; btn.textContent = "Sending...";
        cancelBtn.disabled = true; spinner.style.display = "block";
        google.script.run
          .withSuccessHandler(function() {
            spinner.style.display = "none";
            btn.style.display = "none"; cancelBtn.style.display = "none";
            document.getElementById("formFields").style.display = "none";
            document.getElementById("successBox").style.display = "block";
            setTimeout(function() { google.script.host.close(); }, 2500);
          })
          .withFailureHandler(function(err) {
            spinner.style.display = "none";
            btn.disabled = false; btn.textContent = "Send Email"; cancelBtn.disabled = false;
            showErr("FAILED: " + (err ? (err.message || JSON.stringify(err)) : "Unknown error"));
          })
          .sendComparisonEmailFromDialog(name, toEmail, ccEmails);
      }
    <\/script>
  </body></html>`;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(460).setHeight(530),
    "Send Monthly Comparison Email"
  );
}

// ──────────────────────────────────────────────
//  CALLED BY DIALOG
// ──────────────────────────────────────────────
function sendComparisonEmailFromDialog(recipientName, toEmail, ccEmails) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const data = buildComparisonData(ss);

  const subject = "GMR Monthly Comparison Report | " + data.currentPeriod + " vs " + data.prevPeriod;
  const html    = buildComparisonEmailHTML(data, recipientName);

  const ccArr = ccEmails
    ? ccEmails.split(",").map(function(e) { return e.trim(); }).filter(function(e) { return e.indexOf("@") !== -1; })
    : [];

  const opts = { to: toEmail, subject: subject, htmlBody: html, name: "GMR HR Operations Team" };
  if (ccArr.length) opts.cc = ccArr.join(",");

  MailApp.sendEmail(opts);

  let logSheet = ss.getSheetByName(MCD_SHEETS.LOG);
  if (!logSheet) logSheet = ss.insertSheet(MCD_SHEETS.LOG);
  const tz = Session.getScriptTimeZone();
  logSheet.appendRow([
    Utilities.formatDate(new Date(), tz, "dd MMM yyyy HH:mm"),
    "SENT", toEmail, subject
  ]);

  return "ok";
}

// ──────────────────────────────────────────────
//  PREVIEW
// ──────────────────────────────────────────────
function previewComparisonEmail() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const data = buildComparisonData(ss);
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(buildComparisonEmailHTML(data, "Operations Head")).setWidth(980).setHeight(800),
    "Preview — Monthly Comparison | " + data.currentPeriod + " vs " + data.prevPeriod
  );
}

// ──────────────────────────────────────────────
//  DATA BUILDER
//  KEY CHANGE: prevFairDays = same number of days as current month
//              so comparison is apples-to-apples (e.g. 01-08 May only)
// ──────────────────────────────────────────────
function buildComparisonData(ss) {
  const tz = Session.getScriptTimeZone();

  function fmt(d)      { return Utilities.formatDate(d, tz, "dd MMM yyyy"); }
  function fmtShort(d) { return Utilities.formatDate(d, tz, "dd MMM"); }
  function fmtDay(d)   { return Utilities.formatDate(d, tz, "EEE"); }

  function parseSheet(sheetName, cutoffDay) {
    // cutoffDay: if set, only include date columns where day-of-month <= cutoffDay
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error('Sheet "' + sheetName + '" not found.');
    const raw     = sheet.getDataRange().getValues();
    const headers = raw[0];

    function col(n) { return headers.indexOf(n); }

    const IDX = {
      name:      col("User Name"),
      role:      col("User Role"),
      project:   col("Project"),
      status:    col("Status"),
      totalPres: col("Total Present"),
      totalAbs:  col("Total Absent"),
      weekOff:   col("Total Week Off"),
      circle:    col("Cricle"),
    };

    // Collect all date columns; apply cutoffDay filter for previous sheet
    let dateCols = [];
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] instanceof Date) {
        const d = new Date(headers[i]);
        if (!cutoffDay || d.getDate() <= cutoffDay) {
          dateCols.push({ col: i, date: d });
        }
      }
    }
    dateCols.sort(function(a, b) { return a.date - b.date; });

    const periodLabel = dateCols.length > 0
      ? fmt(dateCols[0].date) + " \u2013 " + fmt(dateCols[dateCols.length - 1].date)
      : sheetName;

    const projectMap = {};

    for (let r = 1; r < raw.length; r++) {
      const row    = raw[r];
      const status = String(row[IDX.status] || "").trim();
      const pName  = String(row[IDX.project] || "").trim();
      const cName  = String(row[IDX.circle]  || "").trim();
      const uRole  = String(row[IDX.role]    || "").trim().toUpperCase();
      if (!pName || !status || (status !== "Active" && status !== "Inactive")) continue;

      if (!projectMap[pName]) {
        projectMap[pName] = {
          name: pName,
          total: 0, active: 0, inactive: 0,
          totalPresent: 0, totalAbsent: 0, totalWeekOff: 0,
          active0Present: 0,
          roleBreakdown: {}, circleBreakdown: {}, daywise: {},
        };
      }
      const p = projectMap[pName];
      p.total++;

      if (status === "Active") {
        p.active++;

        // For fair comparison: recalculate present/absent/weekoff from date columns only
        let tp = 0, ta = 0, wo = 0;
        if (cutoffDay) {
          // Re-derive from day columns (fair window)
          for (let di = 0; di < dateCols.length; di++) {
            const val = String(row[dateCols[di].col] || "").trim();
            if      (val === "Present")  tp++;
            else if (val === "Absent")   ta++;
            else if (val === "Week Off") wo++;
          }
        } else {
          // Current month: use summary columns
          tp = parseFloat(row[IDX.totalPres]) || 0;
          ta = parseFloat(row[IDX.totalAbs])  || 0;
          wo = parseFloat(row[IDX.weekOff])   || 0;
        }

        p.totalPresent  += tp;
        p.totalAbsent   += ta;
        p.totalWeekOff  += wo;
        if (tp === 0) p.active0Present++;

        // Role breakdown
        if (!p.roleBreakdown[uRole]) {
          p.roleBreakdown[uRole] = { active: 0, totalPresent: 0, totalAbsent: 0, totalWeekOff: 0, active0: 0 };
        }
        const rb = p.roleBreakdown[uRole];
        rb.active++; rb.totalPresent += tp; rb.totalAbsent += ta; rb.totalWeekOff += wo;
        if (tp === 0) rb.active0++;

        // Circle breakdown
        if (cName) {
          if (!p.circleBreakdown[cName]) {
            p.circleBreakdown[cName] = { active: 0, inactive: 0, totalPresent: 0, totalAbsent: 0, totalWeekOff: 0, active0: 0 };
          }
          const cb = p.circleBreakdown[cName];
          cb.active++; cb.totalPresent += tp; cb.totalAbsent += ta; cb.totalWeekOff += wo;
          if (tp === 0) cb.active0++;
        }

        // Day-wise
        for (let di = 0; di < dateCols.length; di++) {
          const dc  = dateCols[di];
          const dk  = fmtShort(dc.date);
          if (!p.daywise[dk]) {
            p.daywise[dk] = { present: 0, absent: 0, weekOff: 0, dayName: fmtDay(dc.date), dateObj: dc.date };
          }
          const val = String(row[dc.col] || "").trim();
          if      (val === "Present")  p.daywise[dk].present++;
          else if (val === "Absent")   p.daywise[dk].absent++;
          else if (val === "Week Off") p.daywise[dk].weekOff++;
        }

      } else {
        p.inactive++;
        if (cName) {
          if (!p.circleBreakdown[cName]) {
            p.circleBreakdown[cName] = { active: 0, inactive: 0, totalPresent: 0, totalAbsent: 0, totalWeekOff: 0, active0: 0 };
          }
          p.circleBreakdown[cName].inactive++;
        }
      }
    }

    // Derived metrics
    for (const p of Object.values(projectMap)) {
      const total    = p.totalPresent + p.totalAbsent + p.totalWeekOff;
      const achieved = p.totalPresent + p.totalWeekOff;
      p.attendancePct = total > 0 ? achieved / total : 0;
      p.absencePct    = total > 0 ? p.totalAbsent / total : 0;
      for (const rb of Object.values(p.roleBreakdown)) {
        const rt = rb.totalPresent + rb.totalAbsent + rb.totalWeekOff;
        rb.attendancePct = rt > 0 ? (rb.totalPresent + rb.totalWeekOff) / rt : 0;
        rb.absencePct    = rt > 0 ? rb.totalAbsent / rt : 0;
      }
      for (const cb of Object.values(p.circleBreakdown)) {
        const ct = cb.totalPresent + cb.totalAbsent + cb.totalWeekOff;
        cb.attendancePct = ct > 0 ? (cb.totalPresent + cb.totalWeekOff) / ct : 0;
        cb.absencePct    = ct > 0 ? cb.totalAbsent / ct : 0;
      }
    }

    return {
      projects: Object.values(projectMap).sort(function(a, b) { return a.name.localeCompare(b.name); }),
      period:   periodLabel,
      dateCols: dateCols,
    };
  }

  // Step 1: parse current to find out how many days it spans
  const current = parseSheet(MCD_SHEETS.CURRENT, null);

  // Step 2: figure out the max day-of-month in current period
  const currentDayCols = current.dateCols;
  const maxDayOfMonth  = currentDayCols.length > 0
    ? Math.max.apply(null, currentDayCols.map(function(dc) { return dc.date.getDate(); }))
    : 31;

  // Step 3: parse previous with cutoff = same max day (fair window)
  const previous = parseSheet(MCD_SHEETS.PREVIOUS, maxDayOfMonth);

  return {
    current:       current.projects,
    previous:      previous.projects,
    currentPeriod: current.period,
    prevPeriod:    previous.period,
    generatedOn:   Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy, hh:mm a"),
    fairDayCutoff: maxDayOfMonth,
  };
}

// ──────────────────────────────────────────────
//  EMAIL HTML BUILDER — v2 Professional
// ──────────────────────────────────────────────
function buildComparisonEmailHTML(data, recipientName) {
  const curr     = data.current;
  const prev     = data.previous;
  const genOn    = data.generatedOn;
  const curPer   = data.currentPeriod;
  const prePer   = data.prevPeriod;
  const headName = recipientName || "Operations Head";
  const fairNote = data.fairDayCutoff
    ? "Previous month figures reflect only 01\u2013" + data.fairDayCutoff + " of that month, matching the current reporting window."
    : "";

  // ── Colour tokens ───────────────────────────
  const C = {
    bg:       "#F4F6F9",
    surface:  "#FFFFFF",
    border:   "#DDE3ED",
    head:     "#0B1E3D",
    sub:      "#1E3A8A",
    text:     "#1E293B",
    muted:    "#64748B",
    faint:    "#94A3B8",
    green:    "#15803D",
    greenBg:  "#DCFCE7",
    red:      "#B91C1C",
    redBg:    "#FEE2E2",
    amber:    "#B45309",
    amberBg:  "#FEF3C7",
    blue:     "#1D4ED8",
    blueBg:   "#EFF6FF",
    kashi:    "#1D4ED8",
    triveni:  "#6D28D9",
  };

  function findProject(list, name) {
    return list.find(function(p) { return p.name === name; }) || null;
  }

  const allProjects = [];
  const seen = {};
  curr.forEach(function(p) { if (!seen[p.name]) { seen[p.name] = 1; allProjects.push(p.name); } });
  prev.forEach(function(p) { if (!seen[p.name]) { seen[p.name] = 1; allProjects.push(p.name); } });

  function pct(val) { return (val * 100).toFixed(1); }

  function deltaSpan(cur, pre, invertGood) {
    if (pre === null || pre === undefined || isNaN(pre)) return "";
    const d = cur - pre;
    if (Math.abs(d) < 0.05) return '<span style="color:' + C.faint + ';font-size:10px;">No change</span>';
    let good  = d > 0;
    if (invertGood) good = !good;
    const color = good ? C.green : C.red;
    const arrow = d > 0 ? "▲" : "▼";
    return '<span style="color:' + color + ';font-size:10px;font-weight:700;">' + arrow + " " + (d > 0 ? "+" : "") + d.toFixed(1) + "</span>";
  }

  function statusBadge(pctVal) {
    const v = parseFloat(pctVal) || 0;
    if (v >= 90) return { bg: C.greenBg, fg: C.green,  label: "HEALTHY" };
    if (v >= 80) return { bg: C.amberBg, fg: C.amber,  label: "MODERATE" };
    if (v >= 70) return { bg: "#FFF7ED", fg: "#C2410C", label: "NEEDS IMPROVEMENT" };
    return        { bg: C.redBg,   fg: C.red,   label: "CRITICAL" };
  }

  // ── HTML Table Bar Chart (Gmail-safe, no SVG) ──
  // Uses nested tables with fixed heights — works in all email clients.
  // Each bar = a tall <td> (spacer above) + coloured <td> (the bar itself).
  // The chart area height is fixed at CHART_H px; bars scale within it.
  function htmlBarChart(daywise, barColor, absentColor) {
    const entries = Object.entries(daywise).sort(function(a, b) {
      return (a[1].dateObj || 0) - (b[1].dateObj || 0);
    });
    if (!entries.length) {
      return '<p style="color:' + C.faint + ';font-size:11px;padding:12px 0;">No daily data available.</p>';
    }

    const CHART_H = 80;  // px — total bar column height
    const BAR_W   = 28;  // px — width of each bar column (incl. small gap)

    const maxPresent = Math.max.apply(null, entries.map(function(e) { return e[1].present || 0; })) || 1;

    // Build one <td> per day inside a single outer table row
    // Structure per cell (top-to-bottom):
    //   [value label row]  — 12px font, colour of bar
    //   [spacer td]        — fills remaining height above bar
    //   [present bar td]   — green/blue solid block
    //   [absent bar td]    — red thin block (proportional)
    //   [date label td]    — "01 Jun"
    //   [day label td]     — "Mon"

    let cols = "";
    entries.forEach(function(entry) {
      const dk  = entry[0];
      const d   = entry[1];
      const isWO = d.present === 0 && d.absent === 0 && d.weekOff > 0;

      if (isWO) {
        // Week-off column: grey stub at bottom
        const stubH = 6;
        const spH   = CHART_H - stubH;
        cols += '<td style="width:' + BAR_W + 'px;padding:0 2px;vertical-align:bottom;text-align:center;">'
          + '<table width="' + (BAR_W - 4) + '" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin:0 auto;">'
            // value label (blank for WO)
            + '<tr><td style="height:12px;font-size:8px;font-weight:700;color:#CBD5E1;text-align:center;font-family:Arial,sans-serif;">WO</td></tr>'
            // spacer
            + '<tr><td style="height:' + (spH - 12) + 'px;font-size:0;">&nbsp;</td></tr>'
            // stub bar
            + '<tr><td style="height:' + stubH + 'px;background:#CBD5E1;border-radius:2px 2px 0 0;font-size:0;">&nbsp;</td></tr>'
          + '</table>'
          // date label
          + '<div style="font-size:7px;color:#64748B;text-align:center;margin-top:3px;line-height:1.3;font-family:Arial,sans-serif;">'
            + dk.replace(" ", "<br>")
          + '</div>'
          + '<div style="font-size:6px;color:#94A3B8;text-align:center;font-family:Arial,sans-serif;">' + (d.dayName || "") + '</div>'
        + '</td>';
        return;
      }

      const presVal  = d.present || 0;
      const absVal   = d.absent  || 0;
      // Scale present bar height (min 2px so it's always visible if > 0)
      const presH    = presVal > 0 ? Math.max(2, Math.round(presVal / maxPresent * CHART_H * 0.75)) : 0;
      // Absent bar: proportional to present scale (max 20% of chart height)
      const absH     = absVal  > 0 ? Math.max(2, Math.round(absVal  / maxPresent * CHART_H * 0.75)) : 0;
      const barTotal = presH + absH;
      const spacerH  = Math.max(0, CHART_H - barTotal - 12); // 12 = value label row

      cols += '<td style="width:' + BAR_W + 'px;padding:0 2px;vertical-align:bottom;text-align:center;">'
        + '<table width="' + (BAR_W - 4) + '" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin:0 auto;">'
          // value label
          + '<tr><td style="height:12px;font-size:8px;font-weight:700;color:' + barColor + ';text-align:center;font-family:Arial,sans-serif;">' + presVal + '</td></tr>'
          // spacer (pushes bars to bottom)
          + (spacerH > 0 ? '<tr><td style="height:' + spacerH + 'px;font-size:0;">&nbsp;</td></tr>' : '')
          // present bar
          + (presH > 0 ? '<tr><td style="height:' + presH + 'px;background:' + barColor + ';border-radius:' + (absH > 0 ? "2px 2px 0 0" : "2px") + ';font-size:0;">&nbsp;</td></tr>' : '')
          // absent bar (stacked below)
          + (absH  > 0 ? '<tr><td style="height:' + absH  + 'px;background:' + absentColor + ';border-radius:0 0 2px 2px;font-size:0;">&nbsp;</td></tr>' : '')
        + '</table>'
        // date label
        + '<div style="font-size:7px;color:#64748B;text-align:center;margin-top:3px;line-height:1.3;font-family:Arial,sans-serif;">'
          + dk.replace(" ", "<br>")
        + '</div>'
        + '<div style="font-size:6px;color:#94A3B8;text-align:center;font-family:Arial,sans-serif;">' + (d.dayName || "") + '</div>'
      + '</td>';
    });

    return '<div style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:6px;padding:10px 8px 6px;overflow-x:auto;">'
      + '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;min-width:100%;">'
        + '<tr style="vertical-align:bottom;">' + cols + '</tr>'
      + '</table>'
      + '<div style="margin-top:6px;font-size:9px;color:' + C.muted + ';font-family:Arial,sans-serif;">'
        + '<span style="display:inline-block;width:8px;height:8px;background:' + barColor   + ';border-radius:1px;margin-right:3px;vertical-align:middle;"></span>Present&nbsp;&nbsp;'
        + '<span style="display:inline-block;width:8px;height:8px;background:' + absentColor + ';border-radius:1px;margin-right:3px;vertical-align:middle;"></span>Absent&nbsp;&nbsp;'
        + '<span style="display:inline-block;width:8px;height:8px;background:#CBD5E1;border-radius:1px;margin-right:3px;vertical-align:middle;"></span>Week Off'
      + '</div>'
    + '</div>';
  }

  // ── Section header (no emoji) ────────────────
  function secHdr(title, color) {
    return '<div style="background:' + (color || C.head) + ';color:#FFF;font-size:10px;font-weight:700;'
      + 'letter-spacing:1px;padding:7px 14px;border-radius:4px;margin:18px 0 8px;text-transform:uppercase;'
      + 'font-family:Arial,sans-serif;">' + title + '</div>';
  }

  // ── Thin divider ─────────────────────────────
  function divider() {
    return '<div style="border-top:1px solid ' + C.border + ';margin:18px 0;"></div>';
  }

  // ── KPI Pill (used in the 2-col KPI grid below charts) ──
  function kpiPill(label, curVal, prevVal, isPercent, invertGood) {
    const cv   = parseFloat(curVal)  || 0;
    const pv   = parseFloat(prevVal);
    const dispCur  = isPercent ? cv.toFixed(1) + "%" : Math.round(cv).toLocaleString();
    const dispPrev = isNaN(pv) ? "—" : (isPercent ? pv.toFixed(1) + "%" : Math.round(pv).toLocaleString());
    const dHtml    = isNaN(pv) ? "" : deltaSpan(cv, pv, invertGood);
    return '<td style="padding:8px 10px;vertical-align:top;width:20%;">'
      + '<div style="background:' + C.surface + ';border:1px solid ' + C.border + ';border-radius:6px;padding:10px 8px;text-align:center;">'
        + '<div style="font-size:18px;font-weight:800;color:' + C.text + ';line-height:1.1;">' + dispCur + '</div>'
        + '<div style="font-size:9px;font-weight:700;color:' + C.faint + ';text-transform:uppercase;letter-spacing:.4px;margin:4px 0 5px;">' + label + '</div>'
        + '<div style="font-size:10px;color:' + C.muted + ';">Prev: ' + dispPrev + '</div>'
        + (dHtml ? '<div style="margin-top:3px;">' + dHtml + '</div>' : '')
      + '</div></td>';
  }

  // ═══════════════════════════════════════════════════════════
  //  BUILD: Top side-by-side Project Header Cards
  //  Left = Triveni, Right = Kashi
  // ═══════════════════════════════════════════════════════════
  function projectCard(pName, curP, prevP, themeColor, bgGrad) {
    const cAtt  = curP  ? parseFloat(pct(curP.attendancePct))  : 0;
    const pAtt  = prevP ? parseFloat(pct(prevP.attendancePct)) : null;
    const diff  = pAtt !== null ? cAtt - pAtt : null;
    const st    = statusBadge(cAtt);
    const dStr  = diff !== null && Math.abs(diff) >= 0.05
      ? (diff > 0 ? "▲ +" : "▼ ") + Math.abs(diff).toFixed(1) + "% vs prev period"
      : "No change vs prev period";
    const dColor = diff !== null && diff > 0 ? "#86EFAC" : diff !== null && diff < 0 ? "#FCA5A5" : "#CBD5E1";

    return '<td style="width:50%;padding:4px;vertical-align:top;">'
      + '<div style="background:' + bgGrad + ';border-radius:8px;padding:18px 20px;">'
        // Project label
        + '<div style="font-size:9px;color:#BAE6FD;letter-spacing:1.5px;text-transform:uppercase;font-family:Arial,sans-serif;margin-bottom:6px;">' + pName.toUpperCase() + ' PROJECT</div>'
        // Big attendance %
        + '<div style="font-size:36px;font-weight:900;color:#FFF;line-height:1;font-family:Arial,sans-serif;">' + cAtt.toFixed(1) + '%</div>'
        + '<div style="font-size:10px;color:#BFDBFE;margin:3px 0 10px;font-family:Arial,sans-serif;">Current Period Attendance</div>'
        // Progress bar (table-based for email compat)
        + '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin-bottom:10px;"><tr>'
          + '<td width="' + Math.min(100, Math.max(0, cAtt)).toFixed(0) + '%" style="height:5px;background:#60A5FA;border-radius:3px 0 0 3px;font-size:0;">&nbsp;</td>'
          + '<td style="height:5px;background:rgba(255,255,255,0.15);border-radius:0 3px 3px 0;font-size:0;">&nbsp;</td>'
        + '</tr></table>'
        // Status badge + delta
        + '<table cellpadding="0" cellspacing="0"><tr>'
          + '<td style="padding-right:8px;"><span style="background:' + st.bg + ';color:' + st.fg + ';font-size:9px;font-weight:700;padding:3px 10px;border-radius:10px;font-family:Arial,sans-serif;">' + st.label + '</span></td>'
          + '<td style="font-size:10px;font-weight:700;color:' + dColor + ';font-family:Arial,sans-serif;">' + dStr + '</td>'
        + '</tr></table>'
        + '<hr style="border:none;border-top:1px solid rgba(255,255,255,0.15);margin:12px 0 10px;">'
        // Mini stats row
        + '<table width="100%" cellpadding="0" cellspacing="4"><tr>'
          + '<td style="text-align:center;">'
            + '<div style="font-size:18px;font-weight:800;color:#FFF;font-family:Arial,sans-serif;">' + (curP ? curP.active : 0) + '</div>'
            + '<div style="font-size:9px;color:#BAE6FD;font-family:Arial,sans-serif;">Active Staff</div>'
          + '</td>'
          + '<td style="text-align:center;">'
            + '<div style="font-size:18px;font-weight:800;color:#FDE68A;font-family:Arial,sans-serif;">' + (pAtt !== null ? pAtt.toFixed(1) + "%" : "—") + '</div>'
            + '<div style="font-size:9px;color:#BAE6FD;font-family:Arial,sans-serif;">Prev Period</div>'
          + '</td>'
          + '<td style="text-align:center;">'
            + '<div style="font-size:18px;font-weight:800;color:#FCA5A5;font-family:Arial,sans-serif;">' + (curP ? curP.active0Present : 0) + '</div>'
            + '<div style="font-size:9px;color:#BAE6FD;font-family:Arial,sans-serif;">Zero Present</div>'
          + '</td>'
        + '</tr></table>'
      + '</div>'
    + '</td>';
  }

  // ═══════════════════════════════════════════════════════════
  //  BUILD: Full per-project detail (KPIs + Bar Chart + tables)
  //  shown BELOW the header cards
  // ═══════════════════════════════════════════════════════════
  function projectDetail(pName, curP, prevP, themeColor) {
    const theme = PROJECT_THEME[pName] || { color: themeColor || "#334155", light: "#F8FAFC", label: pName };

    const cAtt    = curP  ? parseFloat(pct(curP.attendancePct))  : 0;
    const pAtt    = prevP ? parseFloat(pct(prevP.attendancePct)) : NaN;
    const cAbsPct = curP  ? parseFloat(pct(curP.absencePct))     : 0;
    const pAbsPct = prevP ? parseFloat(pct(prevP.absencePct))    : NaN;
    const cAbs0   = curP  ? curP.active0Present   : 0;
    const pAbs0   = prevP ? prevP.active0Present  : NaN;
    const cInact  = curP  ? curP.inactive          : 0;
    const pInact  = prevP ? prevP.inactive          : NaN;
    const cActive = curP  ? curP.active             : 0;
    const pActive = prevP ? prevP.active             : NaN;
    const cPWO    = curP  ? curP.totalPresent + curP.totalWeekOff : 0;
    const pPWO    = prevP ? prevP.totalPresent + prevP.totalWeekOff : NaN;

    // Improvements
    const attDiff  = !isNaN(pAtt)    ? cAtt    - pAtt    : null;
    const absDiff  = !isNaN(pAbsPct) ? cAbsPct - pAbsPct : null;
    const imps = [];
    if (attDiff  !== null && attDiff  >  0)  imps.push({ good: true,  msg: "Attendance improved by <b>" + attDiff.toFixed(1) + "%</b> vs previous period." });
    if (attDiff  !== null && attDiff  < -5)  imps.push({ good: false, msg: "Attendance dropped by <b>" + Math.abs(attDiff).toFixed(1) + "%</b>. Immediate review required." });
    if (!isNaN(pActive) && cActive > pActive) imps.push({ good: true,  msg: "Active headcount increased by <b>" + (cActive - pActive) + "</b>." });
    if (!isNaN(pActive) && cActive < pActive) imps.push({ good: false, msg: "Active headcount reduced by <b>" + (pActive - cActive) + "</b>." });
    if (!isNaN(pAbs0)  && cAbs0   < pAbs0)   imps.push({ good: true,  msg: "Zero-present staff reduced from <b>" + pAbs0 + "</b> to <b>" + cAbs0 + "</b>." });
    if (!isNaN(pAbs0)  && cAbs0   > pAbs0)   imps.push({ good: false, msg: "Zero-present staff increased from <b>" + pAbs0 + "</b> to <b>" + cAbs0 + "</b>." });
    if (absDiff  !== null && absDiff  < -2)   imps.push({ good: true,  msg: "Absence rate decreased by <b>" + Math.abs(absDiff).toFixed(1) + "%</b>." });
    if (absDiff  !== null && absDiff  >  2)   imps.push({ good: false, msg: "Absence rate increased by <b>" + absDiff.toFixed(1) + "%</b>." });
    if (!isNaN(pInact) && cInact < pInact)    imps.push({ good: true,  msg: "Inactive count reduced by <b>" + (pInact - cInact) + "</b>." });
    if (!isNaN(pInact) && cInact > pInact)    imps.push({ good: false, msg: "Inactive count grew by <b>" + (cInact - pInact) + "</b>. Verify attrition." });

    const impHTML = imps.length === 0
      ? '<p style="background:' + C.greenBg + ';border:1px solid #BBF7D0;border-radius:4px;padding:10px 12px;font-size:11px;color:' + C.green + ';margin:0;font-family:Arial,sans-serif;">No significant changes detected.</p>'
      : imps.map(function(im) {
          return '<div style="display:table;width:100%;border-radius:4px;margin-bottom:5px;background:'
            + (im.good ? C.greenBg : C.redBg) + ';border:1px solid ' + (im.good ? "#BBF7D0" : "#FCA5A5") + ';">'
            + '<div style="display:table-cell;width:26px;text-align:center;padding:8px 4px;vertical-align:top;">'
              + '<span style="font-size:9px;font-weight:900;color:' + (im.good ? C.green : C.red) + ';">' + (im.good ? "OK" : "!!") + '</span>'
            + '</div>'
            + '<div style="display:table-cell;font-size:11px;color:' + C.text + ';padding:8px 10px 8px 2px;vertical-align:middle;font-family:Arial,sans-serif;">' + im.msg + '</div>'
            + '</div>';
        }).join("");

    // Bar charts
    const barColorCur  = theme.color;
    const absentColor  = "#EF4444";
    const curChart     = curP  && Object.keys(curP.daywise).length  > 0 ? htmlBarChart(curP.daywise,  barColorCur, absentColor)  : '<p style="color:' + C.faint + ';font-size:11px;">No daily data.</p>';
    const prevBarColor = "#94A3B8";
    const prevChart    = prevP && Object.keys(prevP.daywise).length > 0 ? htmlBarChart(prevP.daywise, prevBarColor, "#EF4444")    : '<p style="color:' + C.faint + ';font-size:11px;">No daily data.</p>';

    // Circle table
    const allCircles = {};
    if (curP)  Object.keys(curP.circleBreakdown).forEach(function(c) { allCircles[c] = 1; });
    if (prevP) Object.keys(prevP.circleBreakdown).forEach(function(c) { allCircles[c] = 1; });
    const circles = Object.keys(allCircles).sort();

    const thHead = 'style="padding:7px 10px;font-size:9px;font-weight:700;color:#94A3B8;text-transform:uppercase;letter-spacing:.5px;text-align:center;font-family:Arial,sans-serif;"';
    const tbl = function(rows) {
      return '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid ' + C.border + ';border-radius:6px;overflow:hidden;">'
        + '<thead><tr style="background:#1E293B;">'
          + '<th ' + thHead + ' style="text-align:left;">Circle</th>'
          + '<th ' + thHead + '>Active (Cur)</th>'
          + '<th ' + thHead + '>Active (Prev)</th>'
          + '<th ' + thHead + '>Att% (Cur)</th>'
          + '<th ' + thHead + '>Att% (Prev)</th>'
          + '<th ' + thHead + '>Change</th>'
          + '<th ' + thHead + '>Zero Present</th>'
        + '</tr></thead>'
        + '<tbody>' + rows + '</tbody>'
        + '</table>';
    };

    let circleRows = "";
    circles.forEach(function(cName, ci) {
      const cc   = curP  && curP.circleBreakdown[cName]  ? curP.circleBreakdown[cName]  : null;
      const pc   = prevP && prevP.circleBreakdown[cName] ? prevP.circleBreakdown[cName] : null;
      const cA2  = cc ? pct(cc.attendancePct) : "—";
      const pA2  = pc ? pct(pc.attendancePct) : "—";
      const d2   = (cc && pc) ? parseFloat(cA2) - parseFloat(pA2) : null;
      const dC2  = d2 === null ? C.muted : (d2 > 0 ? C.green : d2 < 0 ? C.red : C.muted);
      const dT2  = d2 === null ? "—" : (d2 > 0 ? "▲ +" : d2 < 0 ? "▼ " : "→ ") + Math.abs(d2).toFixed(1) + "%";
      const zDiff = (cc && pc) ? cc.active0 - pc.active0 : null;
      const zCol  = zDiff === null ? C.text : (zDiff < 0 ? C.green : zDiff > 0 ? C.red : C.text);
      const tdS  = 'style="padding:8px 10px;font-size:11px;text-align:center;font-family:Arial,sans-serif;border-bottom:1px solid #F1F5F9;"';

      circleRows += '<tr style="background:' + (ci % 2 === 0 ? "#FFF" : "#F8FAFC") + ';">'
        + '<td style="padding:8px 12px;font-size:11px;font-weight:700;color:' + theme.color + ';font-family:Arial,sans-serif;border-bottom:1px solid #F1F5F9;">' + cName + '</td>'
        + '<td ' + tdS + '>' + (cc ? cc.active : "—") + '</td>'
        + '<td ' + tdS + ' style="padding:8px 10px;font-size:11px;text-align:center;font-family:Arial,sans-serif;border-bottom:1px solid #F1F5F9;color:' + C.muted + ';">' + (pc ? pc.active : "—") + '</td>'
        + '<td ' + tdS + ' style="padding:8px 10px;font-size:11px;text-align:center;font-family:Arial,sans-serif;border-bottom:1px solid #F1F5F9;font-weight:700;color:' + C.blue + ';">' + cA2 + (cA2 !== "—" ? "%" : "") + '</td>'
        + '<td ' + tdS + ' style="padding:8px 10px;font-size:11px;text-align:center;font-family:Arial,sans-serif;border-bottom:1px solid #F1F5F9;color:' + C.muted + ';">' + pA2 + (pA2 !== "—" ? "%" : "") + '</td>'
        + '<td ' + tdS + ' style="padding:8px 10px;font-size:11px;text-align:center;font-family:Arial,sans-serif;border-bottom:1px solid #F1F5F9;font-weight:700;color:' + dC2 + ';">' + dT2 + '</td>'
        + '<td ' + tdS + ' style="padding:8px 10px;font-size:11px;text-align:center;font-family:Arial,sans-serif;border-bottom:1px solid #F1F5F9;font-weight:700;color:' + zCol + ';">' + (cc ? cc.active0 : "—") + '</td>'
      + '</tr>';
    });

    // Role table
    const allRoles = { "TECHNICIAN": 1, "HELPER": 1, "DEO": 1, "SDA": 1 };
    if (curP)  Object.keys(curP.roleBreakdown).forEach(function(r) { allRoles[r] = 1; });
    if (prevP) Object.keys(prevP.roleBreakdown).forEach(function(r) { allRoles[r] = 1; });

    let roleRows = "";
    Object.keys(allRoles).sort().forEach(function(rName, ri) {
      const cr   = curP  && curP.roleBreakdown[rName]  ? curP.roleBreakdown[rName]  : null;
      const pr   = prevP && prevP.roleBreakdown[rName] ? prevP.roleBreakdown[rName] : null;
      const cRA  = cr ? pct(cr.attendancePct) : "—";
      const pRA  = pr ? pct(pr.attendancePct) : "—";
      const rd   = (cr && pr) ? parseFloat(cRA) - parseFloat(pRA) : null;
      const rCol = rd === null ? C.muted : (rd > 0 ? C.green : rd < 0 ? C.red : C.muted);
      const rTxt = rd === null ? "—" : (rd > 0 ? "▲ +" : rd < 0 ? "▼ " : "→ ") + Math.abs(rd).toFixed(1) + "%";
      const tdR  = 'style="padding:8px 10px;font-size:11px;text-align:center;font-family:Arial,sans-serif;border-bottom:1px solid #F1F5F9;"';

      roleRows += '<tr style="background:' + (ri % 2 === 0 ? "#FFF" : "#F8FAFC") + ';">'
        + '<td style="padding:8px 12px;font-size:11px;font-weight:700;color:' + theme.color + ';font-family:Arial,sans-serif;border-bottom:1px solid #F1F5F9;">' + rName + '</td>'
        + '<td ' + tdR + '>' + (cr ? cr.active : "—") + '</td>'
        + '<td ' + tdR + ' style="padding:8px 10px;font-size:11px;text-align:center;font-family:Arial,sans-serif;border-bottom:1px solid #F1F5F9;color:' + C.muted + ';">' + (pr ? pr.active : "—") + '</td>'
        + '<td ' + tdR + ' style="padding:8px 10px;font-size:11px;text-align:center;font-family:Arial,sans-serif;border-bottom:1px solid #F1F5F9;font-weight:700;color:' + C.blue + ';">' + cRA + (cRA !== "—" ? "%" : "") + '</td>'
        + '<td ' + tdR + ' style="padding:8px 10px;font-size:11px;text-align:center;font-family:Arial,sans-serif;border-bottom:1px solid #F1F5F9;color:' + C.muted + ';">' + pRA + (pRA !== "—" ? "%" : "") + '</td>'
        + '<td ' + tdR + ' style="padding:8px 10px;font-size:11px;text-align:center;font-family:Arial,sans-serif;border-bottom:1px solid #F1F5F9;font-weight:700;color:' + rCol + ';">' + rTxt + '</td>'
        + '<td ' + tdR + ' style="padding:8px 10px;font-size:11px;text-align:center;font-family:Arial,sans-serif;border-bottom:1px solid #F1F5F9;font-weight:700;color:' + C.red + ';">' + (cr ? cr.active0 : "—") + '</td>'
      + '</tr>';
    });

    return '<div style="margin-bottom:28px;">'
      // KPI strip
      + secHdr("KPI Overview — " + pName + " | " + curPer + " vs " + prePer, theme.color)
      + '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;"><tr>'
        + kpiPill("Active Staff",   cActive, pActive, false, false)
        + kpiPill("Inactive",       cInact,  pInact,  false, true)
        + kpiPill("Present + W/O",  cPWO,    pPWO,    false, false)
        + kpiPill("Absence Rate %", cAbsPct, pAbsPct, true,  true)
        + kpiPill("Zero Present",   cAbs0,   pAbs0,   false, true)
      + '</tr></table>'

      // Improvement summary
      + secHdr("Improvement Summary — " + pName, C.sub)
      + '<div style="padding:2px 0 6px;">' + impHTML + '</div>'

      // Bar charts — side by side (current | previous) using a 2-col table
      + secHdr("Daily Attendance Chart — " + pName, theme.color)
      + '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;"><tr>'
          + '<td style="width:50%;padding-right:8px;vertical-align:top;">'
            + '<div style="font-size:10px;font-weight:700;color:' + C.muted + ';text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px;font-family:Arial,sans-serif;">Current Period: ' + curPer + '</div>'
            + curChart
          + '</td>'
          + '<td style="width:50%;padding-left:8px;vertical-align:top;">'
            + '<div style="font-size:10px;font-weight:700;color:' + C.muted + ';text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px;font-family:Arial,sans-serif;">Comparison Period: ' + prePer + '</div>'
            + prevChart
          + '</td>'
      + '</tr></table>'

      // Circle breakdown
      + secHdr("Circle-wise Breakdown — " + pName, theme.color)
      + (circles.length > 0 ? tbl(circleRows) : '<p style="color:' + C.faint + ';font-size:11px;">No circle data.</p>')

      // Role breakdown
      + secHdr("Role-wise Breakdown — " + pName, theme.color)
      + tbl(roleRows)
    + '</div>';
  }

  // ═══════════════════════════════════════════════════════════
  //  ASSEMBLE FULL EMAIL
  // ═══════════════════════════════════════════════════════════
  // Project refs (Triveni left, Kashi right — alphabetically: Kashi < Triveni so we force order)
  const kashi   = findProject(curr, "Kashi");
  const kashiP  = findProject(prev, "Kashi");
  const triveni  = findProject(curr, "Triveni");
  const triveniP = findProject(prev, "Triveni");
  const kashiGrad   = "linear-gradient(135deg,#1E3A8A,#1D4ED8)";
  const triveniGrad = "linear-gradient(135deg,#4C1D95,#6D28D9)";

  // Side-by-side header: Triveni (left) | Kashi (right)
  const headerCards = '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin-bottom:16px;"><tr>'
    + projectCard("Triveni", triveni,  triveniP, C.triveni, triveniGrad)
    + projectCard("Kashi",   kashi,    kashiP,   C.kashi,   kashiGrad)
  + '</tr></table>';

  // Per-project detail sections
  // Order: Kashi first, then Triveni (or use allProjects order)
  let detailsHTML = "";
  // Ensure Kashi first
  const orderedProjects = [];
  if (allProjects.indexOf("Kashi") !== -1)   orderedProjects.push("Kashi");
  if (allProjects.indexOf("Triveni") !== -1)  orderedProjects.push("Triveni");
  allProjects.forEach(function(p) { if (orderedProjects.indexOf(p) === -1) orderedProjects.push(p); });

  orderedProjects.forEach(function(pName) {
    const curP  = findProject(curr, pName);
    const prevP = findProject(prev, pName);
    const theme = PROJECT_THEME[pName] || { color: "#334155" };
    if (orderedProjects.indexOf(pName) > 0) detailsHTML += divider();
    detailsHTML += projectDetail(pName, curP, prevP, theme.color);
  });

  // ── Full HTML email ─────────────────────────
  return '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1"></head>'
    + '<body style="margin:0;padding:16px 0;background:' + C.bg + ';font-family:Arial,sans-serif;color:' + C.text + ';">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="max-width:900px;margin:0 auto;">'

    // ─ Email Header ─────────────────────────────
    + '<tr><td style="background:linear-gradient(135deg,#0B1E3D 0%,#1E3A8A 55%,#4C1D95 100%);border-radius:8px 8px 0 0;padding:26px 30px 22px;">'
      + '<div style="font-size:9px;color:#93C5FD;letter-spacing:2px;text-transform:uppercase;margin-bottom:5px;font-family:Arial,sans-serif;">GMR Performance Review — Monthly Comparison</div>'
      + '<div style="font-size:22px;font-weight:900;color:#FFFFFF;margin-bottom:4px;font-family:Arial,sans-serif;line-height:1.1;">Monthly Attendance Comparison Report</div>'
      + '<div style="font-size:11px;color:#BFDBFE;margin-bottom:14px;font-family:Arial,sans-serif;">'
        + curPer + ' &nbsp;<span style="color:#FDE68A;font-weight:700;">vs</span>&nbsp; ' + prePer
        + '&nbsp;&nbsp;&middot;&nbsp;&nbsp;KPIs &middot; Improvements &middot; Daily Charts &middot; Circle &amp; Role Analysis'
      + '</div>'
      + '<table cellpadding="0" cellspacing="0"><tr>'
        + '<td style="padding-right:8px;"><span style="border:1px solid rgba(52,211,153,0.5);color:#D1FAE5;font-size:10px;font-weight:700;padding:4px 14px;border-radius:16px;font-family:Arial,sans-serif;">Current: ' + curPer + '</span></td>'
        + '<td><span style="border:1px solid rgba(253,224,71,0.4);color:#FDE68A;font-size:10px;font-weight:700;padding:4px 14px;border-radius:16px;font-family:Arial,sans-serif;">Comparison: ' + prePer + '</span></td>'
      + '</tr></table>'
    + '</td></tr>'

    // ─ Meta bar ─────────────────────────────────
    + '<tr><td style="background:#1E293B;padding:8px 30px;">'
      + '<table width="100%" cellpadding="0" cellspacing="0"><tr>'
        + '<td style="font-size:10px;color:#CBD5E1;font-family:Arial,sans-serif;"><b style="color:#94A3B8;">To:</b> ' + headName + '</td>'
        + '<td style="font-size:10px;color:#CBD5E1;text-align:center;font-family:Arial,sans-serif;"><b style="color:#94A3B8;">Reporting Window:</b> ' + curPer + '</td>'
        + '<td style="font-size:10px;color:#CBD5E1;text-align:right;font-family:Arial,sans-serif;"><b style="color:#94A3B8;">Generated:</b> ' + genOn + '</td>'
      + '</tr></table>'
    + '</td></tr>'

    // ─ Body ─────────────────────────────────────
    + '<tr><td style="background:' + C.surface + ';padding:22px 30px;border-radius:0 0 8px 8px;">'

      // Greeting
      + '<p style="margin:0 0 6px;font-size:13px;color:' + C.text + ';font-family:Arial,sans-serif;">'
        + 'Dear <b>' + headName + '</b>,'
      + '</p>'
      + '<p style="margin:0 0 6px;font-size:12px;color:' + C.muted + ';line-height:1.6;font-family:Arial,sans-serif;">'
        + 'Please find below the <b>Monthly Attendance Comparison Report</b> for the period '
        + '<b>' + curPer + '</b> compared against <b>' + prePer + '</b>. '
        + 'The report includes KPI comparisons, daily attendance charts, and project/circle/role-wise breakdowns.'
      + '</p>'
      + (fairNote ? '<p style="margin:0 0 14px;font-size:11px;color:' + C.faint + ';font-family:Arial,sans-serif;font-style:italic;">'
        + 'Note: ' + fairNote + '</p>' : '<div style="margin-bottom:14px;"></div>')

      // Side-by-side project header cards
      + headerCards

      // Per-project detail
      + detailsHTML

      // Footer note
      + '<div style="border-top:1px solid ' + C.border + ';padding-top:14px;margin-top:4px;">'
        + '<p style="margin:0;font-size:11px;color:' + C.faint + ';line-height:1.7;font-family:Arial,sans-serif;">'
          + 'Please review flagged areas and coordinate with the field team for corrective actions. '
          + 'Contact HR Operations for any data discrepancies.<br>'
          + '<i>Auto-generated by GMR Master Dashboard — Monthly Comparison Module. Do not reply directly.</i>'
        + '</p>'
      + '</div>'
    + '</td></tr>'

    // ─ Email Footer ──────────────────────────────
    + '<tr><td style="padding:10px 0;text-align:center;font-size:10px;color:#94A3B8;font-family:Arial,sans-serif;">'
      + 'GMR Manpower Tracking System&nbsp;&middot;&nbsp;' + genOn + '&nbsp;&middot;&nbsp;<b>Confidential &mdash; Internal Use Only</b>'
    + '</td></tr>'
    + '</table></body></html>';
}

// ──────────────────────────────────────────────
//  SUMMARY SHEET WRITER (unchanged logic, kept for compatibility)
// ──────────────────────────────────────────────
function writeComparisonSummarySheet() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const data = buildComparisonData(ss);
  let sheet  = ss.getSheetByName(MCD_SHEETS.REPORT);
  if (!sheet) sheet = ss.insertSheet(MCD_SHEETS.REPORT);
  else { sheet.clearContents(); sheet.clearFormats(); }

  const headers = [
    "Project","Circle","Active (Cur)","Active (Prev)","Active Delta",
    "Present+WO (Cur)","Present+WO (Prev)","Absent (Cur)","Absent (Prev)",
    "Att% (Cur)","Att% (Prev)","Att% Delta","Zero Present (Cur)","Zero Present (Prev)","Status"
  ];
  const numCols = headers.length;

  sheet.getRange(1, 1, 1, numCols)
    .setValues([["GMR MONTHLY COMPARISON | " + data.currentPeriod + " vs " + data.prevPeriod + " | " + data.generatedOn]])
    .merge()
    .setBackground("#0F172A").setFontColor("#FFFFFF").setFontWeight("bold")
    .setFontSize(12).setHorizontalAlignment("center");
  sheet.setRowHeight(1, 36);

  sheet.getRange(2, 1, 1, numCols)
    .setValues([headers])
    .setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold")
    .setFontSize(11).setHorizontalAlignment("center");
  sheet.setRowHeight(2, 28);
  sheet.setFrozenRows(2);

  let row = 3;
  const allProjects = [];
  const seen = {};
  data.current.forEach(function(p)  { if (!seen[p.name]) { seen[p.name] = 1; allProjects.push(p.name); } });
  data.previous.forEach(function(p) { if (!seen[p.name]) { seen[p.name] = 1; allProjects.push(p.name); } });

  for (let pi = 0; pi < allProjects.length; pi++) {
    const pName = allProjects[pi];
    const curP  = data.current.find(function(p)  { return p.name === pName; });
    const prevP = data.previous.find(function(p) { return p.name === pName; });
    const theme = PROJECT_THEME[pName] || { color: "#334155" };

    const allCircles = {};
    if (curP)  Object.keys(curP.circleBreakdown).forEach(function(c)  { allCircles[c] = 1; });
    if (prevP) Object.keys(prevP.circleBreakdown).forEach(function(c) { allCircles[c] = 1; });

    const circles = Object.keys(allCircles).sort();
    if (!circles.length) circles.push("(All)");

    circles.forEach(function(cName) {
      const cc = (cName === "(All)") ? null : (curP  && curP.circleBreakdown[cName]  ? curP.circleBreakdown[cName]  : null);
      const pc = (cName === "(All)") ? null : (prevP && prevP.circleBreakdown[cName] ? prevP.circleBreakdown[cName] : null);

      const cA   = curP  ? (cName === "(All)" ? curP.active  : (cc ? cc.active : 0)) : 0;
      const pA   = prevP ? (cName === "(All)" ? prevP.active : (pc ? pc.active : 0)) : 0;
      const cPWO = curP  ? (cName === "(All)" ? curP.totalPresent  + curP.totalWeekOff  : (cc ? cc.totalPresent  + cc.totalWeekOff  : 0)) : 0;
      const pPWO = prevP ? (cName === "(All)" ? prevP.totalPresent + prevP.totalWeekOff : (pc ? pc.totalPresent  + pc.totalWeekOff  : 0)) : 0;
      const cAbs = curP  ? (cName === "(All)" ? curP.totalAbsent   : (cc ? cc.totalAbsent   : 0)) : 0;
      const pAbs = prevP ? (cName === "(All)" ? prevP.totalAbsent  : (pc ? pc.totalAbsent   : 0)) : 0;
      const cAtt = curP  ? (cName === "(All)" ? (curP.attendancePct  * 100).toFixed(1) : (cc ? (cc.attendancePct  * 100).toFixed(1) : "—")) : "—";
      const pAtt = prevP ? (cName === "(All)" ? (prevP.attendancePct * 100).toFixed(1) : (pc ? (pc.attendancePct * 100).toFixed(1) : "—")) : "—";
      const attD = (cAtt !== "—" && pAtt !== "—") ? (parseFloat(cAtt) - parseFloat(pAtt)).toFixed(1) : "—";
      const cZ   = curP  ? (cName === "(All)" ? curP.active0Present  : (cc ? cc.active0 : 0)) : 0;
      const pZ   = prevP ? (cName === "(All)" ? prevP.active0Present : (pc ? pc.active0 : 0)) : 0;

      const status = cAtt === "—" ? "—"
        : parseFloat(cAtt) >= 90 ? "Healthy"
        : parseFloat(cAtt) >= 80 ? "Moderate"
        : parseFloat(cAtt) >= 70 ? "Needs Improvement"
        : "Critical";

      const rowData = [
        pName, cName === "(All)" ? "All Circles" : cName,
        cA, pA, cA - pA,
        cPWO, pPWO, cAbs, pAbs,
        cAtt + (cAtt !== "—" ? "%" : ""),
        pAtt + (pAtt !== "—" ? "%" : ""),
        attD + (attD !== "—" ? "%" : ""),
        cZ, pZ, status
      ];

      sheet.getRange(row, 1, 1, numCols).setValues([rowData]);
      sheet.setRowHeight(row, 24);
      sheet.getRange(row, 1, 1, numCols).setBackground(row % 2 === 0 ? "#F8FAFC" : "#FFFFFF").setFontSize(11);
      sheet.getRange(row, 1).setFontColor(theme.color).setFontWeight("bold");

      const attDV = attD !== "—" ? parseFloat(attD) : 0;
      if (cA - pA > 0) sheet.getRange(row, 5).setFontColor("#16A34A").setFontWeight("bold");
      if (cA - pA < 0) sheet.getRange(row, 5).setFontColor("#DC2626").setFontWeight("bold");
      if (attDV > 0) sheet.getRange(row, 12).setBackground("#DCFCE7").setFontColor("#16A34A").setFontWeight("bold");
      if (attDV < 0) sheet.getRange(row, 12).setBackground("#FEE2E2").setFontColor("#DC2626").setFontWeight("bold");

      const stColors = {
        "Critical":          ["#FEE2E2","#DC2626"],
        "Needs Improvement": ["#FFF7ED","#C2410C"],
        "Moderate":          ["#FEFCE8","#CA8A04"],
        "Healthy":           ["#DCFCE7","#166534"],
      };
      if (stColors[status]) {
        sheet.getRange(row, 15).setBackground(stColors[status][0]).setFontColor(stColors[status][1]).setFontWeight("bold");
      }
      row++;
    });
  }

  const colWidths = [90,130,90,90,70,110,110,90,90,90,90,80,120,120,120];
  colWidths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('"Monthly Comparison" summary sheet updated successfully.');
}
