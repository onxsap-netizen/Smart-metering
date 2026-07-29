/**
 * ============================================================================
 *  APRAAVA WORKFORCE MANDATE & ATTENDANCE REPORT
 *  Google Apps Script — reads the workbook, builds the styled HTML report,
 *  and emails it (manually or on a schedule).
 * ============================================================================
 *
 *  SETUP
 *  1. Open your Google Sheet (the one with these tabs):
 *       - "Team deatils"
 *       - "Required VS Achieved"
 *       - "Client Attendance Gujarat July"
 *       - "Client Attendance MP July"
 *       - "Client Attendance Rajasthan  Ju"   <-- note the sheet name has
 *         two spaces before "Ju" in the original file; CONFIG.SHEETS below
 *         lets you fix the exact name to match your sheet.
 *  2. Extensions -> Apps Script. Delete any boilerplate code, paste this
 *     whole file in.
 *  3. Edit the CONFIG block below: recipient emails, sheet names (if
 *     different), and the reporting period label.
 *  4. Run `sendApraavaReport` once from the editor (Run button) to test.
 *     The first run will ask you to authorize Gmail + Sheets access.
 *  5. To automate: run `createDailyTrigger` once (or use the Triggers
 *     clock icon in the left sidebar) to schedule it — see bottom of file
 *     for schedule options (daily / weekly / specific weekday).
 *
 * ============================================================================
 */

// ============================= CONFIG ======================================
const CONFIG = {
  // Who receives the report. Comma-separate multiple addresses.
  RECIPIENTS: 'workforce.ops@apraava.com, manager@apraava.com',
  CC: '', // optional, leave blank if none
  SUBJECT_PREFIX: 'Apraava — Workforce Mandate & Attendance Report',

  // Reporting period label shown in the report header (edit each cycle,
  // or leave the auto-generated one — see getPeriodLabel())
  PERIOD_LABEL: '15 Jun – 14 Jul 2026',

  // Exact tab names in your spreadsheet
  SHEETS: {
    TEAM_DETAILS: 'Team deatils',
    REQUIRED_VS_ACHIEVED: 'Required VS Achieved',
    ATTENDANCE_GUJARAT: 'Client Attendance Gujarat July',
    ATTENDANCE_MP: 'Client Attendance MP July',
    ATTENDANCE_RAJASTHAN: 'Client Attendance Rajasthan  Ju', // keep exact spacing
  },

  // Attendance status tokens used in the sheet (case sensitive)
  STATUS_TOKENS: {
    PRESENT: 'P',
    ABSENT: 'A',
    WEEK_OFF: ['Wo', 'WO'], // Gujarat/MP use "Wo", Rajasthan uses "WO"
  },
};
// ============================================================================


/**
 * Entry point — call this to generate + send the report.
 * Also callable from a time-driven trigger.
 */
function sendApraavaReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const stateData = buildStateWiseData_(ss);
  const deploymentData = buildDeploymentPipelineData_(ss);
  const dailySeries = buildDailyAttendanceSeries_(ss);
  const chartBlob = buildAttendanceChart_(dailySeries);

  const html = renderReportHtml_(stateData, deploymentData, dailySeries, chartBlob);

  const subject = `${CONFIG.SUBJECT_PREFIX} — ${getPeriodLabel_()}`;

  GmailApp.sendEmail(CONFIG.RECIPIENTS, subject, 'This report requires an HTML-capable email client.', {
    htmlBody: html,
    cc: CONFIG.CC || undefined,
    inlineImages: { attendanceChart: chartBlob },
    name: 'Apraava Workforce Analytics',
  });

  Logger.log('Report sent to: ' + CONFIG.RECIPIENTS);
}


// ============================= DATA LAYER ==================================

/**
 * Reads Team Details (for sanctioned mandate) + all 3 attendance sheets
 * (for achieved headcount, present/absent mandays), and returns a
 * state-wise + location-wise structure.
 */
function buildStateWiseData_(ss) {
  const teamSheet = ss.getSheetByName(CONFIG.SHEETS.TEAM_DETAILS);

  // ---- Sanctioned mandate from the "Summary" block (State name + count) ----
  // Mirrors the layout: column L = 'State' labels, column M = counts,
  // starting a couple of rows below the "Summary" header.
  const sanctioned = {}; // { 'Gujarat': 139, 'Rajasthan': 14, 'MP': 22 }
  const teamValues = teamSheet.getDataRange().getValues();
  for (let r = 0; r < teamValues.length; r++) {
    const label = teamValues[r][11]; // column L (0-indexed 11)
    const count = teamValues[r][12]; // column M (0-indexed 12)
    if (typeof label === 'string' && typeof count === 'number') {
      const key = normalizeStateName_(label);
      if (key) sanctioned[key] = count;
    }
  }

  // ---- Achieved headcount + present/absent mandays, per state & location ---
  const locationSummary = {}; // key: "State|Location" -> {headcount, present, absent}

  ingestAttendanceSheet_(ss, CONFIG.SHEETS.ATTENDANCE_GUJARAT, 'Gujarat', locationSummary, {
    headerRow: 2, cityCol: 'City', statusCol: 'EMP Status', presentCol: 'Present', absentCol: 'Absent',
    weekOffTokens: ['Wo'],
  });
  ingestAttendanceSheet_(ss, CONFIG.SHEETS.ATTENDANCE_MP, 'MP', locationSummary, {
    headerRow: 2, cityCol: 'City', statusCol: 'EMP Status', presentCol: 'Present', absentCol: 'Absent',
    weekOffTokens: ['Wo'],
  });
  ingestAttendanceSheet_(ss, CONFIG.SHEETS.ATTENDANCE_RAJASTHAN, 'Rajasthan', locationSummary, {
    headerRow: 1, cityCol: 'circle', statusCol: 'Status', presentCol: 'Present', absentCol: 'Absent',
    weekOffTokens: ['WO'],
  });

  // ---- Roll up to state level ----
  const stateAgg = {}; // 'Gujarat' -> {headcount, present, absent}
  Object.keys(locationSummary).forEach(key => {
    const [state] = key.split('|');
    if (!stateAgg[state]) stateAgg[state] = { headcount: 0, present: 0, absent: 0 };
    stateAgg[state].headcount += locationSummary[key].headcount;
    stateAgg[state].present += locationSummary[key].present;
    stateAgg[state].absent += locationSummary[key].absent;
  });

  const states = ['Gujarat', 'Rajasthan', 'MP'].map(state => {
    const agg = stateAgg[state] || { headcount: 0, present: 0, absent: 0 };
    const required = sanctioned[state] || 0;
    const total = agg.present + agg.absent;
    const attendancePct = total ? round1_(100 * agg.present / total) : 0;
    return {
      state: state === 'MP' ? 'Madhya Pradesh' : state,
      required: required,
      achieved: agg.headcount,
      deficit: required - agg.headcount,
      present: agg.present,
      absent: agg.absent,
      attendancePct: attendancePct,
    };
  });

  const totals = states.reduce((acc, s) => {
    acc.required += s.required;
    acc.achieved += s.achieved;
    acc.present += s.present;
    acc.absent += s.absent;
    return acc;
  }, { required: 0, achieved: 0, present: 0, absent: 0 });
  totals.deficit = totals.required - totals.achieved;
  totals.fulfillmentPct = totals.required ? round1_(100 * totals.achieved / totals.required) : 0;
  totals.attendancePct = (totals.present + totals.absent)
    ? round1_(100 * totals.present / (totals.present + totals.absent)) : 0;

  // ---- Location-level table, sorted by state then headcount desc ----
  const locations = Object.keys(locationSummary).map(key => {
    const [state, loc] = key.split('|');
    const v = locationSummary[key];
    const total = v.present + v.absent;
    return {
      state: state === 'MP' ? 'Madhya Pradesh' : state,
      location: loc,
      headcount: v.headcount,
      present: v.present,
      absent: v.absent,
      attendancePct: total ? round1_(100 * v.present / total) : 0,
    };
  }).sort((a, b) => (a.state > b.state ? 1 : a.state < b.state ? -1 : b.headcount - a.headcount));

  return { states, totals, locations };
}


/**
 * Reads the "Required VS Achieved" sheet — this tracks NEW deployment
 * (fresh hiring mandate), separate from the sanctioned strength above.
 */
function buildDeploymentPipelineData_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.REQUIRED_VS_ACHIEVED);
  const values = sheet.getDataRange().getValues();
  const header = values[0].map(h => String(h).trim());
  const locIdx = header.indexOf('Location');
  const reqIdx = header.indexOf('Requirement');
  const fulIdx = header.findIndex(h => h.toLowerCase().indexOf('fillment') > -1 || h.toLowerCase().indexOf('fulfillment') > -1);
  const defIdx = header.indexOf('Deficit');

  const rows = [];
  let totalReq = 0, totalFul = 0, totalDef = 0;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const loc = row[locIdx];
    const req = Number(row[reqIdx]) || 0;
    const ful = Number(row[fulIdx]) || 0;
    const def = row[defIdx] !== '' && row[defIdx] != null ? Number(row[defIdx]) : (req - ful);

    if (!loc && req && !rows.length) continue; // skip stray blanks
    if (loc) {
      rows.push({ location: loc, required: req, deployed: ful, pending: def });
    } else if (req && !loc) {
      // this is the totals row (no location, only numbers)
      totalReq = req; totalFul = ful; totalDef = def;
    }
  }

  if (!totalReq) {
    totalReq = rows.reduce((a, r) => a + r.required, 0);
    totalFul = rows.reduce((a, r) => a + r.deployed, 0);
    totalDef = rows.reduce((a, r) => a + r.pending, 0);
  }

  rows.sort((a, b) => b.pending - a.pending);

  return {
    rows: rows,
    totalRequired: totalReq,
    totalDeployed: totalFul,
    totalPending: totalDef,
    fulfilledPct: totalReq ? round1_(100 * totalFul / totalReq) : 0,
  };
}


/**
 * Builds a day-by-day attendance % series across all three states,
 * excluding week-off days, for the trend chart.
 */
function buildDailyAttendanceSeries_(ss) {
  const dailyMap = {}; // 'yyyy-MM-dd' -> {present, absent}

  accumulateDailySeries_(ss, CONFIG.SHEETS.ATTENDANCE_GUJARAT, dailyMap, {
    headerRow: 2, dateHeaderRow: 1, weekOffTokens: ['Wo'],
  });
  accumulateDailySeries_(ss, CONFIG.SHEETS.ATTENDANCE_MP, dailyMap, {
    headerRow: 2, dateHeaderRow: 1, weekOffTokens: ['Wo'],
  });
  accumulateDailySeries_(ss, CONFIG.SHEETS.ATTENDANCE_RAJASTHAN, dailyMap, {
    headerRow: 1, dateHeaderRow: 1, weekOffTokens: ['WO'],
  });

  const dates = Object.keys(dailyMap).sort();
  const series = dates
    .map(d => {
      const v = dailyMap[d];
      const total = v.present + v.absent;
      return {
        date: d,
        label: Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'dd-MMM'),
        present: v.present,
        absent: v.absent,
        pct: total ? round1_(100 * v.present / total) : null,
      };
    })
    .filter(d => d.present !== 0); // drop week-off days (all-zero present)

  return series;
}


// ============================= SHEET HELPERS ===============================

function ingestAttendanceSheet_(ss, sheetName, stateLabel, locationSummary, opts) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);
  const values = sheet.getDataRange().getValues();
  const header = values[opts.headerRow - 1];

  const cityIdx = header.indexOf(opts.cityCol);
  const statusIdx = header.indexOf(opts.statusCol);
  const presentIdx = header.indexOf(opts.presentCol);
  const absentIdx = header.indexOf(opts.absentCol);

  for (let r = opts.headerRow; r < values.length; r++) {
    const row = values[r];
    if (!row[cityIdx]) continue;
    const status = String(row[statusIdx] || '').trim();
    if (status !== 'Active') continue;

    const loc = String(row[cityIdx]).trim();
    const key = stateLabel + '|' + normalizeLocationName_(loc);
    if (!locationSummary[key]) locationSummary[key] = { headcount: 0, present: 0, absent: 0 };

    locationSummary[key].headcount += 1;
    locationSummary[key].present += Number(row[presentIdx]) || 0;
    locationSummary[key].absent += Number(row[absentIdx]) || 0;
  }
}

function accumulateDailySeries_(ss, sheetName, dailyMap, opts) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);
  const values = sheet.getDataRange().getValues();
  const dateHeader = values[opts.dateHeaderRow - 1];
  const dataStartRow = opts.headerRow; // 0-indexed row where data begins

  // find columns that contain real Date objects
  const dateCols = [];
  dateHeader.forEach((val, idx) => {
    if (val instanceof Date) dateCols.push({ idx: idx, date: val });
  });

  for (let r = dataStartRow; r < values.length; r++) {
    const row = values[r];
    if (!row[0] && !row[1] && !row[2]) continue; // skip blank rows

    dateCols.forEach(dc => {
      const val = String(row[dc.idx] || '').trim();
      const key = Utilities.formatDate(dc.date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (!dailyMap[key]) dailyMap[key] = { present: 0, absent: 0 };

      if (val === CONFIG.STATUS_TOKENS.PRESENT) {
        dailyMap[key].present += 1;
      } else if (val === CONFIG.STATUS_TOKENS.ABSENT) {
        dailyMap[key].absent += 1;
      }
      // week-offs intentionally not counted in present/absent
    });
  }
}

function normalizeStateName_(label) {
  const clean = label.trim();
  if (/gujarat/i.test(clean)) return 'Gujarat';
  if (/raj/i.test(clean)) return 'Rajasthan';
  if (/^mp$/i.test(clean) || /madhya/i.test(clean)) return 'MP';
  return null;
}

function normalizeLocationName_(loc) {
  // Title-case for consistent display (sheet has mixed casing e.g. "RAJKOT", "junagadh")
  return loc.replace(/\w\S*/g, w => w.charAt(0).toUpperCase() + w.substr(1).toLowerCase());
}

function round1_(n) {
  return Math.round(n * 10) / 10;
}

function getPeriodLabel_() {
  return CONFIG.PERIOD_LABEL;
}


// ============================= CHART LAYER ==================================

/**
 * Builds the daily attendance % line chart as an inline image blob using
 * the built-in Apps Script Charts service (no external libraries needed).
 */
function buildAttendanceChart_(series) {
  const dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Date')
    .addColumn(Charts.ColumnType.NUMBER, 'Attendance %');

  series.forEach(pt => dataTable.addRow([pt.label, pt.pct]));

  const avg = round1_(series.reduce((a, p) => a + p.pct, 0) / series.length);

  const chart = Charts.newLineChart()
    .setDataTable(dataTable)
    .setDimensions(900, 380)
    .setColors(['#0d6e6e'])
    .setTitle('Daily Attendance % — Working Days')
    .setBackgroundFill('#ffffff')
    .setPointStyle(Charts.PointStyle.MEDIUM)
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { textStyle: { fontSize: 9, color: '#5b6577' }, slantedText: true, slantedTextAngle: 45 })
    .setOption('vAxis', { textStyle: { fontSize: 10, color: '#5b6577' }, viewWindow: { min: 60, max: 100 }, gridlines: { color: '#e9edf0' } })
    .setOption('backgroundColor', '#ffffff')
    .setOption('chartArea', { left: 60, top: 40, width: '85%', height: '65%' })
    .build();

  return chart.getAs('image/png').setName('attendanceChart');
}


// ============================= HTML RENDER ==================================

/**
 * Renders the full HTML email body. Mirrors the Apraava dashboard-style
 * design: hero header, KPI strip, state-wise mandate cards, new-deployment
 * pipeline table, location drill-down, chart, insights, and action list.
 */
function renderReportHtml_(stateData, deploymentData, dailySeries, chartBlob) {
  const t = stateData.totals;

  const stateCardsHtml = stateData.states.map(s => {
    const flagClass = s.deficit <= 0 ? 'flag-green' : 'flag-red';
    const attClass = s.attendancePct >= 90 ? 'att-green' : (s.attendancePct >= 80 ? 'att-amber' : 'att-red');
    const barWidth = s.required ? Math.min(100, Math.round(100 * s.achieved / s.required)) : 100;
    const barClass = s.deficit > 0 ? 'progress-fill short' : 'progress-fill';
    const chipClass = s.deficit > 0 ? 'chip-warn' : 'chip-ok';
    const chipText = s.deficit > 0 ? `Deficit of ${s.deficit}` : 'Fully Met';
    return `
      <div class="state-card">
        <div class="state-head">
          <div class="state-name"><span class="state-flag ${flagClass}"></span>${escapeHtml_(s.state)}</div>
          <div class="state-att ${attClass}">${s.attendancePct}% attendance</div>
        </div>
        <div class="progress-track"><div class="${barClass}" style="width:${barWidth}%;"></div></div>
        <div class="mandate-foot">
          <span>Required: <b>${s.required}</b> &nbsp;&middot;&nbsp; Achieved: <b>${s.achieved}</b></span>
          <span class="deficit-chip ${chipClass}">${chipText}</span>
        </div>
      </div>`;
  }).join('');

  const locationRowsHtml = stateData.locations.map((l, i) => {
    const badgeClass = l.attendancePct >= 90 ? 'badge-green' : (l.attendancePct >= 80 ? 'badge-amber' : 'badge-red');
    return `<tr><td>${escapeHtml_(l.state)}</td><td>${escapeHtml_(l.location)}</td><td>${l.headcount}</td><td>${l.present.toLocaleString()}</td><td>${l.absent.toLocaleString()}</td><td><span class="badge ${badgeClass}">${l.attendancePct}%</span></td></tr>`;
  }).join('');

  const deployRowsHtml = deploymentData.rows.map(r => {
    const pct = r.required ? round1_(100 * r.deployed / r.required) : 0;
    const badgeClass = pct >= 100 ? 'badge-green' : (pct > 0 ? 'badge-amber' : 'badge-red');
    return `<tr><td>${escapeHtml_(r.location)}</td><td>${r.required}</td><td>${r.deployed}</td><td>${r.pending}</td><td><span class="badge ${badgeClass}">${pct}%</span></td></tr>`;
  }).join('');

  const avgAttendance = round1_(dailySeries.reduce((a, p) => a + p.pct, 0) / dailySeries.length);
  const firstHalf = dailySeries.slice(0, Math.floor(dailySeries.length / 2));
  const secondHalf = dailySeries.slice(Math.floor(dailySeries.length / 2));
  const firstAvg = round1_(firstHalf.reduce((a, p) => a + p.pct, 0) / firstHalf.length);
  const secondAvg = round1_(secondHalf.reduce((a, p) => a + p.pct, 0) / secondHalf.length);

  // Identify weakest location for auto-generated action items
  const worstLoc = [...stateData.locations].sort((a, b) => a.attendancePct - b.attendancePct)[0];
  const worstDeploy = [...deploymentData.rows].sort((a, b) => b.pending - a.pending)[0];
  const worstState = [...stateData.states].sort((a, b) => a.attendancePct - b.attendancePct)[0];

  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
  :root{--ink:#0b1f22;--teal-900:#083b3d;--teal-700:#0d6e6e;--teal-500:#12907f;--teal-100:#e3f3ef;
    --amber:#e8a33d;--amber-100:#fdf1de;--red:#c94a3c;--red-100:#fbeae7;--paper:#f6f8f7;--card:#ffffff;
    --line:#e6ece9;--muted:#5b6f6d;}
  *{box-sizing:border-box;}
  body{margin:0;background:var(--paper);font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,Arial,sans-serif;color:var(--ink);}
  .wrap{max-width:760px;margin:0 auto;padding:28px 16px 60px;}
  .card{background:var(--card);border-radius:16px;overflow:hidden;box-shadow:0 1px 3px rgba(8,59,61,0.06),0 12px 32px rgba(8,59,61,0.06);border:1px solid var(--line);}
  .hero{background:linear-gradient(135deg,#083b3d 0%,#0d6e6e 55%,#128f7a 100%);padding:34px 32px 28px;position:relative;}
  .eyebrow{color:#a9e0d4;font-size:11px;letter-spacing:1.6px;text-transform:uppercase;font-weight:600;position:relative;z-index:1;}
  .brandrow{display:flex;align-items:center;justify-content:space-between;margin-top:14px;position:relative;z-index:1;}
  .brand{display:flex;align-items:center;gap:10px;}
  .logo-mark{width:34px;height:34px;border-radius:9px;background:linear-gradient(135deg,var(--amber),#f3c274);
    display:flex;align-items:center;justify-content:center;font-weight:800;color:#4a2e05;font-size:15px;box-shadow:0 2px 6px rgba(0,0,0,0.15);}
  .brand-name{color:#fff;font-size:15px;font-weight:700;letter-spacing:0.2px;}
  .brand-sub{color:#bfe6dd;font-size:10.5px;margin-top:1px;}
  .period-pill{background:rgba(255,255,255,0.12);border:1px solid rgba(255,255,255,0.2);color:#eafaf5;font-size:11px;padding:6px 12px;border-radius:20px;}
  .h1{color:#fff;font-size:25px;font-weight:800;margin-top:20px;position:relative;z-index:1;letter-spacing:-0.3px;}
  .h1-sub{color:#cdece3;font-size:13.5px;margin-top:6px;position:relative;z-index:1;line-height:1.5;max-width:480px;}
  .tag-row{margin-top:18px;display:flex;gap:8px;flex-wrap:wrap;position:relative;z-index:1;}
  .tag{font-size:11.5px;padding:6px 13px;border-radius:20px;font-weight:600;}
  .tag-gold{background:var(--amber);color:#4a2e05;}
  .tag-ghost{background:rgba(255,255,255,0.14);color:#eafaf5;border:1px solid rgba(255,255,255,0.22);}
  .section{padding:26px 32px 4px;}
  .section-label{font-size:12.5px;font-weight:700;color:var(--teal-900);text-transform:uppercase;letter-spacing:0.6px;display:flex;align-items:center;gap:8px;margin-bottom:14px;}
  .section-label .dot{width:6px;height:6px;border-radius:50%;background:var(--amber);}
  .section-note{font-size:12px;color:var(--muted);margin:-6px 0 16px;line-height:1.6;}
  .stat-strip{display:flex;gap:12px;padding:0 32px;margin-top:22px;}
  .stat-box{flex:1;background:var(--teal-100);border-radius:12px;padding:16px 14px;text-align:center;border:1px solid #d8ece6;}
  .stat-num{font-size:23px;font-weight:800;color:var(--teal-900);letter-spacing:-0.5px;}
  .stat-label{font-size:10px;color:var(--muted);margin-top:5px;text-transform:uppercase;letter-spacing:0.4px;font-weight:600;}
  .stat-box.gap .stat-num{color:var(--red);}
  .stat-box.gap{background:var(--red-100);border-color:#f3d9d4;}
  .stat-box.pct .stat-num{color:#9a6b12;}
  .stat-box.pct{background:var(--amber-100);border-color:#f2ddb8;}
  .state-card{border:1px solid var(--line);border-radius:14px;padding:18px 18px 16px;margin-bottom:14px;background:linear-gradient(180deg,#ffffff,#fbfdfc);}
  .state-head{display:flex;justify-content:space-between;align-items:flex-start;}
  .state-name{font-size:15px;font-weight:800;color:var(--ink);display:flex;align-items:center;gap:8px;}
  .state-flag{width:8px;height:8px;border-radius:50%;}
  .flag-green{background:#1f9d55;} .flag-red{background:var(--red);}
  .state-att{font-size:13px;font-weight:700;}
  .att-green{color:#1f9d55;} .att-amber{color:#b8790f;} .att-red{color:var(--red);}
  .progress-track{height:9px;background:#eef2f0;border-radius:6px;margin-top:8px;overflow:hidden;position:relative;}
  .progress-fill{height:100%;border-radius:6px;background:linear-gradient(90deg,var(--teal-700),var(--teal-500));}
  .progress-fill.short{background:linear-gradient(90deg,#c9822f,var(--amber));}
  .mandate-foot{display:flex;justify-content:space-between;margin-top:8px;font-size:11px;color:var(--muted);}
  .deficit-chip{font-size:10.5px;font-weight:700;padding:2px 9px;border-radius:12px;}
  .chip-ok{background:#e6f7ec;color:#1f9d55;} .chip-warn{background:var(--red-100);color:var(--red);}
  table.loc{width:100%;border-collapse:collapse;font-size:11.8px;}
  table.loc thead th{background:var(--teal-900);color:#eafaf5;text-align:left;padding:10px 10px;font-weight:600;font-size:10.8px;letter-spacing:0.3px;text-transform:uppercase;}
  table.loc thead th:first-child{border-radius:8px 0 0 0;} table.loc thead th:last-child{border-radius:0 8px 0 0;}
  table.loc tbody td{padding:8px 10px;border-bottom:1px solid #f0f3f1;color:#374050;}
  table.loc tbody tr:nth-child(even){background:#fafcfb;}
  .badge{font-size:10.5px;font-weight:700;padding:2px 8px;border-radius:10px;display:inline-block;}
  .badge-green{background:#e6f7ec;color:#1f9d55;} .badge-amber{background:var(--amber-100);color:#9a6b12;} .badge-red{background:var(--red-100);color:var(--red);}
  .chart-card{border:1px solid var(--line);border-radius:14px;padding:18px;background:#fff;}
  .chart-card img{width:100%;display:block;border-radius:8px;}
  .chart-legend{display:flex;gap:18px;margin-top:12px;font-size:11px;color:var(--muted);}
  .legend-dot{width:9px;height:9px;border-radius:50%;display:inline-block;margin-right:5px;vertical-align:middle;}
  .insight-grid{display:flex;gap:10px;flex-wrap:wrap;margin-top:4px;}
  .insight{flex:1;min-width:210px;border-radius:12px;padding:14px 15px;font-size:12px;line-height:1.55;}
  .insight-title{font-weight:700;font-size:11.5px;margin-bottom:5px;display:flex;align-items:center;gap:6px;}
  .insight-up{background:#e6f7ec;color:#136c3a;} .insight-flat{background:#eef4f7;color:#2c5470;} .insight-down{background:var(--red-100);color:#9c2f24;}
  .action-box{background:var(--amber-100);border:1px solid #f0dcb4;border-radius:12px;padding:16px 18px;}
  .action-title{font-size:12.5px;font-weight:800;color:#8a5a00;margin-bottom:10px;display:flex;align-items:center;gap:7px;}
  .action-item{display:flex;gap:10px;font-size:12px;color:#4a4032;line-height:1.6;margin-bottom:9px;}
  .action-num{flex-shrink:0;width:20px;height:20px;border-radius:50%;background:var(--amber);color:#4a2e05;font-size:10.5px;font-weight:800;display:flex;align-items:center;justify-content:center;margin-top:1px;}
  .footer{padding:22px 32px 30px;border-top:1px solid var(--line);margin-top:8px;}
  .footer-text{font-size:11px;color:#93a19e;line-height:1.7;}
  .footer-brand{font-size:11px;color:var(--teal-700);font-weight:700;margin-top:8px;}
  @media (max-width:480px){.wrap{padding:14px 8px 40px;}.hero{padding:24px 20px 22px;}.section{padding:20px 20px 4px;}.stat-strip{padding:0 20px;flex-wrap:wrap;}.stat-box{min-width:44%;}.footer{padding:20px 20px 26px;}}
</style></head>
<body>
<div class="wrap"><div class="card">

  <div class="hero">
    <div class="eyebrow">Apraava Energy &nbsp;&bull;&nbsp; Field Workforce Program</div>
    <div class="brandrow">
      <div class="brand"><div class="logo-mark">A</div>
        <div><div class="brand-name">Apraava</div><div class="brand-sub">Technician Deployment Report</div></div>
      </div>
      <div class="period-pill">${escapeHtml_(getPeriodLabel_())}</div>
    </div>
    <div class="h1">Workforce Mandate &amp; Attendance</div>
    <div class="h1-sub">Consolidated state-wise view of sanctioned mandate vs. achieved headcount, new deployment pipeline, and field attendance performance across Gujarat, Rajasthan &amp; Madhya Pradesh.</div>
    <div class="tag-row">
      <span class="tag tag-gold">All States</span>
      <span class="tag tag-ghost">${t.achieved} Active Technicians</span>
      <span class="tag tag-ghost">Generated ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy')}</span>
    </div>
  </div>

  <div class="stat-strip">
    <div class="stat-box"><div class="stat-num">${t.required}</div><div class="stat-label">Required Mandate</div></div>
    <div class="stat-box"><div class="stat-num">${t.achieved}</div><div class="stat-label">Achieved Headcount</div></div>
    <div class="stat-box gap"><div class="stat-num">${t.deficit}</div><div class="stat-label">Net Deficit</div></div>
    <div class="stat-box pct"><div class="stat-num">${t.fulfillmentPct}%</div><div class="stat-label">Fulfillment</div></div>
  </div>

  <div class="section" style="margin-top:8px;">
    <div class="section-label"><span class="dot"></span>State-wise Required vs. Achieved</div>
    <div class="section-note">Sanctioned mandate compared against current active headcount, with period attendance rate for each state.</div>
    ${stateCardsHtml}
  </div>

  <div class="section">
    <div class="section-label"><span class="dot"></span>New Deployment — Pipeline</div>
    <div class="section-note">Fresh recruitment mandate for new sites, separate from the existing sanctioned strength above.</div>
    <div class="stat-strip" style="padding:0;margin-top:0;margin-bottom:16px;">
      <div class="stat-box"><div class="stat-num">${deploymentData.totalRequired}</div><div class="stat-label">New Requirement</div></div>
      <div class="stat-box"><div class="stat-num">${deploymentData.totalDeployed}</div><div class="stat-label">Deployed So Far</div></div>
      <div class="stat-box gap"><div class="stat-num">${deploymentData.totalPending}</div><div class="stat-label">Yet to Deploy</div></div>
      <div class="stat-box pct" style="background:#fbeae7;border-color:#f3d9d4;"><div class="stat-num" style="color:#c94a3c;">${deploymentData.fulfilledPct}%</div><div class="stat-label">Fulfilled</div></div>
    </div>
    <table class="loc">
      <thead><tr><th>Location</th><th>Requirement</th><th>Deployed</th><th>Pending</th><th>Progress</th></tr></thead>
      <tbody>${deployRowsHtml}</tbody>
    </table>
  </div>

  <div class="section">
    <div class="section-label"><span class="dot"></span>Sub-division (Location) wise Detail</div>
    <div class="section-note">Present / absent mandays and attendance rate for every operating location, grouped by state.</div>
    <table class="loc">
      <thead><tr><th>State</th><th>Location</th><th>HC</th><th>Present</th><th>Absent</th><th>Attendance</th></tr></thead>
      <tbody>${locationRowsHtml}</tbody>
    </table>
  </div>

  <div class="section">
    <div class="section-label"><span class="dot"></span>Daily Attendance Trend</div>
    <div class="section-note">Working-day attendance % across all three states, week-offs excluded.</div>
    <div class="chart-card">
      <img src="cid:attendanceChart" alt="Daily attendance trend chart">
      <div class="chart-legend">
        <span><span class="legend-dot" style="background:#0d6e6e;"></span>Daily attendance %</span>
        <span><span class="legend-dot" style="background:#e8a33d;"></span>Period average (${avgAttendance}%)</span>
      </div>
    </div>
  </div>

  <div class="section">
    <div class="section-label"><span class="dot"></span>Key Insights</div>
    <div class="insight-grid">
      <div class="insight insight-up">
        <div class="insight-title">&check; Auto-generated summary</div>
        Overall fulfillment is ${t.fulfillmentPct}% against sanctioned mandate; average field attendance stands at ${avgAttendance}% for the period.
      </div>
      <div class="insight insight-down">
        <div class="insight-title">&#9888; Weakest location</div>
        ${escapeHtml_(worstLoc.location)} (${escapeHtml_(worstLoc.state)}) has the lowest attendance at ${worstLoc.attendancePct}%.
      </div>
      <div class="insight insight-flat">
        <div class="insight-title">&rarr; Trend</div>
        Attendance moved from ${firstAvg}% (first half) to ${secondAvg}% (second half) of the period.
      </div>
      <div class="insight insight-down">
        <div class="insight-title">&#9888; Deployment pipeline</div>
        Only ${deploymentData.totalDeployed} of ${deploymentData.totalRequired} new positions deployed (${deploymentData.fulfilledPct}%) — ${deploymentData.totalPending} roles still open, largest gap at ${escapeHtml_(worstDeploy.location)} (${worstDeploy.pending}).
      </div>
    </div>
  </div>

  <div class="section" style="padding-bottom:26px;">
    <div class="section-label"><span class="dot"></span>Action Required</div>
    <div class="action-box">
      <div class="action-title">&#9889; Priority follow-ups this week</div>
      <div class="action-item"><span class="action-num">1</span><span><b>New deployment pipeline</b> &mdash; ${deploymentData.totalPending} of ${deploymentData.totalRequired} new positions still pending, led by ${escapeHtml_(worstDeploy.location)} (${worstDeploy.pending} open). Needs an accelerated hiring push.</span></div>
      <div class="action-item"><span class="action-num">2</span><span><b>${escapeHtml_(worstState.state)}</b> &mdash; attendance at ${worstState.attendancePct}%, the lowest of all states. Escalate to the local supervisor for a root-cause review.</span></div>
      <div class="action-item"><span class="action-num">3</span><span><b>${escapeHtml_(worstLoc.location)}</b> &mdash; lowest-performing location at ${worstLoc.attendancePct}% attendance; monitor closely.</span></div>
    </div>
  </div>

  <div class="footer">
    <div class="footer-text">Generated automatically from the Apraava workforce &amp; attendance spreadsheet for the reporting period. Figures reflect active technicians only. For queries or corrections, contact the Apraava Workforce Operations desk.</div>
    <div class="footer-brand">Apraava Energy &mdash; Field Workforce Analytics</div>
  </div>

</div></div>
</body></html>`;
}

function escapeHtml_(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}


// ============================= TRIGGERS ====================================

/**
 * Run ONCE to schedule the report to auto-send every Monday at 8 AM.
 * Change ".onWeekDay(...)" / ".atHour(...)" to fit your cadence, e.g.
 * daily: ScriptApp.newTrigger('sendApraavaReport').timeBased().everyDays(1).atHour(8).create();
 */
function createWeeklyTrigger() {
  deleteExistingTriggers_(); // avoid duplicate triggers on re-run
  ScriptApp.newTrigger('sendApraavaReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  Logger.log('Weekly trigger created: every Monday at 8 AM.');
}

/** Alternative: run once for a DAILY 8 AM trigger instead of weekly. */
function createDailyTrigger() {
  deleteExistingTriggers_();
  ScriptApp.newTrigger('sendApraavaReport')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  Logger.log('Daily trigger created: every day at 8 AM.');
}

function deleteExistingTriggers_() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendApraavaReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}
