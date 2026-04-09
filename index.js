'use strict';

const express = require('express');
const { google } = require('googleapis');
const nodemailer = require('nodemailer');

const app = express();
app.use(express.json());

// ── Config from environment variables ────────────────────────────────────────
const SPREADSHEET_ID = process.env.SPREADSHEET_ID || '1h7h20__qofEqjUv6Eg2XB46TZ78NRZaish3nzOJbYd8';
const SHEET_NAME     = 'Registration Automation';
const PORT           = process.env.PORT || 8080;

// SMTP / email config
const SMTP_HOST = process.env.SMTP_HOST;
const SMTP_PORT = parseInt(process.env.SMTP_PORT || '587', 10);
const SMTP_USER = process.env.SMTP_USER;
const SMTP_PASS = process.env.SMTP_PASS;
const EMAIL_FROM = process.env.EMAIL_FROM || SMTP_USER;
const EMAIL_TO   = process.env.EMAIL_TO;   // comma-separated for multiple recipients

// Slack config — comma-separated for multiple webhooks
const SLACK_WEBHOOK_URLS = (process.env.SLACK_WEBHOOK_URL || '')
  .split(',')
  .map(u => u.trim())
  .filter(Boolean);

// ── Google Sheets auth (uses ADC / service account on Cloud Run) ─────────────
async function getSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
  const authClient = await auth.getClient();
  return google.sheets({ version: 'v4', auth: authClient });
}

// ── Fetch raw values from the Registration Automation sheet ──────────────────
async function fetchSheetData(sheets) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!A:H`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  return res.data.values || [];
}

// ── Parse the sheet into structured sections ─────────────────────────────────
//
// Sheet layout (row groups after the header row):
//   Col A-B  → Marquee Events  (A=name/label, B=Paid Reg value)
//   Col D    → Webinars        (D=name/label)
//   Col F    → Agentic Bootcamps (F=name/label)
//   Col H    → Workshops       (H=name/label)
//
// Each event occupies 3 rows:  [Event Name] → [Labels row] → [Values row]
// Row 0 is the category header row.
//
function parseSheetData(rows) {
  // Helpers
  const cell = (row, col) => (row && row[col] != null ? String(row[col]).trim() : '');
  const num  = (row, col) => {
    const v = cell(row, col);
    return v === '' ? null : v;
  };

  const sections = {
    marqueeEvents:   [],  // { name, totalReg, paidReg }
    webinars:        [],  // { name, totalReg }
    agenticBootcamps:[],  // { name, totalReg }
    workshops:       [],  // { name, totalReg }
  };

  // Start at row index 1 (skip header row), step through groups of 3
  for (let i = 1; i + 2 < rows.length; i += 3) {
    const nameRow  = rows[i];
    const labelRow = rows[i + 1];
    const valRow   = rows[i + 2];

    // Marquee Events (cols A=0, B=1)
    const marqueeName = cell(nameRow, 0);
    if (marqueeName && marqueeName !== 'Total Reg' && marqueeName !== 'Paid Reg') {
      sections.marqueeEvents.push({
        name:     marqueeName,
        totalReg: num(valRow, 0),
        paidReg:  num(valRow, 1),
      });
    }

    // Webinars (col D=3)
    const webinarName = cell(nameRow, 3);
    if (webinarName && webinarName !== 'Total Reg') {
      sections.webinars.push({
        name:     webinarName,
        totalReg: num(valRow, 3),
      });
    }

    // Agentic Bootcamps (col F=5)
    const bootcampName = cell(nameRow, 5);
    if (bootcampName && bootcampName !== 'Total Reg') {
      sections.agenticBootcamps.push({
        name:     bootcampName,
        totalReg: num(valRow, 5),
      });
    }

    // Workshops (col H=7)
    const workshopName = cell(nameRow, 7);
    if (workshopName && workshopName !== 'Total Reg') {
      sections.workshops.push({
        name:     workshopName,
        totalReg: num(valRow, 7),
      });
    }
  }

  // Handle any trailing Agentic Bootcamp rows that start at odd offsets
  // (the sheet may have bootcamp-only rows after the marquee section ends)
  return sections;
}

// ── Build HTML email body (Slack-email compatible) ────────────────────────────
// Slack renders emails in a ~540px modal. Rules applied:
//   - max-width 540px, no fixed widths
//   - flat single-level tables (no nested wrapper tables)
//   - larger font & padding so text isn't cramped on the narrow panel
//   - value column is right-aligned and fixed-width so numbers line up
function buildEmailHtml(sections) {
  const today = new Date().toLocaleDateString('en-US', {
    year: 'numeric', month: 'long', day: 'numeric',
  });

  const val = (v) => (v == null || v === '' ? '—' : v);

  // Shared styles
  const th = (width = null) =>
    `background:#1a3a5c;color:#fff;padding:8px 12px;text-align:left;font-size:13px;border:1px solid #1a3a5c;` +
    (width ? `width:${width};` : '');
  const thR = (width = null) =>
    `background:#1a3a5c;color:#fff;padding:8px 12px;text-align:right;font-size:13px;border:1px solid #1a3a5c;` +
    (width ? `width:${width};` : '');
  const td  = (alt) => `padding:7px 12px;border:1px solid #dde3ec;font-size:13px;color:#222;${alt ? 'background:#f4f7fb;' : ''}`;
  const tdR = (alt) => `padding:7px 12px;border:1px solid #dde3ec;font-size:13px;color:#222;text-align:right;${alt ? 'background:#f4f7fb;' : ''}`;

  const sectionHead = (title) =>
    `<p style="margin:20px 0 6px;font-size:14px;font-weight:bold;color:#1a3a5c;` +
    `padding-bottom:4px;border-bottom:2px solid #1a3a5c;">${title}</p>`;

  const table = (head, body) =>
    `<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;margin-bottom:4px;">` +
    `<thead>${head}</thead><tbody>${body}</tbody></table>`;

  // ── Marquee Events (3 cols: Event | Total Reg | Paid Reg) ──────────────────
  const marqueeBody = sections.marqueeEvents.map((e, i) =>
    `<tr>
      <td style="${td(i%2)}">${e.name}</td>
      <td style="${tdR(i%2)}">${val(e.totalReg)}</td>
      <td style="${tdR(i%2)}">${val(e.paidReg)}</td>
    </tr>`).join('') || `<tr><td colspan="3" style="${td(false)}">No data</td></tr>`;

  const marqueeTable = sectionHead('Marquee Events') + table(
    `<tr><th style="${th()}">Event</th><th style="${thR('110px')}">Total Reg</th><th style="${thR('110px')}">Paid Reg</th></tr>`,
    marqueeBody
  );

  // ── Webinars (2 cols: Webinar | Total Reg) — skip TBD rows ────────────────
  const webinarItems = sections.webinars.filter(e => e.name !== 'TBD');
  const webinarBody = webinarItems.map((e, i) =>
    `<tr>
      <td style="${td(i%2)}">${e.name}</td>
      <td style="${tdR(i%2)}">${val(e.totalReg)}</td>
    </tr>`).join('') || `<tr><td colspan="2" style="${td(false)}">No data</td></tr>`;

  const webinarTable = sectionHead('Webinars') + table(
    `<tr><th style="${th()}">Webinar</th><th style="${thR('110px')}">Total Reg</th></tr>`,
    webinarBody
  );

  // ── Agentic Bootcamps (2 cols: Session | Total Reg) ───────────────────────
  const bootcampBody = sections.agenticBootcamps.map((e, i) =>
    `<tr>
      <td style="${td(i%2)}">${e.name}</td>
      <td style="${tdR(i%2)}">${val(e.totalReg)}</td>
    </tr>`).join('') || `<tr><td colspan="2" style="${td(false)}">No data</td></tr>`;

  const bootcampTable = sectionHead('Agentic Bootcamps') + table(
    `<tr><th style="${th()}">Session</th><th style="${thR('110px')}">Total Reg</th></tr>`,
    bootcampBody
  );

  // ── Workshops (2 cols: Workshop | Total Reg) — skip TBD rows ─────────────
  const workshopItems = sections.workshops.filter(e => e.name !== 'TBD');
  const workshopBody = workshopItems.map((e, i) =>
    `<tr>
      <td style="${td(i%2)}">${e.name}</td>
      <td style="${tdR(i%2)}">${val(e.totalReg)}</td>
    </tr>`).join('') || `<tr><td colspan="2" style="${td(false)}">No data</td></tr>`;

  const workshopTable = sectionHead('Workshops') + table(
    `<tr><th style="${th()}">Workshop</th><th style="${thR('110px')}">Total Reg</th></tr>`,
    workshopBody
  );

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
</head>
<body style="margin:0;padding:0;background:#f0f2f5;font-family:Arial,sans-serif;">
  <div style="max-width:540px;margin:24px auto;background:#fff;border-radius:6px;overflow:hidden;border:1px solid #dde3ec;">

    <!-- Header -->
    <div style="background:#1a3a5c;padding:20px 24px;">
      <div style="font-size:17px;font-weight:bold;color:#fff;line-height:1.3;">
        Event Registration Report
      </div>
      <div style="font-size:12px;color:#a8c4e0;margin-top:4px;">${today}</div>
    </div>

    <!-- Content -->
    <div style="padding:20px 24px 24px;">
      ${marqueeTable}
      ${webinarTable}
      ${bootcampTable}
      ${workshopTable}
    </div>

    <!-- Footer -->
    <div style="background:#f4f7fb;padding:12px 24px;border-top:1px solid #dde3ec;">
      <span style="font-size:11px;color:#999;">
        Pulled from the
        <a href="https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}" style="color:#1a3a5c;">IAB Tech Lab Events spreadsheet</a>
      </span>
    </div>

  </div>
</body>
</html>`;
}

// ── Send email via nodemailer ─────────────────────────────────────────────────
async function sendEmail(html) {
  if (!SMTP_USER || !SMTP_PASS || !EMAIL_TO) {
    throw new Error('Missing required env vars: SMTP_USER, SMTP_PASS, EMAIL_TO');
  }

  const transporter = nodemailer.createTransport({
    host: SMTP_HOST || 'smtp.gmail.com',
    port: SMTP_PORT,
    secure: SMTP_PORT === 465,
    auth: { user: SMTP_USER, pass: SMTP_PASS },
  });

  const today = new Date().toLocaleDateString('en-US', {
    year: 'numeric', month: 'long', day: 'numeric',
  });

  const info = await transporter.sendMail({
    from: `"IAB Tech Lab Reports" <${EMAIL_FROM}>`,
    to: EMAIL_TO,
    subject: `Registration Automation Report — ${today}`,
    html,
  });

  return info.messageId;
}

// ── Build Slack Block Kit message ─────────────────────────────────────────────
function buildSlackBlocks(sections) {
  const today = new Date().toLocaleDateString('en-US', {
    year: 'numeric', month: 'long', day: 'numeric',
  });

  const val = (v) => (v == null || v === '' ? '—' : v);

  const tableRows = (items, cols) =>
    items.length === 0
      ? '_No data_'
      : items.map(e => cols.map(c => `*${c.label}:* ${val(e[c.key])}`).join('   ')).join('\n');

  const marqueeText = tableRows(
    sections.marqueeEvents,
    [{ label: 'Event', key: 'name' }, { label: 'Total Reg', key: 'totalReg' }, { label: 'Paid Reg', key: 'paidReg' }]
  );

  const webinarText = tableRows(
    sections.webinars.filter(e => e.name !== 'TBD'),
    [{ label: 'Webinar', key: 'name' }, { label: 'Total Reg', key: 'totalReg' }]
  );

  const bootcampText = tableRows(
    sections.agenticBootcamps,
    [{ label: 'Session', key: 'name' }, { label: 'Total Reg', key: 'totalReg' }]
  );

  const workshopText = tableRows(
    sections.workshops.filter(e => e.name !== 'TBD'),
    [{ label: 'Workshop', key: 'name' }, { label: 'Total Reg', key: 'totalReg' }]
  );

  return {
    blocks: [
      {
        type: 'header',
        text: { type: 'plain_text', text: 'Event Registration Report', emoji: false },
      },
      {
        type: 'context',
        elements: [{ type: 'mrkdwn', text: today }],
      },
      { type: 'divider' },
      { type: 'section', text: { type: 'mrkdwn', text: '*Marquee Events*' } },
      { type: 'section', text: { type: 'mrkdwn', text: marqueeText } },
      { type: 'divider' },
      { type: 'section', text: { type: 'mrkdwn', text: '*Webinars*' } },
      { type: 'section', text: { type: 'mrkdwn', text: webinarText } },
      { type: 'divider' },
      { type: 'section', text: { type: 'mrkdwn', text: '*Agentic Bootcamps*' } },
      { type: 'section', text: { type: 'mrkdwn', text: bootcampText } },
      { type: 'divider' },
      { type: 'section', text: { type: 'mrkdwn', text: '*Workshops*' } },
      { type: 'section', text: { type: 'mrkdwn', text: workshopText } },
    ],
  };
}

// ── Build Slack message with monospace table layout ───────────────────────────
function buildSlackTable(sections) {
  const today = new Date().toLocaleDateString('en-US', {
    year: 'numeric', month: 'long', day: 'numeric',
  });

  const val = (v) => (v == null || v === '' ? '—' : String(v));
  const padL = (s, w) => String(s).padEnd(w);
  const padR = (s, w) => String(s).padStart(w);

  // cols: [{ label, key, width, align }]
  function makeTable(items, cols) {
    if (items.length === 0) return '  (no data)';
    const top    = '┌' + cols.map(c => '─'.repeat(c.width + 2)).join('┬') + '┐';
    const header = '│ ' + cols.map(c => padL(c.label, c.width)).join(' │ ') + ' │';
    const divider= '├' + cols.map(c => '─'.repeat(c.width + 2)).join('┼') + '┤';
    const bottom = '└' + cols.map(c => '─'.repeat(c.width + 2)).join('┴') + '┘';
    const rows = items.map(e =>
      '│ ' + cols.map(c =>
        c.align === 'right'
          ? padR(val(e[c.key]), c.width)
          : padL(val(e[c.key]), c.width)
      ).join(' │ ') + ' │'
    );
    return [top, header, divider, ...rows, bottom].join('\n');
  }

  const marqueeTable = makeTable(sections.marqueeEvents, [
    { label: 'Event',     key: 'name',     width: 32, align: 'left'  },
    { label: 'Total Reg', key: 'totalReg', width: 9,  align: 'right' },
    { label: 'Paid Reg',  key: 'paidReg',  width: 8,  align: 'right' },
  ]);

  const webinarTable = makeTable(
    sections.webinars.filter(e => e.name !== 'TBD'),
    [
      { label: 'Webinar',   key: 'name',     width: 36, align: 'left'  },
      { label: 'Total Reg', key: 'totalReg', width: 9,  align: 'right' },
    ]
  );

  const bootcampTable = makeTable(sections.agenticBootcamps, [
    { label: 'Session',   key: 'name',     width: 36, align: 'left'  },
    { label: 'Total Reg', key: 'totalReg', width: 9,  align: 'right' },
  ]);

  const workshopTable = makeTable(
    sections.workshops.filter(e => e.name !== 'TBD'),
    [
      { label: 'Workshop',  key: 'name',     width: 36, align: 'left'  },
      { label: 'Total Reg', key: 'totalReg', width: 9,  align: 'right' },
    ]
  );

  const section = (title, table) => ([
    { type: 'section', text: { type: 'mrkdwn', text: `*${title}*` } },
    { type: 'section', text: { type: 'mrkdwn', text: `\`\`\`${table}\`\`\`` } },
  ]);

  return {
    blocks: [
      {
        type: 'header',
        text: { type: 'plain_text', text: 'Event Registration Report', emoji: false },
      },
      { type: 'context', elements: [{ type: 'mrkdwn', text: today }] },
      { type: 'divider' },
      ...section('Marquee Events',    marqueeTable),
      { type: 'divider' },
      ...section('Webinars',          webinarTable),
      { type: 'divider' },
      ...section('Agentic Bootcamps', bootcampTable),
      { type: 'divider' },
      ...section('Workshops',         workshopTable),
    ],
  };
}

// ── Send Slack message via Incoming Webhook ───────────────────────────────────
async function sendSlack(payload) {
  if (SLACK_WEBHOOK_URLS.length === 0) throw new Error('Missing required env var: SLACK_WEBHOOK_URL');

  await Promise.all(SLACK_WEBHOOK_URLS.map(async (url) => {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Slack webhook failed for ${url}: ${res.status} ${text}`);
    }
  }));
}

// ── HTTP handler ──────────────────────────────────────────────────────────────
app.post('/', async (req, res) => {
  try {
    console.log('Fetching Google Sheet data...');
    const sheets  = await getSheetsClient();
    const rows    = await fetchSheetData(sheets);

    console.log(`Fetched ${rows.length} rows from "${SHEET_NAME}"`);
    const sections = parseSheetData(rows);

    // ── Slack notification ──────────────────────────────────────────────────
    const slackPayload = buildSlackBlocks(sections);
    console.log('Sending Slack message...');
    await sendSlack(slackPayload);
    console.log('Slack message sent.');

    // ── Email notification (disabled) ──────────────────────────────────────
    // const html = buildEmailHtml(sections);
    // console.log('Sending email...');
    // const messageId = await sendEmail(html);
    // console.log(`Email sent. messageId=${messageId}`);

    res.status(200).json({ status: 'ok' });
  } catch (err) {
    console.error('Error:', err);
    res.status(500).json({ status: 'error', message: err.message });
  }
});

// Health check
app.get('/', (req, res) => {
  res.status(200).json({ status: 'ok', service: 'sheet-automation' });
});

app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
