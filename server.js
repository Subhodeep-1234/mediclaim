require("dotenv").config();
const path = require("path");
const express = require("express");
const { google } = require("googleapis");
const PDFDocument = require("pdfkit");

const PORT = process.env.PORT || 3000;
const SPREADSHEET_ID = "1CrItGvLoD31VYSiCf4HQDoA7UUdkEdIvKzbd2s7xsuc";
const SHEET_NAME = "Additions";

const app = express();
app.use(express.json());
// index: false -- otherwise this auto-serves index.html for GET "/" before
// our own route below ever runs, bypassing the ?submit= query-string check.
app.use(express.static(__dirname, { index: false }));

// Wide-open CORS: only relevant once the frontend posts to a different
// origin than the one it's served from (e.g. a direct platform URL that
// isn't behind the GET-only reverse proxy at employees.alcoverealty.in).
// Same-origin requests ignore these headers entirely.
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.header("Access-Control-Allow-Headers", "Content-Type, X-Mediclaim-Submit");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

app.get("/health", (req, res) => res.status(200).send("ok"));

app.get("/", (req, res) => {
  // The reverse proxy in front of this app only forwards GET requests to
  // "/", so a form submission rides along as a GET request, but the
  // payload goes in a header (base64-encoded) rather than the query
  // string/URL, so it doesn't end up in browser history or access logs.
  const submitHeader = req.get("X-Mediclaim-Submit");
  if (submitHeader) {
    return handleSubmitRequest(submitHeader, res);
  }

  res.sendFile(path.join(__dirname, "index.html"), (err) => {
    if (err) {
      console.error("Could not serve index.html from", __dirname, "-", err.message);
      res.status(500).send("index.html not found next to server.js -- check the deploy build output.");
    }
  });
});

if (!process.env.GOOGLE_SERVICE_ACCOUNT_KEY) {
  console.error(
    "Missing GOOGLE_SERVICE_ACCOUNT_KEY env var. Set it in the deploy " +
    "platform's environment/secrets config -- .env is gitignored and never " +
    "reaches the container."
  );
  process.exit(1);
}

const serviceAccountKey = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY);
const auth = new google.auth.GoogleAuth({
  credentials: serviceAccountKey,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});
const sheetsPromise = auth.getClient().then((authClient) => google.sheets({ version: "v4", auth: authClient }));

// Serialize writes so two near-simultaneous submissions can't read the same
// "last Sl. No." before either has appended, and collide.
let writeQueue = Promise.resolve();

// Kept in case a direct (non-proxied) URL for this app ever turns up --
// POSTed to "/" rather than a separate path, since the reverse proxy in
// front of employees.alcoverealty.in only forwards the root path.
app.post("/", (req, res) => {
  submitAndRespond(req.body, res);
});

function handleSubmitRequest(base64Json, res) {
  let payload;
  try {
    payload = JSON.parse(Buffer.from(base64Json, "base64").toString("utf-8"));
  } catch (err) {
    return res.status(400).json({ ok: false, error: "Invalid payload" });
  }
  submitAndRespond(payload, res);
}

function submitAndRespond(payload, res) {
  const result = writeQueue.then(() => handleSubmit(payload));
  writeQueue = result.catch(() => {});
  result
    .then(() => res.json({ ok: true }))
    .catch((err) => {
      console.error("submit failed:", err);
      res.status(500).json({ ok: false, error: String((err && err.message) || err) });
    });
}

function nowISTString() {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Asia/Kolkata",
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false
  }).formatToParts(new Date());
  const get = (type) => parts.find((p) => p.type === type).value;
  return `${get("day")}-${get("month")}-${get("year")} ${get("hour")}:${get("minute")}:${get("second")}`;
}

async function handleSubmit(payload) {
  if (
    !payload ||
    !payload.company ||
    !payload.empId ||
    !payload.additionDate ||
    !Array.isArray(payload.members) ||
    !payload.members.length
  ) {
    throw new Error("Incomplete payload");
  }

  const sheets = await sheetsPromise;

  const existing = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A2:A`,
  });

  let nextSlNo = 0;
  (existing.data.values || []).forEach((row) => {
    const n = parseInt(row[0], 10);
    if (!isNaN(n) && n > nextSlNo) nextSlNo = n;
  });

  const submittedAt = nowISTString();

  const rows = payload.members.map((m) => {
    nextSlNo++;
    return [
      nextSlNo,
      payload.company,
      "",
      payload.empId,
      m.fullName,
      payload.additionDate,
      m.dob,
      m.gender,
      m.relationship,
      "",
      "",
      "",
      submittedAt
    ];
  });

  const appendResult = await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:M`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: rows }
  });

  await clearBoldOnRange(sheets, appendResult.data.updates.updatedRange);

  // The sheet write above is the critical action -- a failure here (PDF
  // generation or the email provider being unreachable) is logged but must
  // not fail the submission the employee already sees confirmed.
  try {
    const pdfBuffer = await generateAdditionPdf(payload);
    await sendAdditionEmail(payload, pdfBuffer);
  } catch (err) {
    console.error("Addition PDF/email failed:", err);
  }
}

const PDF_COLUMNS = [
  { header: "Corporate_name", width: 135 },
  { header: "Full Name", width: 84 },
  { header: "DOJ/DOM", width: 55 },
  { header: "DOB", width: 58 },
  { header: "Gender", width: 40 },
  { header: "Relationship", width: 75 },
  { header: "Sum Insured", width: 68 }
];

function generateAdditionPdf(payload) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ margin: 40, size: "A4" });
    const chunks = [];
    doc.on("data", (chunk) => chunks.push(chunk));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);

    doc.fontSize(14).font("Helvetica-Bold").text("Group Mediclaim - Addition Request", { align: "center" });
    doc.moveDown(0.5);
    doc.fontSize(10).font("Helvetica").text(`Employee ID: ${payload.empId}`);
    doc.text(`Addition Date: ${payload.additionDate}`);
    doc.moveDown();

    const startX = doc.page.margins.left;
    let y = doc.y;
    const rowHeight = 22;

    function drawRow(values, isHeader) {
      let x = startX;
      doc.font(isHeader ? "Helvetica-Bold" : "Helvetica").fontSize(9);
      if (isHeader) doc.rect(startX, y, PDF_COLUMNS.reduce((s, c) => s + c.width, 0), rowHeight).fill("#eef2f8");
      doc.fillColor("#16324e");
      PDF_COLUMNS.forEach((col, i) => {
        doc.rect(x, y, col.width, rowHeight).stroke("#dde4ee");
        doc.text(String(values[i] || ""), x + 4, y + 6, { width: col.width - 8, height: rowHeight - 8, ellipsis: true });
        x += col.width;
      });
      y += rowHeight;
    }

    drawRow(PDF_COLUMNS.map((c) => c.header), true);
    payload.members.forEach((m) => {
      drawRow([payload.company, m.fullName, payload.additionDate, m.dob, m.gender, m.relationship, ""], false);
    });

    doc.end();
  });
}

async function sendAdditionEmail(payload, pdfBuffer) {
  if (!process.env.RESEND_API_KEY) {
    console.error("Missing RESEND_API_KEY env var -- skipping addition email.");
    return;
  }

  // Resend's shared sandbox sender works without verifying a domain, but
  // set RESEND_FROM to a verified @alcoverealty.in address once one exists
  // -- it reads better to recipients and is far less likely to be marked spam.
  const fromAddress = process.env.RESEND_FROM || "onboarding@resend.dev";

  const res = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${process.env.RESEND_API_KEY}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      from: fromAddress,
      to: ["manager.hr@alcoverealty.in"],
      cc: ["hr@alcoverealty.in"],
      subject: "Request for Addition of Member(s) under Group Mediclaim Policy",
      text:
        "Dear Sir/Madam,\n" +
        "We would like to request the addition of the following member(s) under our Group Mediclaim Policy.\n" +
        "Kindly confirm the addition and share the e-card(s) at the earliest.",
      attachments: [
        {
          filename: `Mediclaim-Addition-${payload.empId}.pdf`,
          content: pdfBuffer.toString("base64")
        }
      ]
    })
  });

  if (!res.ok) {
    throw new Error(`Resend API error ${res.status}: ${await res.text()}`);
  }
}

let additionsSheetIdPromise = null;
function getAdditionsSheetId(sheets) {
  if (!additionsSheetIdPromise) {
    additionsSheetIdPromise = sheets.spreadsheets
      .get({ spreadsheetId: SPREADSHEET_ID })
      .then((res) => res.data.sheets.find((s) => s.properties.title === SHEET_NAME).properties.sheetId);
  }
  return additionsSheetIdPromise;
}

// New rows appended via INSERT_ROWS pick up the formatting of the row above
// them, which was bold here -- explicitly reset just the written range back
// to normal weight rather than touching the rest of the sheet.
async function clearBoldOnRange(sheets, updatedRange) {
  const match = updatedRange.match(/![A-Z]+(\d+):[A-Z]+(\d+)/);
  if (!match) return;
  const startRowIndex = parseInt(match[1], 10) - 1;
  const endRowIndex = parseInt(match[2], 10);
  const sheetId = await getAdditionsSheetId(sheets);

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          repeatCell: {
            range: { sheetId, startRowIndex, endRowIndex, startColumnIndex: 0, endColumnIndex: 13 },
            cell: { userEnteredFormat: { textFormat: { bold: false } } },
            fields: "userEnteredFormat.textFormat.bold"
          }
        }
      ]
    }
  });
}

app.listen(PORT, "0.0.0.0", () => console.log(`Mediclaim server listening on 0.0.0.0:${PORT}`));
