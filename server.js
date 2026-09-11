require("dotenv").config();
const path = require("path");
const express = require("express");
const { google } = require("googleapis");

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

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:M`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: rows }
  });
}

app.listen(PORT, "0.0.0.0", () => console.log(`Mediclaim server listening on 0.0.0.0:${PORT}`));
