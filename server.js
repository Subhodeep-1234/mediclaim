require("dotenv").config();
const path = require("path");
const express = require("express");
const { google } = require("googleapis");

const PORT = process.env.PORT || 3000;
const SPREADSHEET_ID = "1CrItGvLoD31VYSiCf4HQDoA7UUdkEdIvKzbd2s7xsuc";
const SHEET_NAME = "Additions";

const app = express();
app.use(express.json());
app.use(express.static(__dirname));

if (!process.env.GOOGLE_SERVICE_ACCOUNT_KEY_B64) {
  console.error(
    "Missing GOOGLE_SERVICE_ACCOUNT_KEY_B64 env var. Set it in the deploy " +
    "platform's environment/secrets config -- .env is gitignored and never " +
    "reaches the container."
  );
  process.exit(1);
}

const serviceAccountKey = JSON.parse(
  Buffer.from(process.env.GOOGLE_SERVICE_ACCOUNT_KEY_B64, "base64").toString("utf-8")
);
const auth = new google.auth.GoogleAuth({
  credentials: serviceAccountKey,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});
const sheetsPromise = auth.getClient().then((authClient) => google.sheets({ version: "v4", auth: authClient }));

// Serialize writes so two near-simultaneous submissions can't read the same
// "last Sl. No." before either has appended, and collide.
let writeQueue = Promise.resolve();

app.post("/api/submit", (req, res) => {
  const result = writeQueue.then(() => handleSubmit(req.body));
  writeQueue = result.catch(() => {});
  result
    .then(() => res.json({ ok: true }))
    .catch((err) => {
      console.error("submit failed:", err);
      res.status(500).json({ ok: false, error: String((err && err.message) || err) });
    });
});

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
      ""
    ];
  });

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:L`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: rows }
  });
}

app.listen(PORT, "0.0.0.0", () => console.log(`Mediclaim server listening on 0.0.0.0:${PORT}`));
