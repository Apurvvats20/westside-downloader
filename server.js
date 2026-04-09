const express = require("express");
const path = require("path");
const fs = require("fs");
const { spawn } = require("child_process");
const multer = require("multer");
const { getAuthUrl, handleCallback, isConnected, uploadSmallFile, downloadSmallFile, listFilesInFolder, getUploadsFolderId } = require("./gdrive");

// ── Bootstrap config.json from template + env vars if missing ────────────────
const CONFIG_PATH = path.join(__dirname, "config.json");
if (!fs.existsSync(CONFIG_PATH)) {
  const template = JSON.parse(fs.readFileSync(path.join(__dirname, "config.template.json"), "utf8"));
  if (process.env.GOOGLE_CLIENT_ID)     template.google.client_id     = process.env.GOOGLE_CLIENT_ID;
  if (process.env.GOOGLE_CLIENT_SECRET) template.google.client_secret = process.env.GOOGLE_CLIENT_SECRET;
  if (process.env.GOOGLE_REDIRECT_URI)  template.google.redirect_uri  = process.env.GOOGLE_REDIRECT_URI;
  if (process.env.SP_COOKIE)            template.auth.cookie           = process.env.SP_COOKIE;
  fs.mkdirSync(path.join(__dirname, "uploads"), { recursive: true });
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(template, null, 2), "utf8");
  console.log("✅ config.json created from template");
}

const app = express();
app.use(express.json({ limit: "10mb" }));
app.use(express.static(path.join(__dirname, "public")));

const upload = multer({ dest: path.join(__dirname, "uploads") });
let runningProcess = null;

// ── Config (internal use only) ───────────────────────────────────────────────
function readConfig() {
  return JSON.parse(fs.readFileSync(path.join(__dirname, "config.json"), "utf8"));
}
function writeConfig(cfg) {
  fs.writeFileSync(path.join(__dirname, "config.json"), JSON.stringify(cfg, null, 2), "utf8");
}

// ── Upload CSV/Excel → save locally + upload to Drive ────────────────────────
app.post("/api/upload", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file uploaded" });

  const ext = path.extname(req.file.originalname).toLowerCase();
  if (![".csv", ".xlsx", ".xls"].includes(ext)) {
    fs.unlinkSync(req.file.path);
    return res.status(400).json({ error: "Only CSV or Excel files are supported" });
  }

  const destPath = path.join(__dirname, "uploads", req.file.originalname);
  fs.renameSync(req.file.path, destPath);

  // Update config with new file path
  const cfg = readConfig();
  cfg.excel.file = destPath;

  // Upload to Drive if connected
  let driveFileId = null;
  if (isConnected()) {
    try {
      const folderId = await getUploadsFolderId();
      driveFileId = await uploadSmallFile(destPath, req.file.originalname, folderId);
      // Store mapping filename → driveFileId in config
      if (!cfg.uploadedFiles) cfg.uploadedFiles = {};
      cfg.uploadedFiles[req.file.originalname] = { driveFileId, uploadedAt: new Date().toISOString() };
    } catch (e) {
      console.error("Drive upload failed:", e.message);
    }
  }
  writeConfig(cfg);

  res.json({ success: true, filename: req.file.originalname, driveFileId });
});

// ── List previously uploaded CSVs ─────────────────────────────────────────────
app.get("/api/files", async (req, res) => {
  if (!isConnected()) return res.json([]);
  try {
    const folderId = await getUploadsFolderId();
    const files = await listFilesInFolder(folderId);
    // Filter only CSV/Excel files (not report JSONs)
    const csvFiles = files.filter(f => /\.(csv|xlsx|xls)$/i.test(f.name));
    res.json(csvFiles);
  } catch (e) {
    res.json([]);
  }
});

// ── Select a previously uploaded file ────────────────────────────────────────
app.post("/api/select-file", async (req, res) => {
  const { fileId, filename } = req.body;
  if (!fileId || !filename) return res.status(400).json({ error: "Missing fileId or filename" });
  try {
    // Download file from Drive to local uploads folder
    const destPath = path.join(__dirname, "uploads", filename);
    const content = await downloadSmallFile(fileId);
    fs.writeFileSync(destPath, typeof content === "string" ? content : JSON.stringify(content), "utf8");

    const cfg = readConfig();
    cfg.excel.file = destPath;
    writeConfig(cfg);

    // Load report — try Drive fileId key first, then filename key
    const report = cfg.reports?.[`report_${fileId}`]
      || cfg.reports?.[`report_${filename}`]
      || (fs.existsSync(REPORT_PATH) ? JSON.parse(fs.readFileSync(REPORT_PATH, "utf8")) : null);

    res.json({ success: true, filename, report });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── Update cookie (admin only) ───────────────────────────────────────────────
app.post("/api/update-cookie", (req, res) => {
  try {
    const { cookie } = req.body;
    const cfg = readConfig();
    cfg.auth.cookie = cookie;
    writeConfig(cfg);
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── Google Drive Auth ─────────────────────────────────────────────────────────
app.get("/auth/google", (req, res) => res.redirect(getAuthUrl()));

app.get("/auth/google/callback", async (req, res) => {
  try {
    await handleCallback(req.query.code);
    res.send(`
      <html><body style="font-family:'Google Sans',Roboto,sans-serif;text-align:center;padding:60px;background:#f8f9fa">
        <div style="background:white;border-radius:12px;padding:40px;display:inline-block;border:1px solid #e0e0e0">
          <div style="font-size:48px">✅</div>
          <h2 style="color:#1a73e8;margin:16px 0 8px">Google Drive Connected!</h2>
          <p style="color:#5f6368">You can close this tab and go back to the app.</p>
          <script>setTimeout(() => window.close(), 2000)</script>
        </div>
      </body></html>
    `);
  } catch (e) {
    res.status(500).send(`Auth failed: ${e.message}`);
  }
});

app.get("/api/gdrive-status", (req, res) => res.json({ connected: isConnected() }));

app.post("/api/gdrive-disconnect", (req, res) => {
  const cfg = readConfig();
  cfg.google.tokens = null;
  writeConfig(cfg);
  res.json({ success: true });
});

// ── Report ───────────────────────────────────────────────────────────────────
const REPORT_PATH = path.join(__dirname, "uploads", "report.json");
const DRIVE_REPORT_NAME = "westside_report.json";
let _reportFolderId = null;

async function getReportFromDrive() {
  if (!isConnected()) return null;
  try {
    const folderId = await getUploadsFolderId();
    _reportFolderId = folderId;
    const files = await listFilesInFolder(folderId);
    const reportFile = files.find(f => f.name === DRIVE_REPORT_NAME);
    if (!reportFile) return null;
    const content = await downloadSmallFile(reportFile.id);
    return typeof content === "string" ? JSON.parse(content) : content;
  } catch { return null; }
}

async function saveReportToDrive(report) {
  if (!isConnected()) return;
  try {
    const folderId = _reportFolderId || await getUploadsFolderId();
    _reportFolderId = folderId;
    fs.mkdirSync(path.join(__dirname, "uploads"), { recursive: true });
    fs.writeFileSync(REPORT_PATH, JSON.stringify(report, null, 2), "utf8");
    await uploadSmallFile(REPORT_PATH, DRIVE_REPORT_NAME, folderId);
  } catch (e) { console.error("Failed to save report to Drive:", e.message); }
}

app.get("/api/report", async (req, res) => {
  // Try local file first (fast)
  if (fs.existsSync(REPORT_PATH)) {
    try { return res.json(JSON.parse(fs.readFileSync(REPORT_PATH, "utf8"))); } catch {}
  }
  // Fallback: fetch from Drive (survives redeploys)
  const report = await getReportFromDrive();
  if (report) {
    // Cache locally for subsequent requests
    fs.mkdirSync(path.join(__dirname, "uploads"), { recursive: true });
    fs.writeFileSync(REPORT_PATH, JSON.stringify(report, null, 2), "utf8");
    return res.json(report);
  }
  res.json(null);
});

// ── Seed report (one-time use to set initial state) ───────────────────────────
app.post("/api/seed-report", async (req, res) => {
  try {
    const report = req.body;
    const cfg = readConfig();
    const reportKey = `report_${report.file}`;
    if (!cfg.reports) cfg.reports = {};
    cfg.reports[reportKey] = report;
    cfg.currentReportKey = reportKey;
    fs.mkdirSync(path.join(__dirname, "uploads"), { recursive: true });
    fs.writeFileSync(REPORT_PATH, JSON.stringify(report, null, 2), "utf8");
    writeConfig(cfg);
    // Persist to Drive so it survives redeploys
    await saveReportToDrive(report);
    res.json({ success: true, foundCount: report.foundCount, notFoundCount: report.notFoundCount });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── Run (SSE) ─────────────────────────────────────────────────────────────────
app.get("/api/run", (req, res) => {
  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache, no-transform");
  res.setHeader("Connection", "keep-alive");
  res.setHeader("X-Accel-Buffering", "no");
  res.flushHeaders();

  const send = (type, message) =>
    res.write(`data: ${JSON.stringify({ type, message })}\n\n`);

  // Check a file is selected
  const cfg = readConfig();
  if (!cfg.excel.file || !fs.existsSync(cfg.excel.file)) {
    send("error", "No file selected. Please upload a CSV or Excel file first.");
    send("done", "1");
    res.end();
    return;
  }

  if (runningProcess) {
    send("error", "A download is already in progress.");
    send("done", "1");
    res.end();
    return;
  }

  spawnWithSSE("download.js", [], res, send);
});

// ── Rescan (SSE) ──────────────────────────────────────────────────────────────
app.get("/api/rescan", (req, res) => {
  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache, no-transform");
  res.setHeader("Connection", "keep-alive");
  res.setHeader("X-Accel-Buffering", "no");
  res.flushHeaders();

  const send = (type, message) =>
    res.write(`data: ${JSON.stringify({ type, message })}\n\n`);

  if (runningProcess) {
    send("error", "A download is already in progress.");
    send("done", "1");
    res.end();
    return;
  }

  spawnWithSSE("download.js", ["--rescan"], res, send);
});

function spawnWithSSE(script, args, res, send) {
  // Keep-alive ping every 20s so Railway doesn't cut the SSE connection
  const ping = setInterval(() => res.write(": ping\n\n"), 20000);

  runningProcess = spawn("node", [script, ...args], { cwd: __dirname });
  runningProcess.stdout.on("data", (d) => send("log", d.toString()));
  runningProcess.stderr.on("data", (d) => send("error", d.toString()));
  runningProcess.on("close", async (code) => {
    clearInterval(ping);
    // Save report to Drive after run so it survives redeploys
    if (fs.existsSync(REPORT_PATH)) {
      try {
        const report = JSON.parse(fs.readFileSync(REPORT_PATH, "utf8"));
        await saveReportToDrive(report);
      } catch {}
    }
    send("done", String(code));
    runningProcess = null;
    res.end();
  });

  res.on("close", () => {
    clearInterval(ping);
    if (runningProcess) { runningProcess.kill(); runningProcess = null; }
  });
}

app.post("/api/stop", (req, res) => {
  if (runningProcess) { runningProcess.kill(); runningProcess = null; res.json({ stopped: true }); }
  else res.json({ stopped: false });
});

// ── Start ─────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`\n  Westside Downloader → http://localhost:${PORT}\n`));
