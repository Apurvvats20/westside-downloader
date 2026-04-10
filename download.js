const fs = require("fs");
const path = require("path");
const os = require("os");
const https = require("https");
const http = require("http");
const xlsx = require("xlsx");
const { isConnected, uploadOutputFolder } = require("./gdrive");

// ============ LOAD CONFIG ============
const config = JSON.parse(
  fs.readFileSync(path.join(__dirname, "config.json"), "utf8"),
);

const SHAREPOINT_SITE = config.sharepoint_site;
const AUTH_COOKIE = config.auth.cookie || "";
const MODE = config.mode;
const RESCAN_MODE = process.argv.includes("--rescan");
const REPORT_PATH = path.join(__dirname, "uploads", "report.json");

// Dated output folder — ~/Downloads locally, /tmp on server
const today = new Date();
const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const DATE_STR = `${String(today.getDate()).padStart(2,"0")} ${MONTHS[today.getMonth()]} ${today.getFullYear()}`;
const RUN_FOLDER_NAME = RESCAN_MODE
  ? `WESTSIDE - ${DATE_STR} (Rescan)`
  : `WESTSIDE - ${DATE_STR}`;
const IS_SERVER = process.env.NODE_ENV === "production";
const OUTPUT_BASE = IS_SERVER
  ? path.join(os.tmpdir(), "westside", RUN_FOLDER_NAME)
  : path.join(os.homedir(), "Downloads", RUN_FOLDER_NAME);

// ============ AUTH HEADERS ============
function buildHeaders(extra = {}) {
  const headers = {
    Accept: "application/json;odata=verbose",
    "Content-Type": "application/json;odata=verbose",
    ...extra,
  };
  if (AUTH_COOKIE) headers.Cookie = `FedAuth=${AUTH_COOKIE}`;
  return headers;
}

// ============ SHAREPOINT HELPERS ============
function parseSharePointUrl(url) {
  if (url.startsWith("http")) {
    const parsed = new URL(url);
    if (parsed.pathname.includes("/_layouts/") && parsed.searchParams.has("id")) {
      const serverRelativeUrl = decodeURIComponent(parsed.searchParams.get("id"));
      const layoutsIdx = parsed.pathname.indexOf("/_layouts/");
      const sitePath = parsed.pathname.slice(0, layoutsIdx);
      return { site: `${parsed.protocol}//${parsed.host}${sitePath}`, serverRelativeUrl };
    }
    const match = parsed.pathname.match(/^(\/sites\/[^/]+)/);
    const sitePath = match ? match[1] : "";
    return {
      site: `${parsed.protocol}//${parsed.host}${sitePath}`,
      serverRelativeUrl: decodeURIComponent(parsed.pathname),
    };
  }
  return { site: SHAREPOINT_SITE, serverRelativeUrl: url };
}

async function checkAuth() {
  if (!AUTH_COOKIE) { console.log("ℹ️  No auth cookie — public access mode"); return true; }
  const result = await spRequest(`${SHAREPOINT_SITE}/_api/web/title`);
  if (!result || !result.d) {
    console.error("\n❌ Auth cookie expired — update via Admin Settings.\n");
    return false;
  }
  console.log("✅ Auth cookie is valid");
  return true;
}

function spRequest(url) {
  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    const lib = parsed.protocol === "https:" ? https : http;
    lib.get(url, { headers: buildHeaders(), method: "GET" }, (res) => {
      let data = "";
      res.on("data", (chunk) => (data += chunk));
      res.on("end", () => {
        if (res.statusCode === 401 || res.statusCode === 403) {
          console.error(`\n❌ Access denied (${res.statusCode}). Cookie may be expired.\n`);
          resolve(null);
          return;
        }
        try { resolve(JSON.parse(data)); } catch { resolve(null); }
      });
    }).on("error", reject);
  });
}

// Concurrency limiter — avoids hammering SharePoint with too many parallel requests
function limitConcurrency(tasks, limit = 5) {
  return new Promise((resolve) => {
    if (tasks.length === 0) return resolve([]);
    const results = [];
    let started = 0, finished = 0;
    function next() {
      if (finished === tasks.length) return resolve(results);
      while (started < tasks.length && started - finished < limit) {
        const i = started++;
        tasks[i]()
          .then((r) => { results[i] = r; })
          .catch(() => { results[i] = []; })
          .finally(() => { finished++; next(); });
      }
    }
    next();
  });
}

async function getFilesRecursive(folderServerRelativeUrl, siteUrl, dpiFilter, depth = 0) {
  const site = siteUrl || SHAREPOINT_SITE;
  try {
    const [filesResponse, foldersResponse] = await Promise.all([
      spRequest(`${site}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelativeUrl)}')/Files?$select=Name,ServerRelativeUrl&$top=5000`),
      spRequest(`${site}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelativeUrl)}')/Folders?$select=Name,ServerRelativeUrl&$top=5000`),
    ]);

    const files = (filesResponse?.d?.results || [])
      .map((f) => ({ name: f.Name, url: f.ServerRelativeUrl }));

    const subFolders = (foldersResponse?.d?.results || [])
      .filter((f) => f.Name !== "Forms" && !(dpiFilter && isDpiSiblingFolder(f.Name, dpiFilter)) && !isQuinnOutputsFolder(f.Name));

    // Log subfolder names at depth 0 and 1 so user sees progress
    if (depth <= 1 && subFolders.length > 0) {
      const folderName = folderServerRelativeUrl.split("/").pop();
      console.log(`   📂 Scanning ${folderName} — ${subFolders.length} subfolders found`);
    }
    if (files.length > 0 && depth >= 1) {
      const folderName = folderServerRelativeUrl.split("/").pop();
      console.log(`   🖼️  ${folderName} — ${files.length} file${files.length > 1 ? "s" : ""}`);
    }

    const subResults = await limitConcurrency(
      subFolders.map((folder) => () => getFilesRecursive(folder.ServerRelativeUrl, site, dpiFilter, depth + 1)),
      5
    );

    return [...files, ...subResults.flat()];
  } catch (e) {
    console.log(`   ⚠️  Error scanning ${folderServerRelativeUrl.split("/").pop()}: ${e.message}`);
    return [];
  }
}

function isDpiSiblingFolder(folderName, dpiFilter) {
  const known = ["96 DPI","300 DPI","72 DPI","TIFF","768 x 1024","1024 x 1366"];
  return known.includes(folderName) && folderName !== dpiFilter;
}

function isQuinnOutputsFolder(folderName) {
  // Normalise: lowercase, strip spaces/hyphens/underscores, then check
  const norm = folderName.toLowerCase().replace(/[\s\-_]+/g, "");
  return norm.includes("quinnoutput") || norm.includes("quinnouput");
}

function downloadFile(fileServerRelativeUrl, destPath, siteUrl) {
  const site = siteUrl || SHAREPOINT_SITE;
  return new Promise((resolve, reject) => {
    const url = `${site}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(fileServerRelativeUrl)}')/$value`;
    const lib = new URL(url).protocol === "https:" ? https : http;
    const file = fs.createWriteStream(destPath);
    lib.get(url, { headers: buildHeaders(), method: "GET" }, (res) => {
      if (res.statusCode === 301 || res.statusCode === 302) {
        https.get(res.headers.location, { headers: buildHeaders() }, (res2) => {
          res2.pipe(file);
          file.on("finish", () => { file.close(); resolve(); });
        }).on("error", reject);
      } else {
        res.pipe(file);
        file.on("finish", () => { file.close(); resolve(); });
      }
    }).on("error", (err) => { fs.unlink(destPath, () => {}); reject(err); });
  });
}

// ============ SHARED: scan SP + download matching SKUs ============
async function scanAndDownload(skusToProcess, cfg) {
  console.log("\nScanning SharePoint folders...");
  const scanResults = await Promise.all(
    cfg.search_folders.map(async (folderUrl) => {
      const { site, serverRelativeUrl } = parseSharePointUrl(folderUrl);
      const label = serverRelativeUrl.split("/").pop();
      console.log(`Scanning: ${serverRelativeUrl}`);
      const files = await getFilesRecursive(serverRelativeUrl, site, cfg.dpi_folder || null);
      console.log(`Found ${files.length} files in ${label}`);
      return files;
    })
  );
  const allFiles = scanResults.flat();
  console.log(`\nTotal files indexed: ${allFiles.length}`);

  const results = { found: [], notFound: [] };

  for (const { sku, brand } of skusToProcess) {
    const skuFolder = path.join(OUTPUT_BASE, brand, sku);
    const matchingFiles = allFiles.filter((file) => {
      const nameWithoutExt = path.parse(file.name).name.toLowerCase();
      return nameWithoutExt.startsWith(sku.toLowerCase() + "_") || nameWithoutExt === sku.toLowerCase();
    });

    if (matchingFiles.length === 0) {
      console.log(`❌ Not found: ${sku}`);
      results.notFound.push({ sku, brand });
      continue;
    }

    fs.mkdirSync(skuFolder, { recursive: true });
    console.log(`\n✅ Found ${matchingFiles.length} files for ${sku}`);
    for (const file of matchingFiles) {
      const { site } = parseSharePointUrl(cfg.search_folders[0]);
      try {
        await downloadFile(file.url, path.join(skuFolder, file.name), site);
        console.log(`   Downloaded: ${file.name}`);
        results.found.push({ sku, brand, file: file.name });
      } catch (e) {
        console.log(`   ❌ Failed: ${file.name} — ${e.message}`);
      }
    }
  }

  return results;
}

// ============ MODE: FULL EXCEL RUN ============
async function runExcelMode() {
  const cfg = config.excel;
  const workbook = xlsx.readFile(cfg.file);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(sheet);
  const skus = data
    .map((row) => ({
      sku: String(row[cfg.sku_column] || row[cfg.sku_column.toLowerCase()] || "").trim(),
      brand: String(row[cfg.brand_column] || row[cfg.brand_column.toLowerCase()] || "Unknown").trim(),
    }))
    .filter((r) => r.sku);

  console.log(`✅ Read ${skus.length} SKUs from file`);

  const results = await scanAndDownload(skus, cfg);
  const foundSkus = [...new Set(results.found.map((r) => r.sku))];

  // Save persistent report
  const reportData = {
    lastRun: new Date().toISOString(),
    mode: "full",
    file: path.basename(cfg.file),
    totalSkus: skus.length,
    foundCount: foundSkus.length,
    notFoundCount: results.notFound.length,
    skus: [
      ...foundSkus.map((sku) => ({
        sku,
        brand: skus.find((s) => s.sku === sku)?.brand || "",
        status: "Found",
        fileCount: results.found.filter((r) => r.sku === sku).length,
      })),
      ...results.notFound.map(({ sku, brand }) => ({ sku, brand, status: "Not Found", fileCount: 0 })),
    ],
  };
  fs.mkdirSync(path.dirname(REPORT_PATH), { recursive: true });
  fs.writeFileSync(REPORT_PATH, JSON.stringify(reportData, null, 2), "utf8");

  // Also persist report inside config.json so it survives restarts
  const fullConfig = JSON.parse(fs.readFileSync(path.join(__dirname, "config.json"), "utf8"));
  const uploadedFiles = fullConfig.uploadedFiles || {};
  const fileEntry = Object.values(uploadedFiles).find(e => cfg.file.endsWith(path.basename(cfg.file)));
  const reportKey = fileEntry ? `report_${fileEntry.driveFileId}` : `report_${path.basename(cfg.file)}`;
  if (!fullConfig.reports) fullConfig.reports = {};
  fullConfig.reports[reportKey] = reportData;
  fullConfig.currentReportKey = reportKey;
  fs.writeFileSync(path.join(__dirname, "config.json"), JSON.stringify(fullConfig, null, 2), "utf8");

  console.log("\n========== SUMMARY ==========");
  console.log(`✅ SKUs found and downloaded: ${foundSkus.length}`);
  console.log(`❌ SKUs not found: ${results.notFound.length}`);
  console.log(`📊 REPORT:${JSON.stringify({ found: foundSkus.length, notFound: results.notFound.length })}`);
  console.log("==============================\n");

  if (foundSkus.length > 0 && isConnected()) {
    try {
      await uploadOutputFolder(OUTPUT_BASE, RUN_FOLDER_NAME, (msg) => process.stdout.write(msg + "\n"));
      if (IS_SERVER) fs.rmSync(OUTPUT_BASE, { recursive: true, force: true });
    } catch (e) {
      console.error(`\n❌ Google Drive upload failed: ${e.message}`);
    }
  } else if (!isConnected()) {
    console.log("ℹ️  Google Drive not connected — skipping upload.");
  }
}

// ============ MODE: RESCAN NOT FOUND ============
async function runRescanMode() {
  // Try local file first; if missing (e.g. after a redeploy), pull from Drive
  if (!fs.existsSync(REPORT_PATH)) {
    console.log("⚠️  Local report not found — fetching from Google Drive...");
    try {
      const { isConnected, listFilesInFolder, downloadSmallFile, getUploadsFolderId } = require("./gdrive");
      if (isConnected()) {
        const folderId = await getUploadsFolderId();
        const files = await listFilesInFolder(folderId);
        const reportFile = files.find(f => f.name === "westside_report.json");
        if (reportFile) {
          const content = await downloadSmallFile(reportFile.id);
          const driveReport = typeof content === "string" ? JSON.parse(content) : content;
          fs.mkdirSync(path.dirname(REPORT_PATH), { recursive: true });
          fs.writeFileSync(REPORT_PATH, JSON.stringify(driveReport, null, 2), "utf8");
          console.log("✅ Report restored from Drive");
        }
      }
    } catch (e) {
      console.error("Failed to fetch report from Drive:", e.message);
    }
  }

  if (!fs.existsSync(REPORT_PATH)) {
    console.error("❌ No previous report found. Run a full scan first.");
    process.exit(1);
  }

  const report = JSON.parse(fs.readFileSync(REPORT_PATH, "utf8"));
  const notFoundSkus = report.skus.filter((s) => s.status === "Not Found");

  if (notFoundSkus.length === 0) {
    console.log("✅ All SKUs from last run were already found. Nothing to rescan.");
    process.exit(0);
  }

  console.log(`🔄 Rescanning ${notFoundSkus.length} previously not-found SKUs...`);

  const cfg = config.excel;
  const results = await scanAndDownload(notFoundSkus, cfg);
  const newlyFoundSkus = [...new Set(results.found.map((r) => r.sku))];

  // Update the persistent report — flip newly found SKUs to "Found"
  for (const entry of report.skus) {
    if (newlyFoundSkus.includes(entry.sku)) {
      entry.status = "Found";
      entry.fileCount = results.found.filter((r) => r.sku === entry.sku).length;
    }
  }
  report.foundCount = report.skus.filter((s) => s.status === "Found").length;
  report.notFoundCount = report.skus.filter((s) => s.status === "Not Found").length;
  report.lastRescan = new Date().toISOString();
  fs.writeFileSync(REPORT_PATH, JSON.stringify(report, null, 2), "utf8");

  console.log("\n========== RESCAN SUMMARY ==========");
  console.log(`✅ Newly found: ${newlyFoundSkus.length}`);
  console.log(`❌ Still not found: ${results.notFound.length}`);
  console.log(`📊 REPORT:${JSON.stringify({ found: newlyFoundSkus.length, notFound: results.notFound.length })}`);
  console.log("=====================================\n");

  if (newlyFoundSkus.length > 0 && isConnected()) {
    try {
      await uploadOutputFolder(OUTPUT_BASE, RUN_FOLDER_NAME, (msg) => process.stdout.write(msg + "\n"));
      if (IS_SERVER) fs.rmSync(OUTPUT_BASE, { recursive: true, force: true });
    } catch (e) {
      console.error(`\n❌ Google Drive upload failed: ${e.message}`);
    }
  } else if (newlyFoundSkus.length === 0) {
    console.log("ℹ️  No new SKUs found — nothing to upload.");
  } else if (!isConnected()) {
    console.log("ℹ️  Google Drive not connected — skipping upload.");
  }
}

// ============ MAIN ============
async function main() {
  console.log(`\n========== Westside Downloader [${RESCAN_MODE ? "RESCAN" : "FULL SCAN"}] ==========\n`);
  const authOk = await checkAuth();
  if (!authOk) process.exit(1);

  if (RESCAN_MODE) {
    await runRescanMode();
  } else if (MODE === "excel") {
    await runExcelMode();
  } else {
    console.error(`Unknown mode "${MODE}".`);
    process.exit(1);
  }
}

main().catch(console.error);
