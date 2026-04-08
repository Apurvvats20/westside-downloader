const fs = require("fs");
const path = require("path");
const https = require("https");
const http = require("http");

// ============ LOAD CONFIG ============
const config = JSON.parse(
  fs.readFileSync(path.join(__dirname, "config.json"), "utf8"),
);

const AUTH_COOKIE = config.auth.cookie || "";
const OUTPUT_BASE = path.join(__dirname, "images"); // Always save to images folder
const SITE_URL = "https://trentlimited.sharepoint.com/sites/E-ComWSXYoshiStudio";

// SKUs to rescan with their brands (10 still not found)
const NOT_FOUND_SKUS = [
  { sku: "301059202TAUPE", brand: "Y&F Boys" },
  { sku: "301064963WHITE", brand: "Hop Baby" },
  { sku: "301058074OFF WHITE", brand: "HOP Kids - Junior Boys" },
  { sku: "301059429WHITE", brand: "Hop Baby" },
  { sku: "301059550WHITE", brand: "Hop Baby" },
  { sku: "301060283BROWN", brand: "HOP Kids - Junior Girls" },
  { sku: "301060302BLACK", brand: "Y&F Girls" },
  { sku: "301060302BROWN", brand: "Y&F Girls" },
  { sku: "301060302SILVER", brand: "Y&F Girls" },
  { sku: "301061733OFF WHITE", brand: "HOP Kids - Junior Girls" },
];

function buildHeaders(extra = {}) {
  const headers = {
    Accept: "application/json;odata=verbose",
    "Content-Type": "application/json;odata=verbose",
    ...extra,
  };
  if (AUTH_COOKIE) headers.Cookie = `FedAuth=${AUTH_COOKIE}`;
  return headers;
}

function parseSharePointUrl(url) {
  if (url.startsWith("http")) {
    const parsed = new URL(url);
    if (parsed.pathname.includes("/_layouts/") && parsed.searchParams.has("id")) {
      const serverRelativeUrl = decodeURIComponent(parsed.searchParams.get("id"));
      const layoutsIdx = parsed.pathname.indexOf("/_layouts/");
      const sitePath = parsed.pathname.slice(0, layoutsIdx);
      const site = `${parsed.protocol}//${parsed.host}${sitePath}`;
      return { site, serverRelativeUrl };
    }
    const site = `${parsed.protocol}//${parsed.host}`;
    const serverRelativeUrl = decodeURIComponent(parsed.pathname);
    return { site, serverRelativeUrl };
  }
  return { site: SITE_URL, serverRelativeUrl: url };
}

async function checkAuth() {
  if (!AUTH_COOKIE) {
    console.log("ℹ️  No auth cookie set — using public access mode");
    return true;
  }
  // Test against the specific site collection (guest accounts don't have tenant root access)
  const testUrl = `https://trentlimited.sharepoint.com/sites/E-ComWSXYoshiStudio/_api/web/title`;
  const result = await spRequest(testUrl);
  if (!result || !result.d) {
    console.error("\n❌ Auth cookie appears to be expired or invalid.");
    console.error("   Go to https://trentlimited.sharepoint.com while logged in,");
    console.error("   copy the FedAuth cookie, and update 'auth.cookie' in config.json.\n");
    return false;
  }
  console.log("✅ Auth cookie is valid");
  return true;
}

function spRequest(url) {
  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    const lib = parsed.protocol === "https:" ? https : http;
    const options = { headers: buildHeaders(), method: "GET" };

    lib.get(url, options, (res) => {
      let data = "";
      res.on("data", (chunk) => (data += chunk));
      res.on("end", () => {
        if (res.statusCode === 401 || res.statusCode === 403) {
          console.error(`\n❌ Access denied (${res.statusCode}). Cookie may be expired — update config.json.\n`);
          resolve(null);
          return;
        }
        try {
          resolve(JSON.parse(data));
        } catch (e) {
          resolve(null);
        }
      });
    }).on("error", reject);
  });
}

function isDpiSiblingFolder(folderName, dpiFilter) {
  const knownDpiFolders = ["96 DPI", "300 DPI", "72 DPI", "TIFF", "768 x 1024", "1024 x 1366"];
  return knownDpiFolders.includes(folderName) && folderName !== dpiFilter;
}

async function getFilesRecursive(folderServerRelativeUrl, siteUrl, dpiFilter) {
  const allFiles = [];
  const site = siteUrl || SITE_URL;

  try {
    const filesUrl = `${site}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelativeUrl)}')/Files?$select=Name,ServerRelativeUrl&$top=5000`;
    const filesResponse = await spRequest(filesUrl);

    if (filesResponse && filesResponse.d && filesResponse.d.results) {
      for (const file of filesResponse.d.results) {
        allFiles.push({ name: file.Name, url: file.ServerRelativeUrl });
      }
    }

    const foldersUrl = `${site}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelativeUrl)}')/Folders?$select=Name,ServerRelativeUrl&$top=5000`;
    const foldersResponse = await spRequest(foldersUrl);

    if (foldersResponse && foldersResponse.d && foldersResponse.d.results) {
      for (const folder of foldersResponse.d.results) {
        if (folder.Name === "Forms") continue;
        if (dpiFilter && isDpiSiblingFolder(folder.Name, dpiFilter)) continue;
        const subFiles = await getFilesRecursive(folder.ServerRelativeUrl, site, dpiFilter);
        allFiles.push(...subFiles);
      }
    }
  } catch (e) {
    console.log(`Error scanning folder ${folderServerRelativeUrl}: ${e.message}`);
  }

  return allFiles;
}

function downloadFile(fileServerRelativeUrl, destPath, siteUrl) {
  const site = siteUrl || SITE_URL;
  return new Promise((resolve, reject) => {
    const url = `${site}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(fileServerRelativeUrl)}')/$value`;
    const options = { headers: buildHeaders(), method: "GET" };
    const parsed = new URL(url);
    const lib = parsed.protocol === "https:" ? https : http;

    const file = fs.createWriteStream(destPath);
    lib.get(url, options, (res) => {
      if (res.statusCode === 301 || res.statusCode === 302) {
        https.get(res.headers.location, options, (res2) => {
          res2.pipe(file);
          file.on("finish", () => { file.close(); resolve(); });
        }).on("error", reject);
      } else {
        res.pipe(file);
        file.on("finish", () => { file.close(); resolve(); });
      }
    }).on("error", (err) => {
      fs.unlink(destPath, () => {});
      reject(err);
    });
  });
}

async function main() {
  console.log("\n========== Rescan: Not Found SKUs ==========\n");
  console.log(`Rescanning ${NOT_FOUND_SKUS.length} SKUs...\n`);

  const authOk = await checkAuth();
  if (!authOk) process.exit(1);

  const cfg = config.excel;
  const searchFolders = cfg.search_folders;
  const dpiFilter = cfg.dpi_folder || null;

  // Scan all SharePoint folders
  console.log("\nScanning SharePoint folders...");
  const allFiles = [];
  for (const folderUrl of searchFolders) {
    const { serverRelativeUrl } = parseSharePointUrl(folderUrl);
    const label = serverRelativeUrl.split("/").pop();
    console.log(`Scanning: ${serverRelativeUrl}`);
    const files = await getFilesRecursive(serverRelativeUrl, SITE_URL, dpiFilter);
    console.log(`Found ${files.length} files in ${label}`);
    allFiles.push(...files);
  }
  console.log(`\nTotal files indexed: ${allFiles.length}`);

  // Match and download
  const nowFound = [];
  const stillNotFound = [];

  for (const { sku, brand } of NOT_FOUND_SKUS) {
    const skuFolder = path.join(OUTPUT_BASE, brand, sku);

    const skuNorm = sku.toLowerCase().replace(/\s+/g, "");
    const matchingFiles = allFiles.filter((file) => {
      const nameWithoutExt = path.parse(file.name).name.toLowerCase().replace(/\s+/g, "");
      return (
        nameWithoutExt.startsWith(skuNorm + "_") ||
        nameWithoutExt === skuNorm
      );
    });

    if (matchingFiles.length === 0) {
      console.log(`❌ Still not found: ${sku}`);
      stillNotFound.push(sku);
      continue;
    }

    fs.mkdirSync(skuFolder, { recursive: true });
    console.log(`\n✅ NOW FOUND ${matchingFiles.length} files for ${sku} (${brand})`);

    for (const file of matchingFiles) {
      const destPath = path.join(skuFolder, file.name);
      try {
        await downloadFile(file.url, destPath, SITE_URL);
        console.log(`   Downloaded: ${file.name}`);
        nowFound.push({ sku, brand, file: file.name });
      } catch (e) {
        console.log(`   ❌ Failed: ${file.name} — ${e.message}`);
      }
    }
  }

  // Update not_found_log.txt with remaining not-found SKUs
  const logPath = path.join(__dirname, "not_found_log.txt");
  if (stillNotFound.length > 0) {
    fs.writeFileSync(logPath, stillNotFound.join("\n"), "utf8");
  } else {
    fs.writeFileSync(logPath, "", "utf8");
  }

  console.log("\n========== RESCAN SUMMARY ==========");
  const newlyFoundSkus = [...new Set(nowFound.map((r) => r.sku))];
  console.log(`✅ Newly found & downloaded: ${newlyFoundSkus.length}`);
  if (newlyFoundSkus.length > 0) {
    newlyFoundSkus.forEach((sku) => console.log(`   - ${sku}`));
  }
  console.log(`❌ Still not found: ${stillNotFound.length}`);
  if (stillNotFound.length > 0) {
    stillNotFound.forEach((sku) => console.log(`   - ${sku}`));
  }
  console.log(`📁 Files saved to: ${OUTPUT_BASE}`);
  console.log("=====================================\n");
}

main().catch(console.error);
