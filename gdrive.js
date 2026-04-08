const { google } = require("googleapis");
const fs = require("fs");
const path = require("path");

const CONFIG_PATH = path.join(__dirname, "config.json");

function getConfig() {
  return JSON.parse(fs.readFileSync(CONFIG_PATH, "utf8"));
}

function saveTokens(tokens) {
  const config = getConfig();
  config.google.tokens = { ...(config.google.tokens || {}), ...tokens };
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(config, null, 2), "utf8");
}

function getOAuth2Client() {
  const { google: gcfg } = getConfig();
  const redirectUri = process.env.GOOGLE_REDIRECT_URI || gcfg.redirect_uri;
  const client = new google.auth.OAuth2(
    gcfg.client_id,
    gcfg.client_secret,
    redirectUri
  );
  if (gcfg.tokens) {
    client.setCredentials(gcfg.tokens);
    client.on("tokens", (tokens) => saveTokens(tokens));
  }
  return client;
}

function getAuthUrl() {
  const client = getOAuth2Client();
  return client.generateAuthUrl({
    access_type: "offline",
    scope: ["https://www.googleapis.com/auth/drive"],
    prompt: "consent",
  });
}

async function handleCallback(code) {
  const client = getOAuth2Client();
  const { tokens } = await client.getToken(code);
  saveTokens(tokens);
  client.setCredentials(tokens);
  return client;
}

function isConnected() {
  const config = getConfig();
  return !!(config.google && config.google.tokens);
}

async function createDriveFolder(name, parentId = null) {
  const auth = getOAuth2Client();
  const drive = google.drive({ version: "v3", auth });
  const meta = { name, mimeType: "application/vnd.google-apps.folder" };
  if (parentId) meta.parents = [parentId];
  const res = await drive.files.create({ requestBody: meta, fields: "id" });
  return res.data.id;
}

async function setPublicEdit(fileId) {
  const auth = getOAuth2Client();
  const drive = google.drive({ version: "v3", auth });
  await drive.permissions.create({
    fileId,
    requestBody: { role: "writer", type: "anyone" },
  });
}

async function uploadFile(localPath, name, parentId) {
  const auth = getOAuth2Client();
  const drive = google.drive({ version: "v3", auth });
  const res = await drive.files.create({
    requestBody: { name, parents: [parentId] },
    media: { body: fs.createReadStream(localPath) },
    fields: "id",
  });
  return res.data.id;
}

// Uploads the full brand/sku/files structure from outputBase to a new dated Drive folder
async function uploadOutputFolder(outputBase, folderName, log = console.log) {

  log(`\n☁️  Creating Google Drive folder: "${folderName}"...`);
  const rootId = await createDriveFolder(folderName);
  await setPublicEdit(rootId);

  const brands = fs.readdirSync(outputBase).filter((f) =>
    fs.statSync(path.join(outputBase, f)).isDirectory()
  );

  let totalUploaded = 0;

  for (const brand of brands) {
    const brandPath = path.join(outputBase, brand);
    const brandFolderId = await createDriveFolder(brand, rootId);

    const skus = fs.readdirSync(brandPath).filter((f) =>
      fs.statSync(path.join(brandPath, f)).isDirectory()
    );

    for (const sku of skus) {
      const skuPath = path.join(brandPath, sku);
      const skuFolderId = await createDriveFolder(sku, brandFolderId);

      const files = fs.readdirSync(skuPath).filter((f) =>
        fs.statSync(path.join(skuPath, f)).isFile()
      );

      for (const file of files) {
        await uploadFile(path.join(skuPath, file), file, skuFolderId);
        totalUploaded++;
      }
      log(`   ☁️  Uploaded: ${brand}/${sku} (${files.length} files)`);
    }
  }

  const link = `https://drive.google.com/drive/folders/${rootId}`;
  log(`\n✅ Google Drive upload complete — ${totalUploaded} files uploaded`);
  log(`🔗 DRIVE_LINK:${link}`);
  return { folderId: rootId, link, folderName };
}

module.exports = { getAuthUrl, handleCallback, isConnected, uploadOutputFolder };
