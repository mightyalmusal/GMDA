// server/publish.js
// Handles Cloudflare R2 upload + manifest versioning for the admin app.

const fs = require("node:fs/promises");
const path = require("node:path");

const DATA_DIR = path.join(process.cwd(), "data");
const DATA_FILES = [
  "meta-insights-cache.json",
  "meta-budget-targets.json",
  "meta-mappings.json",
  "meta-selection-lists.json",
];

function getBearerToken(headers = {}) {
  const value = headers.authorization || headers.Authorization || "";
  const m = String(value).match(/^Bearer\s+(.+)$/i);
  return m ? m[1].trim() : "";
}

// Decode JWT payload without verifying signature (localhost admin use only).
function decodeJwtPayload(token) {
  try {
    const parts = token.split(".");
    if (parts.length < 2) return {};
    const padded = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const json = Buffer.from(padded, "base64").toString("utf8");
    return JSON.parse(json);
  } catch {
    return {};
  }
}

async function getManifest(s3) {
  const { GetObjectCommand } = require("@aws-sdk/client-s3");
  try {
    const response = await s3.send(
      new GetObjectCommand({
        Bucket: process.env.R2_BUCKET_NAME,
        Key: "manifest.json",
      })
    );
    const chunks = [];
    for await (const chunk of response.Body) {
      chunks.push(chunk);
    }
    const raw = Buffer.concat(chunks).toString("utf8");
    return JSON.parse(raw);
  } catch {
    return { version: 0 };
  }
}

async function uploadFile(s3, key, content) {
  const { PutObjectCommand } = require("@aws-sdk/client-s3");
  await s3.send(
    new PutObjectCommand({
      Bucket: process.env.R2_BUCKET_NAME,
      Key: key,
      Body: content,
      ContentType: "application/json",
    })
  );
}

async function publishAll(publishedBy) {
  const { S3Client } = require("@aws-sdk/client-s3");

  const s3 = new S3Client({
    region: "auto",
    endpoint: `https://${process.env.R2_ACCOUNT_ID}.r2.cloudflarestorage.com`,
    credentials: {
      accessKeyId: process.env.R2_ACCESS_KEY_ID,
      secretAccessKey: process.env.R2_SECRET_ACCESS_KEY,
    },
  });

  const manifest = await getManifest(s3);
  const currentVersion = manifest.version || 0;

  for (const filename of DATA_FILES) {
    const filePath = path.join(DATA_DIR, filename);
    let content;
    try {
      content = await fs.readFile(filePath, "utf8");
    } catch {
      // Skip files that don't exist locally
      continue;
    }
    await uploadFile(s3, filename, content);
  }

  const newVersion = currentVersion + 1;
  const publishedAt = new Date().toISOString();
  const newManifest = {
    version: newVersion,
    publishedAt,
    publishedBy,
    files: DATA_FILES,
  };

  await uploadFile(s3, "manifest.json", JSON.stringify(newManifest, null, 2));

  return { success: true, version: newVersion, publishedAt };
}

async function handler(event) {
  try {
    const body = JSON.parse(event.body || "{}");

    if (body.action !== "publish") {
      return {
        statusCode: 400,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ error: 'action must be "publish"' }),
      };
    }

    const token = getBearerToken(event.headers || {});
    let publishedBy = "admin@localhost";

    if (token) {
      const payload = decodeJwtPayload(token);
      publishedBy =
        payload.preferred_username ||
        payload.email ||
        payload.upn ||
        "admin@localhost";
    }

    const result = await publishAll(publishedBy);

    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(result),
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ error: err?.message || "Internal Server Error" }),
    };
  }
}

module.exports = { handler };
