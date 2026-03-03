#!/usr/bin/env node
import fs from "node:fs/promises";
import http from "node:http";
import path from "node:path";
import process from "node:process";
import { spawn } from "node:child_process";

const HOST = "127.0.0.1";
const PORT = Number.parseInt(process.env.ZONE_CARD_UI_PORT || "4177", 10);
const ROOT = process.cwd();
const ZONES_PATH = path.resolve(ROOT, "data/zones.json");
const GENERATOR = path.resolve(ROOT, "scripts/generate-zone-cards-openai.mjs");
const PAGE_PATH = path.resolve(ROOT, "scripts/zone-card-generator-ui.html");
const IMAGE_EXTS = new Set([".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"]);
const SKIP_DIRS = new Set([".git", "node_modules"]);

let generating = false;

function normalizePath(v) {
  return String(v || "").replaceAll("\\", "/").trim();
}

function imageUrl(repoPath) {
  return "/repo-file/" + repoPath.split("/").map((p) => encodeURIComponent(p)).join("/");
}

function imageMime(file) {
  const ext = path.extname(file).toLowerCase();
  if (ext === ".png") return "image/png";
  if (ext === ".jpg" || ext === ".jpeg") return "image/jpeg";
  if (ext === ".webp") return "image/webp";
  if (ext === ".gif") return "image/gif";
  if (ext === ".bmp") return "image/bmp";
  return "application/octet-stream";
}

function sendJson(res, code, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(code, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body)
  });
  res.end(body);
}

function sendHtml(res, code, body) {
  res.writeHead(code, {
    "Content-Type": "text/html; charset=utf-8",
    "Content-Length": Buffer.byteLength(body)
  });
  res.end(body);
}

async function readUtf8(file) {
  const raw = await fs.readFile(file, "utf8");
  return raw.replace(/^\uFEFF/, "");
}

async function exists(file) {
  try {
    await fs.access(file);
    return true;
  } catch {
    return false;
  }
}

async function loadZones() {
  const doc = JSON.parse(await readUtf8(ZONES_PATH));
  const tiles = Array.isArray(doc?.tiles) ? doc.tiles : [];
  return tiles
    .map((t) => ({
      id: String(t.id || ""),
      name: String(t.name || t.id || ""),
      biome: String(t.biome || "unknown"),
      image: normalizePath(t.image || "")
    }))
    .filter((t) => t.id)
    .sort((a, b) => a.id.localeCompare(b.id));
}

async function walkImages(absDir, relDir, out, zoneByImage) {
  const entries = await fs.readdir(absDir, { withFileTypes: true });
  for (const e of entries) {
    if (e.name.startsWith(".") && e.name !== ".env.example") continue;
    if (e.isDirectory() && SKIP_DIRS.has(e.name)) continue;

    const rel = relDir ? `${relDir}/${e.name}` : e.name;
    const abs = path.join(absDir, e.name);

    if (e.isDirectory()) {
      await walkImages(abs, rel, out, zoneByImage);
      continue;
    }
    if (!e.isFile()) continue;
    const ext = path.extname(e.name).toLowerCase();
    if (!IMAGE_EXTS.has(ext)) continue;

    const repoPath = normalizePath(rel);
    const st = await fs.stat(abs);
    const zone = zoneByImage.get(repoPath) || null;
    out.push({
      repoPath,
      fileName: e.name,
      mtimeMs: st.mtimeMs,
      bytes: st.size,
      zoneId: zone?.id || "",
      zoneName: zone?.name || "",
      biome: zone?.biome || "",
      url: `${imageUrl(repoPath)}?v=${Math.floor(st.mtimeMs)}`
    });
  }
}

async function scanImages() {
  const zones = await loadZones();
  const zoneByImage = new Map();
  for (const z of zones) if (z.image) zoneByImage.set(z.image, z);
  const out = [];
  await walkImages(ROOT, "", out, zoneByImage);
  out.sort((a, b) => a.repoPath.localeCompare(b.repoPath));
  return out;
}

async function hasApiKey() {
  if (String(process.env.OPENAI_API_KEY || "").trim()) return true;
  for (const f of [".env.local", ".env"]) {
    try {
      const text = await readUtf8(path.resolve(f));
      if (/^\s*OPENAI_API_KEY\s*=.+/m.test(text)) return true;
    } catch {
      // Ignore missing files.
    }
  }
  return false;
}

function parseBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let total = 0;
    const max = 1024 * 1024;
    req.on("data", (c) => {
      total += c.length;
      if (total > max) {
        reject(new Error("Request body too large"));
        req.destroy();
        return;
      }
      chunks.push(c);
    });
    req.on("end", () => {
      try {
        const text = Buffer.concat(chunks).toString("utf8");
        resolve(text ? JSON.parse(text) : {});
      } catch (err) {
        reject(err);
      }
    });
    req.on("error", reject);
  });
}

function parseSavedPaths(stdoutText) {
  const out = new Set();
  for (const line of String(stdoutText || "").split(/\r?\n/)) {
    const m = line.match(/saved ->\s+(.+)$/);
    if (m) out.add(normalizePath(m[1]));
  }
  return [...out];
}

async function plannedAssets(all, zoneIds) {
  const zones = await loadZones();
  const idSet = new Set(zoneIds || []);
  const selected = all ? zones : zones.filter((z) => idSet.has(z.id));
  return selected.map((z) => normalizePath(z.image)).filter(Boolean);
}

function runGenerator(opts) {
  return new Promise((resolve) => {
    const args = [GENERATOR, "--style", opts.style];
    if (opts.all) args.push("--all");
    else args.push("--zones", opts.zoneIds.join(","));
    if (opts.skipExisting) args.push("--skip-existing");
    if (opts.dryRun) args.push("--dry-run");
    if (opts.model) args.push("--model", opts.model);
    if (opts.size) args.push("--size", opts.size);
    if (opts.quality) args.push("--quality", opts.quality);
    if (Number.isFinite(opts.delayMs) && opts.delayMs > 0) {
      args.push("--delay-ms", String(Math.floor(opts.delayMs)));
    }

    const child = spawn(process.execPath, args, { cwd: ROOT, env: process.env, shell: false });
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (d) => { stdout += d.toString("utf8"); });
    child.stderr.on("data", (d) => { stderr += d.toString("utf8"); });
    child.on("close", (code) => {
      resolve({
        code: Number.isFinite(code) ? code : 1,
        stdout,
        stderr,
        command: `${process.execPath} ${args.map((a) => JSON.stringify(a)).join(" ")}`
      });
    });
  });
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url || "/", `http://${HOST}:${PORT}`);

    if (req.method === "GET" && url.pathname === "/") {
      sendHtml(res, 200, await readUtf8(PAGE_PATH));
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/health") {
      sendJson(res, 200, { ok: true, keyConfigured: await hasApiKey() });
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/zones") {
      sendJson(res, 200, { ok: true, tiles: await loadZones() });
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/images") {
      sendJson(res, 200, { ok: true, images: await scanImages() });
      return;
    }

    if (req.method === "GET" && url.pathname.startsWith("/repo-file/")) {
      const rel = normalizePath(decodeURIComponent(url.pathname.slice("/repo-file/".length)));
      const abs = path.resolve(ROOT, rel);
      const rootWithSep = ROOT.endsWith(path.sep) ? ROOT : `${ROOT}${path.sep}`;
      if (!abs.startsWith(rootWithSep)) {
        sendJson(res, 400, { ok: false, error: "Invalid file path" });
        return;
      }
      const ext = path.extname(abs).toLowerCase();
      if (!IMAGE_EXTS.has(ext)) {
        sendJson(res, 400, { ok: false, error: "Not an image file" });
        return;
      }
      try {
        const buf = await fs.readFile(abs);
        res.writeHead(200, {
          "Content-Type": imageMime(abs),
          "Content-Length": buf.length,
          "Cache-Control": "no-cache"
        });
        res.end(buf);
      } catch {
        sendJson(res, 404, { ok: false, error: "File not found" });
      }
      return;
    }

    if (req.method === "POST" && url.pathname === "/api/generate") {
      if (generating) {
        sendJson(res, 409, { ok: false, error: "Generation already in progress." });
        return;
      }

      const body = await parseBody(req);
      const style = String(body.style || "").trim();
      const all = Boolean(body.all);
      const zoneIds = Array.isArray(body.zoneIds) ? body.zoneIds.map((v) => String(v)) : [];
      const model = String(body.model || "gpt-image-1").trim();
      const size = String(body.size || "1536x1024").trim();
      const quality = String(body.quality || "high").trim();
      const delayMs = Number.parseInt(String(body.delayMs || "0"), 10);
      const skipExisting = Boolean(body.skipExisting);
      const dryRun = Boolean(body.dryRun);

      if (!style) {
        sendJson(res, 400, { ok: false, error: "style is required" });
        return;
      }
      if (!all && zoneIds.length === 0) {
        sendJson(res, 400, { ok: false, error: "Provide zoneIds or set all=true" });
        return;
      }

      generating = true;
      try {
        const start = Date.now();
        const targets = await plannedAssets(all, zoneIds);
        const existed = new Map(await Promise.all(targets.map(async (p) => [p, await exists(path.resolve(ROOT, p))])));
        const result = await runGenerator({
          style,
          all,
          zoneIds,
          model,
          size,
          quality,
          delayMs: Number.isFinite(delayMs) ? Math.max(delayMs, 0) : 0,
          skipExisting,
          dryRun
        });
        const changedImages = parseSavedPaths(result.stdout);
        const newImages = [];
        const updatedImages = [];
        for (const p of changedImages) {
          if (existed.get(p)) updatedImages.push(p);
          else newImages.push(p);
        }
        sendJson(res, result.code === 0 ? 200 : 500, {
          ok: result.code === 0,
          code: result.code,
          stdout: result.stdout,
          stderr: result.stderr,
          command: result.command,
          durationMs: Date.now() - start,
          changedImages,
          newImages,
          updatedImages
        });
      } finally {
        generating = false;
      }
      return;
    }

    sendJson(res, 404, { ok: false, error: "Not found" });
  } catch (err) {
    generating = false;
    sendJson(res, 500, { ok: false, error: err instanceof Error ? err.message : String(err) });
  }
});

server.listen(PORT, HOST, () => {
  console.log(`Zone Card Generator UI running at http://${HOST}:${PORT}`);
  console.log("Press Ctrl+C to stop.");
});
