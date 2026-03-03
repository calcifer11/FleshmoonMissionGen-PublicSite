#!/usr/bin/env node
import fs from "node:fs/promises";
import http from "node:http";
import path from "node:path";
import process from "node:process";
import { spawn } from "node:child_process";

const HOST = "127.0.0.1";
const PORT = Number.parseInt(process.env.ZONE_CARD_UI_PORT || "4177", 10);
const ZONES_PATH = path.resolve(process.cwd(), "data/zones.json");
const GENERATOR_SCRIPT_PATH = path.resolve(process.cwd(), "scripts/generate-zone-cards-openai.mjs");

let generationInProgress = false;

function sendJson(res, statusCode, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body)
  });
  res.end(body);
}

function sendHtml(res, statusCode, html) {
  res.writeHead(statusCode, {
    "Content-Type": "text/html; charset=utf-8",
    "Content-Length": Buffer.byteLength(html)
  });
  res.end(html);
}

async function readTextUtf8(filePath) {
  const raw = await fs.readFile(filePath, "utf8");
  return raw.replace(/^\uFEFF/, "");
}

async function loadZones() {
  const raw = await readTextUtf8(ZONES_PATH);
  const doc = JSON.parse(raw);
  const tiles = Array.isArray(doc?.tiles) ? doc.tiles : [];
  return tiles
    .map((tile) => ({
      id: String(tile.id || ""),
      name: String(tile.name || tile.id || ""),
      biome: String(tile.biome || "unknown"),
      image: String(tile.image || "")
    }))
    .filter((tile) => tile.id.length > 0)
    .sort((a, b) => {
      const biomeCmp = a.biome.localeCompare(b.biome);
      if (biomeCmp !== 0) return biomeCmp;
      return a.name.localeCompare(b.name);
    });
}

async function hasApiKeyConfigured() {
  if (String(process.env.OPENAI_API_KEY || "").trim()) {
    return true;
  }

  const candidates = [path.resolve(".env.local"), path.resolve(".env")];
  for (const candidate of candidates) {
    try {
      const text = await readTextUtf8(candidate);
      const lines = text.split(/\r?\n/);
      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed || trimmed.startsWith("#")) continue;
        if (/^OPENAI_API_KEY\s*=/.test(trimmed)) return true;
      }
    } catch {
      // Ignore missing files.
    }
  }
  return false;
}

function parseRequestBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let totalBytes = 0;
    const maxBytes = 1024 * 1024;

    req.on("data", (chunk) => {
      totalBytes += chunk.length;
      if (totalBytes > maxBytes) {
        reject(new Error("Request body too large"));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });

    req.on("end", () => {
      try {
        const text = Buffer.concat(chunks).toString("utf8");
        const parsed = text ? JSON.parse(text) : {};
        resolve(parsed);
      } catch (err) {
        reject(err);
      }
    });

    req.on("error", (err) => reject(err));
  });
}

function runGenerator(options) {
  return new Promise((resolve) => {
    const args = [GENERATOR_SCRIPT_PATH];
    args.push("--style", options.style);

    if (options.all) {
      args.push("--all");
    } else {
      args.push("--zones", options.zoneIds.join(","));
    }

    if (options.skipExisting) args.push("--skip-existing");
    if (options.dryRun) args.push("--dry-run");
    if (options.model) args.push("--model", options.model);
    if (options.size) args.push("--size", options.size);
    if (options.quality) args.push("--quality", options.quality);
    if (Number.isFinite(options.delayMs) && options.delayMs > 0) {
      args.push("--delay-ms", String(Math.floor(options.delayMs)));
    }

    const child = spawn(process.execPath, args, {
      cwd: process.cwd(),
      env: process.env,
      shell: false
    });

    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (data) => {
      stdout += data.toString("utf8");
    });
    child.stderr.on("data", (data) => {
      stderr += data.toString("utf8");
    });

    child.on("close", (code) => {
      resolve({
        code: Number.isFinite(code) ? code : 1,
        stdout,
        stderr,
        command: `${process.execPath} ${args.map((x) => JSON.stringify(x)).join(" ")}`
      });
    });
  });
}

function getHtml() {
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Zone Card Generator UI</title>
  <style>
    :root {
      --bg: #121214;
      --panel: #1d2128;
      --panel-2: #262d37;
      --text: #e7edf7;
      --muted: #93a0b2;
      --accent: #73c3ff;
      --danger: #ff8d8d;
      --ok: #8ef2a4;
      --border: #3a4655;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: linear-gradient(135deg, #0d0f13, #171b22 60%, #111721);
      color: var(--text);
      font-family: "Segoe UI", Tahoma, sans-serif;
    }
    .wrap {
      max-width: 1200px;
      margin: 0 auto;
      padding: 20px;
    }
    h1 {
      margin: 0 0 12px;
      font-size: 28px;
    }
    .sub {
      margin-bottom: 18px;
      color: var(--muted);
    }
    .grid {
      display: grid;
      grid-template-columns: 1.3fr 1fr;
      gap: 14px;
    }
    .panel {
      background: color-mix(in srgb, var(--panel) 90%, black);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 14px;
    }
    .panel h2 {
      margin: 0 0 10px;
      font-size: 18px;
    }
    label {
      display: block;
      margin-bottom: 8px;
      font-size: 13px;
      color: var(--muted);
    }
    input[type="text"], input[type="number"], select, textarea {
      width: 100%;
      border: 1px solid var(--border);
      border-radius: 8px;
      background: var(--panel-2);
      color: var(--text);
      padding: 8px 10px;
      font-size: 14px;
    }
    textarea {
      min-height: 140px;
      resize: vertical;
      line-height: 1.4;
    }
    .row {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
    }
    .actions {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      margin-top: 10px;
    }
    button {
      border: 1px solid var(--border);
      border-radius: 8px;
      background: #2d3a49;
      color: var(--text);
      padding: 8px 12px;
      font-weight: 600;
      cursor: pointer;
    }
    button.primary {
      background: linear-gradient(135deg, #2f7ed6, #46a9ff);
      color: #05192e;
      border-color: #77bdff;
    }
    button:disabled {
      opacity: 0.6;
      cursor: not-allowed;
    }
    .status {
      margin: 6px 0 0;
      font-size: 13px;
      color: var(--muted);
    }
    .status.ok { color: var(--ok); }
    .status.bad { color: var(--danger); }
    .zone-list {
      max-height: 520px;
      overflow: auto;
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 8px;
      background: #1a2028;
    }
    .biome {
      border: 1px solid #374252;
      border-radius: 8px;
      margin-bottom: 8px;
      padding: 8px;
      background: #202833;
    }
    .biome-head {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 6px;
      gap: 8px;
    }
    .biome-title {
      font-weight: 700;
      font-size: 14px;
      text-transform: capitalize;
    }
    .mini button {
      font-size: 12px;
      padding: 4px 8px;
    }
    .tile {
      display: grid;
      grid-template-columns: auto 1fr;
      gap: 8px;
      align-items: start;
      font-size: 13px;
      margin-bottom: 6px;
    }
    .tile code {
      color: #a7d6ff;
      font-size: 12px;
    }
    pre {
      margin: 0;
      white-space: pre-wrap;
      word-break: break-word;
      border: 1px solid var(--border);
      border-radius: 8px;
      background: #10161d;
      padding: 10px;
      max-height: 320px;
      overflow: auto;
    }
    @media (max-width: 980px) {
      .grid { grid-template-columns: 1fr; }
      .zone-list { max-height: 420px; }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <h1>Zone Card Generator UI</h1>
    <div class="sub">Runs locally at 127.0.0.1 and uses your server-side API key. The browser never needs your key.</div>
    <div id="keyStatus" class="status">Checking API key status...</div>
    <div class="grid">
      <section class="panel">
        <h2>Selection</h2>
        <label>
          Search zones
          <input id="filterInput" type="text" placeholder="Filter by name, id, or biome">
        </label>
        <label>
          <input id="allZonesToggle" type="checkbox"> Generate all zones (ignore manual selection)
        </label>
        <div class="actions mini">
          <button id="selectAllBtn" type="button">Select All Visible</button>
          <button id="clearAllBtn" type="button">Clear Selection</button>
        </div>
        <p id="selectionCount" class="status">Loading zones...</p>
        <div id="zoneList" class="zone-list"></div>
      </section>
      <section class="panel">
        <h2>Generation</h2>
        <label>
          Common style prompt (applies to all selected cards)
          <textarea id="styleInput" placeholder="Example: grimdark retro-futurist horror realism, damp concrete, atmospheric fog, painterly concept art"></textarea>
        </label>
        <div class="row">
          <label>Model
            <input id="modelInput" type="text" value="gpt-image-1">
          </label>
          <label>Size
            <select id="sizeInput">
              <option value="1024x1024">1024x1024</option>
              <option value="1536x1024" selected>1536x1024</option>
              <option value="1024x1536">1024x1536</option>
            </select>
          </label>
        </div>
        <div class="row">
          <label>Quality
            <select id="qualityInput">
              <option value="auto">auto</option>
              <option value="high" selected>high</option>
              <option value="medium">medium</option>
              <option value="low">low</option>
            </select>
          </label>
          <label>Delay (ms)
            <input id="delayInput" type="number" min="0" value="0">
          </label>
        </div>
        <label><input id="skipExistingInput" type="checkbox"> Skip existing files</label>
        <label><input id="dryRunInput" type="checkbox"> Dry run only (no API calls)</label>
        <div class="actions">
          <button id="generateBtn" class="primary" type="button">Generate</button>
        </div>
        <p id="runStatus" class="status"></p>
        <h2>Output</h2>
        <pre id="outputBox">No run yet.</pre>
      </section>
    </div>
  </div>
  <script>
    const state = {
      zones: [],
      selected: new Set()
    };

    const refs = {
      keyStatus: document.getElementById("keyStatus"),
      filterInput: document.getElementById("filterInput"),
      allZonesToggle: document.getElementById("allZonesToggle"),
      selectAllBtn: document.getElementById("selectAllBtn"),
      clearAllBtn: document.getElementById("clearAllBtn"),
      zoneList: document.getElementById("zoneList"),
      selectionCount: document.getElementById("selectionCount"),
      styleInput: document.getElementById("styleInput"),
      modelInput: document.getElementById("modelInput"),
      sizeInput: document.getElementById("sizeInput"),
      qualityInput: document.getElementById("qualityInput"),
      delayInput: document.getElementById("delayInput"),
      skipExistingInput: document.getElementById("skipExistingInput"),
      dryRunInput: document.getElementById("dryRunInput"),
      generateBtn: document.getElementById("generateBtn"),
      runStatus: document.getElementById("runStatus"),
      outputBox: document.getElementById("outputBox")
    };

    function matchesFilter(tile, filter) {
      if (!filter) return true;
      const text = (tile.id + " " + tile.name + " " + tile.biome).toLowerCase();
      return text.includes(filter);
    }

    function groupByBiome(tiles) {
      const map = new Map();
      for (const tile of tiles) {
        const biome = tile.biome || "unknown";
        if (!map.has(biome)) map.set(biome, []);
        map.get(biome).push(tile);
      }
      return [...map.entries()].sort((a, b) => a[0].localeCompare(b[0]));
    }

    function updateSelectionCount() {
      refs.selectionCount.textContent = "Selected zones: " + state.selected.size + " / " + state.zones.length;
    }

    function renderZoneList() {
      const filter = refs.filterInput.value.trim().toLowerCase();
      const visible = state.zones.filter((tile) => matchesFilter(tile, filter));
      const groups = groupByBiome(visible);
      refs.zoneList.innerHTML = "";

      for (const [biome, tiles] of groups) {
        const wrap = document.createElement("div");
        wrap.className = "biome";

        const head = document.createElement("div");
        head.className = "biome-head";
        const title = document.createElement("div");
        title.className = "biome-title";
        title.textContent = biome + " (" + tiles.length + ")";
        head.appendChild(title);

        const mini = document.createElement("div");
        mini.className = "mini";
        const selectBtn = document.createElement("button");
        selectBtn.type = "button";
        selectBtn.textContent = "Select";
        selectBtn.onclick = () => {
          for (const tile of tiles) state.selected.add(tile.id);
          renderZoneList();
          updateSelectionCount();
        };
        const clearBtn = document.createElement("button");
        clearBtn.type = "button";
        clearBtn.textContent = "Clear";
        clearBtn.onclick = () => {
          for (const tile of tiles) state.selected.delete(tile.id);
          renderZoneList();
          updateSelectionCount();
        };
        mini.appendChild(selectBtn);
        mini.appendChild(clearBtn);
        head.appendChild(mini);
        wrap.appendChild(head);

        for (const tile of tiles) {
          const row = document.createElement("label");
          row.className = "tile";
          const cb = document.createElement("input");
          cb.type = "checkbox";
          cb.checked = state.selected.has(tile.id);
          cb.onchange = () => {
            if (cb.checked) state.selected.add(tile.id);
            else state.selected.delete(tile.id);
            updateSelectionCount();
          };
          const info = document.createElement("div");
          const name = document.createElement("div");
          name.textContent = tile.name;
          const id = document.createElement("code");
          id.textContent = tile.id;
          info.appendChild(name);
          info.appendChild(id);
          row.appendChild(cb);
          row.appendChild(info);
          wrap.appendChild(row);
        }
        refs.zoneList.appendChild(wrap);
      }

      if (groups.length === 0) {
        refs.zoneList.textContent = "No zones match the current filter.";
      }
    }

    async function loadData() {
      const [healthRes, zonesRes] = await Promise.all([
        fetch("/api/health"),
        fetch("/api/zones")
      ]);
      const health = await healthRes.json();
      const zones = await zonesRes.json();
      state.zones = zones.tiles || [];
      state.selected = new Set(state.zones.map((z) => z.id));

      refs.keyStatus.textContent = health.keyConfigured
        ? "API key detected: ready to generate."
        : "No OPENAI_API_KEY found in env/.env.local/.env.";
      refs.keyStatus.className = "status " + (health.keyConfigured ? "ok" : "bad");

      renderZoneList();
      updateSelectionCount();
    }

    async function generate() {
      const style = refs.styleInput.value.trim();
      if (!style) {
        refs.runStatus.textContent = "Style prompt is required.";
        refs.runStatus.className = "status bad";
        return;
      }

      const all = refs.allZonesToggle.checked;
      const zoneIds = [...state.selected];
      if (!all && zoneIds.length === 0) {
        refs.runStatus.textContent = "Select at least one zone or enable Generate all zones.";
        refs.runStatus.className = "status bad";
        return;
      }

      const payload = {
        style,
        all,
        zoneIds,
        model: refs.modelInput.value.trim(),
        size: refs.sizeInput.value,
        quality: refs.qualityInput.value,
        delayMs: Number.parseInt(refs.delayInput.value || "0", 10),
        skipExisting: refs.skipExistingInput.checked,
        dryRun: refs.dryRunInput.checked
      };

      refs.generateBtn.disabled = true;
      refs.runStatus.textContent = "Generation running...";
      refs.runStatus.className = "status";
      refs.outputBox.textContent = "";

      try {
        const res = await fetch("/api/generate", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        });
        const result = await res.json();
        if (!res.ok || !result.ok) {
          refs.runStatus.textContent = "Generation failed.";
          refs.runStatus.className = "status bad";
        } else {
          refs.runStatus.textContent = "Generation completed.";
          refs.runStatus.className = "status ok";
        }

        refs.outputBox.textContent = [
          "Exit code: " + result.code,
          "Duration: " + result.durationMs + " ms",
          "",
          "Command:",
          result.command || "",
          "",
          "STDOUT:",
          result.stdout || "",
          "",
          "STDERR:",
          result.stderr || ""
        ].join("\\n");
      } catch (err) {
        refs.runStatus.textContent = "Request error: " + err.message;
        refs.runStatus.className = "status bad";
      } finally {
        refs.generateBtn.disabled = false;
      }
    }

    refs.filterInput.addEventListener("input", renderZoneList);
    refs.selectAllBtn.addEventListener("click", () => {
      const filter = refs.filterInput.value.trim().toLowerCase();
      for (const tile of state.zones) {
        if (matchesFilter(tile, filter)) state.selected.add(tile.id);
      }
      renderZoneList();
      updateSelectionCount();
    });
    refs.clearAllBtn.addEventListener("click", () => {
      state.selected.clear();
      renderZoneList();
      updateSelectionCount();
    });
    refs.generateBtn.addEventListener("click", generate);

    loadData().catch((err) => {
      refs.keyStatus.textContent = "Failed to load UI data: " + err.message;
      refs.keyStatus.className = "status bad";
    });
  </script>
</body>
</html>`;
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url || "/", `http://${HOST}:${PORT}`);

    if (req.method === "GET" && url.pathname === "/") {
      sendHtml(res, 200, getHtml());
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/health") {
      const keyConfigured = await hasApiKeyConfigured();
      sendJson(res, 200, { ok: true, keyConfigured });
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/zones") {
      const tiles = await loadZones();
      sendJson(res, 200, { ok: true, tiles });
      return;
    }

    if (req.method === "POST" && url.pathname === "/api/generate") {
      if (generationInProgress) {
        sendJson(res, 409, { ok: false, error: "Generation already in progress." });
        return;
      }

      const body = await parseRequestBody(req);
      const style = String(body.style || "").trim();
      const all = Boolean(body.all);
      const zoneIds = Array.isArray(body.zoneIds) ? body.zoneIds.map((id) => String(id)) : [];
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

      generationInProgress = true;
      const start = Date.now();
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
      generationInProgress = false;

      sendJson(res, result.code === 0 ? 200 : 500, {
        ok: result.code === 0,
        code: result.code,
        stdout: result.stdout,
        stderr: result.stderr,
        command: result.command,
        durationMs: Date.now() - start
      });
      return;
    }

    sendJson(res, 404, { ok: false, error: "Not found" });
  } catch (err) {
    generationInProgress = false;
    sendJson(res, 500, {
      ok: false,
      error: err instanceof Error ? err.message : String(err)
    });
  }
});

server.listen(PORT, HOST, () => {
  console.log(`Zone Card Generator UI running at http://${HOST}:${PORT}`);
  console.log("Press Ctrl+C to stop.");
});
