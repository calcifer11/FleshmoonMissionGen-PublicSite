#!/usr/bin/env node
import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";

const DEFAULT_ZONE_JSON = "data/zones.json";
const DEFAULT_OUTPUT_ROOT = "ZoneCards";
const DEFAULT_MODEL = "gpt-image-1";
const DEFAULT_SIZE = "1536x1024";
const DEFAULT_QUALITY = "high";
const API_URL = "https://api.openai.com/v1/images/generations";

function printHelp() {
  console.log(`
Usage:
  node scripts/generate-zone-cards-openai.mjs [options]

Required:
  --style "<text>"         Common style prompt applied to every selected card
  OR
  --style-file <path>      File containing the common style prompt

Selection (pick at least one unless using --all):
  --zones <id,id,...>      Generate only specific zone tile IDs
  --zones-file <path>      File with one zone tile ID per line
  --biomes <b1,b2,...>     Generate only tiles in these biomes
  --all                    Generate all zones in data/zones.json (explicit opt-in)

Other:
  --zone-json <path>       Zones file path (default: data/zones.json)
  --output-root <path>     ZoneCards root folder (default: ZoneCards)
  --model <name>           Image model (default: gpt-image-1)
  --size <WxH>             Image size (default: 1536x1024)
  --quality <q>            low | medium | high | auto (default: high)
  --skip-existing          Skip tiles whose output file already exists
  --delay-ms <n>           Delay between API calls (default: 0)
  --dry-run                Print plan only, no API calls
  --help                   Show this help

Examples:
  node scripts/generate-zone-cards-openai.mjs \\
    --zones PumpControl,OverflowJunction \\
    --style "grimdark, wet concrete, practical lighting, painterly concept art"

  node scripts/generate-zone-cards-openai.mjs \\
    --biomes labs,transit \\
    --style-file scripts/style-prompts/base-style.txt \\
    --skip-existing

  node scripts/generate-zone-cards-openai.mjs --all --style "retro-futurist horror realism"
`);
}

function parseArgs(argv) {
  const args = {
    zonesCsv: "",
    zonesFile: "",
    biomesCsv: "",
    all: false,
    style: "",
    styleFile: "",
    zoneJson: DEFAULT_ZONE_JSON,
    outputRoot: DEFAULT_OUTPUT_ROOT,
    model: DEFAULT_MODEL,
    size: DEFAULT_SIZE,
    quality: DEFAULT_QUALITY,
    skipExisting: false,
    delayMs: 0,
    dryRun: false,
    help: false
  };

  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (!token.startsWith("--")) {
      throw new Error(`Unexpected argument: ${token}`);
    }

    const [flag, inlineValue] = token.split("=", 2);
    const readValue = () => {
      if (inlineValue !== undefined) return inlineValue;
      const next = argv[i + 1];
      if (!next || next.startsWith("--")) {
        throw new Error(`Missing value for ${flag}`);
      }
      i += 1;
      return next;
    };

    switch (flag) {
      case "--zones":
        args.zonesCsv = readValue();
        break;
      case "--zones-file":
        args.zonesFile = readValue();
        break;
      case "--biomes":
        args.biomesCsv = readValue();
        break;
      case "--all":
        args.all = true;
        break;
      case "--style":
        args.style = readValue();
        break;
      case "--style-file":
        args.styleFile = readValue();
        break;
      case "--zone-json":
        args.zoneJson = readValue();
        break;
      case "--output-root":
        args.outputRoot = readValue();
        break;
      case "--model":
        args.model = readValue();
        break;
      case "--size":
        args.size = readValue();
        break;
      case "--quality":
        args.quality = readValue();
        break;
      case "--skip-existing":
        args.skipExisting = true;
        break;
      case "--delay-ms":
        args.delayMs = Number.parseInt(readValue(), 10);
        if (!Number.isFinite(args.delayMs) || args.delayMs < 0) {
          throw new Error("--delay-ms must be a non-negative integer");
        }
        break;
      case "--dry-run":
        args.dryRun = true;
        break;
      case "--help":
        args.help = true;
        break;
      default:
        throw new Error(`Unknown flag: ${flag}`);
    }
  }

  return args;
}

function parseCsvSet(value) {
  const out = new Set();
  for (const part of String(value || "").split(",")) {
    const trimmed = part.trim();
    if (trimmed) out.add(trimmed);
  }
  return out;
}

function normalizeAssetPath(input) {
  return String(input || "").replaceAll("\\", "/").trim();
}

function folderNameForBiome(biome) {
  switch (String(biome || "").toLowerCase()) {
    case "sewers":
      return "Sewers";
    case "industrial":
      return "Urban-Industrial";
    case "transit":
      return "Transit";
    case "civic":
      return "Civic";
    case "labs":
      return "Labs";
    case "outskirts":
      return "Outskirts";
    default:
      return "Misc";
  }
}

function buildTilePrompt(tile, stylePrompt) {
  const tags = Array.isArray(tile.tags) ? tile.tags.join(", ") : "";
  const hints = Array.isArray(tile.functionalRoleHints) ? tile.functionalRoleHints.join(", ") : "";
  const intensity = tile.intensityLevel ?? "unknown";

  return [
    "Create a landscape environment illustration for a tactical horror board game zone card background.",
    `GLOBAL STYLE (must be consistent across this batch): ${stylePrompt}`,
    "Tile-specific direction:",
    `- Name: ${tile.name || tile.id}`,
    `- ID: ${tile.id}`,
    `- Biome: ${tile.biome || "unknown"}`,
    `- Tags: ${tags || "none"}`,
    `- Role hints: ${hints || "none"}`,
    `- Intensity: ${intensity}`,
    "Constraints:",
    "- No text, letters, numbers, logos, or watermarks.",
    "- No card frame or UI overlay.",
    "- Prioritize architecture/environment storytelling over character portraiture.",
    "- Cinematic lighting and readable composition."
  ].join("\n");
}

async function readTextUtf8(filePath) {
  const raw = await fs.readFile(filePath, "utf8");
  return raw.replace(/^\uFEFF/, "");
}

async function loadZoneIdsFromFile(filePath) {
  const text = await readTextUtf8(filePath);
  const ids = new Set();
  for (const line of text.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    ids.add(trimmed);
  }
  return ids;
}

async function loadApiKey() {
  const envKey = String(process.env.OPENAI_API_KEY || "").trim();
  if (envKey) return envKey;

  const filesToCheck = [".env.local", ".env"];
  for (const name of filesToCheck) {
    const filePath = path.resolve(name);
    try {
      const text = await readTextUtf8(filePath);
      const lines = text.split(/\r?\n/);
      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed || trimmed.startsWith("#")) continue;
        const match = trimmed.match(/^OPENAI_API_KEY\s*=\s*(.+)$/);
        if (!match) continue;
        const value = match[1].trim().replace(/^['"]|['"]$/g, "");
        if (value) return value;
      }
    } catch {
      // Ignore missing/invalid env file and continue.
    }
  }

  return "";
}

async function requestImageBuffer({ apiKey, model, prompt, size, quality }) {
  const payload = {
    model,
    prompt,
    size,
    quality,
    output_format: "png"
  };

  const response = await fetch(API_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`
    },
    body: JSON.stringify(payload)
  });

  const bodyText = await response.text();
  let body = null;
  try {
    body = JSON.parse(bodyText);
  } catch {
    body = null;
  }

  if (!response.ok) {
    const detail = body?.error?.message || bodyText || `HTTP ${response.status}`;
    throw new Error(`Image API request failed: ${detail}`);
  }

  const item = body?.data?.[0];
  if (!item) {
    throw new Error("Image API returned no image data");
  }

  if (item.b64_json) {
    return Buffer.from(item.b64_json, "base64");
  }

  if (item.url) {
    const imgRes = await fetch(item.url);
    if (!imgRes.ok) {
      throw new Error(`Failed to download generated image URL: HTTP ${imgRes.status}`);
    }
    const arr = await imgRes.arrayBuffer();
    return Buffer.from(arr);
  }

  throw new Error("Image API response did not include b64_json or url");
}

async function sleep(ms) {
  if (ms <= 0) return;
  await new Promise((resolve) => setTimeout(resolve, ms));
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    printHelp();
    return;
  }

  let stylePrompt = args.style.trim();
  if (!stylePrompt && args.styleFile) {
    stylePrompt = (await readTextUtf8(path.resolve(args.styleFile))).trim();
  }
  if (!stylePrompt) {
    throw new Error("Provide --style or --style-file");
  }

  const explicitZoneIds = parseCsvSet(args.zonesCsv);
  if (args.zonesFile) {
    const fileZoneIds = await loadZoneIdsFromFile(path.resolve(args.zonesFile));
    for (const id of fileZoneIds) explicitZoneIds.add(id);
  }
  const biomeFilters = parseCsvSet(args.biomesCsv);

  if (!args.all && explicitZoneIds.size === 0 && biomeFilters.size === 0) {
    throw new Error("Select scope with --zones, --zones-file, --biomes, or explicitly use --all");
  }

  const zoneJsonPath = path.resolve(args.zoneJson);
  const outputRoot = path.resolve(args.outputRoot);
  const zoneRawText = await readTextUtf8(zoneJsonPath);
  const zoneDoc = JSON.parse(zoneRawText);
  const tiles = Array.isArray(zoneDoc?.tiles) ? zoneDoc.tiles : null;
  if (!tiles) {
    throw new Error(`Invalid zone JSON shape in ${zoneJsonPath}: expected { "tiles": [] }`);
  }

  const tileById = new Map(tiles.map((tile) => [String(tile.id), tile]));
  const unknownIds = [...explicitZoneIds].filter((id) => !tileById.has(id));
  if (unknownIds.length > 0) {
    throw new Error(`Unknown zone tile id(s): ${unknownIds.join(", ")}`);
  }

  let selected = tiles.slice();
  if (explicitZoneIds.size > 0) {
    selected = selected.filter((tile) => explicitZoneIds.has(String(tile.id)));
  }
  if (biomeFilters.size > 0) {
    selected = selected.filter((tile) => biomeFilters.has(String(tile.biome)));
  }

  if (selected.length === 0) {
    throw new Error("No zone tiles matched the provided selection filters");
  }

  let zoneJsonChanged = false;
  const generationPlan = selected.map((tile) => {
    const existingPath = normalizeAssetPath(tile.image || "");
    const relativePath = existingPath
      ? (existingPath.startsWith("ZoneCards/") ? existingPath.slice("ZoneCards/".length) : existingPath)
      : `${folderNameForBiome(tile.biome)}/${tile.id}.png`;

    if (!existingPath) {
      tile.image = `ZoneCards/${relativePath}`;
      zoneJsonChanged = true;
    }

    return {
      tile,
      assetPath: `ZoneCards/${relativePath}`.replaceAll("\\", "/"),
      outFile: path.join(outputRoot, relativePath.split("/").join(path.sep))
    };
  });

  console.log(`Tiles selected: ${generationPlan.length}`);
  console.log(`Model: ${args.model}`);
  console.log(`Size: ${args.size}`);
  console.log(`Quality: ${args.quality}`);
  console.log(`Skip existing files: ${args.skipExisting ? "yes" : "no"}`);
  console.log(`Dry run: ${args.dryRun ? "yes" : "no"}`);

  if (args.dryRun) {
    for (const entry of generationPlan) {
      console.log(`- ${entry.tile.id} -> ${entry.assetPath}`);
    }
    if (zoneJsonChanged) {
      console.log("Note: zones.json would be updated with missing image paths.");
    }
    return;
  }

  const apiKey = await loadApiKey();
  if (!apiKey) {
    throw new Error("OPENAI_API_KEY not found in environment, .env.local, or .env");
  }

  let generated = 0;
  let skipped = 0;
  const failed = [];

  for (let i = 0; i < generationPlan.length; i += 1) {
    const { tile, outFile, assetPath } = generationPlan[i];
    process.stdout.write(`[${i + 1}/${generationPlan.length}] ${tile.id} ... `);
    try {
      try {
        await fs.access(outFile);
        if (args.skipExisting) {
          skipped += 1;
          process.stdout.write("skipped (exists)\n");
          continue;
        }
      } catch {
        // File is missing, continue to generation.
      }

      const prompt = buildTilePrompt(tile, stylePrompt);
      const imageBuffer = await requestImageBuffer({
        apiKey,
        model: args.model,
        prompt,
        size: args.size,
        quality: args.quality
      });

      await fs.mkdir(path.dirname(outFile), { recursive: true });
      await fs.writeFile(outFile, imageBuffer);
      generated += 1;
      process.stdout.write(`saved -> ${assetPath}\n`);
      await sleep(args.delayMs);
    } catch (err) {
      failed.push({ id: tile.id, error: err instanceof Error ? err.message : String(err) });
      process.stdout.write("failed\n");
    }
  }

  if (zoneJsonChanged) {
    const serialized = `${JSON.stringify(zoneDoc, null, 2)}\n`;
    await fs.writeFile(zoneJsonPath, serialized, "utf8");
    console.log(`Updated ${path.relative(process.cwd(), zoneJsonPath)} with missing image paths.`);
  }

  console.log("");
  console.log(`Generated: ${generated}`);
  console.log(`Skipped: ${skipped}`);
  console.log(`Failed: ${failed.length}`);
  if (failed.length > 0) {
    for (const item of failed) {
      console.log(`- ${item.id}: ${item.error}`);
    }
    process.exitCode = 1;
  }
}

main().catch((err) => {
  const message = err instanceof Error ? err.message : String(err);
  console.error(`Error: ${message}`);
  process.exit(1);
});
