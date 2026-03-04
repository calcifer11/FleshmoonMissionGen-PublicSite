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
const GOOGLE_DRIVE_BACKUP_ROOT = String(
  process.env.ZONECARD_BACKUP_ROOT
  || "/Users/weaver/Library/CloudStorage/GoogleDrive-hubbabubba.awesome@gmail.com/My Drive/FLESHMOON/2.0/ZoneCards"
).trim();
const BIOME_KEY_ALIASES = {
  civic: "residential",
  labs: "research",
  outskirts: "perimeter"
};
const DEFAULT_PROMPTS = {
  sewers: "Painterly realism, cinematic environmental concept art for a tactical horror board game zone card, horizontal 7:12 tarot aspect ratio, grounded subterranean infrastructure aesthetic, damp concrete and aged metal materials, wet reflective surfaces, utilitarian construction, low industrial lighting tones, subtle humidity/mist in the air, moody but not fantastical, no people, no characters, no text, no UI, no borders, no logos, no readable signage, no extreme fisheye lens.",
  industrial: "Painterly realism, cinematic environmental concept art for a tactical horror board game zone card, horizontal 7:12 tarot aspect ratio, heavy mechanical architecture, exposed steel structures, large-scale industrial forms, strong geometric silhouettes, cold industrial lighting with occasional warm mechanical glow, oil-stained worn surfaces, utilitarian machine-district design language, no people, no characters, no text, no UI, no borders, no logos, no readable signage, no extreme fisheye lens.",
  transit: "Painterly realism, cinematic environmental concept art for a tactical horror board game zone card, horizontal 7:12 tarot aspect ratio, urban transit infrastructure aesthetic, concrete/steel/tile material language, linear civic engineering forms, grounded metropolitan atmosphere, practical city lighting, subtle dust in air, tense but not occult-decorated, no people, no characters, no text, no UI, no borders, no logos, no readable signage, no extreme fisheye lens.",
  residential: "Painterly realism, cinematic environmental concept art for a tactical horror board game zone card, horizontal 7:12 tarot aspect ratio, lived-in urban residential/civic architecture, human-scale spaces, domestic materials (drywall, brick, worn tile, scuffed paint, carpet), practical interior lighting or soft window light, grounded atmosphere of abandonment, NO arcane symbols/occult motifs unless explicitly specified by the zone name/description, no people, no characters, no text, no UI, no borders, no logos, no readable signage, no extreme fisheye lens.",
  research: "Painterly realism, cinematic environmental concept art for a tactical horror board game zone card, horizontal 7:12 tarot aspect ratio, sterile high-security research facility aesthetic, clean institutional materials (sealed panels, reinforced surfaces, glass when appropriate), controlled lighting in cool whites/blues with subtle red emergency accents, clinical and structured tone without decorative symbolism, no people, no characters, no text, no UI, no borders, no logos, no readable signage, no extreme fisheye lens.",
  perimeter: "Painterly realism, cinematic environmental concept art for a tactical horror board game zone card, horizontal 7:12 tarot aspect ratio, outdoor fortified containment aesthetic, concrete barriers/fencing/asphalt/gravel material language, open spatial feel, cold floodlighting or harsh environmental lighting, wind/haze/dust in atmosphere, grounded military quarantine tone, no people, no characters, no text, no UI, no borders, no logos, no readable signage, no extreme fisheye lens."
};

function printHelp() {
  console.log(`
Usage:
  node scripts/generate-zone-cards-openai.mjs [options]

Prompting:
  --extra-description "<text>"  Optional flavor text appended after each zone name
  --style "<text>"              Backward-compatible alias of --extra-description
  --style-file <path>           File-based alias of --extra-description

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
    --extra-description "storm runoff, emergency lights, recent evacuation"

  node scripts/generate-zone-cards-openai.mjs \\
    --biomes research,transit \\
    --extra-description "failing power grid, tense atmosphere" \\
    --skip-existing

  node scripts/generate-zone-cards-openai.mjs --all --extra-description "late-stage containment collapse"
`);
}

function parseArgs(argv) {
  const args = {
    zonesCsv: "",
    zonesFile: "",
    biomesCsv: "",
    all: false,
    extraDescription: "",
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
      case "--extra-description":
        args.extraDescription = readValue();
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

function normalizeBiomeKey(value) {
  const key = String(value || "unknown").trim().toLowerCase() || "unknown";
  return BIOME_KEY_ALIASES[key] || key;
}

function normalizeAssetPath(input) {
  return String(input || "").replaceAll("\\", "/").trim();
}

function escapeRegex(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function safeZoneCardBaseName(tile) {
  const source = String(tile?.id || tile?.name || "ZoneCard").trim();
  return source
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/\s+/g, "_")
    .replace(/^_+|_+$/g, "") || "ZoneCard";
}

async function backupExistingZoneCardVersion({ tile, outFile }) {
  const biomeFolder = folderNameForBiome(tile?.biome);
  const oldDir = path.join(GOOGLE_DRIVE_BACKUP_ROOT, biomeFolder, "old");
  const ext = path.extname(outFile) || ".png";
  const baseName = safeZoneCardBaseName(tile);

  await fs.mkdir(oldDir, { recursive: true });

  let maxVersion = 0;
  const entries = await fs.readdir(oldDir, { withFileTypes: true });
  const fileRegex = new RegExp(`^${escapeRegex(baseName)}_old_(\\d+)${escapeRegex(ext)}$`);
  for (const entry of entries) {
    if (!entry.isFile()) continue;
    const match = entry.name.match(fileRegex);
    if (!match) continue;
    const version = Number.parseInt(match[1], 10);
    if (Number.isFinite(version) && version > maxVersion) maxVersion = version;
  }

  const nextVersion = maxVersion + 1;
  const backupFile = path.join(oldDir, `${baseName}_old_${nextVersion}${ext}`);
  await fs.copyFile(outFile, backupFile);
  return backupFile;
}

function folderNameForBiome(biome) {
  switch (normalizeBiomeKey(biome)) {
    case "sewers":
      return "Sewers";
    case "industrial":
      return "Urban-Industrial";
    case "transit":
      return "Transit";
    case "residential":
      return "Civic";
    case "research":
      return "Labs";
    case "perimeter":
      return "Outskirts";
    default:
      return "Misc";
  }
}

function buildTilePrompt(tile, userExtraDescription) {
  const biomeKey = normalizeBiomeKey(tile.biome);
  const basePrompt = DEFAULT_PROMPTS[biomeKey] || DEFAULT_PROMPTS.industrial;
  const zoneName = String(tile.name || tile.id || "Unknown Zone").trim();
  const promptHint = String(tile.promptHint || "").trim();
  return `${basePrompt}\n\nSpecific scene: ${zoneName}${promptHint ? ` — ${promptHint}` : ""}${userExtraDescription ? ` — ${userExtraDescription}` : ""}`;
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

  let userExtraDescription = args.extraDescription.trim();
  if (!userExtraDescription) userExtraDescription = args.style.trim();
  if (!userExtraDescription && args.styleFile) {
    userExtraDescription = (await readTextUtf8(path.resolve(args.styleFile))).trim();
  }

  const explicitZoneIds = parseCsvSet(args.zonesCsv);
  if (args.zonesFile) {
    const fileZoneIds = await loadZoneIdsFromFile(path.resolve(args.zonesFile));
    for (const id of fileZoneIds) explicitZoneIds.add(id);
  }
  const biomeFilters = parseCsvSet(args.biomesCsv);
  const normalizedBiomeFilters = new Set([...biomeFilters].map((biome) => normalizeBiomeKey(biome)));

  if (!args.all && explicitZoneIds.size === 0 && normalizedBiomeFilters.size === 0) {
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
  if (normalizedBiomeFilters.size > 0) {
    selected = selected.filter((tile) => normalizedBiomeFilters.has(normalizeBiomeKey(tile.biome)));
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
      let existsAlready = false;
      try {
        await fs.access(outFile);
        existsAlready = true;
        if (args.skipExisting) {
          skipped += 1;
          process.stdout.write("skipped (exists)\n");
          continue;
        }
      } catch {
        // File is missing, continue to generation.
      }

      const prompt = buildTilePrompt(tile, userExtraDescription);
      const imageBuffer = await requestImageBuffer({
        apiKey,
        model: args.model,
        prompt,
        size: args.size,
        quality: args.quality
      });

      if (existsAlready) {
        const backupFile = await backupExistingZoneCardVersion({ tile, outFile });
        process.stdout.write(`backup -> ${normalizeAssetPath(backupFile)}\n`);
      }

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
