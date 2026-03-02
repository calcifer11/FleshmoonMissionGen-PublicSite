# Fleshmoon / TYR-113 Mission Concept Generator

A static, local mission generator for tactical horror board game concepts.

## Run

1. Open `index.html` in your browser (double-click works in many browsers).
2. If your browser blocks local JSON loading from `file://`, run a tiny local server instead, for example:
   - `python3 -m http.server`
   - then open `http://localhost:8000`

## Features

- Generates mission concepts across 10 categories.
- Generates a thematic operation codename based on mission vibe/tags.
- Seeded deterministic generation (`same seed => same picks`).
- Weighted option picks (`weight` in JSON, default is `1`).
- Per-category reroll buttons.
- Dedicated `Reroll Name Only` button that ignores seed once for codename variety.
- Per-category lock toggles (reroll-all respects locks).
- Copy formatted mission text to clipboard.
- Export current mission object as `.json`.
- Local history panel (last 20 missions via `localStorage`).
- Optional tag filters to bias picks.
- Optional unique-within-mission reroll mode.
- Print-friendly CSS.
- Sparse-canvas action-area map generator on a fixed `7x9` canvas.
- Action areas are contiguous templates with gameplay-aware sizes (Tiny/Small/Medium/Large pacing mix) and intentional negative space.
- Occupancy target band is gameplay-scaled (`14-20` occupied tiles, hard bounds `10-24`).
- Area graph templates (`spine`, `spine_branch`, `loop`, `hub`, `fork_reconverge`, `gauntlet_pocket`) are generated first, then spatialized.
- Inter-area links are created by adjacency or short corridors (1-2 cells), with one choke gate per graph edge.
- Map cells render real tile art from `ZoneCards/.../*.png` (with deterministic biome fallbacks when a tile has no explicit image).
- Per-edge gate states (`open`, `locked`, `blocked`) with click-to-cycle and default 70/20/10 weighting on non-bridge edges.
- Area legend includes role, tile count, size category, cards returned, dominant tile, and supporting tiles.
- Roles include `START`, `OBJECTIVE`, `PRESSURE`, and optional `EXTRACTION`.
- Tile composition is gameplay/story aware: dominant + supporting mix, phase progression, and depth gradients.
- Map controls: reroll entire map, reroll area themes, reroll area shapes, area size bias, variation toggle, mixed biome toggle.
- Mission export includes full map payload (grid cells, action area IDs/metadata, and gate states).

Map export payload includes:

- `cells[]`: `tileId`, `tileName`, `actionAreaId`, row/col/index, corridor + variation flags.
- `grid[][]`: sparse grid with empty cells preserved and per-cell `tileId` / `actionAreaId`.
- `actionAreas[]`: role, tile count, size category, cards returned, shape type, dominant tile, supporting tiles, biome/tags/cells, role flags (`start`, `objective`, `pressure`, `extraction`).
- `corridors[]`: corridor ids, linked areas, tile metadata, and corridor cell indexes.
- `graph`: area graph template + edge pairs.
- `gates[]`: inter-area edge links and current gate state (`open`, `locked`, `blocked`).
- top-level `templateId`, `seed`, and `totalTiles`.

## Edit Options

All option tables live in `data/options.json`.
Zone tile themes for map generation live in `data/zones.json`.

Data shape:

```json
{
  "locations": [{ "text": "...", "weight": 1, "tags": ["indoor"] }],
  "corruptions": [{ "text": "...", "weight": 1, "tags": ["stricken"] }],
  "twists": [{ "text": "...", "weight": 1, "tags": ["machine"] }],
  "escalations": [{ "text": "...", "weight": 1, "tags": ["timer"] }],
  "primaryObjectives": [{ "text": "...", "weight": 1, "tags": ["objective"] }],
  "secondaryObjectives": [{ "text": "...", "weight": 1, "tags": ["objective"] }],
  "enemyMixHints": [{ "text": "...", "weight": 1, "tags": ["combat"] }],
  "mapNotes": [{ "text": "...", "weight": 1, "tags": ["map"] }],
  "rewardsConsequences": [{ "text": "...", "weight": 1, "tags": ["campaign"] }],
  "toneTags": [{ "text": "claustrophobic", "weight": 1, "tags": ["tone"] }],
  "operationCodenamePrefixes": [{ "text": "Iron", "weight": 1, "tags": ["machine"] }],
  "operationCodenameNouns": [{ "text": "Signal", "weight": 1, "tags": ["machine"] }],
  "operationCodenameSuffixes": [{ "text": "Protocol", "weight": 1, "tags": ["machine"] }]
}
```

Notes:

- Keep each option to one sentence for readability.
- `weight` is optional; missing values are treated as `1`.
- `tags` are optional, but useful with UI filters.
- Category keys must match exactly.
- Codename is generated as `Operation <Prefix> <Noun>` (or variant styles), with picks biased by mission tags and optional active filters.

## Edit Zone Tiles

`data/zones.json` controls action-area tile themes.

Shape:

```json
{
  "tiles": [
    {
      "id": "FloodedSewers",
      "name": "Flooded Sewers",
      "biome": "sewers",
      "weight": 4,
      "tags": ["indoor", "hazard", "sewer"],
      "image": "ZoneCards/Sewers/FloodedSewers.png"
    }
  ]
}
```

Guidance:

- Use stable `id` values (exported map data references tile IDs).
- `biome` drives mixed-vs-single-biome area theming.
- Add tags like `objective`, `spawn`, `machine` for objective-area bias.
- Keep weights higher for common tiles and lower for rare/special tiles.

## PoC Mode

`app.js` includes `POC_MODE = true` for limited tile pools.

In PoC Mode, area tile selection (dominant + supporting + corridor picks) uses relaxed fallback order:

1. required tags + biome
2. biome only
3. any tile

The generator will always choose a tile and continue map generation even with sparse coverage.
