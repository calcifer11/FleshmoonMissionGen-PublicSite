const CATEGORY_CONFIG = [
  { field: "location", source: "locations", label: "Location", outputLabel: "Location" },
  { field: "corruption", source: "corruptions", label: "Corruption Type", outputLabel: "Corruption" },
  { field: "twist", source: "twists", label: "Hidden Truth / Twist", outputLabel: "Hidden Truth" },
  { field: "escalation", source: "escalations", label: "Escalation / Timer Pressure", outputLabel: "Escalation" },
  { field: "primaryObjective", source: "primaryObjectives", label: "Primary Objective", outputLabel: "Primary Objective" },
  { field: "secondaryObjective", source: "secondaryObjectives", label: "Secondary Objective", outputLabel: "Secondary Objective" },
  { field: "enemyMixHint", source: "enemyMixHints", label: "Enemy Mix Hint", outputLabel: "Enemy Mix Hint" },
  { field: "mapNotes", source: "mapNotes", label: "Map Notes", outputLabel: "Map Notes" },
  { field: "rewardConsequence", source: "rewardsConsequences", label: "Reward / Consequence", outputLabel: "Reward/Consequence" },
  { field: "toneTag", source: "toneTags", label: "Tone Tag", outputLabel: "Tone" },
  { field: "operationCodename", source: null, label: "Operation Codename", outputLabel: "Operation Codename" }
];

const MAP_PRESETS = {
  "7x9": { key: "7x9", rows: 7, cols: 9, areaMin: 4, areaMax: 6 }
};

const POC_MODE = true;
const GATE_STATES = ["open", "locked", "blocked"];
const AREA_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
const HISTORY_KEY = "fleshmoon_mission_history_v1";
const MAX_HISTORY = 20;
const ZONECARD_IMAGE_LIBRARY = {
  sewers: [
    "ZoneCards/Sewers/FloodedSewers.png",
    "ZoneCards/Sewers/InfestedTunnels.png",
    "ZoneCards/Sewers/SewerBarricades.png",
    "ZoneCards/Sewers/SewerHideout.png",
    "ZoneCards/Sewers/SewerTunnels.png",
    "ZoneCards/Sewers/SewerWaterfall.png",
    "ZoneCards/Sewers/SpawningPools.png"
  ],
  urbanIndustrial: [
    "ZoneCards/Urban-Industrial/Alleyway.png",
    "ZoneCards/Urban-Industrial/ControlRoom.png",
    "ZoneCards/Urban-Industrial/Dumpsters.png",
    "ZoneCards/Urban-Industrial/FiltrationNode.png",
    "ZoneCards/Urban-Industrial/Industrial Walkway.png",
    "ZoneCards/Urban-Industrial/ProcessingRoom.png",
    "ZoneCards/Urban-Industrial/RuinedRooms.png",
    "ZoneCards/Urban-Industrial/TrashHeaps.png"
  ]
};
const BIOME_ART_THEME = {
  sewers: "sewers",
  industrial: "urbanIndustrial",
  transit: "urbanIndustrial",
  civic: "urbanIndustrial",
  labs: "urbanIndustrial",
  outskirts: "urbanIndustrial",
  unknown: "urbanIndustrial"
};
const ZONECARD_IMAGE_PATHS = Object.values(ZONECARD_IMAGE_LIBRARY).flat();
const ZONECARD_IMAGE_TOKEN_MAP = ZONECARD_IMAGE_PATHS.reduce((map, path) => {
  const fileName = path.split("/").pop() || "";
  const baseName = fileName.replace(/\.png$/i, "");
  map.set(toLookupToken(baseName), path);
  return map;
}, new Map());
const SPARSE_OCCUPANCY_MIN = 0.16;
const SPARSE_OCCUPANCY_MAX = 0.36;
const MAP_ATTEMPT_LIMIT = 40;
const SPARSE_PLAN_ATTEMPT_LIMIT = 64;
const CENTER_PLACEMENT_ATTEMPTS = 72;
const TARGET_OCCUPIED_MIN = 14;
const TARGET_OCCUPIED_MAX = 20;
const HARD_OCCUPIED_MIN = 10;
const HARD_OCCUPIED_MAX = 24;
const MAX_CORRIDOR_AREA_RATIO = 0.3;
const MIN_START_OBJECTIVE_DISTANCE = 4;
const QUALITY_ACCEPT_SCORE = 50;
const QUALITY_RELAXED_ACCEPT_SCORE = 42;
const QUALITY_MAX_ATTEMPTS = 40;
const TILE_SHARE_SOFT_CAP = 0.35;
const TILE_SHARE_FALLBACK_CAP = 0.5;
const ENABLE_EXTRACTION_ROLE = true;
const GRAPH_TEMPLATE_POOL = [
  { id: "spine", weight: 25 },
  { id: "spine_branch", weight: 25 },
  { id: "loop", weight: 15 },
  { id: "hub", weight: 15 },
  { id: "fork_reconverge", weight: 10 },
  { id: "gauntlet_pocket", weight: 10 }
];
const ROOM_SHAPE_ARCHETYPES = ["rectangle", "l_shape", "t_shape", "blob", "wide_hall", "courtyard"];
const CORE_ROOM_SHAPE_ARCHETYPES = ["rectangle", "l_shape", "t_shape", "blob"];
const CORRIDOR_SHAPE_ARCHETYPES = ["straight_corridor", "bent_corridor", "short_choke"];
const PHASE_TAG_WEIGHTS = {
  neutral: ["tunnels", "walkway", "alleyway", "ruined", "urban", "transit", "indoor", "outdoor"],
  shift: ["flooded", "infested", "hazard", "spawn", "sewer", "trash", "machine", "stricken", "anomaly"],
  claimed: ["barricades", "control", "objective", "spawn", "machine", "cult", "ritual", "anomaly", "fortified"]
};
const DEPTH_PHASE_MAP = {
  entrance: "neutral",
  middle: "shift",
  deep: "claimed"
};
const EDITABLE_TILE_FIELDS = [
  "tileId",
  "tileName",
  "tileImage",
  "biome",
  "isVariation",
  "dominantTileId"
];
const CUSTOM_ROOM_NAMES_KEY = "fleshmoon_custom_room_names_v1";
const HIDDEN_ROOM_NAMES_KEY = "fleshmoon_hidden_room_names_v1";

const state = {
  options: null,
  zones: null,
  mission: null,
  map: null,
  mapNonce: 0,
  lockState: {},
  history: [],
  selectedTags: new Set(),
  roomLibrary: {
    customTiles: [],
    hiddenTileIds: new Set()
  },
  mapUi: {
    dragSourceIndex: null,
    menuTargetIndex: null
  }
};

const refs = {
  seedInput: document.getElementById("seedInput"),
  uniqueToggle: document.getElementById("uniqueToggle"),
  status: document.getElementById("status"),
  generateBtn: document.getElementById("generateBtn"),
  rerollAllBtn: document.getElementById("rerollAllBtn"),
  rerollNameBtn: document.getElementById("rerollNameBtn"),
  copyBtn: document.getElementById("copyBtn"),
  exportBtn: document.getElementById("exportBtn"),
  missionHeadline: document.getElementById("missionHeadline"),
  missionMeta: document.getElementById("missionMeta"),
  missionOutput: document.getElementById("missionOutput"),
  categoryGrid: document.getElementById("categoryGrid"),
  categoryTemplate: document.getElementById("categoryTemplate"),
  historyList: document.getElementById("historyList"),
  tagFilters: Array.from(document.querySelectorAll(".tag-filter")),

  mapSizePreset: document.getElementById("mapSizePreset"),
  mixedBiomesToggle: document.getElementById("mixedBiomesToggle"),
  allowVariationToggle: document.getElementById("allowVariationToggle"),
  areaBiasSlider: document.getElementById("areaBiasSlider"),
  areaBiasValue: document.getElementById("areaBiasValue"),
  rerollMapBtn: document.getElementById("rerollMapBtn"),
  rerollAreaThemeBtn: document.getElementById("rerollAreaThemeBtn"),
  rerollAreaShapeBtn: document.getElementById("rerollAreaShapeBtn"),
  mapGrid: document.getElementById("mapGrid"),
  areaLegend: document.getElementById("areaLegend"),
  mapStatus: document.getElementById("mapStatus"),
  mapCardMenu: null,
  mapCardMenuList: null,
  mapCardMenuTitle: null
};

init();

async function init() {
  setupCategoryCards();
  setupEvents();
  loadHistory();
  renderHistory();
  refs.areaBiasValue.textContent = String(getAreaSizeBias());

  const loaded = await loadData();
  if (!loaded) return;

  if (POC_MODE) {
    logDevFallback("PoC Mode: connector validation disabled; assuming all tiles support N/E/S/W connectors");
  }

  generateMap({ mode: "entire", bumpNonce: false, reason: "Initial map generated." });
  generateMission({ mode: "new" });
}

function setupEvents() {
  refs.generateBtn.addEventListener("click", () => generateMission({ mode: "new" }));
  refs.rerollAllBtn.addEventListener("click", () => generateMission({ mode: "rerollAll" }));
  refs.rerollNameBtn.addEventListener("click", rerollOperationCodenameOnce);
  refs.copyBtn.addEventListener("click", copyMissionText);
  refs.exportBtn.addEventListener("click", exportMissionJson);

  refs.tagFilters.forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      state.selectedTags = new Set(
        refs.tagFilters.filter((item) => item.checked).map((item) => item.value)
      );
      setStatus(`Filters active: ${state.selectedTags.size || 0}`);
    });
  });

  refs.mapSizePreset.addEventListener("change", () => {
    generateMap({ mode: "entire", reason: "Map size changed." });
  });

  refs.mixedBiomesToggle.addEventListener("change", () => {
    generateMap({ mode: "entire", reason: "Biome mode changed." });
  });

  refs.allowVariationToggle.addEventListener("change", () => {
    generateMap({ mode: "theme", reason: "Variation tile option changed." });
  });

  refs.areaBiasSlider.addEventListener("input", () => {
    refs.areaBiasValue.textContent = String(getAreaSizeBias());
  });

  refs.areaBiasSlider.addEventListener("change", () => {
    refs.areaBiasValue.textContent = String(getAreaSizeBias());
    generateMap({ mode: "entire", reason: "Area size bias updated." });
  });

  refs.rerollMapBtn.addEventListener("click", () => {
    generateMap({ mode: "entire", reason: "Map rerolled." });
  });

  refs.rerollAreaThemeBtn.addEventListener("click", () => {
    generateMap({ mode: "theme", reason: "Area themes rerolled." });
  });

  refs.rerollAreaShapeBtn.addEventListener("click", () => {
    generateMap({ mode: "shape", reason: "Area footprints rerolled." });
  });

  setupMapEditorUi();
}

function setupMapEditorUi() {
  const menu = document.createElement("div");
  menu.className = "map-card-menu hidden";
  menu.setAttribute("role", "menu");

  const title = document.createElement("p");
  title.className = "map-card-menu-title";
  title.textContent = "Add Card";

  const list = document.createElement("div");
  list.className = "map-card-menu-list";

  menu.appendChild(title);
  menu.appendChild(list);
  document.body.appendChild(menu);

  refs.mapCardMenu = menu;
  refs.mapCardMenuList = list;
  refs.mapCardMenuTitle = title;

  menu.addEventListener("click", (event) => {
    const targetIndex = state.mapUi.menuTargetIndex;
    if (targetIndex === null) return;

    const actionButton = event.target.closest("button[data-menu-action]");
    if (actionButton) {
      const action = actionButton.dataset.menuAction || "";
      if (action === "remove") {
        closeMapCardMenu();
        clearTileFromCell(targetIndex);
        return;
      }

      if (action === "rename-card") {
        const nameInput = refs.mapCardMenu.querySelector(".map-card-rename-input");
        const tagsInput = refs.mapCardMenu.querySelector(".map-card-rename-tags-input");
        const nextName = nameInput ? nameInput.value : "";
        const nextTags = tagsInput ? tagsInput.value : "";
        closeMapCardMenu();
        updateCardLabelAndTags(targetIndex, nextName, nextTags);
        return;
      }

      if (action === "add-room-name") {
        const nameInput = refs.mapCardMenu.querySelector(".map-card-new-name-input");
        const tagsInput = refs.mapCardMenu.querySelector(".map-card-new-tags-input");
        const biomeSelect = refs.mapCardMenu.querySelector(".map-card-new-biome-select");
        const created = addCustomRoomTile({
          name: nameInput ? nameInput.value : "",
          biome: biomeSelect ? biomeSelect.value : "unknown",
          tags: tagsInput ? tagsInput.value : ""
        });
        if (created) {
          if (nameInput) nameInput.value = "";
          if (tagsInput) tagsInput.value = "";
          const targetCell = state.map?.cells?.[targetIndex] || null;
          renderMapCardMenuOptions(targetCell);
        }
        return;
      }

      if (action === "delete-room-name") {
        const tileId = actionButton.dataset.tileId || "";
        const removedName = removeRoomTileFromLibrary(tileId);
        if (removedName) {
          setMapStatus(`Deleted room name from picker: ${removedName}.`);
          const targetCell = state.map?.cells?.[targetIndex] || null;
          renderMapCardMenuOptions(targetCell);
        }
        return;
      }

      if (action === "reset-room-list") {
        resetRoomLibrary();
        const targetCell = state.map?.cells?.[targetIndex] || null;
        renderMapCardMenuOptions(targetCell);
        return;
      }

      return;
    }

    const tileButton = event.target.closest("button[data-tile-id]");
    if (!tileButton) return;
    const tileId = tileButton.dataset.tileId || "";
    closeMapCardMenu();
    setTileForCell(targetIndex, tileId);
  });

  document.addEventListener("click", (event) => {
    if (!refs.mapCardMenu || refs.mapCardMenu.classList.contains("hidden")) return;
    if (event.target.closest(".map-card-menu")) return;
    closeMapCardMenu();
  });

  document.addEventListener("contextmenu", (event) => {
    if (event.target.closest(".map-cell")) return;
    if (event.target.closest(".map-card-menu")) return;
    closeMapCardMenu();
  });

  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return;
    closeMapCardMenu();
  });

  window.addEventListener("resize", closeMapCardMenu);
  document.addEventListener("scroll", (event) => {
    if (!refs.mapCardMenu || refs.mapCardMenu.classList.contains("hidden")) return;
    const target = event.target;
    if (target && typeof target.closest === "function" && target.closest(".map-card-menu")) return;
    closeMapCardMenu();
  }, true);
}

function setupCategoryCards() {
  refs.categoryGrid.innerHTML = "";
  CATEGORY_CONFIG.forEach((config) => {
    state.lockState[config.field] = false;
    const fragment = refs.categoryTemplate.content.cloneNode(true);
    const item = fragment.querySelector(".category-item");
    const label = fragment.querySelector(".category-label");
    const value = fragment.querySelector(".category-value");
    const rerollBtn = fragment.querySelector(".reroll-category-btn");
    const lockInput = fragment.querySelector(".lock-category");

    item.dataset.field = config.field;
    label.textContent = config.label;
    value.textContent = "-";
    rerollBtn.addEventListener("click", () => rerollCategory(config.field));
    lockInput.addEventListener("change", (event) => {
      state.lockState[config.field] = event.target.checked;
    });

    refs.categoryGrid.appendChild(fragment);
  });
}

async function loadData() {
  try {
    const [options, zonesRaw] = await Promise.all([
      fetchJsonWithFallback("data/options.json"),
      fetchJsonWithFallback("data/zones.json")
    ]);

    state.options = options;
    state.zones = normalizeZones(zonesRaw);
    loadRoomLibrary();

    validateOptions();
    validateZones();
    setStatus("Mission options and zone tiles loaded.");
    setMapStatus("Map generator ready.");
    return true;
  } catch (err) {
    setStatus("Could not load required JSON files. Run with a local server (python3 -m http.server)." );
    setMapStatus("Map generator unavailable: failed to load data/zones.json.");
    console.error(err);
    return false;
  }
}

async function fetchJsonWithFallback(path) {
  try {
    const response = await fetch(path, { cache: "no-store" });
    if (!response.ok) throw new Error(`HTTP ${response.status} for ${path}`);
    return await response.json();
  } catch (err) {
    try {
      return await loadJsonViaXhr(path);
    } catch (xhrErr) {
      console.error(err);
      console.error(xhrErr);
      throw xhrErr;
    }
  }
}

function loadJsonViaXhr(path) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open("GET", path, true);
    xhr.onreadystatechange = () => {
      if (xhr.readyState !== 4) return;
      if (xhr.status === 200 || (xhr.status === 0 && xhr.responseText)) {
        try {
          resolve(JSON.parse(xhr.responseText));
        } catch (parseErr) {
          reject(parseErr);
        }
      } else {
        reject(new Error(`XHR ${xhr.status} for ${path}`));
      }
    };
    xhr.onerror = () => reject(new Error(`XHR network error for ${path}`));
    xhr.send();
  });
}

function normalizeZones(raw) {
  const tileList = Array.isArray(raw) ? raw : raw?.tiles;
  if (!Array.isArray(tileList)) throw new Error("Invalid zones JSON: expected array or { tiles: [] }");

  return {
    tiles: tileList.map((tile) => ({
      id: String(tile.id),
      name: String(tile.name || tile.id),
      biome: String(tile.biome || "unknown"),
      tags: Array.isArray(tile.tags) ? tile.tags.map((tag) => String(tag)) : [],
      weight: Number(tile.weight || 1),
      image: resolveZoneTileImage({
        id: String(tile.id),
        name: String(tile.name || tile.id),
        biome: String(tile.biome || "unknown"),
        image: tile.image ? String(tile.image) : ""
      })
    }))
  };
}

function resolveZoneTileImage(tile) {
  const explicitPath = normalizeAssetPath(tile.image);
  if (explicitPath) return explicitPath;

  const idToken = toLookupToken(tile.id);
  if (idToken && ZONECARD_IMAGE_TOKEN_MAP.has(idToken)) {
    return ZONECARD_IMAGE_TOKEN_MAP.get(idToken) || "";
  }

  const nameToken = toLookupToken(tile.name);
  if (nameToken && ZONECARD_IMAGE_TOKEN_MAP.has(nameToken)) {
    return ZONECARD_IMAGE_TOKEN_MAP.get(nameToken) || "";
  }

  const themeKey = BIOME_ART_THEME[tile.biome] || BIOME_ART_THEME.unknown;
  const fallbackPool = ZONECARD_IMAGE_LIBRARY[themeKey] || ZONECARD_IMAGE_PATHS;
  if (fallbackPool.length === 0) return "";

  const fallbackIndex = hashSeed(`${tile.id}:${tile.biome}`) % fallbackPool.length;
  return fallbackPool[fallbackIndex];
}

function normalizeAssetPath(path) {
  return String(path || "").replaceAll("\\", "/").trim();
}

function toAssetUrl(path) {
  const normalized = normalizeAssetPath(path);
  if (!normalized) return "";
  return encodeURI(normalized);
}

function toLookupToken(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function validateOptions() {
  CATEGORY_CONFIG.forEach(({ source }) => {
    if (!source) return;
    if (!Array.isArray(state.options[source]) || state.options[source].length === 0) {
      throw new Error(`Missing or empty category: ${source}`);
    }
  });

  const codenameKeys = ["operationCodenamePrefixes", "operationCodenameNouns", "operationCodenameSuffixes"];
  codenameKeys.forEach((key) => {
    if (!Array.isArray(state.options[key]) || state.options[key].length === 0) {
      throw new Error(`Missing or empty codename bank: ${key}`);
    }
  });
}

function validateZones() {
  const minimumTileCount = POC_MODE ? 1 : 6;
  if (!state.zones || !Array.isArray(state.zones.tiles) || state.zones.tiles.length < minimumTileCount) {
    throw new Error(`Zones data too small. Expected at least ${minimumTileCount} tile(s).`);
  }

  const ids = new Set();
  state.zones.tiles.forEach((tile) => {
    if (ids.has(tile.id)) {
      throw new Error(`Duplicate zone tile id: ${tile.id}`);
    }
    ids.add(tile.id);
  });
}

function loadRoomLibrary() {
  const seenIds = new Set((state.zones?.tiles || []).map((tile) => tile.id));
  const customTiles = [];
  const hiddenTileIds = new Set();

  try {
    const rawCustom = JSON.parse(localStorage.getItem(CUSTOM_ROOM_NAMES_KEY) || "[]");
    if (Array.isArray(rawCustom)) {
      rawCustom.forEach((entry) => {
        const tile = normalizeCustomRoomTile(entry);
        if (!tile || seenIds.has(tile.id)) return;
        seenIds.add(tile.id);
        customTiles.push(tile);
      });
    }
  } catch (err) {
    customTiles.length = 0;
  }

  try {
    const rawHidden = JSON.parse(localStorage.getItem(HIDDEN_ROOM_NAMES_KEY) || "[]");
    if (Array.isArray(rawHidden)) {
      rawHidden.forEach((id) => {
        const text = String(id || "").trim();
        if (text) hiddenTileIds.add(text);
      });
    }
  } catch (err) {
    hiddenTileIds.clear();
  }

  state.roomLibrary.customTiles = customTiles;
  state.roomLibrary.hiddenTileIds = hiddenTileIds;
}

function saveRoomLibrary() {
  try {
    const serializableCustom = state.roomLibrary.customTiles.map((tile) => ({
      id: tile.id,
      name: tile.name,
      biome: tile.biome,
      tags: [...tile.tags],
      image: tile.image || "",
      weight: Number(tile.weight || 1)
    }));
    localStorage.setItem(CUSTOM_ROOM_NAMES_KEY, JSON.stringify(serializableCustom));
    localStorage.setItem(HIDDEN_ROOM_NAMES_KEY, JSON.stringify([...state.roomLibrary.hiddenTileIds]));
  } catch (err) {
    // Persistence is optional; ignore storage failures.
  }
}

function parseTagInput(value) {
  if (Array.isArray(value)) {
    return Array.from(new Set(
      value
        .map((tag) => String(tag || "").trim().toLowerCase())
        .filter(Boolean)
    ));
  }

  return Array.from(new Set(
    String(value || "")
      .split(",")
      .map((tag) => tag.trim().toLowerCase())
      .filter(Boolean)
  ));
}

function buildCustomRoomTileId(name) {
  const base = String(name || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 24) || "room";
  const nonce = Math.random().toString(36).slice(2, 7);
  return `custom_${base}_${Date.now().toString(36)}_${nonce}`;
}

function normalizeCustomRoomTile(raw) {
  if (!raw || typeof raw !== "object") return null;
  const name = String(raw.name || "").trim();
  if (!name) return null;

  const id = String(raw.id || buildCustomRoomTileId(name)).trim();
  if (!id) return null;

  const biome = String(raw.biome || "unknown").trim() || "unknown";
  const tags = parseTagInput(raw.tags);
  const image = resolveZoneTileImage({
    id,
    name,
    biome,
    image: raw.image ? String(raw.image) : ""
  });

  return {
    id,
    name,
    biome,
    tags,
    weight: Math.max(1, Number(raw.weight || 1)),
    image,
    isCustom: true
  };
}

function getEditableRoomTiles() {
  const hidden = state.roomLibrary.hiddenTileIds;
  const baseTiles = (state.zones?.tiles || []).filter((tile) => !hidden.has(tile.id));
  const customTiles = (state.roomLibrary.customTiles || []).filter((tile) => !hidden.has(tile.id));
  return [...baseTiles, ...customTiles];
}

function findEditableRoomTileById(tileId) {
  if (!tileId) return null;
  const lookupId = String(tileId);
  return getEditableRoomTiles().find((tile) => tile.id === lookupId) || null;
}

function getRoomLibraryBiomes() {
  const biomeSet = new Set();
  (state.zones?.tiles || []).forEach((tile) => biomeSet.add(tile.biome || "unknown"));
  (state.roomLibrary.customTiles || []).forEach((tile) => biomeSet.add(tile.biome || "unknown"));
  if (biomeSet.size === 0) biomeSet.add("unknown");
  return Array.from(biomeSet).sort((a, b) => a.localeCompare(b));
}

function addCustomRoomTile({ name, biome, tags }) {
  const normalized = normalizeCustomRoomTile({
    id: buildCustomRoomTileId(name),
    name,
    biome,
    tags
  });
  if (!normalized) {
    setMapStatus("Room name must include at least one visible character.");
    return null;
  }

  state.roomLibrary.customTiles.push(normalized);
  state.roomLibrary.hiddenTileIds.delete(normalized.id);
  saveRoomLibrary();
  setMapStatus(`Added new room name: ${normalized.name}.`);
  return normalized;
}

function removeRoomTileFromLibrary(tileId) {
  const lookupId = String(tileId || "").trim();
  if (!lookupId) return "";

  const customIndex = state.roomLibrary.customTiles.findIndex((tile) => tile.id === lookupId);
  if (customIndex >= 0) {
    const [removedTile] = state.roomLibrary.customTiles.splice(customIndex, 1);
    state.roomLibrary.hiddenTileIds.delete(lookupId);
    saveRoomLibrary();
    return removedTile.name || lookupId;
  }

  const baseTile = state.zones?.tiles?.find((tile) => tile.id === lookupId);
  if (baseTile) {
    state.roomLibrary.hiddenTileIds.add(lookupId);
    saveRoomLibrary();
    return baseTile.name || lookupId;
  }

  return "";
}

function resetRoomLibrary() {
  state.roomLibrary.customTiles = [];
  state.roomLibrary.hiddenTileIds = new Set();
  saveRoomLibrary();
  setMapStatus("Room name picker reset to defaults.");
}

function generateMission({ mode }) {
  if (!state.options) return;

  const seed = refs.seedInput.value.trim();
  const mission = state.mission ? { ...state.mission } : {};
  const usedTexts = new Set();

  CATEGORY_CONFIG.forEach((cfg) => {
    if (mode === "rerollAll" && state.lockState[cfg.field] && mission[cfg.field]) {
      usedTexts.add(mission[cfg.field]);
      return;
    }
    mission[cfg.field] = rollCategory(cfg, seed, usedTexts, mission);
    usedTexts.add(mission[cfg.field]);
  });

  mission.seed = seed || "";
  mission.id = createMissionId(seed);

  state.mission = mission;
  attachMapToMission();
  renderMission();
  pushHistory(state.mission);
  renderHistory();
}

function rerollCategory(field) {
  if (!state.options || !state.mission) return;
  const cfg = CATEGORY_CONFIG.find((item) => item.field === field);
  if (!cfg) return;

  const seed = refs.seedInput.value.trim();
  const usedTexts = new Set();
  CATEGORY_CONFIG.forEach((c) => {
    if (c.field !== field && state.mission[c.field]) {
      usedTexts.add(state.mission[c.field]);
    }
  });

  state.mission[field] = rollCategory(cfg, seed, usedTexts, state.mission);
  if (field !== "operationCodename" && !state.lockState.operationCodename) {
    const opCfg = CATEGORY_CONFIG.find((item) => item.field === "operationCodename");
    if (opCfg) {
      state.mission.operationCodename = rollCategory(opCfg, seed, usedTexts, state.mission);
    }
  }

  state.mission.seed = seed || "";
  state.mission.id = createMissionId(seed);
  attachMapToMission();
  renderMission();
  pushHistory(state.mission);
  renderHistory();
}

function rerollOperationCodenameOnce() {
  if (!state.options || !state.mission) return;
  state.mission.operationCodename = generateOperationCodename("", state.mission, { forceRandom: true });
  state.mission.id = createMissionId(refs.seedInput.value.trim());
  attachMapToMission();
  renderMission();
  pushHistory(state.mission);
  renderHistory();
  setStatus("Operation codename rerolled (non-seeded one-time roll).");
}

function rollCategory(cfg, seed, usedTexts, mission) {
  if (cfg.field === "operationCodename") {
    return generateOperationCodename(seed, mission);
  }

  const pool = state.options[cfg.source] || [];
  if (pool.length === 0) return "";

  const filtered = filterByTags(pool, state.selectedTags);
  const candidatePool = filtered.length ? filtered : pool;
  const deterministicSeed = seed ? `${seed}:${cfg.field}` : "";
  const rng = deterministicSeed ? createRng(deterministicSeed) : Math.random;

  if (!refs.uniqueToggle.checked) {
    return weightedPick(candidatePool, rng).text;
  }

  for (let i = 0; i < 50; i += 1) {
    const pick = weightedPick(candidatePool, rng).text;
    if (!usedTexts.has(pick)) {
      return pick;
    }
  }

  return weightedPick(candidatePool, rng).text;
}

function generateOperationCodename(seed, mission, options = {}) {
  const prefixPool = state.options.operationCodenamePrefixes || [];
  const nounPool = state.options.operationCodenameNouns || [];
  const suffixPool = state.options.operationCodenameSuffixes || [];
  if (!prefixPool.length || !nounPool.length || !suffixPool.length) return "Operation Silent Vector";

  const missionTags = collectMissionTags(mission);
  const filterTags = Array.from(state.selectedTags);
  const targetTags = new Set([...missionTags, ...filterTags]);

  const rng = options.forceRandom ? Math.random : (seed ? createRng(`${seed}:operationCodename`) : Math.random);
  const prefix = pickCodenamePart(prefixPool, targetTags, rng);
  const noun = pickCodenamePart(nounPool, targetTags, rng);
  const suffix = pickCodenamePart(suffixPool, targetTags, rng);
  const styleRoll = randomFloat(rng);

  if (styleRoll < 0.17) return `Operation ${noun} ${suffix}`;
  if (styleRoll < 0.33) return `Operation ${prefix}-${suffix}`;
  return `Operation ${prefix} ${noun}`;
}

function pickCodenamePart(pool, targetTags, rng) {
  const themed = filterByTags(pool, targetTags);
  const candidatePool = themed.length ? themed : pool;
  return weightedPick(candidatePool, rng).text;
}

function collectMissionTags(mission) {
  const tags = new Set();
  if (!mission) return tags;

  CATEGORY_CONFIG.forEach((cfg) => {
    if (!cfg.source || !mission[cfg.field]) return;
    const pool = state.options[cfg.source] || [];
    const match = pool.find((entry) => entry.text === mission[cfg.field]);
    if (!match || !Array.isArray(match.tags)) return;
    match.tags.forEach((tag) => tags.add(tag));
  });

  return tags;
}

function generateMap({ mode = "entire", bumpNonce = true, reason = "" } = {}) {
  if (!state.zones?.tiles?.length) return;

  try {
    if (bumpNonce) state.mapNonce += 1;

    const preset = MAP_PRESETS[refs.mapSizePreset.value] || MAP_PRESETS["7x9"];
    const seed = refs.seedInput.value.trim();
    const rng = seed ? createRng(`${seed}:map:${state.mapNonce}:${mode}`) : Math.random;
    const tracker = {
      best: null,
      bestStrict: null,
      bestRelaxed: null
    };

    const considerCandidate = (candidateMap, relaxed = false) => {
      if (!candidateMap) return null;
      const quality = validateMapQuality(candidateMap, { relaxed });
      if (!tracker.best || quality.score > tracker.best.quality.score) {
        tracker.best = { map: candidateMap, quality };
      }
      if (!relaxed && (!tracker.bestStrict || quality.score > tracker.bestStrict.quality.score)) {
        tracker.bestStrict = { map: candidateMap, quality };
      }
      if (relaxed && (!tracker.bestRelaxed || quality.score > tracker.bestRelaxed.quality.score)) {
        tracker.bestRelaxed = { map: candidateMap, quality };
      }
      if (quality.ok) return quality;
      return null;
    };

    let nextMap = null;
    let acceptedQuality = null;

    if (mode === "entire" || !state.map || state.map.rows !== preset.rows || state.map.cols !== preset.cols) {
      for (let attempt = 0; attempt < QUALITY_MAX_ATTEMPTS; attempt += 1) {
        const candidateMap = buildEntireMapCandidate({ preset, seed, rng });
        const quality = considerCandidate(candidateMap, false);
        if (!quality) continue;
        nextMap = candidateMap;
        acceptedQuality = quality;
        break;
      }

      if (!nextMap) {
        for (let attempt = 0; attempt < Math.ceil(QUALITY_MAX_ATTEMPTS / 2); attempt += 1) {
          const candidateMap = buildEntireMapCandidate({ preset, seed, rng });
          const quality = considerCandidate(candidateMap, true);
          if (!quality) continue;
          nextMap = candidateMap;
          acceptedQuality = quality;
          break;
        }
      }

      if (!nextMap) {
        const emergency = tracker.best?.map || buildEmergencyMapCandidate({ preset, seed, rng }) || buildLastResortMapCandidate({ preset, seed, rng });
        if (!emergency) {
          setMapStatus("Sparse map generation failed to produce any valid candidate.");
          return;
        }
        nextMap = emergency;
        acceptedQuality = tracker.best?.quality || validateMapQuality(nextMap, { relaxed: true });
      }
    } else if (mode === "theme") {
      for (let attempt = 0; attempt < 24; attempt += 1) {
        const candidateMap = buildThemeMapCandidate({ preset, seed, rng });
        const quality = considerCandidate(candidateMap, false);
        if (!quality) continue;
        nextMap = candidateMap;
        acceptedQuality = quality;
        break;
      }

      if (!nextMap) {
        for (let attempt = 0; attempt < 12; attempt += 1) {
          const candidateMap = buildThemeMapCandidate({ preset, seed, rng });
          const quality = considerCandidate(candidateMap, true);
          if (!quality) continue;
          nextMap = candidateMap;
          acceptedQuality = quality;
          break;
        }
      }

      if (!nextMap) {
        nextMap = tracker.best?.map || buildEmergencyMapCandidate({ preset, seed, rng }) || buildLastResortMapCandidate({ preset, seed, rng });
        acceptedQuality = tracker.best?.quality || (nextMap ? validateMapQuality(nextMap, { relaxed: true }) : null);
      }
    } else if (mode === "shape") {
      for (let attempt = 0; attempt < QUALITY_MAX_ATTEMPTS; attempt += 1) {
        const candidateMap = buildShapeMapCandidate({ preset, seed, rng });
        const quality = considerCandidate(candidateMap, false);
        if (!quality) continue;
        nextMap = candidateMap;
        acceptedQuality = quality;
        break;
      }

      if (!nextMap) {
        for (let attempt = 0; attempt < Math.ceil(QUALITY_MAX_ATTEMPTS / 2); attempt += 1) {
          const candidateMap = buildShapeMapCandidate({ preset, seed, rng });
          const quality = considerCandidate(candidateMap, true);
          if (!quality) continue;
          nextMap = candidateMap;
          acceptedQuality = quality;
          break;
        }
      }

      if (!nextMap) {
        nextMap = tracker.best?.map || buildEmergencyMapCandidate({ preset, seed, rng }) || buildLastResortMapCandidate({ preset, seed, rng });
        acceptedQuality = tracker.best?.quality || (nextMap ? validateMapQuality(nextMap, { relaxed: true }) : null);
      }
    }

    if (!nextMap) {
      setMapStatus("Sparse map generation did not return a map.");
      return;
    }

    state.map = nextMap;
    renderMap();
    attachMapToMission();
    if (state.mission) renderMission();

    if (reason) {
      setMapStatus(reason);
    } else if (acceptedQuality && !acceptedQuality.ok) {
      const firstReason = acceptedQuality.reasons?.[0] || "best-effort fallback accepted";
      setMapStatus(`PoC fallback map used (score ${acceptedQuality.score}): ${firstReason}`);
    }
  } catch (err) {
    console.error(err);
    setMapStatus("Map generation error encountered; previous map preserved.");
  }
}

function buildEntireMapCandidate({ preset, seed, rng }) {
  const areaProfiles = buildGameplayAreaProfiles(getAreaSizeBias(), rng);
  const areaCount = areaProfiles.length;
  const areaSizes = areaProfiles.map((profile) => profile.size);
  const graph = buildAreaGraph(areaCount, rng);
  const plan = buildSparsePlan({
    rows: preset.rows,
    cols: preset.cols,
    areaSizes,
    areaProfiles,
    graph,
    rng,
    maxAttempts: SPARSE_PLAN_ATTEMPT_LIMIT
  });

  if (!plan) return null;

  const dominantTiles = pickDominantTiles(areaCount, rng, isMixedBiomesEnabled(), [], areaProfiles, graph);
  return buildMapFromSparsePlan({
    preset,
    seed,
    plan,
    graph,
    areaSizes,
    areaProfiles,
    dominantTiles,
    allowVariation: allowVariationTiles(),
    rng,
    previousGateStates: null
  });
}

function buildThemeMapCandidate({ preset, seed, rng }) {
  const previousPlan = state.map?.plan;
  if (!previousPlan) return null;

  const graph = createGraphFromEdges(previousPlan.graph.template, previousPlan.graph.nodeCount, previousPlan.graph.edges);
  const planSnapshot = {
    areaCellsByIndex: previousPlan.areaCellsByIndex.map((cells) => [...cells]),
    corridors: previousPlan.corridors.map((corridor) => ({ ...corridor, cells: [...corridor.cells] })),
    occupiedCount: previousPlan.occupiedCount,
    areaProfiles: previousPlan.areaProfiles
      ? previousPlan.areaProfiles.map((profile) => ({ ...profile }))
      : previousPlan.areaSizes.map((size, idx) => ({
        index: idx,
        size,
        category: getAreaSizeInfo(size).category,
        shapeArchetype: "blob",
        shapeKind: "room"
      }))
  };

  const dominantTiles = pickDominantTiles(
    previousPlan.areaSizes.length,
    rng,
    isMixedBiomesEnabled(),
    state.map.areas.map((area) => area.dominantTileId),
    planSnapshot.areaProfiles,
    graph
  );

  return buildMapFromSparsePlan({
    preset,
    seed,
    plan: planSnapshot,
    graph,
    areaSizes: [...previousPlan.areaSizes],
    areaProfiles: planSnapshot.areaProfiles,
    dominantTiles,
    allowVariation: allowVariationTiles(),
    rng,
    previousGateStates: state.map.gates
  });
}

function buildShapeMapCandidate({ preset, seed, rng }) {
  const previousPlan = state.map?.plan;
  if (!previousPlan) return null;

  const graph = createGraphFromEdges(previousPlan.graph.template, previousPlan.graph.nodeCount, previousPlan.graph.edges);
  const areaProfiles = previousPlan.areaProfiles
    ? previousPlan.areaProfiles.map((profile) => ({ ...profile }))
    : previousPlan.areaSizes.map((size, idx) => ({
      index: idx,
      size,
      category: getAreaSizeInfo(size).category,
      shapeArchetype: "blob",
      shapeKind: "room"
    }));

  const rebuiltPlan = buildSparsePlan({
    rows: preset.rows,
    cols: preset.cols,
    areaSizes: [...previousPlan.areaSizes],
    areaProfiles,
    graph,
    rng,
    maxAttempts: SPARSE_PLAN_ATTEMPT_LIMIT
  });

  if (!rebuiltPlan) return null;

  const dominantTiles = state.map.areas.map((area) => {
    const tile = state.zones.tiles.find((item) => item.id === area.dominantTileId);
    return tile || weightedPick(state.zones.tiles, rng);
  });

  return buildMapFromSparsePlan({
    preset,
    seed,
    plan: rebuiltPlan,
    graph,
    areaSizes: [...previousPlan.areaSizes],
    areaProfiles,
    dominantTiles,
    allowVariation: allowVariationTiles(),
    rng,
    previousGateStates: state.map.gates
  });
}

function buildEmergencyMapCandidate({ preset, seed, rng }) {
  for (let attempt = 0; attempt < 140; attempt += 1) {
    const fallbackProfiles = [
      { index: 0, category: "Medium", size: 6, shapeArchetype: "blob", shapeKind: "room" },
      { index: 1, category: "Small", size: 4, shapeArchetype: "rectangle", shapeKind: "room" },
      { index: 2, category: "Small", size: 3, shapeArchetype: "l_shape", shapeKind: "room" },
      { index: 3, category: "Tiny", size: 2, shapeArchetype: "short_choke", shapeKind: "corridor" }
    ];
    assignAreaShapeArchetypes(fallbackProfiles, rng);
    const areaSizes = fallbackProfiles.map((profile) => profile.size);
    const graph = createGraphFromEdges("spine_branch", 4, buildSpineBranchEdges(4));
    const plan = buildSparsePlan({
      rows: preset.rows,
      cols: preset.cols,
      areaSizes,
      areaProfiles: fallbackProfiles,
      graph,
      rng,
      maxAttempts: SPARSE_PLAN_ATTEMPT_LIMIT * 2
    });

    if (!plan) continue;

    const dominantTiles = pickDominantTiles(4, rng, isMixedBiomesEnabled(), [], fallbackProfiles, graph);
    return buildMapFromSparsePlan({
      preset,
      seed,
      plan,
      graph,
      areaSizes,
      areaProfiles: fallbackProfiles,
      dominantTiles,
      allowVariation: true,
      rng,
      previousGateStates: null
    });
  }

  return null;
}

function buildLastResortMapCandidate({ preset, seed, rng }) {
  const rows = preset.rows;
  const cols = preset.cols;
  if (rows < 6 || cols < 7) return null;

  const areaProfiles = [
    { index: 0, category: "Medium", size: 6, shapeArchetype: "rectangle", shapeKind: "room" },
    { index: 1, category: "Small", size: 4, shapeArchetype: "l_shape", shapeKind: "room" },
    { index: 2, category: "Small", size: 3, shapeArchetype: "blob", shapeKind: "room" },
    { index: 3, category: "Tiny", size: 2, shapeArchetype: "short_choke", shapeKind: "corridor" }
  ];

  const areaCellsByIndex = [
    [toCellIndex(2, 1, cols), toCellIndex(2, 2, cols), toCellIndex(3, 1, cols), toCellIndex(3, 2, cols), toCellIndex(4, 1, cols), toCellIndex(4, 2, cols)],
    [toCellIndex(1, 5, cols), toCellIndex(2, 5, cols), toCellIndex(2, 6, cols), toCellIndex(3, 5, cols)],
    [toCellIndex(5, 4, cols), toCellIndex(5, 5, cols), toCellIndex(5, 6, cols)],
    [toCellIndex(1, 3, cols), toCellIndex(1, 4, cols)]
  ];
  const corridors = [
    { id: "C1", from: 0, to: 3, cells: [toCellIndex(2, 3, cols)], shapeArchetype: "short_choke" },
    { id: "C2", from: 3, to: 1, cells: [toCellIndex(2, 4, cols)], shapeArchetype: "short_choke" },
    { id: "C3", from: 1, to: 2, cells: [toCellIndex(4, 5, cols)], shapeArchetype: "short_choke" }
  ];
  const areaSizes = areaProfiles.map((profile) => profile.size);
  const occupiedCount = areaSizes.reduce((sum, value) => sum + value, 0) + corridors.reduce((sum, item) => sum + item.cells.length, 0);
  const graph = createGraphFromEdges("spine_branch", 4, [[0, 3], [3, 1], [1, 2]]);
  const plan = {
    areaCellsByIndex,
    corridors,
    occupiedCount,
    areaProfiles
  };
  const dominantTiles = pickDominantTiles(4, rng, isMixedBiomesEnabled(), [], areaProfiles, graph);

  return buildMapFromSparsePlan({
    preset,
    seed,
    plan,
    graph,
    areaSizes,
    areaProfiles,
    dominantTiles,
    allowVariation: true,
    rng,
    previousGateStates: null
  });
}

function getAreaSizeInfo(size) {
  if (size <= 2) return { category: "Tiny", cardsReturned: 0 };
  if (size <= 4) return { category: "Small", cardsReturned: 1 };
  if (size <= 10) return { category: "Medium", cardsReturned: 2 };
  if (size <= 16) return { category: "Large", cardsReturned: 3 };
  if (size <= 32) return { category: "Huge", cardsReturned: 5 };
  return { category: "Humongous", cardsReturned: "all" };
}

function getCategoryBounds(category) {
  if (category === "Tiny") return { min: 1, max: 2 };
  if (category === "Small") return { min: 3, max: 4 };
  if (category === "Medium") return { min: 5, max: 10 };
  if (category === "Large") return { min: 11, max: 16 };
  if (category === "Huge") return { min: 17, max: 32 };
  return { min: 33, max: 40 };
}

function areaCategoryRank(category) {
  if (category === "Tiny") return 1;
  if (category === "Small") return 2;
  if (category === "Medium") return 3;
  if (category === "Large") return 4;
  if (category === "Huge") return 5;
  return 6;
}

function buildGameplayAreaProfiles(areaSizeBias, rng) {
  for (let attempt = 0; attempt < 120; attempt += 1) {
    const areaCount = weightedPickFromValues([4, 5, 6], [3.4, 2.2, 0.9], rng);
    const categories = ["Medium", "Small", "Small", "Tiny"];

    if (areaCount >= 5) {
      categories.push(weightedPickFromValues(["Small", "Medium"], [2.8, 1.0], rng));
    }
    if (areaCount >= 6) {
      categories.push(weightedPickFromValues(["Small", "Medium", "Tiny"], [2.3, 0.9, 0.45], rng));
    }

    while (categories.length > areaCount) categories.pop();

    const tinyIndexes = categories
      .map((item, index) => ({ item, index }))
      .filter((entry) => entry.item === "Tiny")
      .map((entry) => entry.index);
    if (tinyIndexes.length > 2) {
      tinyIndexes.slice(2).forEach((index) => {
        categories[index] = "Small";
      });
    }

    const mediumPlusIndexes = categories
      .map((item, index) => ({ item, index }))
      .filter((entry) => areaCategoryRank(entry.item) >= 3)
      .map((entry) => entry.index);
    if (mediumPlusIndexes.length > 2 && areaCount <= 5) {
      mediumPlusIndexes.slice(2).forEach((index) => {
        categories[index] = "Small";
      });
    }

    const mediumIndexes = categories
      .map((item, index) => ({ item, index }))
      .filter((entry) => entry.item === "Medium")
      .map((entry) => entry.index);
    if (mediumIndexes.length > 0 && randomFloat(rng) < 0.08) {
      const promoteIndex = mediumIndexes[randomInt(0, mediumIndexes.length - 1, rng)];
      categories[promoteIndex] = "Large";
    }

    const profiles = categories.map((category, index) => ({
      index,
      category,
      size: rollAreaSizeForCategory(category, areaSizeBias, rng),
      shapeArchetype: "blob",
      shapeKind: "room"
    }));

    if (!normalizeAreaProfileTileBudget(profiles, rng, { min: 11, max: 16 })) continue;

    profiles.forEach((profile) => {
      profile.category = getAreaSizeInfo(profile.size).category;
    });

    if (profiles.filter((profile) => profile.category === "Tiny").length > 2) continue;
    if (!profiles.some((profile) => areaCategoryRank(profile.category) >= 3)) continue;
    if (profiles.filter((profile) => profile.category === "Large").length > 1 && randomFloat(rng) < 0.92) continue;
    if (profiles.filter((profile) => areaCategoryRank(profile.category) >= 3).length > 2 && areaCount <= 5) continue;

    assignAreaShapeArchetypes(profiles, rng);
    if (!passesAreaProfileShapeRules(profiles)) continue;

    return profiles;
  }

  const fallback = [
    { index: 0, category: "Medium", size: 6, shapeArchetype: "blob", shapeKind: "room" },
    { index: 1, category: "Small", size: 4, shapeArchetype: "rectangle", shapeKind: "room" },
    { index: 2, category: "Small", size: 3, shapeArchetype: "l_shape", shapeKind: "room" },
    { index: 3, category: "Tiny", size: 1, shapeArchetype: "straight_corridor", shapeKind: "corridor" }
  ];
  assignAreaShapeArchetypes(fallback, rng);
  return fallback;
}

function rollAreaSizeForCategory(category, areaSizeBias, rng) {
  let bounds = getCategoryBounds(category);
  if (category === "Medium") bounds = { min: 5, max: 8 };
  if (category === "Large") bounds = { min: 11, max: 14 };
  const min = bounds.min;
  const max = bounds.max;
  if (min === max) return min;

  const normalizedBias = Math.max(-1, Math.min(1, areaSizeBias / 100));
  const center = normalizedBias >= 0
    ? min + ((max - min) * (0.45 + (normalizedBias * 0.35)))
    : min + ((max - min) * (0.55 + (normalizedBias * 0.35)));

  const values = [];
  const weights = [];
  for (let value = min; value <= max; value += 1) {
    values.push(value);
    const distance = Math.abs(value - center);
    weights.push(2.8 - Math.min(2.4, distance * 0.6));
  }

  return weightedPickFromValues(values, weights, rng);
}

function normalizeAreaProfileTileBudget(profiles, rng, target = { min: TARGET_OCCUPIED_MIN, max: TARGET_OCCUPIED_MAX }) {
  const targetMin = Number(target.min ?? TARGET_OCCUPIED_MIN);
  const targetMax = Number(target.max ?? TARGET_OCCUPIED_MAX);
  let total = profiles.reduce((sum, profile) => sum + profile.size, 0);
  let guard = 0;

  while (total < targetMin && guard < 80) {
    guard += 1;
    const candidates = profiles.filter((profile) => {
      const bounds = getCategoryBounds(profile.category);
      return profile.size < bounds.max;
    });
    if (candidates.length === 0) break;

    const weights = candidates.map((profile) => {
      const rank = areaCategoryRank(profile.category);
      if (rank === 1) return 0.8;
      if (rank === 2) return 2.0;
      if (rank === 3) return 2.4;
      return 0.55;
    });

    const pick = weightedPickFromValues(candidates, weights, rng);
    pick.size += 1;
    total += 1;
  }

  while (total > targetMax && guard < 160) {
    guard += 1;
    const candidates = profiles.filter((profile) => {
      const bounds = getCategoryBounds(profile.category);
      return profile.size > bounds.min;
    });
    if (candidates.length === 0) break;

    const weights = candidates.map((profile) => {
      const rank = areaCategoryRank(profile.category);
      if (rank >= 4) return 2.9;
      if (rank === 3) return 2.0;
      if (rank === 2) return 1.0;
      return 0.5;
    });

    const pick = weightedPickFromValues(candidates, weights, rng);
    pick.size -= 1;
    total -= 1;
  }

  return total >= targetMin && total <= targetMax;
}

function assignAreaShapeArchetypes(profiles, rng) {
  const maxCorridorAreas = Math.min(
    Math.floor(profiles.length * MAX_CORRIDOR_AREA_RATIO),
    Math.max(0, profiles.length - 2)
  );
  const desiredCorridorAreas = maxCorridorAreas <= 0
    ? 0
    : weightedPickFromValues(
      [0, 1, maxCorridorAreas],
      [1.7, 2.5, 0.55],
      rng
    );

  const sortedCandidates = profiles
    .map((profile, index) => ({ profile, index }))
    .sort((a, b) => {
      const rankA = areaCategoryRank(a.profile.category);
      const rankB = areaCategoryRank(b.profile.category);
      if (rankA !== rankB) return rankA - rankB;
      return b.profile.size - a.profile.size;
    });

  const corridorIndices = new Set();
  for (const candidate of sortedCandidates) {
    if (corridorIndices.size >= desiredCorridorAreas) break;
    const rank = areaCategoryRank(candidate.profile.category);
    if (rank <= 2 && randomFloat(rng) < 0.7) {
      corridorIndices.add(candidate.index);
    }
  }

  profiles.forEach((profile, index) => {
    if (corridorIndices.has(index)) {
      profile.shapeKind = "corridor";
      if (profile.size > 3) profile.size = 3;
      profile.category = getAreaSizeInfo(profile.size).category;
      profile.shapeArchetype = weightedPickFromValues(
        CORRIDOR_SHAPE_ARCHETYPES,
        [1.7, 1.3, 1.6],
        rng
      );
      return;
    }

    profile.shapeKind = "room";
    profile.shapeArchetype = pickRoomShapeArchetype(profile, rng);
  });

  if (!profiles.some((profile) => CORE_ROOM_SHAPE_ARCHETYPES.includes(profile.shapeArchetype))) {
    const fallback = profiles.find((profile) => profile.shapeKind === "room") || profiles[0];
    fallback.shapeKind = "room";
    fallback.shapeArchetype = weightedPickFromValues(
      CORE_ROOM_SHAPE_ARCHETYPES,
      [1.3, 1.2, 1.1, 1.6],
      rng
    );
  }

  const roomCount = profiles.filter((profile) => profile.shapeKind === "room").length;
  const requiredRooms = Math.max(2, Math.ceil(profiles.length * 0.7));
  if (roomCount < requiredRooms) {
    const corridorProfiles = profiles
      .map((profile, index) => ({ profile, index }))
      .filter((entry) => entry.profile.shapeKind === "corridor")
      .sort((a, b) => b.profile.size - a.profile.size);

    for (const entry of corridorProfiles) {
      if (profiles.filter((profile) => profile.shapeKind === "room").length >= requiredRooms) break;
      entry.profile.shapeKind = "room";
      entry.profile.shapeArchetype = pickRoomShapeArchetype(entry.profile, rng);
      if (entry.profile.size < 3) entry.profile.size = 3;
      entry.profile.category = getAreaSizeInfo(entry.profile.size).category;
    }
  }
}

function pickRoomShapeArchetype(profile, rng) {
  const category = profile.category;
  const size = profile.size;

  if (category === "Tiny") {
    return weightedPickFromValues(
      ["rectangle", "l_shape", "blob"],
      [1.8, 1.2, 1.1],
      rng
    );
  }

  if (category === "Small") {
    return weightedPickFromValues(
      ["rectangle", "l_shape", "blob", "wide_hall"],
      [2.0, 1.4, 1.2, 1.0],
      rng
    );
  }

  if (size >= 8) {
    return weightedPickFromValues(
      ROOM_SHAPE_ARCHETYPES,
      [1.8, 1.3, 1.2, 1.5, 1.0, 0.9],
      rng
    );
  }

  return weightedPickFromValues(
    ["rectangle", "l_shape", "t_shape", "blob", "wide_hall"],
    [1.8, 1.3, 1.1, 1.4, 0.9],
    rng
  );
}

function passesAreaProfileShapeRules(profiles) {
  const corridorOnlyCount = profiles.filter((profile) => profile.shapeKind === "corridor").length;
  const corridorRatio = corridorOnlyCount / profiles.length;
  if (corridorRatio > MAX_CORRIDOR_AREA_RATIO) return false;
  if (profiles.filter((profile) => profile.shapeKind === "room").length < 2) return false;
  if ((profiles.filter((profile) => profile.shapeKind === "room").length / profiles.length) < 0.7) return false;
  if (profiles.some((profile) => profile.shapeKind === "corridor" && profile.size > 3)) return false;
  if (!profiles.some((profile) => CORE_ROOM_SHAPE_ARCHETYPES.includes(profile.shapeArchetype))) return false;
  return true;
}

function scoreOccupancyDistance(occupancyRatio) {
  if (occupancyRatio >= SPARSE_OCCUPANCY_MIN && occupancyRatio <= SPARSE_OCCUPANCY_MAX) {
    return Math.abs(occupancyRatio - 0.5);
  }
  if (occupancyRatio < SPARSE_OCCUPANCY_MIN) {
    return (SPARSE_OCCUPANCY_MIN - occupancyRatio) + 0.5;
  }
  return (occupancyRatio - SPARSE_OCCUPANCY_MAX) + 1;
}

function generateAreaSizes(areaCount, totalAreaCells, bias, rng) {
  const sizes = Array(areaCount).fill(2);
  let remaining = totalAreaCells - (areaCount * 2);

  while (remaining > 0) {
    const candidates = [];
    const weights = [];

    for (let i = 0; i < sizes.length; i += 1) {
      if (sizes[i] >= 6) continue;
      candidates.push(i);
      weights.push(sizeGrowthWeight(sizes[i] + 1, bias));
    }

    if (candidates.length === 0) break;
    const chosenIdx = weightedPickFromValues(candidates, weights, rng);
    sizes[chosenIdx] += 1;
    remaining -= 1;
  }

  return sizes;
}

function sizeGrowthWeight(nextSize, bias) {
  const base = {
    2: 1,
    3: 3.3,
    4: 3.3,
    5: 1.45,
    6: 0.8
  };

  const normalizedBias = Math.max(-1, Math.min(1, bias / 100));
  const sizeFactor = (nextSize - 2) / 4;
  const largeBoost = 1 + Math.max(0, normalizedBias) * sizeFactor * 1.35;
  const smallBoost = 1 + Math.max(0, -normalizedBias) * (1 - sizeFactor) * 1.35;

  return (base[nextSize] || 1) * largeBoost * smallBoost;
}

function buildSparsePlan({
  rows,
  cols,
  areaSizes,
  areaProfiles = null,
  graph,
  rng,
  maxAttempts = SPARSE_PLAN_ATTEMPT_LIMIT
}) {
  const areaCount = areaSizes.length;
  const totalCells = rows * cols;
  const adjacency = graph.adjacency || buildGraphAdjacency(graph, graph.nodeCount);
  const bfsOrder = buildBfsOrder(adjacency, 0, areaCount);
  const profiles = areaProfiles
    ? areaProfiles.map((profile) => ({ ...profile }))
    : areaSizes.map((size, index) => ({
      index,
      size,
      category: getAreaSizeInfo(size).category,
      shapeArchetype: "blob",
      shapeKind: "room"
    }));

  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    const centers = chooseAreaCenters({ rows, cols, areaCount, adjacency, bfsOrder, rng });
    if (!centers) continue;

    const ownerByCell = Array(totalCells).fill(-1);
    const areaCellsByIndex = Array.from({ length: areaCount }, () => []);

    const placementOrder = Array.from({ length: areaCount }, (_, index) => index)
      .sort((a, b) => areaSizes[b] - areaSizes[a]);

    let placementFailed = false;

    for (const areaIndex of placementOrder) {
      const profile = profiles[areaIndex] || {
        index: areaIndex,
        size: areaSizes[areaIndex],
        category: getAreaSizeInfo(areaSizes[areaIndex]).category,
        shapeArchetype: "blob",
        shapeKind: "room"
      };

      const shape = placeSingleAreaBlob({
        areaIndex,
        targetSize: areaSizes[areaIndex],
        profile,
        center: centers[areaIndex],
        ownerByCell,
        rows,
        cols,
        rng
      });

      if (!shape) {
        placementFailed = true;
        break;
      }

      areaCellsByIndex[areaIndex] = shape;
    }

    if (placementFailed) continue;

    const corridorResult = connectGraphWithCorridors({
      graph,
      areaCellsByIndex,
      ownerByCell,
      rows,
      cols,
      rng
    });

    if (!corridorResult) continue;

    const occupiedCount = areaSizes.reduce((sum, value) => sum + value, 0)
      + corridorResult.corridors.reduce((sum, corridor) => sum + corridor.cells.length, 0);

    const occupancyRatio = occupiedCount / totalCells;
    if (occupancyRatio > SPARSE_OCCUPANCY_MAX) continue;
    if (occupancyRatio < SPARSE_OCCUPANCY_MIN) continue;

    return {
      areaCellsByIndex,
      corridors: corridorResult.corridors,
      occupiedCount,
      areaProfiles: profiles.map((profile) => ({ ...profile }))
    };
  }

  return null;
}

function buildAreaGraph(areaCount, rng) {
  if (areaCount <= 2) {
    return createGraphFromEdges("spine", areaCount, buildSpineEdges(areaCount));
  }

  const templatePool = getGraphTemplatePool(areaCount);
  const templateId = weightedPickFromValues(
    templatePool.map((item) => item.id),
    templatePool.map((item) => item.weight),
    rng
  );

  if (templateId === "spine_branch") {
    return createGraphFromEdges(templateId, areaCount, buildSpineBranchEdges(areaCount));
  }
  if (templateId === "loop") {
    return createGraphFromEdges(templateId, areaCount, buildLoopTemplateEdges(areaCount));
  }
  if (templateId === "hub") {
    return createGraphFromEdges(templateId, areaCount, buildHubTemplateEdges(areaCount));
  }
  if (templateId === "fork_reconverge") {
    return createGraphFromEdges(templateId, areaCount, buildForkReconvergeEdges(areaCount));
  }
  if (templateId === "gauntlet_pocket") {
    return createGraphFromEdges(templateId, areaCount, buildGauntletPocketEdges(areaCount));
  }

  return createGraphFromEdges("spine", areaCount, buildSpineEdges(areaCount));
}

function createGraphFromEdges(template, nodeCount, edges) {
  return {
    template,
    nodeCount,
    edges,
    adjacency: buildGraphAdjacency({ edges }, nodeCount)
  };
}

function getGraphTemplatePool(areaCount) {
  if (areaCount <= 3) {
    return [
      { id: "spine", weight: 65 },
      { id: "spine_branch", weight: 25 },
      { id: "hub", weight: 10 }
    ];
  }

  if (areaCount === 4) {
    return GRAPH_TEMPLATE_POOL.map((item) => {
      if (item.id === "fork_reconverge" || item.id === "gauntlet_pocket") {
        return { ...item, weight: item.weight * 0.45 };
      }
      return { ...item };
    });
  }

  return GRAPH_TEMPLATE_POOL.map((item) => ({ ...item }));
}

function buildSpineEdges(count) {
  const edges = [];
  for (let i = 0; i < count - 1; i += 1) {
    edges.push([i, i + 1]);
  }
  return edges;
}

function buildSpineBranchEdges(count) {
  if (count < 4) return buildSpineEdges(count);

  const edges = [];
  if (count === 4) {
    addGraphEdge(edges, 0, 1);
    addGraphEdge(edges, 1, 2);
    addGraphEdge(edges, 1, 3);
    return edges;
  }

  for (let i = 0; i < count - 2; i += 1) {
    addGraphEdge(edges, i, i + 1);
  }
  addGraphEdge(edges, 1, count - 1);

  return edges;
}

function buildLoopTemplateEdges(count) {
  if (count < 4) return buildSpineEdges(count);

  const edges = buildSpineEdges(count);
  addGraphEdge(edges, count - 1, 1);
  if (count >= 6) {
    addGraphEdge(edges, 2, 4);
  }
  return edges;
}

function buildHubTemplateEdges(count) {
  if (count < 4) return buildSpineBranchEdges(count);

  const edges = [];
  const hub = count > 2 ? 1 : 0;

  for (let i = 0; i < count; i += 1) {
    if (i === hub) continue;
    addGraphEdge(edges, hub, i);
  }

  if (count >= 6) {
    addGraphEdge(edges, 2, 3);
  }
  return edges;
}

function buildForkReconvergeEdges(count) {
  if (count < 5) return buildSpineBranchEdges(count);

  const edges = [];
  addGraphEdge(edges, 0, 1);
  addGraphEdge(edges, 1, 2);
  addGraphEdge(edges, 1, 3);
  addGraphEdge(edges, 2, 4);
  addGraphEdge(edges, 3, 4);
  if (count >= 6) {
    addGraphEdge(edges, 4, 5);
  }
  return edges;
}

function buildGauntletPocketEdges(count) {
  if (count < 5) return buildSpineEdges(count);

  const edges = [];
  addGraphEdge(edges, 0, 1);
  addGraphEdge(edges, 1, 2);
  addGraphEdge(edges, 2, 3);
  if (count === 5) {
    addGraphEdge(edges, 2, 4);
    return edges;
  }

  addGraphEdge(edges, 3, 4);
  addGraphEdge(edges, 2, 5);
  return edges;
}

function addGraphEdge(edges, a, b) {
  if (a === b) return;
  if (edges.some((edge) => (edge[0] === a && edge[1] === b) || (edge[0] === b && edge[1] === a))) {
    return;
  }
  edges.push([a, b]);
}

function buildGraphAdjacency(graph, nodeCount) {
  const adjacency = Array.from({ length: nodeCount }, () => new Set());
  graph.edges.forEach(([a, b]) => {
    if (a >= 0 && b >= 0 && a < nodeCount && b < nodeCount) {
      adjacency[a].add(b);
      adjacency[b].add(a);
    }
  });
  return adjacency;
}

function buildBfsOrder(adjacency, startNode, nodeCount) {
  const visited = new Set([startNode]);
  const order = [startNode];
  const queue = [startNode];

  while (queue.length) {
    const current = queue.shift();
    adjacency[current].forEach((neighbor) => {
      if (visited.has(neighbor)) return;
      visited.add(neighbor);
      order.push(neighbor);
      queue.push(neighbor);
    });
  }

  for (let index = 0; index < nodeCount; index += 1) {
    if (!visited.has(index)) order.push(index);
  }

  return order;
}

function chooseAreaCenters({ rows, cols, areaCount, adjacency, bfsOrder, rng }) {
  const minRow = 1;
  const maxRow = Math.max(minRow, rows - 2);
  const minCol = 1;
  const maxCol = Math.max(minCol, cols - 2);
  const coreRowMin = Math.max(minRow, Math.floor(rows / 2) - 1);
  const coreRowMax = Math.min(maxRow, Math.floor(rows / 2) + 1);
  const coreColMin = Math.max(minCol, Math.floor(cols / 2) - 1);
  const coreColMax = Math.min(maxCol, Math.floor(cols / 2) + 1);

  for (let attempt = 0; attempt < CENTER_PLACEMENT_ATTEMPTS; attempt += 1) {
    const centers = Array(areaCount).fill(null);
    const useCore = randomFloat(rng) < 0.78;
    const rootRow = useCore ? randomInt(coreRowMin, coreRowMax, rng) : randomInt(minRow, maxRow, rng);
    const rootCol = useCore ? randomInt(coreColMin, coreColMax, rng) : randomInt(minCol, maxCol, rng);
    centers[0] = toCellIndex(rootRow, rootCol, cols);

    let failed = false;

    for (const nodeIndex of bfsOrder) {
      if (nodeIndex === 0) continue;

      const placedNeighbors = Array.from(adjacency[nodeIndex]).filter((neighbor) => centers[neighbor] !== null);
      const parentIndex = placedNeighbors.length
        ? placedNeighbors[randomInt(0, placedNeighbors.length - 1, rng)]
        : 0;

      const parentCenter = centers[parentIndex];
      const candidates = getCenterCandidates(parentCenter, rows, cols, rng);

      let chosen = null;
      for (const candidate of candidates) {
        if (isValidCenterCandidate(candidate, centers, cols)) {
          chosen = candidate;
          break;
        }
      }

      if (chosen === null) {
        failed = true;
        break;
      }

      centers[nodeIndex] = chosen;
    }

    if (!failed) return centers;
  }

  return null;
}

function getCenterCandidates(parentIndex, rows, cols, rng) {
  const parentRow = Math.floor(parentIndex / cols);
  const parentCol = parentIndex % cols;

  const offsets = [];
  const distances = [2, 3, 4, 1];
  distances.forEach((distance) => {
    offsets.push([distance, 0], [-distance, 0], [0, distance], [0, -distance]);
    offsets.push([distance, distance], [distance, -distance], [-distance, distance], [-distance, -distance]);
  });

  shuffleInPlace(offsets, rng);

  const candidates = [];
  offsets.forEach(([rowOffset, colOffset]) => {
    const row = parentRow + rowOffset;
    const col = parentCol + colOffset;
    if (row < 1 || row > rows - 2 || col < 1 || col > cols - 2) return;
    candidates.push(toCellIndex(row, col, cols));
  });

  if (candidates.length < 8) {
    const fallback = [];
    for (let row = 1; row <= rows - 2; row += 1) {
      for (let col = 1; col <= cols - 2; col += 1) {
        fallback.push(toCellIndex(row, col, cols));
      }
    }
    shuffleInPlace(fallback, rng);
    fallback.forEach((value) => {
      if (!candidates.includes(value)) candidates.push(value);
    });
  }

  return candidates;
}

function isValidCenterCandidate(candidateIndex, centers, cols) {
  const row = Math.floor(candidateIndex / cols);
  const col = candidateIndex % cols;

  for (const existing of centers) {
    if (existing === null) continue;
    const existingRow = Math.floor(existing / cols);
    const existingCol = existing % cols;
    const manhattan = Math.abs(existingRow - row) + Math.abs(existingCol - col);
    if (manhattan < 2) return false;
  }

  return true;
}

function placeSingleAreaBlob({ areaIndex, targetSize, profile, center, ownerByCell, rows, cols, rng }) {
  const shapeArchetype = profile?.shapeArchetype || "blob";

  if (shapeArchetype !== "blob") {
    const templated = placeTemplateAreaShape({
      areaIndex,
      targetSize,
      center,
      ownerByCell,
      rows,
      cols,
      rng,
      shapeArchetype
    });

    if (templated) return templated;
  }

  for (let attempt = 0; attempt < 36; attempt += 1) {
    const startCell = pickAreaStartCell(center, areaIndex, ownerByCell, rows, cols, rng);
    if (startCell === null) continue;

    const shape = [startCell];
    ownerByCell[startCell] = areaIndex;

    let failed = false;

    while (shape.length < targetSize) {
      const frontierSet = new Set();
      shape.forEach((cellIndex) => {
        getNeighbors(cellIndex, rows, cols).forEach((neighborIndex) => {
          if (shape.includes(neighborIndex)) return;
          if (!canPlaceAreaCell(neighborIndex, areaIndex, ownerByCell, rows, cols)) return;
          frontierSet.add(neighborIndex);
        });
      });

      const frontier = Array.from(frontierSet);
      if (frontier.length === 0) {
        failed = true;
        break;
      }

      const centerRow = Math.floor(center / cols);
      const centerCol = center % cols;
      const weights = frontier.map((cellIndex) => {
        const row = Math.floor(cellIndex / cols);
        const col = cellIndex % cols;
        const dist = Math.abs(row - centerRow) + Math.abs(col - centerCol);
        const sameNeighbors = getNeighbors(cellIndex, rows, cols)
          .filter((neighborIndex) => ownerByCell[neighborIndex] === areaIndex)
          .length;
        let score = Math.max(0.01, (2.2 - (dist * 0.2)) + (sameNeighbors * 0.7));

        if (profile?.shapeKind === "corridor") {
          const rowsInShape = shape.map((value) => Math.floor(value / cols));
          const colsInShape = shape.map((value) => value % cols);
          const rowSpan = Math.max(...rowsInShape) - Math.min(...rowsInShape);
          const colSpan = Math.max(...colsInShape) - Math.min(...colsInShape);
          const extendLinearly = rowSpan >= colSpan
            ? Math.abs(col - centerCol) <= 1
            : Math.abs(row - centerRow) <= 1;
          if (extendLinearly) score += 1.1;
        }

        return score;
      });

      const nextCell = weightedPickFromValues(frontier, weights, rng);
      shape.push(nextCell);
      ownerByCell[nextCell] = areaIndex;
    }

    if (!failed && shape.length === targetSize) {
      return shape;
    }

    shape.forEach((cellIndex) => {
      ownerByCell[cellIndex] = -1;
    });
  }

  return null;
}

function placeTemplateAreaShape({ areaIndex, targetSize, center, ownerByCell, rows, cols, rng, shapeArchetype }) {
  const offsets = buildShapeOffsets(shapeArchetype, targetSize, rng);
  if (!offsets || offsets.length === 0) return null;

  const bounds = getOffsetBounds(offsets);
  const centerRow = Math.floor(center / cols);
  const centerCol = center % cols;
  const shapeCenterRow = (bounds.minRow + bounds.maxRow) / 2;
  const shapeCenterCol = (bounds.minCol + bounds.maxCol) / 2;
  const baseAnchorRow = Math.round(centerRow - shapeCenterRow);
  const baseAnchorCol = Math.round(centerCol - shapeCenterCol);

  const anchors = [];
  for (let rowOffset = -2; rowOffset <= 2; rowOffset += 1) {
    for (let colOffset = -2; colOffset <= 2; colOffset += 1) {
      anchors.push({ row: baseAnchorRow + rowOffset, col: baseAnchorCol + colOffset });
    }
  }
  for (let i = 0; i < 12; i += 1) {
    anchors.push({
      row: randomInt(0, rows - 1, rng),
      col: randomInt(0, cols - 1, rng)
    });
  }
  shuffleInPlace(anchors, rng);

  for (const anchor of anchors) {
    const placed = [];
    let valid = true;

    for (const offset of offsets) {
      const row = anchor.row + offset.row;
      const col = anchor.col + offset.col;
      if (row < 0 || row >= rows || col < 0 || col >= cols) {
        valid = false;
        break;
      }

      const index = toCellIndex(row, col, cols);
      if (!canPlaceAreaCell(index, areaIndex, ownerByCell, rows, cols)) {
        valid = false;
        break;
      }

      placed.push(index);
    }

    if (!valid) continue;

    placed.forEach((cellIndex) => {
      ownerByCell[cellIndex] = areaIndex;
    });
    return placed;
  }

  return null;
}

function buildShapeOffsets(shapeArchetype, targetSize, rng) {
  let baseOffsets = [];

  if (shapeArchetype === "rectangle") {
    baseOffsets = buildRectangleOffsets(targetSize, rng);
  } else if (shapeArchetype === "l_shape") {
    baseOffsets = buildLShapeOffsets(targetSize, rng);
  } else if (shapeArchetype === "t_shape") {
    baseOffsets = buildTShapeOffsets(targetSize, rng);
  } else if (shapeArchetype === "wide_hall") {
    baseOffsets = buildWideHallOffsets(targetSize, rng);
  } else if (shapeArchetype === "courtyard") {
    baseOffsets = buildCourtyardOffsets(targetSize, rng);
  } else if (shapeArchetype === "straight_corridor") {
    baseOffsets = buildStraightCorridorOffsets(targetSize, rng);
  } else if (shapeArchetype === "bent_corridor") {
    baseOffsets = buildBentCorridorOffsets(targetSize, rng);
  } else if (shapeArchetype === "short_choke") {
    baseOffsets = buildShortChokeOffsets(targetSize);
  } else {
    baseOffsets = buildBlobOffsets(targetSize, rng);
  }

  return normalizeOffsetsToTarget(baseOffsets, targetSize, rng);
}

function buildRectangleOffsets(targetSize, rng) {
  if (targetSize <= 2) {
    return Array.from({ length: targetSize }, (_, index) => ({ row: 0, col: index }));
  }

  let bestWidth = 2;
  let bestHeight = Math.ceil(targetSize / 2);
  let bestArea = bestWidth * bestHeight;

  for (let width = 2; width <= Math.min(6, targetSize); width += 1) {
    const height = Math.max(2, Math.ceil(targetSize / width));
    const area = width * height;
    if (area < bestArea) {
      bestArea = area;
      bestWidth = width;
      bestHeight = height;
      continue;
    }
    if (area === bestArea && Math.abs(width - height) < Math.abs(bestWidth - bestHeight)) {
      bestWidth = width;
      bestHeight = height;
    }
  }

  if (randomFloat(rng) < 0.35 && bestWidth !== bestHeight) {
    const swap = bestWidth;
    bestWidth = bestHeight;
    bestHeight = swap;
  }

  const offsets = [];
  for (let row = 0; row < bestHeight; row += 1) {
    for (let col = 0; col < bestWidth; col += 1) {
      offsets.push({ row, col });
    }
  }
  return offsets;
}

function buildLShapeOffsets(targetSize, rng) {
  if (targetSize <= 3) return buildRectangleOffsets(targetSize, rng);
  const armA = Math.max(2, Math.ceil((targetSize + 1) / 2));
  const armB = Math.max(2, (targetSize - armA) + 1);
  const offsets = [];

  for (let row = 0; row < armA; row += 1) offsets.push({ row, col: 0 });
  for (let col = 0; col < armB; col += 1) offsets.push({ row: 0, col });
  return offsets;
}

function buildTShapeOffsets(targetSize, rng) {
  if (targetSize <= 4) return buildLShapeOffsets(targetSize, rng);

  const topWidth = Math.max(3, Math.min(5, targetSize - 1));
  const stemLength = Math.max(2, (targetSize - topWidth) + 1);
  const centerCol = Math.floor(topWidth / 2);
  const offsets = [];

  for (let col = 0; col < topWidth; col += 1) offsets.push({ row: 0, col });
  for (let row = 1; row < stemLength; row += 1) offsets.push({ row, col: centerCol });
  return offsets;
}

function buildBlobOffsets(targetSize, rng) {
  const offsets = [{ row: 0, col: 0 }];
  const taken = new Set(["0,0"]);

  while (offsets.length < targetSize) {
    const source = offsets[randomInt(0, offsets.length - 1, rng)];
    const neighbors = [
      { row: source.row - 1, col: source.col },
      { row: source.row + 1, col: source.col },
      { row: source.row, col: source.col - 1 },
      { row: source.row, col: source.col + 1 }
    ].filter((cell) => !taken.has(`${cell.row},${cell.col}`));

    if (neighbors.length === 0) continue;
    const pick = neighbors[randomInt(0, neighbors.length - 1, rng)];
    taken.add(`${pick.row},${pick.col}`);
    offsets.push(pick);
  }

  return offsets;
}

function buildWideHallOffsets(targetSize, rng) {
  if (targetSize <= 3) return buildRectangleOffsets(targetSize, rng);

  const length = Math.max(2, Math.ceil(targetSize / 2));
  const offsets = [];
  for (let col = 0; col < length; col += 1) {
    offsets.push({ row: 0, col });
    offsets.push({ row: 1, col });
  }

  if (targetSize % 2 !== 0) {
    offsets.push({ row: 2, col: Math.floor(length / 2) });
  }

  return offsets;
}

function buildCourtyardOffsets(targetSize, rng) {
  if (targetSize < 8) return buildRectangleOffsets(targetSize, rng);

  const offsets = [];
  for (let row = 0; row < 3; row += 1) {
    for (let col = 0; col < 3; col += 1) {
      if (row === 1 && col === 1) continue;
      offsets.push({ row, col });
    }
  }

  if (targetSize > 8) {
    const expansion = [
      { row: -1, col: 1 },
      { row: 1, col: -1 },
      { row: 1, col: 3 },
      { row: 3, col: 1 },
      { row: -1, col: 0 },
      { row: -1, col: 2 },
      { row: 3, col: 0 },
      { row: 3, col: 2 }
    ];
    shuffleInPlace(expansion, rng);
    expansion.slice(0, targetSize - 8).forEach((item) => offsets.push(item));
  }

  return offsets;
}

function buildStraightCorridorOffsets(targetSize, rng) {
  const horizontal = randomFloat(rng) < 0.5;
  return Array.from({ length: targetSize }, (_, index) => (
    horizontal ? { row: 0, col: index } : { row: index, col: 0 }
  ));
}

function buildBentCorridorOffsets(targetSize, rng) {
  if (targetSize <= 2) return buildStraightCorridorOffsets(targetSize, rng);
  const armA = Math.max(2, Math.ceil((targetSize + 1) / 2));
  const armB = Math.max(2, (targetSize - armA) + 1);
  const offsets = [];

  for (let i = 0; i < armA; i += 1) offsets.push({ row: i, col: 0 });
  for (let i = 0; i < armB; i += 1) offsets.push({ row: 0, col: i });
  return offsets;
}

function buildShortChokeOffsets(targetSize) {
  if (targetSize <= 2) return Array.from({ length: targetSize }, (_, index) => ({ row: 0, col: index }));
  if (targetSize === 3) return [{ row: 0, col: 0 }, { row: 0, col: 1 }, { row: 0, col: 2 }];
  if (targetSize === 4) return [{ row: 0, col: 0 }, { row: 0, col: 1 }, { row: 0, col: 2 }, { row: 1, col: 1 }];
  return [{ row: 0, col: 0 }, { row: 0, col: 1 }, { row: 0, col: 2 }, { row: 1, col: 1 }, { row: -1, col: 1 }];
}

function normalizeOffsetsToTarget(offsets, targetSize, rng) {
  const unique = new Map();
  offsets.forEach((offset) => {
    unique.set(`${offset.row},${offset.col}`, { row: offset.row, col: offset.col });
  });

  while (unique.size < targetSize) {
    const cells = Array.from(unique.values());
    const frontier = [];
    const seen = new Set();
    cells.forEach((cell) => {
      getOffsetOrthNeighbors(cell).forEach((neighbor) => {
        const key = `${neighbor.row},${neighbor.col}`;
        if (unique.has(key) || seen.has(key)) return;
        seen.add(key);
        frontier.push(neighbor);
      });
    });
    if (frontier.length === 0) break;
    const pick = frontier[randomInt(0, frontier.length - 1, rng)];
    unique.set(`${pick.row},${pick.col}`, pick);
  }

  while (unique.size > targetSize) {
    const cells = Array.from(unique.values());
    const removable = cells.filter((cell) => {
      const neighbors = getOffsetOrthNeighbors(cell)
        .filter((neighbor) => unique.has(`${neighbor.row},${neighbor.col}`));
      return neighbors.length <= 1;
    });
    const pool = removable.length > 0 ? removable : cells;
    const pick = pool[randomInt(0, pool.length - 1, rng)];
    unique.delete(`${pick.row},${pick.col}`);

    const current = Array.from(unique.values());
    if (current.length > 0 && !isOffsetShapeConnected(current)) {
      unique.set(`${pick.row},${pick.col}`, pick);
      break;
    }
  }

  let normalized = Array.from(unique.values());
  if (normalized.length > targetSize) {
    normalized = extractConnectedSubset(normalized, targetSize);
  }
  if (normalized.length < targetSize) {
    normalized = buildBlobOffsets(targetSize, rng);
  }

  if (!isOffsetShapeConnected(normalized)) {
    normalized = extractConnectedSubset(normalized, Math.min(targetSize, normalized.length));
  }

  return normalized;
}

function extractConnectedSubset(offsets, targetSize) {
  if (offsets.length <= targetSize) return offsets;
  const set = new Set(offsets.map((offset) => `${offset.row},${offset.col}`));
  const sorted = [...offsets].sort((a, b) => {
    const distA = Math.abs(a.row) + Math.abs(a.col);
    const distB = Math.abs(b.row) + Math.abs(b.col);
    return distA - distB;
  });
  const start = sorted[0];
  const queue = [start];
  const picked = new Map([[`${start.row},${start.col}`, start]]);

  while (queue.length && picked.size < targetSize) {
    const current = queue.shift();
    getOffsetOrthNeighbors(current).forEach((neighbor) => {
      const key = `${neighbor.row},${neighbor.col}`;
      if (!set.has(key) || picked.has(key)) return;
      picked.set(key, neighbor);
      queue.push(neighbor);
    });
  }

  return Array.from(picked.values());
}

function getOffsetOrthNeighbors(cell) {
  return [
    { row: cell.row - 1, col: cell.col },
    { row: cell.row + 1, col: cell.col },
    { row: cell.row, col: cell.col - 1 },
    { row: cell.row, col: cell.col + 1 }
  ];
}

function isOffsetShapeConnected(offsets) {
  if (offsets.length <= 1) return true;
  const set = new Set(offsets.map((offset) => `${offset.row},${offset.col}`));
  const queue = [offsets[0]];
  const visited = new Set([`${offsets[0].row},${offsets[0].col}`]);

  while (queue.length) {
    const current = queue.shift();
    getOffsetOrthNeighbors(current).forEach((neighbor) => {
      const key = `${neighbor.row},${neighbor.col}`;
      if (!set.has(key) || visited.has(key)) return;
      visited.add(key);
      queue.push(neighbor);
    });
  }

  return visited.size === offsets.length;
}

function getOffsetBounds(offsets) {
  const rows = offsets.map((item) => item.row);
  const cols = offsets.map((item) => item.col);
  return {
    minRow: Math.min(...rows),
    maxRow: Math.max(...rows),
    minCol: Math.min(...cols),
    maxCol: Math.max(...cols)
  };
}

function pickAreaStartCell(center, areaIndex, ownerByCell, rows, cols, rng) {
  const centerRow = Math.floor(center / cols);
  const centerCol = center % cols;

  const candidates = [center];
  for (let radius = 1; radius <= 2; radius += 1) {
    for (let rowOffset = -radius; rowOffset <= radius; rowOffset += 1) {
      for (let colOffset = -radius; colOffset <= radius; colOffset += 1) {
        if ((Math.abs(rowOffset) + Math.abs(colOffset)) > radius) continue;
        const row = centerRow + rowOffset;
        const col = centerCol + colOffset;
        if (row < 0 || row >= rows || col < 0 || col >= cols) continue;
        const index = toCellIndex(row, col, cols);
        if (!candidates.includes(index)) candidates.push(index);
      }
    }
  }

  shuffleInPlace(candidates, rng);

  for (const candidate of candidates) {
    if (canPlaceAreaCell(candidate, areaIndex, ownerByCell, rows, cols)) {
      return candidate;
    }
  }

  return null;
}

function canPlaceAreaCell(cellIndex, areaIndex, ownerByCell, rows, cols) {
  if (ownerByCell[cellIndex] !== -1 && ownerByCell[cellIndex] !== areaIndex) return false;

  const neighbor8 = getNeighbor8(cellIndex, rows, cols);
  for (const neighborIndex of neighbor8) {
    const owner = ownerByCell[neighborIndex];
    if (owner !== -1 && owner !== areaIndex) return false;
  }

  return true;
}

function getNeighbor8(index, rows, cols) {
  const row = Math.floor(index / cols);
  const col = index % cols;
  const neighbors = [];

  for (let rowOffset = -1; rowOffset <= 1; rowOffset += 1) {
    for (let colOffset = -1; colOffset <= 1; colOffset += 1) {
      if (rowOffset === 0 && colOffset === 0) continue;
      const nextRow = row + rowOffset;
      const nextCol = col + colOffset;
      if (nextRow < 0 || nextRow >= rows || nextCol < 0 || nextCol >= cols) continue;
      neighbors.push(toCellIndex(nextRow, nextCol, cols));
    }
  }

  return neighbors;
}

function connectGraphWithCorridors({ graph, areaCellsByIndex, ownerByCell, rows, cols, rng }) {
  const corridorOccupied = new Set();
  const corridors = [];

  for (let edgeIndex = 0; edgeIndex < graph.edges.length; edgeIndex += 1) {
    const [fromArea, toArea] = graph.edges[edgeIndex];

    if (areAreasAdjacent(areaCellsByIndex[fromArea], toArea, ownerByCell, rows, cols)) {
      corridors.push({
        id: `C${edgeIndex + 1}`,
        from: fromArea,
        to: toArea,
        cells: [],
        shapeArchetype: "short_choke"
      });
      continue;
    }

    const path = findCorridorPath({
      fromArea,
      toArea,
      areaCellsByIndex,
      ownerByCell,
      corridorOccupied,
      rows,
      cols
    });

    if (!path) {
      return null;
    }

    path.forEach((cellIndex) => corridorOccupied.add(cellIndex));
    corridors.push({
      id: `C${edgeIndex + 1}`,
      from: fromArea,
      to: toArea,
      cells: path,
      shapeArchetype: classifyCorridorPathShape(path, cols)
    });
  }

  return { corridors };
}

function classifyCorridorPathShape(path, cols) {
  if (!Array.isArray(path) || path.length <= 1) return "short_choke";
  if (path.length <= 2) return "short_choke";

  const rows = path.map((index) => Math.floor(index / cols));
  const colsInPath = path.map((index) => index % cols);
  const uniqueRows = new Set(rows).size;
  const uniqueCols = new Set(colsInPath).size;

  if (uniqueRows === 1 || uniqueCols === 1) return "straight_corridor";
  return "bent_corridor";
}

function areAreasAdjacent(areaCells, targetAreaIndex, ownerByCell, rows, cols) {
  for (const cellIndex of areaCells) {
    const touching = getNeighbors(cellIndex, rows, cols)
      .some((neighborIndex) => ownerByCell[neighborIndex] === targetAreaIndex);
    if (touching) return true;
  }
  return false;
}

function getAreaBoundaryNeighbors(areaCells, ownerByCell, corridorOccupied, rows, cols) {
  const boundary = new Set();

  areaCells.forEach((cellIndex) => {
    getNeighbors(cellIndex, rows, cols).forEach((neighborIndex) => {
      if (ownerByCell[neighborIndex] !== -1) return;
      if (corridorOccupied.has(neighborIndex)) return;
      boundary.add(neighborIndex);
    });
  });

  return Array.from(boundary);
}

function findCorridorPath({ fromArea, toArea, areaCellsByIndex, ownerByCell, corridorOccupied, rows, cols }) {
  const starts = getAreaBoundaryNeighbors(
    areaCellsByIndex[fromArea],
    ownerByCell,
    corridorOccupied,
    rows,
    cols
  );

  if (starts.length === 0) return null;

  const queue = [];
  const visited = new Set();

  starts.forEach((startCell) => {
    if (!corridorCellAllowed(startCell, fromArea, toArea, ownerByCell, corridorOccupied, rows, cols)) {
      return;
    }
    queue.push({ cell: startCell, path: [startCell] });
    visited.add(startCell);
  });

  while (queue.length) {
    const { cell, path } = queue.shift();

    const touchesTarget = getNeighbors(cell, rows, cols)
      .some((neighborIndex) => ownerByCell[neighborIndex] === toArea);
    if (touchesTarget && path.length >= 1 && path.length <= 2) {
      return path;
    }

    if (path.length >= 2) continue;

    getNeighbors(cell, rows, cols).forEach((neighborIndex) => {
      if (visited.has(neighborIndex)) return;
      if (!corridorCellAllowed(neighborIndex, fromArea, toArea, ownerByCell, corridorOccupied, rows, cols)) return;
      visited.add(neighborIndex);
      queue.push({ cell: neighborIndex, path: [...path, neighborIndex] });
    });
  }

  return null;
}

function corridorCellAllowed(cellIndex, fromArea, toArea, ownerByCell, corridorOccupied, rows, cols) {
  if (ownerByCell[cellIndex] !== -1) return false;
  if (corridorOccupied.has(cellIndex)) return false;

  const orthNeighbors = getNeighbors(cellIndex, rows, cols);
  for (const neighborIndex of orthNeighbors) {
    const owner = ownerByCell[neighborIndex];
    if (owner !== -1 && owner !== fromArea && owner !== toArea) return false;
    if (corridorOccupied.has(neighborIndex)) return false;
  }

  const neighbors8 = getNeighbor8(cellIndex, rows, cols);
  for (const neighborIndex of neighbors8) {
    const owner = ownerByCell[neighborIndex];
    if (owner !== -1 && owner !== fromArea && owner !== toArea) return false;
  }

  return true;
}

function buildMapFromSparsePlan({
  preset,
  seed = "",
  plan,
  graph,
  areaSizes,
  areaProfiles = null,
  dominantTiles,
  allowVariation,
  rng,
  previousGateStates
}) {
  const rows = preset.rows;
  const cols = preset.cols;
  const totalCells = rows * cols;

  const cells = Array.from({ length: totalCells }, (_, index) => ({
    index,
    row: Math.floor(index / cols),
    col: index % cols,
    areaId: "",
    tileId: "",
    tileName: "",
    tileImage: "",
    biome: "",
    tags: [],
    isVariation: false,
    isCorridor: false,
    corridorId: "",
    dominantTileId: ""
  }));

  const normalizedProfiles = areaProfiles
    ? areaProfiles.map((profile) => ({ ...profile }))
    : areaSizes.map((size, index) => ({
      index,
      size,
      category: getAreaSizeInfo(size).category,
      shapeArchetype: "blob",
      shapeKind: "room"
    }));

  const areas = plan.areaCellsByIndex.map((areaCells, areaIndex) => {
    const dominantTile = dominantTiles[areaIndex] || weightedPick(state.zones.tiles, rng) || {
      id: "FallbackTile",
      name: "Fallback Tile",
      biome: "unknown",
      tags: [],
      weight: 1
    };
    const profile = normalizedProfiles[areaIndex] || {
      index: areaIndex,
      size: areaSizes[areaIndex],
      category: getAreaSizeInfo(areaSizes[areaIndex]).category,
      shapeArchetype: "blob",
      shapeKind: "room"
    };
    const sizeInfo = getAreaSizeInfo(areaSizes[areaIndex]);
    const areaId = toAreaId(areaIndex);

    areaCells.forEach((cellIndex) => {
      const cell = cells[cellIndex];
      cell.areaId = areaId;
      cell.tileId = dominantTile.id;
      cell.tileName = dominantTile.name;
      cell.tileImage = dominantTile.image || "";
      cell.biome = dominantTile.biome;
      cell.tags = [...dominantTile.tags];
      cell.dominantTileId = dominantTile.id;
    });

    return {
      id: areaId,
      index: areaIndex,
      size: areaSizes[areaIndex],
      tileCount: areaSizes[areaIndex],
      sizeCategory: sizeInfo.category,
      cardsReturnedOnClear: sizeInfo.cardsReturned,
      shapeArchetype: profile.shapeArchetype || "blob",
      shapeKind: profile.shapeKind || "room",
      cells: [...areaCells],
      dominantTileId: dominantTile.id,
      dominantTileName: dominantTile.name,
      dominantTileImage: dominantTile.image || "",
      supportingTileIds: [],
      supportingTileNames: [],
      biome: dominantTile.biome,
      tags: [...dominantTile.tags],
      role: {
        start: false,
        objective: false,
        objectiveRank: 0,
        extraction: false,
        pressure: false
      },
      phase: ""
    };
  });

  const graphAdjacency = buildAreaAdjacencyFromGraph(graph, areas.length);
  const roleMeta = assignAreaRoles(areas, graphAdjacency, rng, { cols, rows });
  applyNarrativeTileComposition({
    areas,
    cells,
    cols,
    graphAdjacency,
    graph,
    roleMeta,
    dominantTiles,
    allowVariation,
    rng
  });

  const corridors = plan.corridors.map((corridor) => {
    const fromArea = areas[corridor.from];
    const toArea = areas[corridor.to];
    const corridorTile = pickCorridorTile(fromArea, toArea, rng);

    corridor.cells.forEach((cellIndex) => {
      const cell = cells[cellIndex];
      cell.areaId = "";
      cell.isCorridor = true;
      cell.corridorId = corridor.id;
      cell.tileId = corridorTile.id;
      cell.tileName = corridorTile.name;
      cell.tileImage = corridorTile.image || "";
      cell.biome = corridorTile.biome;
      cell.tags = [...new Set([...(corridorTile.tags || []), "corridor"])];
      cell.isVariation = false;
      cell.dominantTileId = "";
    });

    return {
      id: corridor.id,
      fromAreaId: fromArea.id,
      toAreaId: toArea.id,
      fromAreaIndex: corridor.from,
      toAreaIndex: corridor.to,
      shapeArchetype: corridor.shapeArchetype || classifyCorridorPathShape(corridor.cells, cols),
      tileId: corridorTile.id,
      tileName: corridorTile.name,
      tileImage: corridorTile.image || "",
      cells: [...corridor.cells]
    };
  });
  const gates = buildSparseGates({
    cells,
    rows,
    cols,
    areas,
    corridors,
    graph,
    rng,
    previousGateStates
  });

  return {
    preset: preset.key,
    seed,
    rows,
    cols,
    cells,
    areas,
    corridors,
    graph: {
      template: graph.template,
      nodeCount: graph.nodeCount,
      edges: graph.edges.map(([a, b]) => [a, b])
    },
    plan: {
      areaSizes: [...areaSizes],
      areaProfiles: normalizedProfiles.map((profile) => ({ ...profile })),
      areaCellsByIndex: plan.areaCellsByIndex.map((areaCells) => [...areaCells]),
      corridors: plan.corridors.map((corridor) => ({ ...corridor, cells: [...corridor.cells] })),
      occupiedCount: plan.occupiedCount,
      graph: {
        template: graph.template,
        nodeCount: graph.nodeCount,
        edges: graph.edges.map(([a, b]) => [a, b])
      }
    },
    roleMeta,
    areaAdjacency: graphAdjacency,
    gates,
    settings: {
      mixedBiomes: isMixedBiomesEnabled(),
      allowVariationTiles: allowVariationTiles(),
      areaSizeBias: getAreaSizeBias(),
      sparseCanvas: true
    }
  };
}

function applyNarrativeTileComposition({
  areas,
  cells,
  cols,
  graphAdjacency,
  graph,
  roleMeta,
  dominantTiles,
  allowVariation,
  rng
}) {
  if (!areas.length) return;

  const areaById = new Map(areas.map((area) => [area.id, area]));
  const centroids = computeAreaCentroids(areas, Math.max(1, cols || 1));
  const mainPathAreaIds = roleMeta?.mainPathAreaIds || [];
  const predecessorByAreaId = roleMeta?.predecessorByAreaId || {};
  const biasTags = getAreaBiasTags();

  assignAreaPhases(areas, graphAdjacency, mainPathAreaIds);

  let progressionOk = false;
  for (let attempt = 0; attempt < 6; attempt += 1) {
    areas.forEach((area, areaIndex) => {
      composeAreaTiles({
        area,
        areaIndex,
        areas,
        areaById,
        cells,
        cols,
        centroids,
        predecessorByAreaId,
        dominantTiles,
        biasTags,
        allowVariation,
        rng
      });
    });

    enforceMapTileDiversity({
      areas,
      cells,
      rng
    });

    progressionOk = evaluateMainPathProgression(areas, mainPathAreaIds);
    if (progressionOk) break;
  }

  if (!progressionOk) {
    logDevFallback("Fallback: progression check failed after composition rerolls; keeping best-effort tile narrative");
  }
}

function assignAreaPhases(areas, graphAdjacency, mainPathAreaIds) {
  const phaseByAreaId = {};
  const path = mainPathAreaIds.length > 0 ? [...mainPathAreaIds] : areas.map((area) => area.id);

  path.forEach((areaId, idx) => {
    const ratio = path.length <= 1 ? 1 : idx / (path.length - 1);
    let phase = "neutral";
    if (ratio > 0.34 && ratio <= 0.67) phase = "shift";
    if (ratio > 0.67) phase = "claimed";
    phaseByAreaId[areaId] = phase;
  });

  areas.forEach((area) => {
    if (phaseByAreaId[area.id]) {
      area.phase = phaseByAreaId[area.id];
      return;
    }

    const nearest = findNearestMainPathArea(area.id, path, graphAdjacency);
    const inherited = phaseByAreaId[nearest] || "shift";
    area.phase = inherited;
  });
}

function findNearestMainPathArea(areaId, mainPathAreaIds, graphAdjacency) {
  const queue = [areaId];
  const visited = new Set([areaId]);
  while (queue.length) {
    const current = queue.shift();
    if (mainPathAreaIds.includes(current)) return current;
    const neighbors = Array.from(graphAdjacency[current] || []);
    neighbors.forEach((neighbor) => {
      if (visited.has(neighbor)) return;
      visited.add(neighbor);
      queue.push(neighbor);
    });
  }
  return mainPathAreaIds[0] || areaId;
}

function composeAreaTiles({
  area,
  areaIndex,
  areas,
  areaById,
  cells,
  cols,
  centroids,
  predecessorByAreaId,
  dominantTiles,
  biasTags,
  allowVariation,
  rng
}) {
  const fallbackTile = weightedPick(state.zones.tiles, rng);
  const seededDominant = dominantTiles[areaIndex] || state.zones.tiles.find((tile) => tile.id === area.dominantTileId) || fallbackTile;
  const roleTags = getRoleBiasTags(area.role);
  const phaseTags = PHASE_TAG_WEIGHTS[area.phase] || [];
  const dominantSelection = pickTileWithRelaxedFallback({
    biomeHint: seededDominant?.biome || area.biome,
    requiredTags: phaseTags,
    biasTags: [...biasTags, ...roleTags, ...phaseTags],
    preferredId: seededDominant?.id || "",
    rng,
    contextLabel: `area ${area.id} dominant`
  });
  const dominantTile = dominantSelection.tile || seededDominant || fallbackTile;

  const targetSupportTypes = area.size >= 8 ? 2 : area.size >= 3 ? 1 : 0;
  const supportTiles = chooseSupportingTiles({
    area,
    dominantTile,
    desiredTypes: targetSupportTypes,
    biasTags: [...biasTags, ...roleTags],
    phaseTags,
    rng
  });

  const supportRatio = allowVariation
    ? weightedPickFromValues([0.22, 0.28, 0.34, 0.4], [2.3, 3.1, 2.0, 1.0], rng)
    : 0.2;
  const minSupportCells = area.size >= 3 ? 1 : 0;
  let supportCellCount = Math.max(minSupportCells, Math.round(area.size * supportRatio));
  supportCellCount = Math.min(supportCellCount, Math.max(0, area.size - 1));
  if (supportTiles.length === 0) supportCellCount = 0;

  const predecessorId = predecessorByAreaId[area.id] || "";
  const predecessorArea = predecessorId ? areaById.get(predecessorId) : null;
  const entranceCell = deriveEntranceCell(area, predecessorArea, centroids, cells);
  const depthMap = computeAreaDepthMap(area.cells, entranceCell, cells, cols);

  const cellsByDepth = [...area.cells].sort((a, b) => (depthMap.get(a) || 0) - (depthMap.get(b) || 0));
  const entranceCut = Math.max(1, Math.ceil(cellsByDepth.length * 0.4));
  const middleCut = Math.max(entranceCut, Math.ceil(cellsByDepth.length * 0.75));

  const supportPriority = [
    ...cellsByDepth.slice(middleCut).reverse(),
    ...cellsByDepth.slice(entranceCut, middleCut).reverse(),
    ...cellsByDepth.slice(0, entranceCut).reverse()
  ];

  const supportAssignments = new Map();
  let supportAssigned = 0;
  for (const cellIndex of supportPriority) {
    if (supportAssigned >= supportCellCount) break;
    if (supportTiles.length === 0) break;
    const depth = depthMap.get(cellIndex) || 0;
    const depthZone = depth <= (depthMap.size <= 1 ? 0 : 1) ? "entrance"
      : (depth >= 3 ? "deep" : "middle");
    const supportTile = chooseSupportTileForDepth({
      supportTiles,
      depthZone,
      areaPhase: area.phase,
      role: area.role,
      rng
    });
    if (!supportTile) continue;
    supportAssignments.set(cellIndex, supportTile);
    supportAssigned += 1;
  }

  area.cells.forEach((cellIndex) => {
    const tile = supportAssignments.get(cellIndex) || dominantTile;
    const cell = cells[cellIndex];
    cell.tileId = tile.id;
    cell.tileName = tile.name;
    cell.tileImage = tile.image || "";
    cell.biome = tile.biome;
    cell.tags = [...tile.tags];
    cell.dominantTileId = dominantTile.id;
    cell.isVariation = tile.id !== dominantTile.id;
  });

  if (area.size >= 3 && !areaHasVariation(area, cells)) {
    const fallbackVariation = chooseSupportingTiles({
      area,
      dominantTile,
      desiredTypes: 1,
      biasTags: [...biasTags, ...phaseTags],
      phaseTags,
      rng
    })[0];
    if (fallbackVariation) {
      const deepestCell = cellsByDepth[cellsByDepth.length - 1];
      const cell = cells[deepestCell];
      cell.tileId = fallbackVariation.id;
      cell.tileName = fallbackVariation.name;
      cell.tileImage = fallbackVariation.image || "";
      cell.biome = fallbackVariation.biome;
      cell.tags = [...fallbackVariation.tags];
      cell.dominantTileId = dominantTile.id;
      cell.isVariation = true;
    }
  }

  updateAreaTileMetadata(area, cells, dominantTile);
}

function getRoleBiasTags(role = {}) {
  const tags = [];
  if (role.objective) tags.push("objective", "spawn", "machine");
  if (role.start) tags.push("holdout", "urban", "transit");
  if (role.pressure) tags.push("hazard", "spawn", "machine");
  if (role.extraction) tags.push("extraction", "outdoor", "urban");
  return tags;
}

function deriveEntranceCell(area, predecessorArea, centroids, cells) {
  const defaultCell = area.cells[0];
  if (!defaultCell && defaultCell !== 0) return 0;

  let target = null;
  if (predecessorArea) {
    target = centroids[predecessorArea.id] || null;
  } else if (area.role?.start) {
    const areaCentroid = centroids[area.id] || { row: 0, col: 0 };
    target = {
      row: areaCentroid.row < (Math.max(...cells.map((cell) => cell.row)) / 2) ? -1 : (Math.max(...cells.map((cell) => cell.row)) + 1),
      col: areaCentroid.col < (Math.max(...cells.map((cell) => cell.col)) / 2) ? -1 : (Math.max(...cells.map((cell) => cell.col)) + 1)
    };
  }

  if (!target) return defaultCell;

  let bestCell = defaultCell;
  let bestDistance = Number.POSITIVE_INFINITY;
  area.cells.forEach((cellIndex) => {
    const cell = cells[cellIndex];
    const distance = Math.abs(cell.row - target.row) + Math.abs(cell.col - target.col);
    if (distance < bestDistance) {
      bestDistance = distance;
      bestCell = cellIndex;
    }
  });
  return bestCell;
}

function computeAreaDepthMap(areaCellIndexes, entranceCellIndex, cells, cols) {
  const areaSet = new Set(areaCellIndexes);
  const depthMap = new Map();
  const queue = [entranceCellIndex];
  depthMap.set(entranceCellIndex, 0);
  const rows = Math.max(1, Math.floor(cells.length / Math.max(1, cols)));

  while (queue.length) {
    const current = queue.shift();
    const currentDepth = depthMap.get(current) || 0;
    const neighbors = getNeighbors(current, rows, cols);
    neighbors.forEach((neighborIndex) => {
      if (!areaSet.has(neighborIndex)) return;
      if (depthMap.has(neighborIndex)) return;
      depthMap.set(neighborIndex, currentDepth + 1);
      queue.push(neighborIndex);
    });
  }

  areaCellIndexes.forEach((cellIndex) => {
    if (!depthMap.has(cellIndex)) depthMap.set(cellIndex, 0);
  });
  return depthMap;
}

function chooseSupportingTiles({ area, dominantTile, desiredTypes, biasTags, phaseTags, rng }) {
  if (desiredTypes <= 0) return [];
  const support = [];
  for (let i = 0; i < desiredTypes; i += 1) {
    const picked = pickTileWithRelaxedFallback({
      biomeHint: dominantTile.biome || area.biome,
      requiredTags: phaseTags,
      biasTags: [...biasTags, ...PHASE_TAG_WEIGHTS.shift],
      excludedIds: new Set([dominantTile.id, ...support.map((tile) => tile.id)]),
      rng,
      contextLabel: `area ${area.id} support`
    });
    if (!picked.tile) continue;
    if (picked.tile.id === dominantTile.id) continue;
    if (support.some((tile) => tile.id === picked.tile.id)) continue;
    support.push(picked.tile);
  }

  if (support.length < desiredTypes) {
    const fallbackPool = state.zones.tiles.filter((tile) => tile.id !== dominantTile.id && !support.some((item) => item.id === tile.id));
    while (support.length < desiredTypes && fallbackPool.length > 0) {
      const pick = weightedPick(fallbackPool, rng);
      support.push(pick);
      const idx = fallbackPool.findIndex((tile) => tile.id === pick.id);
      if (idx >= 0) fallbackPool.splice(idx, 1);
    }
  }

  return support;
}

function chooseSupportTileForDepth({ supportTiles, depthZone, areaPhase, role, rng }) {
  if (!supportTiles.length) return null;
  const phase = DEPTH_PHASE_MAP[depthZone] || areaPhase || "shift";
  const targetTags = new Set([
    ...(PHASE_TAG_WEIGHTS[phase] || []),
    ...getRoleBiasTags(role)
  ]);
  const weights = supportTiles.map((tile) => {
    const overlap = tile.tags.filter((tag) => targetTags.has(tag)).length;
    return Math.max(0.2, 1 + (overlap * 0.8));
  });
  return weightedPickFromValues(supportTiles, weights, rng);
}

function areaHasVariation(area, cells) {
  const tileIds = new Set(area.cells.map((cellIndex) => cells[cellIndex].tileId));
  return tileIds.size > 1;
}

function updateAreaTileMetadata(area, cells, dominantTile) {
  const usage = new Map();
  area.cells.forEach((cellIndex) => {
    const cell = cells[cellIndex];
    usage.set(cell.tileId, (usage.get(cell.tileId) || 0) + 1);
  });

  let dominantId = dominantTile?.id || area.dominantTileId;
  let dominantCount = usage.get(dominantId) || 0;
  usage.forEach((count, tileId) => {
    if (count > dominantCount) {
      dominantId = tileId;
      dominantCount = count;
    }
  });

  const dominant = state.zones.tiles.find((tile) => tile.id === dominantId) || dominantTile || state.zones.tiles[0];
  const supporting = Array.from(usage.entries())
    .filter(([tileId]) => tileId !== dominant.id)
    .sort((a, b) => b[1] - a[1])
    .map(([tileId]) => state.zones.tiles.find((tile) => tile.id === tileId))
    .filter(Boolean);

  const mergedTags = new Set();
  [dominant, ...supporting].forEach((tile) => {
    (tile.tags || []).forEach((tag) => mergedTags.add(tag));
  });

  area.dominantTileId = dominant.id;
  area.dominantTileName = dominant.name;
  area.dominantTileImage = dominant.image || "";
  area.biome = dominant.biome;
  area.tags = Array.from(mergedTags);
  area.supportingTileIds = supporting.map((tile) => tile.id);
  area.supportingTileNames = supporting.map((tile) => tile.name);
}

function pickTileWithRelaxedFallback({
  biomeHint,
  requiredTags = [],
  biasTags = [],
  preferredId = "",
  excludedIds = new Set(),
  rng,
  contextLabel = "tile"
}) {
  const allTiles = state.zones.tiles || [];
  if (!allTiles.length) {
    return {
      tile: {
        id: "FallbackTile",
        name: "Fallback Tile",
        biome: "unknown",
        tags: [],
        weight: 1
      },
      fallbackLevel: 3
    };
  }

  const requiredTagList = Array.isArray(requiredTags) ? requiredTags : [];
  const matchesRequiredTags = (tile) => {
    if (!requiredTagList.length) return true;
    return tile.tags.some((tag) => requiredTagList.includes(tag));
  };

  const noExcluded = (tile) => !excludedIds.has(tile.id);

  const attempt1 = allTiles.filter((tile) => (!biomeHint || tile.biome === biomeHint) && matchesRequiredTags(tile) && noExcluded(tile));
  if (attempt1.length > 0) {
    return {
      tile: pickWeightedNarrativeTile(attempt1, biasTags, preferredId, rng),
      fallbackLevel: 1
    };
  }

  const attempt2 = allTiles.filter((tile) => (!biomeHint || tile.biome === biomeHint) && noExcluded(tile));
  if (attempt2.length > 0) {
    logDevFallback(`Fallback: tag+biome failed, using biome-only (${contextLabel})`);
    return {
      tile: pickWeightedNarrativeTile(attempt2, biasTags, preferredId, rng),
      fallbackLevel: 2
    };
  }

  const attempt3 = allTiles.filter(noExcluded);
  if (attempt3.length > 0) {
    logDevFallback(`Fallback: biome-only failed, using any-tile (${contextLabel})`);
    return {
      tile: pickWeightedNarrativeTile(attempt3, biasTags, preferredId, rng),
      fallbackLevel: 3
    };
  }

  logDevFallback(`Fallback: exclusions exhausted pool, using any tile (${contextLabel})`);
  return {
    tile: pickWeightedNarrativeTile(allTiles, biasTags, preferredId, rng),
    fallbackLevel: 3
  };
}

function pickWeightedNarrativeTile(pool, biasTags, preferredId, rng) {
  const weights = pool.map((tile) => {
    let score = Number(tile.weight || 1);
    const overlap = tile.tags.filter((tag) => biasTags.includes(tag)).length;
    score += overlap * 0.65;
    if (preferredId && tile.id === preferredId) score += 0.9;
    return Math.max(0.01, score);
  });
  return weightedPickFromValues(pool, weights, rng);
}

function evaluateMainPathProgression(areas, mainPathAreaIds) {
  if (!mainPathAreaIds || mainPathAreaIds.length <= 1) return true;

  const areaById = new Map(areas.map((area) => [area.id, area]));
  const stageByArea = mainPathAreaIds.map((areaId) => {
    const area = areaById.get(areaId);
    if (!area) return 0;
    const tags = new Set(area.tags || []);
    const neutralScore = PHASE_TAG_WEIGHTS.neutral.filter((tag) => tags.has(tag)).length;
    const shiftScore = PHASE_TAG_WEIGHTS.shift.filter((tag) => tags.has(tag)).length;
    const claimedScore = PHASE_TAG_WEIGHTS.claimed.filter((tag) => tags.has(tag)).length + (area.role?.objective ? 1.5 : 0);

    if (claimedScore >= shiftScore && claimedScore >= neutralScore) return 2;
    if (shiftScore >= neutralScore) return 1;
    return 0;
  });

  const first = stageByArea[0];
  const last = stageByArea[stageByArea.length - 1];
  const hasTransition = stageByArea.some((value, idx) => idx > 0 && value !== stageByArea[idx - 1]);
  return hasTransition && last >= 1 && last >= first;
}

function enforceMapTileDiversity({ areas, cells, rng }) {
  const areaCellIndexes = areas.flatMap((area) => area.cells);
  if (areaCellIndexes.length === 0) return;

  const usage = new Map();
  areaCellIndexes.forEach((cellIndex) => {
    const tileId = cells[cellIndex].tileId;
    usage.set(tileId, (usage.get(tileId) || 0) + 1);
  });

  const softCap = Math.ceil(areaCellIndexes.length * TILE_SHARE_SOFT_CAP);
  let overCap = Array.from(usage.entries()).filter(([, count]) => count > softCap);

  let rebalanceGuard = 0;
  while (overCap.length > 0 && rebalanceGuard < 480) {
    rebalanceGuard += 1;
    const [overTileId, overCount] = overCap.sort((a, b) => b[1] - a[1])[0];
    if (overCount <= softCap) break;

    const candidateAreas = areas.filter((area) => area.cells.some((cellIndex) => cells[cellIndex].tileId === overTileId));
    if (!candidateAreas.length) break;
    const area = candidateAreas[randomInt(0, candidateAreas.length - 1, rng)];
    const candidateCell = [...area.cells]
      .filter((cellIndex) => cells[cellIndex].tileId === overTileId)
      .sort((a, b) => (cells[b].isVariation ? 1 : 0) - (cells[a].isVariation ? 1 : 0))[0];
    if (candidateCell === undefined) break;

    const alternatives = state.zones.tiles.filter((tile) => tile.id !== overTileId);
    if (!alternatives.length) break;
    const replacement = weightedPick(alternatives, rng);
    cells[candidateCell].tileId = replacement.id;
    cells[candidateCell].tileName = replacement.name;
    cells[candidateCell].tileImage = replacement.image || "";
    cells[candidateCell].biome = replacement.biome;
    cells[candidateCell].tags = [...replacement.tags];
    cells[candidateCell].isVariation = replacement.id !== area.dominantTileId;

    usage.set(overTileId, Math.max(0, (usage.get(overTileId) || 0) - 1));
    usage.set(replacement.id, (usage.get(replacement.id) || 0) + 1);
    overCap = Array.from(usage.entries()).filter(([, count]) => count > softCap);
  }

  const hardFallbackCap = Math.ceil(areaCellIndexes.length * TILE_SHARE_FALLBACK_CAP);
  const maxCount = Math.max(...Array.from(usage.values()));
  if (maxCount > softCap && maxCount <= hardFallbackCap) {
    logDevFallback("Fallback: tile share cap relaxed from 35% to 50% due limited pool");
  }

  if (maxCount > hardFallbackCap && state.zones.tiles.length > 1) {
    logDevFallback("Fallback: forcing emergency tile diversity to avoid single-type collapse");
    const dominantTileId = Array.from(usage.entries()).sort((a, b) => b[1] - a[1])[0][0];
    const replacements = state.zones.tiles.filter((tile) => tile.id !== dominantTileId);
    const forcedReplacements = Math.max(1, Math.ceil((maxCount - hardFallbackCap) / 2));
    const eligibleCells = areaCellIndexes.filter((cellIndex) => cells[cellIndex].tileId === dominantTileId);
    shuffleInPlace(eligibleCells, rng);
    for (let i = 0; i < Math.min(forcedReplacements, eligibleCells.length); i += 1) {
      const cellIndex = eligibleCells[i];
      const replacement = weightedPick(replacements, rng);
      cells[cellIndex].tileId = replacement.id;
      cells[cellIndex].tileName = replacement.name;
      cells[cellIndex].tileImage = replacement.image || "";
      cells[cellIndex].biome = replacement.biome;
      cells[cellIndex].tags = [...replacement.tags];
      cells[cellIndex].isVariation = true;
    }
  }

  areas.forEach((area) => {
    const dominant = state.zones.tiles.find((tile) => tile.id === area.dominantTileId) || state.zones.tiles[0];
    updateAreaTileMetadata(area, cells, dominant);
  });
}

function pickCorridorTile(fromArea, toArea, rng) {
  const allTiles = state.zones.tiles;
  const biomeHints = new Set([fromArea?.biome, toArea?.biome].filter(Boolean));
  const corridorTagHints = new Set(["corridor", "transit", "sewer", "machine", "industrial", "urban"]);

  const attempt1 = allTiles.filter((tile) => biomeHints.has(tile.biome) && tile.tags.some((tag) => corridorTagHints.has(tag)));
  if (attempt1.length > 0) return weightedPick(attempt1, rng);

  const attempt2 = allTiles.filter((tile) => biomeHints.has(tile.biome));
  if (attempt2.length > 0) return weightedPick(attempt2, rng);

  const attempt3 = allTiles.filter((tile) => tile.tags.some((tag) => corridorTagHints.has(tag)));
  if (attempt3.length > 0) return weightedPick(attempt3, rng);

  return weightedPick(allTiles, rng);
}

function pickDominantTiles(areaCount, rng, mixedBiomes, avoidIds = [], areaProfiles = [], graph = null) {
  const allTiles = state.zones.tiles;
  const result = [];
  const fallbackStats = {
    level1: 0,
    level2: 0,
    level3: 0
  };

  let globalBiome = "";
  if (!mixedBiomes) {
    const biomeWeights = new Map();
    allTiles.forEach((tile) => {
      biomeWeights.set(tile.biome, (biomeWeights.get(tile.biome) || 0) + Number(tile.weight || 1));
    });
    globalBiome = weightedPickFromValues(Array.from(biomeWeights.keys()), Array.from(biomeWeights.values()), rng);
  }

  const requiredTags = getAreaRequiredTags();
  const missionBiasTags = getAreaBiasTags();
  const usedIds = new Set();
  const areaPhaseHints = buildAreaPhaseHintsForDominants(areaCount, graph);

  for (let i = 0; i < areaCount; i += 1) {
    const avoidId = avoidIds[i] || "";
    const profile = areaProfiles[i] || {};
    const phaseHint = areaPhaseHints[i] || "neutral";
    const profileTags = profile.shapeKind === "corridor"
      ? ["transit", "machine"]
      : (phaseHint === "claimed" ? ["objective", "machine", "spawn"]
        : phaseHint === "shift" ? ["hazard", "spawn", "sewer"]
          : ["urban", "transit", "indoor"]);
    const biasTags = [...missionBiasTags, ...profileTags];
    const pickedResult = pickTileWithFallback({
      allTiles,
      requiredTags,
      biasTags,
      biomeHint: globalBiome,
      avoidId,
      usedIds,
      rng
    });

    fallbackStats[`level${pickedResult.fallbackLevel}`] += 1;
    usedIds.add(pickedResult.tile.id);
    let picked = pickedResult.tile;

    if (avoidId && allTiles.length > 1 && picked.id === avoidId) {
      const withoutAvoid = allTiles.filter((tile) => tile.id !== avoidId);
      if (withoutAvoid.length > 0) picked = weightedPick(withoutAvoid, rng);
    }

    result.push(picked);
  }

  if (fallbackStats.level2 > 0) {
    logDevFallback("Fallback: tag+biome failed, using biome-only");
  }
  if (fallbackStats.level3 > 0) {
    logDevFallback("Fallback: biome-only failed, using any-tile");
  }

  if (POC_MODE && fallbackStats.level3 > 0) {
    enforceDominantTileDiversity(result, allTiles, rng);
  }

  return result;
}

function buildAreaPhaseHintsForDominants(areaCount, graph) {
  if (!graph || !Array.isArray(graph.edges)) {
    return Array.from({ length: areaCount }, (_, idx) => (idx === areaCount - 1 ? "claimed" : "neutral"));
  }

  const adjacency = buildGraphAdjacency(graph, graph.nodeCount);
  const startNode = 0;
  const distances = Array(areaCount).fill(Number.POSITIVE_INFINITY);
  const queue = [startNode];
  distances[startNode] = 0;

  while (queue.length) {
    const current = queue.shift();
    const nextDistance = distances[current] + 1;
    adjacency[current].forEach((neighbor) => {
      if (nextDistance >= distances[neighbor]) return;
      distances[neighbor] = nextDistance;
      queue.push(neighbor);
    });
  }

  const maxDistance = Math.max(...distances.filter((value) => Number.isFinite(value)), 1);
  return distances.map((distance) => {
    const ratio = maxDistance <= 0 ? 0 : distance / maxDistance;
    if (ratio > 0.67) return "claimed";
    if (ratio > 0.34) return "shift";
    return "neutral";
  });
}

function getAreaRequiredTags() {
  const zoneTagSet = new Set();
  state.zones.tiles.forEach((tile) => {
    tile.tags.forEach((tag) => zoneTagSet.add(tag));
  });

  const sourceTags = new Set(state.selectedTags);
  const filtered = Array.from(sourceTags).filter((tag) => zoneTagSet.has(tag));
  if (filtered.length <= 3) return filtered;

  return filtered.slice(0, 3);
}

function getAreaBiasTags() {
  const zoneTagSet = new Set();
  state.zones.tiles.forEach((tile) => {
    tile.tags.forEach((tag) => zoneTagSet.add(tag));
  });

  const sourceTags = new Set(state.selectedTags);
  if (state.mission) {
    collectMissionTags(state.mission).forEach((tag) => sourceTags.add(tag));
  }

  const filtered = Array.from(sourceTags).filter((tag) => zoneTagSet.has(tag));
  if (filtered.length <= 6) return filtered;

  return filtered.slice(0, 6);
}

function pickTileWithFallback({ allTiles, requiredTags, biasTags, biomeHint, avoidId, usedIds, rng }) {
  const hasRequiredTags = (tile) => {
    if (!requiredTags || requiredTags.length === 0) return true;
    return tile.tags.some((tag) => requiredTags.includes(tag));
  };

  const attempt1 = allTiles.filter((tile) => (!biomeHint || tile.biome === biomeHint) && hasRequiredTags(tile));
  if (attempt1.length > 0) {
    return { tile: pickWeightedTileFromPool(attempt1, biasTags, usedIds, avoidId, rng), fallbackLevel: 1 };
  }

  const attempt2 = allTiles.filter((tile) => !biomeHint || tile.biome === biomeHint);
  if (attempt2.length > 0) {
    return { tile: pickWeightedTileFromPool(attempt2, biasTags, usedIds, avoidId, rng), fallbackLevel: 2 };
  }

  const attempt3 = allTiles.length > 0 ? allTiles : [];
  if (attempt3.length > 0) {
    return { tile: pickWeightedTileFromPool(attempt3, biasTags, usedIds, avoidId, rng), fallbackLevel: 3 };
  }

  return {
    tile: {
      id: "FallbackTile",
      name: "Fallback Tile",
      biome: "unknown",
      tags: [],
      weight: 1
    },
    fallbackLevel: 3
  };
}

function pickWeightedTileFromPool(pool, biasTags, usedIds, avoidId, rng) {
  const weights = pool.map((tile) => {
    let score = Number(tile.weight || 1);

    if (biasTags && biasTags.length > 0) {
      const overlap = tile.tags.filter((tag) => biasTags.includes(tag)).length;
      score += overlap * 0.65;
    }

    if (usedIds && usedIds.size > 0 && !usedIds.has(tile.id)) {
      score += 0.25;
    }

    if (avoidId && tile.id === avoidId) {
      score *= 0.2;
    }

    return Math.max(0.01, score);
  });

  return weightedPickFromValues(pool, weights, rng);
}

function enforceDominantTileDiversity(resultTiles, allTiles, rng) {
  if (resultTiles.length <= 1 || allTiles.length <= 1) return;

  const unique = new Set(resultTiles.map((tile) => tile.id));
  if (unique.size >= 2) return;

  const baseId = resultTiles[0]?.id;
  const alternatives = allTiles.filter((tile) => tile.id !== baseId);
  if (alternatives.length === 0) return;

  const replaceIndex = randomInt(1, resultTiles.length - 1, rng);
  resultTiles[replaceIndex] = weightedPick(alternatives, rng);
  logDevFallback("Fallback: any-tile mode diversified dominant tiles to avoid single-type collapse");
}

function logDevFallback(message) {
  if (!POC_MODE) return;
  console.debug(message);
}

function buildAreaAdjacencyFromGraph(graph, areaCount) {
  const adjacency = {};
  for (let index = 0; index < areaCount; index += 1) {
    adjacency[toAreaId(index)] = new Set();
  }

  graph.edges.forEach(([from, to]) => {
    const fromId = toAreaId(from);
    const toId = toAreaId(to);
    adjacency[fromId].add(toId);
    adjacency[toId].add(fromId);
  });

  return adjacency;
}

function buildSparseGates({ cells, rows, cols, areas, corridors, graph, rng, previousGateStates }) {
  const gates = {};
  const bridgeEdgeSet = buildBridgeEdgeSet(graph.nodeCount, graph.edges);

  graph.edges.forEach(([fromAreaIndex, toAreaIndex]) => {
    const fromArea = areas[fromAreaIndex];
    const toArea = areas[toAreaIndex];
    if (!fromArea || !toArea) return;

    const gatePair = findGatePairForConnection({
      fromAreaIndex,
      toAreaIndex,
      areas,
      corridors,
      rows,
      cols
    });

    if (!gatePair) return;

    const key = edgeKey(gatePair.a, gatePair.b);
    const graphEdgeKey = areaEdgeKey(fromAreaIndex, toAreaIndex);
    const previousState = previousGateStates?.[key]?.state;
    const defaultState = previousState || pickDefaultGateState({
      isBridge: bridgeEdgeSet.has(graphEdgeKey),
      rng
    });

    gates[key] = {
      key,
      a: gatePair.a,
      b: gatePair.b,
      state: defaultState,
      orientation: gatePair.orientation || determineGateOrientation(gatePair.a, gatePair.b, cols),
      anchorIndex: gatePair.anchorIndex ?? gatePair.a,
      areaAId: fromArea.id,
      areaBId: toArea.id,
      graphEdgeKey
    };
  });

  return gates;
}

function findGatePairForConnection({ fromAreaIndex, toAreaIndex, areas, corridors, rows, cols }) {
  const corridor = corridors.find((item) => (
    (item.fromAreaIndex === fromAreaIndex && item.toAreaIndex === toAreaIndex)
    || (item.fromAreaIndex === toAreaIndex && item.toAreaIndex === fromAreaIndex)
  ));

  if (corridor && corridor.cells.length > 0) {
    const normalized = (corridor.fromAreaIndex === fromAreaIndex)
      ? corridor
      : {
        ...corridor,
        fromAreaIndex: corridor.toAreaIndex,
        toAreaIndex: corridor.fromAreaIndex,
        cells: [...corridor.cells].reverse()
      };

    const firstCorridorCell = normalized.cells[0];
    const fromTouch = findAreaNeighboringCell(areas[fromAreaIndex], firstCorridorCell, rows, cols);
    if (fromTouch !== null) {
      return {
        a: fromTouch,
        b: firstCorridorCell,
        orientation: determineGateOrientation(fromTouch, firstCorridorCell, cols),
        anchorIndex: fromTouch
      };
    }

    const lastCorridorCell = normalized.cells[normalized.cells.length - 1];
    const toTouch = findAreaNeighboringCell(areas[toAreaIndex], lastCorridorCell, rows, cols);
    if (toTouch !== null) {
      return {
        a: lastCorridorCell,
        b: toTouch,
        orientation: determineGateOrientation(lastCorridorCell, toTouch, cols),
        anchorIndex: lastCorridorCell
      };
    }
  }

  const direct = findTouchingPairBetweenAreas(areas[fromAreaIndex], areas[toAreaIndex], rows, cols);
  if (direct) {
    return {
      ...direct,
      orientation: determineGateOrientation(direct.a, direct.b, cols),
      anchorIndex: direct.a
    };
  }

  return null;
}

function findAreaNeighboringCell(area, targetCellIndex, rows, cols) {
  if (!area) return null;
  const areaSet = new Set(area.cells || []);
  const neighbors = getNeighbors(targetCellIndex, rows, cols);
  for (const neighborIndex of neighbors) {
    if (areaSet.has(neighborIndex)) return neighborIndex;
  }
  return null;
}

function findTouchingPairBetweenAreas(areaA, areaB, rows, cols) {
  if (!areaA || !areaB) return null;
  const areaBSet = new Set(areaB.cells || []);
  for (const cellIndex of areaA.cells || []) {
    const neighbors = getNeighbors(cellIndex, rows, cols);
    for (const neighborIndex of neighbors) {
      if (areaBSet.has(neighborIndex)) {
        return { a: cellIndex, b: neighborIndex };
      }
    }
  }
  return null;
}

function determineGateOrientation(a, b, cols) {
  const colA = a % cols;
  const colB = b % cols;
  return colA !== colB ? "vertical" : "horizontal";
}

function pickDefaultGateState({ isBridge, rng }) {
  if (isBridge) return "open";
  return weightedPickFromValues(["open", "locked", "blocked"], [70, 20, 10], rng);
}

function buildBridgeEdgeSet(nodeCount, edges) {
  const adjacency = Array.from({ length: nodeCount }, () => []);
  edges.forEach(([a, b], idx) => {
    adjacency[a].push({ node: b, edgeIndex: idx });
    adjacency[b].push({ node: a, edgeIndex: idx });
  });

  const visited = Array(nodeCount).fill(false);
  const disc = Array(nodeCount).fill(0);
  const low = Array(nodeCount).fill(0);
  const bridgeEdgeSet = new Set();
  let time = 0;

  const dfs = (u, parentEdgeIndex) => {
    visited[u] = true;
    time += 1;
    disc[u] = time;
    low[u] = time;

    adjacency[u].forEach(({ node: v, edgeIndex }) => {
      if (edgeIndex === parentEdgeIndex) return;
      if (!visited[v]) {
        dfs(v, edgeIndex);
        low[u] = Math.min(low[u], low[v]);
        if (low[v] > disc[u]) {
          const [a, b] = edges[edgeIndex];
          bridgeEdgeSet.add(areaEdgeKey(a, b));
        }
      } else {
        low[u] = Math.min(low[u], disc[v]);
      }
    });
  };

  for (let i = 0; i < nodeCount; i += 1) {
    if (!visited[i]) dfs(i, -1);
  }

  return bridgeEdgeSet;
}

function cycleGateState(gateKey) {
  if (!state.map || !state.map.gates[gateKey]) return;
  const gate = state.map.gates[gateKey];
  const currentIndex = GATE_STATES.indexOf(gate.state);
  const nextState = GATE_STATES[(currentIndex + 1) % GATE_STATES.length];
  gate.state = nextState;

  renderMap();
  attachMapToMission();
  if (state.mission) renderMission();
  setMapStatus(`Gate ${gate.areaAId}↔${gate.areaBId} set to ${nextState}.`);
}

function bindMapCellInteractions(cellEl, cell) {
  cellEl.dataset.cellIndex = String(cell.index);

  if (cell.tileId) {
    cellEl.classList.add("draggable");
    cellEl.draggable = true;
    cellEl.addEventListener("dragstart", onMapCellDragStart);
    cellEl.addEventListener("dragend", onMapCellDragEnd);
  }

  cellEl.addEventListener("contextmenu", (event) => {
    event.preventDefault();
    event.stopPropagation();
    openMapCardMenuForCell(cell.index, event.clientX, event.clientY);
  });

  cellEl.addEventListener("dragover", onMapCellDragOver);
  cellEl.addEventListener("dragleave", onMapCellDragLeave);
  cellEl.addEventListener("drop", onMapCellDrop);
}

function onMapCellDragStart(event) {
  const cellEl = event.currentTarget;
  const sourceIndex = Number(cellEl.dataset.cellIndex);
  const sourceCell = state.map?.cells?.[sourceIndex];
  if (!sourceCell || !sourceCell.tileId) {
    event.preventDefault();
    return;
  }

  closeMapCardMenu();
  clearMapDragVisualState();
  state.mapUi.dragSourceIndex = sourceIndex;
  cellEl.classList.add("drag-source");
  requestAnimationFrame(() => {
    cellEl.classList.add("dragging");
  });

  if (event.dataTransfer) {
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", String(sourceIndex));
  }
}

function onMapCellDragOver(event) {
  const targetEl = event.currentTarget;
  const targetIndex = Number(targetEl.dataset.cellIndex);
  const sourceIndex = state.mapUi.dragSourceIndex;
  if (!Number.isInteger(sourceIndex) || !Number.isInteger(targetIndex) || sourceIndex === targetIndex) return;

  const sourceCell = state.map?.cells?.[sourceIndex];
  const targetCell = state.map?.cells?.[targetIndex];
  if (!sourceCell || !sourceCell.tileId || !targetCell) return;

  event.preventDefault();
  targetEl.classList.add("drag-over");
  if (event.dataTransfer) {
    event.dataTransfer.dropEffect = "move";
  }
}

function onMapCellDragLeave(event) {
  const targetEl = event.currentTarget;
  targetEl.classList.remove("drag-over");
}

function onMapCellDrop(event) {
  const targetEl = event.currentTarget;
  const targetIndex = Number(targetEl.dataset.cellIndex);
  const dragDataRaw = event.dataTransfer?.getData("text/plain") || "";
  const dragData = /^\d+$/.test(dragDataRaw) ? Number(dragDataRaw) : NaN;
  const sourceIndex = Number.isInteger(state.mapUi.dragSourceIndex)
    ? state.mapUi.dragSourceIndex
    : (Number.isInteger(dragData) ? dragData : NaN);

  clearMapDragVisualState();
  state.mapUi.dragSourceIndex = null;
  targetEl.classList.remove("drag-over");

  if (!Number.isInteger(sourceIndex) || !Number.isInteger(targetIndex) || sourceIndex === targetIndex) return;
  event.preventDefault();
  swapOrMoveMapCellTile(sourceIndex, targetIndex);
}

function onMapCellDragEnd() {
  clearMapDragVisualState();
  state.mapUi.dragSourceIndex = null;
}

function clearMapDragVisualState() {
  refs.mapGrid.querySelectorAll(".map-cell.dragging, .map-cell.drag-source, .map-cell.drag-over")
    .forEach((el) => {
      el.classList.remove("dragging", "drag-source", "drag-over");
    });
}

function extractEditableTilePayload(cell) {
  const payload = {};
  EDITABLE_TILE_FIELDS.forEach((field) => {
    payload[field] = cell[field];
  });
  payload.tags = Array.isArray(cell.tags) ? [...cell.tags] : [];
  return payload;
}

function applyEditableTilePayload(cell, payload) {
  EDITABLE_TILE_FIELDS.forEach((field) => {
    const value = payload[field];
    if (typeof value === "string") {
      cell[field] = value;
    } else if (typeof value === "boolean") {
      cell[field] = value;
    } else {
      cell[field] = field === "isVariation" ? false : "";
    }
  });
  cell.tags = Array.isArray(payload.tags) ? [...payload.tags] : [];
}

function swapOrMoveMapCellTile(sourceIndex, targetIndex) {
  if (!state.map) return;
  const source = state.map.cells[sourceIndex];
  const target = state.map.cells[targetIndex];
  if (!source || !target || !source.tileId) return;

  const sourcePayload = extractEditableTilePayload(source);
  if (target.tileId) {
    const targetPayload = extractEditableTilePayload(target);
    applyEditableTilePayload(source, targetPayload);
    applyEditableTilePayload(target, sourcePayload);

    attachMapToMission();
    renderMap();
    if (state.mission) renderMission();
    setMapStatus(`Swapped ${sourcePayload.tileName} with ${targetPayload.tileName}.`);
    return;
  }

  applyEditableTilePayload(target, sourcePayload);
  applyEditableTilePayload(source, {});

  attachMapToMission();
  renderMap();
  if (state.mission) renderMission();
  setMapStatus(`Moved ${sourcePayload.tileName} to row ${target.row + 1}, col ${target.col + 1}.`);
}

function setTileForCell(cellIndex, tileId) {
  if (!state.map) return;
  const cell = state.map.cells[cellIndex];
  if (!cell) return;

  const tile = findEditableRoomTileById(tileId);
  if (!tile) return;
  const previousName = cell.tileName || "";
  const hadTile = !!cell.tileId;

  applyEditableTilePayload(cell, {
    tileId: tile.id,
    tileName: tile.name,
    tileImage: tile.image || "",
    biome: tile.biome,
    isVariation: false,
    dominantTileId: tile.id,
    tags: [...tile.tags]
  });

  attachMapToMission();
  renderMap();
  if (state.mission) renderMission();
  if (hadTile) {
    setMapStatus(`Replaced ${previousName} with ${tile.name}.`);
  } else {
    setMapStatus(`Added ${tile.name} at row ${cell.row + 1}, col ${cell.col + 1}.`);
  }
}

function updateCardLabelAndTags(cellIndex, nextNameInput, tagInput) {
  if (!state.map) return;
  const cell = state.map.cells[cellIndex];
  if (!cell || !cell.tileId) return;

  const nextName = String(nextNameInput || "").trim();
  if (!nextName) {
    setMapStatus("Card name cannot be empty.");
    return;
  }

  cell.tileName = nextName;
  cell.tags = parseTagInput(tagInput);

  attachMapToMission();
  renderMap();
  if (state.mission) renderMission();
  setMapStatus(`Updated card name to "${nextName}".`);
}

function clearTileFromCell(cellIndex) {
  if (!state.map) return;
  const cell = state.map.cells[cellIndex];
  if (!cell || !cell.tileId) return;

  const previousName = cell.tileName || "tile";
  applyEditableTilePayload(cell, {});

  attachMapToMission();
  renderMap();
  if (state.mission) renderMission();
  setMapStatus(`Removed ${previousName} from row ${cell.row + 1}, col ${cell.col + 1}.`);
}

function openMapCardMenuForCell(cellIndex, clientX, clientY) {
  if (!state.map || !refs.mapCardMenu || !refs.mapCardMenuList) return;
  const cell = state.map.cells[cellIndex];
  if (!cell) return;

  state.mapUi.menuTargetIndex = cellIndex;
  renderMapCardMenuOptions(cell);

  refs.mapGrid.querySelectorAll(".map-cell.menu-target")
    .forEach((el) => el.classList.remove("menu-target"));
  const targetEl = refs.mapGrid.querySelector(`.map-cell[data-cell-index="${cellIndex}"]`);
  if (targetEl) targetEl.classList.add("menu-target");

  refs.mapCardMenu.classList.remove("hidden");
  refs.mapCardMenu.style.left = "0px";
  refs.mapCardMenu.style.top = "0px";

  const menuRect = refs.mapCardMenu.getBoundingClientRect();
  const left = Math.max(8, Math.min(clientX + 8, window.innerWidth - menuRect.width - 8));
  const top = Math.max(8, Math.min(clientY + 8, window.innerHeight - menuRect.height - 8));

  refs.mapCardMenu.style.left = `${left}px`;
  refs.mapCardMenu.style.top = `${top}px`;
}

function closeMapCardMenu() {
  if (!refs.mapCardMenu) return;
  refs.mapCardMenu.classList.add("hidden");
  state.mapUi.menuTargetIndex = null;
  refs.mapGrid.querySelectorAll(".map-cell.menu-target")
    .forEach((el) => el.classList.remove("menu-target"));
}

function renderMapCardMenuOptions(targetCell = null) {
  if (!refs.mapCardMenuList) return;
  refs.mapCardMenuList.innerHTML = "";
  if (refs.mapCardMenuTitle) {
    refs.mapCardMenuTitle.textContent = targetCell?.tileId ? "Edit Card and Room Names" : "Add Card and Room Names";
  }

  const tools = document.createElement("section");
  tools.className = "map-card-menu-tools";

  if (targetCell?.tileId) {
    const editTitle = document.createElement("p");
    editTitle.className = "map-card-menu-subtitle";
    editTitle.textContent = "Selected Card";

    const renameInput = document.createElement("input");
    renameInput.type = "text";
    renameInput.className = "map-card-menu-input map-card-rename-input";
    renameInput.value = targetCell.tileName || "";
    renameInput.placeholder = "Card name";

    const renameTagsInput = document.createElement("input");
    renameTagsInput.type = "text";
    renameTagsInput.className = "map-card-menu-input map-card-rename-tags-input";
    renameTagsInput.value = Array.isArray(targetCell.tags) ? targetCell.tags.join(", ") : "";
    renameTagsInput.placeholder = "tags (comma-separated)";

    const cardActions = document.createElement("div");
    cardActions.className = "map-card-menu-tool-actions";

    const renameButton = document.createElement("button");
    renameButton.type = "button";
    renameButton.className = "map-card-menu-option";
    renameButton.dataset.menuAction = "rename-card";
    renameButton.textContent = "Save Card Name/Tags";

    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className = "map-card-menu-option map-card-menu-danger";
    removeButton.dataset.menuAction = "remove";
    removeButton.textContent = "Remove Card";
    removeButton.title = `Remove ${targetCell.tileName || "card"}`;

    cardActions.appendChild(renameButton);
    cardActions.appendChild(removeButton);

    tools.appendChild(editTitle);
    tools.appendChild(renameInput);
    tools.appendChild(renameTagsInput);
    tools.appendChild(cardActions);
  }

  const libraryTitle = document.createElement("p");
  libraryTitle.className = "map-card-menu-subtitle";
  libraryTitle.textContent = "Room Name Library";

  const newNameInput = document.createElement("input");
  newNameInput.type = "text";
  newNameInput.className = "map-card-menu-input map-card-new-name-input";
  newNameInput.placeholder = "New room name";

  const newTagsInput = document.createElement("input");
  newTagsInput.type = "text";
  newTagsInput.className = "map-card-menu-input map-card-new-tags-input";
  newTagsInput.placeholder = "tags (comma-separated)";

  const biomeSelect = document.createElement("select");
  biomeSelect.className = "map-card-menu-select map-card-new-biome-select";
  const biomeOptions = getRoomLibraryBiomes();
  biomeOptions.forEach((biome) => {
    const option = document.createElement("option");
    option.value = biome;
    option.textContent = biome.replace(/[_-]+/g, " ").replace(/\b\w/g, (ch) => ch.toUpperCase());
    biomeSelect.appendChild(option);
  });

  const libraryActions = document.createElement("div");
  libraryActions.className = "map-card-menu-tool-actions";

  const addNameButton = document.createElement("button");
  addNameButton.type = "button";
  addNameButton.className = "map-card-menu-option";
  addNameButton.dataset.menuAction = "add-room-name";
  addNameButton.textContent = "Add Room Name";

  const resetButton = document.createElement("button");
  resetButton.type = "button";
  resetButton.className = "map-card-menu-option map-card-menu-secondary";
  resetButton.dataset.menuAction = "reset-room-list";
  resetButton.textContent = "Reset Picker";

  libraryActions.appendChild(addNameButton);
  libraryActions.appendChild(resetButton);

  tools.appendChild(libraryTitle);
  tools.appendChild(newNameInput);
  tools.appendChild(newTagsInput);
  tools.appendChild(biomeSelect);
  tools.appendChild(libraryActions);
  refs.mapCardMenuList.appendChild(tools);

  const divider = document.createElement("div");
  divider.className = "map-card-menu-divider";
  refs.mapCardMenuList.appendChild(divider);

  const tiles = getEditableRoomTiles()
    .sort((a, b) => a.biome.localeCompare(b.biome) || a.name.localeCompare(b.name));
  if (tiles.length === 0) {
    const emptyMessage = document.createElement("p");
    emptyMessage.className = "map-card-menu-empty";
    emptyMessage.textContent = "No room names available. Add one above.";
    refs.mapCardMenuList.appendChild(emptyMessage);
    return;
  }

  let currentBiome = "";
  let currentGroup = null;

  tiles.forEach((tile) => {
    if (tile.biome !== currentBiome) {
      currentBiome = tile.biome;
      currentGroup = document.createElement("section");
      currentGroup.className = "map-card-menu-group";

      const label = document.createElement("p");
      label.className = "map-card-menu-group-label";
      label.textContent = tile.biome
        .replace(/[_-]+/g, " ")
        .replace(/\b\w/g, (ch) => ch.toUpperCase());
      currentGroup.appendChild(label);
      refs.mapCardMenuList.appendChild(currentGroup);
    }

    const optionRow = document.createElement("div");
    optionRow.className = "map-card-menu-option-row";

    const option = document.createElement("button");
    option.type = "button";
    option.className = "map-card-menu-option map-card-menu-option-main";
    option.dataset.tileId = tile.id;
    option.textContent = tile.isCustom ? `${tile.name} (Custom)` : tile.name;
    option.title = `${tile.name} (${tile.biome})`;

    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.className = "map-card-menu-option-delete";
    deleteButton.dataset.menuAction = "delete-room-name";
    deleteButton.dataset.tileId = tile.id;
    deleteButton.textContent = "Delete";
    deleteButton.title = `Delete room name: ${tile.name}`;

    optionRow.appendChild(option);
    optionRow.appendChild(deleteButton);
    currentGroup.appendChild(optionRow);
  });
}

function renderMap() {
  closeMapCardMenu();
  clearMapDragVisualState();
  state.mapUi.dragSourceIndex = null;

  refs.mapGrid.innerHTML = "";
  refs.areaLegend.innerHTML = "";

  if (!state.map) {
    refs.mapGrid.textContent = "No map generated yet.";
    refs.areaLegend.textContent = "No areas yet.";
    return;
  }

  refs.mapGrid.style.setProperty("--map-cols", String(state.map.cols));

  const areaColors = createAreaColorMap(state.map.areas);
  const areaById = new Map(state.map.areas.map((area) => [area.id, area]));
  const cellElements = {};

  state.map.cells.forEach((cell) => {
    const cellEl = document.createElement("article");
    cellEl.className = "map-cell";
    bindMapCellInteractions(cellEl, cell);

    if (!cell.tileId) {
      cellEl.classList.add("empty");
      cellEl.style.setProperty("--area-color", "#17222d");
      refs.mapGrid.appendChild(cellEl);
      cellElements[cell.index] = cellEl;
      return;
    }

    if (cell.isCorridor) {
      cellEl.classList.add("corridor");
      cellEl.style.setProperty("--area-color", "#41586f");
    } else {
      const areaColor = areaColors[cell.areaId] || "#3a4f63";
      cellEl.style.setProperty("--area-color", areaColor);
    }

    const artUrl = toAssetUrl(cell.tileImage);
    if (artUrl) {
      cellEl.classList.add("has-art");
      cellEl.style.setProperty("--tile-image", `url("${artUrl}")`);
    }

    if (cell.isVariation) cellEl.classList.add("variation");

    const area = cell.areaId ? areaById.get(cell.areaId) : null;

    const badge = document.createElement("span");
    badge.className = `area-badge${cell.isCorridor ? " corridor-badge" : ""}`;
    badge.textContent = cell.isCorridor ? (cell.corridorId || "C") : (cell.areaId || "-");

    const tileName = document.createElement("span");
    tileName.className = "tile-name";
    tileName.textContent = cell.tileName;

    const tileMeta = document.createElement("span");
    tileMeta.className = "tile-meta";
    if (cell.isCorridor) {
      tileMeta.textContent = "corridor";
    } else {
      tileMeta.textContent = cell.isVariation ? "variation" : (cell.biome || "-");
    }

    const roleMeta = document.createElement("span");
    roleMeta.className = "tile-role";
    roleMeta.textContent = cell.isCorridor ? "" : mapAreaRoleLabel(area);

    cellEl.appendChild(badge);
    cellEl.appendChild(tileName);
    cellEl.appendChild(tileMeta);
    if (roleMeta.textContent) cellEl.appendChild(roleMeta);

    cellEl.title = `${cell.tileName}${cell.areaId ? ` | Area ${cell.areaId}` : ""}${cell.isCorridor ? ` | ${cell.corridorId}` : ""}`;

    refs.mapGrid.appendChild(cellEl);
    cellElements[cell.index] = cellEl;
  });

  Object.values(state.map.gates).forEach((gate) => {
    const anchor = cellElements[gate.anchorIndex];
    if (!anchor) return;

    const gateBtn = document.createElement("button");
    gateBtn.className = `gate-edge ${gate.orientation} state-${gate.state}`;
    gateBtn.type = "button";
    gateBtn.title = `Gate ${gate.state} (${gate.areaAId} ↔ ${gate.areaBId})`;
    gateBtn.textContent = gateStateGlyph(gate.state);
    gateBtn.addEventListener("click", (event) => {
      event.stopPropagation();
      cycleGateState(gate.key);
    });

    anchor.appendChild(gateBtn);
  });

  renderAreaLegend(areaColors);
}

function renderAreaLegend(areaColors) {
  if (!state.map) return;

  const areas = [...state.map.areas].sort((a, b) => a.id.localeCompare(b.id));
  if (areas.length === 0) {
    refs.areaLegend.textContent = "No areas generated.";
    return;
  }

  areas.forEach((area) => {
    const row = document.createElement("article");
    row.className = "legend-item";

    const swatch = document.createElement("span");
    swatch.className = "legend-swatch";
    swatch.style.background = areaColors[area.id] || "#2f3c49";

    const text = document.createElement("p");
    text.className = "legend-text";
    const roleLabel = mapAreaRoleLabel(area);
    const roleSuffix = roleLabel ? ` [${roleLabel}]` : "";
    const cardsLabel = formatCardsReturnedOnClear(area.cardsReturnedOnClear);
    const shapeLabel = formatShapeArchetype(area.shapeArchetype);
    const supportLabel = area.supportingTileNames?.length
      ? area.supportingTileNames.join(", ")
      : "none";
    text.textContent = `Area ${area.id}: ${area.dominantTileName}${roleSuffix} | ${area.size} tiles | ${area.sizeCategory} | ${cardsLabel} | ${shapeLabel} | support: ${supportLabel}`;

    row.appendChild(swatch);
    row.appendChild(text);
    refs.areaLegend.appendChild(row);
  });
}

function formatCardsReturnedOnClear(value) {
  if (value === undefined || value === null) return "cards returned: n/a";
  if (value === "all") return "all cards returned";
  if (value === 1) return "1 card returned";
  return `${value} cards returned`;
}

function formatShapeArchetype(shapeArchetype) {
  if (!shapeArchetype) return "shape: unknown";
  const label = shapeArchetype
    .replaceAll("_", " ")
    .replace(/\b\w/g, (ch) => ch.toUpperCase());
  return `shape: ${label}`;
}

function mapAreaRoleLabel(area) {
  if (!area || !area.role) return "";
  const labels = [];
  if (area.role.start) labels.push("START");
  if (area.role.objective) {
    labels.push(area.role.objectiveRank > 0 ? `OBJECTIVE ${area.role.objectiveRank}` : "OBJECTIVE");
  }
  if (area.role.pressure) labels.push("PRESSURE");
  if (area.role.extraction) labels.push("EXTRACTION");
  return labels.join(", ");
}

function gateStateGlyph(state) {
  if (state === "locked") return "L";
  if (state === "blocked") return "B";
  return "O";
}

function createAreaColorMap(areas) {
  const palette = [
    "#3f6b8e",
    "#7d5f2f",
    "#5b7b4d",
    "#6f4a7d",
    "#7a4f4f",
    "#3f7471",
    "#7b6f3f",
    "#4f5f8a",
    "#6a4e64",
    "#4e7a6a",
    "#7e5d44",
    "#5d5d82"
  ];

  const colorMap = {};
  areas.forEach((area, idx) => {
    colorMap[area.id] = palette[idx % palette.length];
  });
  return colorMap;
}

function assignAreaRoles(areas, adjacency, rng, layout = {}) {
  if (areas.length === 0) {
    return {
      startAreaId: "",
      objectiveAreaIds: [],
      mainPathAreaIds: [],
      predecessorByAreaId: {}
    };
  }

  const cols = Number(layout.cols || 1);
  const rows = Number(layout.rows || 7);
  const centroids = computeAreaCentroids(areas, cols);
  const areaById = new Map(areas.map((area) => [area.id, area]));
  const degreeById = {};
  areas.forEach((area) => {
    degreeById[area.id] = (adjacency[area.id] ? adjacency[area.id].size : 0);
  });

  areas.forEach((area) => {
    area.role = {
      start: false,
      objective: false,
      objectiveRank: 0,
      extraction: false,
      pressure: false
    };
  });

  const leaves = areas.filter((area) => (degreeById[area.id] || 0) <= 1);
  const startSource = leaves.length > 0 ? leaves : areas;
  const startArea = weightedPickFromValues(
    startSource,
    startSource.map((area) => {
      const centroid = centroids[area.id] || { row: 0, col: 0 };
      const edgeDistance = Math.min(
        centroid.row,
        Math.abs((rows - 1) - centroid.row),
        centroid.col,
        Math.abs((cols - 1) - centroid.col)
      );
      const degree = degreeById[area.id] || 0;
      let score = 1;
      if (degree <= 1) score += 1.8;
      score += 2 / (1 + edgeDistance);
      if (area.sizeCategory === "Tiny") score -= 0.25;
      return Math.max(0.1, score);
    }),
    rng
  );

  startArea.role.start = true;
  const traversal = computeDistancesAndParents(startArea.id, adjacency);
  const distanceFromStart = traversal.distances;
  const predecessorByAreaId = traversal.parents;
  const startCentroid = centroids[startArea.id] || { row: 0, col: 0 };
  const manhattanFromStart = {};
  Object.entries(centroids).forEach(([areaId, centroid]) => {
    manhattanFromStart[areaId] = Math.round(
      Math.abs(startCentroid.row - centroid.row) + Math.abs(startCentroid.col - centroid.col)
    );
  });

  const nonStart = areas.filter((area) => area.id !== startArea.id);
  if (nonStart.length === 0) {
    startArea.role.objective = true;
    startArea.role.objectiveRank = 1;
    return {
      startAreaId: startArea.id,
      objectiveAreaIds: [startArea.id],
      mainPathAreaIds: [startArea.id],
      predecessorByAreaId,
      distanceFromStart
    };
  }
  const objectivePool = nonStart.filter((area) => area.sizeCategory !== "Tiny");
  const objectiveSource = objectivePool.length > 0 ? objectivePool : nonStart;

  const maxGraphDistance = Math.max(...objectiveSource.map((area) => distanceFromStart[area.id] || 0));
  let farthest = objectiveSource.filter((area) => (distanceFromStart[area.id] || 0) === maxGraphDistance);
  const manhattanQualified = farthest.filter((area) => (manhattanFromStart[area.id] || 0) >= MIN_START_OBJECTIVE_DISTANCE);
  if (manhattanQualified.length > 0) {
    farthest = manhattanQualified;
  }
  if (farthest.length === 0) {
    farthest = objectiveSource;
  }

  const objectiveArea = weightedPickFromValues(
    farthest,
    farthest.map((area) => {
      const tags = new Set(area.tags || []);
      let score = 1;
      if (tags.has("objective")) score += 2.8;
      if (tags.has("spawn")) score += 1.4;
      if (tags.has("machine")) score += 1.0;
      score += (distanceFromStart[area.id] || 0) * 0.9;
      score += (manhattanFromStart[area.id] || 0) * 0.5;
      return score;
    }),
    rng
  );

  objectiveArea.role.objective = true;
  objectiveArea.role.objectiveRank = 1;

  const mainPathAreaIds = findShortestAreaPath(startArea.id, objectiveArea.id, adjacency);
  if (mainPathAreaIds.length >= 3) {
    const pressureIndex = Math.floor(mainPathAreaIds.length / 2);
    const pressureId = mainPathAreaIds[pressureIndex];
    const pressureArea = areaById.get(pressureId);
    if (pressureArea && !pressureArea.role.start && !pressureArea.role.objective) {
      pressureArea.role.pressure = true;
    }
  }

  if (ENABLE_EXTRACTION_ROLE) {
    const startNeighbors = adjacency[startArea.id] || new Set();
    const extractionCandidates = areas.filter((area) => {
      if (area.role.start || area.role.objective) return false;
      if (startNeighbors.has(area.id)) return false;
      return true;
    });

    if (extractionCandidates.length > 0 && randomFloat(rng) < 0.6) {
      const extractionArea = weightedPickFromValues(
        extractionCandidates,
        extractionCandidates.map((area) => {
          const centroid = centroids[area.id] || { row: 0, col: 0 };
          const edgeDistance = Math.min(
            centroid.row,
            Math.abs((rows - 1) - centroid.row),
            centroid.col,
            Math.abs((cols - 1) - centroid.col)
          );
          let score = 1;
          score += (distanceFromStart[area.id] || 0) * 0.85;
          score += 1.4 / (1 + edgeDistance);
          const tags = new Set(area.tags || []);
          if (tags.has("extraction")) score += 1.8;
          if (tags.has("outdoor")) score += 0.5;
          return score;
        }),
        rng
      );
      extractionArea.role.extraction = true;
    }
  }

  return {
    startAreaId: startArea.id,
    objectiveAreaIds: [objectiveArea.id],
    mainPathAreaIds,
    predecessorByAreaId,
    distanceFromStart
  };
}

function computeDistancesAndParents(startAreaId, adjacency) {
  const distances = { [startAreaId]: 0 };
  const parents = { [startAreaId]: "" };
  const queue = [startAreaId];

  while (queue.length) {
    const current = queue.shift();
    const nextDistance = distances[current] + 1;
    const neighbors = Array.from(adjacency[current] || []);
    neighbors.forEach((neighbor) => {
      if (distances[neighbor] !== undefined) return;
      distances[neighbor] = nextDistance;
      parents[neighbor] = current;
      queue.push(neighbor);
    });
  }

  return { distances, parents };
}

function findShortestAreaPath(startAreaId, targetAreaId, adjacency) {
  if (!startAreaId || !targetAreaId) return [];
  if (startAreaId === targetAreaId) return [startAreaId];

  const queue = [startAreaId];
  const parent = { [startAreaId]: "" };

  while (queue.length) {
    const current = queue.shift();
    const neighbors = Array.from(adjacency[current] || []);
    for (const neighbor of neighbors) {
      if (parent[neighbor] !== undefined) continue;
      parent[neighbor] = current;
      if (neighbor === targetAreaId) {
        const path = [targetAreaId];
        let cursor = targetAreaId;
        while (parent[cursor]) {
          cursor = parent[cursor];
          path.push(cursor);
        }
        return path.reverse();
      }
      queue.push(neighbor);
    }
  }

  return [startAreaId, targetAreaId];
}

function computeAreaCentroids(areas, cols) {
  const output = {};
  areas.forEach((area) => {
    if (!area.cells || area.cells.length === 0) {
      output[area.id] = { row: 0, col: 0 };
      return;
    }

    let sumRow = 0;
    let sumCol = 0;
    area.cells.forEach((cellIndex) => {
      sumRow += Math.floor(cellIndex / cols);
      sumCol += cellIndex % cols;
    });

    output[area.id] = {
      row: sumRow / area.cells.length,
      col: sumCol / area.cells.length
    };
  });
  return output;
}

function computeDistancesFrom(startAreaId, adjacency) {
  const distances = { [startAreaId]: 0 };
  const queue = [startAreaId];

  while (queue.length) {
    const current = queue.shift();
    const nextDistance = distances[current] + 1;
    const neighbors = Array.from(adjacency[current] || []);

    neighbors.forEach((neighbor) => {
      if (distances[neighbor] !== undefined) return;
      distances[neighbor] = nextDistance;
      queue.push(neighbor);
    });
  }

  return distances;
}

function validateMapQuality(map, options = {}) {
  const relaxed = !!options.relaxed;
  if (!map) {
    return { ok: false, score: -999, reasons: ["Map object missing"], hardFailures: ["Map object missing"] };
  }

  const reasons = [];
  const hardFailures = [];
  let score = 0;

  const occupiedCells = map.cells.filter((cell) => !!cell.tileId);
  const occupiedCount = occupiedCells.length;
  const areas = map.areas || [];

  if (map.rows > 7 || map.cols > 9) {
    hardFailures.push("canvas exceeds 7x9 cap");
  }

  if (occupiedCount < HARD_OCCUPIED_MIN || occupiedCount > HARD_OCCUPIED_MAX) {
    hardFailures.push(`occupied tiles out of hard range (${occupiedCount}, expected ${HARD_OCCUPIED_MIN}-${HARD_OCCUPIED_MAX})`);
  }
  if (occupiedCount >= TARGET_OCCUPIED_MIN && occupiedCount <= TARGET_OCCUPIED_MAX) {
    score += 10;
  } else {
    reasons.push(`occupied tiles outside target band (${occupiedCount}, target ${TARGET_OCCUPIED_MIN}-${TARGET_OCCUPIED_MAX})`);
  }

  if (areas.length === 0) {
    hardFailures.push("no action areas generated");
  }

  const tinyCount = areas.filter((area) => area.sizeCategory === "Tiny").length;
  const mediumPlusCount = areas.filter((area) => areaCategoryRank(area.sizeCategory) >= 3).length;
  const largePlusCount = areas.filter((area) => areaCategoryRank(area.sizeCategory) >= 4).length;
  const corridorOnlyCount = areas.filter((area) => area.shapeKind === "corridor").length;
  const corridorRatio = areas.length ? (corridorOnlyCount / areas.length) : 1;
  const roomShapeCount = areas.filter((area) => area.shapeKind === "room").length;

  if (areas.length > 0 && (tinyCount / areas.length) > 0.5) {
    hardFailures.push("more than 50% of areas are Tiny");
  }
  if (tinyCount > 2) {
    score -= 20;
    reasons.push("more than two Tiny areas");
  }

  if (mediumPlusCount === 0) {
    hardFailures.push("no Medium+ area exists");
  }

  if (largePlusCount > 2) {
    hardFailures.push("more than two Large+ areas");
  }

  if (corridorRatio > MAX_CORRIDOR_AREA_RATIO) {
    hardFailures.push("more than 30% corridor-only areas");
    score -= 30;
  }

  if (roomShapeCount >= 2) {
    score += 15;
  } else {
    reasons.push("fewer than two room-shaped areas");
  }

  const templateId = map.graph?.template || "unknown";
  const hasBranchOrLoop = ["spine_branch", "loop", "hub", "fork_reconverge", "gauntlet_pocket"].includes(templateId);
  if (hasBranchOrLoop) {
    score += 15;
  }

  const objectiveAreas = areas.filter((area) => area.role?.objective);
  if (objectiveAreas.length === 0) {
    hardFailures.push("no objective area assigned");
  }

  if (objectiveAreas.some((area) => area.sizeCategory === "Tiny")) {
    hardFailures.push("objective area is Tiny");
    score -= 25;
  } else if (objectiveAreas.some((area) => ["Medium", "Large"].includes(area.sizeCategory))) {
    score += 10;
  }

  const startArea = areas.find((area) => area.role?.start);
  if (startArea && objectiveAreas.length > 0) {
    const centroids = computeAreaCentroids(areas, map.cols);
    const startCentroid = centroids[startArea.id];
    const objectiveCentroid = centroids[objectiveAreas[0].id];
    const dist = Math.round(
      Math.abs(startCentroid.row - objectiveCentroid.row) + Math.abs(startCentroid.col - objectiveCentroid.col)
    );
    if (dist >= MIN_START_OBJECTIVE_DISTANCE) {
      score += 20;
    } else {
      reasons.push(`start-objective distance below ${MIN_START_OBJECTIVE_DISTANCE}`);
    }
  }

  const longCorridorExists = (map.corridors || []).some((corridor) => corridor.cells.length > 3);
  if (!longCorridorExists) {
    score += 10;
  } else {
    reasons.push("corridor longer than 3 tiles");
  }

  if (occupiedCells.length > 0) {
    const rows = occupiedCells.map((cell) => cell.row);
    const cols = occupiedCells.map((cell) => cell.col);
    const minRow = Math.min(...rows);
    const maxRow = Math.max(...rows);
    const minCol = Math.min(...cols);
    const maxCol = Math.max(...cols);
    const bboxWidth = (maxCol - minCol) + 1;
    const bboxHeight = (maxRow - minRow) + 1;
    const bboxArea = bboxWidth * bboxHeight;

    if (bboxWidth >= (map.cols - 1) && bboxHeight >= (map.rows - 1)) {
      score -= 15;
      reasons.push("bounding box nearly fills full 7x9 canvas");
    }

    if ((bboxWidth >= (map.cols - 1) && bboxHeight <= 2) || (bboxHeight >= (map.rows - 1) && bboxWidth <= 2)) {
      reasons.push("layout forms long thin spanning path");
    }
    if (bboxArea / Math.max(1, occupiedCount) > 2.7) {
      reasons.push("layout too sparse/snakelike for table footprint");
    }
    if (isSnakeLikeOccupiedLayout(occupiedCells, map.rows, map.cols, bboxWidth, bboxHeight)) {
      reasons.push("layout behaves like a snake path");
    }
  }

  const corridorEdgeSet = new Set(
    (map.corridors || [])
      .filter((corridor) => corridor.cells.length > 0)
      .map((corridor) => areaEdgeKey(corridor.fromAreaIndex, corridor.toAreaIndex))
  );
  (map.graph?.edges || []).forEach(([from, to]) => {
    const areaA = areas[from];
    const areaB = areas[to];
    if (!areaA || !areaB) return;
    if (areaCategoryRank(areaA.sizeCategory) >= 3 && areaCategoryRank(areaB.sizeCategory) >= 3) {
      const key = areaEdgeKey(from, to);
      if (!corridorEdgeSet.has(key)) {
        reasons.push("Medium/Large areas directly connected without choke");
      }
    }
  });

  const acceptScore = relaxed ? QUALITY_RELAXED_ACCEPT_SCORE : QUALITY_ACCEPT_SCORE;
  const ok = hardFailures.length === 0 && score >= acceptScore;
  const mergedReasons = [...hardFailures, ...reasons];

  return {
    ok,
    score,
    reasons: mergedReasons,
    hardFailures,
    relaxed
  };
}

function areaEdgeKey(a, b) {
  return a < b ? `${a}|${b}` : `${b}|${a}`;
}

function isSnakeLikeOccupiedLayout(occupiedCells, rows, cols, bboxWidth, bboxHeight) {
  if (occupiedCells.length < 10) return false;
  if (Math.max(bboxWidth, bboxHeight) < 7) return false;

  const occupiedSet = new Set(occupiedCells.map((cell) => cell.index));
  let lowDegreeCount = 0;
  let branchCount = 0;

  occupiedCells.forEach((cell) => {
    const degree = getNeighbors(cell.index, rows, cols)
      .filter((neighbor) => occupiedSet.has(neighbor))
      .length;
    if (degree <= 2) lowDegreeCount += 1;
    if (degree >= 3) branchCount += 1;
  });

  const lowDegreeRatio = lowDegreeCount / occupiedCells.length;
  return lowDegreeRatio > 0.88 && branchCount <= Math.max(1, Math.floor(occupiedCells.length * 0.07));
}
function renderMission() {
  if (!state.mission) {
    refs.missionHeadline.textContent = "Operation Pending";
    refs.missionMeta.textContent = "Generate a mission to populate briefing details.";
    refs.missionOutput.classList.add("is-empty");
    refs.missionOutput.textContent = "No mission generated yet.";
    return;
  }

  refs.missionHeadline.textContent = state.mission.operationCodename || "Operation Silent Vector";

  const mapMeta = state.map ? ` | Map: ${state.map.rows}x${state.map.cols}, ${state.map.areas.length} areas` : "";
  refs.missionMeta.textContent = `Mission ID: ${state.mission.id}${state.mission.seed ? ` | Seed: ${state.mission.seed}` : ""}${mapMeta}`;

  refs.missionOutput.classList.remove("is-empty");
  refs.missionOutput.innerHTML = renderMissionBriefHtml(state.mission);

  CATEGORY_CONFIG.forEach((cfg) => {
    const card = refs.categoryGrid.querySelector(`[data-field="${cfg.field}"]`);
    if (!card) return;
    const value = card.querySelector(".category-value");
    const lock = card.querySelector(".lock-category");
    value.textContent = state.mission[cfg.field] || "-";
    lock.checked = !!state.lockState[cfg.field];
  });
}

function renderMissionBriefHtml(mission) {
  const mapSummary = mission.map
    ? `${mission.map.rows}x${mission.map.cols} with ${mission.map.actionAreas.length} areas`
    : (state.map ? `${state.map.rows}x${state.map.cols} with ${state.map.areas.length} areas` : "No map");

  const rows = [
    { key: "location", label: "Location", value: mission.location },
    { key: "corruption", label: "Corruption", value: mission.corruption },
    { key: "twist", label: "Hidden Truth", value: mission.twist },
    { key: "escalation", label: "Escalation", value: mission.escalation },
    { key: "primary-objective", label: "Primary Objective", value: mission.primaryObjective },
    { key: "secondary-objective", label: "Secondary Objective", value: mission.secondaryObjective },
    { key: "enemy-mix", label: "Enemy Mix Hint", value: mission.enemyMixHint },
    { key: "map-notes", label: "Map Notes", value: mission.mapNotes },
    { key: "action-areas", label: "Action Areas", value: mapSummary },
    { key: "reward", label: "Reward/Consequence", value: mission.rewardConsequence },
    { key: "tone", label: "Tone", value: mission.toneTag }
  ];

  const items = rows.map((row) => `
    <article class="mission-brief-item ${row.key}">
      <span class="mission-brief-label">${escapeHtml(row.label)}</span>
      <p class="mission-brief-value">${escapeHtml(row.value || "-")}</p>
    </article>
  `);

  return `<div class="mission-brief-grid">${items.join("")}</div>`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&#039;");
}

function formatMissionText(mission) {
  const shortName = shortLocationName(mission.location || "Unknown Site");
  const lines = [
    `MISSION CONCEPT: ${shortName}`,
    `- Operation Codename: ${mission.operationCodename || "Operation Silent Vector"}`,
    `- Location: ${mission.location}`,
    `- Corruption: ${mission.corruption}`,
    `- Hidden Truth: ${mission.twist}`,
    `- Escalation: ${mission.escalation}`,
    `- Primary Objective: ${mission.primaryObjective}`,
    `- Secondary Objective: ${mission.secondaryObjective}`,
    `- Enemy Mix Hint: ${mission.enemyMixHint}`,
    `- Map Notes: ${mission.mapNotes}`,
    `- Reward/Consequence: ${mission.rewardConsequence}`,
    `- Tone: ${mission.toneTag}`
  ];

  const mapData = mission.map || buildMapExportData(state.map);
  if (mapData) {
    lines.push(`- Action Areas: ${mapData.rows}x${mapData.cols}, ${mapData.actionAreas.length} areas`);
    lines.push("- Area Legend:");
    mapData.actionAreas.forEach((area) => {
      const roleParts = [];
      if (area.roles?.start) roleParts.push("START");
      if (area.roles?.objective) roleParts.push(area.roles.objectiveRank > 0 ? `OBJECTIVE ${area.roles.objectiveRank}` : "OBJECTIVE");
      if (area.roles?.pressure) roleParts.push("PRESSURE");
      if (area.roles?.extraction) roleParts.push("EXTRACTION");
      const roleSuffix = roleParts.length ? ` [${roleParts.join(", ")}]` : "";
      const sizeCategory = area.areaSizeCategory || getAreaSizeInfo(area.size || 0).category;
      const cardsReturned = area.cardsReturnedOnClear ?? getAreaSizeInfo(area.size || 0).cardsReturned;
      const cardsLabel = cardsReturned === "all"
        ? "all cards returned"
        : `${cardsReturned} cards returned`;
      const supportLabel = (area.supportingTileNames && area.supportingTileNames.length)
        ? area.supportingTileNames.join(", ")
        : "none";
      lines.push(`  - Area ${area.id}: ${area.dominantTileName} (${area.size} tiles, ${sizeCategory}, ${cardsLabel}, support: ${supportLabel})${roleSuffix}`);
    });
  }

  return lines.join("\n");
}

function shortLocationName(location) {
  const chunk = location.split(/[,:;.!?]/)[0].trim();
  return chunk.length > 50 ? `${chunk.slice(0, 47)}...` : chunk;
}

async function copyMissionText() {
  if (!state.mission) return;
  const text = formatMissionText(state.mission);

  try {
    await navigator.clipboard.writeText(text);
    setStatus("Mission copied to clipboard.");
  } catch (err) {
    const area = document.createElement("textarea");
    area.value = text;
    area.style.position = "fixed";
    area.style.opacity = "0";
    document.body.appendChild(area);
    area.select();
    document.execCommand("copy");
    document.body.removeChild(area);
    setStatus("Mission copied using fallback copy method.");
  }
}

function exportMissionJson() {
  if (!state.mission) return;
  const payload = {
    ...state.mission,
    map: buildMapExportData(state.map)
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `mission-${state.mission.id}.json`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function attachMapToMission() {
  if (!state.mission) return;
  state.mission.map = buildMapExportData(state.map);
}

function buildMapExportData(map) {
  if (!map) return null;

  const occupiedCells = map.cells.filter((cell) => !!cell.tileId).length;
  const totalCells = map.rows * map.cols;
  const totalTiles = occupiedCells;
  const grid = Array.from({ length: map.rows }, (_, row) => (
    Array.from({ length: map.cols }, (_, col) => {
      const index = toCellIndex(row, col, map.cols);
      const cell = map.cells[index];
      return {
        index,
        row,
        col,
        tileId: cell.tileId || "",
        actionAreaId: cell.areaId || "",
        corridorId: cell.corridorId || "",
        isCorridorTile: !!cell.isCorridor
      };
    })
  ));

  return {
    preset: map.preset,
    seed: map.seed || refs.seedInput.value.trim() || "",
    templateId: map.graph?.template || "",
    rows: map.rows,
    cols: map.cols,
    totalTiles,
    occupancy: {
      occupiedCells,
      totalCells,
      ratio: totalCells > 0 ? Number((occupiedCells / totalCells).toFixed(4)) : 0
    },
    grid,
    settings: {
      ...map.settings
    },
    graph: {
      template: map.graph.template,
      nodeCount: map.graph.nodeCount,
      edges: map.graph.edges.map(([from, to]) => ({ fromAreaIndex: from, toAreaIndex: to }))
    },
    cells: map.cells.map((cell) => ({
      index: cell.index,
      row: cell.row,
      col: cell.col,
      tileId: cell.tileId,
      tileName: cell.tileName,
      tileImage: cell.tileImage || "",
      actionAreaId: cell.areaId,
      isCorridorTile: !!cell.isCorridor,
      corridorId: cell.corridorId || "",
      isVariationTile: !!cell.isVariation,
      dominantTileId: cell.dominantTileId,
      biome: cell.biome,
      tags: [...cell.tags]
    })),
    actionAreas: map.areas.map((area) => ({
      id: area.id,
      role: mapAreaRoleLabel(area),
      size: area.size,
      tileCount: area.tileCount ?? area.size,
      areaSizeCategory: area.sizeCategory,
      cardsReturnedOnClear: area.cardsReturnedOnClear,
      shapeType: area.shapeArchetype || "",
      shapeArchetype: area.shapeArchetype || "",
      shapeKind: area.shapeKind || "",
      dominantTileId: area.dominantTileId,
      dominantTileName: area.dominantTileName,
      dominantTileImage: area.dominantTileImage || "",
      supportingTileIds: [...(area.supportingTileIds || [])],
      supportingTileNames: [...(area.supportingTileNames || [])],
      biome: area.biome,
      tags: [...area.tags],
      cells: [...area.cells],
      roles: {
        start: !!area.role.start,
        objective: !!area.role.objective,
        objectiveRank: area.role.objectiveRank || 0,
        extraction: !!area.role.extraction,
        pressure: !!area.role.pressure
      }
    })),
    corridors: map.corridors.map((corridor) => ({
      id: corridor.id,
      fromAreaId: corridor.fromAreaId,
      toAreaId: corridor.toAreaId,
      fromAreaIndex: corridor.fromAreaIndex,
      toAreaIndex: corridor.toAreaIndex,
      shapeArchetype: corridor.shapeArchetype || "",
      tileId: corridor.tileId,
      tileName: corridor.tileName,
      tileImage: corridor.tileImage || "",
      cells: [...corridor.cells]
    })),
    gates: Object.values(map.gates).map((gate) => ({
      edgeKey: gate.key,
      fromCell: gate.a,
      toCell: gate.b,
      areaA: gate.areaAId,
      areaB: gate.areaBId,
      graphEdgeKey: gate.graphEdgeKey || "",
      state: gate.state
    }))
  };
}

function pushHistory(mission) {
  const record = {
    id: mission.id,
    seed: mission.seed,
    text: formatMissionText(mission),
    mission: { ...mission }
  };

  state.history.unshift(record);
  state.history = state.history.slice(0, MAX_HISTORY);
  localStorage.setItem(HISTORY_KEY, JSON.stringify(state.history));
}

function loadHistory() {
  try {
    const raw = localStorage.getItem(HISTORY_KEY);
    state.history = raw ? JSON.parse(raw) : [];
  } catch (err) {
    state.history = [];
  }
}

function renderHistory() {
  refs.historyList.innerHTML = "";
  if (state.history.length === 0) {
    refs.historyList.textContent = "No history yet.";
    return;
  }

  state.history.forEach((record) => {
    const item = document.createElement("article");
    item.className = "history-item";
    const title = document.createElement("h3");
    title.textContent = `${record.id}${record.seed ? ` | seed: ${record.seed}` : ""}`;
    const body = document.createElement("pre");
    body.textContent = record.text;
    item.appendChild(title);
    item.appendChild(body);
    refs.historyList.appendChild(item);
  });
}

function filterByTags(pool, selectedTags) {
  if (!selectedTags || selectedTags.size === 0) return pool;
  return pool.filter((entry) => {
    const tags = Array.isArray(entry.tags) ? entry.tags : [];
    return tags.some((tag) => selectedTags.has(tag));
  });
}

function weightedPick(pool, rng) {
  let totalWeight = 0;
  pool.forEach((entry) => {
    totalWeight += Number(entry.weight || 1);
  });

  let cursor = randomFloat(rng) * totalWeight;
  for (const entry of pool) {
    cursor -= Number(entry.weight || 1);
    if (cursor <= 0) return entry;
  }
  return pool[pool.length - 1];
}

function weightedPickFromValues(values, weights, rng) {
  if (values.length === 1) return values[0];

  const normalizedWeights = weights.map((value) => Math.max(0.0001, Number(value || 0.0001)));
  let total = 0;
  normalizedWeights.forEach((value) => {
    total += value;
  });

  let cursor = randomFloat(rng) * total;
  for (let i = 0; i < values.length; i += 1) {
    cursor -= normalizedWeights[i];
    if (cursor <= 0) return values[i];
  }

  return values[values.length - 1];
}

function pickDistinctWeighted(list, count, rng, weightFn) {
  const pool = [...list];
  const result = [];

  while (pool.length && result.length < count) {
    const weights = pool.map((item) => Math.max(0.001, Number(weightFn(item) || 0.001)));
    const picked = weightedPickFromValues(pool, weights, rng);
    result.push(picked);
    const removeIndex = pool.indexOf(picked);
    if (removeIndex >= 0) pool.splice(removeIndex, 1);
  }

  return result;
}

function randomFloat(rng) {
  return typeof rng === "function" ? rng() : Math.random();
}

function randomInt(min, max, rng) {
  const low = Math.ceil(min);
  const high = Math.floor(max);
  return Math.floor(randomFloat(rng) * (high - low + 1)) + low;
}

function shuffleInPlace(list, rng) {
  for (let i = list.length - 1; i > 0; i -= 1) {
    const j = randomInt(0, i, rng);
    const temp = list[i];
    list[i] = list[j];
    list[j] = temp;
  }
}

function createRng(seedString) {
  const seed = hashSeed(seedString);
  return mulberry32(seed);
}

function hashSeed(str) {
  let h = 2166136261;
  for (let i = 0; i < str.length; i += 1) {
    h ^= str.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return h >>> 0;
}

function mulberry32(a) {
  return function rng() {
    let t = (a += 0x6d2b79f5);
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function createMissionId(seed) {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const datePart = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}`;
  const timePart = `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
  const randRng = seed ? createRng(`${seed}:id`) : Math.random;
  const randValue = Math.floor(randomFloat(randRng) * 10000)
    .toString()
    .padStart(4, "0");
  return `${datePart}-${timePart}-${randValue}`;
}

function getNeighbors(index, rows, cols) {
  const row = Math.floor(index / cols);
  const col = index % cols;
  const neighbors = [];

  if (row > 0) neighbors.push(toCellIndex(row - 1, col, cols));
  if (row < rows - 1) neighbors.push(toCellIndex(row + 1, col, cols));
  if (col > 0) neighbors.push(toCellIndex(row, col - 1, cols));
  if (col < cols - 1) neighbors.push(toCellIndex(row, col + 1, cols));

  return neighbors;
}

function toCellIndex(row, col, cols) {
  return (row * cols) + col;
}

function edgeKey(a, b) {
  return a < b ? `${a}-${b}` : `${b}-${a}`;
}

function toAreaId(index) {
  if (index < AREA_LETTERS.length) return AREA_LETTERS[index];
  return `A${index + 1}`;
}

function getAreaSizeBias() {
  return Number(refs.areaBiasSlider.value || 0);
}

function allowVariationTiles() {
  return !!refs.allowVariationToggle.checked;
}

function isMixedBiomesEnabled() {
  return !!refs.mixedBiomesToggle.checked;
}

function setStatus(message) {
  refs.status.textContent = message;
}

function setMapStatus(message) {
  refs.mapStatus.textContent = message;
}
