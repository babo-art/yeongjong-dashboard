(() => {
  "use strict";

  const STORAGE_KEYS = {
    kakaoKey: "yeongjong.dashboard.kakaoKey.v1",
  };

  const DEFAULT_CENTER = { lat: 37.482, lng: 126.484 };

  const ICONS = {
    reference: assetUrl("./assets/icons/산(기준지점)-map.svg"),
    business: assetUrl("./assets/icons/위험물제조소등-map.svg"),
    fire: assetUrl("./assets/icons/소방기관-map.svg"),
  };

  const MAP_COLORS = {
    reference: "#2E7D32",
    business: "#D97706",
    fire: "#8A0715",
  };

  const ENABLE_JURISDICTION_LAYER = true;

  const state = {
    kakaoKey: "",
    sdkLoaded: false,
    map: null,
    data: {
      businesses: [],
      fireOrgs: [],
      references: [],
      jurisdictions: [],
    },
    rendered: {
      markers: [],
      referenceLabels: [],
      businessLabels: [],
      fireLabels: [],
      jurisdictionPolygons: [],
      circles: [],
      lines: [],
      lineLabels: [],
    },
    anchorId: "",
    anchorKind: "",
    targetId: "",
    targetKind: "",
    fileName: "",
    shouldFitDataBounds: false,
    toastTimer: null,
    resizeTimer: null,
  };

  const dom = {};

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    loadJurisdictionData();
    bindDom();
    bindEvents();
    loadSavedKey();
    renderAll();

    if (state.kakaoKey) {
      dom.kakaoKeyInput.value = state.kakaoKey;
      initializeMapRuntime();
    } else {
      openKeyModal();
    }
  }

  function bindDom() {
    dom.bottomDock = document.getElementById("bottomDock");
    dom.selectionHud = document.getElementById("selectionHud");
    dom.dockPopover = document.getElementById("dockPopover");
    dom.popoverCloseBtn = document.getElementById("popoverCloseBtn");
    dom.openSettingsBtn = document.getElementById("openSettingsBtn");
    dom.excelInput = document.getElementById("excelInput");
    dom.toolsBtn = document.getElementById("toolsBtn");
    dom.selectionSummary = document.getElementById("selectionSummary");
    dom.mapOverlayMessage = document.getElementById("mapOverlayMessage");
    dom.clearSelectionBtn = document.getElementById("clearSelectionBtn");
    dom.keyModal = document.getElementById("keyModal");
    dom.kakaoKeyInput = document.getElementById("kakaoKeyInput");
    dom.saveKeyBtn = document.getElementById("saveKeyBtn");
    dom.removeKeyBtn = document.getElementById("removeKeyBtn");
    dom.keyModalHint = document.getElementById("keyModalHint");
    dom.toast = document.getElementById("toast");
    dom.map = document.getElementById("map");
  }

  function bindEvents() {
    dom.openSettingsBtn.addEventListener("click", () => {
      closeDockPopover();
      openKeyModal();
    });
    dom.saveKeyBtn.addEventListener("click", handleSaveKey);
    dom.removeKeyBtn.addEventListener("click", handleRemoveKey);
    dom.excelInput.addEventListener("change", handleExcelUpload);
    dom.toolsBtn.addEventListener("click", toggleDockPopover);
    dom.popoverCloseBtn.addEventListener("click", closeDockPopover);
    dom.clearSelectionBtn.addEventListener("click", clearSelection);
    dom.kakaoKeyInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        handleSaveKey();
      }
    });
    window.addEventListener("resize", handleWindowResize);
  }

  function loadJurisdictionData() {
    const featureCollection = window.CENTER_JURISDICTION_GEOJSON;
    const features = Array.isArray(featureCollection?.features) ? featureCollection.features : [];
    state.data.jurisdictions = features
      .filter((feature) => feature?.geometry?.coordinates?.length)
      .map((feature) => ({
        wardId: sanitizeString(feature.properties?.ward_id),
        wardName: sanitizeString(feature.properties?.ward_nm),
        geometry: feature.geometry,
      }));
  }

  function loadSavedKey() {
    try {
      state.kakaoKey = localStorage.getItem(STORAGE_KEYS.kakaoKey) || "";
    } catch (error) {
      state.kakaoKey = "";
      showToast("브라우저 로컬 저장소를 사용할 수 없어 지도 키 저장이 제한됩니다.");
    }
  }

  function persistKey(value) {
    localStorage.setItem(STORAGE_KEYS.kakaoKey, value);
  }

  function removeKeyFromStorage() {
    localStorage.removeItem(STORAGE_KEYS.kakaoKey);
  }

  function openKeyModal(message) {
    dom.keyModal.classList.add("is-visible");
    dom.keyModal.setAttribute("aria-hidden", "false");
    dom.kakaoKeyInput.value = state.kakaoKey || "";
    dom.kakaoKeyInput.focus();
    if (message) {
      dom.keyModalHint.textContent = message;
    }
  }

  function closeKeyModal() {
    dom.keyModal.classList.remove("is-visible");
    dom.keyModal.setAttribute("aria-hidden", "true");
    dom.keyModalHint.textContent = "도메인 설정은 Kakao Developers에서 `file://` 또는 허용 출처 조건에 맞게 등록되어 있어야 합니다.";
  }

  async function handleSaveKey() {
    const input = dom.kakaoKeyInput.value.trim();
    if (!input) {
      showToast("카카오 JavaScript 키를 입력해 주세요.");
      return;
    }

    const previousKey = state.kakaoKey;
    state.kakaoKey = input;

    try {
      persistKey(input);
    } catch (error) {
      showToast("지도 키를 저장하지 못했습니다. 이 브라우저에서는 매번 다시 입력해야 할 수 있습니다.");
    }

    if (state.sdkLoaded && previousKey && previousKey !== input) {
      showToast("새 지도 키를 적용하기 위해 페이지를 다시 불러옵니다.");
      setTimeout(() => window.location.reload(), 450);
      return;
    }

    closeKeyModal();

    if (!state.sdkLoaded) {
      await initializeMapRuntime();
    } else {
      renderAll();
    }
  }

  function handleRemoveKey() {
    state.kakaoKey = "";
    try {
      removeKeyFromStorage();
    } catch (error) {
      // noop
    }
    dom.kakaoKeyInput.value = "";
    showToast("저장된 지도 키를 삭제했습니다.");
    openKeyModal();
  }

  async function initializeMapRuntime() {
    renderAll();
    try {
      await loadKakaoSdk(state.kakaoKey);
      setupMap();
      closeKeyModal();
      renderAll();
      renderMapObjects();
      showToast("지도 런타임이 준비되었습니다.");
    } catch (error) {
      state.sdkLoaded = false;
      renderAll();
      openKeyModal("지도 키가 올바른지, Kakao Developers 출처 설정이 맞는지 확인해 주세요.");
      showToast(error.message || "카카오맵 SDK 로드에 실패했습니다.");
    }
  }

  function loadKakaoSdk(key) {
    return new Promise((resolve, reject) => {
      if (!key) {
        reject(new Error("카카오 JavaScript 키가 없습니다."));
        return;
      }

      if (window.kakao?.maps?.load) {
        state.sdkLoaded = true;
        window.kakao.maps.load(() => resolve());
        return;
      }

      const existing = document.querySelector('script[data-kakao-sdk="true"]');
      if (existing) {
        existing.remove();
      }

      const script = document.createElement("script");
      script.async = true;
      script.defer = true;
      script.dataset.kakaoSdk = "true";
      script.src = `https://dapi.kakao.com/v2/maps/sdk.js?autoload=false&appkey=${encodeURIComponent(key)}`;

      const timeoutId = window.setTimeout(() => {
        script.remove();
        reject(new Error("카카오맵 SDK 응답 시간이 초과되었습니다."));
      }, 12000);

      script.onload = () => {
        window.clearTimeout(timeoutId);
        if (!window.kakao?.maps?.load) {
          reject(new Error("카카오맵 SDK가 정상적으로 로드되지 않았습니다."));
          return;
        }
        window.kakao.maps.load(() => {
          state.sdkLoaded = true;
          resolve();
        });
      };

      script.onerror = () => {
        window.clearTimeout(timeoutId);
        reject(new Error("카카오맵 SDK 스크립트를 불러오지 못했습니다."));
      };

      document.head.appendChild(script);
    });
  }

  function setupMap() {
    if (!window.kakao?.maps) {
      return;
    }

    if (!state.map) {
      state.map = new window.kakao.maps.Map(dom.map, {
        center: new window.kakao.maps.LatLng(DEFAULT_CENTER.lat, DEFAULT_CENTER.lng),
        level: 9,
      });

      const zoomControl = new window.kakao.maps.ZoomControl();
      state.map.addControl(zoomControl, window.kakao.maps.ControlPosition.RIGHT);

      window.kakao.maps.event.addListener(state.map, "zoom_changed", () => {
        handleMapZoomChanged();
      });
    }
  }

  function handleMapZoomChanged() {
    if (!state.sdkLoaded || !state.map) {
      return;
    }

    if (!hasRenderableData()) {
      return;
    }

    renderPointLabels();
  }

  function hasRenderableData() {
    return Boolean(
      state.data.references.length
      || state.data.businesses.length
      || state.data.fireOrgs.length,
    );
  }

  async function handleExcelUpload(event) {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    try {
      const workbookData = await parseWorkbookFile(file);
      const transformed = transformWorkbookData(workbookData);
      state.data = {
        ...transformed,
        jurisdictions: state.data.jurisdictions || [],
      };
      state.fileName = file.name;
      state.anchorId = "";
      state.anchorKind = "";
      state.targetId = "";
      state.targetKind = "";
      state.shouldFitDataBounds = true;
      renderAll();
      renderMapObjects();
      showToast("엑셀 데이터를 불러왔습니다.");
    } catch (error) {
      console.error(error);
      showToast(error.message || "엑셀 파일을 읽지 못했습니다.");
    } finally {
      event.target.value = "";
    }
  }

  async function parseWorkbookFile(file) {
    if (!window.JSZip) {
      throw new Error("엑셀 파서를 불러오지 못했습니다.");
    }

    const buffer = await file.arrayBuffer();
    const zip = await window.JSZip.loadAsync(buffer);

    const workbookXml = await readZipText(zip, "xl/workbook.xml");
    const workbookRelsXml = await readZipText(zip, "xl/_rels/workbook.xml.rels");
    const sharedStringsXml = zip.file("xl/sharedStrings.xml")
      ? await readZipText(zip, "xl/sharedStrings.xml")
      : null;

    const workbookDoc = parseXml(workbookXml);
    const relsDoc = parseXml(workbookRelsXml);
    const sharedStrings = sharedStringsXml ? parseSharedStrings(parseXml(sharedStringsXml)) : [];
    const sheets = parseWorkbookSheets(workbookDoc, relsDoc);
    const results = {};

    for (const sheet of sheets) {
      const sheetXml = await readZipText(zip, normalizeWorkbookTarget(sheet.target));
      const sheetDoc = parseXml(sheetXml);
      results[sheet.name] = parseWorksheet(sheetDoc, sharedStrings);
    }

    return results;
  }

  function getCol(row, ...candidates) {
    for (const key of candidates) {
      if (row[key] !== undefined && row[key] !== null && row[key] !== "") {
        return row[key];
      }
    }
    return "";
  }

  function transformWorkbookData(workbookData) {
    const businessRows = workbookData["위험물업체"];
    const fireRows = workbookData["소방관서"] || workbookData["소방기관"];
    const referenceRows = workbookData["기준지점"];

    if (!businessRows || !fireRows || !referenceRows) {
      throw new Error("표준 시트(위험물업체, 소방관서, 기준지점)를 모두 찾지 못했습니다.");
    }

    const businesses = businessRows
      .map((row, index) => ({
        id: sanitizeString(getCol(row, "연번", "entity_id")) || `BIZ_${index + 1}`,
        name: sanitizeString(getCol(row, "업체명", "name")),
        address: sanitizeString(getCol(row, "주소", "address")),
        ...resolveCoords(row),
        facilityType: sanitizeString(getCol(row, "제조소등 구분", "facility_type")),
        hazardClass: sanitizeString(getCol(row, "유별", "종류(류별)", "hazard_class")),
        facilityCount: toNumber(getCol(row, "개소", "facility_count")) || 0,
        itemName: sanitizeString(getCol(row, "품명", "item_name")),
        quantityMultiple: sanitizeString(getCol(row, "지정수량 배수(합계)", "quantity_multiple")),
        displayValue: toNumber(getCol(row, "display_value")) || 0,
        active: normalizeActive(getCol(row, "active")),
        notes: sanitizeString(getCol(row, "notes", "비고")),
        kind: "business",
      }))
      .filter(isRenderablePoint);

    const fireOrgs = fireRows
      .map((row, index) => ({
        id: sanitizeString(getCol(row, "연번", "entity_id")) || `FIRE_${index + 1}`,
        name: sanitizeString(getCol(row, "119안전센터(지역대)/구조대", "소방관서", "name")),
        address: sanitizeString(getCol(row, "주소", "address")),
        orgType: sanitizeString(getCol(row, "org_type", "기관유형")),
        ...resolveCoords(row),
        active: normalizeActive(getCol(row, "active")),
        notes: sanitizeString(getCol(row, "notes", "비고")),
        kind: "fire",
      }))
      .filter(isRenderablePoint);

    const references = referenceRows
      .map((row, index) => ({
        id: sanitizeString(getCol(row, "연번", "entity_id")) || `REF_${index + 1}`,
        name: sanitizeString(getCol(row, "이름", "name")),
        labelCode: sanitizeString(getCol(row, "label_code")),
        pointType: sanitizeString(getCol(row, "point_type", "유형")),
        address: sanitizeString(getCol(row, "주소", "address")),
        ...resolveCoords(row),
        active: normalizeActive(getCol(row, "active")),
        notes: sanitizeString(getCol(row, "notes", "비고")),
        kind: "reference",
      }))
      .filter(isRenderablePoint);

    if (!references.length) {
      throw new Error("기준지점 데이터가 없거나 좌표가 비어 있습니다.");
    }

    if (!businesses.length) {
      throw new Error("위험물업체 데이터가 없거나 좌표가 비어 있습니다.");
    }

    return { businesses, fireOrgs, references };
  }

  function isRenderablePoint(item) {
    if (!item.active || !item.name) return false;
    if (!Number.isFinite(item.lat) || !Number.isFinite(item.lng)) return false;
    if (item.lat === 0 && item.lng === 0) return false;
    if (item.lat < 32 || item.lat > 40) return false;
    if (item.lng < 122 || item.lng > 134) return false;
    return true;
  }

  function resolveCoords(row) {
    const latRaw = toNumber(getCol(row, "위도", "lat"));
    const lngRaw = toNumber(getCol(row, "경도", "lng"));
    const inLat = (v) => Number.isFinite(v) && v >= 32 && v <= 40;
    const inLng = (v) => Number.isFinite(v) && v >= 122 && v <= 134;
    if (inLat(latRaw) && inLng(lngRaw)) return { lat: latRaw, lng: lngRaw };
    if (inLat(lngRaw) && inLng(latRaw)) return { lat: lngRaw, lng: latRaw };
    return { lat: latRaw, lng: lngRaw };
  }

  function normalizeActive(value) {
    const normalized = sanitizeString(value).toUpperCase();
    if (!normalized) {
      return true;
    }
    return normalized !== "N";
  }

  function sanitizeString(value) {
    return String(value ?? "").trim();
  }

  function toNumber(value) {
    if (value == null || value === "") {
      return null;
    }
    const numeric = Number(String(value).replace(/,/g, "").trim());
    return Number.isFinite(numeric) ? numeric : null;
  }

  async function readZipText(zip, path) {
    const file = zip.file(path);
    if (!file) {
      throw new Error(`엑셀 내부 파일을 찾지 못했습니다: ${path}`);
    }
    return file.async("string");
  }

  function parseXml(text) {
    const doc = new DOMParser().parseFromString(text, "application/xml");
    const parserError = doc.getElementsByTagName("parsererror")[0];
    if (parserError) {
      throw new Error("엑셀 XML 파싱에 실패했습니다.");
    }
    return doc;
  }

  function getElementsByLocalName(root, name) {
    return Array.from(root.getElementsByTagName("*")).filter((node) => node.localName === name);
  }

  function parseSharedStrings(doc) {
    return getElementsByLocalName(doc, "si").map((si) => {
      const texts = getElementsByLocalName(si, "t");
      return texts.map((node) => node.textContent || "").join("");
    });
  }

  function parseWorkbookSheets(workbookDoc, relsDoc) {
    const relMap = new Map(
      getElementsByLocalName(relsDoc, "Relationship").map((rel) => [
        rel.getAttribute("Id"),
        rel.getAttribute("Target"),
      ]),
    );

    return getElementsByLocalName(workbookDoc, "sheet").map((sheet) => ({
      name: sheet.getAttribute("name") || "",
      target: relMap.get(sheet.getAttribute("r:id")) || "",
    }));
  }

  function normalizeWorkbookTarget(target) {
    const raw = sanitizeString(target).replace(/\\/g, "/");
    if (!raw) {
      throw new Error("엑셀 시트 경로를 찾지 못했습니다.");
    }

    if (raw.startsWith("/")) {
      return raw.slice(1);
    }

    if (raw.startsWith("xl/")) {
      return raw;
    }

    if (raw.startsWith("../")) {
      const trimmed = raw.replace(/^(\.\.\/)+/, "");
      return trimmed.startsWith("xl/") ? trimmed : `xl/${trimmed}`;
    }

    return `xl/${raw}`;
  }

  function parseWorksheet(sheetDoc, sharedStrings) {
    const rows = [];
    let maxColumn = 0;

    getElementsByLocalName(sheetDoc, "row").forEach((rowNode) => {
      const rowIndex = Math.max((Number(rowNode.getAttribute("r")) || 1) - 1, 0);
      rows[rowIndex] = rows[rowIndex] || [];

      getElementsByLocalName(rowNode, "c").forEach((cell) => {
        const ref = cell.getAttribute("r") || "";
        const columnIndex = columnNameToIndex(ref.replace(/\d+/g, ""));
        maxColumn = Math.max(maxColumn, columnIndex);
        rows[rowIndex][columnIndex] = parseCellValue(cell, sharedStrings);
      });
    });

    const normalizedRows = rows.map((row) => {
      const nextRow = new Array(maxColumn + 1).fill("");
      if (!row) {
        return nextRow;
      }
      row.forEach((value, index) => {
        nextRow[index] = value ?? "";
      });
      return nextRow;
    });

    return sheetToObjects(normalizedRows);
  }

  function parseCellValue(cellNode, sharedStrings) {
    const type = cellNode.getAttribute("t") || "n";
    const valueNode = getElementsByLocalName(cellNode, "v")[0];
    const inlineTextNode = getElementsByLocalName(cellNode, "t")[0];
    const raw = valueNode?.textContent ?? inlineTextNode?.textContent ?? "";

    if (type === "s") {
      return sharedStrings[Number(raw)] ?? "";
    }
    if (type === "b") {
      return raw === "1";
    }
    if (type === "n") {
      const numeric = Number(raw);
      return Number.isFinite(numeric) ? numeric : raw;
    }
    return raw;
  }

  function sheetToObjects(rows) {
    const headers = (rows[0] || []).map((header) => sanitizeString(header));
    return rows
      .slice(1)
      .map((row) => {
        const record = {};
        headers.forEach((header, index) => {
          if (header) {
            record[header] = row[index];
          }
        });
        return record;
      })
      .filter((record) => Object.values(record).some((value) => sanitizeString(value) !== ""));
  }

  function columnNameToIndex(columnName) {
    let result = 0;
    for (const char of columnName.toUpperCase()) {
      result = result * 26 + (char.charCodeAt(0) - 64);
    }
    return Math.max(result - 1, 0);
  }

  function clearSelection() {
    state.anchorId = "";
    state.anchorKind = "";
    state.targetId = "";
    state.targetKind = "";
    clearDistanceGraphics();
    renderAll();
    renderMapObjects();
  }

  function handleAnchorOrTargetClick(item, kind) {
    if (state.anchorId === item.id && state.anchorKind === kind) {
      clearSelection();
      return;
    }
    if (state.targetId === item.id && state.targetKind === kind) {
      clearSelection();
      return;
    }
    if (state.anchorId) {
      state.targetId = item.id;
      state.targetKind = kind;
    } else {
      state.anchorId = item.id;
      state.anchorKind = kind;
      state.targetId = "";
      state.targetKind = "";
    }
    renderAll();
    renderMapObjects();
  }

  function handleFireMarkerClick(fire) {
    if (state.targetId === fire.id && state.targetKind === "fire") {
      clearSelection();
      return;
    }
    if (state.anchorId) {
      state.targetId = fire.id;
      state.targetKind = "fire";
    } else {
      state.anchorId = "";
      state.anchorKind = "";
      state.targetId = fire.id;
      state.targetKind = "fire";
    }
    renderAll();
    renderMapObjects();
  }

  function toggleDockPopover() {
    const isVisible = !dom.dockPopover.classList.contains("is-hidden");
    if (isVisible) {
      closeDockPopover();
      return;
    }
    openDockPopover();
  }

  function openDockPopover() {
    dom.dockPopover.classList.remove("is-hidden");
    dom.dockPopover.setAttribute("aria-hidden", "false");
    dom.toolsBtn.classList.add("is-active");
  }

  function closeDockPopover() {
    dom.dockPopover.classList.add("is-hidden");
    dom.dockPopover.setAttribute("aria-hidden", "true");
    dom.toolsBtn.classList.remove("is-active");
  }

  function handleWindowResize() {
    window.clearTimeout(state.resizeTimer);
    state.resizeTimer = window.setTimeout(() => {
      if (!state.map || !state.sdkLoaded) {
        return;
      }
      if (window.kakao?.maps?.event) {
        window.kakao.maps.event.trigger(state.map, "resize");
      }
      renderMapObjects();
    }, 120);
  }

  function renderAll() {
    renderSelectionSummary();
    renderMapOverlayMessage();
  }

  function renderSelectionSummary() {
    const mode = getSelectionMode();
    const anchor = getAnchor();
    const target = getTarget();

    const isVisible = mode !== "none";
    dom.selectionHud.classList.toggle("is-hidden", !isVisible);
    dom.selectionHud.setAttribute("aria-hidden", String(!isVisible));

    if (mode === "paired" && anchor && target) {
      const distance = haversineKm(anchor, target);
      paintSummaryBar({
        title: `${anchor.name} → ${target.name}`,
        meta: `${typeLabel(target)} · ${target.address}`,
        attrGrid: [["직선거리", `${distance.toFixed(2)} km`]],
        distances: [],
      });
      return;
    }

    if (mode === "fire" && target) {
      const ward = findWardForPoint(target.lng, target.lat);
      const attrGrid = [
        ["기관 유형", orgTypeLabel(target.orgType)],
        ["관할 위험물제조소등", ward ? `${countActiveBusinessesInWard(ward)}개` : ""],
      ];
      paintSummaryBar({
        title: target.name,
        meta: `소방관서 · ${target.address}`,
        attrGrid,
        distances: [],
      });
      return;
    }

    if (mode === "anchor" && anchor && state.anchorKind === "business") {
      const info = buildBusinessSelection(anchor);
      const attrGrid = [
        ["제조소등 구분", anchor.facilityType],
        ["종류(류별)", anchor.hazardClass],
        ["개소", anchor.facilityCount ? String(anchor.facilityCount) : ""],
        ["품명", anchor.itemName],
        ["지정수량 배수(합계)", anchor.quantityMultiple],
      ];
      const distances = [];
      if (info.nearestFire) {
        distances.push(["가까운 소방관서", `${info.nearestFire.target.name} ${info.nearestFire.distanceKm.toFixed(2)}km`]);
      }
      if (info.nearestReference) {
        distances.push(["기준지점", `${info.nearestReference.target.name} ${info.nearestReference.distanceKm.toFixed(2)}km`]);
      }
      paintSummaryBar({
        title: anchor.name,
        meta: `위험물제조소등 · ${anchor.address}`,
        attrGrid,
        distances,
      });
      return;
    }

    if (mode === "anchor" && anchor && state.anchorKind === "reference") {
      const info = buildReferenceSelection(anchor);
      const attrGrid = [
        ["유형", anchor.pointType || "기준지점"],
        ["가까운 위험물제조소등", info.nearestBusiness
          ? `${info.nearestBusiness.target.name} ${info.nearestBusiness.distanceKm.toFixed(2)}km`
          : ""],
        ["가까운 소방관서", info.nearestFire
          ? `${info.nearestFire.target.name} ${info.nearestFire.distanceKm.toFixed(2)}km`
          : ""],
      ];
      paintSummaryBar({
        title: anchor.name,
        meta: `기준지점 · ${anchor.address}`,
        attrGrid,
        distances: [],
      });
      return;
    }

    dom.selectionSummary.className = "selection-summary is-empty";
    dom.selectionSummary.textContent = "지도에서 기준지점 또는 위험물제조소등을 선택하면 요약 정보가 표시됩니다.";
  }

  function paintSummaryBar({ title, meta, attrs, attrGrid, distances }) {
    dom.selectionSummary.className = "selection-summary";
    dom.selectionSummary.textContent = "";

    const titleEl = document.createElement("p");
    titleEl.className = "selection-summary__title";
    titleEl.textContent = title;
    dom.selectionSummary.appendChild(titleEl);

    if (meta) {
      const metaEl = document.createElement("p");
      metaEl.className = "selection-summary__meta";
      metaEl.textContent = meta;
      dom.selectionSummary.appendChild(metaEl);
    }

    if (attrGrid) {
      const grid = buildAttrGrid(attrGrid);
      if (grid) dom.selectionSummary.appendChild(grid);
    } else if (attrs) {
      const attrRow = buildChipsRow(attrs);
      if (attrRow) dom.selectionSummary.appendChild(attrRow);
    }

    const distRow = buildChipsRow(distances, true);
    if (distRow) dom.selectionSummary.appendChild(distRow);
  }

  function buildChipsRow(pairs, isDistance) {
    const row = document.createElement("div");
    row.className = "selection-summary__chips" + (isDistance ? " selection-summary__chips--distance" : "");
    (pairs || []).forEach(([label, value]) => {
      if (value == null || value === "") return;
      const chip = document.createElement("span");
      chip.className = "selection-summary__chip";
      const labelSpan = document.createElement("span");
      labelSpan.className = "selection-summary__chip-label";
      labelSpan.textContent = label;
      chip.appendChild(labelSpan);
      chip.append(String(value));
      row.appendChild(chip);
    });
    return row.children.length ? row : null;
  }

  function buildAttrGrid(pairs) {
    const grid = document.createElement("div");
    grid.className = "selection-summary__grid";
    pairs.forEach(([label, value]) => {
      const labelCell = document.createElement("div");
      labelCell.className = "ssg-cell ssg-cell--label";
      labelCell.textContent = label;
      const valueCell = document.createElement("div");
      valueCell.className = "ssg-cell ssg-cell--value";
      valueCell.textContent = (value == null || value === "") ? "—" : String(value);
      grid.append(labelCell, valueCell);
    });
    return grid.children.length ? grid : null;
  }

  function renderMapOverlayMessage() {
    if (!state.kakaoKey) {
      dom.mapOverlayMessage.textContent = "최초 실행 시 카카오 JavaScript 키를 입력해 주세요.";
      dom.mapOverlayMessage.style.display = "block";
      return;
    }
    if (!state.sdkLoaded) {
      dom.mapOverlayMessage.textContent = "지도를 준비하는 중입니다.";
      dom.mapOverlayMessage.style.display = "block";
      return;
    }
    if (!state.data.references.length && !state.data.businesses.length && !state.data.fireOrgs.length) {
      dom.mapOverlayMessage.textContent = "표준 템플릿 엑셀을 업로드하면 지도와 거리표가 표시됩니다.";
      dom.mapOverlayMessage.style.display = "block";
      return;
    }
    dom.mapOverlayMessage.style.display = "none";
  }

  function getSelectionMode() {
    if (state.anchorId && state.targetId) return "paired";
    if (state.anchorId) return "anchor";
    if (state.targetId && state.targetKind === "fire") return "fire";
    return "none";
  }

  function getAnchor() {
    if (!state.anchorId) return null;
    if (state.anchorKind === "reference") {
      return state.data.references.find((r) => r.id === state.anchorId) || null;
    }
    if (state.anchorKind === "business") {
      return state.data.businesses.find((b) => b.id === state.anchorId) || null;
    }
    return null;
  }

  function getTarget() {
    if (!state.targetId) return null;
    if (state.targetKind === "reference") {
      return state.data.references.find((r) => r.id === state.targetId) || null;
    }
    if (state.targetKind === "business") {
      return state.data.businesses.find((b) => b.id === state.targetId) || null;
    }
    if (state.targetKind === "fire") {
      return state.data.fireOrgs.find((f) => f.id === state.targetId) || null;
    }
    return null;
  }

  function renderMapObjects() {
    if (!state.map || !state.sdkLoaded) {
      return;
    }

    clearMapObjects();

    const bounds = new window.kakao.maps.LatLngBounds();
    let hasData = false;
    const highlightedJurisdictions = getHighlightedJurisdictionNames();

    if (ENABLE_JURISDICTION_LAYER) {
      state.data.jurisdictions.forEach((jurisdiction) => {
        const polygons = createJurisdictionPolygons(jurisdiction);
        polygons.forEach((polygon) => {
          const isHighlighted = jurisdictionMatches(jurisdiction.wardName, highlightedJurisdictions);
          polygon.setOptions({
            strokeWeight: isHighlighted ? 3 : 2,
            strokeColor: MAP_COLORS.fire,
            strokeOpacity: isHighlighted ? 0.9 : 0.34,
            fillColor: MAP_COLORS.fire,
            fillOpacity: 0,
            clickable: false,
            zIndex: 1,
          });
          polygon.setMap(state.map);
          state.rendered.jurisdictionPolygons.push(polygon);
        });
      });
    }

    state.data.references.forEach((reference) => {
      const position = new window.kakao.maps.LatLng(reference.lat, reference.lng);
      const isSelected = isMarkerSelected(reference);
      const markerSize = isSelected ? 48 : 40;
      const marker = new window.kakao.maps.Marker({
        map: state.map,
        position,
        title: reference.name,
        image: createMarkerImage(ICONS.reference, markerSize, markerSize, markerSize / 2, markerSize / 2),
        zIndex: 10,
      });
      window.kakao.maps.event.addListener(marker, "click", () => {
        handleAnchorOrTargetClick(reference, "reference");
      });
      state.rendered.markers.push(marker);
      bounds.extend(position);
      hasData = true;

    });

    const businessScale = buildValueScale(state.data.businesses.map(getBusinessSizeValue));

    state.data.businesses.forEach((business) => {
      const position = new window.kakao.maps.LatLng(business.lat, business.lng);
      const baseSize = scaleBusinessMarker(getBusinessSizeValue(business), businessScale);
      const markerSize = isMarkerSelected(business) ? baseSize + 8 : baseSize;
      const marker = new window.kakao.maps.Marker({
        map: state.map,
        position,
        title: business.name,
        image: createMarkerImage(ICONS.business, markerSize, markerSize, markerSize / 2, markerSize / 2),
        zIndex: 10,
      });
      window.kakao.maps.event.addListener(marker, "click", () => {
        handleAnchorOrTargetClick(business, "business");
      });
      state.rendered.markers.push(marker);
      bounds.extend(position);
      hasData = true;

    });

    state.data.fireOrgs.forEach((fireOrg) => {
      const position = new window.kakao.maps.LatLng(fireOrg.lat, fireOrg.lng);
      const isSelected = isMarkerSelected(fireOrg);
      const markerSize = isSelected ? 42 : 34;
      const marker = new window.kakao.maps.Marker({
        map: state.map,
        position,
        title: fireOrg.name,
        image: createMarkerImage(ICONS.fire, markerSize, markerSize, markerSize / 2, markerSize / 2),
        zIndex: 10,
      });
      window.kakao.maps.event.addListener(marker, "click", () => {
        handleFireMarkerClick(fireOrg);
      });
      state.rendered.markers.push(marker);
      bounds.extend(position);
      hasData = true;

    });

    if (hasData && state.shouldFitDataBounds) {
      state.map.setBounds(bounds, 56, 56, getMapBottomPadding(), 56);
      state.shouldFitDataBounds = false;
    } else {
      if (!hasData) {
        state.map.setCenter(new window.kakao.maps.LatLng(DEFAULT_CENTER.lat, DEFAULT_CENTER.lng));
        state.map.setLevel(9);
      }
    }

    renderPointLabels();
    renderSelectionGraphics();
  }

  function clearMapObjects() {
    state.rendered.markers.forEach((marker) => marker.setMap(null));
    clearPointLabels();
    state.rendered.jurisdictionPolygons.forEach((polygon) => polygon.setMap(null));
    clearDistanceGraphics();
    state.rendered.markers = [];
    state.rendered.jurisdictionPolygons = [];
  }

  function clearPointLabels() {
    state.rendered.referenceLabels.forEach((overlay) => overlay.setMap(null));
    state.rendered.businessLabels.forEach((overlay) => overlay.setMap(null));
    state.rendered.fireLabels.forEach((overlay) => overlay.setMap(null));
    state.rendered.referenceLabels = [];
    state.rendered.businessLabels = [];
    state.rendered.fireLabels = [];
  }

  // zoom 레벨이 이 값 이상이면 위험물제조소등 라벨 숨김
  const HIDE_BUSINESS_LABEL_LEVEL = 7;
  // zoom 레벨이 이 값 이상이면 소방관서 라벨을 약식으로 표시
  const SHORTEN_FIRE_LABEL_LEVEL = 7;

  function shortenFireLabelName(name) {
    return sanitizeString(name)
      .replace(/119안전센터$/, "119")
      .replace(/119지역대$/, "119")
      .replace(/119구조대$/, "구조")
      .trim() || name;
  }

  function renderPointLabels() {
    if (!state.map || !state.sdkLoaded) {
      return;
    }

    clearPointLabels();

    const currentLevel = state.map.getLevel();
    const zoomStage = getLabelZoomStage(currentLevel);

    state.data.references.forEach((reference) => {
      const overlay = new window.kakao.maps.CustomOverlay({
        position: new window.kakao.maps.LatLng(reference.lat, reference.lng),
        yAnchor: 1.62,
        content: createOverlayElement("reference-label", reference.name, zoomStage, isMarkerSelected(reference)),
        clickable: false,
      });
      overlay.setMap(state.map);
      state.rendered.referenceLabels.push(overlay);
    });

    if (currentLevel < HIDE_BUSINESS_LABEL_LEVEL) {
      state.data.businesses.forEach((business) => {
        const overlay = new window.kakao.maps.CustomOverlay({
          position: new window.kakao.maps.LatLng(business.lat, business.lng),
          yAnchor: 1.7,
          content: createOverlayElement("business-label", business.name, zoomStage, isMarkerSelected(business)),
          clickable: false,
        });
        overlay.setMap(state.map);
        state.rendered.businessLabels.push(overlay);
      });
    }

    state.data.fireOrgs.forEach((fireOrg) => {
      const displayName = currentLevel >= SHORTEN_FIRE_LABEL_LEVEL
        ? shortenFireLabelName(fireOrg.name)
        : fireOrg.name;
      const overlay = new window.kakao.maps.CustomOverlay({
        position: new window.kakao.maps.LatLng(fireOrg.lat, fireOrg.lng),
        yAnchor: 1.7,
        content: createOverlayElement("fire-label", displayName, zoomStage, isMarkerSelected(fireOrg)),
        clickable: false,
      });
      overlay.setMap(state.map);
      state.rendered.fireLabels.push(overlay);
    });
  }

  function clearDistanceGraphics() {
    state.rendered.circles.forEach((circle) => circle.setMap(null));
    state.rendered.lines.forEach((line) => line.setMap(null));
    state.rendered.lineLabels.forEach((label) => label.setMap(null));
    state.rendered.circles = [];
    state.rendered.lines = [];
    state.rendered.lineLabels = [];
  }

  function renderSelectionGraphics() {
    clearDistanceGraphics();

    const mode = getSelectionMode();
    const anchor = getAnchor();
    const target = getTarget();

    if (mode === "paired" && anchor && target) {
      drawPairLine(anchor, target);
      return;
    }

    if (mode === "anchor" && anchor) {
      if (state.anchorKind === "reference") {
        const info = buildReferenceSelection(anchor);
        if (info.nearestBusiness) {
          drawDistanceBundle(anchor, info.nearestBusiness.target, info.nearestBusiness.distanceKm, "business");
        }
        if (info.nearestFire) {
          drawDistanceBundle(anchor, info.nearestFire.target, info.nearestFire.distanceKm, "fire", {
            labelOffsetY: -0.6,
          });
        }
      } else if (state.anchorKind === "business") {
        const info = buildBusinessSelection(anchor);
        if (info.nearestFire) {
          drawDistanceBundle(anchor, info.nearestFire.target, info.nearestFire.distanceKm, "fire");
        }
      }
      return;
    }
  }

  function getMapBottomPadding() {
    const dockHeight = dom.bottomDock?.offsetHeight || 0;
    return Math.max(140, Math.min(dockHeight + 32, Math.floor(window.innerHeight * 0.58)));
  }

  function createOverlayElement(className, text, zoomStage = "mid", isSelected = false) {
    const wrapper = document.createElement("div");
    wrapper.className = `${className} map-label map-label--${zoomStage}${isSelected ? " is-selected" : ""}`;
    wrapper.textContent = text;
    return wrapper;
  }

  function getLabelZoomStage(level) {
    if (level <= 4) {
      return "near";
    }
    if (level <= 6) {
      return "mid";
    }
    return "far";
  }

  function createJurisdictionPolygons(jurisdiction) {
    const coordinates = jurisdiction?.geometry?.coordinates || [];
    const geometryType = jurisdiction?.geometry?.type;

    if (geometryType === "Polygon") {
      return [new window.kakao.maps.Polygon({ path: convertPolygonPath(coordinates) })];
    }

    if (geometryType === "MultiPolygon") {
      return coordinates.map((polygonCoords) => new window.kakao.maps.Polygon({
        path: convertPolygonPath(polygonCoords),
      }));
    }

    return [];
  }

  function convertPolygonPath(coordinates) {
    return coordinates.map((ring) => ring.map(([lng, lat]) => new window.kakao.maps.LatLng(lat, lng)));
  }

  function getHighlightedJurisdictionNames() {
    const names = new Set();
    const mode = getSelectionMode();
    const anchor = getAnchor();
    const target = getTarget();

    if (mode === "fire" && target) {
      names.add(normalizeName(target.name));
      return names;
    }

    if (mode === "paired" && target?.kind === "fire") {
      names.add(normalizeName(target.name));
      return names;
    }

    if (mode === "anchor" && anchor) {
      if (state.anchorKind === "reference") {
        const info = buildReferenceSelection(anchor);
        if (info.nearestFire?.target?.name) {
          names.add(normalizeName(info.nearestFire.target.name));
        }
      } else if (state.anchorKind === "business") {
        const info = buildBusinessSelection(anchor);
        if (info.nearestFire?.target?.name) {
          names.add(normalizeName(info.nearestFire.target.name));
        }
      }
    }

    return names;
  }

  function drawDistanceBundle(from, to, distanceKm, kind, options = {}) {
    if (!state.map || !from || !to) {
      return;
    }

    const color = kind === "fire" ? MAP_COLORS.fire : MAP_COLORS.business;
    const fillOpacity = kind === "fire" ? 0.08 : 0.1;
    const strokeStyle = kind === "fire" ? "ShortDash" : "Solid";
    const fromPoint = new window.kakao.maps.LatLng(from.lat, from.lng);
    const toPoint = new window.kakao.maps.LatLng(to.lat, to.lng);

    const circle = new window.kakao.maps.Circle({
      center: fromPoint,
      radius: Math.round(distanceKm * 1000),
      strokeWeight: 2,
      strokeColor: color,
      strokeOpacity: 0.6,
      strokeStyle,
      fillColor: color,
      fillOpacity,
    });
    circle.setMap(state.map);
    state.rendered.circles.push(circle);

    const line = new window.kakao.maps.Polyline({
      map: state.map,
      path: [fromPoint, toPoint],
      strokeWeight: 4,
      strokeColor: color,
      strokeOpacity: 0.86,
      strokeStyle,
    });
    state.rendered.lines.push(line);

    const midpoint = new window.kakao.maps.LatLng(
      (from.lat + to.lat) / 2,
      (from.lng + to.lng) / 2,
    );

    const label = new window.kakao.maps.CustomOverlay({
      map: state.map,
      position: midpoint,
      content: createOverlayElement(`line-distance-label line-distance-label--${kind}`, `${distanceKm.toFixed(2)} km`),
      yAnchor: options.labelOffsetY ?? 1.5,
      clickable: false,
    });
    state.rendered.lineLabels.push(label);
  }

  function drawPairLine(from, to) {
    if (!state.map || !from || !to) return;

    const fromPoint = new window.kakao.maps.LatLng(from.lat, from.lng);
    const toPoint = new window.kakao.maps.LatLng(to.lat, to.lng);
    const distanceKm = haversineKm(from, to);

    const line = new window.kakao.maps.Polyline({
      map: state.map,
      path: [fromPoint, toPoint],
      strokeWeight: 4,
      strokeColor: "#334155",
      strokeOpacity: 0.86,
      strokeStyle: "Solid",
    });
    state.rendered.lines.push(line);

    const midpoint = new window.kakao.maps.LatLng(
      (from.lat + to.lat) / 2,
      (from.lng + to.lng) / 2,
    );

    const label = new window.kakao.maps.CustomOverlay({
      map: state.map,
      position: midpoint,
      content: createOverlayElement("line-distance-label line-distance-label--pair", `${distanceKm.toFixed(2)} km`),
      yAnchor: 1.5,
      clickable: false,
    });
    state.rendered.lineLabels.push(label);
  }

  function getBusinessSizeValue(business) {
    const raw = sanitizeString(business?.quantityMultiple);
    if (!raw) return 0;
    const cleaned = raw.replace(/,/g, "").replace(/[^0-9.]/g, " ");
    const numbers = cleaned.split(/\s+/).map(Number).filter((v) => Number.isFinite(v) && v > 0);
    if (!numbers.length) return 0;
    return Math.max(...numbers);
  }

  function buildValueScale(values) {
    const valid = values.filter((value) => Number.isFinite(value) && value > 0).sort((a, b) => a - b);
    if (!valid.length) {
      return { min: 0, max: 1 };
    }
    return {
      min: valid[0],
      max: valid[valid.length - 1],
    };
  }

  function scaleBusinessMarker(value, scale) {
    const MIN_PX = 22;
    const MAX_PX = 64;
    if (!Number.isFinite(value) || value <= 0 || !Number.isFinite(scale.min) || scale.min <= 0) {
      return MIN_PX;
    }
    const ratio = value / scale.min;
    const multiplier = 1 + Math.log10(ratio);
    const maxMultiplier = MAX_PX / MIN_PX;
    const clamped = Math.min(Math.max(multiplier, 1), maxMultiplier);
    return Math.round(MIN_PX * clamped);
  }

  function createMarkerImage(source, width, height, offsetX, offsetY) {
    return new window.kakao.maps.MarkerImage(
      source,
      new window.kakao.maps.Size(width, height),
      {
        offset: new window.kakao.maps.Point(offsetX, offsetY),
      },
    );
  }

  function assetUrl(relativePath) {
    return new URL(relativePath, window.location.href).href;
  }

  function haversineKm(from, to) {
    const earthRadiusKm = 6371;
    const toRadians = (degree) => degree * (Math.PI / 180);
    const dLat = toRadians(to.lat - from.lat);
    const dLng = toRadians(to.lng - from.lng);
    const lat1 = toRadians(from.lat);
    const lat2 = toRadians(to.lat);

    const a = Math.sin(dLat / 2) ** 2
      + Math.cos(lat1) * Math.cos(lat2) * Math.sin(dLng / 2) ** 2;
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return earthRadiusKm * c;
  }

  function pointInRing(lng, lat, ring) {
    let inside = false;
    for (let i = 0, j = ring.length - 1; i < ring.length; j = i++) {
      const xi = ring[i][0];
      const yi = ring[i][1];
      const xj = ring[j][0];
      const yj = ring[j][1];
      const intersect = ((yi > lat) !== (yj > lat))
        && (lng < (xj - xi) * (lat - yi) / (yj - yi + 1e-12) + xi);
      if (intersect) inside = !inside;
    }
    return inside;
  }

  function pointInGeometry(lng, lat, geometry) {
    if (!geometry) return false;
    if (geometry.type === "Polygon") {
      const rings = geometry.coordinates;
      if (!rings?.length || !pointInRing(lng, lat, rings[0])) return false;
      for (let i = 1; i < rings.length; i++) {
        if (pointInRing(lng, lat, rings[i])) return false;
      }
      return true;
    }
    if (geometry.type === "MultiPolygon") {
      return geometry.coordinates.some((polygon) => {
        if (!polygon?.length || !pointInRing(lng, lat, polygon[0])) return false;
        for (let i = 1; i < polygon.length; i++) {
          if (pointInRing(lng, lat, polygon[i])) return false;
        }
        return true;
      });
    }
    return false;
  }

  function findWardForPoint(lng, lat) {
    return state.data.jurisdictions.find((ward) =>
      pointInGeometry(lng, lat, ward.geometry),
    ) || null;
  }

  function countActiveBusinessesInWard(ward) {
    if (!ward?.geometry) return 0;
    return state.data.businesses.reduce((acc, biz) => {
      if (!biz.active) return acc;
      return acc + (pointInGeometry(biz.lng, biz.lat, ward.geometry) ? 1 : 0);
    }, 0);
  }

  function typeLabel(target) {
    if (target.kind === "business") {
      return "위험물제조소등";
    }
    if (target.kind === "reference") {
      return "기준지점";
    }
    return "소방관서";
  }

  function orgTypeLabel(orgType) {
    const labels = {
      fire_station: "소방서",
      safety_center: "119안전센터",
      local_unit: "119지역대",
      rescue_unit: "119구조대",
      fire_org: "소방관서",
    };
    return labels[orgType] || "소방관서";
  }

  function showToast(message) {
    window.clearTimeout(state.toastTimer);
    dom.toast.textContent = message;
    dom.toast.classList.add("is-visible");
    state.toastTimer = window.setTimeout(() => {
      dom.toast.classList.remove("is-visible");
    }, 2800);
  }

  function nearestTarget(from, targets) {
    if (!from || !targets.length) {
      return null;
    }

    return targets
      .map((target) => ({
        target,
        distanceKm: haversineKm(from, target),
      }))
      .sort((left, right) => {
        if (left.distanceKm !== right.distanceKm) {
          return left.distanceKm - right.distanceKm;
        }
        return left.target.id.localeCompare(right.target.id);
      })[0];
  }

  function buildReferenceSelection(reference) {
    return {
      nearestBusiness: nearestTarget(reference, state.data.businesses),
      nearestFire: nearestTarget(reference, state.data.fireOrgs),
    };
  }

  function buildBusinessSelection(business) {
    return {
      nearestFire: nearestTarget(business, state.data.fireOrgs),
      nearestReference: nearestTarget(business, state.data.references),
    };
  }

  function isMarkerSelected(item) {
    if (!item) return false;
    if (state.anchorKind === item.kind && state.anchorId === item.id) return true;
    if (state.targetKind === item.kind && state.targetId === item.id) return true;
    return false;
  }

  function normalizeName(value) {
    return sanitizeString(value)
      .replace(/\s+/g, "")
      .replace(/[()]/g, "")
      .toLowerCase();
  }

  function jurisdictionMatches(wardName, highlightedNames) {
    const normalizedWard = normalizeName(wardName);
    if (!normalizedWard) {
      return false;
    }

    for (const name of highlightedNames) {
      if (name === normalizedWard || name.includes(normalizedWard) || normalizedWard.includes(name)) {
        return true;
      }
    }
    return false;
  }
})();
