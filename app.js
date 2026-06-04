import {
  buildGminaCatalog,
  buildPowiatCatalog,
  cleanText,
  haversineKm,
  makeLineFeature,
  makePolygonFeature,
  parseGeoJson,
} from "./geo-utils.js";
import { buildWorkbookBlob, readWorkbookRows } from "./xlsx-utils.js";
import { createZip, textToBytes } from "./zip-utils.js";

const state = {
  dataset: null,
  displayToCode: new Map(),
  filteredItems: [],
  currentPowiatCode: "",
};

const dom = {
  gminaFile: document.querySelector("#gminaFile"),
  powiatFile: document.querySelector("#powiatFile"),
  flowFile: document.querySelector("#flowFile"),
  loadButton: document.querySelector("#loadButton"),
  loadSummary: document.querySelector("#loadSummary"),
  scopeMode: document.querySelector("#scopeMode"),
  flowMode: document.querySelector("#flowMode"),
  searchInput: document.querySelector("#searchInput"),
  selectionInput: document.querySelector("#selectionInput"),
  transferThreshold: document.querySelector("#transferThreshold"),
  rangeKm: document.querySelector("#rangeKm"),
  topN: document.querySelector("#topN"),
  validationText: document.querySelector("#validationText"),
  detailsText: document.querySelector("#detailsText"),
  exportButton: document.querySelector("#exportButton"),
  statusLine: document.querySelector("#statusLine"),
  powiatPickerWrap: document.querySelector("#powiatPickerWrap"),
  powiatChecklist: document.querySelector("#powiatChecklist"),
  powiatSummary: document.querySelector("#powiatSummary"),
  selectAllPowiat: document.querySelector("#selectAllPowiat"),
  clearAllPowiat: document.querySelector("#clearAllPowiat"),
};

const distanceCache = new Map();

wireEvents();
refreshSelectionOptions();

function wireEvents() {
  dom.loadButton.addEventListener("click", loadDataset);
  dom.scopeMode.addEventListener("change", onScopeChanged);
  dom.flowMode.addEventListener("change", refreshSelectionSummary);
  dom.searchInput.addEventListener("input", refreshSelectionOptions);
  dom.selectionInput.addEventListener("change", refreshSelectionSummary);
  dom.transferThreshold.addEventListener("input", refreshSelectionSummary);
  dom.rangeKm.addEventListener("input", refreshSelectionSummary);
  dom.topN.addEventListener("input", refreshSelectionSummary);
  dom.exportButton.addEventListener("click", exportBundle);
  dom.selectAllPowiat.addEventListener("click", () => setAllPowiatChecks(true));
  dom.clearAllPowiat.addEventListener("click", () => setAllPowiatChecks(false));
}

async function loadDataset() {
  try {
    setStatus("Loading files...");
    const gminaFile = dom.gminaFile.files[0];
    const powiatFile = dom.powiatFile.files[0];
    const flowFile = dom.flowFile.files[0];
    if (!gminaFile || !powiatFile || !flowFile) {
      throw new Error("Choose gmina GeoJSON, powiat GeoJSON, and flow workbook first.");
    }

    const [gminaText, powiatText, flowRows] = await Promise.all([
      gminaFile.text(),
      powiatFile.text(),
      readWorkbookRows(flowFile),
    ]);

    const gminas = buildGminaCatalog(parseGeoJson(gminaText));
    const powiats = buildPowiatCatalog(parseGeoJson(powiatText));
    const { gminaToPowiat, powiatToGminas, powiatWarnings } = buildPowiatLookup(gminas, powiats);
    const codeToName = new Map(Array.from(gminas.values()).map((record) => [record.code, record.name]));
    const flowLoad = loadFlowRows(flowRows, codeToName);
    const validation = validateDataset(gminas, flowLoad.flows, flowLoad.rawRows, flowLoad.normalizedCells, flowLoad.normalizedRows, flowFile.name, gminaFile.name);
    validation.warnings = [...powiatWarnings, ...flowLoad.warnings, ...validation.warnings];
    validation.infoLines.push(`Powiat file: ${powiatFile.name}`);
    validation.infoLines.push(`Powiat count: ${powiats.size}`);
    validation.infoLines.push(`Gmina count assigned to powiats: ${Array.from(powiatToGminas.values()).reduce((acc, arr) => acc + arr.length, 0)}`);

    state.dataset = {
      gminas,
      powiats,
      gminaToPowiat,
      powiatToGminas,
      flows: flowLoad.flows,
      validation,
      powiatChecks: new Map(),
    };
    distanceCache.clear();

    dom.loadSummary.textContent = `Loaded ${gminas.size} gminas, ${powiats.size} powiats, ${flowLoad.flows.length} aggregated flow rows.`;
    renderValidation();
    refreshSelectionOptions();
    setStatus("Dataset loaded.");
  } catch (error) {
    setStatus("Load failed.");
    dom.validationText.textContent = String(error.message || error);
  }
}

function buildPowiatLookup(gminas, powiats) {
  const gminaToPowiat = new Map();
  const powiatToGminas = new Map(Array.from(powiats.keys()).map((code) => [code, []]));
  const warnings = [];
  const unmatched = [];

  for (const code of Array.from(gminas.keys()).sort()) {
    const powiatCode = code.slice(0, 4);
    if (powiats.has(powiatCode)) {
      gminaToPowiat.set(code, powiatCode);
      powiatToGminas.get(powiatCode).push(code);
    } else {
      unmatched.push(code);
    }
  }

  const emptyPowiats = Array.from(powiatToGminas.entries()).filter(([, items]) => items.length === 0).map(([code]) => code);
  if (unmatched.length) {
    warnings.push(`${unmatched.length} gminas could not be assigned to a powiat from the TERYT prefix.`);
    warnings.push(`Example unmatched gminas: ${unmatched.slice(0, 10).join(", ")}`);
  }
  if (emptyPowiats.length) {
    warnings.push(`${emptyPowiats.length} powiats have no assigned gminas in the current catalog.`);
    warnings.push(`Example empty powiats: ${emptyPowiats.slice(0, 10).join(", ")}`);
  }

  return { gminaToPowiat, powiatToGminas, powiatWarnings: warnings };
}

function loadFlowRows(rows, codeToName) {
  const cleanedRows = rows.filter((row) => row.some((cell) => cleanText(cell)));
  if (!cleanedRows.length) {
    return { flows: [], rawRows: 0, normalizedCells: 0, normalizedRows: 0, warnings: [] };
  }

  const body = cleanedRows.slice(1);
  const warnings = [];
  let normalizedCells = 0;
  let normalizedRows = 0;
  const aggregated = new Map();

  for (const row of body) {
    const [resCodeRaw, resNameRaw, workCodeRaw, workNameRaw, transfersRaw] = row;
    const transfers = Number.parseInt(cleanText(transfersRaw), 10) || 0;
    const res = normalizeExcelCode(resCodeRaw, codeToName);
    const work = normalizeExcelCode(workCodeRaw, codeToName);

    if (res.status !== "direct_match") {
      normalizedCells += 1;
    }
    if (work.status !== "direct_match") {
      normalizedCells += 1;
    }
    if (res.status !== "direct_match" || work.status !== "direct_match") {
      normalizedRows += 1;
    }
    if (res.status.startsWith("flagged")) {
      warnings.push(`Unmatched residence code in workbook: ${JSON.stringify(cleanText(resCodeRaw))} (${JSON.stringify(cleanText(resNameRaw))})`);
    }
    if (work.status.startsWith("flagged")) {
      warnings.push(`Unmatched work code in workbook: ${JSON.stringify(cleanText(workCodeRaw))} (${JSON.stringify(cleanText(workNameRaw))})`);
    }

    const fromName = res.name || cleanText(resNameRaw);
    const toName = work.name || cleanText(workNameRaw);
    const key = [res.code, fromName, work.code, toName].join("|");
    const current = aggregated.get(key) || {
      fromCode: res.code,
      fromName,
      toCode: work.code,
      toName,
      transfers: 0,
    };
    current.transfers += transfers;
    aggregated.set(key, current);
  }

  const flows = Array.from(aggregated.values()).sort((a, b) =>
    a.fromName.localeCompare(b.fromName) ||
    a.toName.localeCompare(b.toName) ||
    b.transfers - a.transfers
  );

  return {
    flows,
    rawRows: body.length,
    normalizedCells,
    normalizedRows,
    warnings: dedupe(warnings),
  };
}

function normalizeExcelCode(rawCode, codeToName) {
  const code = cleanText(rawCode);
  if (!code) {
    return { code: "", name: "", status: "flagged_empty" };
  }
  if (codeToName.has(code)) {
    return { code, name: codeToName.get(code), status: "direct_match" };
  }
  if (code.length === 7 && ["2", "4", "5"].includes(code.at(-1))) {
    const parent = `${code.slice(0, -1)}3`;
    if (codeToName.has(parent)) {
      return { code: parent, name: codeToName.get(parent), status: "rolled_up_to_parent" };
    }
  }
  return { code, name: "", status: "flagged_no_match" };
}

function validateDataset(gminas, flows, rawRows, normalizedCells, normalizedRows, excelFileName, gminaFileName) {
  const warnings = [];
  const infoLines = [];
  const flowCodes = new Set();
  for (const row of flows) {
    flowCodes.add(String(row.fromCode));
    flowCodes.add(String(row.toCode));
  }
  const gpkgCodes = new Set(gminas.keys());

  const missingInExcel = Array.from(gpkgCodes)
    .filter((code) => !flowCodes.has(code))
    .sort()
    .map((code) => [code, gminas.get(code).name]);
  const missingInGmina = Array.from(flowCodes)
    .filter((code) => !gpkgCodes.has(code))
    .sort()
    .map((code) => [code, ""]);

  infoLines.push(`Flow file: ${excelFileName}`);
  infoLines.push(`Gmina file: ${gminaFileName}`);
  infoLines.push(`Gmina count: ${gminas.size}`);
  infoLines.push(`Distinct normalized flow codes: ${flowCodes.size}`);
  infoLines.push(`Loaded workbook rows: ${rawRows}`);
  infoLines.push(`Normalized code cells: ${normalizedCells}`);
  infoLines.push(`Rows requiring normalization: ${normalizedRows}`);

  if (missingInExcel.length) {
    warnings.push(`${missingInExcel.length} gmina codes exist in boundaries but not in the workbook.`);
    warnings.push(`Example missing in workbook: ${missingInExcel.slice(0, 10).map(([code, name]) => `${code} (${name})`).join(", ")}`);
  }
  if (missingInGmina.length) {
    warnings.push(`${missingInGmina.length} flow codes exist in workbook but not in the gmina boundaries.`);
    warnings.push(`Example missing in boundaries: ${missingInGmina.slice(0, 10).map(([code]) => code).join(", ")}`);
  }
  if (!warnings.length) {
    warnings.push("Validation completed successfully: all normalized workbook codes matched the loaded gmina catalog.");
  }

  return {
    warnings,
    infoLines,
    missingInExcel,
    missingInGmina,
  };
}

function renderValidation() {
  if (!state.dataset) {
    dom.validationText.textContent = "No dataset loaded.";
    return;
  }
  const lines = [
    "Validation report",
    "",
    ...state.dataset.validation.infoLines.map((line) => `- ${line}`),
    "",
    "Warnings",
    ...state.dataset.validation.warnings.map((line) => `- ${line}`),
  ];
  dom.validationText.textContent = lines.join("\n");
}

function refreshSelectionOptions() {
  const dataset = state.dataset;
  const scope = dom.scopeMode.value;
  const search = normalizeSearch(dom.searchInput.value);
  state.displayToCode.clear();

  const items = [];
  if (dataset) {
    const source = scope === "powiat" ? Array.from(dataset.powiats.values()) : Array.from(dataset.gminas.values());
    for (const row of source.sort((a, b) => a.code.localeCompare(b.code))) {
      const label = `${row.code} - ${row.name}`;
      if (!search || normalizeSearch(label).includes(search)) {
        items.push(label);
        state.displayToCode.set(label, row.code);
      }
    }
  }

  const previous = dom.selectionInput.value;
  dom.selectionInput.innerHTML = "";
  if (!items.length) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = dataset ? "No matching items" : "Load data first";
    dom.selectionInput.append(option);
  } else {
    items.forEach((item) => {
      const option = document.createElement("option");
      option.value = item;
      option.textContent = item;
      dom.selectionInput.append(option);
    });
  }

  if (items.includes(previous)) {
    dom.selectionInput.value = previous;
  }
  state.filteredItems = items;
  onScopeChanged();
}

function onScopeChanged() {
  if (dom.scopeMode.value === "powiat") {
    dom.powiatPickerWrap.classList.remove("hidden");
    rebuildPowiatChecklist();
  } else {
    dom.powiatPickerWrap.classList.add("hidden");
  }
  refreshSelectionSummary();
}

function rebuildPowiatChecklist() {
  const dataset = state.dataset;
  dom.powiatChecklist.innerHTML = "";
  if (!dataset) {
    dom.powiatSummary.textContent = "Choose a powiat to manage included gminas.";
    return;
  }

  const selectedCode = getSelectedCode();
  if (!selectedCode || !dataset.powiats.has(selectedCode)) {
    dom.powiatSummary.textContent = "Choose a powiat to manage included gminas.";
    return;
  }

  const gminaCodes = dataset.powiatToGminas.get(selectedCode) || [];
  if (!dataset.powiatChecks.has(selectedCode)) {
    dataset.powiatChecks.set(selectedCode, new Map(gminaCodes.map((code) => [code, true])));
  }
  state.currentPowiatCode = selectedCode;
  const checkMap = dataset.powiatChecks.get(selectedCode);

  for (const code of gminaCodes) {
    const gmina = dataset.gminas.get(code);
    const label = document.createElement("label");
    label.className = "check-item";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = Boolean(checkMap.get(code));
    checkbox.addEventListener("change", () => {
      checkMap.set(code, checkbox.checked);
      refreshSelectionSummary();
    });
    const span = document.createElement("span");
    span.textContent = `${gmina.code} - ${gmina.name}`;
    label.append(checkbox, span);
    dom.powiatChecklist.append(label);
  }

  refreshPowiatSummary();
}

function refreshPowiatSummary() {
  const dataset = state.dataset;
  if (!dataset || !state.currentPowiatCode || !dataset.powiatChecks.has(state.currentPowiatCode)) {
    return;
  }
  const checkMap = dataset.powiatChecks.get(state.currentPowiatCode);
  const selected = Array.from(checkMap.values()).filter(Boolean).length;
  dom.powiatSummary.textContent = `Selected ${selected} / ${checkMap.size} gminas.`;
}

function setAllPowiatChecks(value) {
  const dataset = state.dataset;
  if (!dataset || !state.currentPowiatCode) {
    return;
  }
  const checkMap = dataset.powiatChecks.get(state.currentPowiatCode);
  for (const code of checkMap.keys()) {
    checkMap.set(code, value);
  }
  rebuildPowiatChecklist();
  refreshSelectionSummary();
}

function refreshSelectionSummary() {
  const dataset = state.dataset;
  if (!dataset) {
    dom.detailsText.textContent = "Load data to inspect the current selection.";
    return;
  }

  const selectionCode = getSelectedCode();
  if (!selectionCode) {
    dom.detailsText.textContent = "Choose a gmina or powiat.";
    return;
  }

  const filterState = getFilterState();
  if (filterState.error) {
    dom.detailsText.textContent = filterState.error;
    return;
  }

  if (dom.scopeMode.value === "powiat") {
    rebuildPowiatChecklist();
    const powiat = dataset.powiats.get(selectionCode);
    const selectedGminas = getSelectedPowiatGminas(selectionCode);
    const allGminas = dataset.powiatToGminas.get(selectionCode) || [];
    const summaryRows = selectedGminas.map((code) =>
      getSelectionRows(dataset, code, dom.flowMode.value, filterState.transferThreshold, filterState.rangeKm, null)
    );
    const totalRows = summaryRows.reduce((sum, result) => sum + result.count, 0);
    const totalSelfFlow = summaryRows.reduce((sum, result) => sum + result.selfFlow, 0);
    dom.detailsText.textContent = [
      `Powiat: ${powiat.code} - ${powiat.name}`,
      `Selected gminas: ${selectedGminas.length} / ${allGminas.length}`,
      `Flow mode: ${dom.flowMode.value === "from" ? "from gmina" : "to gmina"}`,
      `Transfer filter: ${formatOptional(filterState.transferThreshold, ">", "none")}`,
      `Distance filter: ${formatOptional(filterState.rangeKm, "<= ", "none", " km")}`,
      `Top N: ${filterState.topN ?? "none"} (applied globally during export)`,
      `Total matching rows before Top N: ${totalRows}`,
      `Total internal self-flow: ${totalSelfFlow}`,
      `Example selected gminas: ${selectedGminas.slice(0, 10).map((code) => `${code} - ${dataset.gminas.get(code).name}`).join(", ") || "none"}`,
    ].join("\n");
    refreshPowiatSummary();
    return;
  }

  const result = getSelectionRows(dataset, selectionCode, dom.flowMode.value, filterState.transferThreshold, filterState.rangeKm, filterState.topN);
  const gmina = dataset.gminas.get(selectionCode);
  const powiatCode = dataset.gminaToPowiat.get(selectionCode) || "";
  const powiatName = powiatCode && dataset.powiats.has(powiatCode) ? dataset.powiats.get(powiatCode).name : "";
  dom.detailsText.textContent = [
    `Gmina: ${gmina.code} - ${gmina.name}`,
    `Powiat: ${powiatCode ? `${powiatCode} - ${powiatName}` : "none"}`,
    `Flow mode: ${dom.flowMode.value === "from" ? "from gmina" : "to gmina"}`,
    `Transfer filter: ${formatOptional(filterState.transferThreshold, ">", "none")}`,
    `Distance filter: ${formatOptional(filterState.rangeKm, "<= ", "none", " km")}`,
    `Top N: ${filterState.topN ?? "none"}`,
    `Rows after filters: ${result.count}`,
    `Internal self-flow: ${result.selfFlow}`,
    `Distinct partner gminas: ${new Set(result.rows.map((row) => row.otherCode)).size}`,
    `Excluded by transfer threshold: ${result.filteredByTransfer}`,
    `Excluded by distance filter: ${result.filteredByDistance}`,
    `Excluded by Top N: ${result.filteredByTopN}`,
  ].join("\n");
}

function getSelectionRows(dataset, selectedCode, mode, transferThreshold, rangeKm, topN) {
  const all = dataset.flows;
  let subset;
  let selfFlow;

  if (mode === "from") {
    subset = all.filter((row) => row.fromCode === selectedCode && row.toCode !== selectedCode);
    selfFlow = all.filter((row) => row.fromCode === selectedCode && row.toCode === selectedCode).reduce((sum, row) => sum + row.transfers, 0);
    subset = subset.map((row) => ({
      ...row,
      sourceCode: row.fromCode,
      sourceName: row.fromName,
      targetCode: row.toCode,
      targetName: row.toName,
      direction: "from",
      otherCode: row.toCode,
      otherName: row.toName,
    }));
  } else {
    subset = all.filter((row) => row.toCode === selectedCode && row.fromCode !== selectedCode);
    selfFlow = all.filter((row) => row.fromCode === selectedCode && row.toCode === selectedCode).reduce((sum, row) => sum + row.transfers, 0);
    subset = subset.map((row) => ({
      ...row,
      sourceCode: row.fromCode,
      sourceName: row.fromName,
      targetCode: row.toCode,
      targetName: row.toName,
      direction: "to",
      otherCode: row.fromCode,
      otherName: row.fromName,
    }));
  }

  const rawCount = subset.length;
  let filteredByTransfer = 0;
  let filteredByDistance = 0;
  let filteredByTopN = 0;

  if (transferThreshold !== null) {
    const before = subset.length;
    subset = subset.filter((row) => row.transfers > transferThreshold);
    filteredByTransfer = before - subset.length;
  }

  subset = subset.map((row) => ({
    ...row,
    distanceKm: centroidDistanceKm(dataset, selectedCode, row.otherCode),
  }));

  if (rangeKm !== null) {
    const before = subset.length;
    subset = subset.filter((row) => row.distanceKm <= rangeKm);
    filteredByDistance = before - subset.length;
  }

  subset.sort((a, b) =>
    b.transfers - a.transfers ||
    a.distanceKm - b.distanceKm ||
    a.otherName.localeCompare(b.otherName) ||
    a.otherCode.localeCompare(b.otherCode)
  );

  if (topN !== null) {
    const before = subset.length;
    subset = subset.slice(0, topN);
    filteredByTopN = before - subset.length;
  }

  return {
    rows: subset,
    count: subset.length,
    selfFlow,
    rawCount,
    filteredByTransfer,
    filteredByDistance,
    filteredByTopN,
  };
}

function centroidDistanceKm(dataset, codeA, codeB) {
  if (codeA === codeB) {
    return 0;
  }
  const key = [codeA, codeB].sort().join("|");
  if (distanceCache.has(key)) {
    return distanceCache.get(key);
  }
  const a = dataset.gminas.get(codeA);
  const b = dataset.gminas.get(codeB);
  if (!a || !b) {
    return Number.POSITIVE_INFINITY;
  }
  const distance = haversineKm(a.centroidX, a.centroidY, b.centroidX, b.centroidY);
  distanceCache.set(key, distance);
  return distance;
}

async function exportBundle() {
  try {
    const dataset = state.dataset;
    if (!dataset) {
      throw new Error("Load data before exporting.");
    }
    const selectionCode = getSelectedCode();
    if (!selectionCode) {
      throw new Error("Choose a gmina or powiat before exporting.");
    }

    const filterState = getFilterState();
    if (filterState.error) {
      throw new Error(filterState.error);
    }

    setStatus("Preparing export...");
    let rows = [];
    let selectedName = "";
    let contextLabel = "";
    let selfFlow = 0;
    let inputRows = 0;
    let selectionStats = new Map();

    if (dom.scopeMode.value === "powiat") {
      const powiat = dataset.powiats.get(selectionCode);
      const selectedGminas = getSelectedPowiatGminas(selectionCode);
      if (!selectedGminas.length) {
        throw new Error("Select at least one gmina in the chosen powiat.");
      }
      contextLabel = `powiat ${powiat.code} - ${powiat.name}`;
      selectedName = powiat.name;

      for (const code of selectedGminas) {
        const result = getSelectionRows(dataset, code, dom.flowMode.value, filterState.transferThreshold, filterState.rangeKm, null);
        rows.push(...result.rows.map((row) => ({
          ...row,
          selectedCode: code,
          selectedName: dataset.gminas.get(code).name,
        })));
        selfFlow += result.selfFlow;
        inputRows += result.rawCount + (result.selfFlow ? 1 : 0);
        selectionStats.set(code, [result.rawCount + (result.selfFlow ? 1 : 0), result.selfFlow]);
      }

      rows.sort((a, b) =>
        b.transfers - a.transfers ||
        a.distanceKm - b.distanceKm ||
        a.selectedCode.localeCompare(b.selectedCode) ||
        a.otherName.localeCompare(b.otherName) ||
        a.otherCode.localeCompare(b.otherCode)
      );
      if (filterState.topN !== null) {
        rows = rows.slice(0, filterState.topN);
      }
    } else {
      const result = getSelectionRows(dataset, selectionCode, dom.flowMode.value, filterState.transferThreshold, filterState.rangeKm, filterState.topN);
      const gmina = dataset.gminas.get(selectionCode);
      rows = result.rows.map((row) => ({
        ...row,
        selectedCode: selectionCode,
        selectedName: gmina.name,
      }));
      selectedName = gmina.name;
      contextLabel = `gmina ${gmina.code} - ${gmina.name}`;
      selfFlow = result.selfFlow;
      inputRows = result.rawCount + (result.selfFlow ? 1 : 0);
      selectionStats.set(selectionCode, [inputRows, result.selfFlow]);
    }

    const report = buildExportReport(dataset, rows, selectionCode, selectedName, contextLabel, selfFlow, inputRows, filterState, selectionStats);
    const lineGeoJson = buildLineGeoJson(dataset, rows);
    const polygonGeoJson = buildPolygonGeoJson(dataset, rows, contextLabel);
    const workbook = await buildWorkbookBlob(buildWorkbookSheets(rows, report));
    const txtBytes = textToBytes(report.summaryLines.join("\n") + "\n");
    const timestamp = timestampSlug();
    const baseName = `${timestamp}_${safeFilename(selectionCode || "scope")}_${safeFilename(selectedName || contextLabel || "export")}`;

    const bundle = await createZip([
      { name: `${baseName}/${baseName}_lines.geojson`, data: textToBytes(JSON.stringify(lineGeoJson, null, 2)) },
      { name: `${baseName}/${baseName}_polygons.geojson`, data: textToBytes(JSON.stringify(polygonGeoJson, null, 2)) },
      { name: `${baseName}/${baseName}.txt`, data: txtBytes },
      { name: `${baseName}/${baseName}.xlsx`, data: new Uint8Array(await workbook.arrayBuffer()) },
    ]);

    downloadBlob(new Blob([bundle], { type: "application/zip" }), `${baseName}.zip`);
    setStatus("Export ready.");
  } catch (error) {
    setStatus("Export failed.");
    alert(String(error.message || error));
  }
}

function buildLineGeoJson(dataset, rows) {
  const features = [];
  for (const row of rows) {
    const source = dataset.gminas.get(row.sourceCode);
    const target = dataset.gminas.get(row.targetCode);
    if (!source || !target) {
      continue;
    }
    features.push(makeLineFeature(source, target, {
      source_code: row.sourceCode,
      source_name: row.sourceName,
      target_code: row.targetCode,
      target_name: row.targetName,
      distance_km: round(row.distanceKm, 3),
      transfers: row.transfers,
      direction: row.direction === "from" ? "from gmina" : "to gmina",
      selected_code: row.selectedCode,
      selected_name: row.selectedName,
    }));
  }
  return { type: "FeatureCollection", features };
}

function buildPolygonGeoJson(dataset, rows, contextLabel) {
  const connected = buildConnectedGminaTable(dataset, rows, contextLabel);
  return {
    type: "FeatureCollection",
    features: connected.map((row) => makePolygonFeature(dataset.gminas.get(row.code), {
      code: row.code,
      name: row.name,
      connection_count: row.connectionCount,
      total_transfers: row.totalTransfers,
      outgoing_transfers: row.outgoingTransfers,
      incoming_transfers: row.incomingTransfers,
      partners: row.partners.join("; "),
      scope: contextLabel,
    })),
  };
}

function buildConnectedGminaTable(dataset, rows, contextLabel) {
  const connectedCodes = new Set();
  rows.forEach((row) => {
    connectedCodes.add(row.sourceCode);
    connectedCodes.add(row.targetCode);
  });

  const records = [];
  for (const code of Array.from(connectedCodes).sort()) {
    const gmina = dataset.gminas.get(code);
    if (!gmina) {
      continue;
    }
    const outgoing = rows.filter((row) => row.sourceCode === code);
    const incoming = rows.filter((row) => row.targetCode === code);
    const seen = new Set();
    const partners = [];
    for (const row of [...outgoing, ...incoming]) {
      const partnerCode = row.sourceCode === code ? row.targetCode : row.sourceCode;
      const partnerName = row.sourceCode === code ? row.targetName : row.sourceName;
      const key = `${partnerCode}|${partnerName}`;
      if (partnerCode === code || seen.has(key)) {
        continue;
      }
      seen.add(key);
      partners.push(`${partnerCode} - ${partnerName}`);
    }
    records.push({
      code,
      name: gmina.name,
      connectionCount: partners.length,
      totalTransfers: outgoing.reduce((sum, row) => sum + row.transfers, 0) + incoming.reduce((sum, row) => sum + row.transfers, 0),
      outgoingTransfers: outgoing.reduce((sum, row) => sum + row.transfers, 0),
      incomingTransfers: incoming.reduce((sum, row) => sum + row.transfers, 0),
      partners,
      scope: contextLabel,
    });
  }

  return records.sort((a, b) =>
    b.totalTransfers - a.totalTransfers ||
    b.connectionCount - a.connectionCount ||
    a.name.localeCompare(b.name)
  );
}

function buildExportReport(dataset, rows, selectedCode, selectedName, contextLabel, selfFlow, inputRows, filters, selectionStats) {
  const totalTransfers = rows.reduce((sum, row) => sum + row.transfers, 0);
  const uniqueSelected = new Set(rows.map((row) => row.selectedCode)).size;
  const uniqueOther = new Set(rows.map((row) => row.otherCode)).size;
  const avgTransfers = rows.length ? totalTransfers / rows.length : 0;
  const distances = rows.map((row) => row.distanceKm);
  const weightedDistance = totalTransfers > 0
    ? rows.reduce((sum, row) => sum + (row.distanceKm * row.transfers), 0) / totalTransfers
    : 0;
  const topRows = [...rows].sort((a, b) =>
    b.transfers - a.transfers ||
    a.distanceKm - b.distanceKm ||
    a.otherName.localeCompare(b.otherName) ||
    a.otherCode.localeCompare(b.otherCode)
  ).slice(0, 10);
  const networkMetrics = buildNetworkMetrics(rows, selfFlow, selectionStats);
  const scopeStats = buildScopeStats(rows, selfFlow, selectionStats);
  const statsItems = [
    ["Scope", selectedCode && selectedName ? `${selectedCode} - ${selectedName}` : contextLabel],
    ["Mode", dom.flowMode.value === "from" ? "from gmina" : "to gmina"],
    ["Transfer filter", filters.transferThreshold === null ? "none" : `> ${filters.transferThreshold}`],
    ["Distance filter", filters.rangeKm === null ? "none" : `<= ${filters.rangeKm.toFixed(2)} km`],
    ["Top N", filters.topN === null ? "none" : String(filters.topN)],
    ["Input rows", String(inputRows || rows.length + selfFlow)],
    ["Rows after filters", String(rows.length)],
    ["Total transfers", `${totalTransfers} persons`],
    ["Average transfers per row", avgTransfers.toFixed(2)],
    ["Median transfers", median(rows.map((row) => row.transfers)).toFixed(2)],
    ["Minimum transfers", String(rows.length ? Math.min(...rows.map((row) => row.transfers)) : 0)],
    ["Maximum transfers", String(rows.length ? Math.max(...rows.map((row) => row.transfers)) : 0)],
    ["Average distance", `${mean(distances).toFixed(2)} km`],
    ["Median distance", `${median(distances).toFixed(2)} km`],
    ["Weighted average distance", `${weightedDistance.toFixed(2)} km`],
    ["Minimum distance", `${(rows.length ? Math.min(...distances) : 0).toFixed(2)} km`],
    ["Maximum distance", `${(rows.length ? Math.max(...distances) : 0).toFixed(2)} km`],
    ["Unique selected gminas", String(uniqueSelected || (selectedCode ? 1 : 0))],
    ["Unique partner gminas", String(uniqueOther)],
    ["Internal self-flow", `${selfFlow} persons`],
    ["Missing in workbook", String(dataset.validation.missingInExcel.length)],
    ["Missing in boundaries", String(dataset.validation.missingInGmina.length)],
  ];

  const summaryLines = [
    `Run timestamp: ${new Date().toISOString()}`,
    `Scope: ${selectedCode && selectedName ? `${selectedCode} - ${selectedName}` : contextLabel}`,
    `Mode: ${dom.flowMode.value === "from" ? "from gmina" : "to gmina"}`,
    `Transfer filter: ${filters.transferThreshold === null ? "none" : `> ${filters.transferThreshold}`}`,
    `Distance filter: ${filters.rangeKm === null ? "none" : `<= ${filters.rangeKm.toFixed(2)} km`}`,
    `Top N: ${filters.topN === null ? "none" : filters.topN}`,
    "Note: in powiat mode, Top N is applied globally to the combined result.",
    "Distance is calculated from gmina centroids.",
    "",
    "Validation summary:",
    ...dataset.validation.infoLines.map((line) => `  ${line}`),
    `  Missing in workbook: ${dataset.validation.missingInExcel.length}`,
    `  Missing in boundaries: ${dataset.validation.missingInGmina.length}`,
    "",
    "Query stats:",
    ...statsItems.map(([label, value]) => `  ${label}: ${value}`),
    "",
    "Top flows:",
  ];
  if (!topRows.length) {
    summaryLines.push("  No rows remain after filtering.");
  } else {
    topRows.forEach((row, index) => {
      summaryLines.push(`  ${index + 1}. ${row.sourceCode} - ${row.sourceName} -> ${row.targetCode} - ${row.targetName}: ${row.transfers} persons, ${row.distanceKm.toFixed(2)} km`);
    });
  }
  summaryLines.push("");
  summaryLines.push("Warnings:");
  dataset.validation.warnings.forEach((warning) => summaryLines.push(`  - ${warning}`));

  return { summaryLines, statsItems, topRows, networkMetrics, scopeStats };
}

function buildScopeStats(rows, selfFlow, selectionStats) {
  if (!rows.length) {
    return [];
  }
  const grouped = groupBy(rows, (row) => `${row.selectedCode}|${row.selectedName}`);
  return Array.from(grouped.entries()).map(([key, list], index) => {
    const [selectedCode, selectedName] = key.split("|");
    const totalTransfers = list.reduce((sum, row) => sum + row.transfers, 0);
    return {
      lp: index + 1,
      selectedCode,
      selectedName,
      rowCount: list.length,
      totalTransfers,
      avgTransfers: totalTransfers / list.length,
      medianTransfers: median(list.map((row) => row.transfers)),
      totalDistanceKm: list.reduce((sum, row) => sum + row.distanceKm, 0),
      avgDistanceKm: mean(list.map((row) => row.distanceKm)),
      uniquePartners: new Set(list.map((row) => row.otherCode)).size,
      selfFlow: selectionStats.get(selectedCode)?.[1] ?? selfFlow,
    };
  }).sort((a, b) => b.totalTransfers - a.totalTransfers || b.rowCount - a.rowCount);
}

function buildNetworkMetrics(rows, selfFlow, selectionStats) {
  const stats = buildScopeStats(rows, selfFlow, selectionStats);
  if (!stats.length) {
    return [];
  }
  const maxTotal = Math.max(...stats.map((item) => item.totalTransfers), 1);
  const scopeDenom = Math.max(stats.length - 1, 1);
  return stats.map((item, index) => {
    const strongest = [...rows]
      .filter((row) => row.selectedCode === item.selectedCode)
      .sort((a, b) => b.transfers - a.transfers || a.distanceKm - b.distanceKm)[0];
    return {
      lp: index + 1,
      selectedCode: item.selectedCode,
      selectedName: item.selectedName,
      rowCount: item.rowCount,
      totalTransfers: item.totalTransfers,
      avgTransfers: item.avgTransfers,
      medianTransfers: item.medianTransfers,
      totalDistanceKm: item.totalDistanceKm,
      avgDistanceKm: item.avgDistanceKm,
      uniquePartners: item.uniquePartners,
      strongestPartnerCode: strongest?.otherCode || "",
      strongestPartnerName: strongest?.otherName || "",
      strongestPartnerTransfers: strongest?.transfers || 0,
      strongestPartnerDistanceKm: strongest?.distanceKm || 0,
      degreeCentrality: item.uniquePartners / scopeDenom,
      flowCentrality: item.totalTransfers / maxTotal,
      strongestShare: item.totalTransfers > 0 ? (strongest?.transfers || 0) / item.totalTransfers : 0,
      avgTransfersPerPartner: item.uniquePartners > 0 ? item.totalTransfers / item.uniquePartners : 0,
      selfFlow: item.selfFlow,
    };
  });
}

function buildWorkbookSheets(rows, report) {
  return [
    {
      name: "Flows",
      rows: [
        ["Source code", "Source name", "Target code", "Target name", "Distance km", "Transfers", "Direction", "Selected code", "Selected name"],
        ...rows.map((row) => [
          row.sourceCode,
          row.sourceName,
          row.targetCode,
          row.targetName,
          round(row.distanceKm, 3),
          row.transfers,
          row.direction === "from" ? "from gmina" : "to gmina",
          row.selectedCode,
          row.selectedName,
        ]),
      ],
    },
    {
      name: "Summary",
      rows: report.summaryLines.map((line) => [line]),
    },
    {
      name: "Query Stats",
      rows: [
        ["Metric", "Value"],
        ...report.statsItems,
      ],
    },
    {
      name: "Top Flows",
      rows: [
        ["Lp", "Source code", "Source name", "Target code", "Target name", "Transfers", "Distance km"],
        ...report.topRows.map((row, index) => [index + 1, row.sourceCode, row.sourceName, row.targetCode, row.targetName, row.transfers, round(row.distanceKm, 3)]),
      ],
    },
    {
      name: "Scope Stats",
      rows: [
        ["Lp", "Selected code", "Selected name", "Row count", "Total transfers", "Average transfers", "Median transfers", "Total distance km", "Average distance km", "Unique partners", "Self flow"],
        ...report.scopeStats.map((row) => [row.lp, row.selectedCode, row.selectedName, row.rowCount, row.totalTransfers, round(row.avgTransfers, 2), round(row.medianTransfers, 2), round(row.totalDistanceKm, 3), round(row.avgDistanceKm, 3), row.uniquePartners, row.selfFlow]),
      ],
    },
    {
      name: "Network Metrics",
      rows: [
        ["Lp", "Selected code", "Selected name", "Row count", "Total transfers", "Average transfers", "Median transfers", "Total distance km", "Average distance km", "Unique partners", "Strongest partner code", "Strongest partner name", "Strongest partner transfers", "Strongest partner distance km", "Degree centrality", "Flow centrality", "Strongest share", "Average transfers per partner", "Self flow"],
        ...report.networkMetrics.map((row) => [row.lp, row.selectedCode, row.selectedName, row.rowCount, row.totalTransfers, round(row.avgTransfers, 2), round(row.medianTransfers, 2), round(row.totalDistanceKm, 3), round(row.avgDistanceKm, 3), row.uniquePartners, row.strongestPartnerCode, row.strongestPartnerName, row.strongestPartnerTransfers, round(row.strongestPartnerDistanceKm, 3), round(row.degreeCentrality, 4), round(row.flowCentrality, 4), round(row.strongestShare, 4), round(row.avgTransfersPerPartner, 2), row.selfFlow]),
      ],
    },
  ];
}

function getSelectedCode() {
  const selected = dom.selectionInput.value;
  return state.displayToCode.get(selected) || cleanText(selected.split(" - ", 1)[0]);
}

function getSelectedPowiatGminas(powiatCode) {
  const dataset = state.dataset;
  if (!dataset) {
    return [];
  }
  const checkMap = dataset.powiatChecks.get(powiatCode);
  if (!checkMap) {
    return [];
  }
  return Array.from(checkMap.entries()).filter(([, checked]) => checked).map(([code]) => code);
}

function getFilterState() {
  const transferThreshold = parseOptionalInt(dom.transferThreshold.value, "Transfer threshold");
  if (transferThreshold.error) {
    return transferThreshold;
  }
  const rangeKm = parseOptionalFloat(dom.rangeKm.value, "Max distance");
  if (rangeKm.error) {
    return rangeKm;
  }
  const topN = parseOptionalInt(dom.topN.value, "Top N");
  if (topN.error) {
    return topN;
  }
  return {
    transferThreshold: transferThreshold.value,
    rangeKm: rangeKm.value,
    topN: topN.value,
    error: null,
  };
}

function parseOptionalInt(value, label) {
  const text = cleanText(value);
  if (!text) {
    return { value: null, error: null };
  }
  const parsed = Number.parseInt(text, 10);
  if (!Number.isFinite(parsed)) {
    return { value: null, error: `${label} must be an integer.` };
  }
  if (parsed < 0) {
    return { value: null, error: `${label} cannot be negative.` };
  }
  return { value: parsed, error: null };
}

function parseOptionalFloat(value, label) {
  const text = cleanText(value).replace(",", ".");
  if (!text) {
    return { value: null, error: null };
  }
  const parsed = Number.parseFloat(text);
  if (!Number.isFinite(parsed)) {
    return { value: null, error: `${label} must be numeric.` };
  }
  if (parsed < 0) {
    return { value: null, error: `${label} cannot be negative.` };
  }
  return { value: parsed, error: null };
}

function setStatus(text) {
  dom.statusLine.textContent = text;
}

function normalizeSearch(text) {
  return cleanText(text)
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/ł/gi, "l")
    .toLowerCase();
}

function formatOptional(value, prefix, empty, suffix = "") {
  return value === null ? empty : `${prefix}${value}${suffix}`;
}

function safeFilename(text) {
  return cleanText(text)
    .replace(/\s+/g, "_")
    .replace(/[^A-Za-z0-9_-]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^[_.-]+|[_.-]+$/g, "") || "export";
}

function timestampSlug() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, "0");
  return [
    now.getFullYear(),
    pad(now.getMonth() + 1),
    pad(now.getDate()),
    pad(now.getHours()),
    pad(now.getMinutes()),
    pad(now.getSeconds()),
  ].join("_");
}

function dedupe(items) {
  return Array.from(new Set(items));
}

function round(value, places) {
  return Number.isFinite(value) ? Number(value.toFixed(places)) : 0;
}

function mean(values) {
  return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : 0;
}

function median(values) {
  if (!values.length) {
    return 0;
  }
  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0
    ? (sorted[middle - 1] + sorted[middle]) / 2
    : sorted[middle];
}

function groupBy(items, getKey) {
  const map = new Map();
  for (const item of items) {
    const key = getKey(item);
    if (!map.has(key)) {
      map.set(key, []);
    }
    map.get(key).push(item);
  }
  return map;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
