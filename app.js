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
    setStatus("Trwa wczytywanie plików...");
    const gminaFile = dom.gminaFile.files[0];
    const powiatFile = dom.powiatFile.files[0];
    const flowFile = dom.flowFile.files[0];
    if (!gminaFile || !powiatFile || !flowFile) {
      throw new Error("Najpierw wybierz plik GeoJSON gmin, GeoJSON powiatów oraz skoroszyt przepływów.");
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
    validation.infoLines.push(`Plik powiatów: ${powiatFile.name}`);
    validation.infoLines.push(`Liczba powiatów: ${powiats.size}`);
    validation.infoLines.push(`Liczba gmin przypisanych do powiatów: ${Array.from(powiatToGminas.values()).reduce((acc, arr) => acc + arr.length, 0)}`);

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

    dom.loadSummary.textContent = `Wczytano ${gminas.size} gmin, ${powiats.size} powiatów oraz ${flowLoad.flows.length} zagregowanych wierszy przepływów.`;
    renderValidation();
    refreshSelectionOptions();
    setStatus("Dane zostały wczytane.");
  } catch (error) {
    setStatus("Wczytywanie nie powiodło się.");
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
    warnings.push(`${unmatched.length} gmin nie udało się przypisać do powiatu na podstawie prefiksu TERYT.`);
    warnings.push(`Przykład nieprzypisanych gmin: ${unmatched.slice(0, 10).join(", ")}`);
  }
  if (emptyPowiats.length) {
    warnings.push(`${emptyPowiats.length} powiatów nie ma przypisanych gmin w bieżącym katalogu.`);
    warnings.push(`Przykład pustych powiatów: ${emptyPowiats.slice(0, 10).join(", ")}`);
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
      warnings.push(`Niepasujący kod zamieszkania w skoroszycie: ${JSON.stringify(cleanText(resCodeRaw))} (${JSON.stringify(cleanText(resNameRaw))})`);
    }
    if (work.status.startsWith("flagged")) {
      warnings.push(`Niepasujący kod pracy w skoroszycie: ${JSON.stringify(cleanText(workCodeRaw))} (${JSON.stringify(cleanText(workNameRaw))})`);
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

  infoLines.push(`Plik przepływów: ${excelFileName}`);
  infoLines.push(`Plik gmin: ${gminaFileName}`);
  infoLines.push(`Liczba gmin: ${gminas.size}`);
  infoLines.push(`Liczba unikalnych kodów po normalizacji: ${flowCodes.size}`);
  infoLines.push(`Wczytane wiersze skoroszytu: ${rawRows}`);
  infoLines.push(`Znormalizowane komórki kodów: ${normalizedCells}`);
  infoLines.push(`Wiersze wymagające normalizacji: ${normalizedRows}`);

  if (missingInExcel.length) {
    warnings.push(`${missingInExcel.length} kodów gmin występuje w granicach, ale nie ma ich w skoroszycie.`);
    warnings.push(`Przykład braków w skoroszycie: ${missingInExcel.slice(0, 10).map(([code, name]) => `${code} (${name})`).join(", ")}`);
  }
  if (missingInGmina.length) {
    warnings.push(`${missingInGmina.length} kodów przepływów występuje w skoroszycie, ale brakuje ich w granicach gmin.`);
    warnings.push(`Przykład braków w granicach: ${missingInGmina.slice(0, 10).map(([code]) => code).join(", ")}`);
  }
  if (!warnings.length) {
    warnings.push("Walidacja zakończona pomyślnie: wszystkie znormalizowane kody ze skoroszytu pasują do wczytanego katalogu gmin.");
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
    dom.validationText.textContent = "Nie wczytano jeszcze zestawu danych.";
    return;
  }
  const lines = [
    "Raport walidacji",
    "",
    ...state.dataset.validation.infoLines.map((line) => `- ${line}`),
    "",
    "Ostrzeżenia",
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
    option.textContent = dataset ? "Brak pasujących pozycji" : "Najpierw wczytaj dane";
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
    dom.powiatSummary.textContent = "Wybierz powiat, aby zarządzać zaznaczonymi gminami.";
    return;
  }

  const selectedCode = getSelectedCode();
  if (!selectedCode || !dataset.powiats.has(selectedCode)) {
    dom.powiatSummary.textContent = "Wybierz powiat, aby zarządzać zaznaczonymi gminami.";
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
  dom.powiatSummary.textContent = `Zaznaczono ${selected} / ${checkMap.size} gmin.`;
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
    dom.detailsText.textContent = "Wczytaj dane, aby zobaczyć szczegóły bieżącego wyboru.";
    return;
  }

  const selectionCode = getSelectedCode();
  if (!selectionCode) {
    dom.detailsText.textContent = "Wybierz gminę albo powiat.";
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
      `Zaznaczone gminy: ${selectedGminas.length} / ${allGminas.length}`,
      `Tryb przepływu: ${dom.flowMode.value === "from" ? "z gminy" : "do gminy"}`,
      `Filtr przepływów: ${formatOptional(filterState.transferThreshold, "> ", "bez filtra")}`,
      `Filtr dystansu: ${formatOptional(filterState.rangeKm, "<= ", "bez filtra", " km")}`,
      `Top N: ${filterState.topN ?? "bez filtra"} (stosowane globalnie podczas eksportu)`,
      `Łączna liczba pasujących wierszy przed Top N: ${totalRows}`,
      `Łączny przepływ wewnętrzny: ${totalSelfFlow}`,
      `Przykładowe zaznaczone gminy: ${selectedGminas.slice(0, 10).map((code) => `${code} - ${dataset.gminas.get(code).name}`).join(", ") || "brak"}`,
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
    `Powiat: ${powiatCode ? `${powiatCode} - ${powiatName}` : "brak"}`,
    `Tryb przepływu: ${dom.flowMode.value === "from" ? "z gminy" : "do gminy"}`,
    `Filtr przepływów: ${formatOptional(filterState.transferThreshold, "> ", "bez filtra")}`,
    `Filtr dystansu: ${formatOptional(filterState.rangeKm, "<= ", "bez filtra", " km")}`,
    `Top N: ${filterState.topN ?? "bez filtra"}`,
    `Liczba wierszy po filtrach: ${result.count}`,
    `Przepływ wewnętrzny: ${result.selfFlow}`,
    `Liczba unikalnych gmin po drugiej stronie: ${new Set(result.rows.map((row) => row.otherCode)).size}`,
    `Wykluczono przez filtr przepływów: ${result.filteredByTransfer}`,
    `Wykluczono przez filtr dystansu: ${result.filteredByDistance}`,
    `Wykluczono przez Top N: ${result.filteredByTopN}`,
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
      throw new Error("Najpierw wczytaj dane.");
    }
    const selectionCode = getSelectedCode();
    if (!selectionCode) {
      throw new Error("Przed eksportem wybierz gminę albo powiat.");
    }

    const filterState = getFilterState();
    if (filterState.error) {
      throw new Error(filterState.error);
    }

    setStatus("Przygotowywanie eksportu...");
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
        throw new Error("Zaznacz co najmniej jedną gminę w wybranym powiecie.");
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
    const baseName = `${timestamp}_${safeFilename(selectionCode || "zakres")}_${safeFilename(selectedName || contextLabel || "eksport")}`;

    const bundle = await createZip([
      { name: `${baseName}/${baseName}_lines.geojson`, data: textToBytes(JSON.stringify(lineGeoJson, null, 2)) },
      { name: `${baseName}/${baseName}_polygons.geojson`, data: textToBytes(JSON.stringify(polygonGeoJson, null, 2)) },
      { name: `${baseName}/${baseName}.txt`, data: txtBytes },
      { name: `${baseName}/${baseName}.xlsx`, data: new Uint8Array(await workbook.arrayBuffer()) },
    ]);

    downloadBlob(new Blob([bundle], { type: "application/zip" }), `${baseName}.zip`);
    setStatus("Eksport gotowy.");
  } catch (error) {
    setStatus("Eksport nie powiódł się.");
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
      kod_zrodlowy: row.sourceCode,
      nazwa_zrodlowa: row.sourceName,
      kod_docelowy: row.targetCode,
      nazwa_docelowa: row.targetName,
      odleglosc_km: round(row.distanceKm, 3),
      przeplywy: row.transfers,
      kierunek: row.direction === "from" ? "z gminy" : "do gminy",
      kod_wybranej_gminy: row.selectedCode,
      nazwa_wybranej_gminy: row.selectedName,
    }));
  }
  return { type: "FeatureCollection", features };
}

function buildPolygonGeoJson(dataset, rows, contextLabel) {
  const connected = buildConnectedGminaTable(dataset, rows, contextLabel);
  return {
    type: "FeatureCollection",
    features: connected.map((row) => makePolygonFeature(dataset.gminas.get(row.code), {
      kod_teryt: row.code,
      nazwa_gminy: row.name,
      liczba_polaczen: row.connectionCount,
      suma_przeplywow: row.totalTransfers,
      suma_wychodzacych: row.outgoingTransfers,
      suma_przychodzacych: row.incomingTransfers,
      partnerzy: row.partners.join("; "),
      zakres: contextLabel,
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
    ["Zakres", selectedCode && selectedName ? `${selectedCode} - ${selectedName}` : contextLabel],
    ["Tryb", dom.flowMode.value === "from" ? "z gminy" : "do gminy"],
    ["Filtr przepływów", filters.transferThreshold === null ? "bez filtra" : `> ${filters.transferThreshold}`],
    ["Filtr dystansu", filters.rangeKm === null ? "bez filtra" : `<= ${filters.rangeKm.toFixed(2)} km`],
    ["Top N", filters.topN === null ? "bez filtra" : String(filters.topN)],
    ["Wiersze wejściowe", String(inputRows || rows.length + selfFlow)],
    ["Wiersze po filtrach", String(rows.length)],
    ["Łączna liczba przepływów", `${totalTransfers} osób`],
    ["Średni przepływ na wiersz", avgTransfers.toFixed(2)],
    ["Mediana przepływów", median(rows.map((row) => row.transfers)).toFixed(2)],
    ["Minimalny przepływ", String(rows.length ? Math.min(...rows.map((row) => row.transfers)) : 0)],
    ["Maksymalny przepływ", String(rows.length ? Math.max(...rows.map((row) => row.transfers)) : 0)],
    ["Średnia odległość", `${mean(distances).toFixed(2)} km`],
    ["Mediana odległości", `${median(distances).toFixed(2)} km`],
    ["Średnia ważona odległość", `${weightedDistance.toFixed(2)} km`],
    ["Minimalna odległość", `${(rows.length ? Math.min(...distances) : 0).toFixed(2)} km`],
    ["Maksymalna odległość", `${(rows.length ? Math.max(...distances) : 0).toFixed(2)} km`],
    ["Unikalne gminy zakresu", String(uniqueSelected || (selectedCode ? 1 : 0))],
    ["Unikalne gminy po drugiej stronie", String(uniqueOther)],
    ["Przepływ wewnętrzny", `${selfFlow} osób`],
    ["Braki w skoroszycie", String(dataset.validation.missingInExcel.length)],
    ["Braki w granicach", String(dataset.validation.missingInGmina.length)],
  ];

  const summaryLines = [
    `Data uruchomienia: ${new Date().toISOString()}`,
    `Zakres: ${selectedCode && selectedName ? `${selectedCode} - ${selectedName}` : contextLabel}`,
    `Tryb: ${dom.flowMode.value === "from" ? "z gminy" : "do gminy"}`,
    `Filtr przepływów: ${filters.transferThreshold === null ? "bez filtra" : `> ${filters.transferThreshold}`}`,
    `Filtr dystansu: ${filters.rangeKm === null ? "bez filtra" : `<= ${filters.rangeKm.toFixed(2)} km`}`,
    `Top N: ${filters.topN === null ? "bez filtra" : filters.topN}`,
    "Uwaga: w trybie powiatu Top N jest stosowany globalnie dla wyniku łączonego.",
    "Dystans liczony jest między centroidami gmin.",
    "",
    "Podsumowanie walidacji:",
    ...dataset.validation.infoLines.map((line) => `  ${line}`),
    `  Braki w skoroszycie: ${dataset.validation.missingInExcel.length}`,
    `  Braki w granicach: ${dataset.validation.missingInGmina.length}`,
    "",
    "Statystyki zapytania:",
    ...statsItems.map(([label, value]) => `  ${label}: ${value}`),
    "",
    "Największe przepływy:",
  ];
  if (!topRows.length) {
    summaryLines.push("  Brak danych po zastosowaniu filtrów.");
  } else {
    topRows.forEach((row, index) => {
      summaryLines.push(`  ${index + 1}. ${row.sourceCode} - ${row.sourceName} -> ${row.targetCode} - ${row.targetName}: ${row.transfers} osób, ${row.distanceKm.toFixed(2)} km`);
    });
  }
  summaryLines.push("");
  summaryLines.push("Ostrzeżenia:");
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
      name: "Przepływy",
      rows: [
        ["Kod źródłowy", "Nazwa źródłowa", "Kod docelowy", "Nazwa docelowa", "Odległość km", "Przepływy", "Kierunek", "Kod wybranej gminy", "Nazwa wybranej gminy"],
        ...rows.map((row) => [
          row.sourceCode,
          row.sourceName,
          row.targetCode,
          row.targetName,
          round(row.distanceKm, 3),
          row.transfers,
          row.direction === "from" ? "z gminy" : "do gminy",
          row.selectedCode,
          row.selectedName,
        ]),
      ],
    },
    {
      name: "Podsumowanie",
      rows: report.summaryLines.map((line) => [line]),
    },
    {
      name: "Statystyki zapytania",
      rows: [
        ["Metryka", "Wartość"],
        ...report.statsItems,
      ],
    },
    {
      name: "Największe przepływy",
      rows: [
        ["Lp", "Kod źródłowy", "Nazwa źródłowa", "Kod docelowy", "Nazwa docelowa", "Przepływy", "Odległość km"],
        ...report.topRows.map((row, index) => [index + 1, row.sourceCode, row.sourceName, row.targetCode, row.targetName, row.transfers, round(row.distanceKm, 3)]),
      ],
    },
    {
      name: "Statystyki zakresu",
      rows: [
        ["Lp", "Kod wybranej gminy", "Nazwa wybranej gminy", "Liczba wierszy", "Suma przepływów", "Średni przepływ", "Mediana przepływów", "Suma odległości km", "Średnia odległość km", "Liczba unikalnych partnerów", "Przepływ wewnętrzny"],
        ...report.scopeStats.map((row) => [row.lp, row.selectedCode, row.selectedName, row.rowCount, row.totalTransfers, round(row.avgTransfers, 2), round(row.medianTransfers, 2), round(row.totalDistanceKm, 3), round(row.avgDistanceKm, 3), row.uniquePartners, row.selfFlow]),
      ],
    },
    {
      name: "Metryki sieci",
      rows: [
        ["Lp", "Kod wybranej gminy", "Nazwa wybranej gminy", "Liczba wierszy", "Suma przepływów", "Średni przepływ", "Mediana przepływów", "Suma odległości km", "Średnia odległość km", "Liczba unikalnych partnerów", "Kod najsilniejszego partnera", "Nazwa najsilniejszego partnera", "Przepływ najsilniejszego partnera", "Odległość do najsilniejszego partnera km", "Stopień centralności", "Centralność przepływu", "Udział najsilniejszego partnera", "Średni przepływ na partnera", "Przepływ wewnętrzny"],
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
  const transferThreshold = parseOptionalInt(dom.transferThreshold.value, "filtr przepływów");
  if (transferThreshold.error) {
    return transferThreshold;
  }
  const rangeKm = parseOptionalFloat(dom.rangeKm.value, "maks. dystans");
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
    return { value: null, error: `Pole ${label} musi być liczbą całkowitą.` };
  }
  if (parsed < 0) {
    return { value: null, error: `Pole ${label} nie może być ujemne.` };
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
    return { value: null, error: `Pole ${label} musi być liczbą.` };
  }
  if (parsed < 0) {
    return { value: null, error: `Pole ${label} nie może być ujemne.` };
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
