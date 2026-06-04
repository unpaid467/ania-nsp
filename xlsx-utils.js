import { createZip, textToBytes, unzipEntries } from "./zip-utils.js";

const XML = "application/xml";

export async function readWorkbookRows(file) {
  const lower = file.name.toLowerCase();
  if (lower.endsWith(".csv") || lower.endsWith(".txt")) {
    const text = await file.text();
    return parseDelimitedText(text);
  }

  const entries = await unzipEntries(file);
  const workbookXml = textEntry(entries, "xl/workbook.xml");
  const workbookRelsXml = textEntry(entries, "xl/_rels/workbook.xml.rels");
  const sharedStringsXml = entries.get("xl/sharedStrings.xml");
  const workbookDoc = parseXml(workbookXml);
  const relDoc = parseXml(workbookRelsXml);
  const sheetByName = new Map();

  const relMap = new Map();
  for (const rel of Array.from(relDoc.getElementsByTagName("Relationship"))) {
    relMap.set(rel.getAttribute("Id"), rel.getAttribute("Target"));
  }

  for (const sheet of Array.from(workbookDoc.getElementsByTagName("sheet"))) {
    const name = sheet.getAttribute("name") || "";
    const relId = sheet.getAttribute("r:id");
    const target = relMap.get(relId);
    if (target) {
      sheetByName.set(name, `xl/${target.replace(/^\.\//, "")}`);
    }
  }

  const targetSheetName = pickSheetName(Array.from(sheetByName.keys()));
  const targetPath = sheetByName.get(targetSheetName);
  if (!targetPath) {
    throw new Error("Unable to locate worksheet in workbook.");
  }

  const sharedStrings = sharedStringsXml ? parseSharedStrings(textEntry(entries, "xl/sharedStrings.xml")) : [];
  return parseWorksheetRows(textEntry(entries, targetPath), sharedStrings);
}

export async function buildWorkbookBlob(sheets) {
  const now = new Date().toISOString();
  const contentTypes = buildContentTypes(sheets.length);
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
  const core = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${escapeXml(now)}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${escapeXml(now)}</dcterms:modified>
</cp:coreProperties>`;
  const app = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
</Properties>`;
  const workbook = buildWorkbookXml(sheets);
  const workbookRels = buildWorkbookRelsXml(sheets.length);
  const styles = buildStylesXml();
  const files = [
    { name: "[Content_Types].xml", data: contentTypes },
    { name: "_rels/.rels", data: rels },
    { name: "docProps/core.xml", data: core },
    { name: "docProps/app.xml", data: app },
    { name: "xl/workbook.xml", data: workbook },
    { name: "xl/_rels/workbook.xml.rels", data: workbookRels },
    { name: "xl/styles.xml", data: styles },
  ];

  sheets.forEach((sheet, index) => {
    files.push({
      name: `xl/worksheets/sheet${index + 1}.xml`,
      data: buildWorksheetXml(sheet.rows),
    });
  });

  const zip = await createZip(files.map((file) => ({ ...file, data: textToBytes(file.data) })));
  return new Blob([zip], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function parseDelimitedText(text) {
  const lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n").filter(Boolean);
  if (!lines.length) {
    return [];
  }
  const delimiter = guessDelimiter(lines[0]);
  return lines.map((line) => splitDelimitedLine(line, delimiter));
}

function splitDelimitedLine(line, delimiter) {
  const out = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === "\"") {
      if (inQuotes && line[i + 1] === "\"") {
        current += "\"";
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === delimiter && !inQuotes) {
      out.push(current);
      current = "";
    } else {
      current += ch;
    }
  }
  out.push(current);
  return out;
}

function guessDelimiter(line) {
  const counts = [
    [",", (line.match(/,/g) || []).length],
    [";", (line.match(/;/g) || []).length],
    ["\t", (line.match(/\t/g) || []).length],
  ];
  counts.sort((a, b) => b[1] - a[1]);
  return counts[0][0];
}

function pickSheetName(names) {
  const normalized = names.map((name) => [name, normalizeKey(name)]);
  const preferred = normalized.find(([, key]) => key === "macierzprzeplywow");
  return preferred ? preferred[0] : names[0];
}

function parseWorksheetRows(xmlText, sharedStrings) {
  const doc = parseXml(xmlText);
  const rows = [];
  for (const rowNode of Array.from(doc.getElementsByTagName("row"))) {
    const values = [];
    let currentColumn = 0;
    for (const cell of Array.from(rowNode.getElementsByTagName("c"))) {
      const ref = cell.getAttribute("r") || "";
      const columnIndex = ref ? columnLettersToIndex(ref.replace(/[0-9]/g, "")) : currentColumn;
      while (values.length < columnIndex) {
        values.push("");
      }
      const type = cell.getAttribute("t") || "";
      let value = "";
      const valueNode = cell.getElementsByTagName("v")[0];
      const inlineNode = cell.getElementsByTagName("t")[0];
      if (type === "s" && valueNode) {
        value = sharedStrings[Number(valueNode.textContent || "0")] || "";
      } else if (type === "inlineStr" && inlineNode) {
        value = inlineNode.textContent || "";
      } else if (valueNode) {
        value = valueNode.textContent || "";
      }
      values[columnIndex] = value;
      currentColumn = columnIndex + 1;
    }
    rows.push(values);
  }
  return rows;
}

function parseSharedStrings(xmlText) {
  const doc = parseXml(xmlText);
  return Array.from(doc.getElementsByTagName("si")).map((si) =>
    Array.from(si.getElementsByTagName("t")).map((node) => node.textContent || "").join("")
  );
}

function buildContentTypes(sheetCount) {
  const sheetOverrides = Array.from({ length: sheetCount }, (_, index) =>
    `  <Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
  ).join("\n");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
${sheetOverrides}
</Types>`;
}

function buildWorkbookXml(sheets) {
  const sheetXml = sheets.map((sheet, index) =>
    `    <sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`
  ).join("\n");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
${sheetXml}
  </sheets>
</workbook>`;
}

function buildWorkbookRelsXml(sheetCount) {
  const rels = sheetsRange(sheetCount).map((index) =>
    `  <Relationship Id="rId${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index}.xml"/>`
  ).join("\n");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels}
  <Relationship Id="rId${sheetCount + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
}

function buildStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>`;
}

function buildWorksheetXml(rows) {
  const sheetRows = rows.map((row, rowIndex) => {
    const cells = row.map((value, colIndex) => {
      const ref = `${indexToColumnLetters(colIndex)}${rowIndex + 1}`;
      if (value === null || value === undefined || value === "") {
        return `<c r="${ref}" t="inlineStr"><is><t></t></is></c>`;
      }
      return `<c r="${ref}" t="inlineStr"><is><t>${escapeXml(String(value))}</t></is></c>`;
    }).join("");
    return `    <row r="${rowIndex + 1}">${cells}</row>`;
  }).join("\n");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
${sheetRows}
  </sheetData>
</worksheet>`;
}

function textEntry(entries, path) {
  const bytes = entries.get(path);
  if (!bytes) {
    throw new Error(`Missing ZIP entry: ${path}`);
  }
  return new TextDecoder("utf-8").decode(bytes);
}

function parseXml(text) {
  const doc = new DOMParser().parseFromString(text, XML);
  const errorNode = doc.querySelector("parsererror");
  if (errorNode) {
    throw new Error("Invalid XML while reading workbook.");
  }
  return doc;
}

function normalizeKey(value) {
  return value
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "")
    .toLowerCase();
}

function columnLettersToIndex(letters) {
  let out = 0;
  for (let i = 0; i < letters.length; i += 1) {
    out = (out * 26) + (letters.charCodeAt(i) - 64);
  }
  return Math.max(0, out - 1);
}

function indexToColumnLetters(index) {
  let n = index + 1;
  let out = "";
  while (n > 0) {
    const mod = (n - 1) % 26;
    out = String.fromCharCode(65 + mod) + out;
    n = Math.floor((n - mod) / 26);
  }
  return out;
}

function escapeXml(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function sheetsRange(count) {
  return Array.from({ length: count }, (_, index) => index + 1);
}
