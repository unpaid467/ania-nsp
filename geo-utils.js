export function parseGeoJson(text) {
  const data = JSON.parse(text);
  if (data.type === "FeatureCollection" && Array.isArray(data.features)) {
    return data;
  }
  throw new Error("Expected a GeoJSON FeatureCollection.");
}

export function buildGminaCatalog(featureCollection) {
  const gminas = new Map();
  for (const feature of featureCollection.features) {
    const props = feature.properties || {};
    const code = cleanText(props.JPT_KOD_JE || props.kod_teryt || props.code);
    const name = cleanText(props.JPT_NAZWA_ || props.nazwa_gminy || props.name);
    if (!code || !name || !feature.geometry) {
      continue;
    }
    const centroid = geometryCentroid(feature.geometry);
    const bounds = geometryBounds(feature.geometry);
    gminas.set(code, {
      code,
      name,
      centroidX: centroid[0],
      centroidY: centroid[1],
      layoutX: centroid[0],
      layoutY: centroid[1],
      bounds,
      geometry: feature.geometry,
      feature,
    });
  }
  return gminas;
}

export function buildPowiatCatalog(featureCollection) {
  const powiats = new Map();
  for (const feature of featureCollection.features) {
    const props = feature.properties || {};
    const code = cleanText(props.JPT_KOD_JE || props.kod_teryt || props.code);
    const name = cleanText(props.JPT_NAZWA_ || props.nazwa_powiatu || props.name);
    if (!code || !name) {
      continue;
    }
    powiats.set(code, { code, name, feature });
  }
  return powiats;
}

export function geometryBounds(geometry) {
  const coords = flattenCoordinates(geometry);
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const [x, y] of coords) {
    minX = Math.min(minX, x);
    minY = Math.min(minY, y);
    maxX = Math.max(maxX, x);
    maxY = Math.max(maxY, y);
  }
  if (!Number.isFinite(minX)) {
    return [0, 0, 0, 0];
  }
  return [minX, minY, maxX, maxY];
}

export function geometryCentroid(geometry) {
  if (!geometry) {
    return [0, 0];
  }
  if (geometry.type === "Point") {
    return geometry.coordinates.slice(0, 2);
  }
  if (geometry.type === "Polygon") {
    return polygonCentroid(geometry.coordinates);
  }
  if (geometry.type === "MultiPolygon") {
    return multiPolygonCentroid(geometry.coordinates);
  }
  const coords = flattenCoordinates(geometry);
  if (!coords.length) {
    return [0, 0];
  }
  const sum = coords.reduce((acc, [x, y]) => [acc[0] + x, acc[1] + y], [0, 0]);
  return [sum[0] / coords.length, sum[1] / coords.length];
}

export function haversineKm(lon1, lat1, lon2, lat2) {
  const radiusKm = 6371.0088;
  const phi1 = toRadians(lat1);
  const phi2 = toRadians(lat2);
  const dPhi = toRadians(lat2 - lat1);
  const dLambda = toRadians(lon2 - lon1);
  const a = (Math.sin(dPhi / 2) ** 2) +
    (Math.cos(phi1) * Math.cos(phi2) * Math.sin(dLambda / 2) ** 2);
  return 2 * radiusKm * Math.atan2(Math.sqrt(a), Math.sqrt(Math.max(0, 1 - a)));
}

export function makeLineFeature(source, target, properties) {
  return {
    type: "Feature",
    properties,
    geometry: {
      type: "LineString",
      coordinates: [
        [source.layoutX, source.layoutY],
        [target.layoutX, target.layoutY],
      ],
    },
  };
}

export function makePolygonFeature(record, properties) {
  return {
    type: "Feature",
    properties,
    geometry: record.geometry,
  };
}

export function cleanText(value) {
  if (value === null || value === undefined) {
    return "";
  }
  const text = String(value).trim();
  return text.toLowerCase() === "nan" ? "" : text;
}

function flattenCoordinates(geometry) {
  const out = [];
  collectCoords(geometry.coordinates, out);
  return out;
}

function collectCoords(node, out) {
  if (!Array.isArray(node)) {
    return;
  }
  if (typeof node[0] === "number" && typeof node[1] === "number") {
    out.push([node[0], node[1]]);
    return;
  }
  for (const child of node) {
    collectCoords(child, out);
  }
}

function polygonCentroid(rings) {
  if (!rings || !rings.length) {
    return [0, 0];
  }
  const outer = rings[0];
  return ringCentroid(outer);
}

function multiPolygonCentroid(polygons) {
  let totalArea = 0;
  let cx = 0;
  let cy = 0;
  for (const polygon of polygons) {
    const outer = polygon[0];
    const { area, centroid } = ringAreaAndCentroid(outer);
    const weight = Math.abs(area);
    totalArea += weight;
    cx += centroid[0] * weight;
    cy += centroid[1] * weight;
  }
  if (totalArea === 0) {
    return polygonCentroid(polygons[0] || []);
  }
  return [cx / totalArea, cy / totalArea];
}

function ringCentroid(points) {
  const { area, centroid } = ringAreaAndCentroid(points);
  if (Math.abs(area) < 1e-12) {
    const sum = points.reduce((acc, [x, y]) => [acc[0] + x, acc[1] + y], [0, 0]);
    return [sum[0] / points.length, sum[1] / points.length];
  }
  return centroid;
}

function ringAreaAndCentroid(points) {
  if (!points || points.length < 3) {
    return { area: 0, centroid: [0, 0] };
  }
  let twiceArea = 0;
  let cx = 0;
  let cy = 0;
  for (let i = 0; i < points.length - 1; i += 1) {
    const [x1, y1] = points[i];
    const [x2, y2] = points[i + 1];
    const cross = (x1 * y2) - (x2 * y1);
    twiceArea += cross;
    cx += (x1 + x2) * cross;
    cy += (y1 + y2) * cross;
  }
  if (Math.abs(twiceArea) < 1e-12) {
    return { area: 0, centroid: [points[0][0], points[0][1]] };
  }
  return {
    area: twiceArea / 2,
    centroid: [cx / (3 * twiceArea), cy / (3 * twiceArea)],
  };
}

function toRadians(value) {
  return (value * Math.PI) / 180;
}
