const axios = require("axios");

function json(res, status, body) {
  return {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8" },
    body: JSON.stringify(body),
  };
}

function text(res, status, body) {
  return {
    status,
    headers: { "Content-Type": "text/plain; charset=utf-8" },
    body: String(body),
  };
}

function normalizePlate(s) {
  return (s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function yyyyMmDdToExcelSerial(dateStr) {
  // Convierte "YYYY-MM-DD" a serial de Excel (días desde 1899-12-30)
  // Usamos UTC para evitar desfases.
  const m = String(dateStr || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;

  const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  const dt = new Date(Date.UTC(y, mo - 1, d));
  if (Number.isNaN(dt.getTime())) return null;

  const base = new Date(Date.UTC(1899, 11, 30));
  const diffDays = Math.floor((dt.getTime() - base.getTime()) / 86400000);
  return diffDays;
}

async function getToken() {
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Faltan TENANT_ID/CLIENT_ID/CLIENT_SECRET en variables de entorno.");
  }

  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const form = new URLSearchParams();
  form.append("client_id", clientId);
  form.append("client_secret", clientSecret);
  form.append("grant_type", "client_credentials");
  form.append("scope", "https://graph.microsoft.com/.default");

  const resp = await axios.post(url, form.toString(), {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    timeout: 20000,
  });

  return resp.data.access_token;
}

async function fetchTableValues(token) {
  const driveId = process.env.DRIVE_ID;
  const itemId = process.env.ITEM_ID;
  const tableName = process.env.TABLE_NAME;

  if (!driveId || !itemId || !tableName) {
    throw new Error("Faltan DRIVE_ID/ITEM_ID/TABLE_NAME en variables de entorno.");
  }

  // Obtiene valores de la tabla de Excel (incluye encabezados si existen en la tabla)
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables/${encodeURIComponent(
    tableName
  )}/range`;

  const resp = await axios.get(url, {
    headers: { Authorization: `Bearer ${token}` },
    timeout: 20000,
  });

  const values = resp.data?.values;
  if (!Array.isArray(values) || values.length === 0) {
    return { headers: [], rows: [] };
  }

  // La primera fila del "range" normalmente es el header de la tabla
  const headers = values[0];
  const rows = values.slice(1);

  return { headers, rows };
}

function buildHeaderIndex(headers) {
  const map = new Map();
  headers.forEach((h, i) => {
    const key = String(h || "").trim().replace(/\s+/g, " ").toUpperCase();
    map.set(key, i);
  });
  return map;
}

function cellToNumber(v) {
  if (typeof v === "number") return v;
  const s = String(v || "").replace(/\./g, "").replace(/,/g, ".").trim();
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function rowMatches(row, idx, plate, startSerial, endSerial) {
  // Filtra por PLACA (columna E)
  const iPlaca = idx.get("PLACA");
  if (iPlaca === undefined) return false;

  const placaRow = normalizePlate(row[iPlaca]);
  if (placaRow !== plate) return false;

  // Si no hay fechas, listo
  if (startSerial === null && endSerial === null) return true;

  // FECHAS VIAJES (columna D) puede venir como serial Excel o texto
  const iFechaViajes = idx.get("FECHAS VIAJES");
  if (iFechaViajes === undefined) return true; // si no existe, no bloqueamos

  const raw = row[iFechaViajes];

  let serial = null;
  if (typeof raw === "number") serial = raw;
  else {
    // intenta parsear texto tipo YYYY-MM-DD o DD/MM/YYYY etc. -> lo convertimos a serial aproximado
    const s = String(raw || "").trim();
    // ISO
    const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) serial = yyyyMmDdToExcelSerial(`${iso[1]}-${iso[2]}-${iso[3]}`);
    // DD/MM/YYYY
    const dmy = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
    if (!serial && dmy) serial = yyyyMmDdToExcelSerial(`${dmy[3]}-${String(dmy[2]).padStart(2,"0")}-${String(dmy[1]).padStart(2,"0")}`);
  }

  // Si no logramos serial, no filtramos por fecha (para no perder datos)
  if (serial === null) return true;

  if (startSerial !== null && serial < startSerial) return false;
  if (endSerial !== null && serial > endSerial) return false;

  return true;
}

module.exports = async function (context, req) {
  try {
    const { plate, startDate, endDate } = req.body || {};
    const p = normalizePlate(plate);

    if (!p) {
      return json(context.res, 400, { error: "PLATE_REQUIRED" });
    }

    const t0 = Date.now();
    const token = await getToken();
    const { headers, rows } = await fetchTableValues(token);

    // Si no hay headers reales, devolvemos error claro
    if (!headers || headers.length === 0) {
      return json(context.res, 500, { error: "NO_HEADERS_IN_TABLE", message: "La tabla de Excel no tiene encabezados." });
    }

    const idx = buildHeaderIndex(headers);

    const startSerial = startDate ? yyyyMmDdToExcelSerial(startDate) : null;
    const endSerial = endDate ? yyyyMmDdToExcelSerial(endDate) : null;

    const filtered = rows.filter((r) => Array.isArray(r) && rowMatches(r, idx, p, startSerial, endSerial));

    const ms = Date.now() - t0;

    // 👇 RESPUESTA FINAL: SIEMPRE incluye headers + rows
    return json(context.res, 200, {
      ok: true,
      ms,
      headers,
      rows: filtered,
    });
  } catch (err) {
    // Log interno (se ve en GitHub Actions)
    context.log("API ERROR:", err?.message || err);
    return json(context.res, 500, {
      ok: false,
      error: "API_ERROR",
      message: err?.message || String(err),
    });
  }
};
