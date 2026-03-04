// api/buscar/index.js
const axios = require("axios");

/**
 * Columnas requeridas (por letra Excel): A,B,C,D,E,F,J,K,L,M,N,S
 * Según tu encabezado:
 * A  CUMPLIDO EL
 * B  ESTADO DE CUMPLIDO
 * C  PAGADO EL
 * D  FECHAS VIAJES
 * E  PLACA
 * F  CONDUCTOR
 * J  VALOR TOTAL DEL FLETE
 * K  ANTICIPO
 * L  ICA
 * M  INTER2,4%
 * N  RET
 * S  VALOR TOTAL A PAGAR
 */
const OUT_COLUMNS = [
  { idx: 0, key: "CUMPLIDO_EL" },           // A
  { idx: 1, key: "ESTADO_CUMPLIDO" },       // B
  { idx: 2, key: "PAGADO_EL" },             // C
  { idx: 3, key: "FECHAS_VIAJES" },         // D
  { idx: 4, key: "PLACA" },                 // E
  { idx: 5, key: "CONDUCTOR" },             // F
  { idx: 9, key: "VALOR_TOTAL_FLETE" },     // J
  { idx: 10, key: "ANTICIPO" },             // K
  { idx: 11, key: "ICA" },                  // L
  { idx: 12, key: "INTER_2_4" },            // M
  { idx: 13, key: "RET" },                  // N
  { idx: 18, key: "VALOR_TOTAL_A_PAGAR" },  // S
];

// Encabezados “bonitos” para la tabla del front
const OUT_HEADERS = [
  "Cumplido el",
  "Estado de cumplido",
  "Pagado el",
  "Fechas viajes",
  "Placa",
  "Conductor",
  "Valor total del flete",
  "Anticipo",
  "ICA",
  "Inter 2,4%",
  "RET",
  "Valor total a pagar",
];

function needEnv(name) {
  const v = process.env[name];
  if (!v) throw new Error(`Falta variable de entorno: ${name}`);
  return v;
}

function normalizePlate(s) {
  return (s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function maskFirstLetter(s) {
  const t = String(s ?? "").trim();
  if (!t) return "";
  return t[0] + "*****";
}

function formatDateMasked(value) {
  // Acepta: Date serial, ISO, "dd/mm/aaaa", etc.
  if (value === null || value === undefined || value === "") return "";

  // Si viene como número (Excel serial)
  if (typeof value === "number" && isFinite(value)) {
    // Excel serial (aprox, suficiente para visual)
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(excelEpoch.getTime() + value * 86400000);
    const dd = String(d.getUTCDate()).padStart(2, "0");
    const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
    return `${dd}-${mm}-****`;
  }

  const s = String(value).trim();

  // Intento dd/mm/yyyy o dd-mm-yyyy
  const m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
  if (m) {
    const dd = String(m[1]).padStart(2, "0");
    const mm = String(m[2]).padStart(2, "0");
    return `${dd}-${mm}-****`;
  }

  // Intento ISO
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    return `${dd}-${mm}-****`;
  }

  // Si no se puede parsear, igual enmascara “año”
  return s.replace(/\d{4}/g, "****");
}

function isBetween(dateStr, start, end) {
  if (!dateStr) return true;

  // Intento parsear fecha “real” (sin máscara) para filtrar.
  // Si no se puede, no filtramos por fecha.
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return true;

  if (start) {
    const s = new Date(start);
    if (!isNaN(s.getTime()) && d < s) return false;
  }
  if (end) {
    const e = new Date(end);
    if (!isNaN(e.getTime()) && d > e) return false;
  }
  return true;
}

async function getAccessToken() {
  const tenantId = needEnv("TENANT_ID");
  const clientId = needEnv("CLIENT_ID");
  const clientSecret = needEnv("CLIENT_SECRET");

  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const body = new URLSearchParams();
  body.set("client_id", clientId);
  body.set("client_secret", clientSecret);
  body.set("grant_type", "client_credentials");
  body.set("scope", "https://graph.microsoft.com/.default");

  const resp = await axios.post(url, body.toString(), {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    timeout: 20000,
  });

  return resp.data.access_token;
}

async function getTableRows(accessToken) {
  const driveId = needEnv("DRIVE_ID");
  const itemId = needEnv("ITEM_ID");
  const tableName = needEnv("TABLE_NAME");

  // Lee filas del Excel Table (Graph)
  // Nota: si tienes MUCHAS filas, se puede paginar; por ahora traemos “las que haya”
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables/${encodeURIComponent(
    tableName
  )}/rows?$top=1000`;

  const resp = await axios.get(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
    timeout: 30000,
  });

  // resp.data.value[i].values = [ [col0, col1, ...] ]
  const rows = [];
  for (const r of resp.data.value || []) {
    if (Array.isArray(r.values) && r.values[0]) rows.push(r.values[0]);
  }
  return rows;
}

module.exports = async function (context, req) {
  const t0 = Date.now();

  try {
    const body = req.body || {};
    const plate = normalizePlate(body.plate);
    const startDate = body.startDate || null;
    const endDate = body.endDate || null;

    if (!plate) {
      context.res = {
        status: 400,
        headers: { "Content-Type": "application/json" },
        body: { error: "Falta plate" },
      };
      return;
    }

    const token = await getAccessToken();
    const allRows = await getTableRows(token);

    // Filtrado (por PLACA = columna E -> idx 4)
    const filtered = allRows.filter((r) => {
      const plateCell = normalizePlate(r?.[4]);
      if (plateCell !== plate) return false;

      // Si quieres filtrar por fechas, lo hacemos con COLUMNA D (Fechas viajes idx 3)
      const fechasViajes = r?.[3] ? String(r[3]) : "";
      return isBetween(fechasViajes, startDate, endDate);
    });

    // Transformación: solo columnas deseadas + máscaras
    const outRows = filtered.map((r) => {
      const obj = {};
      for (const c of OUT_COLUMNS) obj[c.key] = r?.[c.idx] ?? "";

      // Enmascarado fechas: A,C,D (idx 0,2,3)
      obj.CUMPLIDO_EL = formatDateMasked(obj.CUMPLIDO_EL);
      obj.PAGADO_EL = formatDateMasked(obj.PAGADO_EL);
      obj.FECHAS_VIAJES = formatDateMasked(obj.FECHAS_VIAJES);

      // Enmascarado E y F: primera letra + *****
      obj.PLACA = maskFirstLetter(obj.PLACA);
      obj.CONDUCTOR = maskFirstLetter(obj.CONDUCTOR);

      return obj;
    });

    const ms = Date.now() - t0;

    context.res = {
      status: 200,
      headers: {
        "Content-Type": "application/json",
        // 👉 IMPORTANTE para tu front: “sí hay headers”
        "X-Has-Headers": "1",
      },
      body: {
        ok: true,
        ms,
        headers: OUT_HEADERS,
        rows: outRows,
        count: outRows.length,
      },
    };
  } catch (err) {
    const ms = Date.now() - t0;
    const message = err?.response?.data
      ? JSON.stringify(err.response.data)
      : String(err?.message || err);

    context.log("ERROR buscar:", message);

    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: false,
        ms,
        error: message,
      },
    };
  }
};
