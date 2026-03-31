const axios = require("axios");

// Columnas a mostrar (A,B,C,D,E,F,J,K,L,M,N,S)
const PICK = [0, 1, 2, 3, 4, 5, 9, 10, 11, 12, 13, 18];

const HEADERS = [
  "CUMPLIDO EL",
  "ESTADO DE CUMPLIDO",
  "PAGADO EL",
  "FECHAS VIAJES",
  "PLACA",
  "CONDUCTOR",
  "VALOR TOTAL DEL FLETE",
  "ANTICIPO",
  "ICA",
  "INTER2,4%",
  "RET",
  "VALOR TOTAL A PAGAR"
];

function normalizePlate(s) {
  return (s || "").toString().toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function maskDate(v) {
  if (v === null || v === undefined || v === "") return "";

  if (typeof v === "number" && isFinite(v)) {
    const base = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(base.getTime() + v * 86400000);
    return `${pad2(d.getUTCDate())}-${pad2(d.getUTCMonth() + 1)}-****`;
  }

  const s = String(v).trim();

  const m1 = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m1) return `${m1[3]}-${m1[2]}-****`;

  const m2 = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m2) return `${pad2(m2[1])}-${pad2(m2[2])}-****`;

  return s;
}

function maskFirst(v) {
  const s = (v ?? "").toString().trim();
  if (!s) return "";
  return s[0].toUpperCase() + "*****";
}

function transformPickedRow(row) {
  const out = PICK.map(i => row?.[i] ?? "");

  // A,C,D -> fecha
  out[0] = maskDate(out[0]);
  out[2] = maskDate(out[2]);
  out[3] = maskDate(out[3]);

  // E,F -> máscara
  out[4] = maskFirst(out[4]);
  out[5] = maskFirst(out[5]);

  return out;
}

function parseDateValue(v) {
  if (v === null || v === undefined || v === "") return null;

  if (typeof v === "number" && isFinite(v)) {
    const base = new Date(Date.UTC(1899, 11, 30));
    return new Date(base.getTime() + v * 86400000);
  }

  const s = String(v).trim();

  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));

  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));

  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

module.exports = async function (context, req) {
  try {
    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;
    const driveId = process.env.DRIVE_ID;
    const itemId = process.env.ITEM_ID;
    const tableName = process.env.TABLE_NAME;

    if (!tenantId || !clientId || !clientSecret || !driveId || !itemId || !tableName) {
      context.res = {
        status: 500,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ok: false,
          error: "Faltan variables de entorno requeridas"
        })
      };
      return;
    }

    const { plate, startDate, endDate } = req.body || {};
    const normalizedPlate = normalizePlate(plate);

    if (!normalizedPlate) {
      context.res = {
        status: 400,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ok: false,
          error: "La placa es obligatoria"
        })
      };
      return;
    }

    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const tokenParams = new URLSearchParams();
    tokenParams.append("client_id", clientId);
    tokenParams.append("scope", "https://graph.microsoft.com/.default");
    tokenParams.append("client_secret", clientSecret);
    tokenParams.append("grant_type", "client_credentials");

    const tokenResp = await axios.post(tokenUrl, tokenParams.toString(), {
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      timeout: 30000
    });

    const accessToken = tokenResp.data.access_token;
    if (!accessToken) {
      throw new Error("No se pudo obtener access token");
    }

    const graphUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables('${encodeURIComponent(tableName)}')/rows?$top=5000`;

    const graphResp = await axios.get(graphUrl, {
      headers: {
        Authorization: `Bearer ${accessToken}`
      },
      timeout: 30000
    });

    const rawRows = (graphResp.data.value || [])
      .map(r => (Array.isArray(r.values) && Array.isArray(r.values[0]) ? r.values[0] : null))
      .filter(Boolean);

    const start = startDate ? new Date(startDate + "T00:00:00Z") : null;
    const end = endDate ? new Date(endDate + "T23:59:59Z") : null;

    const filtered = rawRows.filter(row => {
      const plateCell = normalizePlate(row[4]); // E = PLACA
      if (plateCell !== normalizedPlate) return false;

      if (!start && !end) return true;

      const tripDate = parseDateValue(row[3]); // D = FECHAS VIAJES
      if (!tripDate) return false;

      if (start && tripDate < start) return false;
      if (end && tripDate > end) return false;

      return true;
    });

    // Más reciente primero por FECHAS VIAJES (D = índice 3)
    filtered.sort((a, b) => {
      const da = parseDateValue(a[3]);
      const db = parseDateValue(b[3]);

      if (!da && !db) return 0;
      if (!da) return 1;
      if (!db) return -1;

      return db - da;
    });

    const filteredRows = filtered.map(transformPickedRow);

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        ok: true,
        headers: HEADERS,
        rows: filteredRows
      })
    };
  } catch (error) {
    context.log("ERROR API /buscar:", error?.response?.data || error?.message || error);

    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        ok: false,
        error: error?.response?.data || error?.message || "Error interno"
      })
    };
  }
};
