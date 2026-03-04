// Columnas a mostrar (A,B,C,D,E,F,J,K,L,M,N,S)
const PICK = [0, 1, 2, 3, 4, 5, 9, 10, 11, 12, 13, 18];

const HEADERS = [
  "CUMPLIDO EL", "ESTADO DE CUMPLIDO", "PAGADO EL", "FECHAS VIAJES",
  "PLACA", "CONDUCTOR",
  "VALOR TOTAL DEL FLETE", "ANTICIPO", "ICA", "INTER2,4%", "RET",
  "VALOR TOTAL A PAGAR"
];

function maskDate(v) {
  const s = (v ?? "").toString().trim();
  if (!s) return "";
  // intenta yyyy-mm-dd
  const m1 = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m1) return `${m1[3]}-${m1[2]}-****`;
  // intenta dd/mm/yyyy o dd-mm-yyyy
  const m2 = s.match(/(\d{2})[\/\-](\d{2})[\/\-](\d{4})/);
  if (m2) return `${m2[1]}-${m2[2]}-****`;
  return s; // si no parece fecha, lo deja
}

function maskFirst(v) {
  const s = (v ?? "").toString().trim();
  if (!s) return "";
  return s[0].toUpperCase() + "*****";
}

function transformPickedRow(row) {
  const out = PICK.map(i => row?.[i] ?? "");
  // A,C,D -> mask fecha (A=0, C=2, D=3 dentro del "out" quedan en posiciones 0,2,3)
  out[0] = maskDate(out[0]);
  out[2] = maskDate(out[2]);
  out[3] = maskDate(out[3]);
  // E,F -> mask (en "out" quedan en posiciones 4,5)
  out[4] = maskFirst(out[4]);
  out[5] = maskFirst(out[5]);
  return out;
}

// rows = arreglo de filas completas (todas las columnas)
const filteredRows = (rows || []).map(transformPickedRow);

return {
  status: 200,
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify({
    ok: true,
    headers: HEADERS,
    rows: filteredRows
  })
};
