/**
 * Netlify Serverless Function — NOPI Conta Corrente API
 * GET /.netlify/functions/cc-data
 */

const fetch = require("node-fetch");
const XLSX  = require("xlsx");

const EXCEL_URL = [
  "https://www.dropbox.com/scl/fi/y4i9m6v4q8snd2m3qljoh/Motherboard-2026.xlsx",
  "?rlkey=4px2hpxbg8p6fot2l65bkdamg&st=4h2vu72e&dl=1",
].join("");

const CORS = {
  "Access-Control-Allow-Origin":  "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
  "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
};

function toNum(v) {
  if (v == null || v === "" || String(v) === "nan") return null;
  const n = parseFloat(String(v).replace(",", "."));
  return Number.isFinite(n) ? Math.round(n * 100) / 100 : null;
}

function toStr(v) {
  if (v == null || v === "" || String(v).trim() === "nan") return null;
  return String(v).trim();
}

function json(statusCode, body) {
  return {
    statusCode,
    headers: { ...CORS, "Content-Type": "application/json; charset=utf-8" },
    body: JSON.stringify(body),
  };
}

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 204, headers: CORS, body: "" };
  }

  try {
    const res = await fetch(EXCEL_URL, { timeout: 45_000 });
    if (!res.ok) throw new Error(`Dropbox HTTP ${res.status}`);
    const buf = await res.buffer();
    const wb = XLSX.read(buf, { type: "buffer", cellDates: true });

    const ws = wb.Sheets["CC"];
    if (!ws) throw new Error('Folha "CC" não encontrada.');
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

    const contaCorrente = [];

    for (let i = 5; i < rows.length; i++) {
      const nome = toStr(rows[i][0]);
      const valor = toNum(rows[i][3]);

      if (!nome) continue;
      if (valor === null) continue;
      // Filtrar igual à Motherboard: só valores < -1 ou > 1
      if (valor > -1 && valor < 1) continue;

      contaCorrente.push({ nome, valor });
    }

    return json(200, { contaCorrente });

  } catch (err) {
    console.error("[cc-data]", err.message);
    return json(500, { erro: err.message });
  }
};
