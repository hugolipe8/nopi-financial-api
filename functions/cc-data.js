/**
 * Netlify Serverless Function — NOPI Conta Corrente API
 * GET /.netlify/functions/cc-data
 */
const fetch = require("node-fetch");
const XLSX  = require("xlsx");
const EXCEL_URL = [
  "https://www.dropbox.com/scl/fi/q1e1l6enrinhm8ileg903/Motherboard-2026.xlsx",
  "?rlkey=lke29p1fipcrj8l4dl3hqb8gi&st=hrc3v22k&dl=1",
].join("");
const CORS = {
  "Access-Control-Allow-Origin":  "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
  "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
};
const CATEGORIAS = {
  "c":   "Clientes",
  "ct":  "Consultores",
  "f":   "Entidades Financeiras",
  "fd":  "Reserva",
  "fin": "Financiamentos",
  "pf":  "Participações Financeiras",
};
const IGNORAR = new Set(["q","nopi","total geral","(em branco)"]);
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
    let categoriaAtual = null;
    for (let i = 5; i < rows.length; i++) {
      const nome = toStr(rows[i][0]);
      if (!nome) continue;
      const nomeLower = nome.toLowerCase();
      if (IGNORAR.has(nomeLower)) continue;
      if (/^\d{4}\/\d+$/.test(nome)) continue;
      if (CATEGORIAS[nomeLower]) {
        categoriaAtual = CATEGORIAS[nomeLower];
        continue;
      }
      const valor = toNum(rows[i][3]);
      if (valor === null) continue;
      if (valor > -1 && valor < 1) continue;
      contaCorrente.push({ nome, valor, categoria: categoriaAtual || "Outro" });
    }
    const porCategoria = {};
    contaCorrente.forEach(({ nome, valor, categoria }) => {
      if (!porCategoria[categoria]) porCategoria[categoria] = [];
      porCategoria[categoria].push({ nome, valor });
    });
    return json(200, { contaCorrente, porCategoria });
  } catch (err) {
    console.error("[cc-data]", err.message);
    return json(500, { erro: err.message });
  }
};
