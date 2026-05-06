/**
 * Netlify Serverless Function — NOPI Expenses Detail API
 * GET /.netlify/functions/expenses-data
 * Devolve detalhe de despesas por categoria e mês da folha MOTHER
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

// Colunas relevantes (0-indexed)
const COL = {
  DATA:      59,
  ENTIDADE:  62,
  BASE:      70,
  TOTAL:     73,
  BANCO:     74,
  CATEGORIA: 77,
  MES:       44,
  ANO:       51,
};

// Categorias a incluir
const CATEGORIAS = [
  "SALARIOS",
  "GERENCIA",
  "PUBLICIDADE",
  "COMISSOES",
  "NEGOCIOS",
  "BONUS",
  "IMPOSTOS",
  "ESCRITORIO",
  "FORMACAO",
  "DESLOCACOES",
  "SEGUROS",
  "AUTOMOVEIS",
  "EVENTOS",
  "AVENÇAS",
  "INVESTIMENTO",
  "IMPRESSORAS",
  "AGUA",
];

const MESES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
               "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];

const toNum = (v) => {
  if (v == null || v === "" || String(v) === "nan") return null;
  const n = parseFloat(String(v).replace(",", "."));
  return Number.isFinite(n) ? Math.round(n * 100) / 100 : null;
};

const toStr = (v) => v != null && v !== "" ? String(v).trim() : null;

const toDate = (v) => {
  if (!v) return null;
  if (v instanceof Date) return v.toISOString().split("T")[0];
  if (typeof v === "number") {
    try {
      const d = XLSX.SSF.parse_date_code(v);
      return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;
    } catch { return null; }
  }
  return String(v).split("T")[0];
};

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
    const rows = XLSX.utils.sheet_to_json(wb.Sheets["MOTHER"], { header: 1, defval: null });
    const dataRows = rows.slice(1);

    // Estrutura: { [categoria]: { [ano]: { [mes]: { total: number, linhas: [] } } } }
    const resultado = {};

    CATEGORIAS.forEach(cat => { resultado[cat] = {}; });

    dataRows.forEach(r => {
      const cat = toStr(r[COL.CATEGORIA])?.toUpperCase();
      if (!cat || !CATEGORIAS.includes(cat)) return;

      const data = toDate(r[COL.DATA]);
      if (!data) return;

      const ano = parseInt(data.split("-")[0]);
      const mes = parseInt(data.split("-")[1]);
      if (!ano || !mes) return;

      const total = toNum(r[COL.TOTAL]);
      if (total == null) return;

      const entidade = toStr(r[COL.ENTIDADE]) || "—";
      const nomeMes = MESES[mes - 1];

      if (!resultado[cat][ano]) resultado[cat][ano] = {};
      if (!resultado[cat][ano][nomeMes]) {
        resultado[cat][ano][nomeMes] = { total: 0, linhas: [] };
      }

      resultado[cat][ano][nomeMes].total = Math.round(
        (resultado[cat][ano][nomeMes].total + total) * 100
      ) / 100;

      resultado[cat][ano][nomeMes].linhas.push({
        data,
        entidade,
        total,
      });
    });

    // Converter para formato mais simples para o frontend
    // { categoria: string, anos: { [ano]: { [mes]: { total, linhas } } } }[]
    const despesas = CATEGORIAS
      .filter(cat => Object.keys(resultado[cat]).length > 0)
      .map(cat => ({
        categoria: cat,
        anos: resultado[cat],
      }));

    return json(200, { despesas });

  } catch (err) {
    console.error("[expenses-data]", err.message);
    return json(500, { erro: err.message });
  }
};
