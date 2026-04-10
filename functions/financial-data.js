/**
 * Netlify Serverless Function — NOPI Financial API
 * GET /.netlify/functions/financial-data
 */

const fetch = require("node-fetch");
const XLSX  = require("xlsx");

const EXCEL_URL = [
  "https://www.dropbox.com/scl/fi/y4i9m6v4q8snd2m3qljoh/Motherboard-2026.xlsx",
  "?rlkey=4px2hpxbg8p6fot2l65bkdamg&st=4h2vu72e&dl=1",
].join("");

const MONTHS = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
                "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];

const CORS = {
  "Access-Control-Allow-Origin":  "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

function toNum(v) {
  if (v == null || v === "" || String(v) === "nan") return null;
  const n = parseFloat(String(v).replace(",", "."));
  return Number.isFinite(n) ? Math.round(n * 100) / 100 : null;
}

function json(statusCode, body) {
  return {
    statusCode,
    headers: { ...CORS, "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0" },
    body: JSON.stringify(body),
  };
}

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 204, headers: CORS, body: "" };
  }

  try {
    const res = await fetch(EXCEL_URL, { timeout: 45_000 });
    if (!res.ok) throw new Error(`Dropbox respondeu HTTP ${res.status}`);
    const buf = await res.buffer();
    const wb = XLSX.read(buf, { type: "buffer", cellDates: true });

    // ── Folha MFC ─────────────────────────────────────────────────────────────
    const mfc = XLSX.utils.sheet_to_json(wb.Sheets["MFC"], { header: 1, defval: "" });

    const bancos = [];
    for (let i = 3; i <= 12; i++) {
      const row = mfc[i];
      const nome = String(row[0]).trim();
      if (!nome || nome === "nan") continue;
      const entry = { nome, saldoInicial: toNum(row[1]) };
      MONTHS.forEach((m, idx) => { entry[m] = toNum(row[idx + 2]); });
      bancos.push(entry);
    }

    const fielRow = mfc[13];
    const fielDepositario = { nome: "FIEL DEPOSITÁRIO", saldoInicial: toNum(fielRow[1]) };
    MONTHS.forEach((m, idx) => { fielDepositario[m] = toNum(fielRow[idx + 2]); });

    const saldoRow = mfc[14];
    const saldo = { nome: "SALDO", saldoInicial: toNum(saldoRow[1]) };
    MONTHS.forEach((m, idx) => { saldo[m] = toNum(saldoRow[idx + 2]); });

    const variacaoRow = mfc[15];
    const variacao = { nome: "VARIAÇÃO" };
    MONTHS.forEach((m, idx) => { variacao[m] = toNum(variacaoRow[idx + 2]); });

    const mapaFinanceiros = { bancos, fielDepositario, saldo, variacao };

    const linhasCaixa = [
      { row: 19, nome: "Caixa Inicial", tipo: "header" },
      { row: 21, nome: "Recebimentos de Clientes", tipo: "recebimento" },
      { row: 22, nome: "Pagamento de Comissões", tipo: "pagamento" },
      { row: 23, nome: "Pagamento a Fornecedores", tipo: "pagamento" },
      { row: 24, nome: "Pagamento de Salários", tipo: "pagamento" },
      { row: 25, nome: "Pagamento/Recebimento de Impostos", tipo: "pagamento" },
      { row: 26, nome: "Outros Pagamentos/Recebimentos", tipo: "pagamento" },
      { row: 27, nome: "Fluxos Operacionais", tipo: "subtotal" },
      { row: 31, nome: "Pagamentos de Investimentos Corpóreos", tipo: "pagamento" },
      { row: 33, nome: "Fluxos de Investimento", tipo: "subtotal" },
      { row: 35, nome: "Empréstimos Obtidos", tipo: "pagamento" },
      { row: 36, nome: "Empréstimos Concedidos", tipo: "pagamento" },
      { row: 37, nome: "Fluxos de Financiamento", tipo: "subtotal" },
      { row: 40, nome: "Fluxo de Caixa Líquido", tipo: "total" },
      { row: 41, nome: "Saldo de Caixa", tipo: "total" },
    ];

    const mapaCaixa = linhasCaixa.map(({ row, nome, tipo }) => {
      const r = mfc[row] || [];
      const entry = { nome, tipo };
      MONTHS.forEach((m, idx) => { entry[m] = toNum(r[idx + 2]); });
      entry.total = toNum(r[14]);
      return entry;
    });

    // ── Folha DR ──────────────────────────────────────────────────────────────
    const dr = XLSX.utils.sheet_to_json(wb.Sheets["DR"], { header: 1, defval: "" });

    // Resultado YTD — célula O36 (row 35, col 14)
    const resultadoYTD = toNum(dr[35][14]);

    // Despesas Fixas detalhadas (rows 4-26)
    // col 0 = categoriaId (coluna A), col 1 = nome (coluna B)
    const despesasFixas = [];
    for (let i = 4; i <= 26; i++) {
      const row = dr[i];
      const nome = String(row[1]).trim();
      if (!nome || nome === "nan") continue;
      const entry = {
        nome,
        categoriaId: toNum(row[0]),  // ← coluna A = ID da categoria
      };
      MONTHS.forEach((m, idx) => { entry[m] = toNum(row[idx + 2]); });
      entry.media = toNum(row[15]);
      entry.total = toNum(row[14]);
      despesasFixas.push(entry);
    }

    // Resumo DR — linhas principais
    const linhasDR = [
      { row: 3,  nome: "Total Despesas Fixas", tipo: "subtotal" },
      { row: 27, nome: "Total Despesas Variáveis", tipo: "subtotal" },
      { row: 33, nome: "Total Custos", tipo: "total" },
      { row: 34, nome: "Total Rendimentos", tipo: "total" },
      { row: 35, nome: "Resultado Económico", tipo: "resultado" },
    ];

    const resumoDR = linhasDR.map(({ row, nome, tipo }) => {
      const r = dr[row] || [];
      const entry = { nome, tipo };
      MONTHS.forEach((m, idx) => { entry[m] = toNum(r[idx + 2]); });
      entry.media = toNum(r[15]);
      entry.total = toNum(r[14]);
      return entry;
    });

    // Previsão de Despesas 2026 (rows 4-11, cols 19-25)
    // col 18 = categoriaId (coluna S), col 19 = nome (coluna T)
    const previsao = [];
    for (let i = 4; i <= 11; i++) {
      const row = dr[i];
      const nome = String(row[19]).trim();
      if (!nome || nome === "nan") continue;
      const orcamento = toNum(row[24]);
      const real = toNum(row[25]);
      previsao.push({
        nome,
        categoriaId:     toNum(row[18]),  // ← coluna S = ID da categoria
        nopiI:           toNum(row[20]),
        nopiII:          toNum(row[21]),
        nopiIII:         toNum(row[22]),
        totalNopi:       toNum(row[23]),
        orcamentoMensal: orcamento,
        resultadoMensal: real,
        desvio:          orcamento != null && real != null ? Math.round((real - orcamento) * 100) / 100 : null,
        desvioPercent:   orcamento != null && orcamento !== 0 && real != null
                           ? Math.round(((real - orcamento) / orcamento) * 10000) / 100 : null,
      });
    }

    const totalPrevisao = {
      nome: "TOTAL",
      nopiI:           toNum(dr[11][20]),
      nopiII:          toNum(dr[11][21]),
      nopiIII:         toNum(dr[11][22]),
      totalNopi:       toNum(dr[11][23]),
      orcamentoMensal: toNum(dr[11][24]),
      resultadoMensal: toNum(dr[11][25]),
    };

    return json(200, {
      ano: 2026,
      mfc: { mapaFinanceiros, mapaCaixa },
      dr:  { resultadoYTD, despesasFixas, resumoDR, previsao, totalPrevisao },
    });

  } catch (err) {
    console.error("[financial-data]", err.message);
    return json(500, { erro: err.message });
  }
};
