// ---- Monta o modelo do painel com o CRM Q5 como fonte ----
// Planilha: Calendário, Metas, verba de mídia e Oportunidades. CRM: leads, agendamentos,
// realizadas, vendas, VTV e (com login) ligações.
const fs = require("fs");
const path = require("path");
const { createQ5Client } = require("./q5");
const { crmToDaily, applyToWorkbook } = require("./crm-adapter");
const { buildModel } = require("./model");

const CALLS_CACHE = path.join(__dirname, "..", ".cache", "q5-ligacoes.json");
const SDRS = ["Weslley Nascimento", "Matheus Araujo"];

function hojeFortaleza() {
  return new Intl.DateTimeFormat("en-CA", { timeZone: "America/Fortaleza" }).format(new Date());
}
function diasEntre(a, b) {
  const out = [];
  for (let d = new Date(a + "T12:00:00Z"); d.toISOString().slice(0, 10) <= b; d.setUTCDate(d.getUTCDate() + 1)) out.push(d.toISOString().slice(0, 10));
  return out;
}

// Ligações por dia e SDR. Dias passados ficam em cache; hoje é sempre relido.
// Negócio com marcação de anúncio (UTM ou Tipo_lead = Ads): entra no funil Exp. Carbone mesmo sem porta
function ehAdsDeal(d) { const f = (d && d.field_values) || {}; return !!(f["utm-source"] || f["utm-campaign"] || f["tipo-lead"] === "Ads"); }

async function lerLigacoes(q5, users, primeiroDia, hoje) {
  let cache = {};
  try { cache = JSON.parse(fs.readFileSync(CALLS_CACHE, "utf8")); } catch {}
  const out = {};
  for (const nome of SDRS) {
    const u = users.find((x) => (x.name || "").trim() === nome);
    if (!u) continue;
    out[nome] = {};
    for (const dia of diasEntre(primeiroDia, hoje)) {
      const ck = u.id + "|" + dia;
      if (dia < hoje && cache[ck]) { out[nome][dia] = cache[ck]; continue; }
      const s = await q5.callsSummary({ from: `${dia}T00:00:00-03:00`, to: `${dia}T23:59:59.999-03:00`, agentUserId: u.id });
      out[nome][dia] = cache[ck] = { total: s.total || 0, answered: s.answered || 0 };
    }
  }
  fs.mkdirSync(path.dirname(CALLS_CACHE), { recursive: true });
  fs.writeFileSync(CALLS_CACHE, JSON.stringify(cache));
  return out;
}

// Cada número que o CRM gerou, com a regra e a origem. Fica só na máquina (.cache, fora do git).
const AUDIT_FILE = path.join(__dirname, "..", ".cache", "auditoria-crm.csv");
function salvarAuditoria(linhas) {
  const cols = ["dia", "chave", "valor", "regra", "origem", "nome", "etapaAtual", "conversa"];
  const esc = (v) => `"${String(v == null ? "" : v).replace(/"/g, '""')}"`;
  const csv = [cols.join(";"), ...linhas.sort((a, b) => (a.dia + a.chave).localeCompare(b.dia + b.chave)).map((l) => cols.map((c) => esc(l[c])).join(";"))].join("\n");
  fs.mkdirSync(path.dirname(AUDIT_FILE), { recursive: true });
  fs.writeFileSync(AUDIT_FILE, "﻿" + csv);
}

async function montarPainel(XLSX, wb, { hoje = hojeFortaleza(), inicio, fim, force = false, env = process.env } = {}) {
  const { crm, avisos } = await preencherComCrm(XLSX, wb, { hoje, force, env });
  const M = buildModel(XLSX, wb, hoje, inicio || null, fim || null);
  M.alerts.push(...avisos);
  M.fonte = { crm: crm.resumo, ligacoes: crm.temLigacoes ? "crm" : "planilha", lidoEm: new Date().toISOString() };
  return M;
}

// Lê o CRM e reescreve o Diário do workbook. Devolve o resumo e os avisos de qualidade.
async function preencherComCrm(XLSX, wb, { hoje = hojeFortaleza(), force = false, env = process.env } = {}) {
  const q5 = createQ5Client({ apiKey: env.Q5_API_KEY, email: env.Q5_EMAIL, password: env.Q5_PASSWORD });
  const [users, convs] = await Promise.all([q5.users(), q5.listAllConversations()]);
  const { deals } = await q5.dealsFor(convs, { force });
  const M_avisos = [];

  let calls = null, statuses = null, histories = null, avisoLigacoes = null;
  if (q5.hasLogin) {
    try {
      statuses = await q5.chatStatuses();
      // Histórico só dos negócios que entram na conta (com porta de entrada ou ganhos)
      const relevantes = convs.filter((c) => deals[c.id] && ((deals[c.id].field_values || {})["porta-de-entrada"] || deals[c.id].status === "won" || ehAdsDeal(deals[c.id])));
      histories = await q5.stageHistories(relevantes);
    } catch (e) { M_avisos.push({ level: "warning", text: `Histórico de etapas não lido (${e.message}); usando etapa atual e campos de data.` }); }
    try { calls = await lerLigacoes(q5, users, "2026-09-29", hoje); }
    catch (e) { avisoLigacoes = `Ligações não lidas do CRM (${e.message}). Usando a planilha.`; }
  } else {
    avisoLigacoes = "Ligações ainda vêm da planilha: falta o login do Q5 no .env.";
  }

  const crm = crmToDaily({ convs, deals, users, statuses, histories, calls });
  crm.temLigacoes = !!calls;
  applyToWorkbook(XLSX, wb, crm, hoje);
  salvarAuditoria(crm.auditoria);
  const avisos = [...M_avisos, ...crm.avisos];
  if (avisoLigacoes) avisos.push({ level: "warning", text: avisoLigacoes });
  return { crm, avisos };
}

module.exports = { montarPainel, preencherComCrm, hojeFortaleza };
