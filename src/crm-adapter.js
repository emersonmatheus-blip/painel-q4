// ---- Adaptador CRM Q5 -> aba Diário ----
// Converte negócios, histórico de etapas e ligações do Q5 em valores diários com as mesmas
// chaves do Diário ("<grupo>|<métrica>") e grava no workbook. O motor (model.js) roda sem mudança.
// Metas, Calendário, verba de mídia e Oportunidades continuam vindo da planilha.

const TZ = "America/Fortaleza";
const INICIO_Q4 = "2026-09-29"; // avisos de qualidade olham só o trimestre

// Regras combinadas com o Emerson (18/11/2026). Ordem das etapas = ordem oficial do funil Vendas no Q5.
const REGRAS = {
  pipelineVendas: "55cbe115-ea71-419a-84aa-75162cd5f57a",
  etapaAgendamento: "custom_experiencia_agendada",      // Experiência agendada (ordem 82)
  etapaRealizada: "custom_participou_da_experiencia",   // Participou da experiência (ordem 90)
  etapaReuniaoCloser: "custom_reuniao_realizada",       // Reunião realizada (ordem 125)
  // Venda = Contrato assinado ✓ ou ⚠ Assinado (antigo) (decisão do Emerson, 18/11/2026)
  etapasVenda: ["resolved", "custom_negocio_ganho"],
  ordemLimiteAvanço: 200                                // 200+ = Perdido e etapas antigas
};

const SDR = { "Weslley Nascimento": "Weslley", "Matheus Araujo": "Matheus" };
const CLOSER = { "Josue Oliveira": "Josué | Closer", "Thaiana Dias": "Thaiana | Closer" };
const CLOSER_OUTROS = "Outros | Closer (Emerson e SDRs)";
const PORTA_VENDA = { "Experiência Carbone": "Exp. Carbone", "Experiência Conselho": "Exp. Conselho", "Convidados LIGA": "Convidados LIGA", "Outros": "Outros" };

function dayKey(iso) {
  if (!iso) return null;
  const d = new Date(iso);
  return isNaN(d) ? null : new Intl.DateTimeFormat("en-CA", { timeZone: TZ }).format(d);
}
// Campos de data do CRM chegam como "2026-11-09" ou "17/11/2026 19:57".
function fieldDate(v) {
  if (!v) return null;
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) return `${m[3]}-${m[2]}-${m[1]}`;
  return null;
}
function brl(v) { return "R$ " + Math.round(v).toLocaleString("pt-BR"); }

// Chaves do funil por porta e etapa. Exp. Conselho só existe no grupo do Matheus; LIGA só no do Weslley.
function funilKey(porta, sdr, etapa) {
  if (porta === "Experiência Carbone") {
    if (!sdr) return null;
    return `${sdr} | Experiência Carbone|${{ lead: "Leads", agend: "Agendamentos", real: "Realizadas" }[etapa]}`;
  }
  if (porta === "Experiência Conselho") return `Matheus | Experiência Conselho|${{ lead: "Indicações", agend: "Agendamentos", real: "Realizadas" }[etapa]}`;
  if (porta === "Convidados LIGA") return `Weslley | Convidados LIGA|${{ lead: "Convites", agend: "Confirmados", real: "Presentes" }[etapa]}`;
  if (porta === "Outros") return `Funil Outros|${{ lead: "Leads", agend: "Agendamentos", real: "Realizadas" }[etapa]}`; // sem meta, sem SDR
  return null;
}

/**
 * @param convs     conversas do Q5
 * @param deals     { [conversationId]: deal }
 * @param users     usuários do Q5
 * @param statuses  etapas do Q5 [{key,label,order,pipelineId}] (opcional; sem elas usa só a etapa atual)
 * @param histories { [conversationId]: [{from,to,at}] } (opcional)
 * @param calls     { "<Nome SDR>": { "<AAAA-MM-DD>": {total, answered} } } (opcional)
 */
function crmToDaily({ convs, deals, users, statuses, histories, calls }) {
  const userName = Object.fromEntries((users || []).map((u) => [u.id, (u.name || "").trim()]));
  const ordem = {}, nomeEtapa = {};
  (statuses || []).forEach((s) => { if (s.pipelineId === REGRAS.pipelineVendas) ordem[s.key] = s.order; nomeEtapa[s.key] = s.label; });
  const ord = (k) => ordem[k];
  const avança = (k) => ord(k) !== undefined && ord(k) < REGRAS.ordemLimiteAvanço;

  const daily = {}, keys = new Set(), avisos = [], auditoria = [];
  const add = (key, day, v, r, origem, regra) => {
    if (!key || !day) return;
    keys.add(key);
    (daily[key] = daily[key] || {})[day] = (daily[key][day] || 0) + v;
    auditoria.push({ chave: key, dia: day, valor: v, regra, origem, conversa: r.c.id, nome: (r.c.contact && r.c.contact.name) || "", etapaAtual: nomeEtapa[r.c.status] || r.c.status });
  };

  const rows = convs.map((c) => ({ c, d: deals[c.id], h: (histories && histories[c.id]) || [] })).filter((r) => r.d);
  const fv = (r) => r.d.field_values || {};
  const sdrOf = (r) => SDR[fv(r).sdr] || SDR[userName[r.d.owner_user_id]] || null;
  const contato = (r) => (r.c.contact && (r.c.contact.phoneE164 || r.c.contact.id)) || r.d.contact_id;

  // Quando o negócio chegou pela primeira vez a uma etapa de ordem >= a da etapa alvo.
  // 1º histórico de etapas; 2º se a etapa atual já passou do alvo mas não há histórico, usa o campo de data.
  function chegouEm(r, etapaAlvo, camposData) {
    const alvo = ord(etapaAlvo);
    if (alvo !== undefined) {
      const passos = r.h.filter((x) => avança(x.to) && ord(x.to) >= alvo).sort((a, b) => a.at.localeCompare(b.at));
      if (passos.length) return { dia: dayKey(passos[0].at), origem: `histórico: entrou em "${nomeEtapa[passos[0].to]}" em ${passos[0].at}` };
    }
    const atualPassou = alvo !== undefined ? avança(r.c.status) && ord(r.c.status) >= alvo : false;
    if (atualPassou || alvo === undefined) {
      for (const campo of camposData) {
        const d = fieldDate(fv(r)[campo]);
        if (d) return { dia: d, origem: `etapa atual "${nomeEtapa[r.c.status] || r.c.status}" + campo ${campo}=${fv(r)[campo]}` };
      }
    }
    return null;
  }

  // Duplicados: o mesmo contato pode ter um negócio por WhatsApp de SDR. Fica só a última entrada.
  const comPorta = rows.filter((r) => fv(r)["porta-de-entrada"]);
  const ultimo = new Map();
  for (const r of comPorta) {
    const k = contato(r), prev = ultimo.get(k);
    if (!prev || (r.d.opened_at || "") > (prev.d.opened_at || "")) ultimo.set(k, r);
  }
  const validos = [...ultimo.values()];
  const duplicados = comPorta.filter((r) => !validos.includes(r));

  let semSdr = 0;
  const outrosPorta = [];
  for (const r of validos) {
    const porta = fv(r)["porta-de-entrada"], sdr = sdrOf(r);
    if (!PORTA_VENDA[porta]) { outrosPorta.push(r); continue; }
    if (porta === "Experiência Carbone" && !sdr) { if (dayKey(r.d.opened_at) >= INICIO_Q4) semSdr++; continue; }

    add(funilKey(porta, sdr, "lead"), dayKey(r.d.opened_at), 1, r, `negócio criado em ${r.d.opened_at}`, "Lead: negócio com Porta de entrada (última entrada do contato)");

    // Agendamento: chegou em Experiência agendada ou em qualquer etapa depois dela, inclusive quem foi direto
    // para Reunião agendada para Closer. Não exige Data da reunião (decisão do Emerson, 18/11/2026).
    const a = chegouEm(r, REGRAS.etapaAgendamento, ["data-da-reuniao", "data-da-reuniao-reagendada", "data-da-experiencia", "data-da-experiencia-reagendada"]);
    if (a) add(funilKey(porta, sdr, "agend"), a.dia, 1, r, a.origem, "Agendamento: chegou em Experiência agendada ou Reunião agendada para Closer (ou etapa posterior)");
    const e = chegouEm(r, REGRAS.etapaRealizada, ["data-da-experiencia-reagendada", "data-da-experiencia", "data-da-reuniao"]);
    if (e) add(funilKey(porta, sdr, "real"), e.dia, 1, r, e.origem, "Realizada: chegou em Participou da experiência");
    const m = chegouEm(r, REGRAS.etapaReuniaoCloser, ["data-da-reuniao-reagendada", "data-da-reuniao"]);
    if (m) add(`${CLOSER[fv(r).closer] || CLOSER_OUTROS}|Reuniões realizadas`, m.dia, 1, r, m.origem, "Reunião do closer: chegou em Reunião realizada");
  }

  // Vendas: negócio em Contrato assinado ✓ (ganho) ou ⚠ Assinado (antigo). Valor = valor do negócio.
  // Data: quando foi marcado como ganho; na etapa antiga (sem ganho marcado), quando entrou nela pelo histórico.
  function dataVenda(r) {
    if (r.d.won_at) return { dia: dayKey(r.d.won_at), origem: `marcado como ganho em ${r.d.won_at}` };
    const p = r.h.filter((x) => REGRAS.etapasVenda.includes(x.to)).sort((a, b) => a.at.localeCompare(b.at))[0];
    if (p) return { dia: dayKey(p.at), origem: `histórico: entrou em "${nomeEtapa[p.to] || p.to}" em ${p.at}` };
    const dc = fieldDate(fv(r)["data-do-contrato"]);
    if (dc) return { dia: dc, origem: `campo data-do-contrato=${fv(r)["data-do-contrato"]}` };
    return { dia: dayKey(r.c.updatedAt), origem: `sem data de ganho; usada a última atualização ${r.c.updatedAt}` };
  }
  const vendasSemPorta = [], vendasSemValor = [], vendasLista = [];
  let semCloser = 0;
  const NOME_CLOSER = { "Josue Oliveira": "Josué", "Thaiana Dias": "Thaiana" };
  const produtoDe = (r) => fv(r).produto || ((r.d.products || []).map((p) => p.name || p.product_name || (p.product && p.product.name)).filter(Boolean)[0]) || "";
  for (const r of rows.filter((x) => x.d.status === "won" || REGRAS.etapasVenda.includes(x.c.status))) {
    const tipo = PORTA_VENDA[fv(r)["porta-de-entrada"]];
    const valor = (r.d.value_cents || 0) / 100, { dia, origem: quando } = dataVenda(r);
    if (!tipo) { if (dia >= INICIO_Q4) vendasSemPorta.push({ valor, dia }); continue; }
    const grupo = CLOSER[fv(r).closer] || CLOSER_OUTROS;
    if (dia >= INICIO_Q4) { if (!CLOSER[fv(r).closer]) semCloser++; if (!valor) vendasSemValor.push(r); }
    const etapa = nomeEtapa[r.c.status] || r.c.status;
    const origem = `${quando} · etapa "${etapa}" · valor do negócio ${brl(valor)}`;
    add(`${grupo}|Vendas ${tipo}`, dia, 1, r, origem, `Venda: ${etapa}`);
    add(`${grupo}|VTV ${tipo}`, dia, valor, r, origem, "VTV: valor do negócio");
    if (dia >= INICIO_Q4) vendasLista.push({
      dia, valor, funil: fv(r)["porta-de-entrada"], closer: NOME_CLOSER[fv(r).closer] || "Outros",
      // SDR da venda segue a regra dos funis: Conselho é do Matheus, LIGA do Weslley, os demais pelo campo SDR
      sdr: { "Experiência Conselho": "Matheus", "Convidados LIGA": "Weslley" }[fv(r)["porta-de-entrada"]] || sdrOf(r) || "Sem SDR",
      nome: fv(r)["nome-do-contratante"] || (r.c.contact && r.c.contact.name) || "",
      empresa: fv(r)["nome-empresa"] || fv(r)["razao-social"] || "", produto: produtoDe(r),
      meses: Number(fv(r)["prazo-do-contrato-meses"]) || null, etapa
    });
  }

  // Ligações (relatório Relatórios > Ligações do Q5, por dia e SDR)
  if (calls) {
    for (const [nome, porDia] of Object.entries(calls)) {
      const sdr = SDR[nome] || nome;
      for (const [dia, s] of Object.entries(porDia)) {
        for (const [met, v] of [["Ligações realizadas", s.total || 0], ["Ligações atendidas", s.answered || 0]]) {
          const key = `${sdr} | Ligações|${met}`;
          keys.add(key);
          (daily[key] = daily[key] || {})[dia] = v;
          if (v) auditoria.push({ chave: key, dia, valor: v, regra: "Relatório de ligações do Q5", origem: `calls/summary ${nome}`, conversa: "", nome: "", etapaAtual: "" });
        }
      }
    }
  }

  // Avisos de qualidade do CRM (aparecem na lista de verificações do painel)
  const antigos = rows.filter((r) => r.c.status === "custom_negocio_ganho");
  if (vendasSemValor.length) avisos.push({ level: "warning", text: `CRM: ${vendasSemValor.length} venda(s) com valor do negócio zerado (contam na quantidade, não no VTV).` });
  if (semCloser) avisos.push({ level: "warning", text: `CRM: ${semCloser} venda(s) sem Closer preenchido (entram em Outros).` });
  if (vendasSemPorta.length) avisos.push({ level: "warning", text: `CRM: ${vendasSemPorta.length} venda(s) sem Porta de entrada preenchida (${brl(vendasSemPorta.reduce((s, x) => s + x.valor, 0))}) fora do VTV.` });
  if (semSdr) avisos.push({ level: "warning", text: `CRM: ${semSdr} lead(s) de Exp. Carbone sem SDR preenchido, fora da contagem.` });

  // Pendências de preenchimento no CRM (trimestre). Vão para o quadro "Qualidade do CRM" do painel.
  const qualidade = [
    { problema: "Vendas sem valor", qtd: vendasSemValor.length, valor: null },
    { problema: "Vendas sem closer", qtd: semCloser, valor: null },
    { problema: "Vendas sem porta de entrada", qtd: vendasSemPorta.length, valor: vendasSemPorta.reduce((s, x) => s + x.valor, 0) },
    { problema: "Leads sem SDR", qtd: semSdr, valor: null },
    { problema: "Em \"Assinado (antigo)\": mover para Contrato assinado", qtd: antigos.length, valor: antigos.reduce((s, r) => s + (r.d.value_cents || 0) / 100, 0) }
  ];

  return {
    daily, keys, avisos, auditoria, vendasLista, qualidade,
    resumo: {
      conversas: convs.length, negocios: rows.length, comPorta: comPorta.length, duplicadosDescartados: duplicados.length,
      portaOutros: outrosPorta.length, semSdr, assinadoAntigo: antigos.length, comHistorico: rows.filter((r) => r.h.length).length
    }
  };
}

// Grava os valores no Diário do workbook. Colunas vindas do CRM são reescritas por inteiro:
// 0 até hoje quando não houve nada, vazio no futuro. Colunas que faltarem são criadas no fim.
function applyToWorkbook(XLSX, wb, crm, todayKey) {
  const ws = wb.Sheets["Diário"];
  if (!ws) throw Object.assign(new Error("Aba Diário não encontrada na planilha."), { code: "planilha.estrutura" });
  const d = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const norm = (s) => String(s == null ? "" : s).replace(/\s+/g, " ").trim();
  const gRow = d[1] = d[1] || [], mRow = d[2] = d[2] || [];
  const cols = {};
  let g = "";
  for (let c = 0; c < mRow.length; c++) {
    if (gRow[c] !== null && gRow[c] !== undefined && gRow[c] !== "") g = norm(gRow[c]);
    if (mRow[c] && c >= 2) cols[g + "|" + norm(mRow[c])] = c;
  }
  // Funil Outros: cria as colunas mesmo sem evento, para o painel mostrar 0 e não "vazio"
  ["Funil Outros|Leads", "Funil Outros|Agendamentos", "Funil Outros|Realizadas"].forEach((k) => crm.keys.add(k));
  for (const key of crm.keys) {
    if (cols[key] !== undefined) continue;
    const i = key.lastIndexOf("|"), grupo = key.slice(0, i), metrica = key.slice(i + 1); // o grupo também tem "|"
    const c = Math.max(mRow.length, gRow.length);
    gRow[c] = grupo; mRow[c] = metrica; cols[key] = c;
  }
  // Todas as métricas que o CRM alimenta, mesmo sem nenhum evento ainda
  const crmKeys = new Set([...crm.keys, ...Object.keys(cols).filter((k) => /Experiência Carbone\|(Leads|Agendamentos|Realizadas)$|Experiência Conselho\||Convidados LIGA\||Funil Outros\||Closer.*\|(Reuniões realizadas|Vendas |VTV )/.test(k))]);
  if (crm.temLigacoes) Object.keys(cols).filter((k) => /\| Ligações\|/.test(k)).forEach((k) => crmKeys.add(k));

  for (let r = 3; r < d.length; r++) {
    const v = d[r][0];
    if (typeof v !== "number") continue;
    const day = new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 10);
    for (const key of crmKeys) d[r][cols[key]] = day <= todayKey ? ((crm.daily[key] && crm.daily[key][day]) || 0) : null;
  }
  wb.Sheets["Diário"] = XLSX.utils.aoa_to_sheet(d);
  if (!wb.Sheets["Oportunidades"]) {
    wb.Sheets["Oportunidades"] = XLSX.utils.aoa_to_sheet([["Data", "Nome", "Contato", "Status Reunião", "Closer", "Data Retorno", "Observação"]]);
    wb.SheetNames.push("Oportunidades");
  }
  return wb;
}

module.exports = { crmToDaily, applyToWorkbook, REGRAS, fieldDate, dayKey };
