// ---- Robô do painel: gera site/dados.xlsx para o Netlify ----
// 1. Baixa a planilha publicada no Google (metas, calendário, verba, follow-ups).
// 2. Lê o CRM Q5 e reescreve o Diário (leads, agendamentos, realizadas, vendas, VTV, ligações).
// 3. Remove telefones e abas que o painel não usa, e grava site/dados.xlsx.
// Uso local: node scripts/gerar-dados.js   (lê o .env)
// No GitHub: roda sozinho a cada 15 min (.github/workflows/atualizar-painel.yml)
const path = require("path");
const fs = require("fs");
require("dotenv").config({ path: path.join(__dirname, "..", ".env") });
const XLSX = require("xlsx");
const { preencherComCrm, hojeFortaleza } = require("../src/fonte-crm");

const PLANILHA_PUBLICADA = process.env.PLANILHA_URL ||
  "https://docs.google.com/spreadsheets/d/e/2PACX-1vRm6njo5pQQzANpXaA_zmrd3Q5l_-jSM1pAwYGXKDl6Dno2y0-t403PjQkWYEm_ug_iAZ__7kRXLik2/pub?output=xlsx";
const SAIDA = path.join(__dirname, "..", "site", "dados.xlsx");
const ABAS = ["Diário", "Calendário", "Metas", "Oportunidades"]; // "Fechamentos" sai: as vendas agora vêm do CRM

(async () => {
  const t = Date.now();
  const r = await fetch(PLANILHA_PUBLICADA + "&t=" + t);
  if (!r.ok) throw new Error("Não consegui baixar a planilha publicada (HTTP " + r.status + ").");
  const wb = XLSX.read(Buffer.from(await r.arrayBuffer()), { type: "buffer" });

  const hoje = hojeFortaleza();
  const { crm, avisos } = await preencherComCrm(XLSX, wb, { hoje, force: process.argv.includes("--force") });

  // Oportunidades: apaga a coluna C (Contato/telefone) antes de publicar
  if (wb.Sheets["Oportunidades"]) {
    const op = XLSX.utils.sheet_to_json(wb.Sheets["Oportunidades"], { header: 1, raw: true, defval: null });
    op.forEach((row, i) => { if (i > 0 && row.length > 2) row[2] = null; });
    wb.Sheets["Oportunidades"] = XLSX.utils.aoa_to_sheet(op);
  }

  const out = XLSX.utils.book_new();
  for (const nome of ABAS) if (wb.Sheets[nome]) XLSX.utils.book_append_sheet(out, wb.Sheets[nome], nome);

  // Vendas do CRM: alimenta o quadro Fechamentos (só lista; os totais continuam vindo do Diário)
  const serial = (dia) => Date.parse(dia + "T00:00:00Z") / 864e5 + 25569;
  XLSX.utils.book_append_sheet(out, XLSX.utils.aoa_to_sheet([
    ["Data", "Membro", "Empresa", "Comunidade", "Turma", "Cadeira", "Closer", "Funil", "Valor mensal", "Fidelidade (meses)", "VTV", "Observação", "SDR"],
    ...crm.vendasLista.map((v) => [serial(v.dia), v.nome, v.empresa, v.produto, "", "", v.closer, v.funil, null, v.meses, v.valor, v.etapa, v.sdr])
  ]), "Vendas CRM");
  // Pendências de preenchimento no CRM: alimenta o quadro Qualidade do CRM
  XLSX.utils.book_append_sheet(out, XLSX.utils.aoa_to_sheet([
    ["Problema", "Quantidade", "Valor"],
    ...crm.qualidade.map((q) => [q.problema, q.qtd, q.valor])
  ]), "Qualidade CRM");

  fs.mkdirSync(path.dirname(SAIDA), { recursive: true });
  XLSX.writeFile(out, SAIDA);
  // Versão da página: o painel aberto na TV compara e se recarrega quando muda
  const versao = require("crypto").createHash("sha1").update(fs.readFileSync(path.join(path.dirname(SAIDA), "index.html"))).digest("hex").slice(0, 12);
  fs.writeFileSync(path.join(path.dirname(SAIDA), "versao.txt"), versao);
  // Hora em que estes dados saíram do CRM: o botão Atualizar espera este arquivo mudar
  fs.writeFileSync(path.join(path.dirname(SAIDA), "status.json"), JSON.stringify({ gerado: new Date().toISOString() }));

  console.log(`dados.xlsx gerado em ${((Date.now() - t) / 1000).toFixed(0)} s · hoje ${hoje} · ${JSON.stringify(crm.resumo)}`);
  avisos.forEach((a) => console.log(`[${a.level}] ${a.text}`));
})().catch((e) => { console.error("ERRO:", e.message); process.exit(1); });
