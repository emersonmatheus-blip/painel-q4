// ---- Cliente do CRM Q5 (somente leitura) ----
// API pública (chave q5sk_...): conversas, negócios, usuários, campos.
// API interna (login de usuário): relatório de ligações, que a API pública não expõe.
const fs = require("fs");
const path = require("path");

const API = "https://api.grupoquatro5.com";
const PUBLIC = API + "/public/v1";
const INTERNAL = API + "/api/v1";
const CACHE_FILE = path.join(__dirname, "..", ".cache", "q5-deals.json");
const DEAL_MAX_AGE_MS = 6 * 3600 * 1000; // relê o negócio mesmo sem mudança na conversa a cada 6 h

class Q5Error extends Error {
  constructor(code, message) { super(message); this.code = code; }
}

function createQ5Client({ apiKey, email, password, concurrency = 4 } = {}) {
  if (!apiKey) throw new Q5Error("q5.sem-chave", "Defina Q5_API_KEY no .env.");

  async function pub(p, attempt = 0) {
    const r = await fetch(PUBLIC + p, { headers: { Authorization: "Bearer " + apiKey } });
    if (r.status === 429 && attempt < 6) {
      const wait = Number(r.headers.get("retry-after")) * 1000 || 1000 * 2 ** attempt;
      await new Promise((ok) => setTimeout(ok, wait));
      return pub(p, attempt + 1);
    }
    if (r.status === 401 || r.status === 403) throw new Q5Error("q5.sem-acesso", "Chave do Q5 sem acesso a " + p);
    if (r.status === 404 && /\/deal$/.test(p)) return null; // conversa sem negócio
    if (!r.ok) throw new Q5Error("q5.indisponivel", `Q5 respondeu ${r.status} em ${p}`);
    const text = await r.text();
    return text ? JSON.parse(text) : null;
  }

  async function pullConversations(extra) {
    const all = [];
    for (let page = 1; page < 200; page++) {
      const j = await pub(`/conversations?limit=100&page=${page}${extra}`);
      if (!j.data || !j.data.length) break;
      all.push(...j.data);
      if (j.total && all.length >= j.total) break;
    }
    return all;
  }

  // A listagem padrão não traz conversas encerradas (perdidas) nem resolvidas (ganhas).
  async function listAllConversations() {
    const lists = await Promise.all(["", "&status=closed", "&status=resolved"].map(pullConversations));
    return [...new Map(lists.flat().map((c) => [c.id, c])).values()];
  }

  function loadCache() {
    try { return JSON.parse(fs.readFileSync(CACHE_FILE, "utf8")); } catch { return {}; }
  }
  function saveCache(c) {
    fs.mkdirSync(path.dirname(CACHE_FILE), { recursive: true });
    fs.writeFileSync(CACHE_FILE, JSON.stringify(c));
  }

  // Busca o negócio de cada conversa; reaproveita o cache quando a conversa não mudou.
  async function dealsFor(convs, { force = false } = {}) {
    const cache = loadCache(), now = Date.now(), out = {};
    const todo = convs.filter((c) => {
      const hit = cache[c.id];
      if (!force && hit && hit.updatedAt === c.updatedAt && now - hit.fetchedAt < DEAL_MAX_AGE_MS) { out[c.id] = hit.deal; return false; }
      return true;
    });
    let i = 0;
    async function worker() {
      while (i < todo.length) {
        const c = todo[i++];
        try {
          const deal = await pub(`/conversations/${c.id}/deal`);
          out[c.id] = deal;
          cache[c.id] = { updatedAt: c.updatedAt, fetchedAt: now, deal };
        } catch (e) {
          if (cache[c.id]) out[c.id] = cache[c.id].deal; else throw e;
        }
      }
    }
    try { await Promise.all(Array.from({ length: concurrency }, worker)); }
    finally { saveCache(cache); } // guarda o que já leu, mesmo se parar no meio
    return { deals: out, fetched: todo.length };
  }

  // ---- Sessão de usuário para o relatório de ligações ----
  let cookies = null;
  async function login() {
    if (!email || !password) throw new Q5Error("q5.sem-login", "Defina Q5_EMAIL e Q5_PASSWORD no .env para ler as ligações.");
    const r = await fetch(INTERNAL + "/auth/login", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ email, password })
    });
    const body = await r.json().catch(() => ({}));
    if (!r.ok) throw new Q5Error("q5.login-falhou", "Login no Q5 falhou: " + (body.message || r.status));
    if (body && (body.mfaRequired || body.requiresMfa || body.mfa)) throw new Q5Error("q5.login-mfa", "O usuário do painel tem verificação em duas etapas; use um usuário sem MFA.");
    const set = typeof r.headers.getSetCookie === "function" ? r.headers.getSetCookie() : [r.headers.get("set-cookie")].filter(Boolean);
    cookies = set.map((s) => s.split(";")[0]).join("; ");
    if (!cookies) throw new Q5Error("q5.login-falhou", "Login no Q5 não devolveu sessão.");
  }
  async function internal(p, retry = true) {
    if (!cookies) await login();
    const r = await fetch(INTERNAL + p, { headers: { Cookie: cookies } });
    if (r.status === 401 && retry) { cookies = null; return internal(p, false); }
    if (r.status === 429 && retry) {
      await new Promise((ok) => setTimeout(ok, (Number(r.headers.get("retry-after")) || 5) * 1000));
      return internal(p, false);
    }
    if (!r.ok) throw new Q5Error("q5.indisponivel", `Q5 respondeu ${r.status} em ${p}`);
    return r.json();
  }
  // Total e atendidas de um SDR num intervalo [from, to] (datas ISO com fuso).
  async function callsSummary({ from, to, agentUserId }) {
    const q = new URLSearchParams({ from, to, agentUserId });
    return internal("/conversas/calls/summary?" + q);
  }

  // Etapas do funil com nome e ordem (só a API interna devolve o nome).
  async function chatStatuses() {
    const j = await internal("/conversas/chat-statuses");
    return (j.data || j).map((s) => ({ key: s.key, label: s.label, pipelineId: s.pipelineId, order: s.sortOrder }));
  }
  const HIST_FILE = path.join(path.dirname(CACHE_FILE), "q5-historico.json");
  // Histórico de mudanças de etapa por conversa; reaproveita o cache quando a conversa não mudou.
  async function stageHistories(convs) {
    let cache = {};
    try { cache = JSON.parse(fs.readFileSync(HIST_FILE, "utf8")); } catch {}
    const out = {};
    try {
      for (const c of convs) {
        const hit = cache[c.id];
        if (hit && hit.updatedAt === c.updatedAt) { out[c.id] = hit.h; continue; }
        const h = await internal(`/conversas/chats/${c.id}/deal-stage-history`);
        out[c.id] = (Array.isArray(h) ? h : h.data || []).map((x) => ({ from: x.fromStatusKey, to: x.toStatusKey, at: x.changedAt }));
        cache[c.id] = { updatedAt: c.updatedAt, h: out[c.id] };
      }
    } finally {
      fs.mkdirSync(path.dirname(HIST_FILE), { recursive: true });
      fs.writeFileSync(HIST_FILE, JSON.stringify(cache));
    }
    return out;
  }

  return {
    hasLogin: !!(email && password),
    chatStatuses, stageHistories,
    users: () => pub("/users").then((j) => j.data),
    dealFields: () => pub("/deal-fields").then((j) => j.data),
    listAllConversations, dealsFor, callsSummary
  };
}

module.exports = { createQ5Client, Q5Error };
