// ---- Modelo: lê o workbook (SheetJS) e calcula tudo que o painel mostra ----
var PEOPLE = [
  { name: "Weslley", role: "SDR", groups: ["Weslley | Experiência Carbone", "Weslley | Convidados LIGA", "Weslley | Ligações"] },
  { name: "Matheus", role: "SDR e BDR", groups: ["Matheus | Experiência Carbone", "Matheus | Experiência Conselho", "Matheus | Ligações"] },
  { name: "Josué", role: "Closer", groups: ["Josué | Closer"] },
  { name: "Thaiana", role: "Closer", groups: ["Thaiana | Closer"] }
];
var CLOSER_GROUPS = ["Josué | Closer", "Thaiana | Closer", "Outros | Closer (Emerson e SDRs)"];

function serialToKey(v) {
  if (typeof v !== "number" || !isFinite(v)) return null;
  var d = new Date(Math.round((v - 25569) * 86400000));
  return d.toISOString().slice(0, 10);
}
function keyToLabel(k) { return k.slice(8, 10) + "/" + k.slice(5, 7); }
function num(v) { return typeof v === "number" && isFinite(v) ? v : 0; }
function filled(v) { return v !== null && v !== undefined && v !== ""; }
function norm(s) { return String(s == null ? "" : s).replace(/\s+/g, " ").trim(); }

function rows(XLSX, ws) { return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }); }

function buildModel(XLSX, wb, todayKey, startKey, endKey) {
  // Diário
  var d = rows(XLSX, wb.Sheets["Diário"]);
  var gRow = d[1] || [], mRow = d[2] || [];
  var cols = {}, colGroup = [], g = "";
  for (var c = 0; c < mRow.length; c++) {
    if (filled(gRow[c])) g = norm(gRow[c]);
    colGroup[c] = g;
    if (filled(mRow[c]) && c >= 2) cols[g + "|" + norm(mRow[c])] = c;
  }
  var days = [];
  for (var r = 3; r < d.length; r++) {
    var k = serialToKey(d[r][0]);
    if (!k) continue;
    days.push({ key: k, dow: norm(d[r][1]), row: d[r] });
  }
  // Calendário
  var cal = rows(XLSX, wb.Sheets["Calendário"]);
  var hRow = cal[3] || [], calCols = {}, calByKey = {}, calTotalVtv = 0;
  for (c = 0; c < hRow.length; c++) {
    var h = norm(hRow[c]);
    var m = h.match(/^Meta (.+): (.+)$/);
    if (m) calCols[norm(m[1]) + "|" + norm(m[2])] = c;
    else if (h) calCols[h] = c;
  }
  for (r = 4; r < cal.length; r++) {
    k = serialToKey(cal[r][0]);
    if (!k) continue;
    calByKey[k] = cal[r];
    calTotalVtv += num(cal[r][calCols["Meta VTV total do dia"]]);
  }
  // Metas: meta oficial e limites de cor
  var metas = rows(XLSX, wb.Sheets["Metas"]), metaTri = 1500000, limV = 1, limA = 0.8;
  metas.forEach(function (row) {
    if (row[0] === "META_TRI" && typeof row[3] === "number") metaTri = row[3];
    if (row[0] === "LIM_V" && typeof row[3] === "number") limV = row[3];
    if (row[0] === "LIM_A" && typeof row[3] === "number") limA = row[3];
  });

  var firstKey = days.length ? days[0].key : todayKey;
  var lastKey = days.length ? days[days.length - 1].key : todayKey;
  var start = startKey && startKey > firstKey ? startKey : firstKey;
  var end = endKey && endKey < lastKey ? endKey : lastKey;
  if (end < start) end = start;
  var cutoff = todayKey < end ? todayKey : end;
  var notStarted = cutoff < start;
  var inRange = days.filter(function (x) { return x.key >= start && x.key <= cutoff; });
  var inPeriod = days.filter(function (x) { return x.key >= start && x.key <= end; });
  function isWork(k) { var cr = calByKey[k]; return cr ? num(cr[2]) > 0 : false; }

  function real(key) {
    var c = cols[key]; if (c === undefined) return 0;
    return inRange.reduce(function (s, x) { return s + num(x.row[c]); }, 0);
  }
  function meta(key) {
    var c = calCols[key]; if (c === undefined) return null;
    return inRange.reduce(function (s, x) { var cr = calByKey[x.key]; return s + (cr ? num(cr[c]) : 0); }, 0);
  }
  function metaP(key) {
    var c = calCols[key]; if (c === undefined) return null;
    return inPeriod.reduce(function (s, x) { var cr = calByKey[x.key]; return s + (cr ? num(cr[c]) : 0); }, 0);
  }
  function sumKeys(fn, keys) { return keys.reduce(function (s, k) { var v = fn(k); return s + (v || 0); }, 0); }
  function closerKeys(metric, groups) { return (groups || CLOSER_GROUPS).map(function (g) { return g + "|" + metric; }); }
  var JT = ["Josué | Closer", "Thaiana | Closer"];

  // VTV: série diária e acumulada; meta oficial distribuída pelo calendário do plano
  var vtvKeys = ["VTV Exp. Carbone", "VTV Exp. Conselho", "VTV Convidados LIGA"];
  var saleKeys = ["Vendas Exp. Carbone", "Vendas Exp. Conselho", "Vendas Convidados LIGA"];
  var cumReal = 0, cumPlan = 0, series = [];
  inPeriod.forEach(function (x) {
    var dayV = 0;
    vtvKeys.forEach(function (m) { CLOSER_GROUPS.forEach(function (gg) { var c = cols[gg + "|" + m]; if (c !== undefined) dayV += num(x.row[c]); }); });
    var cr = calByKey[x.key];
    cumPlan += cr ? num(cr[calCols["Meta VTV total do dia"]]) : 0;
    if (x.key <= cutoff) cumReal += dayV;
    series.push({ key: x.key, real: x.key <= cutoff && !notStarted ? cumReal : null, meta: calTotalVtv ? metaTri * cumPlan / calTotalVtv : 0 });
  });
  var vtvReal = cumReal;
  var vtvMetaHoje = 0;
  series.forEach(function (p) { if (p.key <= cutoff) vtvMetaHoje = p.meta; });
  var vendas = sumKeys(real, [].concat.apply([], saleKeys.map(function (m) { return closerKeys(m); })));
  var vendasMeta = sumKeys(meta, [].concat.apply([], saleKeys.map(function (m) { return closerKeys(m, JT); })));
  var verba = real("Mídia|Verba de tráfego"), verbaMeta = meta("Mídia|Verba de tráfego");

  // Funis
  function stage(label, rk, mk) { return { label: label, real: sumKeys(real, rk), meta: sumKeys(meta, mk || rk) }; }
  var funnels = [
    { name: "Experiência Carbone", stages: [
      stage("Leads", ["Weslley | Experiência Carbone|Leads", "Matheus | Experiência Carbone|Leads"]),
      stage("Agendamentos", ["Weslley | Experiência Carbone|Agendamentos", "Matheus | Experiência Carbone|Agendamentos"]),
      stage("Realizadas", ["Weslley | Experiência Carbone|Realizadas", "Matheus | Experiência Carbone|Realizadas"]),
      stage("Vendas", closerKeys("Vendas Exp. Carbone"), closerKeys("Vendas Exp. Carbone", JT))
    ], vtv: stage("VTV", closerKeys("VTV Exp. Carbone"), closerKeys("VTV Exp. Carbone", JT)) },
    { name: "Experiência Conselho", stages: [
      stage("Indicações", ["Matheus | Experiência Conselho|Indicações"]),
      stage("Agendamentos", ["Matheus | Experiência Conselho|Agendamentos"]),
      stage("Realizadas", ["Matheus | Experiência Conselho|Realizadas"]),
      stage("Vendas", closerKeys("Vendas Exp. Conselho"), closerKeys("Vendas Exp. Conselho", JT))
    ], vtv: stage("VTV", closerKeys("VTV Exp. Conselho"), closerKeys("VTV Exp. Conselho", JT)) },
    { name: "Convidados LIGA", stages: [
      stage("Convites", ["Weslley | Convidados LIGA|Convites"]),
      stage("Confirmados", ["Weslley | Convidados LIGA|Confirmados"]),
      stage("Presentes", ["Weslley | Convidados LIGA|Presentes"]),
      stage("Vendas", closerKeys("Vendas Convidados LIGA"), closerKeys("Vendas Convidados LIGA", JT))
    ], vtv: stage("VTV", closerKeys("VTV Convidados LIGA"), closerKeys("VTV Convidados LIGA", JT)) }
  ];

  // Lançamentos por pessoa e dia
  var grid = PEOPLE.map(function (p) {
    var pc = [];
    for (var key in cols) if (p.groups.some(function (gg) { return key.indexOf(gg + "|") === 0; })) pc.push(cols[key]);
    var cells = inRange.map(function (x) {
      var n = pc.filter(function (c) { return filled(x.row[c]); }).length;
      var st = n === 0 ? "vazio" : n === pc.length ? "completo" : "parcial";
      return { key: x.key, dow: x.dow, work: isWork(x.key), state: st, today: x.key === todayKey };
    });
    var req = cells.filter(function (c) { return c.work && !c.today; });
    var missing = req.filter(function (c) { return c.state === "vazio"; });
    var lastFilled = null;
    cells.forEach(function (c) { if (c.state !== "vazio") lastFilled = c.key; });
    return { person: p, cells: cells, required: req.length, missing: missing.length, partial: req.filter(function (c) { return c.state === "parcial"; }).length, lastFilled: lastFilled };
  });

  // Por pessoa
  function line(label, key, money) { return { label: label, real: real(key), meta: meta(key), money: !!money }; }
  function closerLines(g) {
    return [
      line("Reuniões realizadas", g + "|Reuniões realizadas"),
      { label: "Vendas", real: sumKeys(real, saleKeys.map(function (m) { return g + "|" + m; })), meta: sumKeys(meta, saleKeys.map(function (m) { return g + "|" + m; })) },
      { label: "VTV", real: sumKeys(real, vtvKeys.map(function (m) { return g + "|" + m; })), meta: sumKeys(meta, vtvKeys.map(function (m) { return g + "|" + m; })), money: true }
    ];
  }
  var people = [
    { name: "Weslley", lines: [
      line("Leads Exp. Carbone", "Weslley | Experiência Carbone|Leads"),
      line("Agendamentos Exp. Carbone", "Weslley | Experiência Carbone|Agendamentos"),
      line("Realizadas Exp. Carbone", "Weslley | Experiência Carbone|Realizadas"),
      line("Convites LIGA", "Weslley | Convidados LIGA|Convites"),
      line("Presentes LIGA", "Weslley | Convidados LIGA|Presentes"),
      { label: "Ligações (atendidas / feitas)", ratio: [real("Weslley | Ligações|Ligações atendidas"), real("Weslley | Ligações|Ligações realizadas")] }
    ] },
    { name: "Matheus", lines: [
      line("Leads Exp. Carbone", "Matheus | Experiência Carbone|Leads"),
      line("Agendamentos Exp. Carbone", "Matheus | Experiência Carbone|Agendamentos"),
      line("Realizadas Exp. Carbone", "Matheus | Experiência Carbone|Realizadas"),
      line("Indicações Conselho", "Matheus | Experiência Conselho|Indicações"),
      line("Realizadas Conselho", "Matheus | Experiência Conselho|Realizadas"),
      { label: "Ligações (atendidas / feitas)", ratio: [real("Matheus | Ligações|Ligações atendidas"), real("Matheus | Ligações|Ligações realizadas")] }
    ] },
    { name: "Josué", lines: closerLines("Josué | Closer") },
    { name: "Thaiana", lines: closerLines("Thaiana | Closer") }
  ];

  // Ligações por SDR
  var calls = ["Weslley", "Matheus"].map(function (n) {
    var fk = n + " | Ligações|Ligações realizadas", ak = n + " | Ligações|Ligações atendidas";
    var fc = cols[fk], ac = cols[ak];
    var daily = inRange.filter(function (x) { return isWork(x.key) || num(x.row[fc]) > 0; }).slice(-10).map(function (x) {
      return { key: x.key, dow: x.dow, f: num(x.row[fc]), a: num(x.row[ac]), filled: filled(x.row[fc]), today: x.key === todayKey };
    });
    return { name: n, feitas: { real: real(fk), meta: meta(fk), metaPer: metaP(fk) }, atend: { real: real(ak), meta: meta(ak), metaPer: metaP(ak) }, daily: daily };
  });

  // Oportunidades
  var op = rows(XLSX, wb.Sheets["Oportunidades"]), opps = [];
  for (r = 1; r < op.length; r++) {
    var o = op[r];
    if (!filled(o[1])) continue;
    opps.push({
      dateKey: serialToKey(o[0]), dateRaw: norm(o[0]), name: norm(o[1]), status: norm(o[3]),
      closer: norm(o[4]) === "-" ? "" : norm(o[4]), retKey: serialToKey(o[5]), note: norm(o[6])
    });
  }

  // Alertas de qualidade de dado
  var alerts = [];
  grid.forEach(function (gr) {
    if (gr.missing > 0) alerts.push({ level: "critical", text: gr.person.name + " não lançou " + gr.missing + " de " + gr.required + " dias úteis do período" + (gr.lastFilled ? " (último lançamento em " + keyToLabel(gr.lastFilled) + ")" : "") + "." });
  });
  [["Weslley", "Weslley | Experiência Carbone"], ["Matheus", "Matheus | Experiência Carbone"]].forEach(function (p) {
    var l = real(p[1] + "|Leads"), a = real(p[1] + "|Agendamentos");
    if (a > l) alerts.push({ level: "warning", text: p[0] + ": " + a + " agend. para " + l + " leads (Exp. Carbone)" });
  });
  opps.forEach(function (o) {
    if (o.dateKey && o.dateKey > todayKey) alerts.push({ level: "warning", text: "Oportunidade de " + o.name + " com data futura (" + keyToLabel(o.dateKey) + "/" + o.dateKey.slice(0, 4) + ")." });
  });
  var fechadas = opps.filter(function (o) { return /fechad/i.test(o.status); });
  if (fechadas.length && vendas === 0) alerts.push({ level: "warning", text: fechadas.length + " Fechado em Oportunidades sem venda lançada no Diário no período." });

  return {
    todayKey: todayKey, cutoff: cutoff, start: start, end: end, notStarted: notStarted, firstKey: firstKey, lastKey: lastKey,
    metaPeriodo: series.length ? series[series.length - 1].meta : 0,
    metaTri: metaTri, limV: limV, limA: limA,
    vtv: { real: vtvReal, metaHoje: vtvMetaHoje, proj: cutoff >= end ? vtvReal : vtvMetaHoje ? vtvReal / vtvMetaHoje * (series.length ? series[series.length - 1].meta : 0) : 0 },
    vendas: { real: vendas, meta: vendasMeta }, verba: { real: verba, meta: verbaMeta },
    series: series, calls: calls, funnels: funnels, grid: grid, people: people, opps: opps, alerts: alerts,
    workDaysTotal: Object.keys(calByKey).filter(function (k) { return k >= start && k <= end && isWork(k); }).length,
    workDaysElapsed: Object.keys(calByKey).filter(function (k) { return k >= start && k <= cutoff && isWork(k); }).length
  };
}

if (typeof module !== "undefined" && module.exports) module.exports = { buildModel: buildModel };
