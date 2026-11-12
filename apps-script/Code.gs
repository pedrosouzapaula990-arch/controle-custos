const PLANILHA_ID = "1QWnNvsYs8kQDJXolMl6dTxsYGS-sCMvS2QJektqFBGw";
const PASTA_COMPROVANTES = "Comprovantes - Controle de Custos";

// Cada aba da planilha e suas colunas
const ABAS = {
  Custos:  ["id","data","descricao","categoria","valor","pagoPor","status","obs","criadoEm","mesRef","comprovante"],
  Suporte: ["id","tipo","data","hora","responsavel","fila","total","status","dados","obs","criadoEm"]
};
// Colunas que NÃO são texto (o resto fica como texto para a planilha não converter datas/horas)
const NUMERICAS = ["valor","criadoEm","fila","total"];

function nomeAba_(p) {
  return p && p.aba && ABAS[p.aba] ? p.aba : "Custos";
}

function aba_(nome) {
  const cols = ABAS[nome];
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  const head = sh.getRange(1, 1, 1, cols.length).getValues()[0];
  if (head.join("|") !== cols.join("|")) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    sh.setFrozenRows(1);
    cols.forEach((c, i) => {
      if (NUMERICAS.indexOf(c) < 0) sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat("@");
    });
  }
  return sh;
}

function ehData_(v) {
  return v && typeof v.getTime === "function";
}

function listar_(nome) {
  const cols = ABAS[nome];
  const sh = aba_(nome);
  const vals = sh.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  return vals.slice(1).filter(r => r[0]).map(r => {
    const o = {};
    cols.forEach((c, i) => {
      let v = r[i] === undefined ? "" : r[i];
      if (ehData_(v)) {
        if (c === "hora") v = Utilities.formatDate(v, tz, "HH:mm");
        else if (c === "mesRef") v = Utilities.formatDate(v, tz, "yyyy-MM");
        else v = Utilities.formatDate(v, tz, "yyyy-MM-dd");
      }
      o[c] = v;
    });
    if ("valor" in o) o.valor = Number(o.valor) || 0;
    return o;
  });
}

function pasta_() {
  const it = DriveApp.getFoldersByName(PASTA_COMPROVANTES);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_COMPROVANTES);
}

function executar_(action, p) {
  const nome = nomeAba_(p);
  const cols = ABAS[nome];
  if (action === "list") return { ok: true, aba: nome, items: listar_(nome) };
  if (action === "validar") { const o = validar_(p); o.kommo = statusKommo_(); return o; }
  if (action === "kommoStatus") return statusKommo_();
  if (action === "puxarKommo") return puxarBasesKommo();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = aba_(nome);
    const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => String(r[0]));

    if (action === "save") {
      const it = typeof p.item === "string" ? JSON.parse(p.item) : p.item;
      const row = cols.map(c => it[c] === undefined || it[c] === null ? "" : it[c]);
      const idx = ids.indexOf(String(it.id));
      const linha = idx > 0 ? idx + 1 : sh.getLastRow() + 1;
      sh.getRange(linha, 1, 1, cols.length).setValues([row]);
      return { ok: true, aba: nome };
    }

    if (action === "delete") {
      const idx = ids.indexOf(String(p.id));
      if (idx > 0) sh.deleteRow(idx + 1);
      return { ok: true, aba: nome };
    }

    if (action === "upload") {
      const idx = ids.indexOf(String(p.id));
      if (idx < 1) return { ok: false, error: "lançamento não encontrado" };
      const blob = Utilities.newBlob(Utilities.base64Decode(p.base64), p.mime || "application/octet-stream", p.nome || "comprovante");
      const arq = pasta_().createFile(blob);
      arq.setName(p.id + " - " + (p.nome || "comprovante"));
      try { arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
      const url = arq.getUrl();
      if (cols.indexOf("comprovante") >= 0) sh.getRange(idx + 1, cols.indexOf("comprovante") + 1).setValue(url);
      if (nome === "Custos") sh.getRange(idx + 1, cols.indexOf("status") + 1).setValue("pago");
      return { ok: true, url: url };
    }

    return { ok: false, error: "acao invalida" };
  } finally {
    lock.releaseLock();
  }
}

function responder_(obj, callback) {
  const txt = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + txt + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try { out = executar_(p.action || "list", p); }
  catch (err) { out = { ok: false, error: String(err) }; }
  return responder_(out, p.callback);
}

function doPost(e) {
  let out;
  try {
    const body = JSON.parse(e.postData.contents);
    out = executar_(body.action, body);
  } catch (err) { out = { ok: false, error: String(err) }; }
  return responder_(out);
}

// Rode 1x pelo editor (Run) para autorizar e criar as abas
function autorizar() {
  pasta_();
  Object.keys(ABAS).forEach(aba_);
  const sb = basesSS_();
  Logger.log("Planilha das bases (compartilhe esta com o suporte): " + sb.getUrl());
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === "recalcularSeMudou")) {
    ScriptApp.newTrigger("recalcularSeMudou").timeBased().everyMinutes(5).create();
  }
  recalcularSeMudou();
  Logger.log("Pronto: o cálculo das bases roda sozinho a cada 5 minutos, só quando as bases mudam.");
}

// Planilha SEPARADA só com as bases (o suporte tem acesso a ela, não à de custos)
function basesSS_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty("BASES_ID"), sb = null;
  if (id) { try { sb = SpreadsheetApp.openById(id); } catch (e) { sb = null; } }
  if (!sb) {
    sb = SpreadsheetApp.create("Bases do Suporte – Darb Play");
    props.setProperty("BASES_ID", sb.getId());
    sb.getSheets()[0].setName(BASES[0]);
  }
  BASES.forEach(n => { if (!sb.getSheetByName(n)) sb.insertSheet(n); });
  return sb;
}

// ===================== VALIDAÇÃO E CRÉDITO =====================
// O suporte cola a base inteira (exportação do Kommo) nas abas abaixo
const BASES = ["Base Darb Oficial", "Base Motiva"];
const COLS_OPC = ["Última Data Vencimento", "Modificado por"];
const BASE_INICIO_MES = { "2026-11": 5775 }; // clientes ativos no início do mês (informado pelo Pedro)
const COLS_BASE = ["ID", "Contato principal", "Etapa do lead", "Data Cadastro", "Data Renovação", "Plano", "Telas", "Servidor", "Vendedor", "Data de Vencimento"];

// O cálculo roda em segundo plano (gatilho a cada 5 min) e só quando a planilha das bases muda.
// O site só lê o último resultado salvo, então abre na hora. "Recalcular" força o cálculo.
function versaoAtual_(desde) {
  const tz = Session.getScriptTimeZone();
  const mes = Utilities.formatDate(new Date(), tz, "yyyy-MM");
  const alterada = DriveApp.getFileById(basesSS_().getId()).getLastUpdated();
  return { versao: alterada.getTime() + "|" + (desde || "") + "|" + mes, alterada: alterada };
}
function lerSalvo_() {
  try { const x = JSON.parse(arquivoCalculo_().getBlob().getDataAsString()); return x && x.out ? x : null; } catch (e) { return null; }
}
function validar_(p) {
  const desde = p.desde || "";
  const salvo = lerSalvo_();
  if (!p.forcar && salvo && (salvo.out.desde || "") === desde) {
    salvo.out.doCache = true;
    try { salvo.out.desatualizado = salvo.versao !== versaoAtual_(desde).versao; } catch (e) {}
    return salvo.out;
  }
  return calcularESalvar_(desde);
}
function calcularESalvar_(desde) {
  const tz = Session.getScriptTimeZone();
  const agora = new Date();
  const hoje = Utilities.formatDate(agora, tz, "yyyy-MM-dd");
  const ss = basesSS_();
  const va = versaoAtual_(desde);
  const bases = BASES.map(nome => {
    const sh = ss.getSheetByName(nome);
    if (!sh || sh.getLastRow() < 2) return { nome: nome, head: [], cols: {} };
    const lastCol = sh.getLastColumn(), n = sh.getLastRow() - 1;
    const head = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h).trim());
    const cols = {};
    COLS_BASE.concat(COLS_OPC).forEach(c => {
      const i = head.indexOf(c);
      if (i >= 0) cols[c] = sh.getRange(2, i + 1, n, 1).getDisplayValues().map(r => r[0]);
    });
    return { nome: nome, head: head, cols: cols, n: n };
  });
  const ini = inicioAuto_(hoje);
  const out = calcularValidacao(bases, hoje, desde, ini.valor);
  out.ativos.origemInicio = ini.origem;
  try { out.ativos.movimento = fotoAtivos_(hoje, out._ids, tz); } catch (e) { out.ativos.movimentoErro = String(e.message || e); }
  delete out._ids;
  try {
    const vg = JSON.parse(PropertiesService.getScriptProperties().getProperty("KOMMO_VENDA_GANHA") || "[]");
    vg.forEach(x => out.erros.push({ base: x.base, id: x.id, linha: "-", cliente: x.nome, vendedor: x.resp, etapa: "Venda ganha", origem: "Etapa", data: "",
      nivel: "alerta", tipo: "Cliente em Venda ganha", detalhe: "A etapa Venda ganha não deve ter clientes. Mover para a etapa correta no Kommo." }));
  } catch (e) {}
  out.ok = true; out.aba = "Suporte"; out.basesUrl = ss.getUrl();
  out.basesAlteradas = Utilities.formatDate(va.alterada, tz, "dd/MM HH:mm");
  out.basesAlteradasMs = va.alterada.getTime();
  out.calculadoEm = Utilities.formatDate(agora, tz, "dd/MM HH:mm");
  try { arquivoCalculo_().setContent(JSON.stringify({ versao: va.versao, out: out })); } catch (e) {}
  return out;
}
// Foto diária dos IDs ativos/vencidos (uma por dia, sobrescrita ao longo do dia) e comparação com o início do mês e com ontem
function fotoDia_(data, criar) {
  const nome = "ativos-" + data + ".json", pasta = pasta_(), it = pasta.getFilesByName(nome);
  if (it.hasNext()) return it.next();
  return criar ? pasta.createFile(nome, "{}", "application/json") : null;
}
function lerFoto_(data) { const f = fotoDia_(data, false); if (!f) return null; try { const x = JSON.parse(f.getBlob().getDataAsString()); return x && x.a ? x : null; } catch (e) { return null; } }
function fotoAtivos_(hoje, ids, tz) {
  fotoDia_(hoje, true).setContent(JSON.stringify(ids));
  const mes = hoje.slice(0, 7), ini = mes + "-01";
  // referência de início do mês: a primeira foto do mês; senão a mais antiga disponível
  let refData = null, refFoto = null;
  for (let d = 1; d <= +hoje.slice(8, 10); d++) {
    const dd = mes + "-" + ("0" + d).slice(-2), f = lerFoto_(dd);
    if (f) { refData = dd; refFoto = f; break; }
  }
  const mov = { fotoHoje: hoje, refInicio: refData, entradas: null, saidas: null, reconquistas: null, entradasLista: [], saidasLista: [] };
  if (refFoto && refData !== hoje) {
    const antes = {}; refFoto.a.forEach(i => antes[i] = 1);
    const agora = {}; ids.a.forEach(i => agora[i] = 1);
    const venc0 = {}; (refFoto.v || []).forEach(i => venc0[i] = 1);
    const ent = ids.a.filter(i => !antes[i]), sai = refFoto.a.filter(i => !agora[i]);
    mov.entradas = ent.length; mov.saidas = sai.length; mov.reconquistas = ent.filter(i => venc0[i]).length;
    mov.entradasLista = ent.slice(0, 50); mov.saidasLista = sai.slice(0, 50);
    mov.ativosRef = refFoto.a.length;
  }
  // ontem
  const ont = new Date(Date.parse(hoje + "T12:00:00Z") - 86400000).toISOString().slice(0, 10), fo = lerFoto_(ont);
  if (fo) { mov.ontemAtivos = fo.a.length; mov.variacaoDia = ids.a.length - fo.a.length; }
  return mov;
}
// Início do mês automático: ativos da Base Darb Oficial na foto do último dia do mês anterior
// (se não houver, a primeira foto do mês; se não houver, o valor fixo informado).
function inicioAuto_(hoje) {
  const base = "Base Darb Oficial", mes = hoje.slice(0, 7);
  const ultimoAnt = new Date(Date.parse(mes + "-01T12:00:00Z") - 86400000).toISOString().slice(0, 10);
  const f = lerFoto_(ultimoAnt);
  if (f && f.n && f.n[base] != null) return { valor: f.n[base], origem: "foto de " + ultimoAnt.split("-").reverse().join("/") };
  if (BASE_INICIO_MES[mes] != null) return { valor: BASE_INICIO_MES[mes], origem: "valor informado" };
  for (let d = 1; d <= +hoje.slice(8, 10); d++) {
    const dd = mes + "-" + ("0" + d).slice(-2), g = lerFoto_(dd);
    if (g && g.n && g.n[base] != null) return { valor: g.n[base], origem: "primeira foto do mês (" + dd.split("-").reverse().join("/") + ")" };
  }
  return { valor: null, origem: "sem dados" };
}
// Gatilho diário 23:50: fecha a foto do dia (a do último dia do mês vira o início do mês seguinte)
function fotoFimDoDia() {
  const salvo = lerSalvo_();
  calcularESalvar_(salvo ? (salvo.out.desde || "") : "");
}
// Cria sozinho os gatilhos que faltarem (sem precisar abrir o Apps Script)
function garantirGatilhos_() {
  const p = PropertiesService.getScriptProperties(), tz = Session.getScriptTimeZone();
  const hoje = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  if (p.getProperty("GATILHOS_OK") === hoje) return;
  const ex = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  if (ex.indexOf("fotoFimDoDia") < 0) ScriptApp.newTrigger("fotoFimDoDia").timeBased().everyDays(1).atHour(23).nearMinute(50).create();
  if (ex.indexOf("puxarBasesKommo") < 0 && p.getProperty("KOMMO_TOKEN")) { try { configurarKommo(); } catch (e) {} }
  p.setProperty("GATILHOS_OK", hoje);
}
// Gatilho (a cada 5 min): recalcula só se as bases mudaram
function recalcularSeMudou() {
  try { garantirGatilhos_(); } catch (e) {}
  const salvo = lerSalvo_();
  const desde = salvo ? (salvo.out.desde || "") : "";
  if (salvo && salvo.versao === versaoAtual_(desde).versao) return;
  calcularESalvar_(desde);
}
function arquivoCalculo_() {
  const nome = "ultimo-calculo-validacao.json", pasta = pasta_();
  const it = pasta.getFilesByName(nome);
  return it.hasNext() ? it.next() : pasta.createFile(nome, "{}", "application/json");
}

// ---- regra (pura, sem dependência do Google) ----
function calcularValidacao(bases, hoje, desde, inicioAuto) {
  const MESES = { "mensal": 1, "bimestral": 2, "dois meses": 2, "trimestral": 3, "semestral": 6, "anual": 12 };
  const mesAtual = hoje.slice(0, 7);
  const parse = s => {
    s = String(s || "").trim(); if (!s) return null;
    let m = s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})/);
    if (m) return m[3] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2);
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? m[1] + "-" + m[2] + "-" + m[3] : null;
  };
  const addMeses = (iso, k) => {
    const [y, mo, d] = iso.split("-").map(Number);
    const dt = new Date(Date.UTC(y, mo - 1 + k, 1));
    const ult = new Date(Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth() + 1, 0)).getUTCDate();
    dt.setUTCDate(Math.min(d, ult));
    return dt.toISOString().slice(0, 10);
  };
  const dias = (a, b) => Math.round((Date.parse(a) - Date.parse(b)) / 86400000);
  const br = iso => iso ? iso.slice(8, 10) + "/" + iso.slice(5, 7) + "/" + iso.slice(0, 4) : "";

  const consumo = { total: 0, vendas: 0, renovacoes: 0, linhasVendas: 0, linhasRenovacoes: 0, porBase: {}, porServidor: {} };
  let consumoDesde = desde ? 0 : null;
  const erros = [], infoBases = [];
  const fimMes = new Date(Date.UTC(+hoje.slice(0, 4), +hoje.slice(5, 7), 0)).getUTCDate();
  const diaHoje = +hoje.slice(8, 10);
  const ativosIds = [], vencidosIds = [];
  const ativ = { agora: 0, porBase: {}, vendasMes: 0, naoRenovaram: 0, renovados: 0, venceRestoMes: 0, vendasHoje: 0 };
  const vend = {}, detAt = {};

  bases.forEach(b => {
    const faltando = COLS_BASE.filter(c => !b.cols[c]);
    infoBases.push({ nome: b.nome, linhas: b.n || 0, vazia: !b.n, faltando: b.n ? faltando : [] });
    if (!b.n || faltando.length) return;
    consumo.porBase[b.nome] = 0;
    consumo.detalheBase = consumo.detalheBase || {};
    const det = consumo.detalheBase[b.nome] = { vendas: 0, renovacoes: 0, clientesVendas: 0, clientesRenovacoes: 0 };
    for (let r = 0; r < b.n; r++) {
      const v = c => String((b.cols[c] || [])[r] || "").trim();
      if (/duplicad/i.test(v("Etapa do lead"))) continue;
      const dc = parse(v("Data Cadastro")), dr = parse(v("Data Renovação"));
      const inCad = dc && dc.slice(0, 7) === mesAtual && dc <= hoje;
      const inRen = dr && dr.slice(0, 7) === mesAtual && dr <= hoje;
      const planoTxt = v("Plano"), meses = MESES[planoTxt.toLowerCase()];
      const telasTxt = v("Telas"), telasN = /^\d+$/.test(telasTxt) ? +telasTxt : NaN;
      const telasOk = telasN >= 1 && telasN <= 10;
      const cred = (telasOk ? telasN : 1) * (meses || 1);
      const servs = v("Servidor").split(",").map(x => x.trim()).filter(Boolean);
      if (desde) {
        if (dc && dc >= desde && dc <= hoje) consumoDesde += cred;
        if (dr && dr >= desde && dr <= hoje) consumoDesde += cred;
      }
      // Ativos / conciliação (toda a base, sem duplicados)
      const dvA = parse(v("Data de Vencimento"));
      const dB = detAt[b.nome] = detAt[b.nome] || { agora: 0, vendasMes: 0, naoRenovaram: 0, renovados: 0, venceRestoMes: 0 };
      if (dvA) {
        const idA = v("ID");
        if (dvA >= hoje) { ativ.agora++; ativ.porBase[b.nome] = (ativ.porBase[b.nome] || 0) + 1; ativosIds.push(idA);
          dB.agora++; if (dvA.slice(0, 7) === mesAtual && dvA > hoje) { ativ.venceRestoMes++; dB.venceRestoMes++; } }
        else { vencidosIds.push(idA); if (dvA.slice(0, 7) === mesAtual) { ativ.naoRenovaram++; dB.naoRenovaram++; } }
      }
      if (inCad) { ativ.vendasMes++; dB.vendasMes++; if (dc === hoje) ativ.vendasHoje++; }
      if (dr && dr.slice(0, 7) === mesAtual && dr <= hoje) { ativ.renovados++; dB.renovados++; }
      const vnm = v("Vendedor") || "(sem vendedor)";
      if (inCad) { vend[vnm] = vend[vnm] || { vendas: 0, comErro: 0, ids: {} }; vend[vnm].vendas++; }
      // Tabulação: renovou no mês mas o vencimento não avançou / está em etapa de vencido
      const ult = parse(v("Última Data Vencimento"));
      if (inRen && dvA && ult && dvA <= ult) erros.push({ base: b.nome, id: v("ID"), linha: r + 2, cliente: v("Contato principal"), vendedor: v("Vendedor"), etapa: v("Etapa do lead"),
        origem: "Tabulação", data: br(dr), nivel: "alerta", tipo: "Renovou sem avançar vencimento", detalhe: "Renovação " + br(dr) + ", mas o vencimento (" + br(dvA) + ") não passou da última data (" + br(ult) + "). Modificado por: " + (v("Modificado por") || "—") });
      // Etapa x vencimento (vale para toda a base, não só o mês)
      const etp = v("Etapa do lead").toLowerCase(), dvE = parse(v("Data de Vencimento"));
      if (dvE && !/vital/.test(etp)) {
        const etapaAtiva = /clientes ativos|vence em 3|vence em 2|vence amanh|vence hoje/.test(etp);
        const etapaVencida = /venceu ontem|vencidos|reconquista/.test(etp);
        const cadErro = (tipo, det) => erros.push({ base: b.nome, id: v("ID"), linha: r + 2, cliente: v("Contato principal"), vendedor: v("Vendedor"),
          etapa: v("Etapa do lead"), origem: "Etapa x vencimento", data: br(dvE), nivel: "erro", tipo: tipo, detalhe: det });
        if (etapaAtiva && dvE < hoje) cadErro("Vencido em etapa de ativo", "Vencimento " + br(dvE) + " já passou, mas está em " + v("Etapa do lead"));
        else if (etapaVencida && dvE >= hoje) cadErro("Ativo em etapa de vencido", "Vencimento " + br(dvE) + " ainda não passou, mas está em " + v("Etapa do lead"));
      }
      if (!inCad && !inRen) continue;
      const vezes = (inCad ? 1 : 0) + (inRen ? 1 : 0);
      consumo.total += cred * vezes; consumo.porBase[b.nome] += cred * vezes;
      if (inCad) { consumo.vendas += cred; consumo.linhasVendas++; det.vendas += cred; det.clientesVendas++; }
      if (inRen) { consumo.renovacoes += cred; consumo.linhasRenovacoes++; det.renovacoes += cred; det.clientesRenovacoes++; }
      const chaveServ = servs.length ? servs.join(", ") : "Sem servidor";
      consumo.porServidor[chaveServ] = (consumo.porServidor[chaveServ] || 0) + cred * vezes;

      const base = { base: b.nome, id: v("ID"), linha: r + 2, cliente: v("Contato principal"), vendedor: v("Vendedor"), etapa: v("Etapa do lead"),
        origem: inCad && inRen ? "Venda e renovação" : inCad ? "Venda" : "Renovação", data: br(inRen ? dr : dc) };
      const add = (nivel, tipo, detalhe) => erros.push(Object.assign({}, base, { nivel: nivel, tipo: tipo, detalhe: detalhe }));
      if (!servs.length) add("erro", "Servidor vazio", "Preencher o servidor");
      if (!planoTxt) add("erro", "Plano vazio", "Contado como 1 mês");
      else if (!meses) add("erro", "Plano não reconhecido", "Plano: " + planoTxt);
      if (!telasTxt) add("erro", "Telas vazio", "Contado como 1 tela");
      else if (!telasOk) add("erro", "Telas inválido", "Telas: " + telasTxt);
      if (inCad && !v("Vendedor")) add("erro", "Venda sem vendedor", "Preencher o vendedor");
      if (inCad && inRen) add("alerta", "Venda e renovação no mesmo mês", "Confirmar se foi venda proporcional");
      if (servs.length > 1 && telasOk && servs.length > telasN) add("alerta", "Servidores x telas", servs.length + " servidores para " + telasN + " tela(s)");
      const dv = parse(v("Data de Vencimento"));
      const ref = inRen ? dr : dc;
      if (!dv) add("erro", "Sem data de vencimento", "Preencher o vencimento");
      else if (meses) {
        const esp = addMeses(ref, meses), dif = dias(dv, esp);
        if (Math.abs(dif) > 2) add("alerta", "Vencimento fora do plano", "Vence " + br(dv) + ", esperado " + br(esp) + " (" + planoTxt + " a partir de " + br(ref) + ")");
      }
    }
  });
  // qualidade por vendedor: vendas do mês com algum erro/alerta de dados
  erros.forEach(e => { const k = e.vendedor || "(sem vendedor)"; if (vend[k] && /Venda/.test(e.origem) && e.nivel === "erro") vend[k].ids[e.base + e.id] = 1; });
  Object.keys(vend).forEach(k => { vend[k].comErro = Object.keys(vend[k].ids).length; delete vend[k].ids; });
  // Conciliação e previsão: só da Base Darb Oficial (o início do mês informado é dela)
  const D = detAt["Base Darb Oficial"] || { agora: 0, vendasMes: 0, naoRenovaram: 0, renovados: 0, venceRestoMes: 0 };
  const inicio = inicioAuto != null ? inicioAuto : (BASE_INICIO_MES[mesAtual] || null);
  const esperado = inicio == null ? null : inicio + D.vendasMes - D.naoRenovaram;
  const taxaRen = (D.renovados + D.naoRenovaram) ? D.renovados / (D.renovados + D.naoRenovaram) : 1;
  const diasRest = fimMes - diaHoje;
  const prevFim = Math.round(D.agora + (D.vendasMes / Math.max(1, diaHoje)) * diasRest - D.venceRestoMes * (1 - taxaRen));
  const ativos = { base: "Base Darb Oficial", agora: D.agora, total: ativ.agora, porBase: ativ.porBase, inicioMes: inicio, vendasMes: D.vendasMes, naoRenovaram: D.naoRenovaram, esperado: esperado,
    diferenca: esperado == null ? null : D.agora - esperado, renovados: D.renovados, taxaRenovacao: Math.round(taxaRen * 1000) / 10,
    venceRestoMes: D.venceRestoMes, diasRestantes: diasRest, previstoFimMes: prevFim, diaDoMes: diaHoje };
  return { mes: mesAtual, hoje: hoje, desde: desde || "", bases: infoBases, consumo: consumo, consumoDesde: consumoDesde, erros: erros,
    ativos: ativos, vendedores: vend, _ids: { a: ativosIds, v: vencidosIds, n: ativ.porBase } };
}

// ===================== API DO KOMMO =====================
// Puxa as bases direto do Kommo (mesma conta, funis diferentes), grava na planilha das bases,
// guarda uma "foto" datada no Drive e recalcula a validação/crédito.
// Configuração (Apps Script > Configurações do projeto > Propriedades do script):
//   KOMMO_SUBDOMINIO = começo do endereço (ex.: "suaempresa" de suaempresa.kommo.com)
//   KOMMO_TOKEN      = token de longa duração
const KOMMO_FUNIS = { "DarbplayOficial": "Base Darb Oficial", "Suporte": "Base Motiva" };
const COLS_EXPORT = ["ID", "Contato principal", "Etapa do lead", "Funil de vendas", "Lead usuário responsável",
  "Data Criada", "Última modificação", "Modificado por", "Data Cadastro", "Plano", "Telas", "Servidor",
  "Aplicativo Instalado", "Vendedor", "Data de Vencimento", "Última Data Vencimento", "Data Renovação",
  "Renovação", "Fila de Renovação", "Tabulação", "Venda"];
const PASTA_FOTOS = "Fotos das bases (Kommo)";

function kommoCfg_() {
  const pr = PropertiesService.getScriptProperties();
  let sub = String(pr.getProperty("KOMMO_SUBDOMINIO") || "").trim().replace(/^https?:\/\//, "").replace(/\.kommo\.com.*$/, "");
  const token = String(pr.getProperty("KOMMO_TOKEN") || "").trim();
  if (!sub || !token) throw new Error("Falta configurar KOMMO_SUBDOMINIO e KOMMO_TOKEN nas Propriedades do script.");
  return { base: "https://" + sub + ".kommo.com/api/v4/", token: token };
}
function kommoReq_(cfg, path) {
  return { url: cfg.base + path, headers: { Authorization: "Bearer " + cfg.token }, muteHttpExceptions: true };
}
function kommoJson_(resp) {
  const c = resp.getResponseCode();
  if (c === 204) return null;
  if (c === 401) throw new Error("Token do Kommo inválido ou expirado (401).");
  if (c === 429) throw new Error("Kommo pediu para ir mais devagar (429).");
  if (c >= 400) throw new Error("Kommo respondeu " + c + ": " + resp.getContentText().slice(0, 200));
  return JSON.parse(resp.getContentText());
}
// busca todas as páginas de um endpoint, 6 páginas por vez (limite do Kommo: 7 req/s)
function kommoTudo_(cfg, path, chave) {
  const out = [];
  let pag = 1, fim = false;
  while (!fim) {
    const reqs = [];
    for (let i = 0; i < 7; i++) reqs.push(kommoReq_(cfg, path + (path.indexOf("?") >= 0 ? "&" : "?") + "limit=250&page=" + (pag + i)));
    const resps = UrlFetchApp.fetchAll(reqs);
    for (const r of resps) {
      const j = kommoJson_(r);
      const itens = j && j._embedded && j._embedded[chave] || [];
      if (!itens.length) { fim = true; break; }
      out.push.apply(out, itens);
      if (itens.length < 250) { fim = true; break; }
    }
    pag += 7;
    if (!fim) Utilities.sleep(1000);
  }
  return out;
}
const normNome_ = t => String(t || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/\s+/g, " ").trim();
function fmtCampo_(cf, tz) {
  const vals = (cf.values || []).map(v => v.value).filter(v => v !== null && v !== undefined && v !== "");
  if (!vals.length) return "";
  if (cf.field_type === "date" || cf.field_type === "birthday") return vals.map(v => Utilities.formatDate(new Date(+v * 1000), tz, "dd.MM.yyyy")).join(", ");
  if (cf.field_type === "date_time") return vals.map(v => Utilities.formatDate(new Date(+v * 1000), tz, "dd.MM.yyyy HH:mm")).join(", ");
  if (cf.field_type === "checkbox") return vals[0] ? "Sim" : "";
  return vals.join(", ");
}

// Puxa as duas bases EM PARTES: trabalha até ~4 min, salva onde parou e continua sozinha 1 min depois.
// Pode ser chamada pelo gatilho, pelo editor ou pelo botão "Puxar do Kommo" do site.
const KOMMO_ORCAMENTO_MS = 4 * 60 * 1000;
function puxarBasesKommo() { return rodarKommo_(false); }
function continuarKommo() { return rodarKommo_(true); }

function rodarKommo_(continuacao) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return { ok: true, jaRodando: true };
  const props = PropertiesService.getScriptProperties();
  const t0 = Date.now(), tz = Session.getScriptTimeZone();
  let job = null;
  try {
    // remove gatilhos de continuação antigos
    ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "continuarKommo").forEach(t => ScriptApp.deleteTrigger(t));
    job = lerJob_();
    if (job && !continuacao && Date.now() - job.inicio > 30 * 60 * 1000) job = null; // job velho/travado: recomeça
    if (job && !continuacao) { agendarContinuacao_(); return { ok: true, jaRodando: true, job: resumoJob_(job) }; }
    const cfg = kommoCfg_();
    if (!job) job = novoJob_(cfg);
    // trava de segurança: no máximo 12 partes por puxada (evita gastar a cota diária do Google)
    job.partes = (job.partes || 0) + 1;
    if (job.partes > 12) { apagarJob_(); throw new Error("Puxada cancelada: passou de 12 partes. Avise para investigar."); }
    salvarJob_(job);
    props.setProperty("KOMMO_PUXANDO", String(Date.now()));
    const sb = basesSS_(), cache = lerNomesCache_();
    let nomesNovos = 0;
    const tempo = () => Date.now() - t0 < KOMMO_ORCAMENTO_MS;

    // FASE 1: leads, página a página, gravando numa aba temporária
    while (job.fase === "leads" && tempo()) {
      const f = job.funis[job.i];
      if (!f) { job.fase = "nomes"; break; }
      const reqs = [];
      for (let k = 0; k < 5; k++) reqs.push(kommoReq_(cfg, "leads?with=contacts&limit=250&page=" + (f.pag + k) + "&" + f.filtro));
      let acabou = false; const linhas = [];
      UrlFetchApp.fetchAll(reqs).forEach(r => {
        if (acabou) return;
        const j = kommoJson_(r), itens = (j && j._embedded && j._embedded.leads) || [];
        itens.forEach(l => linhas.push(linhaLead_(l, job, tz)));
        if (itens.length < 250) acabou = true;
      });
      if (linhas.length) {
        const sh = abaTemp_(sb, f.aba);
        sh.getRange(sh.getLastRow() + 1, 1, linhas.length, COLS_EXPORT.length + 1).setNumberFormat("@").setValues(linhas);
        f.n += linhas.length;
      }
      f.pag += 5;
      if (acabou) job.i++;
      salvarJob_(job);
      Utilities.sleep(800);
    }

    // FASE 2: nomes dos contatos que ainda não conhecemos
    if (job.fase === "nomes") {
      const faltam = [];
      job.funis.forEach(f => {
        const sh = sb.getSheetByName("_tmp " + f.aba);
        if (!sh || sh.getLastRow() < 1) return;
        sh.getRange(1, COLS_EXPORT.length + 1, sh.getLastRow(), 1).getValues().forEach(r => { const id = String(r[0]); if (id && !(id in cache)) faltam.push(id); });
      });
      const unicos = [...new Set(faltam)], LOTE = 80;
      for (let i = 0; i < unicos.length && tempo(); i += LOTE * 6) {
        const reqs = [];
        for (let k = i; k < Math.min(i + LOTE * 6, unicos.length); k += LOTE)
          reqs.push(kommoReq_(cfg, "contacts?limit=250&" + unicos.slice(k, k + LOTE).map(id => "filter[id][]=" + id).join("&")));
        UrlFetchApp.fetchAll(reqs).forEach(r => { const j = kommoJson_(r); ((j && j._embedded && j._embedded.contacts) || []).forEach(c => { cache[c.id] = c.name; nomesNovos++; }); });
        unicos.slice(i, i + LOTE * 6).forEach(id => { if (!(id in cache)) cache[id] = ""; }); // não procura de novo quem não existe
        Utilities.sleep(900);
      }
      if (nomesNovos) salvarNomesCache_(cache);
      if (unicos.every(id => id in cache)) job.fase = "finalizar";
      salvarJob_(job);
    }

    // FASE 3: troca as bases, guarda a foto no Drive e recalcula
    if (job.fase === "finalizar" && Date.now() - t0 < 2.5 * 60 * 1000) {
      const agora = new Date(), resumo = [], conf = [];
      const iNome = COLS_EXPORT.indexOf("Contato principal");
      job.funis.forEach(f => {
        const tmp = sb.getSheetByName("_tmp " + f.aba);
        const dadosTmp = tmp && tmp.getLastRow() ? tmp.getRange(1, 1, tmp.getLastRow(), COLS_EXPORT.length + 1).getValues() : [];
        // segurança: mantém só os leads do funil certo e fora de DUPLICADOS (caso o Kommo ignore o filtro)
        const funilCerto = (Object.keys(KOMMO_FUNIS).find(k => KOMMO_FUNIS[k] === f.aba) || "").toLowerCase();
        const iFunil = COLS_EXPORT.indexOf("Funil de vendas"), iEtapa = COLS_EXPORT.indexOf("Etapa do lead"), vistos = {};
        const linhas = dadosTmp.filter(r => {
          if (String(r[iFunil]).trim().toLowerCase() !== funilCerto || /duplicad/i.test(r[iEtapa]) || (job.fechadas || []).indexOf(String(r[iEtapa])) >= 0 || vistos[r[0]]) return false;
          vistos[r[0]] = 1; return true;
        }).map(r => { const c = r.slice(0, COLS_EXPORT.length).map(String); c[iNome] = cache[String(r[COLS_EXPORT.length])] || ""; return c; });
        const sh = sb.getSheetByName(f.aba) || sb.insertSheet(f.aba);
        sh.clearContents();
        const dados = [COLS_EXPORT].concat(linhas);
        sh.getRange(1, 1, dados.length, COLS_EXPORT.length).setNumberFormat("@").setValues(dados);
        const csv = dados.map(r => r.map(v => { v = String(v); return /[",\n;]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }).join(";")).join("\n");
        const nomeArq = Utilities.formatDate(agora, tz, "yyyy-MM-dd_HHmm") + " - " + f.aba + ".csv";
        pastaFotos_().createFile(Utilities.gzip(Utilities.newBlob(csv, "text/csv", nomeArq), nomeArq + ".gz"));
        if (tmp) sb.deleteSheet(tmp);
        resumo.push(f.aba + ": " + linhas.length + " leads");
        // conferência: tudo que o Kommo entregou foi gravado? (as etapas fechadas e duplicadas já nem são pedidas)
        const gravadas = Math.max(0, sh.getLastRow() - 1);
        const vazias = COLS_BASE.concat(COLS_OPC).filter(c => { const ix = COLS_EXPORT.indexOf(c); return ix >= 0 && linhas.length > 50 && !linhas.some(r => String(r[ix] || "").trim() !== ""); });
        conf.push({ vazias: vazias, aba: f.aba, buscados: dadosTmp.length, gravados: gravadas, esperados: linhas.length, ok: dadosTmp.length > 0 && gravadas === linhas.length && linhas.length === dadosTmp.length });
      });
      // Venda ganha não deve ter clientes: guarda os que aparecerem para virar alerta (só a 1ª página, 1 chamada por funil)
      try {
        const vg = [];
        job.funis.forEach(f => {
          if (!f.pipeId) return;
          const rq = kommoReq_(cfg, "leads?limit=250&filter[statuses][0][pipeline_id]=" + f.pipeId + "&filter[statuses][0][status_id]=142");
          const j = kommoJson_(UrlFetchApp.fetch(rq.url, rq));
          ((j && j._embedded && j._embedded.leads) || []).forEach(l => vg.push({ base: f.aba, id: String(l.id), nome: String(l.name || ""), resp: (job.usuarios || {})[l.responsible_user_id] || "" }));
        });
        props.setProperty("KOMMO_VENDA_GANHA", JSON.stringify(vg.slice(0, 200)));
      } catch (e) { Logger.log("Venda ganha: " + e); }
      apagarJob_();
      props.setProperty("KOMMO_ULTIMA", Utilities.formatDate(agora, tz, "dd/MM HH:mm"));
      const completa = conf.length === job.funis.length && conf.every(x => x.ok);
      const colVazias = conf.filter(x => x.vazias.length).map(x => x.aba + ": " + x.vazias.join(", "));
      props.setProperty("KOMMO_CONF", JSON.stringify({ quando: Utilities.formatDate(agora, tz, "dd/MM HH:mm"), completa: completa,
        colunasVazias: colVazias.join(" | "), campos: Object.keys(job.campos || {}).slice(0, 80).join(", "),
        detalhe: completa ? "" : conf.filter(x => !x.ok).map(x => x.aba + ": Kommo entregou " + x.buscados + ", gravados " + x.gravados).join(" | ") || "alguma base não foi puxada" }));
      props.setProperty("KOMMO_ERRO", "");
      props.deleteProperty("KOMMO_PUXANDO");
      const salvo = lerSalvo_();
      calcularESalvar_(salvo ? (salvo.out.desde || "") : "");
      Logger.log("Kommo OK — " + resumo.join(" | ") + " | total " + Math.round((Date.now() - job.inicio) / 1000) + "s");
      return { ok: true, resumo: resumo };
    }

    // ainda não acabou: continua sozinho em 1 minuto
    agendarContinuacao_();
    Logger.log("Kommo em andamento: " + resumoJob_(job) + " — continua sozinho em 1 min.");
    return { ok: true, continua: true, job: resumoJob_(job) };
  } catch (e) {
    props.setProperty("KOMMO_ERRO", String(e.message || e));
    props.deleteProperty("KOMMO_PUXANDO");
    Logger.log("Erro ao puxar o Kommo: " + e);
    return { ok: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}
function novoJob_(cfg) {
  const pipes = kommoTudo_(cfg, "leads/pipelines", "pipelines");
  const etapa = {}, funil = {}, usuarios = {};
  pipes.forEach(p => { funil[p.id] = p.name; ((p._embedded || {}).statuses || []).forEach(s => etapa[p.id + ":" + s.id] = s.name); });
  kommoTudo_(cfg, "users", "users").forEach(u => usuarios[u.id] = u.name);
  const sb = basesSS_();
  const funis = Object.keys(KOMMO_FUNIS).map(nomeFunil => {
    const p = pipes.find(x => String(x.name).trim().toLowerCase() === nomeFunil.toLowerCase());
    if (!p) throw new Error("Funil não encontrado no Kommo: " + nomeFunil);
    // pula a etapa DUPLICADOS
    // pula DUPLICADOS e os leads fechados (142 = venda ganha, 143 = venda perdida)
    const etapas = ((p._embedded || {}).statuses || []).filter(st => !/duplicad/i.test(st.name) && st.id !== 142 && st.id !== 143);
    const filtro = etapas.map((st, i) => "filter[statuses][" + i + "][pipeline_id]=" + p.id + "&filter[statuses][" + i + "][status_id]=" + st.id).join("&");
    const aba = KOMMO_FUNIS[nomeFunil];
    const tmp = sb.getSheetByName("_tmp " + aba); if (tmp) sb.deleteSheet(tmp);
    sb.insertSheet("_tmp " + aba).hideSheet();
    return { aba: aba, filtro: filtro, pag: 1, n: 0, pipeId: p.id };
  });
  const fechadas = [];
  pipes.forEach(p => ((p._embedded || {}).statuses || []).forEach(st => { if (st.id === 142 || st.id === 143) fechadas.push(String(st.name)); }));
  const job = { campos: {}, inicio: Date.now(), fase: "leads", i: 0, funis: funis, etapa: etapa, funil: funil, usuarios: usuarios, fechadas: fechadas };
  salvarJob_(job);
  return job;
}
function linhaLead_(l, job, tz) {
  const fmtTs = ts => ts ? Utilities.formatDate(new Date(ts * 1000), tz, "dd.MM.yyyy HH:mm:ss") : "";
  const cf = {};
  (l.custom_fields_values || []).forEach(f => {
    // se dois campos tiverem o mesmo nome, o vazio não pode apagar o preenchido
    const k = normNome_(f.field_name), val = fmtCampo_(f, tz);
    if (job.campos) job.campos[String(f.field_name).trim()] = 1;
    if (val !== "" || !(k in cf)) cf[k] = val;
  });
  const c = ((l._embedded || {}).contacts || []), m = c.find(x => x.is_main) || c[0];
  const base = {
    "ID": l.id, "Contato principal": "", "Etapa do lead": job.etapa[l.pipeline_id + ":" + l.status_id] || "",
    "Funil de vendas": job.funil[l.pipeline_id] || "", "Lead usuário responsável": job.usuarios[l.responsible_user_id] || "",
    "Data Criada": fmtTs(l.created_at), "Última modificação": fmtTs(l.updated_at), "Modificado por": job.usuarios[l.updated_by] || "",
    "Venda": l.price || 0
  };
  return COLS_EXPORT.map(col => String(col in base ? base[col] : (cf[normNome_(col)] || ""))).concat([m ? String(m.id) : ""]);
}
function abaTemp_(sb, aba) { return sb.getSheetByName("_tmp " + aba) || sb.insertSheet("_tmp " + aba); }
function arquivoJob_() {
  const nome = "kommo-andamento.json", pasta = pastaFotos_(), it = pasta.getFilesByName(nome);
  return it.hasNext() ? it.next() : pasta.createFile(nome, "", "application/json");
}
function lerJob_() { try { const t = arquivoJob_().getBlob().getDataAsString(); return t ? JSON.parse(t) : null; } catch (e) { return null; } }
function salvarJob_(job) { arquivoJob_().setContent(JSON.stringify(job)); }
function apagarJob_() { arquivoJob_().setContent(""); }
function agendarContinuacao_() { ScriptApp.newTrigger("continuarKommo").timeBased().after(60 * 1000).create(); }
function resumoJob_(job) {
  const fase = { leads: "buscando leads", nomes: "buscando nomes dos contatos", finalizar: "salvando" }[job.fase] || job.fase;
  return fase + " (" + job.funis.map(f => f.aba.replace("Base ", "") + ": " + f.n.toLocaleString("pt-BR")).join(", ") + ")";
}
function arquivoNomes_() {
  const nome = "nomes-contatos-kommo.json", pasta = pastaFotos_();
  const it = pasta.getFilesByName(nome);
  return it.hasNext() ? it.next() : pasta.createFile(nome, "{}", "application/json");
}
function lerNomesCache_() { try { return JSON.parse(arquivoNomes_().getBlob().getDataAsString()) || {}; } catch (e) { return {}; } }
function salvarNomesCache_(m) { try { arquivoNomes_().setContent(JSON.stringify(m)); } catch (e) {} }
function pastaFotos_() {
  const it = DriveApp.getFoldersByName(PASTA_FOTOS);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_FOTOS);
}
function statusKommo_() {
  const pr = PropertiesService.getScriptProperties();
  const job = lerJob_();
  return { ok: true, configurado: !!(pr.getProperty("KOMMO_SUBDOMINIO") && pr.getProperty("KOMMO_TOKEN")),
    puxando: !!job && Date.now() - job.inicio < 30 * 60 * 1000, andamento: job ? resumoJob_(job) : "",
    ultima: pr.getProperty("KOMMO_ULTIMA") || "", erro: pr.getProperty("KOMMO_ERRO") || "",
    conferencia: (function () { try { return JSON.parse(pr.getProperty("KOMMO_CONF") || "null"); } catch (e) { return null; } })() };
}

// Rode 1x pelo editor depois de preencher as Propriedades do script:
// cria os gatilhos (início do dia e antes de cada validação) e faz a primeira puxada.
function configurarKommo() {
  ScriptApp.getProjectTriggers().filter(t => ["puxarBasesKommo", "continuarKommo"].includes(t.getHandlerFunction())).forEach(t => ScriptApp.deleteTrigger(t));
  apagarJob_();
  [[7, 0], [12, 30], [19, 45]].forEach(([h, m]) => ScriptApp.newTrigger("puxarBasesKommo").timeBased().everyDays(1).atHour(h).nearMinute(m).create());
  const r = puxarBasesKommo();
  Logger.log(JSON.stringify(r));
}
