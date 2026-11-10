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
  if (action === "validar") return validar_(p);

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
const COLS_BASE = ["ID", "Contato principal", "Etapa do lead", "Data Cadastro", "Data Renovação", "Plano", "Telas", "Servidor", "Vendedor", "Data de Vencimento"];

function validar_(p) {
  const tz = Session.getScriptTimeZone();
  const hoje = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const cache = CacheService.getScriptCache();
  const chave = "validar-" + hoje + "-" + (p.desde || "") + "-" + versaoBases_();
  const emCache = cache.get(chave);
  if (emCache && !p.forcar) return JSON.parse(emCache);
  const ss = basesSS_();
  const bases = BASES.map(nome => {
    const sh = ss.getSheetByName(nome);
    if (!sh || sh.getLastRow() < 2) return { nome: nome, head: [], cols: {} };
    const lastCol = sh.getLastColumn(), n = sh.getLastRow() - 1;
    const head = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h).trim());
    const cols = {};
    COLS_BASE.forEach(c => {
      const i = head.indexOf(c);
      if (i >= 0) cols[c] = sh.getRange(2, i + 1, n, 1).getDisplayValues().map(r => r[0]);
    });
    return { nome: nome, head: head, cols: cols, n: n };
  });
  const out = calcularValidacao(bases, hoje, p.desde || "");
  out.ok = true; out.aba = "Suporte"; out.basesUrl = ss.getUrl();
  try { cache.put(chave, JSON.stringify(out), 300); } catch (e) {}
  return out;
}
// muda quando o conteúdo das bases muda (nº de linhas + última célula)
function versaoBases_() {
  const ss = basesSS_();
  return BASES.map(n => { const sh = ss.getSheetByName(n); return sh ? sh.getLastRow() + "x" + sh.getLastColumn() : "0"; }).join("-");
}

// ---- regra (pura, sem dependência do Google) ----
function calcularValidacao(bases, hoje, desde) {
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

  bases.forEach(b => {
    const faltando = COLS_BASE.filter(c => !b.cols[c]);
    infoBases.push({ nome: b.nome, linhas: b.n || 0, vazia: !b.n, faltando: b.n ? faltando : [] });
    if (!b.n || faltando.length) return;
    consumo.porBase[b.nome] = 0;
    consumo.detalheBase = consumo.detalheBase || {};
    const det = consumo.detalheBase[b.nome] = { vendas: 0, renovacoes: 0, clientesVendas: 0, clientesRenovacoes: 0 };
    for (let r = 0; r < b.n; r++) {
      const v = c => String(b.cols[c][r] || "").trim();
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
  return { mes: mesAtual, hoje: hoje, desde: desde || "", bases: infoBases, consumo: consumo, consumoDesde: consumoDesde, erros: erros };
}
