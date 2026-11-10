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
}
