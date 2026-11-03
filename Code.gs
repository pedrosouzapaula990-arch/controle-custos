const ABA = "Custos";
const COLS = ["id","data","descricao","categoria","valor","pagoPor","status","obs","criadoEm","mesRef","comprovante"];
const PLANILHA_ID = "1QWnNvsYs8kQDJXolMl6dTxsYGS-sCMvS2QJektqFBGw";
const PASTA_COMPROVANTES = "Comprovantes - Controle de Custos";

function aba_() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  let sh = ss.getSheetByName(ABA);
  if (!sh) sh = ss.insertSheet(ABA);
  const head = sh.getRange(1, 1, 1, COLS.length).getValues()[0];
  if (head.join("|") !== COLS.join("|")) {
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]);
    sh.setFrozenRows(1);
    // colunas de texto (evita a planilha converter datas)
    ["A:D", "F:H", "J:K"].forEach(r => sh.getRange(r).setNumberFormat("@"));
  }
  return sh;
}

function ehData_(v) {
  return v && typeof v.getTime === "function";
}

function listar_() {
  const sh = aba_();
  const vals = sh.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  return vals.slice(1).filter(r => r[0]).map(r => {
    const o = {};
    COLS.forEach((c, i) => o[c] = r[i] === undefined ? "" : r[i]);
    if (ehData_(o.data)) o.data = Utilities.formatDate(o.data, tz, "yyyy-MM-dd");
    if (ehData_(o.mesRef)) o.mesRef = Utilities.formatDate(o.mesRef, tz, "yyyy-MM");
    o.valor = Number(o.valor) || 0;
    return o;
  });
}

function pasta_() {
  const it = DriveApp.getFoldersByName(PASTA_COMPROVANTES);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_COMPROVANTES);
}

function executar_(action, p) {
  if (action === "list") return { ok: true, items: listar_() };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = aba_();
    const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => String(r[0]));

    if (action === "save") {
      const it = typeof p.item === "string" ? JSON.parse(p.item) : p.item;
      const row = COLS.map(c => it[c] === undefined || it[c] === null ? "" : it[c]);
      const idx = ids.indexOf(String(it.id));
      const linha = idx > 0 ? idx + 1 : sh.getLastRow() + 1;
      sh.getRange(linha, 1, 1, COLS.length).setValues([row]);
      return { ok: true };
    }

    if (action === "delete") {
      const idx = ids.indexOf(String(p.id));
      if (idx > 0) sh.deleteRow(idx + 1);
      return { ok: true };
    }

    if (action === "upload") {
      const idx = ids.indexOf(String(p.id));
      if (idx < 1) return { ok: false, error: "lançamento não encontrado" };
      const blob = Utilities.newBlob(Utilities.base64Decode(p.base64), p.mime || "application/octet-stream", p.nome || "comprovante");
      const arq = pasta_().createFile(blob);
      arq.setName(p.id + " - " + (p.nome || "comprovante"));
      try { arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
      const url = arq.getUrl();
      sh.getRange(idx + 1, COLS.indexOf("comprovante") + 1).setValue(url);
      sh.getRange(idx + 1, COLS.indexOf("status") + 1).setValue("pago");
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
  const p = e.parameter || {};
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

// Rode esta função 1x pelo editor (Run) para autorizar o acesso ao Drive.
function autorizar() {
  pasta_();
  aba_();
}
