// Extração do Kommo (roda no Netlify, sem o limite diário do Google).
// Busca as duas bases (funis DarbplayOficial e Suporte), sem DUPLICADOS e sem etapas fechadas (venda ganha/perdida),
// e envia em partes ao Apps Script, que grava na planilha, guarda a foto no Drive e recalcula.
// Variáveis no Netlify: KOMMO_SUBDOMINIO e KOMMO_TOKEN (o token nunca vai para o código).
const APPS = "https://script.google.com/macros/s/AKfycby4poOb53Ar1FQ9IGSeeMScNtIkbFwJjhDqIWfgPqjZz-8alFa7NNWt9Exg9jnK-AGtCQ/exec";
const FUNIS = { DarbplayOficial: "Base Darb Oficial", Suporte: "Base Motiva" };
const COLS = ["ID", "Contato principal", "Etapa do lead", "Funil de vendas", "Lead usuário responsável",
  "Data Criada", "Última modificação", "Modificado por", "Data Cadastro", "Plano", "Telas", "Servidor",
  "Aplicativo Instalado", "Vendedor", "Data de Vencimento", "Última Data Vencimento", "Data Renovação",
  "Renovação", "Fila de Renovação", "Tabulação", "Venda"];
const TZ = "America/Sao_Paulo";
const dorme = ms => new Promise(r => setTimeout(r, ms));
const norm = t => String(t || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/\s+/g, " ").trim();

const sub = String(process.env.KOMMO_SUBDOMINIO || "").trim().replace(/^https?:\/\//, "").replace(/\.kommo\.com.*$/, "");
const token = String(process.env.KOMMO_TOKEN || "").trim();
const BASE = `https://${sub}.kommo.com/api/v4/`;

async function kommo(path, tent = 0) {
  let r;
  try { r = await fetch(BASE + path, { headers: { Authorization: "Bearer " + token } }); }
  catch (e) {
    const causa = (e.cause && (e.cause.code || e.cause.message)) || "";
    if (tent < 6) { await dorme(2000 * (tent + 1)); return kommo(path, tent + 1); }
    throw new Error("Falha de rede ao chamar o Kommo (" + path.split("?")[0] + "): " + e.message + (causa ? " / " + causa : ""));
  }
  if (r.status === 204) return null;
  if (r.status === 429 && tent < 6) { await dorme(2000 * (tent + 1)); return kommo(path, tent + 1); }
  if (r.status === 401) throw new Error("Token do Kommo inválido ou expirado (401).");
  if (!r.ok) throw new Error("Kommo respondeu " + r.status + ": " + (await r.text()).slice(0, 200));
  return r.json();
}
async function todas(path, chave) {
  const out = [];
  for (let pag = 1; ; pag++) {
    const j = await kommo(path + (path.includes("?") ? "&" : "?") + "limit=250&page=" + pag);
    const it = (j && j._embedded && j._embedded[chave]) || [];
    out.push(...it);
    if (it.length < 250) break;
    await dorme(220);
  }
  return out;
}

async function apps(corpo, tent = 0) {
  try {
    const r = await fetch(APPS, { method: "POST", headers: { "content-type": "text/plain;charset=utf-8" }, body: JSON.stringify({ action: "importar", segredo: process.env.IMPORT_SECRET, ...corpo }), redirect: "follow" });
    const t = await r.text();
    const j = JSON.parse(t);
    if (!j.ok) throw new Error(j.error || "Apps Script recusou");
    return j;
  } catch (e) {
    if (tent < 4 && !/recente|Chave/.test(e.message)) { await dorme(3000 * (tent + 1)); return apps(corpo, tent + 1); }
    throw new Error("Apps Script (" + corpo.fase + "): " + e.message);
  }
}

const fmtData = (ts, hora) => {
  const d = new Date(ts * 1000);
  const p = Object.fromEntries(new Intl.DateTimeFormat("en-GB", { timeZone: TZ, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false }).formatToParts(d).map(x => [x.type, x.value]));
  return `${p.day}.${p.month}.${p.year}` + (hora ? ` ${p.hour === "24" ? "00" : p.hour}:${p.minute}` + (hora === 2 ? `:${p.second}` : "") : "");
};
function fmtCampo(f) {
  const vals = (f.values || []).map(v => v.value).filter(v => v !== null && v !== undefined && v !== "");
  if (!vals.length) return "";
  if (f.field_type === "date" || f.field_type === "birthday") return vals.map(v => fmtData(+v, 0)).join(", ");
  if (f.field_type === "date_time") return vals.map(v => fmtData(+v, 1)).join(", ");
  if (f.field_type === "checkbox") return vals[0] ? "Sim" : "";
  return vals.join(", ");
}

export default async () => {
  if (!sub || !token || !process.env.IMPORT_SECRET) { console.log("Falta KOMMO_SUBDOMINIO / KOMMO_TOKEN / IMPORT_SECRET no Netlify."); return; }
  try { await apps({ fase: "inicio" }); } catch (e) { console.log("Não iniciou:", e.message); return; }
  try {
    const pipes = await todas("leads/pipelines", "pipelines");
    const etapa = {}, funil = {}, usuarios = {};
    pipes.forEach(p => { funil[p.id] = p.name; ((p._embedded || {}).statuses || []).forEach(s => (etapa[p.id + ":" + s.id] = s.name)); });
    (await todas("users", "users")).forEach(u => (usuarios[u.id] = u.name));
    const campos = {}, contagens = {}, vendaGanha = [], contatoDe = {};
    const linhasPorAba = {};

    for (const [nomeFunil, aba] of Object.entries(FUNIS)) {
      const p = pipes.find(x => String(x.name).trim().toLowerCase() === nomeFunil.toLowerCase());
      if (!p) throw new Error("Funil não encontrado no Kommo: " + nomeFunil);
      const sts = ((p._embedded || {}).statuses || []).filter(s => !/duplicad/i.test(s.name) && s.id !== 142 && s.id !== 143);
      const filtro = sts.map((s, i) => `filter[statuses][${i}][pipeline_id]=${p.id}&filter[statuses][${i}][status_id]=${s.id}`).join("&");
      await apps({ fase: "andamento", texto: "Buscando " + aba + "…" });
      const leads = await todas("leads?with=contacts&" + filtro, "leads");
      const rows = [], vistos = new Set();
      for (const l of leads) {
        if (l.id === 110674836) console.log("DIAG lead 110674836:", JSON.stringify((l.custom_fields_values || []).map(f => ({ n: f.field_name, id: f.field_id, t: f.field_type, v: f.values }))));
        if (vistos.has(l.id)) continue;
        vistos.add(l.id);
        const cf = {};
        (l.custom_fields_values || []).forEach(f => {
          const k = norm(f.field_name), v = fmtCampo(f);
          campos[String(f.field_name).trim()] = 1;
          if (v !== "" || !(k in cf)) cf[k] = v;
        });
        const c = (l._embedded || {}).contacts || [], m = c.find(x => x.is_main) || c[0];
        const base = {
          "ID": l.id, "Contato principal": "", "Etapa do lead": etapa[l.pipeline_id + ":" + l.status_id] || "",
          "Funil de vendas": funil[l.pipeline_id] || "", "Lead usuário responsável": usuarios[l.responsible_user_id] || "",
          "Data Criada": l.created_at ? fmtData(l.created_at, 2) : "", "Última modificação": l.updated_at ? fmtData(l.updated_at, 2) : "",
          "Modificado por": usuarios[l.updated_by] || "", "Venda": l.price || 0,
        };
        const linha = COLS.map(col => String(col in base ? base[col] : (cf[norm(col)] || "")));
        if (m) contatoDe[l.id] = m.id;
        rows.push({ linha, contato: m ? m.id : null });
      }
      linhasPorAba[aba] = rows;
      console.log("DIAG preenchimento " + aba + " (" + rows.length + "):", COLS.map((c, i) => c + "=" + rows.filter(r => String(r.linha[i]).trim() !== "").length).join(", "));
      contagens[aba] = rows.length;
      // venda ganha não deve ter clientes: guarda os que aparecerem (1ª página)
      const vg = await kommo(`leads?limit=250&filter[statuses][0][pipeline_id]=${p.id}&filter[statuses][0][status_id]=142`);
      ((vg && vg._embedded && vg._embedded.leads) || []).forEach(l => vendaGanha.push({ base: aba, id: String(l.id), nome: String(l.name || ""), resp: usuarios[l.responsible_user_id] || "" }));
      await dorme(220);
    }

    // nomes dos contatos
    const ids = [...new Set(Object.values(linhasPorAba).flat().map(r => r.contato).filter(Boolean))];
    const nomes = {};
    for (let i = 0; i < ids.length; i += 80) {
      if (i % 800 === 0) await apps({ fase: "andamento", texto: `Nomes dos contatos ${i}/${ids.length}…` });
      const j = await kommo("contacts?limit=250&" + ids.slice(i, i + 80).map(id => "filter[id][]=" + id).join("&"));
      ((j && j._embedded && j._embedded.contacts) || []).forEach(c => (nomes[c.id] = c.name));
      await dorme(220);
    }
    const iNome = COLS.indexOf("Contato principal");

    // envio em partes
    for (const [aba, rows] of Object.entries(linhasPorAba)) {
      const linhas = rows.map(r => { const x = r.linha.slice(); x[iNome] = (r.contato && nomes[r.contato]) || ""; return x; });
      for (let i = 0; i < linhas.length; i += 800) {
        await apps({ fase: "parte", aba, linhas: linhas.slice(i, i + 800), texto: `Enviando ${aba} ${Math.min(i + 800, linhas.length)}/${linhas.length}…` });
      }
    }
    const r = await apps({ fase: "fim", contagens, vendaGanha, campos: Object.keys(campos).slice(0, 80) });
    console.log("Kommo OK", JSON.stringify(contagens), r.resumo || "");
  } catch (e) {
    console.log("Erro na extração:", e.message);
    try { await apps({ fase: "erro", mensagem: e.message }); } catch {}
  }
};
