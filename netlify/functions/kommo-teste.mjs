// Checagem rápida (só leitura): o Netlify alcança o Kommo e o Apps Script? Que campos o Kommo entrega de um lead?
// Protegida por ?s=<início da chave IMPORT_SECRET>. Não devolve dados pessoais (só nomes de campos e valores de campos de cadastro).
const APPS = "https://script.google.com/macros/s/AKfycby4poOb53Ar1FQ9IGSeeMScNtIkbFwJjhDqIWfgPqjZz-8alFa7NNWt9Exg9jnK-AGtCQ/exec";
export default async (req) => {
  const u = new URL(req.url), seg = process.env.IMPORT_SECRET || "";
  if (!seg || u.searchParams.get("s") !== seg.slice(0, 8)) return new Response("não autorizado", { status: 403 });
  const out = {};
  const sub = String(process.env.KOMMO_SUBDOMINIO || "").trim(), token = String(process.env.KOMMO_TOKEN || "").trim();
  const lead = u.searchParams.get("lead") || "110674836";
  try {
    const r = await fetch(`https://${sub}.kommo.com/api/v4/leads/${lead}`, { headers: { Authorization: "Bearer " + token } });
    out.kommoStatus = r.status;
    if (r.ok) {
      const j = await r.json();
      out.lead = { id: j.id, pipeline: j.pipeline_id, status: j.status_id, campos: (j.custom_fields_values || []).map(f => ({ nome: f.field_name, tipo: f.field_type, qtd: (f.values || []).length, valor: /tela|venc|renov|plano|cadastro|servidor|vendedor/i.test(f.field_name) ? (f.values || []).map(v => v.value) : undefined })) };
    }
  } catch (e) { out.kommoErro = e.message + " / " + ((e.cause && (e.cause.code || e.cause.message)) || ""); }
  try {
    const r = await fetch(APPS, { method: "POST", headers: { "content-type": "text/plain;charset=utf-8" }, body: JSON.stringify({ action: "kommoStatus" }), redirect: "follow" });
    const t = await r.text();
    out.appsStatus = r.status;
    try { const j = JSON.parse(t); out.apps = { puxando: j.puxando, andamento: j.andamento, ultima: j.ultima, erro: j.erro, conferencia: j.conferencia && { quando: j.conferencia.quando, completa: j.conferencia.completa, colunasVazias: j.conferencia.colunasVazias } }; } catch { out.appsTexto = t.slice(0, 200); }
  } catch (e) { out.appsErro = e.message + " / " + ((e.cause && (e.cause.code || e.cause.message)) || ""); }
  return new Response(JSON.stringify(out, null, 1), { headers: { "content-type": "application/json" } });
};
export const config = { path: "/kommo-teste" };
