// Ponte do site do SUPORTE com o Apps Script. Só acessa a aba "Suporte" (sem acesso aos custos).
const API_URL = "https://script.google.com/macros/s/AKfycby4poOb53Ar1FQ9IGSeeMScNtIkbFwJjhDqIWfgPqjZz-8alFa7NNWt9Exg9jnK-AGtCQ/exec";
const ACOES = ["list", "save", "delete", "validar", "puxarKommo", "kommoStatus"];
const TIPOS_GRAVAVEIS = ["hora", "check", "valid"];

const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), { status, headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" } });

export default async (req) => {
  if (req.method !== "POST") return json({ ok: false, error: "use POST" }, 405);
  let body;
  try { body = JSON.parse(await req.text()); } catch { return json({ ok: false, error: "requisição inválida" }, 400); }
  if (!ACOES.includes(body.action)) return json({ ok: false, error: "ação não permitida" }, 403);
  body.aba = "Suporte"; // sempre a aba do suporte
  if (body.action === "save") {
    const it = body.item || {};
    if (!TIPOS_GRAVAVEIS.includes(it.tipo) || !/^(r|ck-|v-)/.test(String(it.id || ""))) return json({ ok: false, error: "não permitido" }, 403);
  }
  if (body.action === "delete" && (String(body.id || "").startsWith("__") || !/^(r|ck-|v-)/.test(String(body.id || "")))) return json({ ok: false, error: "não permitido" }, 403);
  try {
    const r = await fetch(API_URL, { method: "POST", headers: { "content-type": "text/plain;charset=utf-8" }, body: JSON.stringify(body), redirect: "follow" });
    const txt = await r.text();
    let out;
    try { out = JSON.parse(txt); } catch {
      return json({ ok: false, error: "Apps Script respondeu algo inesperado (HTTP " + r.status + "): " + txt.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().slice(0, 200) });
    }
    if (out.ok && out.aba !== "Suporte" && !["delete", "puxarKommo", "kommoStatus"].includes(body.action)) return json({ ok: false, error: "Apps Script desatualizado: falta o código novo (Code.gs)." });
    return json(out);
  } catch (e) {
    return json({ ok: false, error: "Falha ao falar com o Apps Script: " + e.message });
  }
};

export const config = { path: "/api" };
