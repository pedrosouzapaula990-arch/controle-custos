// Extração do Meta Ads (tráfego pago): gasto, cliques, conversas e assinaturas por dia e por campanha.
// Variável no Netlify: META_TOKEN (nunca vai para o código). Grava na aba "Meta Ads" via Apps Script.
const APPS = "https://script.google.com/macros/s/AKfycby4poOb53Ar1FQ9IGSeeMScNtIkbFwJjhDqIWfgPqjZz-8alFa7NNWt9Exg9jnK-AGtCQ/exec";
const CONTA = "act_1583593523071280";
const V = "v21.0";
const TZ = "America/Sao_Paulo";
const DIAS = 45;
const CONVERSA = ["onsite_conversion.messaging_conversation_started_7d"];
const ASSINA = ["omni_subscribe", "subscribe_total", "offsite_conversion.fb_pixel_subscribe", "subscribe_website", "subscribe_mobile_app", "onsite_web_subscribe"];
const json = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" } });
const dia = d => new Intl.DateTimeFormat("en-CA", { timeZone: TZ }).format(d);
const pega = (lista, tipos) => { for (const t of tipos) { const a = (lista || []).find(x => x.action_type === t); if (a) return +a.value || 0; } return 0; };

export default async (req) => {
  const token = String(process.env.META_TOKEN || "").trim();
  if (!token) return json({ ok: false, error: "META_TOKEN não configurado no Netlify" }, 500);
  const url0 = new URL(req.url);
  const dias = Math.min(400, +url0.searchParams.get("dias") || DIAS);
  const hoje = new Date(), ini = new Date(hoje.getTime() - dias * 86400000);
  const desde = dia(ini), ate = dia(hoje);
  const campos = "date_start,campaign_id,campaign_name,spend,impressions,clicks,ctr,cpc,actions";
  let u = `https://graph.facebook.com/${V}/${CONTA}/insights?level=campaign&time_increment=1&limit=500&fields=${campos}&time_range=${encodeURIComponent(JSON.stringify({ since: desde, until: ate }))}&access_token=${encodeURIComponent(token)}`;
  const linhas = [], tipos = {};
  try {
    for (let i = 0; i < 40 && u; i++) {
      const r = await fetch(u); const j = await r.json();
      if (j.error) return json({ ok: false, error: "Meta: " + (j.error.message || JSON.stringify(j.error)) }, 502);
      for (const x of j.data || []) {
        (x.actions || []).forEach(a => { tipos[a.action_type] = (tipos[a.action_type] || 0) + (+a.value || 0); });
        linhas.push({ data: x.date_start, campanhaId: String(x.campaign_id), campanha: x.campaign_name, gasto: +x.spend || 0, impressoes: +x.impressions || 0, cliques: +x.clicks || 0, ctr: +x.ctr || 0, cpc: +x.cpc || 0, conversas: pega(x.actions, CONVERSA), assinaturas: pega(x.actions, ASSINA), acoes: JSON.stringify(x.actions || []).slice(0, 1500) });
      }
      u = j.paging && j.paging.next || null;
    }
    const r = await fetch(APPS, { method: "POST", headers: { "content-type": "text/plain;charset=utf-8" }, body: JSON.stringify({ action: "importarMeta", segredo: process.env.IMPORT_SECRET, desde, linhas }), redirect: "follow" });
    const j = JSON.parse(await r.text());
    if (!j.ok) return json({ ok: false, error: j.error || "Apps Script recusou" }, 502);
    return json({ ok: true, desde, ate, linhas: linhas.length, gravadas: j.gravadas, tiposDeAcao: tipos });
  } catch (e) { return json({ ok: false, error: "Falha: " + e.message }, 500); }
};
export const config = { path: "/meta-extrair" };
