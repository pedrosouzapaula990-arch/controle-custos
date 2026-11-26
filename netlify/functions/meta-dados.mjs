// Entrega ao Dashboard os dados do Meta Ads guardados pela extração (rápido, sem passar pelo Google).
import { getStore } from "@netlify/blobs";
export default async () => {
  try {
    const store = getStore("meta");
    const items = (await store.get("linhas", { type: "json" })) || [];
    return new Response(JSON.stringify({ ok: true, items }), { headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" } });
  } catch (e) { return new Response(JSON.stringify({ ok: false, error: e.message }), { status: 500, headers: { "content-type": "application/json" } }); }
};
export const config = { path: "/meta-dados" };
