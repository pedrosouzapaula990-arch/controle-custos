// Ponte entre o site e o Google Apps Script (evita bloqueios do navegador, ex.: iPhone/Safari)
const API_URL = "https://script.google.com/macros/s/AKfycbzdVqdYAElufIOZppKk7nyDx8qczvp8rm7vbqww6yp3BsC_ubpK58oGalvl67h1E1EnpQ/exec";

const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" },
  });

export default async (req) => {
  if (req.method !== "POST") return json({ ok: false, error: "use POST" }, 405);
  const body = await req.text();
  try {
    const r = await fetch(API_URL, {
      method: "POST",
      headers: { "content-type": "text/plain;charset=utf-8" },
      body,
      redirect: "follow",
    });
    const txt = await r.text();
    try {
      return json(JSON.parse(txt));
    } catch {
      const msg = txt.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().slice(0, 200);
      return json({ ok: false, error: "Apps Script respondeu algo inesperado (HTTP " + r.status + "): " + msg });
    }
  } catch (e) {
    return json({ ok: false, error: "Falha ao falar com o Apps Script: " + e.message });
  }
};

export const config = { path: "/api" };
