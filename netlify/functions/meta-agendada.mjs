// Puxa o Meta Ads 3x por dia (8h, 14h e 20h, horário de Brasília).
export default async () => {
  const base = process.env.URL || "https://controlegastos202.netlify.app";
  try { const r = await fetch(base + "/meta-extrair"); console.log("Meta:", r.status, (await r.text()).slice(0, 300)); } catch (e) { console.log("Falha Meta:", e.message); }
};
export const config = { schedule: "0 11,17,23 * * *" };
