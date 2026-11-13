// Dispara a extração do Kommo 3x por dia (7h, 12h e 19h, horário de Brasília). A extração em si roda na função "-background" (até 15 min).
export default async () => {
  const base = process.env.URL || "https://controlegastos202.netlify.app";
  try { await fetch(base + "/.netlify/functions/kommo-extrair-background", { method: "POST" }); } catch (e) { console.log("Falha ao disparar:", e.message); }
};
export const config = { schedule: "0 10,15,22 * * *" };
