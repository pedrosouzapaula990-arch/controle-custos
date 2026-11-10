// Regras compartilhadas do suporte (usado pelo site do suporte e pela aba Suporte do site de custos)
const SupCore = (() => {
  const DEF = {
    disparos: ["07:00 – Vence hoje", "10:15 – Vence amanhã", "11:30 – Venceu ontem", "17:40 – Vence hoje", "19:30 – Venceu ontem"],
    resp: ["Turno manhã (9h–16:40)", "Turno noite (16:40–23:30)"],
    tarefas: [
      "12:50 – Validação diária (base, vendas e indicações)",
      "13:30 – Acompanhar solicitações de conteúdo",
      "20:05 – Validação diária (base, vendas e indicações)",
      "21:00 – Acompanhar solicitações de conteúdo"
    ]
  };
  const TURNOS = [
    { ini: 9 * 60, fim: 16 * 60 + 40, chave: "manh", nome: "Turno manhã", faixa: "9h–16:40" },
    { ini: 16 * 60 + 40, fim: 23 * 60 + 30, chave: "noite", nome: "Turno noite", faixa: "16:40–23:30" }
  ];
  const ST = { aguardando: { t: "Aguardando", e: "⏳" }, andamento: { t: "Em andamento", e: "🔄" }, concluido: { t: "Concluído", e: "✅" }, problema: { t: "Com problema", e: "⚠️" } };

  const esc = s => { const d = document.createElement("div"); d.textContent = s ?? ""; return d.innerHTML; };
  const pad = n => String(n).padStart(2, "0");
  const hojeStr = () => { const t = new Date(); return t.getFullYear() + "-" + pad(t.getMonth() + 1) + "-" + pad(t.getDate()); };
  const minAgora = () => { const a = new Date(); return a.getHours() * 60 + a.getMinutes(); };
  const hm2min = h => { const [a, b] = String(h).split(":").map(Number); return a * 60 + (b || 0); };
  const min2hm = m => pad(Math.floor(m / 60)) + ":" + pad(m % 60);
  const turnoDe = m => TURNOS.find(t => m >= t.ini && m < t.fim) || null;
  const horaDe = n => { const m = String(n).match(/^(\d{1,2})(?:[:h](\d{2}))?/); return m ? (+m[1]) * 60 + (+(m[2] || 0)) : null; };
  const semHora = n => String(n).replace(/^\d{1,2}(?:[:h]\d{2})?h?\s*[–-]\s*/, "");
  const curto = n => String(n).split(/\s+[–-]\s+/)[0];
  const porHora = (a, b) => (horaDe(a) ?? 9999) - (horaDe(b) ?? 9999);

  const disparos = cfg => [...(cfg?.disparos || DEF.disparos)].sort(porHora);
  const resp = cfg => cfg?.resp || DEF.resp;
  const tarefas = cfg => [...(cfg?.tarefas || DEF.tarefas)].sort(porHora);
  const parseDados = r => { try { const d = JSON.parse(r.dados || "[]"); return Array.isArray(d) ? d : []; } catch (e) { return []; } };

  // tarefas do dia = conferir cada disparo + tarefas fixas
  function tarefasDoDia(cfg) {
    const out = [];
    disparos(cfg).forEach(d => {
      let m = horaDe(d); if (m == null) return;
      if (m < TURNOS[0].ini) m = TURNOS[0].ini; // disparo antes do turno: confere no início
      out.push({ min: m, hora: min2hm(m), texto: "Conferir disparo " + d });
    });
    tarefas(cfg).forEach(t => {
      const m = horaDe(t); if (m == null) return;
      out.push({ min: m, hora: min2hm(m), texto: semHora(t) || t });
    });
    out.forEach(t => { const tu = turnoDe(t.min) || (t.min < TURNOS[0].ini ? TURNOS[0] : TURNOS[1]); t.turno = tu.nome; t.tk = tu.chave; });
    return out.sort((a, b) => a.min - b.min);
  }
  function ckId(dia, texto) {
    const slug = texto.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 60);
    return "ck-" + dia + "-" + slug;
  }
  function statusTarefa(t, feito, dia) {
    if (feito) return { cls: "concluido", txt: "✅ Feita às " + feito.hora };
    const hj = hojeStr();
    if (dia < hj) return { cls: "problema", txt: "❌ Não feita" };
    if (dia === hj && minAgora() > t.min + 30) return { cls: "problema", txt: "⚠️ Atrasada" };
    return { cls: "", txt: "⏳ Pendente" };
  }
  // HTML dos dois cards de turno
  function htmlChecklist(cfg, checks, dia, interativo, reportes) {
    const lista = tarefasDoDia(cfg);
    return TURNOS.map(tu => {
      const ts = lista.filter(t => t.tk === tu.chave);
      const rows = ts.map(t => {
        const f = checks.find(c => c.id === ckId(dia, t.texto) && c.data === dia);
        const st = statusTarefa(t, f, dia);
        const btn = interativo
          ? `<button type="button" class="ckbtn ${f ? "on" : ""}" data-ck="${encodeURIComponent(JSON.stringify({ texto: t.texto, hora: t.hora, turno: t.turno }))}" aria-label="${f ? "Desmarcar" : "Marcar como feita"}">${f ? "✓" : ""}</button>`
          : `<span class="ckbtn ${f ? "on" : ""}">${f ? "✓" : ""}</span>`;
        const obs = f && f.obs ? `<div class="hint">📝 ${esc(f.obs)}</div>` : "";
        return `<div class="ckrow">${btn}<div class="ckh">${t.hora}</div><div class="ckt">${esc(t.texto)}${obs}</div><span class="chip ${st.cls}">${st.txt}</span></div>`;
      }).join("") || '<div class="empty">Nenhuma tarefa neste turno.</div>';
      const feitas = ts.filter(t => checks.some(c => c.id === ckId(dia, t.texto) && c.data === dia)).length;
      const pct = ts.length ? Math.round(feitas / ts.length * 100) : 0;
      const nRep = reportes ? reportes.filter(r => r.data === dia && hm2min(r.hora) >= tu.ini && hm2min(r.hora) < tu.fim).length : null;
      return `<div class="card"><div class="toolbar" style="margin-bottom:6px"><h2 style="margin:0">${tu.nome} <span class="hint">· ${tu.faixa}</span></h2><span class="hint"><b>${feitas}/${ts.length}</b> feitas</span></div>
        <div class="prog"><div style="width:${pct}%"></div></div>
        ${nRep != null ? `<div class="hint" style="margin:6px 0 2px">Reportes hora a hora no turno: <b>${nRep}</b></div>` : ""}
        ${rows}</div>`;
    }).join("");
  }
  // alerta: dentro do turno, mais de 1h (+10 min) sem reporte
  function semReporte(dia, diaSel) {
    if (diaSel !== hojeStr()) return null;
    const agora = minAgora(), t = turnoDe(agora); if (!t) return null;
    const noTurno = dia.map(i => hm2min(i.hora)).filter(m => m >= t.ini && m <= agora);
    if (!noTurno.length) return agora - t.ini > 15 ? agora - t.ini : null;
    const gap = agora - Math.max(...noTurno);
    return gap > 70 ? gap : null;
  }
  function atrasado(nome, st, diaSel) {
    if (diaSel !== hojeStr() || st !== "aguardando") return false;
    const h = horaDe(nome); if (h == null) return false;
    return minAgora() > h + 15;
  }
  // Preenche os elementos padrão do hora a hora (ids: sUlt, sUltN, sEnv, sFila, sFilaN, sAlert, sAlertN, supDisp, supFila, supEmpty, supList)
  function renderHora($, cfg, items, diaSel, opts) {
    opts = opts || {};
    const dia = items.filter(i => i.data === diaSel).sort((a, b) => a.hora.localeCompare(b.hora) || (a.criadoEm - b.criadoEm));
    const ult = dia[dia.length - 1];
    $("sUlt").textContent = ult ? ult.hora : "—";
    $("sUltN").textContent = ult ? [ult.responsavel, dia.length + " reporte(s) no dia"].filter(Boolean).join(" · ") : "Sem reportes";
    $("sEnv").textContent = ult ? ult.total.toLocaleString("pt-BR") : "—";
    $("sFila").textContent = ult ? ult.fila : "—";
    const fmax = dia.length ? Math.max(...dia.map(i => i.fila)) : 0;
    $("sFilaN").textContent = dia.length ? "Pico do dia: " + fmax : "";
    const nomes = [...new Set([...disparos(cfg), ...(ult ? ult.disp.map(d => d.nome) : [])])];
    const probs = nomes.map(n => ({ nome: n, status: ult?.disp.find(x => x.nome === n)?.status || "aguardando" })).filter(d => d.status === "problema" || atrasado(d.nome, d.status, diaSel));
    const sr = semReporte(dia, diaSel);
    const nAl = probs.length + (sr != null ? 1 : 0);
    $("sAlert").textContent = (ult || nAl) ? nAl : "—";
    $("sAlert").className = "val" + (nAl ? " up" : "");
    $("sAlertN").textContent = nAl ? [...(sr != null ? [`Sem reporte há ${sr} min`] : []), ...probs.map(d => curto(d.nome) + (d.status === "problema" ? "" : " atrasado"))].join(", ") : (ult ? "Tudo certo" : "");
    $("supDisp").innerHTML = nomes.length ? `<table class="dtab">${nomes.map(n => {
      const d = ult?.disp.find(x => x.nome === n); const st = d?.status || "aguardando"; const at = atrasado(n, st, diaSel);
      return `<tr><td>${esc(n)}</td><td><span class="chip ${at ? "problema" : st}">${at ? "⚠️ Atrasado" : ST[st].e + " " + ST[st].t}</span></td><td class="num">${d && d.qtd !== "" && d.qtd != null ? (+d.qtd).toLocaleString("pt-BR") : ""}</td></tr>`;
    }).join("")}</table>` : '<div class="empty">Nenhum disparo cadastrado.</div>';
    const fm = Math.max(fmax, 1);
    $("supFila").innerHTML = dia.length ? dia.map((i, ix) => `<div class="col ${ix === dia.length - 1 ? "cur" : ""}" title="${i.hora}: ${i.fila}"><div class="t">${i.fila}</div><div class="b" style="height:${i.fila / fm * 100}%"></div><div class="m">${i.hora.replace(":00", "h")}</div></div>`).join("") : '<div class="empty">Sem dados neste dia.</div>';
    $("supEmpty").style.display = dia.length ? "none" : "block";
    $("supList").innerHTML = [...dia].reverse().map(i => {
      const act = !opts.acoes ? "" : opts.confirm === i.id
        ? `<button data-sdel="${i.id}" style="color:var(--red)">Excluir?</button> <button data-snodel>Não</button>`
        : `<button data-scopy="${i.id}">Copiar</button> <button data-sedit="${i.id}">Editar</button> <button data-sask="${i.id}">✕</button>`;
      return `<div class="rep"><div class="h">${esc(i.hora)}<small>${esc(i.responsavel || "")}</small></div>
        <div><div class="chips">${i.disp.map(d => `<span class="chip ${esc(d.status)}" title="${esc(ST[d.status]?.t || "")}">${ST[d.status]?.e || ""} ${esc(curto(d.nome))}${d.qtd !== "" && d.qtd != null ? " · " + (+d.qtd).toLocaleString("pt-BR") : ""}</span>`).join("")}</div>
        <div class="meta">Enviados: <b>${i.total.toLocaleString("pt-BR")}</b> · Fila: <b>${i.fila}</b>${i.obs ? " · " + esc(i.obs) : ""}</div></div>
        <div class="actions">${act}</div></div>`;
    }).join("");
  }
  function textoWhats(i) {
    const [y, m, d] = i.data.split("-");
    const l = [`*Reporte ${i.hora} – ${d}/${m}*`];
    if (i.responsavel) l.push(`Responsável: ${i.responsavel}`);
    l.push("", "📤 *Disparos*");
    i.disp.forEach(x => l.push(`${ST[x.status]?.e || ""} ${x.nome}: ${ST[x.status]?.t || ""}${x.qtd !== "" && x.qtd != null ? " – " + (+x.qtd).toLocaleString("pt-BR") + " enviados" : ""}`));
    l.push(`Total enviado até agora: *${(+i.total).toLocaleString("pt-BR")}*`, "", `👥 Fila de atendimento: *${i.fila}*`);
    if (i.obs) l.push("", `📝 ${i.obs}`);
    return l.join("\n");
  }
  return { DEF, TURNOS, ST, minAgora, hm2min, turnoDe, horaDe, disparos, resp, tarefas, parseDados, tarefasDoDia, ckId, htmlChecklist, renderHora, textoWhats, curto };
})();
