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

  // cada disparo tem 2 conferências: início (começou sem erro) e fim, 3h depois (100% da base)
  const txtIni = d => "Início do disparo " + d + " – começou sem erro";
  const txtFim = d => "Fim do disparo " + d + " – enviado a 100% da base";
  // tarefas do dia = 2 conferências por disparo + tarefas fixas
  function tarefasDoDia(cfg) {
    const out = [];
    disparos(cfg).forEach(d => {
      const h = horaDe(d); if (h == null) return;
      const ini = Math.max(h, TURNOS[0].ini); // disparo antes do turno: confere no início do turno
      out.push({ min: ini, hora: min2hm(ini), texto: txtIni(d) });
      const fim = h + 180;
      out.push({ min: fim, hora: min2hm(fim), texto: txtFim(d), fim: true });
    });
    tarefas(cfg).forEach(t => {
      const m = horaDe(t); if (m == null) return;
      out.push({ min: m, hora: min2hm(m), texto: semHora(t) || t });
    });
    out.forEach(t => { const tu = turnoDe(t.min) || (t.min < TURNOS[0].ini ? TURNOS[0] : TURNOS[1]); t.turno = tu.nome; t.tk = tu.chave; });
    return out.sort((a, b) => a.min - b.min);
  }
  // id da tarefa marcada: dia + horário + texto (tarefas com o mesmo texto em turnos diferentes não se misturam)
  function ckId(dia, texto, hora) {
    const slug = texto.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 60);
    return "ck-" + dia + "-" + String(hora || "").replace(":", "") + "-" + slug;
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
        const f = checks.find(c => c.id === ckId(dia, t.texto, t.hora) && c.data === dia);
        const st = statusTarefa(t, f, dia);
        const btn = interativo
          ? `<button type="button" class="ckbtn ${f ? "on" : ""}" data-ck="${encodeURIComponent(JSON.stringify({ texto: t.texto, hora: t.hora, turno: t.turno, fim: !!t.fim }))}" aria-label="${f ? "Desmarcar" : "Marcar como feita"}">${f ? "✓" : ""}</button>`
          : `<span class="ckbtn ${f ? "on" : ""}">${f ? "✓" : ""}</span>`;
        let extra = {}; try { extra = f ? JSON.parse(f.dados || "{}") : {}; } catch (e) {}
        const nr = f && extra.naoReceberam !== undefined && extra.naoReceberam !== "" ? `<div class="hint">📵 Não receberam: <b>${(+extra.naoReceberam).toLocaleString("pt-BR")}</b></div>` : "";
        const obs = (f && f.obs ? `<div class="hint">📝 ${esc(f.obs)}</div>` : "") + nr;
        return `<div class="ckrow">${btn}<div class="ckh">${t.hora}</div><div class="ckt">${esc(t.texto)}${obs}</div><span class="chip ${st.cls}">${st.txt}</span></div>`;
      }).join("") || '<div class="empty">Nenhuma tarefa neste turno.</div>';
      const feitas = ts.filter(t => checks.some(c => c.id === ckId(dia, t.texto, t.hora) && c.data === dia)).length;
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
    const checks = opts.checks || [];
    const horaTarefa = {}; tarefasDoDia(cfg).forEach(t => horaTarefa[t.texto] = t.hora);
    const conferido = n => checks.find(c => c.data === diaSel && c.id === ckId(diaSel, txtIni(n), horaTarefa[txtIni(n)]));
    const finalizado = n => checks.find(c => c.data === diaSel && c.id === ckId(diaSel, txtFim(n), horaTarefa[txtFim(n)]));
    const probs = nomes.map(n => ({ nome: n, status: ult?.disp.find(x => x.nome === n)?.status || "aguardando" })).filter(d => d.status === "problema" || (atrasado(d.nome, d.status, diaSel) && !conferido(d.nome)));
    const sr = semReporte(dia, diaSel);
    const nAl = probs.length + (sr != null ? 1 : 0);
    $("sAlert").textContent = (ult || nAl) ? nAl : "—";
    $("sAlert").className = "val" + (nAl ? " up" : "");
    $("sAlertN").textContent = nAl ? [...(sr != null ? [dia.length ? `Sem reporte há ${sr} min` : "Nenhum reporte no turno ainda"] : []), ...probs.map(d => curto(d.nome) + (d.status === "problema" ? "" : " atrasado"))].join(", ") : (ult ? "Tudo certo" : "");
    $("supDisp").innerHTML = nomes.length ? `<table class="dtab">${nomes.map(n => {
      const d = ult?.disp.find(x => x.nome === n); const st = d?.status || "aguardando"; const ck = conferido(n); const at = atrasado(n, st, diaSel) && !ck;
      const fz = finalizado(n); let nr = ""; try { const x = fz ? JSON.parse(fz.dados || "{}") : {}; if (x.naoReceberam !== undefined && x.naoReceberam !== "") nr = " · " + (+x.naoReceberam).toLocaleString("pt-BR") + " não receberam"; } catch (e) {}
      const chip = at ? `<span class="chip problema">⚠️ Atrasado</span>`
        : fz ? `<span class="chip concluido">🏁 100% às ${esc(fz.hora)}${nr}</span>`
        : (st === "aguardando" && ck) ? `<span class="chip andamento">✅ Iniciou ok às ${esc(ck.hora)}</span>` : `<span class="chip ${st}">${ST[st].e} ${ST[st].t}</span>`;
      return `<tr><td>${esc(n)}</td><td>${chip}</td><td class="num">${d && d.qtd !== "" && d.qtd != null ? (+d.qtd).toLocaleString("pt-BR") : ""}</td></tr>`;
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
  // ---------- VALIDAÇÃO ----------
  const slug = t => String(t).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 40);
  const vId = e => "v-" + slug(e.base) + "-" + (e.id || ("l" + e.linha)) + "-" + slug(e.tipo);
  const VST = { corrigido: { t: "Corrigido na base", cls: "concluido" }, certo: { t: "Conferido – está certo", cls: "concluido" }, vendedor: { t: "Erro do vendedor", cls: "problema" } };
  function marcasMap(marks) { const m = {}; (marks || []).forEach(x => m[x.id] = x); return m; }
  function htmlBasesInfo(res) {
    if (!res) return "";
    const b = (res.bases || []).map(x => x.vazia ? `<span class="chip problema">${esc(x.nome)}: vazia</span>`
      : x.faltando.length ? `<span class="chip problema">${esc(x.nome)}: faltam colunas ${esc(x.faltando.join(", "))}</span>`
      : `<span class="chip">${esc(x.nome)}: ${x.linhas.toLocaleString("pt-BR")} linhas</span>`).join(" ");
    return `<div class="chips" style="margin-bottom:10px">${b}</div>`;
  }
  function filtrarErros(res, marks, f) {
    const mm = marcasMap(marks);
    return (res?.erros || []).map(e => Object.assign({}, e, { vid: vId(e), marca: mm[vId(e)] })).filter(e =>
      (!f.nivel || e.nivel === f.nivel) && (!f.base || e.base === f.base) &&
      (!f.status || (f.status === "pend" ? !e.marca : !!e.marca)) &&
      (!f.vendedor || (e.vendedor || "(sem vendedor)") === f.vendedor) && (!f.tipo || e.tipo === f.tipo));
  }
  function kpisValidacao(res, marks) {
    const all = filtrarErros(res, marks, {});
    const pend = all.filter(e => !e.marca);
    return { erros: pend.filter(e => e.nivel === "erro").length, alertas: pend.filter(e => e.nivel === "alerta").length,
      tratados: all.length - pend.length, total: all.length,
      vendErros: all.filter(e => e.marca?.status === "vendedor").length };
  }
  function htmlListaValidacao(lista, interativo, limite) {
    if (!lista.length) return '<div class="empty">Nada para mostrar com esses filtros. 🎉</div>';
    const rows = lista.slice(0, limite || 300).map(e => {
      const m = e.marca, st = m ? VST[m.status] || { t: m.status, cls: "" } : null;
      const btn = interativo ? `<button type="button" data-vt="${esc(e.vid)}">${m ? "Alterar" : "Tratar"}</button>` : "";
      return `<div class="vrow">
        <div><span class="chip ${e.nivel === "erro" ? "problema" : "andamento"}">${e.nivel === "erro" ? "Erro" : "Alerta"}</span></div>
        <div><b>${esc(e.tipo)}</b> · ${esc(e.cliente || "(sem nome)")} <span class="hint">ID ${esc(e.id)}</span>
          <div class="hint">${esc(e.detalhe)}</div>
          <div class="hint">${esc(e.origem)} ${esc(e.data)} · ${esc(e.base)} (linha ${e.linha}) · Vendedor: <b>${esc(e.vendedor || "—")}</b></div>
          ${m ? `<div class="hint"><span class="chip ${st.cls}">${esc(st.t)}</span> ${esc(m.responsavel || "")} ${esc(m.data ? m.data.split("-").reverse().join("/") : "")} ${esc(m.hora || "")}${m.obs ? " · 📝 " + esc(m.obs) : ""}</div>` : ""}
        </div><div class="actions">${btn}</div></div>`;
    }).join("");
    const mais = lista.length > (limite || 300) ? `<div class="hint" style="padding:8px 0">Mostrando ${limite || 300} de ${lista.length}. Use os filtros para ver o resto.</div>` : "";
    return rows + mais;
  }
  function htmlPorVendedor(res, marks) {
    const all = filtrarErros(res, marks, {}).filter(e => e.nivel === "erro" || e.marca?.status === "vendedor");
    const c = {}; all.forEach(e => { const k = e.vendedor || "(sem vendedor)"; c[k] = c[k] || { pend: 0, total: 0 }; c[k].total++; if (!e.marca) c[k].pend++; });
    const l = Object.entries(c).sort((a, b) => b[1].total - a[1].total);
    if (!l.length) return '<div class="empty">Nenhum erro.</div>';
    const mx = l[0][1].total;
    return l.map(([k, v]) => `<div class="bar-row"><div class="name">${esc(k)}</div><div class="bar-track"><div class="bar" style="width:${Math.max(3, v.total / mx * 100)}%"></div></div><div class="v">${v.total}${v.pend ? ` <span class="hint">(${v.pend} pend.)</span>` : ""}</div></div>`).join("");
  }
  // ---------- CRÉDITO ----------
  const LIMITE = 450;
  // saldo = último "saldo informado" + recargas depois dele − consumo desde a data dele
  function infoSaldo(recargas) {
    const l = [...(recargas || [])].sort((a, b) => (a.data + (a.criadoEm || 0)).localeCompare(b.data + (b.criadoEm || 0)));
    const ultSaldo = [...l].reverse().find(r => r.status === "saldo");
    const desde = ultSaldo ? ultSaldo.data : null;
    const recargasDepois = l.filter(r => r.status === "recarga" && (!ultSaldo || r.data > ultSaldo.data || (r.data === ultSaldo.data && (+r.criadoEm || 0) > (+ultSaldo.criadoEm || 0))));
    return { ultSaldo, desde, somaRecargas: recargasDepois.reduce((s, r) => s + (+r.total || 0), 0) };
  }
  function calcSaldo(recargas, res) {
    const inf = infoSaldo(recargas);
    if (!inf.ultSaldo || !res || res.consumoDesde == null || res.desde !== inf.desde) return null;
    return (+inf.ultSaldo.total || 0) + inf.somaRecargas - res.consumoDesde;
  }
  function htmlCredito(res, recargas, owner) {
    const c = res?.consumo; if (!c) return '<div class="empty">Carregando…</div>';
    const saldo = calcSaldo(recargas, res), inf = infoSaldo(recargas);
    const alerta = saldo != null && saldo <= LIMITE;
    const mesNome = new Date(res.mes + "-15").toLocaleDateString("pt-BR", { month: "long" });
    const k = (lbl, val, note, cls) => `<div class="card kpi"><div class="lbl">${lbl}</div><div class="val ${cls || ""}">${val}</div><div class="note">${note || ""}</div></div>`;
    const n = x => (+x || 0).toLocaleString("pt-BR");
    let html = alerta ? `<div class="banner" style="display:block">⚠️ Saldo de créditos em <b>${n(saldo)}</b>, abaixo do limite de ${LIMITE}. Pedir recarga.</div>` : "";
    html += `<div class="kpis">
      ${k("Consumo em " + mesNome, n(c.total), "créditos até " + res.hoje.split("-").reverse().slice(0, 2).join("/"))}
      ${k("Vendas", n(c.vendas), n(c.linhasVendas) + " cliente(s)")}
      ${k("Renovações", n(c.renovacoes), n(c.linhasRenovacoes) + " cliente(s)")}
      ${k("Saldo estimado", saldo == null ? "—" : n(saldo), saldo == null ? (owner ? "Registre o saldo atual abaixo" : "Saldo ainda não informado") : "desde " + inf.desde.split("-").reverse().join("/") + " · limite " + LIMITE, alerta ? "up" : "")}
    </div>`;
    const bars = obj => { const e = Object.entries(obj || {}).sort((a, b) => b[1] - a[1]); const mx = e.length ? e[0][1] : 1;
      return e.map(([k2, v]) => `<div class="bar-row"><div class="name" title="${esc(k2)}">${esc(k2)}</div><div class="bar-track"><div class="bar" style="width:${Math.max(2, v / mx * 100)}%"></div></div><div class="v">${n(v)}</div></div>`).join("") || '<div class="empty">Sem consumo.</div>'; };
    html += `<div class="grid2"><div class="card"><h2>Consumo por servidor</h2>${bars(c.porServidor)}</div><div class="card"><h2>Consumo por base</h2>${c.detalheBase ? `<table class="dtab"><tr><td class="hint">Base</td><td class="hint num">Vendas</td><td class="hint num">Renovações</td><td class="hint num">Total</td></tr>${Object.entries(c.detalheBase).map(([nm, d]) => `<tr><td><b>${esc(nm)}</b></td><td class="num">${n(d.vendas)}<div class="hint">${n(d.clientesVendas)} cli.</div></td><td class="num">${n(d.renovacoes)}<div class="hint">${n(d.clientesRenovacoes)} cli.</div></td><td class="num"><b>${n(d.vendas + d.renovacoes)}</b></td></tr>`).join("")}<tr><td><b>Total</b></td><td class="num"><b>${n(c.vendas)}</b></td><td class="num"><b>${n(c.renovacoes)}</b></td><td class="num"><b>${n(c.total)}</b></td></tr></table>` : bars(c.porBase)}
      <div class="hint" style="margin-top:10px">Cada cliente vale telas × meses do plano. Telas ou plano vazio contam como 1. Duplicados ficam de fora.</div></div></div>`;
    const lst = [...(recargas || [])].sort((a, b) => (b.data + (b.criadoEm || 0)).localeCompare(a.data + (a.criadoEm || 0)));
    html += `<div class="card"><div class="toolbar"><h2 style="margin:0">Recargas e saldo informado</h2>${owner ? '<button class="primary" id="crAdd">+ Registrar</button>' : ""}</div>
      ${lst.length ? lst.map(r => `<div class="rep" style="grid-template-columns:90px 1fr auto"><div class="h">${esc(r.data.split("-").reverse().join("/"))}</div>
        <div><span class="chip ${r.status === "saldo" ? "andamento" : "concluido"}">${r.status === "saldo" ? "Saldo informado" : "Recarga"}</span> <b>${r.status === "saldo" ? "" : "+"}${n(r.total)}</b> créditos${r.obs ? `<div class="meta">${esc(r.obs)}</div>` : ""}</div>
        <div class="actions">${owner ? `<button data-crdel="${esc(r.id)}">✕</button>` : ""}</div></div>`).join("") : '<div class="empty">Nenhuma recarga registrada.</div>'}</div>`;
    return html;
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
  return { DEF, TURNOS, ST, VST, LIMITE, minAgora, hm2min, turnoDe, horaDe, disparos, resp, tarefas, parseDados, tarefasDoDia, ckId, htmlChecklist, renderHora, textoWhats, curto,
    vId, filtrarErros, kpisValidacao, htmlListaValidacao, htmlPorVendedor, htmlBasesInfo, infoSaldo, calcSaldo, htmlCredito };
})();
