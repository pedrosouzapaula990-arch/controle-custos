// Regras compartilhadas do suporte (usado pelo site do suporte e pela aba Suporte do site de custos)
const SupCore = (() => {
  const DEF = {
    disparos: ["07:00 – Vence hoje", "10:15 – Vence amanhã", "11:30 – Venceu ontem", "17:40 – Vence hoje", "19:30 – Venceu ontem"],
    resp: ["Diego", "Felipe", "Kaique", "Lígia"],
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
  const RESP_ANTIGO = JSON.stringify(["Turno manhã (9h–16:40)", "Turno noite (16:40–23:30)"]);
  // lista padrão de responsáveis; quem ainda tem a lista antiga (só os turnos) passa a ver os nomes
  const resp = cfg => (cfg?.resp && JSON.stringify(cfg.resp) !== RESP_ANTIGO) ? cfg.resp : DEF.resp;
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
    const quando = res.basesAlteradas ? `<div class="hint" style="margin-top:6px">Bases alteradas em <b>${esc(res.basesAlteradas)}</b> · cálculo de ${esc(res.calculadoEm || "")}.${res.desatualizado ? " <b>As bases mudaram: o cálculo novo sai em até 5 min (ou clique em Recalcular).</b>" : " Só recalcula quando as bases mudarem."}</div>` : "";
    const k = res.kommo;
    const kom = !k ? "" : !k.configurado ? `<div class="hint" style="margin-top:4px">Kommo: ainda não configurado (as bases vêm coladas na planilha).</div>`
      : `${(function () { const c = k.conferencia; if (!c) return `<div class="hint" style="margin-top:6px"><span class="chip andamento">Extração</span> sem conferência ainda (aparece na próxima puxada do Kommo)</div>`;
        return `<div class="hint" style="margin-top:6px"><span class="chip ${c.completa ? "concluido" : "problema"}">${c.completa ? "✔ Extração completa" : "✖ Extração incompleta"}</span> extraída em <b>${esc(c.quando)}</b> · ativos e vencidos${c.completa ? " 100% gravados" : ": " + esc(c.detalhe)}</div>${c.colunasVazias ? `<div class="hint" style="color:var(--red)">⚠ Colunas sem nenhum valor vindo do Kommo — ${esc(c.colunasVazias)}. Campos vistos: ${esc(c.campos || "—")}</div>` : ""}`; })()}
      <div class="hint" style="margin-top:4px">Kommo: ${k.puxando ? "<b>puxando agora" + (k.andamento ? ": " + esc(k.andamento) : "") + "…</b>" : "última puxada " + esc(k.ultima || "—")} · automático às 7h, 12h e 19h${k.erro ? ` · <span style="color:var(--red)">erro: ${esc(k.erro)}</span>` : ""}</div>`;
    return `<div class="chips" style="margin-bottom:4px">${b}</div>${quando}${kom}`;
  }
  function filtrarErros(res, marks, f) {
    const mm = marcasMap(marks);
    const ms = res?.basesAlteradasMs || 0;
    return (res?.erros || []).map(e => {
      let marca = mm[vId(e)], reaberto = false;
      // "Corrigido" mas o erro continua aparecendo depois de as bases mudarem → volta a ficar pendente
      if (marca && marca.status === "corrigido" && ms && +marca.criadoEm && ms > +marca.criadoEm + 60000) { reaberto = true; marca = null; }
      return Object.assign({}, e, { vid: vId(e), marca: marca, reaberto: reaberto, marcaOriginal: mm[vId(e)] });
    }).filter(e =>
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
  // Agrupa os pontos por lead: um lead com vários problemas aparece uma vez só
  const chaveLead = e => e.base + "|" + (e.id || ("l" + e.linha));
  function agruparLeads(lista) {
    const m = new Map();
    lista.forEach(e => { const k = chaveLead(e); if (!m.has(k)) m.set(k, []); m.get(k).push(e); });
    return [...m.values()];
  }
  function htmlListaValidacao(lista, interativo, limite) {
    const grupos = agruparLeads(lista), nL = grupos.length, nErro = grupos.filter(g => g.some(e => e.nivel === "erro")).length;
    const cont = `<div class="hint" style="padding:4px 0 10px"><b>${nL.toLocaleString("pt-BR")}</b> lead${nL === 1 ? "" : "s"} nesta lista${nL ? ` · ${nErro.toLocaleString("pt-BR")} com erro · ${(nL - nErro).toLocaleString("pt-BR")} só com alerta` : ""} <span style="opacity:.7">(${lista.length.toLocaleString("pt-BR")} ponto${lista.length === 1 ? "" : "s"})</span></div>`;
    if (!grupos.length) return cont + '<div class="empty">Nada para mostrar com esses filtros. 🎉</div>';
    const rows = grupos.slice(0, limite || 300).map(g => {
      const e = g[0], temErro = g.some(x => x.nivel === "erro"), todosTratados = g.every(x => x.marca);
      const btn = interativo ? `<button type="button" data-vt="${esc(e.vid)}">${todosTratados ? "Alterar" : "Tratar"}${g.length > 1 ? " tudo" : ""}</button>` : "";
      const linhas = g.map(x => {
        const m = x.marca, st = m ? VST[m.status] || { t: m.status, cls: "" } : null;
        return `<div class="hint" style="margin-top:3px"><span class="chip ${x.nivel === "erro" ? "problema" : "andamento"}" style="font-size:10px">${x.nivel === "erro" ? "Erro" : "Alerta"}</span> <b>${esc(x.tipo)}</b> — ${esc(x.detalhe)}
          ${x.reaberto ? ` <span class="chip problema" style="font-size:10px">↩ Reaberto</span>` : ""}
          ${m ? ` <span class="chip ${st.cls}" style="font-size:10px">${esc(st.t)}</span> ${esc(m.responsavel || "")} ${esc(m.data ? m.data.split("-").reverse().join("/") : "")} ${esc(m.hora || "")}${m.obs ? " · 📝 " + esc(m.obs) : ""}` : ""}</div>`;
      }).join("");
      const origens = [...new Set(g.map(x => x.origem + " " + x.data))].join(" · ");
      return `<div class="vrow">
        <div><span class="chip ${temErro ? "problema" : "andamento"}">${temErro ? "Erro" : "Alerta"}</span>${g.length > 1 ? `<div class="hint" style="margin-top:4px">${g.length} pontos</div>` : ""}</div>
        <div><b>${esc(e.cliente || "(sem nome)")}</b> <span class="hint">ID ${esc(e.id)}</span>
          <div class="hint">${esc(origens)} · ${esc(e.base)} (linha ${e.linha}) · Vendedor: <b>${esc(e.vendedor || "—")}</b></div>
          ${linhas}
        </div><div class="actions">${btn}</div></div>`;
    }).join("");
    const mais = grupos.length > (limite || 300) ? `<div class="hint" style="padding:8px 0">Mostrando ${limite || 300} de ${grupos.length} leads. Use os filtros para ver o resto.</div>` : "";
    return cont + rows + mais;
  }
  function htmlPorVendedor(res, marks) {
    const all = filtrarErros(res, marks, {}).filter(e => e.nivel === "erro" || e.marca?.status === "vendedor");
    const c = {}; all.forEach(e => { const k = e.vendedor || "(sem vendedor)"; c[k] = c[k] || { pend: 0, total: 0 }; c[k].total++; if (!e.marca) c[k].pend++; });
    const l = Object.entries(c).sort((a, b) => b[1].total - a[1].total);
    if (!l.length) return '<div class="empty">Nenhum erro.</div>';
    const mx = l[0][1].total;
    return l.map(([k, v]) => `<div class="bar-row"><div class="name">${esc(k)}</div><div class="bar-track"><div class="bar" style="width:${Math.max(3, v.total / mx * 100)}%"></div></div><div class="v">${v.total}${v.pend ? ` <span class="hint">(${v.pend} pend.)</span>` : ""}</div></div>`).join("");
  }
  // Auditoria e qualidade por vendedor (vendas do mês × vendas com erro de dados)
  function htmlQualidadeVendedor(res) {
    const v = res?.vendedores; if (!v) return "";
    const l = Object.entries(v).filter(([, x]) => x.vendas > 0).map(([k, x]) => ({ k, vendas: x.vendas, erro: x.comErro, q: Math.round((1 - x.comErro / x.vendas) * 100) })).sort((a, b) => b.vendas - a.vendas);
    if (!l.length) return "";
    return `<div class="card"><h2>Vendas e qualidade por vendedor (mês)</h2><table class="dtab"><tr><td class="hint">Vendedor</td><td class="hint num">Vendas</td><td class="hint num">Com erro</td><td class="hint num">Qualidade</td></tr>
      ${l.map(x => `<tr><td><b>${esc(x.k)}</b></td><td class="num">${x.vendas}</td><td class="num">${x.erro}</td><td class="num"><b>${x.q}%</b></td></tr>`).join("")}</table>
      <div class="hint" style="margin-top:8px">Qualidade = vendas do mês sem nenhum erro de cadastro (servidor, plano, telas, vencimento…).</div></div>`;
  }
  // Conciliação de clientes ativos + previsões
  function htmlAtivos(res) {
    const a = res?.ativos; if (!a) return "";
    const n = x => (+x || 0).toLocaleString("pt-BR");
    const mov = a.movimento || {};
    let h = `<div class="card"><h2>Clientes ativos – ${esc(a.base)}</h2><div class="kpis">
      <div class="card kpi"><div class="lbl">Ativos hoje</div><div class="val">${n(a.agora)}</div><div class="note">vencimento de hoje em diante</div></div>`;
    if (a.esperado != null) h += `<div class="card kpi"><div class="lbl">Esperado</div><div class="val">${n(a.esperado)}</div><div class="note">${n(a.inicioMes)} início + ${n(a.vendasMes)} vendas − ${n(a.naoRenovaram)} não renovaram</div></div>
      <div class="card kpi"><div class="lbl">Diferença</div><div class="val ${Math.abs(a.diferenca) > 20 ? "up" : ""}">${a.diferenca > 0 ? "+" : ""}${n(a.diferenca)}</div><div class="note">${a.diferenca === 0 ? "bate certinho" : "ativos hoje − esperado"}</div></div>`;
    h += `<div class="card kpi"><div class="lbl">Previsão fim do mês</div><div class="val">${n(a.previstoFimMes)}</div><div class="note">estimativa · renovação ${a.taxaRenovacao}% · ${a.diasRestantes} dia(s) restantes</div></div></div>`;
    if (mov.entradas != null) h += `<div class="hint" style="margin-top:8px">Desde ${esc(mov.refInicio.split("-").reverse().join("/"))} (foto do Drive): <b>${n(mov.entradas)}</b> entraram (<b>${n(mov.reconquistas)}</b> reconquistas) · <b>${n(mov.saidas)}</b> saíram${mov.variacaoDia != null ? ` · variação vs ontem: <b>${mov.variacaoDia > 0 ? "+" : ""}${n(mov.variacaoDia)}</b>` : ""}.</div>`;
    else h += `<div class="hint" style="margin-top:8px">As fotos diárias começaram agora; a comparação de entradas e saídas aparece a partir do próximo dia.${mov.movimentoErro ? " " + esc(mov.movimentoErro) : ""}</div>`;
    if (a.esperado != null && Math.abs(a.diferenca) > 20) h += `<div class="hint">Diferença acima de 20: conferir cadastros sem data de vencimento, vendas sem data de cadastro ou clientes tabulados na etapa errada.</div>`;
    return h + `</div>`;
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
  function prevCred(saldo, c, res) {
    const dia = +res.hoje.slice(8, 10), media = c.total / Math.max(1, dia);
    if (saldo == null || media <= 0) return "—";
    if (saldo <= LIMITE) return "já no limite";
    const d = Math.floor((saldo - LIMITE) / media);
    return d + " dia(s)";
  }
  function htmlCredito(res, recargas, owner, semAtivos) {
    const c = res?.consumo; if (!c) return '<div class="empty">Buscando o último cálculo das bases…</div>';
    const saldo = calcSaldo(recargas, res), inf = infoSaldo(recargas);
    const alerta = saldo != null && saldo <= LIMITE;
    const mesNome = new Date(res.mes + "-15").toLocaleDateString("pt-BR", { month: "long" });
    const k = (lbl, val, note, cls) => `<div class="card kpi"><div class="lbl">${lbl}</div><div class="val ${cls || ""}">${val}</div><div class="note">${note || ""}</div></div>`;
    const n = x => (+x || 0).toLocaleString("pt-BR");
    let html = res.basesAlteradas ? `<div class="hint" style="margin-bottom:10px">Bases de ${esc(res.basesAlteradas)} · cálculo de ${esc(res.calculadoEm || "")}${res.desatualizado ? " · <b>atualizando (até 5 min)</b>" : ""}</div>` : "";
    if (!semAtivos) html += htmlAtivos(res);
    html += alerta ? `<div class="banner" style="display:block">⚠️ Saldo de créditos em <b>${n(saldo)}</b>, abaixo do limite de ${LIMITE}. Pedir recarga.</div>` : "";
    html += `<div class="kpis">
      ${k("Consumo em " + mesNome, n(c.total), (res.basesAlteradas ? "bases de " + res.basesAlteradas : "créditos até " + res.hoje.split("-").reverse().slice(0, 2).join("/")))}
      ${k("Vendas", n(c.vendas), n(c.linhasVendas) + " cliente(s)")}
      ${k("Renovações", n(c.renovacoes), n(c.linhasRenovacoes) + " cliente(s)")}
      ${k("Previsão do crédito", prevCred(saldo, c, res), "dias até chegar em " + LIMITE)}
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

  // ---------- DASHBOARD (aba só do dono) ----------
  const ultMeses = (mes, k) => { const [y, m] = mes.split("-").map(Number), l = []; for (let i = k - 1; i >= 0; i--) { const d = new Date(y, m - 1 - i, 1); l.push(d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0")); } return l; };
  const mesCurto = m => { const [y, mm] = m.split("-"); return new Date(+y, +mm - 1, 1).toLocaleDateString("pt-BR", { month: "short" }).replace(".", "") + "/" + y.slice(2); };
  function htmlDashboard(sub, c) {
    const res = c.res, brl = c.brl, n = x => (+x || 0).toLocaleString("pt-BR");
    const k = (l, v, note, cls) => `<div class="card kpi"><div class="lbl">${l}</div><div class="val ${cls || ""}">${v}</div><div class="note">${note || ""}</div></div>`;
    const barras = (pares, fmt) => { const mx = Math.max(1, ...pares.map(p => p[1])); return pares.map(([nm, v]) => `<div class="bar-row"><div class="name" title="${esc(nm)}">${esc(nm)}</div><div class="bar-track"><div class="bar" style="width:${Math.max(2, v / mx * 100)}%"></div></div><div class="v">${(fmt || n)(v)}</div></div>`).join("") || '<div class="empty">Sem dados.</div>'; };
    if (!res || !res.consumo) return '<div class="empty">Carregando os números das bases…</div>';
    const custos = c.custos || [], totCustos = custos.reduce((a, i) => a + (+i.valor || 0), 0);
    const trafego = custos.filter(i => /an[uú]ncio|tr[aá]fego|meta ads/i.test(i.categoria || "")), totTraf = trafego.reduce((a, i) => a + (+i.valor || 0), 0);
    const co = res.consumo, vendas = co.linhasVendas, ren = co.linhasRenovacoes, at = res.ativos || {};
    const saldo = calcSaldo(c.recargas, res), kp = kpisValidacao(res, c.marks);
    const cpv = vendas ? totTraf / vendas : 0;
    const meses = ultMeses(res.mes, 6), sr = res.serie || {};
    if (sub === "resumo") {
      const av = [];
      if (saldo != null && saldo <= LIMITE) av.push(`Saldo de créditos em <b>${n(saldo)}</b>, abaixo de ${LIMITE}. Pedir recarga.`);
      if (at.diferenca != null && Math.abs(at.diferenca) > 20) av.push(`Clientes ativos com diferença de <b>${at.diferenca > 0 ? "+" : ""}${n(at.diferenca)}</b> contra o esperado.`);
      if (kp.erros) av.push(`<b>${n(kp.erros)}</b> erros pendentes na validação.`);
      const cf = res.kommo && res.kommo.conferencia; if (cf && !cf.completa) av.push("Última extração do Kommo <b>incompleta</b>.");
      if (res.desatualizado) av.push("As bases mudaram e o cálculo está atualizando.");
      return `<div class="kpis">${k("Custos do mês", brl(totCustos), n(custos.length) + " lançamento(s)")}${k("Vendas do mês", n(vendas), n(ren) + " renovações")}
        ${k("Clientes ativos", n(at.agora), at.esperado != null ? "esperado " + n(at.esperado) : "Base Darb Oficial", at.diferenca != null && Math.abs(at.diferenca) > 20 ? "up" : "")}${k("Consumo de crédito", n(co.total), "saldo " + (saldo == null ? "—" : n(saldo)) + " · " + prevCred(saldo, co, res))}
        ${k("Anúncios", brl(totTraf), cpv ? brl(cpv) + " por venda" : "sem vendas no mês")}${k("Erros pendentes", n(kp.erros), n(kp.alertas) + " alertas")}
        ${k("Previsão ativos (fim do mês)", n(at.previstoFimMes), "estimativa")}${k("Renovação no mês", (at.taxaRenovacao != null ? at.taxaRenovacao : "—") + "%", "Base Darb Oficial")}</div>
        <div class="card"><h2>Atenção</h2>${av.length ? av.map(x => `<div class="hint" style="margin:5px 0">⚠️ ${x}</div>`).join("") : '<div class="hint">Nada fora do normal agora. ✔</div>'}</div>
        <div class="hint" style="margin-top:8px">Bases de ${esc(res.basesAlteradas || "—")} · cálculo de ${esc(res.calculadoEm || "—")}.</div>`;
    }
    if (sub === "meta") { const mt = c.meta || {};
      return `<div class="card"><div class="grid2" style="margin:0"><div class="field"><label for="mtAlvo">Quero ter quantos clientes?</label><input id="mtAlvo" type="number" min="0" step="100" value="${esc(mt.alvo)}"></div>
        <div class="field"><label for="mtMes">Quero chegar nisso até o mês de</label><input id="mtMes" type="month" value="${esc(mt.mes)}"></div></div></div><div id="metaOut">${htmlMeta(res, mt)}</div>`; }
    if (sub === "credito") return htmlCredito(res, c.recargas, false, true);
    if (sub === "vendas") {
      const dias = Object.entries(sr.vendasDia || {}).sort().map(([d, v]) => [d.slice(8, 10) + "/" + d.slice(5, 7), v]);
      const vend = Object.entries(res.vendedores || {}).sort((a, b) => b[1].vendas - a[1].vendas).map(([nm, x]) => [nm, x.vendas]);
      return `<div class="kpis">${k("Vendas no mês", n(vendas), "clientes com data de cadastro no mês")}${k("Vendas em créditos", n(co.vendas), "telas × meses")}${k("Renovações", n(ren), n(co.renovacoes) + " créditos")}${k("Média por dia", n(Math.round(vendas / Math.max(1, +res.hoje.slice(8, 10)))), "vendas/dia")}</div>
        <div class="grid2"><div class="card"><h2>Vendas por vendedor (mês)</h2>${barras(vend)}</div><div class="card"><h2>Vendas por dia (mês)</h2>${barras(dias)}</div></div>
        <div class="card"><h2>Vendas por mês</h2>${barras(meses.map(m => [mesCurto(m), (sr.vendasMes || {})[m] || 0]))}</div>${htmlQualidadeVendedor(res)}`;
    }
    if (sub === "trafego") {
      const l = [...trafego].sort((a, b) => (b.data || "").localeCompare(a.data || ""));
      return `<div class="kpis">${k("Gasto em anúncios", brl(totTraf), "no mês, das categorias de anúncios")}${k("Vendas no mês", n(vendas), "")}${k("Custo por venda", cpv ? brl(cpv) : "—", "anúncios ÷ vendas")}${k("Lançamentos", n(l.length), "de anúncios")}</div>
        <div class="card"><h2>Lançamentos de anúncios</h2>${l.length ? l.map(i => `<div class="rep" style="grid-template-columns:90px 1fr auto"><div class="h">${esc((i.data || "").split("-").reverse().join("/"))}</div><div>${esc(i.descricao || "")}<div class="meta">${esc(i.categoria || "")}</div></div><div><b>${brl(+i.valor)}</b></div></div>`).join("") : '<div class="empty">Nenhum lançamento de anúncios neste mês.</div>'}</div>
        <div class="hint">Conta as categorias com "anúncio", "tráfego" ou "Meta Ads" no nome.</div>`;
    }
    if (sub === "ativos") {
      const mov = at.movimento || {}, tab = (lista, vazio) => `<div style="max-height:46vh;overflow-y:auto"><table class="dtab"><tr><td class="hint">Cliente</td><td class="hint">Motivo</td><td class="hint">Etapa</td><td class="hint">Vencimento</td></tr>${lista.map(x => `<tr><td><b>${esc(x.cliente || "(sem nome)")}</b><div class="hint">ID ${esc(x.id)} · ${esc(x.base)} · ${esc(x.vendedor || "—")}</div></td><td>${esc(x.motivo)}</td><td class="hint">${esc(x.etapa)}</td><td class="hint">${esc(x.venc ? x.venc.split("-").reverse().join("/") : "—")}</td></tr>`).join("") || `<tr><td colspan="4" class="empty">${vazio}</td></tr>`}</table></div>`;
      const ren = meses.map(m => [mesCurto(m), (sr.renovMes || {})[m] || 0]);
      return htmlAtivos(res) + `<div class="grid2"><div class="card"><h2>Renovações por mês</h2>${barras(ren)}<div class="hint" style="margin-top:6px">Clientes com data de renovação em cada mês (as duas bases).</div></div>
        <div class="card"><h2>Ativos por base</h2>${barras(Object.entries(at.porBase || {}))}<div class="hint" style="margin-top:6px">A origem (campanha/indicação) precisa de um campo de origem no Kommo; ainda não vem nas bases.</div></div></div>
        ${mov.entradas != null ? `<div class="card"><h2>Entraram (${n(mov.totalEntradas)})</h2>${tab(mov.entradasLista || [], "Ninguém entrou desde a foto de referência.")}</div><div class="card"><h2>Saíram (${n(mov.totalSaidas)})</h2>${tab(mov.saidasLista || [], "Ninguém saiu.")}</div>` : '<div class="card"><div class="hint">A lista de quem entrou e saiu aparece a partir da segunda foto diária (amanhã).</div></div>'}`;
    }
    return "";
  }

  // ---------- META DE CLIENTES ----------
  // A = ativos hoje; perda diária = A × (1 − renovação) / 30; vendas/dia v para chegar em T em D dias: A·q^D + v·(1−q^D)/(1−q) = T, q = 1 − perda/30
  function calcMeta(res, alvo, mesAlvo) {
    const at = res && res.ativos; if (!at || !/^\d{4}-\d{2}$/.test(mesAlvo || "") || !(alvo > 0)) return null;
    const [y, m] = mesAlvo.split("-").map(Number), fim = new Date(Date.UTC(y, m, 0)), [hy, hm, hd] = res.hoje.split("-").map(Number);
    const D = Math.round((fim - new Date(Date.UTC(hy, hm - 1, hd))) / 86400000);
    const A = at.agora, r = Math.min(1, Math.max(0, (at.taxaRenovacao != null ? at.taxaRenovacao : 100) / 100)), c = 1 - r, ch = c / 30;
    const q = 1 - ch, mantem = A * ch;
    let v = null, fimRitmo = null;
    const ritmo = at.vendasMes / Math.max(1, hd);
    if (D > 0) {
      const qD = Math.pow(q, D);
      v = ch === 0 ? (alvo - A) / D : (alvo - A * qD) * (1 - q) / (1 - qD);
      fimRitmo = ch === 0 ? A + ritmo * D : A * qD + ritmo * (1 - qD) / (1 - q);
    }
    return { A: A, r: r, perdaDia: mantem, perdaMes: A * c, mantem: mantem, D: D, vNecessario: v, ritmo: ritmo, fimRitmo: fimRitmo, totalVendas: v == null ? null : Math.max(0, v) * D, fim: fim, alvo: alvo };
  }
  function htmlMeta(res, cfg) {
    if (!res || !res.ativos) return '<div class="empty">Carregando os números das bases…</div>';
    const x = calcMeta(res, +cfg.alvo, cfg.mes); if (!x) return '<div class="hint">Informe a meta e o mês.</div>';
    const n = v => Math.round(+v || 0).toLocaleString("pt-BR"), n1 = v => (Math.round((+v || 0) * 10) / 10).toLocaleString("pt-BR");
    const k = (l, v, note, cls) => `<div class="card kpi"><div class="lbl">${l}</div><div class="val ${cls || ""}">${v}</div><div class="note">${note || ""}</div></div>`;
    const dataFim = x.fim.toLocaleDateString("pt-BR", { timeZone: "UTC" });
    let sit;
    if (x.D <= 0) sit = '<div class="banner" style="display:block">Esse mês já passou. Escolha o mês atual ou um mês futuro.</div>';
    else if (x.alvo <= x.A && x.vNecessario <= 0) sit = `<div class="hint" style="margin:8px 0">✔ Você já está acima da meta. Para só manter os ${n(x.A)} clientes, precisa de cerca de <b>${n1(x.mantem)}</b> vendas por dia.</div>`;
    else {
      const dif = x.ritmo - x.vNecessario, ok = dif >= 0;
      sit = `<div class="${ok ? "hint" : "banner"}" style="${ok ? "margin:8px 0" : "display:block"}">${ok ? "✔ No ritmo atual você bate a meta." : "⚠️ Abaixo do ritmo necessário."} Hoje você vende <b>${n1(x.ritmo)}</b>/dia e precisa de <b>${n1(x.vNecessario)}</b>/dia (${ok ? "sobra" : "faltam"} <b>${n1(Math.abs(dif))}</b>/dia). No ritmo atual chega a <b>${n(x.fimRitmo)}</b> clientes em ${dataFim}.</div>`;
    }
    return `<div class="kpis">${k("Clientes ativos hoje", n(x.A), "Base Darb Oficial")}${k("Renovação do mês", n1(x.r * 100) + "%", "perde " + n1((1 - x.r) * 100) + "% por mês")}
      ${k("Perda por dia", n1(x.perdaDia), "≈ " + n(x.perdaMes) + " clientes por mês")}${k("Vendas/dia para não cair", n1(x.mantem), "só para manter a base")}</div>
      <div class="kpis">${k("Meta", n(x.alvo), "até " + dataFim)}${k("Vendas/dia para chegar lá", x.vNecessario == null ? "—" : n1(Math.max(0, x.vNecessario)), x.D > 0 ? x.D + " dia(s) restantes" : "")}
      ${k("Total de vendas necessárias", x.totalVendas == null ? "—" : n(x.totalVendas), "até a data")}${k("Ritmo atual", n1(x.ritmo), "vendas/dia neste mês")}</div>${sit}
      <div class="hint">Conta simples: a perda usa a renovação do mês até agora e supõe que todo cliente renova uma vez por mês. É uma estimativa para acompanhar o ritmo.</div>`;
  }

  // Botão "Puxar do Kommo": dispara e acompanha até terminar (a puxada leva ~1-2 min)
  async function puxarKommo(api, btn, aviso, depois) {
    btn.disabled = true; btn.textContent = "Puxando do Kommo…"; aviso("Puxando as bases do Kommo. Pode levar alguns minutos (ele trabalha em partes)…");
    try { await api("puxarKommo"); } catch (e) { /* pode estourar o tempo da requisição; a puxada continua no servidor */ }
    const t0 = Date.now();
    const espera = async () => {
      let st = null; try { st = await api("kommoStatus"); } catch (e) {}
      if (st && st.puxando && st.andamento) aviso("Puxando do Kommo: " + st.andamento + "…");
      if (st && !st.puxando) {
        btn.disabled = false; btn.textContent = "⇣ Puxar do Kommo";
        aviso(st.erro ? "Erro ao puxar do Kommo: " + st.erro : "");
        depois(); return;
      }
      if (Date.now() - t0 > 25 * 60 * 1000) { btn.disabled = false; btn.textContent = "⇣ Puxar do Kommo"; aviso("A puxada está demorando. Tente recarregar em alguns minutos."); return; }
      setTimeout(espera, 10000);
    };
    setTimeout(espera, 5000);
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
    puxarKommo, vId, filtrarErros, kpisValidacao, htmlListaValidacao, htmlPorVendedor, htmlQualidadeVendedor, htmlAtivos, htmlDashboard, htmlMeta, calcMeta, htmlBasesInfo, infoSaldo, calcSaldo, htmlCredito };
})();
