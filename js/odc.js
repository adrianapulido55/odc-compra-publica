(function () {
  "use strict";

  /* ================= Configuración ================= */
  // Enlace del formulario de Microsoft Forms para el reto final (Pantalla 16)
  const FORMS_URL = "https://forms.cloud.microsoft/r/ngXH5SQR55";

  /* ================= Utilidades ================= */
  const $ = (s, el = document) => el.querySelector(s);
  const h = (tag, attrs = {}, ...kids) => {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (k === "class") el.className = v;
      else if (k === "style") el.style.cssText = v;
      else if (k.startsWith("on")) el.addEventListener(k.slice(2), v);
      else if (v !== false && v != null) el.setAttribute(k, v === true ? "" : v);
    }
    kids.flat(Infinity).forEach(k => k != null && k !== false && el.append(k.nodeType ? k : document.createTextNode(k)));
    return el;
  };
  const icon = (d, cls) => { const s = document.createElementNS("http://www.w3.org/2000/svg", "svg"); s.setAttribute("viewBox", "0 0 24 24"); s.setAttribute("aria-hidden", "true"); if (cls) s.setAttribute("class", cls); s.innerHTML = d; return s; };
  const I = {
    play: '<path d="M8 5v14l11-7z"/>', pause: '<path d="M7 5h4v14H7zM13 5h4v14h-4z"/>',
    restart: '<path d="M4 4v6h6"/><path d="M5.5 15a7 7 0 1 0 1.6-7.3L4 10"/>',
    chev: '<path d="M9 6l6 6-6 6"/>', check: '<path d="M5 12.5l4.5 4.5L19 7.5"/>', x: '<path d="M6 6l12 12M18 6L6 18"/>',
    doc: '<path d="M7 3h7l5 5v13H7z"/><path d="M14 3v5h5M10 13h6M10 17h6"/>', task: '<path d="M9 11l2 2 4-4"/><rect x="4" y="4" width="16" height="16" rx="3"/>',
    up: '<path d="M12 19V5M6 11l6-6 6 6"/>', down: '<path d="M12 5v14M6 13l6 6 6-6"/>', grip: '<path d="M9 6h.01M15 6h.01M9 12h.01M15 12h.01M9 18h.01M15 18h.01"/>',
    flip: '<path d="M4 12a8 8 0 0 1 14-5.3L20 9"/><path d="M20 4v5h-5"/><path d="M20 12a8 8 0 0 1-14 5.3L4 15"/><path d="M4 20v-5h5"/>',
    plus: '<path d="M12 5v14M5 12h14"/>', ext: '<path d="M14 4h6v6M20 4l-9 9"/><path d="M18 14v5H5V6h5"/>', menu: '<path d="M4 6h7v5H4zM13 6h7v5h-7zM4 13h7v5H4zM13 13h7v5h-7z"/>',
    send: '<path d="M4 12l16-8-6 16-3-7z"/>', info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7.5v.01"/>', net: '<circle cx="12" cy="12" r="3"/><circle cx="4" cy="5" r="2"/><circle cx="20" cy="5" r="2"/><circle cx="4" cy="19" r="2"/><circle cx="20" cy="19" r="2"/><path d="M6 6l4 4M18 6l-4 4M6 18l4-4M18 18l-4-4"/>'
  };
  const store = {
    get(k) { try { return JSON.parse(localStorage.getItem("odc-cpe:" + k)); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem("odc-cpe:" + k, JSON.stringify(v)); } catch (e) {} }
  };
  const lb = (pair, cls) => h("p", { class: cls || "" }, h("strong", {}, pair[0]), " " + pair[1]);
  const refNode = (segs, cls) => h("p", { class: cls }, segs.map(([t, it]) => it ? h("em", {}, t) : linkify(t)));
  const linkify = t => { const m = t.match(/https?:\/\/\S+$/); return m ? [t.slice(0, m.index), h("a", { href: m[0], target: "_blank", rel: "noopener" }, m[0])] : [t]; };
  const words = t => (t.trim().match(/\S+/g) || []).length;

  /* ================= SCORM 1.2 (opcional) ================= */
  const SCORM = (() => {
    let api = null;
    const find = w => { let n = 0; while (w && !w.API && w.parent && w.parent !== w && n++ < 10) w = w.parent; return w && w.API ? w.API : null; };
    try { api = find(window) || (window.opener && find(window.opener)); } catch (e) { api = null; }
    const call = (fn, ...a) => { try { return api ? api[fn](...a) : ""; } catch (e) { return ""; } };
    if (api) {
      call("LMSInitialize", "");
      if (call("LMSGetValue", "cmi.core.lesson_status") === "not attempted") call("LMSSetValue", "cmi.core.lesson_status", "incomplete");
      window.addEventListener("beforeunload", () => { call("LMSCommit", ""); call("LMSFinish", ""); });
    }
    return {
      location(v) { call("LMSSetValue", "cmi.core.lesson_location", String(v)); call("LMSCommit", ""); },
      score(raw, max) { call("LMSSetValue", "cmi.core.score.raw", String(raw)); call("LMSSetValue", "cmi.core.score.min", "0"); call("LMSSetValue", "cmi.core.score.max", String(max)); call("LMSCommit", ""); },
      status(s) { call("LMSSetValue", "cmi.core.lesson_status", s); call("LMSCommit", ""); }
    };
  })();

  /* ================= Modal ================= */
  const modal = $("#modal");
  let lastFocus = null;
  function openModal(title, nodes, wide) {
    lastFocus = document.activeElement;
    $("#modalTitle").textContent = title;
    const body = $("#modalBody"); body.innerHTML = ""; body.append(...[nodes].flat());
    modal.querySelector(".modal__card").classList.toggle("modal__card--wide", !!wide);
    modal.hidden = false; $("#modalClose").focus();
  }
  function closeModal() { modal.hidden = true; $("#modalBody").innerHTML = ""; lastFocus && lastFocus.focus(); }
  $("#modalClose").addEventListener("click", closeModal);
  modal.addEventListener("click", e => { if (e.target === modal) closeModal(); });
  const fuenteModal = (title, paras, refs) => openModal(title, [
    ...paras.map(p => h("p", {}, p)),
    ...(refs || []).map(r => refNode(r, "modal__ref"))
  ]);

  /* ================= Avatar Licita ================= */
  let activeVoice = null;
  const stopVoice = () => { if (activeVoice) activeVoice.stop(); if ("speechSynthesis" in window) speechSynthesis.cancel(); };

  function licita(fileId, guion, opts = {}) {
    const file = (window.AUDIO_DATA && window.AUDIO_DATA[fileId]) || "audio/licita-pantalla-" + fileId + ".mp3";
    const paras = guion.map(t => h("p", {}, t));
    const playBtn = h("button", { class: "ctrl ctrl--play", "aria-label": "Reproducir" }, icon(I.play));
    const restartBtn = h("button", { class: "ctrl", "aria-label": "Reiniciar" }, icon(I.restart));
    const card = h("section", { class: "licita" + (opts.big ? " licita--big" : "") + (opts.mini ? " licita--mini" : ""), "aria-label": "Licita" },
      h("div", { class: "licita__top" },
        h("div", { class: "licita__img" }, h("img", { src: "img/licita.jpg", alt: "Licita, avatar del ODC: una lupa azul sonriente con mango naranja." })),
        h("div", {},
          h("div", { class: "licita__name" }, opts.titulo || "Licita"),
          h("div", { class: "licita__controls" }, playBtn, restartBtn, h("span", { class: "wave", "aria-hidden": "true" }, h("i"), h("i"), h("i"), h("i"))))),
      h("details", { class: "transcript" }, h("summary", {}, icon(I.chev), "Transcripción"), h("div", { class: "transcript__body" }, paras)));

    const audio = new Audio(); audio.preload = "none";
    let useTTS = false, playing = false, idx = 0;
    const lens = guion.map(t => t.length), total = lens.reduce((a, b) => a + b, 0);
    const setUI = on => {
      playing = on; card.classList.toggle("is-talking", on);
      playBtn.replaceChildren(icon(on ? I.pause : I.play)); playBtn.setAttribute("aria-label", on ? "Pausar" : "Reproducir");
      if (opts.hotspot) opts.hotspot.classList.toggle("is-on", on);
    };
    const hl = i => paras.forEach((p, j) => p.classList.toggle("is-reading", j === i));
    audio.addEventListener("timeupdate", () => {
      if (!audio.duration) return;
      const pos = audio.currentTime / audio.duration * total; let acc = 0, i = 0;
      for (; i < lens.length; i++) { acc += lens[i]; if (pos <= acc) break; }
      hl(Math.min(i, lens.length - 1));
    });
    audio.addEventListener("ended", () => { setUI(false); hl(-1); });
    function tts(from) {
      if (!("speechSynthesis" in window)) { setUI(false); return; }
      speechSynthesis.cancel(); idx = from;
      const vs = speechSynthesis.getVoices();
      const v = vs.find(x => /es[-_](CO|MX|US|419)/i.test(x.lang)) || vs.find(x => /^es/i.test(x.lang));
      const next = () => {
        if (!playing) return;
        if (idx >= guion.length) { setUI(false); hl(-1); idx = 0; return; }
        const u = new SpeechSynthesisUtterance(guion[idx]); u.lang = v ? v.lang : "es-CO"; if (v) u.voice = v;
        u.onend = () => { idx++; next(); };
        hl(idx); speechSynthesis.speak(u);
      };
      next();
    }
    const ctl = {
      play() {
        stopVoice(); activeVoice = ctl; setUI(true);
        if (useTTS) { tts(idx); return; }
        if (!audio.src) audio.src = file;
        audio.play().catch(() => { if (playing) { useTTS = true; tts(0); } });
      },
      pause() { setUI(false); if (useTTS) speechSynthesis.cancel(); else audio.pause(); },
      stop() { ctl.pause(); },
      restart() { if (useTTS) { idx = 0; ctl.play(); } else { audio.currentTime = 0; ctl.play(); } },
      toggle() { playing ? ctl.pause() : ctl.play(); }
    };
    audio.addEventListener("error", () => { if (playing && !useTTS) { useTTS = true; tts(0); } });
    playBtn.addEventListener("click", ctl.toggle);
    restartBtn.addEventListener("click", ctl.restart);
    card._ctl = ctl;
    return card;
  }

  /* ================= Bloques comunes ================= */
  const M = ODC.menu.items;
  const done = id => !!store.get("hecho-" + id);
  function markDone(id) { store.set("hecho-" + id, true); buildMenu(); }

  function retroBox(ok, pair, extra) {
    return h("div", { class: "retro " + (ok ? "is-ok" : "is-bad") },
      h("span", { class: "retro__icon", "aria-hidden": "true" }, icon(ok ? I.check : I.x)),
      h("div", {}, lb(pair), extra || null));
  }

  function actividadShell(a, body, check, onReset) {
    const box = h("section", { class: "actividad", "aria-live": "polite" },
      h("span", { class: "actividad__tag" }, icon(I.task), a.etiqueta));
    let retro = null;
    const api = {
      show(ok) {
        retro && retro.remove();
        retro = retroBox(ok, a.retro, ok ? null : h("button", { class: "reintentar", onclick: () => { retro.remove(); retro = null; onReset && onReset(); } }, "Intentar de nuevo"));
        box.append(retro);
        if (ok) markDone(current);
      }
    };
    box.append(...[body].flat());
    if (check) box.append(h("div", { class: "actividad__actions" }, h("button", { class: "btn btn--primary", onclick: () => api.show(check()) }, icon(I.check), "Comprobar")));
    box._api = api;
    return box;
  }

  function actABC(a) {
    const btns = a.opciones.map((t, i) => h("button", { class: "opcion", onclick: () => answer(i) }, h("span", {}, t), h("span", { class: "marca", "aria-hidden": "true" })));
    const reset = () => btns.forEach(b => { b.disabled = false; b.classList.remove("is-ok", "is-bad"); b.querySelector(".marca").replaceChildren(); });
    const box = actividadShell(a, [h("p", { class: "actividad__q" }, a.pregunta), h("div", { class: "opciones" }, btns)], null, reset);
    function answer(i) {
      const ok = i === a.correcta;
      btns.forEach((b, j) => {
        b.disabled = true;
        if (j === a.correcta) { b.classList.add("is-ok"); b.querySelector(".marca").replaceChildren(icon(I.check)); }
        if (j === i && !ok) { b.classList.add("is-bad"); b.querySelector(".marca").replaceChildren(icon(I.x)); }
      });
      box._api.show(ok);
    }
    return box;
  }

  function actVF(a) {
    const vals = [true, false];
    const btns = a.botones.map((t, i) => h("button", { class: "vf-btn", onclick: () => answer(i) }, t));
    const reset = () => btns.forEach(b => { b.disabled = false; b.classList.remove("is-ok", "is-bad"); });
    const box = actividadShell(a, [h("p", { class: "actividad__q" }, a.pregunta), h("div", { class: "vf" }, btns)], null, reset);
    function answer(i) {
      const ok = vals[i] === a.correcta;
      btns.forEach((b, j) => { b.disabled = true; if (vals[j] === a.correcta) b.classList.add("is-ok"); else if (j === i) b.classList.add("is-bad"); });
      box._api.show(ok);
    }
    return box;
  }

  // Selección por toque/clic + arrastrar y soltar
  function dragSource(el, data) {
    el.draggable = true;
    el.addEventListener("dragstart", e => { e.dataTransfer.setData("text/plain", data); e.dataTransfer.effectAllowed = "move"; el.classList.add("is-dragging"); });
    el.addEventListener("dragend", () => el.classList.remove("is-dragging"));
  }
  function dropTarget(el, fn) {
    el.addEventListener("dragover", e => { e.preventDefault(); el.classList.add("is-over"); });
    el.addEventListener("dragleave", () => el.classList.remove("is-over"));
    el.addEventListener("drop", e => { e.preventDefault(); el.classList.remove("is-over"); fn(e.dataTransfer.getData("text/plain")); });
  }

  function actClasificar(a) {
    let sel = null;
    const place = {};
    const pool = h("div", { class: "pool" });
    const chips = a.fichas.map((t, i) => {
      const c = h("button", { class: "chip", "data-i": i, onclick: () => { sel = sel === i ? null : i; chips.forEach((x, j) => x.classList.toggle("is-sel", j === sel)); } }, icon(I.grip), t);
      dragSource(c, String(i)); return c;
    });
    pool.append(...chips);
    const drop = (bi, i) => { i = +i; place[i] = bi; bins[bi].querySelector(".bin__drop").append(chips[i]); chips[i].classList.remove("is-sel", "is-ok", "is-bad"); sel = null; };
    const bins = a.cajas.map((t, bi) => {
      const b = h("div", { class: "bin", role: "group", "aria-label": t, onclick: e => { if (sel != null && !e.target.closest(".chip")) drop(bi, sel); } },
        h("div", { class: "bin__title" }, t), h("div", { class: "bin__drop" }));
      dropTarget(b, i => drop(bi, i)); return b;
    });
    const hint = h("p", { class: "hint" }, icon(I.info), "Arrastra o toca una tarjeta y luego la casilla.");
    const reset = () => { Object.keys(place).forEach(k => delete place[k]); chips.forEach(c => { c.classList.remove("is-ok", "is-bad"); pool.append(c); }); };
    return actividadShell(a, [h("p", { class: "actividad__q" }, a.pregunta), hint, pool, h("div", { class: "bins" }, bins)], () => {
      let ok = true;
      chips.forEach((c, i) => { const good = place[i] === a.correctas[i]; c.classList.toggle("is-ok", good); c.classList.toggle("is-bad", !good); if (!good) ok = false; });
      return ok;
    }, reset);
  }

  function actOrdenar(a) {
    let order = a.fichas.map((_, i) => i);
    const list = h("ol", { class: "orden" });
    let dragFrom = null;
    const draw = () => {
      list.replaceChildren(...order.map((fi, pos) => {
        const li = h("li", { class: "orden__item", draggable: "true" },
          h("span", { class: "orden__n" }, String(pos + 1)), h("span", { class: "orden__t" }, a.fichas[fi]),
          h("span", { class: "orden__btns" },
            h("button", { class: "mini", "aria-label": "Subir", disabled: pos === 0, onclick: () => move(pos, -1) }, icon(I.up)),
            h("button", { class: "mini", "aria-label": "Bajar", disabled: pos === order.length - 1, onclick: () => move(pos, 1) }, icon(I.down))));
        li.addEventListener("dragstart", () => { dragFrom = pos; li.classList.add("is-dragging"); });
        li.addEventListener("dragend", () => li.classList.remove("is-dragging"));
        li.addEventListener("dragover", e => e.preventDefault());
        li.addEventListener("drop", e => { e.preventDefault(); if (dragFrom == null) return; const [x] = order.splice(dragFrom, 1); order.splice(pos, 0, x); dragFrom = null; draw(); });
        return li;
      }));
    };
    const move = (p, d) => { [order[p], order[p + d]] = [order[p + d], order[p]]; draw(); list.querySelectorAll(".mini")[p * 2 + (d > 0 ? 3 : -2)]?.focus(); };
    draw();
    return actividadShell(a, [h("p", { class: "actividad__q" }, a.pregunta), list], () => {
      let ok = true;
      [...list.children].forEach((li, p) => { const g = order[p] === a.orden[p]; li.classList.toggle("is-ok", g); li.classList.toggle("is-bad", !g); if (!g) ok = false; });
      return ok;
    }, () => [...list.children].forEach(li => li.classList.remove("is-ok", "is-bad")));
  }

  function actCompletar(a) {
    const mk = () => h("select", { class: "blank" }, h("option", { value: "" }, "…"), a.opciones.map(o => h("option", { value: o }, o)));
    const s1 = mk(), s2 = mk();
    const frase = h("p", { class: "actividad__q frase" }, a.partes[0], s1, a.partes[1], s2, a.partes[2]);
    return actividadShell(a, [frase], () => {
      const g1 = s1.value === a.correctas[0], g2 = s2.value === a.correctas[1];
      s1.className = "blank " + (g1 ? "is-ok" : "is-bad"); s2.className = "blank " + (g2 ? "is-ok" : "is-bad");
      return g1 && g2;
    }, () => { s1.className = s2.className = "blank"; });
  }

  function actRelacionar(a) {
    const sels = a.leyes.map(() => h("select", { class: "blank" }, h("option", { value: "" }, "…"), a.opciones.map((o, i) => h("option", { value: i }, o))));
    const rows = a.leyes.map((l, i) => h("div", { class: "rel" }, h("span", { class: "rel__ley" }, l), h("span", { class: "rel__arrow", "aria-hidden": "true" }, "→"), sels[i]));
    rows.forEach((r, i) => sels[i].setAttribute("aria-label", a.leyes[i]));
    return actividadShell(a, [h("p", { class: "actividad__q" }, a.pregunta), h("div", { class: "rels" }, rows)], () => {
      let ok = true;
      sels.forEach((s, i) => { const g = s.value !== "" && +s.value === a.correctas[i]; s.className = "blank " + (g ? "is-ok" : "is-bad"); if (!g) ok = false; });
      return ok;
    }, () => sels.forEach(s => s.className = "blank"));
  }

  const ACT = { abc: actABC, vf: actVF, clasificar: actClasificar, ordenar: actOrdenar, completar: actCompletar, relacionar: actRelacionar };

  function lectura(c) {
    const pop = h("p", { class: "pop", hidden: true }, c.info);
    const btn = h("button", { class: "btn-info", "aria-expanded": "false" }, icon(I.plus), "info");
    btn.addEventListener("click", () => { pop.hidden = !pop.hidden; btn.setAttribute("aria-expanded", String(!pop.hidden)); });
    return h("section", { class: "lectura" },
      c.contenido.map(t => /^Síntesis del equipo:/.test(t) ? lb(t.split(/(?<=:) /), "sintesis") : h("p", {}, t)),
      h("div", { class: "info-row" }, btn), pop);
  }

  function moduloScreen(n, c, visual) {
    const mi = M.findIndex(x => x.pantalla === "Pantalla " + n);
    return h("section", { class: "screen mod" },
      h("div", { class: "mod__main" },
        h("header", { class: "mod__head" },
          h("div", { class: "kicker" }, M[mi].nombre + " · " + M[mi].titulo),
          h("h1", {}, c.titulo),
          lb(c.proposito, "proposito")),
        visual, lectura(c), ACT[c.actividad.tipo](c.actividad)),
      h("aside", { class: "mod__side" }, licita(String(n).padStart(2, "0"), c.guion)));
  }

  /* ================= Pantallas ================= */
  const R = {};

  R[1] = () => {
    const c = ODC.p1;
    const hs = (act, label, x, fn) => h("button", { class: "hotspot", "data-act": act, "aria-label": label, style: `left:${x}%;top:89.8%;width:16%;height:8.4%`, onclick: fn });
    const audioHs = hs("audio", "Audio", 82.3, () => lic._ctl.toggle());
    const lic = licita("01", c.guion, { big: true, hotspot: audioHs });
    const i = c.autoria.indexOf(":");
    return h("section", { class: "screen cover" },
      h("div", { class: "cover__hero" },
        h("img", { src: "img/portada.jpg", alt: "Mapa de Colombia con departamentos conectados por líneas digitales hacia una plataforma (SECOP II)" }),
        hs("start", "Iniciar", 41.8, () => go(2)), hs("help", "Ayuda", 64.2, showHelp), audioHs),
      h("div", { class: "cover__text" },
        h("div", { class: "kicker" }, c.kicker), h("h1", {}, c.titulo),
        h("p", { class: "cover__desc" }, c.descripcion),
        h("p", { class: "cover__autoria" }, h("strong", {}, c.autoria.slice(0, i + 1)), c.autoria.slice(i + 1))),
      h("div", { class: "cover__side" }, lic));
  };
  function showHelp() { openModal("Ayuda", ODC.ayuda.map(t => lb(t.split(/(?<=:) /)))); }

  R[2] = () => {
    const c = ODC.p2;
    const blk = (b, id, cls) => h("article", { class: "rea-block " + cls },
      h("h2", {}, b.titulo), h("p", { class: "rea-block__txt" }, b.texto),
      licita(id, b.guion, { mini: true, titulo: b.audio }));
    return h("section", { class: "screen" },
      h("div", { class: "rea-grid" }, blk(c.rea, "02a", "rea-block--rea"), blk(c.pob, "02b", "rea-block--pob")),
      h("section", { class: "lectura antes" }, h("h2", {}, c.antes.titulo), c.antes.items.map(p => lb(p))),
      h("div", { class: "center" }, h("button", { class: "btn btn--primary btn--lg", onclick: () => go(3) }, icon(I.menu), "Ver menú")));
  };

  R[3] = () => {
    const c = ODC.p3;
    const grupos = [[0, 4], [4, 8], [8, 13]];
    const partes = ODC.menu.partes.map((p, gi) => h("div", { class: "mpart" },
      h("div", { class: "mpart__head" }, h("strong", {}, p.nombre), h("span", {}, p.desc)),
      h("div", { class: "mpart__grid" }, M.slice(...grupos[gi]).map(it => {
        const t = targetOf(it);
        return h("button", { class: "mbtn" + (done(t) ? " is-done" : ""), onclick: () => go(t) },
          done(t) ? h("span", { class: "mbtn__badge", "aria-label": "Completado" }, icon(I.check)) : null,
          h("span", { class: "mbtn__k" }, it.nombre), h("span", { class: "mbtn__t" }, it.titulo), h("span", { class: "mbtn__p" }, it.pantalla));
      }))));
    return h("section", { class: "screen mod" },
      h("div", { class: "mod__main" }, h("header", { class: "mod__head" }, h("h1", {}, c.titulo)), partes,
        h("section", { class: "lectura" }, c.texto.map(t => /^(Parte|Navegación|Recomendación)/.test(t) && t.includes(": ") ? lb(t.split(/(?<=:) /)) : h("p", {}, t)))),
      h("aside", { class: "mod__side" }, licita("03", c.guion)));
  };

  R[4] = () => {
    const c = ODC.p4;
    return moduloScreen(4, c, [
      h("div", { class: "alertas" }, c.alertas.map(a => h("article", { class: "alerta" }, h("h3", {}, a.titulo), h("p", {}, a.texto)))),
      h("div", { class: "cifras cifras--2" }, c.cifras.map(f => h("article", { class: "cifra" },
        h("div", { class: "cifra__valor" }, f.valor), h("p", { class: "cifra__texto" }, f.texto),
        h("button", { class: "btn-fuente", onclick: () => fuenteModal("Ver fuente", [f.fuente[0]], [f.fuente[1]]) }, icon(I.doc), "Ver fuente"))))]);
  };

  R[5] = () => {
    const c = ODC.p5;
    const side = (t, cls) => h("div", { class: "cmp__side " + cls },
      h("div", { class: "cmp__name" }, t[0]), h("div", { class: "cmp__kind" }, t[1]), h("ul", {}, t.slice(2).map(x => h("li", {}, x))));
    const front = side(c.comparador[1], "cmp__side--ii");
    const range = h("input", { type: "range", min: "0", max: "100", value: "50", class: "cmp__range", "aria-label": c.comparador[0][0] + " · " + c.comparador[1][0] });
    const wrap = h("div", { class: "cmp" }, side(c.comparador[0], "cmp__side--i"), front, h("div", { class: "cmp__handle", "aria-hidden": "true" }, h("span", {}, "‹ ›")), range);
    const upd = () => { wrap.style.setProperty("--pos", range.value + "%"); };
    range.addEventListener("input", upd); upd();
    return moduloScreen(5, c, wrap);
  };

  R[6] = () => {
    const c = ODC.p6;
    const det = h("p", { class: "embudo__det", "aria-live": "polite" });
    const widths = [100, 82, 64, 46];
    const bars = c.embudo.map((e, i) => {
      const b = h("button", { class: "embudo__bar", style: `--w:${widths[i]}%`, onclick: () => {
        bars.forEach((x, j) => x.classList.toggle("is-on", j <= i)); bars.forEach((x, j) => x.setAttribute("aria-pressed", String(j === i)));
        det.textContent = e.detalle; det.classList.add("is-shown");
      } }, h("span", {}, e.texto));
      return b;
    });
    return moduloScreen(6, c, h("div", { class: "embudo-wrap" },
      h("div", { class: "embudo" }, bars, det),
      h("div", { class: "fichas" }, c.fichas.map(f => h("article", { class: "ficha" }, h("strong", {}, f[0]), h("span", {}, f[1]))))));
  };

  R[7] = () => {
    const c = ODC.p7;
    const max = Math.max(...c.regiones.map(r => +r.n));
    const rows = c.regiones.map(r => {
      const d = h("p", { class: "barra__det", hidden: true }, r.detalle);
      const b = h("button", { class: "barra", "aria-expanded": "false", onclick: () => { d.hidden = !d.hidden; b.setAttribute("aria-expanded", String(!d.hidden)); } },
        h("span", { class: "barra__lbl" }, r.region),
        h("span", { class: "barra__track" }, h("span", { class: "barra__fill", style: `--w:${(+r.n / max) * 100}%` }, h("b", {}, r.n))),
        icon(I.chev, "barra__chev"));
      return h("div", { class: "barra-row" }, b, d);
    });
    return moduloScreen(7, c, h("div", { class: "evid" },
      h("div", { class: "barras" }, rows),
      h("aside", { class: "vacio" }, h("h3", {}, icon(I.info), c.vacio[0]), h("p", {}, c.vacio[1]))));
  };

  R[8] = () => {
    const c = ODC.p8;
    const cards = c.cifras.map(f => {
      const card = h("button", { class: "flip", "aria-pressed": "false", onclick: () => { card.classList.toggle("is-flipped"); card.setAttribute("aria-pressed", String(card.classList.contains("is-flipped"))); } },
        h("span", { class: "flip__in" },
          h("span", { class: "flip__front" }, h("span", { class: "cifra__valor" }, f.valor), h("span", { class: "cifra__texto" }, f.texto), h("span", { class: "flip__hint" }, icon(I.flip))),
          h("span", { class: "flip__back" }, h("span", {}, f.fuente), h("span", { class: "flip__hint" }, icon(I.flip)))));
      return card;
    });
    return moduloScreen(8, c, [
      h("div", { class: "flips" }, cards),
      h("div", { class: "banner" }, h("strong", {}, c.banner[0]),
        h("button", { class: "btn btn--ghost", onclick: () => fuenteModal("Profundizar", c.profundizar, c.refs) }, icon(I.doc), "Profundizar"))]);
  };

  R[9] = () => {
    const c = ODC.p9;
    return moduloScreen(9, c, [
      h("div", { class: "cifras cifras--3" }, c.cifras.map(f => h("article", { class: "cifra" }, h("div", { class: "cifra__valor" }, f.valor), h("p", { class: "cifra__texto" }, f.texto)))),
      h("div", { class: "banner" }, h("div", {}, h("strong", {}, c.banner[0]), h("p", {}, c.banner[1])),
        h("button", { class: "btn btn--primary", onclick: () => openModal(c.hubs.boton, [redHubs(), h("p", {}, c.hubs.texto)], true) }, icon(I.net), c.hubs.boton))]);
  };

  function redHubs() {
    // Representación didáctica: la mayoría de contratistas con una sola línea; pocos «hubs» con muchas
    const W = 560, H = 320, ns = "http://www.w3.org/2000/svg";
    const svg = document.createElementNS(ns, "svg");
    svg.setAttribute("viewBox", `0 0 ${W} ${H}`); svg.setAttribute("class", "red"); svg.setAttribute("role", "img");
    svg.setAttribute("aria-label", ODC.p9.hubs.texto);
    let seed = 7; const rnd = () => (seed = (seed * 9301 + 49297) % 233280) / 233280;
    const hubs = [[150, 120], [380, 190], [270, 80]];
    const ent = []; for (let i = 0; i < 9; i++) ent.push([40 + rnd() * 480, 30 + rnd() * 260]);
    const g = document.createElementNS(ns, "g"); svg.append(g);
    const line = (a, b, cls, d) => { const l = document.createElementNS(ns, "line"); l.setAttribute("x1", a[0]); l.setAttribute("y1", a[1]); l.setAttribute("x2", b[0]); l.setAttribute("y2", b[1]); l.setAttribute("class", cls); l.style.animationDelay = d + "s"; g.append(l); };
    const dot = (p, r, cls, d) => { const c = document.createElementNS(ns, "circle"); c.setAttribute("cx", p[0]); c.setAttribute("cy", p[1]); c.setAttribute("r", r); c.setAttribute("class", cls); c.style.animationDelay = d + "s"; svg.append(c); };
    let k = 0;
    hubs.forEach(hb => { for (let i = 0; i < 16; i++) { const a = rnd() * Math.PI * 2, r = 45 + rnd() * 70; const p = [hb[0] + Math.cos(a) * r, hb[1] + Math.sin(a) * r * .75]; line(hb, p, "red__l red__l--hub", (k++) * .03); dot(p, 3.2, "red__n", k * .03); } });
    for (let i = 0; i < 26; i++) { const a = [20 + rnd() * 520, 15 + rnd() * 290]; const e = ent[Math.floor(rnd() * ent.length)]; line(e, a, "red__l", (k++) * .03); dot(a, 3.2, "red__n red__n--uno", k * .03); }
    ent.forEach(e => dot(e, 4.5, "red__e", 0));
    hubs.forEach(hb => dot(hb, 11, "red__hub", .2));
    return svg;
  }

  R[10] = () => {
    const c = ODC.p10;
    const map = [0, 0, 0, 1, 2];
    const panel = h("div", { class: "flujo__panel", "aria-live": "polite" });
    const steps = c.pasos.map((p, i) => h("button", { class: "flujo__step", "aria-pressed": "false", onclick: () => {
      steps.forEach((s, j) => { s.classList.toggle("is-on", j === i); s.setAttribute("aria-pressed", String(j === i)); });
      const f = c.fuentes[map[i]];
      panel.replaceChildren(h("p", { class: "flujo__ley" }, p.leyenda),
        h("button", { class: "btn btn--ghost btn--sm", onclick: () => fuenteModal("Ver fuente", [f[0]], [f[1]]) }, icon(I.doc), "Ver fuente"));
      panel.classList.add("is-shown");
    } }, h("span", {}, p.paso)));
    return moduloScreen(10, c, [
      h("div", { class: "flujo" }, h("div", { class: "flujo__steps" }, steps), panel),
      h("div", { class: "cifras cifras--3" }, c.cifras.map(f => h("article", { class: "cifra cifra--light" }, h("div", { class: "cifra__valor" }, f.valor), h("p", { class: "cifra__texto" }, f.texto))))]);
  };

  R[11] = () => {
    const c = ODC.p11;
    const panel = h("div", { class: "formula__panel", "aria-live": "polite" });
    const pieces = [];
    const parts = [];
    c.formula.forEach((f, i) => {
      const b = h("button", { class: "pieza" + (i === 4 ? " pieza--eq" : ""), "aria-pressed": "false", onclick: () => {
        b.classList.add("is-on"); pieces.forEach(x => x.setAttribute("aria-pressed", String(x === b))); pieces.forEach(x => x.classList.toggle("is-cur", x === b));
        panel.replaceChildren(h("strong", {}, f.titulo), h("p", {}, f.estudio)); panel.classList.add("is-shown");
      } }, h("strong", {}, f.titulo), h("span", {}, f.texto));
      pieces.push(b);
      if (i > 0 && i < 4) parts.push(h("span", { class: "op", "aria-hidden": "true" }, "+"));
      parts.push(b);
    });
    return moduloScreen(11, c, [
      h("div", { class: "formula" }, parts), panel,
      h("div", { class: "efectos" }, c.efectos.map((e, i) => h("article", { class: "efecto " + (i ? "efecto--no" : "efecto--si") }, h("h3", {}, e[0]), h("p", {}, e[1])))),
      h("div", { class: "banner banner--key" }, h("strong", {}, c.banner))]);
  };

  R[12] = () => {
    const c = ODC.p12;
    const ficha = h("div", { class: "tl__ficha", "aria-live": "polite" });
    const hitos = c.hitos.map((x, i) => h("button", { class: "tl__hito", "aria-pressed": "false", onclick: () => {
      hitos.forEach((b, j) => { b.classList.toggle("is-on", j === i); b.setAttribute("aria-pressed", String(j === i)); });
      ficha.replaceChildren(h("div", { class: "tl__ficha-k" }, x.anio), h("strong", {}, x.norma), h("p", {}, x.ficha)); ficha.classList.add("is-shown");
    } }, h("span", { class: "tl__anio" }, x.anio), h("span", { class: "tl__dot", "aria-hidden": "true" }), h("span", { class: "tl__norma" }, x.norma)));
    return moduloScreen(12, c, [h("div", { class: "tl" }, h("div", { class: "tl__line" }, hitos), ficha), h("div", { class: "banner banner--key" }, h("strong", {}, c.banner))]);
  };

  /* ---------- Evaluación final ---------- */
  const evalState = () => store.get("eval") || { p13: {}, p14: {} };
  const saveEval = s => {
    store.set("eval", s);
    const n13 = Object.values(s.p13).filter(Boolean).length, n14 = Object.values(s.p14).filter(Boolean).length;
    const answered = Object.keys(s.p13).length === 3 && Object.keys(s.p14).length === 4;
    if (answered) { SCORM.score(n13 + n14, 7); if (n13 + n14 >= 6) markDoneId(13); }
  };
  const markDoneId = id => { store.set("hecho-" + id, true); buildMenu(); };

  function resultado() {
    const s = evalState();
    const n = Object.values(s.p13).filter(Boolean).length + Object.values(s.p14).filter(Boolean).length;
    const answered = Object.keys(s.p13).length === 3 && Object.keys(s.p14).length === 4;
    if (!answered) return null;
    const ok = n >= 6;
    return h("div", { class: "score " + (ok ? "is-ok" : "is-bad") },
      h("span", { class: "retro__icon", "aria-hidden": "true" }, icon(ok ? I.check : I.x)),
      h("div", { class: "score__n" }, String(n), h("small", {}, " / 7")),
      h("p", {}, ODC.p13.puntaje),
      h("button", { class: "reintentar", onclick: () => { store.set("eval", { p13: {}, p14: {} }); go(13); } }, "Intentar de nuevo"));
  }

  R[13] = () => {
    const c = ODC.p13, st = evalState();
    const repaso = h("p", { class: "repaso", hidden: Object.keys(st.p13).length < 3 }, c.repaso);
    const cards = c.items.map((it, i) => {
      const card = h("article", { class: "vfcard" });
      const btns = c.botones.map((t, j) => h("button", { class: "vf-btn", onclick: () => answer(j === 0) }, t));
      card.append(h("p", { class: "vfcard__q" }, it.texto), h("div", { class: "vf" }, btns));
      const answer = v => {
        const ok = v === it.correcta;
        btns.forEach((b, j) => { b.disabled = true; if ((j === 0) === it.correcta) b.classList.add("is-ok"); else b.classList.add("is-bad-dim"); if ((j === 0) === v && !ok) b.classList.add("is-bad"); });
        card.classList.add(ok ? "is-ok" : "is-bad");
        card.append(h("p", { class: "vfcard__retro" }, it.retro));
        const s = evalState(); s.p13[i] = ok; saveEval(s);
        if (Object.keys(s.p13).length === 3) repaso.hidden = false;
      };
      if (i in st.p13) { /* ya respondida: mostrar estado */ const v = st.p13[i] ? it.correcta : !it.correcta; setTimeout(() => answer(v), 0); }
      return card;
    });
    return h("section", { class: "screen" },
      h("header", { class: "mod__head" }, h("h1", {}, c.titulo), h("p", { class: "proposito" }, c.instruccion)),
      h("div", { class: "vfcards" }, cards), repaso);
  };

  R[14] = () => {
    const c = ODC.p14, st = evalState();
    let sel = null; const place = {};
    const chips = c.casos.map((t, i) => {
      const ch = h("button", { class: "caso", onclick: () => { sel = sel === i ? null : i; chips.forEach((x, j) => x.classList.toggle("is-sel", j === sel)); } }, icon(I.grip), t);
      dragSource(ch, String(i)); return ch;
    });
    const pool = h("div", { class: "casos" }, chips);
    const drop = (ti, i) => { i = +i; const prev = Object.keys(place).find(k => place[k] === ti && +k !== i); if (prev != null) { delete place[prev]; pool.append(chips[prev]); } place[i] = ti; targets[ti].querySelector(".hz__drop").append(chips[i]); chips[i].classList.remove("is-sel"); sel = null; };
    const targets = c.hallazgos.map((t, ti) => {
      const el = h("div", { class: "hz", onclick: e => { if (sel != null && !e.target.closest(".caso")) drop(ti, sel); } }, h("div", { class: "hz__t" }, t), h("div", { class: "hz__drop" }));
      dropTarget(el, i => drop(ti, i)); return el;
    });
    const fb = h("div", { class: "hz-fb", "aria-live": "polite" });
    const res = h("div", {});
    const check = () => {
      const s = evalState(); fb.replaceChildren();
      chips.forEach((ch, i) => {
        const ok = place[i] === c.correctas[i];
        ch.classList.toggle("is-ok", ok); ch.classList.toggle("is-bad", !ok); s.p14[i] = ok;
        fb.append(h("p", { class: ok ? "is-ok" : "is-bad" }, icon(ok ? I.check : I.x), c.retro[i]));
      });
      fb.append(h("p", { class: "repaso" }, c.repaso));
      saveEval(s); res.replaceChildren(resultado() || "");
    };
    return h("section", { class: "screen" },
      h("header", { class: "mod__head" }, h("h1", {}, c.titulo), h("p", { class: "proposito" }, c.instruccion)),
      h("p", { class: "hint" }, icon(I.info), "Arrastra o toca un caso y luego el hallazgo."),
      h("div", { class: "match" }, pool, h("div", { class: "hzs" }, targets)),
      h("div", { class: "actividad__actions" }, h("button", { class: "btn btn--primary", onclick: check }, icon(I.check), "Comprobar")),
      fb, res);
  };

  R[15] = () => {
    const c = ODC.p15;
    const niveles = c.niveles.map((n, i) => {
      const d = h("div", { class: "sem__det", hidden: true }, n.estudios.map(e => h("p", {}, e)));
      const b = h("button", { class: "btn btn--ghost btn--sm", "aria-expanded": "false", onclick: () => { d.hidden = !d.hidden; b.setAttribute("aria-expanded", String(!d.hidden)); } }, icon(I.doc), c.boton);
      return h("article", { class: "sem sem--" + i }, h("span", { class: "sem__luz", "aria-hidden": "true" }), h("div", { class: "sem__body" }, h("p", { class: "sem__t" }, n.texto), b, d));
    });
    return h("section", { class: "screen mod" },
      h("div", { class: "mod__main" }, h("header", { class: "mod__head" }, h("h1", {}, c.titulo)), h("div", { class: "semaforo" }, niveles)),
      h("aside", { class: "mod__side" }, licita("15", c.guion)));
  };

  R[16] = () => {
    const c = ODC.p16;
    const saved = store.get("reto") || {};
    let elegido = saved.i ?? null;
    const det = h("p", { class: "reco__det", "aria-live": "polite" });
    const cards = c.tarjetas.map((t, i) => h("button", { class: "reco", role: "radio", "aria-checked": "false", onclick: () => pick(i) }, h("strong", {}, t.titulo), h("span", {}, t.texto)));
    const pick = i => { elegido = i; cards.forEach((x, j) => { x.classList.toggle("is-sel", j === i); x.setAttribute("aria-checked", String(j === i)); }); det.textContent = c.tarjetas[i].detalle; det.classList.add("is-shown"); persist(); };
    const ta = h("textarea", { class: "reto__txt", rows: "6", "aria-label": c.reto }, saved.t || "");
    const cnt = h("span", { class: "reto__cnt" });
    const upd = () => { const n = words(ta.value); cnt.textContent = n + " / 100"; cnt.classList.toggle("is-over", n > 100); persist(); };
    const persist = () => store.set("reto", { i: elegido, t: ta.value });
    ta.addEventListener("input", upd);
    const rub = h("div", { class: "rubrica", hidden: true }, icon(I.check), h("p", {}, c.rubrica));
    const send = () => {
      if (elegido == null) { cards[0].focus(); cards.forEach(x => x.classList.add("is-nudge")); setTimeout(() => cards.forEach(x => x.classList.remove("is-nudge")), 600); return; }
      if (!ta.value.trim() || words(ta.value) > 100) { ta.focus(); return; }
      const txt = c.tarjetas[elegido].titulo + " " + c.tarjetas[elegido].texto + "\n\n" + ta.value.trim();
      try { navigator.clipboard && navigator.clipboard.writeText(txt).catch(() => {}); } catch (e) {}
      if (FORMS_URL) window.open(FORMS_URL, "_blank", "noopener");
      rub.hidden = false; markDone(16); SCORM.status("completed");
    };
    const checks = c.lista[1].map(t => h("label", { class: "check" }, h("input", { type: "checkbox" }), h("span", {}, t)));
    setTimeout(() => { if (elegido != null) pick(elegido); upd(); }, 0);
    return h("section", { class: "screen mod" },
      h("div", { class: "mod__main" },
        h("header", { class: "mod__head" }, h("h1", {}, c.titulo)),
        h("div", { class: "recos", role: "radiogroup", "aria-label": c.titulo }, cards), det,
        h("section", { class: "actividad" },
          h("span", { class: "actividad__tag" }, icon(I.task), c.reto),
          h("p", { class: "reto__lead" }, c.reto_largo), ta,
          h("div", { class: "reto__bar" }, cnt, h("button", { class: "btn btn--primary", onclick: send }, icon(I.send), c.enviar)), rub),
        h("section", { class: "lectura" }, h("h2", {}, c.lista[0]), checks)),
      h("aside", { class: "mod__side" }, licita("16", c.guion)));
  };

  const refsScreen = c => h("section", { class: "screen" },
    h("header", { class: "mod__head" }, h("h1", {}, c.titulo)),
    h("section", { class: "lectura refs" }, c.refs.map(r => refNode(r, "ref"))),
    h("p", { class: "nota" }, icon(I.info), c.nota));
  R[17] = () => refsScreen(ODC.p17);
  R[18] = () => refsScreen(ODC.p18);

  R[19] = () => {
    const c = ODC.p19;
    return h("section", { class: "screen" },
      h("header", { class: "mod__head" }, h("h1", {}, c.titulo)),
      h("div", { class: "creditos" },
        h("table", { class: "meta" }, h("tbody", {}, c.filas.map(f => h("tr", {}, h("th", { scope: "row" }, f[0]), h("td", {}, f[1]))))),
        h("aside", { class: "lic" },
          h("div", { class: "cc" }, h("span", { class: "cc__logo", "aria-hidden": "true" }, "cc"), h("div", {}, h("strong", {}, c.cc[0]), h("span", {}, c.cc[1]))),
          h("p", {}, c.licencia))));
  };

  /* ================= Menú lateral ================= */
  const LABELS = { 1: ODC.p1.label, 2: ODC.p2.label, 3: ODC.p3.label };
  for (let n = 4; n <= 19; n++) LABELS[n] = (ODC["p" + n] || {}).label;
  ["p4", "p5", "p6", "p7", "p8", "p9", "p10", "p11", "p12"].forEach(k => LABELS[+k.slice(1)] = ODC[k].label);
  const TOTAL = 19;
  let current = 1;
  const targetOf = it => parseInt(it.pantalla.match(/\d+/)[0], 10);
  const groupOf = { 14: 13, 16: 15, 18: 17 };

  function buildMenu() {
    const nav = $("#menuNav"); nav.innerHTML = "";
    const item = (num, title, sub, target) => h("button", {
      class: "menu__item" + (done(target) ? " is-done" : ""), "data-go": target,
      onclick: () => { go(target); closeMenu(); }
    }, h("span", { class: "menu__num" }, done(target) ? icon(I.check, "ok") : num), h("span", {}, title, sub ? h("small", {}, sub) : null));
    nav.append(h("div", { class: "menu__group" },
      item("1", LABELS[1].split(" · ")[1], null, 1), item("2", LABELS[2].split(" · ")[1], null, 2), item("3", LABELS[3].split(" · ")[1], null, 3)));
    [[0, 4], [4, 8], [8, 13]].forEach((r, gi) => {
      const p = ODC.menu.partes[gi];
      nav.append(h("div", { class: "menu__group" }, h("div", { class: "menu__part" }, h("strong", {}, p.nombre), h("span", {}, p.desc)),
        M.slice(...r).map(it => {
          const t = targetOf(it), isMod = it.nombre.startsWith("Módulo");
          return item(isMod ? it.nombre.replace("Módulo ", "") : String(t), isMod ? it.titulo : it.nombre, isMod ? it.nombre : it.titulo, t);
        })));
    });
    markMenu();
  }
  function markMenu() {
    const cur = groupOf[current] || current;
    document.querySelectorAll(".menu__item").forEach(b => +b.dataset.go === cur ? b.setAttribute("aria-current", "page") : b.removeAttribute("aria-current"));
  }
  const menu = $("#menu"), scrim = $("#scrim"), openBtn = $("#menuOpen");
  function openMenu() { menu.classList.add("is-open"); scrim.hidden = false; openBtn.setAttribute("aria-expanded", "true"); const f = menu.querySelector(".menu__item"); f && f.focus(); }
  function closeMenu() { menu.classList.remove("is-open"); scrim.hidden = true; openBtn.setAttribute("aria-expanded", "false"); }
  openBtn.addEventListener("click", openMenu);
  $("#menuClose").addEventListener("click", closeMenu);
  scrim.addEventListener("click", closeMenu);
  $("#menuBtn").addEventListener("click", () => go(3));
  document.addEventListener("keydown", e => { if (e.key === "Escape") { if (!modal.hidden) closeModal(); else closeMenu(); } });

  /* ================= Navegación ================= */
  function go(id, push = true) {
    id = Math.min(Math.max(1, id | 0), TOTAL);
    stopVoice();
    current = id;
    const main = $("#contenido");
    main.innerHTML = ""; main.append(R[id]());
    document.body.classList.toggle("is-cover", id === 1);
    $("#screenLabel").textContent = LABELS[id];
    $("#odcTitle").textContent = ODC.p1.titulo;
    $("#progressBar").style.width = (id / TOTAL * 100) + "%";
    $("#prevBtn").disabled = id === 1;
    $("#nextBtn").disabled = id === TOTAL;
    if (push) history.replaceState(null, "", "#pantalla-" + id);
    window.scrollTo(0, 0);
    main.focus({ preventScroll: true });
    markMenu();
    store.set("ultima", id);
    SCORM.location(id);
    if (id === TOTAL) SCORM.status("completed");
  }
  $("#prevBtn").addEventListener("click", () => go(current - 1));
  $("#nextBtn").addEventListener("click", () => go(current + 1));
  document.addEventListener("keydown", e => {
    if (!modal.hidden || /INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName)) return;
    if (e.altKey && e.key === "ArrowRight") go(current + 1);
    if (e.altKey && e.key === "ArrowLeft") go(current - 1);
  });

  buildMenu();
  const m = location.hash.match(/pantalla-(\d+)/);
  go(m ? +m[1] : 1, false);
  if ("speechSynthesis" in window) speechSynthesis.getVoices();
})();
