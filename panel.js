// Scene panel UI. Talks to Word only through window.WordBridge.
(function () {
  var SCENE = "כותרת סצנה";
  var STYLES = [
    { name: "כותרת סצנה", label: "סצנה", key: "1" },
    { name: "תיאור", label: "תיאור", key: "2" },
    { name: "דמות", label: "דמות", key: "3" },
    { name: "דיאלוג", label: "דיאלוג", key: "4" },
    { name: "הערת בימוי", label: "הערה", key: "5" }
  ];
  var COLORS = ["#c9b8f0", "#f4a3a3", "#f7e17a", "#9fe0a5", "#8fe3ea", "#f2a6e3", "#f8c48c", "#a9c8f5", "#d5d0c4"];

  var W = window.WordBridge;
  var scenes = [];      // [{h, s, q, p:[plotId], preview}]
  var plots = [];       // [{id, n, c}]
  var loaded = false;
  var curScene = -1;
  var editing = -1;
  var view = "heads";   // heads | struct
  var showSyn = true;
  var filter = null;    // plot id
  var collapsed = {};   // sequence start index -> true
  var manage = false;
  var sig = "";
  var busy = false;

  function $(id) { return document.getElementById(id); }
  function esc(t) { return String(t == null ? "" : t).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function plotById(id) { for (var i = 0; i < plots.length; i++) if (plots[i].id === id) return plots[i]; return null; }
  function lsGet(k, d) { try { var v = localStorage.getItem(k); return v == null ? d : v; } catch (e) { return d; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }

  showSyn = lsGet("showSyn", "1") !== "0";
  view = lsGet("view", "heads");

  // ---------- data
  function persist() {
    var any = plots.length || scenes.some(function (s) { return s.s || s.q || (s.p && s.p.length); });
    if (!any && !W.loadData()) return;
    W.saveData({
      v: 2,
      plots: plots,
      scenes: scenes.map(function (s) { return { h: s.h, s: s.s, q: s.q, p: s.p }; })
    });
  }

  function refreshScenes(force) {
    if (busy) return Promise.resolve();
    busy = true;
    return W.paragraphs().then(function (ps) {
      var heads = [], previews = [];
      ps.forEach(function (p, i) {
        if (p.style !== SCENE) return;
        heads.push(p.text.trim());
        var pv = "";
        for (var k = i + 1; k < ps.length && ps[k].style !== SCENE; k++) {
          if (ps[k].style === "תיאור" && ps[k].text.trim()) { pv = ps[k].text.trim(); break; }
        }
        previews.push(pv);
      });
      var newSig = JSON.stringify(heads) + "|" + JSON.stringify(previews);
      if (newSig === sig && !force) return;
      var old;
      if (!loaded) {
        var d = W.loadData() || { scenes: [], plots: [] };
        plots = d.plots || [];
        old = d.scenes || [];
        loaded = true;
      } else {
        old = scenes;
      }
      var map = alignScenes(old.map(function (s) { return s.h; }), heads);
      var changed = old.length !== heads.length || map.some(function (m, i) { return m !== i; }) ||
        heads.some(function (h, i) { return !old[i] || old[i].h !== h; });
      scenes = heads.map(function (h, i) {
        var o = map[i] >= 0 ? old[map[i]] : null;
        return { h: h, s: o ? o.s || "" : "", q: o ? o.q || "" : "", p: o ? (o.p || []).slice() : [], preview: previews[i] };
      });
      if (changed && sig) { editing = -1; persist(); }
      sig = newSig;
      render();
    }).catch(showErr).then(function () { busy = false; });
  }

  // ---------- style buttons
  function buildStyleButtons() {
    var box = $("styles");
    box.innerHTML = STYLES.map(function (s) {
      return '<button class="sbtn" data-style="' + esc(s.name) + '" title="' + esc(s.name) + '  ⌃⌥' + s.key + '">' +
        '<span class="sl">' + esc(s.label) + '</span><span class="sk">⌃⌥' + s.key + '</span></button>';
    }).join("");
    box.addEventListener("click", function (e) {
      var b = e.target.closest(".sbtn");
      if (!b) return;
      W.applyStyle(b.dataset.style).then(function () { refreshCursor(); refreshScenes(); }).catch(showErr);
    });
  }

  function markStyle(style) {
    $("curstyle").textContent = style || "—";
    Array.prototype.forEach.call(document.querySelectorAll(".sbtn"), function (b) {
      b.classList.toggle("on", b.dataset.style === style);
    });
  }

  // ---------- rendering
  function chip(p, extra) {
    return '<span class="chip' + (extra || "") + '" data-plot="' + esc(p.id) + '" style="background:' + p.c + '">' + esc(p.n) + '</span>';
  }

  function renderBar() {
    $("bar").innerHTML =
      '<div class="seg">' +
        '<button data-view="heads" class="' + (view === "heads" ? "on" : "") + '">כותרות</button>' +
        '<button data-view="struct" class="' + (view === "struct" ? "on" : "") + '">מבנה</button>' +
      '</div>' +
      '<span class="cnt">' + (scenes.length ? scenes.length + " סצנות" : "") + '</span>' +
      '<span class="sp"></span>' +
      (view === "heads" ? '<label><input type="checkbox" id="showsyn"' + (showSyn ? " checked" : "") + '> תקצירים</label>' :
        '<button class="icon" id="fold" title="כווץ או פתח את כל הסיקוונסים">' + (allCollapsed() ? "⊞" : "⊟") + '</button>') +
      '<button class="icon" id="refresh" title="רענון" aria-label="רענון">↻</button>';
  }

  function renderPlotBar() {
    var box = $("plotbar");
    if (!plots.length && view !== "struct") { box.innerHTML = ""; box.hidden = true; return; }
    box.hidden = false;
    var counts = {};
    scenes.forEach(function (s) { (s.p || []).forEach(function (id) { counts[id] = (counts[id] || 0) + 1; }); });
    var html = '<div class="plots">' +
      (plots.length ? plots.map(function (p) {
        return '<span class="chip big' + (filter === p.id ? " sel" : "") + (filter && filter !== p.id ? " dim" : "") +
          '" data-filter="' + esc(p.id) + '" style="background:' + p.c + '" title="הצג רק את הסצנות של העלילה">' +
          esc(p.n) + ' <small>' + (counts[p.id] || 0) + '</small></span>';
      }).join("") : '<span class="hint">עלילות: פתח סצנה (✎) והוסף עלילה</span>') +
      (plots.length ? '<button class="icon" id="manage" title="ניהול עלילות">⚙</button>' : "") +
      '</div>';
    if (filter) html += '<div class="fnote">מוצגות רק הסצנות של "' + esc((plotById(filter) || {}).n) + '" · <a href="#" id="clearf">הצג הכל</a></div>';
    if (manage && plots.length) {
      html += '<div class="manage">' + plots.map(function (p) {
        return '<div class="mrow" data-id="' + esc(p.id) + '">' +
          '<button class="swatch" style="background:' + p.c + '" title="החלף צבע"></button>' +
          '<input class="mname" value="' + esc(p.n) + '">' +
          '<button class="icon mdel" title="מחק עלילה">✕</button></div>';
      }).join("") + '</div>';
    }
    box.innerHTML = html;
  }

  function sequences() {
    // [{start, title, items:[sceneIndex]}]; scenes before the first sequence go in a group with start -1
    var seqs = [], cur = null;
    scenes.forEach(function (s, i) {
      if (s.q || !cur) { cur = { start: s.q ? i : -1, title: s.q, items: [] }; seqs.push(cur); }
      cur.items.push(i);
    });
    return seqs;
  }
  function allCollapsed() {
    var seqs = sequences().filter(function (q) { return q.start >= 0; });
    return seqs.length > 0 && seqs.every(function (q) { return collapsed[q.start]; });
  }

  function editorHtml(i) {
    var s = scenes[i];
    return '<div class="editor" data-i="' + i + '">' +
      '<textarea class="ed" data-i="' + i + '" rows="2" placeholder="תקציר של שורה אחת: מה קורה בסצנה">' + esc(s.s) + '</textarea>' +
      '<div class="elabel">עלילות</div>' +
      '<div class="eplots">' + plots.map(function (p) {
        return chip(p, (s.p.indexOf(p.id) >= 0 ? " on" : " off") + " tog-plot");
      }).join("") +
        '<input class="newplot" placeholder="+ עלילה חדשה" size="10">' +
      '</div>' +
      '<label class="eseq"><input type="checkbox" class="isseq"' + (s.q ? " checked" : "") + '> תחילת סיקוונס</label>' +
      (s.q ? '<input class="seqname" value="' + esc(s.q) + '" placeholder="שם הסיקוונס">' : "") +
      '<div class="eact">' +
        '<button class="mv" data-d="-1" title="הזז סצנה למעלה"' + (i === 0 ? " disabled" : "") + '>▲</button>' +
        '<button class="mv" data-d="1" title="הזז סצנה למטה"' + (i === scenes.length - 1 ? " disabled" : "") + '>▼</button>' +
        '<span class="sp"></span><button class="done">סגור</button>' +
      '</div></div>';
  }

  function sceneHtml(i) {
    var s = scenes[i];
    var chips = (s.p || []).map(function (id) { var p = plotById(id); return p ? chip(p, " sm") : ""; }).join("");
    var handle = filter ? "" : '<span class="hd" data-i="' + i + '" title="גרור כדי להזיז">⋮⋮</span>';
    var main;
    if (view === "struct") {
      main = '<button class="go" data-i="' + i + '"><span class="num">' + (i + 1) + '.</span> ' +
        (s.s ? esc(s.s) : '<span class="muted">' + esc(s.h || "(כותרת ריקה)") + '</span>') + '</button>';
    } else {
      main = '<button class="go" data-i="' + i + '"><span class="num">' + (i + 1) + '.</span> <b>' + esc(s.h || "(כותרת ריקה)") + '</b></button>';
    }
    var syn = "";
    if (view === "heads" && showSyn && editing !== i) {
      syn = s.s ? '<div class="syn">' + esc(s.s) + '</div>' : (s.preview ? '<div class="syn pv">' + esc(s.preview) + '</div>' : "");
    }
    return '<div class="scene' + (i === curScene ? " cur" : "") + (editing === i ? " open" : "") + '" data-i="' + i + '">' +
      '<div class="row">' + handle + main +
        '<button class="tog" data-i="' + i + '" title="תקציר, עלילות וסיקוונס">' + (editing === i ? "▴" : "✎") + '</button>' +
      '</div>' + syn +
      (chips && editing !== i ? '<div class="chips">' + chips + '</div>' : "") +
      (editing === i ? editorHtml(i) : "") +
    '</div>';
  }

  function renderList() {
    var list = $("scenes");
    list.className = view === "struct" ? "view-struct" : "";
    if (!scenes.length) {
      list.innerHTML = '<div class="empty">אין עדיין סצנות.<br>כותרת סצנה חדשה: ⌃⌥1</div>';
      return;
    }
    var visible = function (i) { return !filter || scenes[i].p.indexOf(filter) >= 0 || editing === i; };
    var html = "";
    if (view === "struct") {
      var seqs = sequences(), n = 0;
      var hasSeq = seqs.some(function (q) { return q.start >= 0; });
      if (!hasSeq) html += '<div class="hint box">אין עדיין סיקוונסים. פתח סצנה (✎) וסמן "תחילת סיקוונס".</div>';
      seqs.forEach(function (q) {
        var items = q.items.filter(visible);
        if (q.start >= 0) {
          n++;
          var col = collapsed[q.start];
          html += '<div class="seq' + (col ? " col" : "") + (q.items.indexOf(curScene) >= 0 ? " cur" : "") + '" data-start="' + q.start + '">' +
            '<span class="arr">' + (col ? "◂" : "▾") + '</span> ' + n + '. ' + esc(q.title) +
            '<span class="scount">' + q.items.length + '</span></div>';
          if (col) return;
        } else if (hasSeq && items.length) {
          html += '<div class="seq none">לפני הסיקוונס הראשון</div>';
        }
        items.forEach(function (i) { html += sceneHtml(i); });
      });
    } else {
      scenes.forEach(function (s, i) { if (visible(i)) html += sceneHtml(i); });
    }
    list.innerHTML = html;
  }

  function render() {
    renderBar();
    renderPlotBar();
    renderList();
  }

  function markCurrent(i) {
    if (i === curScene) return;
    curScene = i;
    Array.prototype.forEach.call(document.querySelectorAll(".scene"), function (el) {
      el.classList.toggle("cur", +el.dataset.i === i);
    });
    Array.prototype.forEach.call(document.querySelectorAll(".seq[data-start]"), function (el) {
      var q = sequences().filter(function (x) { return x.start === +el.dataset.start; })[0];
      el.classList.toggle("cur", !!q && q.items.indexOf(i) >= 0);
    });
    var el = document.querySelector('.scene[data-i="' + i + '"]');
    if (el && editing < 0) {
      if (el.scrollIntoViewIfNeeded) el.scrollIntoViewIfNeeded(false); else el.scrollIntoView({ block: "nearest" });
    }
  }

  // ---------- actions
  function addPlot(name) {
    name = name.trim();
    if (!name) return null;
    for (var i = 0; i < plots.length; i++) if (plots[i].n === name) return plots[i];
    var p = { id: "p" + Date.now().toString(36) + Math.floor(Math.random() * 99), n: name, c: COLORS[plots.length % COLORS.length] };
    plots.push(p);
    return p;
  }

  function moveScene(from, to) {
    // to = index the scene should sit before (scenes.length = end)
    if (to === from || to === from + 1 || busy) return;
    var rec = scenes[from];
    if (rec.q && scenes[from + 1] && !scenes[from + 1].q) { scenes[from + 1].q = rec.q; }
    rec.q = "";
    scenes.splice(from, 1);
    var at = to > from ? to - 1 : to;
    scenes.splice(at, 0, rec);
    editing = editing === from ? at : -1;
    persist();
    render();
    busy = true;
    W.moveScene(from, to).then(function () {
      busy = false; sig = "";
      return refreshScenes().then(refreshCursor);
    }).catch(function (e) { busy = false; showErr(e); sig = ""; refreshScenes(); });
  }

  // ---------- events
  function wire() {
    $("bar").addEventListener("click", function (e) {
      var v = e.target.closest("[data-view]");
      if (v) { view = v.dataset.view; lsSet("view", view); editing = -1; render(); return; }
      if (e.target.closest("#refresh")) { sig = ""; refreshScenes(true).then(refreshCursor); return; }
      if (e.target.closest("#fold")) {
        var all = allCollapsed();
        collapsed = {};
        if (!all) sequences().forEach(function (q) { if (q.start >= 0) collapsed[q.start] = true; });
        render();
      }
    });
    $("bar").addEventListener("change", function (e) {
      if (e.target.id === "showsyn") { showSyn = e.target.checked; lsSet("showSyn", showSyn ? "1" : "0"); renderList(); }
    });

    var pb = $("plotbar");
    pb.addEventListener("click", function (e) {
      var f = e.target.closest("[data-filter]");
      if (f) { filter = filter === f.dataset.filter ? null : f.dataset.filter; editing = -1; render(); return; }
      if (e.target.closest("#clearf")) { e.preventDefault(); filter = null; render(); return; }
      if (e.target.closest("#manage")) { manage = !manage; renderPlotBar(); return; }
      var sw = e.target.closest(".swatch");
      if (sw) {
        var p = plotById(sw.closest(".mrow").dataset.id);
        p.c = COLORS[(COLORS.indexOf(p.c) + 1) % COLORS.length];
        persist(); render(); return;
      }
      var del = e.target.closest(".mdel");
      if (del) {
        var id = del.closest(".mrow").dataset.id, pl = plotById(id);
        if (!confirm('למחוק את העלילה "' + pl.n + '"? היא תוסר מכל הסצנות.')) return;
        plots = plots.filter(function (x) { return x.id !== id; });
        scenes.forEach(function (s) { s.p = s.p.filter(function (x) { return x !== id; }); });
        if (filter === id) filter = null;
        persist(); render();
      }
    });
    pb.addEventListener("change", function (e) {
      if (e.target.classList.contains("mname")) {
        var p = plotById(e.target.closest(".mrow").dataset.id);
        if (e.target.value.trim()) { p.n = e.target.value.trim(); persist(); renderList(); }
      }
    });

    var list = $("scenes");
    list.addEventListener("click", function (e) {
      var go = e.target.closest(".go");
      if (go) { var i = +go.dataset.i; markCurrent(i); W.goToScene(i).catch(showErr); return; }
      var tog = e.target.closest(".tog");
      if (tog) { var j = +tog.dataset.i; editing = editing === j ? -1 : j; renderList(); focusEditor(); return; }
      var seq = e.target.closest(".seq[data-start]");
      if (seq) { var st = +seq.dataset.start; collapsed[st] = !collapsed[st]; render(); return; }
      var tp = e.target.closest(".tog-plot");
      if (tp && editing >= 0) {
        var s = scenes[editing], id = tp.dataset.plot, k = s.p.indexOf(id);
        if (k >= 0) s.p.splice(k, 1); else s.p.push(id);
        persist(); renderPlotBar(); tp.classList.toggle("on"); tp.classList.toggle("off"); return;
      }
      var mv = e.target.closest(".mv");
      if (mv && editing >= 0) {
        var d = +mv.dataset.d;
        moveScene(editing, d < 0 ? editing - 1 : editing + 2);
        return;
      }
      if (e.target.closest(".done")) { editing = -1; render(); }
    });
    list.addEventListener("input", function (e) {
      if (editing < 0) return;
      if (e.target.classList.contains("ed")) { scenes[editing].s = e.target.value.replace(/\n/g, " "); debouncePersist(); }
      if (e.target.classList.contains("seqname")) { scenes[editing].q = e.target.value || " "; debouncePersist(); }
    });
    list.addEventListener("change", function (e) {
      if (editing < 0) return;
      if (e.target.classList.contains("isseq")) {
        scenes[editing].q = e.target.checked ? "סיקוונס חדש" : "";
        persist(); renderList();
        var sn = document.querySelector(".seqname"); if (sn) { sn.focus(); sn.select(); }
      }
      if (e.target.classList.contains("seqname")) { scenes[editing].q = e.target.value.trim() || "סיקוונס"; persist(); }
    });
    list.addEventListener("keydown", function (e) {
      if (editing < 0) return;
      if (e.target.classList.contains("newplot") && e.key === "Enter") {
        e.preventDefault();
        var p = addPlot(e.target.value);
        if (p && scenes[editing].p.indexOf(p.id) < 0) scenes[editing].p.push(p.id);
        persist(); renderPlotBar(); renderList();
        var np = document.querySelector(".newplot"); if (np) np.focus();
        return;
      }
      if (e.target.classList.contains("ed") && e.key === "Enter") { e.preventDefault(); persist(); editing = -1; render(); }
      if (e.key === "Escape") { persist(); editing = -1; render(); }
    });
    wireDrag(list);
  }

  var pT;
  function debouncePersist() { clearTimeout(pT); pT = setTimeout(persist, 600); }
  function focusEditor() {
    var ed = document.querySelector(".ed");
    if (ed) { ed.focus(); ed.setSelectionRange(ed.value.length, ed.value.length); }
  }

  // ---------- drag to reorder
  function wireDrag(list) {
    var from = -1, line = null, target = -1;
    list.addEventListener("pointerdown", function (e) {
      var h = e.target.closest(".hd");
      if (!h || filter) return;
      e.preventDefault();
      from = +h.dataset.i;
      h.setPointerCapture(e.pointerId);
      line = document.createElement("div");
      line.className = "dropline";
      list.appendChild(line);
      var row = h.closest(".scene"); row.classList.add("dragging");
    });
    list.addEventListener("pointermove", function (e) {
      if (from < 0) return;
      var rows = Array.prototype.slice.call(list.querySelectorAll(".scene"));
      var lr = list.getBoundingClientRect();
      target = -1;
      for (var k = 0; k < rows.length; k++) {
        var r = rows[k].getBoundingClientRect();
        if (e.clientY < r.top + r.height / 2) { target = +rows[k].dataset.i; line.style.top = (r.top - lr.top + list.scrollTop - 3) + "px"; break; }
      }
      if (target < 0 && rows.length) {
        var lastR = rows[rows.length - 1].getBoundingClientRect();
        target = +rows[rows.length - 1].dataset.i + 1;
        line.style.top = (lastR.bottom - lr.top + list.scrollTop) + "px";
      }
      // autoscroll
      if (e.clientY < lr.top + 30) list.scrollTop -= 12;
      if (e.clientY > lr.bottom - 30) list.scrollTop += 12;
    });
    function end() {
      if (from < 0) return;
      var f = from, t = target;
      from = -1; target = -1;
      if (line) { line.remove(); line = null; }
      var d = list.querySelector(".dragging"); if (d) d.classList.remove("dragging");
      if (t >= 0) moveScene(f, t);
    }
    list.addEventListener("pointerup", end);
    list.addEventListener("pointercancel", function () { target = -1; end(); });
  }

  // ---------- cursor
  var cT;
  function refreshCursor() {
    return W.cursor().then(function (c) { markStyle(c.style); markCurrent(c.sceneIndex); }).catch(function () {});
  }
  function onSel() {
    clearTimeout(cT);
    cT = setTimeout(function () { refreshCursor(); if (editing < 0) refreshScenes(); }, 350);
  }

  function showErr(e) {
    var box = $("err");
    box.textContent = "שגיאה: " + (e && e.message ? e.message : e);
    box.hidden = false;
    setTimeout(function () { box.hidden = true; }, 7000);
  }

  // ---------- start
  W.ready().then(function () {
    buildStyleButtons();
    wire();
    W.onSelectionChange(onSel);
    refreshScenes().then(refreshCursor);
    setInterval(function () {
      if (document.hidden || editing >= 0) return;
      refreshScenes();
    }, 4000);
  });
})();
