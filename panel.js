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

  var W = window.WordBridge;
  var scenes = [];          // [{h, s, preview}]
  var curScene = -1;
  var open = {};            // index -> synopsis editor open
  var showSyn = true;
  var sig = "";
  var busy = false;

  function $(id) { return document.getElementById(id); }
  function esc(t) { return String(t).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }

  try { showSyn = localStorage.getItem("showSyn") !== "0"; } catch (e) {}

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
      W.applyStyle(b.dataset.style).then(refreshCursor).catch(showErr);
    });
  }

  function markStyle(style) {
    $("curstyle").textContent = style || "—";
    Array.prototype.forEach.call(document.querySelectorAll(".sbtn"), function (b) {
      b.classList.toggle("on", b.dataset.style === style);
    });
  }

  // ---------- scenes
  function refreshScenes() {
    if (busy) return Promise.resolve();
    busy = true;
    return W.paragraphs().then(function (ps) {
      var heads = [], previews = [];
      ps.forEach(function (p, i) {
        if (p.style === SCENE) {
          heads.push(p.text.trim());
          var pv = "";
          for (var k = i + 1; k < ps.length && ps[k].style !== SCENE; k++) {
            if (ps[k].style === "תיאור" && ps[k].text.trim()) { pv = ps[k].text.trim(); break; }
          }
          previews.push(pv);
        }
      });
      var newSig = JSON.stringify(heads) + "|" + JSON.stringify(previews);
      if (newSig === sig) return;
      var stored = scenes.length ? scenes.map(function (s) { return { h: s.h, s: s.s }; }) : W.loadSynopses();
      var syn = alignSynopses(stored, heads);
      var oldHeads = scenes.map(function (s) { return s.h; });
      scenes = heads.map(function (h, i) { return { h: h, s: syn[i] || "", preview: previews[i] }; });
      if (sig && JSON.stringify(oldHeads) !== JSON.stringify(heads)) { open = {}; persist(); }
      sig = newSig;
      renderScenes();
    }).catch(showErr).then(function () { busy = false; });
  }

  function persist() {
    var hasAny = scenes.some(function (s) { return s.s; });
    var stored = W.loadSynopses();
    if (!hasAny && !stored.length) return;
    W.saveSynopses(scenes.map(function (s) { return { h: s.h, s: s.s }; }));
  }

  function renderScenes() {
    var list = $("scenes");
    $("count").textContent = scenes.length ? scenes.length + " סצנות" : "";
    if (!scenes.length) {
      list.innerHTML = '<div class="empty">אין עדיין סצנות.<br>כותרת סצנה חדשה: ⌃⌥1</div>';
      return;
    }
    list.innerHTML = scenes.map(function (s, i) {
      var syn = s.s ? '<div class="syn">' + esc(s.s) + '</div>'
        : (s.preview ? '<div class="syn pv">' + esc(s.preview) + '</div>' : "");
      return '<div class="scene' + (i === curScene ? " cur" : "") + (open[i] ? " open" : "") + '" data-i="' + i + '">' +
        '<div class="row">' +
          '<button class="go" data-i="' + i + '"><span class="num">' + (i + 1) + '.</span> ' + esc(s.h || "(כותרת ריקה)") + '</button>' +
          '<button class="tog" data-i="' + i + '" title="תקציר" aria-label="תקציר">' + (s.s ? "✎" : "+") + '</button>' +
        '</div>' +
        (showSyn && !open[i] ? syn : "") +
        (open[i] ? '<textarea class="ed" data-i="' + i + '" rows="2" placeholder="תקציר של שורה אחת: מה קורה בסצנה">' + esc(s.s) + '</textarea>' : "") +
      '</div>';
    }).join("");
    var ed = list.querySelector(".ed");
    if (ed) { ed.focus(); ed.setSelectionRange(ed.value.length, ed.value.length); }
  }

  function markCurrent(i) {
    if (i === curScene) return;
    curScene = i;
    Array.prototype.forEach.call(document.querySelectorAll(".scene"), function (el) {
      el.classList.toggle("cur", +el.dataset.i === i);
    });
    var el = document.querySelector('.scene[data-i="' + i + '"]');
    if (el && el.scrollIntoViewIfNeeded) el.scrollIntoViewIfNeeded(false);
    else if (el) el.scrollIntoView({ block: "nearest" });
  }

  function wireList() {
    var list = $("scenes");
    list.addEventListener("click", function (e) {
      var go = e.target.closest(".go"), tog = e.target.closest(".tog");
      if (go) { var i = +go.dataset.i; markCurrent(i); W.goToScene(i).catch(showErr); }
      if (tog) {
        var j = +tog.dataset.i;
        open = open[j] ? {} : (function () { var o = {}; o[j] = true; return o; })();
        renderScenes();
      }
    });
    var t;
    list.addEventListener("input", function (e) {
      if (!e.target.classList.contains("ed")) return;
      scenes[+e.target.dataset.i].s = e.target.value.replace(/\n/g, " ");
      clearTimeout(t); t = setTimeout(persist, 600);
    });
    list.addEventListener("keydown", function (e) {
      if (e.target.classList.contains("ed") && (e.key === "Enter" || e.key === "Escape")) {
        e.preventDefault(); persist(); open = {}; renderScenes();
      }
    });
    list.addEventListener("focusout", function (e) {
      if (e.target.classList.contains("ed")) {
        setTimeout(function () {
          if (!document.activeElement || !document.activeElement.classList.contains("ed")) { persist(); open = {}; renderScenes(); }
        }, 150);
      }
    });
  }

  // ---------- cursor
  var cT;
  function refreshCursor() {
    return W.cursor().then(function (c) { markStyle(c.style); markCurrent(c.sceneIndex); }).catch(function () {});
  }
  function onSel() {
    clearTimeout(cT);
    cT = setTimeout(function () { refreshCursor(); refreshScenes(); }, 350);
  }

  function showErr(e) {
    var box = $("err");
    box.textContent = "שגיאה: " + (e && e.message ? e.message : e);
    box.hidden = false;
    setTimeout(function () { box.hidden = true; }, 6000);
  }

  // ---------- start
  W.ready().then(function () {
    buildStyleButtons();
    wireList();
    $("showsyn").checked = showSyn;
    $("showsyn").addEventListener("change", function () {
      showSyn = this.checked;
      try { localStorage.setItem("showSyn", showSyn ? "1" : "0"); } catch (e) {}
      renderScenes();
    });
    $("refresh").addEventListener("click", function () { sig = ""; refreshScenes().then(refreshCursor); });
    W.onSelectionChange(onSel);
    refreshScenes().then(refreshCursor);
    setInterval(function () {
      if (document.hidden || document.querySelector(".ed")) return;
      refreshScenes();
    }, 4000);
  });
})();
