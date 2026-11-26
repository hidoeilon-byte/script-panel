// Scene panel UI. Talks to Word only through window.WordBridge.
(function () {
  var SCENE = "כותרת סצנה";
  var PAREN = "הערת בימוי";
  var lastPara = null;
  var customLocs = [];   // locations added by hand
  var locOpen = false;   // bank manager open
  var STYLES = [
    { name: "כותרת סצנה", label: "סצנה", key: "F1" },
    { name: "תיאור", label: "תיאור", key: "F2" },
    { name: "דמות", label: "דמות", key: "F3" },
    { name: "דיאלוג", label: "דיאלוג", key: "F4" },
    { name: "הערת בימוי", label: "הערה", key: "F8" }
  ];
  var COLORS = ["#c9b8f0", "#f4a3a3", "#f7e17a", "#9fe0a5", "#8fe3ea", "#f2a6e3", "#f8c48c", "#a9c8f5", "#d5d0c4"];

  var W = window.WordBridge;
  var scenes = [];      // [{h, s, q, p:[plotId], preview}]
  var plots = [];       // [{id, n, c}]
  var DEFAULT_MARKS = [
    { id: "m1", n: "לשכתב", c: "#f4a3a3" },
    { id: "m2", n: "לבדוק", c: "#f7e17a" },
    { id: "m3", n: "מוכן", c: "#9fe0a5" }
  ];
  var marks = DEFAULT_MARKS.slice();   // scene color marks [{id, n, c}]
  var mfilter = null;   // mark id
  var picking = -1;     // scene index with open color picker
  var mmanage = false;
  var loaded = false;
  var curScene = -1;
  var editing = -1;
  var view = "heads";   // heads | struct
  var lastSub = { script: "struct", notes: "todo", comparisons: "changes", tools: "tools" };
  var STAGES = ["מחולל", "אקפוזיציה", "מפנה 1", "מפנה 2", "שיא", "התרה"];
  // older names of the parts map to the short ones, so nothing saved before breaks
  var STAGE_OLD = { "אירוע מחולל": "מחולל", "אקספוזיציה": "אקפוזיציה", "נקודת מפנה": "מפנה 1", "נקודת מפנה שנייה": "מפנה 2", "נקודת שיא": "שיא" };
  function stNorm(v) { return STAGE_OLD[v] || v || ""; }
  var stageOpen = {};   // which parts of the structure are open (click the part header)
  var dragI = -1, sumEdit = -1;   // scene whose summary is being edited inside the parts   // remembered sub-tab per group
  var showSyn = true;
  var filter = null;    // plot id
  var collapsed = {};   // sequence start index -> true
  var manage = false;
  var sig = "";
  var busy = false;

  function $(id) { return document.getElementById(id); }
  function esc(t) { return String(t == null ? "" : t).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function markById(id) { for (var i = 0; i < marks.length; i++) if (marks[i].id === id) return marks[i]; return null; }
  function plotById(id) { for (var i = 0; i < plots.length; i++) if (plots[i].id === id) return plots[i]; return null; }
  function lsGet(k, d) { try { var v = localStorage.getItem(k); return v == null ? d : v; } catch (e) { return d; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }

  showSyn = lsGet("showSyn", "1") !== "0";
  view = lsGet("view", "struct");
  if (view === "notes") view = "compare";
  if (view === "heads" || view === "feedback") view = "struct";
  lastSub[groupOf(view)] = view;
  if (lsGet("compact", "") === "1") document.body.classList.add("compact");

  // ---------- data
  function persist() {
    var any = plots.length || customLocs.length || scenes.some(function (s) { return s.s || s.q || s.m || (s.p && s.p.length); });
    if (!any && !W.loadData()) return;
    W.saveData({
      v: 2,
      plots: plots,
      marks: marks,
      locs: customLocs,
      scenes: scenes.map(function (s) { return { h: s.h, s: s.s, q: s.q, p: s.p, m: s.m || "", st: s.st || "" }; })
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
        if (d.marks && d.marks.length) marks = d.marks;
        customLocs = d.locs || [];
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
        return { h: h, s: o ? o.s || "" : "", q: o ? o.q || "" : "", p: o ? (o.p || []).slice() : [], m: o ? o.m || "" : "", st: o ? o.st || "" : "", preview: previews[i] };
      });
      if (changed && sig) { editing = -1; persist(); }
      sig = newSig;
      render();
    }).catch(showErr).then(function () { busy = false; });
  }

  // ---------- style buttons
  // same order as Tab in the Word add-in (CycleType)
  var CYCLE = ["תיאור", "דמות", "דיאלוג", "הערת בימוי", "כותרת סצנה"];
  var IOS = false;   // set after Office.onReady

  function setType(st) {
    return W.applyStyle(st).then(function () {
      if (st === PAREN) { lastPara = { style: PAREN, text: "" }; return W.wrapParen(); }
    }).then(function () { refreshCursor(); refreshScenes(); }).catch(showErr);
  }

  function cycleType(dir) {
    return W.cursor().then(function (c) {
      var k = CYCLE.indexOf(c.style);
      var n = k < 0 ? (dir > 0 ? 0 : CYCLE.length - 1) : (k + dir + CYCLE.length) % CYCLE.length;
      return setType(CYCLE[n]);
    }).catch(showErr);
  }

  function buildStyleButtons() {
    try { IOS = Office.context.platform === Office.PlatformType.iOS; } catch (e) {}
    var box = $("styles");
    box.innerHTML = STYLES.map(function (s) {
      return '<button class="sbtn" data-style="' + esc(s.name) + '" title="' + esc(s.name) + '">' +
        '<span class="sl">' + esc(s.label) + '</span>' + (IOS ? "" : '<span class="sk">' + s.key + '</span>') + '</button>';
    }).join("") +
      '';
    if (IOS) {
      var hint = document.querySelector(".tabhint");
      if (hint) hint.textContent = "סוג הבא מחליף סוג שורה";
    }
    box.addEventListener("click", function (e) {
      var b = e.target.closest(".sbtn");
      if (!b) return;
      if (b.dataset.dir) { cycleType(+b.dataset.dir); return; }
      setType(b.dataset.style);
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

  // three groups; each group has two sub-tabs
  var GROUPS = [
    { g: "script", t: "תסריט", subs: [["struct", "מבנה"]] },
    { g: "notes", t: "הערות", subs: [["todo", "שלי"], ["feedback", "משוב"]] },
    { g: "comparisons", t: "השוואות", subs: [["changes", "דראפטים"], ["compare", "סינופסיס"]] },
    { g: "tools", t: "כלים", subs: [["tools", "כלים"]] }
  ];
  function groupOf(v) {
    return v === "struct" || v === "heads" ? "script" : v === "notes" || v === "todo" || v === "feedback" || v === "answers" ? "notes" :
      v === "changes" || v === "compare" ? "comparisons" : "tools";
  }
  function renderBar() {
    var cur = groupOf(view), gb = GROUPS.filter(function (x) { return x.g === cur; })[0];
    $("bar").innerHTML =
      '<div class="seg">' + GROUPS.map(function (x) {
        return '<button data-group="' + x.g + '" class="' + (x.g === cur ? "on" : "") + '">' + x.t +
          (x.g === "notes" && openNotes() ? ' <span class="badge">' + openNotes() + '</span>' : '') + '</button>';
      }).join("") + '</div>' +
      '<div class="seg sub">' + gb.subs.map(function (sv) {
        return '<button data-view="' + sv[0] + '" class="' + (view === sv[0] ? "on" : "") + '">' + sv[1] + '</button>';
      }).join("") + '</div>' +
      '<span class="cnt">' + (scenes.length ? scenes.length + " סצנות" : "") + '</span>' +
      '<button class="icon" id="barfold" title="הסתר או הצג את סרגל הכלים למעלה" aria-label="הסתר או הצג סרגל">' + (document.body.classList.contains("compact") ? "▾" : "▴") + '</button>' +
      '<span class="sp"></span>' +
      (view === "heads" ? '<label><input type="checkbox" id="showsyn"' + (showSyn ? " checked" : "") + '> תקצירים</label>' :
        view === "struct" ? '<button class="icon" id="fold" title="כווץ או פתח את כל הסיקוונסים">' + (allCollapsed() ? "⊞" : "⊟") + '</button>' : "") +
      '<button class="icon" id="refresh" title="רענון" aria-label="רענון">↻</button>';
  }

  function renderMarkBar() {
    var mc = {};
    scenes.forEach(function (s) { if (s.m) mc[s.m] = (mc[s.m] || 0) + 1; });
    var html = '<div class="plots marks"><span class="lbl">צבע:</span>' + marks.map(function (m) {
      return '<span class="chip big' + (mfilter === m.id ? " sel" : "") + (mfilter && mfilter !== m.id ? " dim" : "") +
        '" data-mfilter="' + esc(m.id) + '" style="background:' + m.c + '" title="הצג רק סצנות בצבע הזה">' +
        esc(m.n) + ' <small>' + (mc[m.id] || 0) + '</small></span>';
    }).join("") + '<button class="icon" id="mmanage" title="ניהול צבעים">⚙</button></div>';
    if (mfilter) html += '<div class="fnote">מוצגות רק סצנות "' + esc((markById(mfilter) || {}).n) + '" · <a href="#" id="clearm">הצג הכל</a></div>';
    if (mmanage) {
      html += '<div class="manage">' + marks.map(function (m) {
        return '<div class="mrow" data-mid="' + esc(m.id) + '">' +
          '<button class="swatch mswatch" style="background:' + m.c + '" title="החלף צבע"></button>' +
          '<input class="mname mmname" value="' + esc(m.n) + '">' +
          '<button class="icon mmdel" title="מחק צבע">✕</button></div>';
      }).join("") + '<div class="mrow"><input class="mname mnew" placeholder="+ צבע חדש (שם ו-Enter)"></div></div>';
    }
    $("markbar").innerHTML = html;
  }

  function renderPlotBar() {
    renderMarkBar();
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
      '<div class="askrow"><button class="lb askclaude" data-i="' + i + '">שאל את Claude על הסצנה</button></div>' +
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
    var handle = (filter || mfilter) ? "" : '<span class="hd" data-i="' + i + '" title="גרור כדי להזיז">⋮⋮</span>';
    var main = '<button class="go" data-i="' + i + '"><span class="num">' + (i + 1) + '.</span> <b>' + esc(s.h || "(כותרת ריקה)") + '</b></button>';
    var syn = "";
    if ((view === "struct" || showSyn) && editing !== i) {
      syn = s.s ? '<div class="syn">' + esc(s.s) + '</div>' : (s.preview ? '<div class="syn pv">' + esc(s.preview) + '</div>' : "");
    }
    var mk = s.m ? markById(s.m) : null;
    var dot = '<button class="dot' + (mk ? "" : " empty") + '" data-i="' + i + '" title="' + (mk ? esc(mk.n) : "סימון בצבע") + '"' +
      (mk ? ' style="background:' + mk.c + '"' : "") + '></button>';
    var picker = picking === i ? '<div class="markpick">' + marks.map(function (m) {
      return '<button class="pick' + (s.m === m.id ? " on" : "") + '" data-m="' + esc(m.id) + '" style="background:' + m.c + '">' + esc(m.n) + '</button>';
    }).join("") + '<button class="pick none" data-m="">ללא</button></div>' : "";
    return '<div class="scene' + (i === curScene ? " cur" : "") + (editing === i ? " open" : "") + (mk ? " marked" : "") + '" data-i="' + i + '"' +
      (mk ? ' style="--mk:' + mk.c + '"' : "") + '>' +
      '<div class="row">' + handle + main + dot +
        '<button class="tog" data-i="' + i + '" title="תקציר, עלילות וסיקוונס">' + (editing === i ? "▴" : "✎") + '</button>' +
      '</div>' + picker + syn +
      (chips && editing !== i ? '<div class="chips">' + chips + '</div>' : "") +
      (editing === i ? editorHtml(i) : "") +
    '</div>';
  }

  // ---------- changes view (talks to the macro through a custom XML part)
  var chg = null;          // {vs, last, n, items:[{i,pos,text}]}
  var chgPending = "";
  var chgSel = 0;
  var marks2 = null;       // [{name,i,text}] read straight from the document
  var markIdx = 0;
  // ---------- tools: the actions that used to be ribbon buttons, sent to the macro through the bridge
  var TOOLS = [
    { h: "סינופסיס", items: [
      { cmd: "syn2scr", t: "עדכן תסריט מסינופסיס", d: "הסינופסיס מעודכן: מוסיף סצנות חסרות ומחליף לוקיישנים בתסריט. תיאור הסינופסיס נכנס כהערה." },
      { cmd: "scr2syn", t: "עדכן סינופסיס מתסריט", d: "התסריט מעודכן: מוסיף לסינופסיס כותרות חסרות ומחליף לוקיישנים." },
      { cmd: "savesyn", t: "שמור סינופסיס חדש", d: "שומר עותק של הסינופסיס הפתוח, במספר הבא, בתיקייה שלו." },
      { cmd: "fixsyn", t: "העתק סגנונות לסינופסיס", d: "מעתיק את סגנונות התסריט (כותרת סצנה, דיאלוג...) לסינופסיס העבודה, כדי שהכפתורים יעבדו בו." }
    ], note: "לחיצה ראשונה מציגה מה ישתנה. לחיצה שנייה על אותו כפתור מבצעת. לפני כל שינוי נשמר גיבוי." },
    { h: "השוואה חכמה עם Claude", items: [
      { cmd: "smart_s2t", t: "סינופסיס ← תסריט", d: "מה שכתוב בסינופסיס וחסר או שונה בתסריט. ההערות נכנסות לתסריט." },
      { cmd: "smart_t2s", t: "תסריט ← סינופסיס", d: "מה שקורה בתסריט וחסר או שונה בסינופסיס. נפתח הסינופסיס וההערות נכנסות אליו." },
      { cmd: "smartclear", t: "סיום השוואה", d: "מוחק את הערות העזר. הערות ההשוואה נשארות." }
    ], note: "1. לוחצים על כיוון. 2. מעתיקים את הבקשה שמופיעה כאן. 3. מדביקים ל-Claude ב-Word (בחלון של המסמך שמקבל את ההערות). 4. ההבדלים מופיעים בלשונית הערות." },
    { h: "ניתוח ביטים", items: [
      { cmd: "beats", t: "ניתוח ביטים בסצנות", d: "בוחרים סצנות. Claude מזהה את המנעד הרגשי ושינויי המצב, והסימונים נכנסים למרקר בצבע לכל סוג." }
    ] },
    { h: "דראפטים", items: [
      { cmd: "force", t: "דראפט חדש עכשיו", d: "שומר דראפט חדש בלי לחכות לשינוי מהותי." },
      { cmd: "prevdraft", t: "דראפט קודם", d: "פותח את הדראפט הקודם לקריאה, לצד הקובץ." }
    ] },
    { h: "הגשה", items: [
      { cmd: "check", t: "לפני הגשה", d: "בודק את הפורמט לפני שליחה לקרן." }
    ] },
    { h: "תסריט", items: [
      { cmd: "new", t: "תסריט חדש", d: "פותח תסריט חדש בפורמט הקרן." }
    ] }
  ];
  var toolPending = "";
  var smartPrompt = "";   // "s2t" | "t2s" | ""
  var smartBundle = { ready: false, msg: "", text: "", scenes: 0, missing: 0 };
  // ---------- beats: pick scenes, copy the request to Claude for Word, paste the answer back and mark it
  var beatsOpen = false;
  var beatSel = {};        // scene index -> true
  var BEAT_COLORS = { "שינוי מצב": "DarkBlue", "צחוק": "Pink", "התקרבות": "Blue", "התרככות": "Gray25" };
  var BEATS_PROMPT = "ניתוח ביטים רגשיים בסצנות. פתח את המסמך הפתוח ב-Word, ומצא את הסצנות הבאות לפי הכותרת שלהן. לכל סצנה: כתוב משפט אחד על המנעד הרגשי שלה, ואיפה הזוג אוהב ואיפה הוא משלים. סמן במדויק רגעים של: שינוי מצב, צחוק של הדמות, התקרבות, והתרככות. אל תשנה את המסמך ואל תכתוב הערות. החזר רשימה: שורה לכל סצנה שמתחילה ב\"סצנה: \" עם שם הסצנה ומשפט המנעד, ואחריה שורה לכל סימון בפורמט:\nסוג | ציטוט מדויק מהטקסט (עד 100 תווים, מילה במילה)\nסוגים מותרים: שינוי מצב, צחוק, התקרבות, התרככות.\n\nהסצנות:\n";
  function buildBeatsPrompt() {
    var names = scenes.map(function (s, i) { return beatSel[i] ? "- " + (s.h || "סצנה " + (i + 1)) : ""; }).filter(Boolean).join("\n");
    return BEATS_PROMPT + (names || "(עוד לא נבחרו סצנות)");
  }
  function beatsBox() {
    var rows = scenes.map(function (s, i) {
      return '<label class="beatrow"><input type="checkbox" class="beatcb" data-i="' + i + '"' + (beatSel[i] ? " checked" : "") + '> ' + esc(s.h || ("סצנה " + (i + 1))) + '</label>';
    }).join("");
    return '<div class="promptbox"><div class="toolh">בחירת סצנות לניתוח</div>' +
      (scenes.length ? '<div class="beatlist">' + rows + '</div><div class="toolnote"><button class="linkbtn" id="beatsall">בחר הכל</button> · <button class="linkbtn" id="beatsnone">נקה בחירה</button></div>'
        : '<div class="toolnote">אין סצנות בתסריט. פתח את המסמך ורענן.</div>') +
      '<div class="toolh">הבקשה ל-Claude (בחלון התסריט)</div>' +
      '<textarea id="beatstext" readonly rows="5">' + esc(buildBeatsPrompt()) + '</textarea>' +
      '<button class="lb on" id="beatscopy">העתק</button> <span class="chgstat" id="beatscopied"></span>' +
      '<div class="toolh">תשובת Claude</div>' +
      '<textarea id="beatsreply" rows="4" placeholder="הדבק כאן את הרשימה ש-Claude החזיר"></textarea>' +
      '<button class="lb on" id="beatsapply">סמן במרקר</button> <button class="lb" id="beatsclear">נקה סימוני ביטים</button>' +
      '<div class="toolnote" id="beatsstat"></div>' +
      '<div class="toolnote">שורה לכל סימון: סוג | ציטוט. שינוי מצב כחול כהה, צחוק ורוד, התקרבות כחול, התרככות אפור.</div></div>';
  }
  function applyBeats() {
    var items = [], bad = 0;
    ($("beatsreply").value || "").split(/\n/).forEach(function (line) {
      var parts = line.split("|");
      if (parts.length < 2) return;
      var type = parts[0].trim().replace(/^[-*\d.\s]+/, "");
      var q = parts.slice(1).join("|").trim().replace(/^"|"$/g, "").trim();
      if (!BEAT_COLORS[type] || !q) { bad++; return; }
      items.push({ text: q, color: BEAT_COLORS[type] });
    });
    if (!items.length) { $("beatsstat").textContent = "לא נמצאו שורות בפורמט: סוג | ציטוט."; return; }
    $("beatsstat").textContent = "מסמן…";
    W.beatMark(items).then(function (r) {
      $("beatsstat").textContent = "סומנו " + r.marked + (r.missing ? ", לא נמצא טקסט ל-" + r.missing : "") + (bad ? ", שורות לא מובנות: " + bad : "");
    }).catch(showErr);
  }
  // the whole material (synopsis helper note + script text, scene by scene) is copied with the request, so Claude reads it at once
  function loadBundle(tries) {
    smartBundle = { ready: false, msg: "אוסף את הסצנות…", text: "", scenes: 0, missing: 0 };
    if (view === "tools" || view === "compare") renderList();
    if (!W.compareBundle) { smartBundle.msg = "החלונית לא יכולה לקרוא את המסמך כאן."; return; }
    W.compareBundle("s2t").then(function (b) {
      if (!b.helpers && tries > 0) {
        smartBundle.msg = "מחכה להערות העזר מהסינופסיס…";
        if (view === "tools" || view === "compare") renderList();
        setTimeout(function () { if (smartPrompt === "s2t") loadBundle(tries - 1); }, 2000);
        return;
      }
      if (!b.helpers) { smartBundle.msg = "לא נמצאו הערות עזר מהסינופסיס. לחץ שוב על " + "\"סינופסיס ← תסריט\"" + "."; }
      else smartBundle = { ready: true, msg: "", text: b.text, scenes: b.scenes, missing: b.missing };
      if (view === "tools" || view === "compare") renderList();
    }).catch(function (er) { smartBundle.msg = "לא הצלחתי לאסוף: " + (er && er.message ? er.message : er); if (view === "tools" || view === "compare") renderList(); });
  }
  var SUGFMT = "פורמט ההצעה: תשובה (Reply) להערה עצמה, בלי שום טקסט נוסף. שורה ראשונה \"הצעה:\". אחריה אחד משני סוגים. (א) תיקון נקודתי, פעולה אחת או יותר לפי הסדר: להחלפה בתוך משפט, שורה \"במקום: \" עם קטע מדויק מהתסריט ושורה \"יבוא: \" עם הטקסט החדש. למחיקת מילים או משפט, שורה \"מחק: \" עם קטע מדויק. למחיקת פסקה שלמה, שורה \"מחק פסקה: \" עם קטע מדויק מהפסקה. להוספת פסקאות, שורה \"אחרי: \" עם קטע מדויק מהפסקה שאחריה הן נכנסות, ואחריה שורה לכל פסקה חדשה שמתחילה בשם הסגנון בסוגריים מרובעים, [תיאור] או [דמות] או [דיאלוג] או [הערת בימוי]; בלי \"אחרי:\" הן נכנסות אחרי הפסקה שההערה מעוגנת אליה. כל קטע מדויק: מילה במילה מהטקסט המקורי של התסריט, עד 100 תווים. (ב) שכתוב של כל הסצנה, כשהפער גדול או כשתיקון נקודתי ייצא טלאי: שורה \"החלף סצנה:\" ואחריה כל הסצנה החדשה, שורה לכל פסקה שמתחילה ב-[תיאור], [דמות], [דיאלוג] או [הערת בימוי], בלי הכותרת (היא נשארת). השכתוב שומר על מה שעובד בתסריט, על הקול של הדמויות ועל פרטים שסצנות אחרות תלויות בהם, ומשנה רק מה שצריך כדי שהסצנה תעשה את מה שהסינופסיס רוצה ממנה. הערה עם \"החלף סצנה\" חייבת להיות מעוגנת על כותרת הסצנה, ובגוף הסצנה אסור שיהיו הערות פתוחות אחרות. אם יש לך רעיון לתיקון, תמיד כתוב אותו כפעולות או כשכתוב, גם אם זו החלטה של הכותב: הוא יכול לדחות אותו. \"הצעה: אין\" רק כשאין לך שום הצעה, ובשורה הבאה הסבר קצר. אם התיקון סותר משהו בסצנה אחרת, הוסף שורה \"רצף: סצנה N – מה צריך לתקן שם | מקום: \" ואחריו קטע מדויק מהתסריט באותו מקום, עד 60 תווים. תיאור בזמן הווה, קצר וויזואלי; דיאלוג קצר, בקול של הדמויות. ";
  var PROMPTS = {
    s2t: "השוואה סינופסיס ← תסריט. בסוף ההודעה הזו מופיע כל החומר: לכל סצנה, הטקסט שלה בסינופסיס ואחריו הטקסט שלה בתסריט, פסקה בשורה עם הסגנון בסוגריים. עבוד מהחומר הזה. אל תקרא את המסמך כדי להשוות; את המסמך פותחים רק בשלב 2, כדי לכתוב הערות. שלב 1, בצ'אט בלבד, בלי לכתוב במסמך: א. רשימת אירועים לכל סצנה, מהראשונה עד האחרונה, בלי לדלג. כאן הצד הראשון הוא הסינופסיס והצד השני הוא התסריט: פרק את הטקסט של הצד הראשון לרשימה של כל האירועים, אחד אחרי השני: מי נכנס ואיך (בדלת, אחרי דפיקה, כבר בפנים), מה כל אחד מחזיק ומביא, מי עושה מה ולמי, איפה זה קורה (ספה, שטיח, מטבח, פתח), באיזה סדר, ומה כל דמות יודעת או מבינה באותו רגע. בדוק כל אירוע מול הצד השני וסמן: ✓ קיים, ≠ שונה, ✗ חסר. אחר כך הוסף שורות + לכל אירוע שקיים רק בצד השני. אין פרט קטן מדי: שינוי במי מגיע, איך, עם מה, איפה או באיזה סדר הוא תמיד שינוי, ומדווחים עליו גם אם המשמעות שלו לא ברורה מיד. הצג טבלה לכל סצנה: # | בסינופסיס | בתסריט | מה זה משנה. בשורות ✓ השאר את העמודה האחרונה ריקה. בכל שורה ≠, ✗ או + כתוב מה השינוי עושה לסיפור, לקומדיה, לבימוי, לקצב או לדמות (למשל: \"יעל מגיעה עם מנורה במקום קניות: הכפילות עם השקיות של חנה נעלמת, והכעס של יעל על אמה מאבד את הסיבה שלו\"). ב. התמונה הגדולה, אחרי כל הטבלאות, עד עשר שורות: הקשת של כל דמות ראשית בכל צד, ואיפה היא השתנתה או נשברה; מעקב אחרי חפצים וסימנים לאורך כל הסיפור (מי מחזיק אותם, איפה הם, מה קורה להם, למשל פתקים, משקפיים, עוגיות); מה הקהל יודע ומתי בכל צד; וסתירות בתוך התסריט עצמו (חפץ שנמצא בשני מקומות, דמות שיודעת משהו לפני שגילתה אותו). ג. עצור ושאל אותי מה לתקן ברשימה, ועל אילו שינויים לכתוב הערות. אל תמשיך לפני שעניתי. כל שאלה אליי היא שאלה אמריקאית: אם יש לך כלי שאלות עם אפשרויות לבחירה, השתמש בו; אחרת כתוב לכל שאלה אפשרויות ממוספרות (1, 2, 3...) שאפשר לענות עליהן במספר, ההמלצה ראשונה ומסומנת (מומלץ). למשל: על אילו שינויים לכתוב הערות? 1. על כל השינויים (מומלץ) 2. רק על ≠ ו-✗ 3. אכתוב לך מספרי שורות. שלב 2, רק אחרי שעניתי: לכל שינוי שאישרתי, הוסף במסמך הערת Word (Comment) חדשה שמתחילה ב\"השוואה:\", ומיד אחריה הצעה. עגן את ההערה על המשפט בתסריט שהשינוי נוגע בו; שינוי שחסר לגמרי, או סצנה שצריכה שכתוב, על כותרת הסצנה. בהערה: שורה ראשונה \"השוואה: \" ומה השתנה, ושורה שנייה \"מה זה משנה: \" כמו בטבלה. כמה שינויים באותה סצנה שתיקון נפרד שלהם ייצא טלאי: הערה אחת על הכותרת עם שכתוב של כל הסצנה. " + SUGFMT + "אל תשנה טקסט בתסריט, אל תמחק הערות ואל תענה בתוך הערות העזר. אם אתה עוצר באמצע שלב 2, כתוב באיזו סצנה עצרת, כדי שאגיד לך להמשיך. בסוף שלב 2 כתוב רק כמה הערות והצעות הוספת, ושאפשר לעבור עליהן בכרטיס בחלונית.",
    t2s: "השוואה תסריט ← סינופסיס. על כל כותרת סצנה בסינופסיס יש הערת עזר שמתחילה ב\"תסריט:\" ובה הטקסט של הסצנה מהתסריט. קרא את כל הסינופסיס ואת כל הערות העזר פעם אחת, מההתחלה עד הסוף, לפני שאתה כותב משהו. שלב 1, בצ'אט בלבד, בלי לכתוב במסמך: א. רשימת אירועים לכל סצנה, מהראשונה עד האחרונה, בלי לדלג. כאן הצד הראשון הוא התסריט (הערת העזר) והצד השני הוא הסינופסיס: פרק את הטקסט של הצד הראשון לרשימה של כל האירועים, אחד אחרי השני: מי נכנס ואיך (בדלת, אחרי דפיקה, כבר בפנים), מה כל אחד מחזיק ומביא, מי עושה מה ולמי, איפה זה קורה (ספה, שטיח, מטבח, פתח), באיזה סדר, ומה כל דמות יודעת או מבינה באותו רגע. בדוק כל אירוע מול הצד השני וסמן: ✓ קיים, ≠ שונה, ✗ חסר. אחר כך הוסף שורות + לכל אירוע שקיים רק בצד השני. אין פרט קטן מדי: שינוי במי מגיע, איך, עם מה, איפה או באיזה סדר הוא תמיד שינוי, ומדווחים עליו גם אם המשמעות שלו לא ברורה מיד. הצג טבלה לכל סצנה: # | בסינופסיס | בתסריט | מה זה משנה. בשורות ✓ השאר את העמודה האחרונה ריקה. בכל שורה ≠, ✗ או + כתוב מה השינוי עושה לסיפור, לקומדיה, לבימוי, לקצב או לדמות (למשל: \"יעל מגיעה עם מנורה במקום קניות: הכפילות עם השקיות של חנה נעלמת, והכעס של יעל על אמה מאבד את הסיבה שלו\"). ב. התמונה הגדולה, אחרי כל הטבלאות, עד עשר שורות: הקשת של כל דמות ראשית בכל צד, ואיפה היא השתנתה או נשברה; מעקב אחרי חפצים וסימנים לאורך כל הסיפור (מי מחזיק אותם, איפה הם, מה קורה להם, למשל פתקים, משקפיים, עוגיות); מה הקהל יודע ומתי בכל צד; וסתירות בתוך התסריט עצמו (חפץ שנמצא בשני מקומות, דמות שיודעת משהו לפני שגילתה אותו). ג. עצור ושאל אותי מה לתקן ברשימה, ועל אילו שינויים לכתוב הערות. אל תמשיך לפני שעניתי. כל שאלה אליי היא שאלה אמריקאית: אם יש לך כלי שאלות עם אפשרויות לבחירה, השתמש בו; אחרת כתוב לכל שאלה אפשרויות ממוספרות (1, 2, 3...) שאפשר לענות עליהן במספר, ההמלצה ראשונה ומסומנת (מומלץ). למשל: על אילו שינויים לכתוב הערות? 1. על כל השינויים (מומלץ) 2. רק על ≠ ו-✗ 3. אכתוב לך מספרי שורות. שלב 2, רק אחרי שעניתי: לכל שינוי שאישרתי, הוסף הערת Word (Comment) חדשה על המשפט בסינופסיס שהשינוי נוגע בו, או על כותרת הסצנה אם הוא חסר לגמרי. שורה ראשונה \"השוואה: \" ומה השתנה, שורה שנייה \"מה זה משנה: \", ושורה שלישית \"לסינופסיס: \" עם ניסוח מוצע למשפט או לפסקה המעודכנים. אל תשנה טקסט ואל תמחק הערות. בסוף כתוב רק כמה הערות הוספת.",
    instr: "ביצוע ההוראות שלי. בהערות ה-Word שמתחילות ב\"השוואה:\" ולא סומנו כטופלו, מצא את ההערות שהתשובה האחרונה שלהן מתחילה ב\"הוראה:\". זו הוראה שלי לתיקון. לכל הוראה: קרא את הסצנה של ההערה, וחפש במסמך את מילות המפתח של התיקון כדי לשמור על רצף בין הסצנות, והוסף להערה הצעה חדשה. אם ההוראה דורשת שינוי גם בסצנה אחרת, הוסף שם הערת Word חדשה שמתחילה ב\"השוואה: המשך – \" ותיאור קצר, ולה הצעה משלה. " + SUGFMT + "אם אתה מוסיף הערה בסצנה אחרת, עגן אותה על אותו קטע. אל תשנה טקסט בתסריט, אל תסמן הערות כטופלו ואל תמחק הערות. בסוף כתוב בצ'אט רק כמה הצעות הוספת.",
    prep: "הכנת הצעות לכרטיס התיקונים. עבור פעם אחת על כל הערות ה-Word שמתחילות ב\"השוואה:\" ולא סומנו כטופלו, ואין להן עדיין תשובה שמתחילה ב\"הצעה:\". בתחילת העבודה קרא פעם אחת את כל התסריט, כדי לשמור על רצף בין הסצנות. לכל הערה הוסף הצעה. " + SUGFMT + "אם אתה מוסיף הערה בסצנה אחרת, עגן אותה על אותו קטע. אל תשנה טקסט בתסריט, אל תסמן הערות כטופלו ואל תמחק הערות. בסוף כתוב בצ'אט רק כמה הצעות הוספת.",
    walk: "מעבר על ההשוואות. עבור איתי על הערות ה-Word שמתחילות ב\"השוואה:\" ועוד לא סומנו כטופלו, אחת אחת, לפי הסדר שלהן בתסריט מההתחלה לסוף (לא לפי הסדר שבו נכתבו). הערה אחת בכל פעם. לכל הערה: 1. בחר במסמך את הטקסט שההערה מעוגנת אליו, כדי שאקפוץ אליו, וסמן אותו במרקר צהוב. 2. קרא רק את הסצנה של ההערה, לא את כל המסמך, וכתוב \"הערה N מתוך M\", את כותרת הסצנה ואת תוכן ההערה. 3. הצע ניסוח קצר לתוספת או לשינוי בתסריט, בפורמט הקרן ובסגנונות של התבנית, ואמור איפה בדיוק הוא נכנס. לפני ההצעה חפש במסמך (חיפוש, לא קריאה מלאה) את מילות המפתח שלה: חפצים, דמויות, מקומות. אם יש סתירה עם סצנה אחרת, כתוב אותה והצע גם שם תיקון. 4. שאל אותי: להכניס כמו שהצעת, לשנות (ואני אגיד מה), לדלג, או לסמן טופל בלי שינוי. אל תשנה טקסט לפני שעניתי. 5. אחרי שהכנסת שינוי, בחר את הטקסט החדש כדי שאקפוץ אליו וסמן אותו במרקר ירוק. לפני שעוברים להערה הבאה, תמיד הורד את המרקר הצהוב ואת הירוק. 6. אחרי שינוי, או כשאמרתי טופל, סמן את ההערה כטופלה (Resolved). אם אי אפשר, הוסף להערה תשובה (Reply) שמתחילה במילה \"טופל\". בדילוג אל תסמן. אל תמחק הערות ואל תשנה טקסט אחר. אם מוצגים סימוני שינויים מול דראפט (מרקר צהוב או טקסט אפור מחוק), בקש ממני קודם ללחוץ \"הסתר\" בלשונית שינויים."
  };
  var delCmpPending = false, delCmpT = 0;
  function isCmp(n) { return /^השוואה:/.test(n.text || ""); }
  function copyFrom(taId, msgId) {
    var ta = $(taId);
    var done = function () { $(msgId).textContent = "הועתק. עכשיו הדבק ב-Claude (Cmd+V)"; };
    ta.focus(); ta.select();
    try { if (navigator.clipboard && navigator.clipboard.writeText) { navigator.clipboard.writeText(ta.value).then(done, function () { document.execCommand("copy"); done(); }); return; } } catch (er) {}
    try { document.execCommand("copy"); done(); } catch (er2) { $(msgId).textContent = "סמן את הטקסט והעתק ידנית (Cmd+C)"; }
  }

  var STYLES_OK = ["תיאור", "דמות", "דיאלוג", "הערת בימוי", "כותרת סצנה"];
  function sugRaw(n) {
    if (waiting(n)) return "";
    var r = (n.replies || []).filter(function (x) { return /^הצעה:/.test(x); });
    return r.length ? r[r.length - 1] : "";
  }
  // the last of הצעה/הוראה replies decides: an instruction newer than any suggestion waits for Claude
  function lastInstr(n) {
    var rs = (n.replies || []).filter(function (x) { return /^(הצעה|הוראה):/.test(x); });
    var l = rs[rs.length - 1] || "";
    return /^הוראה:/.test(l) ? l.replace(/^הוראה:\s*/, "") : "";
  }
  function waiting(n) { return !!lastInstr(n); }
  // suggestion = ordered operations:
  //   במקום: X / יבוא: Y  -> replace text      מחק: X -> delete text      מחק פסקה: X -> delete the paragraph
  //   אחרי: X + [סגנון] lines -> new paragraphs after that paragraph (no אחרי: = after the noted paragraph)
  function parseSug(raw) {
    if (!raw) return null;
    var ls = raw.replace(/[\r\u000b\u2028]/g, "\n")
      .replace(/\s*(\[(?:תיאור|דמות|דיאלוג|הערת בימוי|כותרת סצנה)\]|החלף סצנה:|במקום:|יבוא:|רצף:|אחרי:|מחק פסקה:|מחק:)/g, "\n$1")
      .split("\n").map(function (l) { return l.trim(); }).filter(Boolean);
    var head = (ls.shift() || "").replace(/^הצעה:\s*/, "");
    var sg = { ops: [], warn: [], info: [] }, add = null, rp = null;
    ls.forEach(function (l) {
      var m = l.match(/^\[([^\]]+)\]\s*(.*)$/);
      if (/^החלף סצנה:/.test(l)) { add = { op: "scene", lines: [] }; sg.ops.push(add); rp = null; var rest = l.replace(/^החלף סצנה:\s*/, ""); if (rest) sg.info.push(rest); }
      else if (m && STYLES_OK.indexOf(m[1].trim()) >= 0) {
        if (!add) { add = { op: "add", after: "", lines: [] }; sg.ops.push(add); }
        add.lines.push({ style: m[1].trim(), text: m[2] });
      } else if (/^אחרי:/.test(l)) { add = { op: "add", after: l.replace(/^אחרי:\s*/, ""), lines: [] }; sg.ops.push(add); }
      else if (/^במקום:/.test(l)) { rp = { op: "replace", old: l.replace(/^במקום:\s*/, ""), neu: "" }; sg.ops.push(rp); add = null; }
      else if (/^יבוא:/.test(l)) { if (rp) rp.neu = l.replace(/^יבוא:\s*/, ""); add = null; }
      else if (/^מחק פסקה:/.test(l)) { sg.ops.push({ op: "delpara", old: l.replace(/^מחק פסקה:\s*/, "") }); add = null; }
      else if (/^מחק:/.test(l)) { sg.ops.push({ op: "delete", old: l.replace(/^מחק:\s*/, "") }); add = null; }
      else if (/^רצף:/.test(l)) {
        var wt = l.replace(/^רצף:\s*/, ""), wq = "", mm = wt.match(/\|\s*מקום:\s*(.+)$/);
        if (mm) { wq = mm[1].trim(); wt = wt.slice(0, mm.index).trim(); }
        var sn = (wt.match(/סצנה\s*(\d+)/) || [])[1];
        sg.warn.push({ text: wt, quote: wq, scene: sn ? +sn : 0 });
      }
      else sg.info.push(l);
    });
    sg.ops = sg.ops.filter(function (o) { return o.op === "add" || o.op === "scene" ? o.lines.length : o.op === "replace" ? o.old && o.neu : !!o.old; });
    // a scene rewrite replaces the whole scene, so it is the only operation
    var sc = sg.ops.filter(function (o) { return o.op === "scene"; })[0];
    if (sc) sg.ops = [sc];
    sg.scene = !!sc;
    sg.ok = sg.ops.length > 0;
    sg.kind = sg.ok ? "ops" : (/אין/.test(head) || !sg.ok ? "none" : "ops");
    return sg;
  }
  var cardEdit = false, cardMsg = "", hlId = "", cardMode = false, instrOpen = false;
  var sceneBefore = null;   // {id, paras} or {id, err}: the current text of the scene, shown above a scene rewrite
  var doneNote = null;   // {id, text, where, snap, undo}: the note just inserted, waiting for "טופל" or "צריך תיקון"
  function markOff() {
    var g = W.clearCardMarks ? W.clearCardMarks().catch(function () {}) : Promise.resolve();
    if (!hlId || !W.markNote) return g;
    var id = hlId; hlId = "";
    return g.then(function () { return W.markNote(id, false).catch(function () {}); });
  }
  function selectNote(id, isCompare) {
    cardEdit = false; cardMsg = ""; instrOpen = false; roundDone = false; doneNote = null;
    return markOff().then(function () {
      noteSel = id;
      renderList();
      var se = document.querySelector('.noteitem.sel');
      if (se && se.scrollIntoView) se.scrollIntoView({ block: "nearest" });
      if (isCompare && W.markNote) { hlId = id; return W.markNote(id, true); }
      return W.gotoNote(id);
    }).catch(showErr);
  }
  function openCmp() { return (notes || []).filter(function (n) { return isCmp(n) && !n.resolved; }); }
  var roundDone = false;
  function endRound() { roundDone = true; cardEdit = false; instrOpen = false; markOff(); noteSel = ""; renderList(); }
  function stepCmp(dir) {
    var co = openCmp();
    if (!co.length) { markOff(); noteSel = ""; renderList(); return; }
    var ix = co.map(function (n) { return n.id; }).indexOf(noteSel);
    if (dir > 0) { if (ix === co.length - 1) { endRound(); return; } ix = ix < 0 ? 0 : ix + 1; }
    else ix = ix < 0 ? co.length - 1 : (ix - 1 + co.length) % co.length;
    selectNote(co[ix].id, true);
  }
  // after a note was handled: go to the next open note that came after it in the order before the change
  function snapOrder() { return { ids: openCmp().map(function (n) { return n.id; }), cur: noteSel }; }
  // a note that is no longer in the document (deleted by hand, by "מחק את כל הערות ההשוואה", or rewritten by Claude):
  // the list is refreshed and the card moves on, instead of a stuck "not saved" message
  function noteFail(er, snap) {
    var m = er && er.message ? er.message : String(er);
    if (/לא נמצאה במסמך/.test(m)) {
      cardMsg = "ההערה הזו כבר לא במסמך (כנראה נמחקה), אז עברתי לבאה.";
      return refreshNotes().then(function () { advanceAfter(snap); renderList(); });
    }
    cardMsg = "לא נשמר: " + m; renderList();
  }
  function advanceAfter(snap) {
    var open = openCmp().map(function (n) { return n.id; });
    var k = snap.ids.indexOf(snap.cur);
    for (var i = k + 1; i < snap.ids.length; i++) if (open.indexOf(snap.ids[i]) >= 0) { selectNote(snap.ids[i], true); return; }
    if (!open.length) { markOff(); noteSel = ""; renderList(); return; }
    endRound();
  }
  function sugHtml(sg) {
    if (!sg) return "";
    var h = sg.ops.map(function (o) {
      if (o.op === "scene") {
        var before = sceneBefore && sceneBefore.id === noteSel ? sceneBefore : null;
        return '<div class="sugi"><b>שכתוב של כל הסצנה</b> (הכותרת נשארת). "הכנס" מחליף את הסצנה, ו"צריך תיקון" מחזיר אותה.</div>' +
          '<div class="chgrow"><button class="linkbtn" id="scenebefore">' + (before ? "הסתר את הסצנה הנוכחית" : "הצג את הסצנה הנוכחית (לפני)") + '</button></div>' +
          (before ? '<div class="scnbefore"><div class="sughead">לפני</div>' + (before.err ? '<div class="sugi">' + esc(before.err) + '</div>' :
            before.paras.map(function (l) { return '<div class="sugl"><span class="sugt">' + esc(l.style) + '</span><s>' + esc(l.text) + '</s></div>'; }).join("")) + '</div><div class="sughead">אחרי</div>' : '') +
          o.lines.map(function (l) { return '<div class="sugl"><span class="sugt">' + esc(l.style) + '</span>' + esc(l.text) + '</div>'; }).join("");
      }
      if (o.op === "replace") return '<div class="sugl findable" data-find="' + esc(o.old) + '"><span class="sugt">במקום</span><s>' + esc(o.old) + '</s></div><div class="sugl"><span class="sugt">יבוא</span>' + esc(o.neu) + '</div>';
      if (o.op === "delete") return '<div class="sugl findable" data-find="' + esc(o.old) + '"><span class="sugt">מחק</span><s>' + esc(o.old) + '</s></div>';
      if (o.op === "delpara") return '<div class="sugl findable" data-find="' + esc(o.old) + '"><span class="sugt">מחק פסקה</span><s>' + esc(o.old) + '</s></div>';
      return (o.after ? '<div class="sugi findable" data-find="' + esc(o.after) + '">אחרי: ' + esc(o.after) + '</div>' : '<div class="sugi">אחרי הפסקה המסומנת</div>') +
        o.lines.map(function (l) { return '<div class="sugl"><span class="sugt">' + esc(l.style) + '</span>' + esc(l.text) + '</div>'; }).join("");
    }).join('<div class="sugsep"></div>');
    h += sg.info.map(function (l) { return '<div class="sugi">' + esc(l) + '</div>'; }).join("");
    h += sg.warn.map(function (w) {
      var can = w.quote || w.scene;
      return '<div class="sugw' + (can ? ' warnjump' : '') + '" data-quote="' + esc(w.quote) + '" data-scene="' + (w.scene || "") + '">⚠ רצף: ' + esc(w.text) + (can ? ' <span class="linkbtn">← קפוץ</span>' : '') + '</div>';
    }).join("");
    return h;
  }

  // ---------- notes: Word comments, with the sync notes first-class
  var notes = null;        // [{id, text, resolved, where, style}]
  var noteSel = "";
  var showDone = false;
  var walkOpen = false;
  var cmtShow = false;
  var notesFolded = lsGet("notesFold", "") === "1";   // the macro hides comments every time the script opens
  var synNoteCount = 0;
  function isSynNote(t) { t = (t || "").trim(); return /^סינופסיס \d+:/.test(t) || /^תסריט:/.test(t); }
  function openNotes() { return notes ? notes.filter(function (n) { return !n.resolved; }).length : 0; }
  function noteKind(t) {
    if (/^השוואה:/.test(t)) return { k: "cmp", l: "השוואה" };
    if (/^לפי סינופסיס|^לפי התסריט/.test(t)) return { k: "loc", l: "לוקיישן השתנה" };
    if (/^סצנה חדשה מ|חסרה בסינופסיס/.test(t)) return { k: "new", l: "סצנה חדשה" };
    if (/^הסצנה הזו לא קיימת/.test(t)) return { k: "gone", l: "לא בסינופסיס" };
    return { k: "other", l: "הערה" };
  }
  function refreshNotes() {
    if (!W.listNotes) return Promise.resolve();
    return W.listNotes().then(function (n) {
      if (n === null) { notes = []; return; }
      synNoteCount = n.filter(function (x) { return isSynNote(x.text); }).length;
      n = n.filter(function (x) { return !isSynNote(x.text); });
      // Claude marks a note handled by replying "טופל" when it can't resolve it itself
      n.forEach(function (x) {
        if (!x.resolved && (x.replies || []).some(function (r) { return /^✓?\s*טופל/.test(r); })) {
          x.resolved = true; W.resolveNote(x.id, true).catch(function () {});
        }
      });
      var changed = JSON.stringify(n) !== JSON.stringify(notes);
      notes = n;
      if (changed) { renderBar(); if ((view === "notes" || view === "compare") && !cardEdit && !instrOpen) renderList(); }
    }).catch(function () {});
  }
  function promptBox() {
    if (!walkOpen) return "";
    return '<div class="promptbox"><textarea id="walktext" readonly rows="5">' + esc(PROMPTS[walkOpen]) + '</textarea>' +
      '<button class="lb on" id="walkcopy">העתק</button> <span class="chgstat" id="walkcopied"></span>' +
      '<div class="toolnote">' + (walkOpen === "instr" ? 'מדביקים ל-Claude ב-Word. הוא הופך כל הוראה להצעה חדשה, ואז עוברים שוב בכרטיס ולוחצים "הכנס".' : walkOpen === "prep" ? 'מדביקים ל-Claude ב-Word. הוא מוסיף לכל הערה תשובה עם הצעה, ואז עוברים כאן בכרטיס.' : 'מדביקים ל-Claude ב-Word. הוא מציג הערה אחת, מציע ושואל לפני כל שינוי.') + '</div>' +
      '<div class="toolnote">לחיסכון: לפני שמדביקים, בחר Sonnet בתפריט המודל של Claude.</div></div>';
  }
  function renderCard(list, cmpOpen) {
    var pos = cmpOpen.map(function (n) { return n.id; }).indexOf(noteSel);
    if (pos < 0) pos = 0;
    var cur = cmpOpen[pos];
    var raw = sugRaw(cur), sg = parseSug(raw);
    var pct = Math.round(100 * pos / cmpOpen.length);
    var html = '<div class="card">' +
      '<div class="cardtop"><button class="lb" id="cardlist">☰ רשימה</button>' +
      '<span class="cardpos">השוואה ' + (pos + 1) + ' מתוך ' + cmpOpen.length + '</span>' +
      '<span class="cardnav"><button class="lb" id="cmpprev" title="הקודמת">→</button><button class="lb" id="cmpnext" title="הבאה">←</button></span></div>' +
      '<div class="cardbar"><span style="width:' + pct + '%"></span></div>' +
      '<div class="cardwhere">' + esc(cur.where || "") + '</div>' +
      '<div class="cardwhat">' + esc(cur.text.replace(/^השוואה:\s*/, "")) + '</div>';
    if (waiting(cur)) {
      html += '<div class="sugbox waitbox"><div class="sughead">מחכה ל-Claude</div><div class="sugl">' + esc(lastInstr(cur)) + '</div>' +
        '<div class="sugi">שולחים מלשונית ההערות: "שלח הוראות ל-Claude"</div></div>';
    } else if (raw) {
      html += '<div class="sugbox"><div class="sughead">' + (sg.kind === "none" ? "Claude: צריך החלטה שלך" : "ההצעה של Claude") +
        (cardEdit ? '' : ' <button class="linkbtn" id="cardedit">ערוך</button>') + '</div>' +
        (cardEdit ? '<textarea id="cardtext" rows="6">' + esc(raw) + '</textarea><div class="toolnote">שורה לכל פעולה: במקום: + יבוא: · מחק: · מחק פסקה: · אחרי: ואחריו שורות [תיאור] / [דמות] / [דיאלוג] · או החלף סצנה: ואחריו כל הסצנה</div>' +
          '<div class="editbtns"><button class="lb on" id="cardsave">שמור עריכה</button> <button class="lb" id="cardcancel">בטל עריכה</button></div>' : sugHtml(sg)) + '</div>';
    } else {
      html += '<div class="sugbox empty2">אין עדיין הצעה להערה הזו.<br><button class="linkbtn" id="prepshow">הכן הצעות עם Claude</button></div>';
    }
    if (instrOpen) {
      html += '<div class="instrbox"><div class="sughead">הוראה ל-Claude</div>' +
        '<textarea id="instrtext" rows="4" placeholder="למשל: אור כן רואה את החיבוק, אבל רק בסוף הסצנה. תעדכן בהתאם גם את סצנה 10."></textarea>' +
        '<div class="chgrow"><button class="lb on" id="instrsave">שמור והמשך</button><button class="linkbtn" id="instrcancel">בטל</button></div></div>';
    }
    if (cardMsg) html += '<div class="toolnote cardmsg">' + esc(cardMsg) + '</div>';
    var canIns = raw && (sg.ok || cardEdit);
    if (raw && !sg.ok && !cardEdit && !instrOpen) {
      html += '<div class="decide"><button class="lb on" id="dosug">כמו בסינופסיס</button><button class="lb" id="keepscr">כמו בתסריט</button></div>';
    }
    html += '<div class="cardact">' +
      '<button class="bigbtn on" id="cardins"' + (canIns ? '' : ' disabled') + '>' + (cardEdit ? "הכנס ערוך" : "הכנס") + '</button>' +
      '<button class="bigbtn" id="cardskip">דלג</button>' +
      '<button class="bigbtn" id="cardok">✓ טופל</button></div>' +
      (instrOpen ? '' : '<button class="bigbtn instrbtn" id="instropen">✎ הוראה ל-Claude</button>') +
      (function () { var miss = cmpOpen.filter(function (n) { return !sugRaw(n); }).length;
        return '<div class="cardfoot">' + (miss ? '<button class="linkbtn" id="prepshow">הכן הצעות ל-' + miss + ' הערות שחסרות</button> · ' : '') + '<button class="linkbtn" id="walkshow">מעבר בצ\'אט</button></div>'; })() +
      promptBox() + '</div>';
    var sy = list.scrollTop;
    list.innerHTML = html;
    list.scrollTop = sy;
  }
  function renderNotes(list) {
    if (!notes) { list.innerHTML = '<div class="empty">טוען הערות…</div>'; return; }
    var cmpOpen = notes.filter(function (n) { return isCmp(n) && !n.resolved; });
    var nCmpAll = notes.filter(isCmp).length;
    if (cardMode && doneNote) {
      list.innerHTML = '<div class="card"><div class="cardtop"><button class="lb" id="cardlist">☰ רשימה</button><span class="cardpos">✓ הוכנס</span></div>' +
        '<div class="cardwhere">' + esc(doneNote.where || "") + '</div>' +
        '<div class="cardwhat">' + esc(doneNote.text.replace(/^השוואה:\s*/, "")) + '</div>' +
        '<div class="sugbox donebox">השינוי מסומן בירוק במסמך. בדוק אותו: אם הוא טוב, "טופל". אם לא, "צריך תיקון" מחזיר את הטקסט למה שהיה ופותח הוראה ל-Claude.</div>' +
        (cardMsg ? '<div class="toolnote cardmsg">' + esc(cardMsg) + '</div>' : '') +
        '<div class="cardact" style="grid-template-columns:1fr 1fr"><button class="bigbtn on" id="doneok">✓ טופל</button><button class="bigbtn" id="donefix">צריך תיקון</button></div></div>';
      return;
    }
    if (cardMode && cmpOpen.length && roundDone) {
      list.innerHTML = '<div class="card"><div class="cardtop"><button class="lb" id="cardlist">☰ רשימה</button><span class="cardpos">סיימת סיבוב</span></div>' +
        '<div class="cardbar"><span style="width:100%"></span></div>' +
        '<div class="sugbox empty2">נשארו ' + cmpOpen.length + ' השוואות פתוחות (דילגת עליהן או שהן מחכות ל-Claude).</div>' +
        '<button class="bigbtn on" id="cardrestart">התחל שוב מהראשונה שפתוחה</button></div>';
      return;
    }
    if (cardMode && cmpOpen.length) { renderCard(list, cmpOpen); return; }
    if (!cmpOpen.length) cardMode = false;
    var shown = notes.filter(function (n) { return showDone || !n.resolved; });
    var html = '<div class="chgctl"><div class="chgrow"><span class="chgstat">' +
      (openNotes() ? openNotes() + " הערות פתוחות" : "אין הערות פתוחות") + '</span><span class="sp"></span>' +
      '<label><input type="checkbox" id="notedone"' + (showDone ? " checked" : "") + '> הצג גם שטופלו</label></div>' +
      '<div class="chgrow"><button class="lb" id="cmtview">' + (cmtShow ? "הסתר הערות במסמך" : "הצג הערות גם במסמך") + '</button></div>' +
      (nCmpAll && W.deleteCompareNotes ? '<div class="chgrow"><button class="lb del' + (delCmpPending ? ' on' : '') + '" id="delcmp">' +
        (delCmpPending ? 'למחוק ' + nCmpAll + ' הערות השוואה? לחץ שוב לאישור' : '🗑 מחק את כל הערות ההשוואה (' + nCmpAll + ')') + '</button>' +
        (delCmpPending ? ' <button class="linkbtn" id="delcmpno">ביטול</button>' : '') + '</div>' : '') +
      (synNoteCount ? '<div class="chgrow"><span class="chgstat">' + synNoteCount + ' הערות עזר מוסתרות (להסרה: כלים ← סיום השוואה)</span></div>' : '') + '</div>';
    if (cmpOpen.length) {
      var withSug = cmpOpen.filter(function (n) { return !!sugRaw(n); }).length;
      var nWait = cmpOpen.filter(waiting).length;
      var nSyn = notes.filter(function (n) { return (n.replies || []).some(function (r) { return /לעדכן את הסינופסיס:/.test(r); }); }).length;
      if (nSyn) html += '<div class="chgrow"><span class="chgstat">' + nSyn + ' נקודות סומנו "כמו בתסריט" ומחכות לעדכון בסינופסיס</span></div>';
      if (nWait) html += '<div class="startbox waitbox"><span class="chgstat">' + nWait + (nWait === 1 ? ' הוראה מחכה' : ' הוראות מחכות') + ' ל-Claude</span>' +
        '<button class="bigbtn on" id="instrsend">שלח הוראות ל-Claude</button>' + (walkOpen === "instr" ? promptBox() : '') + '</div>';
      html += '<div class="startbox"><button class="bigbtn on" id="cardstart">▶ עבור על ההשוואות</button>' +
        '<span class="chgstat">' + cmpOpen.length + ' פתוחות' + (withSug ? ', ' + withSug + ' עם הצעה מוכנה' : ', עוד בלי הצעות') + '</span>' +
        '<div class="cardfoot">' + (withSug < cmpOpen.length ? '<button class="linkbtn" id="prepshow">הכן הצעות ל-' + (cmpOpen.length - withSug) + ' הערות שחסרות</button> · ' : '') + '<button class="linkbtn" id="walkshow">מעבר בצ\'אט</button></div>' +
        (walkOpen !== "instr" ? promptBox() : '') + '</div>';
    }
    if (!shown.length) {
      html += '<div class="empty">' + (notes.length ? "כל ההערות טופלו." : "אין הערות במסמך.") +
        '<br>הערות נוספות כשמעדכנים את התסריט מהסינופסיס (כלים).</div>';
    } else {
      html += '<div class="chglist">' + shown.map(function (n, k) {
        var kd = noteKind(n.text);
        var body = n.text.split("\n")[0].replace(/\r/g, " ");
        return '<div class="noteitem' + (noteSel === n.id ? " sel" : "") + (n.resolved ? " done" : "") + '" data-id="' + esc(n.id) + '">' +
          '<button class="notego" data-id="' + esc(n.id) + '"><span class="nk nk-' + kd.k + '">' + kd.l + '</span>' +
          '<span class="nw">' + esc(n.where || "(שורה ריקה)") + '</span>' +
          '<span class="nt">' + esc(body.slice(0, 140)) + '</span></button>' +
          '<button class="lb noteok" data-id="' + esc(n.id) + '" title="' + (n.resolved ? "פתח מחדש" : "סמן כטופל") + '">' + (n.resolved ? "↺" : "✓ טופל") + '</button></div>';
      }).join("") + '</div>';
    }
    var foldHead = '<div class="chgrow"><button class="linkbtn" id="notesfold">' + (notesFolded ? "▸ הצג את כל ההערות" : "▾ מזער את כל ההערות") + '</button></div>';
    html = notesFolded ? foldHead : foldHead + html;
    var y = window.scrollY, sy = list.scrollTop;
    list.innerHTML = html;
    window.scrollTo(0, y); list.scrollTop = sy;
  }
  // ---------- filters for the personal and feedback lists (colour, character, part of the structure)
  var flt = { c: "", ch: "", st: "", cat: "" }, charNames = null, colOpen = false;
  function fltColorOf(t, isFb) {
    return stageColorOf(t);
  }
  function fltStageOf(t) {
    var sc = t.sn ? scenes[t.sn - 1] : (t.scene ? scenes.filter(function (x) { return x.h === t.scene; })[0] : null);
    return sc ? stNorm(sc.st) : "";
  }
  function fltOk(t, isFb) {
    if (flt.c) { var cc = fltColorOf(t, isFb) || "none"; if (cc !== flt.c) return false; }
    if (flt.cat && isFb && fbCatName(t) !== flt.cat) return false;
    if (flt.ch) { var chOk = t.ch ? t.ch.split(",").map(function (x) { return x.trim(); }).indexOf(flt.ch) >= 0 : (t.text || "").indexOf(flt.ch) >= 0; if (!chOk) return false; }
    if (flt.st) { var ss = t.pt ? stNorm(t.pt) : (fltStageOf(t) || "לא שויך"); if (ss !== flt.st) return false; }
    return true;
  }
  function loadCharNames() {
    if (charNames !== null || !W.paragraphs) return;
    charNames = [];
    W.paragraphs().then(function (ps) {
      var seen = {};
      ps.forEach(function (p) {
        if (p.style !== "דמות") return;
        var n = (p.text || "").replace(/\(.*\)/g, "").trim();
        if (n && !seen[n]) { seen[n] = true; charNames.push(n); }
      });
      renderList();
    }).catch(function () {});
  }
  function fbCharOpts(cur) {
    var names = (charNames || []).slice();
    if (cur && names.indexOf(cur) < 0) names.push(cur);
    return names.map(function (n) { return '<option value="' + esc(n) + '"' + (cur === n ? " selected" : "") + '>' + esc(n) + '</option>'; }).join("");
  }
  var partsOpen = false;
  function legendHtml() {
    return '<div class="fblegend"><span class="lbl">מקרא:</span> ' +
      STAGES.map(function (x) {
        var c = STAGE_COLORS[x];
        return '<button class="fbleg' + (flt.c === c ? " on" : "") + '" data-c="' + c + '"><i style="background:' + c + '"></i>' + x + '</button>';
      }).join("") + '</div>';
  }
  function fltBar(isFb) {
    loadCharNames();
    var cOpts = [["", "כל הצבעים"], ["#ef4444", "לשנות"], ["#22c55e", "להוסיף"], ["#f59e0b", "להסיר"], ["#3b82f6", "לבדוק"], ["#a855f7", "כללי"]];
    var stOpts = [""].concat(STAGES, ["לא שויך"]);
    return '<div class="fltbar">' +
      '<select id="fltch"><option value="">כל הדמויות</option>' + (charNames || []).map(function (n) { return '<option value="' + esc(n) + '"' + (flt.ch === n ? " selected" : "") + '>' + esc(n) + '</option>'; }).join("") + '</select>' +
      '</div>';
  }

  // ---------- תשובות: answers from Claude kept in the document, to come back to later
  var ansItems = null, ansOpen = -1, ansTitle = "", ansBody = "";
  function saveAnswers() { W.saveAnswers(ansItems); }
  function renderAnswers(list) {
    if (!ansItems) ansItems = W.loadAnswers() || [];
    var html = '<div class="ansadd"><input id="ansTitle" type="text" placeholder="כותרת קצרה" value="' + esc(ansTitle) + '">' +
      '<textarea id="ansBody" rows="4" placeholder="הדבק כאן את התשובה של Claude">' + esc(ansBody) + '</textarea>' +
      '<button class="lb on" id="ansAdd">שמור תשובה</button></div>';
    html += '<div class="chglist">' + ansItems.map(function (a, k) {
      return '<div class="todoitem ansrow" data-i="' + k + '">' +
        '<div class="todobody"><div class="todotext ansopen" data-i="' + k + '">' + esc(a.title || "(בלי כותרת)") + '</div>' +
        (ansOpen === k ? '<div class="ansbody">' + esc(a.body || "") + '</div>' : '') + '</div>' +
        '<div class="todoacts"><button class="linkbtn ansdel" data-i="' + k + '">מחק</button></div></div>';
    }).join("") + '</div>';
    if (!ansItems.length) html += '<div class="empty">עוד אין תשובות שמורות.</div>';
    list.innerHTML = html;
  }
  // ---------- משוב: a list from a conversation with the lecturer (own tab, own storage)
  var fbItems = null, fbPaste = "", fbPick = -1, fbEdit = -1, todoPick = -1, fbMeta = -1, todoMeta = -1;
  // colour per category Claude writes in brackets: set from the start, the user can still change it per item
  var FB_CAT_COLORS = { "לשנות": "#ef4444", "להוסיף": "#22c55e", "להסיר": "#f59e0b", "לבדוק": "#3b82f6", "כללי": "#a855f7" };
  var STAGE_COLORS = { "מחולל": "#ef4444", "אקפוזיציה": "#f59e0b", "מפנה 1": "#22c55e", "מפנה 2": "#3b82f6", "שיא": "#a855f7", "התרה": "#64748b" };
  function stageColorOf(t) {
    var st = stNorm(t.pt) || fltStageOf(t);
    return STAGE_COLORS[st] || "";
  }
  function fbCatName(t) {
    if (t.cat) return t.cat;
    var m = (t.text || "").match(/\[([^\]]+)\]/);
    return m ? m[1].trim() : "";
  }
  function fbCatColor(t) {
    var n = fbCatName(t);
    return FB_CAT_COLORS[n] ? FB_CAT_COLORS[n] : "";
  }
  function renderFeedback(list) {
    if (!fbItems) fbItems = W.loadFeedback() || [];
    var open = fbItems.filter(function (t) { return !t.done; }).length;
    var html = '<div class="chgrow"><button class="lb on" id="fbopen">העתק סיכום שיחה עם המרצה</button> <span class="chgstat" id="fbopened"></span></div>' +
      '<details class="sumimp"><summary>הוסף רשימה מ-Claude</summary><div class="sumbody">' +
      '<textarea id="fbpaste" rows="4" placeholder="הדבק כאן את הרשימה, שורה לכל פריט. למשל: [לשנות] …">' + esc(fbPaste) + '</textarea>' +
      '<button class="lb on" id="fbpasteok">הוסף לרשימה</button> <span class="chgstat" id="fbres"></span></div></details>' +
      '<div class="chgctl"><div class="chgrow"><span class="chgstat">' + open + ' פתוחים מתוך ' + fbItems.length + '</span></div></div>';
    html += fltBar(true);
    html += legendHtml();
    html += '<datalist id="charlist">' + (charNames || []).map(function (n) { return '<option value="' + esc(n) + '">'; }).join("") + '</datalist>';
    html += '<div class="chglist">' + fbItems.map(function (t, k) {
      if (!fltOk(t, true)) return "";
      var cols = STAGES.map(function (x) { return STAGE_COLORS[x]; });
      var curc = stageColorOf(t);
      var catNow = fbCatName(t);
      var dots = '';
      var curBtn = '<button class="fbcur" data-i="' + k + '" title="שנה צבע" style="background:' + (curc || "transparent") + '"></button>';
      var col = stageColorOf(t);
      return '<div class="todoitem' + (t.done ? " done" : "") + '" data-i="' + k + '"' + (col ? ' style="border-inline-start:6px solid ' + col + '"' : "") + '>' +
        '<button class="askicon askclaude-fb" data-i="' + k + '" title="שלח ל-Claude">✳</button>' +
        '<button class="todocheck fbcheck" data-i="' + k + '">' + (t.done ? "✓" : "") + '</button>' +
        '<div class="todobody"><textarea class="fbta" data-i="' + k + '" rows="2">' + esc(String(t.text || "").replace(/^\s*\[[^\]]+\]\s*/, "")) + '</textarea>' +
        (fbMeta === k ? '<div class="fbmeta"><input class="fbch" list="charlist" data-i="' + k + '" value="' + esc(t.ch || "") + '" placeholder="דמות, דמות">' +
        '</div>' : '') +
        '</div><div class="todoacts">' + (t.sn ? '<button class="linkbtn fbscene" data-sn="' + t.sn + '">📍 סצנה ' + t.sn + '</button>' : '') + '' + '<div class="fbcap" data-i="' + k + '">' + esc([t.ch, (t.pt || fltStageOf(t))].filter(Boolean).join(" · ")) + '</div>' + curBtn + '<div class="fbdots">' + dots + '</div><button class="linkbtn fbdel" data-i="' + k + '">מחק</button></div></div>';
    }).join("") + '</div>';
    if (!fbItems.length) html += '<div class="empty">עוד אין פריטים מהשיחה עם המרצה.</div>';
    list.innerHTML = html;
  }
  function saveFeedback() { W.saveFeedback(fbItems); }

  // ---------- personal list (שלי): thoughts and things to work on, kept in the document
  var todos = null, todoDraft = "", todoLinking = "";
  function saveTodos() { W.saveTodo(todos); }
  function renderTodo(list) {
    if (!todos) { todos = W.loadTodo() || []; }
    var open = todos.filter(function (t) { return !t.done; }).length;
    var html = '<div class="chgctl"><div class="chgrow"><span class="chgstat">' + open + ' פתוחים מתוך ' + todos.length + '</span></div>' +
      '<div class="todoadd"><input id="todoin" type="text" placeholder="מחשבה או משהו לעבוד עליו…" value="' + esc(todoDraft) + '">' +
      '<button class="lb on" id="todoadd">הוסף</button></div></div>';
    html = fltBar(true) + legendHtml() + html;
    html += '<div class="chglist">' + todos.map(function (t, k) {
      if (!fltOk(t, false)) return "";
      return '<div class="todoitem' + (t.done ? " done" : "") + '" data-i="' + k + '"' + (stageColorOf(t) ? ' style="border-inline-start:6px solid ' + stageColorOf(t) + '"' : "") + '>' +
        '<button class="todocheck" data-i="' + k + '">' + (t.done ? "✓" : "") + '</button>' +
        '<div class="todobody"><textarea class="ptta" data-i="' + k + '" rows="2">' + esc(t.text) + '</textarea>' +
        (t.scene ? '<button class="linkbtn todoscene" data-i="' + k + '">📍 ' + esc(t.scene) + '</button>' : '') +
        (todoLinking === String(k) ? '<div class="todolink">בחר כותרת סצנה בתסריט, ואז לחץ "קשר כאן"<br><button class="lb on" id="todolinkok" data-i="' + k + '">קשר לסצנה הנוכחית</button></div>' : '') +
        (todoMeta === k ? '<div class="fbmeta"><input class="todoch" data-i="' + k + '" value="' + esc(t.ch || "") + '" placeholder="דמות, דמות" list="charlist"></div>' : '') +
        '</div><div class="todoacts"><button class="fbcur todocur" data-i="' + k + '" title="שנה צבע (חלק במבנה)" style="background:' + (stageColorOf(t) || "transparent") + '"></button>' +
        '<span class="fbcap todocap" data-i="' + k + '">' + esc([t.ch, stNorm(t.pt)].filter(Boolean).join(" · ")) + '</span>' +
        (todoPick === k ? '<div class="stagepicks">' + STAGES.map(function (x) { var c = STAGE_COLORS[x]; return '<button class="stagepick todocol' + (stageColorOf(t) === c ? " on" : "") + '" data-i="' + k + '" data-s="' + x + '"><i style="background:' + c + '"></i>' + x + '</button>'; }).join("") + '</div>' : '') +
        '<button class="askicon askclaude-todo" data-i="' + k + '" title="שלח ל-Claude">✳</button><button class="linkbtn todolinkbtn" data-i="' + k + '">קשר</button>' +
        '<button class="linkbtn tododel" data-i="' + k + '">מחק</button></div></div>';
    }).join("") + '</div>';
    if (!todos.length) html += '<div class="empty">אין עדיין פריטים. כתוב כאן מחשבה, והיא תישאר במסמך.</div>';
    var sy = list.scrollTop;
    list.innerHTML = html;
    list.scrollTop = sy;
  }
  function copyAndFlash(btn, text) {
    var done = function () { var old = btn.textContent; btn.textContent = "הועתק ✓"; setTimeout(function () { btn.textContent = old; }, 1800); };
    try { if (navigator.clipboard && navigator.clipboard.writeText) { navigator.clipboard.writeText(text).then(done, function () { fallbackCopy(text); done(); }); return; } } catch (e) {}
    fallbackCopy(text); done();
  }
  function fallbackCopy(text) {
    var ta = document.createElement("textarea"); ta.value = text; document.body.appendChild(ta); ta.select();
    try { document.execCommand("copy"); } catch (e) {}
    document.body.removeChild(ta);
  }
  function todoAdd() {
    var t = (todoDraft || "").trim();
    if (!t) return;
    todos.push({ text: t, done: false, scene: "" });
    todoDraft = ""; saveTodos(); renderList();
  }

  function renderTools(list, only) {
    list.innerHTML = TOOLS.filter(function (g) { return only ? g.h.indexOf(only) === 0 : g.h.indexOf("השוואה חכמה") !== 0; }).map(function (g) {
      return '<div class="toolgrp"><div class="toolh">' + esc(g.h) + '</div>' + g.items.map(function (it) {
        return '<button class="toolbtn' + (toolPending === it.cmd ? " busy" : "") + '" data-cmd="' + it.cmd + '"><b>' + esc(it.t) + '</b><span>' +
          esc(toolPending === it.cmd ? "נשלח… (עד 2 שניות)" : it.d) + '</span></button>';
      }).join("") + (g.note ? '<div class="toolnote">' + esc(g.note) + '</div>' : "") +
        (g.h === "ניתוח ביטים" && beatsOpen ? beatsBox() : "") +
        (g.h.indexOf("השוואה חכמה") === 0 && smartPrompt ? '<div class="promptbox"><div class="toolh">הבקשה ל-Claude (' + (smartPrompt === "s2t" ? "בחלון התסריט" : "בחלון הסינופסיס") + ')</div>' +
          (smartPrompt === "s2t" && !smartBundle.ready ? '<div class="toolnote">' + esc(smartBundle.msg || "אוסף את הסצנות…") + '</div>' :
          '<textarea id="smarttext" readonly rows="5">' + esc(smartPrompt === "s2t" ? PROMPTS.s2t + "\n\n===== החומר =====\n\n" + smartBundle.text : PROMPTS[smartPrompt]) + '</textarea>' +
          (smartPrompt === "s2t" ? '<div class="toolnote">בבקשה יש את כל הסצנות: ' + smartBundle.scenes + ' סצנות, סינופסיס מול תסריט' + (smartBundle.missing ? ' (' + smartBundle.missing + ' בלי הערת עזר)' : '') + '. <button class="linkbtn" id="smartrebuild">אסוף שוב</button></div>' : '') +
          '<button class="lb on" id="smartcopy">העתק</button> <span class="chgstat" id="smartcopied"></span>') +
          '<div class="toolnote">Claude עונה קודם בצ\'אט בטבלת אירועים לכל סצנה ובתמונה הגדולה, ושואל אותך. רק אחרי שתאשר הוא כותב הערות והצעות. מומלץ Opus.</div></div>' : "") + '</div>';
    }).join("");
  }

  function renderChanges(list) {
    var last = chg && chg.last ? chg.last : 0;
    var vs = chg && chg.vs ? chg.vs : 0;
    var opts = "";
    for (var k = last; k >= 1; k--) opts += '<option value="' + k + '"' + ((chgSel || vs || last) === k ? " selected" : "") + '>דראפט ' + k + '</option>';
    var html = '<div class="chgctl">' +
      '<div class="chgrow"><span>מול</span><select id="chgdraft">' + (opts || '<option>טוען…</option>') + '</select>' +
      '<button class="lb on" id="chgshow">הצג שינויים</button><button class="lb" id="chghide">הסתר</button></div>' +
      '<div class="chgrow"><button class="lb" id="chgprev" title="F6">▲ הקודם</button><button class="lb" id="chgnext" title="F7">▼ הבא</button>' +
      '<span class="chgstat">' + (chgPending ? "מעבד…" : (marks2 ? (marks2.length ? (markIdx ? "שינוי " + markIdx + " מתוך " : "") + marks2.length + " שינויים" : "אין שינויים מוצגים") : (chg && chg.n > 0 ? chg.n + " שינויים מול דראפט " + vs : (chg ? "אין שינויים מוצגים" : "")))) + '</span></div>' +
      (window.CHGDBG && (!chg || !chg.last) ? '<div class="chgstat" style="direction:ltr">' + esc(window.CHGDBG) + '</div>' : '') + '</div>';
    var items = marks2 || (chg ? chg.items : []);
    if (items && items.length) {
      html += '<div class="chglist">' + items.map(function (it, k) {
        return '<button class="chgitem' + (markIdx === k + 1 ? " sel" : "") + '" data-k="' + (k + 1) + '"' + (it.name ? ' data-name="' + esc(it.name) + '"' : ' data-pos="' + it.pos + '"') + '><span class="ci">' + (k + 1) + '</span><span class="ct">' + esc(it.text || "(שורה ריקה)") + '</span></button>';
      }).join("") + '</div>';
    } else if (chg || marks2) {
      html += '<div class="empty">בחר דראפט ולחץ "הצג שינויים".<br>מה שנוסף יסומן במרקר צהוב, מה שנמחק באפור מחוק.</div>';
    }
    list.innerHTML = html;
  }
  function refreshChanges() {
    if (W.listChangeMarks) {
      W.listChangeMarks().then(function (m) {
        if (m === null) return;
        var sig = JSON.stringify(m);
        if (sig !== JSON.stringify(marks2) || chgPending) {
          if (chgPending && sig !== JSON.stringify(marks2)) chgPending = "";
          marks2 = m;
          if (markIdx > m.length) markIdx = 0;
          if (view === "changes") renderList();
        }
      }).catch(function () {});
    }
    if (!W.readChanges) return Promise.resolve();
    return W.readChanges().then(function (c) {
      var sig = JSON.stringify(c);
      if (sig !== JSON.stringify(chg) || chgPending) {
        if (c && chgPending && c.stamp !== (chg && chg.stamp)) chgPending = "";
        chg = c || { vs: 0, last: 0, n: 0, items: [] };
        if (view === "changes") renderList();
      }
    }).catch(function () {});
  }
  function chgRequest(cmd, arg) {
    chgPending = cmd;
    if (view === "changes") renderList();
    return W.sendRequest(cmd, arg || "").then(function () {
      setTimeout(refreshChanges, 2500);
      setTimeout(function () { chgPending = ""; refreshChanges(); if (view === "changes") renderList(); }, 7000);
    }).catch(showErr);
  }

  function renderList() {
    var list = $("scenes");
    if (view === "changes") { list.className = "view-changes"; renderChanges(list); return; }
    if (view === "tools") { list.className = "view-tools"; renderTools(list); return; }
    if (view === "notes") { view = "compare"; lsSet("view", view); renderList(); return; }
    if (view === "todo") { list.className = "view-notes"; renderTodo(list); return; }
    if (view === "compare") {
      list.className = "view-notes";
      list.innerHTML = '<div id="cmpN"></div><div id="cmpT"></div>';
      renderNotes($("cmpN")); renderTools($("cmpT"), "השוואה חכמה"); return;
    }
    if (view === "feedback") { list.className = "view-notes"; renderFeedback(list); return; }
    if (view === "answers") { list.className = "view-notes"; renderAnswers(list); return; }
    list.className = view === "struct" ? "view-struct" : "";
    if (!scenes.length) {
      list.innerHTML = '<div class="empty">אין עדיין סצנות.<br>כותרת סצנה חדשה: Tab עד שמגיעים לסצנה</div>';
      return;
    }
    var visible = function (i) {
      if (editing === i) return true;
      return (!filter || scenes[i].p.indexOf(filter) >= 0) && (!mfilter || scenes[i].m === mfilter);
    };
    var html = "";
    if (view === "struct") {
      html += '<details class="sumimp"><summary>תקצירים מ-Claude</summary>' +
        '<div class="sumbody"><button class="lb on" id="sumcopy">העתק בקשה: תקציר לכל סצנה</button> <span class="chgstat" id="sumcopied"></span>' +
        '<textarea id="sumpaste" rows="4" placeholder="הדבק כאן את רשימת Claude, שורה לכל סצנה: סצנה 1: ..."></textarea>' +
        '<button class="lb" id="sumapply">הכנס תקצירים לסצנות ריקות</button> <span class="chgstat" id="sumres"></span></div></details>';
    }
    if (view === "struct") {
      var byStage = {}; STAGES.forEach(function (x) { byStage[x] = []; }); byStage[""] = [];
      scenes.forEach(function (sc, i) { var sn2 = stNorm(sc.st); (byStage[STAGES.indexOf(sn2) >= 0 ? sn2 : ""]).push(i); });
      var allOpen = STAGES.every(function (x) { return stageOpen[x]; });
      html += '<div class="partsbar"><span>חלקי המבנה</span><button class="linkbtn" id="partsall">' + (allOpen ? "כווץ הכול" : "פתח הכול") + '</button></div>';
      STAGES.concat([""]).forEach(function (stg) {
        var open = !!stageOpen[stg] && stg !== "";
        if (stg === "" && !byStage[""].length) return;
        html += '<div class="part' + (open ? " open" : "") + '" data-stage="' + esc(stg) + '">' +
          '<div class="parthead"><span class="arr">' + (open ? "▾" : "◂") + '</span> ' + (stg || "לא שויך") + '</div>';
        if (open || stg === "") {
          if (!byStage[stg].length) html += '<div class="stdrop">גרור לכאן סצנה מהרשימה</div>';
          html += byStage[stg].map(function (i) {
            var sc = scenes[i];
            var sumPart = sumEdit === i
              ? '<textarea class="srta" data-i="' + i + '" rows="2">' + esc(sc.s || "") + '</textarea>'
              : (sc.s ? '<div class="srsum" data-i="' + i + '">' + esc(sc.s) + '</div>' : '<div class="srsum srempty" data-i="' + i + '">לחץ כדי לכתוב תקציר</div>');
            var mkc = sc.m ? markById(sc.m) : null;
            return '<div class="strow' + (mkc ? " marked" : "") + '" draggable="true" data-i="' + i + '"' + (mkc ? ' style="border-inline-start:6px solid ' + mkc.c + '"' : "") + '>' +
              '<span class="srnum">' + (i + 1) + '</span><span class="srhead">' + esc(sc.h) + '</span>' +
              sumPart + '</div>';
          }).join("");
        }
        html += '</div>';
      });
    }
    if (view === "struct") {
      // the scenes already show inside the parts above, so the old sequence list is not repeated here
      var seqs = [], n = 0;
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
    document.body.classList.toggle("noheader", view === "notes" || view === "compare" || view === "todo" || view === "feedback");
    renderBar();
    renderPlotBar();
    var hideBars = view === "changes" || view === "tools" || view === "compare" || view === "notes" || view === "todo" || view === "feedback" || view === "answers";
    $("markbar").hidden = hideBars;
    if (hideBars) $("plotbar").hidden = true;
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
      if (e.target.closest("#barfold")) {
        var cmp = !document.body.classList.contains("compact");
        document.body.classList.toggle("compact", cmp); lsSet("compact", cmp ? "1" : "");
        renderBar(); return;
      }
      var gg = e.target.closest("[data-group]");
      if (gg) {
        var grp = GROUPS.filter(function (x) { return x.g === gg.dataset.group; })[0];
        var last = lastSub[grp.g];
        var pick = grp.subs.filter(function (sv) { return sv[0] === last; }).length ? last : grp.subs[0][0];
        if ((view === "notes" || view === "compare") && pick !== "compare") { markOff(); cardEdit = false; cardMode = false; }
        view = pick; lsSet("view", view); editing = -1; render();
        if (view === "changes") { refreshChanges().then(function () { if (!chg || !chg.last) W.sendRequest && W.sendRequest("info", ""); }); }
        if (view === "notes" || view === "compare") { refreshNotes(); }
        return;
      }
      var v = e.target.closest("[data-view]");
      if (v) {
        if ((view === "notes" || view === "compare") && v.dataset.view !== "compare") { markOff(); cardEdit = false; cardMode = false; }
        view = v.dataset.view; lastSub[groupOf(view)] = view; lsSet("view", view); editing = -1; render();
        if (view === "changes") { refreshChanges().then(function () { if (!chg || !chg.last) W.sendRequest && W.sendRequest("info", ""); }); }
        if (view === "notes" || view === "compare") { refreshNotes(); }
        return;
      }
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

    var mb = $("markbar");
    mb.addEventListener("click", function (e) {
      var f = e.target.closest("[data-mfilter]");
      if (f) { mfilter = mfilter === f.dataset.mfilter ? null : f.dataset.mfilter; editing = -1; render(); return; }
      if (e.target.closest("#clearm")) { e.preventDefault(); mfilter = null; render(); return; }
      if (e.target.closest("#mmanage")) { mmanage = !mmanage; renderMarkBar(); return; }
      var sw = e.target.closest(".mswatch");
      if (sw) {
        var m = markById(sw.closest(".mrow").dataset.mid);
        m.c = COLORS[(COLORS.indexOf(m.c) + 1) % COLORS.length];
        persist(); render(); return;
      }
      var del = e.target.closest(".mmdel");
      if (del) {
        var id = del.closest(".mrow").dataset.mid, mm = markById(id);
        if (!confirm('למחוק את הצבע "' + mm.n + '"? הוא יוסר מכל הסצנות.')) return;
        marks = marks.filter(function (x) { return x.id !== id; });
        scenes.forEach(function (s) { if (s.m === id) s.m = ""; });
        if (mfilter === id) mfilter = null;
        persist(); render();
      }
    });
    mb.addEventListener("change", function (e) {
      if (e.target.classList.contains("mmname") && e.target.value.trim()) {
        markById(e.target.closest(".mrow").dataset.mid).n = e.target.value.trim(); persist(); render();
      }
    });
    mb.addEventListener("keydown", function (e) {
      if (e.target.classList.contains("mnew") && e.key === "Enter" && e.target.value.trim()) {
        marks.push({ id: "m" + Date.now().toString(36), n: e.target.value.trim(), c: COLORS[marks.length % COLORS.length] });
        persist(); render();
        var nn = document.querySelector(".mnew"); if (nn) nn.focus();
      }
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
    list.addEventListener("input", function (e) { if (e.target.id === "todoin") todoDraft = e.target.value; });
    list.addEventListener("keydown", function (e) { if (e.target.id === "todoin" && e.key === "Enter") { todoDraft = e.target.value; todoAdd(); } });
    list.addEventListener("dragstart", function (e) {
      var r = e.target.closest && e.target.closest(".strow");
      if (r) { dragI = +r.dataset.i; try { e.dataTransfer.setData("text/plain", String(dragI)); } catch (x) {} }
  });
  list.addEventListener("focusout", function (e) {
      if (!e.target.classList || !e.target.classList.contains("srta")) return;
      var i = +e.target.dataset.i;
      if (scenes[i] && scenes[i].s !== e.target.value) { scenes[i].s = e.target.value.replace(/\n/g, " ").trim(); persist(); }
      sumEdit = -1; renderList();
    });
  list.addEventListener("change", function (e) {
      if (e.target.id === "fltch" || e.target.id === "fltst") {
        flt[e.target.id === "fltch" ? "ch" : "st"] = e.target.value;
        renderList(); return;
      }
      var tc = e.target.classList || { contains: function () { return false; } };
      if (tc.contains("ptta")) {
        var pi = +e.target.dataset.i, pv = e.target.value.trim();
        if (todos && todos[pi] && pv && pv !== todos[pi].text) { todos[pi].text = pv; saveTodos(); }
        return;
      }
      if (tc.contains("todoch")) { todos[+e.target.dataset.i].ch = e.target.value.split(",").map(function (x) { return x.trim(); }).filter(Boolean).join(", "); saveTodos(); return; }
      if (tc.contains("todopt")) { todos[+e.target.dataset.i].pt = e.target.value; saveTodos(); return; }
      if (tc.contains("fbch")) { fbItems[+e.target.dataset.i].ch = e.target.value.split(",").map(function (x) { return x.trim(); }).filter(Boolean).join(", "); saveFeedback(); return; }
      if (tc.contains("fbpt")) { fbItems[+e.target.dataset.i].pt = e.target.value; saveFeedback(); return; }
      if (!tc.contains("fbta")) return;
      var fi2 = +e.target.dataset.i, nv = e.target.value.trim();
      if (fbItems && fbItems[fi2] && nv && nv !== fbItems[fi2].text.replace(/^\s*\[[^\]]+\]\s*/, "")) { fbItems[fi2].cat = fbCatName(fbItems[fi2]); fbItems[fi2].text = nv; saveFeedback(); }
  });
  list.addEventListener("keydown", function (e) {
      if (e.target.classList && e.target.classList.contains("fbta") && e.key === "Enter" && !e.shiftKey) { e.preventDefault(); e.target.blur(); return; }
      if (e.target.classList && e.target.classList.contains("srta") && e.key === "Enter" && !e.shiftKey) { e.preventDefault(); e.target.blur(); }
  });
  list.addEventListener("dragover", function (e) { if (e.target.closest && e.target.closest(".part")) e.preventDefault(); });
  list.addEventListener("drop", function (e) {
      var st = e.target.closest && e.target.closest(".part");
      if (!st || dragI < 0) return;
      e.preventDefault();
      var stg = st.dataset.stage;
      if (scenes[dragI]) { scenes[dragI].st = stg; persist(); }
      dragI = -1; renderList();
  });
  list.addEventListener("click", function (e) {
      if (e.target.closest("#partsall")) {
        var allNow = STAGES.every(function (x) { return stageOpen[x]; });
        STAGES.forEach(function (x) { stageOpen[x] = !allNow; });
        renderList(); return;
      }
      var ph = e.target.closest(".parthead");
      if (ph) { var pg = ph.parentNode.dataset.stage; stageOpen[pg] = !stageOpen[pg]; renderList(); return; }
      var sed = e.target.closest(".srsum");
      if (sed) { sumEdit = +sed.dataset.i; renderList(); var ta = document.querySelector('.srta[data-i="' + sumEdit + '"]'); if (ta) ta.focus(); return; }
      var sr = e.target.closest(".strow");
      if (sr && !e.target.closest("button") && !e.target.closest("textarea")) { W.goToScene(+sr.dataset.i).catch(showErr); return; }
      if (e.target.closest("#sumcopy")) {
        var lines = scenes.map(function (sc, i) { return "סצנה " + (i + 1) + ": " + sc.h; }).join("\n");
        copyAndFlash(e.target.closest("#sumcopy"), "כתוב לכל סצנה בתסריט משפט אחד בזמן הווה, על מה קורה בה. אל תשנה את המסמך. החזר רק רשימה, שורה לכל סצנה, בפורמט: \"סצנה N: המשפט\"\n\nהסצנות:\n" + lines);
        return;
      }
      if (e.target.closest("#sumapply")) {
        var raw = ($("sumpaste").value || ""), done = 0, skipped = 0, miss = 0;
        raw.split("\n").forEach(function (ln) {
          var m = ln.match(/^\s*(?:סצנה\s*)?(\d+)\s*[:.)\-–]\s*(.+)$/);
          if (!m) return;
          var k = +m[1] - 1, txt = m[2].trim();
          if (!scenes[k]) { miss++; return; }
          if (scenes[k].s && scenes[k].s.trim()) { skipped++; return; }
          scenes[k].s = txt; done++;
        });
        if (done) persist();
        $("sumres").textContent = "הוכנסו " + done + (skipped ? ", דולגו " + skipped + " שכבר יש להן תקציר" : "") + (miss ? ", " + miss + " לא נמצאו" : "");
        renderList();
        return;
      }
      var ask = e.target.closest(".askclaude");
      if (ask) {
        var sc = scenes[+ask.dataset.i];
        if (sc) {
          var plotNames = plots.filter(function (p) { return sc.p.indexOf(p.id) >= 0; }).map(function (p) { return p.n; }).join(", ");
          copyAndFlash(ask, "שאלה על סצנה בתסריט (אל תשנה את המסמך):\nסצנה: " + sc.h + "\nמה קורה: " + (sc.s || "(עדיין אין תקציר)") + (plotNames ? "\nעלילות: " + plotNames : "") +
            "\n\nתן לי 3 אפשרויות לחזק או לשנות את הסצנה הזו, כל אחת במשפט אחד עם הסבר קצר ומה היא משנה. אל תכתוב הצעות בפורמט של תיקון בתסריט, רק את האפשרויות.");
        }
        return;
      }
      var askF = e.target.closest(".askclaude-fb");
      if (askF) {
        var ft = fbItems[+askF.dataset.i];
        copyAndFlash(askF, "פריט מסיכום השיחה עם המרצה (אל תשנה את המסמך):\n" + (ft.text || "") + (ft.sn ? "\nסצנה " + ft.sn : "") + (ft.ch ? "\nדמות: " + ft.ch : "") + (fltStageOf(ft) || ft.pt ? "\nחלק: " + (stNorm(ft.pt) || fltStageOf(ft)) : "") +
          "\n\nעזור לי לחשוב על זה. אם צריך תיקון בתסריט, כתוב אותו בפורמט ההצעה הרגיל. אל תשנה את המסמך.");
        return;
      }
      var askT = e.target.closest(".askclaude-todo");
      if (askT) {
        var tt2 = todos[+askT.dataset.i];
        copyAndFlash(askT, "מחשבה מהרשימה שלי על התסריט:\n" + tt2.text + (tt2.scene ? "\nסצנה: " + tt2.scene : "") +
          "\n\nעזור לי לחשוב על זה. אם צריך תיקון בתסריט, כתוב אותו בפורמט ההצעה הרגיל. אל תשנה את המסמך.");
        return;
      }
      if (e.target.closest("#fltcolbtn")) { colOpen = !colOpen; renderList(); return; }
      if (e.target.closest("#fltparts")) { partsOpen = !partsOpen; renderList(); return; }
      var flg = e.target.closest(".fbleg");
      if (flg) { var lc = flg.dataset.c; flt.c = flt.c === lc ? "" : lc; renderList(); return; }
      var fcolB = e.target.closest(".fltcol");
      if (fcolB) { flt.c = fcolB.dataset.c || ""; colOpen = false; renderList(); return; }
      if (view === "answers") {
        if (e.target.closest("#ansAdd")) {
          var at = ($("ansTitle").value || "").trim(), ab = ($("ansBody").value || "").trim();
          if (!ab) return;
          ansItems.push({ title: at || ab.slice(0, 40), body: ab, date: new Date().toISOString().slice(0, 10) });
          ansTitle = ""; ansBody = ""; ansOpen = -1; saveAnswers(); renderList(); return;
        }
        var ad = e.target.closest(".ansdel");
        if (ad) { ansItems.splice(+ad.dataset.i, 1); saveAnswers(); renderList(); return; }
        var ao = e.target.closest(".ansopen");
        if (ao) { var aoi = +ao.dataset.i; ansOpen = ansOpen === aoi ? -1 : aoi; renderList(); return; }
        return;
      }
      if (view === "feedback" || (view === "todo" && e.target.closest("#colR"))) {
        var fl = $("colR") || list;
        if (e.target.closest("#fbopen")) {
          W.sendRequest("copyfb", "").then(function () { $("fbopened").textContent = "הסיכום הועתק. הדבק אותו בצ'אט של Claude עם הבקשה"; }).catch(showErr);
          return;
        }
        if (e.target.closest("#fbpasteok")) {
          var fn = 0;
          ($("fbpaste").value || "").split("\n").forEach(function (ln) {
            var t = ln.trim(); if (!t) return;
            t = t.replace(/^[-•*]\s*/, "");
            var sm = t.match(/סצנ(?:ה|ות)\s*(\d+)/);
            var sn = sm ? +sm[1] : 0;
            var chm = t.match(/דמות:\s*([^|]+)/), ptm = t.match(/חלק:\s*([^|]+)/);
            var ch = chm ? chm[1].trim() : "", pt = ptm ? ptm[1].trim() : "";
            var clean = sm ? t.replace(/^סצנ(?:ה|ות)[^:\[]*:\s*/, "").trim() : t;
            clean = clean.replace(/\|?\s*דמות:[^|]*/g, "").replace(/\|?\s*חלק:[^|]*/g, "").trim();
            var catm = (clean || t).match(/\[([^\]]+)\]/);
            var cat = catm ? catm[1].trim() : "";
            var body = (clean || t).replace(/\[[^\]]+\]\s*/, "").trim();
            fbItems.push({ text: body || clean || t, done: false, sn: sn, cat: cat, ch: ch, pt: STAGES.indexOf(stNorm(pt)) >= 0 ? stNorm(pt) : "" }); fn++;
          });
          fbPaste = "";
          if (fn) saveFeedback();
          $("fbres").textContent = fn ? "נוספו " + fn + " פריטים" : "אין מה להוסיף";
          if (fn) renderFeedback(fl);
          return;
        }
        var fc = e.target.closest(".fbcheck");
        if (fc) { var fi = +fc.dataset.i; fbItems[fi].done = !fbItems[fi].done; saveFeedback(); renderFeedback(fl); return; }
        var fcap = e.target.closest(".fbcap");
        if (fcap) { var fci = +fcap.dataset.i; fbMeta = fbMeta === fci ? -1 : fci; renderFeedback(fl); return; }
        var fcur = e.target.closest(".fbcur");
        if (fcur) { var ci2 = +fcur.dataset.i; fbPick = fbPick === ci2 ? -1 : ci2; renderFeedback(fl); return; }
        var fdot = e.target.closest(".fbdot");
        var fcat = e.target.closest(".fbcatbtn");
        if (fcat) { var dci = +fcat.dataset.i; fbItems[dci].cat = fcat.dataset.cat || ""; fbPick = -1; saveFeedback(); renderFeedback(fl); return; }
        if (fdot) { var di = +fdot.dataset.i; fbItems[di].pt = fdot.dataset.s || ""; fbPick = -1; saveFeedback(); renderFeedback(fl); return; }
        var fsc = e.target.closest(".fbscene");
        if (fsc) { W.goToScene(+fsc.dataset.sn - 1).catch(showErr); return; }
        var fdl = e.target.closest(".fbdel");
        if (fdl) { fbItems.splice(+fdl.dataset.i, 1); saveFeedback(); renderFeedback(fl); return; }
        return;
      }
      if (view === "todo" && !e.target.closest("#colR")) {
        if (e.target.closest("#todoadd")) { todoDraft = $("todoin").value; todoAdd(); return; }
        var tcur = e.target.closest(".todocur");
        // a click on the colored strip at the item's edge also opens the color choice
        var tstrip = e.target.classList && e.target.classList.contains("todoitem") ? e.target : null;
        if (tstrip) {
          var rr = tstrip.getBoundingClientRect(), rtl = getComputedStyle(tstrip).direction === "rtl";
          if (rtl ? e.clientX >= rr.right - 12 : e.clientX <= rr.left + 12) { var tsi = +tstrip.dataset.i; todoPick = todoPick === tsi ? -1 : tsi; renderTodo(list); return; }
        }
        if (tcur) { var tci = +tcur.dataset.i; todoPick = todoPick === tci ? -1 : tci; renderTodo(list); return; }
        var tcol = e.target.closest(".todocol");
        if (tcol) { var tcn = +tcol.dataset.i; todos[tcn].pt = tcol.dataset.s || ""; todoPick = -1; saveTodos(); renderTodo(list); return; }
        var tcap = e.target.closest(".todocap");
        if (tcap) { var tcp = +tcap.dataset.i; todoMeta = todoMeta === tcp ? -1 : tcp; renderTodo(list); return; }
        var tk = e.target.closest(".todocheck");
        if (tk) { var ti = +tk.dataset.i; todos[ti].done = !todos[ti].done; saveTodos(); renderList(); return; }
        var td = e.target.closest(".tododel");
        if (td) { todos.splice(+td.dataset.i, 1); saveTodos(); renderList(); return; }
        var tl = e.target.closest(".todolinkbtn");
        if (tl) { todoLinking = todoLinking === tl.dataset.i ? "" : tl.dataset.i; renderList(); return; }
        var tok = e.target.closest("#todolinkok");
        if (tok) {
          var ti2 = +tok.dataset.i;
          W.currentSceneTitle().then(function (sc) {
            if (sc) { todos[ti2].scene = sc; saveTodos(); }
            todoLinking = ""; renderList();
          }).catch(showErr);
          return;
        }
        var tsc = e.target.closest(".todoscene");
        if (tsc) { W.gotoHeading(todos[+tsc.dataset.i].scene).catch(showErr); return; }
        return;
      }
      if ((view === "notes" || view === "compare") && e.target.closest("#notesfold")) { notesFolded = !notesFolded; lsSet("notesFold", notesFolded ? "1" : ""); renderList(); return; }
      if ((view === "notes" || view === "compare") && !e.target.closest(".toolbtn, #smartcopy")) {
        if (e.target.id === "notedone") { showDone = e.target.checked; renderList(); return; }
        if (e.target.closest("#delcmpno")) { delCmpPending = false; clearTimeout(delCmpT); renderList(); return; }
        if (e.target.closest("#delcmp")) {
          // first click asks, second click deletes; the question goes away by itself after a few seconds
          clearTimeout(delCmpT);
          if (!delCmpPending) { delCmpPending = true; renderList(); delCmpT = setTimeout(function () { delCmpPending = false; renderList(); }, 6000); return; }
          delCmpPending = false;
          markOff().then(function () { cardMode = false; cardEdit = false; return W.deleteCompareNotes(); })
            .then(function (k) { cardMsg = ""; return refreshNotes().then(function () { renderList(); var st = document.querySelector(".chgctl .chgstat"); if (st) st.textContent = "נמחקו " + k.done + " הערות השוואה" + (k.failed ? ", " + k.failed + " לא נמחקו (אפשר למחוק אותן ב-Word: קליק ימני על ההערה ← מחק)" : "") + "."; }); })
            .catch(showErr);
          return;
        }
        if (e.target.closest("#instrsend")) { walkOpen = walkOpen === "instr" ? false : "instr"; renderList(); return; }
        if (e.target.closest("#instropen")) { instrOpen = true; cardMsg = ""; renderList(); var it = $("instrtext"); if (it) it.focus(); return; }
        if (e.target.closest("#instrcancel")) { instrOpen = false; renderList(); return; }
        if (e.target.closest("#keepscr")) {
          var kid = noteSel, snK = snapOrder();
          var kn = (notes || []).filter(function (x) { return x.id === kid; })[0];
          markOff().then(function () { return W.replyNote(kid, "טופל – התסריט נשאר. לעדכן את הסינופסיס: " + (kn ? kn.text.replace(/^השוואה:\s*/, "") : "")); }).then(function () { return W.resolveNote(kid, true); })
            .then(refreshNotes).then(function () { advanceAfter(snK); }).catch(showErr);
          return;
        }
        if (e.target.closest("#dosug")) {
          var did = noteSel, snD = snapOrder();
          cardMsg = "שומר…"; renderList();
          W.replyNote(did, "הוראה: תתאים את התסריט למה שכתוב בסינופסיס בנקודה הזו. כתוב את זה כפעולות שאפשר להכניס (במקום/יבוא, מחק, מחק פסקה, אחרי + שורות עם סגנון), או כ\"החלף סצנה:\" אם צריך שכתוב של כל הסצנה.")
            .then(function () { cardMsg = ""; return refreshNotes().then(function () { advanceAfter(snD); }); })
            .catch(function (er) { noteFail(er, snD); });
          return;
        }
        if (e.target.closest("#instrsave")) {
          var txt = ($("instrtext").value || "").trim(), iid = noteSel, snI = snapOrder();
          if (!txt) { cardMsg = "כתוב הוראה לפני השמירה."; instrOpen = false; renderList(); return; }
          instrOpen = false; cardMsg = "שומר…"; renderList();
          W.replyNote(iid, "הוראה: " + txt.replace(/\s*\n\s*/g, " ")).then(function () {
            cardMsg = "";
            return refreshNotes().then(function () { advanceAfter(snI); });
          }).catch(function (er) { noteFail(er, snI); });
          return;
        }
        if (e.target.closest("#doneok")) {
          var dn = doneNote; if (!dn) return;
          doneNote = null; cardMsg = "";
          markOff().then(function () { return W.resolveNote(dn.id, true); }).then(refreshNotes).then(function () { advanceAfter(dn.snap); }).catch(showErr);
          return;
        }
        if (e.target.closest("#donefix")) {
          var df = doneNote; if (!df) return;
          cardMsg = "מחזיר את הטקסט…"; renderList();
          W.undoSuggestion(df.undo || []).then(function () { cardMsg = ""; }, function (er) { cardMsg = (er && er.message ? er.message : String(er)); })
            .then(function () { return markOff(); }).then(function () {
              doneNote = null; noteSel = df.id; instrOpen = true; renderList();
              var it2 = $("instrtext"); if (it2) it2.focus();
            });
          return;
        }
        var wj = e.target.closest(".warnjump");
        if (wj) {
          var q = wj.dataset.quote || "", scn = +wj.dataset.scene || 0, co2 = openCmp();
          var hit = q ? co2.filter(function (n) { var w = (n.where || "").slice(0, 30); return w && (q.indexOf(w) >= 0 || (n.where || "").indexOf(q.slice(0, 30)) >= 0); })[0] : null;
          if (!hit && scn) hit = co2.filter(function (n) { return n.scene === scn && n.id !== noteSel; })[0];
          if (hit) { selectNote(hit.id, true); return; }
          if (q) { W.markFind(q).catch(showErr); return; }
          if (scn) { W.goToScene(scn - 1).catch(showErr); cardMsg = "אין הערה פתוחה בסצנה " + scn + ". קפצתי לכותרת הסצנה."; renderList(); }
          return;
        }
        var fd = e.target.closest(".findable");
        if (fd) {
          W.markFind(fd.dataset.find).then(function (found) { if (!found) { cardMsg = "הקטע לא נמצא בתסריט."; renderList(); } }).catch(showErr);
          return;
        }
        if (e.target.closest("#cardstart")) {
          cardMode = true; walkOpen = false;
          var co0 = openCmp(), withS = co0.filter(function (n) { return !!sugRaw(n); });
          var startId = co0.map(function (n) { return n.id; }).indexOf(noteSel) >= 0 ? noteSel : (withS[0] || co0[0]).id;
          selectNote(startId, true); return;
        }
        if (e.target.closest("#cardrestart")) { roundDone = false; var co1 = openCmp(); if (co1.length) selectNote(co1[0].id, true); return; }
        if (e.target.closest("#cardlist")) { doneNote = null; roundDone = false; cardMode = false; cardEdit = false; markOff(); renderList(); return; }
        if (e.target.closest("#cmtview")) { cmtShow = !cmtShow; renderList(); W.sendRequest("comments", cmtShow ? "show" : "hide").catch(showErr); return; }
        if (e.target.closest("#walkshow")) { walkOpen = walkOpen === "walk" ? false : "walk"; renderList(); return; }
        if (e.target.closest("#prepshow")) { walkOpen = walkOpen === "prep" ? false : "prep"; renderList(); return; }
        if (e.target.closest("#walkcopy")) { copyFrom("walktext", "walkcopied"); return; }
        if (e.target.closest("#cmpnext")) { stepCmp(1); return; }
        if (e.target.closest("#cmpprev")) { stepCmp(-1); return; }
        if (e.target.closest("#cardskip")) { stepCmp(1); return; }
        if (e.target.closest("#cardedit")) { cardEdit = true; cardMsg = ""; renderList(); var ct = $("cardtext"); if (ct) ct.focus(); return; }
        if (e.target.closest("#cardcancel")) { cardEdit = false; renderList(); return; }
        if (e.target.closest("#cardsave")) {
          // the edited text is saved as a new "הצעה:" reply, so the card (and Claude) use it from now on; nothing goes into the script yet
          var sid = noteSel, sv = ($("cardtext") ? $("cardtext").value : "").trim();
          if (!sid || !sv) return;
          if (!/^הצעה:/.test(sv)) sv = "הצעה:\n" + sv;
          var sgS = parseSug(sv);
          if (!sgS || !sgS.ok) { cardMsg = "העריכה לא בפורמט שאפשר להכניס. תקן אותה, או לחץ \"בטל עריכה\"."; renderList(); return; }
          W.replyNote(sid, sv).then(function () { cardEdit = false; cardMsg = "העריכה נשמרה. אפשר לראות לפני ואחרי, ואז להכניס."; return refreshNotes(); })
            .then(function () { renderList(); }).catch(showErr);
          return;
        }
        if (e.target.closest("#cardok")) {
          var idk = noteSel, snO = snapOrder();
          markOff().then(function () { return W.resolveNote(idk, true); }).then(refreshNotes).then(function () { advanceAfter(snO); }).catch(showErr);
          return;
        }
        if (e.target.closest("#scenebefore")) {
          if (sceneBefore && sceneBefore.id === noteSel) { sceneBefore = null; renderList(); return; }
          var sbId = noteSel;
          W.sceneOfNote(sbId).then(function (r) { sceneBefore = { id: sbId, paras: r.paras }; }, function (er) { sceneBefore = { id: sbId, err: er && er.message ? er.message : String(er) }; })
            .then(function () { renderList(); });
          return;
        }
        if (e.target.closest("#cardins")) {
          var nid = noteSel, snN = snapOrder(), cn = (notes || []).filter(function (x) { return x.id === nid; })[0];
          if (!cn) return;
          var sg2 = parseSug(cardEdit && $("cardtext") ? $("cardtext").value : sugRaw(cn));
          if (!sg2 || !sg2.ok) { cardMsg = "ההצעה לא בפורמט שאפשר להכניס. ערוך אותה או דלג."; renderList(); return; }
          var chk = W.listChangeMarks ? W.listChangeMarks() : Promise.resolve(null);
          chk.then(function (m) {
            if (m && m.length) { cardMsg = "מוצגים סימוני שינויים מול דראפט. קודם לחץ \"הסתר\" בלשונית שינויים."; renderList(); return; }
            cardMsg = "מכניס…"; renderList();
            return markOff().then(function () { return sg2.scene ? W.replaceScene(nid, sg2.ops[0]) : W.applySuggestion(nid, sg2); }).then(function (undo) {
              sceneBefore = null;
              cardEdit = false; cardMsg = "";
              doneNote = { id: nid, text: cn.text, where: cn.where, snap: snN, undo: undo };
              return refreshNotes().then(function () { renderList(); });
            }).catch(function (er) { cardMsg = "לא הוכנס: " + (er && er.message ? er.message : er); renderList(); });
          });
          return;
        }
        var ok = e.target.closest(".noteok");
        if (ok) {
          var nn = (notes || []).filter(function (x) { return x.id === ok.dataset.id; })[0];
          if (nn) { nn.resolved = !nn.resolved; renderBar(); renderList(); W.resolveNote(nn.id, nn.resolved).then(refreshNotes).catch(showErr); }
          return;
        }
        var ng = e.target.closest(".notego");
        if (ng) {
          var gn = (notes || []).filter(function (x) { return x.id === ng.dataset.id; })[0];
          var isC = !!(gn && isCmp(gn) && !gn.resolved);
          if (isC) cardMode = true;
          selectNote(ng.dataset.id, isC);
        }
        return;
      }
      if (view === "tools" || view === "compare") {
        if (e.target.closest("#smartcopy")) { copyFrom("smarttext", "smartcopied"); return; }
        if (e.target.closest("#smartrebuild")) { loadBundle(0); return; }
        if (e.target.closest("#beatscopy")) { copyFrom("beatstext", "beatscopied"); return; }
        if (e.target.closest("#beatsapply")) { applyBeats(); return; }
        if (e.target.closest("#beatsclear")) {
          W.beatClear().then(function () { $("beatsstat").textContent = "הסימונים נוקו."; }).catch(showErr);
          return;
        }
        if (e.target.closest("#beatsall") || e.target.closest("#beatsnone")) {
          var allOn = !!e.target.closest("#beatsall");
          scenes.forEach(function (s, i) { beatSel[i] = allOn; });
          document.querySelectorAll(".beatcb").forEach(function (cb) { cb.checked = allOn; });
          $("beatstext").value = buildBeatsPrompt();
          return;
        }
        if (e.target.classList && e.target.classList.contains("beatcb")) {
          beatSel[+e.target.dataset.i] = e.target.checked;
          $("beatstext").value = buildBeatsPrompt();
          return;
        }
        var tb = e.target.closest(".toolbtn");
        if (tb && tb.dataset.cmd === "beats") { beatsOpen = !beatsOpen; renderList(); return; }
        if (tb) {
          if (tb.dataset.cmd === "smart_s2t") { smartPrompt = "s2t"; setTimeout(function () { loadBundle(10); }, 2500); }
          else if (tb.dataset.cmd === "smart_t2s") smartPrompt = "t2s";
          else if (tb.dataset.cmd === "smartclear") smartPrompt = "";
          toolPending = tb.dataset.cmd; renderList();
          W.sendRequest(tb.dataset.cmd, "").catch(showErr);
          setTimeout(function () { toolPending = ""; if (view === "tools" || view === "compare") renderList(); }, 2500);
        }
        return;
      }
      if (view === "changes") {
        if (e.target.closest("#chgshow")) { var n = +$("chgdraft").value; if (n) chgRequest("compare", String(n)); return; }
        if (e.target.closest("#chghide")) { chgRequest("hide"); return; }
        var goK = function (k) {
          if (!marks2 || !marks2.length) return false;
          markIdx = ((k - 1 + marks2.length) % marks2.length) + 1;
          W.gotoMark(marks2[markIdx - 1].name).catch(showErr);
          renderList();
          var el = document.querySelector('.chgitem[data-k="' + markIdx + '"]');
          if (el && el.scrollIntoViewIfNeeded) el.scrollIntoViewIfNeeded(false);
          return true;
        };
        if (e.target.closest("#chgnext")) { if (!goK(markIdx + 1)) W.sendRequest("next", ""); return; }
        if (e.target.closest("#chgprev")) { if (!goK(markIdx ? markIdx - 1 : (marks2 ? marks2.length : 0))) W.sendRequest("prev", ""); return; }
        var ci = e.target.closest(".chgitem");
        if (ci) {
          if (ci.dataset.name) { goK(+ci.dataset.k); }
          else W.sendRequest("goto", ci.dataset.pos);
        }
        return;
      }
      var go = e.target.closest(".go");
      if (go) { var i = +go.dataset.i; markCurrent(i); W.goToScene(i).catch(showErr); return; }
      var dt = e.target.closest(".dot");
      if (dt) { var di = +dt.dataset.i; picking = picking === di ? -1 : di; renderList(); return; }
      var pk = e.target.closest(".pick");
      if (pk && picking >= 0) { scenes[picking].m = pk.dataset.m; picking = -1; persist(); renderPlotBar(); renderList(); return; }
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
      if (e.target.id === "chgdraft") { chgSel = +e.target.value; return; }
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

  // ---------- auto brackets for parentheticals
  // When a paragraph turns into "הערת בימוי" (shortcut or button), wrap it in ( ).
  var wrapping = false;
  function pollParen() {
    if (wrapping || busy || !W.curPara) return;
    W.curPara().then(function (c) {
      var prev = lastPara;
      lastPara = c;
      updateLocBox(c);
      if (!prev || c.style !== PAREN || prev.style === PAREN) return;
      var t = c.text.trim();
      var inPlace = prev.text === c.text;           // style changed on the same paragraph
      if (t === "" || (inPlace && (t.charAt(0) !== "(" || t.charAt(t.length - 1) !== ")"))) {
        wrapping = true;
        return W.wrapParen().then(function () {
          return W.curPara().then(function (c2) { lastPara = c2; });
        }).catch(function () {}).then(function () { wrapping = false; });
      }
    }).catch(function () {});
  }

  // ---------- location bank
  var TYPES = ["פנים", "חוץ", "פנים/חוץ"];
  var TIMES = ["יום", "לילה", "בוקר", "צהריים", "אחר הצהריים", "ערב", "שחר"];
  var HEAD_RE = /^(פנים\/חוץ|פנים|חוץ)\.?\s*(.*)$/;

  function parseHead(t) {
    t = (t || "").trim();
    var m = HEAD_RE.exec(t);
    if (!m) return { type: "", loc: t, time: "", hasDash: false };
    var rest = m[2], loc = rest, time = "", dash = false;
    var k = rest.search(/\s[–-]\s?|[–]/);
    if (k >= 0) {
      dash = true;
      loc = rest.slice(0, k);
      time = rest.slice(k).replace(/^\s*[–-]\s*/, "");
    }
    return { type: m[1], loc: loc.trim(), time: time.trim(), hasDash: dash };
  }

  function locBank() {
    var count = {};
    scenes.forEach(function (sc) {
      var p = parseHead(sc.h);
      var l = p.loc.replace(/\s*\(.*\)\s*$/, "").trim();
      if (p.type && l) count[l] = (count[l] || 0) + 1;
    });
    customLocs.forEach(function (l) { if (!(l in count)) count[l] = 0; });
    return Object.keys(count).sort(function (a, b) { return count[b] - count[a] || a.localeCompare(b, "he"); })
      .map(function (l) { return { n: l, c: count[l], custom: customLocs.indexOf(l) >= 0 }; });
  }

  var locSig = "";
  function updateLocBox(c) {
    var box = $("locbox");
    if (!box) return;
    if (!c || c.style !== SCENE) {
      if (!locOpen) { box.hidden = true; box.innerHTML = ""; locSig = ""; }
      return;
    }
    var sig = c.text + "|" + locOpen + "|" + customLocs.length + "|" + scenes.length;
    if (sig === locSig && !box.hidden) return;
    locSig = sig;
    box.hidden = false;
    var p = parseHead(c.text);
    var bank = locBank();
    var html = '<div class="lochead"><b>כותרת סצנה</b><span class="sp"></span>' +
      '<button class="icon" id="locmanage" title="בנק לוקיישנים">' + (locOpen ? "▴" : "☰") + ' ' + bank.length + '</button></div>';
    if (!p.type) {
      html += '<div class="lochint">מתחילים ב:</div><div class="locbtns">' + TYPES.map(function (t) {
        return '<button class="lb" data-type="' + esc(t) + '">' + esc(t) + '.</button>';
      }).join("") + '</div>';
    } else if (!p.hasDash) {
      var f = p.loc;
      var list = bank.filter(function (l) { return !f || l.n.indexOf(f) >= 0; }).slice(0, 12);
      html += '<div class="lochint">' + (f ? 'לוקיישנים שמתאימים ל"' + esc(f) + '":' : 'בחר לוקיישן או המשך להקליד:') + '</div>';
      html += '<div class="loclist">' + (list.length ? list.map(function (l) {
        return '<button class="li" data-loc="' + esc(l.n) + '">' + esc(l.n) + (l.c ? ' <small>' + l.c + '</small>' : '') + '</button>';
      }).join("") : '<span class="lochint">אין לוקיישן כזה בבנק עדיין.</span>') + '</div>';
      if (f && !bank.some(function (l) { return l.n === f; })) {
        html += '<button class="lb add" data-addloc="' + esc(f) + '">+ הוסף "' + esc(f) + '" לבנק</button>';
      }
    } else {
      html += '<div class="lochint">זמן:</div><div class="locbtns">' + TIMES.map(function (t) {
        return '<button class="lb' + (p.time === t ? " on" : "") + '" data-time="' + esc(t) + '">' + esc(t) + '</button>';
      }).join("") + '</div>';
    }
    if (locOpen) {
      html += '<div class="locman">' + bank.map(function (l) {
        return '<div class="lmrow"><span class="lmloc" data-loc="' + esc(l.n) + '" title="הכנס לכותרת">' + esc(l.n) + '</span><small>' + (l.c ? l.c + " סצנות" : "ידני") + '</small>' +
          (l.custom ? '<button class="icon" data-delloc="' + esc(l.n) + '" title="הסר מהבנק">✕</button>' : '') + '</div>';
      }).join("") + '<input class="newloc" placeholder="+ לוקיישן חדש לבנק (Enter)"></div>';
    }
    box.innerHTML = html;
  }

  function setHead(type, loc, time) {
    var t = type + ". " + (loc || "");
    if (time !== undefined) t += " – " + time;
    else if (loc) t += " – ";
    return W.setParaText(t).then(function () { locSig = ""; return W.curPara().then(updateLocBox); }).catch(showErr);
  }

  function wireLoc() {
    var box = $("locbox");
    if (!box) return;
    box.addEventListener("click", function (e) {
      var c = lastPara || { text: "" };
      var p = parseHead(c.text);
      var b;
      if ((b = e.target.closest("[data-type]"))) {
        // a location picked before the type stays, and the time buttons come next
        if (p.loc) { setHead(b.dataset.type, p.loc); return; }
        W.setParaText(b.dataset.type + ". ").then(function () { locSig = ""; return W.curPara().then(updateLocBox); }).catch(showErr); return;
      }
      if ((b = e.target.closest("[data-loc]"))) {
        // the bank list can be open while the cursor is elsewhere: never overwrite a paragraph that is not a scene heading
        if (!lastPara || lastPara.style !== SCENE) { showErr("שים את הסמן על כותרת סצנה, ואז לחץ על הלוקיישן."); return; }
        if (b.classList.contains("lmloc")) locOpen = false;
        if (!p.type) { W.setParaText(b.dataset.loc).then(function () { locSig = ""; return W.curPara().then(updateLocBox); }).catch(showErr); return; }
        setHead(p.type, b.dataset.loc, p.hasDash && p.time ? p.time : undefined); return;
      }
      if ((b = e.target.closest("[data-time]"))) { setHead(p.type, p.loc, b.dataset.time); return; }
      if ((b = e.target.closest("[data-addloc]"))) {
        if (customLocs.indexOf(b.dataset.addloc) < 0) customLocs.push(b.dataset.addloc);
        persist(); setHead(p.type, b.dataset.addloc); return;
      }
      if ((b = e.target.closest("[data-delloc]"))) {
        customLocs = customLocs.filter(function (x) { return x !== b.dataset.delloc; });
        persist(); locSig = ""; updateLocBox(lastPara); return;
      }
      if (e.target.closest("#locmanage")) { locOpen = !locOpen; locSig = ""; updateLocBox(lastPara && lastPara.style === SCENE ? lastPara : { style: SCENE, text: "" }); }
    });
    box.addEventListener("keydown", function (e) {
      if (e.target.classList.contains("newloc") && e.key === "Enter" && e.target.value.trim()) {
        var v = e.target.value.trim();
        if (customLocs.indexOf(v) < 0) customLocs.push(v);
        persist(); locSig = ""; updateLocBox(lastPara && lastPara.style === SCENE ? lastPara : { style: SCENE, text: "" });
        var n = document.querySelector(".newloc"); if (n) n.focus();
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
    wireLoc();
    W.onSelectionChange(onSel);
    setInterval(function () { if (!document.hidden) pollParen(); }, 700);
    refreshScenes().then(refreshCursor);
    setInterval(function () {
      if (document.hidden || editing >= 0) return;
      refreshNotes();
      if (view === "changes") { refreshChanges(); return; }
      if (view === "notes" || view === "compare" || view === "feedback" || view === "todo") return;
      refreshScenes();
    }, 3000);
    if (view === "changes") refreshChanges();
    refreshNotes();
  });
})();
