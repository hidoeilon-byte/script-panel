// Real Word connection (Office.js). Same interface as the dev mock.
(function () {
  var SCENE = "כותרת סצנה";
  var KEY = "scenePanel.data";
  var OLD_KEY = "scenePanel.synopses";

  function run(fn) { return Word.run(fn); }

  function sceneRanges(ps) {
    var idx = [];
    ps.items.forEach(function (p, i) { if (p.style === SCENE) idx.push(i); });
    return {
      idx: idx,
      range: function (k) {
        var start = ps.items[idx[k]];
        var end = k + 1 < idx.length ? ps.items[idx[k + 1] - 1] : ps.items[ps.items.length - 1];
        return start.getRange("Whole").expandTo(end.getRange("Whole"));
      }
    };
  }

  window.WordBridge = {
    ready: function () {
      return new Promise(function (res) {
        Office.onReady(function () {
          try {
            var st = Office.context.document.settings;
            if (!st.get("Office.AutoShowTaskpaneWithDocument")) {
              st.set("Office.AutoShowTaskpaneWithDocument", true);
              st.saveAsync();
            }
          } catch (e) {}
          res();
        });
      });
    },

    paragraphs: function () {
      return run(function (ctx) {
        var ps = ctx.document.body.paragraphs;
        ps.load("items/text,items/style");
        return ctx.sync().then(function () {
          return ps.items.map(function (p) { return { text: p.text, style: p.style }; });
        });
      });
    },

    cursor: function () {
      return run(function (ctx) {
        var sel = ctx.document.getSelection();
        var cur = sel.paragraphs.getFirst();
        cur.load("style");
        var before = ctx.document.body.getRange("Start").expandTo(sel.getRange("Start"));
        var ps = before.paragraphs;
        ps.load("items/style");
        return ctx.sync().then(function () {
          var n = 0;
          ps.items.forEach(function (p) { if (p.style === SCENE) n++; });
          return { style: cur.style, sceneIndex: n - 1 };
        });
      });
    },

    applyStyle: function (name) {
      return run(function (ctx) {
        var ps = ctx.document.getSelection().paragraphs;
        ps.load("items");
        return ctx.sync().then(function () {
          ps.items.forEach(function (p) { p.style = name; });
          return ctx.sync();
        });
      });
    },

    goToScene: function (i) {
      return run(function (ctx) {
        var ps = ctx.document.body.paragraphs;
        ps.load("items/style");
        return ctx.sync().then(function () {
          var heads = ps.items.filter(function (p) { return p.style === SCENE; });
          if (heads[i]) { heads[i].select("Start"); }
          return ctx.sync();
        });
      });
    },

    // Move scene `from` so it sits before scene `to` (to === count -> end).
    moveScene: function (from, to) {
      return run(function (ctx) {
        var ps = ctx.document.body.paragraphs;
        ps.load("items/style,items/text");
        return ctx.sync().then(function () {
          var sr = sceneRanges(ps);
          var n = sr.idx.length;
          if (from < 0 || from >= n || to === from || to === from + 1) return null;
          var src = sr.range(from);
          var ox = src.getOoxml();
          return ctx.sync().then(function () {
            var inserted;
            if (to < n) {
              inserted = ps.items[sr.idx[to]].getRange("Whole").insertOoxml(ox.value, "Before");
            } else {
              inserted = ps.items[ps.items.length - 1].getRange("Whole").insertOoxml(ox.value, "After");
            }
            src.delete();
            var ips = inserted.paragraphs;
            ips.load("items/text,items/style");
            return ctx.sync().then(function () {
              // insertOoxml can leave one extra empty paragraph at the end of the block
              var last = ips.items[ips.items.length - 1];
              var srcLastText = ps.items[(from + 1 < n ? sr.idx[from + 1] : ps.items.length) - 1].text;
              if (last && ips.items.length > 1 && last.text === "" && srcLastText !== "") {
                last.delete();
              }
              ips.items[0].select("Start");
              return ctx.sync();
            });
          });
        });
      });
    },

    // {style, text} of the cursor paragraph (cheap, for polling)
    curPara: function () {
      return run(function (ctx) {
        var p = ctx.document.getSelection().paragraphs.getFirst();
        p.load("style,text");
        return ctx.sync().then(function () { return { style: p.style, text: p.text }; });
      });
    },

    // Replace the text of the cursor paragraph (keeps its style) and put the cursor at the end
    setParaText: function (t) {
      return run(function (ctx) {
        var p = ctx.document.getSelection().paragraphs.getFirst();
        var r = p.insertText(t, "Replace");
        r.getRange("End").select();
        return ctx.sync();
      });
    },

    // Wrap the cursor paragraph in round brackets; empty -> "()" with cursor inside
    wrapParen: function () {
      return run(function (ctx) {
        var p = ctx.document.getSelection().paragraphs.getFirst();
        p.load("text");
        return ctx.sync().then(function () {
          var t = p.text.trim();
          if (t === "") {
            p.insertText("()", "Start");
            return ctx.sync().then(function () {
              var r = p.search("(", { matchCase: true }).getFirst();
              r.getRange("End").select();
              return ctx.sync();
            });
          }
          if (t.charAt(0) !== "(") p.insertText("(", "Start");
          if (t.charAt(t.length - 1) !== ")") p.insertText(")", "End");
          return ctx.sync().then(function () {
            var r = p.search(")", { matchCase: true }).getLast();
            r.getRange("Start").select();
            return ctx.sync();
          });
        });
      });
    },

    // changes list written by the macro
    readChanges: function () {
      return new Promise(function (res) {
        try {
          Office.context.document.customXmlParts.getByNamespaceAsync("urn:scripttool:chg", function (r) {
            window.CHGDBG = "ns:" + r.status + " n=" + (r.value ? r.value.length : "-") + (r.error ? " " + r.error.message : "");
            if (r.status !== "succeeded" || !r.value || !r.value.length) return res(null);
            r.value[r.value.length - 1].getXmlAsync(function (x) {
              window.CHGDBG += " xml:" + x.status + (x.error ? " " + x.error.message : "") + " len=" + (x.value ? x.value.length : 0);
              if (x.status !== "succeeded") return res(null);
              var doc = new DOMParser().parseFromString(x.value, "text/xml");
              var root = doc.documentElement;
              var dEl = root.getElementsByTagName("d")[0];
              if (dEl) {
                var lines = dEl.textContent.split("\n");
                var h = lines[0].split("|");
                var items = lines.slice(1).filter(Boolean).map(function (ln) {
                  var p = ln.split("|");
                  return { i: +p[0], pos: +p[1], text: p.slice(2).join("|") };
                });
                return res({ vs: +h[0], last: +h[1], n: +h[2], stamp: h[3], items: items });
              }
              var items2 = Array.prototype.map.call(root.getElementsByTagName("c"), function (c) {
                return { i: +c.getAttribute("i"), pos: +c.getAttribute("pos"), text: c.textContent };
              });
              res({ vs: +root.getAttribute("vs"), last: +root.getAttribute("last"), n: +root.getAttribute("n"), stamp: root.getAttribute("stamp"), items: items2 });
            });
          });
        } catch (e) { res(null); }
      });
    },

    // change regions marked by the macro (bookmarks zChgP#, higher number = earlier in the document)
    listChangeMarks: function () {
      if (!Office.context.requirements.isSetSupported("WordApi", "1.4")) return Promise.resolve(null);
      return run(function (ctx) {
        var names = ctx.document.body.getRange("Whole").getBookmarks(true, true);
        return ctx.sync().then(function () {
          var list = names.value.filter(function (n) { return /^zChgP\d+$/.test(n); })
            .sort(function (a, b) { return +b.slice(5) - +a.slice(5); });
          var paras = list.map(function (n) {
            var p = ctx.document.getBookmarkRangeOrNullObject(n).paragraphs.getFirst();
            p.load("text");
            return p;
          });
          return ctx.sync().then(function () {
            return list.map(function (n, i) { return { name: n, i: i + 1, text: (paras[i].isNullObject ? "" : paras[i].text).trim().slice(0, 90) }; });
          });
        });
      });
    },

    // Word comments (sync notes and any other notes) for the notes tab
    listNotes: function () {
      if (!Office.context.requirements.isSetSupported("WordApi", "1.4")) return Promise.resolve(null);
      return run(function (ctx) {
        var cs = ctx.document.body.getComments();
        cs.load("items/id,items/content,items/resolved,items/authorName");
        return ctx.sync().then(function () {
          var paras = cs.items.map(function (c) {
            var p = c.getRange().paragraphs.getFirst();
            p.load("text,style");
            return p;
          });
          var reps = cs.items.map(function (c) {
            try { var r = c.replies; r.load("items/content"); return r; } catch (e) { return null; }
          });
          return ctx.sync().then(function () {
            return cs.items.map(function (c, i) {
              var rr = reps[i] && reps[i].items ? reps[i].items.map(function (x) { return (x.content || "").trim(); }) : [];
              return { id: c.id, text: (c.content || "").trim(), resolved: !!c.resolved, author: c.authorName || "",
                       replies: rr,
                       where: (paras[i].text || "").trim().slice(0, 90), style: paras[i].style || "" };
            });
          });
        });
      });
    },
    gotoNote: function (id) {
      return run(function (ctx) {
        var cs = ctx.document.body.getComments();
        cs.load("items/id");
        return ctx.sync().then(function () {
          var c = cs.items.filter(function (x) { return x.id === id; })[0];
          if (c) c.getRange().select("Select");
          return ctx.sync();
        });
      });
    },
    resolveNote: function (id, done) {
      return run(function (ctx) {
        var cs = ctx.document.body.getComments();
        cs.load("items/id");
        return ctx.sync().then(function () {
          var c = cs.items.filter(function (x) { return x.id === id; })[0];
          if (c) c.resolved = done;
          return ctx.sync();
        });
      });
    },

    // --- review card: yellow mark on the note being worked on, and inserting a prepared suggestion
    markNote: function (id, on) {
      var self = window.WordBridge;
      self._hl = self._hl || {};
      return run(function (ctx) {
        var cs = ctx.document.body.getComments();
        cs.load("items/id");
        return ctx.sync().then(function () {
          var c = cs.items.filter(function (x) { return x.id === id; })[0];
          if (!c) return;
          var r = c.getRange();
          if (on) {
            r.font.load("highlightColor");
            return ctx.sync().then(function () {
              var h = r.font.highlightColor;
              if (h === null || h === "" || h === undefined) {   // only mark text that has no highlight of its own
                r.font.highlightColor = "Yellow"; self._hl[id] = true;
              }
              r.select("Select");
              return ctx.sync();
            });
          }
          if (self._hl[id]) { r.font.highlightColor = null; delete self._hl[id]; }
          return ctx.sync();
        });
      });
    },
    // sug: {kind:"add", after:"", lines:[{style,text}]} | {kind:"replace", old:"", neu:""}
    applySuggestion: function (id, sug) {
      return run(function (ctx) {
        var cs = ctx.document.body.getComments();
        cs.load("items/id");
        return ctx.sync().then(function () {
          var c = cs.items.filter(function (x) { return x.id === id; })[0];
          if (!c) throw new Error("ההערה לא נמצאה במסמך");
          var para = c.getRange().paragraphs.getLast();
          if (sug.kind === "replace") {
            if (!sug.old || sug.old.length > 250) throw new Error("הקטע להחלפה חסר או ארוך מדי");
            var inPara = c.getRange().paragraphs.getFirst().search(sug.old, { matchCase: true });
            var inBody = ctx.document.body.search(sug.old, { matchCase: true });
            inPara.load("items"); inBody.load("items");
            return ctx.sync().then(function () {
              var hit = inPara.items[0] || inBody.items[0];
              if (!hit) throw new Error("הקטע \"" + sug.old.slice(0, 40) + "…\" לא נמצא בתסריט");
              hit.insertText(sug.neu, "Replace");
              c.resolved = true;
              return ctx.sync();
            });
          }
          var findAfter = sug.after ? ctx.document.body.search(sug.after.slice(0, 250), { matchCase: true }) : null;
          if (findAfter) findAfter.load("items");
          return ctx.sync().then(function () {
            var base = findAfter && findAfter.items[0] ? findAfter.items[0].paragraphs.getLast() : para;
            sug.lines.forEach(function (ln) {
              base = base.insertParagraph(ln.text, "After");
              base.style = ln.style;
              base.font.highlightColor = null;
            });
            c.resolved = true;
            return ctx.sync();
          });
        });
      });
    },

    gotoMark: function (name) {
      return run(function (ctx) {
        var r = ctx.document.getBookmarkRangeOrNullObject(name);
        r.select("Start");
        return ctx.sync();
      });
    },

    // leave a request for the macro (picked up within ~2 seconds)
    sendRequest: function (cmd, arg) {
      return new Promise(function (res, rej) {
        try {
          var esc = function (t) { return String(t).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;"); };
          Office.context.document.customXmlParts.addAsync('<req xmlns="urn:scripttool:req" cmd="' + esc(cmd) + '" arg="' + esc(arg) + '"/>', function (r) {
            if (r.status === "succeeded") res(); else rej(new Error("request failed"));
          });
        } catch (e) { rej(e); }
      });
    },

    loadData: function () {
      try {
        var st = Office.context.document.settings;
        var d = st.get(KEY);
        if (d) return d;
        var old = st.get(OLD_KEY);
        if (old) return { scenes: old.map(function (o) { return { h: o.h, s: o.s, q: "", p: [] }; }), plots: [] };
      } catch (e) {}
      return null;
    },

    saveData: function (data) {
      try {
        var st = Office.context.document.settings;
        st.set(KEY, data);
        st.saveAsync();
      } catch (e) {}
    },

    onSelectionChange: function (cb) {
      try {
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, cb);
      } catch (e) {}
    }
  };
})();
