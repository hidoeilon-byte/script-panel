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
