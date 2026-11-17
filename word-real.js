// Real Word connection (Office.js). Same interface as the dev mock.
(function () {
  var SCENE = "כותרת סצנה";
  var KEY = "scenePanel.synopses";

  function run(fn) { return Word.run(fn); }

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

    // [{text, style}] for every paragraph in the body
    paragraphs: function () {
      return run(function (ctx) {
        var ps = ctx.document.body.paragraphs;
        ps.load("items/text,items/style");
        return ctx.sync().then(function () {
          return ps.items.map(function (p) { return { text: p.text, style: p.style }; });
        });
      });
    },

    // {style, sceneIndex} of the paragraph that holds the cursor
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

    loadSynopses: function () {
      try { return Office.context.document.settings.get(KEY) || []; } catch (e) { return []; }
    },

    saveSynopses: function (list) {
      try {
        var st = Office.context.document.settings;
        st.set(KEY, list);
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
