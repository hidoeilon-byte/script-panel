// Keeps each scene's synopsis attached to the right scene while scenes are
// added, deleted, renamed or moved. Pure function, no Word dependency.
// old: [{h: headingText, s: synopsis}], heads: [headingText]
// returns: [synopsis] aligned to heads
(function (root) {
  function alignSynopses(old, heads) {
    var n = old.length, m = heads.length, i, j;
    var out = new Array(m).fill("");
    var usedOld = new Array(n).fill(false);
    var gotNew = new Array(m).fill(false);
    // 1. LCS on exact heading text
    var L = [];
    for (i = 0; i <= n; i++) { L.push(new Array(m + 1).fill(0)); }
    for (i = n - 1; i >= 0; i--) {
      for (j = m - 1; j >= 0; j--) {
        L[i][j] = old[i].h === heads[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
      }
    }
    var pairs = [];
    i = 0; j = 0;
    while (i < n && j < m) {
      if (old[i].h === heads[j]) { pairs.push([i, j]); i++; j++; }
      else if (L[i + 1][j] >= L[i][j + 1]) { i++; } else { j++; }
    }
    pairs.forEach(function (p) { out[p[1]] = old[p[0]].s; usedOld[p[0]] = true; gotNew[p[1]] = true; });
    // 2. Gaps of equal size between anchors = renamed headings, map by position
    var anchors = [[-1, -1]].concat(pairs, [[n, m]]);
    for (var k = 0; k < anchors.length - 1; k++) {
      var a = anchors[k], b = anchors[k + 1];
      var go = b[0] - a[0] - 1, gn = b[1] - a[1] - 1;
      if (go > 0 && go === gn) {
        for (var t = 1; t <= go; t++) {
          out[a[1] + t] = old[a[0] + t].s;
          usedOld[a[0] + t] = true; gotNew[a[1] + t] = true;
        }
      }
    }
    // 3. Moved scenes: same text, not matched yet
    for (j = 0; j < m; j++) {
      if (gotNew[j]) continue;
      for (i = 0; i < n; i++) {
        if (!usedOld[i] && old[i].h === heads[j]) {
          out[j] = old[i].s; usedOld[i] = true; gotNew[j] = true; break;
        }
      }
    }
    return out;
  }
  root.alignSynopses = alignSynopses;
  if (typeof module !== "undefined") module.exports = alignSynopses;
})(typeof window !== "undefined" ? window : globalThis);
