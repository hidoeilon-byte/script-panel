// Keeps each scene's data (synopsis, plots, sequence) attached to the right
// scene while scenes are added, deleted, renamed or moved.
// alignScenes(oldHeads, newHeads) -> for each new heading, the index of the
// old scene it continues, or -1 for a new scene.
(function (root) {
  function alignScenes(old, heads) {
    var n = old.length, m = heads.length, i, j;
    var map = new Array(m).fill(-1);
    var usedOld = new Array(n).fill(false);
    // 1. LCS on exact heading text
    var L = [];
    for (i = 0; i <= n; i++) { L.push(new Array(m + 1).fill(0)); }
    for (i = n - 1; i >= 0; i--) {
      for (j = m - 1; j >= 0; j--) {
        L[i][j] = old[i] === heads[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
      }
    }
    var pairs = [];
    i = 0; j = 0;
    while (i < n && j < m) {
      if (old[i] === heads[j]) { pairs.push([i, j]); i++; j++; }
      else if (L[i + 1][j] >= L[i][j + 1]) { i++; } else { j++; }
    }
    pairs.forEach(function (p) { map[p[1]] = p[0]; usedOld[p[0]] = true; });
    // 2. Gaps of equal size between anchors = renamed headings, map by position
    var anchors = [[-1, -1]].concat(pairs, [[n, m]]);
    for (var k = 0; k < anchors.length - 1; k++) {
      var a = anchors[k], b = anchors[k + 1];
      var go = b[0] - a[0] - 1, gn = b[1] - a[1] - 1;
      if (go > 0 && go === gn) {
        for (var t = 1; t <= go; t++) { map[a[1] + t] = a[0] + t; usedOld[a[0] + t] = true; }
      }
    }
    // 3. Moved scenes: same text, not matched yet
    for (j = 0; j < m; j++) {
      if (map[j] >= 0) continue;
      for (i = 0; i < n; i++) {
        if (!usedOld[i] && old[i] === heads[j]) { map[j] = i; usedOld[i] = true; break; }
      }
    }
    return map;
  }
  root.alignScenes = alignScenes;
  if (typeof module !== "undefined") module.exports = alignScenes;
})(typeof window !== "undefined" ? window : globalThis);
