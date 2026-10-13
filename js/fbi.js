/* =====================================================================
   LOGIKA PERHITUNGAN FBI  (murni, tanpa DOM)  — dipakai index.html & test
   ===================================================================== */
(function (root, factory) {
  if (typeof module !== "undefined") module.exports = factory(require("./config.js"));
  else root.FBI = factory(root.FBI_CONFIG);
})(typeof self !== "undefined" ? self : this, function (CFG) {

  const norm = (s) => String(s ?? "").trim().toUpperCase().replace(/\s+/g, " ");
  const num  = (v) => { const n = typeof v === "number" ? v : parseFloat(String(v ?? "").replace(/[^\d.\-eE]/g, "")); return isFinite(n) ? n : 0; };

  /* Ambil nilai kolom dari object dengan pencocokan nama case-insensitive */
  function pick(row, colName) {
    if (!row) return undefined;
    const target = norm(colName);
    for (const k in row) if (norm(k) === target) return row[k];
    return undefined;
  }

  /* Bangun mapping efektif:
     1. Mode override (OFF_US / ON_US): semua kategori QRIS disamakan dengan baris "QRIS OFF US" / "QRIS ON US".
     2. Normalisasi QRIS: KLASIFIKASI "QRIS" => "OFF US" jika punya KET INTERCHANGE, selain itu "ON US";
        KET GABUNGAN "QRIS" => "QRIS OFF US" / "QRIS ON US".
     Mapping asli (JSON_KET_MAPPING) tidak diubah — hasil normalisasi disimpan di _KLASIFIKASI / _GABUNGAN. */
  function effectiveMapping(qrisMode) {
    const m = CFG.JSON_KET_MAPPING.map((r) => ({ ...r }));
    if (qrisMode === "OFF_US" || qrisMode === "ON_US") {
      const ref = m.find((r) => norm(r.KATEGORI) === (qrisMode === "OFF_US" ? "QRIS OFF US" : "QRIS ON US"));
      if (ref) m.forEach((r) => {
        if (norm(r.GROUPING) === "QRIS") { r["KET MDR"] = ref["KET MDR"]; if (ref["KET INTERCHANGE"]) r["KET INTERCHANGE"] = ref["KET INTERCHANGE"]; else delete r["KET INTERCHANGE"]; }
      });
    }
    m.forEach((r) => {
      let kls = norm(r.KLASIFIKASI), gab = norm(r["KET GABUNGAN"]);
      if (kls === "QRIS") kls = r["KET INTERCHANGE"] ? "OFF US" : "ON US";
      if (gab === "QRIS") gab = kls === "OFF US" ? "QRIS OFF US" : "QRIS ON US";
      r._KLASIFIKASI = kls; r._GABUNGAN = gab;
    });
    return m;
  }

  /* Index master: key -> baris master (key pertama yang ketemu menang) */
  function indexMaster(masterRows, joinKey) {
    const idx = new Map();
    for (const r of masterRows) {
      const k = norm(pick(r, joinKey));
      if (k && !idx.has(k)) idx.set(k, r);
    }
    return idx;
  }

  /* Hitung satu baris raw. Mengembalikan {gross:{col:val}, itc:{}, net:{}, total, ket, detail:[]} */
  function computeRow(raw, master, mapping) {
    const zero = () => Object.fromEntries(CFG.OUTPUT_COLUMNS.map((c) => [c, 0]));
    const out = { gross: zero(), itc: zero(), net: zero(), total: 0, ket: "", detail: [] };
    if (!master) { out.ket = "MDR TIDAK DITEMUKAN"; return out; }

    const warn = new Set();
    for (const m of mapping) {
      const sv = num(pick(raw, m.KATEGORI + CFG.SV_SUFFIX));
      if (!sv) continue;
      const outCol = CFG.GABUNGAN_TO_OUTPUT[m._GABUNGAN];
      if (!outCol || !m["KET MDR"]) continue;                 // kategori tanpa KET MDR tidak dihitung

      const mdrRaw = pick(master, m["KET MDR"]);
      if (mdrRaw === undefined) warn.add("KOLOM " + m["KET MDR"] + " TIDAK ADA");
      const mdr = num(mdrRaw);
      const gross = sv * mdr;

      let nett = 0, nettCol = "";
      if (m._KLASIFIKASI === "OFF US") {
        nettCol = m["KET INTERCHANGE"] || CFG.NETT_COLUMN_BY_GROUPING[m.GROUPING] || "";
        const nRaw = nettCol ? pick(master, nettCol) : undefined;
        if (nettCol && nRaw === undefined) warn.add("KOLOM " + nettCol + " TIDAK ADA");
        nett = num(nRaw);
      }
      const itc = sv * nett;
      const net = gross - itc;

      out.gross[outCol] += gross; out.itc[outCol] += itc; out.net[outCol] += net;
      out.detail.push({ kategori: m.KATEGORI, sv, mdr, nettCol, nett, gross, itc, net, outCol });
    }
    out.total = CFG.OUTPUT_COLUMNS.reduce((s, c) => s + out.net[c], 0);
    out.ket = [...warn].join("; ");
    return out;
  }

  /* Proses seluruh dataset */
  function compute(rawRows, masterRows, opts = {}) {
    const joinKey  = opts.joinKey  || CFG.JOIN_KEY_DEFAULT;
    const qrisMode = opts.qrisMode || CFG.QRIS_MODE_DEFAULT;
    const mapping  = effectiveMapping(qrisMode);
    const idx      = indexMaster(masterRows, joinKey);
    const rows = rawRows.map((raw) => {
      const master = idx.get(norm(pick(raw, joinKey)));
      return { raw, res: computeRow(raw, master, mapping) };
    });
    const summary = { gross: {}, itc: {}, net: {}, total: 0, rows: rows.length, unmatched: 0 };
    for (const c of CFG.OUTPUT_COLUMNS) { summary.gross[c] = 0; summary.itc[c] = 0; summary.net[c] = 0; }
    for (const { res } of rows) {
      if (res.ket === "MDR TIDAK DITEMUKAN") summary.unmatched++;
      for (const c of CFG.OUTPUT_COLUMNS) { summary.gross[c] += res.gross[c]; summary.itc[c] += res.itc[c]; summary.net[c] += res.net[c]; }
      summary.total += res.total;
    }
    return { rows, summary, joinKey, qrisMode, mapping };
  }

  /* Susun array-of-arrays (AOA) untuk sheet hasil: 2 baris header + data, plus daftar merge */
  function buildResultAOA(result, rawHeaders, bulan, rawTitleRow) {
    const L = (t) => t.replace("{BULAN}", bulan || "");
    const cols = CFG.OUTPUT_COLUMNS, n = cols.length, base = rawHeaders.length;
    const h1 = new Array(base).fill(null); if (rawTitleRow) rawTitleRow.forEach((v, i) => (h1[i] = v ?? null));
    const h2 = [...rawHeaders];
    const merges = [];
    const blocks = [["GROSS", "gross"], ["ITC", "itc"], ["NET", "net"]];
    let c = base;
    for (const [lab] of blocks) {
      h1.push(L(CFG.BLOCK_LABELS[lab]), ...new Array(n - 1).fill(null));
      h2.push(...cols);
      merges.push({ s: { r: 0, c }, e: { r: 0, c: c + n - 1 } });
      c += n;
    }
    h1.push(L(CFG.BLOCK_LABELS.TOTAL), CFG.BLOCK_LABELS.KET);
    h2.push(L(CFG.BLOCK_LABELS.TOTAL), CFG.BLOCK_LABELS.KET);
    merges.push({ s: { r: 0, c }, e: { r: 1, c } }, { s: { r: 0, c: c + 1 }, e: { r: 1, c: c + 1 } });

    const aoa = [h1, h2];
    for (const { raw, res } of result.rows) {
      const line = rawHeaders.map((h) => raw[h] ?? null);
      for (const [, key] of blocks) line.push(...cols.map((k) => res[key][k]));
      line.push(res.total, res.ket || null);
      aoa.push(line);
    }
    return { aoa, merges, firstCalcCol: base };
  }

  /* Rekap per kolom grup (default GROUP) */
  function buildRekapAOA(result, groupCol, bulan) {
    const cols = CFG.OUTPUT_COLUMNS;
    const agg = new Map();
    for (const { raw, res } of result.rows) {
      const g = String(pick(raw, groupCol) ?? "(kosong)");
      if (!agg.has(g)) agg.set(g, { tid: 0, gross: 0, itc: 0, net: Object.fromEntries(cols.map((c) => [c, 0])), total: 0 });
      const a = agg.get(g); a.tid++;
      for (const c of cols) { a.gross += res.gross[c]; a.itc += res.itc[c]; a.net[c] += res.net[c]; }
      a.total += res.total;
    }
    const header = [groupCol, "JML TID", "GROSS FBI", "INTERCHANGE FEE", ...cols.map((c) => "NET " + c), "TOTAL NET FBI " + (bulan || "")];
    const aoa = [header];
    [...agg.entries()].sort((a, b) => b[1].total - a[1].total).forEach(([g, a]) => aoa.push([g, a.tid, a.gross, a.itc, ...cols.map((c) => a.net[c]), a.total]));
    const s = result.summary;
    aoa.push(["TOTAL", result.rows.length, cols.reduce((x, c) => x + s.gross[c], 0), cols.reduce((x, c) => x + s.itc[c], 0), ...cols.map((c) => s.net[c]), s.total]);
    return aoa;
  }

  return { norm, num, pick, effectiveMapping, indexMaster, computeRow, compute, buildResultAOA, buildRekapAOA, CFG };
});
