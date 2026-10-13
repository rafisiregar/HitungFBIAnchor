/* Unit test: node test/test_fbi.js
   expected.json dihasilkan python/fbi_calc.py (python test/test_fbi.py --regen) — memastikan JS & Python identik. */
const FBI = require("../js/fbi.js");
const fx = require("./fixture.json");
const exp = require("./expected.json");
let bad = 0;
const r = FBI.compute(fx.raw, fx.master, { joinKey: "TID", qrisMode: "MAPPING" });
r.rows.forEach(({ raw, res }, k) => {
  const e = exp[k];
  if (!e) { if (res.ket !== "MDR TIDAK DITEMUKAN") { bad++; console.log("Harus unmatched:", raw.TID); } return; }
  for (const c of FBI.CFG.OUTPUT_COLUMNS)
    if (Math.abs(res.gross[c] - e[0][c]) > 1e-6 || Math.abs(res.itc[c] - e[1][c]) > 1e-6 || Math.abs(res.net[c] - e[2][c]) > 1e-6) { bad++; console.log("MISMATCH", raw.TID, c); }
  if (Math.abs(res.total - e[3]) > 1e-6) { bad++; console.log("TOTAL MISMATCH", raw.TID); }
});
// Mode MAPPING: QRIS SV -> On Us, QRIS-CB/NFC/TITO -> Off Us
if (!(r.summary.net["QRIS ON US"] > 0 && r.summary.net["QRIS OFF US"] > 0)) { bad++; console.log("MAPPING: QRIS On/Off harus terisi"); }
const off = FBI.compute(fx.raw, fx.master, { joinKey: "TID", qrisMode: "OFF_US" });
if (off.summary.net["QRIS ON US"] !== 0) { bad++; console.log("OFF_US: QRIS ON US harus 0"); }
const on = FBI.compute(fx.raw, fx.master, { joinKey: "TID", qrisMode: "ON_US" });
if (on.summary.net["QRIS OFF US"] !== 0 || on.summary.itc["QRIS ON US"] !== 0) { bad++; console.log("ON_US: QRIS OFF US & ITC harus 0"); }
// Join pakai GROUP: hanya ANVAYA GROUP yang tidak ada di master dummy
const g = FBI.compute(fx.raw, fx.master, { joinKey: "GROUP" });
if (g.summary.unmatched !== 1) { bad++; console.log("Join GROUP: unmatched harus 1"); }
const aoa = FBI.buildResultAOA(r, fx.raw_headers, "Juni", null);
if (aoa.aoa[1].length !== fx.raw_headers.length + 7 * 3 + 2) { bad++; console.log("Jumlah kolom output salah"); }
console.log(bad ? `FAIL (${bad})` : `OK — ${r.rows.length} baris, total net ${r.summary.total.toFixed(2)}`);
process.exit(bad ? 1 : 0);
