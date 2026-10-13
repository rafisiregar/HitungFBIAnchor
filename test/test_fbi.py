"""python test/test_fbi.py          -> bandingkan python/fbi_calc.py dengan expected.json
   python test/test_fbi.py --regen  -> tulis ulang expected.json dari hasil Python"""
import json, sys, os, pandas as pd
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "python"))
import fbi_calc as F
here = os.path.dirname(__file__)
fx = json.load(open(os.path.join(here, "fixture.json")))
raw, mas = pd.DataFrame(fx["raw"]), pd.DataFrame(fx["master"])
h, _ = F.FBICalculator().compute(raw, mas, bulan="Juni")
out = []
for _, r in h.iterrows():
    if r[("KETERANGAN", "KETERANGAN")] == F.UNMATCHED_LABEL: out.append(None); continue
    out.append([{c: r[("Gross FBI Juni", c)] for c in F.OUTPUT_COLUMNS}, {c: r[("Interchange Fee Juni", c)] for c in F.OUTPUT_COLUMNS},
                {c: r[("Net FBI Juni", c)] for c in F.OUTPUT_COLUMNS}, r[("Total Net FBI Juni", "Total Net FBI Juni")]])
if "--regen" in sys.argv:
    json.dump(out, open(os.path.join(here, "expected.json"), "w")); print("expected.json diperbarui"); sys.exit(0)
exp = json.load(open(os.path.join(here, "expected.json")))
bad = 0
for e, o in zip(exp, out):
    if (e is None) != (o is None): bad += 1; continue
    if e is None: continue
    for i in range(3):
        for c in F.OUTPUT_COLUMNS:
            if abs(e[i][c] - o[i][c]) > 1e-6: bad += 1
    if abs(e[3] - o[3]) > 1e-6: bad += 1
print("FAIL" if bad else f"OK — {len(out)} baris identik dengan expected.json"); sys.exit(1 if bad else 0)
