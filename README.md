# Kalkulator FBI Merchant

Aplikasi web statis (HTML + JavaScript, SheetJS) untuk menghitung **Fee Based Income (FBI)** merchant EDC dari file raw Sales Volume bulanan dan master **MDR & Pengali Nett**. Seluruh perhitungan berjalan di browser — tidak ada data yang dikirim ke server, aman dipakai dari jaringan kantor, dan `index.html` bisa dibuka langsung (double-click) tanpa internet karena SheetJS sudah dibundel di `js/xlsx.full.min.js`.

Tersedia juga **modul Python** (`python/fbi_calc.py`) dengan logika identik untuk batch / verifikasi — unit test memastikan keduanya menghasilkan angka yang sama.

## Struktur repo

```
index.html                   UI: upload raw & master, parameter (dropdown), ringkasan, download Excel
js/config.js                 SEMUA aturan yang bisa diubah (mapping kategori, kunci join, mode QRIS, label, master tertanam)
js/fbi.js                    Logika hitung murni (dipakai index.html & unit test)
js/xlsx.full.min.js          SheetJS (offline)
python/fbi_calc.py           Versi Python: objek FBICalculator + CLI, konfigurasi sama dengan config.js
contoh/RAW_mentah_JUNI_26.xlsx                 Contoh raw (sheet "mentah")
contoh/MASTER_MDR_DAN_NETT_dummy.xlsx          Contoh master (sheet "MDR DAN NETT") — rate dummy untuk 6 dari 7 TID contoh
contoh/CONTOH_FBI_JUNI_26_hasil_7kolom.xlsx    Hasil hitung contoh (3 sheet: HASIL PERHITUNGAN, REKAP, PARAMETER)
test/                        fixture + expected + test JS & Python (jalan otomatis di GitHub Actions)
.github/workflows/pages.yml  Test JS + Python, lalu deploy ke GitHub Pages setiap push ke main
```

## Deploy ke GitHub Pages

1. Buat repo baru di GitHub, push seluruh isi folder ini ke branch `main`.
2. Di repo: **Settings → Pages → Build and deployment → Source: GitHub Actions**.
3. Setiap push ke `main`, workflow menjalankan unit test lalu mempublikasikan ke `https://<user>.github.io/<repo>/`.

> Catatan: GitHub Pages hanya melayani file statis — PHP/Python tidak bisa berjalan di sana. Karena itu perhitungan dilakukan di browser (JS), sedangkan Python dipakai untuk batch lokal / CI.

## Cara pakai (web)

| Langkah | Keterangan |
|---|---|
| 1. Raw | Upload file seperti sheet `mentah` (baris 1 = nama bulan, baris 2 = header, kolom `<KATEGORI> SV`). Label bulan terisi otomatis dari baris 1, bisa diedit. |
| 2. Master | **Upload file** dengan sheet `MDR DAN NETT` (header baris 2). Opsi lain: **Tersimpan di browser** (klik *Simpan master ini di browser* sekali, bulan depan cukup upload raw) atau **Tertanam di aplikasi** (aktif otomatis bila `EMBEDDED_MASTER` di `config.js` diisi). |
| 3. Parameter | **Kunci join** (TID / MID / GROUP / NAMA_MERCHANT), **mode QRIS**, kolom untuk sheet REKAP. |
| Hitung → Download | Excel: `HASIL PERHITUNGAN` (raw + Gross / Interchange / Net × 7 kolom + Total Net + KETERANGAN, header 2 baris merge seperti template), `REKAP` per grup, `PARAMETER` (jejak audit mapping & setting). |

## Cara pakai (Python)

```bash
pip install pandas openpyxl
python python/fbi_calc.py contoh/RAW_mentah_JUNI_26.xlsx contoh/MASTER_MDR_DAN_NETT_dummy.xlsx -o hasil.xlsx
# opsi: --join GROUP   --qris OFF_US   --bulan Juli   --rekap ANCHOR
```
```python
from fbi_calc import FBICalculator, read_sheet
calc = FBICalculator(join_key="TID", qris_mode="MAPPING")   # ganti join_key="GROUP" bila perlu
df_hasil, rekap = calc.compute(df_raw, df_master, bulan="Juni")
```

## Aturan hitung

| Klasifikasi | Kategori raw | Gross | Interchange Fee | Net |
|---|---|---|---|---|
| **ON US** | BLANK (→ CC On Us), BRIZZI, KD-BRI, KD-BRI-TARTUN, KK-BRI-VISA/MC/JCB, QRIS | SV × MDR | 0 | = Gross |
| **OFF US Kredit** | KB-VISA/MC/JCB/LAINNYA, KK-LAINNYA | SV × MDR Off Us | SV × ITC VISA/MC/JCB/LAINNYA | Gross − ITC |
| **OFF US Debit** | KB-VISA-NPG, KB-MC-NPG, KB-LAINNYA-NPG | SV × MDR NPG | SV × DEBIT | Gross − ITC |
| **OFF US QRIS** | QRIS-CB, QRIS-NFC, QRIS-TITO | SV × MDR QRIS Off Us | SV × QRIS | Gross − ITC |

Kolom output per blok: `BRIZZI · CC ON US · CC OFF US · DEBIT ON US · DEBIT OFF US · QRIS ON US · QRIS OFF US`.
Baris raw yang kuncinya tidak ada di master → FBI 0 dan `KETERANGAN = MDR TIDAK DITEMUKAN`. Kolom master yang tidak dipakai (mis. `MDR Debit Private Label`) dibiarkan apa adanya.

### Mode QRIS
- `MAPPING` (default) — ikuti `JSON_KET_MAPPING`: `QRIS SV` = On Us, `QRIS-CB/NFC/TITO SV` = Off Us
- `OFF_US` / `ON_US` — override: semua kolom QRIS raw dihitung sebagai Off Us / On Us

## Mengubah aturan
Semua ada di `js/config.js` (dan padanannya di bagian atas `python/fbi_calc.py`): `JSON_KET_MAPPING`, `JOIN_KEYS`, `QRIS_MODE_DEFAULT`, `OUTPUT_COLUMNS`, `BLOCK_LABELS`. Untuk menanam master MDR permanen, isi `EMBEDDED_MASTER` dengan array of object (kolom sama persis seperti sheet `MDR DAN NETT`).

Setelah mengubah logika: `node test/test_fbi.js` dan `python test/test_fbi.py` (atau `--regen` untuk memperbarui expected.json dari Python).
