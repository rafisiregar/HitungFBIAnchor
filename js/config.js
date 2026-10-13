/* =====================================================================
   KONFIGURASI PERHITUNGAN FBI
   File ini SATU-SATUNYA yang perlu diubah jika aturan hitung berubah.
   Semua nilai di sini juga bisa dipilih user lewat dropdown di index.html.
   ===================================================================== */

const FBI_CONFIG = {

  /* ---------- 1. KUNCI JOIN RAW -> MASTER MDR ----------
     Nama kolom yang ada di KEDUA file (raw & master).
     Pilihan yang muncul di dropdown; "default" dipakai saat halaman dibuka. */
  JOIN_KEYS: ["TID", "MID", "GROUP", "NAMA_MERCHANT"],
  JOIN_KEY_DEFAULT: "TID",

  /* ---------- 2. PERLAKUAN KOLOM QRIS DI RAW ----------
     Raw hanya punya "QRIS SV", "QRIS-CB SV", "QRIS-NFC SV", "QRIS-TITO SV"
     tanpa pemisahan On Us / Off Us. Pilih perlakuannya:
       "MAPPING" -> ikuti JSON_KET_MAPPING apa adanya
                    (4 kategori itu tidak punya KET MDR => FBI QRIS = 0)
       "OFF_US"  -> semua kategori QRIS raw dihitung sebagai QRIS OFF US
       "ON_US"   -> semua kategori QRIS raw dihitung sebagai QRIS ON US  */
  QRIS_MODES: {
    MAPPING: "Ikuti mapping: QRIS = On Us; QRIS-CB/NFC/TITO = Off Us",
    OFF_US:  "Override: semua QRIS raw = QRIS Off Us",
    ON_US:   "Override: semua QRIS raw = QRIS On Us",
  },
  QRIS_MODE_DEFAULT: "MAPPING",

  /* ---------- 3. SUFFIX KOLOM SV DI RAW ----------  "KB-MC SV" = "KB-MC" + " SV" */
  SV_SUFFIX: " SV",

  /* ---------- 4. NAMA KOLOM PENGALI NETT DI MASTER ----------
     Dipakai untuk KLASIFIKASI OFF US yang tidak punya KET INTERCHANGE sendiri. */
  NETT_COLUMN_BY_GROUPING: {
    DEBIT: "DEBIT",
    QRIS:  "QRIS",
    // KREDIT memakai KET INTERCHANGE per kategori (ITC VISA / MC / JCB / LAINNYA)
  },

  /* ---------- 5. KOLOM OUTPUT (urutan = urutan di Excel hasil) ---------- */
  OUTPUT_COLUMNS: ["BRIZZI", "CC ON US", "CC OFF US", "DEBIT ON US", "DEBIT OFF US", "QRIS ON US", "QRIS OFF US"],

  /* KET GABUNGAN -> kolom output */
  GABUNGAN_TO_OUTPUT: {
    "BRIZZI ON US":  "BRIZZI",
    "KREDIT ON US":  "CC ON US",
    "KREDIT OFF US": "CC OFF US",
    "DEBIT ON US":   "DEBIT ON US",
    "DEBIT OFF US":  "DEBIT OFF US",
    "QRIS ON US":    "QRIS ON US",
    "QRIS OFF US":   "QRIS OFF US",
  },

  /* ---------- 6. LABEL BLOK HEADER OUTPUT ----------  {BULAN} diganti otomatis */
  BLOCK_LABELS: {
    GROSS: "Gross FBI {BULAN}",
    ITC:   "Interchange Fee {BULAN}",
    NET:   "Net FBI {BULAN}",
    TOTAL: "Total Net FBI {BULAN}",
    KET:   "KETERANGAN",
  },

  /* ---------- 7. MAPPING KATEGORI (sama persis dengan JSON_KET_MAPPING Python) ----------
     Aturan hitung per baris:
       ON US  : gross = net = SV x MDR ; itc = 0
       OFF US : gross = SV x MDR ; itc = SV x NETT ; net = gross - itc
                NETT = master[KET INTERCHANGE] jika ada, selain itu master[NETT_COLUMN_BY_GROUPING[GROUPING]]
       (Debit: SV x (MDR - NETT) dan Kredit: SVxMDR - SVxNETT secara matematis identik)       */
  JSON_KET_MAPPING: [
    {"KATEGORI": "BLANK",          "GROUPING": "KREDIT", "KLASIFIKASI": "ON US",  "KET GABUNGAN": "KREDIT ON US",  "KET MDR": "MDR On Us"},
    {"KATEGORI": "BRIZZI",         "GROUPING": "BRIZZI", "KLASIFIKASI": "ON US",  "KET GABUNGAN": "BRIZZI ON US",  "KET MDR": "MDR Uang Elektronik"},
    {"KATEGORI": "KB-JCB",         "GROUPING": "KREDIT", "KLASIFIKASI": "OFF US", "KET GABUNGAN": "KREDIT OFF US", "KET MDR": "MDR Off Us", "KET INTERCHANGE": "ITC JCB"},
    {"KATEGORI": "KB-LAINNYA",     "GROUPING": "KREDIT", "KLASIFIKASI": "OFF US", "KET GABUNGAN": "KREDIT OFF US", "KET MDR": "MDR Off Us", "KET INTERCHANGE": "ITC LAINNYA"},
    {"KATEGORI": "KB-LAINNYA-NPG", "GROUPING": "DEBIT",  "KLASIFIKASI": "OFF US", "KET GABUNGAN": "DEBIT OFF US",  "KET MDR": "MDR NPG"},
    {"KATEGORI": "KB-MC",          "GROUPING": "KREDIT", "KLASIFIKASI": "OFF US", "KET GABUNGAN": "KREDIT OFF US", "KET MDR": "MDR Off Us", "KET INTERCHANGE": "ITC MC"},
    {"KATEGORI": "KB-MC-NPG",      "GROUPING": "DEBIT",  "KLASIFIKASI": "OFF US", "KET GABUNGAN": "DEBIT OFF US",  "KET MDR": "MDR NPG"},
    {"KATEGORI": "KB-VISA",        "GROUPING": "KREDIT", "KLASIFIKASI": "OFF US", "KET GABUNGAN": "KREDIT OFF US", "KET MDR": "MDR Off Us", "KET INTERCHANGE": "ITC VISA"},
    {"KATEGORI": "KB-VISA-NPG",    "GROUPING": "DEBIT",  "KLASIFIKASI": "OFF US", "KET GABUNGAN": "DEBIT OFF US",  "KET MDR": "MDR NPG"},
    {"KATEGORI": "KD-BRI",         "GROUPING": "DEBIT",  "KLASIFIKASI": "ON US",  "KET GABUNGAN": "DEBIT ON US",   "KET MDR": "MDR Debit"},
    {"KATEGORI": "KD-BRI-TARTUN",  "GROUPING": "DEBIT",  "KLASIFIKASI": "ON US",  "KET GABUNGAN": "DEBIT ON US",   "KET MDR": "MDR Debit"},
    {"KATEGORI": "KK-BRI-JCB",     "GROUPING": "KREDIT", "KLASIFIKASI": "ON US",  "KET GABUNGAN": "KREDIT ON US",  "KET MDR": "MDR On Us"},
    {"KATEGORI": "KK-BRI-MC",      "GROUPING": "KREDIT", "KLASIFIKASI": "ON US",  "KET GABUNGAN": "KREDIT ON US",  "KET MDR": "MDR On Us"},
    {"KATEGORI": "KK-BRI-VISA",    "GROUPING": "KREDIT", "KLASIFIKASI": "ON US",  "KET GABUNGAN": "KREDIT ON US",  "KET MDR": "MDR On Us"},
    {"KATEGORI": "KK-LAINNYA",     "GROUPING": "KREDIT", "KLASIFIKASI": "OFF US", "KET GABUNGAN": "KREDIT OFF US", "KET MDR": "MDR Off Us", "KET INTERCHANGE": "ITC LAINNYA"},
    /* QRIS: KLASIFIKASI "QRIS" dibaca engine sebagai: ada KET INTERCHANGE => OFF US, tidak ada => ON US.
       KET GABUNGAN "QRIS" otomatis dipetakan ke kolom QRIS ON US / QRIS OFF US sesuai hasil itu. */
    {"KATEGORI": "QRIS",           "GROUPING": "QRIS",   "KLASIFIKASI": "QRIS",   "KET GABUNGAN": "QRIS", "KET MDR": "MDR QRIS On Us"},
    {"KATEGORI": "QRIS-CB",        "GROUPING": "QRIS",   "KLASIFIKASI": "QRIS",   "KET GABUNGAN": "QRIS", "KET MDR": "MDR QRIS Off US", "KET INTERCHANGE": "QRIS"},
    {"KATEGORI": "QRIS-NFC",       "GROUPING": "QRIS",   "KLASIFIKASI": "QRIS",   "KET GABUNGAN": "QRIS", "KET MDR": "MDR QRIS Off US", "KET INTERCHANGE": "QRIS"},
    {"KATEGORI": "QRIS-TITO",      "GROUPING": "QRIS",   "KLASIFIKASI": "QRIS",   "KET GABUNGAN": "QRIS", "KET MDR": "MDR QRIS Off US", "KET INTERCHANGE": "QRIS"},
    {"KATEGORI": "QRIS ON US",     "GROUPING": "QRIS",   "KLASIFIKASI": "QRIS",   "KET GABUNGAN": "QRIS", "KET MDR": "MDR QRIS On US"},
    {"KATEGORI": "QRIS OFF US",    "GROUPING": "QRIS",   "KLASIFIKASI": "QRIS",   "KET GABUNGAN": "QRIS", "KET MDR": "MDR QRIS Off US", "KET INTERCHANGE": "QRIS"},
  ],

  /* ---------- 8. MASTER MDR TERTANAM (opsional, untuk masa depan) ----------
     Jika diisi (array of object dengan kolom sama seperti sheet "MDR DAN NETT"),
     aplikasi bisa jalan tanpa upload master. Kosongkan [] = wajib upload master. */
  EMBEDDED_MASTER: [],
  MASTER_SHEET_NAME: "MDR DAN NETT",   // sheet yang dicari di file master
  MASTER_HEADER_ROW: 2,                // baris header (baris 1 = judul blok)
  RAW_HEADER_ROW: 2,                   // baris header raw (baris 1 = nama bulan)
};

if (typeof module !== "undefined") module.exports = FBI_CONFIG;
if (typeof window !== "undefined") window.FBI_CONFIG = FBI_CONFIG;
