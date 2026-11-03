# Rekap Supervisi Konstruksi (Google Sheet + Apps Script)

`Code.gs` merekap semua file di folder Drive sumber ke satu sheet **Rekap**,
dengan kolom **Diambil Pada**, **Nama File Sumber**, **File Diubah Terakhir**,
**ID File**, **Catatan** di kanan tabel, dan sheet **Log** untuk riwayat tiap run.

## Pasang (sekali, ±3 menit)
1. Buka Google Sheet "Rekap Supervisi Konstruksi" → **Extensions → Apps Script**.
2. Ganti isi `Code.gs` dengan file `Code.gs` di repo ini.
3. **Project Settings** → centang *Show appsscript.json* → ganti isinya dengan `appsscript.json` di repo
   (mengaktifkan layanan Drive API v3, dipakai untuk konversi .xlsx sementara).
4. Pilih fungsi `updateRekap` → **Run** → setujui izin (akun harus bisa membaca folder sumber).
5. Menu **Rekap → Pasang jadwal harian** agar jalan otomatis tiap pukul 19.00.

## Cara kerja
- Tiap file `.xlsx`/Google Sheet di folder dibuka, lalu sheet berawalan `List_` (mis. `List_SKTT`) dibaca.
  Kolom dicocokkan **berdasarkan nama header**, bukan posisi (`PST` di sumber dibaca sebagai `PSO`).
- File .xlsx dikonversi sementara ke Google Sheet lalu dihapus. File berukuran 10–17 MB, jadi
  satu file bisa makan puluhan detik. Script berhenti di 4,5 menit dan **melanjutkan otomatis** 1 menit kemudian.
- Hanya file yang **berubah sejak pengambilan terakhir** yang diproses ulang. **Rekap → Ambil ulang SEMUA file** memaksa semuanya.
- `Jenis` = akhiran nama sheet (`List_SKTT` → `SKTT`). `to decimal actual/plan` = progres ÷ 100
  (atau apa adanya bila sel berformat %). `Deviasi` = actual − plan.

## Yang perlu dicek setelah run pertama
- **Kolom Catatan** menandai sel rumus error (`#NAME?`) dan header yang tidak ditemukan.
  Pada file contoh, banyak sel di `List_SKTT` berupa rumus yang menghasilkan `#NAME?` saat dibaca di luar Excel.
  Jika setelah konversi hasilnya tetap error, nilainya kosong dan perlu diperbaiki di file sumber
  (atau simpan sel itu sebagai nilai).
- `Milestone` dan `Bullet` tidak ada di sheet `List_SKTT` yang diperiksa; kolom ini kosong kecuali ada header
  bernama sama di sumber.
- Bila file lain punya nama sheet data yang berbeda dari `List_*`, ubah `LIST_SHEET_PREFIX`.
