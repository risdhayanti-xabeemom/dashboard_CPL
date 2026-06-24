# Dashboard Asesmen CPL dan CQI D3 Teknik Elektronika

Aplikasi Streamlit untuk mengolah nilai CPMK dosen menjadi capaian CPMK, IK, CPL, status ketercapaian CPL, laporan CQI, tren multi-semester, dan evaluasi efektivitas tindak lanjut berbasis OBE/IABEE.

Format input Excel lama tetap dipakai. Fitur multi-semester tidak menambah atau mengubah sheet input.

## Struktur File Aplikasi

- `app.py`: aplikasi utama Streamlit.
- `requirements.txt`: dependensi Python.
- `template_input.xlsx`: template Excel kosong dengan header lengkap, dibuat otomatis saat aplikasi dijalankan bila belum ada.
- `sample_data.xlsx`: data contoh kecil untuk uji coba, dibuat otomatis saat aplikasi dijalankan bila belum ada.
- `README.md`: panduan penggunaan.

Logo sidebar dapat memakai file berikut bila tersedia:

- `logo_polinema.png`, `logo_polinema.jpg`, atau file `OIP.jpg` dari folder Downloads.
- `logo_jte.png`, `logo_jte.jpg`, atau file `OIP (1).jpg` dari folder Downloads.

Jika file `Master_CPMK_44_MK_D3_PSTE_draft.xlsx` tersedia di folder aplikasi atau Downloads, aplikasi dapat menggunakannya sebagai sumber master CPL/IK/CPMK untuk membuat `template_input.xlsx`. Sheet nilai mahasiswa tetap perlu diisi pada `Nilai_CPMK` sebelum asesmen dihitung.

## Cara Menjalankan

Instal dependensi:

```bash
pip install -r requirements.txt
```

Jalankan aplikasi:

```bash
streamlit run app.py
```

## Struktur Excel Input

File input wajib memiliki sheet berikut.

### Master_CPL

Kolom:

- `Kode CPL`
- `Rumusan CPL`
- `GA IABEE`

### Master_IK

Kolom:

- `Kode IK`
- `Rumusan IK`
- `Kode CPL`

### Mapping_CPMK

Kolom:

- `Semester`
- `Kode MK`
- `Mata Kuliah`
- `Kode CPMK`
- `Rumusan CPMK`
- `Kode IK`
- `Kode CPL`
- `Bobot CPMK`

Kolom `Semester` pada file lama tetap berarti semester mata kuliah. Pada rekap multi-semester, aplikasi menambahkan metadata periode asesmen dan menyimpan semester mata kuliah sebagai `Semester MK`.

### Nilai_CPMK

Kolom:

- `NIM`
- `Nama Mahasiswa`
- `Kode MK`
- `Mata Kuliah`
- `Kode CPMK`
- `Nilai`

### Target

Kolom:

- `Parameter`
- `Nilai`

Parameter yang digunakan:

- `Batas Nilai Minimum`, contoh `70`
- `Target Ketercapaian CPL`, contoh `70`

## Mode Single Semester

Gunakan mode ini untuk alur lama:

1. Pilih `Single Semester` di sidebar.
2. Upload satu file Excel asesmen.
3. Sesuaikan `Batas Nilai Minimum` dan `Target Ketercapaian CPL` bila diperlukan.
4. Baca tab `Dashboard CPL`, `Rekap CPMK`, `Rekap IK`, `Rekap CPL`, dan `CQI`.
5. Unduh laporan `Rekap_CPMK.xlsx`, `Rekap_IK.xlsx`, `Rekap_CPL.xlsx`, `Laporan_CQI.xlsx`, atau `Semua_Rekap_Asesmen_CPL.xlsx`.

Jika tidak ada file yang diupload, aplikasi mencoba membaca `sample_data.xlsx`.

## Mode Multi Semester

Gunakan mode ini untuk membandingkan beberapa periode asesmen:

1. Pilih `Multi Semester` di sidebar.
2. Isi `Nama Periode/Semester`, misalnya `2025 Ganjil`, `2025 Genap`, atau `2026 Ganjil`.
3. Upload file Excel periode tersebut dengan format lama yang sama.
4. Klik `Tambah Periode`.
5. Ulangi untuk periode lain.

Data periode disimpan di `session_state`, sehingga beberapa file dapat dianalisis dalam satu sesi selama halaman belum di-reset.

## Cara Membaca Dashboard

- KPI menunjukkan jumlah CPL `Tercapai`, `Perlu Perhatian`, dan `Belum Tercapai`.
- Bar chart menunjukkan capaian CPL per periode terpilih.
- Radar chart menunjukkan profil capaian CPL.
- Heatmap menunjukkan capaian IK per CPL.
- Status ditentukan dari target:
  - `Tercapai`: capaian >= target.
  - `Perlu Perhatian`: target - 10 <= capaian < target.
  - `Belum Tercapai`: capaian < target - 10.

## Cara Membaca Tren CPL

Tab `Tren CPL` menampilkan line chart capaian CPL antar periode. Gunakan filter CPL untuk melihat satu atau beberapa CPL tertentu.

Tabel tren berisi:

- `Periode`
- `Kode CPL`
- `Rumusan CPL`
- `Capaian Aktual`
- `Target`
- `Status`

Interpretasi sederhana: CPL yang naik menunjukkan tindak lanjut pembelajaran atau asesmen mulai berdampak. CPL yang stagnan atau turun perlu dibahas sebagai prioritas CQI.

## Cara Membaca Tren IK

Tab `Tren IK` menampilkan indikator kinerja dalam CPL terpilih. Gunakan filter `Kode CPL`, lalu lihat IK mana yang paling rendah atau tidak membaik antar semester.

Tabel tren IK berisi:

- `Periode`
- `Kode CPL`
- `Kode IK`
- `Rumusan IK`
- `Capaian IK`
- `Status`

## CQI Tracking

Tab `CQI` menampilkan CPL yang berstatus `Perlu Perhatian` atau `Belum Tercapai`. Pada mode multi-semester, tabel dilengkapi:

- `Periode`
- `Status Awal`
- `Rekomendasi CQI`
- `PIC`
- `Target Waktu`
- `Status Tindak Lanjut`
- `Capaian Semester Berikutnya`
- `Efektivitas Tindak Lanjut`

Logika efektivitas:

- Naik minimal 5%: `Efektif`
- Naik 0 sampai kurang dari 5%: `Perlu Penguatan`
- Turun: `Belum Efektif`
- Belum ada periode berikutnya: `Belum Ada Periode Berikutnya`

## Evaluasi CQI

Tab `Evaluasi CQI` membandingkan CPL bermasalah pada periode awal dengan capaian periode berikutnya.

Tabel evaluasi berisi:

- `Periode Awal`
- `Periode Berikut`
- `Kode CPL`
- `Indikator Lemah`
- `Capaian Awal`
- `Capaian Berikut`
- `Perubahan`
- `Efektivitas`
- `Rekomendasi Lanjutan`

## Download Multi Semester

Mode multi-semester menyediakan:

- `Tren_CPL.xlsx`
- `Tren_IK.xlsx`
- `Evaluasi_CQI.xlsx`
- `Semua_Rekap_Multi_Semester.xlsx`

File `Semua_Rekap_Multi_Semester.xlsx` berisi sheet:

- `Rekap_CPMK_All`
- `Rekap_IK_All`
- `Rekap_CPL_All`
- `Tren_CPL`
- `Tren_IK`
- `CQI_All`
- `Evaluasi_CQI`

## Contoh Interpretasi untuk Rapat Evaluasi Kurikulum IABEE

Misalnya CPL2 pada `2025 Ganjil` berstatus `Belum Tercapai` dengan indikator lemah `IK2.1` dan `IK2.2`. Tim kurikulum dapat menetapkan tindak lanjut berupa penambahan studi kasus troubleshooting, penyesuaian rubrik analisis masalah, dan penguatan praktikum.

Pada `2025 Genap`, tab `Evaluasi CQI` menunjukkan perubahan CPL2. Jika naik lebih dari 5%, tindak lanjut dapat distandarkan ke RPS dan rubrik. Jika hanya naik sedikit, implementasi perlu diperkuat. Jika turun, akar penyebab perlu dianalisis ulang bersama dosen pengampu, koordinator rumpun MK, dan tim kurikulum.
# dashboard_CPL
