## Python Code Filter Data Excel untuk Publikasi UTP
Source Code untuk memisahkan file sesuai filter provinsi atau kabupaten dalam rangka penyusunan Publikasi ST2023 UTP Tahap II.

**Ekspektasi keluaran fungsi :** Membuat file berformat XLSX sesuai dengan filter provinsi dan kabupaten/kota yang diinginkan beserta agregatnya.

**Catatan Atas**
- Harap dijalankan dengan bahasa Python dengan versi >= 3.6
- Code ini masih dalam tahap pengembangan, untuk fungsi dasar sudah berjalan sebagaimana mestinya.

### **Persiapan**
Sebelum melakukan filtering, perlu diperhatikan beberapa step awal :

1. Menggabungkan **SELURUH** data yang akan difilter (boleh dengan subfolder) menjadi satu folder
2. Membuat folder baru untuk data hasil filter
3. Menginstall package pandas (jika sudah ada, silakan upgrade ke versi terbaru)

### **Fungsi-Fungsi**

Fungsi **make_sheet_name()** men-generate nama sheet yang akan dimasukkan ke Excel

Fungsi **make_xlsx_name()** men-generate nama file Excel hasil filter

Fungsi **filter_data()** akan memfilter sesuai masukan yang diinginkan

---
**Parameter untuk filter_data():**

| Parameter | Rincian |
|-----------|---------|
| inp | path untuk membaca folder data mentah |
| outp | path untuk menyimpan folder hasil filter |
| set | 'prov' untuk filter tingkat provinsi dan 'kab' untuk kabupaten |
| kode | kode wilayah |
| nama | nama wilayah |
---

Contoh keluaran nama file :
- Untuk tabel per komoditas dari tabel 8.11 komoditas 6209, berubah nama menjadi 8_11_6209.xlsx

Contoh output tabel :

![img](https://i.ibb.co.com/s9HPhwG/Screenshot-2024-08-11-081757.png)

_Catatan : File dan sheet name sudah otomatis mengambil dari nama tabel mentah secara default, sehingga pengguna tidak perlu mengotak-atik nama tabel mentah yang akan difilter._
