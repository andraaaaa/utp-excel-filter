**[!!] Update 18 Agustus** : Reformatter sudah terintegrasi dengan fungsi inti

## Python Code Filter dan Merger Data Excel untuk Publikasi UTP
Source Code untuk memisahkan file sesuai filter provinsi atau kabupaten dalam rangka penyusunan Publikasi ST2023 UTP Tahap II.

**Ekspektasi keluaran fungsi :** Membuat file berformat XLSX sesuai dengan filter provinsi dan kabupaten/kota yang diinginkan beserta angka agregat dan format angkanya.

**Minimum Requirements**
- Python versi >= 3.9
- pandas 2.2.2
- openpyxl 3.1.5
- xlrd 2.0.1
- pip v24.2

One-time line install :
`pip install "pandas==2.2.2" "openpyxl==3.1.5" "xlrd==2.0.1"`
  
**Catatan Atas**
- Nama file yang digunakan sebagai input (nonsingle) :

  `Nasional_Tabulasi_UTP_BAB_[nomorbab]_tabel_[nomor_tabel_dgn_underscore]_komoditas_[kode_komoditas]`
- Code ini masih dalam tahap pengembangan, untuk fungsi dasar sudah berjalan sebagaimana mestinya.

### **Persiapan**
Sebelum melakukan filtering, perlu diperhatikan beberapa step awal :

1. Menggabungkan **SELURUH** data yang akan difilter (boleh dengan subfolder) menjadi satu folder
2. Membuat folder baru untuk data hasil filter
3. Menginstall package yang diperlukan

### **Fungsi-Fungsi**

Fungsi **make_sheet_name()** men-generate nama sheet yang akan dimasukkan ke Excel

Fungsi **make_xlsx_name()** men-generate nama file Excel hasil filter

Fungsi **filter_data()** akan memfilter sesuai masukan yang diinginkan

Fungsi **merge_data()** menggabungkan seluruh file excel menjadi sheet dalam satu file baru

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

Contoh output tabel provinsi :

![img](https://i.ibb.co.com/0tKrVNw/Screenshot-2024-08-18-203510.png)

Contoh output tabel kabupaten :

![img](https://i.ibb.co.com/mvmXWBR/Screenshot-2024-08-18-204018-n.png)

_Catatan : File dan sheet name sudah otomatis mengambil dari nama tabel mentah secara default, sehingga pengguna tidak perlu mengotak-atik nama tabel mentah yang akan difilter._

### Data Merging (Optional)
Untuk kemudahan melakukan copy paste data dengan format yang sudah sesuai dan manajemen file lebih efisien, dapat dilakukan penggabungan seluruh file menjadi satu sheet dalam satu file Excel.

Contoh output merging data :

![img](https://i.ibb.co.com/KqV6kzn/merger.png)

Contoh penggunaan sudah terlampir dalam kode sumber.
