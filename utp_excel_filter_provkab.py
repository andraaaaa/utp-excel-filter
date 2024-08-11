# script KAB version

import pandas as pd
import re, os

rm_prov = [0, 2, 3]
rm_kab = [0, 1, 3, 4, 5]

def format_int(x):
    return f"{x:,.0f}"

def make_sheet_name(x):
    try:
        c = re.search('jenis_komoditas', x)
        if c:
            sn = re.split(r'_', x)
            v = sn[10]
        else:
            sn = re.split(r'_', x)
            v = sn[9]
        v = v.replace('.xlsx', '')
    except IndexError:
        sn = re.split(r'_', x)
        v = sn[6] + '_' + sn[7]
    return v

def make_xlsx_name(x):
    sn = re.split(r'_', x)
    v = sn[6] + '_' + sn[7]
    v = v.replace('.xlsx', '')
    return v

def check_folders(p):
    if os.path.exists(p): os.makedirs(p)
    else: pass 

folder_data = "D:\\Publikasi UTP\\data prov" # Ganti dengan folder data yang akan diinput
folder_outp = "D:\\Publikasi UTP\\filter prov" # Ganti dengan path folder output hasil run filter

def filter_data(inp, outp, set, kode, nama):
    for index, dirs, files in os.walk(inp):
        for a in files:
            print("Reading %s"%(a))
            getsheet = ''
            ex = pd.ExcelFile(inp + "\\" + a)

            # Generate nama file excel dan nama sheet dari nama file
            for q in ex.sheet_names:
                try:
                    if set == 'prov':
                        c = re.search('kab', q)
                        if c: getsheet = q
                    if set == 'kab':
                        c = re.search('kec', q)
                        if c: getsheet = q
                except:
                    print('Marker sheet tidak ada')

            # Filter data berdasarkan sheet
            df = ex.parse(getsheet)
            try:
                if set == 'prov':
                    df6401 = df.loc[df.iloc[:, 2] == kode]
                elif set == 'kab':
                    df6401 = df.loc[df.iloc[:, 4] == kode]
            except: print("Switcher harus 'prov' untuk tabel provinsi dan 'kab' untuk tabel kabupaten")

            # Generate nilai agregat
            kabsum = df6401.sum()
            df6401 = df6401._append(kabsum, ignore_index=True)
            sz = len(df6401)

            # Menambahkan baris agregat ke tabel dan drop kolom yang tidak diperlukan
            if set == 'prov':
                df6401.iloc[sz-1, 1] = nama
                df6401 = df6401.drop(df6401.columns[rm_prov], axis=1)
            elif set == 'kab':
                df6401.iloc[sz-1, 2] = nama
                df6401 = df6401.drop(df6401.columns[rm_kab], axis=1)

            df6401.fillna("–", inplace=True)
            df6401.style.format(formatter='{:.0f}', thousands=".")
            df6401.to_excel("%s\\%s_%s.xlsx"%(outp, make_xlsx_name(a), make_sheet_name(a)), sheet_name=make_sheet_name(a))

# contoh pemakaian :
filter_data(inp=folder_data, outp=folder_outp, set='prov', kode=32, nama='JAWA BARAT')
filter_data(inp=folder_data, outp=folder_outp, set='kab', kode=6401, nama='PASER')

# Expected output : tabel XLSX yang memuat data sejumlah kecamatan terfilter dengan sum dan nama kabupaten di paling bawah
#                   ditambah dengan En Dash untuk missing value, nama file adalah nomor tabel_kode komoditas
#                   untuk nama tabel nasional menjadi nama
#                   (contoh : 8_11_6209.xlsx)
