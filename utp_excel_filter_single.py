import pandas as pd
import re

rm_prov = [0, 2, 3]
rm_kab = [0, 1, 3, 4, 5]

data_in = "D:\\Publikasi UTP\\Nasional_Tabulasi_UTP_BAB_4_tabel_4_53_komoditas__2133_2139.xlsx"
data_out = "D:\\Publikasi UTP\\filtered.xlsx"

def filter_data(inp, outp, set, kode, nama):
    print("Reading %s"%(inp))
    getsheet = ''
    ex = pd.ExcelFile(inp)
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

    df = ex.parse(getsheet)
    try:
        if set == 'prov':
            df6401 = df.loc[df.iloc[:, 2] == kode]
        elif set == 'kab':
            df6401 = df.loc[df.iloc[:, 4] == kode]
    except: print("Switcher harus 'prov' untuk tabel provinsi dan 'kab' untuk tabel kabupaten")

    kabsum = df6401.sum()
    df6401 = df6401._append(kabsum, ignore_index=True)
    sz = len(df6401)

    if set == 'prov':
        df6401.iloc[sz-1, 1] = nama
        df6401 = df6401.drop(df6401.columns[rm_prov], axis=1)
    elif set == 'kab':
        df6401.iloc[sz-1, 2] = nama
        df6401 = df6401.drop(df6401.columns[rm_kab], axis=1)

    df6401.fillna("–", inplace=True)    
    df6401.to_excel(outp, sheet_name="Nama Sheet") # Edit nama sheet atau hapus sheet_name untuk default

# contoh pemakaian :
filter_data(inp=data_in, outp=data_out, set='prov', kode=32, nama='JAWA BARAT')
filter_data(inp=data_in, outp=data_out, set='kab', kode=6401, nama='PASER')