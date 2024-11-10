import os
import pandas as pd

csv_folder = "Q:/_HiSenetl/_PARYA/MKK/MKK_INDIRILEN_DOSYALAR/165/TAKAS"

# Klasordekı tüm dosyaları lısteleyıp okuma, bırlestırme
csv_files = [file for file in os.listdir(csv_folder) if file.endswith('.csv')]
dfs = []
for file in csv_files:
    file_path = os.path.join(csv_folder, file)
    df = pd.read_csv(file_path, encoding='latin-1')
    dfs.append(df)

# Tum dosyaları tek bir dataframe haline getirme
combined_df = pd.concat(dfs, ignore_index=True)
combined_df.columns=["ÜyeKod","Müşteri No","Tanim","Grup","Hesap","Adet","Tutar", "Sözleşme Türü "]

# Hesap Numaralarına gore Adet toplama
summed_df = combined_df.groupby("ÜyeKod")["Hesap"].sum().reset_index()

# TIB Hesaplarını ayrıstırma ve komısyon hesaplama
tib_komisyon_takas = summed_df[summed_df['ÜyeKod'].astype(str).str.len() == 11].copy()
tib_komisyon_takas_son = tib_komisyon_takas[["ÜyeKod", "Hesap"]].copy()
tib_komisyon_takas_son['Hesap'] = tib_komisyon_takas_son['Hesap'] / 365 * 0.0015

# IYM Hesaplarını ayrıstırma ve komısyon hesaplama
iym_komisyon_takas = summed_df[~summed_df['ÜyeKod'].astype(str).str.len().isin([8, 11])].copy()
iym_komisyon_takas_son = iym_komisyon_takas[["ÜyeKod", "Hesap"]].copy()
iym_komisyon_takas_son['Hesap'] = iym_komisyon_takas_son['Hesap'] / 365 * 0.0015

# Olusacak dosyaının yerı
excel_file_path = "Q:/_HiSenetl/_PARYA/MKK/MKK_INDIRILEN_DOSYALAR/165/Takas_Komisyon.xlsx"

# TIB degerlerını dosyaya yazma
with pd.ExcelWriter(excel_file_path) as writer:
    tib_komisyon_takas_son.to_excel(writer, sheet_name='TIB', index=False)

# IYM degerlerını dosyaya yazma
with pd.ExcelWriter(excel_file_path, mode='a') as writer:
    iym_komisyon_takas_son.to_excel(writer, sheet_name='IYM', index=False)

# IYM/TIB ıcın toplam komısyonu hesaplama
tib_total_hesap = tib_komisyon_takas_son['Hesap'].sum()
iym_total_hesap = iym_komisyon_takas_son['Hesap'].sum()

# Ozet sayfasını olusturup toplam komısyonları yazma
summary_df = pd.DataFrame({'Sheet': ['TIB', 'IYM', 'Toplam'], 'Total Hesap': [tib_total_hesap, iym_total_hesap, tib_total_hesap + iym_total_hesap]})

with pd.ExcelWriter(excel_file_path, mode='a') as writer:
    summary_df.to_excel(writer, sheet_name='Özet', index=False)

