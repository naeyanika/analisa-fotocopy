import streamlit as st
import pandas as pd
import re

st.title("Analisa Pembayaran Fotocopy")
st.write("""File yang dibutuhkan ATK.xlsx""")
st.write("""Buat Kolom |No.|VOUCHER NO.|TRANS. DATE|DESCRIPTION|Nominal|Invoice, Kuitansi, Nota|Voucher|KETERANGAN (Kelemahan)|""")
st.write("""Penamaan kolom harus sama persis seperti di atas""")

# Input harga per lembar fotocopy
harga_per_lembar = st.number_input("Harga per lembar fotocopy (Rp):", min_value=0, value=300)

# Upload file Excel
uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx"])

if uploaded_file:
    # Baca file Excel
    df = pd.read_excel(uploaded_file)
    
    # Tampilkan nama kolom
    st.write("Nama Kolom di File:")
    st.write(df.columns.tolist())
    
    # Standardisasi nama kolom
    df.columns = df.columns.str.strip().str.upper()
    
    # Kata kunci pencarian fotocopy
    kata_kunci_fc = ["FC", "fc", "Fotocopy", "fotocopy", "Foto copy", "foto copy", "foto kopi", "fotokopi"]
    
    # Filter data berdasarkan kata kunci di kolom DESCRIPTION
    if 'DESCRIPTION' in df.columns and 'NOMINAL' in df.columns:
        df_fc = df[df['DESCRIPTION'].str.contains('|'.join(kata_kunci_fc), case=False, na=False)]
        
        # Fungsi untuk menghitung jumlah lembar
        def hitung_lembar(deskripsi):
            angka = re.findall(r'\d+', deskripsi)
            return sum(map(int, angka)) if angka else 0
        
        df_fc['TOTAL LEMBAR'] = df_fc['DESCRIPTION'].apply(hitung_lembar)
        df_fc['ESTIMASI BIAYA'] = df_fc['TOTAL LEMBAR'] * harga_per_lembar
        
        st.subheader("Rincian Transaksi Fotocopy")
        
        # Tampilkan DataFrame dengan kolom yang sesuai
        kolom_tampil = ['VOUCHER NO.', 'TRANS. DATE', 'DESCRIPTION', 'NOMINAL', 'TOTAL LEMBAR', 'ESTIMASI BIAYA']
        kolom_tersedia = [kolom for kolom in kolom_tampil if kolom in df_fc.columns]
        st.dataframe(df_fc[kolom_tersedia])
        
        # Total Biaya Aktual dan Estimasi
        total_biaya_aktual = df_fc['NOMINAL'].sum()
        total_biaya_estimasi = df_fc['ESTIMASI BIAYA'].sum()
        
        col1, col2 = st.columns(2)
        with col1:
            st.success(f"Total Biaya Aktual: Rp {total_biaya_aktual:,.2f}")
        with col2:
            st.info(f"Total Biaya Estimasi: Rp {total_biaya_estimasi:,.2f}")
            
        # Hitung dan tampilkan selisih
        selisih = total_biaya_aktual - total_biaya_estimasi
        if selisih > 0:
            st.warning(f"Biaya Aktual lebih tinggi dari Estimasi: Rp {abs(selisih):,.2f}")
        elif selisih < 0:
            st.warning(f"Biaya Aktual lebih rendah dari Estimasi: Rp {abs(selisih):,.2f}")
        else:
            st.success("Biaya Aktual sama dengan Estimasi")
            
    else:
        st.error("Kolom 'DESCRIPTION' atau 'NOMINAL' tidak ditemukan. Periksa file Excel Anda.")
