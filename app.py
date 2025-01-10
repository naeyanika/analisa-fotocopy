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
    if 'DESCRIPTION' in df.columns:
        df_fc = df[df['DESCRIPTION'].str.contains('|'.join(kata_kunci_fc), case=False, na=False)]

        # Fungsi untuk menghitung jumlah lembar
        def hitung_lembar(deskripsi):
            angka = re.findall(r'\d+', deskripsi)
            return sum(map(int, angka)) if angka else 0

        df_fc['TOTAL LEMBAR'] = df_fc['DESCRIPTION'].apply(hitung_lembar)
        df_fc['TOTAL BIAYA'] = df_fc['TOTAL LEMBAR'] * harga_per_lembar

        st.subheader("Rincian Transaksi Fotocopy")

        # Tampilkan DataFrame dengan kolom yang sesuai
        kolom_tampil = ['VOUCHER NO.', 'TRANS. DATE', 'DESCRIPTION', 'TOTAL LEMBAR', 'TOTAL BIAYA']
        kolom_tersedia = [kolom for kolom in kolom_tampil if kolom in df_fc.columns]

        st.dataframe(df_fc[kolom_tersedia])

        # Total Biaya
        total_biaya = df_fc['TOTAL BIAYA'].sum()
        st.success(f"Total Biaya Fotocopy: Rp {total_biaya:,}")
    else:
        st.error("Kolom 'DESCRIPTION' tidak ditemukan. Periksa file Excel Anda.")
