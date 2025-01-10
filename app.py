import pandas as pd
import streamlit as st

st.title("Analisa Pembayaran Fotocopy")
st.write("""File yang dibutuhkan ATK.xlsx""")
st.write("""Buat Kolom |No.|VOUCHER NO.|TRANS. DATE|DESCRIPTION|Nominal|Invoice, Kuitansi, Nota|Voucher|KETERANGAN (Kelemahan)|""")
st.write("""Penamaan kolom harus sama persis seperti di atas""")

# Upload file Excel
uploaded_file = st.file_uploader("Unggah file Excel", type=["xlsx"])

# Input harga per lembar fotokopi
harga_per_lembar = st.number_input("Masukkan harga per lembar fotokopi (Rp)", min_value=0, value=300, step=50)

if uploaded_file:
    # Membaca file Excel
    df = pd.read_excel(uploaded_file)

    # Daftar kata kunci untuk mendeteksi transaksi fotokopi
    kata_kunci_fc = ["FC", "fc", "Fotocopy", "fotocopy", "Foto copy", "foto copy", "foto kopi", "fotokopi"]

    # Filter transaksi yang berkaitan dengan fotokopi
    df_fc = df[df['DESCRIPTION'].str.contains('|'.join(kata_kunci_fc), case=False, na=False)]

    # Ekstrak jumlah lembar dari deskripsi
    import re
    def hitung_lembar(desc):
        angka = re.findall(r'\((\d+)\)', desc)
        return sum(map(int, angka)) if angka else 0

    df_fc['Total Lembar'] = df_fc['DESCRIPTION'].apply(hitung_lembar)
    df_fc['Total Biaya'] = df_fc['Total Lembar'] * harga_per_lembar

    # Menampilkan hasil analisis
    st.subheader("Detail Transaksi Fotokopi")
    st.dataframe(df_fc[['VOUCHER NO.', 'TRANS. DATE', 'DESCRIPTION', 'Total Lembar', 'Total Biaya']])

    # Total keseluruhan
    total_lembar = df_fc['Total Lembar'].sum()
    total_biaya = df_fc['Total Biaya'].sum()
    st.write(f"**Total Lembar:** {total_lembar}")
    st.write(f"**Total Biaya:** Rp {total_biaya:,.0f}")

    # Download hasil analisis
    output = io.BytesIO()
    df_fc.to_excel(output, index=False)
    output.seek(0)
    st.download_button(
        label="Unduh Hasil Analisis",
        data=output,
        file_name="analisis_fotocopy.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
