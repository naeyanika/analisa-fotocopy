import streamlit as st
import pandas as pd
import re
from io import BytesIO

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
        
        # Tambah nomor urut
        df_fc = df_fc.reset_index(drop=True)
        df_fc.index = df_fc.index + 1
        df_fc.index.name = 'NO.'
        
        # Fungsi untuk menghitung jumlah lembar
        def hitung_lembar(deskripsi):
            angka = re.findall(r'\d+', deskripsi)
            return sum(map(int, angka)) if angka else 0
        
        # Format tanggal
        df_fc['TRANS. DATE'] = pd.to_datetime(df_fc['TRANS. DATE']).dt.strftime('%d/%m/%Y')
        
        # Hitung total lembar dan estimasi biaya
        df_fc['TOTAL LEMBAR'] = df_fc['DESCRIPTION'].apply(hitung_lembar)
        df_fc['ESTIMASI BIAYA'] = df_fc['TOTAL LEMBAR'] * harga_per_lembar
        df_fc['SELISIH'] = df_fc['NOMINAL'] - df_fc['ESTIMASI BIAYA']
        
        # Tampilkan DataFrame dengan kolom yang diinginkan
        kolom_tampil = ['NO.', 'VOUCHER NO.', 'TRANS. DATE', 'DESCRIPTION', 'TOTAL LEMBAR', 
                       'NOMINAL', 'ESTIMASI BIAYA', 'SELISIH']
        
        st.subheader("Rincian Transaksi Fotocopy")
        st.dataframe(df_fc[kolom_tampil])
        
        # Export to Excel
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
                # Get the xlsxwriter workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                
                # Add number format for currency columns
                money_fmt = workbook.add_format({'num_format': '#,##0'})
                
                # Apply the format to the currency columns
                worksheet.set_column('F:H', 15, money_fmt)  # Adjust column width and format
                
                # Auto-adjust columns' width
                for idx, col in enumerate(df.columns):
                    series = df[col]
                    max_len = max(
                        series.astype(str).apply(len).max(),
                        len(str(series.name))
                    ) + 2
                    worksheet.set_column(idx, idx, max_len)
                
            processed_data = output.getvalue()
            return processed_data
        
        # Create download button
        excel_file = to_excel(df_fc[kolom_tampil])
        st.download_button(
            label="📥 Download Excel file",
            data=excel_file,
            file_name='analisa_fotocopy.xlsx',
            mime='application/vnd.ms-excel'
        )
        
    else:
        st.error("Kolom 'DESCRIPTION' atau 'NOMINAL' tidak ditemukan. Periksa file Excel Anda.")
