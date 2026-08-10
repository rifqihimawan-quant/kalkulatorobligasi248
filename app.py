import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io

# Configure the Web App UI
st.set_page_config(page_title="Kalkulator Obligasi", layout="centered")

st.title("📈 Kalkulator Obligasi")
st.write("Aplikasi simulasi web berdasarkan **Kalkulator Obligasi 2.4.8.xls**")
st.markdown("---")

# 1. User Interface for Inputs
jenis_simulasi = st.selectbox(
    "Pilih Jenis Simulasi", 
    ["Simulasi Beli", "Simulasi Jual", "Simulasi Switching"]
)

col1, col2 = st.columns(2)

with col1:
    mata_uang = st.selectbox("Mata Uang", ["IDR", "USD"])
    nominal = st.number_input("Nominal Transaksi", min_value=0.0, value=1000000.0, step=10000.0)
    harga = st.number_input("Harga (%)", min_value=0.0, value=100.0, step=0.1)

with col2:
    tgl_setelmen = st.date_input("Tanggal Setelmen")
    kupon = st.number_input("Tingkat Kupon (%)", min_value=0.0, value=5.0, step=0.1)
    pajak = st.number_input("Pajak (%)", min_value=0.0, value=10.0, step=1.0)

# 2. Mock Calculation Logic (Replace with exact formulas from Excel)
st.markdown("### Hasil Simulasi")

# Simple placeholder math representing bond calculations
nilai_pokok = nominal * (harga / 100)
bunga_berjalan = nominal * (kupon / 100) * (1/12) # Simplified
pajak_nominal = bunga_berjalan * (pajak / 100)
total_penyelesaian = nilai_pokok + bunga_berjalan - pajak_nominal

# 3. Present Results in a Dataframe
hasil_df = pd.DataFrame({
    "Deskripsi": ["Nilai Pokok", "Bunga Berjalan", "Pajak", "Total Penyelesaian"],
    "Nominal": [f"{nilai_pokok:,.2f}", f"{bunga_berjalan:,.2f}", f"-{pajak_nominal:,.2f}", f"{total_penyelesaian:,.2f}"]
})

st.table(hasil_df)

# 4. Function to Export Dataframe as JPG using Matplotlib
def convert_df_to_jpg(df, title):
    fig, ax = plt.subplots(figsize=(6, 3))
    ax.axis('off')
    
    # Create Table Image
    table = ax.table(cellText=df.values, colLabels=df.columns, loc='center', cellLoc='center')
    table.scale(1, 1.5)
    plt.title(title, weight='bold', size=14)
    
    # Save to buffer
    buf = io.BytesIO()
    plt.savefig(buf, format='jpg', bbox_inches='tight', dpi=300)
    buf.seek(0)
    return buf

# 5. Download Button for JPG
st.markdown("---")
st.write("Unduh hasil simulasi ini sebagai gambar (JPG) untuk referensi.")

jpg_buffer = convert_df_to_jpg(hasil_df, f"Hasil {jenis_simulasi}")

st.download_button(
    label="⬇️ Unduh Simulasi (JPG)",
    data=jpg_buffer,
    file_name="simulasi_obligasi.jpg",
    mime="image/jpeg"
)