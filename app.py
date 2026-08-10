import streamlit as st
import pandas as pd
from datetime import date

# Configure the Web App UI - Set to "wide" to fit the Excel-like layout
st.set_page_config(page_title="Kalkulator Obligasi", layout="wide")

st.title("📈 Kalkulator Obligasi")
st.markdown("Aplikasi simulasi web berdasarkan **Kalkulator Obligasi 2.4.8.xls**")
st.markdown("---")

# ==========================================
# 1. USER FRIENDLY INPUT SECTION
# ==========================================
st.subheader("📝 Form Input Data")

nama_obligasi = st.text_input("Nama Obligasi (Kode Seri)", value="INDON47NEW")

col_in1, col_in2, col_in3 = st.columns(3)

with col_in1:
    jenis_simulasi = st.selectbox("Pilih Jenis Simulasi", ["Simulasi Beli", "Simulasi Jual", "Simulasi Switching"])
    mata_uang = st.selectbox("Mata Uang", ["USD", "IDR"])
    jenis_transaksi = st.selectbox("Jenis Transaksi", ["Pasar Sekunder", "Pasar Perdana"])
    
with col_in2:
    tgl_transaksi = st.date_input("Tanggal Transaksi", value=date(2026, 8, 10))
    tgl_setelmen = st.date_input("Tanggal Setelmen", value=date(2026, 8, 12))
    tgl_kupon_terakhir = st.date_input("Tanggal Kupon Terakhir", value=date(2026, 7, 18))

with col_in3:
    nominal = st.number_input("Nilai Nominal", min_value=0.0, value=60000.0, step=1000.0)
    harga = st.number_input("Harga Nasabah Beli (%)", min_value=0.0, value=84.1500, step=0.01, format="%.4f")
    kupon = st.number_input("Kupon (%)", min_value=0.0, value=4.750, step=0.1, format="%.3f")

# Placeholder button to run simulation
st.button("Hitung Simulasi")
st.markdown("---")

# ==========================================
# 2. EXCEL-LIKE SIMULATION OUTPUT SECTION
# ==========================================
st.subheader("📊 Hasil Simulasi")

# Create two main columns to mirror the Excel layout
col_left, col_right = st.columns([1, 1.2])

with col_left:
    # Yellow Header for Bond Name
    st.markdown(f"""
        <div style="background-color: yellow; padding: 10px; border: 2px solid black; text-align: left;">
            <h3 style="color: black; margin: 0; font-weight: bold;">{nama_obligasi.upper()}</h3>
        </div>
        """, unsafe_allow_html=True)
    
    st.write("") # Spacer

    # Left Data Table
    data_kiri = {
        "Data Produk Obligasi (NASABAH BELI)": [
            "Mata Uang", "Jenis Transaksi", "Tanggal Transaksi", "Tanggal Setelmen", 
            "Tanggal Kupon Terakhir", "Tanggal Kupon Berikutnya", "Tanggal Jatuh Tempo (JT)", 
            "Kupon", "Nilai Nominal", "Harga Nasabah Beli"
        ],
        "Nilai": [
            mata_uang, jenis_transaksi, tgl_transaksi.strftime('%d-%b-%y'), 
            tgl_setelmen.strftime('%d-%b-%y'), tgl_kupon_terakhir.strftime('%d-%b-%y'), 
            "18-Jan-27", "18-Jul-47", f"{kupon:.3f}%", f"{nominal:,.0f}", f"{harga:.4f}%"
        ]
    }
    df_kiri = pd.DataFrame(data_kiri)
    st.dataframe(df_kiri, hide_index=True, use_container_width=True)

    # Calculated Fields
    calc_data = {
        "Keterangan": ["Imbal Hasil hingga JT / Yield To Maturity (gross)", "Hari Kupon Berjalan / Days of Accrued Interest", "Kupon Berjalan / Accrued Interest"],
        "Hasil": ["6.101%", "24", "190.00"]
    }
    df_calc = pd.DataFrame(calc_data)
    st.dataframe(df_calc, hide_index=True, use_container_width=True)

    # Grand Total 
    st.markdown("""
        <div style="background-color: #333333; padding: 15px; border: 1px solid black; text-align: right; color: white;">
            <span style="float: left; font-size: 18px; font-weight: bold;">Jumlah Indikatif yang Dibayar</span>
            <span style="font-size: 20px; font-weight: bold;">50,680.00</span>
        </div>
        """, unsafe_allow_html=True)

with col_right:
    # Header
    st.markdown("""
        <div style="background-color: #333333; padding: 10px; border: 1px solid black; text-align: center; color: white;">
            <h4 style="margin: 0; font-size: 16px;">Proyeksi Pendapatan yang Diterima Nasabah (nett)</h4>
        </div>
        """, unsafe_allow_html=True)
    
    # Right Data Table (Proyeksi)
    data_kanan = {
        "Keterangan": [
            "Nominal yang dibayar pada tanggal :", 
            "Kupon (tidak termasuk pajak capital gain/loss) :",
            "1", "2", "3", "4", "5", "6", "7", "...", "42",
            "Kupon yang diterima hingga JT1 :",
            "Pengembalian hingga JT2 :",
            "Nominal yang diterima saat JT3 :",
            "• Pengembalian pokok",
            "• Kupon saat jatuh tempo (gross)",
            "• Capital gain/loss",
            "• Total Pajak",
            "  ○ Pajak Kupon",
            "  ○ Pajak capital gain/loss",
            "Total keuntungan/kerugian4 :",
            "• Nominal kumulatif",
            "• Persentase kumulatif"
        ],
        "Tanggal": [
            tgl_setelmen.strftime('%d-%b-%y'), "", 
            "18-Jan-27", "18-Jul-27", "18-Jan-28", "18-Jul-28", "18-Jan-29", "18-Jul-29", "18-Jan-30", "", "18-Jul-47",
            "", "", "", "", "", "", "", "", "", "", "", ""
        ],
        "Nilai": [
            "(50,680.00)", "", 
            "1,425.00", "1,425.00", "1,425.00", "1,425.00", "1,425.00", "1,425.00", "1,425.00", "...", "1,425.00",
            "59,660.00", "119,660.00", "61,425.00", "60,000.00", "1,425.00", "9,510.00", "0.00", "0.00", "0.00", 
            "", "69,170.00", "137.00%"
        ]
    }
    df_kanan = pd.DataFrame(data_kanan)
    
    # We use st.table here because it renders raw data without the interactive scrolling dataframe view, 
    # making it look much closer to an Excel printout.
    st.table(df_kanan)

# Footer Notes
st.markdown("---")
st.markdown("""
**Keterangan :**
* Kolom yang di-highlight kuning **WAJIB** diisi dengan data terkini.
1. Kupon yang diterima hingga JT = total kupon (net, tidak termasuk pajak capital gain/loss) - accrued kupon beli
2. Pengembalian hingga JT = nominal + total kupon (net) yang diterima hingga JT - pajak capital gain/loss

**DISCLAIMER :**
* Kalkulator ini disediakan hanya sebagai alat bantu simulasi Obligasi dan tidak dimaksudkan untuk menyediakan rekomendasi.
* Simulasi, harga dan YTM Obligasi yang ditampilkan hanya bersifat indikatif.
""")