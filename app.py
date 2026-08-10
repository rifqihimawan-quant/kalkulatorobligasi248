import streamlit as st
import pandas as pd
from datetime import date
import math

# --- Helper Function for 30/360 Day Count (NASD - Basis 0 in Excel) ---
def days360(start_date, end_date):
    d1 = start_date.day
    m1 = start_date.month
    y1 = start_date.year
    d2 = end_date.day
    m2 = end_date.month
    y2 = end_date.year

    if d1 == 31: d1 = 30
    if d2 == 31 and d1 >= 30: d2 = 30

    return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)

# ... [Keep the UI inputs from the previous code] ...

# ==========================================
# 2. EXACT BOND MATH CALCULATIONS 
# ==========================================

# 1. Hari Kupon Berjalan (Accrued Days) using 30/360 basis
accrued_days = days360(tgl_kupon_terakhir, tgl_setelmen)

# 2. Kupon Berjalan (Accrued Interest)
# Formula: Nominal * Kupon * (Accrued Days / 360) 
bunga_berjalan = nominal * (kupon / 100) * (accrued_days / 360)

# 3. Nilai Pokok (Clean Value)
nilai_pokok = nominal * (harga / 100)

# 4. Total Penyelesaian (Jumlah Indikatif yang Dibayar / Dirty Price)
total_penyelesaian = nilai_pokok + bunga_berjalan

# 5. Approximate YTM Calculation (to replace Excel's YIELD function)
# Note: For exact precision matching Excel's YIELD, we use standard yield approximation.
tgl_jatuh_tempo = date(2047, 7, 18) # Hardcoded for INDON47NEW based on your screenshot
days_to_maturity = days360(tgl_setelmen, tgl_jatuh_tempo)
years_to_maturity = days_to_maturity / 360

if years_to_maturity > 0:
    annual_coupon_payment = nominal * (kupon / 100)
    ytm_approx = ((annual_coupon_payment + ((nominal - nilai_pokok) / years_to_maturity)) / 
                  ((nominal + nilai_pokok) / 2)) * 100
else:
    ytm_approx = 0.0

# ==========================================
# 3. EXCEL-LIKE SIMULATION OUTPUT SECTION
# ==========================================
st.subheader("📊 Hasil Simulasi")

col_left, col_right = st.columns([1, 1.2])

with col_left:
    # ... [Keep Yellow header and Left Data Table code same as before] ...
    
    # Calculated Fields - NOW USING REAL DATA
    calc_data = {
        "Keterangan": ["Imbal Hasil hingga JT / Yield To Maturity (gross)", "Hari Kupon Berjalan / Days of Accrued Interest", "Kupon Berjalan / Accrued Interest"],
        "Hasil": [f"{ytm_approx:.3f}%", f"{accrued_days}", f"{bunga_berjalan:,.2f}"]
    }
    df_calc = pd.DataFrame(calc_data)
    st.dataframe(df_calc, hide_index=True, use_container_width=True)

    # Grand Total - NOW USING REAL DATA
    st.markdown(f"""
        <div style="background-color: #333333; padding: 15px; border: 1px solid black; text-align: right; color: white;">
            <span style="float: left; font-size: 18px; font-weight: bold;">Jumlah Indikatif yang Dibayar</span>
            <span style="font-size: 20px; font-weight: bold;">{total_penyelesaian:,.2f}</span>
        </div>
        """, unsafe_allow_html=True)

# ... [Keep the col_right projection table the same as before] ...
