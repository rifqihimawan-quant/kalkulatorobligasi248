import streamlit as st
import pandas as pd
from datetime import date

# Extracted Bond Database from the Calculation Engine
BOND_DB = {
    'FR0056': {'maturity': '2026-09-15', 'coupon': 0.08375, 'freq': 'Semi Annually'},
    'FR0059': {'maturity': '2027-05-15', 'coupon': 0.07, 'freq': 'Semi Annually'},
    'FR0064': {'maturity': '2028-05-15', 'coupon': 0.06125, 'freq': 'Semi Annually'},
    'FR0071': {'maturity': '2029-03-15', 'coupon': 0.09, 'freq': 'Semi Annually'},
    'FR0078': {'maturity': '2029-05-15', 'coupon': 0.0825, 'freq': 'Semi Annually'},
    'FR0073': {'maturity': '2031-05-15', 'coupon': 0.0875, 'freq': 'Semi Annually'},
    'FR0058': {'maturity': '2032-06-15', 'coupon': 0.0825, 'freq': 'Semi Annually'},
    'FR0074': {'maturity': '2032-08-15', 'coupon': 0.075, 'freq': 'Semi Annually'},
    'FR0065': {'maturity': '2033-05-15', 'coupon': 0.06625, 'freq': 'Semi Annually'},
    'FR0068': {'maturity': '2034-03-15', 'coupon': 0.08375, 'freq': 'Semi Annually'},
    'SR018T3': {'maturity': '2026-03-10', 'coupon': 0.0625, 'freq': 'Monthly'},
    'SR019T3': {'maturity': '2026-09-10', 'coupon': 0.0595, 'freq': 'Monthly'},
    'SR020T3': {'maturity': '2027-03-10', 'coupon': 0.063, 'freq': 'Monthly'},
    'SR021T3': {'maturity': '2027-09-10', 'coupon': 0.0635, 'freq': 'Monthly'},
    'SR020T5': {'maturity': '2029-03-10', 'coupon': 0.064, 'freq': 'Monthly'},
    'ST016T2': {'maturity': '2028-06-10', 'coupon': 0.0605, 'freq': 'Monthly'},
    'ST016T4': {'maturity': '2030-05-10', 'coupon': 0.0625, 'freq': 'Monthly'}
}

def calculate_bond_metrics(bond_name, sett_date, nominal, price):
    bond = BOND_DB[bond_name]
    maturity = pd.to_datetime(bond['maturity']).date()
    coupon_rate = bond['coupon']
    freq = bond['freq']
    
    months_step = 1 if freq == 'Monthly' else 6
    freq_div = 12 if freq == 'Monthly' else 2
    
    # 1. Find the last and next coupon dates
    current = maturity
    payments_remaining = 0
    while current > sett_date:
        payments_remaining += 1
        current = (pd.to_datetime(current) - pd.DateOffset(months=months_step)).date()
        
    last_coupon = current
    next_coupon = (pd.to_datetime(last_coupon) + pd.DateOffset(months=months_step)).date()
    
    # 2. Accrued Interest (Bunga Berjalan) using Actual/Actual method
    accrued_days = (sett_date - last_coupon).days
    period_days = (next_coupon - last_coupon).days
    
    bunga_berjalan = nominal * (coupon_rate / freq_div) * (accrued_days / period_days)
    
    # 3. Total Debet (Principal + Accrued Interest)
    principal_cost = nominal * (price / 100)
    total_debet = principal_cost + bunga_berjalan
    
    # 4. Total Coupons to Maturity
    total_kupon = payments_remaining * (coupon_rate / freq_div) * nominal
    
    return total_debet, bunga_berjalan, total_kupon, last_coupon, next_coupon

st.set_page_config(page_title="Kalkulator Obligasi", layout="centered")
st.title("📈 Kalkulator Simulasi Beli Obligasi")
st.markdown("---")

# Input Fields
col1, col2 = st.columns(2)

with col1:
    bond_selection = st.selectbox("1. Pilih Seri Obligasi", options=list(BOND_DB.keys()))
    tanggal_beli = st.date_input("2. Tanggal Beli", value=date.today())
    tanggal_settlement = st.date_input("3. Tanggal Settlement", value=date.today())

with col2:
    nominal_pembelian = st.number_input("4. Nominal Pembelian (IDR)", min_value=1000000.0, value=100000000.0, step=1000000.0)
    harga_beli = st.number_input("5. Harga Beli (%)", min_value=0.0, max_value=200.0, value=100.0, step=0.05)

st.markdown("---")

# Execution & Output
if tanggal_settlement < tanggal_beli:
    st.error("⚠️ Tanggal Settlement tidak boleh mendahului Tanggal Beli.")
elif tanggal_settlement >= pd.to_datetime(BOND_DB[bond_selection]['maturity']).date():
    st.error("⚠️ Obligasi sudah jatuh tempo pada tanggal settlement tersebut.")
else:
    total_debet, bunga_berjalan, total_kupon, last_coupon, next_coupon = calculate_bond_metrics(
        bond_selection, tanggal_settlement, nominal_pembelian, harga_beli
    )
    
    st.subheader("📋 Ringkasan Transaksi")
    
    res_col1, res_col2 = st.columns(2)
    with res_col1:
        st.metric(label="Total Debet", value=f"Rp {total_debet:,.2f}")
        st.metric(label="Bunga Berjalan (Accrued Interest)", value=f"Rp {bunga_berjalan:,.2f}")
    with res_col2:
        st.metric(label="Total Kupon Hingga Jatuh Tempo", value=f"Rp {total_kupon:,.2f}")
        st.caption(f"*Kupon terakhir: {last_coupon.strftime('%d %b %Y')} | Kupon berikutnya: {next_coupon.strftime('%d %b %Y')}*")