import streamlit as st
import pandas as pd
import io

# Configurazione della pagina Web
st.set_page_config(page_title="Misano Logistics Master Portal", layout="wide", page_icon="🚢")

st.title("🚢 Misano Group - Master Tracking Portal")
st.markdown("Carica i report Excel di **JAS**, **DSV** e **Geodis** per generare il file unico **MISANO MASTER**.")

st.sidebar.header("📁 Caricamento File Spedizionieri")
file_jas = st.sidebar.file_uploader("Report JAS (.xlsx)", type=["xlsx"])
file_dsv = st.sidebar.file_uploader("Report DSV (.xlsx)", type=["xlsx"])
file_geodis = st.sidebar.file_uploader("Report Geodis (.xlsx)", type=["xlsx"])

# Definizione della struttura Master Misano
MASTER_COLUMNS = [
    'SPEDIZIONIERE', 'BOOKING NUMBER', 'ETD', 'ETA', 'POL', 
    'SUPPLIER', 'BOXES', 'IRIS', 'T.T.', 'INVOICE', 
    'PESO LORDO', 'CRD DATE', 'BOOKING RICEVUTO', 'NOLO', 'TOT. SPEDIZ', 'NOTE'
]

def parse_jas(file):
    if file is None:
        return pd.DataFrame()
    df = pd.read_excel(file)
    master_jas = pd.DataFrame()
    master_jas['SPEDIZIONIERE'] = ['JAS'] * len(df)
    master_jas['BOOKING NUMBER'] = df.get('Booking No', df.get('House Bill', ''))
    master_jas['ETD'] = pd.to_datetime(df.get('Estimated Departure Date', None), errors='coerce').dt.strftime('%Y-%m-%d')
    master_jas['ETA'] = pd.to_datetime(df.get('Estimated Arrival Date', None), errors='coerce').dt.strftime('%Y-%m-%d')
    master_jas['POL'] = df.get('Port of Loading', '')
    master_jas['SUPPLIER'] = df.get('Shipper Name', '')
    master_jas['BOXES'] = df.get('Packages', 0)
    master_jas['INVOICE'] = df.get('Commercial Invoice No', '')
    master_jas['PESO LORDO'] = df.get('Gross Weight (KG)', 0)
    master_jas['CRD DATE'] = pd.to_datetime(df.get('Cargo Received Date', None), errors='coerce').dt.strftime('%Y-%m-%d')
    return master_jas

def parse_dsv(file):
    if file is None:
        return pd.DataFrame()
    df = pd.read_excel(file)
    master_dsv = pd.DataFrame()
    master_dsv['SPEDIZIONIERE'] = ['DSV'] * len(df)
    master_dsv['BOOKING NUMBER'] = df.get('myDSV Booking ID', df.get('House Bill', ''))
    master_dsv['ETD'] = pd.to_datetime(df.get('Estimated departure date', None), errors='coerce').dt.strftime('%Y-%m-%d')
    master_dsv['ETA'] = pd.to_datetime(df.get('Estimated arrival date', None), errors='coerce').dt.strftime('%Y-%m-%d')
    master_dsv['POL'] = df.get('Load Port (First)', df.get('Origin - Name', ''))
    master_dsv['SUPPLIER'] = df.get('Sender Name', '')
    master_dsv['BOXES'] = df.get('Total Colli', 0)
    master_dsv['INVOICE'] = df.get('Invoicing Reference', df.get('Order Reference', ''))
    master_dsv['PESO LORDO'] = df.get('Total Weight', 0)
    master_dsv['CRD DATE'] = pd.to_datetime(df.get('Origin Goods Ready for Shipping date', None), errors='coerce').dt.strftime('%Y-%m-%d')
    return master_dsv

def parse_geodis(file):
    if file is None:
        return pd.DataFrame()
    df = pd.read_excel(file)
    master_geodis = pd.DataFrame()
    master_geodis['SPEDIZIONIERE'] = ['Geodis'] * len(df)
    master_geodis['BOOKING NUMBER'] = df.get('Booking Ref', df.get('House Waybill', ''))
    master_geodis['ETD'] = pd.to_datetime(df.get('ETD', None), errors='coerce').dt.strftime('%Y-%m-%d')
    master_geodis['ETA'] = pd.to_datetime(df.get('ETA', None), errors='coerce').dt.strftime('%Y-%m-%d')
    master_geodis['POL'] = df.get('POL', df.get('Port of Origin', ''))
    master_geodis['SUPPLIER'] = df.get('Shipper', df.get('Supplier', ''))
    master_geodis['BOXES'] = df.get('Packages', 0)
    master_geodis['INVOICE'] = df.get('Invoice No', '')
    master_geodis['PESO LORDO'] = df.get('Gross Weight', 0)
    master_geodis['CRD DATE'] = pd.to_datetime(df.get('CRD', None), errors='coerce').dt.strftime('%Y-%m-%d')
    return master_geodis

if st.button("🔄 Elabora e Unifica i Report", type="primary"):
    df_jas = parse_jas(file_jas)
    df_dsv = parse_dsv(file_dsv)
    df_geodis = parse_geodis(file_geodis)
    
    df_master = pd.concat([df_jas, df_dsv, df_geodis], ignore_index=True)
    
    for col in MASTER_COLUMNS:
        if col not in df_master.columns:
            df_master[col] = ""
            
    df_master = df_master[MASTER_COLUMNS]
    
    st.success(f"✅ Unificazione completata! Totale spedizioni attive aggregate: {len(df_master)}")
    st.dataframe(df_master, use_container_width=True)
    
    # Export Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_master.to_excel(writer, index=False, sheet_name='MISANO MASTER')
    
    st.download_button(
        label="📥 Scarica MISANO MASTER (.xlsx)",
        data=buffer.getvalue(),
        file_name="MISANO_MASTER_CONSOLIDATO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
