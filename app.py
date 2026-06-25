import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO
import base64
from PIL import Image

# === KONFIGURASI ===
st.set_page_config(
    page_title="Aplikasi RPD Kanwil Ditjenpas Babel",
    page_icon="📊",
    layout="wide"
)

# === INISIALISASI SESSION STATE ===
months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des']

# Data default (kosong)
default_excel_data = {
    'rencana': {
        '51': {month: 0 for month in months},
        '52': {month: 0 for month in months},
        '53': {month: 0 for month in months}
    },
    'penyerapan': {
        '51': {month: 0 for month in months},
        '52': {month: 0 for month in months},
        '53': {month: 0 for month in months}
    },
    'pagu': {'51': 0, '52': 0, '53': 0}
}

if 'data' not in st.session_state:
    st.session_state.data = {
        'pagu': default_excel_data['pagu'].copy(),
        'rencana': {
            '51': default_excel_data['rencana']['51'].copy(),
            '52': default_excel_data['rencana']['52'].copy(),
            '53': default_excel_data['rencana']['53'].copy()
        },
        'penyerapan': {
            '51': default_excel_data['penyerapan']['51'].copy(),
            '52': default_excel_data['penyerapan']['52'].copy(),
            '53': default_excel_data['penyerapan']['53'].copy()
        }
    }

if 'excel_uploaded' not in st.session_state:
    st.session_state.excel_uploaded = False

if 'info_satker' not in st.session_state:
    st.session_state.info_satker = {
        'kode': '-',
        'nama': '-',
        'kppn': '-'
    }

# === FUNGSI UTILITAS ===
def format_rp(value):
    return f"Rp {value:,.0f}".replace(",", ".")

def parse_excel_file(uploaded_file):
    """Parse file Excel dan ekstrak data rencana & penyerapan (HANYA 51, 52, 53)"""
    try:
        df = pd.read_excel(uploaded_file, sheet_name='Data', header=2)
        
        data_start = 0
        for idx, row in df.iterrows():
            if str(row.iloc[0]).strip() == '1':
                data_start = idx
                break
        
        data_df = df.iloc[data_start:].reset_index(drop=True)
        
        rencana_data = {
            '51': {m: 0 for m in months},
            '52': {m: 0 for m in months},
            '53': {m: 0 for m in months}
        }
        
        penyerapan_data = {
            '51': {m: 0 for m in months},
            '52': {m: 0 for m in months},
            '53': {m: 0 for m in months}
        }
        
        month_map = {
            1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'Mei', 6: 'Jun',
            7: 'Jul', 8: 'Agu', 9: 'Sep', 10: 'Okt', 11: 'Nov', 12: 'Des'
        }
        
        info_satker = {
            'kode': '-',
            'nama': '-',
            'kppn': '-'
        }
        
        if not data_df.empty:
            try:
                if pd.notna(data_df.iloc[0, 1]):
                    info_satker['kode'] = str(data_df.iloc[0, 1]).strip()
                if pd.notna(data_df.iloc[0, 2]):
                    info_satker['nama'] = str(data_df.iloc[0, 2]).strip()
                if pd.notna(data_df.iloc[0, 3]):
                    info_satker['kppn'] = str(data_df.iloc[0, 3]).strip()
            except:
                pass
        
        data_loaded = 0
        for _, row in data_df.iterrows():
            try:
                periode_val = row.iloc[4]
                if pd.isna(periode_val):
                    continue
                
                if isinstance(periode_val, str):
                    periode = int(periode_val.strip())
                else:
                    periode = int(periode_val)
                
                bulan = month_map.get(periode, '')
                
                if bulan:
                    rencana_data['51'][bulan] = int(row.iloc[5]) if pd.notna(row.iloc[5]) else 0
                    rencana_data['52'][bulan] = int(row.iloc[6]) if pd.notna(row.iloc[6]) else 0
                    rencana_data['53'][bulan] = int(row.iloc[7]) if pd.notna(row.iloc[7]) else 0
                    
                    penyerapan_data['51'][bulan] = int(row.iloc[9]) if pd.notna(row.iloc[9]) else 0
                    penyerapan_data['52'][bulan] = int(row.iloc[10]) if pd.notna(row.iloc[10]) else 0
                    penyerapan_data['53'][bulan] = int(row.iloc[11]) if pd.notna(row.iloc[11]) else 0
                    
                    data_loaded += 1
            except Exception as e:
                continue
        
        st.success(f"✅ Berhasil memuat {data_loaded} baris data")
        
        st.subheader(f"📊 Data yang berhasil di-load (Satker: {info_satker['kode']})")
        data_preview = []
        for bulan in months:
            data_preview.append({
                'Bulan': bulan,
                'Rencana 51': rencana_data['51'][bulan],
                'Rencana 52': rencana_data['52'][bulan],
                'Rencana 53': rencana_data['53'][bulan],
                'Penyerapan 51': penyerapan_data['51'][bulan],
                'Penyerapan 52': penyerapan_data['52'][bulan],
                'Penyerapan 53': penyerapan_data['53'][bulan],
            })
        df_preview = pd.DataFrame(data_preview)
        st.dataframe(df_preview)
        
        return rencana_data, penyerapan_data, info_satker
    
    except Exception as e:
        st.error(f"❌ Error membaca file Excel: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None, None, None

# === FUNGSI TARGET TRIWULAN ===
def get_target_triwulan(pagu, triwulan, jenis):
    target_persen = {
        'I': {'51': 0.10, '52': 0.15, '53': 0.10},
        'II': {'51': 0.40, '52': 0.50, '53': 0.40},
        'III': {'51': 0.60, '52': 0.65, '53': 0.60},
        'IV': {'51': 1.00, '52': 1.00, '53': 1.00}
    }
    persen = target_persen.get(triwulan, {}).get(jenis, 0)
    return pagu * persen

# === FUNGSI MENENTUKAN BULAN BERJALAN ===
def get_bulan_berjalan():
    for m in reversed(months):
        if (st.session_state.data['rencana']['51'][m] > 0 or 
            st.session_state.data['rencana']['52'][m] > 0 or 
            st.session_state.data['rencana']['53'][m] > 0 or
            st.session_state.data['penyerapan']['51'][m] > 0 or 
            st.session_state.data['penyerapan']['52'][m] > 0 or 
            st.session_state.data['penyerapan']['53'][m] > 0):
            return m
    return 'Des'

# === CSS ===
st.markdown("""
    <style>
    .main-header { 
        font-size: 2rem; 
        color: #1f77b4; 
        text-align: center; 
        margin-bottom: 1rem; 
    }
    .stProgress > div > div > div > div {
        display: none !important;
    }
    .logo-container {
        display: flex;
        justify-content: center;
        margin-bottom: 1rem;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="main-header">📊 Aplikasi Analisis RPD Kanwil Ditjenpas Babel</h1>', unsafe_allow_html=True)
st.markdown("---")

# === SIDEBAR ===
status_relaksasi_bulan = {}

with st.sidebar:
    # === LOGO ===
    st.markdown('<div class="logo-container">', unsafe_allow_html=True)
    try:
        # Coba load logo dari file
        logo = Image.open('FA_Logo_Kementrian_Imigrasi_dan_Pemasyarakatan (1).png')
        st.image(logo, width=150)
    except:
        # Jika file tidak ditemukan, tampilkan teks
        st.markdown("### 🏛️ KEMENTERIAN IMIGRASI DAN PEMASYARAKATAN")
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("### 🏢 Informasi Satker")
    st.markdown(f"**Kode Satker:** {st.session_state.info_satker['kode']}")
    st.markdown(f"**Nama Satker:** {st.session_state.info_satker['nama']}")
    st.markdown(f"**KPPN:** {st.session_state.info_satker['kppn']}")
    st.markdown("**Periode:** 2026")
    st.markdown("---")
    
    st.markdown("### 📤 Upload Data Excel")
    uploaded_file = st.file_uploader("Pilih file Excel", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        if st.button("📥 Proses Upload", use_container_width=True):
            with st.spinner("Memproses data..."):
                rencana_data, penyerapan_data, info_satker = parse_excel_file(uploaded_file)
                
                if rencana_data and penyerapan_data:
                    for akun in ['51', '52', '53']:
                        for bulan in months:
                            st.session_state.data['rencana'][akun][bulan] = 0
                            st.session_state.data['penyerapan'][akun][bulan] = 0
                    
                    for akun in ['51', '52', '53']:
                        if akun in rencana_data:
                            for bulan in months:
                                if bulan in rencana_data[akun]:
                                    st.session_state.data['rencana'][akun][bulan] = rencana_data[akun][bulan]
                        if akun in penyerapan_data:
                            for bulan in months:
                                if bulan in penyerapan_data[akun]:
                                    st.session_state.data['penyerapan'][akun][bulan] = penyerapan_data[akun][bulan]
                    
                    if info_satker:
                        st.session_state.info_satker = info_satker
                    
                    st.session_state.excel_uploaded = True
                    st.success("✅ Data berhasil diupload dari Excel!")
                    st.rerun()
    
    if st.session_state.excel_uploaded:
        st.info("📊 Menggunakan data Excel")
        if st.button("🔄 Kembali ke Default", use_container_width=True):
            st.session_state.data = {
                'pagu': default_excel_data['pagu'].copy(),
                'rencana': {
                    '51': default_excel_data['rencana']['51'].copy(),
                    '52': default_excel_data['rencana']['52'].copy(),
                    '53': default_excel_data['rencana']['53'].copy()
                },
                'penyerapan': {
                    '51': default_excel_data['penyerapan']['51'].copy(),
                    '52': default_excel_data['penyerapan']['52'].copy(),
                    '53': default_excel_data['penyerapan']['53'].copy()
                }
            }
            st.session_state.info_satker = {
                'kode': '-',
                'nama': '-',
                'kppn': '-'
            }
            st.session_state.excel_uploaded = False
            st.rerun()
    
    st.markdown("---")
    st.markdown("### ⚙️ Opsi Relaksasi")
    for m in months:
        is_default_relaksasi = m in ['Jan', 'Feb', 'Mar']
        status_relaksasi_bulan[m] = st.checkbox(f"Relaksasi {m}", value=is_default_relaksasi, key=f"chk_{m}")
    
    st.markdown("---")
    if st.button("🧹 Reset Semua Data", use_container_width=True):
        st.session_state.data = {
            'pagu': {'51': 0, '52': 0, '53': 0},
            'rencana': {'51': {month: 0 for month in months}, '52': {month: 0 for month in months}, '53': {month: 0 for month in months}},
            'penyerapan': {'51': {month: 0 for month in months}, '52': {month: 0 for month in months}, '53': {month: 0 for month in months}}
        }
        st.session_state.info_satker = {
            'kode': '-',
            'nama': '-',
            'kppn': '-'
        }
        st.session_state.excel_uploaded = False
        st.rerun()

# === MAIN CONTENT ===

# 1. INPUT PAGU DIPA
st.markdown("### 💰 Pagu Utama DIPA")

if st.session_state.excel_uploaded:
    st.info(f"📊 **Status:** Menggunakan data dari file Excel yang diupload (Satker: {st.session_state.info_satker['kode']})")

col1, col2, col3 = st.columns(3)

with col1:
    pagu51 = st.number_input("Pagu 51 (Belanja Pegawai)", min_value=0, value=st.session_state.data['pagu']['51'], step=1000000, format="%d")
    st.session_state.data['pagu']['51'] = pagu51

with col2:
    pagu52 = st.number_input("Pagu 52 (Belanja Barang)", min_value=0, value=st.session_state.data['pagu']['52'], step=1000000, format="%d")
    st.session_state.data['pagu']['52'] = pagu52

with col3:
    pagu53 = st.number_input("Pagu 53 (Belanja Modal)", min_value=0, value=st.session_state.data['pagu']['53'], step=1000000, format="%d")
    st.session_state.data['pagu']['53'] = pagu53

total_pagu = pagu51 + pagu52 + pagu53
st.info(f"**Total Pagu DIPA Komposit:** {format_rp(total_pagu)}")

proporsi_51 = pagu51 / total_pagu if total_pagu > 0 else 0
proporsi_52 = pagu52 / total_pagu if total_pagu > 0 else 0
proporsi_53 = pagu53 / total_pagu if total_pagu > 0 else 0

st.markdown("---")

# 2. TABEL INPUT DATA BULANAN (EDITABLE)
st.markdown("### 📅 Input Rencana & Penyerapan Bulanan")
st.caption("💡 Klik sel untuk mengedit nilai. Data dari Excel akan otomatis terisi setelah upload.")

edit_data = []
for m in months:
    edit_data.append({
        'Bulan': m,
        'Rencana 51': st.session_state.data['rencana']['51'][m],
        'Penyerapan 51': st.session_state.data['penyerapan']['51'][m],
        'Rencana 52': st.session_state.data['rencana']['52'][m],
        'Penyerapan 52': st.session_state.data['penyerapan']['52'][m],
        'Rencana 53': st.session_state.data['rencana']['53'][m],
        'Penyerapan 53': st.session_state.data['penyerapan']['53'][m],
    })

df_editor = pd.DataFrame(edit_data)

df_edited = st.data_editor(
    df_editor,
    use_container_width=True,
    hide_index=True,
    column_config={
        'Bulan': st.column_config.TextColumn('Bulan', width='small'),
        'Rencana 51': st.column_config.NumberColumn('Rencana 51', format='%d', step=100000),
        'Penyerapan 51': st.column_config.NumberColumn('Penyerapan 51', format='%d', step=100000),
        'Rencana 52': st.column_config.NumberColumn('Rencana 52', format='%d', step=100000),
        'Penyerapan 52': st.column_config.NumberColumn('Penyerapan 52', format='%d', step=100000),
        'Rencana 53': st.column_config.NumberColumn('Rencana 53', format='%d', step=100000),
        'Penyerapan 53': st.column_config.NumberColumn('Penyerapan 53', format='%d', step=100000),
    },
    num_rows="fixed"
)

# Update data dari editor ke session state
for _, row in df_edited.iterrows():
    bulan = row['Bulan']
    if bulan in months:
        st.session_state.data['rencana']['51'][bulan] = row['Rencana 51']
        st.session_state.data['penyerapan']['51'][bulan] = row['Penyerapan 51']
        st.session_state.data['rencana']['52'][bulan] = row['Rencana 52']
        st.session_state.data['penyerapan']['52'][bulan] = row['Penyerapan 52']
        st.session_state.data['rencana']['53'][bulan] = row['Rencana 53']
        st.session_state.data['penyerapan']['53'][bulan] = row['Penyerapan 53']

st.markdown("---")

# 3. PROSES SIMULASI PERHITUNGAN (SESUAI EXCEL)
st.markdown("### 📋 Tabel Analisis Deviasi & Nilai IKPA")
st.caption("Perhitungan sesuai dengan formula pada Detail Indikator Halaman 3 DIPA")

rows = []
kumulatif_dev_seluruh = 0
dev_seluruh_terakhir = 0
rata_rata_kumulatif_terakhir = 0
nilai_ikpa_terakhir = 0

# Simpan data deviasi rata-rata kumulatif untuk grafik
deviasi_rata_rata_kumulatif = []

for idx, m in enumerate(months):
    r51 = st.session_state.data['rencana']['51'][m]
    s51 = st.session_state.data['penyerapan']['51'][m]
    r52 = st.session_state.data['rencana']['52'][m]
    s52 = st.session_state.data['penyerapan']['52'][m]
    r53 = st.session_state.data['rencana']['53'][m]
    s53 = st.session_state.data['penyerapan']['53'][m]
    
    dev51 = s51 - r51
    dev52 = s52 - r52
    dev53 = s53 - r53
    
    pct_dev51 = (abs(dev51) / r51) if r51 > 0 else 0
    pct_dev52 = (abs(dev52) / r52) if r52 > 0 else 0
    pct_dev53 = (abs(dev53) / r53) if r53 > 0 else 0
    
    tertimbang_51 = pct_dev51 * proporsi_51
    tertimbang_52 = pct_dev52 * proporsi_52
    tertimbang_53 = pct_dev53 * proporsi_53
    
    dev_seluruh = tertimbang_51 + tertimbang_52 + tertimbang_53
    
    if status_relaksasi_bulan.get(m, False):
        dev_seluruh = 0.0
        status_bulan = "🔵 Relaksasi"
    else:
        status_bulan = "🟢 Normal"
    
    kumulatif_dev_seluruh += dev_seluruh
    rata_rata_kumulatif = kumulatif_dev_seluruh / (idx + 1)
    
    # Simpan untuk grafik
    deviasi_rata_rata_kumulatif.append(rata_rata_kumulatif * 100)
    
    if rata_rata_kumulatif <= 0.05:
        nilai_ikpa = 1.0
    else:
        nilai_ikpa = max(0.0, 1.0 - rata_rata_kumulatif)
    
    dev_seluruh_terakhir = dev_seluruh
    rata_rata_kumulatif_terakhir = rata_rata_kumulatif
    nilai_ikpa_terakhir = nilai_ikpa
    
    rows.append({
        'No': idx + 1,
        'Bulan': m,
        'Status': status_bulan,
        'Rencana 51': r51, 'Penyerapan 51': s51, 'Deviasi 51': dev51,
        'Rencana 52': r52, 'Penyerapan 52': s52, 'Deviasi 52': dev52,
        'Rencana 53': r53, 'Penyerapan 53': s53, 'Deviasi 53': dev53,
        '% Deviasi 51': pct_dev51, '% Deviasi 52': pct_dev52, '% Deviasi 53': pct_dev53,
        'Proporsi 51': proporsi_51, 'Proporsi 52': proporsi_52, 'Proporsi 53': proporsi_53,
        'Tertimbang 51': tertimbang_51, 'Tertimbang 52': tertimbang_52, 'Tertimbang 53': tertimbang_53,
        'Dev Seluruh (P)': dev_seluruh,
        'Rata-Rata Kumulatif (Q)': rata_rata_kumulatif,
        'Nilai IKPA': nilai_ikpa * 100
    })

df_simulasi = pd.DataFrame(rows)

# Tampilkan tabel hasil
df_display = df_simulasi[[
    'No', 'Bulan', 'Status',
    'Rencana 51', 'Penyerapan 51', 'Deviasi 51',
    'Rencana 52', 'Penyerapan 52', 'Deviasi 52',
    'Rencana 53', 'Penyerapan 53', 'Deviasi 53',
    '% Deviasi 51', '% Deviasi 52', '% Deviasi 53',
    'Proporsi 51', 'Proporsi 52', 'Proporsi 53',
    'Tertimbang 51', 'Tertimbang 52', 'Tertimbang 53',
    'Dev Seluruh (P)', 'Rata-Rata Kumulatif (Q)', 'Nilai IKPA'
]]

st.dataframe(
    df_display.style.format({
        'Rencana 51': lambda x: format_rp(x),
        'Penyerapan 51': lambda x: format_rp(x),
        'Deviasi 51': lambda x: format_rp(x),
        'Rencana 52': lambda x: format_rp(x),
        'Penyerapan 52': lambda x: format_rp(x),
        'Deviasi 52': lambda x: format_rp(x),
        'Rencana 53': lambda x: format_rp(x),
        'Penyerapan 53': lambda x: format_rp(x),
        'Deviasi 53': lambda x: format_rp(x),
        '% Deviasi 51': lambda x: f"{x*100:.2f}%",
        '% Deviasi 52': lambda x: f"{x*100:.2f}%",
        '% Deviasi 53': lambda x: f"{x*100:.2f}%",
        'Proporsi 51': lambda x: f"{x*100:.2f}%",
        'Proporsi 52': lambda x: f"{x*100:.2f}%",
        'Proporsi 53': lambda x: f"{x*100:.2f}%",
        'Tertimbang 51': lambda x: f"{x*100:.2f}%",
        'Tertimbang 52': lambda x: f"{x*100:.2f}%",
        'Tertimbang 53': lambda x: f"{x*100:.2f}%",
        'Dev Seluruh (P)': lambda x: f"{x*100:.2f}%",
        'Rata-Rata Kumulatif (Q)': lambda x: f"{x*100:.2f}%",
        'Nilai IKPA': '{:.2f}'
    }),
    use_container_width=True, 
    hide_index=True
)

# === 4. RINGKASAN HASIL (AKUMULASI JANUARI - SAMPAI BULAN BERJALAN) ===
st.markdown("---")
st.markdown("### 📌 Ringkasan Hasil dan Target Penyerahan")

bulan_berjalan = get_bulan_berjalan()
idx_bulan = months.index(bulan_berjalan) + 1

total_rencana_51_akum = sum([st.session_state.data['rencana']['51'][m] for m in months[:idx_bulan]])
total_rencana_52_akum = sum([st.session_state.data['rencana']['52'][m] for m in months[:idx_bulan]])
total_rencana_53_akum = sum([st.session_state.data['rencana']['53'][m] for m in months[:idx_bulan]])

total_penyerapan_51_akum = sum([st.session_state.data['penyerapan']['51'][m] for m in months[:idx_bulan]])
total_penyerapan_52_akum = sum([st.session_state.data['penyerapan']['52'][m] for m in months[:idx_bulan]])
total_penyerapan_53_akum = sum([st.session_state.data['penyerapan']['53'][m] for m in months[:idx_bulan]])

tw_map = {'Jan':'I', 'Feb':'I', 'Mar':'I', 'Apr':'II', 'Mei':'II', 'Jun':'II',
          'Jul':'III', 'Agu':'III', 'Sep':'III', 'Okt':'IV', 'Nov':'IV', 'Des':'IV'}
tw_berjalan = tw_map[bulan_berjalan]

st.info(f"📌 **Periode Analisis:** Januari s.d. {bulan_berjalan} (Triwulan {tw_berjalan})")

st.markdown("#### 📊 Akumulasi Penyerapan (Januari - {})".format(bulan_berjalan))

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("**Belanja Pegawai (51)**")
    st.metric("Akumulasi Rencana", format_rp(total_rencana_51_akum))
    st.metric("Akumulasi Penyerapan", format_rp(total_penyerapan_51_akum))
    
    target_51_akum = get_target_triwulan(pagu51, tw_berjalan, '51')
    st.metric("Target (Triwulan {})".format(tw_berjalan), format_rp(target_51_akum))
    
    selisih_51 = total_penyerapan_51_akum - target_51_akum
    if target_51_akum > 0:
        capaian_51 = (total_penyerapan_51_akum / target_51_akum) * 100
        if capaian_51 >= 100:
            st.markdown(f'<p style="color: #28a745; font-weight: bold;">Capaian: {capaian_51:.1f}% ✅</p>', unsafe_allow_html=True)
        else:
            st.markdown(f'<p style="color: #dc3545; font-weight: bold;">Capaian: {capaian_51:.1f}% ❌</p>', unsafe_allow_html=True)
        st.caption(f"Selisih: {format_rp(selisih_51)}")
    else:
        st.markdown('<p style="color: #dc3545; font-weight: bold;">Capaian: 0%</p>', unsafe_allow_html=True)

with col2:
    st.markdown("**Belanja Barang (52)**")
    st.metric("Akumulasi Rencana", format_rp(total_rencana_52_akum))
    st.metric("Akumulasi Penyerapan", format_rp(total_penyerapan_52_akum))
    
    target_52_akum = get_target_triwulan(pagu52, tw_berjalan, '52')
    st.metric("Target (Triwulan {})".format(tw_berjalan), format_rp(target_52_akum))
    
    selisih_52 = total_penyerapan_52_akum - target_52_akum
    if target_52_akum > 0:
        capaian_52 = (total_penyerapan_52_akum / target_52_akum) * 100
        if capaian_52 >= 100:
            st.markdown(f'<p style="color: #28a745; font-weight: bold;">Capaian: {capaian_52:.1f}% ✅</p>', unsafe_allow_html=True)
        else:
            st.markdown(f'<p style="color: #dc3545; font-weight: bold;">Capaian: {capaian_52:.1f}% ❌</p>', unsafe_allow_html=True)
        st.caption(f"Selisih: {format_rp(selisih_52)}")
    else:
        st.markdown('<p style="color: #dc3545; font-weight: bold;">Capaian: 0%</p>', unsafe_allow_html=True)

with col3:
    st.markdown("**Belanja Modal (53)**")
    st.metric("Akumulasi Rencana", format_rp(total_rencana_53_akum))
    st.metric("Akumulasi Penyerapan", format_rp(total_penyerapan_53_akum))
    
    target_53_akum = get_target_triwulan(pagu53, tw_berjalan, '53')
    st.metric("Target (Triwulan {})".format(tw_berjalan), format_rp(target_53_akum))
    
    selisih_53 = total_penyerapan_53_akum - target_53_akum
    if target_53_akum > 0:
        capaian_53 = (total_penyerapan_53_akum / target_53_akum) * 100
        if capaian_53 >= 100:
            st.markdown(f'<p style="color: #28a745; font-weight: bold;">Capaian: {capaian_53:.1f}% ✅</p>', unsafe_allow_html=True)
        else:
            st.markdown(f'<p style="color: #dc3545; font-weight: bold;">Capaian: {capaian_53:.1f}% ❌</p>', unsafe_allow_html=True)
        st.caption(f"Selisih: {format_rp(selisih_53)}")
    else:
        st.markdown('<p style="color: #dc3545; font-weight: bold;">Capaian: 0%</p>', unsafe_allow_html=True)

st.markdown("---")

# === TARGET PER TRIWULAN ===
st.markdown("#### 📋 Target Penyerapan per Triwulan (Akumulasi)")

triwulan_data = {'I': {}, 'II': {}, 'III': {}, 'IV': {}}
for tw in ['I', 'II', 'III', 'IV']:
    triwulan_data[tw] = {
        'rencana': {'51': 0, '52': 0, '53': 0},
        'penyerapan': {'51': 0, '52': 0, '53': 0},
        'target': {'51': 0, '52': 0, '53': 0}
    }

month_to_tw = {'I': 3, 'II': 6, 'III': 9, 'IV': 12}
for tw in ['I', 'II', 'III', 'IV']:
    idx_akhir = month_to_tw[tw]
    for m in months[:idx_akhir]:
        for akun in ['51', '52', '53']:
            triwulan_data[tw]['rencana'][akun] += st.session_state.data['rencana'][akun][m]
            triwulan_data[tw]['penyerapan'][akun] += st.session_state.data['penyerapan'][akun][m]

for tw in ['I', 'II', 'III', 'IV']:
    for akun in ['51', '52', '53']:
        pagu = st.session_state.data['pagu'][akun]
        target = get_target_triwulan(pagu, tw, akun)
        triwulan_data[tw]['target'][akun] = target

for tw in ['I', 'II', 'III', 'IV']:
    bulan_akhir = {'I': 'Mar', 'II': 'Jun', 'III': 'Sep', 'IV': 'Des'}[tw]
    with st.expander(f"Triwulan {tw} (Jan - {bulan_akhir})"):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("**Belanja Pegawai (51)**")
            rencana_tw = triwulan_data[tw]['rencana']['51']
            serap_tw = triwulan_data[tw]['penyerapan']['51']
            target_tw = triwulan_data[tw]['target']['51']
            selisih_tw = serap_tw - target_tw
            
            st.metric("Akumulasi Rencana", format_rp(rencana_tw))
            st.metric("Akumulasi Penyerapan", format_rp(serap_tw))
            st.metric("Target", format_rp(target_tw))
            
            if target_tw > 0:
                capaian = (serap_tw / target_tw) * 100
                if capaian >= 100:
                    st.markdown(f'<p style="color: #28a745; font-weight: bold;">Capaian: {capaian:.1f}% ✅</p>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<p style="color: #dc3545; font-weight: bold;">Capaian: {capaian:.1f}% ❌</p>', unsafe_allow_html=True)
                st.caption(f"Selisih: {format_rp(selisih_tw)}")
            else:
                st.markdown('<p style="color: #dc3545; font-weight: bold;">Capaian: 0%</p>', unsafe_allow_html=True)
        
        with col2:
            st.markdown("**Belanja Barang (52)**")
            rencana_tw = triwulan_data[tw]['rencana']['52']
            serap_tw = triwulan_data[tw]['penyerapan']['52']
            target_tw = triwulan_data[tw]['target']['52']
            selisih_tw = serap_tw - target_tw
            
            st.metric("Akumulasi Rencana", format_rp(rencana_tw))
            st.metric("Akumulasi Penyerapan", format_rp(serap_tw))
            st.metric("Target", format_rp(target_tw))
            
            if target_tw > 0:
                capaian = (serap_tw / target_tw) * 100
                if capaian >= 100:
                    st.markdown(f'<p style="color: #28a745; font-weight: bold;">Capaian: {capaian:.1f}% ✅</p>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<p style="color: #dc3545; font-weight: bold;">Capaian: {capaian:.1f}% ❌</p>', unsafe_allow_html=True)
                st.caption(f"Selisih: {format_rp(selisih_tw)}")
            else:
                st.markdown('<p style="color: #dc3545; font-weight: bold;">Capaian: 0%</p>', unsafe_allow_html=True)
        
        with col3:
            st.markdown("**Belanja Modal (53)**")
            rencana_tw = triwulan_data[tw]['rencana']['53']
            serap_tw = triwulan_data[tw]['penyerapan']['53']
            target_tw = triwulan_data[tw]['target']['53']
            selisih_tw = serap_tw - target_tw
            
            st.metric("Akumulasi Rencana", format_rp(rencana_tw))
            st.metric("Akumulasi Penyerapan", format_rp(serap_tw))
            st.metric("Target", format_rp(target_tw))
            
            if target_tw > 0:
                capaian = (serap_tw / target_tw) * 100
                if capaian >= 100:
                    st.markdown(f'<p style="color: #28a745; font-weight: bold;">Capaian: {capaian:.1f}% ✅</p>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<p style="color: #dc3545; font-weight: bold;">Capaian: {capaian:.1f}% ❌</p>', unsafe_allow_html=True)
                st.caption(f"Selisih: {format_rp(selisih_tw)}")
            else:
                st.markdown('<p style="color: #dc3545; font-weight: bold;">Capaian: 0%</p>', unsafe_allow_html=True)
        
        total_target_tw = sum(triwulan_data[tw]['target'].values())
        total_serap_tw = sum(triwulan_data[tw]['penyerapan'].values())
        if total_target_tw > 0:
            progress_tw = min(total_serap_tw / total_target_tw, 1)
            st.progress(progress_tw)
            capaian_tw = progress_tw * 100
            if capaian_tw >= 100:
                st.markdown(f'<p style="color: #28a745; font-weight: bold;">Total Capaian Triwulan {tw}: {capaian_tw:.1f}% ✅</p>', unsafe_allow_html=True)
            else:
                st.markdown(f'<p style="color: #dc3545; font-weight: bold;">Total Capaian Triwulan {tw}: {capaian_tw:.1f}% ❌</p>', unsafe_allow_html=True)

# === GRAFIK ===
st.markdown("---")
st.markdown("### 📈 Visualisasi Deviasi Rata-Rata Kumulatif per Bulan")

# Buat grafik deviasi rata-rata kumulatif
fig_deviasi = go.Figure()

fig_deviasi.add_trace(go.Scatter(
    x=months,
    y=deviasi_rata_rata_kumulatif,
    name='Deviasi Rata-rata Kumulatif (%)',
    mode='lines+markers',
    line=dict(width=3, color='#ff7f0e'),
    marker=dict(size=10)
))

# Tambahkan garis threshold 5%
fig_deviasi.add_hline(y=5, line_dash="dash", line_color="red", annotation_text="Threshold 5%")

fig_deviasi.update_layout(
    title='Deviasi Rata-rata Kumulatif per Bulan (%)',
    height=400,
    yaxis_title='Deviasi (%)',
    xaxis_title='Bulan'
)

st.plotly_chart(fig_deviasi, use_container_width=True)

# === DOWNLOAD ===
st.markdown("---")
st.markdown("### 💾 Download Data")

col1, col2 = st.columns(2)

with col1:
    csv = df_simulasi.to_csv(index=False)
    st.download_button(
        label="📥 Download CSV",
        data=csv,
        file_name=f"RPD_Analisis_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv",
        use_container_width=True
    )

with col2:
    with BytesIO() as output:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_simulasi.to_excel(writer, sheet_name='Analisis RPD', index=False)
            
            monthly_data = []
            for m in months:
                monthly_data.append({
                    'Bulan': m,
                    'Rencana 51': st.session_state.data['rencana']['51'][m],
                    'Penyerapan 51': st.session_state.data['penyerapan']['51'][m],
                    'Rencana 52': st.session_state.data['rencana']['52'][m],
                    'Penyerapan 52': st.session_state.data['penyerapan']['52'][m],
                    'Rencana 53': st.session_state.data['rencana']['53'][m],
                    'Penyerapan 53': st.session_state.data['penyerapan']['53'][m]
                })
            df_monthly = pd.DataFrame(monthly_data)
            df_monthly.to_excel(writer, sheet_name='Data Bulanan', index=False)
            
            # Data per triwulan
            rekap_data = []
            for tw in ['I', 'II', 'III', 'IV']:
                for akun in ['51', '52', '53']:
                    rekap_data.append({
                        'Triwulan': tw,
                        'Jenis Belanja': f"{akun}",
                        'Rencana': triwulan_data[tw]['rencana'][akun],
                        'Penyerapan': triwulan_data[tw]['penyerapan'][akun],
                        'Target': triwulan_data[tw]['target'][akun],
                        'Selisih': triwulan_data[tw]['penyerapan'][akun] - triwulan_data[tw]['target'][akun],
                        'Capaian (%)': (triwulan_data[tw]['penyerapan'][akun] / triwulan_data[tw]['target'][akun] * 100) if triwulan_data[tw]['target'][akun] > 0 else 0
                    })
            df_rekap = pd.DataFrame(rekap_data)
            df_rekap.to_excel(writer, sheet_name='Rekap Triwulan', index=False)
            
            summary = pd.DataFrame({
                'Metrik': [
                    'Periode Analisis',
                    'Total Pagu',
                    'Akumulasi Rencana (51)', 'Akumulasi Penyerapan (51)', 'Target (51)',
                    'Akumulasi Rencana (52)', 'Akumulasi Penyerapan (52)', 'Target (52)',
                    'Akumulasi Rencana (53)', 'Akumulasi Penyerapan (53)', 'Target (53)',
                    'Total Rencana', 'Total Penyerapan', 'Total Target',
                    'Selisih Total', 'Capaian Total',
                    'Deviasi Seluruh (P)', 'Rata-rata Deviasi Kumulatif', 'Nilai IKPA'
                ],
                'Nilai': [
                    f"Jan - {bulan_berjalan} (Triwulan {tw_berjalan})",
                    total_pagu,
                    total_rencana_51_akum, total_penyerapan_51_akum, target_51_akum,
                    total_rencana_52_akum, total_penyerapan_52_akum, target_52_akum,
                    total_rencana_53_akum, total_penyerapan_53_akum, target_53_akum,
                    total_rencana_51_akum + total_rencana_52_akum + total_rencana_53_akum,
                    total_penyerapan_51_akum + total_penyerapan_52_akum + total_penyerapan_53_akum,
                    target_51_akum + target_52_akum + target_53_akum,
                    (total_penyerapan_51_akum + total_penyerapan_52_akum + total_penyerapan_53_akum) - (target_51_akum + target_52_akum + target_53_akum),
                    f"{((total_penyerapan_51_akum + total_penyerapan_52_akum + total_penyerapan_53_akum) / (target_51_akum + target_52_akum + target_53_akum) * 100) if (target_51_akum + target_52_akum + target_53_akum) > 0 else 0:.1f}%",
                    f"{dev_seluruh_terakhir * 100:.2f}%",
                    f"{rata_rata_kumulatif_terakhir * 100:.2f}%",
                    f"{nilai_ikpa_terakhir * 100:.2f}"
                ]
            })
            summary.to_excel(writer, sheet_name='Ringkasan', index=False)
        
        excel_data = output.getvalue()
        st.download_button(
            label="📥 Download Excel",
            data=excel_data,
            file_name=f"RPD_Analisis_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# Footer
st.markdown("---")
st.caption(f"© 2026 Aplikasi RPD Kanwil Ditjenpas Babel | Bulan Berjalan: {bulan_berjalan} | Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")