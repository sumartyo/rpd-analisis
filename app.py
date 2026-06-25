import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO

# === KONFIGURASI ===
st.set_page_config(
    page_title="Aplikasi RPD Kanwil Ditjenpas Babel",
    page_icon="📊",
    layout="wide"
)

# === INISIALISASI SESSION STATE ===
months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des']

# Data default dari Excel (HANYA 51, 52, 53)
default_excel_data = {
    'rencana': {
        '51': {
            'Jan': 525859698, 'Feb': 578254746, 'Mar': 1085986124,
            'Apr': 563128746, 'Mei': 607667173, 'Jun': 1177149901,
            'Jul': 0, 'Agu': 0, 'Sep': 0, 'Okt': 0, 'Nov': 0, 'Des': 0
        },
        '52': {
            'Jan': 270899999, 'Feb': 256999997, 'Mar': 90000000,
            'Apr': 237110337, 'Mei': 274897332, 'Jun': 139383437,
            'Jul': 0, 'Agu': 0, 'Sep': 0, 'Okt': 0, 'Nov': 0, 'Des': 0
        },
        '53': {
            'Jan': 0, 'Feb': 0, 'Mar': 0,
            'Apr': 0, 'Mei': 0, 'Jun': 0,
            'Jul': 0, 'Agu': 0, 'Sep': 0, 'Okt': 0, 'Nov': 0, 'Des': 0
        }
    },
    'penyerapan': {
        '51': {
            'Jan': 525859698, 'Feb': 566329084, 'Mar': 1138767257,
            'Apr': 563128746, 'Mei': 608391335, 'Jun': 908723396,
            'Jul': 0, 'Agu': 0, 'Sep': 0, 'Okt': 0, 'Nov': 0, 'Des': 0
        },
        '52': {
            'Jan': 270900000, 'Feb': 256745788, 'Mar': 6983892,
            'Apr': 237110337, 'Mei': 274357200, 'Jun': 112084549,
            'Jul': 0, 'Agu': 0, 'Sep': 0, 'Okt': 0, 'Nov': 0, 'Des': 0
        },
        '53': {
            'Jan': 0, 'Feb': 0, 'Mar': 0,
            'Apr': 0, 'Mei': 0, 'Jun': 0,
            'Jul': 0, 'Agu': 0, 'Sep': 0, 'Okt': 0, 'Nov': 0, 'Des': 0
        }
    },
    'pagu': {'51': 7460964000, '52': 2787775000, '53': 0}
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
        data_df = data_df[data_df.iloc[:, 1] == 692669]
        
        rencana_data = {'51': {}, '52': {}, '53': {}}
        penyerapan_data = {'51': {}, '52': {}, '53': {}}
        
        month_map = {
            1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'Mei', 6: 'Jun',
            7: 'Jul', 8: 'Agu', 9: 'Sep', 10: 'Okt', 11: 'Nov', 12: 'Des'
        }
        
        for _, row in data_df.iterrows():
            periode = int(row.iloc[4])
            bulan = month_map.get(periode, '')
            
            if bulan:
                rencana_data['51'][bulan] = int(row.iloc[5]) if pd.notna(row.iloc[5]) else 0
                rencana_data['52'][bulan] = int(row.iloc[6]) if pd.notna(row.iloc[6]) else 0
                rencana_data['53'][bulan] = int(row.iloc[7]) if pd.notna(row.iloc[7]) else 0
                
                penyerapan_data['51'][bulan] = int(row.iloc[9]) if pd.notna(row.iloc[9]) else 0
                penyerapan_data['52'][bulan] = int(row.iloc[10]) if pd.notna(row.iloc[10]) else 0
                penyerapan_data['53'][bulan] = int(row.iloc[11]) if pd.notna(row.iloc[11]) else 0
        
        return rencana_data, penyerapan_data
    
    except Exception as e:
        st.error(f"Error membaca file Excel: {str(e)}")
        return None, None

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
    </style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="main-header">📊 Aplikasi Analisis RPD Kanwil Ditjenpas Babel</h1>', unsafe_allow_html=True)
st.markdown("---")

# === SIDEBAR ===
status_relaksasi_bulan = {}

with st.sidebar:
    st.markdown("KANWIL DITJEN PEMASYARAKATAN BANGKA BELITUNG")
    st.markdown("**KPPN:** 015")
    st.markdown("**Periode:** 2026")
    st.markdown("---")
    
    st.markdown("### 📤 Upload Data Excel")
    uploaded_file = st.file_uploader("Pilih file Excel", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        if st.button("📥 Proses Upload", use_container_width=True):
            with st.spinner("Memproses data..."):
                rencana_data, penyerapan_data = parse_excel_file(uploaded_file)
                
                if rencana_data and penyerapan_data:
                    for akun in ['51', '52', '53']:
                        if akun in rencana_data:
                            for bulan in months:
                                if bulan in rencana_data[akun]:
                                    st.session_state.data['rencana'][akun][bulan] = rencana_data[akun][bulan]
                        if akun in penyerapan_data:
                            for bulan in months:
                                if bulan in penyerapan_data[akun]:
                                    st.session_state.data['penyerapan'][akun][bulan] = penyerapan_data[akun][bulan]
                    
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
        st.session_state.excel_uploaded = False
        st.rerun()

# === MAIN CONTENT ===

# 1. INPUT PAGU DIPA
st.markdown("### 💰 Pagu Utama DIPA")

if st.session_state.excel_uploaded:
    st.info("📊 **Status:** Menggunakan data dari file Excel yang diupload")

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
    
    if rata_rata_kumulatif <= 0.05:
        nilai_ikpa = 1.0
    else:
        nilai_ikpa = max(0.0, 1.0 - rata_rata_kumulatif)
    
    # Simpan nilai terakhir untuk digunakan di luar loop
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

# Tentukan bulan berjalan
bulan_berjalan = 'Jun'
for m in reversed(months):
    if (st.session_state.data['rencana']['51'][m] > 0 or 
        st.session_state.data['rencana']['52'][m] > 0 or 
        st.session_state.data['rencana']['53'][m] > 0):
        bulan_berjalan = m
        break

idx_bulan = months.index(bulan_berjalan) + 1

# Hitung AKUMULASI dari Januari sampai bulan berjalan
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

# === RINGKASAN TOTAL KESELURUHAN ===
st.markdown("#### 📌 Ringkasan Total Keseluruhan")

total_rencana_all = total_rencana_51_akum + total_rencana_52_akum + total_rencana_53_akum
total_penyerapan_all = total_penyerapan_51_akum + total_penyerapan_52_akum + total_penyerapan_53_akum
target_all = target_51_akum + target_52_akum + target_53_akum
selisih_all = total_penyerapan_all - target_all

col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.metric("Total Rencana", format_rp(total_rencana_all))
with col2:
    st.metric("Total Penyerapan", format_rp(total_penyerapan_all))
with col3:
    st.metric("Total Target", format_rp(target_all))
with col4:
    st.metric("Total Selisih", format_rp(selisih_all))
with col5:
    if target_all > 0:
        capaian_total = (total_penyerapan_all / target_all) * 100
        if capaian_total >= 100:
            st.markdown(f'<p style="color: #28a745; font-size: 1.5rem; font-weight: bold;">{capaian_total:.1f}% ✅</p>', unsafe_allow_html=True)
        else:
            st.markdown(f'<p style="color: #dc3545; font-size: 1.5rem; font-weight: bold;">{capaian_total:.1f}% ❌</p>', unsafe_allow_html=True)
    else:
        st.markdown('<p style="color: #dc3545; font-size: 1.5rem; font-weight: bold;">0%</p>', unsafe_allow_html=True)

# === TARGET PER TRIWULAN (DENGAN WARNA MERAH UNTUK CAPAIAN < 100%) ===
st.markdown("---")
st.markdown("#### 📋 Target Penyerapan per Triwulan (Akumulasi)")

# Hitung akumulasi per triwulan
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

# Tampilkan per triwulan dengan expander dan warna
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

# === REKAP TOTAL SEMUA TRIWULAN ===
st.markdown("---")
st.markdown("### 📊 Rekap Total Semua Triwulan (Akumulasi)")

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

pivot_rencana = df_rekap.pivot(index='Triwulan', columns='Jenis Belanja', values='Rencana')
pivot_penyerapan = df_rekap.pivot(index='Triwulan', columns='Jenis Belanja', values='Penyerapan')
pivot_target = df_rekap.pivot(index='Triwulan', columns='Jenis Belanja', values='Target')
pivot_capaian = df_rekap.pivot(index='Triwulan', columns='Jenis Belanja', values='Capaian (%)')

col1, col2 = st.columns(2)

with col1:
    st.subheader("Akumulasi Rencana per Triwulan")
    st.dataframe(pivot_rencana.style.format(lambda x: format_rp(x)), use_container_width=True)
    
    st.subheader("Akumulasi Penyerapan per Triwulan")
    st.dataframe(pivot_penyerapan.style.format(lambda x: format_rp(x)), use_container_width=True)

with col2:
    st.subheader("Target per Triwulan")
    st.dataframe(pivot_target.style.format(lambda x: format_rp(x)), use_container_width=True)
    
    st.subheader("Capaian per Triwulan")
    st.dataframe(
        pivot_capaian.style.format('{:.1f}%')
        .map(lambda x: 'color: #dc3545; font-weight: bold;' if x < 100 else 'color: #28a745; font-weight: bold;')
    ),
    use_container_width=True

# === GRAFIK ===
st.markdown("---")
st.markdown("### 📈 Visualisasi Data")

col1, col2 = st.columns(2)

with col1:
    fig_bar = go.Figure()
    fig_bar.add_trace(go.Bar(
        x=df_simulasi['Bulan'], 
        y=df_simulasi['Dev Seluruh (P)'] * 100,
        name='Deviasi Seluruh (%)',
        marker_color='#ff7f0e',
        text=[f"{x*100:.2f}%" for x in df_simulasi['Dev Seluruh (P)']],
        textposition='auto'
    ))
    fig_bar.update_layout(
        title='Deviasi Seluruh Belanja per Bulan (%)',
        height=400,
        yaxis_title='Deviasi (%)'
    )
    st.plotly_chart(fig_bar, use_container_width=True)

with col2:
    fig_line = go.Figure()
    fig_line.add_trace(go.Scatter(
        x=df_simulasi['Bulan'], 
        y=df_simulasi['Nilai IKPA'],
        name='Nilai IKPA',
        mode='lines+markers',
        line=dict(width=3, color='#1f77b4'),
        marker=dict(size=10)
    ))
    fig_line.add_hline(y=95, line_dash="dash", line_color="green", annotation_text="Threshold 95")
    fig_line.update_layout(
        title='Nilai IKPA per Bulan',
        height=400,
        yaxis_title='Nilai IKPA',
        yaxis_range=[0, 105]
    )
    st.plotly_chart(fig_line, use_container_width=True)

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
            
            df_rekap.to_excel(writer, sheet_name='Rekap Triwulan', index=False)
            
            # Semua variabel sudah didefinisikan di atas
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
                    total_rencana_all, total_penyerapan_all, target_all,
                    selisih_all,
                    f"{(total_penyerapan_all/target_all*100) if target_all > 0 else 0:.1f}%",
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
st.caption(f"© 2026 Aplikasi RPD Kanwil Ditjenpas Babel | Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")