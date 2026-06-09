import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import google.generativeai as genai
import pydeck as pdk
import requests
from io import BytesIO
import re
import numpy as np

# =====================================================
# KONFIGURASI HALAMAN
# =====================================================
st.set_page_config(
    page_title="Dashboard Penjualan TM & TT",
    layout="wide"
)

# =====================================================
# CSS EXECUTIVE DASHBOARD
# =====================================================
st.markdown("""
<style>
[data-testid="stAppViewContainer"] .main .block-container {
    padding-top: 1.4rem;
    padding-left: 2rem;
    padding-right: 2rem;
    max-width: 1500px;
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #f5f7fb 0%, #eef4ff 100%);
}

.hero-sales {
    background: linear-gradient(135deg, #0f172a 0%, #1d4ed8 48%, #16a34a 100%);
    padding: 28px 32px;
    border-radius: 24px;
    color: white;
    box-shadow: 0 12px 32px rgba(15,23,42,0.20);
    margin-bottom: 18px;
}

.hero-title {
    font-size: 2.15rem;
    font-weight: 850;
    margin-bottom: 8px;
}

.hero-subtitle {
    font-size: 1rem;
    opacity: 0.95;
    max-width: 980px;
    line-height: 1.5;
}

.kpi-card {
    padding: 18px 20px;
    border-radius: 18px;
    color: white;
    min-height: 125px;
    box-shadow: 0 10px 26px rgba(0,0,0,0.12);
}

.kpi-label {
    font-size: 0.95rem;
    font-weight: 600;
    opacity: 0.92;
    margin-bottom: 8px;
}

.kpi-main {
    font-size: 1.75rem;
    font-weight: 850;
    line-height: 1.1;
    margin-bottom: 8px;
}

.kpi-caption {
    font-size: 0.82rem;
    opacity: 0.93;
}

.card-navy {
    background: linear-gradient(135deg, #0f172a, #1e3a8a);
}

.card-blue {
    background: linear-gradient(135deg, #2563eb, #06b6d4);
}

.card-green {
    background: linear-gradient(135deg, #16a34a, #22c55e);
}

.card-red {
    background: linear-gradient(135deg, #dc2626, #f97316);
}

.card-purple {
    background: linear-gradient(135deg, #7c3aed, #a855f7);
}

.card-orange {
    background: linear-gradient(135deg, #ea580c, #f59e0b);
}

.info-strip {
    padding: 13px 16px;
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 15px;
    color: #334155;
    margin: 14px 0 16px 0;
}

.section-title {
    font-size: 1.35rem;
    font-weight: 800;
    color: #0f172a;
    margin: 10px 0 14px 0;
}

.small-caption {
    color: #64748b;
    font-size: 0.9rem;
}

div[data-testid="stMetricValue"] {
    font-size: 1.65rem;
}
</style>
""", unsafe_allow_html=True)

# =====================================================
# DEFAULT LINK DATA
# =====================================================
DEFAULT_DATA_URL = "https://ptpln365-my.sharepoint.com/:x:/g/personal/irham_tantowi_ptpln365_onmicrosoft_com/IQAUglusWltYS5lAaGWg8N8UAQcg4oRKsBIujfrMF60IfGQ?e=4zID9r"

# =====================================================
# FUNGSI BANTUAN
# =====================================================
def convert_google_drive_url(url):
    match = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
    if match:
        file_id = match.group(1)
        return f"https://drive.google.com/uc?export=download&id={file_id}"

    match = re.search(r"id=([a-zA-Z0-9_-]+)", url)
    if match:
        file_id = match.group(1)
        return f"https://drive.google.com/uc?export=download&id={file_id}"

    return url


def convert_sharepoint_url(url):
    if "download=1" in url:
        return url
    if "?" in url:
        return url + "&download=1"
    return url + "?download=1"


def read_excel_from_url(url):
    if "drive.google.com" in url:
        url = convert_google_drive_url(url)
    elif "sharepoint.com" in url or "onedrive" in url:
        url = convert_sharepoint_url(url)

    response = requests.get(
        url,
        headers={"User-Agent": "Mozilla/5.0"},
        allow_redirects=True
    )

    if response.status_code != 200:
        raise Exception(f"Gagal download file. Status code: {response.status_code}")

    return pd.read_excel(BytesIO(response.content), sheet_name="TM", header=1)


def kpi_card(label, value, caption, css_class):
    st.markdown(
        f"""
        <div class="kpi-card {css_class}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-main">{value}</div>
            <div class="kpi-caption">{caption}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def safe_numeric(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)


def format_gwh(value):
    try:
        return f"{float(value):,.2f} GWh"
    except:
        return "0.00 GWh"


def format_pct(value):
    try:
        return f"{float(value):,.2f}%"
    except:
        return "0.00%"


def warna_delta(delta):
    return "Positif" if delta >= 0 else "Negatif"


def get_secret_key(key_name):
    try:
        return st.secrets[key_name]
    except:
        return ""


# =====================================================
# SUMBER DATA
# =====================================================
st.sidebar.header("📁 Sumber Data")

file_upload = st.sidebar.file_uploader(
    "Upload Excel jika ingin data baru",
    type=["xlsx"]
)

try:
    if file_upload is not None:
        df = pd.read_excel(file_upload, sheet_name="TM", header=1)
        st.sidebar.success("Sukses file upload manual.")
    else:
        df = read_excel_from_url(DEFAULT_DATA_URL)
        st.sidebar.success("Sukses data default dari link.")
except Exception as e:
    st.error(f"Gagal membaca data: {e}")
    st.stop()

df.columns = df.columns.astype(str).str.strip()

# =====================================================
# PERIODE DATA
# =====================================================
bulan_map = {
    "Januari": "Jan",
    "Februari": "Feb",
    "Maret": "Mar",
    "April": "Apr",
    "Mei": "Mei",
    "Juni": "Jun",
    "Juli": "Jul",
    "Agustus": "Agu",
    "September": "Sep",
    "Oktober": "Okt",
    "November": "Nov",
    "Desember": "Des"
}

urutan_bulan = list(bulan_map.values())

st.sidebar.header("🗓️ Periode Data")

tahun_lalu = st.sidebar.number_input("Tahun Lalu", value=2025, step=1)
tahun_ini = st.sidebar.number_input("Tahun Ini", value=2026, step=1)

mode_periode = st.sidebar.selectbox(
    "Mode Periode",
    ["Kumulatif s.d. Bulan", "Bulanan YoY"]
)

pilih_bulan = st.sidebar.selectbox(
    "Pilih Bulan",
    list(bulan_map.keys()),
    index=3
)

kode_bulan = bulan_map[pilih_bulan]
index_bulan = urutan_bulan.index(kode_bulan)


def hitung_kwh(df, tahun, kode_bulan, mode):
    if mode == "Bulanan YoY":
        kolom_bulanan = f"kWh {kode_bulan} {tahun}"
        if kolom_bulanan in df.columns:
            return safe_numeric(df[kolom_bulanan]), kolom_bulanan
        return pd.Series(0, index=df.index), f"{kolom_bulanan} belum tersedia"

    kolom_kumulatif = f"kWh sd {kode_bulan} {tahun}"
    if kolom_kumulatif in df.columns:
        return safe_numeric(df[kolom_kumulatif]), kolom_kumulatif

    bulan_sampai = urutan_bulan[:index_bulan + 1]
    kolom_bulanan = [
        f"kWh {b} {tahun}"
        for b in bulan_sampai
        if f"kWh {b} {tahun}" in df.columns
    ]

    if len(kolom_bulanan) > 0:
        nilai = df[kolom_bulanan].apply(
            pd.to_numeric, errors="coerce"
        ).fillna(0).sum(axis=1)
        return nilai, " + ".join(kolom_bulanan)

    return pd.Series(0, index=df.index), f"Data {kode_bulan} {tahun} belum tersedia"


nilai_lalu, sumber_lalu = hitung_kwh(df, tahun_lalu, kode_bulan, mode_periode)
nilai_ini, sumber_ini = hitung_kwh(df, tahun_ini, kode_bulan, mode_periode)

st.sidebar.markdown("### 📊 Parameter Data")
st.sidebar.success(f"Mode: {mode_periode}")
st.sidebar.info(f"Sumber {tahun_lalu}: {sumber_lalu}")
st.sidebar.info(f"Sumber {tahun_ini}: {sumber_ini}")

required_cols = ["UP3", "TARIF", "KLUSTER USAHA"]
missing_cols = [col for col in required_cols if col not in df.columns]

if missing_cols:
    st.error(f"Kolom wajib tidak ditemukan: {missing_cols}")
    st.write("Kolom yang terbaca:", list(df.columns))
    st.stop()

df = df.dropna(subset=["UP3", "TARIF", "KLUSTER USAHA"], how="any")

df["UP3"] = df["UP3"].astype(str).str.strip()
df["TARIF"] = df["TARIF"].astype(str).str.strip()
df["KLUSTER USAHA"] = df["KLUSTER USAHA"].astype(str).str.strip()

if "NAMA PELANGGAN" in df.columns:
    df["NAMA PELANGGAN"] = df["NAMA PELANGGAN"].astype(str).str.strip()

if "IDPEL" in df.columns:
    df["IDPEL"] = df["IDPEL"].astype(str).str.strip()

df["kWh Tahun Lalu"] = nilai_lalu
df["kWh Tahun Ini"] = nilai_ini

df["GWh Tahun Lalu"] = df["kWh Tahun Lalu"] / 1_000_000
df["GWh Tahun Ini"] = df["kWh Tahun Ini"] / 1_000_000
df["Delta GWh"] = df["GWh Tahun Ini"] - df["GWh Tahun Lalu"]

df["Growth %"] = df.apply(
    lambda x: (x["Delta GWh"] / x["GWh Tahun Lalu"] * 100)
    if x["GWh Tahun Lalu"] != 0 else 0,
    axis=1
)

df["Status Delta"] = df["Delta GWh"].apply(warna_delta)

# =====================================================
# FILTER DATA
# =====================================================
st.sidebar.header("🔎 Filter Data")

pilih_tarif = st.sidebar.multiselect(
    "Tarif",
    sorted(df["TARIF"].dropna().unique()),
    default=sorted(df["TARIF"].dropna().unique())
)

pilih_up3 = st.sidebar.multiselect(
    "UP3",
    sorted(df["UP3"].dropna().unique()),
    default=sorted(df["UP3"].dropna().unique())
)

pilih_cluster = st.sidebar.multiselect(
    "Cluster KBLI",
    sorted(df["KLUSTER USAHA"].dropna().unique()),
    default=sorted(df["KLUSTER USAHA"].dropna().unique())
)

if "GOLONGAN TARIF" in df.columns:
    df["GOLONGAN TARIF"] = df["GOLONGAN TARIF"].astype(str).str.strip()
    pilih_golongan = st.sidebar.multiselect(
        "Golongan Tarif",
        sorted(df["GOLONGAN TARIF"].dropna().unique()),
        default=sorted(df["GOLONGAN TARIF"].dropna().unique())
    )
else:
    pilih_golongan = []

st.sidebar.header("🔍 Cari Pelanggan")

search_idpel = st.sidebar.text_input(
    "Cari IDPEL / Nama Pelanggan",
    placeholder="Masukkan IDPEL atau nama pelanggan"
)

st.sidebar.header("📊 Filter Visual Cluster")

mode_cluster = st.sidebar.selectbox(
    "Mode Tampilan",
    ["Top saja", "Bottom saja", "Top + Bottom"]
)

top_n_cluster = st.sidebar.slider("Jumlah Top Cluster", 3, 30, 10, 1)
bottom_n_cluster = st.sidebar.slider("Jumlah Bottom Cluster", 3, 30, 10, 1)

metrik_cluster = st.sidebar.selectbox(
    "Metrik Ranking",
    ["GWh Tahun Ini", "Delta GWh", "Growth %"]
)

st.sidebar.header("🎚️ Filter Nilai")

kondisi = st.sidebar.selectbox(
    "Kondisi Filter Nilai",
    ["Semua", "Lebih dari", "Kurang dari", "Antara"]
)

nilai_min = 0.0
nilai_max = 0.0

if kondisi == "Lebih dari":
    nilai_min = st.sidebar.number_input("Nilai lebih dari", value=0.0)
elif kondisi == "Kurang dari":
    nilai_max = st.sidebar.number_input("Nilai kurang dari", value=0.0)
elif kondisi == "Antara":
    nilai_min = st.sidebar.number_input("Nilai minimum", value=0.0)
    nilai_max = st.sidebar.number_input("Nilai maksimum", value=1000.0)

st.sidebar.header("🗺️ Map Sebaran")
tampilkan_map = st.sidebar.checkbox("Tampilkan Map Pelanggan", value=False)

st.sidebar.header("🧠 AI Insight")

api_key = get_secret_key("GOOGLE_API_KEY")

input_pelanggan_ai = st.sidebar.text_input(
    "Pelanggan yang ingin dianalisis",
    placeholder="Isi IDPEL atau nama pelanggan"
)

instruksi_analisis_pasar = st.sidebar.text_area(
    "Instruksi analisis pasar",
    placeholder=(
        "Contoh: Analisis penyebab kenaikan/penurunan pelanggan ini. "
        "Pertimbangkan hari kerja efektif, hari besar agama, karakter musiman produksi dari pabrik, "
        "tren sektor usaha, kondisi pasar, isu sosial, politik, geopolitik, dan teknis non teknis."
    ),
    height=140
)

instruksi_ai_user = st.sidebar.text_area(
    "Instruksi tambahan untuk AI",
    placeholder="Analisis pelanggan yang naik dan turun:",
    height=120
)

# =====================================================
# APPLY FILTER
# =====================================================
df_filter = df[
    (df["TARIF"].isin(pilih_tarif)) &
    (df["UP3"].isin(pilih_up3)) &
    (df["KLUSTER USAHA"].isin(pilih_cluster))
].copy()

if "GOLONGAN TARIF" in df.columns:
    df_filter = df_filter[df_filter["GOLONGAN TARIF"].isin(pilih_golongan)]

if kondisi == "Lebih dari":
    df_filter = df_filter[df_filter[metrik_cluster] > nilai_min]
elif kondisi == "Kurang dari":
    df_filter = df_filter[df_filter[metrik_cluster] < nilai_max]
elif kondisi == "Antara":
    df_filter = df_filter[
        (df_filter[metrik_cluster] >= nilai_min) &
        (df_filter[metrik_cluster] <= nilai_max)
    ]

# =====================================================
# KPI EXECUTIVE
# =====================================================
total_lalu = df_filter["GWh Tahun Lalu"].sum()
total_ini = df_filter["GWh Tahun Ini"].sum()
delta = df_filter["Delta GWh"].sum()
growth = (delta / total_lalu * 100) if total_lalu != 0 else 0

jumlah_pelanggan = len(df_filter)
rata_gwh = total_ini / jumlah_pelanggan if jumlah_pelanggan > 0 else 0

jumlah_naik = len(df_filter[df_filter["Delta GWh"] > 0])
jumlah_turun = len(df_filter[df_filter["Delta GWh"] < 0])

status_delta = "Naik" if delta >= 0 else "Turun"
warna_kpi_delta = "card-green" if delta >= 0 else "card-red"

st.markdown(
    f"""
    <div class="hero-sales">
        <div class="hero-title">⚡ Dashboard Penjualan Kluster B & I UID Jawa Timur</div>
        <div class="hero-subtitle">
            Monitoring penjualan TM & TT berbasis GWh, YoY growth, delta pelanggan, cluster usaha, UP3, tarif, dan detail pelanggan.
            Periode: <b>{mode_periode} {pilih_bulan}</b> | Perbandingan <b>{tahun_lalu} vs {tahun_ini}</b>.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

k1, k2, k3, k4, k5 = st.columns(5)

with k1:
    kpi_card(
        f"Penjualan {tahun_lalu}",
        format_gwh(total_lalu),
        "Baseline pembanding YoY",
        "card-navy"
    )

with k2:
    kpi_card(
        f"Penjualan {tahun_ini}",
        format_gwh(total_ini),
        "Realisasi periode berjalan",
        "card-blue"
    )

with k3:
    kpi_card(
        "Delta",
        format_gwh(delta),
        f"Kondisi: {status_delta}",
        warna_kpi_delta
    )

with k4:
    kpi_card(
        "Growth",
        format_pct(growth),
        "Pertumbuhan dibanding tahun lalu",
        "card-purple"
    )

with k5:
    kpi_card(
        "Jumlah Pelanggan",
        f"{jumlah_pelanggan:,.0f}",
        f"Naik: {jumlah_naik:,} | Turun: {jumlah_turun:,}",
        "card-green"
    )

st.markdown(
    f"""
    <div class="info-strip">
        💡 <b>Insight cepat:</b> total penjualan {tahun_ini} sebesar <b>{total_ini:,.2f} GWh</b>, 
        dengan delta <b>{delta:,.2f} GWh</b> dan growth <b>{growth:,.2f}%</b>. 
        Gunakan filter di sidebar untuk melihat kontribusi per UP3, tarif, cluster usaha, dan pelanggan.
    </div>
    """,
    unsafe_allow_html=True
)

# =====================================================
# DATA REKAP
# =====================================================
cluster_summary = df_filter.groupby(
    "KLUSTER USAHA", as_index=False
).agg({
    "GWh Tahun Ini": "sum",
    "GWh Tahun Lalu": "sum",
    "Delta GWh": "sum"
})

cluster_summary["Growth %"] = cluster_summary.apply(
    lambda x: (x["Delta GWh"] / x["GWh Tahun Lalu"] * 100)
    if x["GWh Tahun Lalu"] != 0 else 0,
    axis=1
)

cluster_summary["Status Delta"] = cluster_summary["Delta GWh"].apply(warna_delta)

up3_summary = df_filter.groupby(
    "UP3", as_index=False
).agg({
    "GWh Tahun Ini": "sum",
    "GWh Tahun Lalu": "sum",
    "Delta GWh": "sum"
})

up3_summary["Growth %"] = up3_summary.apply(
    lambda x: (x["Delta GWh"] / x["GWh Tahun Lalu"] * 100)
    if x["GWh Tahun Lalu"] != 0 else 0,
    axis=1
)

up3_summary["Status Delta"] = up3_summary["Delta GWh"].apply(warna_delta)

tarif_summary = df_filter.groupby(
    "TARIF", as_index=False
).agg({
    "GWh Tahun Ini": "sum",
    "GWh Tahun Lalu": "sum",
    "Delta GWh": "sum"
})

tarif_summary["Growth %"] = tarif_summary.apply(
    lambda x: (x["Delta GWh"] / x["GWh Tahun Lalu"] * 100)
    if x["GWh Tahun Lalu"] != 0 else 0,
    axis=1
)

# =====================================================
# SEARCH IDPEL / PELANGGAN
# =====================================================
if search_idpel:
    st.subheader("🔍 Hasil Pencarian Pelanggan")

    df_search = df_filter.copy()

    if "IDPEL" not in df_search.columns:
        st.warning("Kolom IDPEL tidak ditemukan di data.")
    else:
        df_search["IDPEL"] = df_search["IDPEL"].astype(str)

        if "NAMA PELANGGAN" in df_search.columns:
            df_search["NAMA PELANGGAN"] = df_search["NAMA PELANGGAN"].astype(str)

            hasil = df_search[
                df_search["IDPEL"].str.contains(search_idpel, case=False, na=False) |
                df_search["NAMA PELANGGAN"].str.contains(search_idpel, case=False, na=False)
            ]
        else:
            hasil = df_search[
                df_search["IDPEL"].str.contains(search_idpel, case=False, na=False)
            ]

        if hasil.empty:
            st.info("Data pelanggan tidak ditemukan.")
        else:
            kolom_tampil = [
                "UP3",
                "NAMA PELANGGAN",
                "IDPEL",
                "TARIF",
                "DAYA",
                "KLUSTER USAHA",
                "GWh Tahun Lalu",
                "GWh Tahun Ini",
                "Delta GWh",
                "Growth %"
            ]

            kolom_tampil = [kol for kol in kolom_tampil if kol in hasil.columns]

            st.dataframe(
                hasil[kolom_tampil].sort_values("GWh Tahun Ini", ascending=False),
                use_container_width=True,
                hide_index=True
            )

# =====================================================
# TABS DASHBOARD
# =====================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📈 Executive Overview",
    "🏢 Analisis UP3",
    "🏭 Analisis Cluster",
    "📋 Detail Pelanggan",
    "🗺️ Map & AI"
])

# =====================================================
# TAB 1 EXECUTIVE OVERVIEW
# =====================================================
with tab1:
    st.markdown('<div class="section-title">Executive Overview</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)

    with c1:
        if not tarif_summary.empty:
            fig_tarif = px.bar(
                tarif_summary.sort_values("GWh Tahun Ini", ascending=False),
                x="TARIF",
                y=["GWh Tahun Lalu", "GWh Tahun Ini"],
                barmode="group",
                text_auto=".2f",
                title=f"Perbandingan Penjualan per Tarif - {mode_periode} {pilih_bulan}"
            )
            fig_tarif.update_layout(
                height=430,
                yaxis_title="GWh",
                xaxis_title="Tarif",
                legend_title="Periode",
                margin=dict(t=60, l=10, r=10, b=10)
            )
            st.plotly_chart(fig_tarif, use_container_width=True)

    with c2:
        if not cluster_summary.empty and cluster_summary["GWh Tahun Ini"].sum() > 0:
            donut = cluster_summary.sort_values("GWh Tahun Ini", ascending=False).head(12)

            fig_donut = px.pie(
                donut,
                names="KLUSTER USAHA",
                values="GWh Tahun Ini",
                hole=0.52,
                title=f"Komposisi Top 12 Cluster - {tahun_ini}"
            )
            fig_donut.update_traces(
                textinfo="percent",
                textposition="inside",
                hovertemplate="<b>%{label}</b><br>GWh: %{value:.2f}<br>Share: %{percent}<extra></extra>"
            )
            fig_donut.update_layout(
                height=430,
                legend_title="Cluster",
                margin=dict(t=60, l=10, r=10, b=10)
            )
            st.plotly_chart(fig_donut, use_container_width=True)

    st.markdown('<div class="section-title">Top 10 Pelanggan Berdasarkan Penjualan Tahun Ini</div>', unsafe_allow_html=True)

    nama_col = "NAMA PELANGGAN" if "NAMA PELANGGAN" in df_filter.columns else "IDPEL"

    top_pelanggan = (
        df_filter
        .sort_values("GWh Tahun Ini", ascending=False)
        .head(10)
        .copy()
    )

    if not top_pelanggan.empty:
        fig_top_pelanggan = px.bar(
            top_pelanggan.sort_values("GWh Tahun Ini", ascending=True),
            x="GWh Tahun Ini",
            y=nama_col,
            orientation="h",
            color="Delta GWh",
            color_continuous_scale="RdYlGn",
            text="GWh Tahun Ini",
            title="Top 10 Pelanggan Berdasarkan GWh Tahun Ini"
        )
        fig_top_pelanggan.update_traces(
            texttemplate="%{text:.2f}",
            textposition="outside"
        )
        fig_top_pelanggan.update_layout(
            height=520,
            xaxis_title="GWh Tahun Ini",
            yaxis_title="",
            margin=dict(t=60, l=10, r=30, b=10)
        )
        st.plotly_chart(fig_top_pelanggan, use_container_width=True)

# =====================================================
# TAB 2 ANALISIS UP3
# =====================================================
with tab2:
    st.markdown('<div class="section-title">Analisis UP3</div>', unsafe_allow_html=True)

    c3, c4 = st.columns(2)

    with c3:
        up3_chart = up3_summary.sort_values("Delta GWh", ascending=True).copy()

        if not up3_chart.empty:
            fig_up3_delta = px.bar(
                up3_chart,
                x="Delta GWh",
                y="UP3",
                orientation="h",
                color="Status Delta",
                color_discrete_map={
                    "Positif": "#16a34a",
                    "Negatif": "#dc2626"
                },
                text="Delta GWh",
                title="Delta GWh per UP3"
            )
            fig_up3_delta.update_traces(
                texttemplate="%{text:.2f}",
                textposition="outside"
            )
            fig_up3_delta.update_layout(
                height=max(500, len(up3_chart) * 35),
                xaxis_title="Delta GWh",
                yaxis_title="UP3",
                margin=dict(t=60, l=10, r=30, b=10)
            )
            st.plotly_chart(fig_up3_delta, use_container_width=True)

    with c4:
        up3_growth = up3_summary.sort_values("Growth %", ascending=False).copy()

        if not up3_growth.empty:
            fig_up3_growth = px.scatter(
                up3_growth,
                x="GWh Tahun Ini",
                y="Growth %",
                size="GWh Tahun Ini",
                color="Delta GWh",
                hover_name="UP3",
                color_continuous_scale="RdYlGn",
                title="Matrix UP3: GWh Tahun Ini vs Growth"
            )
            fig_up3_growth.update_layout(
                height=500,
                xaxis_title=f"GWh {tahun_ini}",
                yaxis_title="Growth %",
                margin=dict(t=60, l=10, r=10, b=10)
            )
            st.plotly_chart(fig_up3_growth, use_container_width=True)

    st.markdown('<div class="section-title">Rekap UP3</div>', unsafe_allow_html=True)

    up3_tabel = up3_summary.copy()
    up3_tabel = up3_tabel.sort_values("GWh Tahun Ini", ascending=False).reset_index(drop=True)
    up3_tabel.insert(0, "No", range(1, len(up3_tabel) + 1))

    st.dataframe(
        up3_tabel,
        use_container_width=True,
        hide_index=True
    )

# =====================================================
# TAB 3 ANALISIS CLUSTER
# =====================================================
with tab3:
    st.markdown('<div class="section-title">Top & Bottom Cluster KBLI</div>', unsafe_allow_html=True)

    top_cluster = cluster_summary.sort_values(metrik_cluster, ascending=False).head(top_n_cluster)
    bottom_cluster = cluster_summary.sort_values(metrik_cluster, ascending=True).head(bottom_n_cluster)

    if mode_cluster == "Top saja":
        cluster_tampil = top_cluster.copy()
        cluster_tampil["Kategori"] = "Top"
        judul_cluster = f"Top {top_n_cluster} Cluster berdasarkan {metrik_cluster}"

    elif mode_cluster == "Bottom saja":
        cluster_tampil = bottom_cluster.copy()
        cluster_tampil["Kategori"] = "Bottom"
        judul_cluster = f"Bottom {bottom_n_cluster} Cluster berdasarkan {metrik_cluster}"

    else:
        top_cluster = top_cluster.copy()
        bottom_cluster = bottom_cluster.copy()
        top_cluster["Kategori"] = "Top"
        bottom_cluster["Kategori"] = "Bottom"
        cluster_tampil = pd.concat([top_cluster, bottom_cluster], ignore_index=True)
        cluster_tampil = cluster_tampil.drop_duplicates(subset=["KLUSTER USAHA"])
        judul_cluster = f"Top {top_n_cluster} dan Bottom {bottom_n_cluster} Cluster berdasarkan {metrik_cluster}"

    cluster_tampil = cluster_tampil.sort_values(metrik_cluster, ascending=True)
    cluster_tampil["Warna"] = cluster_tampil[metrik_cluster].apply(
        lambda x: "Positif" if x >= 0 else "Negatif"
    )

    fig_cluster = px.bar(
        cluster_tampil,
        x=metrik_cluster,
        y="KLUSTER USAHA",
        orientation="h",
        text=metrik_cluster,
        color="Warna",
        color_discrete_map={
            "Positif": "#16a34a",
            "Negatif": "#dc2626"
        },
        title=f"{judul_cluster} - {mode_periode} {pilih_bulan}"
    )

    fig_cluster.update_traces(
        texttemplate="%{text:.2f}",
        textposition="outside"
    )

    fig_cluster.update_layout(
        height=max(650, len(cluster_tampil) * 42),
        xaxis_title=metrik_cluster,
        yaxis_title="Cluster Usaha",
        margin=dict(t=60, l=10, r=30, b=10)
    )

    st.plotly_chart(fig_cluster, use_container_width=True)

    st.markdown('<div class="section-title">Heatmap UP3 vs Cluster</div>', unsafe_allow_html=True)

    heatmap_data = df_filter.pivot_table(
        index="UP3",
        columns="KLUSTER USAHA",
        values="Delta GWh",
        aggfunc="sum",
        fill_value=0
    )

    if not heatmap_data.empty:
        top_cols = (
            heatmap_data.abs().sum(axis=0)
            .sort_values(ascending=False)
            .head(15)
            .index
        )

        heatmap_data = heatmap_data[top_cols]

        fig_heatmap = px.imshow(
            heatmap_data,
            aspect="auto",
            color_continuous_scale="RdYlGn",
            title="Heatmap Delta GWh UP3 vs Top Cluster"
        )
        fig_heatmap.update_layout(
            height=650,
            xaxis_title="Cluster Usaha",
            yaxis_title="UP3",
            margin=dict(t=60, l=10, r=10, b=10)
        )
        st.plotly_chart(fig_heatmap, use_container_width=True)

# =====================================================
# TAB 4 DETAIL PELANGGAN
# =====================================================
with tab4:
    st.markdown('<div class="section-title">Detail Pelanggan Terfilter</div>', unsafe_allow_html=True)

    kolom_detail = [
        "No",
        "UP3",
        "NAMA PELANGGAN",
        "IDPEL",
        "TARIF",
        "GOLONGAN TARIF",
        "DAYA",
        "Daya baru (VA)",
        "DAYA BARU (VA)",
        "KLUSTER USAHA",
        "DETAIL KLUSTER USAHA",
        "GWh Tahun Lalu",
        "GWh Tahun Ini",
        "Delta GWh",
        "Growth %",
        "Status Delta"
    ]

    kolom_detail = [kol for kol in kolom_detail if kol in df_filter.columns]

    tabel_detail = df_filter[kolom_detail].copy()

    if "GWh Tahun Ini" in tabel_detail.columns:
        tabel_detail = tabel_detail.sort_values("GWh Tahun Ini", ascending=False)

    tabel_detail = tabel_detail.reset_index(drop=True)

    if "No" in tabel_detail.columns:
        tabel_detail = tabel_detail.drop(columns=["No"])

    tabel_detail.insert(0, "No", range(1, len(tabel_detail) + 1))

    c5, c6 = st.columns(2)

    with c5:
        st.markdown("#### 🔼 Top 10 Pelanggan Naik")
        top_naik = tabel_detail.sort_values("Delta GWh", ascending=False).head(10)
        st.dataframe(top_naik, use_container_width=True, hide_index=True)

    with c6:
        st.markdown("#### 🔽 Top 10 Pelanggan Turun")
        top_turun = tabel_detail.sort_values("Delta GWh", ascending=True).head(10)
        st.dataframe(top_turun, use_container_width=True, hide_index=True)

    st.markdown("#### 📋 Tabel Detail Lengkap")

    search_detail = st.text_input(
        "🔍 Cari di tabel detail",
        placeholder="Ketik nama pelanggan, IDPEL, UP3, tarif, atau cluster"
    )

    tabel_tampil = tabel_detail.copy()

    if search_detail:
        tabel_tampil = tabel_tampil[
            tabel_tampil.astype(str).apply(
                lambda row: row.str.contains(search_detail, case=False, na=False).any(),
                axis=1
            )
        ]

    st.dataframe(
        tabel_tampil,
        use_container_width=True,
        height=560,
        hide_index=True
    )

    csv = tabel_tampil.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="⬇️ Download Data Terfilter CSV",
        data=csv,
        file_name="dashboard_penjualan_tm_tt_filtered.csv",
        mime="text/csv"
    )

    # ANALISIS PASAR MANUAL PER PELANGGAN
    st.markdown("### 🧠 Analisis Pasar Manual per Pelanggan")

    if not input_pelanggan_ai:
        st.info("Isi IDPEL atau nama pelanggan di sidebar untuk membuat analisis pasar.")
    else:
        tabel_ai = tabel_detail.copy()

        if "IDPEL" in tabel_ai.columns:
            tabel_ai["IDPEL"] = tabel_ai["IDPEL"].astype(str)

        if "NAMA PELANGGAN" in tabel_ai.columns:
            tabel_ai["NAMA PELANGGAN"] = tabel_ai["NAMA PELANGGAN"].astype(str)

            pelanggan_ai = tabel_ai[
                tabel_ai["IDPEL"].str.contains(input_pelanggan_ai, case=False, na=False) |
                tabel_ai["NAMA PELANGGAN"].str.contains(input_pelanggan_ai, case=False, na=False)
            ].copy()
        else:
            pelanggan_ai = tabel_ai[
                tabel_ai["IDPEL"].str.contains(input_pelanggan_ai, case=False, na=False)
            ].copy()

        if pelanggan_ai.empty:
            st.warning("Pelanggan yang diminta tidak ditemukan pada data terfilter.")
        else:
            st.write("Pelanggan ditemukan:")
            st.dataframe(
                pelanggan_ai,
                use_container_width=True,
                hide_index=True
            )

            if st.button("🧠 Generate Analisis Pasar Pelanggan Ini"):
                if not api_key:
                    st.warning("API Key belum tersedia di Streamlit Secrets.")
                else:
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel("gemini-1.5-flash")

                    prompt_analisis_pasar = f"""
Anda adalah analis pasar dan analis penjualan tenaga listrik PLN UID Jawa Timur.

TUGAS:
Buat analisis pasar untuk pelanggan yang diminta user berdasarkan data pelanggan berikut.

KONTEKS PERIODE:
- Mode periode: {mode_periode}
- Bulan: {pilih_bulan}
- Perbandingan: {tahun_lalu} vs {tahun_ini}

INSTRUKSI USER:
{instruksi_analisis_pasar if instruksi_analisis_pasar else "Analisis kenaikan/penurunan konsumsi listrik pelanggan dengan mempertimbangkan faktor pasar, sosial, budaya, hari kerja, hari besar, dan karakteristik sektor usaha."}

DATA PELANGGAN:
{pelanggan_ai.to_string(index=False)}

KETENTUAN:
1. Analisis harus spesifik sesuai nama pelanggan, cluster usaha, detail cluster, tarif, daya, Delta GWh, dan Growth %.
2. Jelaskan kemungkinan faktor pasar/sosial/budaya/operasional yang relevan.
3. Pertimbangkan faktor hari kerja efektif, hari besar nasional/keagamaan, cuti bersama, Ramadan/Idulfitri, libur sekolah, pola konsumsi masyarakat, musim produksi industri, dan kondisi sektor usaha.
4. Jangan mengarang fakta spesifik yang tidak tersedia.
5. Jika faktor eksternal belum pasti, gunakan frasa "berpotensi", "indikasi", atau "perlu dikonfirmasi".
6. Berikan rekomendasi tindak lanjut singkat untuk UP3/ULP.
7. Bahasa formal, ringkas, dan cocok untuk bahan paparan manajemen PLN.

OUTPUT:
Buat dalam format:
- Nama Pelanggan:
- Ringkasan Kondisi:
- Analisis Pasar:
- Risiko/Peluang:
- Rekomendasi Tindak Lanjut:
"""

                    with st.spinner("AI sedang membuat analisis pasar pelanggan..."):
                        response = model.generate_content(prompt_analisis_pasar)

                    st.success("Analisis pasar berhasil dibuat.")
                    st.markdown(response.text)

# =====================================================
# TAB 5 MAP & AI
# =====================================================
with tab5:
    st.markdown('<div class="section-title">Map Sebaran dan AI Insight</div>', unsafe_allow_html=True)

    if tampilkan_map:
        st.subheader("🗺️ Map Sebaran Pelanggan / Cluster")

        if "LATITUDE" not in df_filter.columns or "LONGITUDE" not in df_filter.columns:
            st.warning("Kolom LATITUDE dan LONGITUDE belum tersedia di Excel.")
        else:
            map_df = df_filter.copy()

            map_df["LATITUDE"] = pd.to_numeric(map_df["LATITUDE"], errors="coerce")
            map_df["LONGITUDE"] = pd.to_numeric(map_df["LONGITUDE"], errors="coerce")

            map_df = map_df.dropna(subset=["LATITUDE", "LONGITUDE"])

            if map_df.empty:
                st.warning("Data koordinat kosong atau tidak valid.")
            else:
                map_df["Radius"] = map_df["GWh Tahun Ini"].clip(lower=0.05) * 150
                map_df["Warna_Map"] = map_df["Delta GWh"].apply(
                    lambda x: [220, 38, 38, 170] if x < 0 else [22, 163, 74, 170]
                )

                if "NAMA PELANGGAN" not in map_df.columns:
                    map_df["NAMA PELANGGAN"] = "-"

                view_state = pdk.ViewState(
                    latitude=map_df["LATITUDE"].mean(),
                    longitude=map_df["LONGITUDE"].mean(),
                    zoom=7,
                    pitch=0
                )

                scatter_layer = pdk.Layer(
                    "ScatterplotLayer",
                    data=map_df,
                    get_position="[LONGITUDE, LATITUDE]",
                    get_radius="Radius",
                    get_fill_color="Warna_Map",
                    pickable=True,
                    auto_highlight=True
                )

                tooltip = {
                    "html": """
                    <b>{NAMA PELANGGAN}</b><br/>
                    UP3: {UP3}<br/>
                    Tarif: {TARIF}<br/>
                    Cluster: {KLUSTER USAHA}<br/>
                    GWh Tahun Ini: {GWh Tahun Ini}<br/>
                    Delta GWh: {Delta GWh}<br/>
                    Growth: {Growth %}%
                    """,
                    "style": {
                        "backgroundColor": "white",
                        "color": "black"
                    }
                }

                st.pydeck_chart(
                    pdk.Deck(
                        layers=[scatter_layer],
                        initial_view_state=view_state,
                        tooltip=tooltip,
                        map_style="light"
                    ),
                    use_container_width=True
                )

                st.caption("Hijau = naik/positif, merah = turun/negatif. Ukuran bubble mengikuti GWh Tahun Ini.")
    else:
        st.info("Aktifkan checkbox 'Tampilkan Map Pelanggan' di sidebar untuk menampilkan map.")

    st.divider()
    st.subheader("🧠 Generate AI Insight")

    index_pivot = ["UP3", "TARIF", "KLUSTER USAHA"]

    if "GOLONGAN TARIF" in df_filter.columns:
        index_pivot = ["UP3", "TARIF", "GOLONGAN TARIF", "KLUSTER USAHA"]

    pivot = df_filter.pivot_table(
        index=index_pivot,
        values=["GWh Tahun Lalu", "GWh Tahun Ini", "Delta GWh"],
        aggfunc="sum"
    ).reset_index()

    pivot["Growth %"] = pivot.apply(
        lambda x: (x["Delta GWh"] / x["GWh Tahun Lalu"] * 100)
        if x["GWh Tahun Lalu"] != 0 else 0,
        axis=1
    )

    if st.button("🔍 Generate AI Insight"):
        if not api_key:
            st.warning("API Key belum tersedia di Streamlit Secrets.")
        else:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel("gemini-1.5-flash")

            data_ai_top = pivot.sort_values("Delta GWh", ascending=False).head(15)
            data_ai_bottom = pivot.sort_values("Delta GWh", ascending=True).head(15)
            data_ai = pd.concat([data_ai_top, data_ai_bottom], ignore_index=True)

            prompt = f"""
Analisis PLN UID Jawa Timur.

KONTEKS:
- Dashboard Penjualan Tenaga Listrik Tegangan Menengah (TM) dan Tegangan Tinggi (TT)
- Mode periode: {mode_periode}
- Bulan: {pilih_bulan}
- Perbandingan: {tahun_lalu} vs {tahun_ini}

DATA KPI:
- Penjualan {tahun_lalu}: {total_lalu:,.2f} GWh
- Penjualan {tahun_ini}: {total_ini:,.2f} GWh
- Delta: {delta:,.2f} GWh
- Growth: {growth:,.2f}%
- Jumlah Pelanggan: {jumlah_pelanggan:,.0f}

DATA UTAMA:
{data_ai.to_string(index=False)}

INSTRUKSI TAMBAHAN USER:
{instruksi_ai_user if instruksi_ai_user else "Tidak ada instruksi tambahan dari user."}

INSTRUKSI ANALISIS:
1. Fokus pola naik dan turunnya penjualan GWh, Delta GWh, dan Growth.
2. Fokus pada perubahan signifikan, baik positif maupun negatif.
3. Gunakan angka spesifik dalam GWh dan persen.
4. Identifikasi top dan bottom delta GWh.
5. Identifikasi cluster yang naik dan turun dengan mempertimbangkan faktor sosial, budaya, hari kerja, hari besar, dan karakteristik industri.
6. Identifikasi UP3 mana yang naik dan turun.
7. Jangan mengarang fakta eksternal yang tidak tersedia.
8. Jika ada faktor eksternal yang belum pasti, gunakan frasa "berpotensi", "indikasi", atau "perlu dikonfirmasi".

OUTPUT WAJIB:
1. Executive Summary singkat dalam 1 paragraf.
2. Top 5 pendorong kenaikan.
3. Top 5 penyebab penurunan.
4. Analisis per UP3.
5. Analisis per cluster.
6. Rekomendasi tindak lanjut untuk manajemen, UP3, dan PAE.

GAYA BAHASA:
Formal, tajam, ringkas, dan cocok untuk bahan penyampaian ke manajemen PLN.
"""

            with st.spinner("AI sedang menganalisis data..."):
                response = model.generate_content(prompt)

            st.success("AI Insight berhasil dibuat.")
            st.markdown(response.text)
