import streamlit as st
import pandas as pd
import plotly.express as px
import google.generativeai as genai
import pydeck as pdk

st.set_page_config(page_title="Dashboard Penjualan TM", layout="wide")

st.title("⚡ Dashboard Penjualan kluster B & I UID Jawa Timur")

import requests
from io import BytesIO
import re

# =========================
# FUNGSI BACA FILE DARI LINK
# =========================
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

    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    response = requests.get(url, headers=headers, allow_redirects=True)

    if response.status_code != 200:
        raise Exception(f"Gagal download file. Status code: {response.status_code}")

    return pd.read_excel(BytesIO(response.content), sheet_name="TM", header=1)


# =========================
# SUMBER DATA
# =========================
st.sidebar.header("📁 Sumber Data")

sumber_data = st.sidebar.radio(
    "Pilih sumber data",
    ["Upload Manual", "SharePoint / Google Drive"]
)

file_upload = st.file_uploader("Upload Excel", type=["xlsx"])

link_file = st.sidebar.text_input(
    "https://ptpln365-my.sharepoint.com/:x:/g/personal/irham_tantowi_ptpln365_onmicrosoft_com/IQAqM9WM9C3ySolssm7NJXPhAW14_mQHJp1ImR630d8zSUY?e=aaKOf4%22",
    placeholder="https://ptpln365-my.sharepoint.com/:x:/g/personal/irham_tantowi_ptpln365_onmicrosoft_com/IQAqM9WM9C3ySolssm7NJXPhAW14_mQHJp1ImR630d8zSUY?e=aaKOf4%22"
)

df = None

if sumber_data == "Upload Manual":
    if file_upload is not None:
        df = pd.read_excel(file_upload, sheet_name="TM", header=1)
        st.success("Menggunakan file upload manual.")
    else:
        st.warning("Silakan upload file Excel terlebih dahulu.")
        st.stop()

elif sumber_data == "SharePoint / Google Drive":
    if link_file:
        try:
            df = read_excel_from_url(link_file)
            st.success("Data berhasil dibaca dari link.")
        except Exception as e:
            st.error("Gagal membaca file dari link.")
            st.write(e)
            st.stop()
    else:
        st.warning("Masukkan link Excel terlebih dahulu.")
        st.stop()

df.columns = df.columns.astype(str).str.strip()

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
    index=2
)

kode_bulan = bulan_map[pilih_bulan]
index_bulan = urutan_bulan.index(kode_bulan)

def hitung_kwh(df, tahun, kode_bulan, mode):
    if mode == "Bulanan YoY":
        kolom_bulanan = f"kWh {kode_bulan} {tahun}"

        if kolom_bulanan in df.columns:
            return pd.to_numeric(df[kolom_bulanan], errors="coerce").fillna(0), kolom_bulanan
        else:
            return pd.Series(0, index=df.index), f"{kolom_bulanan} belum tersedia"

    kolom_kumulatif = f"kWh sd {kode_bulan} {tahun}"

    if kolom_kumulatif in df.columns:
        return pd.to_numeric(df[kolom_kumulatif], errors="coerce").fillna(0), kolom_kumulatif

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

st.success(f"Mode: {mode_periode}")
st.info(f"Sumber {tahun_lalu}: {sumber_lalu}")
st.info(f"Sumber {tahun_ini}: {sumber_ini}")

df = df.dropna(subset=["UP3", "TARIF", "KLUSTER USAHA"], how="any")

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
    pilih_golongan = st.sidebar.multiselect(
        "Golongan Tarif",
        sorted(df["GOLONGAN TARIF"].dropna().unique()),
        default=sorted(df["GOLONGAN TARIF"].dropna().unique())
    )
else:
    pilih_golongan = []

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
api_key = st.secrets["GOOGLE_API_KEY"]

df_filter = df[
    (df["TARIF"].isin(pilih_tarif)) &
    (df["UP3"].isin(pilih_up3)) &
    (df["KLUSTER USAHA"].isin(pilih_cluster))
]

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

total_lalu = df_filter["GWh Tahun Lalu"].sum()
total_ini = df_filter["GWh Tahun Ini"].sum()
delta = df_filter["Delta GWh"].sum()
growth = (delta / total_lalu * 100) if total_lalu != 0 else 0

col1, col2, col3, col4 = st.columns(4)
col1.metric(f"{mode_periode} {pilih_bulan} {tahun_lalu}", f"{total_lalu:,.2f} GWh")
col2.metric(f"{mode_periode} {pilih_bulan} {tahun_ini}", f"{total_ini:,.2f} GWh")
col3.metric("Delta", f"{delta:,.2f} GWh")
col4.metric("Growth YoY", f"{growth:.2f}%")

st.divider()

# =========================
# MAP SEBARAN PELANGGAN
# =========================
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
                lambda x: [231, 76, 60, 170] if x < 0 else [46, 134, 193, 170]
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

            st.caption("Biru = naik/positif, Merah = turun/negatif. Ukuran bubble mengikuti GWh Tahun Ini.")

# =========================
# GRAFIK GOLONGAN TARIF
# =========================
if "GOLONGAN TARIF" in df_filter.columns:
    st.subheader("📊 Penjualan per Golongan Tarif")

    grafik_golongan = df_filter.groupby(
        "GOLONGAN TARIF", as_index=False
    )[["GWh Tahun Lalu", "GWh Tahun Ini"]].sum()

    fig_golongan = px.bar(
        grafik_golongan,
        x="GOLONGAN TARIF",
        y=["GWh Tahun Lalu", "GWh Tahun Ini"],
        barmode="group",
        text_auto=".2f",
        title=f"Penjualan per Golongan Tarif - {mode_periode} {pilih_bulan}"
    )

    fig_golongan.update_layout(
        title=dict(font=dict(size=22)),
        xaxis=dict(title_font=dict(size=16), tickfont=dict(size=14)),
        yaxis=dict(title_font=dict(size=16), tickfont=dict(size=14)),
        legend=dict(font=dict(size=14))
    )

    st.plotly_chart(fig_golongan, use_container_width=True)

# =========================
# CLUSTER SUMMARY
# =========================
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
    top_cluster["Kategori"] = "Top"
    bottom_cluster["Kategori"] = "Bottom"
    cluster_tampil = pd.concat([top_cluster, bottom_cluster], ignore_index=True)
    cluster_tampil = cluster_tampil.drop_duplicates(subset=["KLUSTER USAHA"])
    judul_cluster = f"Top {top_n_cluster} dan Bottom {bottom_n_cluster} Cluster berdasarkan {metrik_cluster}"

cluster_tampil = cluster_tampil.sort_values(metrik_cluster, ascending=True)

cluster_tampil["Warna"] = cluster_tampil[metrik_cluster].apply(
    lambda x: "Positif" if x >= 0 else "Negatif"
)

st.subheader("🏭 Top & Bottom Cluster KBLI")

fig_cluster = px.bar(
    cluster_tampil,
    x=metrik_cluster,
    y="KLUSTER USAHA",
    orientation="h",
    text=metrik_cluster,
    color="Warna",
    color_discrete_map={
        "Positif": "#5DADE2",
        "Negatif": "#E74C3C"
    },
    title=f"{judul_cluster} - {mode_periode} {pilih_bulan}"
)

fig_cluster.update_traces(
    texttemplate="%{text:.1f}",
    textposition="inside",
    textfont=dict(size=16, color="white")
)

fig_cluster.update_layout(
    height=max(650, len(cluster_tampil) * 45),
    title=dict(font=dict(size=24)),
    xaxis=dict(
        title=metrik_cluster,
        title_font=dict(size=18),
        tickfont=dict(size=15),
        zeroline=True,
        zerolinewidth=2,
        zerolinecolor="black"
    ),
    yaxis=dict(
        title="KLUSTER USAHA",
        title_font=dict(size=18),
        tickfont=dict(size=15)
    ),
    legend=dict(font=dict(size=15), title_font=dict(size=16))
)

st.plotly_chart(fig_cluster, use_container_width=True)

st.subheader("🍩 Komposisi Penjualan per Cluster")

donut = cluster_summary.sort_values("GWh Tahun Ini", ascending=False).head(15)

fig_donut = px.pie(
    donut,
    names="KLUSTER USAHA",
    values="GWh Tahun Ini",
    hole=0.45,
    title=f"Share Top 15 Cluster - {mode_periode} {pilih_bulan} {tahun_ini}"
)

st.plotly_chart(fig_donut, use_container_width=True)

st.subheader("📋 Tabel Pivot Dinamis")

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

st.dataframe(pivot, use_container_width=True)

st.divider()
st.subheader("🧠 Generate AI Insight")

if st.button("🔍 Generate AI Insight"):
    if not api_key:
        st.warning("Masukkan API Key terlebih dahulu di sidebar.")
    else:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-3-flash-preview")

        data_ai_top = pivot.sort_values("Delta GWh", ascending=False).head(15)
        data_ai_bottom = pivot.sort_values("Delta GWh", ascending=True).head(15)
        data_ai = pd.concat([data_ai_top, data_ai_bottom], ignore_index=True)

        prompt = f"""
        Anda adalah Senior Business Analyst PLN UID Jawa Timur.

        KONTEKS:
        - Dashboard Penjualan Tenaga Listrik Tegangan Menengah (TM)
        - Mode periode: {mode_periode}
        - Bulan: {pilih_bulan}
        - Perbandingan: {tahun_lalu} vs {tahun_ini}

        DATA UTAMA:
        {data_ai.to_string(index=False)}

        INSTRUKSI ANALISIS:
        1. Fokus pada perubahan signifikan, baik positif maupun negatif.
        2. Gunakan angka spesifik dalam GWh dan persen.
        3. Identifikasi cluster/customer segment yang menjadi growth driver.
        4. Identifikasi cluster yang turun dan perlu perhatian.
        5. Berikan rekomendasi aksi yang konkret untuk UP3/ULP.
        6. Analisis bisa diselaraskan dengan isu/kondisi yang ada di media secara valid.
        7. anda bisa mencantumkan sumber data jika ambil dari external

        OUTPUT WAJIB:
        1. Executive Summary maksimal 3 kalimat.
        2. Top Growth Drivers.
        3. Underperforming Cluster.
        4. Warning Area.
        5. Rekomendasi Aksi.
        6. Narasi singkat untuk bahan paparan GM UID Jawa Timur.

        GAYA BAHASA:
        Formal, tajam, ringkas, dan cocok untuk bahan paparan manajemen PLN.
        """

        with st.spinner("AI sedang menganalisis data..."):
            response = model.generate_content(prompt)

        st.success("AI Insight berhasil dibuat.")
        st.markdown(response.text)
