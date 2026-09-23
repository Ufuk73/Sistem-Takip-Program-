import streamlit as st
import pandas as pd
import numpy as np

# Sayfa yapılandırması
st.set_page_config(page_title="Sistem Takip Programı", layout="wide")

st.title("💻 Sistem Takip Programı")

# Örnek veri seti (Kendi verinizle değiştirebilirsiniz)
@st.cache_data
def load_data():
    data = pd.DataFrame({
        "Tarih": pd.date_range(start="2026-01-01", periods=10, freq="D"),
        "Sistem / Cihaz": [f"Sunucu-{i}" for i in range(1, 11)],
        "Durum": ["Aktif", "Pasif", "Aktif", "Bakımda", "Aktif", "Aktif", "Pasif", "Bakımda", "Aktif", "Aktif"],
        "Kullanım Oranı (%)": [45, 12, 78, 0, 88, 62, 5, 0, 91, 53],
        "Arıza Sayısı": [0, 2, 1, 0, 3, 0, 1, 0, 4, 1]
    })
    return data

df = load_data()

# Özet Metrik Kartları
st.subheader("📊 Genel Durum Özeti")
col1, col2, col3, col4 = st.columns(4)
col1.metric("Toplam Cihaz", len(df))
col2.metric("Aktif Sistemler", len(df[df["Durum"] == "Aktif"]))
col3.metric("Bakımdaki Sistemler", len(df[df["Durum"] == "Bakımda"]))
col4.metric("Ort. Kullanım", f"%{df['Kullanım Oranı (%)'].mean():.1f}")

st.divider()

# Ekstra kütüphane gerektirmeyen yerel Streamlit Grafikleri
st.subheader("📈 Performans ve Kullanım Analizi")

col_chart1, col_chart2 = st.columns(2)

with col_chart1:
    st.write("**Cihaz Başına Kullanım Oranları (%)**")
    chart_data = df.set_index("Sistem / Cihaz")[["Kullanım Oranı (%)"]]
    st.bar_chart(chart_data)  # Doğrudan Streamlit sütun grafiği

with col_chart2:
    st.write("**Zaman İçindeki Kullanım Trendi**")
    line_data = df.set_index("Tarih")[["Kullanım Oranı (%)"]]
    st.line_chart(line_data)  # Doğrudan Streamlit çizgi grafiği

st.divider()

# Veri Tablosu ve Filtreleme
st.subheader("📋 Sistem Listesi ve Detaylar")

durum_filtresi = st.multiselect(
    "Duruma Göre Filtrele:",
    options=df["Durum"].unique(),
    default=df["Durum"].unique()
)

filtered_df = df[df["Durum"].isin(durum_filtresi)]
st.dataframe(filtered_df, use_container_width=True)