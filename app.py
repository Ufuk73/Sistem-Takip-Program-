import os
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# SAYFA YAPILANDIRMASI
# ---------------------------------------------------------
st.set_page_config(
    page_title="Sistem Takip Programı",
    page_icon="💻",
    layout="wide",
    initial_sidebar_state="expanded",
)

VERI_DOSYASI = "sistem_verileri.csv"


# 1. OTOMATİK VERİ YÜKLEME FONKSİYONU
def veri_yukle():
    if os.path.exists(VERI_DOSYASI):
        try:
            return pd.read_csv(VERI_DOSYASI)
        except Exception:
            pass

    # Dosya henüz yoksa varsayılan başlangıç verisi oluşturulur
    varsayilan_veri = pd.DataFrame(
        [
            {
                "Seri No": "SN-1001",
                "Cihaz / Sistem": "Ana Sunucu",
                "Kategori": "Sunucu",
                "Bölge / Lokasyon": "Sistem Odası",
                "Durum": "Aktif",
                "Sorumlu": "Ahmet Yılmaz",
                "Notlar": "Düzenli çalışıyor",
            },
            {
                "Seri No": "SN-1002",
                "Cihaz / Sistem": "Kenar Anahtar (Switch)",
                "Kategori": "Ağ Cihazı",
                "Bölge / Lokasyon": "Kat 1 Pano",
                "Durum": "Aktif",
                "Sorumlu": "Mehmet Demir",
                "Notlar": "Portlar dolu",
            },
            {
                "Seri No": "SN-1003",
                "Cihaz / Sistem": "Yedek Güç Kaynağı (UPS)",
                "Kategori": "Güç",
                "Bölge / Lokasyon": "Sistem Odası",
                "Durum": "Bakımda",
                "Sorumlu": "Ali Kaya",
                "Notlar": "Akü değişimi bekleniyor",
            },
        ]
    )
    varsayilan_veri.to_csv(VERI_DOSYASI, index=False)
    return varsayilan_veri


# 2. OTOMATİK KAYDETME FONKSİYONU
def veri_kaydet(df):
    df.to_csv(VERI_DOSYASI, index=False)


# Uygulama hafızasına veriyi yükle
if "df" not in st.session_state:
    st.session_state.df = veri_yukle()

# ---------------------------------------------------------
# ANA BAŞLIK
# ---------------------------------------------------------
st.title("💻 Sistem Takip Programı")
st.caption(
    "💾 **Otomatik Kayıt Devrede:** Tablodan yapılan değişiklikler veya yeni eklenen cihazlar anında kaydedilir."
)

# ---------------------------------------------------------
# YAN MENÜ (SIDEBAR) - YENİ CİHAZ EKLEME
# ---------------------------------------------------------
st.sidebar.header("➕ Yeni Cihaz / Sistem Ekle")

with st.sidebar.form("yeni_cihaz_formu", clear_on_submit=True):
    yeni_sn = st.text_input("Seri No / Kod", placeholder="Örn: SN-1004")
    yeni_ad = st.text_input("Cihaz / Sistem Adı", placeholder="Örn: Güvenlik Kamerası")
    yeni_kat = st.selectbox(
        "Kategori",
        ["Sunucu", "Ağ Cihazı", "Güç", "Bilgisayar", "Saha Cihazı", "Diğer"],
    )
    yeni_lok = st.text_input("Bölge / Lokasyon", placeholder="Örn: Merkez Bina")
    yeni_durum = st.selectbox("Durum", ["Aktif", "Pasif", "Bakımda", "Arızalı"])
    yeni_sorumlu = st.text_input("Sorumlu Personel", placeholder="Örn: Ayşe Öztürk")
    yeni_not = st.text_area("Notlar / Açıklama", placeholder="Varsa notlar...")

    kaydet_btn = st.form_submit_button("Sisteme Kaydet")

    if kaydet_btn:
        if yeni_ad.strip() == "":
            st.sidebar.error("Lütfen cihaz/sistem adını boş bırakmayın!")
        else:
            yeni_satir = {
                "Seri No": yeni_sn,
                "Cihaz / Sistem": yeni_ad,
                "Kategori": yeni_kat,
                "Bölge / Lokasyon": yeni_lok,
                "Durum": yeni_durum,
                "Sorumlu": yeni_sorumlu,
                "Notlar": yeni_not,
            }
            # Yeni kaydı listeye ekle ve anında CSV'ye yaz
            st.session_state.df = pd.concat(
                [st.session_state.df, pd.DataFrame([yeni_satir])],
                ignore_index=True,
            )
            veri_kaydet(st.session_state.df)
            st.sidebar.success("Yeni cihaz eklendi ve otomatik kaydedildi! 💾")
            st.rerun()

# ---------------------------------------------------------
# ÖZET METRİK KARTLARI
# ---------------------------------------------------------
col1, col2, col3, col4 = st.columns(4)
toplam_cihaz = len(st.session_state.df)
aktif_cihaz = len(st.session_state.df[st.session_state.df["Durum"] == "Aktif"])
bakim_cihaz = len(st.session_state.df[st.session_state.df["Durum"] == "Bakımda"])
ariza_cihaz = len(
    st.session_state.df[st.session_state.df["Durum"].isin(["Arızalı", "Pasif"])]
)

col1.metric("Toplam Sistem / Cihaz", toplam_cihaz)
col2.metric("Aktif Cihazlar", aktif_cihaz)
col3.metric("Bakımdakiler", bakim_cihaz)
col4.metric("Arızalı / Pasif", ariza_cihaz)

st.divider()

# ---------------------------------------------------------
# FİLTRELEME VE ARAMA
# ---------------------------------------------------------
col_ara, col_durum, col_kat = st.columns([2, 1, 1])

with col_ara:
    arama_metni = st.text_input("🔍 Cihaz, Seri No veya Lokasyon Ara:")

with col_durum:
    durum_filtresi = st.multiselect(
        "Durum Filtresi:",
        options=list(st.session_state.df["Durum"].unique()),
        default=list(st.session_state.df["Durum"].unique()),
    )

with col_kat:
    kat_filtresi = st.multiselect(
        "Kategori Filtresi:",
        options=list(st.session_state.df["Kategori"].unique()),
        default=list(st.session_state.df["Kategori"].unique()),
    )

# Filtrelerin Uygulanması
filtreli_df = st.session_state.df[
    (st.session_state.df["Durum"].isin(durum_filtresi))
    & (st.session_state.df["Kategori"].isin(kat_filtresi))
]

if arama_metni:
    filtreli_df = filtreli_df[
        filtreli_df.apply(
            lambda row: arama_metni.lower()
            in row.astype(str).str.lower().str.cat(sep=" "),
            axis=1,
        )
    ]

# ---------------------------------------------------------
# CANLI DÜZENLENEBİLİR TABLO VE OTOMATİK KAYIT
# ---------------------------------------------------------
st.subheader("📋 Sistem Listesi")
st.caption(
    "💡 *Hücrelere tıklayarak doğrudan düzenleme yapabilirsiniz. Alt kısımdan yeni satır ekleyebilir veya satır silebilirsiniz.*"
)

# st.data_editor ile interaktif tablo
guncellenen_df = st.data_editor(
    filtreli_df,
    num_rows="dynamic",  # Satır ekleme/silmeye izin verir
    use_container_width=True,
    key="editor_tablosu",
)

# Tabloda bir değişiklik yapıldığında OTOMATİK KAYDET
if not guncellenen_df.equals(filtreli_df):
    if (
        len(durum_filtresi) == len(st.session_state.df["Durum"].unique())
        and len(kat_filtresi) == len(st.session_state.df["Kategori"].unique())
        and not arama_metni
    ):
        st.session_state.df = guncellenen_df
    else:
        st.session_state.df.update(guncellenen_df)

    veri_kaydet(st.session_state.df)
    st.toast("Değişiklikler otomatik kaydedildi! 💾", icon="✅")

st.divider()

# ---------------------------------------------------------
# DAHİLİ GRAFİKLER (Sadece Streamlit)
# ---------------------------------------------------------
st.subheader("📊 Özet Grafikler")
col_g1, col_g2 = st.columns(2)

with col_g1:
    st.write("**Duruma Göre Dağılım**")
    durum_sayilari = st.session_state.df["Durum"].value_counts()
    st.bar_chart(durum_sayilari)

with col_g2:
    st.write("**Kategoriye Göre Dağılım**")
    kat_sayilari = st.session_state.df["Kategori"].value_counts()
    st.bar_chart(kat_sayilari)