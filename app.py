import streamlit as st
import traceback

# Sayfa Yapılandırması
st.set_page_config(
    page_title="Sistem ve Parça Takip Sistemi",
    page_icon="⚙️",
    layout="wide"
)

try:
    import sqlite3
    import datetime
    import pandas as pd
    import os
    import io

    DB_DOSYASI = "sistem_takip.db"
    UPLOAD_FOLDER = "yuklenen_dosyalar"
    
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)

    # --- TÜRKÇE BÜYÜK HARF DÖNÜŞÜMÜ ---
    def tr_upper(text):
        if not isinstance(text, str):
            return str(text) if text is not None else ""
        return (
            text.replace("i", "İ")
            .replace("ı", "I")
            .replace("ş", "Ş")
            .replace("ğ", "Ğ")
            .replace("ü", "Ü")
            .replace("ö", "Ö")
            .replace("ç", "Ç")
            .upper()
        )

    # --- VERİTABANI BAŞLATMA VE MIGRATION ---
    def veritabanini_hazirla():
        conn = sqlite3.connect(DB_DOSYASI)
        cursor = conn.cursor()
        
        cursor.execute("CREATE TABLE IF NOT EXISTS parcalar (id INTEGER PRIMARY KEY AUTOINCREMENT)")
        cursor.execute("CREATE TABLE IF NOT EXISTS islem_loglari (id INTEGER PRIMARY KEY AUTOINCREMENT, zaman TEXT, islem_turu TEXT, detay TEXT)")
        cursor.execute("CREATE TABLE IF NOT EXISTS gunluk_notlar (id INTEGER PRIMARY KEY AUTOINCREMENT, tarih TEXT, bolge TEXT, detay TEXT)")
        cursor.execute("CREATE TABLE IF NOT EXISTS parca_gecmis (id INTEGER PRIMARY KEY AUTOINCREMENT, parca_sn TEXT, tarih TEXT, islem TEXT, aciklama TEXT)")
        
        beklenen_kolonlar = {
            "bolge": "TEXT",
            "sistem_adi": "TEXT",
            "sistem_pn": "TEXT",
            "sistem_sn": "TEXT",
            "parca_adi": "TEXT",
            "parca_pn": "TEXT",
            "parca_sn": "TEXT",
            "durum": "TEXT",
            "onarim_tarih": "TEXT",
            "aciklama": "TEXT",
            "dosya_adi": "TEXT"
        }
        
        cursor.execute("PRAGMA table_info(parcalar)")
        mevcut_kolonlar = [kol[1] for kol in cursor.fetchall()]
        
        for kolon_adi, kolon_tipi in beklenen_kolonlar.items():
            if kolon_adi not in mevcut_kolonlar:
                cursor.execute(f"ALTER TABLE parcalar ADD COLUMN {kolon_adi} {kolon_tipi}")
            
        conn.commit()
        conn.close()

    def log_yaz(islem_turu, detay):
        try:
            zaman = datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")
            conn = sqlite3.connect(DB_DOSYASI)
            cursor = conn.cursor()
            cursor.execute("INSERT INTO islem_loglari (zaman, islem_turu, detay) VALUES (?, ?, ?)", (zaman, islem_turu, detay))
            conn.commit()
            conn.close()
        except Exception as e:
            print(f"Log yazılamadı: {e}")

    veritabanini_hazirla()

    # --- ÜST BAŞLIK ---
    st.markdown("### ⚙️ Sistem ve Parça Takip Sistemi")

    # Verileri Çek
    conn = sqlite3.connect(DB_DOSYASI)
    df_parcalar = pd.read_sql_query("SELECT * FROM parcalar", conn)
    conn.close()

    # İstatistik Hesaplama
    toplam = len(df_parcalar)
    faal = len(df_parcalar[df_parcalar["durum"] == "FAAL"]) if not df_parcalar.empty and "durum" in df_parcalar.columns else 0
    yedek = len(df_parcalar[df_parcalar["durum"] == "YEDEK PARÇA"]) if not df_parcalar.empty and "durum" in df_parcalar.columns else 0
    onarimda = len(df_parcalar[df_parcalar["durum"] == "ONARIMDA"]) if not df_parcalar.empty and "durum" in df_parcalar.columns else 0
    gayri = len(df_parcalar[df_parcalar["durum"] == "GAYRI FAAL"]) if not df_parcalar.empty and "durum" in df_parcalar.columns else 0

    # Kritik onarım hesaplama (30 gün+)
    kritik = 0
    kritik_liste = []
    if not df_parcalar.empty and "onarim_tarih" in df_parcalar.columns and "durum" in df_parcalar.columns:
        bugun = datetime.date.today()
        for _, row in df_parcalar[df_parcalar["durum"] == "ONARIMDA"].iterrows():
            o_tarih = row["onarim_tarih"]
            if o_tarih:
                try:
                    baslangic = datetime.datetime.strptime(str(o_tarih).strip(), "%d.%m.%Y").date()
                    gecen_gun = (bugun - baslangic).days
                    if gecen_gun >= 30:
                        kritik += 1
                        kritik_liste.append(f"• **Bölge: {row.get('bolge', 'Bölge Yok')}** ({row.get('parca_adi', 'Parça')} - SN: {row.get('parca_sn', '-')}) -> {gecen_gun} gündür onarımda!")
                except ValueError:
                    pass

    if kritik > 0:
        st.error(f"⚠️ **DİKKAT:** 30 Günü Aşan Onarımda Bekleyen **{kritik}** Adet Parça Bulunuyor!")
        with st.expander("Kritik Parçaları Listele", expanded=False):
            for k_bilgi in kritik_liste:
                st.markdown(k_bilgi)

    # --- ÇERÇEVELİ (KART) GÖRÜNÜM İÇİN ÖZEL CSS ---
    st.markdown("""
    <style>
    .metric-card {
        background-color: #f8f9fa;
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        padding: 12px 15px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    .metric-title {
        font-size: 13px;
        color: #6c757d;
        font-weight: 600;
        margin-bottom: 4px;
        text-transform: uppercase;
    }
    .metric-value {
        font-size: 22px;
        color: #333333;
        font-weight: 700;
    }
    </style>
    """, unsafe_allow_html=True)

    # Metrikleri Çerçeveli Sütunlar Halinde Yerleştirme
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    
    with c1:
        st.markdown(f'<div class="metric-card"><div class="metric-title">Toplam</div><div class="metric-value">{toplam}</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card"><div class="metric-title">Faal</div><div class="metric-value" style="color: #28a745;">{faal}</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="metric-card"><div class="metric-title">Yedek</div><div class="metric-value" style="color: #17a2b8;">{yedek}</div></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="metric-card"><div class="metric-title">Onarımda</div><div class="metric-value" style="color: #ffc107;">{onarimda}</div></div>', unsafe_allow_html=True)
    with c5:
        st.markdown(f'<div class="metric-card"><div class="metric-title">Kritik (30+)</div><div class="metric-value" style="color: #dc3545;">{kritik}</div></div>', unsafe_allow_html=True)
    with c6:
        st.markdown(f'<div class="metric-card"><div class="metric-title">Gayri Faal</div><div class="metric-value" style="color: #343a40;">{gayri}</div></div>', unsafe_allow_html=True)

    st.markdown("<hr style='margin:15px 0 10px 0;'>", unsafe_allow_html=True)

    tab_takip, tab_ekle, tab_notlar, tab_loglar, tab_gecmis, tab_yonetim = st.tabs([
        "📋 Sistem Takip", 
        "➕ Yeni Kayıt", 
        "📝 Günlük Notlar", 
        "📜 Loglar", 
        "🔍 Parça Geçmişi",
        "⚙️ Yönetim"
    ])

    # Dialog (Açılır Pencere) ile Düzenleme Fonksiyonu
    @st.dialog("Kayıt Düzenle", width="large")
    def duzenle_dialog(r_id):
        conn = sqlite3.connect(DB_DOSYASI)
        df_tekil = pd.read_sql_query("SELECT * FROM parcalar WHERE id = ?", conn, params=(r_id,))
        conn.close()
        
        if not df_tekil.empty:
            row = df_tekil.iloc[0]
            with st.form(key=f"dialog_form_{r_id}"):
                e_bolge = st.text_input("Bölge", value=row.get("bolge", "") if pd.notna(row.get("bolge", "")) else "")
                e_sistem = st.text_input("Sistem Adı", value=row.get("sistem_adi", "") if pd.notna(row.get("sistem_adi", "")) else "")
                
                col_d1, col_d2 = st.columns(2)
                e_s_pn = col_d1.text_input("Sistem PN", value=row.get("sistem_pn", "") if pd.notna(row.get("sistem_pn", "")) else "")
                e_s_sn = col_d2.text_input("Sistem SN", value=row.get("sistem_sn", "") if pd.notna(row.get("sistem_sn", "")) else "")
                
                e_parca = st.text_input("Parça Adı", value=row.get("parca_adi", "") if pd.notna(row.get("parca_adi", "")) else "")
                
                col_d3, col_d4 = st.columns(2)
                e_p_pn = col_d3.text_input("Parça PN", value=row.get("parca_pn", "") if pd.notna(row.get("parca_pn", "")) else "")
                e_p_sn = col_d4.text_input("Parça SN", value=row.get("parca_sn", "") if pd.notna(row.get("parca_sn", "")) else "")
                
                mevcut_d = row.get("durum", "FAAL")
                d_list = ["FAAL", "YEDEK PARÇA", "ONARIMDA", "GAYRI FAAL"]
                d_idx = d_list.index(mevcut_d) if mevcut_d in d_list else 0
                e_durum = st.selectbox("Durum", d_list, index=d_idx)
                
                e_tarih = st.text_input("Onarım Tarihi (GG.AA.YYYY)", value=row.get("onarim_tarih", "") if pd.notna(row.get("onarim_tarih", "")) else "")
                e_aciklama = st.text_area("Açıklama / Not", value=row.get("aciklama", "") if pd.notna(row.get("aciklama", "")) else "")
                
                if st.form_submit_button("Güncellemeyi Kaydet", type="primary"):
                    conn = sqlite3.connect(DB_DOSYASI)
                    cursor = conn.cursor()
                    cursor.execute("""
                        UPDATE parcalar SET bolge=?, sistem_adi=?, sistem_pn=?, sistem_sn=?, parca_adi=?, parca_pn=?, parca_sn=?, durum=?, onarim_tarih=?, aciklama=?
                        WHERE id=?
                    """, (tr_upper(e_bolge), tr_upper(e_sistem), tr_upper(e_s_pn), tr_upper(e_s_sn), tr_upper(e_parca), tr_upper(e_p_pn), tr_upper(e_p_sn), e_durum, e_tarih, tr_upper(e_aciklama), r_id))
                    
                    cursor.execute("INSERT INTO parca_gecmis (parca_sn, tarih, islem, aciklama) VALUES (?, ?, ?, ?)", 
                                   (tr_upper(e_p_sn), datetime.datetime.now().strftime("%d.%m.%Y %H:%M"), f"GÜNCELLEME ({e_durum})", tr_upper(e_aciklama)))
                    conn.commit()
                    conn.close()
                    log_yaz("GÜNCELLEME", f"ID {r_id} güncellendi.")
                    st.success("Kayıt güncellendi!")
                    st.rerun()

    # 1. SEKME: TAKİP & FİLTRELEME
    with tab_takip:
        if not df_parcalar.empty:
            col_f1, col_f2, col_f3, col_f4 = st.columns(4)
            
            bolgeler = ["TÜMÜ"] + list(df_parcalar["bolge"].dropna().unique()) if "bolge" in df_parcalar.columns else ["TÜMÜ"]
            parcalar = ["TÜMÜ"] + list(df_parcalar["parca_adi"].dropna().unique()) if "parca_adi" in df_parcalar.columns else ["TÜMÜ"]
            durumlar = ["TÜMÜ", "FAAL", "YEDEK PARÇA", "ONARIMDA", "GAYRI FAAL", "30 GÜN+ KRİTİK"]
            
            f_bolge = col_f1.selectbox("Bölge", bolgeler)
            f_parca = col_f2.selectbox("Parça Adı", parcalar)
            f_durum = col_f3.selectbox("Durum", durumlar)
            f_arama = col_f4.text_input("Hızlı Arama")
            
            filt_df = df_parcalar.copy()
            if f_bolge != "TÜMÜ" and "bolge" in filt_df.columns:
                filt_df = filt_df[filt_df["bolge"] == f_bolge]
            if f_parca != "TÜMÜ" and "parca_adi" in filt_df.columns:
                filt_df = filt_df[filt_df["parca_adi"] == f_parca]
            if f_durum != "TÜMÜ" and f_durum != "30 GÜN+ KRİTİK" and "durum" in filt_df.columns:
                filt_df = filt_df[filt_df["durum"] == f_durum]
            elif f_durum == "30 GÜN+ KRİTİK" and "durum" in filt_df.columns and "onarim_tarih" in filt_df.columns:
                kritik_idler = []
                bugun = datetime.date.today()
                for _, r in filt_df[filt_df["durum"] == "ONARIMDA"].iterrows():
                    if r["onarim_tarih"]:
                        try:
                            b_t = datetime.datetime.strptime(str(r["onarim_tarih"]).strip(), "%d.%m.%Y").date()
                            if (bugun - b_t).days >= 30:
                                kritik_idler.append(r["id"])
                        except ValueError:
                            pass
                filt_df = filt_df[filt_df["id"].isin(kritik_idler)]
                
            if f_arama:
                a_upper = tr_upper(f_arama)
                filt_df = filt_df[filt_df.apply(lambda row: row.astype(str).str.upper().str.contains(a_upper).any(), axis=1)]
                
            st.markdown("<hr style='margin:5px 0;'>", unsafe_allow_html=True)
            
            # --- NOT FORMATINDA İNDİRME ÖZELLİĞİ ---
            if not filt_df.empty:
                col_exp1, col_exp2 = st.columns([7, 3])
                with col_exp2:
                    tarih_str = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
                    not_icerik = f"SİSTEM VE PARÇA LİSTESİ NOTU\n"
                    not_icerik += f"Tarih: {tarih_str}\n"
                    not_icerik += f"--------------------------------------------------\n"
                    not_icerik += f"Seçilen Filtreler -> Bölge: {f_bolge} | Parça: {f_parca} | Durum: {f_durum}"
                    if f_arama:
                        not_icerik += f" | Arama: {f_arama}"
                    not_icerik += f"\nToplam Kayıt: {len(filt_df)}\n"
                    not_icerik += f"--------------------------------------------------\n\n"

                    for idx, (_, row) in enumerate(filt_df.iterrows(), 1):
                        not_icerik += f"{idx}. Bölge: {row.get('bolge', '-')}\n"
                        not_icerik += f"   Sistem: {row.get('sistem_adi', '-')} (PN: {row.get('sistem_pn', '-')}, SN: {row.get('sistem_sn', '-')})\n"
                        not_icerik += f"   Parça : {row.get('parca_adi', '-')} (PN: {row.get('parca_pn', '-')}, SN: {row.get('parca_sn', '-')})\n"
                        not_icerik += f"   Durum : {row.get('durum', '-')} | Onarım Tar.: {row.get('onarim_tarih', '-')}\n"
                        if pd.notna(row.get('aciklama')) and str(row.get('aciklama')).strip():
                            not_icerik += f"   Not   : {row.get('aciklama')}\n"
                        not_icerik += f"--------------------------------------------------\n"

                    st.download_button(
                        "📝 Not Olarak İndir", 
                        data=not_icerik.encode('utf-8'), 
                        file_name=f"sistem_notu_{datetime.date.today().strftime('%d_%m_%Y')}.txt", 
                        mime="text/plain",
                        help="Filtrelenen listeyi sade bir not dosyası olarak indirir."
                    )
            
            if not filt_df.empty and "bolge" in filt_df.columns:
                for b_adi, b_grubu in filt_df.groupby("bolge"):
                    with st.expander(f"📍 BÖLGE: {b_adi} ({len(b_grubu)} Kayıt)", expanded=False):
                        h1, h2, h3, h4, h5 = st.columns([2.2, 2.2, 1.8, 0.9, 0.9])
                        h1.markdown("<small><b>SİSTEM ADI</b></small>", unsafe_allow_html=True)
                        h2.markdown("<small><b>PARÇA & SERİ NO</b></small>", unsafe_allow_html=True)
                        h3.markdown("<small><b>DURUM / TARİH</b></small>", unsafe_allow_html=True)
                        h4.markdown("<small><b>DÜZENLE</b></small>", unsafe_allow_html=True)
                        h5.markdown("<small><b>SİL</b></small>", unsafe_allow_html=True)
                        st.markdown("<hr style='margin:2px 0 5px 0;'>", unsafe_allow_html=True)

                        for _, row in b_grubu.iterrows():
                            r_id = row["id"]
                            durum = row.get("durum", "FAAL")
                            durum_badge = f"🟢 {durum}" if durum == "FAAL" else f"🟡 {durum}" if durum == "ONARIMDA" else f"🔴 {durum}" if durum == "GAYRI FAAL" else f"🔵 {durum}"

                            c1, c2, c3, c4, c5 = st.columns([2.2, 2.2, 1.8, 0.9, 0.9])
                            
                            with c1:
                                st.markdown(f"<div style='line-height: 1.1;'><small><b>{row.get('sistem_adi', '-')}</b><br><span style='color:gray;'>PN: {row.get('sistem_pn', '-')}</span></small></div>", unsafe_allow_html=True)
                            with c2:
                                st.markdown(f"<div style='line-height: 1.1;'><small>{row.get('parca_adi', '-')}<br><span style='color:gray;'>SN: {row.get('parca_sn', '-')}</span></small></div>", unsafe_allow_html=True)
                            with c3:
                                st.markdown(f"<div style='line-height: 1.1;'><small>{durum_badge}<br><span style='color:gray;'>{row.get('onarim_tarih', '-')}</span></small></div>", unsafe_allow_html=True)
                                
                            with c4:
                                if st.button("✏️", key=f"edit_{r_id}", help="Düzenle"):
                                    duzenle_dialog(r_id)
                                    
                            with c5:
                                if st.button("🗑️", key=f"del_{r_id}", help="Sil"):
                                    conn = sqlite3.connect(DB_DOSYASI)
                                    cursor = conn.cursor()
                                    cursor.execute("DELETE FROM parcalar WHERE id=?", (r_id,))
                                    conn.commit()
                                    conn.close()
                                    log_yaz("SİLME", f"ID {r_id} silindi.")
                                    st.success("Silindi!")
                                    st.rerun()

                            st.markdown("<hr style='margin:3px 0;'>", unsafe_allow_html=True)
            else:
                st.info("Filtreleme kriterlerine uygun kayıt bulunamadı.")
        else:
            st.info("Kayıt bulunmuyor.")

    # 2. SEKME: YENİ KAYIT EKLE
    with tab_ekle:
        kayitli_bolgeler = sorted(list(df_parcalar["bolge"].dropna().unique())) if not df_parcalar.empty and "bolge" in df_parcalar.columns else []
        
        secim_tipi = st.radio("Bölge Giriş", ["Kayıtlı Seç", "Yeni Yaz"], horizontal=True, key="form_secim_tipi")
        
        aktif_bolge = ""
        if secim_tipi == "Kayıtlı Seç" and len(kayitli_bolgeler) > 0:
            aktif_bolge = st.selectbox("Bölge Seç", kayitli_bolgeler, key="form_secilen_bolge")
        else:
            aktif_bolge = st.text_input("Yeni Bölge Adı", key="form_yeni_bolge")

        varsayilan_sistem = ""
        varsayilan_s_pn = ""
        varsayilan_s_sn = ""
        
        if aktif_bolge and not df_parcalar.empty and "bolge" in df_parcalar.columns:
            eslesen_kayitlar = df_parcalar[df_parcalar["bolge"].str.upper() == str(aktif_bolge).upper()]
            if not eslesen_kayitlar.empty:
                son_kayit = eslesen_kayitlar.iloc[-1]
                varsayilan_sistem = son_kayit.get("sistem_adi", "") if pd.notna(son_kayit.get("sistem_adi", "")) else ""
                varsayilan_s_pn = son_kayit.get("sistem_pn", "") if pd.notna(son_kayit.get("sistem_pn", "")) else ""
                varsayilan_s_sn = son_kayit.get("sistem_sn", "") if pd.notna(son_kayit.get("sistem_sn", "")) else ""

        with st.form("yeni_kayit_formu", clear_on_submit=True):
            st.markdown(f"📌 **Seçilen/Girilen Bölge:** `{aktif_bolge if aktif_bolge else 'Henüz seçilmedi'}`")
            
            e_sistem = st.text_input("Sistem Adı", value=varsayilan_sistem)
            c1, c2 = st.columns(2)
            e_s_pn = c1.text_input("Sistem PN", value=varsayilan_s_pn)
            e_s_sn = c2.text_input("Sistem SN", value=varsayilan_s_sn)
            
            e_parca = st.text_input("Parça Adı")
            c3, c4 = st.columns(2)
            e_p_pn = c3.text_input("Parça PN", value="")
            e_p_sn = c4.text_input("Parça SN", value="")
            
            e_durum = st.selectbox("Durum", ["FAAL", "YEDEK PARÇA", "ONARIMDA", "GAYRI FAAL"])
            e_tarih = st.text_input("Onarım Tarihi (GG.AA.YYYY)", value=datetime.datetime.now().strftime("%d.%m.%Y"))
            e_aciklama = st.text_area("Açıklama")
            yuklenen_dosya_form = st.file_uploader("Belge/Fotoğraf", type=["png", "jpg", "jpeg", "pdf"])
            
            if st.form_submit_button("Sisteme Kaydet", type="primary"):
                if not aktif_bolge or not e_sistem or not e_parca:
                    st.warning("Bölge, Sistem Adı ve Parça Adı zorunludur!")
                else:
                    dosya_ismi = None
                    if yuklenen_dosya_form is not None:
                        dosya_ismi = f"{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{yuklenen_dosya_form.name}"
                        with open(os.path.join(UPLOAD_FOLDER, dosya_ismi), "wb") as f:
                            f.write(yuklenen_dosya_form.getbuffer())

                    conn = sqlite3.connect(DB_DOSYASI)
                    cursor = conn.cursor()
                    cursor.execute("""
                        INSERT INTO parcalar (bolge, sistem_adi, sistem_pn, sistem_sn, parca_adi, parca_pn, parca_sn, durum, onarim_tarih, aciklama, dosya_adi)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (tr_upper(aktif_bolge), tr_upper(e_sistem), tr_upper(e_s_pn), tr_upper(e_s_sn), tr_upper(e_parca), tr_upper(e_p_pn), tr_upper(e_p_sn), e_durum, e_tarih, tr_upper(e_aciklama), dosya_ismi))
                    
                    cursor.execute("INSERT INTO parca_gecmis (parca_sn, tarih, islem, aciklama) VALUES (?, ?, ?, ?)", 
                                   (tr_upper(e_p_sn), datetime.datetime.now().strftime("%d.%m.%Y %H:%M"), f"İLK KAYIT ({e_durum})", tr_upper(e_aciklama)))
                    conn.commit()
                    conn.close()
                    log_yaz("YENİ KAYIT", f"Bölge: {aktif_bolge}, Sistem: {e_sistem} eklendi.")
                    st.success("Kayıt eklendi!")
                    st.rerun()

    # 3. SEKME: GÜNLÜK NOTLAR
    with tab_notlar:
        kayitli_not_bolgeler = sorted(list(df_parcalar["bolge"].dropna().unique())) if not df_parcalar.empty and "bolge" in df_parcalar.columns else []
        with st.form("not_form", clear_on_submit=True):
            n_tarih = st.text_input("Tarih", value=datetime.datetime.now().strftime("%d.%m.%Y"))
            n_bolge = st.selectbox("Bölge", kayitli_not_bolgeler) if len(kayitli_not_bolgeler) > 0 else st.text_input("Bölge")
            n_detay = st.text_area("Not Detayı")
            if st.form_submit_button("Günlük Not Ekle"):
                conn = sqlite3.connect(DB_DOSYASI)
                cursor = conn.cursor()
                cursor.execute("INSERT INTO gunluk_notlar (tarih, bolge, detay) VALUES (?, ?, ?)", (n_tarih, tr_upper(n_bolge), tr_upper(n_detay)))
                conn.commit()
                conn.close()
                st.success("Günlük not eklendi!")
                st.rerun()
                
        conn = sqlite3.connect(DB_DOSYASI)
        df_notlar = pd.read_sql_query("SELECT * FROM gunluk_notlar ORDER BY id DESC", conn)
        conn.close()
        if not df_notlar.empty:
            st.dataframe(df_notlar, use_container_width=True, hide_index=True)

    # 4. SEKME: LOGLAR
    with tab_loglar:
        conn = sqlite3.connect(DB_DOSYASI)
        df_loglar = pd.read_sql_query("SELECT * FROM islem_loglari ORDER BY id DESC", conn)
        conn.close()
        if not df_loglar.empty:
            st.dataframe(df_loglar, use_container_width=True, hide_index=True)

    # 5. SEKME: PARÇA GEÇMİŞİ
    with tab_gecmis:
        st.subheader("Bölge Bazlı Parça ve İşlem Geçmişi")
        
        if not df_parcalar.empty and "bolge" in df_parcalar.columns:
            benzersiz_bolgeler = sorted(list(df_parcalar["bolge"].dropna().unique()))
            
            if benzersiz_bolgeler:
                secilen_gecmis_bolge = st.selectbox("İncelemek İçin Bölge Seçin", ["TÜM BÖLGELER"] + benzersiz_bolgeler)
                
                bolge_filtreli_df = df_parcalar.copy()
                if secilen_gecmis_bolge != "TÜM BÖLGELER":
                    bolge_filtreli_df = bolge_filtreli_df[bolge_filtreli_df["bolge"] == secilen_gecmis_bolge]
                
                for bolge_adi, bolge_grubu in bolge_filtreli_df.groupby("bolge"):
                    with st.expander(f"📍 BÖLGE: {bolge_adi} ({len(bolge_grubu)} Parça)", expanded=False):
                        for _, p_row in bolge_grubu.iterrows():
                            p_ad = p_row.get("parca_adi", "-")
                            p_sn = p_row.get("parca_sn", "-")
                            s_ad = p_row.get("sistem_adi", "-")
                            durum = p_row.get("durum", "-")
                            
                            st.markdown(f"**Sistem:** `{s_ad}` | **Parça:** `{p_ad}` | **SN:** `{p_sn}` | **Durum:** `{durum}`")
                            
                            d_adi = p_row.get("dosya_adi")
                            if pd.notna(d_adi) and d_adi:
                                d_yolu = os.path.join(UPLOAD_FOLDER, d_adi)
                                if os.path.exists(d_yolu):
                                    with open(d_yolu, "rb") as file_in:
                                        st.download_button(f"📥 Belgeyi İndir ({p_sn})", data=file_in, file_name=d_adi, key=f"dl_gecmis_{p_row['id']}")

                            conn = sqlite3.connect(DB_DOSYASI)
                            df_p_gecmis = pd.read_sql_query("SELECT tarih, islem, aciklama FROM parca_gecmis WHERE parca_sn = ? ORDER BY id DESC", conn, params=(p_sn,))
                            conn.close()
                            
                            if not df_p_gecmis.empty:
                                for _, g_row in df_p_gecmis.iterrows():
                                    st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;🕒 <small>{g_row['tarih']} — 📌 **{g_row['islem']}** : {g_row['aciklama']}</small>", unsafe_allow_html=True)
                            else:
                                st.markdown("&nbsp;&nbsp;&nbsp;&nbsp;<small style='color:gray;'>Geçmiş işlem kaydı bulunmuyor.</small>", unsafe_allow_html=True)
                                
                            st.markdown("<hr style='margin:5px 0; border-top: 1px dashed #ddd;'>", unsafe_allow_html=True)
            else:
                st.info("Kayıtlı bölge bulunamadı.")
        else:
            st.info("Henüz hiç parça kaydı bulunmuyor.")

    # 6. SEKME: YÖNETİM
    with tab_yonetim:
        st.subheader("Veri Yönetimi ve Raporlama")
        
        col_yonetim1, col_yonetim2 = st.columns(2)
        
        with col_yonetim1:
            st.markdown("##### 📤 Dışa Aktar")
            st.markdown("Mevcut tüm parça listesini Excel formatında bilgisayarınıza indirebilirsiniz.")
            if not df_parcalar.empty:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_parcalar.to_excel(writer, index=False, sheet_name='Veriler')
                st.download_button("📥 Excel Raporu İndir", data=output.getvalue(), file_name="sistem_takip_rapor.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("Dışa aktarılacak kayıt bulunmuyor.")
                
        with col_yonetim2:
            st.markdown("##### 📥 İçe Aktar (Excel / CSV)")
            st.markdown("Daha önceden dışa aktardığınız veya sisteme uygun başlıklarla hazırladığınız excel/csv dosyasını yükleyin.")
            
            yuklenen_excel = st.file_uploader("Dosya Seç", type=["xlsx", "xls", "csv"], key="excel_ice_aktarim")
            if yuklenen_excel is not None:
                if st.button("Verileri Sisteme Aktar", type="primary"):
                    try:
                        if yuklenen_excel.name.endswith('.csv'):
                            df_gelen = pd.read_csv(yuklenen_excel)
                        else:
                            df_gelen = pd.read_excel(yuklenen_excel)
                            
                        if 'id' in df_gelen.columns:
                            df_gelen = df_gelen.drop(columns=['id'])
                            
                        conn = sqlite3.connect(DB_DOSYASI)
                        cursor = conn.cursor()
                        
                        eklenen_sayisi = 0
                        for _, row in df_gelen.iterrows():
                            b_bolge = tr_upper(row.get("bolge", ""))
                            b_sistem = tr_upper(row.get("sistem_adi", ""))
                            b_s_pn = tr_upper(row.get("sistem_pn", ""))
                            b_s_sn = tr_upper(row.get("sistem_sn", ""))
                            b_parca = tr_upper(row.get("parca_adi", ""))
                            b_p_pn = tr_upper(row.get("parca_pn", ""))
                            b_p_sn = tr_upper(row.get("parca_sn", ""))
                            b_durum = str(row.get("durum", "FAAL"))
                            b_tarih = str(row.get("onarim_tarih", ""))
                            b_aciklama = tr_upper(row.get("aciklama", ""))
                            b_dosya = row.get("dosya_adi", None)
                            if pd.isna(b_dosya):
                                b_dosya = None
                                
                            if b_bolge and b_sistem and b_parca:
                                cursor.execute("""
                                    INSERT INTO parcalar (bolge, sistem_adi, sistem_pn, sistem_sn, parca_adi, parca_pn, parca_sn, durum, onarim_tarih, aciklama, dosya_adi)
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                """, (b_bolge, b_sistem, b_s_pn, b_s_sn, b_parca, b_p_pn, b_p_sn, b_durum, b_tarih, b_aciklama, b_dosya))
                                
                                cursor.execute("INSERT INTO parca_gecmis (parca_sn, tarih, islem, aciklama) VALUES (?, ?, ?, ?)", 
                                               (b_p_sn, datetime.datetime.now().strftime("%d.%m.%Y %H:%M"), f"EXCEL İÇE AKTARIM ({b_durum})", b_aciklama))
                                eklenen_sayisi += 1
                                
                        conn.commit()
                        conn.close()
                        log_yaz("İÇE AKTARIM", f"Excelden {eklenen_sayisi} adet kayıt başarıyla aktarıldı.")
                        st.success(f"Başarıyla {eklenen_sayisi} kayıt içe aktarıldı!")
                        st.rerun()
                    except Exception as ex:
                        st.error(f"Dosya içe aktarılırken hata oluştu: {ex}")

except Exception as e:
    st.error("Bir hata oluştu:")
    st.code(traceback.format_exc())