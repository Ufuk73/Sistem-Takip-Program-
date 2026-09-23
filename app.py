import datetime
import io
import os
import shutil
import sqlite3
import traceback
import pandas as pd
import streamlit as st

# Sayfa Yapılandırması
st.set_page_config(
    page_title="Sistem ve Parça Takip Sistemi", page_icon="⚙️", layout="wide"
)

try:
  DB_DOSYASI = "sistem_takip.db"
  UPLOAD_FOLDER = "yuklenen_dosyalar"
  AUTOSAVE_FOLDER = "otomatik_yedekler"

  for folder in [UPLOAD_FOLDER, AUTOSAVE_FOLDER]:
    if not os.path.exists(folder):
      os.makedirs(folder)

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

  # --- OTOMATİK EXCEL YEDEKLEME FONKSİYONU ---
  def otomatik_yedekle():
    """Her veri işleminde verilerin otomatik Excel yedeğini alır (Son 10 yedek saklanır)."""
    try:
      timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
      backup_path = os.path.join(AUTOSAVE_FOLDER, f"autosave_{timestamp}.xlsx")
      latest_backup = os.path.join(
          AUTOSAVE_FOLDER, "sistem_takip_latest_autosave.xlsx"
      )

      with sqlite3.connect(DB_DOSYASI) as conn:
        df_p = pd.read_sql_query("SELECT * FROM parcalar", conn)
        df_n = pd.read_sql_query("SELECT * FROM gunluk_notlar", conn)
        df_l = pd.read_sql_query("SELECT * FROM islem_loglari", conn)
        df_g = pd.read_sql_query("SELECT * FROM parca_gecmis", conn)

      with pd.ExcelWriter(backup_path, engine="openpyxl") as writer:
        df_p.to_excel(writer, sheet_name="Parcalar", index=False)
        df_n.to_excel(writer, sheet_name="Gunluk_Notlar", index=False)
        df_l.to_excel(writer, sheet_name="Loglar", index=False)
        df_g.to_excel(writer, sheet_name="Parca_Gecmisi", index=False)

      shutil.copy2(backup_path, latest_backup)

      # Son 10 otomatik Excel yedeğini tut, eskileri temizle
      yedekler = sorted(
          [
              os.path.join(AUTOSAVE_FOLDER, f)
              for f in os.listdir(AUTOSAVE_FOLDER)
              if f.startswith("autosave_") and f.endswith(".xlsx")
          ],
          key=os.path.getmtime,
      )
      if len(yedekler) > 10:
        for eski_yedek in yedekler[:-10]:
          os.remove(eski_yedek)
    except Exception as ex:
      print(f"Otomatik Excel yedekleme hatası: {ex}")

  # --- TARİH DÖNÜŞÜM YARDIMCILARI ---
  def str_to_date(tarih_str):
    if not tarih_str or pd.isna(tarih_str):
      return datetime.date.today()
    try:
      return datetime.datetime.strptime(
          str(tarih_str).strip(), "%d.%m.%Y"
      ).date()
    except ValueError:
      return datetime.date.today()

  # --- VERİTABANI BAŞLATMA VE MIGRATION ---
  def veritabanini_hazirla():
    with sqlite3.connect(DB_DOSYASI) as conn:
      cursor = conn.cursor()
      cursor.execute(
          "CREATE TABLE IF NOT EXISTS parcalar (id INTEGER PRIMARY KEY"
          " AUTOINCREMENT)"
      )
      cursor.execute(
          "CREATE TABLE IF NOT EXISTS islem_loglari (id INTEGER PRIMARY KEY"
          " AUTOINCREMENT, zaman TEXT, islem_turu TEXT, detay TEXT)"
      )
      cursor.execute(
          "CREATE TABLE IF NOT EXISTS gunluk_notlar (id INTEGER PRIMARY KEY"
          " AUTOINCREMENT, tarih TEXT, bolge TEXT, detay TEXT)"
      )
      cursor.execute(
          "CREATE TABLE IF NOT EXISTS parca_gecmis (id INTEGER PRIMARY KEY"
          " AUTOINCREMENT, parca_sn TEXT, tarih TEXT, islem TEXT, aciklama"
          " TEXT)"
      )

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
          "dosya_adi": "TEXT",
      }

      cursor.execute("PRAGMA table_info(parcalar)")
      mevcut_kolonlar = [kol[1] for kol in cursor.fetchall()]

      for kolon_adi, kolon_tipi in beklenen_kolonlar.items():
        if kolon_adi not in mevcut_kolonlar:
          cursor.execute(
              f"ALTER TABLE parcalar ADD COLUMN {kolon_adi} {kolon_tipi}"
          )

  def log_yaz(islem_turu, detay):
    try:
      zaman = datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")
      with sqlite3.connect(DB_DOSYASI) as conn:
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO islem_loglari (zaman, islem_turu, detay) VALUES (?,"
            " ?, ?)",
            (zaman, islem_turu, detay),
        )
      # Her işlem sonrası otomatik Excel yedeklemesi yap
      otomatik_yedekle()
    except Exception as e:
      print(f"Log yazılamadı: {e}")

  veritabanini_hazirla()

  # --- ÖZEL SİYAH TEMA VE ÇERÇEVE CSS ---
  st.markdown(
      """
    <style>
    .stApp {
        background-color: #121212;
        color: #e0e0e0;
    }
    .metric-card {
        background-color: #1e1e1e;
        border: 1px solid #333333;
        border-radius: 8px;
        padding: 12px 15px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }
    .metric-title {
        font-size: 13px;
        color: #a0a0a0;
        font-weight: 600;
        margin-bottom: 4px;
        text-transform: uppercase;
    }
    .metric-value {
        font-size: 22px;
        color: #ffffff;
        font-weight: 700;
    }
    .autosave-badge {
        font-size: 12px;
        color: #2ecc71;
        background-color: #1b382b;
        padding: 4px 8px;
        border-radius: 4px;
        border: 1px solid #2ecc71;
    }
    </style>
    """,
      unsafe_allow_html=True,
  )

  # --- ÜST BAŞLIK VE OTOMATİK KAYIT İNDİKATÖRÜ ---
  col_head1, col_head2 = st.columns([7, 3])
  with col_head1:
    st.markdown("### ⚙️ Sistem ve Parça Takip Sistemi")
  with col_head2:
    st.markdown(
        "<div style='text-align: right; padding-top: 10px;'><span"
        ' class="autosave-badge">🟢 Otomatik Excel Yedekleme Aktif</span></div>',
        unsafe_allow_html=True,
    )

  # Verileri Çek
  with sqlite3.connect(DB_DOSYASI) as conn:
    df_parcalar = pd.read_sql_query("SELECT * FROM parcalar", conn)

  # İstatistik Hesaplama
  toplam = len(df_parcalar)
  faal = (
      len(df_parcalar[df_parcalar["durum"] == "FAAL"])
      if not df_parcalar.empty and "durum" in df_parcalar.columns
      else 0
  )
  yedek = (
      len(df_parcalar[df_parcalar["durum"] == "YEDEK PARÇA"])
      if not df_parcalar.empty and "durum" in df_parcalar.columns
      else 0
  )
  onarimda = (
      len(df_parcalar[df_parcalar["durum"] == "ONARIMDA"])
      if not df_parcalar.empty and "durum" in df_parcalar.columns
      else 0
  )
  gayri = (
      len(df_parcalar[df_parcalar["durum"] == "GAYRI FAAL"])
      if not df_parcalar.empty and "durum" in df_parcalar.columns
      else 0
  )

  # Kritik onarım hesaplama (30 gün+)
  kritik = 0
  kritik_liste = []
  if (
      not df_parcalar.empty
      and "onarim_tarih" in df_parcalar.columns
      and "durum" in df_parcalar.columns
  ):
    bugun = datetime.date.today()
    for _, row in df_parcalar[df_parcalar["durum"] == "ONARIMDA"].iterrows():
      o_tarih = row["onarim_tarih"]
      if o_tarih:
        try:
          baslangic = datetime.datetime.strptime(
              str(o_tarih).strip(), "%d.%m.%Y"
          ).date()
          gecen_gun = (bugun - baslangic).days
          if gecen_gun >= 30:
            kritik += 1
            kritik_liste.append(
                f"• **Bölge: {row.get('bolge', 'Bölge Yok')}**"
                f" ({row.get('parca_adi', 'Parça')} - SN:"
                f" {row.get('parca_sn', '-')}) -> {gecen_gun} gündür onarımda!"
            )
        except ValueError:
          pass

  if kritik > 0:
    st.error(
        f"⚠️ **DİKKAT:** 30 Günü Aşan Onarımda Bekleyen **{kritik}** Adet Parça"
        " Bulunuyor!"
    )
    with st.expander("Kritik Parçaları Listele", expanded=False):
      for k_bilgi in kritik_liste:
        st.markdown(k_bilgi)

  # Metrikleri Çerçeveli Sütunlar Halinde Yerleştirme
  c1, c2, c3, c4, c5, c6 = st.columns(6)

  with c1:
    st.markdown(
        '<div class="metric-card"><div'
        ' class="metric-title">Toplam</div><div'
        f' class="metric-value">{toplam}</div></div>',
        unsafe_allow_html=True,
    )
  with c2:
    st.markdown(
        '<div class="metric-card"><div class="metric-title">Faal</div><div'
        f' class="metric-value" style="color: #2ecc71;">{faal}</div></div>',
        unsafe_allow_html=True,
    )
  with c3:
    st.markdown(
        '<div class="metric-card"><div class="metric-title">Yedek</div><div'
        f' class="metric-value" style="color: #3498db;">{yedek}</div></div>',
        unsafe_allow_html=True,
    )
  with c4:
    st.markdown(
        '<div class="metric-card"><div'
        ' class="metric-title">Onarımda</div><div class="metric-value"'
        f' style="color: #f1c40f;">{onarimda}</div></div>',
        unsafe_allow_html=True,
    )
  with c5:
    st.markdown(
        '<div class="metric-card"><div class="metric-title">Kritik'
        ' (30+)</div><div class="metric-value" style="color:'
        f' #e74c3c;">{kritik}</div></div>',
        unsafe_allow_html=True,
    )
  with c6:
    st.markdown(
        '<div class="metric-card"><div class="metric-title">Gayri'
        ' Faal</div><div class="metric-value" style="color:'
        f' #95a5a6;">{gayri}</div></div>',
        unsafe_allow_html=True,
    )

  st.markdown(
      "<hr style='margin:15px 0 10px 0; border-color: #333333;'>",
      unsafe_allow_html=True,
  )

  tab_takip, tab_ekle, tab_notlar, tab_loglar, tab_gecmis, tab_yonetim = st.tabs(
      [
          "📋 Kart Görünümü",
          "➕ Yeni Kayıt",
          "📝 Günlük Notlar",
          "📜 Loglar",
          "🔍 Parça Geçmişi",
          "⚙️ Yönetim",
      ]
  )

  # Dialog ile Düzenleme Fonksiyonu
  @st.dialog("Kayıt Düzenle", width="large")
  def duzenle_dialog(r_id):
    with sqlite3.connect(DB_DOSYASI) as conn:
      df_tekil = pd.read_sql_query(
          "SELECT * FROM parcalar WHERE id = ?", conn, params=(r_id,)
      )

    if not df_tekil.empty:
      row = df_tekil.iloc[0]
      with st.form(key=f"dialog_form_{r_id}"):
        e_bolge = st.text_input(
            "Bölge",
            value=(
                row.get("bolge", "") if pd.notna(row.get("bolge", "")) else ""
            ),
        )
        e_sistem = st.text_input(
            "Sistem Adı",
            value=(
                row.get("sistem_adi", "")
                if pd.notna(row.get("sistem_adi", ""))
                else ""
            ),
        )

        col_d1, col_d2 = st.columns(2)
        e_s_pn = col_d1.text_input(
            "Sistem PN",
            value=(
                row.get("sistem_pn", "")
                if pd.notna(row.get("sistem_pn", ""))
                else ""
            ),
        )
        e_s_sn = col_d2.text_input(
            "Sistem SN",
            value=(
                row.get("sistem_sn", "")
                if pd.notna(row.get("sistem_sn", ""))
                else ""
            ),
        )

        e_parca = st.text_input(
            "Parça Adı",
            value=(
                row.get("parca_adi", "")
                if pd.notna(row.get("parca_adi", ""))
                else ""
            ),
        )

        col_d3, col_d4 = st.columns(2)
        e_p_pn = col_d3.text_input(
            "Parça PN",
            value=(
                row.get("parca_pn", "")
                if pd.notna(row.get("parca_pn", ""))
                else ""
            ),
        )
        e_p_sn = col_d4.text_input(
            "Parça SN",
            value=(
                row.get("parca_sn", "")
                if pd.notna(row.get("parca_sn", ""))
                else ""
            ),
        )

        mevcut_d = row.get("durum", "FAAL")
        d_list = ["FAAL", "YEDEK PARÇA", "ONARIMDA", "GAYRI FAAL"]
        d_idx = d_list.index(mevcut_d) if mevcut_d in d_list else 0
        e_durum = st.selectbox("Durum", d_list, index=d_idx)

        varsayilan_tarih = str_to_date(row.get("onarim_tarih"))
        e_tarih_obj = st.date_input(
            "Onarım Tarihi", value=varsayilan_tarih, format="DD.MM.YYYY"
        )
        e_tarih_str = e_tarih_obj.strftime("%d.%m.%Y")

        e_aciklama = st.text_area(
            "Açıklama / Not",
            value=(
                row.get("aciklama", "")
                if pd.notna(row.get("aciklama", ""))
                else ""
            ),
        )

        if st.form_submit_button("Güncellemeyi Kaydet", type="primary"):
          with sqlite3.connect(DB_DOSYASI) as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                            UPDATE parcalar SET bolge=?, sistem_adi=?, sistem_pn=?, sistem_sn=?, parca_adi=?, parca_pn=?, parca_sn=?, durum=?, onarim_tarih=?, aciklama=?
                            WHERE id=?
                        """,
                (
                    tr_upper(e_bolge),
                    tr_upper(e_sistem),
                    tr_upper(e_s_pn),
                    tr_upper(e_s_sn),
                    tr_upper(e_parca),
                    tr_upper(e_p_pn),
                    tr_upper(e_p_sn),
                    e_durum,
                    e_tarih_str,
                    tr_upper(e_aciklama),
                    r_id,
                ),
            )

            cursor.execute(
                "INSERT INTO parca_gecmis (parca_sn, tarih, islem, aciklama)"
                " VALUES (?, ?, ?, ?)",
                (
                    tr_upper(e_p_sn),
                    datetime.datetime.now().strftime("%d.%m.%Y %H:%M"),
                    f"GÜNCELLEME ({e_durum})",
                    tr_upper(e_aciklama),
                ),
            )

          log_yaz("GÜNCELLEME", f"ID {r_id} güncellendi.")
          st.success("Kayıt güncellendi ve Excel yedeği alındı!")
          st.rerun()

  # Silme Onay Dialogu
  @st.dialog("Silme Onayı")
  def sil_onay_dialog(r_id, parca_adi, parca_sn):
    st.warning(
        f"**{parca_adi}** (SN: `{parca_sn}`) kaydını silmek istediğinizden emin"
        " misiniz?"
    )
    col_s1, col_s2 = st.columns(2)
    if col_s1.button("Evet, Sil", type="primary", use_container_width=True):
      with sqlite3.connect(DB_DOSYASI) as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM parcalar WHERE id=?", (r_id,))
      log_yaz("SİLME", f"ID {r_id} ({parca_adi}) silindi.")
      st.success("Kayıt başarıyla silindi ve Excel yedeği alındı.")
      st.rerun()
    if col_s2.button("İptal", use_container_width=True):
      st.rerun()

  # 1. SEKME: TAKİP & FİLTRELEME (KART GÖRÜNÜMÜ)
  with tab_takip:
    if not df_parcalar.empty:
      col_f1, col_f2, col_f3, col_f4 = st.columns(4)

      bolgeler = (
          ["TÜMÜ"] + list(df_parcalar["bolge"].dropna().unique())
          if "bolge" in df_parcalar.columns
          else ["TÜMÜ"]
      )
      parcalar = (
          ["TÜMÜ"] + list(df_parcalar["parca_adi"].dropna().unique())
          if "parca_adi" in df_parcalar.columns
          else ["TÜMÜ"]
      )
      durumlar = [
          "TÜMÜ",
          "FAAL",
          "YEDEK PARÇA",
          "ONARIMDA",
          "GAYRI FAAL",
          "30 GÜN+ KRİTİK",
      ]

      f_bolge = col_f1.selectbox("Bölge", bolgeler)
      f_parca = col_f2.selectbox("Parça Adı", parcalar)
      f_durum = col_f3.selectbox("Durum", durumlar)
      f_arama = col_f4.text_input("Hızlı Arama")

      filt_df = df_parcalar.copy()
      if f_bolge != "TÜMÜ" and "bolge" in filt_df.columns:
        filt_df = filt_df[filt_df["bolge"] == f_bolge]
      if f_parca != "TÜMÜ" and "parca_adi" in filt_df.columns:
        filt_df = filt_df[filt_df["parca_adi"] == f_parca]
      if (
          f_durum != "TÜMÜ"
          and f_durum != "30 GÜN+ KRİTİK"
          and "durum" in filt_df.columns
      ):
        filt_df = filt_df[filt_df["durum"] == f_durum]
      elif (
          f_durum == "30 GÜN+ KRİTİK"
          and "durum" in filt_df.columns
          and "onarim_tarih" in filt_df.columns
      ):
        kritik_idler = []
        bugun = datetime.date.today()
        for _, r in filt_df[filt_df["durum"] == "ONARIMDA"].iterrows():
          if r["onarim_tarih"]:
            try:
              b_t = datetime.datetime.strptime(
                  str(r["onarim_tarih"]).strip(), "%d.%m.%Y"
              ).date()
              if (bugun - b_t).days >= 30:
                kritik_idler.append(r["id"])
            except ValueError:
              pass
        filt_df = filt_df[filt_df["id"].isin(kritik_idler)]

      if f_arama:
        a_upper = tr_upper(f_arama)
        filt_df = filt_df[
            filt_df.apply(
                lambda row: row.astype(str)
                .str.upper()
                .str.contains(a_upper)
                .any(),
                axis=1,
            )
        ]

      st.markdown("<hr style='margin:5px 0;'>", unsafe_allow_html=True)

      if not filt_df.empty:
        col_exp1, col_exp2 = st.columns([7, 3])
        with col_exp2:
          tarih_str = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
          not_icerik = "SİSTEM VE PARÇA LİSTESİ NOTU\n"
          not_icerik += f"Tarih: {tarih_str}\n"
          not_icerik += "--------------------------------------------------\n"
          not_icerik += (
              f"Seçilen Filtreler -> Bölge: {f_bolge} | Parça: {f_parca} |"
              f" Durum: {f_durum}"
          )
          if f_arama:
            not_icerik += f" | Arama: {f_arama}"
          not_icerik += f"\nToplam Kayıt: {len(filt_df)}\n"
          not_icerik += "--------------------------------------------------\n\n"

          for idx, (_, row) in enumerate(filt_df.iterrows(), 1):
            not_icerik += f"{idx}. Bölge: {row.get('bolge', '-')}\n"
            not_icerik += (
                f"   Sistem: {row.get('sistem_adi', '-')} (PN:"
                f" {row.get('sistem_pn', '-')}, SN: {row.get('sistem_sn', '-')})\n"
            )
            not_icerik += (
                f"   Parça : {row.get('parca_adi', '-')} (PN:"
                f" {row.get('parca_pn', '-')}, SN: {row.get('parca_sn', '-')})\n"
            )
            not_icerik += (
                f"   Durum : {row.get('durum', '-')} | Onarım Tar.:"
                f" {row.get('onarim_tarih', '-')}\n"
            )
            if pd.notna(row.get("aciklama")) and str(
                row.get("aciklama")
            ).strip():
              not_icerik += f"   Not   : {row.get('aciklama')}\n"
            not_icerik += (
                "--------------------------------------------------\n"
            )

          st.download_button(
              "📝 Not Olarak İndir",
              data=not_icerik.encode("utf-8"),
              file_name=(
                  "sistem_notu_"
                  f"{datetime.date.today().strftime('%d_%m_%Y')}.txt"
              ),
              mime="text/plain",
              help="Filtrelenen listeyi sade bir not dosyası olarak indirir.",
          )

      if not filt_df.empty and "bolge" in filt_df.columns:
        for b_adi, b_grubu in filt_df.groupby("bolge"):
          with st.expander(
              f"📍 BÖLGE: {b_adi} ({len(b_grubu)} Kayıt)", expanded=False
          ):
            h1, h2, h3, h4, h5 = st.columns([2.2, 2.2, 1.8, 0.9, 0.9])
            h1.markdown(
                "<small><b>SİSTEM ADI</b></small>", unsafe_allow_html=True
            )
            h2.markdown(
                "<small><b>PARÇA & SERİ NO</b></small>", unsafe_allow_html=True
            )
            h3.markdown(
                "<small><b>DURUM / TARİH</b></small>", unsafe_allow_html=True
            )
            h4.markdown(
                "<small><b>DÜZENLE</b></small>", unsafe_allow_html=True
            )
            h5.markdown("<small><b>SİL</b></small>", unsafe_allow_html=True)
            st.markdown(
                "<hr style='margin:2px 0 5px 0; border-color: #333;'>",
                unsafe_allow_html=True,
            )

            for _, row in b_grubu.iterrows():
              r_id = row["id"]
              durum = row.get("durum", "FAAL")
              durum_badge = (
                  f"🟢 {durum}"
                  if durum == "FAAL"
                  else (
                      f"🟡 {durum}"
                      if durum == "ONARIMDA"
                      else (
                          f"🔴 {durum}"
                          if durum == "GAYRI FAAL"
                          else f"🔵 {durum}"
                      )
                  )
              )

              c1, c2, c3, c4, c5 = st.columns([2.2, 2.2, 1.8, 0.9, 0.9])

              with c1:
                st.markdown(
                    "<div style='line-height: 1.1;'><small><b>"
                    f"{row.get('sistem_adi', '-')}"
                    "</b><br><span style='color:#a0a0a0;'>PN:"
                    f" {row.get('sistem_pn', '-')}</span></small></div>",
                    unsafe_allow_html=True,
                )
              with c2:
                st.markdown(
                    "<div style='line-height: 1.1;'><small>"
                    f"{row.get('parca_adi', '-')}"
                    "<br><span style='color:#a0a0a0;'>SN:"
                    f" {row.get('parca_sn', '-')}</span></small></div>",
                    unsafe_allow_html=True,
                )
              with c3:
                st.markdown(
                    "<div style='line-height: 1.1;'><small>"
                    f"{durum_badge}<br><span style='color:#a0a0a0;'>"
                    f"{row.get('onarim_tarih', '-')}</span></small></div>",
                    unsafe_allow_html=True,
                )

              with c4:
                if st.button("✏️", key=f"edit_{r_id}", help="Düzenle"):
                  duzenle_dialog(r_id)

              with c5:
                if st.button("🗑️", key=f"del_{r_id}", help="Sil"):
                  sil_onay_dialog(
                      r_id,
                      row.get("parca_adi", "Parça"),
                      row.get("parca_sn", "-"),
                  )

              st.markdown(
                  "<hr style='margin:3px 0; border-color: #333;'>",
                  unsafe_allow_html=True,
              )
      else:
        st.info("Filtreleme kriterlerine uygun kayıt bulunamadı.")
    else:
      st.info("Kayıt bulunmuyor.")

  # 2. SEKME: YENİ KAYIT EKLE
  with tab_ekle:
    kayitli_bolgeler = (
        sorted(list(df_parcalar["bolge"].dropna().unique()))
        if not df_parcalar.empty and "bolge" in df_parcalar.columns
        else []
    )

    secim_tipi = st.radio(
        "Bölge Giriş",
        ["Kayıtlı Seç", "Yeni Yaz"],
        horizontal=True,
        key="form_secim_tipi",
    )

    aktif_bolge = ""
    if secim_tipi == "Kayıtlı Seç" and len(kayitli_bolgeler) > 0:
      aktif_bolge = st.selectbox(
          "Bölge Seç", kayitli_bolgeler, key="form_secilen_bolge"
      )
    else:
      aktif_bolge = st.text_input("Yeni Bölge Adı", key="form_yeni_bolge")

    varsayilan_sistem = ""
    varsayilan_s_pn = ""
    varsayilan_s_sn = ""

    if aktif_bolge and not df_parcalar.empty and "bolge" in df_parcalar.columns:
      eslesen_kayitlar = df_parcalar[
          df_parcalar["bolge"].str.upper() == str(aktif_bolge).upper()
      ]
      if not eslesen_kayitlar.empty:
        son_kayit = eslesen_kayitlar.iloc[-1]
        varsayilan_sistem = (
            son_kayit.get("sistem_adi", "")
            if pd.notna(son_kayit.get("sistem_adi", ""))
            else ""
        )
        varsayilan_s_pn = (
            son_kayit.get("sistem_pn", "")
            if pd.notna(son_kayit.get("sistem_pn", ""))
            else ""
        )
        varsayilan_s_sn = (
            son_kayit.get("sistem_sn", "")
            if pd.notna(son_kayit.get("sistem_sn", ""))
            else ""
        )

    with st.form("yeni_kayit_formu", clear_on_submit=True):
      st.markdown(
          "📌 **Seçilen/Girilen Bölge:**"
          f" `{aktif_bolge if aktif_bolge else 'Henüz seçilmedi'}`"
      )

      e_sistem = st.text_input("Sistem Adı", value=varsayilan_sistem)
      c1, c2 = st.columns(2)
      e_s_pn = c1.text_input("Sistem PN", value=varsayilan_s_pn)
      e_s_sn = c2.text_input("Sistem SN", value=varsayilan_s_sn)

      e_parca = st.text_input("Parça Adı")
      c3, c4 = st.columns(2)
      e_p_pn = c3.text_input("Parça PN", value="")
      e_p_sn = c4.text_input("Parça SN", value="")

      e_durum = st.selectbox(
          "Durum", ["FAAL", "YEDEK PARÇA", "ONARIMDA", "GAYRI FAAL"]
      )

      e_tarih_obj = st.date_input(
          "Onarım Tarihi", value=datetime.date.today(), format="DD.MM.YYYY"
      )
      e_tarih_str = e_tarih_obj.strftime("%d.%m.%Y")

      e_aciklama = st.text_area("Açıklama")
      yuklenen_dosya_form = st.file_uploader(
          "Belge/Fotoğraf", type=["png", "jpg", "jpeg", "pdf"]
      )

      if st.form_submit_button("Sisteme Kaydet", type="primary"):
        if not aktif_bolge or not e_sistem or not e_parca:
          st.warning("Bölge, Sistem Adı ve Parça Adı zorunludur!")
        else:
          dosya_ismi = None
          if yuklenen_dosya_form is not None:
            dosya_ismi = (
                f"{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{yuklenen_dosya_form.name}"
            )
            with open(os.path.join(UPLOAD_FOLDER, dosya_ismi), "wb") as f:
              f.write(yuklenen_dosya_form.getbuffer())

          with sqlite3.connect(DB_DOSYASI) as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                            INSERT INTO parcalar (bolge, sistem_adi, sistem_pn, sistem_sn, parca_adi, parca_pn, parca_sn, durum, onarim_tarih, aciklama, dosya_adi)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                (
                    tr_upper(aktif_bolge),
                    tr_upper(e_sistem),
                    tr_upper(e_s_pn),
                    tr_upper(e_s_sn),
                    tr_upper(e_parca),
                    tr_upper(e_p_pn),
                    tr_upper(e_p_sn),
                    e_durum,
                    e_tarih_str,
                    tr_upper(e_aciklama),
                    dosya_ismi,
                ),
            )

            cursor.execute(
                "INSERT INTO parca_gecmis (parca_sn, tarih, islem, aciklama)"
                " VALUES (?, ?, ?, ?)",
                (
                    tr_upper(e_p_sn),
                    datetime.datetime.now().strftime("%d.%m.%Y %H:%M"),
                    f"İLK KAYIT ({e_durum})",
                    tr_upper(e_aciklama),
                ),
            )

          log_yaz(
              "YENİ KAYIT", f"Bölge: {aktif_bolge}, Sistem: {e_sistem} eklendi."
          )
          st.success("Kayıt eklendi ve otomatik Excel yedeği alındı!")
          st.rerun()

  # 3. SEKME: GÜNLÜK NOTLAR
  with tab_notlar:
    kayitli_not_bolgeler = (
        sorted(list(df_parcalar["bolge"].dropna().unique()))
        if not df_parcalar.empty and "bolge" in df_parcalar.columns
        else []
    )
    with st.form("not_form", clear_on_submit=True):
      n_tarih_obj = st.date_input(
          "Tarih", value=datetime.date.today(), format="DD.MM.YYYY"
      )
      n_tarih_str = n_tarih_obj.strftime("%d.%m.%Y")
      n_bolge = (
          st.selectbox("Bölge", kayitli_not_bolgeler)
          if len(kayitli_not_bolgeler) > 0
          else st.text_input("Bölge")
      )
      n_detay = st.text_area("Not Detayı")
      if st.form_submit_button("Günlük Not Ekle"):
        with sqlite3.connect(DB_DOSYASI) as conn:
          cursor = conn.cursor()
          cursor.execute(
              "INSERT INTO gunluk_notlar (tarih, bolge, detay) VALUES (?,"
              " ?, ?)",
              (n_tarih_str, tr_upper(n_bolge), tr_upper(n_detay)),
          )
        log_yaz("GÜNLÜK NOT", f"Bölge: {n_bolge} için not eklendi.")
        st.success("Günlük not eklendi!")
        st.rerun()

    with sqlite3.connect(DB_DOSYASI) as conn:
      df_notlar = pd.read_sql_query(
          "SELECT * FROM gunluk_notlar ORDER BY id DESC", conn
      )
    if not df_notlar.empty:
      st.dataframe(df_notlar, use_container_width=True, hide_index=True)

  # 4. SEKME: LOGLAR
  with tab_loglar:
    with sqlite3.connect(DB_DOSYASI) as conn:
      df_loglar = pd.read_sql_query(
          "SELECT * FROM islem_loglari ORDER BY id DESC", conn
      )
    if not df_loglar.empty:
      st.dataframe(df_loglar, use_container_width=True, hide_index=True)

  # 5. SEKME: PARÇA GEÇMİŞİ
  with tab_gecmis:
    st.subheader("Bölge Bazlı Parça ve İşlem Geçmişi")

    if not df_parcalar.empty and "bolge" in df_parcalar.columns:
      benzersiz_bolgeler = sorted(list(df_parcalar["bolge"].dropna().unique()))

      if benzersiz_bolgeler:
        secilen_gecmis_bolge = st.selectbox(
            "İncelemek İçin Bölge Seçin",
            ["TÜM BÖLGELER"] + benzersiz_bolgeler,
        )

        bolge_filtreli_df = df_parcalar.copy()
        if secilen_gecmis_bolge != "TÜM BÖLGELER":
          bolge_filtreli_df = bolge_filtreli_df[
              bolge_filtreli_df["bolge"] == secilen_gecmis_bolge
          ]

        for bolge_adi, bolge_grubu in bolge_filtreli_df.groupby("bolge"):
          with st.expander(
              f"📍 BÖLGE: {bolge_adi} ({len(bolge_grubu)} Parça)",
              expanded=False,
          ):
            for _, p_row in bolge_grubu.iterrows():
              p_ad = p_row.get("parca_adi", "-")
              p_sn = p_row.get("parca_sn", "-")
              s_ad = p_row.get("sistem_adi", "-")
              durum = p_row.get("durum", "-")

              st.markdown(
                  f"**Sistem:** `{s_ad}` | **Parça:** `{p_ad}` | **SN:**"
                  f" `{p_sn}` | **Durum:** `{durum}`"
              )

              d_adi = p_row.get("dosya_adi")
              if pd.notna(d_adi) and d_adi:
                d_yolu = os.path.join(UPLOAD_FOLDER, d_adi)
                if os.path.exists(d_yolu):
                  with open(d_yolu, "rb") as file_in:
                    st.download_button(
                        f"📥 Belgeyi İndir ({p_sn})",
                        data=file_in,
                        file_name=d_adi,
                        key=f"dl_gecmis_{p_row['id']}",
                    )

              with sqlite3.connect(DB_DOSYASI) as conn:
                df_p_gecmis = pd.read_sql_query(
                    "SELECT tarih, islem, aciklama FROM parca_gecmis WHERE"
                    " parca_sn = ? ORDER BY id DESC",
                    conn,
                    params=(p_sn,),
                )

              if not df_p_gecmis.empty:
                for _, g_row in df_p_gecmis.iterrows():
                  st.markdown(
                      f"&nbsp;&nbsp;&nbsp;&nbsp;🕒 <small>{g_row['tarih']} — 📌"
                      f" **{g_row['islem']}** : {g_row['aciklama']}</small>",
                      unsafe_allow_html=True,
                  )
              else:
                st.markdown(
                    "&nbsp;&nbsp;&nbsp;&nbsp;<small"
                    " style='color:#a0a0a0;'>Geçmiş işlem kaydı"
                    " bulunmuyor.</small>",
                    unsafe_allow_html=True,
                )

              st.markdown(
                  "<hr style='margin:5px 0; border-top: 1px dashed #444;'>",
                  unsafe_allow_html=True,
              )
      else:
        st.info("Kayıtlı bölge bulunamadı.")
    else:
      st.info("Henüz hiç parça kaydı bulunmuyor.")

  # 6. SEKME: YÖNETİM & OTOMATİK EXCEL YEDEK YÖNETİMİ
  with tab_yonetim:
    st.subheader("Veri Yönetimi, Raporlama ve Otomatik Excel Yedekleri")

    col_yonetim1, col_yonetim2 = st.columns(2)

    with col_yonetim1:
      st.markdown("##### 📤 Manuel Dışa Aktar & Yedek Al")
      st.markdown(
          "Mevcut parça listesini **Excel** olarak indirebilir veya"
          " veritabanının tam **Yedeğini (.db)** alabilirsiniz."
      )

      if not df_parcalar.empty:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          df_parcalar.to_excel(writer, index=False, sheet_name="Veriler")
        st.download_button(
            "📥 Güncel Excel Raporu İndir",
            data=output.getvalue(),
            file_name="sistem_takip_rapor.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      else:
        st.info("Dışa aktarılacak kayıt bulunmuyor.")

      st.markdown(
          "<hr style='margin:10px 0;'>", unsafe_allow_html=True
      )

      # --- VERİTABANI YEDEĞİ İNDİR ---
      if os.path.exists(DB_DOSYASI):
        with open(DB_DOSYASI, "rb") as db_file:
          st.download_button(
              "💾 Veritabanı Yedeği İndir (.db)",
              data=db_file,
              file_name=(
                  "sistem_takip_yedek_"
                  f"{datetime.date.today().strftime('%d_%m_%Y')}.db"
              ),
              mime="application/x-sqlite3",
              help="Veritabanının tam kopyasını indirir.",
          )

      # --- OTOMATİK EXCEL YEDEKLERİNİ LİSTELE / İNDİR ---
      st.markdown(
          "<hr style='margin:10px 0;'>", unsafe_allow_html=True
      )
      st.markdown("##### 🟢 Otomatik Alınan Son Excel Yedekleri")
      if os.path.exists(AUTOSAVE_FOLDER):
        oto_yedek_listesi = sorted(
            [
                f
                for f in os.listdir(AUTOSAVE_FOLDER)
                if f.startswith("autosave_") and f.endswith(".xlsx")
            ],
            reverse=True,
        )
        if oto_yedek_listesi:
          secilen_oto_yedek = st.selectbox(
              "Excel Yedek Seçin", oto_yedek_listesi[:10]
          )
          y_path = os.path.join(AUTOSAVE_FOLDER, secilen_oto_yedek)
          with open(y_path, "rb") as y_file:
            st.download_button(
                f"📥 Otomatik Excel Yedeğini İndir ({secilen_oto_yedek})",
                data=y_file,
                file_name=secilen_oto_yedek,
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
        else:
          st.caption("Henüz otomatik Excel yedeği oluşturulmadı.")

    with col_yonetim2:
      st.markdown("##### 📥 İçe Aktar & Geri Yükle")

      tab_ice_excel, tab_ice_db = st.tabs(
          ["Excel / CSV Yükle", "Veritabanı (.db) Geri Yükle"]
      )

      with tab_ice_excel:
        yuklenen_excel = st.file_uploader(
            "Excel Dosyası Seç",
            type=["xlsx", "xls", "csv"],
            key="excel_ice_aktarim",
        )
        if yuklenen_excel is not None:
          if st.button("Verileri Sisteme Aktar", type="primary"):
            try:
              if yuklenen_excel.name.endswith(".csv"):
                df_gelen = pd.read_csv(yuklenen_excel)
              else:
                df_gelen = pd.read_excel(yuklenen_excel)

              if "id" in df_gelen.columns:
                df_gelen = df_gelen.drop(columns=["id"])

              eklenen_sayisi = 0
              with sqlite3.connect(DB_DOSYASI) as conn:
                cursor = conn.cursor()
                for _, row in df_gelen.iterrows():
                  bolge = tr_upper(row.get("bolge", ""))
                  sistem_adi = tr_upper(row.get("sistem_adi", ""))
                  sistem_pn = tr_upper(row.get("sistem_pn", ""))
                  sistem_sn = tr_upper(row.get("sistem_sn", ""))
                  parca_adi = tr_upper(row.get("parca_adi", ""))
                  parca_pn = tr_upper(row.get("parca_pn", ""))
                  parca_sn = tr_upper(row.get("parca_sn", ""))
                  durum = str(row.get("durum", "FAAL")).strip()
                  if durum not in [
                      "FAAL",
                      "YEDEK PARÇA",
                      "ONARIMDA",
                      "GAYRI FAAL",
                  ]:
                    durum = "FAAL"

                  o_tar = row.get("onarim_tarih")
                  onarim_tarih = (
                      str(o_tar).strip()
                      if pd.notna(o_tar)
                      else datetime.date.today().strftime("%d.%m.%Y")
                  )

                  aciklama = tr_upper(
                      row.get("aciklama", "")
                      if pd.notna(row.get("aciklama"))
                      else ""
                  )
                  dosya_adi = (
                      str(row.get("dosya_adi", "")).strip()
                      if pd.notna(row.get("dosya_adi"))
                      else None
                  )

                  cursor.execute(
                      """
                                        INSERT INTO parcalar (bolge, sistem_adi, sistem_pn, sistem_sn, parca_adi, parca_pn, parca_sn, durum, onarim_tarih, aciklama, dosya_adi)
                                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                    """,
                      (
                          bolge,
                          sistem_adi,
                          sistem_pn,
                          sistem_sn,
                          parca_adi,
                          parca_pn,
                          parca_sn,
                          durum,
                          onarim_tarih,
                          aciklama,
                          dosya_adi,
                      ),
                  )

                  cursor.execute(
                      "INSERT INTO parca_gecmis (parca_sn, tarih, islem,"
                      " aciklama) VALUES (?, ?, ?, ?)",
                      (
                          parca_sn,
                          datetime.datetime.now().strftime("%d.%m.%Y %H:%M"),
                          f"İÇE AKTARILDI ({durum})",
                          aciklama,
                      ),
                  )
                  eklenen_sayisi += 1

              log_yaz(
                  "İÇE AKTARMA",
                  f"Excel/CSV ile {eklenen_sayisi} kayıt eklendi.",
              )
              st.success(
                  f"{eklenen_sayisi} adet kayıt başarıyla aktarıldı ve Excel"
                  " yedeği alındı!"
              )
              st.rerun()
            except Exception as ex:
              st.error(f"Aktarım sırasında bir hata oluştu: {ex}")

      with tab_ice_db:
        uploaded_db = st.file_uploader(
            "Yedek Veritabanı Dosyası Seç (.db)",
            type=["db", "sqlite", "sqlite3"],
            key="db_ice_aktarim",
        )
        if uploaded_db is not None:
          st.warning(
              "⚠️ Veritabanını geri yüklemek mevcut verilerin üzerine yazılmasına"
              " neden olabilir. Devam etmek istediğinize emin misiniz?"
          )
          if st.button("Veritabanını Geri Yükle", type="primary"):
            try:
              with open(DB_DOSYASI, "wb") as f:
                f.write(uploaded_db.getbuffer())
              log_yaz("YEDEK GERİ YÜKLEME", "Veritabanı yedekten geri yüklendi.")
              st.success(
                  "Veritabanı başarıyla geri yüklendi! Sayfa yenileniyor..."
              )
              st.rerun()
            except Exception as ex:
              st.error(f"Geri yükleme sırasında hata oluştu: {ex}")

except Exception as e:
  st.error("Uygulama çalışırken beklenmeyen bir hata oluştu:")
  st.code(traceback.format_exc())