import time
import os
import sys
import streamlit as st
import pandas as pd
from datetime import datetime
from fpdf import FPDF  # fpdf2 kütüphanesi yüklü olmalıdır: pip install fpdf2
import plotly.express as px
import re
import shutil

# --- GENEL AYARLAR ---
APARTMAN_ADI = "SALİM APARTMANI YÖNETİM PLANI (2026)" 
EXCEL_FILE = "apartman_verileri.xlsx" 
BELGE_KLASORU = "Apartman Yönetim"

def klasor_olustur():
    """Belge klasörünü oluştur"""
    if not os.path.exists(BELGE_KLASORU):
        os.makedirs(BELGE_KLASORU)

def verileri_yukle(): 
    if os.path.exists(EXCEL_FILE): 
        try: 
            sheets = pd.read_excel(EXCEL_FILE, sheet_name=None) 
            for s_name in ["Gelirler", "Giderler"]: 
                if s_name in sheets: 
                    df = sheets[s_name] 
                    if isinstance(df, pd.DataFrame) and not df.empty: 
                        if "Tarih" in df.columns: 
                            df["Tarih"] = pd.to_datetime(df["Tarih"], errors='coerce') 
                            df = df.dropna(subset=["Tarih"]) 
                            df = df.sort_values(by="Tarih").reset_index(drop=True) 
                            sheets[s_name] = df 
            
            # Giderler tablosunda 'Belge' sütun kontrolü
            if "Giderler" in sheets and "Belge" not in sheets["Giderler"].columns:
                sheets["Giderler"]["Belge"] = "YOK"

            if "Sakinler" not in sheets: 
                sheets["Sakinler"] = pd.DataFrame([{"No": i, "Ad": "Sakin", "Soyad": str(i), "Sifre": f"123{i}"} for i in range(1, 9)]) 
            else:
                if "Sifre" not in sheets["Sakinler"].columns:
                    sheets["Sakinler"]["Sifre"] = [f"123{i}" for i in sheets["Sakinler"]["No"]]
            
            if "Demirbaslar" not in sheets:
                sheets["Demirbaslar"] = pd.DataFrame(columns=["Demirbas Adi", "Adet", "Alis Tarihi", "Durumu", "Aciklama"])

            if "Ayarlar" not in sheets: 
                sheets["Ayarlar"] = pd.DataFrame([ 
                    {"Tip": "Daire", "Tutar": 500}, 
                    {"Tip": "İşyeri", "Tutar": 250}, 
                    {"Tip": "Devir", "Tutar": 0},
                    {"Tip": "GecikmeFaiziOrani", "Tutar": 5},
                    {"Tip": "YoneticiSifresi", "Tutar": "sahin2026"},
                    {"Tip": "AktifDuyuru", "Tutar": "Apartman sakinlerimize önemle duyurulur: Aidatların son ödeme gününe riayet edilmesi rica olunur."} 
                ]) 
            else:
                df_ayar = sheets["Ayarlar"]
                if not any(df_ayar["Tip"] == "GecikmeFaiziOrani"):
                    yeni_ayar = pd.DataFrame([{"Tip": "GecikmeFaiziOrani", "Tutar": 5}])
                    sheets["Ayarlar"] = pd.concat([df_ayar, yeni_ayar], ignore_index=True)
                if not any(df_ayar["Tip"] == "YoneticiSifresi"):
                    yeni_ayar = pd.DataFrame([{"Tip": "YoneticiSifresi", "Tutar": "sahin2026"}])
                    sheets["Ayarlar"] = pd.concat([df_ayar, yeni_ayar], ignore_index=True)
                if not any(df_ayar["Tip"] == "AktifDuyuru"):
                    yeni_ayar = pd.DataFrame([{"Tip": "AktifDuyuru", "Tutar": "Apartman sakinlerimize önemle duyurulur."}])
                    sheets["Ayarlar"] = pd.concat([df_ayar, yeni_ayar], ignore_index=True)

            return sheets 
        except Exception as e: 
            st.error(f"Dosya okuma hatası: {e}") 
            pass 
            
    return { 
        "Gelirler": pd.DataFrame(columns=["Tarih", "No", "Miktar", "Tur"]), 
        "Giderler": pd.DataFrame(columns=["Tarih", "Kategori", "Aciklama", "Miktar", "Belge"]), 
        "Ayarlar": pd.DataFrame([ 
            {"Tip": "Daire", "Tutar": 500}, 
            {"Tip": "İşyeri", "Tutar": 250}, 
            {"Tip": "Devir", "Tutar": 0},
            {"Tip": "GecikmeFaiziOrani", "Tutar": 5},
            {"Tip": "YoneticiSifresi", "Tutar": "sahin2026"},
            {"Tip": "AktifDuyuru", "Tutar": "Apartman sakinlerimize önemle duyurulur."}
        ]), 
        "Sakinler": pd.DataFrame([{"No": i, "Ad": "Sakin", "Soyad": str(i), "Sifre": f"123{i}"} for i in range(1, 9)]),
        "Demirbaslar": pd.DataFrame(columns=["Demirbas Adi", "Adet", "Alis Tarihi", "Durumu", "Aciklama"])
    } 

def verileri_kaydet(data_dict): 
    with pd.ExcelWriter(EXCEL_FILE) as writer: 
        for sheet_name, df in data_dict.items(): 
            df.to_excel(writer, sheet_name=sheet_name, index=False) 

def get_config(data, tip): 
    ayarlar = data.get("Ayarlar", pd.DataFrame())
    if ayarlar.empty or "Tip" not in ayarlar.columns:
        if tip == "YoneticiSifresi": return "sahin2026"
        if tip == "AktifDuyuru": return "Duyuru bulunmamaktadır."
        return 0.0
    res = ayarlar[ayarlar["Tip"].astype(str).str.contains(tip, case=False, na=False)]
    if res.empty:
        if tip == "YoneticiSifresi": return "sahin2026"
        if tip == "AktifDuyuru": return "Duyuru bulunmamaktadır."
        return 0.0
    val = res.iloc[0]["Tutar"]
    if tip in ["YoneticiSifresi", "AktifDuyuru"]:
        return str(val)
    return float(pd.to_numeric(val, errors="coerce") or 0)

def tr_temizle(metin): 
    return str(metin)

def sakin_bilgisi(data, no): 
    res = data["Sakinler"][data["Sakinler"]["No"] == no] 
    if not res.empty: 
        ad = res.iloc[0]["Ad"] 
        soyad = res.iloc[0]["Soyad"] 
        return f"{ad} {soyad}" 
    return "Sakin"

def daire_borc_hesapla(data, no): 
    da = get_config(data, "Daire") 
    ia = get_config(data, "İşyeri") 
    ay_sayisi = datetime.now().month 
    birim_aidat = ia if no in [7, 8] else da 
    toplam_borc = birim_aidat * ay_sayisi 
    
    odenen = pd.to_numeric(data["Gelirler"][data["Gelirler"]["No"] == no]["Miktar"], errors="coerce").sum() 
    anapara_bakiye = odenen - toplam_borc 
    
    faiz_orani = get_config(data, "GecikmeFaiziOrani") / 100.0 
    gecikme_tutari = 0
    if anapara_bakiye < 0:
        gecikme_tutari = abs(anapara_bakiye) * (faiz_orani)
        
    return anapara_bakiye, gecikme_tutari

def aidat_durum_tablosu_olustur(data):
    aidat_durumu = []
    daire_aidati = get_config(data, "Daire")
    isyeri_aidati = get_config(data, "İşyeri")
    ay_sayisi = datetime.now().month

    for no in range(1, 9):
        birim_aidati = isyeri_aidati if no in [7, 8] else daire_aidati
        toplam_borc = birim_aidati * ay_sayisi
        odenen = pd.to_numeric(data["Gelirler"][data["Gelirler"]["No"] == no]["Miktar"], errors="coerce").sum() 
        
        kalan_anapara = toplam_borc - odenen
        faiz_orani = get_config(data, "GecikmeFaiziOrani") / 100.0
        
        gecikme_faizi = 0.0
        if kalan_anapara > 0:
            gecikme_faizi = kalan_anapara * faiz_orani

        aidat_durumu.append({
            "Daire No": no,
            "Sakin": sakin_bilgisi(data, no),
            "Birim": "İşyeri" if no in [7, 8] else "Daire",
            "Ödenmesi Gereken": toplam_borc,
            "Ödenen": odenen,
            "Kalan Borç": max(kalan_anapara, 0),
            "Gecikme Faizi": gecikme_faizi,
            "Toplam Borç (Faizli)": max(kalan_anapara, 0) + gecikme_faizi
        })

    return pd.DataFrame(aidat_durumu)

def sakin_listesi_pdf_olustur(sakinler_df, baslik="KAT MALİKLERİ LİSTESİ"):
    pdf = pdf_hazirla()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    today = datetime.now().strftime("%d.%m.%Y")
    
    pdf.set_font(pdf.font_family, 'B', 16)
    pdf.cell(190, 10, tr_temizle(APARTMAN_ADI), ln=True, align='C')
    pdf.cell(190, 10, tr_temizle(baslik), ln=True, align='C')
    pdf.ln(5)
    
    pdf.set_font(pdf.font_family, 'B', 10)
    pdf.cell(0, 10, tr_temizle(f"Rapor Oluşturulma Tarihi: {today}"), ln=True, align='L')
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(5)
    
    pdf.set_font(pdf.font_family, 'B', 11)
    pdf.cell(20, 10, "No", 1, 0, 'C')
    pdf.cell(85, 10, "Ad", 1, 0, 'L')
    pdf.cell(85, 10, "Soyad", 1, 1, 'L')
    
    pdf.set_font(pdf.font_family, '', 10)
    for _, row in sakinler_df.iterrows():
        pdf.cell(20, 10, str(row.get("No", "")), 1, 0, 'C')
        pdf.cell(85, 10, tr_temizle(row.get("Ad", "")), 1, 0, 'L')
        pdf.cell(85, 10, tr_temizle(row.get("Soyad", "")), 1, 1, 'L')
    
    pdf.ln(10)
    pdf.set_font(pdf.font_family, 'I', 10)
    pdf.cell(0, 10, tr_temizle("-Yönetici / İmza-"), ln=True, align='C')
    pdf.cell(0, 10, tr_temizle("Şahin ERGEZER"), ln=True, align='C')
    return bytes(pdf.output())

def pdf_hazirla():
    pdf = FPDF()
    font_yolu = "arial.ttf"
    font_bold_yolu = "arialbd.ttf"
    font_italic_yolu = "ariali.ttf"
    if os.path.exists(font_yolu):
        try:
            pdf.add_font('ArialTR', '', font_yolu, uni=True)
            pdf.add_font('ArialTR', 'B', font_bold_yolu if os.path.exists(font_bold_yolu) else font_yolu, uni=True)
            pdf.add_font('ArialTR', 'I', font_italic_yolu if os.path.exists(font_italic_yolu) else font_yolu, uni=True)
            pdf.set_font('ArialTR', '', 12)
        except:
            pdf.set_font("Arial", size=12)
    else:
        pdf.set_font("Arial", size=12)
    return pdf

def tek_daire_pdf_olustur(data, no): 
    pdf = pdf_hazirla()
    pdf.set_auto_page_break(auto=True, margin=15) 
    today = datetime.now().strftime("%d.%m.%Y") 
    pdf.add_page() 
    tur = "DAİRE" if no <= 6 else "İŞYERİ" 
    sak = sakin_bilgisi(data, no) 
    pdf.set_font(pdf.font_family, 'B', 15) 
    pdf.cell(190, 10, tr_temizle(APARTMAN_ADI), ln=True, align='C') 
    pdf.set_font(pdf.font_family, '', 12) 
    pdf.cell(0, 10, tr_temizle(f"No : {no} ({sak}) {tur}"), ln=True, align='C') 
    pdf.line(10, pdf.get_y(), 200, pdf.get_y()) 
    pdf.ln(1) 
    daire_verisi = data["Gelirler"][data["Gelirler"]["No"] == no] 
    pdf.set_font(pdf.font_family, 'B', 10) 
    pdf.cell(0, 10, tr_temizle(f"  Rapor Oluşturulma Tarihi: {today}"), ln=True) 
    pdf.cell(45, 10, "Tarih", 1, 0, 'C') 
    pdf.cell(95, 10, "Açıklama", 1, 0, 'L') 
    pdf.cell(50, 10, "Miktar (TL)", 1, 1, 'R') 
    pdf.ln() 
    pdf.set_font(pdf.font_family, '', 10) 
    toplam = 0 
    for _, row in daire_verisi.iterrows(): 
        t = row['Tarih'].strftime('%d.%m.%Y') 
        m = float(row['Miktar']) 
        pdf.cell(45, 10, t, 1, 0, 'C') 
        pdf.cell(95, 10, "Aidat Tahsilatı", 1, 0, 'L') 
        pdf.cell(50, 10, f"{m:,.2f}", 1, 1, 'R') 
        toplam += m 
    pdf.ln(2) 
    pdf.set_font(pdf.font_family, 'B', 12) 
    pdf.cell(140, 10, "   TOPLAM TAHSİL EDİLEN:", 0) 
    pdf.cell(50, 10, f"{toplam:,.2f} TL", 0) 
    pdf.ln(10) 
    
    bakiye, faiz = daire_borc_hesapla(data, no)
    if bakiye < 0: 
        pdf.set_font(pdf.font_family, '', 10) 
        pdf.cell(140, 10, tr_temizle("   Ana Para Borç Miktarınız :"), 0) 
        pdf.cell(50, 10, f"{abs(bakiye):,.2f} TL", 0) 
        pdf.ln()
        if faiz > 0:
            pdf.cell(140, 10, tr_temizle("   Gecikme Faizi / Tazminatı :"), 0) 
            pdf.cell(50, 10, f"{faiz:,.2f} TL", 0) 
            pdf.ln()
            pdf.set_font(pdf.font_family, 'B', 10)
            pdf.cell(140, 10, tr_temizle("   TOPLAM BORÇ (Faizli):"), 0)
            pdf.cell(50, 10, f"{abs(bakiye)+faiz:,.2f} TL", 0)
            pdf.ln()
            
    pdf.ln(5) 
    pdf.set_font(pdf.font_family, 'I', 10) 
    pdf.cell(0, 10, tr_temizle("-Yönetici / İmza-"), ln=True, align='C') 
    pdf.cell(0, 10, tr_temizle("Şahin ERGEZER"), ln=True, align='C') 
    return bytes(pdf.output())

def genel_ozet_pdf_olustur(data): 
    pdf = pdf_hazirla()
    pdf.set_auto_page_break(auto=True, margin=15)
    today = datetime.now().strftime("%d.%m.%Y")
    
    pdf.add_page()
    pdf.set_font(pdf.font_family, 'B', 16)
    pdf.cell(190, 10, tr_temizle(APARTMAN_ADI), ln=True, align='C')
    pdf.cell(190, 10, tr_temizle("GENEL ÖZET RAPORU"), ln=True, align='C')
    pdf.ln(3)
    pdf.set_font(pdf.font_family, 'I', 10)
    pdf.cell(190, 8, tr_temizle("KASA DURUMU"), ln=True, align='C')
    pdf.ln(5)
    
    devir = get_config(data, "Devir")
    gelir = pd.to_numeric(data["Gelirler"]["Miktar"], errors='coerce').sum()
    gider = pd.to_numeric(data["Giderler"]["Miktar"], errors='coerce').sum()
    kasa = devir + gelir - gider
    
    pdf.set_font(pdf.font_family, 'B', 10)
    pdf.cell(120, 10, tr_temizle("Açıklama"), 1)
    pdf.cell(70, 10, tr_temizle("Tutar (TL)"), 1)
    pdf.ln()
    
    oz_veriler = [
        ("Geçen Yıldan Devir", devir),
        ("Toplam Gelir (+)", gelir),
        ("Toplam Gider (-)", gider),
        ("GÜNCEL KASA", kasa)
    ]
    pdf.set_font(pdf.font_family, '', 10)
    for aciklama, tutar in oz_veriler:
        pdf.cell(120, 10, tr_temizle(aciklama), 1)
        pdf.cell(70, 10, f"{tutar:,.2f}", 1)
        pdf.ln()

    pdf.ln(15)
    pdf.set_font(pdf.font_family, 'I', 10)
    pdf.cell(0, 10, tr_temizle("-Yönetici / İmza-"), ln=True, align='C')
    pdf.cell(0, 10, tr_temizle("Şahin ERGEZER"), ln=True, align='C')

    pdf.add_page()
    pdf.set_font(pdf.font_family, 'B', 14)
    pdf.cell(190, 10, tr_temizle(APARTMAN_ADI), ln=True, align='C')
    pdf.cell(190, 10, tr_temizle("BÖLÜMLERİN AİDAT VE FAİZ DURUMU"), ln=True, align='C')
    pdf.ln(5)
    
    pdf.set_font(pdf.font_family, 'B', 8)
    pdf.cell(12, 8, "No", 1, 0, 'C')
    pdf.cell(51, 8, "Sakin", 1, 0, 'L')
    pdf.cell(18, 8, "Birim", 1, 0, 'C')
    pdf.cell(27, 8, "Öd.Gereken", 1, 0, 'R')
    pdf.cell(27, 8, "Ödenen", 1, 0, 'R')
    pdf.cell(27, 8, "Kalan Borç", 1, 0, 'R')
    pdf.cell(28, 8, "Faizli Toplam", 1, 1, 'R')
    
    pdf.set_font(pdf.font_family, '', 8)
    aidat_df = aidat_durum_tablosu_olustur(data)
    for _, row in aidat_df.iterrows():
        pdf.cell(12, 8, str(int(row["Daire No"])), 1, 0, 'C')
        pdf.cell(51, 8, tr_temizle(row["Sakin"]), 1, 0, 'L')
        pdf.cell(18, 8, tr_temizle(row["Birim"]), 1, 0, 'C')
        pdf.cell(27, 8, f"{row['Ödenmesi Gereken']:,.2f}", 1, 0, 'R')
        pdf.cell(27, 8, f"{row['Ödenen']:,.2f}", 1, 0, 'R')
        pdf.cell(27, 8, f"{row['Kalan Borç']:,.2f}", 1, 0, 'R')
        pdf.cell(28, 8, f"{row['Toplam Borç (Faizli)']:,.2f}", 1, 1, 'R')

    pdf.ln(15)
    pdf.set_font(pdf.font_family, 'I', 10)
    pdf.cell(0, 10, tr_temizle("-Yönetici / İmza-"), ln=True, align='C')
    pdf.cell(0, 10, tr_temizle("Şahin ERGEZER"), ln=True, align='C')

    pdf.add_page()
    pdf.set_font(pdf.font_family, 'B', 14)
    pdf.cell(190, 10, tr_temizle(APARTMAN_ADI), ln=True, align='C')
    pdf.cell(190, 10, tr_temizle("GENEL GİDERLER LİSTESİ"), ln=True, align='C')
    pdf.ln(5)
    
    pdf.set_font(pdf.font_family, 'B', 9)
    pdf.cell(28, 8, tr_temizle("Tarih"), 1, 0, 'C')
    pdf.cell(42, 8, tr_temizle("Kategori"), 1, 0, 'L')
    pdf.cell(62, 8, tr_temizle("Açıklama"), 1, 0, 'L')
    pdf.cell(33, 8, tr_temizle("Miktar (TL)"), 1, 0, 'R')
    pdf.cell(25, 8, tr_temizle("Belge"), 1, 1, 'C')
    
    pdf.set_font(pdf.font_family, '', 9)
    toplam_gider = 0
    if not data["Giderler"].empty:
        for _, row in data["Giderler"].iterrows():
            t = row['Tarih'].strftime('%d.%m.%Y')
            k = tr_temizle(row['Kategori'])
            a = tr_temizle(row['Aciklama'])
            m = float(row['Miktar'])
            b = tr_temizle(row.get('Belge', 'YOK'))
            
            pdf.cell(28, 8, t, 1, 0, 'C')
            pdf.cell(42, 8, k, 1, 0, 'L')
            pdf.cell(62, 8, a, 1, 0, 'L')
            pdf.cell(33, 8, f"{m:,.2f}", 1, 0, 'R')
            pdf.cell(25, 8, b, 1, 1, 'C')
            toplam_gider += m
            
    pdf.ln(2)
    pdf.set_font(pdf.font_family, 'B', 10)
    pdf.cell(132, 8, tr_temizle("TOPLAM GİDER:"), 0, 0, 'R')
    pdf.cell(58, 8, f"{toplam_gider:,.2f} TL", 0, 1, 'R')

    pdf.ln(15)
    pdf.set_font(pdf.font_family, 'I', 10)
    pdf.cell(0, 10, tr_temizle("-Yönetici / İmza-"), ln=True, align='C')
    pdf.cell(0, 10, tr_temizle("Şahin ERGEZER"), ln=True, align='C')
    
    return bytes(pdf.output())

def dosya_yukle_belge(uploaded_file):
    klasor_olustur()
    dosya_yolu = os.path.join(BELGE_KLASORU, uploaded_file.name)
    with open(dosya_yolu, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return dosya_yolu

def klasordeki_dosyalari_listele():
    klasor_olustur()
    if not os.path.exists(BELGE_KLASORU):
        return []
    dosyalar = []
    for root, dirs, files in os.walk(BELGE_KLASORU):
        for file in files:
            dosya_yolu = os.path.join(root, file)
            dosyalar.append({
                "ad": file,
                "yol": dosya_yolu,
                "boyut": os.path.getsize(dosya_yolu),
                "tarih": datetime.fromtimestamp(os.path.getmtime(dosya_yolu)).strftime("%d.%m.%Y %H:%M")
            })
    return dosyalar

def dosya_sil(dosya_yolu):
    if os.path.exists(dosya_yolu):
        os.remove(dosya_yolu)
        return True
    return False

# --- UI BAŞLANGIÇ ---
st.set_page_config(page_title="SALİM APARTMANI", layout="wide") 
data = verileri_yukle() 

# --- GİRİŞ (LOGIN) EKRANI KONTROLÜ ---
if "giris_yapildi" not in st.session_state:
    st.session_state.giris_yapildi = False
    st.session_state.kullanici_tipi = None 
    st.session_state.daire_no = None

if not st.session_state.giris_yapildi:
    st.title(f"🏢 {APARTMAN_ADI}")
    st.subheader("🔑 Oturum Açma Ekranı")
    
    aktif_duyuru_ilk = get_config(data, "AktifDuyuru")
    if aktif_duyuru_ilk and str(aktif_duyuru_ilk).strip() != "":
        st.info(f"📢 **Yönetim Duyurusu:** {aktif_duyuru_ilk}")

    tab_yonetici, tab_sakin = st.tabs(["👨‍💼 Yönetici Girişi", "🚪 Daire Sakini Girişi"])
    
    with tab_yonetici:
        with st.form("yonetim_giris_form"):
            yonetici_sifre_input = st.text_input("Yönetici Şifresi", type="password")
            yonetici_submit = st.form_submit_button("Yönetici Olarak Giriş Yap")
            if yonetici_submit:
                gercek_sifre = get_config(data, "YoneticiSifresi")
                if yonetici_sifre_input == gercek_sifre:
                    st.session_state.giris_yapildi = True
                    st.session_state.kullanici_tipi = "yonetici"
                    st.success("Giriş başarılı! Yönlendiriliyorsunuz...")
                    st.rerun()
                else:
                    st.error("Hatalı yönetici şifresi!")

    with tab_sakin:
        with st.form("sakin_giris_form"):
            secilen_daire = st.selectbox("Daire / Bölüm Seçin", [1, 2, 3, 4, 5, 6, 7, 8])
            sakin_sifre_input = st.text_input("Daire Şifresi", type="password", help="Varsayılan şifre: 123 + Daire Numarası")
            sakin_submit = st.form_submit_button("Daire Sakini Olarak Giriş Yap")
            if sakin_submit:
                sakin_row = data["Sakinler"][data["Sakinler"]["No"] == secilen_daire]
                kayitli_sifre = f"123{secilen_daire}"
                if not sakin_row.empty and "Sifre" in sakin_row.columns:
                    kayitli_sifre = str(sakin_row.iloc[0]["Sifre"])
                
                if sakin_sifre_input == kayitli_sifre:
                    st.session_state.giris_yapildi = True
                    st.session_state.kullanici_tipi = "sakin"
                    st.session_state.daire_no = secilen_daire
                    st.success(f"Bölüm {secilen_daire} olarak giriş yapıldı!")
                    st.rerun()
                else:
                    st.error("Hatalı daire şifresi!")
    st.stop()

# ==========================================
# 1. EĞER GİRİŞ YAPAN DAİRE SAKİNİ İSE
# ==========================================
if st.session_state.kullanici_tipi == "sakin":
    no = st.session_state.daire_no
    sak = sakin_bilgisi(data, no)
    tur = "DAİRE" if no <= 6 else "İŞYERİ"
    
    st.sidebar.title(f"🚪 BÖLÜM {no} PANELİ")
    st.sidebar.write(f"**Sakini:** {sak} ({tur})")
    
    if st.sidebar.button("🔒 Oturumu Kapat", use_container_width=True):
        st.session_state.giris_yapildi = False
        st.session_state.kullanici_tipi = None
        st.session_state.daire_no = None
        st.rerun()

    st.title(f"🏠 {APARTMAN_ADI}")
    
    aktif_duyuru = get_config(data, "AktifDuyuru")
    if aktif_duyuru and str(aktif_duyuru).strip() != "":
        st.info(f"📢 **Yönetim Duyurusu:** {aktif_duyuru}")

    # --- EKLENEN GÜNCEL TARİH İBARESİ ---
    bugun_tarih = datetime.now().strftime("%d.%m.%Y")
    st.subheader(f"Bölüm {no} - {sak} - Borç ve Ödeme Durumu")
    st.caption(f"Rapor Tarihi: {bugun_tarih}")
    
    bakiye, faiz = daire_borc_hesapla(data, no)
    da = get_config(data, "Daire")
    ia = get_config(data, "İşyeri")
    ay_sayisi = datetime.now().month
    birim_aidat = ia if no in [7, 8] else da
    toplam_borc = birim_aidat * ay_sayisi
    odenen_toplam = pd.to_numeric(data["Gelirler"][data["Gelirler"]["No"] == no]["Miktar"], errors="coerce").sum()
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Ödenmesi Gereken", f"{toplam_borc:,.2f} TL")
    col2.metric("Yapılan Ödemeler", f"{odenen_toplam:,.2f} TL")
    
    kalan_anapara = toplam_borc - odenen_toplam
    toplam_faizli = max(kalan_anapara, 0) + (faiz if kalan_anapara > 0 else 0)
    col3.metric("Toplam Borç (Faiz Dahil)", f"{toplam_faizli:,.2f} TL", delta=f"-{toplam_faizli:,.2f} TL" if toplam_faizli > 0 else "0.00 TL", delta_color="inverse")

    st.markdown("---")
    st.subheader("📑 Yapılan Ödeme Geçmişiniz")
    daire_verisi = data["Gelirler"][data["Gelirler"]["No"] == no]
    if not daire_verisi.empty:
        display_df = daire_verisi.copy()
        display_df["Tarih"] = display_df["Tarih"].dt.strftime('%d.%m.%Y')
        st.dataframe(display_df, use_container_width=True, hide_index=True)
    else:
        st.info("Henüz sisteme işlenmiş ödemeniz bulunmamaktadır.")

    st.markdown("---")
    st.subheader("📥 Daire Gelir Raporu & Makbuz İndir")
    pdf_bytes = tek_daire_pdf_olustur(data, no)
    st.download_button(
        label="📄 Gelir Raporumu ve Borç Durumumu İndir (PDF)", 
        data=pdf_bytes, 
        file_name=f"Gelir_Raporu_Bolum_{no}_{sak.replace(' ', '_')}.pdf", 
        mime="application/pdf", 
        use_container_width=True
    )
    st.stop()


# ==========================================
# 2. EĞER GİRİŞ YAPAN YÖNETİCİ İSE
# ==========================================
st.sidebar.title("🏢 YÖNETİM PANELİ") 
st.sidebar.success("👨‍💼 Yönetici olarak giriş yapıldı.")

menu = st.sidebar.selectbox(
    "Menü",
    [
        "📊 Genel Özet",
        "🚨 Borç Takibi & Faiz",
        "📥 Aidat Tahsilatı",
        "💸 Gider Girişi",
        "🛋️ Demirbaş Envanteri",
        "📢 Duyuru Yönetimi",
        "📄 Raporlar",
        "📋 Kat Malikleri Listesi",
        "📑 Genel Kurul Belgeleri",
        "⚙️ Ayarlar"
    ]
)

if menu == "📊 Genel Özet": 
    st.title(APARTMAN_ADI) 
    devir = get_config(data, "Devir") 
    gelir = pd.to_numeric(data["Gelirler"]["Miktar"], errors='coerce').sum() 
    gider = pd.to_numeric(data["Giderler"]["Miktar"], errors='coerce').sum() 
    kasa = devir + gelir - gider 
    df_ozet = pd.DataFrame({ 
        "Açıklama": ["Geçen Yıldan Devir", "Toplam Gelir (+)", "Toplam Gider (-)", "GÜNCEL KASA"], 
        "Tutar (TL)": [f"{devir:,.2f}", f"{gelir:,.2f}", f"{gider:,.2f}", f"{kasa:,.2f}"] 
    }) 
    def kasa_renklendir(row): 
        if row["Açıklama"] == "GÜNCEL KASA": 
            return ["color: red; font-weight: bold"] * len(row) 
        return [""] * len(row) 
    st.dataframe(df_ozet.style.apply(kasa_renklendir, axis=1), use_container_width=True, hide_index=True) 
    
    st.markdown("---")
    st.subheader("📊 Bölümlerin Aidat ve Gecikme Faizi Durumu")
    aidat_df = aidat_durum_tablosu_olustur(data)
    st.dataframe(
        aidat_df.style.format({
            "Ödenmesi Gereken": "{:,.2f} TL",
            "Ödenen": "{:,.2f} TL",
            "Kalan Borç": "{:,.2f} TL",
            "Gecikme Faizi": "{:,.2f} TL",
            "Toplam Borç (Faizli)": "{:,.2f} TL"
        }),
        use_container_width=True,
        hide_index=True
    )
    
    if st.button("📄 Genel Özet Raporu (PDF)", use_container_width=True): 
        pdf_bytes = genel_ozet_pdf_olustur(data) 
        st.download_button( 
            "📥 Genel Özet Raporu İndir (PDF)", 
            data=pdf_bytes, 
            file_name="Genel_Ozet_Raporu.pdf", 
            mime="application/pdf", 
            use_container_width=True 
        ) 

    st.markdown("---")
    c1, c2 = st.columns([1, 1.5]) 
    with c1: 
        if not data["Giderler"].empty: 
            fig = px.pie(data["Giderler"], values='Miktar', names='Kategori', title="Harcama Dağılımı") 
            st.plotly_chart(fig, use_container_width=True) 
    with c2: 
        st.subheader("📝 Son İşlemler") 
        g_list = [] 
        if not data["Gelirler"].empty: 
            t_g = data["Gelirler"].copy(); t_g["Tür"] = "GELİR"; t_g["Detay"] = t_g["No"].apply(lambda x: f"Bölüm {x}") 
            g_list.append(t_g[["Tarih", "Tür", "Detay", "Miktar"]]) 
        if not data["Giderler"].empty: 
            t_gi = data["Giderler"].copy(); t_gi["Tür"] = "GİDER"; t_gi["Detay"] = t_gi["Kategori"]; t_gi["Miktar"] = t_gi["Miktar"] * -1 
            g_list.append(t_gi[["Tarih", "Tür", "Detay", "Miktar"]]) 
        if g_list: 
            son_on = pd.concat(g_list).sort_values(by="Tarih", ascending=False).head(20) 
            son_on["Tarih"] = son_on["Tarih"].dt.strftime('%d.%m.%Y') 
            st.dataframe(son_on, use_container_width=True, hide_index=True) 

elif menu == "🚨 Borç Takibi & Faiz": 
    st.title("🚨 Güncel Borç ve Gecikme Faizi Takibi") 
    st.info("Kat Mülkiyeti Kanunu Madde 20 gereği ödenmeyen aidatlara aylık gecikme faizi/tazminatı uygulanmaktadır.")
    
    faiz_orani = get_config(data, "GecikmeFaiziOrani")
    st.write(f"⚙️ **Uygulanan Aylık Gecikme Faizi Oranı:** %{faiz_orani}")
    
    aidat_tablosu = aidat_durum_tablosu_olustur(data)
    
    for _, row in aidat_tablosu.iterrows():
        no = int(row["Daire No"])
        sak = row["Sakin"]
        kalan = row["Kalan Borç"]
        faiz = row["Gecikme Faizi"]
        toplam_borcli = row["Toplam Borç (Faizli)"]
        
        c1, c2, c3 = st.columns([1, 2, 1]) 
        c1.write(f"**Bölüm {no} - {sak}**") 
        
        if kalan <= 0: 
            c2.success(f"✅ Borcu Bulunmamaktadır") 
            durum_metni = f"Bakiyeniz günceldir, borcunuz bulunmamaktadır." 
        else:
            c2.error(f"🚨 Anapara: {kalan:,.2f} TL | Faiz: {faiz:,.2f} TL | **Toplam: {toplam_borcli:,.2f} TL**") 
            durum_metni = f"Mevcut Borcunuz: {kalan:,.2f} TL, Gecikme Faizi ile Toplam: {toplam_borcli:,.2f} TL'dir." 
            
        msg = f"{APARTMAN_ADI}\\nBolum: {no} - {sak}\\nDurum: {durum_metni}\\nTarih: {datetime.now().strftime('%d.%m.%Y')}" 
        c3.markdown(f"[💬 Tekli WP Hatırlat](https://wa.me/?text={msg.replace(' ', '%20').replace('\\n', '%0A')})")
        
    st.markdown("---")
    st.subheader("📢 Toplu WhatsApp Bildirim Metni Üretici")
    toplu_mesaj = f"Sayın Kat Maliklerimiz,\n\n{APARTMAN_ADI} gereği güncel aidat ve gecikme faizi borçlarınızı kontrol etmenizi rica ederiz.\n\nİyi günler dileriz.\nŞahin ERGEZER - Yönetici"
    st.text_area("Toplu Mesaj Şablonu", value=toplu_mesaj, height=100)

elif menu == "📥 Aidat Tahsilatı": 
    st.title("📥 Tahsilat Girişi") 
    if "makbuz_hazir" not in st.session_state: 
        st.session_state.makbuz_hazir = False 
        st.session_state.son_makbuz = None 
        st.session_state.son_no = None 
        
    with st.form("t_makbuz_form"): 
        no = st.selectbox("Bölüm Seçin", [1,2,3,4,5,6,7,8]) 
        varsayilan_aidat = get_config(data, "İşyeri") if no in [7, 8] else get_config(data, "Daire")
        
        miktar = st.number_input("Tutar (TL)", min_value=0.0, value=float(varsayilan_aidat)) 
        tarih = st.date_input("İşlem Tarihi", value=datetime.now().date(), format="DD/MM/YYYY") 
        submit = st.form_submit_button("Kaydet ve Makbuz Hazırla") 
        
        if submit: 
            yeni = pd.DataFrame([[pd.to_datetime(tarih), no, miktar, "Aidat"]], columns=["Tarih", "No", "Miktar", "Tur"]) 
            data["Gelirler"] = pd.concat([data["Gelirler"], yeni], ignore_index=True) 
            verileri_kaydet(data) 
            
            pdf = pdf_hazirla()
            pdf.add_page() 
            pdf.set_font(pdf.font_family, 'B', 16) 
            pdf.cell(190, 10, txt=tr_temizle(APARTMAN_ADI), ln=True, align='C') 
            pdf.ln(10) 
            pdf.set_font(pdf.font_family, '', 12) 
            sak = sakin_bilgisi(data, no) 
            pdf.cell(190, 10, txt=tr_temizle(f"Sayın {sak} (Bölüm {no}),"), ln=True) 
            pdf.cell(190, 10, txt=tr_temizle(f"Tarih: {tarih.strftime('%d.%m.%Y')}"), ln=True) 
            pdf.cell(190, 10, txt=tr_temizle(f"Tahsil Edilen Tutar: {miktar:,.2f} TL"), ln=True) 
            
            bakiye, faiz = daire_borc_hesapla(data, no) 
            if bakiye >= 0: 
                pdf.cell(190, 10, txt=tr_temizle("Açıklama: Aidat Tahsilatı. Geçmiş borcunuz bulunmamaktadır."), ln=True) 
            else: 
                pdf.cell(190, 10, txt=tr_temizle(f"Kalan Anapara Borcunuz: {-bakiye:,.2f} TL, Faiz: {faiz:,.2f} TL"), ln=True) 
                
            pdf.ln(10) 
            pdf.set_font(pdf.font_family, 'I', 10)
            pdf.cell(190, 10, txt=tr_temizle("-Yönetici / İmza-"), ln=True, align='C') 
            pdf.cell(190, 10, txt=tr_temizle("Şahin ERGEZER"), ln=True, align='C') 
            
            st.session_state.son_makbuz = bytes(pdf.output()) 
            st.session_state.son_no = no 
            st.session_state.makbuz_hazir = True 
            st.success(f"Bölüm {no} için {miktar} TL başarıyla kaydedildi!") 
            
    if st.session_state.makbuz_hazir: 
        st.download_button( 
            label="📄 Makbuzu İndir (PDF)", 
            data=st.session_state.son_makbuz, 
            file_name=f"Makbuz_Bolum_{st.session_state.son_no}.pdf", 
            mime="application/pdf", 
            use_container_width=True 
        ) 
        if st.button("Yeni Kayıt Yap"): 
            st.session_state.makbuz_hazir = False 
            st.rerun() 

elif menu == "💸 Gider Girişi": 
    st.title("💸 Gider Girişi") 
    cats = ["Elektrik", "Su", "Asansör", "Temizlik", "Tamirat", "Bahçe", "Diğer"] 
    with st.form("gider_form"): 
        kat = st.selectbox("Kategori", cats) 
        miktar = st.number_input("Tutar (TL)", min_value=0.0) 
        aciklama = st.text_input("Açıklama") 
        belge = st.selectbox("Belge Durumu", ["Var", "YOK"], index=1)
        tarih = st.date_input("Tarih", value=datetime.now().date(), format="DD/MM/YYYY") 
        if st.form_submit_button("Gideri Kaydet"): 
            yeni = pd.DataFrame([[pd.to_datetime(tarih), kat, aciklama, miktar, belge]], columns=["Tarih", "Kategori", "Aciklama", "Miktar", "Belge"]) 
            data["Giderler"] = pd.concat([data["Giderler"], yeni], ignore_index=True) 
            verileri_kaydet(data) 
            st.success("Gider kaydedildi.") 
            st.rerun() 

elif menu == "🛋️ Demirbaş Envanteri":
    st.title("🛋️ Apartman Demirbaş Envanter Listesi")
    edited_demirbas = st.data_editor(
        data["Demirbaslar"], 
        num_rows="dynamic", 
        use_container_width=True,
        column_config={
            "Durumu": st.column_config.SelectboxColumn(
                "Durumu",
                options=["Sağlam", "Bakım Gerekiyor", "Hurda/Kullanılamaz"],
                required=True
            )
        }
    )
    if st.button("💾 Demirbaş Listesini Kaydet"):
        data["Demirbaslar"] = edited_demirbas
        verileri_kaydet(data)
        st.success("Demirbaş envanteri güncellendi!")
        st.rerun()

elif menu == "📢 Duyuru Yönetimi":
    st.title("📢 Apartman Duyuru Paneli")
    st.info("Buraya yazacağınız duyuru, daire sakinleri şifreleriyle sisteme giriş yaptıklarında ekranda anında görünecektir.")
    
    mevcut_duyuru = get_config(data, "AktifDuyuru")
    
    with st.form("duyuru_form"):
        yeni_duyuru = st.text_area("Daire Sakinlerine İletilecek Duyuru Metni", value=str(mevcut_duyuru), height=150)
        duyuru_submit = st.form_submit_button("📢 Duyuruyu Yayınla ve Kaydet")
        
        if duyuru_submit:
            ayarlar_df = data["Ayarlar"]
            if any(ayarlar_df["Tip"] == "AktifDuyuru"):
                ayarlar_df.loc[ayarlar_df["Tip"] == "AktifDuyuru", "Tutar"] = yeni_duyuru
            else:
                yeni_satir = pd.DataFrame([{"Tip": "AktifDuyuru", "Tutar": yeni_duyuru}])
                ayarlar_df = pd.concat([ayarlar_df, yeni_satir], ignore_index=True)
            
            data["Ayarlar"] = ayarlar_df
            verileri_kaydet(data)
            st.success("✅ Duyuru başarıyla güncellendi ve sakinler için yayına alındı!")
            st.rerun()

elif menu == "📄 Raporlar": 
    st.title("📄 Finansal Raporlar") 
    st.subheader("🏠 Daire Bazlı Gelir Raporları")
    for no in range(1, 9):
        sak = sakin_bilgisi(data, no)
        pdf_bytes = tek_daire_pdf_olustur(data, no)
        st.download_button(
            label=f"No:{no} ({sak}) - İndir Gelir Raporu (PDF)", 
            data=pdf_bytes, 
            file_name=f"Gelir_Raporu_No_{no}_{sak.replace(' ', '_')}.pdf", 
            mime="application/pdf", 
            use_container_width=True,
            key=f"btn_pdf_{no}"
        )
    st.markdown("---")
    st.subheader("📊 Genel Apartman Raporları")
    if st.button("Genel Özet, Faiz ve Gider Raporu (PDF)", use_container_width=True):
        pdf_bytes = genel_ozet_pdf_olustur(data)
        st.download_button("📥 İndir (Genel Özet Raporu)", data=pdf_bytes, file_name="Genel_Ozet_Faiz_Gider_Raporu.pdf", mime="application/pdf", use_container_width=True)

elif menu == "📋 Kat Malikleri Listesi":
    st.title("📋 Kat Malikleri ve Sakinler")
    sakinler = data["Sakinler"].copy()
    if sakinler.empty:
        st.warning("Henüz kayıtlı sakin bulunmamaktadır.")
    else:
        st.dataframe(sakinler, use_container_width=True, hide_index=True)
        pdf_bytes = sakin_listesi_pdf_olustur(sakinler)
        st.download_button(
            label="📄 Sakin Listesini PDF Olarak İndir",
            data=pdf_bytes,
            file_name="kat_malikleri_listesi.pdf",
            mime="application/pdf",
            use_container_width=True
        )

elif menu == "📑 Genel Kurul Belgeleri":
    st.title("📑 GENEL KURUL TOPLANTI BELGELERİ")
    uploaded_file = st.file_uploader("Dosya seçin (PDF, DOC, DOCX, XLS, XLSX)", type=["pdf", "doc", "docx", "xls", "xlsx"])
    if uploaded_file is not None:
        if st.button("📥 Dosyayı Yükle", use_container_width=True):
            dosya_yolu = dosya_yukle_belge(uploaded_file)
            st.success(f"✅ Dosya başarıyla yüklendi: {dosya_yolu}")
            st.rerun()
            
    st.subheader("📁 Yüklenen Dosyalar")
    dosyalar = klasordeki_dosyalari_listele()
    if dosyalar:
        for dosya in dosyalar:
            c1, c2, c3 = st.columns([3, 1, 1])
            with c1:
                st.write(f"**📄 {dosya['ad']}**")
                st.caption(f"Boyut: {dosya['boyut'] / 1024:.1f} KB | Tarih: {dosya['tarih']}")
            with c2:
                with open(dosya['yol'], "rb") as f:
                    st.download_button("📥 İndir", data=f.read(), file_name=dosya['ad'], key=f"d_{dosya['ad']}", use_container_width=True)
            with c3:
                if st.button("🗑️ Sil", key=f"del_{dosya['ad']}", use_container_width=True):
                    dosya_sil(dosya['yol'])
                    st.rerun()
            st.markdown("---")

elif menu == "⚙️ Ayarlar": 
    st.title("⚙️ Yönetim Ayarları") 
    t1, t2, t3, t4, t5 = st.tabs(["💰 Aidat/Devir/Faiz", "📥 Tahsilat", "💸 Gider", "🛋️ Demirbaş", "👥 Sakin"]) 
    with t1: 
        with st.form("ayarlar_form"): 
            d_a = st.number_input("Daire Aidatı (TL)", value=float(get_config(data, "Daire"))) 
            i_a = st.number_input("İşyeri Aidatı (TL)", value=float(get_config(data, "İşyeri"))) 
            dev = st.number_input("Geçen Yıldan Devir (Kasa)", value=float(get_config(data, "Devir"))) 
            faiz_or = st.number_input("Aylık Gecikme Faizi Oranı (%)", value=float(get_config(data, "GecikmeFaiziOrani")))
            yon_sifre = st.text_input("Yönetici Giriş Şifresi", value=str(get_config(data, "YoneticiSifresi")))
            
            if st.form_submit_button("Ayarları Güncelle"): 
                ayarlar_df = data["Ayarlar"]
                yeni_degerler = [
                    {"Tip": "Daire", "Tutar": d_a}, 
                    {"Tip": "İşyeri", "Tutar": i_a}, 
                    {"Tip": "Devir", "Tutar": dev},
                    {"Tip": "GecikmeFaiziOrani", "Tutar": faiz_or},
                    {"Tip": "YoneticiSifresi", "Tutar": yon_sifre}
                ]
                for item in yeni_degerler:
                    tip = item["Tip"]
                    tutar = item["Tutar"]
                    if any(ayarlar_df["Tip"] == tip):
                        ayarlar_df.loc[ayarlar_df["Tip"] == tip, "Tutar"] = tutar
                    else:
                        ayarlar_df = pd.concat([ayarlar_df, pd.DataFrame([item])], ignore_index=True)
                
                data["Ayarlar"] = ayarlar_df
                verileri_kaydet(data) 
                st.success("Ayarlar kaydedildi!") 
                st.rerun() 
    with t2: 
        st.subheader("📥 Tüm Tahsilatlar") 
        if not data["Gelirler"].empty: 
            gelir_display = data["Gelirler"].copy() 
            gelir_display["Sıra No"] = range(1, len(gelir_display) + 1) 
            gelir_display["Tarih"] = pd.to_datetime(gelir_display["Tarih"], errors='coerce').dt.strftime('%d.%m.%Y') 
            gelir_display = gelir_display[["Sıra No", "Tarih", "No", "Miktar", "Tur"]] 
        else: 
            gelir_display = pd.DataFrame(columns=["Sıra No", "Tarih", "No", "Miktar", "Tur"]) 
        edited_gelir = st.data_editor(gelir_display.reset_index(drop=True), num_rows="dynamic", use_container_width=True, hide_index=True) 
        if st.button("Tahsilat Değişikliklerini Kaydet"): 
            edited_gelir_copy = edited_gelir.copy() 
            edited_gelir_copy["Tarih"] = pd.to_datetime(edited_gelir_copy["Tarih"], format='%d.%m.%Y', errors='coerce') 
            data["Gelirler"] = edited_gelir_copy.drop(columns=["Sıra No"]).reset_index(drop=True) 
            verileri_kaydet(data) 
            st.success("✅ Tahsilatlar kaydedildi!") 
            st.rerun() 
    with t3: 
        st.subheader("💸 Tüm Giderler") 
        if not data["Giderler"].empty: 
            df_gider = data["Giderler"].copy() 
            df_gider["Sıra No"] = range(1, len(df_gider) + 1) 
            df_gider["Tarih"] = pd.to_datetime(df_gider["Tarih"]).dt.strftime('%d.%m.%Y') 
            gider_display = df_gider[["Sıra No", "Tarih", "Kategori", "Aciklama", "Miktar", "Belge"]] 
        else: 
            gider_display = pd.DataFrame(columns=["Sıra No", "Tarih", "Kategori", "Aciklama", "Miktar", "Belge"]) 
            
        edited_gider = st.data_editor(
            gider_display.reset_index(drop=True), 
            num_rows="dynamic", 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Belge": st.column_config.SelectboxColumn("Belge", options=["Var", "YOK"], required=True)
            }
        ) 
        if st.button("💾 Giderleri Kaydet"): 
            data["Giderler"] = edited_gider.drop(columns=["Sıra No"]).reset_index(drop=True) 
            if not data["Giderler"].empty: 
                data["Giderler"]["Tarih"] = pd.to_datetime(data["Giderler"]["Tarih"], format='%d.%m.%Y', errors='coerce') 
            verileri_kaydet(data) 
            st.success("Gider tablosu güncellendi!") 
            st.rerun() 
    with t4:
        st.subheader("🛋️ Demirbaş Yönetimi")
        edited_deb = st.data_editor(data["Demirbaslar"], num_rows="dynamic", use_container_width=True)
        if st.button("Demirbaş Tablosunu Kaydet"):
            data["Demirbaslar"] = edited_deb
            verileri_kaydet(data)
            st.success("Demirbaş listesi kaydedildi.")
            st.rerun()
    with t5: 
        st.subheader("👥 Sakin Listesi ve Daire Şifreleri") 
        if "Sifre" not in data["Sakinler"].columns:
            data["Sakinler"]["Sifre"] = [f"123{i}" for i in data["Sakinler"]["No"]]
        edited_sakinler = st.data_editor(data["Sakinler"], num_rows="dynamic", use_container_width=True) 
        if st.button("Sakin Listesini Kaydet"): 
            data["Sakinler"] = edited_sakinler 
            verileri_kaydet(data) 
            st.success("Sakin listesi güncellendi!") 
            st.rerun() 

st.sidebar.markdown("---")
if st.sidebar.button("🔒 Güvenli Çıkış / Oturumu Kapat", use_container_width=True):
    verileri_kaydet(data)
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.caption("👨‍💼 SALİM APARTMANI YÖNETİMİ")
st.sidebar.caption("📅 " + datetime.now().strftime("%d.%m.%Y %H:%M"))
st.sidebar.markdown("---")
st.sidebar.markdown("<p style='color: red; text-align: center; font-weight: bold;'>Bu program Şahin ERGEZER tarafından hazırlanmıştır (2026)</p>", unsafe_allow_html=True)