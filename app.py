import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
import requests
import xml.etree.ElementTree as ET
import streamlit as st

def check_password():
    correct = st.secrets.get("APP_PASSWORD", "")

    if not correct:
        st.warning("Şifre ayarlanmamış. Streamlit Secrets içine APP_PASSWORD ekle.")
        st.stop()

    def password_entered():
        if st.session_state.get("password", "") == correct:
            st.session_state["password_ok"] = True
            st.session_state["password"] = ""
        else:
            st.session_state["password_ok"] = False

    if st.session_state.get("password_ok") is True:
        return True

    st.text_input("Şifre", type="password", key="password", on_change=password_entered)
    if st.session_state.get("password_ok") is False:
        st.error("Yanlış şifre")
    return False

if not check_password():
    st.stop()

st.set_page_config(page_title="Stok & Fiyat", layout="wide")
st.title("Stok & Fiyatlandırma (BLT / Havuztek)")

# -------------------------
# 1) Excel okuma ve temizleme
# -------------------------
@st.cache_data
def load_blt_excel(uploaded_file) -> pd.DataFrame:
    # Dosyada başlık/boş satırlar var: header=None okuyup “SIRA” sayısal olan satırları alıyoruz
    raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    # Excel bazen az sütun okur; biz 10 sütunu garanti edelim (0..9)
    raw = raw.reindex(columns=range(10))

    # İlk kolon (SIRA) tam sayı olan satırlar = ürün satırı
    mask = raw[0].apply(lambda x: isinstance(x, (int, np.integer, float)) and pd.notna(x) and float(x).is_integer())
    data = raw.loc[mask].copy()

    # 10 kolon adını güvenle ver
    data.columns = [
        "SIRA","KOD","EBAT","URUN","MALIYET","GENEL_GIDER",
        "HAVUZTEK_SATIS_USD","BLT_SATIS_USD","KUR","BLT_SATIS_TL"
    ]

    # Sayısal temizlik
    data["BLT_SATIS_USD"] = pd.to_numeric(data["BLT_SATIS_USD"], errors="coerce")
    data["KUR"] = pd.to_numeric(data["KUR"], errors="coerce").ffill()

    # Ürün adı: URUN + EBAT
    data["UrunAdi"] = (data["URUN"].astype(str).str.strip() + " " + data["EBAT"].astype(str).str.strip()).str.strip()

    return data[["UrunAdi","KOD","EBAT","BLT_SATIS_USD","KUR"]].copy()
def tcmb_usd_kuru():
    url = "https://www.tcmb.gov.tr/kurlar/today.xml"
    r = requests.get(url, timeout=10)
    r.raise_for_status()

    root = ET.fromstring(r.content)

    # Tarih bilgisini de alalım
    tarih = root.attrib.get("Tarih")

    # USD satıs kuru: BanknoteSelling (yoksa ForexSelling)
    usd = None
    for cur in root.findall("Currency"):
        if cur.get("Kod") == "USD":
            val = cur.findtext("BanknoteSelling") or cur.findtext("ForexSelling")
            if val:
                usd = float(val.replace(",", "."))
            break

    if usd is None:
        raise ValueError("TCMB XML içinde USD kuru bulunamadı.")

    return usd, tarih


# -------------------------
# 2) Sol panel: firma ve kur
# -------------------------
st.sidebar.header("Ayarlar")

firma = st.sidebar.selectbox("Firma seç", ["BLT", "HAVUZTEK"])

kur_kaynagi = st.sidebar.radio("Kur", ["TCMB (otomatik)", "Elle"], index=0)

if kur_kaynagi == "TCMB (otomatik)":
    try:
        otomatik_kur, tarih = tcmb_usd_kuru()
        st.sidebar.success(f"TCMB USD: {otomatik_kur:.4f} (Tarih: {tarih})")
        kur = st.sidebar.number_input("USD/TRY (istersen değiştir)", min_value=0.0, value=float(otomatik_kur), step=0.01)
    except Exception as e:
        st.sidebar.error(f"Kur çekilemedi: {e}")
        kur = st.sidebar.number_input("USD/TRY (elle gir)", min_value=0.0, value=0.0, step=0.01)
else:
    kur = st.sidebar.number_input("USD/TRY (elle gir)", min_value=0.0, value=34.0, step=0.01)

havuztek_indirim = st.sidebar.number_input(
    "Havuztek indirim (%)",
    min_value=0.0,
    max_value=100.0,
    value=20.0,
    step=1.0
)

st.sidebar.divider()
st.sidebar.caption("Havuztek fiyatı = BLT fiyatı × (1 - indirim)")

# -------------------------
# 3) Dosyayı yüklet (kolay kullanım)
# -------------------------
uploaded = st.sidebar.file_uploader("BLT Excel dosyanı yükle (.xlsx)", type=["xlsx"])

import os

DEFAULT_EXCEL = "BLT GÜNCEL FİYATLLAR.xlsx"

st.sidebar.header("Dosya")
dosya_modu = st.sidebar.radio("Excel Kaynağı", ["Klasördeki dosya (otomatik)", "Elle seç (opsiyonel)"], index=0)

excel_kaynagi = None

if dosya_modu == "Klasördeki dosya (otomatik)":
    if not os.path.exists(DEFAULT_EXCEL):
        st.error(f"Klasörde '{DEFAULT_EXCEL}' bulunamadı. Dosyayı stok_app içine koy.")
        st.stop()
    excel_kaynagi = DEFAULT_EXCEL
    st.sidebar.success(f"Otomatik dosya: {DEFAULT_EXCEL}")
else:
    uploaded = st.sidebar.file_uploader("Excel seç (.xlsx)", type=["xlsx"])
    if not uploaded:
        st.info("Soldan Excel dosyasını seç.")
        st.stop()
    excel_kaynagi = uploaded

# Excel oku
try:
    fiyatlar = load_blt_excel(excel_kaynagi)
except Exception as e:
    st.error(f"Excel okunamadı: {e}")
    st.stop()

# TL fiyatı üret
fiyatlar["BLT_Fiyat_TL"] = (fiyatlar["BLT_SATIS_USD"] * kur).round(2)

if firma == "BLT":
    fiyatlar["Fiyat_TL"] = fiyatlar["BLT_Fiyat_TL"]
else:
    fiyatlar["Fiyat_TL"] = (fiyatlar["BLT_Fiyat_TL"] * (1 - havuztek_indirim/100)).round(2)

# -------------------------
# 4) Ürün arama ve sepet
# -------------------------
st.header("Ürünler")

q = st.text_input("Ürün ara (isimde geçen)")
view = fiyatlar.copy()
if q.strip():
    view = view[view["UrunAdi"].str.lower().str.contains(q.strip().lower(), na=False)]

st.dataframe(
    view[["UrunAdi","KOD","EBAT","BLT_SATIS_USD","Fiyat_TL"]].head(300),
    use_container_width=True
)

# Sepet state
if "sepet" not in st.session_state:
    st.session_state.sepet = []

st.header("Sepet")

urun = st.selectbox("Ürün seç", options=view["UrunAdi"].dropna().unique().tolist())
miktar = st.number_input("Miktar", min_value=1.0, step=1.0, value=1.0)

c1, c2 = st.columns(2)
with c1:
    if st.button("Sepete ekle"):
        st.session_state.sepet.append({"UrunAdi": urun, "Miktar": float(miktar)})
with c2:
    if st.button("Sepeti temizle"):
        st.session_state.sepet = []
        st.rerun()

if not st.session_state.sepet:
    st.info("Sepete ürün ekleyince hesap burada görünecek.")
    st.stop()

sepet_df = pd.DataFrame(st.session_state.sepet)

detay = sepet_df.merge(
    fiyatlar[["UrunAdi", "Fiyat_TL"]],
    on="UrunAdi",
    how="left"
)

detay["AraTutar"] = detay["Miktar"] * detay["Fiyat_TL"]
toplam = float(detay["AraTutar"].sum())

st.subheader("Hesap Detay")
st.dataframe(detay, use_container_width=True)

st.metric("Genel Toplam (TL)", f"{toplam:,.2f}")
# -------------------------
# Excel çıktısı
# -------------------------
output = BytesIO()

# Detay sayfasını daha okunur yapalım
detay_excel = detay.copy()
detay_excel["Firma"] = firma
detay_excel["Kur"] = kur
detay_excel["Havuztek_Indirim_%"] = havuztek_indirim if firma == "HAVUZTEK" else 0

ozet = pd.DataFrame([{
    "Firma": firma,
    "Kur": kur,
    "Havuztek_Indirim_%": havuztek_indirim if firma == "HAVUZTEK" else 0,
    "GenelToplam_TL": toplam
}])

with pd.ExcelWriter(output, engine="openpyxl") as writer:
    detay_excel.to_excel(writer, sheet_name="Detay", index=False)
    ozet.to_excel(writer, sheet_name="Ozet", index=False)

st.download_button(
    label="Teklifi Excel indir",
    data=output.getvalue(),
    file_name=f"teklif_{firma}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


