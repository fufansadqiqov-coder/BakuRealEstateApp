# =========================================================
# STREAMLIT + OSS + SCRAPING + ANALIZ + EMAIL + UI
# =========================================================

import streamlit as st
import pandas as pd
import re
import time
import random
from io import BytesIO
from typing import Dict, Any, Optional
import requests
from bs4 import BeautifulSoup
from geopy.distance import geodesic
from email.message import EmailMessage
import smtplib
import oss2
import plotly.express as px
from concurrent.futures import ThreadPoolExecutor

# =========================================================
# KONFİQURASİYA
try:
    ACCESS_KEY_ID = st.secrets["ACCESS_KEY_ID"]
    ACCESS_KEY_SECRET = st.secrets["ACCESS_KEY_SECRET"]
    APP_PASSWORD = st.secrets["APP_PASSWORD"]
except KeyError:
    st.error("Secrets konfiqurasiyası tapılmadı. `.streamlit/secrets.toml` faylını yaradın.")
    st.stop()

ENDPOINT = "oss-ap-southeast-1.aliyuncs.com"
BUCKET_NAME = "emlak-bot-demo"
HOME_SALES_KEY = "Home Sales Statistika and Location.xlsx"
OUTPUT_KEY = "Depo.xlsx"
REGISTERED_USERS_EXCEL_KEY = "Gmail.xlsx"
SENDER_EMAIL = "fufansadqiqov@gmail.com"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}
BASE_URL = "https://bina.az"

# =========================================================
# YARDIMÇI FUNKSİYALAR
@st.cache_resource
def get_oss_bucket() -> oss2.Bucket:
    auth = oss2.Auth(ACCESS_KEY_ID, ACCESS_KEY_SECRET)
    bucket = oss2.Bucket(auth, ENDPOINT, BUCKET_NAME)
    return bucket

@st.cache_data
def load_initial_data(bucket: oss2.Bucket):
    obj = bucket.get_object(HOME_SALES_KEY)
    xls = pd.ExcelFile(BytesIO(obj.read()))
    statistik = pd.read_excel(xls, sheet_name='Statistic')
    metro_data = pd.read_excel(xls, sheet_name='Location')
    statistik["Qiymet"] = pd.to_numeric(statistik["Qiymet"].astype(str).str.replace(r"[^\d]", "", regex=True), errors='coerce')
    statistik["Sahə"] = pd.to_numeric(statistik["Sahə"].astype(str).str.extract(r"(\d+\.?\d*)")[0], errors='coerce')
    statistik.dropna(subset=["Qiymet", "Sahə"], inplace=True)
    statistik = statistik[statistik["Sahə"] != 0].copy()
    statistik["Qiymet_m2"] = statistik["Qiymet"] / statistik["Sahə"]
    ortalama_m2 = statistik.groupby("Erazi")["Qiymet_m2"].median().reset_index()
    ortalama_m2.rename(columns={"Qiymet_m2": "Ortalama_Qiymet_m2"}, inplace=True)
    return statistik, metro_data, ortalama_m2

def load_registered_users(bucket: oss2.Bucket):
    if bucket.object_exists(REGISTERED_USERS_EXCEL_KEY):
        obj = bucket.get_object(REGISTERED_USERS_EXCEL_KEY)
        df = pd.read_excel(BytesIO(obj.read()))
        return set(df["Gmail"].dropna().tolist()) if "Gmail" in df.columns else set()
    return set()

def save_new_user(bucket: oss2.Bucket, email: str):
    existing = load_registered_users(bucket)
    if email not in existing:
        existing.add(email)
        df = pd.DataFrame(list(existing), columns=["Gmail"])
        buf = BytesIO()
        df.to_excel(buf, index=False)
        buf.seek(0)
        bucket.put_object(REGISTERED_USERS_EXCEL_KEY, buf)

def send_email(subject: str, body: str, receivers: list):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = SENDER_EMAIL
    msg['To'] = ", ".join(receivers)
    msg.set_content(body)
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(SENDER_EMAIL, APP_PASSWORD)
            server.send_message(msg)
        return True
    except Exception as e:
        st.error(f"Email göndərmə xətası: {e}")
        return False

# =========================================================
# SCRAPING & ANALİZ
def scrape_item_details(link: str) -> Optional[Dict[str, Any]]:
    try:
        r = requests.get(link, headers=HEADERS, timeout=15)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        map_div = soup.find("div", {"id": "item_map"})
        details = {
            "Qiymet": soup.find("div", class_="product-price__i").get_text(strip=True) if soup.find("div", class_="product-price__i") else None,
            "Erazi": soup.find("h1", class_="product-title").get_text(strip=True).split(",")[-1].strip() if soup.find("h1", class_="product-title") else None,
            "Link": link,
            "Lat": float(map_div.get("data-lat")) if map_div and map_div.get("data-lat") else None,
            "Lng": float(map_div.get("data-lng")) if map_div and map_div.get("data-lng") else None
        }
        return details
    except:
        return None

def process_item(details: Dict[str, Any], ortalama_m2: pd.DataFrame, metro_data: pd.DataFrame):
    if not details: return None
    if not details.get("Qiymet") or not details.get("Erazi"): return None
    try:
        details['Qiymet'] = int(re.sub(r'[^\d]', '', str(details['Qiymet'])))
    except: return None
    details['qiymet_m2'] = details['Qiymet']/random.randint(40,100)  # dummy Sahə
    row = ortalama_m2[ortalama_m2["Erazi"].str.lower()==details["Erazi"].lower()]
    avg_price = row["Ortalama_Qiymet_m2"].values[0] if not row.empty else None
    details['ortalama_qiymet_m2'] = avg_price
    # Metro məsafəsi
    lat,lng = details.get("Lat"),details.get("Lng")
    if lat and lng:
        try:
            distances = [geodesic((lat,lng),(row["Lat"],row["Long"])).km for _,row in metro_data.iterrows()]
            details['Metroya məsafə (km)'] = round(min(distances),2) if distances else None
        except: details['Metroya məsafə (km)'] = None
    return details

def run_process(bucket, metro_data, ortalama_m2, max_pages, selected_areas):
    st.info("Prosess başlayır...")
    depo_df = pd.DataFrame(columns=["Link","Qiymet","Erazi","qiymet_m2","ortalama_qiymet_m2","Metroya məsafə (km)"])
    # Dummy scraping links
    links = [f"{BASE_URL}/items/{i}" for i in range(max_pages*5)]
    results=[]
    with ThreadPoolExecutor(max_workers=5) as executor:
        for det in executor.map(scrape_item_details, links):
            processed = process_item(det, ortalama_m2, metro_data)
            if processed:
                results.append(processed)
    if results:
        new_df = pd.DataFrame(results)
        save_df_to_oss(bucket, new_df)
        st.success(f"{len(new_df)} elan işlənib və OSS-ə yazıldı")
        st.dataframe(new_df)

def save_df_to_oss(bucket, df):
    buf = BytesIO()
    df.to_excel(buf,index=False)
    buf.seek(0)
    bucket.put_object(OUTPUT_KEY, buf)

# =========================================================
# STREAMLIT UI
st.set_page_config(page_title="Əmlak Analizatoru", layout="wide", page_icon="🏙️")
bucket = get_oss_bucket()
statistik, metro_data, ortalama_m2 = load_initial_data(bucket)

# Session State
if 'verification_step' not in st.session_state: st.session_state.verification_step='enter_email'
if 'user_email' not in st.session_state: st.session_state.user_email=''
if 'verification_code' not in st.session_state: st.session_state.verification_code=''

# Login / Qeydiyyat
if st.session_state.verification_step != 'verified':
    st.markdown('<div style="padding:3rem;max-width:500px;margin:auto;background:#fff;border-radius:20px;">', unsafe_allow_html=True)
    registered_emails = load_registered_users(bucket)
    if st.session_state.verification_step == 'enter_email':
        email = st.text_input("Email ünvanınız:")
        if st.button("Davam et"):
            if re.match(r"[^@]+@[^@]+\.[^@]+", email):
                if email in registered_emails:
                    st.session_state.user_email, st.session_state.verification_step = email,'verified'
                    st.rerun()
                else:
                    code=f"{random.randint(100000,999999)}"
                    if send_email("Qeydiyyat kodu", f"Kodunuz: {code}", [email]):
                        st.session_state.user_email, st.session_state.verification_code, st.session_state.verification_step = email, code, 'enter_code'
                        st.success(f"Kod göndərildi: {code} (test üçün ekranda)")
            else: st.error("Düzgün email daxil edin")
    elif st.session_state.verification_step=='enter_code':
        code_input = st.text_input("Təsdiq kodu:")
        if st.button("Qeydiyyatı tamla"):
            if code_input==st.session_state.verification_code:
                save_new_user(bucket, st.session_state.user_email)
                st.session_state.verification_step='verified'
                st.balloons()
                st.rerun()
            else: st.error("Kod yanlışdır")
    st.markdown('</div>', unsafe_allow_html=True)

else:  # Main App
    st.sidebar.info(f"Giriş edildi: {st.session_state.user_email}")
    max_pages = st.sidebar.slider("Yoxlanılacaq səhifə sayı",1,20,3)
    all_areas = sorted(ortalama_m2["Erazi"].unique().tolist())
    selected_areas = st.sidebar.multiselect("Ərazi filtri", options=all_areas)
    if st.sidebar.button("🚀 Axtarışı Başlat"):
        run_process(bucket, metro_data, ortalama_m2, max_pages, selected_areas)

    st.title("İdarə Paneli")
    st.markdown("### Ümumi Baxış")
    cols = st.columns(4)
    card_data=[
        ("Ərazi Sayı", len(ortalama_m2),"Analiz edilən"),
        ("Metro Stansiyaları", len(metro_data),"Məlumat bazasında"),
        ("Elan Sayı", len(statistik),"Data bazasında"),
        ("Ortalama Qiymət", f"{ortalama_m2['Ortalama_Qiymet_m2'].median():.0f}","AZN/m² (Median)")
    ]
    card_classes=["#667eea","#f093fb","#4facfe","#43e97b"]
    for i,col in enumerate(cols):
        with col:
            st.markdown(f"<div style='background:{card_classes[i]};padding:2rem;border-radius:15px;color:#fff;text-align:center;'><h4>{card_data[i][0]}</h4><h2>{card_data[i][1]}</h2><p>{card_data[i][2]}</p></div>",unsafe_allow_html=True)

    st.markdown("### Ən Bahalı Ərazilər")
    top_areas = ortalama_m2.nlargest(10,"Ortalama_Qiymet_m2")
    fig_bar = px.bar(top_areas.sort_values("Ortalama_Qiymet_m2"), y="Erazi", x="Ortalama_Qiymet_m2", orientation='h', text="Ortalama_Qiymet_m2", template="plotly_white")
    st.plotly_chart(fig_bar,use_container_width=True)
