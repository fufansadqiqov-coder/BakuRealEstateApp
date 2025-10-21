# =========================================================
# BÜTÜN KOD (Streamlit + OSS + Scraping + Analiz + Email + UI)
# =========================================================
import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
from rapidfuzz import process, fuzz
from geopy.distance import geodesic
from email.message import EmailMessage
import smtplib
from io import BytesIO
import oss2
import random
from typing import Dict, Any, Optional
import plotly.express as px

# =========================================================
# KONFİQURASİYA
try:
    ACCESS_KEY_ID = st.secrets["ACCESS_KEY_ID"]
    ACCESS_KEY_SECRET = st.secrets["ACCESS_KEY_SECRET"]
    APP_PASSWORD = st.secrets["APP_PASSWORD"]
except (KeyError, FileNotFoundError):
    st.error("Secrets konfiqurasiyası tapılmadı. `.streamlit/secrets.toml` faylını yaradın.")
    st.stop()

ENDPOINT = "oss-ap-southeast-1.aliyuncs.com"
BUCKET_NAME = "emlak-bot-demo"
HOME_SALES_KEY = "Home Sales Statistika and Location.xlsx"
OUTPUT_KEY = "Depo.xlsx"
REGISTERED_USERS_EXCEL_KEY = "Gmail.xlsx"
SENDER_EMAIL = "fufansadqiqov@gmail.com"
HEADERS = { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36" }
BASE_URL = "https://bina.az"

# =========================================================
# YARDIMÇI FUNKSİYALAR
def get_oss_bucket() -> oss2.Bucket:
    try:
        auth = oss2.Auth(ACCESS_KEY_ID, ACCESS_KEY_SECRET)
        bucket = oss2.Bucket(auth, ENDPOINT, BUCKET_NAME)
        bucket.get_bucket_info()
        return bucket
    except Exception as e:
        st.error(f"Aliyun OSS bağlantı xətası: {e}")
        st.stop()

def load_initial_data(_bucket: oss2.Bucket):
    try:
        obj = _bucket.get_object(HOME_SALES_KEY)
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
    except Exception as e:
        st.error(f"İlkin data yüklənməsi xətası ('{HOME_SALES_KEY}'): {e}")
        st.stop()

def send_email(subject: str, body: str, receivers: list):
    if not all([SENDER_EMAIL, APP_PASSWORD, receivers]):
        st.warning("Email konfiqurasiyası tam deyil, bildiriş göndərilmədi.")
        return False
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
        st.error(f"Email göndərilərkən xəta baş verdi: {e}")
        return False

def load_registered_users_from_excel(bucket: oss2.Bucket) -> set:
    try:
        if bucket.object_exists(REGISTERED_USERS_EXCEL_KEY):
            obj = bucket.get_object(REGISTERED_USERS_EXCEL_KEY)
            df = pd.read_excel(BytesIO(obj.read()))
            return set(df["Gmail"].dropna().tolist()) if "Gmail" in df.columns else set()
        return set()
    except Exception as e:
        st.error(f"Qeydiyyatlı istifadəçilər yüklənərkən xəta: {e}")
        return set()

def save_new_user_to_excel(bucket: oss2.Bucket, new_email: str):
    try:
        existing_emails = load_registered_users_from_excel(bucket)
        if new_email not in existing_emails:
            existing_emails.add(new_email)
            df = pd.DataFrame(list(existing_emails), columns=["Gmail"])
            output_stream = BytesIO()
            df.to_excel(output_stream, index=False)
            output_stream.seek(0)
            bucket.put_object(REGISTERED_USERS_EXCEL_KEY, output_stream)
    except Exception as e:
        st.error(f"Yeni istifadəçi yadda saxlanılarkən xəta: {e}")

# =========================================================
# STREAMLIT UI
st.set_page_config(page_title="Əmlak Analizatoru", layout="wide", page_icon="🏙️")

# Session state
if 'verification_step' not in st.session_state: st.session_state.verification_step = 'enter_email'
if 'user_email' not in st.session_state: st.session_state.user_email = ''
if 'verification_code' not in st.session_state: st.session_state.verification_code = ''

bucket = get_oss_bucket()
statistik, metro_data, ortalama_m2 = load_initial_data(bucket)

# --- Login / Qeydiyyat ---
if st.session_state.verification_step != 'verified':
    st.markdown('<div class="login-container">', unsafe_allow_html=True)
    if st.session_state.verification_step == 'enter_email':
        st.markdown("<h1>Giriş və ya Qeydiyyat</h1>", unsafe_allow_html=True)
        registered_emails = load_registered_users_from_excel(bucket)
        with st.form("email_form"):
            email = st.text_input("Email ünvanınız:", placeholder="Email ünvanınızı daxil edin", label_visibility="collapsed")
            if st.form_submit_button("Davam Et"):
                if re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', email):
                    if email in registered_emails:
                        st.session_state.user_email, st.session_state.verification_step = email, 'verified'
                        st.rerun()
                    else:
                        code = f"{random.randint(100000, 999999):06d}"
                        if send_email("Əmlak Analizatoru - Qeydiyyat Kodu", f"Qeydiyyatı tamamlamaq üçün kodunuz: {code}", [email]):
                            st.session_state.user_email, st.session_state.verification_code, st.session_state.verification_step = email, code, 'enter_code'
                            st.rerun()
                        else: st.error("Email göndərilə bilmədi. Ünvanı yoxlayın.")
                else: st.error("Zəhmət olmasa, düzgün email formatı daxil edin.")

    elif st.session_state.verification_step == 'enter_code':
        st.markdown("<h3>Qeydiyyatı Tamamlayın</h3>", unsafe_allow_html=True)
        st.info(f"**{st.session_state.user_email}** ünvanına göndərilən 6 rəqəmli kodu daxil edin.")
        with st.form("code_form"):
            code_input = st.text_input("Təsdiq kodu:", placeholder="6 rəqəmli kod", max_chars=6, label_visibility="collapsed")
            if st.form_submit_button("Qeydiyyatı Bitir"):
                if code_input == st.session_state.verification_code:
                    save_new_user_to_excel(bucket, st.session_state.user_email)
                    st.session_state.verification_step = 'verified'
                    st.balloons()
                    st.rerun()
                else: st.error("Kod yanlışdır.")
    st.markdown('</div>', unsafe_allow_html=True)

else:
    st.title("Əmlak Analizatoru - İdarə Paneli")
    st.markdown("### Ümumi Baxış")
    st.write(f"Salam, {st.session_state.user_email}!")

    # Statistika
    cols = st.columns(4)
    card_data = [
        ("Ərazi Sayı", len(ortalama_m2), "Analiz edilən"),
        ("Metro Stansiyaları", len(metro_data), "Məlumat bazasında"),
        ("Elan Sayı", len(statistik), "Data bazasında"),
        ("Ortalama Qiymət", f"{ortalama_m2['Ortalama_Qiymet_m2'].median():.0f}", "AZN/m² (Median)")
    ]
    for col, data in zip(cols, card_data):
        title, number, label = data
        with col:
            st.metric(label=title, value=number, delta=label)
