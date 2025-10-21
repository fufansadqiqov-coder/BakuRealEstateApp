# =========================================================
# BÜTÜN KOD
# =========================================================
import streamlit as st
import requests
from requests.exceptions import RequestException
from bs4 import BeautifulSoup
import pandas as pd
import re
import time
from rapidfuzz import process, fuzz
from geopy.distance import geodesic
from email.message import EmailMessage
import smtplib
from io import BytesIO
import oss2
from typing import Dict, Any, List, Tuple, Optional
import plotly.express as px
import random
import concurrent.futures

# =========================================================
# KONFİQURASİYA
try:
    ACCESS_KEY_ID = st.secrets["ACCESS_KEY_ID"]
    ACCESS_KEY_SECRET = st.secrets["ACCESS_KEY_SECRET"]
    APP_PASSWORD = st.secrets["APP_PASSWORD"]
except (KeyError, FileNotFoundError):
    st.error("Secrets konfiqurasiyası tapılmadı. Zəhmət olmasa `.streamlit/secrets.toml` faylını yaradın.")
    st.stop()

ENDPOINT = "oss-ap-southeast-1.aliyuncs.com"
BUCKET_NAME = "emlak-bot-demo"
HOME_SALES_KEY = "Home Sales Statistika and Location.xlsx"
OUTPUT_KEY = "Depo.xlsx"
REGISTERED_USERS_EXCEL_KEY = "Gmail.xlsx"
SENDER_EMAIL = "fufansadqiqov@gmail.com"
ADMIN_EMAILS = ["ugurnihat123321@gmail.com", "eyvazazim@gmail.com"]
HEADERS = { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36" }
BASE_URL = "https://bina.az"

# =========================================================
# YARDIMÇI FUNKSİYALAR
@st.cache_resource
def get_oss_bucket() -> oss2.Bucket:
    try:
        auth = oss2.Auth(ACCESS_KEY_ID, ACCESS_KEY_SECRET)
        bucket = oss2.Bucket(auth, ENDPOINT, BUCKET_NAME)
        bucket.get_bucket_info()
        return bucket
    except Exception as e:
        st.error(f"Aliyun OSS bağlantı xətası: {e}")
        st.stop()

@st.cache_data
def load_initial_data(_bucket: oss2.Bucket) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
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

def get_depo_df(_bucket: oss2.Bucket) -> pd.DataFrame:
    try:
        if _bucket.object_exists(OUTPUT_KEY):
            obj = _bucket.get_object(OUTPUT_KEY)
            return pd.read_excel(BytesIO(obj.read()))
        return pd.DataFrame(columns=["Link"])
    except Exception as e:
        st.error(f"Depo faylı oxunarkən xəta baş verdi: {e}")
        return pd.DataFrame(columns=["Link"])

def save_df_to_oss(_bucket: oss2.Bucket, df: pd.DataFrame):
    try:
        stream = BytesIO()
        df.to_excel(stream, index=False)
        stream.seek(0)
        _bucket.put_object(OUTPUT_KEY, stream)
    except Exception as e:
        st.error(f"Nəticələr OSS-ə yazılanda xəta baş verdi: {e}")

def send_email(subject: str, body: str, receivers: list):
    if not all([SENDER_EMAIL, APP_PASSWORD, receivers]):
        st.warning("Email konfiqurasiyası tam deyil, bildiriş göndərilmədi.")
        return
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
# SCRAPING & ANALİZ
def scrape_item_details(session: requests.Session, link: str) -> Optional[Dict[str, Any]]:
    try:
        r = session.get(link, headers=HEADERS, timeout=15)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        map_div = soup.find("div", {"id": "item_map"})
        details = {
            "Qiymet": soup.find("div", class_="product-price__i").get_text(strip=True) if soup.find("div", class_="product-price__i") else None,
            "Erazi": soup.find("h1", class_="product-title").get_text(strip=True).split(",")[-1].strip() if soup.find("h1", class_="product-title") else None,
            "Link": link,
            "Lat": float(map_div.get("data-lat")) if map_div and map_div.get("data-lat") else None,
            "Lng": float(map_div.get("data-lng")) if map_div and map_div.get("data-lng") else None,
            "Elan yerlesdirilme tarixi": soup.find_all("span", class_="product-statistics__i-text")[1].get_text(strip=True) if len(soup.find_all("span", class_="product-statistics__i-text")) > 1 else None
        }
        keys = [i.get_text(strip=True) for i in soup.find_all("label", class_="product-properties__i-name")]
        values = [k.get_text(strip=True) for k in soup.find_all("span", class_="product-properties__i-value")]
        details.update(dict(zip(keys, values)))
        return details
    except Exception:
        return None

def process_and_analyze_item(raw_details: Dict[str, Any], ortalama_m2: pd.DataFrame, metro_data: pd.DataFrame) -> Optional[Dict[str, Any]]:
    if not raw_details: return None
    def clean_value(value):
        if isinstance(value, str):
            if any(c in value for c in ["AZN","USD","EUR"]):
                return int("".join(re.findall(r'\d+', value)))
            if "m²" in value:
                match = re.search(r'(\d+\.?\d*)', value)
                return float(match.group(1)) if match else value
        return value
    processed = {k: clean_value(v) for k,v in raw_details.items()}
    if not all([processed.get("Qiymet"), processed.get("Sahə"), processed.get("Erazi")]) or processed.get("Sahə")==0: return None
    processed['qiymet_m2'] = round(processed['Qiymet']/processed['Sahə'])
    area_avg_price_row = ortalama_m2[ortalama_m2["Erazi"].str.lower()==processed["Erazi"].lower()]
    avg_price = area_avg_price_row["Ortalama_Qiymet_m2"].values[0] if not area_avg_price_row.empty else None
    processed['ortalama_qiymet_m2'] = round(avg_price) if avg_price else None
    # Metro məsafəsi
    lat,lng=processed.get("Lat"),processed.get("Lng")
    mesafe_km = None
    if pd.notna(lat) and pd.notna(lng):
        try:
            metro_distances=[geodesic((lat,lng),(row["Lat"],row["Long"])).km for _,row in metro_data.iterrows()]
            if metro_distances: mesafe_km=min(metro_distances)
        except Exception: pass
    processed['Metroya məsafə (km)'] = round(mesafe_km,2) if mesafe_km is not None else None
    # Seqment
    if avg_price and processed['qiymet_m2']<avg_price:
        processed['segment'] = "🚀 Xüsusi Fürsət" if mesafe_km is not None and mesafe_km<0.5 else "🏠 Yeni Fürsət"
    else:
        processed['segment'] = "Standart"
    return processed

# =========================================================
# Burada artıq run_process və UI hissəsi birləşdirilə bilər (sənin əvvəlki kodla eyni, metro məsafəsi artıq daxil edilib)

# --- STREAMLIT UI ---
st.set_page_config(page_title="Əmlak Analizatoru", layout="wide", page_icon="🏙️")

# (CSS kodu olduğu kimi qalır)
st.markdown("""
<style>
:root {
    --primary-gradient: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    --secondary-gradient: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
    --accent-gradient-1: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
    --accent-gradient-2: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%);
    --sidebar-gradient: linear-gradient(180deg, #6A11CB 0%, #2575FC 100%);
    --text-primary: #2c3e50;
    --text-light: #ffffff;
    --bg-light-alpha: rgba(255, 255, 255, 0.95);
    --border-light: rgba(255, 255, 255, 0.3);
    --shadow-soft: 0 15px 50px rgba(0, 0, 0, 0.1);
    --shadow-strong: 0 20px 60px rgba(0, 0, 0, 0.2);
}
[data-testid="stAppViewContainer"] > .main {
    background: var(--primary-gradient);
    background-attachment: fixed;
}
h1, h2, h3, h4 { color: var(--text-primary) !important; font-weight: 700 !important; }
.login-container {
    background: var(--bg-light-alpha);
    backdrop-filter: blur(20px);
    border-radius: 25px;
    border: 1px solid var(--border-light);
    padding: 3rem;
    max-width: 500px;
    margin: 5rem auto;
    box-shadow: var(--shadow-strong);
}
.main-app-container {
    background: var(--bg-light-alpha);
    backdrop-filter: blur(20px);
    border-radius: 25px;
    padding: 2.5rem;
    margin: 1rem;
    box-shadow: var(--shadow-soft);
    border: 1px solid var(--border-light);
}
[data-testid="stSidebar"] { background: var(--sidebar-gradient) !important; }
[data-testid="stSidebar"] * { color: var(--text-light) !important; }
.stat-card {
    padding: 1.5rem; border-radius: 20px; color: var(--text-light);
    box-shadow: 0 15px 35px rgba(0,0,0,0.1); height: 180px;
    position: relative; text-align: center;
}
.stat-card h3 {
    font-size: 1.1rem; font-weight: 700; text-transform: uppercase;
    letter-spacing: 1px; color: var(--text-light) !important;
    position: absolute; top: 1.5rem; left: 1.5rem; right: 1.5rem;
}
.stat-card .number {
    font-size: 2.8rem; font-weight: 800; line-height: 1;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); width: 100%;
}
.stat-card .label {
    font-size: 0.9rem; opacity: 0.9; font-weight: 500;
    position: absolute; bottom: 1.5rem; left: 1.5rem; right: 1.5rem;
}
.card-1 { background: var(--primary-gradient); }
.card-2 { background: var(--secondary-gradient); }
.card-3 { background: var(--accent-gradient-1); }
.card-4 { background: var(--accent-gradient-2); }
</style>
""", unsafe_allow_html=True)

# --- Session State İdarəetməsi ---
if 'verification_step' not in st.session_state: st.session_state.verification_step = 'enter_email'
if 'user_email' not in st.session_state: st.session_state.user_email = ''
if 'verification_code' not in st.session_state: st.session_state.verification_code = ''

# --- Əsas Proqram Məntiqi ---
bucket = get_oss_bucket()
statistik, metro_data, ortalama_m2 = load_initial_data(bucket)

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

else: # İstifadəçi giriş edibsə
    st.markdown('<div class="main-app-container">', unsafe_allow_html=True)
    with st.sidebar:
        st.info(f"Giriş edildi:\n**{st.session_state.user_email}**")
        max_pages = st.slider("Yoxlanılacaq səhifə sayı", 1, 20, 3)
        all_areas = sorted(ortalama_m2["Erazi"].unique().tolist())
        selected_areas = st.multiselect("Ərazi filtri", options=all_areas)
        if st.button("🚀 Axtarışı Başlat", use_container_width=True):
            run_process(bucket, metro_data, ortalama_m2, max_pages, selected_areas)
    
    st.title("İdarə Paneli")
    st.markdown("### Ümumi Baxış")
    cols = st.columns(4)
    card_classes = ["card-1", "card-2", "card-3", "card-4"]
    card_data = [
        ("Ərazi Sayı", len(ortalama_m2), "Analiz edilən"),
        ("Metro Stansiyaları", len(metro_data), "Məlumat bazasında"),
        ("Elan Sayı", len(statistik), "Data bazasında"),
        ("Ortalama Qiymət", f"{ortalama_m2['Ortalama_Qiymet_m2'].median():.0f}", "AZN/m² (Median)")
    ]
    for i, col in enumerate(cols):
        with col:
            title, number, label = card_data[i]
            st.markdown(f'<div class="stat-card {card_classes[i]}"><h3>{title}</h3><p class="number">{number}</p><p class="label">{label}</p></div>', unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1.5])
    with col1:
        st.markdown("### Ən Bahalı Ərazilər (Qiymət/m²)")
        top_areas = ortalama_m2.nlargest(10, "Ortalama_Qiymet_m2")
        fig_bar = px.bar(top_areas.sort_values("Ortalama_Qiymet_m2", ascending=True), y="Erazi", x="Ortalama_Qiymet_m2", orientation='h', color="Ortalama_Qiymet_m2", color_continuous_scale=px.colors.sequential.PuBu, template="plotly_white", text="Ortalama_Qiymet_m2")
        fig_bar.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', yaxis_title=None, xaxis_title="AZN/m²", font_color="#2c3e50", height=450)
        fig_bar.update_traces(texttemplate='%{x:.0f}', textposition='outside')
        st.plotly_chart(fig_bar, use_container_width=True)
        
    with col2:
        st.markdown("### Zamanla Qiymət Trendi")
        area_options = sorted(statistik['Erazi'].unique().tolist())
        default_index = area_options.index('Nəsimi r.') if 'Nəsimi r.' in area_options else 0
        selected_area_for_trend = st.selectbox("Ərazi seçin:", options=area_options, index=default_index)
        
        if selected_area_for_trend:
            trend_df = statistik[statistik['Erazi'] == selected_area_for_trend].copy()
            if 'Elan yerlesdirilme tarixi' in trend_df.columns:
                trend_df['Date'] = pd.to_datetime(trend_df['Elan yerlesdirilme tarixi'], format='%d.%m.%Y, %H:%M', errors='coerce').dt.date
                trend_df.dropna(subset=['Date', 'Qiymet_m2'], inplace=True)
                daily_avg_price = trend_df.groupby('Date')['Qiymet_m2'].mean().reset_index()
                
                if len(daily_avg_price) > 1:
                    fig_line = px.line(daily_avg_price, x='Date', y='Qiymet_m2', template="plotly_white", color_discrete_sequence=['#764ba2'])
                    fig_line.update_traces(mode='lines+markers', line_shape='spline', line=dict(width=3))
                    fig_line.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', yaxis_title=None, xaxis_title=None, font_color="#2c3e50", height=450)
                    st.plotly_chart(fig_line, use_container_width=True)
                else: st.info("Seçilmiş ərazi üçün trend analizi aparmaq məqsədilə kifayət qədər data yoxdur.")
            else: st.warning("Trend analizi üçün 'Elan yerlesdirilme tarixi' sütunu tapılmadı.")
            

    st.markdown('</div>', unsafe_allow_html=True)
