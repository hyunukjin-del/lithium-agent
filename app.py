import streamlit as st
import pandas as pd
import numpy as np
import io
import json
import re
import smtplib
from datetime import datetime
from PIL import Image, ImageOps
import google.generativeai as genai
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# 화학 분자량 상수 (g/mol)
MW_LI2CO3 = 73.89
MW_CAO = 56.08
MW_CAOH2 = 74.09
MW_LIOH = 23.95
MW_CACO3 = 100.09
MW_LI = 6.941
MW_CA = 40.08
MW_NA = 22.99
MW_SI = 28.09
MW_MG = 24.31
MW_K = 39.10

ELEMENT_ORDER = ["Li", "Ca", "Na", "Si", "Mg", "K"]
AGENT_TITLE = "LC-LH전환반응 M/B자동화 및 거동예측 Agent tool"

st.set_page_config(page_title=AGENT_TITLE, page_icon="🧪", layout="wide")

# --------------------------------------------------------------------------
# [1] 기본 세션 상태 초기화
# --------------------------------------------------------------------------
DEFAULT_DATA = {
    "run_no": 1,
    "tab1_run_no": 1,
    "tab2_run_no": 1,
    "li2co3_mass": 95.34,
    "li2co3_water": 1040.0,
    "fresh_cao_mass": 92.42,
    "recycled_cao_mass": 0.0,
    "slurry_water": 831.0,
    "temp_c": 80.0,
    "time_h": 2.0,
    "primary_filtrate_mass": 1646.0,
    "primary_filtrate_sg": 1.035,
    "primary_filtrate_ph": 12.81,
    "wet_cake_mass": 311.0,
    "sample_wet": 27.7,
    "sample_dry": 14.8,
    "wash_water_in": 850.0,
    "wash_sol_mass": 832.0,
    "wash_sol_sg": 1.000,
    "wash_sol_ph": 13.70,
    "test_dry_cake": 40.6,
    "calcined_cao": 23.9,
    "calc_temp": 1000.0,
    "calc_time": 1.0,
    "icp_li_1": 10500.0,
    "icp_ca_1": 120.0,
    "icp_na_1": 45.0,
    "icp_si_1": 8.5,
    "icp_mg_1": 1.2,
    "icp_k_1": 15.0,
    "icp_li_w": 1400.0,
    "icp_ca_w": 80.0,
    "icp_na_w": 6.0,
    "icp_si_w": 2.1,
    "icp_mg_w": 0.3,
    "icp_k_w": 2.0,
    "solid_li_wt": 0.38,
    "solid_ca_wt": 38.20,
    "solid_na_wt": 0.015,
    "solid_si_wt": 0.045,
    "solid_mg_wt": 0.008,
    "solid_k_wt": 0.010,
}

for k, v in DEFAULT_DATA.items():
    if k not in st.session_state:
        st.session_state[k] = v

def sync_run_from_tab1():
    st.session_state.run_no = st.session_state.tab1_run_no
    st.session_state.tab2_run_no = st.session_state.tab1_run_no

def sync_run_from_tab2():
    st.session_state.run_no = st.session_state.tab2_run_no
    st.session_state.tab1_run_no = st.session_state.tab2_run_no

secret_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        secret_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

if "gemini_api_key" not in st.session_state or not st.session_state.gemini_api_key:
    st.session_state.gemini_api_key = secret_key

if "email_recipients" not in st.session_state:
    st.session_state.email_recipients = "user@company.com"
if "email_sender" not in st.session_state:
    st.session_state.email_sender = "sender@gmail.com"
if "email_password" not in st.session_state:
    st.session_state.email_password = ""
if "smtp_server" not in st.session_state:
    st.session_state.smtp_server = "smtp.gmail.com"
if "smtp_port" not in st.session_state:
    st.session_state.smtp_port = 587
if "auto_email_on_save" not in st.session_state:
    st.session_state.auto_email_on_save = True
if "email_logs" not in st.session_state:
    st.session_state.email_logs = []

if "history" not in st.session_state:
    st.session_state.history = pd.DataFrame([
        {
            "회차 (Run)": 1, 
            "구분": "실측치 (Actual)",
            "Li 회수율 (%)": 95.80, 
            "LiOH용액 Li농도 (mg/L)": 10500.0,
            "LiOH용액 농도 (g/L)": round(10500.0 * (MW_LIOH / MW_LI) / 1000, 2),
            "M/B 닫힘율 (%)": 95.88, 
            "하소 감율 LOI (%)": 41.13, 
            "CaO 활성도 (%)": 100.0,
            "신품 CaO 보충량 (g)": 68.52
        }
    ])

if "chat_messages" not in st.session_state:
    st.session_state.chat_messages = [
        {"role": "assistant", "content": f"안녕하세요! **{AGENT_TITLE}**입니다. 수기 일지 사진 인식, LiOH 용액(mg/L) 및 CaCO₃(wt%) 통합 분석, M/B 연산에 대해 질문해 주세요."}
    ]

# --------------------------------------------------------------------------
# [2] 스마트 키 정규화 및 고속 Vision OCR 엔진
# --------------------------------------------------------------------------
def clean_float(val):
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace(",", "")
    match = re.search(r"[-+]?\d*\.?\d+", val_str)
    if match:
        try:
            return float(match.group())
        except ValueError:
            return None
    return None

def normalize_parsed_keys(raw_dict):
    mapping = {
        "run_no": ["run_no", "회차", "run", "회차번호", "실험회차"],
        "li2co3_mass": ["li2co3_mass", "li2co3", "lc", "탄산리튬", "탄산리튬투입량", "li2co3투입량"],
        "li2co3_water": ["li2co3_water", "용매수", "lc물", "용해수", "탄산리튬물"],
        "fresh_cao_mass": ["fresh_cao_mass", "fresh_cao", "신품cao", "생석회신품", "신품생석회"],
        "recycled_cao_mass": ["recycled_cao_mass", "recycled_cao", "재생cao", "생석회재생", "재생생석회"],
        "slurry_water": ["slurry_water", "슬러리수", "소화수", "조제수", "슬러리조제수"],
        "temp_c": ["temp_c", "temp", "온도", "반응온도"],
        "time_h": ["time_h", "time", "시간", "반응시간"],
        "primary_filtrate_mass": ["primary_filtrate_mass", "filtrate_mass", "여액무게", "lioh용액무게", "1차여액", "여액질량"],
        "primary_filtrate_sg": ["primary_filtrate_sg", "filtrate_sg", "여액비중", "비중1"],
        "primary_filtrate_ph": ["primary_filtrate_ph", "filtrate_ph", "여액ph", "ph1"],
        "wet_cake_mass": ["wet_cake_mass", "wet_cake", "습케이크", "caco3무게", "습중량", "caco3습중량"],
        "sample_wet": ["sample_wet", "함수율습", "샘플습중량", "습중량샘플"],
        "sample_dry": ["sample_dry", "함수율건", "샘플건중량", "건중량샘플"],
        "wash_water_in": ["wash_water_in", "수세수", "세척수", "수세수투입량", "세척수무게"],
        "wash_sol_mass": ["wash_sol_mass", "수세액무게", "회수수세액", "수세액질량"],
        "wash_sol_sg": ["wash_sol_sg", "수세액비중"],
        "wash_sol_ph": ["wash_sol_ph", "수세액ph"],
        "test_dry_cake": ["test_dry_cake", "소성투입", "건조케익", "caco3샘플"],
        "calcined_cao": ["calcined_cao", "회수cao", "소성후cao", "cao회수량"],
        "calc_temp": ["calc_temp", "소성온도", "하소온도"],
        "calc_time": ["calc_time", "소성시간", "하소시간"],
        "icp_li_1": ["icp_li_1", "li_1", "li_여액"], "icp_ca_1": ["icp_ca_1", "ca_1", "ca_여액"],
        "icp_na_1": ["icp_na_1", "na_1", "na_여액"], "icp_si_1": ["icp_si_1", "si_1", "si_여액"],
        "icp_mg_1": ["icp_mg_1", "mg_1", "mg_여액"], "icp_k_1": ["icp_k_1", "k_1", "k_여액"],
        "icp_li_w": ["icp_li_w", "li_w", "li_수세"], "icp_ca_w": ["icp_ca_w", "ca_w", "ca_수세"],
        "icp_na_w": ["icp_na_w", "na_w", "na_수세"], "icp_si_w": ["icp_si_w", "si_w", "si_수세"],
        "icp_mg_w": ["icp_mg_w", "mg_w", "mg_수세"], "icp_k_w": ["icp_k_w", "k_w", "k_수세"],
        "solid_li_wt": ["solid_li_wt", "li_wt", "li_고체"], "solid_ca_wt": ["solid_ca_wt", "ca_wt", "ca_고체"],
        "solid_na_wt": ["solid_na_wt", "na_wt", "na_고체"], "solid_si_wt": ["solid_si_wt", "si_wt", "si_고체"],
        "solid_mg_wt": ["solid_mg_wt", "mg_wt", "mg_고체"], "solid_k_wt": ["solid_k_wt", "k_wt", "k_고체"]
    }
    
    normalized = {}
    for target_key, aliases in mapping.items():
        for raw_k, raw_v in raw_dict.items():
            raw_k_lower = str(raw_k).lower().strip()
            if raw_k_lower == target_key.lower() or raw_k_lower in [a.lower() for a in aliases]:
                val = clean_float(raw_v)
                if val is not None:
                    normalized[target_key] = val
                    break
    return normalized

def optimize_image_for_vision(image_bytes):
    img = Image.open(io.BytesIO(image_bytes))
    img = ImageOps.exif_transpose(img)
    if img.mode != "RGB":
        img = img.convert("RGB")
    max_dim = 1280
    if max(img.size) > max_dim:
        img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
    return img

def get_gemini_model_candidates(api_key):
    genai.configure(api_key=api_key)
    return [
        "gemini-1.5-flash",
        "gemini-2.0-flash",
        "gemini-1.5-flash-latest",
        "gemini-1.5-pro"
    ]

def parse_image_with_vision(image_bytes, doc_type="lab_note"):
    api_key = st.session_state.gemini_api_key.strip()
    if not api_key:
        return None, "사이드바에 Google Gemini API Key를 입력하거나 Secrets에 등록해 주세요."

    try:
        img = optimize_image_for_vision(image_bytes)

        if doc_type == "lab_note":
            prompt = """당신은 탄산리튬(LC) 가성화 및 Ca-Loop 습식 제련 공정의 수기 실험 일지 전문 데이터 분석가입니다.
이미지에서 아래 항목들의 수치를 정확히 판독하여 JSON 형식으로만 응답해 주세요. 단위나 문자는 제외하고 순수 숫자(float)만 추출해야 합니다.

{
  "run_no": number, "li2co3_mass": number, "li2co3_water": number,
  "fresh_cao_mass": number, "recycled_cao_mass": number, "slurry_water": number,
  "temp_c": number, "time_h": number, "primary_filtrate_mass": number,
  "primary_filtrate_sg": number, "primary_filtrate_ph": number,
  "wet_cake_mass": number, "sample_wet": number, "sample_dry": number,
  "wash_water_in": number, "wash_sol_mass": number, "wash_sol_sg": number, "wash_sol_ph": number,
  "test_dry_cake": number, "calcined_cao": number, "calc_temp": number, "calc_time": number
}"""
        else:
            prompt = """이 이미지는 LiOH 용액 및 수세액(mg/L), 그리고 CaCO3 고체(wt%) 성적서입니다.
추출 가능한 수치를 아래 JSON 포맷으로 반환해 주세요.

{
  "icp_li_1": number, "icp_ca_1": number, "icp_na_1": number, "icp_si_1": number, "icp_mg_1": number, "icp_k_1": number,
  "icp_li_w": number, "icp_ca_w": number, "icp_na_w": number, "icp_si_w": number, "icp_mg_w": number, "icp_k_w": number,
  "solid_li_wt": number, "solid_ca_wt": number, "solid_na_wt": number, "solid_si_wt": number, "solid_mg_wt": number, "solid_k_wt": number
}"""

        model_names = get_gemini_model_candidates(api_key)
        last_error = None

        for m_name in model_names:
            try:
                model = genai.GenerativeModel(m_name)
                response = model.generate_content(
                    [img, prompt],
                    generation_config={"response_mime_type": "application/json"},
                    request_options={"timeout": 30}
                )
                if response and response.text:
                    raw_text = response.text.replace("```json", "").replace("
