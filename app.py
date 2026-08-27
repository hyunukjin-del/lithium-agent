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

# 원소 표준 정렬 순서 정의
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
    # 액체 ICP 분석 (mg/L)
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
    # 고체 CaCO3 분석 (wt%)
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

# Secrets 키 연동
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
# [2] 강력한 숫자 정제 함수 및 고속 Vision OCR 엔진
# --------------------------------------------------------------------------
def clean_float(val):
    """'95.34g', '1,040', '80℃' 등에서 순수 숫자 float만 추출"""
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

def optimize_image_for_vision(image_bytes):
    """대용량 스마트폰 사진을 1280px 초경량 JPEG로 변환하여 전송 지연 방지"""
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
이미지에서 아래 한국어 수기 항목들을 찾아 해당하는 영문 키에 순수 숫자(float)로 매핑하여 JSON으로 반환하세요.
단위(g, mL, ℃ 등)는 제외하고 숫자만 추출하세요. 적혀있지 않은 항목은 null로 설정하세요.

[수기 표기어와 영문 키 매핑 가이드]
- run_no: 회차, Run 번호
- li2co3_mass: 탄산리튬, LC, Li2CO3, 원료 투입량 (g)
- li2co3_water: 용매수, LC 물, 용해수 (g)
- fresh_cao_mass: 신품 생석회, 신품 CaO, 생석회(신) (g)
- recycled_cao_mass: 재생 생석회, 재생 CaO, 생석회(재) (g)
- slurry_water: 슬러리수, 소화수, 슬러리 조제수 (g)
- temp_c: 반응온도, 온도 (℃)
- time_h: 반응시간, 시간 (h)
- primary_filtrate_mass: 1차 여액, 여액, LiOH 용액 무게 (g)
- primary_filtrate_sg: 여액 비중, 비중 (g/mL)
- primary_filtrate_ph: 여액 pH, pH
- wet_cake_mass: 1차 습케이크, 습케익, CaCO3 무게, 케익 습중량 (g)
- sample_wet: 함수율 습샘플, 습중량 (g)
- sample_dry: 함수율 건샘플, 건중량 (g)
- wash_water_in: 투입 수세수, 세척수, 수세수 투입량 (g)
- wash_sol_mass: 회수 수세액, 수세액 무게 (g)
- wash_sol_sg: 수세액 비중 (g/mL)
- wash_sol_ph: 수세액 pH
- test_dry_cake: 소성 투입 샘플, 건조케익 (g)
- calcined_cao: 소성 후 CaO, 회수 생석회 (g)
- calc_temp: 소성온도, 하소온도 (℃)
- calc_time: 소성시간, 하소시간 (h)

반드시 아래 JSON 포맷으로만 응답하세요:
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
이미지에서 각 원소별 수치를 추출하여 순수 숫자(float)만 JSON으로 반환하세요.

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
