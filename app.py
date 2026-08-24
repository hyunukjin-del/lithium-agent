import streamlit as st
import pandas as pd
import plotly.express as px
import json
from openai import OpenAI

# =========================================================
# 1. 페이지 설정
# =========================================================
st.set_page_config(
    page_title="Li2CO3 가성화 & CaO 칼슘 리사이클링 AI 에이전트",
    page_icon="⚗️",
    layout="wide"
)

if "messages" not in st.session_state:
    st.session_state.messages = [
        {"role": "assistant", "content": "안녕하세요! 탄산리튬 가성화 및 CaO 리사이클링 공정 AI 에이전트입니다. 상단에 API 키를 입력하시면 GPT-4o가 정밀 진단을 시작합니다."}
    ]

if "cycle_history" not in st.session_state:
    st.session_state.cycle_history = [
        {"Cycle": 1, "Feed_Li2CO3_g": 10.0, "Fresh_CaO_g": 7.97, "Recycle_CaO_g": 0.00, "Li_Recovery_pct": 98.0, "Conversion_pct": 96.5, "Solid_Si_wt": 0.58, "Solid_Mg_wt": 0.31, "Solid_Al_wt": 0.24, "Filtration_Time_min": 3.2, "Dry_CaCO3_g": 12.95},
        {"Cycle": 2, "Feed_Li2CO3_g": 10.0, "Fresh_CaO_g": 1.17, "Recycle_CaO_g": 6.80, "Li_Recovery_pct": 96.8, "Conversion_pct": 95.2, "Solid_Si_wt": 0.85, "Solid_Mg_wt": 0.46, "Solid_Al_wt": 0.38, "Filtration_Time_min": 3.9, "Dry_CaCO3_g": 12.90},
        {"Cycle": 3, "Feed_Li2CO3_g": 10.0, "Fresh_CaO_g": 1.21, "Recycle_CaO_g": 6.76, "Li_Recovery_pct": 94.5, "Conversion_pct": 93.8, "Solid_Si_wt": 1.15, "Solid_Mg_wt": 0.62, "Solid_Al_wt": 0.52, "Filtration_Time_min": 4.8, "Dry_CaCO3_g": 12.85},
        {"Cycle": 4, "Feed_Li2CO3_g": 10.0, "Fresh_CaO_g": 1.27, "Recycle_CaO_g": 6.70, "Li_Recovery_pct": 92.1, "Conversion_pct": 90.5, "Solid_Si_wt": 1.38, "Solid_Mg_wt": 0.77, "Solid_Al_wt": 0.64, "Filtration_Time_min": 6.2, "Dry_CaCO3_g": 12.78},
        {"Cycle": 5, "Feed_Li2CO3_g": 10.0, "Fresh_CaO_g": 1.32, "Recycle_CaO_g": 6.65, "Li_Recovery_pct": 89.4, "Conversion_pct": 87.2, "Solid_Si_wt": 1.58, "Solid_Mg_wt": 0.91, "Solid_Al_wt": 0.73, "Filtration_Time_min": 8.5, "Dry_CaCO3_g": 12.70},
    ]

# =========================================================
# 2. 화학공학 도구
# =========================================================
def calculate_reaction_mass_balance(feed_li2co3_g, filtrate_li_g_l=1.78, filtrate_vol_l=1.05, recovered_dry_caco3_g=12.85, **kwargs):
    MW_LI2CO3, MW_CACO3, MW_LIOH, MW_LI = 73.89, 100.09, 23.95, 6.94
    feed_moles = feed_li2co3_g / MW_LI2CO3
    inlet_li_g = feed_li2co3_g * (2 * MW_LI / MW_LI2CO3)
    theo_caco3_g = feed_moles * MW_CACO3
    theo_lioh_g = 2 * feed_moles * MW_LIOH
    recovered_li_g = filtrate_li_g_l * filtrate_vol_l
    recovered_lioh_g = recovered_li_g * (MW_LIOH / MW_LI)

    return {
        "inlet_li_g": round(inlet_li_g, 3),
        "recovered_li_g": round(recovered_li_g, 3),
        "li_recovery_pct": round((recovered_li_g / inlet_li_g) * 100, 2),
        "conversion_pct": round((recovered_lioh_g / theo_lioh_g) * 100, 2),
        "theoretical_caco3_g": round(theo_caco3_g, 2),
        "actual_caco3_g": round(recovered_dry_caco3_g, 2),
        "equivalent_lioh_g": round(recovered_lioh_g, 2)
    }

def diagnose_impurity_and_operability(solid_si_wt, solid_al_wt, solid_mg_wt, filtration_time_min, **kwargs):
    diagnostics = []
    status = "Normal"
    if solid_si_wt >= 1.20:
        diagnostics.append(f"Si 농도({solid_si_wt} wt%) 임계치(1.20 wt%) 초과: C-S-H 형성 및 활성도 저하")
        status = "Critical"
    if solid_al_wt >= 0.50:
        diagnostics.append(f"Al 농도({solid_al_wt} wt%) 상승: 소성 시 비활성 클링커 형성 위험")
        if status != "Critical": status = "Warning"
    if solid_mg_wt >= 0.60 or filtration_time_min >= 6.0:
        diagnostics.append(f"여과시간({filtration_time_min}분) 및 Mg 농도({solid_mg_wt} wt%) 상승: 여과포 눈막힘 발생")
        status = "Critical"

    return {
        "process_status": status,
        "diagnostics": diagnostics if diagnostics else ["모든 수치 정상"]
    }

def calculate_optimal_purge_and_makeup(target_feed_li2co3_g, current_recycled_cao_g, solid_si_wt, **kwargs):
    req_total_cao = (target_feed_li2co3_g / 73.89) * 56.08 * 1.05
    purge_ratio = 0.05 if solid_si_wt < 0.8 else (0.15 if solid_si_wt < 1.2 else 0.30)
    purged_cao = current_recycled_cao_g * purge_ratio
    usable_recycled_cao = current_recycled_cao_g - purged_cao
    makeup_fresh_cao = max(0.0, req_total_cao - usable_recycled_cao)

    return {
        "required_total_cao_g": round(req_total_cao, 2),
        "recommended_purge_ratio_pct": round(purge_ratio * 100, 1),
        "purge_cao_amount_g": round(purged_cao, 2),
        "reusable_recycled_cao_g": round(usable_recycled_cao, 2),
        "makeup_fresh_cao_g": round(makeup_fresh_cao, 2)
    }

TOOL_DEFINITIONS = [
    {
        "type": "function",
        "function": {
            "name": "calculate_reaction_mass_balance",
            "description": "탄산리튬 가성화 반응의 리튬 회수율, 반응 전환율 및 CaCO3 수율 계산",
            "parameters": {
                "type": "object",
                "properties": {"feed_li2co3_g": {"type": "number"}},
                "required": ["feed_li2co3_g"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "diagnose_impurity_and_operability",
            "description": "재생 CaO 불순물(Si, Al, Mg) 상태 및 여과 지연 진단",
            "parameters": {
                "type": "object",
                "properties": {
                    "solid_si_wt": {"type": "number"},
                    "solid_al_wt": {"type": "number"},
                    "solid_mg_wt": {"type": "number"},
                    "filtration_time_min": {"type": "number"}
                },
                "required": ["solid_si_wt", "solid_al_wt", "solid_mg_wt", "filtration_time_min"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "calculate_optimal_purge_and_makeup",
            "description": "최적 고형분 퍼지(Purge) 비율 및 신규 Fresh CaO 보충량 산출",
            "parameters": {
                "type": "object",
                "properties": {
                    "target_feed_li2co3_g": {"type": "number"},
                    "current_recycled_cao_g": {"type": "number"},
                    "solid_si_wt": {"type": "number"}
                },
                "required": ["target_feed_li2co3_g", "current_recycled_cao_g", "solid_si_wt"]
            }
        }
    }
]

AVAILABLE_FUNCTIONS = {
    "calculate_reaction_mass_balance": calculate_reaction_mass_balance,
    "diagnose_impurity_and_operability": diagnose_impurity_and_operability,
    "calculate_optimal_purge_and_makeup": calculate_optimal_purge_and_makeup
}

# =========================================================
# 3. 메인 상단 화면 (API 키 입력창을 메인에 배치)
# =========================================================
st.title("⚗️ Li₂CO₃ 가성화 & CaO 칼슘 루핑 AI 공정 에이전트")

# 화면 메인에 바로 보이는 API 키 입력 카드
with st.container():
    col_k1, col_k2 = st.columns([3, 1])
    with col_k1:
        api_key_input = st.text_input("🔑 OpenAI API Key 입력창", value="", type="password", placeholder="sk-proj-... 키를 여기에 붙여넣으세요")
    with col_k2:
        model_name = st.selectbox("LLM 모델", ["gpt-4o", "gpt-4o-mini"], index=0)

clean_key = api_key_input.strip().replace('"', '').replace("'", "")
is_valid_key = clean_key.startswith("sk-") and clean_key.isascii() and len(clean_key) > 20

if is_valid_key:
    st.success("✅ OpenAI API 연결 완료! 아래에서 질문을 클릭하거나 입력하세요.")
else:
    st.info("💡 위 입력창에 OpenAI API Key('sk-...')를 붙여넣으시면 에이전트가 활성화됩니다.")

st.markdown("---")

df = pd.DataFrame(st.session_state.cycle_history)
last_row = df.iloc[-1]

col1, col2, col3, col4 = st.columns(4)
col1.metric("현재 진행 회차", f"Cycle {int(last_row['Cycle'])}")
col1.metric("최근 Li 회수율", f"{last_row['Li_Recovery_pct']}%")
col2.metric("고상 Si 축적 농도", f"{last_row['Solid_Si_wt']} wt%")
col3.metric("총 여과 소요시간", f"{last_row['Filtration_Time_min']} min")
col4.metric("재생 CaO 사용률", f"{round(last_row['Recycle_CaO_g'] / (last_row['Fresh_CaO_g'] + last_row['Recycle_CaO_g']) * 100, 1)}%")

# =========================================================
# 4. 메인 탭
# =========================================================
tab_chat, tab_charts, tab_data = st.tabs(["💬 AI 공정 에이전트 대화", "📈 불순물 및 공정 시각화", "📋 회차별 실측 데이터"])

with tab_chat:
    st.subheader("🤖 공정 진단 및 처방 대화창")
    
    st.markdown("**💡 추천 빠른 질문 버튼:**")
    qc1, qc2, qc3 = st.columns(3)
    quick_input = None
    if qc1.button("📌 최근 Cycle 종합 진단"):
        quick_input = f"현재 Cycle {int(last_row['Cycle'])}까지 진행됐어. 고상 Si는 {last_row['Solid_Si_wt']}%, Mg는 {last_row['Solid_Mg_wt']}%, Al은 {last_row['Solid_Al_wt']}%, 여과시간은 {last_row['Filtration_Time_min']}분이야. 공정 이상 유무를 진단해줘."
    if qc2.button("⚠️ Si 농축 원인 및 여과 지연 대책"):
        quick_input = "Si와 Mg 농축이 여과 속도와 전환율에 미치는 영향을 분석하고 해결 방안을 알려줘."
    if qc3.button("🎯 다음 회차 Purge 및 CaO 처방"):
        quick_input = f"다음 회차에 Li2CO3 10g을 처리할 예정이야. 현재 재생 CaO가 {last_row['Recycle_CaO_g']}g이고 Si가 {last_row['Solid_Si_wt']}%인데 최적 퍼지율과 Fresh CaO 보충량을 계산해줘."

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    user_query = st.chat_input("질문이나 분석할 실험 데이터를 입력하세요...")
    if quick_input:
        user_query = quick_input

    if user_query:
        if not is_valid_key:
            st.error("⚠️ 화면 맨 위 '🔑 OpenAI API Key 입력창'에 sk-... 키를 먼저 붙여넣어 주세요.")
        else:
            st.session_state.messages.append({"role": "user", "content": user_query})
            with st.chat_message("user"):
                st.markdown(user_query)

            with st.chat_message("assistant"):
                with st.spinner("GPT-4o가 공정 수지 도구를 호출하여 분석 중입니다..."):
                    try:
                        client = OpenAI(api_key=clean_key)
                        system_prompt = (
                            "당신은 탄산리튬 가성화(Causticizing) 및 CaO 칼슘 루핑 전문 공정 최적화 AI 에이전트입니다. "
                            "제공된 함수(Tool)를 적극 활용하여 정확한 수치 계산과 진단을 수행하세요. "
                            "답변은 한국어로 작성하며, 공정 엔지니어 관점에서 명확하고 실행 가능한 수치와 처방(Prescription)을 제시하세요."
                        )
                        messages_payload = [{"role": "system", "content": system_prompt}]
                        for m in st.session_state.messages:
                            messages_payload.append({"role": m["role"], "content": str(m["content"])})

                        response = client.chat.completions.create(
                            model=model_name,
                            messages=messages_payload,
                            tools=TOOL_DEFINITIONS,
                            tool_choice="auto",
                            temperature=0
                        )
                        response_msg = response.choices[0].message

                        if response_msg.tool_calls:
                            messages_payload.append({
                                "role": "assistant",
                                "tool_calls": [
                                    {
                                        "id": tc.id,
                                        "type": "function",
                                        "function": {
                                            "name": tc.function.name,
                                            "arguments": tc.function.arguments
                                        }
                                    } for tc in response_msg.tool_calls
                                ]
                            })
                            for tool_call in response_msg.tool_calls:
                                fn_name = tool_call.function.name
                                fn_args = json.loads(tool_call.function.arguments)
                                if fn_name in AVAILABLE_FUNCTIONS:
                                    tool_result = AVAILABLE_FUNCTIONS[fn_name](**fn_args)
                                    messages_payload.append({
                                        "role": "tool",
                                        "tool_call_id": tool_call.id,
                                        "content": json.dumps(tool_result, ensure_ascii=True)
                                    })
                            final_response = client.chat.completions.create(
                                model=model_name,
                                messages=messages_payload,
                                temperature=0
                            )
                            answer = final_response.choices[0].message.content
                        else:
                            answer = response_msg.content

                        st.markdown(answer)
                        st.session_state.messages.append({"role": "assistant", "content": answer})
                    except Exception as e:
                        st.error(f"오류가 발생했습니다: {str(e)}")

with tab_charts:
    st.subheader("📊 다주기(n-Cycle) 거동 추이 분석")
    chart_col1, chart_col2 = st.columns(2)
    with chart_col1:
        fig_imp = px.line(df, x="Cycle", y=["Solid_Si_wt", "Solid_Mg_wt", "Solid_Al_wt"], markers=True, title="재생 CaO 내 불순물 축적 추이 (wt%)")
        fig_imp.add_hline(y=1.20, line_dash="dash", line_color="red", annotation_text="Si 임계치 (1.20 wt%)")
        st.plotly_chart(fig_imp, use_container_width=True)
        fig_time = px.bar(df, x="Cycle", y="Filtration_Time_min", title="회차별 여과 시간 (분)", color="Filtration_Time_min", color_continuous_scale="Reds")
        st.plotly_chart(fig_time, use_container_width=True)
    with chart_col2:
        fig_rec = px.line(df, x="Cycle", y=["Li_Recovery_pct", "Conversion_pct"], markers=True, title="회차별 Li 회수율 및 반응 전환율 (%)")
        st.plotly_chart(fig_rec, use_container_width=True)
        fig_cao = px.bar(df, x="Cycle", y=["Recycle_CaO_g", "Fresh_CaO_g"], title="회차별 투입 CaO 구성 (재생 vs 신규)", barmode="stack")
        st.plotly_chart(fig_cao, use_container_width=True)

with tab_data:
    st.subheader("📋 전체 실험 회차 데이터")
    st.dataframe(df, use_container_width=True)
