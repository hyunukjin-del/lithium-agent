import streamlit as st
import pandas as pd
import plotly.express as px
import json
from openai import OpenAI

# =========================================================
# 1. 페이지 기본 설정 및 세션 초기화
# =========================================================
st.set_page_config(
    page_title="Li2CO3 가성화 & CaO 칼슘루핑 AI 에이전트",
    page_icon="⚗️",
    layout="wide"
)

if "messages" not in st.session_state:
    st.session_state.messages = [
        {"role": "assistant", "content": "안녕하세요! 탄산리튬 가성화 및 CaO 리사이클링 공정 에이전트입니다. 실험 데이터 진단이나 다음 회차 투입량 처방이 필요하시면 언제든 질문해 주세요."}
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
# 2. 화학공학 정밀 연산 함수
# =========================================================
def calculate_reaction_mass_balance(feed_li2co3_g, filtrate_li_g_l, filtrate_vol_l, recovered_dry_caco3_g, **kwargs):
    MW_LI2CO3, MW_CACO3, MW_LIOH, MW_LI = 73.89, 100.09, 23.95, 6.94
    feed_moles = feed_li2co3_g / MW_LI2CO3
    inlet_li_g = feed_li2co3_g * (2 * MW_LI / MW_LI2CO3)
    theo_caco3_g = feed_moles * MW_CACO3
    theo_lioh_g = 2 * feed_moles * MW_LIOH

    recovered_li_g = filtrate_li_g_l * filtrate_vol_l
    recovered_lioh_g = recovered_li_g * (MW_LIOH / MW_LI)
    li_recovery_pct = (recovered_li_g / inlet_li_g) * 100
    conversion_pct = (recovered_lioh_g / theo_lioh_g) * 100
    caco3_yield_pct = (recovered_dry_caco3_g / theo_caco3_g) * 100

    return {
        "input_li_g": round(inlet_li_g, 3),
        "recovered_li_g": round(recovered_li_g, 3),
        "li_recovery_pct": round(li_recovery_pct, 2),
        "conversion_pct": round(conversion_pct, 2),
        "theoretical_caco3_g": round(theo_caco3_g, 2),
        "actual_caco3_g": round(recovered_dry_caco3_g, 2),
        "caco3_yield_pct": round(caco3_yield_pct, 2),
        "equivalent_lioh_g": round(recovered_lioh_g, 2)
    }

def diagnose_impurity_and_operability(solid_si_wt, solid_al_wt, solid_mg_wt, filtration_time_min, **kwargs):
    diagnostics = []
    status = "Normal"

    if solid_si_wt >= 1.20:
        diagnostics.append(f"Si 농도({solid_si_wt} wt%)가 임계치(1.20 wt%) 초과: C-S-H 형성 및 CaO 활성도 저하")
        status = "Critical"
    if solid_al_wt >= 0.50:
        diagnostics.append(f"Al 농도({solid_al_wt} wt%) 상승: 소성 시 비활성 클링커(Ca3Al2O6) 형성 위험")
        if status != "Critical": status = "Warning"
    if solid_mg_wt >= 0.60 or filtration_time_min >= 6.0:
        diagnostics.append(f"여과시간({filtration_time_min}분) 및 Mg 농도({solid_mg_wt} wt%) 상승: 여과포 눈막힘 및 케이크 슬라임화")
        status = "Critical"

    return {
        "process_status": status,
        "diagnostics": diagnostics if diagnostics else ["모든 불순물 수치 및 여과 속도가 정상 범위 내에 있습니다."]
    }

def calculate_optimal_purge_and_makeup(target_feed_li2co3_g, current_recycled_cao_g, solid_si_wt, **kwargs):
    req_total_cao = (target_feed_li2co3_g / 73.89) * 56.08 * 1.05

    if solid_si_wt < 0.8:
        purge_ratio = 0.05
    elif solid_si_wt < 1.2:
        purge_ratio = 0.15
    else:
        purge_ratio = 0.30

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
            "description": "탄산리튬 가성화 반응의 리튬 회수율, 반응 전환율 및 CaCO3 수율을 계산합니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "feed_li2co3_g": {"type": "number", "description": "투입된 탄산리튬 무게 (g)"},
                    "filtrate_li_g_l": {"type": "number", "description": "여액 평균 Li 농도 (g/L)"},
                    "filtrate_vol_l": {"type": "number", "description": "총 회수된 여액 부피 (L)"},
                    "recovered_dry_caco3_g": {"type": "number", "description": "회수된 건조 CaCO3 무게 (g)"}
                },
                "required": ["feed_li2co3_g", "filtrate_li_g_l", "filtrate_vol_l", "recovered_dry_caco3_g"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "diagnose_impurity_and_operability",
            "description": "재생 CaO 내 불순물(Si, Al, Mg) 축적 상태 및 여과 지연 현상을 진단합니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "solid_si_wt": {"type": "number", "description": "소성 재생 CaO 내 Si 함량 (wt%)"},
                    "solid_al_wt": {"type": "number", "description": "소성 재생 CaO 내 Al 함량 (wt%)"},
                    "solid_mg_wt": {"type": "number", "description": "소성 재생 CaO 내 Mg 함량 (wt%)"},
                    "filtration_time_min": {"type": "number", "description": "감압여과 소요시간 (분)"}
                },
                "required": ["solid_si_wt", "solid_al_wt", "solid_mg_wt", "filtration_time_min"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "calculate_optimal_purge_and_makeup",
            "description": "불순물 제어를 위한 최적 고형분 퍼지(Purge) 비율 및 신규 Fresh CaO 보충량을 산출합니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "target_feed_li2co3_g": {"type": "number", "description": "다음 회차 Li2CO3 투입 목표량 (g)"},
                    "current_recycled_cao_g": {"type": "number", "description": "소성 후 회수된 재생 CaO 양 (g)"},
                    "solid_si_wt": {"type": "number", "description": "현재 재생 CaO 내 Si 함량 (wt%)"}
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
# 3. 사이드바 UI
# =========================================================
with st.sidebar:
    st.header("⚙️ 에이전트 설정")
    default_key = ""
    try:
        default_key = st.secrets.get("OPENAI_API_KEY", "")
    except Exception:
        pass
        
    api_key = st.text_input("OpenAI API Key", value=default_key, type="password", help="sk-로 시작하는 실제 영문/숫자 API 키를 입력하세요")
    model_name = st.selectbox("LLM 모델", ["gpt-4o", "gpt-4o-mini"], index=0)
    
    st.markdown("---")
    st.header("🧪 신규 Cycle 데이터 등록")
    with st.form("add_cycle_form"):
        new_cycle = len(st.session_state.cycle_history) + 1
        st.subheader(f"Cycle {new_cycle} 데이터 입력")
        c_feed = st.number_input("Li2CO3 투입량 (g)", value=10.0, step=0.5)
        c_fresh = st.number_input("투입 Fresh CaO (g)", value=1.35, step=0.1)
        c_recy = st.number_input("투입 재생 CaO (g)", value=6.60, step=0.1)
        c_rec_pct = st.number_input("Li 회수율 (%)", value=88.5, step=0.5)
        c_conv_pct = st.number_input("전환율 (%)", value=86.0, step=0.5)
        c_si = st.number_input("고상 Si (wt%)", value=1.72, step=0.05)
        c_mg = st.number_input("고상 Mg (wt%)", value=0.98, step=0.05)
        c_al = st.number_input("고상 Al (wt%)", value=0.79, step=0.05)
        c_time = st.number_input("여과 시간 (min)", value=9.2, step=0.2)
        c_caco3 = st.number_input("건조 CaCO3 (g)", value=12.65, step=0.1)
        
        submitted = st.form_submit_button("Cycle 데이터 추가")
        if submitted:
            st.session_state.cycle_history.append({
                "Cycle": new_cycle,
                "Feed_Li2CO3_g": c_feed,
                "Fresh_CaO_g": c_fresh,
                "Recycle_CaO_g": c_recy,
                "Li_Recovery_pct": c_rec_pct,
                "Conversion_pct": c_conv_pct,
                "Solid_Si_wt": c_si,
                "Solid_Mg_wt": c_mg,
                "Solid_Al_wt": c_al,
                "Filtration_Time_min": c_time,
                "Dry_CaCO3_g": c_caco3
            })
            st.success(f"Cycle {new_cycle} 데이터가 추가되었습니다!")
            st.rerun()

# =========================================================
# 4. 상단 대시보드 지표 카드
# =========================================================
st.title("⚗️ Li₂CO₃ 가성화 & CaO 칼슘 루핑 AI 공정 에이전트")
st.markdown("자연어로 대화하며 **물질수지 자동 계산, 불순물 농축 진단, 최적 퍼지/보충량 처방**을 수행합니다.")

df = pd.DataFrame(st.session_state.cycle_history)
last_row = df.iloc[-1]

col1, col2, col3, col4 = st.columns(4)
col1.metric("현재 진행 회차", f"Cycle {int(last_row['Cycle'])}")
col1.metric("최근 Li 회수율", f"{last_row['Li_Recovery_pct']}%", delta=f"{round(last_row['Li_Recovery_pct'] - df.iloc[-2]['Li_Recovery_pct'], 1)}%" if len(df) > 1 else None)
col2.metric("고상 Si 축적 농도", f"{last_row['Solid_Si_wt']} wt%", delta=f"{round(last_row['Solid_Si_wt'] - df.iloc[-2]['Solid_Si_wt'], 2)} wt%" if len(df) > 1 else None, delta_color="inverse")
col3.metric("총 여과 소요시간", f"{last_row['Filtration_Time_min']} min", delta=f"{round(last_row['Filtration_Time_min'] - df.iloc[-2]['Filtration_Time_min'], 1)} min" if len(df) > 1 else None, delta_color="inverse")
col4.metric("재생 CaO 사용률", f"{round(last_row['Recycle_CaO_g'] / (last_row['Fresh_CaO_g'] + last_row['Recycle_CaO_g']) * 100, 1)}%")

# =========================================================
# 5. 메인 탭 구성
# =========================================================
tab_chat, tab_charts, tab_data = st.tabs(["💬 AI 공정 에이전트 대화", "📈 불순물 및 공정 시각화", "📋 회차별 실측 데이터"])

with tab_chat:
    st.subheader("🤖 공정 진단 및 처방 대화창")
    
    st.markdown("**💡 추천 빠른 질문:**")
    qc1, qc2, qc3 = st.columns(3)
    quick_input = None
    if qc1.button("📌 최근 Cycle 종합 진단"):
        quick_input = f"현재 Cycle {int(last_row['Cycle'])}까지 진행됐어. 고상 Si는 {last_row['Solid_Si_wt']}%, Mg는 {last_row['Solid_Mg_wt']}%, Al은 {last_row['Solid_Al_wt']}%, 여과시간은 {last_row['Filtration_Time_min']}분이야. 공정 이상 유무를 진단해줘."
    if qc2.button("⚠️ Si 농축 원인 및 여과 지연 대책"):
        quick_input = "Si와 Mg 농축이 여과 속도와 전환율에 미치는 영향을 분석하고 공정 해결 방안을 알려줘."
    if qc3.button("🎯 다음 회차 Purge 및 CaO 처방"):
        quick_input = f"다음 회차에 Li2CO3 10g을 처리할 예정이야. 현재 재생 CaO가 {last_row['Recycle_CaO_g']}g이고 Si가 {last_row['Solid_Si_wt']}%인데 최적 퍼지율과 Fresh CaO 보충량을 계산해줘."

    # 이전 대화 출력
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    user_query = st.chat_input("질문이나 분석할 실험 데이터를 입력하세요...")
    if quick_input:
        user_query = quick_input

    if user_query:
        # API Key 유효성 사전 검사 (한글/공백/따옴표 방어)
        clean_key = api_key.strip().replace('"', '').replace("'", "")
        
        if not clean_key:
            st.error("⚠️ 좌측 사이드바에서 OpenAI API Key를 입력해주세요!")
        elif not clean_key.startswith("sk-") or not clean_key.isascii():
            st.error("⚠️ 입력된 API Key가 올바르지 않습니다. 한글이나 따옴표가 없는 순수 영문/숫자 형태의 'sk-...' 키를 입력해주세요.")
        else:
            st.session_state.messages.append({"role": "user", "content": user_query})
            with st.chat_message("user"):
                st.markdown(user_query)

            with st.chat_message("assistant"):
                with st.spinner("물질수지 및 불순물 거동 분석 중..."):
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
        fig_imp = px.line(
            df, x="Cycle", y=["Solid_Si_wt", "Solid_Mg_wt", "Solid_Al_wt"],
            markers=True, title="재생 CaO 내 불순물 축적 추이 (wt%)",
            labels={"value": "농도 (wt%)", "variable": "원소 구분"}
        )
        fig_imp.add_hline(y=1.20, line_dash="dash", line_color="red", annotation_text="Si 임계치 (1.20 wt%)")
        st.plotly_chart(fig_imp, use_container_width=True)

        fig_time = px.bar(
            df, x="Cycle", y="Filtration_Time_min",
            title="회차별 총 여과 소요시간 (분)",
            color="Filtration_Time_min", color_continuous_scale="Reds"
        )
        fig_time.add_hline(y=6.0, line_dash="dash", line_color="orange", annotation_text="여과 지연 기준 (6.0 min)")
        st.plotly_chart(fig_time, use_container_width=True)

    with chart_col2:
        fig_rec = px.line(
            df, x="Cycle", y=["Li_Recovery_pct", "Conversion_pct"],
            markers=True, title="회차별 Li 회수율 및 반응 전환율 (%)",
            labels={"value": "비율 (%)", "variable": "지표"}
        )
        fig_rec.add_hline(y=90.0, line_dash="dash", line_color="darkred", annotation_text="수율 한계 (90%)")
        st.plotly_chart(fig_rec, use_container_width=True)

        fig_cao = px.bar(
            df, x="Cycle", y=["Recycle_CaO_g", "Fresh_CaO_g"],
            title="회차별 투입 CaO 구성 (재생 vs 신규)",
            labels={"value": "투입량 (g)", "variable": "구분"},
            barmode="stack"
        )
        st.plotly_chart(fig_cao, use_container_width=True)

with tab_data:
    st.subheader("📋 전체 실험 회차 데이터")
    st.dataframe(df.style.format({
        "Feed_Li2CO3_g": "{:.2f}",
        "Fresh_CaO_g": "{:.2f}",
        "Recycle_CaO_g": "{:.2f}",
        "Li_Recovery_pct": "{:.1f}%",
        "Conversion_pct": "{:.1f}%",
        "Solid_Si_wt": "{:.2f}",
        "Solid_Mg_wt": "{:.2f}",
        "Solid_Al_wt": "{:.2f}",
        "Filtration_Time_min": "{:.1f}",
        "Dry_CaCO3_g": "{:.2f}"
    }), use_container_width=True)

    csv_data = df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="📥 CSV 데이터 다운로드",
        data=csv_data,
        file_name="calcium_looping_experiment_data.csv",
        mime="text/csv"
    )
