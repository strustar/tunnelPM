import streamlit as st
import pandas as pd
import numpy as np

def check_shear(In, R):
    """
    🛡️ 전단설계 최적화 보고서 - KDS 14 20 기준 (v3.2 - 가독성 및 근거 강화)
    - 요청사항 반영: 요약표 가독성, 단계별 구분, 판정/계산 근거 명시
    - Av, Vs, Av/s 산정 근거 추가
    - 최종 판정 로직 강화 (OK/NG 및 사유 명시)
    - 최대 전단강도(단면) 검토 과정 분리 및 계산 근거 제시
    - 수식 스타일 통일 (LaTeX) 및 UI/UX 개선
    """

    # =================================================================
    # 0. 페이지 헤더 및 기본 UI 설정 (기존과 동일)
    # =================================================================
    st.markdown("""
    <style>
        /* 전체 폰트 및 기본 스타일 재정의 */
        .main .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
        }
        h1, h2, h3, h4, h5, h6 {
            font-weight: 700;
        }
        /* 계산 결과 블록 스타일 */
        .calc-block {
            background-color: #f0f2f6;
            padding: 20px;
            border-radius: 10px;
            border-left: 6px solid;
            margin-top: 10px;
            margin-bottom: 30px;
        }
        .calc-block p {
            font-size: 1.2em; /* 글자 크기 상향 */
            margin: 0;
            color: #333;
            font-weight: 500;
        }
        .calc-block strong {
            color: #0056b3;
            font-size: 1.3em; /* 결과값 강조 */
            font-weight: 700;
        }
        /* LaTeX 수식 스타일 */
        .stLatex {
            font-size: 1.15em; /* 수식 크기 상향 */
            padding: 15px;
            background-color: #e9ecef;
            border-radius: 5px;
            text-align: center;
            margin-bottom: 15px;
        }
    </style>
    <div style="background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
                padding: 50px 40px; border-radius: 20px; margin-bottom: 40px;
                box-shadow: 0 12px 30px rgba(0,0,0,0.15);">
        <h1 style="color: white; text-align: center; margin: 0; font-size: 3.5em;
                   font-weight: 900; text-shadow: 2px 2px 5px rgba(0,0,0,0.2);">
            🛡️ 전단설계 최적화 보고서
        </h1>
        <p style="color: #e0e0e0; text-align: center; margin: 15px 0 0 0;
                  font-size: 1.4em; opacity: 0.95;">
            KDS 14 20 콘크리트구조설계기준 적용
        </p>
    </div>
    """, unsafe_allow_html=True)

    # =================================================================
    # 1. 설계 기준 및 이론 (기존과 동일)
    # =================================================================
    st.markdown("## 📋 **전단철근 판정 기준 선택**")
    check_type = st.radio(
        "판정 기준을 선택하십시오:",
        ('일반 (2단계)', '프리캐스트 (3단계)'),
        horizontal=True,
        label_visibility="collapsed"
    )
    In.check_type = check_type
    st.markdown("---")

    if check_type == '프리캐스트 (3단계)':
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("""
            <div style="background: linear-gradient(135deg, #2196f3, #1976d2); color: white; padding: 25px; border-radius: 15px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); margin-bottom: 20px; height: 100%;">
                <h3 style="margin: 0; text-align: center; font-size: 1.5em;">🔵 전단철근 불필요</h3>
                <div style="text-align: center; margin: 20px 0; font-size: 1.3em; font-weight: bold;">V<sub>u</sub> ≤ ½φV<sub>c</sub></div>
                <p style="margin: 0; text-align: center; font-size: 1.1em; opacity: 0.9;">이론적으로 전단철근 불필요</p>
            </div>""", unsafe_allow_html=True)
        with col2:
            st.markdown("""
            <div style="background: linear-gradient(135deg, #5f6dbc, #48627e); color: white; padding: 25px; border-radius: 15px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); margin-bottom: 20px; height: 100%;">
                <h3 style="margin: 0; text-align: center; font-size: 1.5em;">🟡 최소전단철근</h3>
                <div style="text-align: center; margin: 20px 0; font-size: 1.2em; font-weight: bold;">½φV<sub>c</sub> &lt; V<sub>u</sub> ≤ φV<sub>c</sub></div>
                <p style="margin: 0; text-align: center; font-size: 1.1em; opacity: 0.9;">규정 최소량 적용</p>
            </div>""", unsafe_allow_html=True)
        with col3:
            st.markdown("""
            <div style="background: linear-gradient(135deg, #f44336, #d32f2f); color: white; padding: 25px; border-radius: 15px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); margin-bottom: 20px; height: 100%;">
                <h3 style="margin: 0; text-align: center; font-size: 1.5em;">🔴 설계전단철근</h3>
                <div style="text-align: center; margin: 20px 0; font-size: 1.3em; font-weight: bold;">V<sub>u</sub> > φV<sub>c</sub></div>
                <p style="margin: 0; text-align: center; font-size: 1.1em; opacity: 0.9;">계산에 의한 철근량</p>
            </div>""", unsafe_allow_html=True)
    else: # 일반 (2단계)
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("""
            <div style="background: linear-gradient(135deg, #2196f3, #1976d2); color: white; padding: 25px; border-radius: 15px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); margin-bottom: 20px; height: 100%;">
                <h3 style="margin: 0; text-align: center; font-size: 1.5em;">🔵 전단철근 불필요</h3>
                <div style="text-align: center; margin: 20px 0; font-size: 1.3em; font-weight: bold;">V<sub>u</sub> ≤ φV<sub>c</sub></div>
                <p style="margin: 0; text-align: center; font-size: 1.1em; opacity: 0.9;">규정에 의한 최소철근 배근 또는 불필요</p>
            </div>""", unsafe_allow_html=True)
        with col2:
            st.markdown("""
            <div style="background: linear-gradient(135deg, #f44336, #d32f2f); color: white; padding: 25px; border-radius: 15px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); margin-bottom: 20px; height: 100%;">
                <h3 style="margin: 0; text-align: center; font-size: 1.5em;">🔴 설계전단철근</h3>
                <div style="text-align: center; margin: 20px 0; font-size: 1.3em; font-weight: bold;">V<sub>u</sub> > φV<sub>c</sub></div>
                <p style="margin: 0; text-align: center; font-size: 1.1em; opacity: 0.9;">계산에 의한 철근량</p>
            </div>""", unsafe_allow_html=True)
            
    st.markdown("<br>", unsafe_allow_html=True)

    # =================================================================
    # 2. 공통 설계 조건 및 계산 로직 (기존과 동일)
    # =================================================================
    def format_number(num, decimal_places=1):
        return f"{num:,.{decimal_places}f}"

    phi_v = 0.75
    lamda = 1.0
    fy_shear = 400
    bar_dia = 13
    legs = 2
    bar_area = np.pi * (bar_dia / 2)**2
    Av_stirrup = bar_area * legs

    bw, d, fck, Ag = In.be, In.depth, In.fck, R.Ag
    
    results = [] # 결과를 저장할 리스트

    # =================================================================
    # 3. 각 하중 케이스별 계산 (기존과 동일)
    # =================================================================
    for i in range(len(In.Vu)):
        Vu = In.Vu[i]
        Pu = In.Pu_shear[i]

        p_factor = 1 + (Pu * 1000) / (14 * Ag) if Pu != 0 else 1.0
        Vc = (1/6) * p_factor * lamda * np.sqrt(fck) * bw * d
        phi_Vc = phi_v * Vc
        half_phi_Vc = 0.5 * phi_Vc

        if check_type == '프리캐스트 (3단계)':
            if Vu * 1000 <= half_phi_Vc:
                shear_category = "전단철근 불필요"
                category_color_hex = "#1976d2"
            elif Vu * 1000 <= phi_Vc:
                shear_category = "최소전단철근"
                category_color_hex = "#48627e"
            else:
                shear_category = "설계전단철근"
                category_color_hex = "#d32f2f"
        else: # 일반 (2단계)
            if Vu * 1000 <= phi_Vc:
                shear_category = "전단철근 불필요"
                category_color_hex = "#1976d2"
            else:
                shear_category = "설계전단철근"
                category_color_hex = "#d32f2f"

        min_Av_s_1_val = 0.0625 * np.sqrt(fck)
        min_Av_s_2_val = 0.35
        min_Av_s_1 = min_Av_s_1_val * (bw / fy_shear)
        min_Av_s_2 = min_Av_s_2_val * (bw / fy_shear)

        min_Av_s_req = max(min_Av_s_1, min_Av_s_2)
        s_from_min_req = Av_stirrup / min_Av_s_req

        Vs_req = (Vu * 1000 - phi_Vc) / phi_v if shear_category == "설계전단철근" else 0
        s_from_vs_req = (Av_stirrup * fy_shear * d) / Vs_req if Vs_req > 0 else float('inf')
        
        Vs_limit_d4 = (1/3) * np.sqrt(fck) * bw * d
        s_max_code = min(d / 4, 300) if Vs_req > Vs_limit_d4 else min(d / 2, 600)
        
        if Vs_req > Vs_limit_d4:
            s_max_reason = f"""
            <div style='background-color:#fff3cd; padding:15px; border-radius:8px; border-left:5px solid #ffc107; font-size: 18px'>
            <b>[판단근거]</b><br>
            <b>🔍 기준 구분:</b><br>
            - V<sub>s</sub> > (1/3)√f<sub>ck</sub>·b<sub>w</sub>·d 인 경우 → 더 촘촘한 간격 (d/4, 300mm)<br>
            - V<sub>s</sub> ≤ (1/3)√f<sub>ck</sub>·b<sub>w</sub>·d 인 경우 → 일반 간격 (d/2, 600mm)<br><br>
            <b>📊 현재 상황:</b><br>
            V<sub>s,req</sub> = {format_number(Vs_req/1000, 1)} kN <b>> 기준값 {format_number(Vs_limit_d4/1000, 1)} kN</b><br>
            → <b>기준값 초과</b>이므로 <b>더 촘촘한 최대 간격 (d/4, 300mm) 적용</b>
            </div>
            """
        else:
            s_max_reason = f"""
            <div style='background-color:#fff3cd; padding:15px; border-radius:8px; border-left:5px solid #ffc107; font-size: 18px'>
            <b>[판단근거]</b><br>
            <b>🔍 기준 구분:</b><br>
            - V<sub>s</sub> > (1/3)√f<sub>ck</sub>·b<sub>w</sub>·d 인 경우 → 더 촘촘한 간격 (d/4, 300mm)<br>
            - V<sub>s</sub> ≤ (1/3)√f<sub>ck</sub>·b<sub>w</sub>·d 인 경우 → 일반 간격 (d/2, 600mm)<br><br>
            <b>📊 현재 상황:</b><br>
            V<sub>s,req</sub> = {format_number(Vs_req/1000, 1)} kN <b>≤ 기준값 {format_number(Vs_limit_d4/1000, 1)} kN</b><br>
            → <b>기준값 이하</b>이므로 <b>일반 최대 간격 (d/2, 600mm) 적용</b>
            </div>
            """


        if shear_category == "전단철근 불필요":
            actual_s = s_max_code
            stirrups_needed = "전단철근 불필요"
        elif shear_category == "최소전단철근":
            s_calc = s_from_min_req
            actual_s = min(s_calc, s_max_code)
            actual_s = np.floor(actual_s / 5) * 5
            stirrups_needed = f"H{bar_dia}-{legs}leg @{actual_s:.0f}"
        else: # 설계전단철근
            s_calc = min(s_from_min_req, s_from_vs_req)
            actual_s = min(s_calc, s_max_code)
            actual_s = np.floor(actual_s / 5) * 5
            stirrups_needed = f"H{bar_dia}-{legs}leg @{actual_s:.0f}"

        if shear_category == "전단철근 불필요":
            phi_Vs = 0
        else:
            phi_Vs = (phi_v * Av_stirrup * fy_shear * d) / actual_s if actual_s > 0 else 0
        
        phi_Vn = phi_Vc + phi_Vs
        is_safe_strength = (phi_Vn >= Vu * 1000)
        Vs_max_limit = (2/3) * np.sqrt(fck) * bw * d
        Vs_provided = phi_Vs / phi_v if phi_Vs > 0 else 0
        is_safe_section = (Vs_provided <= Vs_max_limit)
        is_safe_total = is_safe_strength and is_safe_section
        
        stirrups_per_meter = 1000 / actual_s if actual_s > 0 and shear_category != "전단철근 불필요" else 0
        
        final_status_text = ""
        ng_reason = ""
        if not is_safe_section:
            final_status_text = "❌ NG (단면 부족)"
            ng_reason = f"전단철근이 부담하는 강도(Vs = {format_number(Vs_provided/1000, 1)} kN)가 최대 허용치(Vs,max = {format_number(Vs_max_limit/1000, 1)} kN)를 초과하여 단면 파괴가 우려됩니다. 단면 크기 상향이 필요합니다."
        elif not is_safe_strength:
            final_status_text = "❌ NG (강도 부족)"
            ng_reason = f"설계 전단강도(φVn = {format_number(phi_Vn/1000, 1)} kN)가 요구 전단강도(Vu = {format_number(Vu, 1)} kN)보다 작아 안전하지 않습니다."
        else:
            final_status_text = "✅ OK"

        results.append({
            'case': i + 1, 'Pu': Pu, 'Vu': Vu, 'shear_category': shear_category,
            'category_color': category_color_hex, 'phi_Vn_kN': phi_Vn / 1000, 
            'is_safe': is_safe_total, 'is_safe_section': is_safe_section, 'actual_s': actual_s,
            'stirrups_needed': stirrups_needed, 'stirrups_per_meter': stirrups_per_meter,
            'p_factor':p_factor, 'Vc_N':Vc, 'phi_Vc_N':phi_Vc, 'half_phi_Vc_N':half_phi_Vc,
            'Vs_req_N':Vs_req, 'min_Av_s_req':min_Av_s_req, 's_from_min_req':s_from_min_req, 
            's_from_vs_req': s_from_vs_req, 's_max_code':s_max_code, 's_max_reason': s_max_reason,
            'Vs_limit_d4_N':Vs_limit_d4, 'phi_Vs_N':phi_Vs, 
            'Vs_provided_N':Vs_provided, 'Vs_max_limit_N':Vs_max_limit,
            'final_status_text': final_status_text, 'ng_reason': ng_reason,
            'min_Av_s_1_val': min_Av_s_1_val, 'min_Av_s_2_val': min_Av_s_2_val
        })

    # =================================================================
    # 4. 전체 설계 결과 요약 (⭐ 스타일 수정)
    # =================================================================
    st.markdown("## 📊 **전체 설계 결과 요약**")
    st.markdown("---")

    summary_data = []
    for r in results:
        summary_data.append({
            'Case': f"Case {r['case']}",
            '하중조건 (kN)': f"Pu = {format_number(r['Pu'], 0)}\nVu = {format_number(r['Vu'])}",
            '판정결과': r['shear_category'],
            '최적 설계': f"{r['stirrups_needed']}",
            '1m당 개수': f"{r['stirrups_per_meter']:.1f}개" if r['stirrups_per_meter'] > 0 else "—",
            '설계강도 (kN)': f"φVn = {format_number(r['phi_Vn_kN'])}",
            '최종 판정': r['final_status_text']
        })
    
    df_summary = pd.DataFrame(summary_data)

    def style_safety(val):
        color = '#155724' if "✅ OK" in str(val) else '#721c24'
        bg_color = '#d4edda' if "✅ OK" in str(val) else '#f8d7da'
        border = '2px solid #c62828' if "❌ NG" in str(val) else 'none'
        return f'background-color: {bg_color}; color: {color}; font-weight: bold; text-align: center; border: {border};'

    def style_category(val):
        color_map = {"불필요": "#1565c0", "최소전단철근": "#2b385f", "설계전단철근": "#c62828"}
        bg_color_map = {"불필요": "#e3f2fd", "최소전단철근": "#e9ebee", "설계전단철근": "#ffebee"}
        for key in color_map:
            if key in str(val):
                return f'background-color: {bg_color_map[key]}; color: {color_map[key]}; font-weight: bold; text-align: center;'
        return 'text-align: center;'

    # DataFrame Styler 객체를 HTML로 변환하여 markdown으로 출력
    styled_df_html = (df_summary.style
        .applymap(style_safety, subset=['최종 판정'])
        .applymap(style_category, subset=['판정결과'])
        .set_properties(**{
            'text-align': 'center',
            'white-space': 'pre-line',
            'padding': '15px',
            'font-size': '20px',
            'font-weight': '900',
            'border': '2px solid #000000'  # 셀 테두리
        })
        .set_table_styles([
            # 표 전체 스타일
            {'selector': '', 'props': [
                ('width', '100%'),           # 화면 너비에 맞춤
                ('border-collapse', 'collapse'),
                ('margin', '0 auto'),        # 가운데 정렬
                ('box-shadow', '0 2px 10px rgba(0,0,0,0.1)')  # 그림자 효과
            ]},
            # 헤더 스타일
            {'selector': 'th', 'props': [
                ('background-color', '#39c561'),
                ('color', 'black'),
                ('font-weight', 'bold'),
                ('text-align', 'center'),
                ('padding', '18px'),
                ('font-size', '22px'),
                ('border', '2px solid #000000'),  # 헤더 테두리
            ]},
            # 데이터 셀 호버 효과
            {'selector': 'tr:hover', 'props': [
                ('background-color', '#f1f1f1'),
                ('transform', 'scale(1.01)'),     # 살짝 확대
                ('transition', 'all 0.2s ease')  # 부드러운 전환
            ]},
            # 홀수/짝수 행 구분
            {'selector': 'tr:nth-child(even)', 'props': [
                ('background-color', '#fafafa')
            ]},
            # 외곽 테두리 강화
            {'selector': 'table', 'props': [
                ('border', '3px solid #39c561'),
                ('border-radius', '8px'),         # 모서리 둥글게
                ('overflow', 'hidden')            # 둥근 모서리 적용
            ]}
        ])
        .hide(axis="index")
        .to_html())

    st.markdown(styled_df_html, unsafe_allow_html=True)

    # =================================================================
    # 5. 케이스별 상세 계산 과정 (⭐ 스타일 및 내용 수정)
    # =================================================================
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("## ⚙️ **케이스별 상세 계산 과정**")
    
    def format_N_to_kN(value, dp=2):
        return f"{value/1000:,.{dp}f}"

    for i, r in enumerate(results):
        num_symbols = ["❶", "❷", "❸", "❹", "❺", "❻", "❼", "❽", "❾", "❿"]
        st.markdown("---")
        st.markdown(f"### **{num_symbols[i]} Case {r['case']} 검토** (Pu = {format_number(r['Pu'], 0)} kN, Vu = {format_number(r['Vu'])} kN)")
        st.markdown(f"<h4 style='color: {r['category_color']}; margin-bottom: 25px;'>결과: {r['shear_category']} / {r['stirrups_needed']}</h4>", unsafe_allow_html=True)

        # 단계별 제목 스타일 변경
        def step_header(text):
            st.markdown(f"#### **{text}**")

        step_header("1단계: 축력 영향 계수 ($P_{증가}$)")
        st.latex(r"P_{증가} = 1 + \frac{P_u}{14 \cdot A_g}")
        st.markdown(f"""
        <div class="calc-block" style="border-color: #007bff;">
            <p>P<sub>증가</sub> = 1 + {format_number(r['Pu']*1000, 0)} &divide; (14 &times; {format_number(Ag, 0)}) = <strong>{r['p_factor']:.3f}</strong></p>
        </div>
        """, unsafe_allow_html=True)

        step_header("2단계: 콘크리트 설계 전단강도 ($\phi V_c$)")
        st.latex(r"\phi V_c = \phi_v \times \left( \frac{1}{6} P_{증가} \lambda \sqrt{f_{ck}} b_w d \right)")
        st.markdown(f"""
        <div class="calc-block" style="border-color: #17a2b8;">
            <p>φV<sub>c</sub> = {phi_v} &times; (1/6 &times; {r['p_factor']:.3f} &times; {lamda} &times; &radic;{fck} &times; {format_number(bw, 0)} &times; {format_number(d, 0)}) 
            = <strong>{format_N_to_kN(r['phi_Vc_N'])} kN</strong></p>
        </div>
        """, unsafe_allow_html=True)
        
        step_header("3단계: 전단철근 필요성 판정")
        if check_type == '프리캐스트 (3단계)':
            st.latex(r"V_u \text{ vs } \phi V_c, \quad \frac{1}{2}\phi V_c")
            
            judgement_str = ""
            if r['shear_category'] == "전단철근 불필요":
                judgement_str = f"V<sub>u</sub> = {format_number(r['Vu'])} kN &nbsp; ≤ &nbsp; ½φV<sub>c</sub> = {format_N_to_kN(r['half_phi_Vc_N'])} kN"
            elif r['shear_category'] == "최소전단철근":
                judgement_str = f"½φV<sub>c</sub> = {format_N_to_kN(r['half_phi_Vc_N'])} kN &nbsp; &lt; &nbsp; V<sub>u</sub> = {format_number(r['Vu'])} kN &nbsp; ≤ &nbsp; φV<sub>c</sub> = {format_N_to_kN(r['phi_Vc_N'])} kN"
            else:
                judgement_str = f"V<sub>u</sub> = {format_number(r['Vu'])} kN &nbsp; > &nbsp; φV<sub>c</sub> = {format_N_to_kN(r['phi_Vc_N'])} kN"
            
            st.markdown(f"""
            <div class="calc-block" style="border-color: {r['category_color']};">
                <p>{judgement_str}</p>
                <hr style='margin: 10px 0;'>
                <p style='font-size: 1.25em;'>판정: <strong style='color:{r['category_color']};'>{r['shear_category']}</strong></p>
            </div>
            """, unsafe_allow_html=True)
        else: # 일반 (2단계)
            st.latex(r"V_u \text{ vs } \phi V_c")
            judgement_str = ""
            if r['shear_category'] == "전단철근 불필요":
                judgement_str = f"V<sub>u</sub> = {format_number(r['Vu'])} kN &nbsp; ≤ &nbsp; φV<sub>c</sub> = {format_N_to_kN(r['phi_Vc_N'])} kN"
            else:
                judgement_str = f"V<sub>u</sub> = {format_number(r['Vu'])} kN &nbsp; > &nbsp; φV<sub>c</sub> = {format_N_to_kN(r['phi_Vc_N'])} kN"

            st.markdown(f"""
            <div class="calc-block" style="border-color: {r['category_color']};">
                <p>{judgement_str}</p>
                <hr style='margin: 10px 0;'>
                <p style='font-size: 1.25em;'>판정: <strong style='color:{r['category_color']};'>{r['shear_category']}</strong></p>
            </div>
            """, unsafe_allow_html=True)

        if r['shear_category'] != "전단철근 불필요":
            step_header("4단계: 필요 전단철근량 및 간격 계산")

            st.markdown("<h5><b>■ 전단철근 단면적 (A<sub>v</sub>) 산정</b></h5>", unsafe_allow_html=True)
            st.latex(r"A_v = n \times \frac{\pi D^2}{4}")
            st.markdown(f"""
            <div class="calc-block" style="border-color: #6f42c1;">
                <p>A<sub>v</sub> = {legs} &times; (π &times; {bar_dia}<sup>2</sup> &divide; 4) = <strong>{Av_stirrup:.1f} mm²</strong></p>
            </div>
            """, unsafe_allow_html=True)

            if r['shear_category'] == "설계전단철근":
                st.markdown("<h5><b>■ 전단철근이 부담할 필요 전단강도 (V<sub>s</sub>) 산정</b></h5>", unsafe_allow_html=True)
                st.latex(r"V_s = \frac{V_u - \phi V_c}{\phi_v}")
                st.markdown(f"""
                <div class="calc-block" style="border-color: #6f42c1;">
                    <p>V<sub>s,req</sub> = ({format_number(r['Vu'])} - {format_N_to_kN(r['phi_Vc_N'])}) &divide; {phi_v} = <strong>{format_N_to_kN(r['Vs_req_N'])} kN</strong></p>
                </div>
                """, unsafe_allow_html=True)
                st.markdown("<h5><b>■ 강도 요구조건에 의한 간격 (s<sub>강도요구</sub>) 산정</b></h5>", unsafe_allow_html=True)
                st.latex(r"s_{강도요구} \leq \frac{A_v f_{yt} d}{V_s}")
                st.markdown(f"""
                <div class="calc-block" style="border-color: #6f42c1;">
                    <p>간격 (강도) = ({Av_stirrup:.1f} &times; {fy_shear} &times; {format_number(d,0)}) &divide; {format_number(r['Vs_req_N'], 0)} = <strong>{format_number(r['s_from_vs_req'])} mm</strong></p>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<h5><b>■ 최소 철근량 규정에 의한 간격 (s<sub>최소철근</sub>) 산정</b></h5>", unsafe_allow_html=True)
            st.latex(r"(\frac{A_v}{s})_{min} = \max\left(0.0625\sqrt{f_{ck}}\frac{b_w}{f_{yt}}, 0.35\frac{b_w}{f_{yt}}\right)")
            st.markdown(f"""
            <div class="calc-block" style="border-color: #6610f2;">
                <p>(A<sub>v</sub>/s)<sub>min</sub> = max( {r['min_Av_s_1_val']:.4f} &times; {bw}/{fy_shear}, {r['min_Av_s_2_val']} &times; {bw}/{fy_shear} )</p>
                <p>= max( {min_Av_s_1:.4f}, {min_Av_s_2:.4f} ) = <strong>{r['min_Av_s_req']:.4f}</strong></p>
            </div>
            """, unsafe_allow_html=True)
            st.latex(r"s_{최소철근} \leq \frac{A_v}{(\frac{A_v}{s})_{min}}")
            st.markdown(f"""
            <div class="calc-block" style="border-color: #6610f2;">
                <p>간격 (최소철근) = {Av_stirrup:.1f} &divide; {r['min_Av_s_req']:.4f} = <strong>{format_number(r['s_from_min_req'])} mm</strong></p>
            </div>
            """, unsafe_allow_html=True)

            step_header("5단계: 최대 허용 간격 결정")
            st.markdown(f"""            
                <p style="font-size:1.1em; color:#856404; margin:0;">{r['s_max_reason']}</p>            
            """, unsafe_allow_html=True)
            st.latex(r"s_{max} = \min(d/2, 600) \quad \text{or} \quad \min(d/4, 300)")
            st.markdown(f"""
            <div class="calc-block" style="border-color: #fd7e14;">
                <p>간격 (최대허용 기준) = <strong>{format_number(r['s_max_code'])} mm</strong></p>
            </div>
            """, unsafe_allow_html=True)
            
            step_header("6단계: 최종 전단철근 간격 결정")
            st.latex(r"s_{final} = \text{floor}\left( \min(s_{계산값}, s_{max}) \right)")
            
            s_final_calc_str = ""
            if r['shear_category'] == "설계전단철근":
                s_final_calc_str = f"min({format_number(r['s_from_vs_req'])}, {format_number(r['s_from_min_req'])}, {format_number(r['s_max_code'])})"
            elif r['shear_category'] == "최소전단철근":
                s_final_calc_str = f"min({format_number(r['s_from_min_req'])}, {format_number(r['s_max_code'])})"

            st.markdown(f"""
            <div class="calc-block" style="border-color: #20c997;">
                <p>s<sub>final</sub> = floor( {s_final_calc_str} ) = <strong>{r['actual_s']:.0f} mm</strong></p>
                <p style="font-size:0.95em; color:#6c757d; margin-top:8px;">* 계산된 간격은 시공성을 고려하여 5mm 단위로 내림하여 적용합니다.</p>
            </div>
            """, unsafe_allow_html=True)

        # 단계 번호 조정
        final_check_start_num = 7
        if r['shear_category'] == "전단철근 불필요":
            final_check_start_num = 4

        step_header(f"{final_check_start_num}단계: 최대 전단강도 검토 (단면 안전성)")
        st.latex(r"V_s \leq V_{s,max} = \frac{2}{3}\sqrt{f_{ck}}b_w d")
        
        section_check_color = "#28a745" if r['is_safe_section'] else "#dc3545"
        section_check_status = "OK" if r['is_safe_section'] else "NG"
        st.markdown(f"""
        <div class="calc-block" style="border-color: {section_check_color};">
            <p>V<sub>s,배근</sub> = φV<sub>s</sub> / φ<sub>v</sub> = {format_N_to_kN(r['phi_Vs_N'])} &divide; {phi_v} = <strong>{format_N_to_kN(r['Vs_provided_N'])} kN</strong></p>
            <p>V<sub>s,max</sub> = (2/3) &times; &radic;{fck} &times; {format_number(bw, 0)} &times; {format_number(d, 0)} = <strong>{format_N_to_kN(r['Vs_max_limit_N'])} kN</strong></p>
            <hr style='margin: 10px 0;'>
            <p style='font-size: 1.25em;'>V<sub>s,배근</sub> ≤ V<sub>s,max</sub> 판정: <strong style='color:{section_check_color};'>{section_check_status}</strong></p>
        </div>
        """, unsafe_allow_html=True)

        step_header(f"{final_check_start_num + 1}단계: 최종 안전성 검토")
        result_color = "#28a745" if r['is_safe'] else "#dc3545"
        result_icon = "✅" if r['is_safe'] else "❌"
        
        st.markdown(f"""
        <div style="background-color: {result_color}1A; padding: 25px; border-radius: 10px; 
                     border: 2px solid {result_color}; margin-bottom: 40px;">
            <p style="font-size: 1.5em; font-weight: 700; color: {result_color}; margin-bottom: 15px; text-align: center;">
            {result_icon} 최종 검토 결과
            </p>
            <p style="font-size: 1.2em; line-height: 1.8;">
                <b>최종 배근:</b> <span style="font-size: 1.3em; font-weight: 700;">{r['stirrups_needed']}</span>
                (1m당 <b>{r['stirrups_per_meter']:.1f}개</b>)<br>
                <b>설계 전단강도 (φVn):</b> {format_N_to_kN(r['phi_Vc_N'])} + {format_N_to_kN(r['phi_Vs_N'])} = 
                <span style="font-size: 1.3em; font-weight: 700;">{format_number(r['phi_Vn_kN'])} kN</span><br>
                <b>요구 강도 (Vu):</b> <span style="font-weight: 700;">{format_number(r['Vu'])} kN</span><br>
                <hr style="margin: 10px 0; border-color: {result_color}80;">
                <b>최종 판정:</b> <span style="font-size: 1.35em; font-weight: 700;">{r['final_status_text']}</span>
            </p>
            {f'''<div style="border-top: 2px solid {result_color}80; margin-top: 15px; padding-top: 15px;">
                    <p style="color: {result_color}; font-weight: bold; margin:0; font-size: 1.15em;">
                        ⚠️ <b>사유:</b> {r["ng_reason"]}
                    </p>
                </div>''' if r["ng_reason"] else ''}
        </div>
        """, unsafe_allow_html=True)
        
    # placeholder 업데이트 (기존 코드와 호환)
    if hasattr(In, 'placeholder_shear'):
        for i, r in enumerate(results):
            if i < len(In.placeholder_shear) and In.placeholder_shear[i]:
                if r['is_safe']:
                    In.placeholder_shear[i].markdown(
                        f":white_check_mark: **:green[OK]**"
                    )
                else:
                    In.placeholder_shear[i].markdown(
                        f":x: **:red[NG]**"
                    )

    return results