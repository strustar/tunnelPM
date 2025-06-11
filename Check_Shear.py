import streamlit as st
import pandas as pd
import numpy as np

def check_shear(In, R):
    """
    🛡️ 전단설계 최적화 보고서 - KDS 14 20 기준 (스타일 최적화 버전)
    - Expander 제거 및 상세과정 기본 표시
    - 가독성을 위한 폰트 크기 및 레이아웃 조정
    - 계산식 및 결과값 스타일 강화
    """

    # =================================================================
    # 0. 페이지 헤더 및 기본 UI 설정
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
            font-size: 1.15em; /* 글자 크기 상향 */
            margin: 0;
            color: #333;
            font-weight: 500;
        }
        .calc-block strong {
            color: #0056b3;
            font-size: 1.25em; /* 결과값 강조 */
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
    # 1. 설계 기준 및 이론
    # =================================================================
    st.markdown("## 📋 **전단철근 판정 기준 (3단계)**")
    st.markdown("---")

    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #2196f3, #1976d2); 
                    color: white; padding: 25px; border-radius: 15px; 
                    box-shadow: 0 5px 15px rgba(0,0,0,0.1); margin-bottom: 20px; height: 100%;">
            <h3 style="margin: 0; text-align: center; font-size: 1.5em;">
                🔵 전단철근 불필요
            </h3>
            <div style="text-align: center; margin: 20px 0; font-size: 1.3em; font-weight: bold;">
                V<sub>u</sub> ≤ ½φV<sub>c</sub>
            </div>
            <p style="margin: 0; text-align: center; font-size: 1.1em; opacity: 0.9;">
                이론적으로 전단철근 불필요
            </p>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #ffc107, #f57c00); 
                    color: white; padding: 25px; border-radius: 15px; 
                    box-shadow: 0 5px 15px rgba(0,0,0,0.1); margin-bottom: 20px; height: 100%;">
            <h3 style="margin: 0; text-align: center; font-size: 1.5em;">
                🟡 최소전단철근
            </h3>
            <div style="text-align: center; margin: 20px 0; font-size: 1.2em; font-weight: bold;">
                ½φV<sub>c</sub> &lt; V<sub>u</sub> ≤ φV<sub>c</sub>
            </div>
            <p style="margin: 0; text-align: center; font-size: 1.1em; opacity: 0.9;">
                규정 최소량 적용
            </p>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #f44336, #d32f2f); 
                    color: white; padding: 25px; border-radius: 15px; 
                    box-shadow: 0 5px 15px rgba(0,0,0,0.1); margin-bottom: 20px; height: 100%;">
            <h3 style="margin: 0; text-align: center; font-size: 1.5em;">
                🔴 설계전단철근
            </h3>
            <div style="text-align: center; margin: 20px 0; font-size: 1.3em; font-weight: bold;">
                V<sub>u</sub> > φV<sub>c</sub>
            </div>
            <p style="margin: 0; text-align: center; font-size: 1.1em; opacity: 0.9;">
                계산에 의한 철근량
            </p>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("<br>", unsafe_allow_html=True)

    # =================================================================
    # 2. 공통 설계 조건 및 계산 로직
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
    # 3. 각 하중 케이스별 계산 (내부 로직)
    # =================================================================
    for i in range(len(In.Vu)):
        Vu = In.Vu[i]
        Pu = In.Pu_shear[i]

        p_factor = 1 + (Pu * 1000) / (14 * Ag) if Pu != 0 else 1.0
        Vc = (1/6) * p_factor * lamda * np.sqrt(fck) * bw * d
        phi_Vc = phi_v * Vc
        half_phi_Vc = 0.5 * phi_Vc

        if Vu * 1000 <= half_phi_Vc:
            shear_category = "전단철근 불필요"
            category_color_hex = "#1976d2"
        elif Vu * 1000 <= phi_Vc:
            shear_category = "최소전단철근"
            category_color_hex = "#f57c00"
        else:
            shear_category = "설계전단철근"
            category_color_hex = "#d32f2f"

        min_Av_s_1 = 0.0625 * np.sqrt(fck) * (bw / fy_shear)
        min_Av_s_2 = 0.35 * (bw / fy_shear)
        min_Av_s_req = max(min_Av_s_1, min_Av_s_2)
        s_from_min_req = Av_stirrup / min_Av_s_req

        Vs_req = (Vu * 1000 - phi_Vc) / phi_v if shear_category == "설계전단철근" else 0
        Vs_limit_d4 = (1/3) * np.sqrt(fck) * bw * d
        s_max_code = min(d / 4, 300) if Vs_req > Vs_limit_d4 else min(d / 2, 600)
        
        if shear_category == "전단철근 불필요":
            actual_s = s_max_code
            stirrups_needed = "전단철근 불필요"
        else:
            if shear_category == "설계전단철근":
                s_from_vs_req = (Av_stirrup * fy_shear * d) / Vs_req if Vs_req > 0 else float('inf')
                s_calc = min(s_from_min_req, s_from_vs_req)
            else: # 최소전단철근
                s_calc = s_from_min_req
            
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
        safety_ratio = phi_Vn / (Vu * 1000) if Vu > 0 else float('inf')
        stirrups_per_meter = 1000 / actual_s if actual_s > 0 and shear_category != "전단철근 불필요" else 0
        
        results.append({
            'case': i + 1, 'Pu': Pu, 'Vu': Vu, 'shear_category': shear_category,
            'category_color': category_color_hex, 'phi_Vn_kN': phi_Vn / 1000, 
            'safety_ratio': safety_ratio, 'is_safe': is_safe_total,
            'is_safe_section': is_safe_section, 'actual_s': actual_s,
            'stirrups_needed': stirrups_needed, 'stirrups_per_meter': stirrups_per_meter,
            'p_factor':p_factor, 'Vc_N':Vc, 'phi_Vc_N':phi_Vc, 'half_phi_Vc_N':half_phi_Vc,
            'Vs_req_N':Vs_req, 'min_Av_s_req':min_Av_s_req, 's_from_min_req':s_from_min_req, 
            'Vs_limit_d4_N':Vs_limit_d4, 's_max_code':s_max_code, 'phi_Vs_N':phi_Vs, 
            'Vs_provided_N':Vs_provided, 'Vs_max_limit_N':Vs_max_limit
        })

    # =================================================================
    # 4. 전체 설계 결과 요약
    # =================================================================
    st.markdown("## 📊 **전체 설계 결과 요약**")
    st.markdown("---")

    summary_data = []
    for r in results:
        status_icon = "✅ 안전" if r['is_safe'] else "❌ NG"
        if not r['is_safe_section']: status_icon += " (단면!)"

        summary_data.append({
            'Case': f"Case {r['case']}",
            '하중조건 (kN)': f"Pu = {format_number(r['Pu'], 0)}\nVu = {format_number(r['Vu'])}",
            '판정결과': r['shear_category'],
            '최적 설계': r['stirrups_needed'],
            '1m당 개수': f"{r['stirrups_per_meter']:.1f}개" if r['stirrups_per_meter'] > 0 else "—",
            '설계강도 (kN)': f"φVn = {format_number(r['phi_Vn_kN'])}",
            '안전율': f"{r['safety_ratio']:.3f}",
            '최종 판정': status_icon
        })
    
    df_summary = pd.DataFrame(summary_data)

    def style_safety(val):
        color = '#155724' if "✅ 안전" in str(val) else '#721c24'
        bg_color = '#d4edda' if "✅ 안전" in str(val) else '#f8d7da'
        border = '2px solid #dc3545' if "❌ NG" in str(val) else 'none'
        return f'background-color: {bg_color}; color: {color}; font-weight: bold; text-align: center; border: {border};'

    def style_category(val):
        color_map = {"불필요": "#1565c0", "최소전단철근": "#e65100", "설계전단철근": "#c62828"}
        bg_color_map = {"불필요": "#e3f2fd", "최소전단철근": "#fff8e1", "설계전단철근": "#ffebee"}
        for key in color_map:
            if key in str(val):
                return f'background-color: {bg_color_map[key]}; color: {color_map[key]}; font-weight: bold; text-align: center;'
        return 'text-align: center;'

    st.dataframe(df_summary.style
        .applymap(style_safety, subset=['최종 판정'])
        .applymap(style_category, subset=['판정결과'])
        .set_properties(**{'text-align': 'center', 'white-space': 'pre-line', 'padding': '12px', 'font-size': '1.05em'})
        .set_table_styles([
            {'selector': 'th', 'props': [('background-color', '#343a40'), ('color', 'white'), ('font-weight', 'bold'),
                                         ('text-align', 'center'), ('padding', '15px'), ('font-size', '1.1em')]},
            {'selector': 'tr:hover', 'props': [('background-color', '#f1f1f1')]}
        ]), use_container_width=True, hide_index=True)

    # =================================================================
    # 5. 케이스별 상세 계산 과정
    # =================================================================
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("## ⚙️ **케이스별 상세 계산 과정**")
    st.markdown("---")

    def format_N_to_kN(value, dp=2):
        return f"{value/1000:,.{dp}f}"

    for i, r in enumerate(results):
        num_symbols = ["❶", "❷", "❸", "❹", "❺", "❻", "❼", "❽", "❾", "❿"]
        st.markdown(f"### **{num_symbols[i]} Case {r['case']} 검토** (Vu = {format_number(r['Vu'])} kN)")
        st.markdown(f"<h4 style='color: {r['category_color']}; margin-bottom: 25px;'>결과: {r['shear_category']} / {r['stirrups_needed']}</h4>", unsafe_allow_html=True)

        # 단계별 계산 과정 표시
        st.markdown("##### **1단계: 축력 영향 계수 ($P_{증가}$)**")
        st.latex(r"P_{증가} = 1 + \frac{P_u}{14 \cdot A_g}")
        st.markdown(f"""
        <div class="calc-block" style="border-color: #007bff;">
            <p>P<sub>증가</sub> = 1 + {format_number(r['Pu']*1000, 0)} &divide; (14 &times; {format_number(Ag, 0)}) = <strong>{r['p_factor']:.3f}</strong></p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("##### **2단계: 콘크리트 설계 전단강도 ($\phi V_c$)**")
        st.latex(r"\phi V_c = \phi_v \times \left( \frac{1}{6} P_{증가} \lambda \sqrt{f_{ck}} b_w d \right)")
        st.markdown(f"""
        <div class="calc-block" style="border-color: #28a745;">
            <p>φV<sub>c</sub> = {phi_v} &times; (1/6 &times; {r['p_factor']:.3f} &times; {lamda} &times; &radic;{fck} &times; {format_number(bw, 0)} &times; {format_number(d, 0)}) 
            = <strong>{format_N_to_kN(r['phi_Vc_N'])} kN</strong></p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("##### **3단계: 전단철근 필요성 판정**")
        st.latex(r"V_u \text{ vs } \phi V_c, \quad \frac{1}{2}\phi V_c")
        st.markdown(f"""
        <div class="calc-block" style="border-color: {r['category_color']};">
            <p>V<sub>u</sub> = <strong>{format_number(r['Vu'])} kN</strong></p>
            <p>φV<sub>c</sub> = {format_N_to_kN(r['phi_Vc_N'])} kN</p>
            <p>½φV<sub>c</sub> = {format_N_to_kN(r['half_phi_Vc_N'])} kN</p>
            <hr style='margin: 10px 0;'>
            <p style='font-size: 1.2em;'>판정: <strong style='color:{r['category_color']};'>{r['shear_category']}</strong></p>
        </div>
        """, unsafe_allow_html=True)

        if r['shear_category'] != "전단철근 불필요":
            st.markdown("##### **4단계: 필요 전단철근량 및 간격 계산**")
            st.latex(r"s \leq \frac{A_v f_{yt} d}{V_s} \quad \text{and} \quad s \leq s_{최대허용}")
            
            # 최소철근량 간격
            st.latex(r"(\frac{A_v}{s})_{min} = \max\left(0.0625\sqrt{f_{ck}}\frac{b_w}{f_{yt}}, 0.35\frac{b_w}{f_{yt}}\right)")
            st.markdown(f"""
            <div class="calc-block" style="border-color: #6f42c1;">
                <p>간격 (최소철근량 기준) = {Av_stirrup:.1f} &divide; {r['min_Av_s_req']:.4f} = <strong>{format_number(r['s_from_min_req'])} mm</strong></p>
            </div>
            """, unsafe_allow_html=True)

            # 최대간격
            st.latex(r"s_{max} = \min(d/2, 600) \quad \text{or} \quad \min(d/4, 300)")
            st.markdown(f"""
            <div class="calc-block" style="border-color: #fd7e14;">
                <p>간격 (최대허용 기준) = <strong>{format_number(r['s_max_code'])} mm</strong></p>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("##### **5단계: 최종 배근 및 강도 검토**")
        st.latex(r"s_{final} \quad \rightarrow \quad \phi V_n = \phi V_c + \phi V_s \geq V_u")
        
        result_color = "#28a745" if r['is_safe'] else "#dc3545"
        result_icon = "✅" if r['is_safe'] else "❌"
        
        st.markdown(f"""
        <div style="background-color: {result_color}1A; padding: 25px; border-radius: 10px; 
                    border: 2px solid {result_color}; margin-bottom: 25px;">
            <p style="font-size: 1.3em; font-weight: 700; color: {result_color}; margin-bottom: 15px; text-align: center;">
            {result_icon} 최종 검토 결과
            </p>
            <p style="font-size: 1.1em; line-height: 1.8;">
                <b>최종 배근:</b> <span style="font-size: 1.2em; font-weight: 700;">{r['stirrups_needed']}</span>
                (1m당 <b>{r['stirrups_per_meter']:.1f}개</b>)<br>
                <b>설계 전단강도 (φVn):</b> {format_N_to_kN(r['phi_Vc_N'])} + {format_N_to_kN(r['phi_Vs_N'])} = 
                <span style="font-size: 1.2em; font-weight: 700;">{format_number(r['phi_Vn_kN'])} kN</span><br>
                <b>안전성:</b> φVn = {format_number(r['phi_Vn_kN'])} kN {'≥' if r['is_safe'] else '<'} Vu = {format_number(r['Vu'])} kN<br>
                <b>안전율 (S.F):</b> {format_number(r['phi_Vn_kN'])} &divide; {format_number(r['Vu'])} = 
                <span style="font-size: 1.2em; font-weight: 700;">{r['safety_ratio']:.3f}</span>
            </p>
        </div>
        """, unsafe_allow_html=True)

        if not r['is_safe_section']:
            st.error(f"⚠️ **단면 검토 경고:** 전단철근이 부담하는 강도(Vs = {format_N_to_kN(r['Vs_provided_N'])} kN)가 최대 허용치(Vs,max = {format_N_to_kN(r['Vs_max_limit_N'])} kN)를 초과했습니다. 단면 크기 상향이 필요합니다.")

    # placeholder 업데이트 (기존 코드와 호환)
    if hasattr(In, 'placeholder_shear'):
        for i, r in enumerate(results):
            if i < len(In.placeholder_shear) and In.placeholder_shear[i]:
                if r['is_safe']:
                    In.placeholder_shear[i].markdown(
                        f"<div style='text-align:center; color:#38a169; font-weight:900; "
                        f"font-size:18px; padding: 4px; border-radius: 12px; background-color: #f0fff4; "
                        f"border: 2px solid #38a169;'>"
                        f"✅ {r['safety_ratio']:.2f}</div>",
                        unsafe_allow_html=True
                    )
                else:
                    In.placeholder_shear[i].markdown(
                        f"<div style='text-align:center; color:#e53e3e; font-weight:900; "
                        f"font-size:18px; padding: 4px; border-radius: 12px; background-color: #fed7d7; "
                        f"border: 2px solid #e53e3e;'>"
                        f"❌ NG </div>",
                        unsafe_allow_html=True
                    )

    return results