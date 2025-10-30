import streamlit as st
import pandas as pd
import numpy as np


def check_shear(In, R):
    """
    🛡️ 전단설계 최적화 보고서 - KDS 14 20 기준 (v4.0 - Mm 개념 및 축력 고려 반영)
    - Mm (수정 모멘트) 개념 추가
    - Mm < 0: 축력 고려식 적용
    - Mm > 0: 정밀식 적용
    - 계산 근거 및 수식 명시 강화
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
            font-size: 1.2em;
            margin: 0;
            color: #333;
            font-weight: 500;
        }
        .calc-block strong {
            color: #0056b3;
            font-size: 1.3em;
            font-weight: 700;
        }
        /* LaTeX 수식 스타일 */
        .stLatex {
            font-size: 1.15em;
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
            KDS 14 20 콘크리트구조설계기준 적용 (축력 고려)
        </p>
    </div>
    """, unsafe_allow_html=True)

    # =================================================================
    # 1. 설계 기준 및 이론
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
    # 2. 공통 설계 조건 및 계산 로직
    # =================================================================
    def format_number(num, decimal_places=1):
        return f"{num:,.{decimal_places}f}"

    phi_v = 0.75
    lamda = 1.0
    bar_dia = 13
    legs = 2
    bar_area = np.pi * (bar_dia / 2)**2
    Av_stirrup = bar_area * legs

    bw, d, h, fck, Ag, fy_shear = In.be, In.depth, In.height, In.fck, R.Ag, In.fy_hollow
    Ast_tension = R.Ast_tension[0]/2
    Ast_compression = R.Ast_compression[0]/2
    As = Ast_tension + Ast_compression
    rho_w = As/(bw*d)
    
    results = [] # 결과를 저장할 리스트

    # =================================================================
    # 3. 각 하중 케이스별 계산 (Mm 개념 추가)
    # =================================================================
    for i in range(len(In.Vu)):
        Vu = In.Vu[i]
        Pu = In.Pu[i]  # kN (압축 +, 인장 -)
        Mu = In.Mu[i]  # kN·m
        
        Nu = Pu * 1000  # N 단위로 변환
        Mu_Nmm = Mu * 1e6  # N·mm 단위로 변환

        # ============================================================
        # 🔹 Mm (수정 모멘트) 계산
        # ============================================================
        Mm = Mu_Nmm - Nu * (4 * h - d) / 8
        
        # ============================================================
        # 🔹 Vc 계산 (Mm 값에 따라 식 선택)
        # ============================================================
        if Mm < 0:
            # Case 1: 축력이 지배적 (Mm < 0)
            # Vc = 0.29 λ √fck B d √(1 + Nu/(3.5·Ag))
            Vc = 0.29 * lamda * np.sqrt(fck) * bw * d * np.sqrt(1 + Nu / (3.5 * Ag))
            vc_method = "축력 고려식 (Mm < 0)"
            vc_formula = r"V_c = 0.29 \lambda \sqrt{f_{ck}} b_w d \sqrt{1 + \frac{N_u}{3.5 A_g}}"
        else:
            # Case 2: 휨모멘트가 지배적 (Mm ≥ 0) - 정밀식
            # # Vc = (0.16 √fck + 17.6 ρw Vu d / Mu) b d
            
            # Vu_N = Vu * 1000  # N 단위
            # term2 = 17.6 * rho_w * Vu_N * d / Mu_Nmm if Mu_Nmm != 0 else 0
            # term2 = min(term2, 1.0)  # Vu·d/Mu ≤ 1.0 제한
            
            # Vc = (0.16 * np.sqrt(fck) + term2) * bw * d 

            # 상한값 제한
            Vc = (1/6 * np.sqrt(fck) + 17.6 * rho_w * Vu*1000 * d / (Mm)) * bw * d
            # Vc_max = (1/3) * np.sqrt(fck) * bw * d
            # Vc = min(Vc, Vc_max)
            
            vc_method = "정밀식 (Mm ≥ 0)"
            vc_formula = r"V_c = \left(0.16\sqrt{f_{ck}} + 17.6\rho_w\frac{V_u d}{M_u}\right)b_w d \leq \frac{1}{3}\sqrt{f_{ck}}b_w d"


        phi_Vc = phi_v * Vc
        half_phi_Vc = 0.5 * phi_Vc

        # 판정 로직
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

        # 최소 전단철근량
        min_Av_s_1_val = 0.0625 * np.sqrt(fck)
        min_Av_s_2_val = 0.35
        min_Av_s_1 = min_Av_s_1_val * (bw / fy_shear)
        min_Av_s_2 = min_Av_s_2_val * (bw / fy_shear)

        min_Av_s_req = max(min_Av_s_1, min_Av_s_2)
        s_from_min_req = Av_stirrup / min_Av_s_req

        # 필요 전단철근량
        Vs_req = (Vu * 1000 - phi_Vc) / phi_v if shear_category == "설계전단철근" else 0
        s_from_vs_req = (Av_stirrup * fy_shear * d) / Vs_req if Vs_req > 0 else float('inf')
        
        # 최대 간격
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

        # 최종 간격 결정
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

        # 최종 강도 계산
        if shear_category == "전단철근 불필요":
            phi_Vs = 0
        else:
            phi_Vs = (phi_v * Av_stirrup * fy_shear * d) / actual_s if actual_s > 0 else 0
        
        phi_Vn = phi_Vc + phi_Vs
        is_safe_strength = (phi_Vn >= Vu * 1000)
        
        # 단면 안전성
        Vs_max_limit = (2/3) * np.sqrt(fck) * bw * d
        Vs_provided = phi_Vs / phi_v if phi_Vs > 0 else 0
        is_safe_section = (Vs_provided <= Vs_max_limit)
        is_safe_total = is_safe_strength and is_safe_section
        
        stirrups_per_meter = 1000 / actual_s if actual_s > 0 and shear_category != "전단철근 불필요" else 0
        
        # 최종 판정
        final_status_text = ""
        ng_reason = ""
        if not is_safe_section:
            final_status_text = "❌ NG (단면 부족)"
            ng_reason = f"전단철근이 부담하는 강도(Vs = {format_number(Vs_provided/1000, 1)} kN)가 최대 허용치(Vs,max = {format_number(Vs_max_limit/1000, 1)} kN)를 초과하여 단면 파괴가 우려됩니다. 단면 크기 상향이 필요합니다."
        elif not is_safe_strength:
            final_status_text = "❌ NG (강도 부족)"
            ng_reason = f"설계 전단강도(φVn = {format_number(phi_Vn/1000, 1)} kN)가 요구 강도(Vu = {format_number(Vu, 1)} kN)보다 작습니다. 철근량 증대 또는 단면 상향이 필요합니다."
        else:
            final_status_text = "✅ OK (안전)"
            ng_reason = ""

        # 결과 저장
        results.append({
            'case_num': i + 1,
            'Vu': Vu,
            'Pu': Pu,
            'Mu': Mu,
            'Mm': Mm / 1e6,  # kN·m 단위로 변환
            'vc_method': vc_method,
            'vc_formula': vc_formula,
            'Vc_N': Vc,
            'phi_Vc_N': phi_Vc,
            'shear_category': shear_category,
            'category_color': category_color_hex,
            'Vs_req_N': Vs_req,
            's_from_vs_req': s_from_vs_req,
            'min_Av_s_1_val': min_Av_s_1_val,
            'min_Av_s_2_val': min_Av_s_2_val,
            'min_Av_s_req': min_Av_s_req,
            's_from_min_req': s_from_min_req,
            's_max_code': s_max_code,
            's_max_reason': s_max_reason,
            'actual_s': actual_s,
            'stirrups_needed': stirrups_needed,
            'phi_Vs_N': phi_Vs,
            'phi_Vn_kN': phi_Vn / 1000,
            'Vs_provided_N': Vs_provided,
            'Vs_max_limit_N': Vs_max_limit,
            'is_safe_section': is_safe_section,
            'is_safe': is_safe_total,
            'final_status_text': final_status_text,
            'ng_reason': ng_reason,
            'stirrups_per_meter': stirrups_per_meter
        })

    # =================================================================
    # 4. 결과 출력 (요약표)
    # =================================================================
    st.markdown("---")
    st.markdown("## 📊 **전단설계 결과 요약**")
    
    def format_N_to_kN(value):
        return f"{value/1000:,.1f}"
    
    summary_data = []
    for r in results:
        summary_data.append({
            'Case': r['case_num'],
            'Vu (kN)': f"{r['Vu']:.1f}",
            'Pu (kN)': f"{r['Pu']:.1f}",
            'Mu (kN·m)': f"{r['Mu']:.1f}",
            'Mm (kN·m)': f"{r['Mm']:.1f}",
            'Vc 계산법': r['vc_method'],
            'φVc (kN)': format_N_to_kN(r['phi_Vc_N']),
            '판정': r['shear_category'],
            '배근': r['stirrups_needed'],
            'φVn (kN)': f"{r['phi_Vn_kN']:.1f}",
            '최종': '✅ OK' if r['is_safe'] else '❌ NG'
        })
    
    df_summary = pd.DataFrame(summary_data)
    st.dataframe(df_summary, use_container_width=True, hide_index=True)

    # =================================================================
    # 5. 상세 계산 과정 출력
    # =================================================================
    st.markdown("---")
    st.markdown("## 📝 **상세 계산 과정**")
    
    def step_header(text):
        st.markdown(f"""
        <div style="background: linear-gradient(90deg, green 0%, #764ba2 100%); 
                    color: white; padding: 15px 25px; border-radius: 10px; 
                    margin: 30px 0 20px 0; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
            <h3 style="margin: 0; font-size: 1.6em;">📌 {text}</h3>
        </div>
        """, unsafe_allow_html=True)

    for r in results:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                    padding: 10px; border-radius: 15px; margin: 40px 0; 
                    box-shadow: 0 8px 16px rgba(0,0,0,0.15);">
            <h2 style="color: white; margin: 0; font-size: 2.2em; text-align: center;">
                ⚙️ Case {r['case_num']} 상세 계산
            </h2>
        </div>
        """, unsafe_allow_html=True)

        step_header("1단계 : 설계 조건 확인")        
        st.markdown("<h5><b>■ 하중 조건</b></h5>", unsafe_allow_html=True)
        st.markdown(rf"""
            <div style="font-size:18px; line-height:1.8;">
            
            $\displaystyle
            \quad\quad \boldsymbol{{V_u}} = {r['Vu']:,.1f}\,\text{{kN}} \quad
            \boldsymbol{{P_u}} = {r['Pu']:,.1f}\,\text{{kN}} \quad
            \boldsymbol{{M_u}} = {r['Mu']:,.1f}\,\text{{kN}}·\text{{m}}
            $
            
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<h5><b>■ 부재 제원</b></h5>", unsafe_allow_html=True)
        st.markdown(rf"""
            <div style="font-size:18px; line-height:1.8;">
            
            $\displaystyle
            \quad\quad \boldsymbol{{b_w}} = {bw:,.0f}\,\text{{mm}} \quad
            \boldsymbol{{d}} = {d:,.0f}\,\text{{mm}} \quad
            \boldsymbol{{h}} = {h:,.0f}\,\text{{mm}} \quad
            (d_c = {In.dc[0]:,.1f}\,\text{{mm}})
            $
            
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<h5><b>■ 재료 특성</b></h5>", unsafe_allow_html=True)
        st.markdown(rf"""
            <div style="font-size:18px; line-height:1.8;">
            
            $\displaystyle
            \quad\quad \boldsymbol{{f_{{ck}}}} = {fck:.0f}\,\text{{MPa}} \quad
            \boldsymbol{{f_{{ys}}}} = {fy_shear:.0f}\,\text{{MPa}} \quad
            \boldsymbol{{\lambda}} = {lamda:.1f}
            $
            
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<h5><b>■ 배근 정보</b></h5>", unsafe_allow_html=True)
        st.markdown(rf"""
            <div style="font-size:18px; line-height:1.8;">
            
            &nbsp;&nbsp; • **인장측 철근**: $\boldsymbol{{A_{{st}}}} = {Ast_tension:,.1f}\,\text{{mm}}^2$
            
            &nbsp;&nbsp; • **압축측 철근**: $\boldsymbol{{A_{{sc}}}} = {Ast_compression:,.1f}\,\text{{mm}}^2$
            
            &nbsp;&nbsp; • **전단철근**: H{bar_dia}-{legs}leg $\quad (\boldsymbol{{A_v}} = {Av_stirrup:,.1f}\,\text{{mm}}^2)$
            
            </div>
            """, unsafe_allow_html=True)
        
        step_header("2단계 : M<sub>m</sub>에 의한 φV<sub>c</sub> 산정")
        st.markdown("<h5><b>■ M<sub>m</sub> (수정 모멘트) 계산</b></h5>", unsafe_allow_html=True)
        st.markdown(rf"""
            <div style="font-size:18px; line-height:1.8;">

            &nbsp;&nbsp; ① **일반식**  
            $\displaystyle
            \quad\quad \boldsymbol{{M_m}} = \boldsymbol{{M_u}} - \boldsymbol{{P_u}} \times \frac{{4h - d}}{{8}}
            $

            &nbsp;&nbsp;② **값 대입 및 계산**  
            $\displaystyle
            \quad\quad \boldsymbol{{M_m}} =
            {r['Mu']:,.1f} - {r['Pu']:,.1f} \times
            \frac{{(4\times{h:,.0f}-{d:,.0f})}}{{8 \times 1,000}}
            = \mathbf{{{r['Mm']:,.1f}}}\,\text{{kN}}·\text{{m}}
            $
            </div>
            """, unsafe_allow_html=True)


        step_header("3단계: 콘크리트 부담 전단강도 (φV<sub>c</sub>) 계산")
        st.markdown("<h5><b>■ φV<sub>c</sub> 산정식 선택</b></h5>", unsafe_allow_html=True)
        st.markdown(rf"""
            <div style="font-size:18px; line-height:1.8;">

            &nbsp;&nbsp; • <b>M<sub>m</sub> 값 확인</b> : 
            {r['Mm']:,.1f} kN·m 
            {('<span style="color:#059669; font-weight:700;">&lt; 0 <strong>(축력 고려식 적용)</strong></span>'
            if r['Mm'] < 0 
            else '<span style="color:#059669; font-weight:700;">&ge; 0 <strong>(정밀식 적용)</strong></span>')}
            </div>
            """, unsafe_allow_html=True)

        
        st.markdown("<h5><b>■ 축력 고려식 (M<sub>m</sub> < 0)</b></h5>", unsafe_allow_html=True)
        st.markdown(rf"""
            <div style="font-size:18px; line-height:1.8;">
            
            &nbsp;&nbsp; ① **일반식**  
            $\displaystyle
            \quad\quad \boldsymbol{{\phi V_c}} = \phi \times 0.29 \lambda \sqrt{{f_{{ck}}}} \times b_w \times d \times \sqrt{{1 + \frac{{N_u}}{{3.5 A_g}}}}
            $
            
            &nbsp;&nbsp;② **값 대입 및 계산**  
            $\displaystyle
            \quad\quad \boldsymbol{{\phi V_c}} = 
            0.75 \times 0.29 \times 1.0 \times \sqrt{{{fck}}} \times {bw:,.0f} \times {d:,.0f} \times \sqrt{{1 + \frac{{{r['Pu']:,.1f} \times 1,000}}{{3.5 \times {Ag:,.0f}}}}}
            $
            
            $\displaystyle
            \quad\quad\quad\quad\quad\,\,\, = \mathbf{{{r['phi_Vc_N']/1000:,.1f}}}\,\text{{kN}}
            $
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<h5><b>■ 정밀식 (전단력과 휨 모멘트 고려) (M<sub>m</sub> > 0)</b></h5>", unsafe_allow_html=True)
        phi_Vc = 0.75 * (1/6 * np.sqrt(fck) + 17.6 * rho_w * r['Vu'] * d / (r['Mm'] * 1000)) * bw * d
        # print("Mm : ", r['Mm'])
        st.markdown(rf"""
            <div style="font-size:18px; line-height:1.8;">
            
            &nbsp;&nbsp; ① **일반식**  
            $\displaystyle
            \quad\quad \boldsymbol{{\phi V_c}} = \phi \times \left(\frac{{1}}{{6}}\lambda\sqrt{{f_{{ck}}}}\; + \;17.6 \rho_w \frac{{V_u  d}}{{M_u}}\right) b_w \times d
            $
            
            &nbsp;&nbsp;② **철근비 계산**  
            $\displaystyle
            \quad\quad \rho_w = \frac{{A_s}}{{b_w \times d}} = \frac{{{As:,.0f}}}{{{bw:,.0f} \times {d:,.0f}}} = {rho_w:.4f}
            $
            
            &nbsp;&nbsp;③ **값 대입 및 계산**  
            $\displaystyle
            \quad\quad \boldsymbol{{\phi V_c}} = 
            0.75 \times \left(\frac{{1}}{{6}} \times 1.0 \times \sqrt{{{fck}}} \; + \; 17.6 \times {rho_w:.4f} \times \frac{{{r['Vu']:,.1f} \times {d:,.0f}}}{{{r['Mm']:,.1f} \times 1,000}}\right) \times {bw:,.0f} \times {d:,.0f}
            $
            
            $\displaystyle
            \quad\quad\quad\quad\quad\,\,\, = \mathbf{{{phi_Vc/1000:,.1f}}}\,\text{{kN}}
            $
            </div>
            """, unsafe_allow_html=True)

        step_header("4단계: 전단철근 판정")
        # # st.latex(r['vc_formula'])
        st.markdown(f"""
        <div class="calc-block" style="border-color: #4ecdc4;">
            <p>V<sub>c</sub> = <strong>{format_N_to_kN(r['Vc_N'])} kN</strong></p>
            <p>φV<sub>c</sub> = {phi_v} × {format_N_to_kN(r['Vc_N'])} = <strong>{format_N_to_kN(r['phi_Vc_N'])} kN</strong></p>
        </div>
        """, unsafe_allow_html=True)

        # 판정
        judgement_str = ""
        if r['shear_category'] == "전단철근 불필요":
            judgement_str = f"V<sub>u</sub> = {format_number(r['Vu'])} kN &nbsp; ≤ &nbsp; φV<sub>c</sub> = {format_N_to_kN(r['phi_Vc_N'])} kN"
        else:
            judgement_str = f"V<sub>u</sub> = {format_number(r['Vu'])} kN &nbsp; > &nbsp; φV<sub>c</sub> = {format_N_to_kN(r['phi_Vc_N'])} kN"

        st.markdown(f"""
        <div class="calc-block" style="border-color: {r['category_color']};">
            <p>{judgement_str}</p>
            <hr style='margin: 10px 0;'>
            <p style='font-size: 1.25em;'>판정 : <strong style='color:{r['category_color']};'>{r['shear_category']}</strong></p>
        </div>
        """, unsafe_allow_html=True)

        if r['shear_category'] != "전단철근 불필요":
            step_header("5단계: 필요 전단철근량 및 간격 계산")

            st.markdown("<h5><b>■ 전단철근 단면적 (A<sub>v</sub>) 산정</b></h5>", unsafe_allow_html=True)
            st.latex(r"A_v = n \times \frac{\pi D^2}{4}")
            st.markdown(f"""
            <div class="calc-block" style="border-color: #6f42c1;">
                <p>A<sub>v</sub> = {legs} × (π × {bar_dia}<sup>2</sup> ÷ 4) = <strong>{Av_stirrup:.1f} mm²</strong></p>
            </div>
            """, unsafe_allow_html=True)

            if r['shear_category'] == "설계전단철근":
                st.markdown("<h5><b>■ 전단철근이 부담할 필요 전단강도 (V<sub>s</sub>) 산정</b></h5>", unsafe_allow_html=True)
                st.latex(r"V_s = \frac{V_u - \phi V_c}{\phi_v}")
                st.markdown(f"""
                <div class="calc-block" style="border-color: #6f42c1;">
                    <p>V<sub>s,req</sub> = ({format_number(r['Vu'])} - {format_N_to_kN(r['phi_Vc_N'])}) ÷ {phi_v} = <strong>{format_N_to_kN(r['Vs_req_N'])} kN</strong></p>
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown("<h5><b>■ 강도 요구조건에 의한 간격 (s<sub>강도요구</sub>) 산정</b></h5>", unsafe_allow_html=True)
                st.latex(r"s_{강도요구} \leq \frac{A_v f_{yt} d}{V_s}")
                st.markdown(f"""
                <div class="calc-block" style="border-color: #6f42c1;">
                    <p>간격 (강도) = ({Av_stirrup:.1f} × {fy_shear} × {format_number(d,0)}) ÷ {format_number(r['Vs_req_N'], 0)} = <strong>{format_number(r['s_from_vs_req'])} mm</strong></p>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<h5><b>■ 최소 철근량 규정에 의한 간격 (s<sub>최소철근</sub>) 산정</b></h5>", unsafe_allow_html=True)
            st.latex(r"(\frac{A_v}{s})_{min} = \max\left(0.0625\sqrt{f_{ck}}\frac{b_w}{f_{yt}}, 0.35\frac{b_w}{f_{yt}}\right)")
            st.markdown(f"""
            <div class="calc-block" style="border-color: #6610f2;">
                <p>(A<sub>v</sub>/s)<sub>min</sub> = max( {r['min_Av_s_1_val']:.4f} × {bw}/{fy_shear}, {r['min_Av_s_2_val']} × {bw}/{fy_shear} )</p>
                <p>= max( {min_Av_s_1:.4f}, {min_Av_s_2:.4f} ) = <strong>{r['min_Av_s_req']:.4f}</strong></p>
            </div>
            """, unsafe_allow_html=True)
            
            st.latex(r"s_{최소철근} \leq \frac{A_v}{(\frac{A_v}{s})_{min}}")
            st.markdown(f"""
            <div class="calc-block" style="border-color: #6610f2;">
                <p>간격 (최소철근) = {Av_stirrup:.1f} ÷ {r['min_Av_s_req']:.4f} = <strong>{format_number(r['s_from_min_req'])} mm</strong></p>
            </div>
            """, unsafe_allow_html=True)

            step_header("6단계: 최대 허용 간격 결정")
            st.markdown(f"""            
                <p style="font-size:1.1em; color:#856404; margin:0;">{r['s_max_reason']}</p>            
            """, unsafe_allow_html=True)
            st.latex(r"s_{max} = \min(d/2, 600) \quad \text{or} \quad \min(d/4, 300)")
            st.markdown(f"""
            <div class="calc-block" style="border-color: #fd7e14;">
                <p>간격 (최대허용 기준) = <strong>{format_number(r['s_max_code'])} mm</strong></p>
            </div>
            """, unsafe_allow_html=True)
            
            step_header("7단계: 최종 전단철근 간격 결정")
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

            step_header(f"8단계: 최대 전단강도 검토 (단면 안전성)")
            st.latex(r"V_s \leq V_{s,max} = \frac{2}{3}\sqrt{f_{ck}}b_w d")
            
            section_check_color = "#28a745" if r['is_safe_section'] else "#dc3545"
            section_check_status = "OK" if r['is_safe_section'] else "NG"
            st.markdown(f"""
            <div class="calc-block" style="border-color: {section_check_color};">
                <p>V<sub>s,배근</sub> = φV<sub>s</sub> / φ<sub>v</sub> = {format_N_to_kN(r['phi_Vs_N'])} ÷ {phi_v} = <strong>{format_N_to_kN(r['Vs_provided_N'])} kN</strong></p>
                <p>V<sub>s,max</sub> = (2/3) × √{fck} × {format_number(bw, 0)} × {format_number(d, 0)} = <strong>{format_N_to_kN(r['Vs_max_limit_N'])} kN</strong></p>
                <hr style='margin: 10px 0;'>
                <p style='font-size: 1.25em;'>V<sub>s,배근</sub> ≤ V<sub>s,max</sub> 판정: <strong style='color:{section_check_color};'>{section_check_status}</strong></p>
            </div>

            """, unsafe_allow_html=True)
        # 단계 번호 조정
        final_check_start_num = 9
        if r['shear_category'] == "전단철근 불필요":
            final_check_start_num = 5

        step_header(f"{final_check_start_num}단계: 최종 안전성 검토")
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