import streamlit as st
import pandas as pd
import numpy as np

def check_shear(In, R):
    """
    🛡️ 전단설계 최적화 보고서 - KDS 14 20 기준 (최종 통합 최적화)
    기존 심플한 구조 + 요구사항 통합 + 가독성 극대화
    """

    # =================================================================
    # 0. 페이지 헤더 및 기본 UI 설정
    # =================================================================
    st.markdown("""
    <div style="background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
                 padding: 40px; border-radius: 20px; margin-bottom: 40px;
                 box-shadow: 0 10px 25px rgba(0,0,0,0.1);">
        <h1 style="color: white; text-align: center; margin: 0; font-size: 3.2em;
                   font-weight: 900; text-shadow: 1px 1px 3px rgba(0,0,0,0.2);">
            🛡️ 전단설계 최적화 보고서
        </h1>
        <p style="color: #e0e0e0; text-align: center; margin: 15px 0 0 0;
                  font-size: 1.3em; opacity: 0.9;">
            KDS 14 20 콘크리트구조설계기준 적용
        </p>
    </div>
    """, unsafe_allow_html=True)

    # =================================================================
    # 1. 설계 기준 및 이론 (컬러 구분)
    # =================================================================
    st.markdown("## 📋 **전단철근 판정 기준 (3단계)**")
    st.markdown("---")

    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #2196f3, #1976d2); 
                    color: white; padding: 25px; border-radius: 15px; 
                    box-shadow: 0 4px 8px rgba(0,0,0,0.15); margin-bottom: 20px;">
            <h3 style="margin: 0; text-align: center; font-size: 20px;">
                🔵 전단철근 불필요
            </h3>
            <div style="text-align: center; margin: 15px 0; font-size: 18px; font-weight: bold;">
                V<sub>u</sub> ≤ ½φV<sub>c</sub>
            </div>
            <p style="margin: 0; text-align: center; font-size: 14px;">
                이론적으로 전단철근 불필요
            </p>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #ffc107, #f57c00); 
                    color: white; padding: 25px; border-radius: 15px; 
                    box-shadow: 0 4px 8px rgba(0,0,0,0.15); margin-bottom: 20px;">
            <h3 style="margin: 0; text-align: center; font-size: 20px;">
                🟡 최소전단철근
            </h3>
            <div style="text-align: center; margin: 15px 0; font-size: 16px; font-weight: bold;">
                ½φV<sub>c</sub> < V<sub>u</sub> ≤ φV<sub>c</sub>
            </div>
            <p style="margin: 0; text-align: center; font-size: 14px;">
                규정 최소량 적용
            </p>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #f44336, #d32f2f); 
                    color: white; padding: 25px; border-radius: 15px; 
                    box-shadow: 0 4px 8px rgba(0,0,0,0.15); margin-bottom: 20px;">
            <h3 style="margin: 0; text-align: center; font-size: 20px;">
                🔴 설계전단철근
            </h3>
            <div style="text-align: center; margin: 15px 0; font-size: 18px; font-weight: bold;">
                V<sub>u</sub> > φV<sub>c</sub>
            </div>
            <p style="margin: 0; text-align: center; font-size: 14px;">
                계산에 의한 철근량
            </p>
        </div>
        """, unsafe_allow_html=True)

    # 간격 결정 원칙 (요구사항 반영: 2조건만 비교)
    st.markdown("### **⚙️ 간격 결정 원칙 (최적화)**")
    
    col1, col2 = st.columns([1, 1])
    with col1:
        st.latex(r"s_{final} = \min(s_{최소철근}, s_{최대허용})")
        st.markdown("**2조건만 비교하여 최적화**")
        
    with col2:
        st.latex(r"s_{최소철근} = \frac{A_v}{(A_v/s)_{min}}")
        st.latex(r"s_{최대허용} = \min(\frac{d}{2}, 600\text{mm})")

    # =================================================================
    # 2. 공통 설계 조건 및 계산 로직
    # =================================================================
    
    # 천단위 구분 함수
    def format_number(num, decimal_places=1):
        if decimal_places == 0:
            return f"{num:,.0f}"
        else:
            return f"{num:,.{decimal_places}f}"

    # --- 설계 상수 정의 ---
    phi_v = 0.75
    lamda = 1.0
    fy_shear = 400
    bar_dia = 13
    legs = 2
    bar_area = np.pi * (bar_dia / 2)**2
    Av_stirrup = bar_area * legs

    # --- 입력 변수 추출 ---
    bw, d, fck, Ag = In.be, In.depth, In.fck, R.Ag
    
    st.markdown("### **🔧 설계 상수 및 입력값**")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(f"""
        **설계 상수**
        - φᵥ = {phi_v}
        - λ = {lamda}  
        - fᵧₜ = {format_number(fy_shear, 0)} MPa
        """)
    
    with col2:
        st.markdown(f"""
        **부재 제원**
        - bw = {format_number(bw, 0)} mm
        - d = {format_number(d, 0)} mm
        - fck = {fck} MPa
        - Ag = {format_number(Ag, 0)} mm²
        """)
    
    with col3:
        st.markdown(f"""
        **전단철근 제원**
        - H{bar_dia} {legs}-leg stirrup
        - Av = {Av_stirrup:.1f} mm²
        """)

    results = [] # 결과를 저장할 리스트

    # =================================================================
    # 3. 각 하중 케이스별 계산 (요구사항 반영: 2조건만 비교)
    # =================================================================
    for i in range(len(In.Vu)):
        Vu = In.Vu[i]
        Pu = In.Pu_shear[i]

        # 1. 콘크리트 전단강도 (Vc, φVc)
        p_factor = 1 + (Pu * 1000) / (14 * Ag) if Pu != 0 else 1.0
        Vc = (1/6) * p_factor * lamda * np.sqrt(fck) * bw * d
        phi_Vc = phi_v * Vc

        # 2. 전단철근 필요성 판정
        half_phi_Vc = 0.5 * phi_Vc
        if Vu * 1000 <= half_phi_Vc:
            shear_category = "전단철근 불필요"
            category_color = "#2196f3"
        elif Vu * 1000 <= phi_Vc:
            shear_category = "최소전단철근"
            category_color = "#ffc107"
        else:
            shear_category = "설계전단철근"
            category_color = "#f44336"

        # 3. 간격 계산 (요구사항 반영: 2조건만 비교)
        # 3a. 최소철근량 요구사항
        min_Av_s_1 = 0.0625 * np.sqrt(fck) * (bw / fy_shear)
        min_Av_s_2 = 0.35 * (bw / fy_shear)
        min_Av_s_req = max(min_Av_s_1, min_Av_s_2)
        s_from_min_req = Av_stirrup / min_Av_s_req

        # 3b. 최대간격 규정
        Vs_req = (Vu * 1000 - phi_Vc) / phi_v if shear_category == "설계전단철근" else 0
        Vs_limit_d4 = (1/3) * np.sqrt(fck) * bw * d
        s_max_code = min(d / 4, 300) if Vs_req > Vs_limit_d4 else min(d / 2, 600)

        # 4. 최종 배근 간격 결정 (2조건만 비교)
        if shear_category == "전단철근 불필요":
            actual_s = s_max_code  # 최대허용간격 적용
            stirrups_needed = "전단철근 불필요"
        else:
            actual_s = min(s_from_min_req, s_max_code)  # 2조건만 비교
            # 시공성 고려, 5mm 단위 내림
            actual_s = np.floor(actual_s / 5) * 5
            stirrups_needed = f"H{bar_dia}-{legs}leg @{actual_s:.0f}mm"

        # 5. 최종 설계강도 및 안전성 검토
        if shear_category == "전단철근 불필요":
            phi_Vs = 0
        else:
            phi_Vs = (phi_v * Av_stirrup * fy_shear * d) / actual_s
        
        phi_Vn = phi_Vc + phi_Vs
        
        is_safe_strength = (phi_Vn >= Vu * 1000)
        Vs_max_limit = (2/3) * np.sqrt(fck) * bw * d
        Vs_provided = phi_Vs / phi_v if phi_Vs > 0 else 0
        is_safe_section = (Vs_provided <= Vs_max_limit)
        is_safe_total = is_safe_strength and is_safe_section
        
        safety_ratio = phi_Vn / (Vu * 1000) if Vu > 0 else float('inf')
        
        # 1m당 개수 계산
        stirrups_per_meter = 1000 / actual_s if actual_s > 0 else 0
        
        # 결과 저장
        results.append({
            'case': i + 1, 'Pu': Pu, 'Vu': Vu, 'shear_category': shear_category,
            'category_color': category_color, 'phi_Vn_kN': phi_Vn / 1000, 
            'safety_ratio': safety_ratio, 'is_safe': is_safe_total,
            'is_safe_section': is_safe_section, 'actual_s': actual_s,
            'stirrups_needed': stirrups_needed, 'stirrups_per_meter': stirrups_per_meter,
            # 상세 계산 과정용 데이터
            'p_factor':p_factor, 'Vc_N':Vc, 'phi_Vc_N':phi_Vc, 'half_phi_Vc_N':half_phi_Vc,
            'Vs_req_N':Vs_req, 'min_Av_s_req':min_Av_s_req, 's_from_min_req':s_from_min_req, 
            'Vs_limit_d4_N':Vs_limit_d4, 's_max_code':s_max_code, 'phi_Vs_N':phi_Vs, 
            'Vs_provided_N':Vs_provided, 'Vs_max_limit_N':Vs_max_limit
        })

    # =================================================================
    # 4. 전체 설계 결과 요약 (스타일링 개선)
    # =================================================================
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("## 📊 **전체 설계 결과 요약**")
    st.markdown("---")

    summary_data = []
    for r in results:
        if r['is_safe']:
            status_icon = "✅ 안전"
            status_color = "#28a745"
        else:
            status_icon = "❌ NG"
            status_color = "#dc3545"
            if not r['is_safe_section']:
                status_icon += " (단면!)"

        summary_data.append({
            'Case': f"Case {r['case']}",
            '하중조건 (kN)': f"Pu = {format_number(r['Pu'])}\nVu = {format_number(r['Vu'])}",
            '판정결과': r['shear_category'],
            '최적 설계': r['stirrups_needed'],
            '1m당 개수': f"{r['stirrups_per_meter']:.1f}개" if r['stirrups_per_meter'] > 0 else "-",
            '설계강도 (kN)': f"φVn = {format_number(r['phi_Vn_kN'])}",
            '안전율': f"{r['safety_ratio']:.3f}",
            '최종 판정': status_icon
        })
    
    df_summary = pd.DataFrame(summary_data)

    # 스타일 함수
    def style_safety(val):
        if "✅ 안전" in str(val):
            return 'background-color: #d4edda; color: #155724; font-weight: bold; text-align: center;'
        elif "❌ NG" in str(val):
            return 'background-color: #f8d7da; color: #721c24; font-weight: bold; text-align: center; border: 2px solid #dc3545;'
        return 'text-align: center;'

    def style_category(val):
        if "불필요" in str(val):
            return 'background-color: #e3f2fd; color: #1565c0; font-weight: bold; text-align: center;'
        elif "최소전단철근" in str(val):
            return 'background-color: #fff8e1; color: #e65100; font-weight: bold; text-align: center;'
        elif "설계전단철근" in str(val):
            return 'background-color: #ffebee; color: #c62828; font-weight: bold; text-align: center;'
        return 'text-align: center;'

    styled_summary = df_summary.style\
        .applymap(style_safety, subset=['최종 판정'])\
        .applymap(style_category, subset=['판정결과'])\
        .set_properties(**{'text-align': 'center', 'white-space': 'pre-line', 'padding': '10px'})\
        .set_table_styles([
            {'selector': 'th', 'props': [('background-color', '#495057'),
                                         ('color', 'white'),
                                         ('font-weight', 'bold'),
                                         ('text-align', 'center'),
                                         ('padding', '12px')]},
        ])

    st.dataframe(styled_summary, use_container_width=True, hide_index=True)

    # =================================================================
    # 5. 케이스별 상세 계산 과정 (LaTeX 수식 적용)
    # =================================================================
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("## ⚙️ **케이스별 상세 계산 과정**")
    st.markdown("---")

    def format_N_to_kN(value):
        return f"{value/1000:,.2f}"

    for i, r in enumerate(results):
        num_symbols = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"]
        st.markdown(f"# **{num_symbols[i]}번 검토**")
        with st.expander(f"### **Case {r['case']} 상세 계산** (Vu = {format_number(r['Vu'])} kN) - {r['shear_category']}"):
            
            # 단계별 계산 과정 표시
            st.markdown("#### **1단계: 축력 영향 계수**")
            st.latex(r"P_{증가} = 1 + \frac{P_u}{14A_g}")
            st.markdown(f"""
            <div style="background-color: #f8f9fa; padding: 15px; border-radius: 8px; border-left: 4px solid #007bff;">
                P증가 = 1 + {format_number(r['Pu']*1000, 0)} ÷ (14 × {format_number(Ag, 0)}) = <strong>{r['p_factor']:.3f}</strong>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("#### **2단계: 콘크리트 설계 전단강도**")
            st.latex(r"\phi V_c = \phi_v \times \frac{1}{6} \times P_{증가} \times \lambda \times \sqrt{f_{ck}} \times b_w \times d")
            st.markdown(f"""
            <div style="background-color: #f8f9fa; padding: 15px; border-radius: 8px; border-left: 4px solid #28a745;">
                φVc = {phi_v} × (1/6) × {r['p_factor']:.3f} × {lamda} × √{fck} × {format_number(bw, 0)} × {format_number(d, 0)} 
                = <strong>{format_N_to_kN(r['phi_Vc_N'])} kN</strong>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("#### **3단계: 전단철근 필요성 판정**")
            st.latex(r"V_u \text{ vs } \phi V_c, \frac{1}{2}\phi V_c")
            
            comparison_text = f"Vu = {format_number(r['Vu'])} kN, φVc = {format_N_to_kN(r['phi_Vc_N'])} kN, ½φVc = {format_N_to_kN(r['half_phi_Vc_N'])} kN"
            
            st.markdown(f"""
            <div style="background-color: {r['category_color']}22; padding: 15px; border-radius: 8px; 
                        border-left: 4px solid {r['category_color']};">
                <strong>비교:</strong> {comparison_text}<br>
                <strong>판정:</strong> <span style="color: {r['category_color']}; font-weight: bold; font-size: 18px;">{r['shear_category']}</span>
            </div>
            """, unsafe_allow_html=True)

            if r['shear_category'] != "전단철근 불필요":
                st.markdown("#### **4단계: 최소전단철근량 계산**")
                st.latex(r"\frac{A_v}{s} = \max\left(0.0625\sqrt{f_{ck}}\frac{b_w}{f_{yt}}, 0.35\frac{b_w}{f_{yt}}\right)")
                st.markdown(f"""
                <div style="background-color: #f8f9fa; padding: 15px; border-radius: 8px; border-left: 4px solid #6610f2;">
                    최소철근량: Av/s = <strong>{r['min_Av_s_req']:.4f} mm²/mm</strong><br>
                    최소철근 간격: s = {Av_stirrup:.1f} ÷ {r['min_Av_s_req']:.4f} = <strong>{format_number(r['s_from_min_req'])} mm</strong>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("#### **5단계: 간격 결정 (2조건 비교)**")
            st.latex(r"s_{final} = \min(s_{최소철근}, s_{최대허용})")
            
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"""
                **최소철근 간격**  
                s_min = {format_number(r['s_from_min_req'])} mm
                """)
            with col2:
                st.markdown(f"""
                **최대허용 간격**  
                s_max = {format_number(r['s_max_code'])} mm
                """)
            
            st.markdown(f"""
            <div style="background-color: #e8f5e8; padding: 15px; border-radius: 8px; border-left: 4px solid #28a745;">
                <strong>최종 적용간격:</strong> s = min({format_number(r['s_from_min_req'])}, {format_number(r['s_max_code'])}) = <strong>{format_number(r['actual_s'])} mm</strong><br>
                <strong>1m당 개수:</strong> 1,000 ÷ {format_number(r['actual_s'])} = <strong>{r['stirrups_per_meter']:.1f}개</strong>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("#### **6단계: 최종 안전성 검토**")
            st.latex(r"\phi V_n = \phi V_c + \phi V_s \geq V_u")
            st.latex(r"S.F. = \frac{\phi V_n}{V_u}")
            
            result_color = "#28a745" if r['is_safe'] else "#dc3545"
            result_icon = "✅" if r['is_safe'] else "❌ NG"
            
            st.markdown(f"""
            <div style="background-color: {result_color}22; padding: 20px; border-radius: 10px; 
                        border-left: 6px solid {result_color}; border: 2px solid {result_color};">
                <div style="font-size: 16px; font-weight: 600;">
                    <strong>최종 설계강도:</strong> φVn = {format_N_to_kN(r['phi_Vc_N'])} + {format_N_to_kN(r['phi_Vs_N'])} = <strong>{format_number(r['phi_Vn_kN'])} kN</strong><br>
                    <strong>안전성 확인:</strong> φVn = {format_number(r['phi_Vn_kN'])} kN {'≥' if r['is_safe'] else '<'} Vu = {format_number(r['Vu'])} kN<br>
                    <strong>안전율:</strong> S.F = {format_number(r['phi_Vn_kN'])} ÷ {format_number(r['Vu'])} = <strong>{r['safety_ratio']:.3f}</strong><br>
                    <strong style="color: {result_color}; font-size: 18px;">{result_icon} {r['shear_category']}</strong>
                </div>
            </div>
            """, unsafe_allow_html=True)
            st.markdown("---")

            if not r['is_safe_section']:
                st.error("⚠️ **경고**: 전단철근이 부담하는 강도(Vs)가 최대 허용치를 초과했습니다. 단면 크기 확대가 필요합니다.")

    # =================================================================
    # 6. 종합 평가 및 통계
    # =================================================================
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("## 🏆 **종합 평가 및 통계**")
    st.markdown("---")

    total_cases = len(results)
    safe_cases = sum(1 for r in results if r['is_safe'])
    unsafe_cases = total_cases - safe_cases

    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("총 케이스", f"{total_cases}개")
    with col2:
        st.metric("안전 케이스", f"{safe_cases}개", f"{safe_cases/total_cases*100:.1f}%")
    with col3:
        st.metric("NG 케이스", f"{unsafe_cases}개", f"{unsafe_cases/total_cases*100:.1f}%" if unsafe_cases > 0 else "0%")
    with col4:
        avg_safety = sum(r['safety_ratio'] for r in results) / len(results)
        st.metric("평균 안전율", f"{avg_safety:.3f}")

    # 최종 결론
    if unsafe_cases == 0:
        st.success("🎉 **모든 케이스가 안전합니다!** 제시된 전단철근 배근으로 시공 가능합니다.")
    else:
        ng_cases = [f"Case {r['case']}" for r in results if not r['is_safe']]
        st.error(f"❌ **{unsafe_cases}개 케이스가 NG입니다.** NG 케이스: {', '.join(ng_cases)} → 즉시 보강 설계 필요!")

    # placeholder 업데이트 (기존 코드와 호환)
    if hasattr(In, 'placeholder_shear'):
        for i, r in enumerate(results):
            if i < len(In.placeholder_shear) and In.placeholder_shear[i]:
                if r['is_safe']:
                    In.placeholder_shear[i].markdown(
                        f"<div style='text-align:center; color:#28a745; font-weight:900; "
                        f"font-size:16px; padding: 8px; border-radius: 8px; background-color: #d4edda; "
                        f"border: 2px solid #28a745;'>"
                        f"✅ {r['safety_ratio']:.2f}</div>",
                        unsafe_allow_html=True
                    )
                else:
                    In.placeholder_shear[i].markdown(
                        f"<div style='text-align:center; color:#dc3545; font-weight:900; "
                        f"font-size:16px; padding: 8px; border-radius: 8px; background-color: #f8d7da; "
                        f"border: 3px solid #dc3545;'>"
                        f"❌ NG </div>",
                        unsafe_allow_html=True
                    )

    return results

