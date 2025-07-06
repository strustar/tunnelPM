import streamlit as st
import pandas as pd
import numpy as np

def RC_and_AASHTO(Section_Type, Reinforcement_Type, beta1, c, eta, fck, Layer, ni, ep_si, ep_cu, dsi, fsi, Es, fy, Fsi, Asi, *bhD):
    """
    Calculates the nominal axial force (P) and moment (M) capacity of a reinforced concrete section.
    This function is based on strain compatibility and equilibrium of forces.
    """
    a = beta1 * c
    if 'Rectangle' in Section_Type:
        [hD, b, h] = bhD
        # Area of concrete in compression
        Ac = a * b if a < h else h * b
        # Centroid of the concrete compressive force from the section's geometric center
        y_bar = (h / 2) - (a / 2) if a < h else 0
    else: # Fallback for other shapes if needed
        [hD] = bhD
        Ac = 0 # Placeholder for other shapes
        y_bar = 0 # Placeholder

    # Compressive force from concrete
    Cc = eta * (0.85 * fck) * Ac / 1e3  # in kN
    M = 0

    # Loop through each layer of reinforcement
    # For this implementation, Layer=1 and ni=[2] representing top and bottom steel
    for L in range(Layer):
        for i in range(ni[L]):
            if c <= 0:  # Avoid division by zero or invalid neutral axis
                continue
            
            # Strain in steel layer based on linear strain distribution (strain compatibility)
            ep_si[L, i] = ep_cu * (c - dsi[L, i]) / c
            
            # Stress in steel based on linear elastic-perfectly plastic material model
            fsi[L, i] = Es * ep_si[L, i]
            
            # Apply yield strength limits
            fsi[L, i] = np.clip(fsi[L, i], -fy, fy)

            # Force in steel reinforcement
            if 'RC' in Reinforcement_Type or 'hollow' in Reinforcement_Type:
                # If steel is in compression zone (c >= dsi)
                if c >= dsi[L, i]:
                    # Subtract the force in the concrete area displaced by the rebar
                    Fsi[L, i] = Asi[L, i] * (fsi[L, i] - eta * 0.85 * fck) / 1e3  # in kN
                # If steel is in tension zone (c < dsi)
                else:
                    Fsi[L, i] = Asi[L, i] * fsi[L, i] / 1e3  # in kN
            
            # Sum moments of steel forces about the section's geometric center
            M = M + Fsi[L, i] * (hD / 2 - dsi[L, i])

    # Total nominal axial capacity (sum of forces)
    P = np.sum(Fsi) + Cc
    # Total nominal moment capacity (sum of moments)
    M = (M + Cc * y_bar) / 1e3  # in kN·m
    
    return P, M

def check_column(In, R, F):
    """
    Generates a real-time column strength check report in Streamlit
    using data from In, R, and F objects, with integrated KDS-2021 logic.
    """

    # =================================================================
    # 스타일 - 심플하고 깔끔한 디자인 (가독성 최우선)
    # =================================================================
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;600;700&display=swap');
    * { font-family: 'Noto Sans KR', sans-serif; }

    /* 최상단 레이아웃 */
    .main-container {
        background: #fdfdf8;
        padding: 30px;
        border-radius: 12px;
        margin: 20px 0;
        box-shadow: 0 3px 8px rgba(0, 0, 0, 0.08);
    }
    .main-header {
        background: #1e40af;
        color: #ffffff;
        font-size: 2.2em;
        font-weight: 700;
        padding: 25px;
        border-radius: 8px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin-bottom: 15px;           /* 하단 여백 축소 */
    }

    /* 공통 설계 조건 박스 */
    .common-conditions {
        text-align: center;
        background: #155e75;
        color: #e0f2fe;
        padding: 20px;                 /* 안쪽 여백 약간 축소 */
        border-radius: 8px;
        font-size: 1.8em;
        margin-top: 10px;              /* 상단과 간격 조정 */
        margin-bottom: 20px;           /* 하단 간격 조정 */
        box-shadow: 0 2px 4px rgba(0,0,0,0.08);
    }

    /* 리포트 컨테이너 */
    .report-container {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 30px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
        margin-bottom: 20px;
    }

    /* 섹션 헤더 */
    .section-header {
        background: #2563eb;
        color: #ffffff;
        font-size: 1.8em;             /* 폰트 크기 소폭 축소 */
        font-weight: 600;
        padding: 10px 18px;            /* 패딩 조정 */
        border-radius: 6px;
        text-align: center;
        margin: 15px 0;                /* 위아래 간격 통일 */
    }

    /* 소제목 */
    .sub-section-header {
        background: #1e3a8a;
        color: #ffffff;
        padding: 8px 16px;             /* 패딩 소폭 축소 */
        border-left: 4px solid #3b82f6;
        border-radius: 4px;
        font-size: 1.2em;             /* 폰트 크기 조정 */
        font-weight: 900;
        margin: 12px 0 8px;            /* 위아래 간격 조정 */
    }

    
    /* --- 상세 계산 박스 및 내부 요소 스타일 (여기를 중심으로 수정됨) --- */
    
    /* 1. 상세 계산 박스 전체 (기본 스타일) */
    .detailed-calc-container {
        background-color: #fafafc;
        border-left: 4px solid #3b82f6;
        box-shadow: 0 3px 6px rgba(0,0,0,0.08);
        border-radius: 8px;
        padding: 22px;
        margin-top: 15px;
        line-height: 1.8; /* 줄간격 조정 */
        
        /* 기본 텍스트: Noto Sans KR, 검은색 계열, 20px */
        font-family: 'Noto Sans KR', sans-serif;
        font-size: 20px; 
        font-weight: 600; /* 보통 굵기 */
        color: #212529;   /* 검은색 계열 */
    }

    /* 2. 일반 텍스트의 제목 (b 태그) */
    .detailed-calc-container b {
        font-weight: 600; /* 굵게 */
        color: #212529;   /* 검은색 계열 */
    }

    /* 3. 수식 스타일 (분홍색) */
    .math-expr {
        font-family: 'Times New Roman', serif; /* 수식만 다른 글꼴 적용 */
        font-style: italic;
        font-size: 1.05em; /* 주변 글자와 크기 조화 */
        font-weight: 600;
        color: #000000;   /* << 원하시는 분홍색 */
        margin: 0 2px; /* 좌우 여백 */
    }

    /* 4. 코드 블록 스타일 (분홍색) */
    .detailed-calc-container code {
        font-family: 'Consolas', monospace; /* 코드용 글꼴 */
        padding: 3px 6px;
        border-radius: 4px;
        font-size: 1.05em;
        font-weight: 600;
        color:#000000;   /* << 원하시는 분홍색 */
    }
    .detailed-calc-container li {
        font-weight: 600;
    }

    /* 5. 상태 표시 'OK' (분홍색 테마) */
    .ok, .pass-badge {
        color: #0000ff;          /* 진한 분홍색 텍스트 */
        # background-color: #ffffff;  /* 연한 분홍색 배경 */
        padding: 2px 6px;
        border-radius: 3px;
        border: 2px solid #0000ff;
        font-weight: 600;
        display: inline-block;
        font-family: 'Noto Sans KR', sans-serif; /* 폰트 일관성 유지 */
    }
    
    /* 6. 상태 표시 'NG' (가독성을 위해 적색 유지) */
    .ng, .fail-badge {
        color: #dc2626;
        background-color: #fee2e2;
        padding: 2px 6px;
        border-radius: 3px;
        border: 2px solid #ef4444;
        font-weight: 600;
        display: inline-block;
        font-family: 'Noto Sans KR', sans-serif; /* 폰트 일관성 유지 */
    }            

    </style>
    """, unsafe_allow_html=True)


    # =================================================================
    # 계산 헬퍼 함수
    # =================================================================
    def render_detailed_strength_check(In, PM_obj, material_type, case_idx):
        """모든 하중조합에 대한 상세 강도 검토 과정을 수식과 함께 HTML로 렌더링"""
        try:
            # --- 1. 데이터 추출 ---
            Pu_values, Mu_values = getattr(In, 'Pu', []), getattr(In, 'Mu', [])
            if hasattr(Pu_values, 'tolist'): Pu_values = Pu_values.tolist()
            if hasattr(Mu_values, 'tolist'): Mu_values = Mu_values.tolist()

            Reinforcement_Type = 'hollow' if material_type == '중공철근' else 'RC'

            if material_type == '이형철근':
                c_values, phiPn_values, phiMn_values, SF_values = getattr(In, 'c_RC', []), getattr(In, 'Pd_RC', []), getattr(In, 'Md_RC', []), getattr(In, 'safe_RC', [])
                e_min, fy, Es = getattr(R, 'e', [0,20,20])[1], getattr(In, 'fy', 400.0), getattr(In, 'Es', 200000.0)
            else: # 중공철근
                c_values, phiPn_values, phiMn_values, SF_values = getattr(In, 'c_FRP', []), getattr(In, 'Pd_FRP', []), getattr(In, 'Md_FRP', []), getattr(In, 'safe_FRP', [])
                e_min, fy, Es = getattr(F, 'e', [0,20,20])[1], getattr(In, 'fy_hollow', 800.0), getattr(In, 'Es_hollow', 200000.0)  # 중공철근 항복강도 800 MPa
            
            if hasattr(c_values, 'tolist'): c_values = c_values.tolist()
            if hasattr(phiPn_values, 'tolist'): phiPn_values = phiPn_values.tolist()
            if hasattr(phiMn_values, 'tolist'): phiMn_values = phiMn_values.tolist()
            if hasattr(SF_values, 'tolist'): SF_values = SF_values.tolist()

            Pu, Mu, c_assumed, phiPn, phiMn, SF = Pu_values[case_idx], Mu_values[case_idx], c_values[case_idx], phiPn_values[case_idx], phiMn_values[case_idx], SF_values[case_idx]
            e_actual = (Mu / Pu) * 1000 if Pu != 0 else 0

            # --- 2. 계산을 위한 재료 및 단면 속성 설정 (사용자 제공 로직 통합) ---
            h, b, fck = getattr(In, 'height', 300), getattr(In, 'be', 1000), getattr(In, 'fck', 40.0)
            RC_Code = getattr(In, 'RC_Code', 'KDS-2021')
            Column_Type = getattr(In, 'Column_Type', 'Tied Column')
            
            # --- 2a. KDS-2021 기준 계수 계산 ---
            if 'KDS-2021' in RC_Code:
                [n, ep_co, ep_cu] = [2, 0.002, 0.0033]
                if fck > 40:
                    n = 1.2 + 1.5 * ((100 - fck) / 60) ** 4
                    ep_co = 0.002 + (fck - 40) / 1e5
                    ep_cu = 0.0033 - (fck - 40) / 1e5
                if n >= 2:
                    n = 2
                n = round(n * 100) / 100

                alpha = 1 - 1 / (1 + n) * (ep_co / ep_cu)
                temp = 1 / (1 + n) / (2 + n) * (ep_co / ep_cu) ** 2
                if fck <= 40:
                    alpha = 0.8
                beta = 1 - (0.5 - temp) / alpha
                if fck <= 50:
                    beta = 0.4

                [alpha, beta] = [round(alpha * 100) / 100, round(beta * 100) / 100]
                beta1 = 2 * beta        
                eta = alpha / beta1
                eta = round(eta * 100) / 100
                if fck == 50:
                    eta = 0.97
                if fck == 80:
                    eta = 0.87

            # --- 2b. 철근 배치 설정 (사용자 제공 로직 통합) ---
            Layer_in, dia, dc, nh, nb, nD, sb, dia1, dc1 = In.Layer, In.dia, In.dc, In.nh, In.nb, In.nD, In.sb, In.dia1, In.dc1
            
            Layer = 1
            ni = [2] # 압축측, 인장측 철근 그룹
            
            nst = b / sb[0]
            nst1 = b / sb[0]
            
            # 중공철근의 경우 단면적 절반 적용
            area_factor = 0.5 if 'hollow' in Reinforcement_Type else 1.0
            
            Ast = [np.pi * d**2 / 4 * area_factor for d in dia]
            Ast1 = [np.pi * d**2 / 4 * area_factor for d in dia1]

            dsi = np.zeros((Layer, ni[0]))
            Asi = np.zeros((Layer, ni[0]))
            
            dsi[0, :] = [dc1[0], h - dc[0]]
            Asi[0, :] = [Ast1[0] * nst1, Ast[0] * nst]

            ep_si, fsi, Fsi = np.zeros_like(dsi), np.zeros_like(dsi), np.zeros_like(dsi)

            # --- 3. 공칭강도(Pn, Mn) 계산 ---
            [Pn, Mn] = RC_and_AASHTO('Rectangle', Reinforcement_Type, beta1, c_assumed, eta, fck, Layer, ni, ep_si, ep_cu, dsi, fsi, Es, fy, Fsi, Asi, h, b, h)
            e_calc = (Mn * 1000 / Pn) if Pn != 0 else 0
            equilibrium_diff = abs(e_calc - e_actual)
            equilibrium_check = equilibrium_diff / max(abs(e_actual), 1) <= 0.01  # 1% 이하 오차 허용

            # --- 4. 계산 과정에 필요한 중간값 추출 ---
            a = beta1 * c_assumed
            Ac = min(a, h) * b  # 콘크리트 압축 면적
            Cc = eta * (0.85 * fck) * Ac / 1000  # 콘크리트 압축력
            y_bar = (h / 2) - (a / 2) if a < h else 0  # 콘크리트 압축력 중심
            
            Cs_force = Fsi[0, 0] # 압축측 철근 힘
            Ts_force = Fsi[0, 1] # 인장측 철근 힘
            
            # 모멘트 계산 상세 과정
            Cc_moment = Cc * y_bar  # 콘크리트 압축력 모멘트
            Cs_moment = Cs_force * (h/2 - dsi[0, 0])  # 압축철근 모멘트
            Ts_moment = Ts_force * (h/2 - dsi[0, 1])  # 인장철근 모멘트
            
            # 철근 단면적 계산
            As1_calc = Ast1[0] * nst1  # 압축측 철근 단면적
            As_calc = Ast[0] * nst  # 인장측 철근 단면적

            # --- 5. 강도감소계수(φ) 계산 (통합된 로직) ---
            dt = dsi[0, 1] # 인장측 철근 깊이
            eps_t = ep_cu * (dt - c_assumed) / c_assumed if c_assumed > 0 else 0
            eps_y = fy / Es
            phi_factor, phi_basis = 0.65, ""

            if 'RC' in Reinforcement_Type or 'hollow' in Reinforcement_Type:
                phi0 = 0.70 if 'Spiral' in Column_Type else 0.65
                ep_tccl = eps_y # 압축지배 한계 변형률
                ep_ttcl = 0.005 if fy < 400 else 2.5 * eps_y # 인장지배 한계 변형률
                
                if eps_t <= ep_tccl:
                    phi_factor = phi0
                    phi_basis = f"<span class='math-expr'>ε<sub>t</sub></span> = {eps_t:.5f} ≤ <span class='math-expr'>ε<sub>ty</sub></span> = {ep_tccl:.5f} 이므로, <b>압축지배단면 (φ={phi0:.2f})</b>입니다."
                elif eps_t >= ep_ttcl:
                    phi_factor = 0.85
                    phi_basis = f"<span class='math-expr'>ε<sub>t</sub></span> = {eps_t:.5f} ≥ {ep_ttcl:.5f} 이므로, <b>인장지배단면 (φ=0.85)</b>입니다."
                else: # 변화구간
                    phi_factor = phi0 + (0.85 - phi0) * (eps_t - ep_tccl) / (ep_ttcl - ep_tccl)
                    phi_basis = f"<span class='math-expr'>ε<sub>ty</sub></span>({ep_tccl:.5f}) < <span class='math-expr'>ε<sub>t</sub></span>({eps_t:.5f}) < {ep_ttcl:.5f} 이므로, <b>변화구간</b>에 해당합니다."

            # --- 6. 개선된 안전율 계산 (교점 거리비 방식) ---
            safety_factor = np.sqrt(phiPn**2 + phiMn**2) / np.sqrt(Pu**2 + Mu**2) if Pu > 0 and Mu > 0 else 0
            sf_status = "안전" if safety_factor >= 1.0 else "위험"
            sf_color = "ok" if safety_factor >= 1.0 else "ng"

            # --- 7. 부등호 및 판정 설정 ---
            p_inequality = "≤" if Pu <= phiPn else ">"
            p_status = "O.K." if Pu <= phiPn else "N.G."
            p_color = "ok" if Pu <= phiPn else "ng"
            
            m_inequality = "≤" if Mu <= phiMn else ">"
            m_status = "O.K." if Mu <= phiMn else "N.G."
            m_color = "ok" if Mu <= phiMn else "ng"

            # --- 8. 상세 계산 과정 HTML 생성 ---
            html = f"""
            <div class="detailed-calc-container">
                <div style="font-size: 1.3em; font-weight: 800; color: #1e40af; background: linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%); padding: 12px 20px; border-radius: 8px; margin-bottom: 20px; text-align: center; border: 2px solid #3b82f6;">
                    [하중조합 LC-{case_idx+1} 상세 계산 과정]
                </div>
                <br>
                <b>1. 기본 정보 및 설계계수</b>
                <ul>
                    <li>적용 기준: <code>{RC_Code}</code>, 기둥 형식: <code>{Column_Type}</code></li>
                    <li>콘크리트 계수: <code><span class='math-expr'>β₁</span> = {beta1:.3f}</code>, <code><span class='math-expr'>η</span> = {eta:.3f}</code>, <code><span class='math-expr'>ε<sub>cu</sub></span> = {ep_cu:.5f}</code></li>
                    <li>철근 재료: <code><span class='math-expr'>f<sub>y</sub></span> = {fy:,.0f} MPa</code>, <code><span class='math-expr'>E<sub>s</sub></span> = {Es:,.0f} MPa</code> {'(중공철근)' if 'hollow' in Reinforcement_Type else '(이형철근)'}</li>
                    <li>작용 하중: <code><span class='math-expr'>P<sub>u</sub></span> = {Pu:,.1f} kN</code>, <code><span class='math-expr'>M<sub>u</sub></span> = {Mu:,.1f} kN·m</code> (편심 <code>e = {e_actual:,.3f} mm</code>)</li>
                    <li>가정된 중립축: <code>c = {c_assumed:,.3f} mm</code> (시행착오법으로 결정)</li>
                </ul><hr>
                <b>2. 변형률 호환 및 응력 계산</b>
                <ul>
                    <li><b>변형률 계산:</b> <code><span class='math-expr'>ε<sub>s</sub> = ε<sub>cu</sub> × (c - d<sub>s</sub>) / c</span></code></li>
                    <li>압축측 철근 (d<sub>s</sub>={dsi[0,0]:.1f}mm): <code><span class='math-expr'>ε<sub>sc</sub></span> = {ep_si[0,0]:.5f}</code> → <code><span class='math-expr'>f<sub>sc</sub></span> = {fsi[0,0]:,.2f} MPa</code></li>
                    <li>인장측 철근 (d<sub>t</sub>={dsi[0,1]:.1f}mm): <code><span class='math-expr'>ε<sub>st</sub></span> = {ep_si[0,1]:.5f}</code> → <code><span class='math-expr'>f<sub>st</sub></span> = {fsi[0,1]:,.2f} MPa</code></li>
                </ul><hr>
                <b>3. 단면력 평형 및 공칭강도 계산</b>
                <ul>
                    <li>등가응력블록 깊이: <code>a = <span class='math-expr'>β₁</span> × c = {beta1:.3f} × {c_assumed:.3f} = {a:.3f} mm</code></li>
                    <li>콘크리트 압축면적: <code><span class='math-expr'>A<sub>c</sub></span> = min(a, h) × b = {min(a, h):.1f} × {b:.1f} = {Ac:,.1f} mm²</code></li>
                    <li>콘크리트 압축력: <code><span class='math-expr'>C<sub>c</sub></span> = η × 0.85 × <span class='math-expr'>f<sub>ck</sub></span> × <span class='math-expr'>A<sub>c</sub></span> = {eta:.3f} × 0.85 × {fck:.1f} × {Ac:,.1f} = {Cc:,.1f} kN</code></li>
                    <li>압축측 철근 단면적: <code><span class='math-expr'>A<sub>s1</sub></span> = {As1_calc:,.1f} mm²</code></li>
                    <li>압축측 철근 합력: <code><span class='math-expr'>C<sub>s</sub></span> = <span class='math-expr'>A<sub>s1</sub></span> × (<span class='math-expr'>f<sub>sc</sub></span> - η × 0.85 × <span class='math-expr'>f<sub>ck</sub></span>) = {As1_calc:,.1f} × ({fsi[0,0]:,.2f} - {eta:.3f} × 0.85 × {fck:.1f}) = {Cs_force:,.1f} kN</code></li>
                    <li>인장측 철근 단면적: <code><span class='math-expr'>A<sub>s</sub></span> = {As_calc:,.1f} mm²</code></li>
                    <li>인장측 철근 합력: <code><span class='math-expr'>T<sub>s</sub></span> = <span class='math-expr'>A<sub>s</sub></span> × <span class='math-expr'>f<sub>st</sub></span> = {As_calc:,.1f} × {fsi[0,1]:,.2f} = {Ts_force:,.1f} kN</code></li>
                    <li><b>공칭 축강도:</b> <code><span class='math-expr'>P<sub>n</sub></span> = <span class='math-expr'>C<sub>c</sub></span> + <span class='math-expr'>C<sub>s</sub></span> + <span class='math-expr'>T<sub>s</sub></span> = {Cc:,.1f}{Cs_force:+.1f}{Ts_force:+.1f} = {Pn:,.1f} kN</code></li>
                </ul><hr>
                <b>4. 공칭 휨강도 계산</b>
                <ul>
                    <li>콘크리트 압축력 중심: <code><span class='math-expr'>ȳ</span> = (h/2) - (a/2) = ({h:.1f}/2) - ({a:.1f}/2) = {y_bar:.1f} mm</code></li>
                    <li>콘크리트 압축력 모멘트: <code><span class='math-expr'>M<sub>c</sub></span> = <span class='math-expr'>C<sub>c</sub></span> × <span class='math-expr'>ȳ</span> = {Cc:,.1f} × {y_bar:.1f} = {Cc_moment:,.1f} kN·mm</code></li>
                    <li>압축철근 모멘트팔: <code>(h/2) - d<sub>s1</sub> = ({h:.1f}/2) - {dsi[0,0]:.1f} = {(h/2 - dsi[0,0]):.1f} mm</code></li>
                    <li>압축철근 모멘트: <code><span class='math-expr'>M<sub>s1</sub></span> = <span class='math-expr'>C<sub>s</sub></span> × (h/2 - d<sub>s1</sub>) = {Cs_force:,.1f} × {(h/2 - dsi[0,0]):.1f} = {Cs_moment:,.1f} kN·mm</code></li>
                    <li>인장철근 모멘트팔: <code>(h/2) - d<sub>t</sub> = ({h:.1f}/2) - {dsi[0,1]:.1f} = {(h/2 - dsi[0,1]):.1f} mm</code></li>
                    <li>인장철근 모멘트: <code><span class='math-expr'>M<sub>s</sub></span> = <span class='math-expr'>T<sub>s</sub></span> × (h/2 - d<sub>t</sub>) = {Ts_force:,.1f} × {(h/2 - dsi[0,1]):.1f} = {Ts_moment:,.1f} kN·mm</code></li>
                    <li><b>공칭 휨강도:</b> <code><span class='math-expr'>M<sub>n</sub></span> = (<span class='math-expr'>M<sub>c</sub></span> + <span class='math-expr'>M<sub>s1</sub></span> + <span class='math-expr'>M<sub>s</sub></span>) / 1000 = ({Cc_moment:,.1f} + {Cs_moment:,.1f} + {Ts_moment:,.1f}) / 1000 = {Mn:,.1f} kN·m</code></li>
                </ul><hr>
                <b>5. 강도감소계수(φ) 및 설계강도</b>
                <ul>
                    <li><b>판단 근거:</b> {phi_basis}</li>
                    <li>결정된 강도감소계수: <code>φ = {phi_factor:.3f}</code></li>
                    <li><b>설계 축강도:</b> <code>φ<span class='math-expr'>P<sub>n</sub></span> = {phi_factor:.3f} × {Pn:,.1f} = {Pn*phi_factor:,.1f} kN</code></li>
                    <li><b>설계 휨강도:</b> <code>φ<span class='math-expr'>M<sub>n</sub></span> = {phi_factor:.3f} × {Mn:,.1f} = {Mn*phi_factor:,.1f} kN·m</code></li>
                </ul><hr>
                <b>6. 최종 검토 및 안전성 평가</b>
                <ul>
                    <li><b>평형조건 검토:</b> 
                        <ul>
                            <li>계산편심: <code>e' = <span class='math-expr'>M<sub>n</sub></span> / <span class='math-expr'>P<sub>n</sub></span> = {Mn:,.1f} / {Pn:,.1f} × 1000 = {e_calc:.3f} mm</code></li>
                            <li>작용편심: <code>e = <span class='math-expr'>M<sub>u</sub></span> / <span class='math-expr'>P<sub>u</sub></span> × 1000 = {Mu:,.1f} / {Pu:,.1f} × 1000 = {e_actual:.3f} mm</code></li>
                            <li>상대오차: <code>|e' - e| / e = |{e_calc:.3f} - {e_actual:.3f}| / {e_actual:.3f} = {equilibrium_diff/max(abs(e_actual), 1)*100:.2f}%</code> <span class="{'ok' if equilibrium_check else 'ng'}">{'≤ 1% (O.K.)' if equilibrium_check else '> 1%'}</span></li>
                        </ul>
                    </li>
                    <li><b>강도조건 검토:</b>
                        <ul>
                            <li>축력 검토: <code><span class='math-expr'>P<sub>u</sub></span> = {Pu:,.1f} kN {p_inequality} φ<span class='math-expr'>P<sub>n</sub></span> = {phiPn:,.1f} kN</code> <span class="{p_color}"><b>∴ {p_status}</b></span></li>
                            <li>휨강도 검토: <code><span class='math-expr'>M<sub>u</sub></span> = {Mu:,.1f} kN·m {m_inequality} φ<span class='math-expr'>M<sub>n</sub></span> = {phiMn:,.1f} kN·m</code> <span class="{m_color}"><b>∴ {m_status}</b></span></li>
                            <li><b>PM 상관도 교점 안전율:</b> <code>S.F. = √[(φ<span class='math-expr'>P<sub>n</sub></span>)² + (φ<span class='math-expr'>M<sub>n</sub></span>)²] / √[<span class='math-expr'>P<sub>u</sub></span>² + <span class='math-expr'>M<sub>u</sub></span>²] = √[{phiPn:,.1f}² + {phiMn:,.1f}²] / √[{Pu:,.1f}² + {Mu:,.1f}²] = {safety_factor:.3f}</code> <span class="{sf_color}"><b>({sf_status})</b></span></li>
                        </ul>
                    </li>
                </ul>
            </div>
            <br><br>
            """
            return html

        except Exception as e:
            import traceback
            st.error(f"LC-{case_idx+1} 상세 검토 중 오류 발생: {e}")
            st.code(traceback.format_exc())
            return f'<div class="detailed-calc-container">⚠️ LC-{case_idx+1} 상세 검토 중 오류가 발생했습니다.</div>'

    def extract_pm_data(PM_obj, material_type, In):
        try:
            def safe_extract(attr_name, default=[]):
                values = getattr(PM_obj, attr_name, default)
                if hasattr(values, 'tolist'): return values.tolist()
                elif isinstance(values, (list, tuple)): return list(values)
                else: return [float(values)] if values is not None else []
            
            pm_data = {
                'e_mm': safe_extract('Ze'), 'c_mm': safe_extract('Zc'), 'Pn_kN': safe_extract('ZPn'),
                'Mn_kNm': safe_extract('ZMn'), 'phi': safe_extract('Zphi'), 'phiPn_kN': safe_extract('ZPd'),
                'phiMn_kNm': safe_extract('ZMd')
            }
            
            try:
                Pd_values, Md_values, e_values, c_values = safe_extract('Pd'), safe_extract('Md'), safe_extract('e'), safe_extract('c')
                balanced_data = {
                    'Pb_kN': float(Pd_values[3]) if len(Pd_values) > 3 else 0.0,
                    'Mb_kNm': float(Md_values[3]) if len(Md_values) > 3 else 0.0,
                    'eb_mm': float(e_values[3]) if len(e_values) > 3 else 0.0,
                    'cb_mm': float(c_values[3]) if len(c_values) > 3 else 0.0
                }
            except (IndexError, ValueError, TypeError):
                balanced_data = {'Pb_kN': 0.0, 'Mb_kNm': 0.0, 'eb_mm': 0.0, 'cb_mm': 0.0}
            
            if material_type == '이형철근':
                material_props = {
                    'fy': float(getattr(In, 'fy', 400.0)),
                    'Es': float(getattr(In, 'Es', 200000.0)) / 1000
                }
            else:  # 중공철근
                material_props = {
                    'fy': float(getattr(In, 'fy_hollow', 800.0)),  # 중공철근 항복강도 800 MPa
                    'Es': float(getattr(In, 'Es_hollow', 200000.0)) / 1000
                }
            
            return pm_data, balanced_data, material_props
            
        except Exception as e:
            st.error(f"데이터 추출 중 오류 발생: {e}")
            return {}, {'Pb_kN': 0.0, 'Mb_kNm': 0.0, 'eb_mm': 0.0, 'cb_mm': 0.0}, {'fy': 0.0, 'Es': 0.0}

    def calculate_strength_check(In, material_type):
        try:
            Pu_values, Mu_values = getattr(In, 'Pu', []), getattr(In, 'Mu', [])
            if hasattr(Pu_values, 'tolist'): Pu_values = Pu_values.tolist()
            if hasattr(Mu_values, 'tolist'): Mu_values = Mu_values.tolist()
            
            if material_type == '이형철근':
                safety_factors, Pd_values, Md_values = getattr(In, 'safe_RC', []), getattr(In, 'Pd_RC', []), getattr(In, 'Md_RC', [])
            else:
                safety_factors, Pd_values, Md_values = getattr(In, 'safe_FRP', []), getattr(In, 'Pd_FRP', []), getattr(In, 'Md_FRP', [])
            
            if hasattr(safety_factors, 'tolist'): safety_factors = safety_factors.tolist()
            if hasattr(Pd_values, 'tolist'): Pd_values = Pd_values.tolist()
            if hasattr(Md_values, 'tolist'): Md_values = Md_values.tolist()
            
            if not Pu_values or not Mu_values: return []
            
            checks = []
            for i in range(min(len(Pu_values), len(Mu_values))):
                try:
                    Pu, Mu = float(Pu_values[i]), float(Mu_values[i])
                    Pd, Md = float(Pd_values[i]), float(Md_values[i])
                    
                    # PM 상관도 교점 거리비 안전율 계산
                    safety_factor = np.sqrt(Pd**2 + Md**2) / np.sqrt(Pu**2 + Mu**2) if Pu > 0 and Mu > 0 else 0
                    
                    e = (Mu / Pu) * 1000 if Pu != 0 else 0
                    verdict = 'PASS' if safety_factor >= 1.0 else 'FAIL'
                    
                    checks.append({
                        'LC': f'LC-{i+1}', 
                        'Pu/phiPn': f'{Pu:,.1f} / {Pd:,.1f}', 
                        'Mu/phiMn': f'{Mu:,.1f} / {Md:,.1f}',
                        'e_mm': e, 
                        'SF': safety_factor, 
                        'Verdict': verdict
                    })
                except (ValueError, TypeError, ZeroDivisionError, IndexError):
                    checks.append({
                        'LC': f'LC-{i+1}', 
                        'Pu/phiPn': 'ERROR', 
                        'Mu/phiMn': 'ERROR', 
                        'e_mm': 0.0, 
                        'SF': 0.0, 
                        'Verdict': 'FAIL'
                    })
            return checks
        except Exception as e:
            st.error(f"강도 검토 계산 중 오류 발생: {e}")
            return []
            
    def render_common_conditions(In):
        st.markdown('<div class="common-conditions"><div class="common-header">🏗️ 공통 설계 조건</div>', unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            be, h, sb = getattr(In, 'be', 1000), getattr(In, 'height', 300), getattr(In, 'sb', [150.0])[0]
            st.markdown(f'''<table class="common-table"><tr><td colspan="2" style="text-align:center;font-weight:bold;">📐 단면 제원</td></tr><tr><td><span class="icon">📏</span>단위폭 b<sub>e</sub></td><td>{be:,.1f} mm</td></tr><tr><td><span class="icon">📏</span>단면 두께 h</td><td>{h:,.1f} mm</td></tr><tr><td><span class="icon">📐</span>공칭 철근간격 s</td><td>{sb:,.1f} mm</td></tr></table>''', unsafe_allow_html=True)
        
        with col2:
            fck, Ec = getattr(In, 'fck', 40.0), getattr(In, 'Ec', 30000.0) / 1000
            st.markdown(f'''<table class="common-table"><tr><td colspan="2" style="text-align:center;font-weight:bold;">🏭 콘크리트 재료</td></tr><tr><td><span class="icon">💪</span>압축강도 f<sub>ck</sub></td><td>{fck:,.1f} MPa</td></tr><tr><td><span class="icon">⚡</span>탄성계수 E<sub>c</sub></td><td>{Ec:,.1f} GPa</td></tr><tr><td style="opacity:0;"></td><td style="opacity:0;"></td></tr></table>''', unsafe_allow_html=True)
        
        with col3:
            dm, code, ct = getattr(In, 'Design_Method', 'USD').split('(')[0].strip(), getattr(In, 'RC_Code', 'KDS-2021'), getattr(In, 'Column_Type', 'Tied Column')
            st.markdown(f'''<table class="common-table"><tr><td colspan="2" style="text-align:center;font-weight:bold;">📋 설계 조건</td></tr><tr><td><span class="icon">🔧</span>설계방법</td><td>{dm}</td></tr><tr><td><span class="icon">📖</span>설계기준</td><td>{code}</td></tr><tr><td><span class="icon">🏛️</span>기둥형식</td><td>{ct}</td></tr></table>''', unsafe_allow_html=True)
        
        with col4:
            dia, dc = getattr(In, 'dia', [22.0])[0], getattr(In, 'dc', [60.0])[0]
            rebar_count = f"{In.be / In.sb[0]:,.2f}"  # 1000/150 = 6.67 → 각 7개
            st.markdown(f'''<table class="common-table"><tr><td colspan="2" style="text-align:center;font-weight:bold;">🔩 철근 배치</td></tr><tr><td><span class="icon">⭕</span>철근 직경 D</td><td>{dia:,.1f} mm</td></tr><tr><td><span class="icon">🛡️</span>피복두께 d<sub>c</sub></td><td>{dc:,.1f} mm</td></tr><tr><td><span class="icon">📊</span>압축/인장측</td><td>각 {rebar_count}개</td></tr></table>''', unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)

    def create_report_column(column_ui, title, In, PM_obj, material_type):
        with column_ui:
            st.markdown('<div class="report-container">', unsafe_allow_html=True)
            st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)
            
            pm_data, balanced_data, material_props = extract_pm_data(PM_obj, material_type, In)
            
            st.markdown('<div class="sub-section-header">🔧 철근 재료 특성</div>', unsafe_allow_html=True)
            steel_type_note = " (중공철근 - 단면적 50%)" if material_type == '중공철근' else " (이형철근)"
            st.markdown(f'''<table class="param-table"><tr><td><span class="icon">💪</span>항복강도 f<sub>y</sub>{steel_type_note}</td><td>{material_props['fy']:,.1f} MPa</td></tr><tr><td><span class="icon">⚡</span>탄성계수 E<sub>s</sub></td><td>{material_props['Es']:,.1f} GPa</td></tr></table>''', unsafe_allow_html=True)
            
            st.markdown('<div class="sub-section-header">⚖️ 평형상태(Balanced) 검토</div>', unsafe_allow_html=True)
            st.markdown(f'''<table class="param-table"><tr><td><span class="icon">⚖️</span>축력 P<sub>b</sub></td><td>{balanced_data.get('Pb_kN', 0):,.1f} kN</td></tr><tr><td><span class="icon">📏</span>모멘트 M<sub>b</sub></td><td>{balanced_data.get('Mb_kNm', 0):,.1f} kN·m</td></tr><tr><td><span class="icon">📐</span>편심 e<sub>b</sub></td><td>{balanced_data.get('eb_mm', 0):,.1f} mm</td></tr><tr><td><span class="icon">🎯</span>중립축 깊이 c<sub>b</sub></td><td>{balanced_data.get('cb_mm', 0):,.1f} mm</td></tr></table>''', unsafe_allow_html=True)
            
            st.markdown('<div class="sub-section-header">📊 기둥강도 검토 결과 (요약)</div>', unsafe_allow_html=True)
            check_results = calculate_strength_check(In, material_type)
            
            if check_results:
                def render_html_table(results):
                    html = '''<table class="results-table"><tr><th>하중조합</th><th>P<sub>u</sub> / φP<sub>n</sub> [kN]</th><th>M<sub>u</sub> / φM<sub>n</sub> [kN·m]</th><th>편심 e [mm]</th><th>PM 교점 안전율</th><th>판정</th></tr>'''
                    all_passed = True
                    for r in results:
                        vc = "pass" if r['Verdict'] == 'PASS' else "fail"
                        if vc == "fail": all_passed = False
                        html += f'''<tr><td><b>{r['LC']}</b></td><td>{r['Pu/phiPn']}</td><td>{r['Mu/phiMn']}</td><td>{r['e_mm']:.1f}</td><td>{r['SF']:.3f}</td><td class="{vc}">{r['Verdict']} {'✅' if vc == 'pass' else '❌'}</td></tr>'''
                    return html + '</table>', all_passed

                html_table, all_passed = render_html_table(check_results)
                st.markdown(html_table, unsafe_allow_html=True)

                st.markdown('<div class="sub-section-header">🔍 상세 강도 검토 (모든 하중조합)</div>', unsafe_allow_html=True)
                
                num_cases = len(getattr(In, 'Pu', []))
                st.info(f"##### 📋 총 {num_cases}개 하중조합에 대해 상세 검토를 수행합니다.")
                
                for case_idx in range(num_cases):
                    detailed_html = render_detailed_strength_check(In, PM_obj, material_type, case_idx)
                    st.markdown(detailed_html, unsafe_allow_html=True)

                st.markdown('<div class="sub-section-header">🎯 최종 종합 판정</div>', unsafe_allow_html=True)
                if all_passed:
                    st.markdown('<div class="final-verdict-container final-pass">🎉 전체 조건 만족 - 구조 안전</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="final-verdict-container final-fail">⚠️ 일부 조건 불만족 - 보강 검토 필요</div>', unsafe_allow_html=True)
            else:
                st.warning("⚠️ 검토 데이터를 계산할 수 없습니다.", icon="⚠️")
            
            st.markdown('</div>', unsafe_allow_html=True)

    # =================================================================
    # 메인 렌더링
    # =================================================================
    try:
        st.markdown('<div class="main-container">', unsafe_allow_html=True)
        st.markdown('<div class="main-header">🏗️ 기둥 강도 검토 보고서</div>', unsafe_allow_html=True)
        render_common_conditions(In)
        col1, col2 = st.columns(2, gap="large")
        create_report_column(col1, "📊 이형철근 검토", In, R, "이형철근")
        create_report_column(col2, "📊 중공철근 검토", In, F, "중공철근")
        st.markdown('</div>', unsafe_allow_html=True)
        
    except Exception as e:
        st.error(f"⚠️ 보고서 생성 중 오류 발생: {e}")
        import traceback
        with st.expander("🔍 디버깅 정보 보기"):
            st.write(f"Error Type: {type(e).__name__}")
            st.write(f"Error Details: {e}")
            st.code(traceback.format_exc())
