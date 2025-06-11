import streamlit as st

def apply_css_styles():
    """공통 CSS 스타일을 적용합니다."""
    st.markdown("""
        <style>
        /* 전역 설정 */
        .main-container {
            padding: 1rem;
            max-width: 1200px;
            margin: 0 auto;
        }
        
        /* 메인 타이틀 */
        .title-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 2rem;
            border-radius: 15px;
            text-align: center;
            margin-bottom: 2rem;
            color: white;
        }
        
        .title-header h1 {
            color: white !important;
            margin: 0;
            font-size: 2.2rem;
            font-weight: 700;
        }
        
        .title-header p {
            color: rgba(255,255,255,0.9) !important;
            margin: 0.5rem 0 0 0;
            font-size: 1.5rem;
        }
        
        /* 섹션 박스 */
        .intro-box {
            background: #f8f9fa;
            padding: 1.5rem;
            border-radius: 10px;
            border-left: 4px solid #667eea;
            margin-bottom: 1rem;
        }
        
        .assumption-box {
            background: #f8f9fa;
            padding: 1.5rem;
            border-radius: 10px;
            border-left: 4px solid #4facfe;
            margin-bottom: 1rem;
        }
        
        /* 케이스 박스 */
        .case-special-box {
            background: linear-gradient(135deg, #2E7D32 0%, #1B5E20 100%);
            padding: 1.5rem;
            border-radius: 55px;
            color: white;
            margin-bottom: 1rem;
        }

        .case-general-box {
            background: linear-gradient(135deg, #1565C0 0%, #0D47A1 100%);
            padding: 1.5rem;
            border-radius: 55px;
            color: white;
            margin-bottom: 1rem;
        }
        
        /* 균열 제어 박스 */
        .crack-box {
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
            padding: 1.5rem;
            border-radius: 10px;
            color: white;
            margin-bottom: 1rem;
        }
        
        /* 텍스트 색상 강제 적용 */
        .case-special-box h3, .case-special-box h4, .case-special-box p {
            color: white !important;
        }
        
        .case-general-box h3, .case-general-box h4, .case-general-box p {
            color: white !important;
        }
        
        .crack-box h3, .crack-box p {
            color: white !important;
        }
        
        /* 수식 배경 */
        .math-bg {
            background: rgba(255,255,255,0.15);
            padding: 0.5rem;
            border-radius: 5px;
            margin: 0.5rem 0;
        }
        
        /* 결과 박스 */
        .result-success {
            background: #28a745;
            color: white;
            padding: 1rem;
            border-radius: 8px;
            text-align: center;
            margin: 0.5rem 0;
        }
        
        .result-error {
            background: #dc3545;
            color: white;
            padding: 1rem;
            border-radius: 8px;
            text-align: center;
            margin: 0.5rem 0;
        }
                
        </style>
    """, unsafe_allow_html=True)

def display_header():
    """메인 타이틀을 표시합니다."""
    st.markdown("""
    <div class="title-header">
        <h1>🏗️ RC 사용성 검토</h1>
        <p>응력 및 균열 제어 완전 가이드</p>
    </div>
    """, unsafe_allow_html=True)

def display_theory_background():
    """이론적 배경을 표시합니다."""
    col1, col2 = st.columns([1, 1], gap="large")
    
    with col1:
        st.markdown("""
        <div class="intro-box">
            <h3>📋 I. 개요</h3>
            <p><strong>🎯 강도 설계:</strong> 계수하중에 대한 구조물의 <strong>파괴 방지</strong> 목적</p>
            <p><strong>🔍 사용성 검토:</strong> 사용하중에 대한 <strong>기능 및 내구성 확보</strong> 목적 (균열, 처짐 제어)</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="assumption-box">
            <h3>⚙️ II. 해석의 기본 가정 (탄성 이론)</h3>
            <p><strong>1.</strong> <strong>평면 유지 가정:</strong> 변형률(ε)은 중립축에서 선형 분포</p>
            <p><strong>2.</strong> <strong>선형 탄성 거동:</strong> 응력-변형률 관계는 선형 (f = E·ε)</p>
            <p><strong>3.</strong> <strong>콘크리트 인장 무시:</strong> 인장력은 전량 철근이 부담</p>
        </div>
        """, unsafe_allow_html=True)

def display_basic_theory():
    """기본 이론만 표시하는 함수 (단독 사용 가능)"""
    apply_css_styles()
    display_header()
    display_theory_background()

    # --- Part 1 타이틀 ---
    st.markdown("""
    <div style="display: inline-block; background:linear-gradient(135deg, #FF8C00, #FFA500); padding: 1rem; border-radius: 10px; margin-bottom: 1rem;">
        <h2 style="color:black; text-align: left; margin:0;">Part 1. 탄성 이론 기반 응력 해석 비교</h2>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("")

    # --- 응력 해석 비교 ---
    col_left, col_right = st.columns(2, gap="large")

    # --- 특수한 경우 ---
    with col_left:
        st.markdown("""
        <div class="case-special-box">
            <h3 style="text-align: center; margin-top: 0;">🎯 Case Ⅰ: 특수한 경우</h3>
            <h4 style="text-align: center;">순수 휨 (축력 P<sub>0</sub> = 0)</h4>
            <p style="text-align: center;">📊 보(Beam)에 해당</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("**🔹 1. 핵심 원리: 힘의 평형**")
        st.markdown("- 외부 축력 부재로, 내부 압축력(C)과 인장력(T)은 동일")
        
        st.latex("C = T")

        st.markdown("**🔹 2. 중립축($x$) 계산: `해석적 풀이`**")
        st.markdown("- 환산단면의 단면 1차 모멘트 평형식으로부터 $x$에 대한 **2차 방정식** 유도")
        
        st.latex(r"\frac{1}{2} b x^2 = n A_s (d-x)")
        
        st.markdown("- 근의 공식을 통해 $x$를 직접 계산 가능")
        
        st.latex(r"x = kd \quad \text{where} \quad k = \sqrt{(n\rho)^2 + 2n\rho} - n\rho")
        
        st.markdown("**🔹 3. 응력 계산**")
        st.markdown("- 계산된 $x$를 이용하여 응력 직접 산출")
        
        st.latex(r"f_s = \frac{M_o}{A_s (d - x/3)}")
        
        st.success("✅ **특징**: 2차방정식으로 직접 해결 가능")

    # --- 일반적인 경우 ---
    with col_right:
        st.markdown("""
        <div class="case-general-box">
            <h3 style="text-align: center; margin-top: 0;">⚙️ Case Ⅱ: 일반적인 경우</h3>
            <h4 style="text-align: center;">축력+휨 (축력 P<sub>0</sub> ≠ 0)</h4>
            <p style="text-align: center;">🏛️ 기둥(Column)에 해당</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("**🔹 1. 핵심 원리: 2개의 평형 방정식**")
        st.markdown("- 축력과 모멘트 평형을 **동시에** 만족시켜야 함")
        
        st.latex(r"\text{축력: } P_0 = C - T")
        st.latex(r"\text{모멘트: } M_0 = C(\frac{h}{2}-\frac{x}{3}) + T(d-\frac{h}{2})")

        st.markdown("**🔹 2. 중립축($x$) 계산: `수치 해석`**")
        st.markdown("- 미지수 2개($x$, $\epsilon_c$)를 갖는 **비선형 연립방정식** 문제")
        st.markdown("- 직접 풀이 불가. `fsolve`와 같은 컴퓨터 Solver 필수")
        
        st.markdown("**🔹 3. 응력 계산**")
        st.markdown("- Solver로 찾은 해($x$, $\epsilon_c$)로부터 응력 산출")
        
        st.latex(r"\epsilon_s = \epsilon_c \frac{d-x}{x} \quad \implies \quad f_s = E_s \epsilon_s")
        
        st.warning("⚠️ **특징**: 수치해석 반복 계산 필수")

    # --- Part 2: 균열 제어 ---
    st.markdown("---")
    # st.markdown("")
    st.markdown("""
    <div style="display: inline-block; background:linear-gradient(135deg, #FF8C00, #FFA500); padding: 1rem; border-radius: 10px; margin-bottom: 1rem;">
        <h2 style="color:black; text-align: left; margin:0;">Part 2. 휨균열 제어를 위한 철근 간격 검토</h2>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("")

    st.markdown("### 🔬 2.1 최외단 철근 응력 $f_{st}$ 산정")
    st.markdown("- 균열폭에 가장 지배적인 최외단 철근의 응력을 계산")

    st.latex(r"f_{st} = f_s \cdot \frac{h - d_c - x}{d - x}")

    st.markdown("### 📏 2.2 최대 허용 간격 (s) 산정 [KDS 기준]")
    st.markdown("- 아래 두 식으로 계산된 값 중 **작은 값**을 허용 간격으로 적용")
    
    col1, col2 = st.columns(2)
    with col1:
        st.latex(r"s \leq 375 \left( \frac{210}{f_{st}} \right) - 2.5 C_c")
    with col2:
        st.latex(r"s \leq 300 \left( \frac{210}{f_{st}} \right)")

    st.markdown("### 🎯 2.3 판정")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        <div class="result-success">
            ✅ <strong>배근 간격 ≤ 허용 간격 (s) → O.K.</strong>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="result-error">
            ❌ <strong>배근 간격 > 허용 간격 (s) → N.G. (설계 변경 필요)</strong>
        </div>
        """, unsafe_allow_html=True)



def render_case_analysis(data, In, i, num_symbols):
    """단일 케이스 해석을 수행하고 결과를 현재 컬럼에 표시"""
    
    fs_case, x_case = data.fs[i], data.x[i]
    P0_case, M0_case = In.P0[i], In.M0[i]

    st.markdown(f"# **{num_symbols[i]}번 검토**")

    # --- 케이스 분류 및 해석 방법 선택 ---
    if P0_case == 0:
        # 특수한 경우: 순수 휨
        st.markdown(f"""
        <div class="case-special-box">
            <h3 style="text-align: center; margin-top: 0;">🎯 Case Ⅰ: 특수한 경우</h3>
            <h4 style="text-align: center;">순수 휨 (P₀ = {P0_case:,.1f} kN, M₀ = {M0_case:,.1f} kN·m)</h4>
            <p style="text-align: center;">📊 보(Beam)에 해당 - 해석적 풀이 적용</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("### 🔬 **A. 탄성 해석 과정 (이론적 접근)**")
        
        st.markdown("#### **Step 1: 힘의 평형 조건**")
        st.markdown("외부 축력이 없으므로, 내부 압축력(C)과 인장력(T)이 평형:")
        st.latex("C = T")
        
        st.markdown("#### **Step 2: 중립축 위치 계산 (해석적 풀이)**")
        st.markdown("환산단면의 단면 1차 모멘트 평형으로부터 2차 방정식 유도:")
        st.latex(r"\frac{1}{2} b x^2 = n A_s (d-x)")
        st.latex(r"x = kd \quad \text{where} \quad k = \sqrt{(n\rho)^2 + 2n\rho} - n\rho")
        st.latex(f"x = {x_case:.1f} \\text{{ mm}} \\quad (해석적 해)")
        
        st.markdown("#### **Step 3: 철근 응력 계산**")
        st.markdown("계산된 중립축을 이용하여 직접 응력 산출:")
        st.latex(r"f_s = \frac{M_0}{A_s (d - x/3)}")
        st.latex(f"f_s = {fs_case:.1f} \\text{{ MPa}}")

    else:
        # 일반적인 경우: 축력+휨
        st.markdown(f"""
        <div class="case-general-box">
            <h3 style="text-align: center; margin-top: 0;">⚙️ Case Ⅱ: 일반적인 경우</h3>
            <h4 style="text-align: center;">축력+휨 (P₀ = {P0_case:,.1f} kN, M₀ = {M0_case:,.1f} kN·m)</h4>
            <p style="text-align: center;">🏛️ 기둥(Column)에 해당 - 수치해석 필요</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("### 🔬 **A. 탄성 해석 과정 (수치해석 접근)**")
        
        st.markdown("#### **Step 1: 연립 평형방정식 설정**")
        st.markdown("축력과 모멘트 평형을 동시에 만족하는 미지수 $(x, \\epsilon_c)$ 계산:")
        
        st.markdown("**축력 평형:**")
        st.latex(r"P_0 = C - T = \frac{1}{2} f_c b x - A_s f_s")
        
        st.markdown("**모멘트 평형:**")
        st.latex(r"M_0 = C(\frac{h}{2}-\frac{x}{3}) + T(d-\frac{h}{2})")
        
        st.markdown("#### **Step 2: 수치해석 결과**")
        st.markdown("비선형 연립방정식을 수치적 방법(fsolve 등)으로 해결:")
        st.latex(f"x = {x_case:.1f} \\text{{ mm}} \\quad (수치해)")
        st.latex(f"f_s = {fs_case:.1f} \\text{{ MPa}} \\quad (수치해)")
        
        st.warning("⚠️ **주의**: 직접 풀이 불가능하여 반복계산 통해 해 도출")

    st.markdown("---")
    
    # --- 균열 검토 (공통) ---
    st.markdown("### 📏 **B. 휨균열 제어 검토**")
    
    # 최외단 철근 응력 계산
    st.markdown("#### **Step 1: 최외단 철근 응력 $f_{st}$ 산정**")
    fst_case = fs_case  # 1단 배근 가정으로 fs ≈ fst
    st.latex(r"f_{st} = f_s \cdot \frac{h - d_c - x}{d - x} \approx f_s")
    st.latex(f"f_{{st}} = {fst_case:.1f} \\text{{ MPa}}")
    
    # 허용 간격 계산
    st.markdown("#### **Step 2: 최대 허용 간격 산정 [KDS 기준]**")
    
    # 계산식 1
    s_allowed_1 = 375 * (210 / fst_case) - 2.5 * In.Cc if fst_case > 0 else float('inf')
    st.markdown("**조건 1:**")
    st.latex(r"s_1 = 375 \left( \frac{210}{f_{st}} \right) - 2.5 C_c")
    st.latex(f"s_1 = 375 \\left( \\frac{{210}}{{{fst_case:.1f}}}\\right) - 2.5 \\times {In.Cc:.1f} = {s_allowed_1:.1f} \\text{{ mm}}")
    
    # 계산식 2
    s_allowed_2 = 300 * (210 / fst_case) if fst_case > 0 else float('inf')
    st.markdown("**조건 2:**")
    st.latex(r"s_2 = 300 \left( \frac{210}{f_{st}} \right)")
    st.latex(f"s_2 = 300 \\left( \\frac{{210}}{{{fst_case:.1f}}} \\right) = {s_allowed_2:.1f} \\text{{ mm}}")
    
    # 최종 허용 간격
    s_allowed_final = min(s_allowed_1, s_allowed_2)
    st.latex(f"s_{{allow}} = \\min(s_1, s_2) = {s_allowed_final:.1f} \\text{{ mm}}")

    # --- 최종 판정 ---
    st.markdown("#### **Step 3: 최종 판정**")
    
    result_col1, result_col2 = st.columns(2)
    result_col1.metric("최종 허용 간격", f"{s_allowed_final:.1f} mm", delta="Min(s1, s2)", delta_color="off")
    result_col2.metric("실제 배근 간격", f"{In.sb[0]:.1f} mm")

    if In.sb[0] <= s_allowed_final:
        is_safe = True
        st.markdown(f"""
        <div class="result-success">
            ✅ <strong>O.K. (배근 간격 {In.sb[0]:.1f} mm ≤ 허용 간격 {s_allowed_final:.1f} mm)</strong>
        </div>
        """, unsafe_allow_html=True)
    else:
        is_safe = False
        st.markdown(f"""
        <div class="result-error">
            ❌ <strong>N.G. (배근 간격 {In.sb[0]:.1f} mm > 허용 간격 {s_allowed_final:.1f} mm)</strong>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    return is_safe


def serviceability_check_results(In, R, F):
    """
    RC 사용성 검토 - 이론 설명과 실제 계산을 통합하여 표시
    """
    
    # --- Part 3: 실제 계산 결과 ---
    st.markdown("---")
    st.markdown("""
    <div style="display: inline-block; background:linear-gradient(135deg, #FF8C00, #FFA500); padding: 1rem; border-radius: 10px; margin-bottom: 1rem;">
        <h2 style="color:black; text-align: left; margin:0;">Part 3. 하중 케이스별 상세 해석 및 균열 검토</h2>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("")
    
    num_symbols = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"]

    # 좌우 컬럼 생성
    col_R, col_F = st.columns(2, gap="large")
    
    # 왼쪽 컬럼: R 데이터
    is_safe_R = []
    with col_R:
        for i in range(len(In.P0)):
            is_safe_R.append(render_case_analysis(R, In, i, num_symbols))

    # 오른쪽 컬럼: F 데이터
    is_safe_F = []
    with col_F:
        for i in range(len(In.P0)):
            is_safe_F.append(render_case_analysis(F, In, i, num_symbols))

    for i in range(len(In.Vu)):
        # 1) R 판정 텍스트
        r_status = ":green[OK]" if is_safe_R[i] else ":red[NG]"
        # 2) F 판정 텍스트
        f_status = ":green[OK]" if is_safe_F[i] else ":red[NG]"
        # 3) 슬래시 구분하여 출력
        In.placeholder_serviceability[i].write(f"{r_status} / {f_status}")

