import streamlit as st

hover_bc = 'lightblue'
bc = 'linen'
basic_style = f'font-weight: bold; font-size: 18px; background-color: {bc};'
border_style = 'border: 1px solid black; border-radius: 30px;'


# 메인바 윗쪽 여백 줄이기 & 텍스트, 숫자 상자, 사이드바, 선택상자, 라디오 버튼 스타일  !! 인쇄 설정
def input_box(In):
    input_style = f""" <style>
        /* =============================================================
            메인 컨테이너 및 레이아웃 설정
        ============================================================= */
        .block-container {{
            padding-top: 0.8rem;
            margin-top: 20px;                        
            max-width: {In.max_width};
        }}
        
        .element-container {{
            white-space: nowrap;
            overflow-x: visible;
        }}

        /* =============================================================
            폰트 및 텍스트 설정
        ============================================================= */
        body, h1, h2, h3, h4, h5, h6, p, blockquote {{
            font-family: 'Nanum Gothic', sans-serif;
            font-weight: bold;
        }}
        
        h1 {{ font-size: {In.font_h1} !important; }}
        h2 {{ font-size: {In.font_h2} !important; }}
        h3 {{ font-size: {In.font_h3} !important; }}
        h4 {{ font-size: {In.font_h4} !important; }}
        h5 {{ font-size: {In.font_h5} !important; }}
        h6 {{ font-size: {In.font_h6} !important; }}

        /* =============================================================
            사이드바 스타일
        ============================================================= */
        [data-testid="stSidebar"] {{
            padding: 5px;
            margin-top: -40px !important;
            background-color: azure;
            border: 3px dashed purple;
            height: 110% !important;
            max-width: 600px;
        }}

        /* =============================================================
            입력 위젯 공통 스타일
        ============================================================= */
        
        /* 숫자 입력 */
        .stNumberInput input {{
            margin-top: 5px;
            margin-bottom: 5px;
            padding: 5px;
            padding-left: 15px;
            text-align: left;
            {basic_style}
            {border_style}
        }}
        
        /* +/- 버튼 컨테이너 크기 조정 */
        .stNumberInput button[data-testid="stNumberInputStepUp"],
        .stNumberInput button[data-testid="stNumberInputStepDown"] {{
        visibility: visible !important;
            width: 30px !important;
            height: 20px !important;
            min-width: 20px !important;
            min-height: 20px !important;
            padding: 0 !important;
            margin: 0px !important;
            display: flex !important;
                    align-items: center !important;
                    justify-content: center !important;
                }}

        # /* 버튼 내부 SVG 아이콘 중앙 정렬 */
        # .stNumberInput button[data-testid="stNumberInputStepUp"] svg,
        # .stNumberInput button[data-testid="stNumberInputStepDown"] svg {{
        #     margin: auto !important;
        # }}

        /* 버튼 컨테이너 자체도 중앙 정렬 */
        .stNumberInput > div > div:last-child {{
            display: flex !important;
            flex-direction: column !important;
            justify-content: center !important;
            align-items: center !important;
        }}

        /* 텍스트 입력 */
        .stTextInput input {{
            padding: 6px;
            padding-left: 12px;
            text-align: left;
            {basic_style}
            {border_style}
        }}
        
        /* 텍스트 입력 플레이스홀더 */
        .stTextInput input::placeholder {{
            padding-left: 5px;
            {basic_style}
        }}
        
        /* 선택 상자 */
        .stSelectbox > div > div {{
            padding-left: 5px;
            text-align: left;
            {basic_style}
            {border_style}
        }}
        
        /* 라디오 버튼 */
        .stRadio > div {{
            display: flex;
            justify-content: space-between;
            margin-top: -5px;
            padding: 7px;
            padding-left: 15px;
            {basic_style}
            {border_style}
        }}

        /* =============================================================
            호버 효과 (마우스 올렸을 때)
        ============================================================= */
        .stNumberInput input:hover,
        .stTextInput input:hover,
        .stSelectbox > div > div:hover,
        .stRadio > div[role='radiogroup'] > label:hover {{
            {basic_style}
            background-color: {hover_bc};
        }}
        
        /* =============================================================
            인쇄 설정
        ============================================================= */
        @media print {{
            [data-testid="stSidebar"], 
            header, 
            footer, 
            .no-print {{
                display: none;
            }}
            @page {{
                size: A4;
                margin-left: 50px;
            }}
            body {{
                width: 100%;
            }}
        }}
        
        /* 페이지 브레이크 */
        .page-break {{
            page-break-before: always;
        }}
    </style> """
    
    st.markdown(input_style, unsafe_allow_html=True)
    # 탭 스타일
    st.markdown("""
<style>
.stTabs [data-baseweb="tab-list"] {
    gap: 24px;
    justify-content: center;
    background: rgba(186, 186, 59, 0.5);
    padding: 12px;
    border-radius: 20px;
    border: 1px solid #dee2e6;
    box-shadow: inset 0 2px 4px rgba(0,0,0,0.06);
    margin: 10px 0;
}

.stTabs [data-baseweb="tab"] {
    height: 55px !important;
    background: linear-gradient(145deg, #ffffff, #f8f9fa) !important;
    border-radius: 16px !important;
    border: 2px solid #2196f3 !important;
    color: #495057 !important;
    font-size: 18px !important;
    font-weight: 600 !important;
    padding: 0 24px !important;
    margin: 0 2px !important;
    box-shadow: 0 3px 6px rgba(0,0,0,0.08), inset 0 1px 0 rgba(255,255,255,0.8) !important;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    position: relative !important;
}

.stTabs [data-baseweb="tab"]:hover {
    background: #bdc51b !important;
    border: 2px solid #2196f3 !important;
    color: #1565c0 !important;
    # transform: translateY(-3px) scale(1.02) !important;
    box-shadow: 0 6px 20px rgba(33, 150, 243, 0.25), inset 0 1px 0 rgba(255,255,255,0.9) !important;
    font-weight: 700 !important;
}

.stTabs [aria-selected="true"] {
    background: linear-gradient(145deg, #4caf50, #66bb6a) !important;
    color: black !important;
    border: 2px solid #4caf50 !important;
    font-size: 18px !important;
    font-weight: 700 !important;
    transform: translateY(-4px) scale(1.05) !important;
    box-shadow: 0 8px 25px rgba(76, 175, 80, 0.4), inset 0 1px 0 rgba(255,255,255,0.2) !important;
}

.stTabs [data-baseweb="tab"] p {
    font-size: 18px !important;
    margin: 0 !important;
    text-shadow: 0 1px 2px rgba(0,0,0,0.1) !important;
}
</style>
""", unsafe_allow_html=True)
