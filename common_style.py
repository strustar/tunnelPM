import streamlit as st

hover_bc = 'lightblue'
bc = 'linen'
basic_style = f'font-weight: bold; font-size: 18px; background-color: {bc};'
border_style = 'border: 1px solid black; border-radius: 30px;'


# 메인바 윗쪽 여백 줄이기 & 텍스트, 숫자 상자, 사이드바, 선택상자, 라디오 버튼 스타일  !! 인쇄 설정
def input_box(In):
    input_style = f""" <style>
        /* 메인 페이지의 상단 패딩 조정 */
        .block-container{{padding-top: 0.2rem;}}

        .block-container {{
            margin-top: 20px;                        
            max-width: {In.max_width};
        }}
        .element-container {{
            white-space: nowrap;
            overflow-x: visible;
        }}

        /* CSS to set font for everything except code blocks */
        body, h1, h2, h3, h4, h5, h6, p, blockquote {{
            font-family: 'Nanum Gothic', sans-serif;
            font-weight: bold;
        }}
        /* Font size for titles (h1 to h6) */
        h1 {{font-size: {In.font_h1} !important;}}
        h2 {{font-size: {In.font_h2} !important;}}
        h3 {{font-size: {In.font_h3} !important;}}
        h4 {{font-size: {In.font_h4} !important;}}
        h5 {{font-size: {In.font_h5} !important;}}
        h6 {{font-size: {In.font_h6} !important;}}

        [data-testid=stSidebar] {{
            padding: 5px;
            margin-top: -80px !important;
            background-color: azure;
            border: 3px dashed purple;
            height: 110% !important;
            max-width: 600px !important;  /* 사이드바의 최대 크기를 조절합니다 */
            width: 100% !important;  /* 이렇게 하면 사이드 바 폭을 고정할수 있음. */
        }}
        input[type="number"] {{
            margin-top: 5px;
            margin-bottom: 5px;
            padding: 5px;
            padding-left: 15px;            
            {basic_style}
            {border_style}
        }}
        div.row-widget.stTextInput > div > div > input {{
            padding: 6px;
            padding-left: 12px;            
            {basic_style}
            {border_style}
        }}
        div.row-widget.stTextInput > div > div > input::placeholder {{
            padding-left: 5px;
            {basic_style}
            # color: purple;
        }}

        div.row-widget.stSelectbox > div > div {{
            padding-left: 5px;
            {basic_style}
            {border_style}
        }}

        /* 라디오 버튼 */
        # div.row-widget.stRadio > div[role='radiogroup'] > label {{
        #     padding: 3px;
        #     padding-left: 10px;
        # }}
        div.row-widget.stRadio > div {{
            display: flex;
            justify-content: space-between;
            # flex-direction: row;
            margin-top: -5px;
            padding: 7px;
            padding-left: 15px;
            # padding-right: 15px;
            {basic_style}
            {border_style}            
        }}

        /* 호버 스타일 : 선택상자, 텍스트, 라디오, 숫자 버튼 */
        div.row-widget.stSelectbox > div[data-baseweb='select'] > div:hover,
        div.row-widget.stTextInput > div > div > input:hover,        
        div.row-widget.stRadio > div[role='radiogroup'] > label:hover,
        input[type="number"]:hover {{
            {basic_style}
            background-color: {hover_bc};            
        }}
    </style> """
    st.markdown(input_style, unsafe_allow_html=True)

    # 모든 글씨 및 라텍스 수식 진하게 설정
    st.markdown(
        """<style>
        .main * {
            font-weight: bold !important;
            # font-family: Nanum Gothic !important;
        }
    </style>""",
        unsafe_allow_html=True,
    )

    # 왼쪽 사이드바 인쇄하지 않기 설정 등, # 페이지 간 구분을 위한 CSS 스타일 정의
    st.markdown(
        """
    <style>
        @media print {
            [data-testid=stSidebar], header, footer, .no-print {
                display: none;
            }
            @page {
                size: A4;
                margin-left: 50px; /* 왼쪽 여백 설정 */
            }
            body {
                width: 100%; /* 전체 너비 사용 */
            }
        }
        /* 페이지 브레이크 스타일 */
        .page-break {
            page-break-before: always;
        }
    </style> """,
        unsafe_allow_html=True,
    )


#! 워터마크 이미지, 투명도 및 크기를 설정하는 함수
# @st.cache_resource
def watermark(In):
    if '0811' not in In.watermark:
        st.markdown(
            """
            <style>
                .watermark-container {
                    position: fixed;                    
                    top: 50%;
                    left: 50%;
                    transform: translate(-17%, -50%);
                    z-index: 9999;  /* 다른 요소들보다 상위에 위치하도록 설정 */
                    pointer-events: none;  /* 워터마크가 클릭 이벤트를 방해하지 않도록 설정 */
                }
                .watermark-container img {
                    max-width: 100%;  /* 이미지의 최대 폭을 화면 너비에 맞춤 */
                    height: auto;  /* 이미지 높이를 자동으로 조절하여 비율을 유지 */
                    opacity: 0.15;  /* 이미지에만 투명도 적용 */
                }
                    
                @media print {
                    .watermark-container {
                        width: 100%; /* 인쇄 시 워터마크 이미지의 폭 */
                        top: 50%;
                        left: 50%;
                        transform: translate(-50%, -50%);                        
                        position: fixed; /* 인쇄 시 위치 고정 */
                    }
                    .watermark-container img {
                        max-width: 100%;  /* 인쇄 시 이미지 폭을 부모 요소의 폭에 맞춤 */
                        opacity: 0.25;  /* 인쇄 시 이미지에만 투명도 조정 */
                    }
                }
            </style>
            <div class="watermark-container">
                <img src="https://github.com/strustar/Support/blob/main/Images/watermark.png?raw=true">                
            </div>
        """,
            unsafe_allow_html=True,
        )
