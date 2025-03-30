import streamlit as st
import numpy as np
import pandas as pd
import Column_Sidebar, Column_Calculate, Column_Result

import os, sys, importlib
os.system('cls')  # 터미널 창 청소, clear screen 
sys.path.append( "D:\\Work_Python\\Common")  # 공통 스타일 변수 디렉토리 추가
import commonStyle    # print(sys.path)
importlib.reload(commonStyle) # 다른 폴더 자동 변경이 안됨? ㅠ

### * -- Set page config
st.set_page_config(page_title = "Paper 2024", page_icon = "column.png", layout = "centered",    # centered, wide
                    initial_sidebar_state="expanded",  # runOnSave = True,
                    menu_items = {
                        # 'Get Help': 'https://www.extremelycoolapp.com/help',
                        # 'Report a bug': "https://www.extremelycoolapp.com/bug",
                        # 'About': "# This is a header. This is an *extremely* cool app!"
                    })

st.session_state.fck = 40
rho = 0.015;  b = 400;  h = 400
d = np.sqrt(4*rho*b*h/(8*np.pi))
st.session_state.dia1 = d

ep_fu = 0.01
Ef = 100;  f_fu = ep_fu*Ef*1e3;  #ep_fu = f_fu/Ef/1e3
'ep_fu : '; ep_fu
st.session_state.f_fu = f_fu
st.session_state.Ef = Ef

In = Column_Sidebar.Sidebar()
commonStyle.input_box(In)
commonStyle.watermark(In)

R = Column_Calculate.Cal(In, 'RC')
F = Column_Calculate.Cal(In, 'FRP')

# ===========================================================
areas_pos = []
areas_neg = []
areas_total = []
for i in range(2):
    if i == 0:
        x = R.ZMd
        y = R.ZPd
        curve_name = "RC"
    else:
        x = F.ZMd
        y = F.ZPd
        curve_name = "FRP"

    # 양수 부분 계산
    positive_indices = y > 0
    x_pos = x[positive_indices]
    y_pos = y[positive_indices]
    area_pos = np.trapz(y_pos, x_pos)
    
    # 음수 부분 계산
    negative_indices = y < 0
    x_neg = x[negative_indices]
    y_neg = y[negative_indices]
    area_neg = np.trapz(y_neg, x_neg)
    
    # 전체 면적
    total_area = area_pos + area_neg
    
    areas_pos.append(area_pos)
    areas_neg.append(area_neg)
    areas_total.append(total_area)

# 결과를 데이터프레임으로 정리
if len(areas_total) == 2:
    Balance_loc_RC = '압축측' if R.Pd[3] > 0 else '인장측'
    Balance_loc_FRP = '압축측' if F.Pd[3] > 0 else '인장측'
    data = {
        '구분': ['압축측 면적', '인장측 면적', '전체 면적', 'Balance Point (Pd)', 'Balance Point (Md)', 'Balance Point (Location)', 'c/d', 'e/(h/2)'],
        'RC': [areas_pos[0], areas_neg[0], areas_total[0], R.Pd[3], R.Md[3], Balance_loc_RC, R.c[3]/(In.height-In.dc[0]), R.e[3]/(In.height/2)],
        'FRP': [areas_pos[1], areas_neg[1], areas_total[1], F.Pd[3], F.Md[3], Balance_loc_FRP, F.c[3]/(In.height-In.dc[0]), F.e[3]/(In.height/2)],
        # 'Pd': [R.Pd[3], F.Pd[3], 'R.Pd[3]+F.Pd[3]'],
    }
    df = pd.DataFrame(data)

    # 비교 (%) 열 추가
    df['비교 (%)'] = 0.0  # 초기값으로 0.0을 설정
    # 퍼센트 차이 계산
    df['비교 (%)'][:5] = (df['FRP'][:5] - df['RC'][:5]) / df['RC'][:5] * 100
    df['비교 (%)'][5] = f'ep_fu = {ep_fu}'
    df['비교 (%)'][6] = f'f_fu = {f_fu} MPa'
    df['비교 (%)'][7] = f'Ef = {Ef} GPa'
    # # 어느 쪽이 큰지 표시
    # df['비교'] = df.apply(lambda row: 'FRP' if row['FRP'] > row['RC'] else ('RC' if row['RC'] > row['FRP'] else '동일'), axis=1)
    
    # 결과 출력
    # st.write("면적 비교 결과:")
    # edited_df = st.data_editor(df, num_rows="dynamic")
    edited_df = st.data_editor(df.transpose(), num_rows="dynamic")


Column_Result.Fig(In, R, F)




# ========================================================
# import streamlit as st
# import requests
# import fitz  # PyMuPDF
# import io

# # 구글 드라이브 파일 링크
file_link = "https://drive.google.com/file/d/1BrWtUEHPNzupNnfyTv9sUHryBbkpXEsn/view?usp=sharing"

# # 파일 ID 추출
# file_id = file_link.split('/d/')[1].split('/')[0]
# file_link
# # 메타데이터 URL 생성
# metadata_url = f"https://drive.google.com/file/d/{file_id}/view"

# metadata_url
# # 직접 다운로드 링크 생성
# download_url = f"https://drive.google.com/uc?export=download&id={file_id}"

# # PDF 파일 다운로드 함수 (캐싱 적용)
# @st.cache_data
# def download_pdf(url):
#     response = requests.get(url)
#     if response.status_code == 200:
#         return response.content
#     else:
#         return None

# # PDF 다운로드
# pdf_content = download_pdf(download_url)

# if pdf_content:
#     # PDF를 메모리에서 열기
#     pdf_stream = io.BytesIO(pdf_content)
#     doc = fitz.open(stream=pdf_stream, filetype='pdf')
    
#     # 첫 페이지 가져오기
#     page = doc.load_page(0)  # 페이지 번호는 0부터 시작
    
#     # 페이지를 Pixmap으로 렌더링
#     pix = page.get_pixmap()
    
#     # 이미지를 바이트로 변환
#     img_bytes = pix.tobytes()
    
#     # 이미지 표시
#     st.image(img_bytes)
# else:
#     st.error("PDF 파일을 다운로드하지 못했습니다.")

# import re

# def get_filename_from_drive_link(file_link):
#     # 파일 ID 추출
#     file_id = file_link.split('/d/')[1].split('/')[0]
    
#     # 메타데이터 URL 생성
#     metadata_url = f"https://drive.google.com/file/d/{file_id}/view"
    
#     try:
#         # 메타데이터 페이지 요청
#         response = requests.get(metadata_url)
#         response.raise_for_status()  # 오류 발생 시 예외 발생
        
#         # 정규 표현식을 사용하여 파일 이름 추출
#         match = re.search(r'<title>(.*?) - Google Drive</title>', response.text)
#         if match:
#             filename = match.group(1)
#             return filename
#         else:
#             return "파일 이름을 찾을 수 없습니다."
#     except requests.RequestException as e:
#         return f"오류 발생: {str(e)}"

# # 구글 드라이브 파일 링크
# file_link = "https://drive.google.com/file/d/1B1cQWgBh1eQukcH58jLD2Jr50yY4RanT/view?usp=sharing"

# # 파일 이름 추출
# filename = get_filename_from_drive_link(file_link)
# (f"파일 이름: {filename}")

# import sys
# sys.exit()
