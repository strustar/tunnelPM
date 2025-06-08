import streamlit as st
import Column_Sidebar, Column_Calculate, Column_Result, Excel
import Check_Shear, Check_Column, Check_Serviceability, Check_Steel_Stress_Fcn

import os, sys, importlib

# os.system('cls')  # 터미널 창 청소, clear screen
# sys.path.append("D:\\Work_Python\\Common")  # 공통 스타일 변수 디렉토리 추가
import common_style  # print(sys.path)

# importlib.reload(common_style)  # 다른 폴더 자동 변경이 안됨? ㅠ

### * -- Set page config
st.set_page_config(
    page_title="Column Design (FRP vs. Rebar)",
    page_icon="column.png",
    layout="wide",  # centered, wide
    initial_sidebar_state="expanded",  # runOnSave = True,
    menu_items={
        # 'Get Help': 'https://www.extremelycoolapp.com/help',
        # 'Report a bug': "https://www.extremelycoolapp.com/bug",
        # 'About': "# This is a header. This is an *extremely* cool app!"
    },
)

In = Column_Sidebar.Sidebar()
common_style.input_box(In)
# commonStyle.watermark(In)

R = Column_Calculate.Cal(In, 'RC')  # 이형철근
F = Column_Calculate.Cal(In, 'RC_hollow')  # 중공철근
# R.Ast_total
# F.Ast_total

if '기둥' in In.Option:
    Column_Result.Fig(In, R, F)
    Check_Column.check_column(In, R, F)
elif '전단' in In.Option:
    Check_Shear.check_shear(In, R)    # 전단철근은 이형철근으로 검토
elif '사용성' in In.Option:
    R.fy = In.fy
    F.fy = In.fy_hollow
    
    results_R = Check_Steel_Stress_Fcn.calculate_steel_stress(In, R)
    results_F = Check_Steel_Stress_Fcn.calculate_steel_stress(In, F)
    # 실패한 케이스가 있는지 체크
    for res in results_R:
        if not res.get("success", False):
            st.error(f"❌ 오류 발생: {res.get('message', '알 수 없는 오류')}")
            st.error(f"❌ 오류 발생: 사용성 검토 실패 (사용 모멘트를 줄이세요)")
            st.stop()
    In.Cc = In.dc[0] - In.dia[0]/2
    R.fs = [x['fs'] for x in results_R]  # 모든 요소의 'fs' 값을 리스트로 추출
    R.x = [x['x'] for x in results_R]  # 모든 요소의 'x' 값을 리스트로 추출
    F.fs = [x['fs'] for x in results_F]  # 모든 요소의 'fs' 값을 리스트로 추출
    F.x = [x['x'] for x in results_F]  # 모든 요소의 'x' 값을 리스트로 추출

    Check_Serviceability.display_basic_theory()
    Check_Serviceability.serviceability_check_results(In, R, F)

# Excel.Excel(In, R, F)

# import sys
# sys.exit()
