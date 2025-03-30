import streamlit as st
import Column_Sidebar, Column_Calculate, Column_Result

import os, sys, importlib

os.system('cls')  # 터미널 창 청소, clear screen
sys.path.append("D:\\Work_Python\\Common")  # 공통 스타일 변수 디렉토리 추가
import commonStyle  # print(sys.path)

importlib.reload(commonStyle)  # 다른 폴더 자동 변경이 안됨? ㅠ

### * -- Set page config
st.set_page_config(
    page_title="Column Design (FRP vs. Rebar)",
    page_icon="column.png",
    layout="centered",  # centered, wide
    initial_sidebar_state="expanded",  # runOnSave = True,
    menu_items={
        # 'Get Help': 'https://www.extremelycoolapp.com/help',
        # 'Report a bug': "https://www.extremelycoolapp.com/bug",
        # 'About': "# This is a header. This is an *extremely* cool app!"
    },
)

In = Column_Sidebar.Sidebar()
commonStyle.input_box(In)
# commonStyle.watermark(In)

R = Column_Calculate.Cal(In, 'RC')
F = Column_Calculate.Cal(In, 'FRP')

Column_Result.Fig(In, R, F)


# import sys
# sys.exit()
