import streamlit as st


class In:
    pass


In.ok = ':blue[∴ OK] (🆗✅)'
In.ng = ':red[∴ NG] (❌)'
In.col_span_ref = [1, 1]
In.col_span_okng = [5, 1]  # 근거, OK(NG) 등 2열 배열 간격 설정
In.max_width = '1800px'

In.font_h1 = '32px'
In.font_h2 = '28px'
In.font_h3 = '26px'
In.font_h4 = '24px'
In.font_h5 = '20px'
In.font_h6 = '16px'
In.h2 = '## '
In.h3 = '### '
In.h4 = '#### '
In.h5 = '##### '
In.h6 = '###### '
In.s1 = In.h5 + '$\quad$'
In.s2 = In.h5 + '$\qquad$'
In.s3 = In.h5 + '$\quad \qquad$'

In.border1 = f'<hr style="border-top: 2px solid green; margin-top:30px; margin-bottom:30px; margin-right: -30px">'  # 1줄
In.border2 = f'<hr style="border-top: 5px double green; margin-top: 0px; margin-bottom:30px; margin-right: -30px">'  # 2줄


def Sidebar():
    sb = st.sidebar
    side_border = '<hr style="border-top: 2px solid purple; margin-top:15px; margin-bottom:15px;">'
    h5 = In.h5
    h4 = h5

    html_code = """
        <div style="background-color: lightblue; margin-top: 10px; padding: 10px; padding-top: 20px; padding-bottom:0px; font-weight:bold; border: 2px solid black; border-radius: 20px;">
            <h5>문의 사항은 언제든지 아래 이메일로 문의 주세요^^</h5>
            <h5>📧📬 : <a href='mailto:strustar@konyang.ac.kr' style='color: blue;'>strustar@konyang.ac.kr</a> (건양대 손병직)</h5>
        </div>
    """
    sb.markdown(html_code, unsafe_allow_html=True)

    sb.write('')
    sb.write('## ', ':blue[[Information : 입력값 📘]]')
    sb.write('')
    # sb.write(h4, '✤ 워터마크(watermark) 제거*')
    # col = sb.columns(2)
    # with col[0]:
    #     In.watermark = st.text_input(h5 + '✦ 숨김', type='password', placeholder='password 입력하세요' , label_visibility='collapsed')  # , type='password'
    # sb.write('###### $\,$', ':blue[*워터마크를 제거 하시려면 메일로 문의주세요]')

    sb.markdown(side_border, unsafe_allow_html=True)  #  구분선 ------------------------------------
    sb.write(h4, ':green[✤ Design Code]')
    col = sb.columns([1, 1.2], gap='medium')
    with col[0]:
        In.RC_Code = st.selectbox(h5 + '￭ RC Code', ('KDS-2021', 'KCI-2012'), key='RC_Code')
    with col[1]:
        In.FRP_Code = st.selectbox(h5 + '￭ FRP Code', ('AASHTO-2018', 'ACI 440.1R-06(15)', 'ACI 440.11-22'), key='FRP_Code')

    sb.markdown(side_border, unsafe_allow_html=True)  #  구분선 ------------------------------------
    col = sb.columns([1, 1.2])
    with col[0]:
        st.write(h4, ':green[✤ Column Type]')
        In.Column_Type = st.radio(h5 + '￭ Section Type', ('Tied Column', 'Spiral Column'), key='Column_Type', label_visibility='collapsed')
    with col[1]:
        st.write(h4, ':green[✤ PM Diagram Option]')
        In.PM_Type = st.radio('PM Type', ('RC vs. FRP', 'Pₙ-Mₙ vs. ϕPₙ-ϕMₙ'), horizontal=True, label_visibility='collapsed', key='PM_Type')

    sb.markdown(side_border, unsafe_allow_html=True)  #  구분선 ------------------------------------
    sb.write(h4, ':green[✤ Section Dimensions]')
    col = sb.columns([2, 1])
    with col[0]:
        In.Section_Type = st.radio(h5 + '￭ Section Type', ('Rectangle', 'Circle'), key='Section_Type', horizontal=True, label_visibility='collapsed')

    col = sb.columns([1, 1, 1])
    if 'Rectangle' in In.Section_Type:
        disabledR = False
        disabledC = True
    if 'Circle' in In.Section_Type:
        disabledR = True
        disabledC = False

    with col[0]:
        In.be = st.number_input(h5 + r'￭ $\bm{{\small{{b_e}} }}$ [mm]', min_value=10.0, value=1000.0, step=10.0, format='%f', key='be', disabled=disabledR)
    with col[1]:
        In.height = st.number_input(h5 + r'￭ $\bm{{\small{{h}} }}$ [mm]', min_value=10.0, value=300.0, step=10.0, format='%f', key='height', disabled=disabledR)
    with col[2]:
        In.D = st.number_input(h5 + r'￭ $\bm{{\small{{D}} }}$ [mm]', min_value=10.0, value=600.0, step=10.0, format='%f', key='D', disabled=disabledC)

    sb.markdown(side_border, unsafe_allow_html=True)  #  구분선 ------------------------------------
    sb.write(h4, ':green[✤ Material Properties]')
    col = sb.columns(2, gap='medium')
    with col[0]:
        In.fck = st.number_input(h5 + r'￭ $\bm{{\small{{f_{ck}}} }}$ [MPa]', min_value=10.0, value=40.0, step=1.0, format='%f', key='fck')
        In.fy = st.number_input(h5 + r'￭ $\bm{{\small{{f_{y}}} }}$ [MPa]', min_value=10.0, value=800.0, step=10.0, format='%f', key='fy')
        In.f_fu = st.number_input(h5 + r'￭ $\bm{{\small{{f_{fu}}} }}$ [MPa]', min_value=10.0, value=1000.0, step=10.0, format='%f', key='f_fu')
        Ec = 8500 * (In.fck + 4) ** (1 / 3) / 1e3
    with col[1]:  # MPa로 변경 *1e3
        In.Ec = st.number_input(h5 + r'￭ $\bm{{\small{{E_{c}}} }}$ [GPa]', min_value=10.0, value=Ec, step=1.0, format='%.1f', disabled=True, key='Ec') * 1e3
        In.Es = st.number_input(h5 + r'￭ $\bm{{\small{{E_{s}}} }}$ [GPa]', min_value=10.0, value=200.0, step=10.0, format='%f', key='Es') * 1e3
        In.Ef = st.number_input(h5 + r'￭ $\bm{{\small{{E_{f}}} }}$ [GPa]', min_value=10.0, value=100.0, step=10.0, format='%f', key='Ef') * 1e3

    sb.markdown(side_border, unsafe_allow_html=True)  #  구분선 ------------------------------------
    sb.write(h4, ':green[✤ Reinforcement Layer (Rebar & FRP)]')
    Layer = sb.radio('숨김', ('Layer 1', 'Layer 2', 'Layer 3'), horizontal=True, label_visibility='collapsed', key='Layer')
    if '1' in Layer:
        In.Layer = 1
    if '2' in Layer:
        In.Layer = 2
    if '3' in Layer:
        In.Layer = 3

    In.dia, In.dc, In.nh, In.nb, In.nD = [], [], [], [], []
    col = sb.columns(3)
    for i in range(In.Layer):
        key_dia = 'dia' + str(i + 1)
        key_dc = 'dc' + str(i + 1)
        key_nh = 'nh' + str(i + 1)
        key_nb = 'nb' + str(i + 1)
        key_nD = 'nD' + str(i + 1)
        label_visibility = 'visible' if i == 0 else 'hidden'
        with col[i]:
            In.dia.append(st.number_input(h5 + '￭ dia [mm] : 보강재 직경', min_value=1.0, value=22.2, step=1.0, format='%f', key=key_dia, label_visibility=label_visibility))
            In.dc.append(st.number_input(h5 + r'￭ $\bm{d_c}$ [mm] : 피복 두께', min_value=1.0, value=60.0 + 40 * i, step=1.0, format='%f', key=key_dc, label_visibility=label_visibility))

            if 'Rectangle' in In.Section_Type:
                In.nh.append(st.number_input(h5 + r'￭ $\bm{n_h}$ [EA] : $\bm{{\small{h}}}$방향 보강재 개수', min_value=2, value=2, step=1, format='%d', key=key_nh, label_visibility=label_visibility))
                In.nb.append(st.number_input(h5 + r'￭ $\bm{n_b}$ [EA] : $\bm{{\small{b}}}$방향 보강재 개수', min_value=2, value=6, step=1, format='%d', key=key_nb, label_visibility=label_visibility))
            if 'Circle' in In.Section_Type:
                In.nD.append(st.number_input(h5 + r'￭ $\bm{n_D}$ [EA] : 원형 단면 총 보강재 개수', min_value=2, value=8, step=1, format='%d', key=key_nh, label_visibility=label_visibility))

    return In
