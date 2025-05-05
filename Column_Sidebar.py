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

In.border1 = (
    f'<hr style="border-top: 2px solid green; margin-top:30px; margin-bottom:30px; margin-right: -30px">'  # 1줄
)
In.border2 = (
    f'<hr style="border-top: 5px double green; margin-top: 0px; margin-bottom:30px; margin-right: -30px">'  # 2줄
)


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
    sb.write(h4, ':green[✤ Design Method]')
    In.Design_Method = sb.radio(
        h5 + '￭ Design Method',
        ('USD (Ultimate Strength Design)', 'LSD (Limit State Design)'),
        key='Design_Method',
        horizontal=True,
        label_visibility='collapsed',
    )

    # sb.write(h4, ':green[✤ Design Code]')
    # col = sb.columns([1, 1.2], gap='medium')
    # with col[0]:
    #     In.RC_Code = st.selectbox(h5 + '￭ RC Code', ('KDS-2021', 'KCI-2012'), key='RC_Code')
    # with col[1]:
    #     In.FRP_Code = st.selectbox(h5 + '￭ FRP Code', ('AASHTO-2018', 'ACI 440.1R-06(15)', 'ACI 440.11-22'), key='FRP_Code')
    In.RC_Code = 'KDS-2021'
    In.FRP_Code = 'AASHTO-2018'

    sb.markdown(side_border, unsafe_allow_html=True)  #  구분선 ------------------------------------
    col = sb.columns([1, 1.2])
    with col[0]:
        st.write(h4, ':green[✤ Column Type]')
        In.Column_Type = st.radio(
            h5 + '￭ Section Type', ('Tied Column', 'Spiral Column'), key='Column_Type', label_visibility='collapsed'
        )
    with col[1]:
        st.write(h4, ':green[✤ PM Diagram Option]')
        In.PM_Type = st.radio(
            'PM Type',
            ('이형철근 \u00a0 vs. \u00a0 중공철근', 'Pₙ-Mₙ \u00a0 vs. \u00a0 ϕPₙ-ϕMₙ'),
            horizontal=True,
            label_visibility='collapsed',
            key='PM_Type',
        )

    sb.markdown(side_border, unsafe_allow_html=True)  #  구분선 ------------------------------------
    sb.write(h4, ':green[✤ Section Dimensions]')
    In.Section_Type = 'Rectangle'

    col = sb.columns([1, 1])
    with col[0]:
        In.be = st.number_input(
            h5 + r'￭ $\bm{{\small{{b_e}} }}$ (단위폭) [mm]',
            min_value=10.0,
            value=1000.0,
            step=10.0,
            format='%f',
            key='be',
        )
    with col[1]:
        In.height = st.number_input(
            h5 + r'￭ $\bm{{\small{{h}} }}$ [mm]', min_value=10.0, value=300.0, step=10.0, format='%f', key='height'
        )
    In.D = 500
    # with col[2]:
    #     In.D = st.number_input(h5 + r'￭ $\bm{{\small{{D}} }}$ [mm]', min_value=10.0, value=600.0, step=10.0, format='%f', key='D', disabled=disabledC)

    sb.markdown(side_border, unsafe_allow_html=True)  #  구분선 ------------------------------------
    sb.write(h4, ':green[✤ Material Properties]')
    col = sb.columns(3, gap='medium')
    with col[0]:
        st.write(h5, ':blue[✦ 콘크리트]')
        In.fck = st.number_input(
            h5 + r'$\bm{{\small{{f_{ck}}} }}$ [MPa]', min_value=10.0, value=40.0, step=1.0, format='%f', key='fck'
        )
        Ec = 8500 * (In.fck + 4) ** (1 / 3) / 1e3
        In.Ec = (
            st.number_input(
                h5 + r'$\bm{{\small{{E_{c}}} }}$ [GPa]',
                min_value=10.0,
                value=Ec,
                step=1.0,
                format='%.1f',
                disabled=True,
                key='Ec',
            )
            * 1e3
        )
    with col[1]:  # MPa로 변경 *1e3
        st.write(h5, ':blue[✦ 이형철근]')
        In.fy = st.number_input(
            h5 + r'$\bm{{\small{{f_{y}}} }}$ [MPa]', min_value=10.0, value=400.0, step=10.0, format='%f', key='fy'
        )
        In.Es = (
            st.number_input(
                h5 + r'$\bm{{\small{{E_{s}}} }}$ [GPa]', min_value=10.0, value=200.0, step=10.0, format='%f', key='Es'
            )
            * 1e3
        )
    with col[2]:
        st.write(h5, ':blue[✦ 중공철근]')
        In.f_fu = st.number_input(
            h5 + r'$\bm{{\small{{f_{y}}} }}$ [MPa]', min_value=10.0, value=800.0, step=10.0, format='%f', key='f_fu'
        )
        In.Ef = (
            st.number_input(
                h5 + r'$\bm{{\small{{E_{s}}} }}$ [GPa]', min_value=10.0, value=200.0, step=10.0, format='%f', key='Ef'
            )
            * 1e3
        )
        In.fy_hollow = In.f_fu
        In.Es_hollow = In.Ef

    In.Layer = 1
    In.nD = [8]
    sb.markdown(side_border, unsafe_allow_html=True)  #  구분선 ------------------------------------
    sb.write(h4, ':green[✤ Reinforcement Layer (Rebar & FRP)]')
    In.nb = [6]
    In.nh = [2]
    col = sb.columns(2, gap='medium')
    with col[0]:
        In.sb = [
            st.number_input(
                h5 + r'$\bm{s_b}$ [mm] : $\bm{{\small{b}}}$(단위폭) 방향 보강재 간격',
                min_value=10.0,
                value=150.0,
                step=10.0,
                format='%f',
                key='sb',
                label_visibility='visible',
            )
        ]

    col = sb.columns(2, gap='large')
    with col[0]:
        st.write(h5, ':blue[✦ 인장측]')
        In.dia = [
            st.number_input(
                h5 + 'dia [mm] : 보강재 외경',
                min_value=1.0,
                value=22.2,
                step=1.0,
                format='%f',
                key='dia',
                label_visibility='visible',
            )
        ]
        In.dc = [
            st.number_input(
                h5 + r'$\bm{d_c}$ [mm] : 피복 두께',
                min_value=1.0,
                value=60.0,
                step=2.0,
                format='%f',
                key='dc',
                label_visibility='visible',
            )
        ]

    with col[1]:
        st.write(h5, ':blue[✦ 압축측]')
        In.dia1 = [
            st.number_input(
                h5 + 'dia [mm] : 보강재 외경',
                min_value=1.0,
                value=22.2,
                step=1.0,
                format='%f',
                key='dia1',
                label_visibility='hidden',
            )
        ]
        In.dc1 = [
            st.number_input(
                h5 + r'$\bm{d_c}$ [mm] : 피복 두께',
                min_value=1.0,
                value=60.0,
                step=2.0,
                format='%f',
                key='dc1',
                label_visibility='hidden',
            )
        ]
    return In
