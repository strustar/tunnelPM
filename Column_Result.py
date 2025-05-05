import streamlit as st
import plotly.graph_objects as go
import numpy as np
import pandas as pd
from Column_Aggrid import display_data_with_aggrid

family = 'sans-serif, Arial, Nanumgothic, Georgia'
size8 = 8
size15 = 15
size17 = 17
size20 = 20


def Fig(In, R, F):
    if 'ACI 440.1' in In.FRP_Code:  #! for ACI 440.1R**  Only Only, 추가점 c=0(x=0 일때)
        F.ZPn = np.insert(F.ZPn, -1, F.Pn8)
        F.ZMn = np.insert(F.ZMn, -1, F.Mn8)
        F.ZPd = np.insert(F.ZPd, -1, F.Pd8)
        F.ZMd = np.insert(F.ZMd, -1, F.Md8)
        F.Pn = np.append(F.Pn, F.Pn8)
        F.Mn = np.append(F.Mn, F.Mn8)
        F.Pd = np.append(F.Pd, F.Pd8)
        F.Md = np.append(F.Md, F.Md8)

    PM_RC, PM_FRP = st.columns(2, gap="small")
    with PM_RC:
        container_PM_RC = st.container()
    with PM_FRP:
        container_PM_FRP = st.container()

    column_RC, column_Common, column_FRP = st.columns([1.0, 1, 1.0], gap="small")
    with column_Common:
        container_column_Common = st.container()
    with column_RC:
        container_column_RC = st.container()
    with column_FRP:
        container_column_FRP = st.container()

    table_RC, table_Common, table_FRP = st.columns([1.0, 1, 1.0], gap="small")
    with table_Common:
        if '중공' not in In.PM_Type:
            st.session_state.selected_row = None
        placeholder = st.empty()

        if '중공' in In.PM_Type:
            st.info('아래 라디오 버튼을 클릭하세요', icon="ℹ️")
        else:
            st.warning(':green[좌측 사이드바에서 PM Diagram Option을 "이형철근 vs. 중공철근"으로 설정하세요]', icon="⚠️")

        txt = 'D - ' if 'Circle' in In.Section_Type else 'h - '
        selected_row = st.radio(
            '#### ￭ Select one below',
            (
                'A (Pure Compression) : $M_n$ = 0, $e$ = 0, $c$ = inf',
                'B (Minimum Eccentricity) : $e$ = $e_{min}$',
                f'C (Zero Tension) : $\epsilon_t$ = 0, $c$ = {txt}$d_c$',
                'D (Balance Point) : $e$ = $e_{b}$',
                'E : $\epsilon_t$ = 2.5$\epsilon_y$ [RC]',
                'F (Pure Moment) : $P_n$ = 0, $e$ = inf',
                'G (Pure Tension) : $M_n$ = 0, $e$ = 0, $c$ = -inf',
            ),
            key='selected_row',
            index=None,
            label_visibility='collapsed',
        )
    with table_RC:
        Table(In, R, F, 'RC', selected_row)
    with table_FRP:
        Table(In, R, F, 'RC_hollow', selected_row)

    st.write('')
    Table(In, R, F, 'Full Data', selected_row)
    # data = np.array([[PM.Ze[n], PM.Zc[n], PM.ZPn[n], PM.ZMn[n], PM.Zphi[n], PM.ZPd[n], PM.ZMd[n], PM.Zep_s[n,1], PM.Zfs[n,1], PM.Zep_s[n,0], PM.Zfs[n,0]] for n in range(len(PM.Ze))])
    # st.write(data)

    # selected_row가 설정된 상태에서 PM 상관도를 그리기 위해서
    fig = PM_plot(In, R, F, 'RC', selected_row)
    container_PM_RC.plotly_chart(fig, config={'displayModeBar': False})
    fig = PM_plot(In, R, F, 'RC_hollow', selected_row)
    container_PM_FRP.plotly_chart(fig, config={'displayModeBar': False})

    # selected_row가 설정된 상태에서 기둥을 그리기 위해서
    fig = Column(In, R, F, 'Common', None)
    container_column_Common.plotly_chart(fig, config={'displayModeBar': False})
    fig = Column(In, R, F, 'RC', selected_row)
    container_column_RC.plotly_chart(fig, config={'displayModeBar': False})
    fig = Column(In, R, F, 'RC_hollow', selected_row)
    container_column_FRP.plotly_chart(fig, config={'displayModeBar': False})


def Table(In, R, F, loc, selected_row):
    headers = [
        '<b>   e<br>[mm]',
        '<b>   c<br>[mm]',
        '<b>  P<sub>n</sub> <br>[kN]',
        '<b>   M<sub>n</sub> <br>[kN·m]',
        '<b>ϕ',
        '<b> ϕP<sub>n</sub> <br>[kN]',
        '<b>  ϕM<sub>n</sub> <br>[kN·m]',
    ]

    PM = F if 'hollow' in loc else R
    data = np.array([[PM.e[n], PM.c[n], PM.Pn[n], PM.Mn[n], PM.phi[n], PM.Pd[n], PM.Md[n]] for n in range(7)])
    data_str = [
        [f'<b>{num:,.3f}</b>' if idx == 4 else f'<b>{num:,.1f}</b>' for idx, num in enumerate(row)] for row in data
    ]
    color = red_text(data)

    columnwidth = [1, 1, 1, 1, 1, 1, 1]
    height = 350
    width = 550
    left = 1
    cells_align = 'right'
    n = 100
    if selected_row != None:
        if 'A (P' in selected_row:
            n = 0
        if 'B (M' in selected_row:
            n = 1
        if 'C (Z' in selected_row:
            n = 2
        if 'D (B' in selected_row:
            n = 3
        if 'E ' in selected_row:
            n = 4
        if 'F (P' in selected_row:
            n = 5
        if 'G (P' in selected_row:
            n = 6
    cells_fill_color = [['lightblue' if i == n else 'white' for i in range(7)] for _ in range(7)]

    if 'Full Data' in loc:
        headers = headers + [
            '<b>ε<sub>t</sub>',
            '<b>   f<sub>t</sub> <br>[MPa]',
            '<b>ε<sub>c</sub>',
            '<b>   f<sub>c</sub> <br>[MPa]',
        ]
        # columnwidth = [1, 0.9, 1, 0.8, 0.8, 1, 0.8, 1, 1, 1, 1]
        height = 800
        width = 840
        cells_fill_color = 'lightblue'

        col_RC, col_center, col_FRP = st.columns([7, 0.1, 7], gap='small')
        with col_RC:
            st.write('### :blue[[이형철근]]')
            display_data_with_aggrid(headers, R, height=height, width=width)
        with col_FRP:
            st.write('### :blue[[중공철근]]')
            display_data_with_aggrid(headers, F, height=height, width=width)
    else:
        common_table(
            headers,
            data_str,
            columnwidth,
            cells_align,
            cells_fill_color,
            height,
            left,
            color=color,
            fill_color='gainsboro',
            width=width,
        )


def red_text(data):
    # 값이 음수일때, 빨간색으로 설정하기 위해서
    sz = data.shape
    sgn = np.zeros([sz[0], sz[1]])
    for i in range(sz[0]):
        for j in range(sz[1]):
            if data[i][j] < 0:
                sgn[i, j] = 1
    color = [['red' if sgn[i][j] == 1 else 'black' for i in range(sz[0])] for j in range(sz[1])]

    return color


def common_table(headers, data, columnwidth, cells_align, cells_fill_color, height, left, **kargs):
    if np.ndim(data) == 1:
        data_dict = {header: [value] for header, value in zip(headers, data)}  # 행이 한개 일때
    else:
        data_dict = {header: values for header, values in zip(headers, zip(*data))}  # 행이 여러개(2개 이상) 일때
    df = pd.DataFrame(data_dict)

    fill_color = ['gainsboro']
    lw = 2
    if len(kargs) > 0:
        fill_color = kargs['fill_color']
        color = kargs['color']
        width = kargs['width']

    fig = go.Figure(
        data=[
            go.Table(
                columnwidth=columnwidth,
                header=dict(
                    values=list(df.columns),
                    align=['center'],
                    font_color='black',
                    fill_color=fill_color,
                    font=dict(size=size17, color='black', family=family),
                    line=dict(color='black', width=lw),
                ),  #'darkgray'  # 글꼴 변경  # 셀 경계색, 두께
                cells=dict(
                    values=[df[col] for col in df.columns],
                    align=cells_align,
                    font_color=color,
                    fill_color=cells_fill_color,  # 셀 배경색 변경
                    font=dict(size=size17, color='black', family=family),  # 글꼴 변경
                    line=dict(color='black', width=lw),  # 셀 경계색, 두께
                    # format=['.1f']  # '나이' 열의 데이터를 실수 형태로 변환하여 출력  '.2f'
                ),
            )
        ]
    )
    fig.update_layout(autosize=False, width=width, height=height, margin=dict(l=left, r=1, t=1, b=0))
    st.plotly_chart(fig, config={'displayModeBar': False})


def shape(fig, typ, x0, y0, x1, y1, fillcolor, color, width, **kargs):
    dash = 'solid'
    if len(kargs) > 0:
        dash = kargs['LineStyle']
    fig.add_shape(
        type=typ, x0=x0, y0=y0, x1=x1, y1=y1, fillcolor=fillcolor, line=dict(color=color, width=width, dash=dash)
    )  # dash = solid, dot, dash, longdash, dashdot, longdashdot, '5px 10px'


def annotation(fig, x, y, color, txt, xanchor, yanchor, **kargs):
    bgcolor = None
    size = size17
    if len(kargs) > 0:
        bgcolor = kargs['bgcolor']
        size = kargs['size']
    fig.add_annotation(
        x=x,
        y=y,
        text=txt,
        showarrow=False,
        bgcolor=bgcolor,
        font=dict(color=color, size=size, family=family),
        xanchor=xanchor,
        yanchor=yanchor,
    )


def dimension(fig, x0, y0, length, fillcolor, color, width, txt, loc, opt):
    shift = 35
    arrow_size = 12
    if 'Doubl' in opt:
        shift = 95
    if 'force' in opt:
        shift = 0
        arrow_size = 18
    arrow_width = arrow_size / 4  # 화살표 두께 고정 !!!

    if ('bottom' in loc) or ('top' in loc):  # 가로 치수 (bottom & top)
        y0 = y0 - shift if 'bottom' in loc else y0 + shift - 13

        y1 = y0
        x1 = x0 + length
        tx = (x0 + x1) / 2
        ty = y0 + 13 if 'bottom' in loc else y0 + 20
        locx = 'center'
        locy = 'middle'
        if '1' in opt:
            tx = x1
            locx = 'right'
        if '2' in opt:
            tx = x1
            locx = 'left'
        if 'up' in opt:
            ty = y0
            locy = 'bottom'
        if 'down' in opt:
            ty = y0
            locy = 'top'

        if 'tension' in opt and '1' in opt:
            tx = x0
            locx = 'right'
        if 'tension' in opt and '2' in opt:
            tx = x0
            locx = 'left'
    else:  # 세로 치수 (left, right)
        if 'left' in loc:
            sgn = 1
            locx = 'right'
            locy = 'middle'
        if 'right' in loc:
            sgn = -1
            locx = 'left'
            locy = 'middle'
        x0 = x0 - sgn * shift
        x1 = x0
        y1 = y0 + length
        tx = x0 - sgn * 10
        ty = (y0 + y1) / 2

    shape(fig, 'line', x0, y0, x1, y1, fillcolor, color, width)  # 치수선
    if 'False' not in opt:
        annotation(fig, tx, ty, color, txt, locx, locy)  # 치수 문자

    for i in [1, 2]:
        if ('bottom' in loc) or ('top' in loc):  # 가로 치수 (bottom & top)
            if x0 <= x1:
                sgn = 1
            if x0 >= x1:
                sgn = -1
            b0 = y0
            b1 = y0 - arrow_width
            b2 = y0 + arrow_width
            if i == 1:
                a0 = x0
                a1 = x0 + sgn * arrow_size
            if i == 2:
                a0 = x1
                a1 = x1 - sgn * arrow_size
            a2 = a1
            p0 = a0
            p1 = a0
            q0 = y0 - arrow_size
            q1 = y0 + arrow_size
        else:  # 세로 치수 (left, right)
            if y0 <= y1:
                sgn = 1
            if y0 >= y1:
                sgn = -1
            a0 = x0
            a1 = x0 - arrow_width
            a2 = x0 + arrow_width
            if i == 1:
                b0 = y0
                b1 = y0 + sgn * arrow_size
            if i == 2:
                b0 = y1
                b1 = y1 - sgn * arrow_size
            b2 = b1
            p0 = x0 - arrow_size
            p1 = x0 + arrow_size
            q0 = b0
            q1 = b0

        if 'force' not in opt:
            shape(fig, 'line', p0, q0, p1, q1, fillcolor, color, width)  # 치수 보조선
        if i == 2 and 'force' in opt:
            continue
        fig.add_shape(
            type="path",
            path=f"M {a0} {b0} L {a1} {b1} L {a2} {b2} Z",
            fillcolor=fillcolor,
            line=dict(color=color, width=width),
        )  # 화살표  # M = move to, L = line to, Z = close path


def PM_plot(In, R, F, loc, selected_row):
    fig = go.Figure()
    fig_width = 800
    fig_height = 620
    [titleRC, titleFRP] = ['이형철근', '중공철근']
    [titlePM, titlephiPM] = ['P<sub>n</sub>-M<sub>n</sub> Diagram', 'ϕP<sub>n</sub>-ϕM<sub>n</sub> Diagram']

    if '중공' in In.PM_Type:
        [xtxt, ytxt] = ['M<sub>n</sub> or ϕM<sub>n</sub> [kN·m]', 'P<sub>n</sub> or ϕP<sub>n</sub> [kN]']
        [legend1, legend2] = [titlePM, titlephiPM]
        if 'RC' in loc:
            title = titleRC
            color1 = 'red'
            dash1 = 'dot'
            color2 = 'green'
            dash2 = 'solid'
            PM_x1 = R.ZMn
            PM_y1 = R.ZPn
            PM_x2 = R.ZMd
            PM_y2 = R.ZPd
            PM_x7 = R.Mn
            PM_y7 = R.Pn
            PM_x8 = R.Md
            PM_y8 = R.Pd
        if 'hollow' in loc:
            title = titleFRP
            color1 = 'magenta'
            dash1 = 'dot'
            color2 = 'royalblue'
            dash2 = 'solid'
            PM_x1 = F.ZMn
            PM_y1 = F.ZPn
            PM_x2 = F.ZMd
            PM_y2 = F.ZPd
            PM_x7 = F.Mn
            PM_y7 = F.Pn
            PM_x8 = F.Md
            PM_y8 = F.Pd

    if '중공' not in In.PM_Type:
        [legend1, legend2] = [titleRC, titleFRP]
        if 'RC' in loc:
            title = titlePM
            color1 = 'red'
            dash1 = 'dot'
            color2 = 'magenta'
            dash2 = 'dot'
            xtxt = 'M<sub>n</sub> [kN·m]'
            ytxt = 'P<sub>n</sub> [kN]'
            PM_x1 = R.ZMn
            PM_y1 = R.ZPn
            PM_x2 = F.ZMn
            PM_y2 = F.ZPn
            PM_x7 = R.Mn
            PM_y7 = R.Pn
            PM_x8 = F.Mn
            PM_y8 = F.Pn
        if 'hollow' in loc:
            title = titlephiPM
            color1 = 'green'
            dash1 = 'solid'
            color2 = 'royalblue'
            dash2 = 'solid'
            xtxt = 'ϕM<sub>n</sub> [kN·m]'
            ytxt = 'ϕP<sub>n</sub> [kN]'
            PM_x1 = R.ZMd
            PM_y1 = R.ZPd
            PM_x2 = F.ZMd
            PM_y2 = F.ZPd
            PM_x7 = R.Md
            PM_y7 = R.Pd
            PM_x8 = F.Md
            PM_y8 = F.Pd

    fig.add_trace(
        go.Scatter(
            x=PM_x1,
            y=PM_y1,
            mode='lines',
            name=legend1,
            line=dict(width=3, color=color1, dash=dash1, shape='spline'),
            fill='tozeroy',
            fillcolor='rgba(255,0,0,0.1)',
        )
    )
    fig.add_trace(
        go.Scatter(
            x=PM_x2,
            y=PM_y2,
            mode='lines',
            name=legend2,
            line=dict(width=3, color=color2, dash=dash2, shape='spline'),
            fill='tozeroy',
            fillcolor='rgba(0,255,0,0.1)',
        )
    )

    xmax = 1.19 * max(PM_x1)
    ymin = 1.25 * min([min(PM_y1), min(PM_y2)])
    ymax = 1.15 * max(PM_y1)
    shape(fig, 'rect', 0, ymin, xmax, ymax, 'rgba(0,0,0,0)', 'RoyalBlue', 3)  # 그림 상자 외부 박스

    # Update the layout properties
    fig.update_layout(
        autosize=False,
        width=fig_width,
        height=fig_height,
        margin=dict(l=0, r=0, t=0, b=0),  # hovermode='x unified',
        hoverlabel=dict(bgcolor='lightgray', font_size=size20, font_color='blue'),
        legend=dict(
            x=0.99,
            y=0.99,
            xanchor='right',
            yanchor='top',
            font_size=size15,
            font_color='black',
            borderwidth=1,
            bordercolor='black',
        ),
        annotations=[
            dict(
                text=title,
                x=xmax / 2,
                y=0.99 * (ymax - ymin) + ymin,
                xanchor='center',
                yanchor='top',
                bgcolor='gold',
                showarrow=False,
                font=dict(size=size20, family=family, color='black'),
            )
        ],
    )
    fig.update_traces(hovertemplate='M : %{x}<br>P : %{y}')
    fig.update_xaxes(
        range=[0, xmax],
        showgrid=True,
        gridwidth=2,  # gridcolor='gray', griddash='solid',
        showspikes=True,
        spikecolor="gray",
        spikesnap="cursor",
        spikemode="across",
        spikethickness=3,
        title=dict(text=xtxt, standoff=20, font=dict(size=size17, family=family, color='black')),
        tickfont=dict(size=size17, color='black'),
        tickformat=',.0f',
        nticks=10,
        tickmode='auto',
    )
    fig.update_yaxes(
        range=[ymin, ymax],
        showgrid=True,
        gridwidth=2,  # gridcolor='gray', griddash='solid',
        showspikes=True,
        spikecolor="gray",
        spikesnap="cursor",
        spikemode="across",
        spikethickness=3,
        title=dict(text=ytxt, standoff=20, font=dict(size=size17, family=family, color='black')),
        tickfont=dict(size=size17, color='black'),
        tickformat=',.0f',
        nticks=10,
        tickmode='auto',
        zeroline=True,
        zerolinewidth=2,
        zerolinecolor='black',
    )

    # phi*Pn(max), e_min(B), Zero Tension(C), e_b(D)
    if '중공' in In.PM_Type:  #! PM_Type = RC vs. FRP 일때만 Only
        # phi*Pn(max)
        if 'RC' in loc:
            [x, y, color] = [R.Md[2 - 1], R.Pd[2 - 1], 'green']
        if 'hollow' in loc:
            [x, y, color] = [F.Md[2 - 1], F.Pd[2 - 1], 'royalblue']
        fig.add_trace(
            go.Scatter(x=[0, x], y=[y, y], line=dict(width=3, color=color), showlegend=False, name='delete_hover')
        )
        txt = f'ϕP<sub>n(max)</sub> = {y:,.1f} kN'
        annotation(fig, x * 0.04, 0.97 * y, 'black', txt, 'left', 'top')

        # e_min(B), Zero Tension(C), e_b(D)
        for i in [1, 2, 3]:
            if 'RC' in loc:
                [x, y] = [R.Mn[i], R.Pn[i]]
            if 'hollow' in loc:
                [x, y] = [F.Mn[i], F.Pn[i]]
            shape(fig, 'line', 0, 0, x, y, None, 'black', 1)

            if 'RC' in loc:
                val = R.e[i]
            if 'hollow' in loc:
                val = F.e[i]
            f = 0.45 if i == 1 else 0.35
            color = 'blue' if val > 0 else 'red'
            if i == 1:
                txt = f'e<sub>min</sub> = {val:,.1f} mm'
            if i == 2:
                txt = f'e<sub>0</sub> = {val:,.1f} mm'
            if i == 3:
                txt = f'e<sub>b</sub> = {val:,.1f} mm'
            annotation(fig, f * x, f * y, color, txt, 'center', 'middle', bgcolor='yellow', size=size17)

    # A(0), B(1), C(2), D(3), E(4), F(5), G(6) 점 찍기
    for i in [1, 2]:  # 1 : Pn-Mn,  2 : Pd-Md
        for z1 in range(len(PM_x7)):
            if i == 1:
                [x1, y1] = [PM_x7[z1], PM_y7[z1]]
                if z1 == len(PM_x7) - 1:
                    txt = 'c = 0'
                if z1 == 1 - 1:
                    txt = 'A (Pure Compression)'  # \;, \quad : 공백 넣기
                if z1 == 2 - 1:
                    txt = 'B (e<sub>min</sub>)'
                if z1 == 3 - 1:
                    txt = 'C (e<sub>0</sub>) Zero Tension'
                if z1 == 4 - 1:
                    txt = 'D (e<sub>b</sub>)          <br>Balance      <br>Point          '
                if z1 == 5 - 1:
                    txt = 'E (ε<sub>t</sub> = 2.5ε<sub>y</sub>)'
                if z1 == 6 - 1:
                    txt = 'F (Pure Moment)'
                if z1 == 7 - 1:
                    txt = 'G (Pure Tension)'

                if 'RC' in loc:
                    bgcolor = 'lightgreen' if z1 == 4 - 1 else None
                if 'hollow' in loc:
                    bgcolor = 'lightblue' if z1 == 4 - 1 else None
                [sgnx, xanchor] = [1, 'left']
                sgny = -1 if z1 >= 4 - 1 else 1
                # if 'hollow' in loc:
                #     # if z1 == 5 - 1:
                #     #     [sgnx, xanchor, txt] = [-1, 'right', 'E (ε<sub>t</sub> = 0.8ε<sub>fu</sub>)']
                #     if z1 == 6 - 1:
                #         txt = 'F'
                if '중공' not in In.PM_Type:
                    if z1 == 5 - 1:
                        txt = 'E'
                    if z1 == 4 - 1 and 'hollow' in loc:
                        [sgny, txt] = [1, 'D']

                [x, y] = [x1 + sgnx * 0.015 * xmax, y1 + sgny * 0.02 * ymax]
                annotation(fig, x, y, 'black', txt, xanchor, 'middle', bgcolor=bgcolor, size=size17)
            elif i == 2:
                [x1, y1] = [PM_x8[z1], PM_y8[z1]]

            if z1 == len(PM_x7) - 1:
                color = 'black'
            if z1 == 1 - 1:
                color = 'red'
            if z1 == 2 - 1:
                color = 'green'
            if z1 == 3 - 1:
                color = 'blue'
            if z1 == 4 - 1:
                color = 'cyan'
            if z1 == 5 - 1:
                color = 'magenta'
            if z1 == 6 - 1:
                color = 'yellow'
            if z1 == 7 - 1:
                color = 'darkred'
            fig.add_trace(
                go.Scatter(
                    x=[x1],
                    y=[y1],
                    mode='markers',
                    marker=dict(size=size8, color=color, line=dict(width=2, color='black')),
                    showlegend=False,
                    name='delete_hover',
                )
            )

    if selected_row != None:
        if 'A (P' in selected_row:
            [n, txt] = [0, 'A (Pure Compression)']
        if 'B (M' in selected_row:
            [n, txt] = [1, 'B (Minimum Eccentricity)']
        if 'C (Z' in selected_row:
            [n, txt] = [2, 'C (Zero Tension)']
        if 'D (B' in selected_row:
            [n, txt] = [3, 'D (Balance Point)']
        if 'E ' in selected_row:
            n = 4
        if 'F (P' in selected_row:
            [n, txt] = [5, 'F (Pure Moment)']
        if 'G (P' in selected_row:
            [n, txt] = [6, 'G (Pure Tension)']

        if 'RC' in loc:
            [x, y] = [R.Mn[n], R.Pn[n]]
        if 'hollow' in loc:
            [x, y] = [F.Mn[n], F.Pn[n]]
        shape(fig, 'line', 0, 0, x, y, None, 'purple', 3)

    fig.update_traces(hoverinfo="skip", selector=dict(name='delete_hover'))
    return fig


def Column(In, R, F, loc, selected_row):
    ####! 공통
    fig = go.Figure()
    f = 1.06
    fig_width = 550
    fig_height = f * fig_width

    # 그림 상자 외부 박스
    xlim = 500
    ylim = f * xlim
    shape(fig, 'rect', 0, 0, xlim, ylim, 'white', 'purple', 2)

    # 공통 기본 치수
    base_scale = 0.5 * xlim
    if 'Rectangle' in In.Section_Type:
        mr = base_scale / min([In.be, In.height])
        w = In.height * mr
        h = In.be * mr
        h = w
        typ = 'rect'
        txt = f'h = {In.hD:,.0f} mm'
    x = (xlim - w) / 2
    y = (ylim - h) / 1.25

    # 기둥
    y_col = 90
    xc = x + w / 2  # 기둥
    shape(fig, 'line', x, 0, x, y_col, None, 'black', 2)
    shape(fig, 'line', x + w, 0, x + w, y_col, None, 'black', 2)
    shape(fig, 'line', x, y_col, x + w, y_col, None, 'black', 2)

    # 기둥에서 철근 형상
    for L in range(In.Layer):
        dc = In.dc[L] * mr
        dc1 = In.dc1[L] * mr
        dia = In.dia[L] * mr
        dia1 = In.dia1[L] * mr
        cr = 1.2 * dia / 2
        cr1 = 1.2 * dia1 / 2
        gap = dc - cr  # 원래 반경(cr)보다 20% 크게~ 보이게
        gap1 = dc1 - cr1

        for i in range(2):
            if i == 0:
                color = 'rgba(255,0,0, 0.3)'
                x1 = x + dc - cr
                shape(fig, 'rect', x1, 0, x1 + 2 * cr, y_col, color, color, 1.5)
            if i == 1:
                color = 'rgba(0,0,255, 0.3)'
                x1 = x + w - dc1 - cr1
                shape(fig, 'rect', x1, 0, x1 + 2 * cr1, y_col, color, color, 1.5)
    ####! 공통

    if 'Common' in loc:
        ### 콘크리트 단면
        shape(fig, typ, x, y, x + w, y + h, 'yellow', 'black', 2)

        dimension(fig, x, y + h, w, 'black', 'black', 1.5, txt, 'top', [])  # 가로 치수
        if 'Rectangle' in In.Section_Type:
            dimension(
                fig, x, y, h, 'black', 'black', 1.5, f'{In.be:,.0f} mm<br>(b<sub>e</sub>)', 'left', []
            )  # 세로 치수
        ### 콘크리트 단면

        ### 철근
        for L in range(In.Layer):
            dc = In.dc[L] * mr
            dc1 = In.dc1[L] * mr
            dia = In.dia[L] * mr
            dia1 = In.dia1[L] * mr
            cr = 1.2 * dia / 2
            cr1 = 1.2 * dia1 / 2
            gap = dc - cr  # 원래 반경(cr)보다 20% 크게~ 보이게
            gap1 = dc1 - cr1

            for i in range(2):  # 전체 원형 등  2개
                if i == 0:
                    color = 'red'
                    x1 = x + dc - cr
                    y1 = y + h / 2 - cr
                    shape(fig, 'circle', x1, y1, x1 + 2 * cr, y1 + 2 * cr, color, 'black', 2)
                if i == 1:
                    color = 'blue'
                    x1 = x + w - dc1 - cr1
                    y1 = y + h / 2 - cr1
                    shape(fig, 'circle', x1, y1, x1 + 2 * cr1, y1 + 2 * cr1, color, 'black', 2)

                if i == 0:
                    txt = f'보강재 간격 : {sum(In.sb):,.0f} mm <br>보강재 개수 : 좌우 {In.be/sum(In.sb):,.2f} 개씩'
                    annotation(fig, x + w / 2, y + h, 'black', txt, 'center', 'top')

                    txt = f'인장측'
                    annotation(fig, x + gap, y + 0.7 * h, 'red', txt, 'left', 'middle')
                    txt = f'외경 : {sum(In.dia):,.1f} mm'
                    annotation(fig, x + gap, y + 0.6 * h, 'red', txt, 'left', 'middle')

                if i == 1:
                    shape(
                        fig,
                        'circle',
                        x1 + 2 * cr1 * 0.2,
                        y1 + 2 * cr1 * 0.2,
                        x1 + 2 * cr1 * 0.8,
                        y1 + 2 * cr1 * 0.8,
                        'yellow',
                        'black',
                        2,
                    )
                    txt = f'압축측'
                    annotation(fig, x + w - gap1, y + 0.3 * h, 'blue', txt, 'right', 'middle')
                    txt = f'외경 : {sum(In.dia1):,.1f} mm'
                    annotation(fig, x + w - gap1, y + 0.4 * h, 'blue', txt, 'right', 'middle')

            # 띠 또는 나선철근
            x1 = x + gap
            y1 = y + h * 0.25
            if 'Rectangle' in In.Section_Type:
                w1 = w - gap - gap1
                h1 = h * 0.5
            shape(fig, typ, x1, y1, x1 + w1, y1 + h1, None, 'black', 1)

            # 가로 치수 (dc)
            y1 = y - 35 * L
            for i in range(2):
                if i == 0:
                    x1 = x
                    txt = f'{In.dc[L]:,.0f} mm'
                    color = 'red'
                    dimension(
                        fig, x1, y1, dc, color, color, 1.5, txt, 'bottom', 'annotationFalse'
                    )  # 가로 치수 (치수 문자 제거)
                    annotation(fig, x1 + dc + 5, y1 - 33, color, txt, 'left', 'middle')  # 치수 문자 (다시 정렬)
                    # rho, 철근량, 보강량 정보 (이형철근, 중공철근)
                    txt = '[이형철근]'
                    annotation(fig, x1 - 5, y1 - 63, 'green', txt, 'right', 'middle')
                    txt = f'{sum(R.Ast_total[:]):,.0f} mm<sup>2</sup> <br>(총 보강량)'
                    annotation(fig, x1 - 5, y1 - 99, 'green', txt, 'right', 'middle')
                    txt = f'ρ = {R.rho:.4f} <br>(보강비)'
                    annotation(fig, x1 + 5, y1 - 99, 'green', txt, 'left', 'middle')

                if i == 1:
                    x1 = x + w - dc1
                    txt = f'{In.dc1[L]:,.0f} mm'
                    color = 'blue'
                    dimension(
                        fig, x1, y1, dc1, color, color, 1.5, txt, 'bottom', 'annotationFalse'
                    )  # 가로 치수 (치수 문자 제거)
                    annotation(fig, x1 - 5, y1 - 33, color, txt, 'right', 'middle')  # 치수 문자 (다시 정렬)
                    # rho, 철근량, 보강량 정보 (이형철근, 중공철근)
                    txt = '[중공철근]'
                    annotation(fig, x + w + 5, y1 - 63, 'green', txt, 'left', 'middle')
                    txt = f'{sum(F.Ast_total[:]):,.0f} mm<sup>2</sup> <br>(총 보강량)'
                    annotation(fig, x + w + 5, y1 - 99, 'green', txt, 'left', 'middle')
                    txt = f'ρ = {F.rho:.4f} <br>(보강비)'
                    annotation(fig, x + w - 5, y1 - 99, 'green', txt, 'right', 'middle')
        ### 철근
    else:  #! RC & FRP
        ###! 기둥 Pn, e, c & 변형률
        if 'RC' in loc:
            [bgcolor, ep_cu] = ['lightgreen', R.ep_cu]
        if 'hollow' in loc:
            [bgcolor, ep_cu] = ['lightblue', F.ep_cu]
        dc = In.dc[0] * mr
        In.depth = In.hD - In.dc[0]

        if selected_row == None:  # 초기값 (startupFcn)
            txt = 'Select one from A to G below'
            txtPn = 'ϕP<sub>n'
            txte = 'e'
            ee = 0.25 * In.hD
            xee = x + ee * mr
            Pd = 0
            ep_t = 0
            fc = 0
            ft = 0
            if 'RC' in loc:
                ep_t = -0.002  # 인장(-)
            if 'hollow' in loc:
                ep_t = -0.004
            cc = ep_cu / (abs(ep_t) + ep_cu) * In.depth
        else:
            if 'A (P' in selected_row:
                [n, txt] = [0, 'A (Pure Compression)']
            if 'B (M' in selected_row:
                [n, txt] = [1, 'B (Minimum Eccentricity)']
            if 'C (Z' in selected_row:
                [n, txt] = [2, 'C (Zero Tension)']
            if 'D (B' in selected_row:
                [n, txt] = [3, 'D (Balance Point)']
            if 'E ' in selected_row:
                if 'RC' in loc:
                    [n, txt] = [4, 'E (ε<sub>t</sub> = 2.5ε<sub>y</sub>)']
                # if 'hollow' in loc:
                #     [n, txt] = [4, 'E (ε<sub>t</sub> = 0.8ε<sub>fu</sub>)']
            if 'F (P' in selected_row:
                [n, txt] = [5, 'F (Pure Moment)']
            if 'G (P' in selected_row:
                [n, txt] = [6, 'G (Pure Tension)']

            if 'RC' in loc:
                [ee, Pd, cc, ep_t, fc, ft] = [R.e[n], R.Pd[n], R.c[n], R.ep_s[n, 1], R.fs[n, 0], R.fs[n, 1]]
            if 'hollow' in loc:
                [ee, Pd, cc, ep_t, fc, ft] = [F.e[n], F.Pd[n], F.c[n], F.ep_s[n, 1], F.fs[n, 0], F.fs[n, 1]]
            txtPn = f'{Pd:,.1f} kN'
            if ee == np.inf:
                txte = 'inf'
            elif ee == 0:
                txte = ''
            else:
                txte = f'{ee:,.1f} mm'

        y_col2 = y_col + 70
        y_ep = 0.6 * ylim
        shape(fig, 'line', xc, 0, xc, 160, None, 'black', 2, LineStyle='dash')  # Center Line (기둥)
        shape(fig, 'line', x, y_ep, x + w, y_ep, None, 'black', 2)  # 변형률 기준선
        annotation(fig, xc, y_col / 2, 'blue', txt, 'center', 'middle')  # 기둥 속 텍스트 A, B...

        ### 기둥 Pn, e
        xee = xc + ee * mr if abs(ee) < 3 * In.height / 4 else xc + np.sign(ee) * 3 * In.height / 4 * mr
        yee = y_col + (y_col2 - y_col) / 1.8
        arrow_size = 18 * 1.2
        arrow_width = arrow_size / 3  # 화살표 두께 고정 !!!
        color = 'black' if ee >= 0 else 'red'
        if ee != 0:
            [xanchor, x1] = ['center', (xc + xee) / 2] if abs(ee) > In.height / 4 else ['right', (xc + 2 * xee) / 3]
            shape(fig, 'line', xc, yee, xee, yee, None, 'black', 1)
            annotation(fig, x1, yee, color, txte, xanchor, 'middle', bgcolor=bgcolor, size=size17)  # text e

        if Pd >= 0:
            b0 = y_col
            sgn = 1
            shape(fig, 'line', xee, y_col + arrow_size, xee, y_col2, None, color, 3)
        if Pd < 0:
            b0 = y_col2
            sgn = -1
            color = 'red'
            shape(fig, 'line', xee, y_col, xee, y_col2 - arrow_size, None, color, 3)

        annotation(fig, xee, y_col2 + 3, color, txtPn, 'center', 'bottom', bgcolor=bgcolor, size=size17)  # text Pn
        a0 = xee
        a1 = a0 - arrow_width / 2
        a2 = a0 + arrow_width / 2
        b1 = b0 + sgn * arrow_size
        b2 = b1
        fig.add_shape(
            type="path", path=f"M {a0} {b0} L {a1} {b1} L {a2} {b2} Z", fillcolor=color, line=dict(color=color, width=2)
        )  # 화살표  # M = move to, L = line to, Z = close path

        # 압축 영역 기둥에 음영 표시
        shadecolor = 'rgba(0,0,255, 0.3)'
        if selected_row != None:
            xcc = cc * mr if cc < In.height else In.height * mr
            shape(fig, 'rect', x + w - xcc, 0, x + w, y_col, shadecolor, None, 1)
            shape(fig, 'rect', x + 1.2 * w, y_col / 1.2, x + 1.3 * w, y_col / 1.7, shadecolor, None, 1)
            annotation(
                fig,
                x + 1.25 * w,
                y_col / 2.5,
                'black',
                '압축 영역',
                'center',
                'middle',
                bgcolor=shadecolor,
                size=size17,
            )
            annotation(fig, x + 1.25 * w, y_col / 5.8, 'black', f'c = {cc:,.1f} mm', 'center', 'middle')
        ### 기둥 Pn, e

        ### 기둥 변형률, 응력
        fix = R.ep_cu
        y_ep_cu = 80
        xd = x + w - In.depth * mr
        xcc = x + w - cc * mr if cc <= In.depth else xd
        y_ep_c = y_ep - y_ep_cu * ep_cu / fix
        y_ep_t = y_ep - y_ep_cu * ep_t / fix

        if selected_row != None:
            if 'RC' in loc:
                phi = R.phi[n]
                phi_Status = R.phi_Status[n]
            if 'hollow' in loc:
                phi = F.phi[n]
                phi_Status = F.phi_Status[n]
            color = 'black'
            if 'Ten' in phi_Status:
                color = 'red'
            if 'Com' in phi_Status:
                color = 'blue'
            annotation(
                fig, x + w / 2, y_ep + y_ep_cu * 2.3, color, f'ϕ = {phi:,.3f} <br>[{phi_Status}]', 'center', 'middle'
            )  # phi
        if cc != -np.inf:
            annotation(fig, x + w + 5, y_ep - y_ep_cu / 2, 'blue', f'ε<sub>cu</sub> = {ep_cu:,.4f}', 'left', 'middle')
            if selected_row != None:
                annotation(
                    fig,
                    x + w,
                    y_ep - y_ep_cu,
                    'blue',
                    f'f<sub>c</sub> = {fc:,.1f} MPa <br>(압축보강재 응력)',
                    'left',
                    'middle',
                )
        txt = 'Zero Tension' if ep_t == 0 else f'ε<sub>t</sub> = {ep_t:,.4f}'
        if selected_row == None:
            txt = 'ε<sub>t</sub>'
        [yanchor, color] = ['bottom', 'red']
        if ep_t > 0:  # 보강재가 압축(+)일때
            [yanchor, color] = ['top', 'blue']
        if cc != np.inf:
            annotation(fig, x + dc - 5, y_ep, color, txt, 'right', yanchor)
            if 'Zero' not in txt and selected_row != None:
                annotation(
                    fig, x + dc - 5, y_ep - np.sign(ft) * 30, color, f'f<sub>t</sub> = {ft:,.1f} MPa', 'right', yanchor
                )

        # fmt: off
        a0 = xcc;  a1 = x + w
        # fmt: on
        a2 = a1
        a3 = a2
        b0 = y_ep
        b1 = b0
        b2 = y_ep_c
        b3 = b2
        if cc > In.depth:
            [a3, b3] = [xcc, y_ep_t]
        fig.add_shape(
            type="path",
            path=f"M {a0} {b0} L {a1} {b1} L {a2} {b2} L {a3} {b3} Z",
            fillcolor=shadecolor,
            line=dict(color='black', width=2),
        )  # 압축 변형률 삼각형, 사각형 음영  # M = move to, L = line to, Z = close path

        shadecolor = 'rgba(255,0,0, 0.3)'
        a0 = xcc
        a1 = x + dc
        a2 = a1
        a3 = a2
        b0 = y_ep
        b1 = b0
        b2 = y_ep_t
        b3 = b2
        if cc == -np.inf:
            a0 = xd
            a1 = a0
            a2 = x + w
            a3 = a2
            b1 = y_ep_t
            b2 = b1
            b3 = b0
        fig.add_shape(
            type="path",
            path=f"M {a0} {b0} L {a1} {b1} L {a2} {b2} L {a3} {b3} Z",
            fillcolor=shadecolor,
            line=dict(color='black', width=2),
        )  # 인장 변형률 삼각형 음영  # M = move to, L = line to, Z = close path
        ### 기둥 변형률
        ###! 기둥 Pn, e, c & 변형률

    # Update the layout properties
    fig.update_layout(autosize=False, width=fig_width, height=fig_height, margin=dict(l=0, r=2, t=0, b=0))

    fig.update_xaxes(range=[0, xlim])
    fig.update_yaxes(range=[0, ylim])
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    return fig
