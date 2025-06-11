import streamlit as st

class In:
    ok, ng = ':blue[∴ OK] (🆗✅)', ':red[∴ NG] (❌)'
    col_span_ref, col_span_okng, max_width = [1,1], [5,1], '1800px'
    font_h1, font_h2, font_h3, font_h4, font_h5, font_h6 = '32px', '28px', '26px', '24px', '20px', '16px'
    h2, h3, h4, h5, h6 = '## ', '### ', '#### ', '##### ', '###### '
    s1, s2, s3 = '##### $\\quad$', '##### $\\qquad$', '##### $\\quad \\qquad$'

In.border1 = (
    f'<hr style="border-top: 2px solid green; margin-top:30px; margin-bottom:30px; margin-right: -30px">'
)
In.border2 = (
    f'<hr style="border-top: 5px double green; margin-top: 0px; margin-bottom:30px; margin-right: -30px">'
)

def initialize_material_properties():
    """재료 속성 세션 상태 초기화"""
    defaults = {
        'fck': 27.0,
        'fy': 400.0,
        'Es': 200.0,
        'f_fu': 800.0,
        'Ef': 200.0,
        'be': 1000.0,
        'height': 300.0,
        'sb': 150.0,
        'dia': 22.2,
        'dc': 60.0,
        'dia1': 22.2,
        'dc1': 60.0,
    }
    
    for key, default_value in defaults.items():
        if f'material_{key}' not in st.session_state:
            st.session_state[f'material_{key}'] = default_value

def create_material_input(label, key, min_val=10.0, step=1.0, format_str='%f', **kwargs):
    """재료 속성 입력 필드 생성 (세션 상태 관리)"""
    session_key = f'material_{key}'
    
    value = st.number_input(
        label,
        min_value=min_val,
        value=st.session_state.get(session_key, kwargs.get('default', 10.0)),
        step=step,
        format=format_str,
        key=f'input_{key}',
        **{k: v for k, v in kwargs.items() if k != 'default'}
    )
    
    # 세션 상태 업데이트
    st.session_state[session_key] = value
    return value

def Sidebar():
    sb = st.sidebar
    side_border = '<hr style="border-top: 2px solid purple; margin-top:15px; margin-bottom:15px;">'
    h5 = In.h5
    h4 = "#### "
    In.Design_Method = 'KDS-2021'
    
    # 재료 속성 세션 상태 초기화
    initialize_material_properties()

    sb.write('## ', ':blue[[Information : 입력값 📘]]')

    # sb.markdown(side_border, unsafe_allow_html=True)
    # sb.write(h4, ':green[✤ Option]')
    # In.Option = sb.radio('Option', ('기둥 검토', '전단 검토', '사용성 검토'), horizontal=True, label_visibility='collapsed', index=0)   

    # 동적 UI 섹션 (가장 먼저 배치)
    from Column_Sidebar_Fcn import create_column_ui
    create_column_ui(In, sb, side_border, h4)

    # 고정 속성들 (세션 상태로 관리)
    In.RC_Code = 'KDS-2021'
    In.FRP_Code = 'AASHTO-2018'

    sb.markdown(side_border, unsafe_allow_html=True)
    
    # Column Type & PM Diagram (세션 상태 자동 관리)
    col = sb.columns([1, 1.2])
    with col[0]:
        st.write(h4, ':green[✤ Column Type]')
        In.Column_Type = st.radio(
            h5 + '￭ Section Type', 
            ('Tied Column', 'Spiral Column'), 
            key='persistent_column_type', 
            label_visibility='collapsed'
        )
    with col[1]:
        st.write(h4, ':green[✤ PM Diagram Option]')
        In.PM_Type = st.radio(
            'PM Type',
            ('이형철근 \u00a0 vs. \u00a0 중공철근', 'Pₙ-Mₙ \u00a0 vs. \u00a0 ϕPₙ-ϕMₙ'),
            horizontal=True,
            label_visibility='collapsed',
            key='persistent_pm_type',
        )

    sb.markdown(side_border, unsafe_allow_html=True)
    
    # Section Dimensions (세션 상태로 관리)
    sb.write(h4, ':green[✤ Section Dimensions]')
    In.Section_Type = 'Rectangle'

    col = sb.columns([1, 1])
    with col[0]:
        In.be = create_material_input(
            h5 + r'￭ $\bm{{\small{{b_e}} }}$ (단위폭) [mm]',
            'be',
            min_val=10.0,
            step=10.0,
            default=1000.0
        )
    with col[1]:
        In.height = create_material_input(
            h5 + r'￭ $\bm{{\small{{h}} }}$ [mm]',
            'height',
            min_val=10.0,
            step=10.0,
            default=300.0
        )
    
    In.D = 500

    sb.markdown(side_border, unsafe_allow_html=True)
    
    # Material Properties (세션 상태로 관리)
    sb.write(h4, ':green[✤ Material Properties]')
    col = sb.columns(3, gap='medium')
    
    with col[0]:
        st.write(h5, ':blue[✦ 콘크리트]')
        In.fck = create_material_input(
            h5 + r'$\bm{{\small{{f_{ck}}} }}$ [MPa]',
            'fck',
            min_val=10.0,
            step=1.0,
            default=27.0
        )
        
        # Ec 계산 (의존적 값)
        Ec = 8500 * (In.fck + 4) ** (1 / 3) / 1e3
        In.Ec = st.number_input(
            h5 + r'$\bm{{\small{{E_{c}}} }}$ [GPa]',
            min_value=10.0,
            value=Ec,
            step=1.0,
            format='%.1f',
            disabled=True,
            key='persistent_Ec',
        ) * 1e3
        
    with col[1]:
        st.write(h5, ':blue[✦ 이형철근]')
        In.fy = create_material_input(
            h5 + r'$\bm{{\small{{f_{y}}} }}$ [MPa]',
            'fy',
            min_val=10.0,
            step=10.0,
            default=400.0
        )
        In.Es = create_material_input(
            h5 + r'$\bm{{\small{{E_{s}}} }}$ [GPa]',
            'Es',
            min_val=10.0,
            step=10.0,
            default=200.0
        ) * 1e3
        
    with col[2]:
        st.write(h5, ':blue[✦ 중공철근]')
        In.f_fu = create_material_input(
            h5 + r'$\bm{{\small{{f_{y}}} }}$ [MPa]',
            'f_fu',
            min_val=10.0,
            step=10.0,
            default=800.0
        )
        In.Ef = create_material_input(
            h5 + r'$\bm{{\small{{E_{s}}} }}$ [GPa]',
            'Ef',
            min_val=10.0,
            step=10.0,
            default=200.0
        ) * 1e3
        
        In.fy_hollow = In.f_fu
        In.Es_hollow = In.Ef

    # Reinforcement Layer (세션 상태로 관리)
    In.Layer = 1
    In.nD = [8]
    In.nb = [6]
    In.nh = [2]
    
    sb.markdown(side_border, unsafe_allow_html=True)
    sb.write(h4, ':green[✤ Reinforcement Layer (Rebar & FRP)]')
    
    col = sb.columns(2, gap='medium')
    with col[0]:
        In.sb = [create_material_input(
            h5 + r'$\bm{s_b}$ [mm] : $\bm{{\small{b}}}$(단위폭) 방향 보강재 간격',
            'sb',
            min_val=10.0,
            step=10.0,
            default=150.0,
            label_visibility='visible'
        )]

    col = sb.columns(2, gap='large')
    with col[0]:
        st.write(h5, ':blue[✦ 인장측]')
        In.dia = [create_material_input(
            h5 + 'dia [mm] : 보강재 외경',
            'dia',
            min_val=1.0,
            step=1.0,
            default=22.2,
            label_visibility='visible'
        )]
        In.dc = [create_material_input(
            h5 + r'$\bm{d_c}$ [mm] : 피복 두께',
            'dc',
            min_val=1.0,
            step=2.0,
            default=60.0,
            label_visibility='visible'
        )]

    with col[1]:
        st.write(h5, ':blue[✦ 압축측]')
        In.dia1 = [create_material_input(
            h5 + 'dia [mm] : 보강재 외경',
            'dia1',
            min_val=1.0,
            step=1.0,
            default=22.2,
            label_visibility='hidden'
        )]
        In.dc1 = [create_material_input(
            h5 + r'$\bm{d_c}$ [mm] : 피복 두께',
            'dc1',
            min_val=1.0,
            step=2.0,
            default=60.0,
            label_visibility='hidden'
        )]
    
    return In
