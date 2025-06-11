import streamlit as st

def create_column_ui(In, sb, side_border="", h4=""):
    # =================================================================
    # 공통 유틸리티 함수들
    # =================================================================
    def initialize_session_states():
        """모든 세션 상태 초기화"""
        session_defaults = {
            'row_count': 3,
            'serviceability_count': 3,
            'shear_strength_count': 3
        }
        for key, default in session_defaults.items():
            if key not in st.session_state:
                st.session_state[key] = default

    def get_default_value(attr, row_index, defaults):
        """행별 기본값 계산 (공통)"""
        if row_index < len(defaults):
            return defaults[row_index]
        return defaults[2] * (1 + 0.1 * (row_index - 2))

    def adjust_dynamic_lists(attrs, target_size):
        """동적 리스트 크기 조정 (공통)"""
        for attr in attrs:
            if not hasattr(In, attr):
                setattr(In, attr, [])
            current_list = getattr(In, attr)
            if len(current_list) < target_size:
                current_list.extend([0.0] * (target_size - len(current_list)))
            elif len(current_list) > target_size:
                current_list[:] = current_list[:target_size]

    def ensure_list_size(attr, index):
        """특정 속성의 리스트 크기 보장"""
        if not hasattr(In, attr):
            setattr(In, attr, [])
        current_list = getattr(In, attr)
        while len(current_list) <= index:
            current_list.append(0.0)

    def apply_common_styles():
        """공통 CSS 스타일 적용"""
        st.markdown("""
        <style>
            .stButton > button[key*="add_row"] {
                background-color: #28a745;
                color: white;
                border: none;
                border-radius: 20px;
                padding: 0.25rem 0.75rem;
                font-size: 0.9rem;
                font-weight: bold;
                transition: all 0.2s ease;
            }
            .stButton > button[key*="add_row"]:hover {
                background-color: #218838;
                transform: translateY(-1px);
            }
            .stButton > button[key*="delete_"] {
                background-color: #dc3545;
                color: white;
                border: none;
                border-radius: 50%;
                padding: 0.3rem 0.5rem;
                font-size: 0.9rem;
                width: 35px;
                height: 35px;
                transition: all 0.2s ease;
            }
            .stButton > button[key*="delete_"]:hover {
                background-color: #c82333;
                transform: scale(1.1);
            }
            .stButton > button:focus {
                outline: none;
            }
            .stButton > button[key*="add_row"]:focus {
                box-shadow: 0 0 0 0.2rem rgba(40, 167, 69, 0.25);
            }
            .stButton > button[key*="delete_"]:focus {
                box-shadow: 0 0 0 0.2rem rgba(220, 53, 69, 0.25);
            }
            .stNumberInput input {
                text-align: center;
            }
            .stMarkdown div[style*="text-align:center"] {
                display: flex;
                align-items: center;
                justify-content: center;
                height: 50px;
                font-weight: bold;
                color: #2e7d32;
            }
        </style>
        """, unsafe_allow_html=True)

    def create_number_symbol(index):
        """넘버링 기호 생성"""
        num_symbols = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"]
        return num_symbols[index] if index < len(num_symbols) else f"{index+1}"

    def render_section_header(title, button_key, add_callback, show_checkbox=False):
        """섹션 헤더 렌더링 (공통)"""
        cols = sb.columns([2, 1, 1] if show_checkbox else [3, 1])
        with cols[0]:
            if h4:
                st.write(h4, f':green[✤ {title}]')
            else:
                st.markdown(f"#### :green[✤ {title}]")
        if show_checkbox:
            with cols[1]:
                In.check = st.checkbox(':green[선 보이기]', value=True)
            with cols[2]:
                if st.button('➕ 추가', key=button_key):
                    add_callback()
                    st.rerun()
        else:
            with cols[1]:
                if st.button('➕ 추가', key=button_key):
                    add_callback()
                    st.rerun()

    def render_table_header(spec_items, col_width, special_header=None):
        """테이블 헤더 렌더링 (공통)"""
        header_cols = sb.columns(col_width, gap='small')
        header_cols[0].write("")
        for i, (attr, (tex, unit)) in enumerate(spec_items):
            header_cols[i + 1].markdown(rf"$\bm{{\small{{{tex}}}}}$ {unit}")
        if special_header and len(header_cols) > len(spec_items) + 1:
            header_cols[-1].markdown(special_header)

    def render_input_field(col, attr, defaults, step, row_index, section_key, row_count):
        """입력 필드 렌더링 (공통)"""
        default_val = get_default_value(attr, row_index, defaults)
        val = col.number_input(
            label="",
            min_value=0.0,
            value=default_val,
            step=step,
            format="%.0f",
            key=f"{attr}_{section_key}_{row_index}_{row_count}",
            label_visibility="collapsed"
        )
        ensure_list_size(attr, row_index)
        getattr(In, attr)[row_index] = val

    def render_delete_button(col, row_index, row_count, section_key, delete_callback):
        """삭제 버튼 렌더링 (공통)"""
        if st.button('🗑️', key=f'delete_{section_key}_row_{row_index}_{row_count}',
                help=f'{row_index+1}번 행 삭제'):
            delete_callback(row_index)
            st.rerun()

    def create_generic_section(config):
        """범용 섹션 생성 함수"""
        if side_border:
            sb.markdown(side_border, unsafe_allow_html=True)

        render_section_header(
            config['title'], 
            config['button_key'], 
            config['add_callback'], 
            config.get('show_checkbox', False)
        )
        adjust_dynamic_lists(config['attrs'], config['row_count'])
        render_table_header(
            config['spec_items'], 
            config['col_width'], 
            config.get('special_header')
        )

        # placeholder 이름에 따라 In.<placeholder_name> 초기화
        if config.get('has_placeholder'):
            placeholder_name = config['placeholder_name']
            if not hasattr(In, placeholder_name):
                setattr(In, placeholder_name, [None] * config['row_count'])
            else:
                current = getattr(In, placeholder_name)
                if len(current) < config['row_count']:
                    current.extend([None] * (config['row_count'] - len(current)))
                elif len(current) > config['row_count']:
                    setattr(In, placeholder_name, current[:config['row_count']])

        for i in range(config['row_count']):
            cols = sb.columns(config['col_width'])
            # 넘버링
            cols[0].markdown(
                f"<div style='text-align:center; font-size:1.4em;'>{create_number_symbol(i)}</div>",
                unsafe_allow_html=True
            )
            col_idx = 1
            for attr, (tex, unit, defaults, step) in config['spec_full'].items():
                col = cols[col_idx]
                # placeholder 컬럼이면 empty()를 저장
                if config.get('has_placeholder') and attr == config.get('placeholder_attr'):
                    placeholder_name = config['placeholder_name']
                    getattr(In, placeholder_name)[i] = col.empty()
                else:
                    render_input_field(
                        col, attr, defaults, step, i, 
                        config['section_key'], config['row_count']
                    )
                col_idx += 1

            if config.get('special_column_func'):
                config['special_column_func'](cols, i, config)

    # =================================================================
    # 강도 검토 섹션 함수들 (Vu 제거됨)
    # =================================================================
    def add_strength_row():
        """강도 검토 행 추가"""
        st.session_state.row_count += 1

    def delete_strength_row(row_index):
        """강도 검토 행 삭제"""
        if st.session_state.row_count > 3 and row_index >= 3:
            for attr in ['Pu', 'Mu', 'safe_RC', 'safe_FRP', 'Pd_RC', 'Pd_FRP', 'Md_RC', 'Md_FRP']:
                current_list = getattr(In, attr, [])
                if row_index < len(current_list):
                    current_list.pop(row_index)
            st.session_state.row_count -= 1

    def strength_special_column(cols, row_index, config):
        """강도 검토 특수 컬럼 처리 (삭제 버튼만)"""
        if config['row_count'] > 3 and row_index >= 3:
            with cols[4]:
                render_delete_button(cols[4], row_index, config['row_count'], 
                                config['section_key'], delete_strength_row)

    def create_strength_section():
        """기둥 강도 검토 섹션 전체 생성"""
        strength_spec_full = {
            "Pu":    ("P_u",    "[kN]",   [2000.0, 3000.0, 6000.0], 200.0),
            "Mu":    ("M_u",    "[kN·m]", [220.0,  250.0,  300.0],  20.0),
            "검토":  ("검토",   "",       [0.0,    0.0,    0.0],    0.0),  # placeholder
        }
        strength_spec_items = [
            ("Pu",   ("P_u", "[kN]")),
            ("Mu",   ("M_u", "[kN·m]")),
            ("검토", ("검토", "")),
        ]
        special_header = "**삭제**" if st.session_state.row_count > 3 else None

        config = {
            'title'             : '기둥 강도 검토',
            'button_key'        : 'add_row_btn',
            'add_callback'      : add_strength_row,
            'show_checkbox'     : True,
            'attrs'             : ['Pu', 'Mu', 'safe_RC', 'safe_FRP', 'Pd_RC', 'Pd_FRP', 'Md_RC', 'Md_FRP'],
            'row_count'         : st.session_state.row_count,
            'col_width'         : [0.3, 1.2, 1, 0.8, 1],
            'spec_items'        : strength_spec_items,
            'spec_full'         : strength_spec_full,
            'special_header'    : special_header,
            'section_key'       : 'strength',
            'has_placeholder'   : True,
            'placeholder_attr'  : '검토',                    # "검토" 컬럼 placeholder
            'placeholder_name'  : 'placeholder_strength',    # In.placeholder_strength
            'special_column_func': strength_special_column,
        }
        create_generic_section(config)

    # =================================================================
    # 전단 강도 검토 섹션 함수들 (검토 컬럼 추가)
    # =================================================================
    def add_shear_strength_row():
        """전단 강도 검토 행 추가"""
        st.session_state.shear_strength_count += 1

    def delete_shear_strength_row(row_index):
        """전단 강도 검토 행 삭제"""
        if st.session_state.shear_strength_count > 3 and row_index >= 3:
            for attr in ['Pu_shear', 'Vu']:
                current_list = getattr(In, attr, [])
                if row_index < len(current_list):
                    current_list.pop(row_index)
            st.session_state.shear_strength_count -= 1

    def shear_strength_special_column(cols, row_index, config):
        """전단 강도 검토 특수 컬럼 처리 (삭제 버튼)"""
        if config['row_count'] > 3 and row_index >= 3:
            with cols[4]:
                render_delete_button(cols[4], row_index, config['row_count'], 
                                config['section_key'], delete_shear_strength_row)

    def create_shear_strength_section():
        """전단 강도 검토 섹션 전체 생성"""
        shear_strength_spec_full = {
            "Pu_shear": ("P_u", "[kN]", [2000.0, 3000.0, 6000.0], 200.0),
            "Vu":      ("V_u", "[kN]", [150.0,  180.0,  250.0],  10.0),
            "검토":    ("검토", "",     [0.0,    0.0,    0.0],   0.0),  # placeholder
        }
        shear_strength_spec_items = [
            ("Pu_shear", ("P_u", "[kN]")),
            ("Vu",      ("V_u", "[kN]")),
            ("검토",    ("검토", "")),
        ]
        col_width = [0.3, 1.2, 1.2, 0.8, 1]
        special_header = "**삭제**" if st.session_state.shear_strength_count > 3 else None

        config = {
            'title'             : '전단 검토',
            'button_key'        : 'add_row_btn2',
            'add_callback'      : add_shear_strength_row,
            'show_checkbox'     : False,
            'attrs'             : ['Pu_shear', 'Vu'],
            'row_count'         : st.session_state.shear_strength_count,
            'col_width'         : col_width,
            'spec_items'        : shear_strength_spec_items,
            'spec_full'         : shear_strength_spec_full,
            'special_header'    : special_header,
            'section_key'       : 'shear_strength',
            'has_placeholder'   : True,
            'placeholder_attr'  : '검토',                  # "검토" 컬럼 placeholder
            'placeholder_name'  : 'placeholder_shear',     # In.placeholder_shear
            'special_column_func': shear_strength_special_column,
        }
        create_generic_section(config)

    # =================================================================
    # Serviceability 검토 섹션 함수들 (검토 컬럼 추가)
    # =================================================================
    def add_serviceability_row():
        """Serviceability 검토 행 추가"""
        st.session_state.serviceability_count += 1

    def delete_serviceability_row(row_index):
        """Serviceability 검토 행 삭제"""
        if st.session_state.serviceability_count > 3 and row_index >= 3:
            for attr in ['P0', 'M0']:
                current_list = getattr(In, attr, [])
                if row_index < len(current_list):
                    current_list.pop(row_index)
            st.session_state.serviceability_count -= 1

    def serviceability_special_column(cols, row_index, config):
        """Serviceability 검토 특수 컬럼 처리 (삭제 버튼)"""
        if config['row_count'] > 3 and row_index >= 3:
            with cols[4]:
                render_delete_button(cols[4], row_index, config['row_count'], 
                                config['section_key'], delete_serviceability_row)

    def create_serviceability_section():
        """Serviceability 검토 섹션 전체 생성"""
        serviceability_spec_full = {
            "P0":   ("P_0", "[kN]",   [0.0, 100.0, 200.0], 10.0),
            "M0":   ("M_0", "[kN·m]", [100.0,  150.0,  200.0],  10.0),
            "검토": ("검토", "",      [0.0,    0.0,    0.0],   0.0),  # placeholder
        }
        serviceability_spec_items = [
            ("P0",   ("P_0", "[kN]")),
            ("M0",   ("M_0", "[kN·m]")),
            ("검토", ("검토", "")),
        ]
        col_width = [0.3, 1.2, 1.2, 0.8, 1]
        special_header = "**삭제**" if st.session_state.serviceability_count > 3 else None

        config = {
            'title'             : '사용성 검토',
            'button_key'        : 'add_row_btn1',
            'add_callback'      : add_serviceability_row,
            'show_checkbox'     : False,
            'attrs'             : ['P0', 'M0'],
            'row_count'         : st.session_state.serviceability_count,
            'col_width'         : col_width,
            'spec_items'        : serviceability_spec_items,
            'spec_full'         : serviceability_spec_full,
            'special_header'    : special_header,
            'section_key'       : 'serviceability',
            'has_placeholder'   : True,
            'placeholder_attr'  : '검토',                          # "검토" 컬럼 placeholder
            'placeholder_name'  : 'placeholder_serviceability',   # In.placeholder_serviceability
            'special_column_func': serviceability_special_column,
        }
        create_generic_section(config)

    # =================================================================
    # 메인 실행부
    # =================================================================
    initialize_session_states()
    apply_common_styles()

    # 라디오 옵션에 따라 섹션 렌더링
    option = getattr(In, 'Option', None)
    
    # 1. 기둥 강도 검토 섹션 (placeholder_strength)
    create_strength_section()  
    # 2. 전단 검토 섹션 (placeholder_shear)
    create_shear_strength_section()
    # 3. 사용성 검토 섹션 (placeholder_serviceability)
    create_serviceability_section()
