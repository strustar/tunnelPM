# Column_Sidebar.py (세션 충돌 수정 / 사용법·레이아웃 개선 버전)
import streamlit as st
import json
import os
import re
from datetime import datetime
from pathlib import Path

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

# ============================================
# 프리셋 파일 관리 (개별 JSON 파일)
# ============================================
def _preset_dir() -> str:
    d = Path.home() / "Downloads" / "column_presets"
    d.mkdir(parents=True, exist_ok=True)
    return str(d)

def _slugify(name: str) -> str:
    s = name.strip()
    s = re.sub(r"[^\w\-\.\s]", "_", s, flags=re.UNICODE)
    s = re.sub(r"\s+", "_", s)
    return s[:80] if len(s) > 80 else s

def make_preset_filepath(preset_name: str) -> str:
    return os.path.join(_preset_dir(), f"{_slugify(preset_name)}.json")

def list_preset_files() -> list[Path]:
    d = Path(_preset_dir())
    return sorted(list(d.glob("*.json")))

def next_preset_name() -> str:
    files = list_preset_files()
    nums = []
    for p in files:
        m = re.match(r'^프리셋_(\d+)$', p.stem)
        if m:
            try: nums.append(int(m.group(1)))
            except: pass
    n = max(nums)+1 if nums else 1
    return f"프리셋_{n}"

def _snapshot_current_values() -> dict:
    ss = st.session_state
    row_count   = ss.get('row_count', 3)
    shear_count = ss.get('shear_strength_count', 3)
    serv_count  = ss.get('serviceability_count', 3)

    scalar_keys = [
        'fck','fy','Es','f_fu','Ef','be','height',
        'sb','dia','dc','dia1','dc1','column_type','pm_type'
    ]
    scalar_values = {k: ss.get(k) for k in scalar_keys}

    dynamic = {
        'strength': {
            'row_count': row_count,
            'Pu': [ss.get(f"Pu_strength_{i}_{row_count}") for i in range(row_count)],
            'Mu': [ss.get(f"Mu_strength_{i}_{row_count}") for i in range(row_count)],
        },
        'shear_strength': {
            'row_count': shear_count,
            'Pu_shear': [ss.get(f"Pu_shear_shear_strength_{i}_{shear_count}") for i in range(shear_count)],
            'Vu': [ss.get(f"Vu_shear_strength_{i}_{shear_count}") for i in range(shear_count)],
        },
        'serviceability': {
            'row_count': serv_count,
            'P0': [ss.get(f"P0_serviceability_{i}_{serv_count}") for i in range(serv_count)],
            'M0': [ss.get(f"M0_serviceability_{i}_{serv_count}") for i in range(serv_count)],
        },
    }
    return {
        'preset_name': ss.get('new_preset_name', ''),
        'saved_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'values': scalar_values,
        'dynamic': dynamic,
    }

def save_preset_to_file(preset_name: str) -> bool:
    try:
        fp = make_preset_filepath(preset_name)
        snap = _snapshot_current_values()
        snap['preset_name'] = preset_name
        with open(fp, 'w', encoding='utf-8') as f:
            json.dump(snap, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        st.sidebar.error(f"프리셋 저장에 실패했습니다: {e}")
        return False

def load_preset_from_dict(preset_data: dict) -> bool:
    try:
        ss = st.session_state
        values  = preset_data.get('values', {})
        dynamic = preset_data.get('dynamic', {})

        for k, v in values.items():
            ss[k] = v

        def _apply(section_key, section_obj):
            rc = section_obj.get('row_count', 3)
            for attr, arr in section_obj.items():
                if attr == 'row_count': continue
                for i, v in enumerate(arr):
                    if v is None: continue
                    ss[f"{attr}_{section_key}_{i}_{rc}"] = v
            if section_key == 'strength':        ss['row_count'] = rc
            elif section_key == 'shear_strength': ss['shear_strength_count'] = rc
            elif section_key == 'serviceability': ss['serviceability_count'] = rc

        if 'strength' in dynamic:       _apply('strength',       dynamic['strength'])
        if 'shear_strength' in dynamic: _apply('shear_strength', dynamic['shear_strength'])
        if 'serviceability' in dynamic: _apply('serviceability', dynamic['serviceability'])
        return True
    except Exception as e:
        st.sidebar.error(f"프리셋 로드에 실패했습니다: {e}")
        return False

def load_preset_from_file(file_path: str) -> bool:
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            pdata = json.load(f)
        return load_preset_from_dict(pdata)
    except Exception as e:
        st.sidebar.error(f"프리셋 파일 로드에 실패했습니다: {e}")
        return False

# ============================================
# 초기 세션 상태
# ============================================
def initialize_material_properties():
    defaults = {
        'fck': 27.0, 'fy': 400.0, 'Es': 200.0,
        'f_fu': 800.0, 'Ef': 200.0,
        'be': 1000.0, 'height': 300.0,
        'sb': 150.0, 'dia': 22.2, 'dc': 60.0,
        'dia1': 22.2, 'dc1': 60.0,
        'column_type': 'Tied Column',
        'pm_type': '이형철근 \u00a0 vs. \u00a0 중공철근',
        'row_count': 3, 'shear_strength_count': 3, 'serviceability_count': 3,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v
    if "has_result" not in st.session_state:
        st.session_state.has_result = False
    if "run_calculation" not in st.session_state:
        st.session_state.run_calculation = True

# ============================================
# 사이드바 본체
# ============================================
def Sidebar():
    sb = st.sidebar
    side_border = '<hr style="border-top: 2px solid purple; margin-top:15px; margin-bottom:15px;">'
    h5 = In.h5
    h4 = "#### "

    # --- 초기화 ---
    initialize_material_properties()

    # --- 저장 직후 이름 자동증가 처리(플래그 방식) ---
    if '_bump_preset_name' in st.session_state and st.session_state['_bump_preset_name']:
        # 다음 렌더 사이클 초기에 키 갱신 → 위젯 생성 전에 처리됨
        st.session_state['new_preset_name'] = next_preset_name()
        st.session_state['_bump_preset_name'] = False
    # 최초 진입 시 기본 이름 준비
    if 'new_preset_name' not in st.session_state:
        st.session_state['new_preset_name'] = next_preset_name()

    # --- 사이드바 텍스트 넘침 방지 CSS ---
    st.markdown("""
    <style>
    .sidebar-wrap, .sidebar-wrap * {
        white-space: normal !important;
        overflow-wrap: anywhere !important;
        word-break: break-word !important;
        line-height: 1.36;
    }
    </style>
    """, unsafe_allow_html=True)

    # ===== 계산 실행 사용법 (버튼 위) =====
    with sb.expander('📖 **계산 실행 사용법**', expanded=False):
        st.markdown("""
<div class="sidebar-wrap">

### 🎯 간단한 사용법

**1단계: 입력값 설정**  
- 아래 입력 필드에서 콘크리트 강도, 철근 강도 등을 입력하세요  
- 테이블에서 ➕ 버튼으로 행을 추가할 수 있습니다  

**2단계: 계산 실행**  
- 모든 값을 입력한 후 **🚀 계산 실행** 버튼을 클릭하세요  
- 메인 화면에 결과가 표시됩니다  

**3단계: 결과 확인**  
- PM 상관도, 강도 검토, 전단 검토 등을 탭에서 확인하세요  
- 엑셀 파일로 다운로드할 수 있습니다  

---

### 💡 팁
- 입력값을 변경해도 결과는 그대로 유지됩니다  
- 새로 계산하려면 다시 **🚀** 버튼을 누르세요
</div>
        """, unsafe_allow_html=True)

    # ===== 계산 실행 버튼 =====
    if sb.button("🚀 계산 실행", use_container_width=True, type="primary"):
        st.session_state.run_calculation = True
        st.session_state.has_result = False
    In.should_run = st.session_state.run_calculation
    if st.session_state.run_calculation:
        st.session_state.run_calculation = False

    sb.markdown(side_border, unsafe_allow_html=True)

    # ================================
    # 💾 프리셋 관리 (불러오기/저장)
    # ================================
    with sb.expander('💾 **프리셋 관리**', expanded=False):
        st.caption(f"📁 저장 경로: `{_preset_dir()}`")

        tabs = st.tabs(['📥 불러오기', '💾 저장'])

        # --- 불러오기 ---
        with tabs[0]:
            st.info("아래 목록에서 프리셋을 선택하신 뒤 **불러오기**를 눌러 주세요.")
            st.markdown("---")
            files = list_preset_files()
            if files:
                display_names = [p.stem for p in files]  # .json 숨김
                sel = st.selectbox("저장된 프리셋 선택", display_names, key='preset_select_from_folder')
                if st.button("✅ 불러오기", use_container_width=True, key='btn_load_from_folder'):
                    fp = str(files[display_names.index(sel)])
                    if load_preset_from_file(fp):
                        st.success(f"✅ '{sel}' 프리셋을 불러왔습니다.")
                        st.rerun()
            else:
                st.info("저장된 프리셋 파일이 없습니다. 먼저 저장해 주세요.")

        # --- 저장 ---
        with tabs[1]:
            preset_name = st.text_input(
                '프리셋 이름 (기본값 권장: 프리셋_N)',
                value=st.session_state['new_preset_name'],
                key='new_preset_name',
                placeholder='예: 프리셋_1, 프리셋_2 …'
            )
            st.caption("📝 UI에서는 확장자를 숨기지만, 실제 파일은 `.json`으로 저장됩니다.")

            if st.button('💾 저장', use_container_width=True, type='primary', key='save_btn_simple'):
                name = (preset_name or "").strip()
                if not name:
                    st.error('❌ 프리셋 이름을 입력해 주세요.')
                else:
                    if save_preset_to_file(name):
                        st.success(f"💾 '{name}' 저장을 완료했습니다.")
                        # 바로 키를 바꾸지 말고, 플래그만 세우고 재실행 → 다음 사이클 초기에 안전 갱신
                        st.session_state['_bump_preset_name'] = True
                        st.rerun()

        # --- 프리셋 사용법 (넘침 방지 래퍼 포함) ---
        st.markdown("---")
        with sb.expander("ℹ️ 프리셋 사용법", expanded=False):
            st.markdown(f"""
<div class="sidebar-wrap">

- **저장**: 현재 사이드바 입력값으로 저장합니다.  
  경로: Downloads/column_presets/`프리셋_N`.json

- **불러오기**: 목록에서 **프리셋 이름**을 선택 → **불러오기** 버튼을 누르면 저장 당시의 값이 그대로 복원됩니다(행 개수 포함).

- **이름 규칙**: 기본 이름은 항상 `프리셋_N`(자동 증가)이며, 원하시면 직접 다른 이름으로 저장하셔도 됩니다.

- **확장자 표시**: UI에서는 `.json`을 숨기지만, 실제 파일은 `.json`으로 저장됩니다.

</div>
            """, unsafe_allow_html=True)

    sb.markdown(side_border, unsafe_allow_html=True)

    # ============================================
    # 나머지 입력 UI
    # ============================================
    In.Design_Method = 'KDS-2021'
    sb.write('## ', ':blue[[Information : 입력값 📘]]')

    from Column_Sidebar_Fcn import create_column_ui
    create_column_ui(In, sb, side_border, h4)

    In.RC_Code = 'KDS-2021'
    In.FRP_Code = 'AASHTO-2018'

    sb.markdown(side_border, unsafe_allow_html=True)

    # Column Type & PM Diagram
    col = sb.columns([1, 1.2])
    with col[0]:
        st.write(h4, ':green[✤ Column Type]')
        In.Column_Type = st.radio(
            h5 + '￭ Section Type',
            ('Tied Column', 'Spiral Column'),
            key='column_type',
            label_visibility='collapsed'
        )
    with col[1]:
        st.write(h4, ':green[✤ PM Diagram Option]')
        In.PM_Type = st.radio(
            'PM Type',
            ('이형철근 \u00a0 vs. \u00a0 중공철근', 'Pₙ-Mₙ \u00a0 vs. \u00a0 ϕPₙ-ϕMₙ'),
            horizontal=True,
            label_visibility='collapsed',
            key='pm_type',
        )

    sb.markdown(side_border, unsafe_allow_html=True)

    # Section Dimensions
    sb.write(h4, ':green[✤ Section Dimensions]')
    In.Section_Type = 'Rectangle'
    col = sb.columns([1, 1])
    with col[0]:
        In.be = st.number_input(
            h5 + r'￭ $\bm{{\small{{b_e}} }}$ (단위폭) [mm]',
            min_value=10.0,
            step=10.0,
            key='be',
            format='%f'
        )
    with col[1]:
        In.height = st.number_input(
            h5 + r'￭ $\bm{{\small{{h}} }}$ [mm]',
            min_value=10.0,
            step=10.0,
            key='height',
            format='%f'
        )
    In.D = 500

    sb.markdown(side_border, unsafe_allow_html=True)

    # Material Properties
    sb.write(h4, ':green[✤ Material Properties]')
    col = sb.columns(3, gap='medium')

    with col[0]:
        st.write(h5, ':blue[✦ 콘크리트]')
        In.fck = st.number_input(
            h5 + r'$\bm{{\small{{f_{ck}}} }}$ [MPa]',
            min_value=10.0,
            step=1.0,
            key='fck',
            format='%f'
        )
        Ec = 8500 * (In.fck + 4) ** (1 / 3) / 1e3
        In.Ec = st.number_input(
            h5 + r'$\bm{{\small{{E_{c}}} }}$ [GPa]',
            min_value=10.0,
            value=Ec,
            step=1.0,
            format='%.1f',
            disabled=True,
            key='Ec',
        ) * 1e3

    with col[1]:
        st.write(h5, ':blue[✦ 이형철근]')
        In.fy = st.number_input(
            h5 + r'$\bm{{\small{{f_{y}}} }}$ [MPa]',
            min_value=10.0,
            step=10.0,
            key='fy',
            format='%f'
        )
        In.Es = st.number_input(
            h5 + r'$\bm{{\small{{E_{s}}} }}$ [GPa]',
            min_value=10.0,
            step=10.0,
            key='Es',
            format='%f'
        ) * 1e3

    with col[2]:
        st.write(h5, ':blue[✦ 중공철근]')
        In.f_fu = st.number_input(
            h5 + r'$\bm{{\small{{f_{y}}} }}$ [MPa]',
            min_value=10.0,
            step=10.0,
            key='f_fu',
            format='%f'
        )
        In.Ef = st.number_input(
            h5 + r'$\bm{{\small{{E_{s}}} }}$ [GPa]',
            min_value=10.0,
            step=10.0,
            key='Ef',
            format='%f'
        ) * 1e3
        In.fy_hollow = In.f_fu
        In.Es_hollow = In.Ef

    In.Layer = 1
    In.nD = [8]; In.nb = [6]; In.nh = [2]

    sb.markdown(side_border, unsafe_allow_html=True)
    sb.write(h4, ':green[✤ Reinforcement Layer (Rebar & FRP)]')

    col = sb.columns(2, gap='medium')
    with col[0]:
        In.sb = [st.number_input(
            h5 + r'$\bm{s_b}$ [mm] : $\bm{{\small{b}}}$(단위폭) 방향 보강재 간격',
            min_value=10.0,
            step=10.0,
            key='sb',
            format='%f'
        )]

    col = sb.columns(2, gap='large')
    with col[0]:
        st.write(h5, ':blue[✦ 인장측]')
        In.dia = [st.number_input(
            h5 + 'dia [mm] : 보강재 외경',
            min_value=1.0,
            step=1.0,
            key='dia',
            format='%f'
        )]
        In.dc = [st.number_input(
            h5 + r'$\bm{d_c}$ [mm] : 피복 두께',
            min_value=1.0,
            step=2.0,
            key='dc',
            format='%f'
        )]
    with col[1]:
        st.write(h5, ':blue[✦ 압축측]')
        In.dia1 = [st.number_input(
            h5 + 'dia [mm] : 보강재 외경',
            min_value=1.0,
            step=1.0,
            key='dia1',
            format='%f',
            label_visibility='hidden'
        )]
        In.dc1 = [st.number_input(
            h5 + r'$\bm{d_c}$ [mm] : 피복 두께',
            min_value=1.0,
            step=2.0,
            key='dc1',
            format='%f',
            label_visibility='hidden'
        )]

    return In
