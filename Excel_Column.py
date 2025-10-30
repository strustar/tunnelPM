import numpy as np
import xlsxwriter
import traceback

def RC_and_AASHTO(Section_Type, Reinforcement_Type, beta1, c, eta, fck, Layer, ni, ep_si, ep_cu, dsi, fsi, Es, fy, Fsi, Asi, *bhD):
    """
    Calculates the nominal axial force (P) and moment (M) capacity of a reinforced concrete section.
    Handles 1D or 2D input arrays for reinforcement properties.
    """
    a = beta1 * c
    hD = bhD[0]
    if 'Rectangle' in Section_Type:
        [_, b, h] = bhD
        Ac = a * b if a < h else h * b
        y_bar = (h / 2) - (a / 2) if a < h else 0
    else:
        b, h = 0, 0
        Ac = 0
        y_bar = 0
    Cc = eta * (0.85 * fck) * Ac / 1e3
    M = 0
    
    # Input Array Handling
    expected_shape = (Layer, ni[0])
    def _ensure_2d(arr, name):
        arr = np.asarray(arr)
        if arr.size == np.prod(expected_shape) and arr.shape != expected_shape:
             return arr.reshape(expected_shape)
        elif arr.shape != expected_shape:
             raise ValueError(f"Array '{name}' shape {arr.shape} is incompatible with expected {expected_shape}")
        return arr
    
    try:
        dsi = _ensure_2d(dsi, 'dsi')
        Asi = _ensure_2d(Asi, 'Asi')
        ep_si = _ensure_2d(ep_si, 'ep_si')
        fsi = _ensure_2d(fsi, 'fsi')
        Fsi = _ensure_2d(Fsi, 'Fsi')
    except ValueError as e:
         print(f"Error in RC_and_AASHTO input shapes: {e}")
         return 0, 0
    
    for L in range(Layer):
        for i in range(ni[L]):
            dsi_val = dsi[L, i]
            Asi_val = Asi[L, i]
            if c <= 0: continue
            current_ep_si = ep_cu * (c - dsi_val) / c
            current_fsi = Es * current_ep_si
            current_fsi = np.clip(current_fsi, -fy, fy)
            
            ep_si[L, i] = current_ep_si
            fsi[L, i] = current_fsi
            
            if 'RC' in Reinforcement_Type or 'hollow' in Reinforcement_Type:
                if c >= dsi_val:
                    current_Fsi = Asi_val * (current_fsi - eta * 0.85 * fck) / 1e3
                else:
                    current_Fsi = Asi_val * current_fsi / 1e3
            else:
                 current_Fsi = 0
            
            Fsi[L, i] = current_Fsi
            M = M + current_Fsi * (hD / 2 - dsi_val)
    
    P = np.sum(Fsi) + Cc
    M = (M + Cc * y_bar) / 1e3
    return P, M

def safe_extract(obj, attr_name, default=[]):
    values = getattr(obj, attr_name, default)
    if hasattr(values, 'tolist'): return values.tolist()
    elif isinstance(values, (list, tuple)): return list(values)
    elif isinstance(values, (int, float, np.number)): return [values]
    elif values is None: return []
    else:
        try: return [float(values)]
        except (ValueError, TypeError): return default

def create_column_sheet(wb, In, R, F):
    """기둥 강도 검토 시트 생성 - 최적화된 폰트 및 열 너비"""
    column_ws = wb.add_worksheet('기둥 강도 검토')
    
    # ─── 스타일 정의 (폰트 크기 최적화) ───────────────────────────
    base_font = {'font_name': 'Noto Sans KR', 'font_size': 12, 'border': 1, 'valign': 'vcenter'}
    
    styles = {
        'title': {**base_font, 'bold': True, 'font_size': 16, 'bg_color': '#1e40af', 'font_color': 'white', 'border': 3, 'align': 'center'},
        'main_header': {**base_font, 'bold': True, 'font_size': 14, 'bg_color': '#2563eb', 'font_color': 'white', 'border': 2, 'align': 'center'},
        'common_section': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#155e75', 'font_color': '#e0f2fe', 'align': 'center'},
        'section': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#1e3a8a', 'font_color': 'white', 'align': 'center'},
        'sub_header': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#3b82f6', 'font_color': 'white', 'align': 'center'},
        
        # 기본 텍스트 스타일
        'label': {**base_font, 'font_size': 12, 'bold': True, 'align': 'left'},
        'value': {**base_font, 'bold': True, 'font_size': 12, 'align': 'center'},
        'number': {**base_font, 'bold': True, 'font_size': 12, 'num_format': '#,##0.0', 'align': 'center'},
        'unit': {**base_font, 'font_size': 12, 'align': 'center', 'bold': True},
        'combo': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#fef3c7', 'font_color': '#92400e', 'align': 'center'},
        
        # OK/NG 스타일
        'ok': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#ffffff', 'font_color': '#0000ff', 'border': 2, 'align': 'center'},
        'ng': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#fee2e2', 'font_color': '#dc2626', 'border': 2, 'align': 'center'},
        'final_ok': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#d1fae5', 'font_color': '#065f46', 'border': 3, 'align': 'center'},
        'final_ng': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#fee2e2', 'font_color': '#991b1b', 'border': 3, 'align': 'center'},
        
        # 상세 계산 박스 스타일
        'calc_title': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#dbeafe', 'font_color': '#1e40af', 'border': 2, 'align': 'center'},
        'calc_section': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#fafafc', 'align': 'left', 'text_wrap': True},
        'calc_text': {**base_font, 'font_size': 11, 'bg_color': '#fafafc', 'align': 'left', 'text_wrap': True, 'bold': True},
        'calc_formula': {**base_font, 'font_size': 11, 'bg_color': '#fafafc', 'align': 'left', 'italic': True, 'font_color': '#000000', 'bold': True},
        'calc_value': {**base_font, 'font_size': 11, 'bg_color': '#f0f9ff', 'align': 'center', 'bold': True, 'num_format': '#,##0.0'},
        'calc_value_strain': {**base_font, 'font_size': 11, 'bg_color': '#f0f9ff', 'align': 'center', 'bold': True, 'num_format': '0.00000'},
        
        # 테이블 스타일
        'table_header': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#3b82f6', 'font_color': 'white', 'align': 'center'},
        'table_data': {**base_font, 'font_size': 11, 'align': 'center', 'bold': True},
        'note_content': {**base_font, 'font_size': 11, 'align': 'left', 'text_wrap': True, 'bg_color': '#fafafc', 'valign': 'top', 'bold': True},
        
        # 이형/중공철근 구분 강조
        'rc_marker': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#f59e0b', 'font_color': 'white', 'align': 'center', 'border': 2},
        'hollow_marker': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#10b981', 'font_color': 'white', 'align': 'center', 'border': 2},
    }
    
    formats = {name: wb.add_format(props) for name, props in styles.items()}
    
    # ─── 컬럼 너비 설정 (최적화) ────────────────────────────────
    # 좌우 대칭 레이아웃: 이형철근 (A-G) | 구분선 | 중공철근 (I-O)
    column_ws.set_column('A:A', 21)    # 왼쪽 여백
    column_ws.set_column('B:C', 15)   # 이형 라벨
    column_ws.set_column('D:D', 10)   # 이형 단위
    column_ws.set_column('E:E', 15)   # 이형 라벨
    column_ws.set_column('F:G', 10)    # 이형 우측 여백
    
    column_ws.set_column('H:H', 10)    # 중앙 구분선
    
    column_ws.set_column('I:I', 21)    # 중공 좌측 여백
    column_ws.set_column('J:K', 15)   # 중공 라벨
    column_ws.set_column('L:L', 10)   # 중공 단위
    column_ws.set_column('M:N', 15)    # 중공 우측 여백
    column_ws.set_column('O:O', 10)   
    
    row = 0
    max_col = 14  # O 컬럼 인덱스
    
    # ═══════════════════════════════════════════════════════════
    # 섹션 1: 메인 타이틀
    # ═══════════════════════════════════════════════════════════
    column_ws.merge_range(row, 0, row, max_col, '🏗️ 기둥 강도 검토 보고서', formats['title'])
    column_ws.set_row(row, 35)
    row += 2
    
    # ═══════════════════════════════════════════════════════════
    # 섹션 2: 공통 설계 조건
    # ═══════════════════════════════════════════════════════════
    column_ws.merge_range(row, 0, row, max_col, '🏗️ 공통 설계 조건', formats['common_section'])
    column_ws.set_row(row, 28)
    row += 1
    
    section_data = [
        ['📐 단면 제원', [
            ['📏 단위폭 be', getattr(In, 'be', 1000), 'mm'],
            ['📏 단면 두께 h', getattr(In, 'height', 300), 'mm'],
            ['📐 공칭 철근간격 s', getattr(In, 'sb', [150.0])[0], 'mm']
        ]],
        ['🏭 콘크리트 재료', [
            ['💪 압축강도 fck', getattr(In, 'fck', 27.0), 'MPa'],
            ['⚡ 탄성계수 Ec', getattr(In, 'Ec', 30000.0)/1000, 'GPa'],
            ['', '', '']
        ]],
        ['📋 설계 조건', [
            ['🔧 설계방법', getattr(In, 'Design_Method', 'USD').split('(')[0].strip(), ''],
            ['📖 설계기준', getattr(In, 'RC_Code', 'KDS-2021'), ''],
            ['🏛️ 기둥형식', getattr(In, 'Column_Type', 'Tied Column'), '']
        ]],
        ['🔩 철근 배치', [
            ['⭕ 철근 직경 D', getattr(In, 'dia', [22.0])[0], 'mm'],
            ['🛡️ 피복두께 dc', getattr(In, 'dc', [60.0])[0], 'mm'],
            ['📊 압축/인장측', f'각 {getattr(In, "be", 1000) / getattr(In, "sb", [150.0])[0]:,.1f}개' if getattr(In, "sb", [0])[0] > 0 else 'N/A', '']
        ]]
    ]
    
    start_cols = [0, 4, 9, 12]
    for i, (section_title, items) in enumerate(section_data):
        col_start = start_cols[i]
        col_end = col_start + 2
        column_ws.merge_range(row, col_start, row, col_end, section_title, formats['section'])
        column_ws.set_row(row, 22)
        
        for j, (label, value, unit) in enumerate(items):
            if label:
                current_data_row = row + j + 1
                column_ws.write(current_data_row, col_start, label, formats['label'])
                fmt = formats['number'] if isinstance(value, (int, float, np.number)) and unit else formats['value']
                column_ws.write(current_data_row, col_start + 1, value, fmt)
                column_ws.write(current_data_row, col_start + 2, unit, formats['unit'])
                column_ws.set_row(current_data_row, 20)
    
    row += 5
    
    # ═══════════════════════════════════════════════════════════
    # 섹션 3 & 4: 재료 비교 (이형철근 vs 중공철근 명확히 구분)
    # ═══════════════════════════════════════════════════════════
    # 구분 마커 추가
    column_ws.merge_range(row, 0, row, 6, '🔶 이형철근 (SD400/500)', formats['rc_marker'])
    column_ws.merge_range(row, 8, row, max_col, '🔷 중공철근 (면적50%, fy=800MPa)', formats['hollow_marker'])
    column_ws.set_row(row, 25)
    row += 1
    
    column_ws.merge_range(row, 0, row, 6, '📊 이형철근 검토', formats['main_header'])
    column_ws.merge_range(row, 8, row, max_col, '📊 중공철근 검토', formats['main_header'])
    column_ws.set_row(row, 25)
    row += 1
    
    column_ws.merge_range(row, 0, row, 6, '🔧 철근 재료 특성', formats['sub_header'])
    column_ws.merge_range(row, 8, row, max_col, '🔧 철근 재료 특성', formats['sub_header'])
    column_ws.set_row(row, 22)
    row += 1
    
    material_data = [
        ['💪 항복강도 fy', getattr(In, 'fy', 400.0), getattr(In, 'fy_hollow', 800.0), 'MPa', '(이형)', '(중공-면적50%)'],
        ['⚡ 탄성계수 Es', getattr(In, 'Es', 200000.0), getattr(In, 'Es_hollow', 200000.0), 'MPa', '', '']
    ]
    
    for label, vR, vF, unit, note_R, note_F in material_data:
        column_ws.merge_range(row, 0, row, 1, f'{label} {note_R}', formats['label'])
        column_ws.write(row, 2, vR, formats['number'])
        column_ws.write(row, 3, unit, formats['unit'])
        
        column_ws.merge_range(row, 8, row, 9, f'{label} {note_F}', formats['label'])
        column_ws.write(row, 10, vF, formats['number'])
        column_ws.write(row, 11, unit, formats['unit'])
        
        column_ws.set_row(row, 20)
        row += 1
    
    row += 1
    
    # ═══════════════════════════════════════════════════════════
    # 섹션 5: 평형상태 검토
    # ═══════════════════════════════════════════════════════════
    column_ws.merge_range(row, 0, row, 6, '⚖️ 평형상태(Balanced) 검토', formats['sub_header'])
    column_ws.merge_range(row, 8, row, max_col, '⚖️ 평형상태(Balanced) 검토', formats['sub_header'])
    column_ws.set_row(row, 22)
    row += 1
    
    try:
        R_Pd_b, R_Md_b, R_e_b, R_c_b = getattr(R,'Pd',[0]*6)[3], getattr(R,'Md',[0]*6)[3], getattr(R,'e',[0]*6)[3], getattr(R,'c',[0]*6)[3]
        F_Pd_b, F_Md_b, F_e_b, F_c_b = getattr(F,'Pd',[0]*6)[3], getattr(F,'Md',[0]*6)[3], getattr(F,'e',[0]*6)[3], getattr(F,'c',[0]*6)[3]
    except (AttributeError, IndexError, TypeError):
        R_Pd_b,R_Md_b,R_e_b,R_c_b=0,0,0,0
        F_Pd_b,F_Md_b,F_e_b,F_c_b=0,0,0,0
    
    equilibrium_data = [
        ['⚖️ 축력 Pb', R_Pd_b, F_Pd_b, 'kN'],
        ['📏 모멘트 Mb', R_Md_b, F_Md_b, 'kN·m'],
        ['📐 편심 eb', R_e_b, F_e_b, 'mm'],
        ['🎯 중립축 깊이 cb', R_c_b, F_c_b, 'mm']
    ]
    
    for label, vR, vF, unit in equilibrium_data:
        column_ws.merge_range(row, 0, row, 1, label, formats['label'])
        column_ws.write(row, 2, vR, formats['number'])
        column_ws.write(row, 3, unit, formats['unit'])
        
        column_ws.merge_range(row, 8, row, 9, label, formats['label'])
        column_ws.write(row, 10, vF, formats['number'])
        column_ws.write(row, 11, unit, formats['unit'])
        
        column_ws.set_row(row, 20)
        row += 1
    
    row += 1
    
    # ═══════════════════════════════════════════════════════════
    # 섹션 6: 강도 검토 결과 요약
    # ═══════════════════════════════════════════════════════════
    column_ws.merge_range(row, 0, row, 6, '📊 기둥강도 검토 결과 (요약)', formats['sub_header'])
    column_ws.merge_range(row, 8, row, max_col, '📊 기둥강도 검토 결과 (요약)', formats['sub_header'])
    column_ws.set_row(row, 22)
    row += 1
    
    # 테이블 헤더
    headers = ['LC', 'Pu/φPn [kN]', 'Mu/φMn [kN·m]', 'e [mm]', 'S.F.', '판정']
    
    for col_idx, header in enumerate(headers):
        column_ws.write(row, col_idx, header, formats['table_header'])
        column_ws.write(row, 8 + col_idx, header, formats['table_header'])
    
    column_ws.set_row(row, 22)
    row += 1
    
    all_results = {'R': [], 'F': []}
    
    try:
        Pu_values = safe_extract(In, 'Pu')
        Mu_values = safe_extract(In, 'Mu')
        Pd_RC_ends = getattr(R, 'Pd', [0]*6)
        Md_RC_ends = getattr(R, 'Md', [0]*6)
        Pd_FRP_ends = getattr(F, 'Pd', [0]*6)
        Md_FRP_ends = getattr(F, 'Md', [0]*6)
        Pd_RC_iter = safe_extract(In, 'Pd_RC')
        Md_RC_iter = safe_extract(In, 'Md_RC')
        Pd_FRP_iter = safe_extract(In, 'Pd_FRP')
        Md_FRP_iter = safe_extract(In, 'Md_FRP')
        
        num_load_cases = len(Pu_values)
        
        for i in range(num_load_cases):
            try:
                Pu, Mu = Pu_values[i], Mu_values[i]
                e = (Mu / Pu) * 1000 if Pu != 0 else np.inf
                
                if np.isclose(Pu, 0):
                    Pd_R, Md_R = Pd_RC_ends[5], Md_RC_ends[5]
                    Pd_F, Md_F = Pd_FRP_ends[5], Md_FRP_ends[5]
                elif np.isclose(Mu, 0):
                    Pd_R, Md_R = Pd_RC_ends[0], Md_RC_ends[0]
                    Pd_F, Md_F = Pd_FRP_ends[0], Md_FRP_ends[0]
                else:
                    Pd_R = Pd_RC_iter[i] if i < len(Pd_RC_iter) else 0
                    Md_R = Md_RC_iter[i] if i < len(Md_RC_iter) else 0
                    Pd_F = Pd_FRP_iter[i] if i < len(Pd_FRP_iter) else 0
                    Md_F = Md_FRP_iter[i] if i < len(Md_FRP_iter) else 0
                
                sR = np.sqrt(Pd_R**2+Md_R**2)/np.sqrt(Pu**2+Mu**2) if (Pu**2+Mu**2)>1e-9 else np.inf
                sF = np.sqrt(Pd_F**2+Md_F**2)/np.sqrt(Pu**2+Mu**2) if (Pu**2+Mu**2)>1e-9 else np.inf
                
                R_pass = sR >= 1.0
                F_pass = sF >= 1.0
                
                all_results['R'].append(R_pass)
                all_results['F'].append(F_pass)
                
                # 이형철근 행
                fmt_bg_R = formats['ok'] if R_pass else formats['ng']
                column_ws.write(row, 0, f'LC-{i+1}', formats['combo'])
                column_ws.write(row, 1, f'{Pu:,.1f}/{Pd_R:,.1f}', formats['table_data'])
                column_ws.write(row, 2, f'{Mu:,.1f}/{Md_R:,.1f}', formats['table_data'])
                column_ws.write(row, 3, e if np.isfinite(e) else "∞", formats['number'])
                column_ws.write(row, 4, sR if np.isfinite(sR) else "∞", formats['number'])
                column_ws.write(row, 5, 'PASS ✅' if R_pass else 'FAIL ❌', fmt_bg_R)
                
                # 중공철근 행
                fmt_bg_F = formats['ok'] if F_pass else formats['ng']
                column_ws.write(row, 8, f'LC-{i+1}', formats['combo'])
                column_ws.write(row, 9, f'{Pu:,.1f}/{Pd_F:,.1f}', formats['table_data'])
                column_ws.write(row, 10, f'{Mu:,.1f}/{Md_F:,.1f}', formats['table_data'])
                column_ws.write(row, 11, e if np.isfinite(e) else "∞", formats['number'])
                column_ws.write(row, 12, sF if np.isfinite(sF) else "∞", formats['number'])
                column_ws.write(row, 13, 'PASS ✅' if F_pass else 'FAIL ❌', fmt_bg_F)
                
                column_ws.set_row(row, 20)
            except Exception as e_inner:
                error_msg = f'LC-{i+1} Err: {str(e_inner)}'
                column_ws.write(row, 0, error_msg, formats['ng'])
                column_ws.write(row, 8, error_msg, formats['ng'])
                all_results['R'].append(False)
                all_results['F'].append(False)
            
            row += 1
            
    except Exception as e_outer:
        column_ws.write(row, 0, f'데이터 로드 오류', formats['ng'])
        column_ws.write(row, 8, f'데이터 로드 오류', formats['ng'])
        row += 1
    
    row += 1
    
    # ═══════════════════════════════════════════════════════════
    # 섹션 7: 상세 강도 검토 (모든 하중조합)
    # ═══════════════════════════════════════════════════════════
    column_ws.merge_range(row, 0, row, max_col, '🔍 상세 강도 검토 (모든 하중조합)', formats['common_section'])
    column_ws.set_row(row, 28)
    row += 2
    
    def write_detailed_calculation(case_idx, material_type, PM_obj, start_col):
        """스트림릿 UI를 그대로 재현한 상세 계산 과정"""
        nonlocal row
        initial_row = row
        
        try:
            Pu_values = safe_extract(In, 'Pu')
            Mu_values = safe_extract(In, 'Mu')
            
            if case_idx >= len(Pu_values) or case_idx >= len(Mu_values):
                return 0
            
            Pu, Mu = Pu_values[case_idx], Mu_values[case_idx]
            is_pure_bending = np.isclose(Pu, 0)
            is_pure_compression = np.isclose(Mu, 0)
            
            # 타이틀
            title_text = f'[하중조합 LC-{case_idx+1} 상세 계산 과정]'
            column_ws.merge_range(row, start_col, row, start_col + 6, title_text, formats['calc_title'])
            column_ws.set_row(row, 25)
            row += 1
            
            # ─────────────────────────────────────────────────
            # 순수 휨/압축 케이스 처리
            # ─────────────────────────────────────────────────
            if is_pure_bending or is_pure_compression:
                idx = 5 if is_pure_bending else 0
                c_assumed = safe_extract(PM_obj, 'c', [0]*6)[idx]
                phiPn = safe_extract(PM_obj, 'Pd', [0]*6)[idx]
                phiMn = safe_extract(PM_obj, 'Md', [0]*6)[idx]
                
                condition_str = "순수 휨 상태 (Pu = 0)" if is_pure_bending else "순수 압축 상태 (Mu = 0)"
                
                p_pass = Pu <= phiPn
                m_pass = Mu <= phiMn
                
                safety_factor = np.sqrt(phiPn**2 + phiMn**2) / np.sqrt(Pu**2 + Mu**2) if (Pu**2 + Mu**2) > 0 else np.inf
                sf_pass = safety_factor >= 1.0
                
                # 1. 기본 정보
                column_ws.merge_range(row, start_col, row, start_col + 6, '1. 기본 정보 및 설계계수', formats['calc_section'])
                column_ws.set_row(row, 20)
                row += 1
                
                calc_items = [
                    ('📌 특별 조건', condition_str),
                    ('💪 작용 축력 Pu', f'{Pu:,.1f} kN'),
                    ('💪 작용 모멘트 Mu', f'{Mu:,.1f} kN·m'),
                    ('🎯 중립축 c (사전 계산)', f'{c_assumed:,.1f} mm'),
                ]
                
                for label, value in calc_items:
                    column_ws.write(row, start_col, label, formats['calc_text'])
                    column_ws.merge_range(row, start_col + 1, row, start_col + 6, value, formats['calc_text'])
                    column_ws.set_row(row, 18)
                    row += 1
                
                row += 1
                
                # 6. 최종 검토
                column_ws.merge_range(row, start_col, row, start_col + 6, '6. 최종 검토 및 안전성 평가', formats['calc_section'])
                column_ws.set_row(row, 20)
                row += 1
                
                final_items = [
                    ('➡️ 설계 축강도 φPn', f'{phiPn:,.1f} kN'),
                    ('➡️ 설계 휨강도 φMn', f'{phiMn:,.1f} kN·m'),
                    ('⚖️ 축력 검토', f'Pu = {Pu:,.1f} kN {"≤" if p_pass else ">"} φPn = {phiPn:,.1f} kN → {"O.K. ✅" if p_pass else "N.G. ❌"}'),
                    ('⚖️ 휨강도 검토', f'Mu = {Mu:,.1f} kN·m {"≤" if m_pass else ">"} φMn = {phiMn:,.1f} kN·m → {"O.K. ✅" if m_pass else "N.G. ❌"}'),
                    ('📊 PM 교점 안전율', f'S.F. = {safety_factor:.1f} → {"안전 ✅" if sf_pass else "위험 ⚠️"}'),
                ]
                
                for label, value in final_items:
                    column_ws.write(row, start_col, label, formats['calc_text'])
                    column_ws.merge_range(row, start_col + 1, row, start_col + 6, value, formats['calc_text'])
                    column_ws.set_row(row, 18)
                    row += 1
                
                row += 2
                return row - initial_row
            
            # ─────────────────────────────────────────────────
            # 일반 케이스 처리 (전체 계산 과정)
            # ─────────────────────────────────────────────────
            Reinforcement_Type = 'hollow' if material_type == '중공철근' else 'RC'
            
            if material_type == '이형철근':
                c_values = safe_extract(In, 'c_RC')
                fy = getattr(In, 'fy', 400.0)
                Es = getattr(In, 'Es', 200000.0)
            else:
                c_values = safe_extract(In, 'c_FRP')
                fy = getattr(In, 'fy_hollow', 800.0)
                Es = getattr(In, 'Es_hollow', 200000.0)
            
            if not c_values or case_idx >= len(c_values):
                raise IndexError(f"LC-{case_idx+1}: 계산된 c 값 없음")
            
            c_assumed = c_values[case_idx]
            e_actual = (Mu / Pu) * 1000 if Pu != 0 else np.inf
            
            h, b, fck = getattr(In, 'height', 300), getattr(In, 'be', 1000), getattr(In, 'fck', 27.0)
            RC_Code = getattr(In, 'RC_Code', 'KDS-2021')
            Column_Type = getattr(In, 'Column_Type', 'Tied Column')
            
            # KDS-2021 계수 계산
            if 'KDS-2021' in RC_Code:
                n, ep_co, ep_cu = 2, 0.002, 0.0033
                if fck > 40:
                    n = 1.2 + 1.5 * ((100 - fck) / 60) ** 4
                    ep_co = 0.002 + (fck - 40) / 1e5
                    ep_cu = 0.0033 - (fck - 40) / 1e5
                n = min(n, 2)
                n = round(n * 100) / 100
                
                alpha = 1 - (1 / (1 + n)) * (ep_co / ep_cu)
                temp = (1 / (1 + n) / (2 + n)) * (ep_co / ep_cu) ** 2
                if fck <= 40:
                    alpha = 0.8
                beta = 1 - (0.5 - temp) / alpha
                if fck <= 50:
                    beta = 0.4
                
                alpha, beta = round(alpha * 100) / 100, round(beta * 100) / 100
                beta1 = 2 * beta
                eta = alpha / beta1
                eta = round(eta * 100) / 100
                if fck == 50:
                    eta = 0.97
                if fck == 80:
                    eta = 0.87
            else:
                beta1, eta, ep_cu = 0.85, 1.0, 0.003
            
            # 철근 배치
            dia, dc, sb, dia1, dc1 = In.dia, In.dc, In.sb, In.dia1, In.dc1
            Layer, ni = 1, [2]
            nst = b / sb[0] if sb[0] > 0 else 0
            nst1 = nst
            area_factor = 0.5 if 'hollow' in Reinforcement_Type else 1.0
            
            Ast = [np.pi * d**2 / 4 * area_factor for d in dia]
            Ast1 = [np.pi * d**2 / 4 * area_factor for d in dia1]
            
            dsi_2d = np.array([[dc1[0], h - dc[0]]])
            Asi_2d = np.array([[Ast1[0]*nst1, Ast[0]*nst]])
            
            ep_si_out = np.zeros_like(dsi_2d)
            fsi_out = np.zeros_like(dsi_2d)
            Fsi_out = np.zeros_like(dsi_2d)
            
            # 공칭강도 계산
            Pn, Mn = RC_and_AASHTO('Rectangle', Reinforcement_Type, beta1, c_assumed, eta, fck, Layer, ni, ep_si_out, ep_cu, dsi_2d, fsi_out, Es, fy, Fsi_out, Asi_2d, h, b, h)
            
            ep_si_calc, fsi_calc, Fsi_calc = ep_si_out, fsi_out, Fsi_out
            
            a = beta1 * c_assumed
            Ac = min(a, h) * b
            Cc = eta * (0.85 * fck) * Ac / 1000
            y_bar = (h / 2) - (a / 2) if a < h else 0
            
            Cs_force, Ts_force = Fsi_calc[0, 0], Fsi_calc[0, 1]
            Cc_moment = Cc * y_bar
            Cs_moment = Cs_force * (h/2 - dsi_2d[0, 0])
            Ts_moment = Ts_force * (h/2 - dsi_2d[0, 1])
            
            As1_calc, As_calc = Asi_2d[0, 0], Asi_2d[0, 1]
            
            # 강도감소계수
            dt = dsi_2d[0, 1]
            eps_t = ep_cu * (dt - c_assumed) / c_assumed if c_assumed > 0 else -np.inf
            eps_y = fy / Es
            phi_factor, phi_basis = 0.65, ""
            
            if 'RC' in Reinforcement_Type or 'hollow' in Reinforcement_Type:
                phi0 = 0.70 if 'Spiral' in Column_Type else 0.65
                ep_tccl = eps_y
                ep_ttcl = max(0.005, 2.5 * eps_y) if 'KDS-2021' in RC_Code and fy >= 400 else 0.005
                
                if eps_t <= ep_tccl:
                    phi_factor = phi0
                    phi_basis = f"εt = {eps_t:.5f} ≤ εty = {ep_tccl:.5f} → 압축지배 (φ={phi0:.2f})"
                elif eps_t >= ep_ttcl:
                    phi_factor = 0.85
                    phi_basis = f"εt = {eps_t:.5f} ≥ {ep_ttcl:.5f} → 인장지배 (φ=0.85)"
                else:
                    phi_factor = phi0 + (0.85 - phi0) * (eps_t - ep_tccl) / (ep_ttcl - ep_tccl)
                    phi_basis = f"εty({ep_tccl:.5f}) < εt({eps_t:.5f}) < {ep_ttcl:.5f} → 변화구간"
            
            phiPn, phiMn = Pn * phi_factor, Mn * phi_factor
            
            e_calc = (Mn * 1000 / Pn) if Pn != 0 else 0
            equilibrium_diff = abs(e_calc - e_actual)
            equilibrium_check = equilibrium_diff / max(abs(e_actual), 1) <= 0.01
            
            safety_factor = np.sqrt(phiPn**2 + phiMn**2) / np.sqrt(Pu**2 + Mu**2) if (Pu**2 + Mu**2) > 1e-9 else np.inf
            sf_pass = safety_factor >= 1.0
            
            p_pass = Pu <= phiPn
            m_pass = Mu <= phiMn
            
            # ═══ 1. 기본 정보 ═══
            column_ws.merge_range(row, start_col, row, start_col + 6, '1. 기본 정보 및 설계계수', formats['calc_section'])
            column_ws.set_row(row, 20)
            row += 1
            
            info_items = [
                ('📌 적용 기준', f'{RC_Code}, {Column_Type}'),
                ('🧱 콘크리트 계수', f'β₁={beta1:.2f}, η={eta:.2f}, εcu={ep_cu:.4f}, fck={fck:,.0f} MPa'),
                ('⛓️ 철근 재료', f'fy={fy:,.0f} MPa, Es={Es:,.0f} MPa ({material_type})'),
                ('💪 작용 하중', f'Pu={Pu:,.1f} kN, Mu={Mu:,.1f} kN·m (편심 e={e_actual:,.1f} mm)'),
                ('🎯 중립축', f'c = {c_assumed:,.1f} mm (시행착오법)'),
            ]
            
            for label, value in info_items:
                column_ws.write(row, start_col, label, formats['calc_text'])
                column_ws.merge_range(row, start_col + 1, row, start_col + 6, value, formats['calc_text'])
                column_ws.set_row(row, 18)
                row += 1
            
            row += 1
            
            # ═══ 2. 변형률 및 응력 ═══
            column_ws.merge_range(row, start_col, row, start_col + 6, '2. 변형률 및 응력 계산', formats['calc_section'])
            column_ws.set_row(row, 20)
            row += 1
            
            strain_items = [
                ('📜 변형률 계산식', 'εs = εcu × (c - ds) / c'),
                ('🔼 압축측 철근', f'εsc = {ep_si_calc[0,0]:.5f}, fsc = {fsi_calc[0,0]:,.1f} MPa (ds={dsi_2d[0,0]:.1f}mm)'),
                ('🔽 인장측 철근', f'εst = {ep_si_calc[0,1]:.5f}, fst = {fsi_calc[0,1]:,.1f} MPa (dt={dsi_2d[0,1]:.1f}mm)'),
            ]
            
            for label, value in strain_items:
                column_ws.write(row, start_col, label, formats['calc_text'])
                column_ws.merge_range(row, start_col + 1, row, start_col + 6, value, formats['calc_formula'] if '=' in value and 'ε' in value else formats['calc_text'])
                column_ws.set_row(row, 18)
                row += 1
            
            row += 1
            
            # ═══ 3. 단면력 평형 ═══
            column_ws.merge_range(row, start_col, row, start_col + 6, '3. 단면력 평형 및 공칭 축강도', formats['calc_section'])
            column_ws.set_row(row, 20)
            row += 1
            
            force_items = [
                ('📏 등가응력블록 깊이', f'a = β₁ × c = {beta1:.2f} × {c_assumed:.1f} = {a:.1f} mm'),
                ('🧱 콘크리트 압축면적', f'Ac = min(a,h) × b = {min(a,h):.1f} × {b:.1f} = {Ac:,.1f} mm²'),
                ('🧱 콘크리트 압축력', f'Cc = η × 0.85 × fck × Ac = {eta:.2f} × 0.85 × {fck:.1f} × {Ac:,.1f} = {Cc:,.1f} kN'),
                ('🔼 압축측 철근 면적', f"A's = {As1_calc:,.1f} mm²"),
                ('🔼 압축측 철근 합력', f"Cs = A's × (fsc - η×0.85×fck) = {As1_calc:,.1f} × ({fsi_calc[0,0]:,.1f} - {eta:.2f}×0.85×{fck:.1f}) = {Cs_force:,.1f} kN"),
                ('🔽 인장측 철근 면적', f'As = {As_calc:,.1f} mm²'),
                ('🔽 인장측 철근 합력', f'Ts = As × fst = {As_calc:,.1f} × {fsi_calc[0,1]:,.1f} = {Ts_force:,.1f} kN'),
                ('➡️ 공칭 축강도', f'Pn = Cc + Cs + Ts = {Cc:,.1f} + {Cs_force:,.1f} + {Ts_force:,.1f} = {Pn:,.1f} kN'),
            ]
            
            for label, value in force_items:
                column_ws.write(row, start_col, label, formats['calc_text'])
                column_ws.merge_range(row, start_col + 1, row, start_col + 6, value, formats['calc_formula'] if '=' in value else formats['calc_text'])
                column_ws.set_row(row, 18)
                row += 1
            
            row += 1
            
            # ═══ 4. 공칭 휨강도 ═══
            column_ws.merge_range(row, start_col, row, start_col + 6, '4. 공칭 휨강도 계산', formats['calc_section'])
            column_ws.set_row(row, 20)
            row += 1
            
            moment_items = [
                ('📍 콘크리트 압축력 중심', f'ȳ = (h/2) - (a/2) = ({h:.1f}/2) - ({a:.1f}/2) = {y_bar:.1f} mm'),
                ('🧱 콘크리트 모멘트', f'Mc = Cc × ȳ = {Cc:,.1f} × {y_bar:.1f} = {Cc_moment:,.1f} kN·mm'),
                ('🔼 압축철근 모멘트', f"M's = Cs × (h/2-d's) = {Cs_force:,.1f} × {(h/2-dsi_2d[0,0]):.1f} = {Cs_moment:,.1f} kN·mm"),
                ('🔽 인장철근 모멘트', f'Ms = Ts × (h/2-dt) = {Ts_force:,.1f} × {(h/2-dsi_2d[0,1]):.1f} = {Ts_moment:,.1f} kN·mm'),
                ('➡️ 공칭 휨강도', f"Mn = (Mc + M's + Ms) / 1000 = ({Cc_moment:,.1f} + {Cs_moment:,.1f} + {Ts_moment:,.1f}) / 1000 = {Mn:,.1f} kN·m"),
            ]
            
            for label, value in moment_items:
                column_ws.write(row, start_col, label, formats['calc_text'])
                column_ws.merge_range(row, start_col + 1, row, start_col + 6, value, formats['calc_formula'] if '=' in value else formats['calc_text'])
                column_ws.set_row(row, 18)
                row += 1
            
            row += 1
            
            # ═══ 5. 강도감소계수 ═══
            column_ws.merge_range(row, start_col, row, start_col + 6, '5. 강도감소계수(φ) 및 설계강도', formats['calc_section'])
            column_ws.set_row(row, 20)
            row += 1
            
            phi_items = [
                ('🧐 판단 근거', phi_basis),
                ('📉 강도감소계수', f'φ = {phi_factor:.3f}'),
                ('➡️ 설계 축강도', f'φPn = {phi_factor:.3f} × {Pn:,.1f} = {phiPn:,.1f} kN'),
                ('➡️ 설계 휨강도', f'φMn = {phi_factor:.3f} × {Mn:,.1f} = {phiMn:,.1f} kN·m'),
            ]
            
            for label, value in phi_items:
                column_ws.write(row, start_col, label, formats['calc_text'])
                column_ws.merge_range(row, start_col + 1, row, start_col + 6, value, formats['calc_formula'] if '=' in value and 'φ' in value else formats['calc_text'])
                column_ws.set_row(row, 18)
                row += 1
            
            row += 1
            
            # ═══ 6. 최종 검토 ═══
            column_ws.merge_range(row, start_col, row, start_col + 6, '6. 최종 검토 및 안전성 평가', formats['calc_section'])
            column_ws.set_row(row, 20)
            row += 1
            
            final_items = [
                ('📐 계산 편심', f"e' = Mn / Pn × 1000 = {Mn:,.1f} / {Pn:,.1f} × 1000 = {e_calc:.1f} mm"),
                ('💪 작용 편심', f'e = Mu / Pu × 1000 = {Mu:,.1f} / {Pu:,.1f} × 1000 = {e_actual:.1f} mm'),
                ('✔️ 평형 검토', f"|e'-e|/e×100 = {equilibrium_diff/max(abs(e_actual),1)*100:.1f}% {'≤ 1% (O.K.)' if equilibrium_check else '> 1%'}"),
                ('⚖️ 축력 검토', f'Pu = {Pu:,.1f} kN {"≤" if p_pass else ">"} φPn = {phiPn:,.1f} kN → {"O.K. ✅" if p_pass else "N.G. ❌"}'),
                ('⚖️ 휨강도 검토', f'Mu = {Mu:,.1f} kN·m {"≤" if m_pass else ">"} φMn = {phiMn:,.1f} kN·m → {"O.K. ✅" if m_pass else "N.G. ❌"}'),
                ('📊 PM 교점 안전율', f'S.F. = √[(φPn)²+(φMn)²] / √[Pu²+Mu²] = √[{phiPn:,.1f}²+{phiMn:,.1f}²] / √[{Pu:,.1f}²+{Mu:,.1f}²] = {safety_factor:.1f} → {"안전 ✅" if sf_pass else "위험 ⚠️"}'),
            ]
            
            for label, value in final_items:
                column_ws.write(row, start_col, label, formats['calc_text'])
                column_ws.merge_range(row, start_col + 1, row, start_col + 6, value, formats['calc_formula'] if '=' in value else formats['calc_text'])
                column_ws.set_row(row, 18)
                row += 1
            
            row += 2
            return row - initial_row
            
        except Exception as e:
            error_text = f'LC-{case_idx+1} 상세 검토 중 오류: {str(e)}'
            print(f"Error details for LC-{case_idx+1}:\n{traceback.format_exc()}")
            column_ws.merge_range(row, start_col, row + 1, start_col + 6, error_text, formats['ng'])
            column_ws.set_row(row, 25)
            row += 3
            return 3
    
    # 상세 계산 루프
    try:
        num_cases = len(safe_extract(In, 'Pu'))
        
        if num_cases == 0:
            column_ws.merge_range(row, 0, row, max_col, '검토할 하중조합이 없습니다.', formats['ng'])
            row += 1
        else:
            for case_idx in range(num_cases):
                initial_row_for_case = row
                
                # 이형철근 (왼쪽)
                write_detailed_calculation(case_idx, '이형철근', R, 0)
                
                # 중공철근 (오른쪽) - 같은 시작 행에서
                row = initial_row_for_case
                write_detailed_calculation(case_idx, '중공철근', F, 8)
                
                # 다음 케이스는 더 긴 쪽에 맞춤
                # row는 이미 write_detailed_calculation에서 업데이트됨
                row += 1  # 케이스 간 간격
                
    except Exception as e:
        print(f"Error during detailed calculation loop:\n{traceback.format_exc()}")
        error_msg = f'상세 계산 작성 중 오류: {str(e)}'
        column_ws.merge_range(row, 0, row + 2, max_col, error_msg, formats['ng'])
        row += 3
    
    # ═══════════════════════════════════════════════════════════
    # 섹션 8: 최종 종합 판정
    # ═══════════════════════════════════════════════════════════
    column_ws.merge_range(row, 0, row, 6, '🎯 최종 종합 판정', formats['sub_header'])
    column_ws.merge_range(row, 8, row, max_col, '🎯 최종 종합 판정', formats['sub_header'])
    column_ws.set_row(row, 22)
    row += 1
    
    for key, col_start, material_name in [('R', 0, '이형철근'), ('F', 8, '중공철근')]:
        if key in all_results and all_results.get(key) and len(all_results.get(key)) > 0:
            final_pass = all(all_results[key])
            text = f'🎉 {material_name}: 전체 조건 만족 (구조 안전)' if final_pass else f'⚠️ {material_name}: 일부 조건 불만족 (보강 검토 필요)'
            fmt = formats['final_ok'] if final_pass else formats['final_ng']
        else:
            text = f'⚠️ {material_name}: 검토 데이터 계산 불가'
            fmt = formats['final_ng']
        
        column_ws.merge_range(row, col_start, row, col_start + 6, text, fmt)
        column_ws.set_row(row, 28)
    
    row += 2
    
    # ═══════════════════════════════════════════════════════════
    # 섹션 9: 참고사항
    # ═══════════════════════════════════════════════════════════
    note_text = (
        "📋 검토 기준 및 참고사항\n\n"
        "🔍 PM 교점 안전율 판정 기준:\n"
        " • S.F. = √[(φPn)² + (φMn)²] / √[Pu² + Mu²]\n"
        " • S.F. ≥ 1.0 → PASS (구조적으로 안전)\n\n"
        "🔧 철근 종류별 특성:\n"
        " • 이형철근: 일반적인 SD400/SD500 철근\n"
        " • 중공철근: 단면적 50% 적용, 항복강도 800 MPa\n\n"
        f"📖 설계 기준: {getattr(In, 'RC_Code', 'KDS 41 (2021)')}\n"
        "⚡ 강도감소계수(φ)는 KDS-2021 기준에 따라 변형률 조건별로 적용"
    )
    
    column_ws.merge_range(row, 0, row + 10, max_col, note_text, formats['note_content'])
    
    return column_ws
