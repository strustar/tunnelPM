import numpy as np
import xlsxwriter # 이 코드를 실행하려면 'pip install XlsxWriter'가 필요합니다.

def RC_and_AASHTO(Section_Type, Reinforcement_Type, beta1, c, eta, fck, Layer, ni, ep_si, ep_cu, dsi, fsi, Es, fy, Fsi, Asi, *bhD):
    """
    변형률 호환과 힘의 평형을 기반으로 한 철근콘크리트 단면의 공칭 축력(P)과 모멘트(M) 용량 계산
    (주어진 함수를 그대로 사용합니다)
    """
    a = beta1 * c
    if 'Rectangle' in Section_Type:
        [hD, b, h] = bhD
        Ac = a * b if a < h else h * b
        y_bar = (h / 2) - (a / 2) if a < h else 0
    else:
        [hD] = bhD
        Ac = 0
        y_bar = 0

    Cc = eta * (0.85 * fck) * Ac / 1e3  # kN
    M = 0

    for L in range(Layer):
        for i in range(ni[L]):
            if c <= 0:
                continue
            
            ep_si[L, i] = ep_cu * (c - dsi[L, i]) / c
            fsi[L, i] = Es * ep_si[L, i]
            fsi[L, i] = np.clip(fsi[L, i], -fy, fy)

            if 'RC' in Reinforcement_Type or 'hollow' in Reinforcement_Type:
                if c >= dsi[L, i]:
                    Fsi[L, i] = Asi[L, i] * (fsi[L, i] - eta * 0.85 * fck) / 1e3
                else:
                    Fsi[L, i] = Asi[L, i] * fsi[L, i] / 1e3
            
            M = M + Fsi[L, i] * (hD / 2 - dsi[L, i])

    P = np.sum(Fsi) + Cc
    M = (M + Cc * y_bar) / 1e3
    
    return P, M

import numpy as np
import xlsxwriter

import numpy as np
import xlsxwriter # 이 코드를 실행하려면 'pip install XlsxWriter'가 필요합니다.

def create_column_sheet(wb, In, R, F):
    """기둥 강도 검토 시트 생성 - 스트림릿 웹과 동일한 상세 검토 포함 (순수 휨/압축 조건 추가)"""
    
    column_ws = wb.add_worksheet('기둥 강도 검토')
    
    # ─── 스타일 정의 ─────────────────────────────
    base_font = {'font_name': 'Noto Sans KR', 'border': 1, 'valign': 'vcenter'}
    
    styles = {
        'title': {**base_font, 'bold': True, 'font_size': 24, 'bg_color': '#1e40af', 
                'font_color': 'white', 'border': 3, 'align': 'center'},
        'main_header': {**base_font, 'bold': True, 'font_size': 18, 'bg_color': '#2563eb', 
                    'font_color': 'white', 'border': 2, 'align': 'center'},
        'common_section': {**base_font, 'bold': True, 'font_size': 16, 'bg_color': '#155e75', 
                          'font_color': '#e0f2fe', 'align': 'center'},
        'section': {**base_font, 'bold': True, 'font_size': 15, 'bg_color': '#1e3a8a', 
                   'font_color': 'white', 'align': 'center'},
        'sub_header': {**base_font, 'bold': True, 'font_size': 13, 'bg_color': '#3b82f6', 
                      'font_color': 'white', 'align': 'center'},
        'label': {**base_font, 'font_size': 12, 'bold': True, 'align': 'left'},
        'value': {**base_font, 'bold': True, 'font_size': 12, 'align': 'center'},
        'number': {**base_font, 'bold': True, 'font_size': 12, 'num_format': '#,##0.000', 
                  'align': 'center'},
        'unit': {**base_font, 'font_size': 12, 'align': 'center', 'bold': True},
        'combo': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#fef3c7', 
                 'font_color': '#92400e', 'align': 'center'},
        'ok': {**base_font, 'bold': True, 'font_size': 13, 'bg_color': '#dcfce7', 
              'font_color': '#166534', 'border': 2, 'align': 'center'},
        'ng': {**base_font, 'bold': True, 'font_size': 13, 'bg_color': '#fecaca', 
              'font_color': '#dc2626', 'border': 2, 'align': 'center'},
        'final_ok': {**base_font, 'bold': True, 'font_size': 16, 'bg_color': '#d1fae5', 
                    'font_color': '#065f46', 'border': 3, 'align': 'center'},
        'final_ng': {**base_font, 'bold': True, 'font_size': 16, 'bg_color': '#fee2e2', 
                    'font_color': '#991b1b', 'border': 3, 'align': 'center'},
        'calc_title': {**base_font, 'bold': True, 'font_size': 14, 'bg_color': '#dbeafe', 
                      'font_color': '#1e40af', 'border': 2, 'align': 'center'},
        'calc_content': {**base_font, 'font_size': 11, 'align': 'left', 'text_wrap': True,
                        'bg_color': '#fafafc', 'valign': 'top', 'bold': True},
        'calc_result': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#f0f9ff',
                       'font_color': '#0c4a6e', 'align': 'center'},
        'table_header': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#3b82f6',
                        'font_color': 'white', 'align': 'center'},
        'table_data': {**base_font, 'font_size': 11, 'align': 'center', 'bold': True},
        'icon_cell': {**base_font, 'font_size': 12, 'align': 'center', 'bold': True,
                     'bg_color': '#f8fafc'}
    }
    
    formats = {name: wb.add_format(props) for name, props in styles.items()}
    
    # ─── 컬럼 너비 설정 ─────────────────────────────
    col_widths = ['A:A', 'B:B', 'C:C', 'D:D', 'E:E', 'F:F', 'G:G', 'H:H', 'I:I', 'J:J', 'K:K', 'L:L', 'M:M', 'N:N']
    col_width_values = [22, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16]
    
    for col_range, width in zip(col_widths, col_width_values):
        column_ws.set_column(col_range, width)
    
    row = 0
    max_col = 13  # Column N
    
    # ─── 1. 메인 타이틀 ───────────────────────────────
    column_ws.merge_range(row, 0, row, max_col, '🏗️ 기둥 강도 검토 보고서', formats['title'])
    column_ws.set_row(row, 50)
    row += 2
    
    # ─── 2. 공통 설계 조건 ─────────────────────────────
    column_ws.merge_range(row, 0, row, max_col, '◈ 공통 설계 조건', formats['common_section'])
    column_ws.set_row(row, 35)
    row += 1
    
    section_data = [
        ['📐 단면 제원', [
            ['📏 단위폭 be', getattr(In, 'be', 1000), 'mm'],
            ['📏 단면 두께 h', getattr(In, 'height', 300), 'mm'],
            ['📐 공칭 철근간격 s', getattr(In, 'sb', [150.0])[0], 'mm']
        ]],
        ['🏭 콘크리트 재료', [
            ['💪 압축강도 fck', getattr(In, 'fck', 40.0), 'MPa'],
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
            ['📊 압축/인장측', f'{In.be / In.sb[0]:.0f}개씩', '']
        ]]
    ]
    
    start_cols = [0, 3, 7, 11]
    for i, (section_title, items) in enumerate(section_data):
        col_start = start_cols[i]
        column_ws.merge_range(row, col_start, row, col_start + 2, section_title, formats['section'])
        column_ws.set_row(row, 25)
        for j, (label, value, unit) in enumerate(items):
            if label:
                column_ws.write(row + j + 1, col_start, label, formats['label'])
                fmt = formats['number'] if isinstance(value, (int, float)) and unit else formats['value']
                column_ws.write(row + j + 1, col_start + 1, value, fmt)
                column_ws.write(row + j + 1, col_start + 2, unit, formats['unit'])
                column_ws.set_row(row + j + 1, 22)
    row += 5
    
    # ─── 3. 이형철근 vs 중공철근 비교 섹션 ─────────────
    column_ws.merge_range(row, 0, row, 6, '📊 이형철근 검토', formats['main_header'])
    column_ws.merge_range(row, 8, row, max_col, '📊 중공철근 검토', formats['main_header'])
    column_ws.set_row(row, 32)
    row += 1
    
    # ─── 4. 재료 특성 비교 ─────────────────────────────
    column_ws.merge_range(row, 0, row, 6, '🔧 철근 재료 특성', formats['sub_header'])
    column_ws.merge_range(row, 8, row, max_col, '🔧 철근 재료 특성', formats['sub_header'])
    column_ws.set_row(row, 25)
    row += 1
    
    material_data = [
        ['💪 항복강도 fy', getattr(In, 'fy', 400.0), getattr(In, 'fy_hollow', 800.0), 'MPa', '(이형철근)', '(중공철근 - 단면적 50%)'],
        ['⚡ 탄성계수 Es', getattr(In, 'Es', 200000.0)/1000, getattr(In, 'Es_hollow', 200000.0)/1000, 'GPa', '', '']
    ]
    
    for label, vR, vF, unit, note_R, note_F in material_data:
        column_ws.write(row, 0, f'{label} {note_R}', formats['label'])
        column_ws.write(row, 1, vR, formats['number'])
        column_ws.write(row, 2, unit, formats['unit'])
        column_ws.write(row, 8, f'{label} {note_F}', formats['label'])
        column_ws.write(row, 9, vF, formats['number'])
        column_ws.write(row, 10, unit, formats['unit'])
        column_ws.set_row(row, 22)
        row += 1
    row += 1
    
    # ─── 5. 평형상태 검토 ─────────────────────────────
    column_ws.merge_range(row, 0, row, 6, '⚖️ 평형상태(Balanced) 검토', formats['sub_header'])
    column_ws.merge_range(row, 8, row, max_col, '⚖️ 평형상태(Balanced) 검토', formats['sub_header'])
    column_ws.set_row(row, 25)
    row += 1
    
    try:
        R_data = [getattr(R, attr, [0,0,0,0]) for attr in ['Pd', 'Md', 'e', 'c']]
        F_data = [getattr(F, attr, [0,0,0,0]) for attr in ['Pd', 'Md', 'e', 'c']]
        R_vals = [data[3] if len(data) > 3 else 0.0 for data in R_data]
        F_vals = [data[3] if len(data) > 3 else 0.0 for data in F_data]
        equilibrium_data = [
            ['⚖️ 축력 Pb', R_vals[0], F_vals[0], 'kN'],
            ['📏 모멘트 Mb', R_vals[1], F_vals[1], 'kN·m'],
            ['📐 편심 eb', R_vals[2], F_vals[2], 'mm'],
            ['🎯 중립축 깊이 cb', R_vals[3], F_vals[3], 'mm']
        ]
    except (AttributeError, IndexError, TypeError):
        equilibrium_data = [['⚖️ 축력 Pb', 0.0, 0.0, 'kN'], ['📏 모멘트 Mb', 0.0, 0.0, 'kN·m'], ['📐 편심 eb', 0.0, 0.0, 'mm'], ['🎯 중립축 깊이 cb', 0.0, 0.0, 'mm']]
    
    for label, vR, vF, unit in equilibrium_data:
        column_ws.write(row, 0, label, formats['label'])
        column_ws.write(row, 1, vR, formats['number'])
        column_ws.write(row, 2, unit, formats['unit'])
        column_ws.write(row, 8, label, formats['label'])
        column_ws.write(row, 9, vF, formats['number'])
        column_ws.write(row, 10, unit, formats['unit'])
        column_ws.set_row(row, 22)
        row += 1
    row += 1
    
    # ─── 6. 강도 검토 결과 요약 ─────────────────────────
    column_ws.merge_range(row, 0, row, 6, '📊 기둥강도 검토 결과 (요약)', formats['sub_header'])
    column_ws.merge_range(row, 8, row, max_col, '📊 기둥강도 검토 결과 (요약)', formats['sub_header'])
    column_ws.set_row(row, 25)
    row += 1
    
    headers = ['하중조합', 'Pu/φPn [kN]', 'Mu/φMn [kN·m]', '편심 e [mm]', 'PM교점 안전율', '판정']
    for i, hdr in enumerate(headers):
        column_ws.write(row, i, hdr, formats['table_header'])
        column_ws.write(row, i + 8, hdr, formats['table_header'])
    column_ws.set_row(row, 25)
    row += 1
    
    all_results = {'R': [], 'F': []}
    
    try:
        Pu_values = safe_extract(In, 'Pu')
        Mu_values = safe_extract(In, 'Mu')
        
        # 순수 휨/압축 케이스를 위한 PM 다이어그램 양 끝 값
        Pd_RC_ends = getattr(R, 'Pd', [0]*6)
        Md_RC_ends = getattr(R, 'Md', [0]*6)
        Pd_FRP_ends = getattr(F, 'Pd', [0]*6)
        Md_FRP_ends = getattr(F, 'Md', [0]*6)
        
        # 반복 계산으로 얻은 PM 교점 값
        Pd_RC_iter = safe_extract(In, 'Pd_RC')
        Md_RC_iter = safe_extract(In, 'Md_RC')
        Pd_FRP_iter = safe_extract(In, 'Pd_FRP')
        Md_FRP_iter = safe_extract(In, 'Md_FRP')
        
        num_load_cases = len(Pu_values)
        
        for i in range(num_load_cases):
            Pu, Mu = Pu_values[i], Mu_values[i]
            e = (Mu / Pu) * 1000 if Pu != 0 else np.inf
            
            # 조건에 따라 설계강도 (Pd, Md) 결정
            if np.isclose(Pu, 0): # 순수 휨
                Pd_R, Md_R = Pd_RC_ends[5], Md_RC_ends[5]
                Pd_F, Md_F = Pd_FRP_ends[5], Md_FRP_ends[5]
            elif np.isclose(Mu, 0): # 순수 압축
                Pd_R, Md_R = Pd_RC_ends[0], Md_RC_ends[0]
                Pd_F, Md_F = Pd_FRP_ends[0], Md_FRP_ends[0]
            else: # 일반 경우
                Pd_R = Pd_RC_iter[i] if i < len(Pd_RC_iter) else 0
                Md_R = Md_RC_iter[i] if i < len(Md_RC_iter) else 0
                Pd_F = Pd_FRP_iter[i] if i < len(Pd_FRP_iter) else 0
                Md_F = Md_FRP_iter[i] if i < len(Md_FRP_iter) else 0

            sR = np.sqrt(Pd_R**2 + Md_R**2) / np.sqrt(Pu**2 + Mu**2) if (Pu**2 + Mu**2) > 0 else np.inf
            sF = np.sqrt(Pd_F**2 + Md_F**2) / np.sqrt(Pu**2 + Mu**2) if (Pu**2 + Mu**2) > 0 else np.inf

            R_pass = sR >= 1.0
            F_pass = sF >= 1.0
            all_results['R'].append(R_pass)
            all_results['F'].append(F_pass)
            
            row_data_R = [[f'LC-{i+1}', formats['combo']], [f'{Pu:,.1f} / {Pd_R:,.1f}', formats['table_data']], [f'{Mu:,.1f} / {Md_R:,.1f}', formats['table_data']], [e, formats['number']], [f'{sR:.1f}', formats['number']], ['PASS ✅' if R_pass else 'FAIL ❌', formats['ok'] if R_pass else formats['ng']]]
            row_data_F = [[f'LC-{i+1}', formats['combo']], [f'{Pu:,.1f} / {Pd_F:,.1f}', formats['table_data']], [f'{Mu:,.1f} / {Md_F:,.1f}', formats['table_data']], [e, formats['number']], [f'{sF:.1f}', formats['number']], ['PASS ✅' if F_pass else 'FAIL ❌', formats['ok'] if F_pass else formats['ng']]]

            for j, (val, fmt) in enumerate(row_data_R): column_ws.write(row, j, val, fmt)
            for j, (val, fmt) in enumerate(row_data_F): column_ws.write(row, j + 8, val, fmt)
                
            column_ws.set_row(row, 22)
            row += 1
            
    except Exception as e:
        column_ws.merge_range(row, 0, row, 6, f'데이터 처리 오류: {e}', formats['ng'])
        column_ws.merge_range(row, 8, row, max_col, f'데이터 처리 오류: {e}', formats['ng'])
        row += 1
    
    row += 1
    
    # ─── 7. 상세 계산 과정 ─────────────────────────────
    column_ws.merge_range(row, 0, row, max_col, '🔍 상세 강도 검토 과정 (모든 하중조합)', formats['common_section'])
    column_ws.set_row(row, 35)
    row += 2
    
    # 상세 계산 작성 함수
    def write_detailed_calculation(case_idx, material_type, PM_obj, start_col):
        nonlocal row
        
        try:
            Pu_values = safe_extract(In, 'Pu')
            Mu_values = safe_extract(In, 'Mu')
            if case_idx >= len(Pu_values) or case_idx >= len(Mu_values): return 1
            Pu, Mu = Pu_values[case_idx], Mu_values[case_idx]

            is_pure_bending = np.isclose(Pu, 0)
            is_pure_compression = np.isclose(Mu, 0)
            
            calc_contents = []
            title_text = f'[LC-{case_idx+1}] {material_type} 상세 계산 과정'
            column_ws.merge_range(row, start_col, row, start_col + 5, title_text, formats['calc_title'])
            column_ws.set_row(row, 30)

            if is_pure_bending or is_pure_compression:
                if is_pure_bending:
                    c_assumed = getattr(PM_obj, 'c', [0]*6)[5]
                    phiPn = getattr(PM_obj, 'Pd', [0]*6)[5]
                    phiMn = getattr(PM_obj, 'Md', [0]*6)[5]
                    condition_str = "순수 휨 상태 (Pu = 0)"
                else: # is_pure_compression
                    c_assumed = getattr(PM_obj, 'c', [0]*6)[0]
                    phiPn = getattr(PM_obj, 'Pd', [0]*6)[0]
                    phiMn = getattr(PM_obj, 'Md', [0]*6)[0]
                    condition_str = "순수 압축 상태 (Mu = 0)"

                p_status = "O.K." if Pu <= phiPn else "N.G."
                m_status = "O.K." if Mu <= phiMn else "N.G."
                safety_factor = np.sqrt(phiPn**2 + phiMn**2) / np.sqrt(Pu**2 + Mu**2) if (Pu**2 + Mu**2) > 0 else np.inf
                sf_status = "안전" if safety_factor >= 1.0 else "위험"
                
                calc_contents = [
                    '1. 기본 정보 및 설계계수',
                    f'   • 특별 조건: {condition_str}',
                    f'   • 작용 하중: Pu={Pu:,.1f} kN, Mu={Mu:,.1f} kN·m',
                    f'   • 결정된 중립축: c={c_assumed:,.1f} mm (사전 계산값)',
                    '',
                    '2. 최종 검토 및 안전성 평가 (요약)',
                    f'   • 축력 검토: Pu={Pu:,.1f} kN {"≤" if p_status == "O.K." else ">"} φPn={phiPn:,.1f} kN (∴ {p_status})',
                    f'   • 휨강도 검토: Mu={Mu:,.1f} kN·m {"≤" if m_status == "O.K." else ">"} φMn={phiMn:,.1f} kN·m (∴ {m_status})',
                    f'   • PM 교점 안전율: S.F. = {safety_factor:.1f} ({sf_status})'
                ]
            
            else:
                if material_type == '이형철근':
                    c_values, phiPn_values, phiMn_values = safe_extract(In, 'c_RC'), safe_extract(In, 'Pd_RC'), safe_extract(In, 'Md_RC')
                    fy, Es, steel_note = getattr(In, 'fy', 400.0), getattr(In, 'Es', 200000.0), '(이형철근)'
                else:
                    c_values, phiPn_values, phiMn_values = safe_extract(In, 'c_FRP'), safe_extract(In, 'Pd_FRP'), safe_extract(In, 'Md_FRP')
                    fy, Es, steel_note = getattr(In, 'fy_hollow', 800.0), getattr(In, 'Es_hollow', 200000.0), '(중공철근 - 단면적 50%)'

                if case_idx >= len(c_values): return 1
                c_assumed, phiPn, phiMn = c_values[case_idx], phiPn_values[case_idx], phiMn_values[case_idx]
                e_actual = (Mu / Pu) * 1000 if Pu != 0 else np.inf
                
                h, b, fck = getattr(In, 'height', 300), getattr(In, 'be', 1000), getattr(In, 'fck', 40.0)
                RC_Code, Column_Type = getattr(In, 'RC_Code', 'KDS-2021'), getattr(In, 'Column_Type', 'Tied Column')
                
                if 'KDS-2021' in RC_Code:
                    [n, ep_co, ep_cu] = [2, 0.002, 0.0033]
                    if fck > 40: n, ep_co, ep_cu = 1.2 + 1.5 * ((100 - fck) / 60)**4, 0.002 + (fck - 40)/1e5, 0.0033 - (fck - 40)/1e5
                    if n >= 2: n = 2
                    n = round(n * 100) / 100
                    alpha = 1 - 1/(1+n)*(ep_co/ep_cu)
                    temp = 1/(1+n)/(2+n)*(ep_co/ep_cu)**2
                    if fck <= 40: alpha = 0.8
                    beta = 1 - (0.5 - temp)/alpha
                    if fck <= 50: beta = 0.4
                    beta1, eta = 2 * round(beta*100)/100, round((round(alpha*100)/100) / (2 * round(beta*100)/100)*100)/100
                    if fck == 50: eta = 0.97
                    if fck == 80: eta = 0.87
                
                Layer, ni = 1, [2]
                dia, dc, dia1, dc1, sb = In.dia, In.dc, In.dia1, In.dc1, In.sb
                nst = b / sb[0]
                area_factor = 0.5 if material_type == '중공철근' else 1.0
                Ast, Ast1 = [np.pi*d**2/4*area_factor for d in dia], [np.pi*d**2/4*area_factor for d in dia1]
                dsi, Asi = np.zeros((Layer, ni[0])), np.zeros((Layer, ni[0]))
                dsi[0,:], Asi[0,:] = [dc1[0], h-dc[0]], [Ast1[0]*nst, Ast[0]*nst]
                ep_si, fsi, Fsi = np.zeros_like(dsi), np.zeros_like(dsi), np.zeros_like(dsi)
                Reinforcement_Type = 'hollow' if material_type == '중공철근' else 'RC'
                
                [Pn, Mn] = RC_and_AASHTO('Rectangle', Reinforcement_Type, beta1, c_assumed, eta, fck, Layer, ni, ep_si, ep_cu, dsi, fsi, Es, fy, Fsi, Asi, h, b, h)
                
                a, Ac = beta1 * c_assumed, min(beta1*c_assumed, h) * b
                Cc = eta * (0.85 * fck) * Ac / 1000
                y_bar = (h/2) - (a/2) if a < h else 0
                Cs_force, Ts_force = Fsi[0,0], Fsi[0,1]
                
                dt, eps_y = dsi[0,1], fy/Es
                eps_t = ep_cu * (dt-c_assumed)/c_assumed if c_assumed > 0 else 0
                phi0 = 0.70 if 'Spiral' in Column_Type else 0.65
                ep_tccl, ep_ttcl = eps_y, 0.005 if fy < 400 else 2.5 * eps_y
                
                if eps_t <= ep_tccl: phi_factor, phi_basis = phi0, f"압축지배단면 (φ={phi0:.2f})"
                elif eps_t >= ep_ttcl: phi_factor, phi_basis = 0.85, f"인장지배단면 (φ=0.85)"
                else: phi_factor, phi_basis = phi0 + (0.85-phi0)*(eps_t-ep_tccl)/(ep_ttcl-ep_tccl), f"변화구간 (φ={phi_factor:.3f})"
                
                safety_factor = np.sqrt(phiPn**2 + phiMn**2) / np.sqrt(Pu**2 + Mu**2) if (Pu**2 + Mu**2) > 0 else np.inf
                sf_status = "안전" if safety_factor >= 1.0 else "위험"
                
                calc_contents = [
                    '1. 기본 정보 및 설계계수', f'   • 작용 하중: Pu={Pu:,.1f} kN, Mu={Mu:,.1f} kN·m (편심 e={e_actual:.1f} mm)', f'   • 가정된 중립축: c={c_assumed:.1f} mm', '',
                    '2. 변형률 호환 및 응력 계산', f'   • 압축측 철근 (ds={dsi[0,0]:.1f}mm): εsc={ep_si[0,0]:.4f} → fsc={fsi[0,0]:,.1f} MPa', f'   • 인장측 철근 (dt={dsi[0,1]:.1f}mm): εst={ep_si[0,1]:.4f} → fst={fsi[0,1]:,.1f} MPa', '',
                    '3. 단면력 평형 및 공칭강도 계산', f'   • 등가응력블록 깊이: a = {a:.1f} mm', f'   • 콘크리트 압축력: Cc = {Cc:,.1f} kN', f'   • 압축/인장 철근 합력: Cs = {Cs_force:,.1f} kN, Ts = {Ts_force:,.1f} kN', f'   • 공칭 축강도: Pn = {Pn:,.1f} kN', '',
                    '4. 공칭 휨강도 계산', f'   • 공칭 휨강도: Mn = {Mn:,.1f} kN·m', '',
                    '5. 강도감소계수 및 설계강도', f'   • 판단 근거: {phi_basis}', f'   • 설계 축/휨강도: φPn = {phiPn:,.1f} kN, φMn = {phiMn:,.1f} kN·m', '',
                    '6. 최종 검토 및 안전성 평가', f'   • 축력 검토: Pu={Pu:,.1f} {"≤" if Pu <= phiPn else ">"} φPn={phiPn:,.1f} kN', f'   • 휨강도 검토: Mu={Mu:,.1f} {"≤" if Mu <= phiMn else ">"} φMn={phiMn:,.1f} kN·m', f'   • PM 교점 안전율: S.F. = {safety_factor:.1f} ({sf_status})'
                ]
            
            calc_start_row = row + 1
            for i, content in enumerate(calc_contents):
                current_row = calc_start_row + i
                if content:
                    column_ws.merge_range(current_row, start_col, current_row, start_col + 5, content, formats['calc_content'])
                    column_ws.set_row(current_row, 25)
                else:
                    column_ws.set_row(current_row, 12)
            
            return len(calc_contents) + 1
            
        except Exception as e:
            error_text = f'LC-{case_idx+1} {material_type} 계산 오류: {str(e)[:100]}'
            column_ws.merge_range(row, start_col, row, start_col + 5, error_text, formats['ng'])
            return 1
    
    try:
        num_cases = len(safe_extract(In, 'Pu'))
        for case_idx in range(num_cases):
            lines_used_R = write_detailed_calculation(case_idx, '이형철근', R, 0)
            lines_used_F = write_detailed_calculation(case_idx, '중공철근', F, 8)
            row += max(lines_used_R, lines_used_F) + 2
    except Exception as e:
        column_ws.merge_range(row, 0, row, max_col, f'상세 계산 작성 중 오류: {e}', formats['ng'])
        row += 1
    
    # ─── 8. 최종 종합 판정 ────────────────────────────
    column_ws.merge_range(row, 0, row, 6, '🎯 최종 종합 판정', formats['sub_header'])
    column_ws.merge_range(row, 8, row, max_col, '🎯 최종 종합 판정', formats['sub_header'])
    column_ws.set_row(row, 25)
    row += 1
    
    for key, col_start, material_name in [('R', 0, '이형철근'), ('F', 8, '중공철근')]:
        if key in all_results and all_results[key]:
            final_pass = all(all_results[key])
            text = f'🎉 {material_name} - 전체 조건 만족 (구조 안전)' if final_pass else f'⚠️ {material_name} - 일부 조건 불만족 (보강 검토 필요)'
            fmt = formats['final_ok'] if final_pass else formats['final_ng']
        else:
            text = f'❓ {material_name} - 검토 데이터 부족'
            fmt = formats['ng']
        column_ws.merge_range(row, col_start, row, col_start + 5, text, fmt)
        column_ws.set_row(row, 35)
    row += 2
    
    # ─── 9. 참고사항 ─────────────────────────────
    note_text = ("📋 검토 기준 및 참고사항\n\n" "🔍 PM 교점 안전율 판정 기준:\n" "  • S.F. = √[(φPn)² + (φMn)²] / √[Pu² + Mu²]\n" "  • S.F. ≥ 1.0 → PASS (구조적으로 안전)\n\n" "🔧 철근 종류별 특성:\n" "  • 이형철근: 일반적인 SD400/SD500 철근\n" "  • 중공철근: 단면적 50% 적용, 항복강도 800 MPa\n\n" f"📖 설계 기준: {getattr(In, 'RC_Code', 'KDS 41 17 00 (2021)')}\n" "⚡ 강도감소계수(φ)는 KDS-2021 기준에 따라 변형률 조건별로 적용")
    column_ws.merge_range(row, 0, row + 8, max_col, note_text, formats['calc_content'])
    
    return column_ws

# 헬퍼 함수 (클래스 외부에서 사용)
def safe_extract(obj, attr):
    """객체에서 안전하게 속성을 추출하는 헬퍼 함수"""
    try:
        val = getattr(obj, attr, [])
        if hasattr(val, 'tolist'):
            return val.tolist()
        elif isinstance(val, (list, tuple)):
            return list(val)
        else:
            return [val] if val is not None else []
    except Exception:
        return []

            