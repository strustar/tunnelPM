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

def create_column_sheet(wb, In, R, F):
    """기둥 강도 검토 시트 생성 - 스트림릿 웹과 동일한 상세 검토 포함"""
    
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
    
    # 공통 조건 4개 섹션을 가로로 배치
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
            ['📊 압축/인장측', f'각 {In.be / In.sb[0]:.0f}개', '']
        ]]
    ]
    
    # 4개 섹션을 3열씩 배치
    start_cols = [0, 3, 7, 11]
    for i, (section_title, items) in enumerate(section_data):
        col_start = start_cols[i]
        
        # 섹션 헤더
        column_ws.merge_range(row, col_start, row, col_start + 2, section_title, formats['section'])
        column_ws.set_row(row, 25)
        
        # 각 항목
        for j, (label, value, unit) in enumerate(items):
            if label:  # 빈 행이 아닌 경우만
                column_ws.write(row + j + 1, col_start, label, formats['label'])
                fmt = formats['number'] if isinstance(value, (int, float)) and unit else formats['value']
                column_ws.write(row + j + 1, col_start + 1, value, fmt)
                column_ws.write(row + j + 1, col_start + 2, unit, formats['unit'])
                column_ws.set_row(row + j + 1, 22)
    
    row += 5  # 4행 데이터 + 1행 여백
    
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
        # 이형철근 (좌측)
        column_ws.write(row, 0, f'{label} {note_R}', formats['label'])
        column_ws.write(row, 1, vR, formats['number'])
        column_ws.write(row, 2, unit, formats['unit'])
        # 중공철근 (우측)
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
    
    # 평형상태 데이터 추출
    try:
        R_data = [getattr(R, attr, [0,0,0,0]) for attr in ['Pd', 'Md', 'e', 'c']]
        F_data = [getattr(F, attr, [0,0,0,0]) for attr in ['Pd', 'Md', 'e', 'c']]
        
        # 안전하게 4번째 인덱스 추출
        R_vals = [data[3] if len(data) > 3 else 0.0 for data in R_data]
        F_vals = [data[3] if len(data) > 3 else 0.0 for data in F_data]
        
        equilibrium_data = [
            ['⚖️ 축력 Pb', R_vals[0], F_vals[0], 'kN'],
            ['📏 모멘트 Mb', R_vals[1], F_vals[1], 'kN·m'],
            ['📐 편심 eb', R_vals[2], F_vals[2], 'mm'],
            ['🎯 중립축 깊이 cb', R_vals[3], F_vals[3], 'mm']
        ]
    except (AttributeError, IndexError, TypeError):
        equilibrium_data = [
            ['⚖️ 축력 Pb', 0.0, 0.0, 'kN'],
            ['📏 모멘트 Mb', 0.0, 0.0, 'kN·m'],
            ['📐 편심 eb', 0.0, 0.0, 'mm'],
            ['🎯 중립축 깊이 cb', 0.0, 0.0, 'mm']
        ]
    
    for label, vR, vF, unit in equilibrium_data:
        # 이형철근 (좌측)
        column_ws.write(row, 0, label, formats['label'])
        column_ws.write(row, 1, vR, formats['number'])
        column_ws.write(row, 2, unit, formats['unit'])
        # 중공철근 (우측)
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
    
    # 테이블 헤더
    headers = ['하중조합', 'Pu/φPn [kN]', 'Mu/φMn [kN·m]', '편심 e [mm]', 'PM교점 안전율', '판정']
    for i, hdr in enumerate(headers):
        if i < 6:  # 좌측 이형철근
            column_ws.write(row, i, hdr, formats['table_header'])
        if i < 6:  # 우측 중공철근
            column_ws.write(row, i + 8, hdr, formats['table_header'])
    column_ws.set_row(row, 25)
    row += 1
    
    # 결과 데이터 처리
    all_results = {'R': [], 'F': []}
    
    try:
        # 데이터 안전하게 추출
        def safe_extract(obj, attr):
            val = getattr(obj, attr, [])
            return val.tolist() if hasattr(val, 'tolist') else list(val) if val else []
        
        Pu_values = safe_extract(In, 'Pu')
        Mu_values = safe_extract(In, 'Mu')
        safe_RC = safe_extract(In, 'safe_RC')
        safe_FRP = safe_extract(In, 'safe_FRP')
        Pd_RC = safe_extract(In, 'Pd_RC')
        Md_RC = safe_extract(In, 'Md_RC')
        Pd_FRP = safe_extract(In, 'Pd_FRP')
        Md_FRP = safe_extract(In, 'Md_FRP')
        
        num_load_cases = min(len(Pu_values), len(Mu_values)) if Pu_values and Mu_values else 0
        
        for i in range(num_load_cases):
            if i >= len(Pu_values) or i >= len(Mu_values):
                break
                
            Pu, Mu = Pu_values[i], Mu_values[i]
            e = (Mu / Pu) * 1000 if Pu != 0 else 0
            
            # PM 교점 거리비 안전율 계산
            if i < len(Pd_RC) and i < len(Md_RC) and Pu > 0 and Mu > 0:
                sR = np.sqrt(Pd_RC[i]**2 + Md_RC[i]**2) / np.sqrt(Pu**2 + Mu**2)
            else:
                sR = safe_RC[i] if i < len(safe_RC) else 0
                
            if i < len(Pd_FRP) and i < len(Md_FRP) and Pu > 0 and Mu > 0:
                sF = np.sqrt(Pd_FRP[i]**2 + Md_FRP[i]**2) / np.sqrt(Pu**2 + Mu**2)
            else:
                sF = safe_FRP[i] if i < len(safe_FRP) else 0
            
            R_pass = sR >= 1.0
            F_pass = sF >= 1.0
            all_results['R'].append(R_pass)
            all_results['F'].append(F_pass)
            
            # 데이터 테이블 작성
            row_data = [
                [f'LC-{i+1}', formats['combo']],
                [f'{Pu:,.1f} / {Pd_RC[i]:,.1f}' if i < len(Pd_RC) else f'{Pu:,.1f} / 0.0', formats['table_data']],
                [f'{Mu:,.1f} / {Md_RC[i]:,.1f}' if i < len(Md_RC) else f'{Mu:,.1f} / 0.0', formats['table_data']],
                [e, formats['number']],
                [f'{sR:.3f}', formats['number']],
                ['PASS ✅' if R_pass else 'FAIL ❌', formats['ok'] if R_pass else formats['ng']]
            ]
            
            # 이형철근 결과 (좌측)
            for j, (val, fmt) in enumerate(row_data):
                column_ws.write(row, j, val, fmt)
            
            # 중공철근 결과 (우측)
            row_data_F = [
                [f'LC-{i+1}', formats['combo']],
                [f'{Pu:,.1f} / {Pd_FRP[i]:,.1f}' if i < len(Pd_FRP) else f'{Pu:,.1f} / 0.0', formats['table_data']],
                [f'{Mu:,.1f} / {Md_FRP[i]:,.1f}' if i < len(Md_FRP) else f'{Mu:,.1f} / 0.0', formats['table_data']],
                [e, formats['number']],
                [f'{sF:.3f}', formats['number']],
                ['PASS ✅' if F_pass else 'FAIL ❌', formats['ok'] if F_pass else formats['ng']]
            ]
            
            for j, (val, fmt) in enumerate(row_data_F):
                column_ws.write(row, j + 8, val, fmt)
                
            column_ws.set_row(row, 22)
            row += 1
            
    except Exception as e:
        # 오류 발생 시 기본값 표시
        column_ws.merge_range(row, 0, row, 6, f'데이터 처리 오류: {e}', formats['ng'])
        column_ws.merge_range(row, 8, row, max_col, f'데이터 처리 오류: {e}', formats['ng'])
        column_ws.set_row(row, 25)
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
            # 기본 데이터 추출
            Pu_values = safe_extract(In, 'Pu')
            Mu_values = safe_extract(In, 'Mu')
            
            if case_idx >= len(Pu_values) or case_idx >= len(Mu_values):
                return 1
                
            Pu, Mu = Pu_values[case_idx], Mu_values[case_idx]
            e_actual = (Mu / Pu) * 1000 if Pu != 0 else 0
            
            # 재료별 데이터 추출
            if material_type == '이형철근':
                c_values = safe_extract(In, 'c_RC')
                phiPn_values = safe_extract(In, 'Pd_RC')
                phiMn_values = safe_extract(In, 'Md_RC')
                fy, Es = getattr(In, 'fy', 400.0), getattr(In, 'Es', 200000.0)
                steel_note = '(이형철근)'
            else:
                c_values = safe_extract(In, 'c_FRP')
                phiPn_values = safe_extract(In, 'Pd_FRP')
                phiMn_values = safe_extract(In, 'Md_FRP')
                fy, Es = getattr(In, 'fy_hollow', 800.0), getattr(In, 'Es_hollow', 200000.0)
                steel_note = '(중공철근 - 단면적 50%)'
            
            if case_idx >= len(c_values) or case_idx >= len(phiPn_values) or case_idx >= len(phiMn_values):
                return 1
                
            c_assumed = c_values[case_idx]
            phiPn = phiPn_values[case_idx]
            phiMn = phiMn_values[case_idx]
            
            # 설계 계수 계산
            h, b, fck = getattr(In, 'height', 300), getattr(In, 'be', 1000), getattr(In, 'fck', 40.0)
            RC_Code = getattr(In, 'RC_Code', 'KDS-2021')
            Column_Type = getattr(In, 'Column_Type', 'Tied Column')
            
            # KDS-2021 계수 계산
            if 'KDS-2021' in RC_Code:
                [n, ep_co, ep_cu] = [2, 0.002, 0.0033]
                if fck > 40:
                    n = 1.2 + 1.5 * ((100 - fck) / 60) ** 4
                    ep_co = 0.002 + (fck - 40) / 1e5
                    ep_cu = 0.0033 - (fck - 40) / 1e5
                if n >= 2: n = 2
                n = round(n * 100) / 100
                
                alpha = 1 - 1 / (1 + n) * (ep_co / ep_cu)
                temp = 1 / (1 + n) / (2 + n) * (ep_co / ep_cu) ** 2
                if fck <= 40: alpha = 0.8
                beta = 1 - (0.5 - temp) / alpha
                if fck <= 50: beta = 0.4
                
                [alpha, beta] = [round(alpha * 100) / 100, round(beta * 100) / 100]
                beta1 = 2 * beta
                eta = alpha / beta1
                eta = round(eta * 100) / 100
                if fck == 50: eta = 0.97
                if fck == 80: eta = 0.87
            
            # 철근 배치 설정
            Layer = 1
            ni = [2]
            dia, dc = getattr(In, 'dia', [22.0]), getattr(In, 'dc', [60.0])
            dia1, dc1 = getattr(In, 'dia1', [22.0]), getattr(In, 'dc1', [60.0])
            sb = getattr(In, 'sb', [150.0])
            
            nst = b / sb[0]
            area_factor = 0.5 if material_type == '중공철근' else 1.0
            
            Ast = [np.pi * d**2 / 4 * area_factor for d in dia]
            Ast1 = [np.pi * d**2 / 4 * area_factor for d in dia1]
            
            dsi = np.zeros((Layer, ni[0]))
            Asi = np.zeros((Layer, ni[0]))
            dsi[0, :] = [dc1[0], h - dc[0]]
            Asi[0, :] = [Ast1[0] * nst, Ast[0] * nst]
            
            ep_si, fsi, Fsi = np.zeros_like(dsi), np.zeros_like(dsi), np.zeros_like(dsi)
            
            # 공칭강도 계산
            Reinforcement_Type = 'hollow' if material_type == '중공철근' else 'RC'
            [Pn, Mn] = RC_and_AASHTO('Rectangle', Reinforcement_Type, beta1, c_assumed, eta, fck, Layer, ni, ep_si, ep_cu, dsi, fsi, Es, fy, Fsi, Asi, h, b, h)
            
            # 중간값 계산
            a = beta1 * c_assumed
            Ac = min(a, h) * b
            Cc = eta * (0.85 * fck) * Ac / 1000
            y_bar = (h / 2) - (a / 2) if a < h else 0
            
            Cs_force = Fsi[0, 0]
            Ts_force = Fsi[0, 1]
            
            # 강도감소계수 계산
            dt = dsi[0, 1]
            eps_t = ep_cu * (dt - c_assumed) / c_assumed if c_assumed > 0 else 0
            eps_y = fy / Es
            
            phi0 = 0.70 if 'Spiral' in Column_Type else 0.65
            ep_tccl = eps_y
            ep_ttcl = 0.005 if fy < 400 else 2.5 * eps_y
            
            if eps_t <= ep_tccl:
                phi_factor = phi0
                phi_basis = f"압축지배단면 (φ={phi0:.2f})"
            elif eps_t >= ep_ttcl:
                phi_factor = 0.85
                phi_basis = f"인장지배단면 (φ=0.85)"
            else:
                phi_factor = phi0 + (0.85 - phi0) * (eps_t - ep_tccl) / (ep_ttcl - ep_tccl)
                phi_basis = f"변화구간 (φ={phi_factor:.3f})"
            
            # PM 교점 안전율
            safety_factor = np.sqrt(phiPn**2 + phiMn**2) / np.sqrt(Pu**2 + Mu**2) if Pu > 0 and Mu > 0 else 0
            sf_status = "안전" if safety_factor >= 1.0 else "위험"
            
            # Excel에 상세 계산 내용 작성
            calc_start_row = row
            
            # 제목
            title_text = f'[LC-{case_idx+1}] {material_type} 상세 계산 과정'
            column_ws.merge_range(calc_start_row, start_col, calc_start_row, start_col + 5, title_text, formats['calc_title'])
            column_ws.set_row(calc_start_row, 30)
            calc_start_row += 1
                        # 계산 과정 내용
            calc_contents = [
                '1. 기본 정보 및 설계계수',
                f'   • 적용 기준: {RC_Code}, 기둥 형식: {Column_Type}',
                f'   • 콘크리트 계수: β₁={beta1:.3f}, η={eta:.3f}, εcu={ep_cu:.5f}',
                f'   • 철근 재료: fy={fy:,.0f} MPa, Es={Es:,.0f} MPa {steel_note}',
                f'   • 작용 하중: Pu={Pu:,.1f} kN, Mu={Mu:,.1f} kN·m (편심 e={e_actual:.3f} mm)',
                f'   • 가정된 중립축: c={c_assumed:.3f} mm',
                '',
                '2. 변형률 호환 및 응력 계산',
                f'   • 변형률 계산: εs = εcu × (c - ds) / c',
                f'   • 압축측 철근 (ds={dsi[0,0]:.1f}mm): εsc={ep_si[0,0]:.5f} → fsc={fsi[0,0]:,.2f} MPa',
                f'   • 인장측 철근 (dt={dsi[0,1]:.1f}mm): εst={ep_si[0,1]:.5f} → fst={fsi[0,1]:,.2f} MPa',
                '',
                '3. 단면력 평형 및 공칭강도 계산',
                f'   • 등가응력블록 깊이: a = β₁ × c = {beta1:.3f} × {c_assumed:.3f} = {a:.3f} mm',
                f'   • 콘크리트 압축면적: Ac = min(a, h) × b = {min(a, h):.1f} × {b:.1f} = {Ac:,.1f} mm²',
                f'   • 콘크리트 압축력: Cc = η × 0.85 × fck × Ac = {eta:.3f} × 0.85 × {fck:.1f} × {Ac:,.1f} = {Cc:,.1f} kN',
                f'   • 압축측 철근 합력: Cs = {Cs_force:,.1f} kN',
                f'   • 인장측 철근 합력: Ts = {Ts_force:,.1f} kN',
                f'   • 공칭 축강도: Pn = Cc + Cs + Ts = {Cc:,.1f}{Cs_force:+.1f}{Ts_force:+.1f} = {Pn:,.1f} kN',
                '',
                '4. 공칭 휨강도 계산',
                f'   • 콘크리트 압축력 중심: ȳ = (h/2) - (a/2) = ({h:.1f}/2) - ({a:.1f}/2) = {y_bar:.1f} mm',
                f'   • 압축철근 모멘트팔: (h/2) - ds1 = ({h:.1f}/2) - {dsi[0,0]:.1f} = {(h/2 - dsi[0,0]):.1f} mm',
                f'   • 인장철근 모멘트팔: (h/2) - dt = ({h:.1f}/2) - {dsi[0,1]:.1f} = {(h/2 - dsi[0,1]):.1f} mm',
                f'   • 공칭 휨강도: Mn = {Mn:,.1f} kN·m',
                '',
                '5. 강도감소계수 및 설계강도',
                f'   • 판단 근거: {phi_basis}',
                f'   • 설계 축강도: φPn = {phi_factor:.3f} × {Pn:,.1f} = {phiPn:,.1f} kN',
                f'   • 설계 휨강도: φMn = {phi_factor:.3f} × {Mn:,.1f} = {phiMn:,.1f} kN·m',
                '',
                '6. 최종 검토 및 안전성 평가',
                f'   • 축력 검토: Pu={Pu:,.1f} {"≤" if Pu <= phiPn else ">"} φPn={phiPn:,.1f} kN',
                f'   • 휨강도 검토: Mu={Mu:,.1f} {"≤" if Mu <= phiMn else ">"} φMn={phiMn:,.1f} kN·m',
                f'   • PM 교점 안전율: S.F. = {safety_factor:.3f} ({sf_status})',
                f'   • 계산편심: e\' = Mn/Pn × 1000 = {Mn:,.1f}/{Pn:,.1f} × 1000 = {(Mn*1000/Pn) if Pn != 0 else 0:.3f} mm',
                f'   • 작용편심: e = Mu/Pu × 1000 = {Mu:,.1f}/{Pu:,.1f} × 1000 = {e_actual:.3f} mm'
            ]
            
            for i, content in enumerate(calc_contents):
                current_row = calc_start_row + i
                if content:  # 빈 줄이 아닌 경우
                    column_ws.merge_range(current_row, start_col, current_row, start_col + 5, content, formats['calc_content'])
                    column_ws.set_row(current_row, 25)
                else:  # 빈 줄
                    column_ws.set_row(current_row, 12)
                    
            return len(calc_contents)
            
        except Exception as e:
            error_text = f'LC-{case_idx+1} {material_type} 계산 오류: {str(e)[:100]}'
            column_ws.merge_range(row, start_col, row, start_col + 5, error_text, formats['ng'])
            column_ws.set_row(row, 25)
            return 1
    
    # 각 하중조합별로 상세 계산 과정 작성
    try:
        num_cases = len(safe_extract(In, 'Pu'))
        for case_idx in range(num_cases):
            # 좌측: 이형철근
            lines_used_R = write_detailed_calculation(case_idx, '이형철근', R, 0)
            # 우측: 중공철근  
            lines_used_F = write_detailed_calculation(case_idx, '중공철근', F, 8)
            
            row += max(lines_used_R, lines_used_F) + 2  # 여백 추가
            
    except Exception as e:
        column_ws.merge_range(row, 0, row, max_col, f'상세 계산 작성 중 오류: {e}', formats['ng'])
        column_ws.set_row(row, 25)
        row += 1
    
    # ─── 8. 최종 종합 판정 ────────────────────────────
    column_ws.merge_range(row, 0, row, 6, '🎯 최종 종합 판정', formats['sub_header'])
    column_ws.merge_range(row, 8, row, max_col, '🎯 최종 종합 판정', formats['sub_header'])
    column_ws.set_row(row, 25)
    row += 1
    
    # 최종 판정
    for key, col_start, material_name in [('R', 0, '이형철근'), ('F', 8, '중공철근')]:
        if key in all_results and all_results[key]:
            final_pass = all(all_results[key])
            if final_pass:
                text = f'🎉 {material_name} - 전체 조건 만족 (구조 안전)'
                fmt = formats['final_ok']
            else:
                text = f'⚠️ {material_name} - 일부 조건 불만족 (보강 검토 필요)'
                fmt = formats['final_ng']
        else:
            text = f'❓ {material_name} - 검토 데이터 부족'
            fmt = formats['ng']
        
        column_ws.merge_range(row, col_start, row, col_start + 5, text, fmt)
        column_ws.set_row(row, 35)
    
    row += 2
    
    # ─── 9. 참고사항 ─────────────────────────────
    note_text = (
        "📋 검토 기준 및 참고사항\n\n"
        "🔍 PM 교점 안전율 판정 기준:\n"
        "  • S.F. = √[(φPn)² + (φMn)²] / √[Pu² + Mu²]\n"
        "  • S.F. ≥ 1.0 → PASS (구조적으로 안전)\n"
        "  • S.F. < 1.0 → FAIL (보강 검토 필요)\n\n"
        "🔧 철근 종류별 특성:\n"
        "  • 이형철근: 일반적인 SD400/SD500 철근 (표준 단면적)\n"
        "  • 중공철근: 내부가 비어있는 철근 (단면적 50% 적용, 항복강도 800 MPa)\n\n"
        f"📖 설계 기준: {getattr(In, 'RC_Code', 'KDS 41 17 00 (2021)')} (콘크리트구조 설계기준)\n"
        "📊 상세 분석 데이터: P-M Interaction Diagram 참조\n\n"
        "💡 변형률 호환 및 힘의 평형을 기반으로 한 정밀 해석 수행\n"
        "⚡ 강도감소계수(φ)는 KDS-2021 기준에 따라 변형률 조건별로 적용\n"
        "🎯 각 하중조합별 상세 계산 과정을 통한 투명한 검토 절차"
    )
    
    column_ws.merge_range(row, 0, row + 12, max_col, note_text, formats['calc_content'])
    for i in range(13):
        column_ws.set_row(row + i, 20)
    
    # ─── 10. 추가 스타일 적용 ─────────────────────────
    # 전체 시트에 격자 스타일 적용
    column_ws.set_margins(left=0.5, right=0.5, top=0.75, bottom=0.75)
    column_ws.set_header('&C&"맑은 고딕,Bold"&18🏗️ 기둥 강도 검토 보고서')
    column_ws.set_footer('&L&D &T&C&P / &N&R&"맑은 고딕"&12구조설계 전문가')
    
    # 인쇄 설정
    column_ws.set_landscape()
    column_ws.set_paper(9)  # A4
    column_ws.fit_to_pages(1, 0)  # 가로 1페이지에 맞춤
    
    # 보호 설정 (선택사항)
    # column_ws.protect('password', {'select_locked_cells': True, 'select_unlocked_cells': True})
    
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

            