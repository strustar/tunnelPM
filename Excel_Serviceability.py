import xlsxwriter

def create_serviceability_sheet(wb, In, R, F):
    """
    Streamlit 앱의 RC 사용성 검토 내용을 그대로 Excel 시트로 생성합니다.
    (이론적 배경, 상세 해석 과정, 균열 검토 결과 포함)
    """
    svc_ws = wb.add_worksheet('사용성 검토')

    # --- 1. 스타일 정의 (기존 코드와 동일) ---
    base_font = {'font_name': '맑은 고딕', 'border': 1, 'valign': 'vcenter', 'text_wrap': True}
    styles = {
        'title': {'bold': True, 'font_size': 18, 'bg_color': '#1F4E79', 'font_color': 'white', 'align': 'center'},
        'header_main': {'bold': True, 'font_size': 14, 'bg_color': '#2E75B6', 'font_color': 'white', 'align': 'center'},
        'header_part': {'bold': True, 'font_size': 16, 'bg_color': '#FFA500', 'font_color': 'black', 'align': 'center'},
        'header_theory': {'bold': True, 'font_size': 12, 'align': 'center'},
        'theory_box': {'bg_color': '#F8F9FA', 'border_color': '#DDDDDD'},
        'formula_box': {'bg_color': '#F0F2F6', 'align': 'center'},
        'case_special': {'bold': True, 'font_size': 12, 'bg_color': '#2E7D32', 'font_color': 'white', 'align': 'center'},
        'case_general': {'bold': True, 'font_size': 12, 'bg_color': '#1565C0', 'font_color': 'white', 'align': 'center'},
        'label': {'font_size': 11, 'align': 'left'},
        'value': {'bold': True, 'font_size': 11, 'align': 'right', 'num_format': '#,##0.0'},
        'ok': {'bold': True, 'font_size': 12, 'bg_color': '#D5EDDA', 'font_color': '#155724', 'align': 'center'},
        'ng': {'bold': True, 'font_size': 12, 'bg_color': '#F8D7DA', 'font_color': '#721C24', 'align': 'center'},
        'note': {'font_size': 10, 'valign': 'top', 'bg_color': '#F8F9FA'},
    }
    # 기본 폰트 및 테두리 일괄 적용
    formats = {name: wb.add_format({**base_font, **prop}) for name, prop in styles.items()}
    subscript_format = wb.add_format({'font_name': '맑은 고딕', 'font_script': 2}) # 아래첨자 전용

    # --- 2. 컬럼 너비 설정 ---
    col_widths = ['A:A', 'B:B', 'C:C', 'D:D', 'E:E', 'F:F', 'G:G']
    col_values = [4, 28, 12, 4, 28, 12, 4]
    for i, (col, width) in enumerate(zip(col_widths, col_values)):
        svc_ws.set_column(col, width)
        svc_ws.set_column(chr(ord('A') + i + 7) + ':' + chr(ord('A') + i + 7), width) # H열부터 동일하게 적용
    
    # --- 3. 헬퍼 함수 ---
    def write_rich(r, c, base, sub, unit="", fmt=formats['label']):
        """아래첨자 텍스트 작성"""
        svc_ws.write_rich_string(r, c, fmt, base, subscript_format, sub, fmt, unit)

    # --- 4. 시트 내용 작성 ---
    row = 0
    max_col = 14
    
    # ─── 메인 타이틀 ───
    svc_ws.merge_range(row, 0, row, max_col-1, '🏗️ RC 사용성 검토: 응력 및 균열 제어', formats['title'])
    svc_ws.set_row(row, 36)
    row += 2

    # ─── Part 1 & 2: 이론적 배경 ───
    ws_theory = wb.add_worksheet('이론적 배경') # 별도 시트로 분리하여 깔끔하게 정리
    ws_theory.set_column('A:A', 60)
    ws_theory.set_column('B:B', 60)
    
    # --- 이론 시트 내용 작성 ---
    theory_row = 0
    ws_theory.merge_range(theory_row, 0, theory_row, 1, 'Part 1. 탄성 이론 기반 응력 해석', formats['header_part'])
    theory_row +=1
    
    # Case I
    ws_theory.write(theory_row, 0, '🎯 Case Ⅰ: 특수한 경우 (순수 휨)', formats['case_special'])
    ws_theory.write(theory_row+1, 0, '핵심 원리: C = T (내부 압축력 = 인장력)', formats['header_theory'])
    ws_theory.write(theory_row+2, 0, '중립축(x) 계산: ½ b x² = n As (d-x)', formats['formula_box'])
    ws_theory.write(theory_row+3, 0, '응력(fs) 계산: fs = Mo / [ As (d - x/3) ]', formats['formula_box'])

    # Case II
    ws_theory.write(theory_row, 1, '⚙️ Case Ⅱ: 일반적인 경우 (축력+휨)', formats['case_general'])
    ws_theory.write(theory_row+1, 1, '핵심 원리: 축력/모멘트 동시 평형', formats['header_theory'])
    ws_theory.write(theory_row+2, 1, '축력: P₀ = C - T\n모멘트: M₀ = C(h/2-x/3) + T(d-h/2)', formats['formula_box'])
    ws_theory.write(theory_row+3, 1, '해법: 비선형 연립방정식으로 수치해석 필요', formats['formula_box'])
    theory_row += 5

    ws_theory.merge_range(theory_row, 0, theory_row, 1, 'Part 2. 휨균열 제어를 위한 철근 간격 검토', formats['header_part'])
    theory_row += 1
    ws_theory.write(theory_row, 0, '최외단 철근응력 (fst) 산정', formats['header_theory'])
    ws_theory.write(theory_row+1, 0, 'fst = fs · (h - dc - x) / (d - x)', formats['formula_box'])
    ws_theory.write(theory_row, 1, '최대 허용간격 (s) 산정 [KDS 기준]', formats['header_theory'])
    ws_theory.write(theory_row+1, 1, 's ≤ min [ 375(210/fst) - 2.5Cc ,  300(210/fst) ]', formats['formula_box'])
    
    # ─── Part 3: 상세 해석 및 균열 검토 ───
    svc_ws.merge_range(row, 0, row, max_col-1, 'Part 3. 하중 케이스별 상세 해석 및 균열 검토', formats['header_part'])
    svc_ws.set_row(row, 28)
    row += 1

    datasets = {'R': {'data': R, 'col_offset': 0, 'name': '이형철근'},
                'F': {'data': F, 'col_offset': 7, 'name': '중공철근'}}
    
    # R과 F 데이터의 상세 결과를 그리는 내부 함수
    def render_analysis_block(start_row, col_offset, rebar_name, data_source, case_idx):
        r = start_row
        c = col_offset
        
        fs, x = data_source.fs[case_idx], data_source.x[case_idx]
        P0, M0 = In.P0[case_idx], In.M0[case_idx]
        
        # 케이스 헤더
        case_format = formats['case_special'] if P0 == 0 else formats['case_general']
        case_title = f"{'🎯' if P0 == 0 else '⚙️'} {rebar_name} {num_symbols[case_idx]}번 검토"
        svc_ws.merge_range(r, c, r, c + 5, case_title, case_format)
        svc_ws.set_row(r, 22)
        r += 1

        # A. 탄성 해석
        svc_ws.merge_range(r, c, r, c+5, '🔬 A. 탄성 해석 과정', formats['header_main'])
        r += 1
        write_rich(r, c, '하중조건 P', '₀', f' = {P0:,.1f} kN', fmt=formats['label'])
        write_rich(r, c, ' / M', '₀', f' = {M0:,.1f} kN·m', fmt=formats['label'])
        r += 1
        write_rich(r, c+1, '중립축 x = ', '', f'{x:.1f} mm', fmt=formats['value'])
        write_rich(r, c+1, '철근응력 f', 's', f' = {fs:.1f} MPa', fmt=formats['value'])
        r += 2

        # B. 휨균열 제어 검토
        svc_ws.merge_range(r, c, r, c+5, '📏 B. 휨균열 제어 검토', formats['header_main'])
        r += 1

        fst = fs # 1단 배근 가정
        s1 = 375 * (210 / fst) - 2.5 * In.Cc if fst > 0 else float('inf')
        s2 = 300 * (210 / fst) if fst > 0 else float('inf')
        s_final = min(s1, s2)
        is_ok = In.sb[0] <= s_final
        
        write_rich(r, c, '최외단 응력 f', 'st', f' = {fst:.1f} MPa', fmt=formats['label'])
        r += 1
        write_rich(r, c, '허용간격 s', '1', f' = {s1:.1f} mm', fmt=formats['label'])
        write_rich(r, c, ' / s', '2', f' = {s2:.1f} mm', fmt=formats['label'])
        r += 1
        write_rich(r, c, '최종 허용간격 s', 'allow', f' = {s_final:.1f} mm', fmt=formats['label'])
        write_rich(r+1, c, '실제 배근간격 s', 'actual', f' = {In.sb[0]:.1f} mm', fmt=formats['label'])
        r += 2

        # 최종 판정
        result_text = f"✅ O.K. ({In.sb[0]:.1f} ≤ {s_final:.1f})" if is_ok else f"❌ N.G. ({In.sb[0]:.1f} > {s_final:.1f})"
        result_format = formats['ok'] if is_ok else formats['ng']
        svc_ws.merge_range(r, c, r, c+5, result_text, result_format)
        svc_ws.set_row(r, 24)
        r += 2 # 다음 블록을 위한 간격
        return r

    num_symbols = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"]
    current_row_r, current_row_f = row, row # 각 컬럼의 행 위치를 독립적으로 관리

    for i in range(len(In.P0)):
        # R 데이터 (이형철근) 블록 렌더링
        next_row_r = render_analysis_block(current_row_r, datasets['R']['col_offset'], datasets['R']['name'], R, i)
        # F 데이터 (중공철근) 블록 렌더링
        next_row_f = render_analysis_block(current_row_f, datasets['F']['col_offset'], datasets['F']['name'], F, i)
        
        # 각 컬럼의 행 위치 업데이트
        current_row_r = next_row_r
        current_row_f = next_row_f

    return svc_ws