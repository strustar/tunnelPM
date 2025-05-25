def create_review_sheet(wb, In, R, F):
    """검토결과 시트 생성 - 아래첨자 적용 및 최적화"""
    
    review_ws = wb.add_worksheet('검토결과')
    # review_ws.activate()
    
    # ─── 스타일 정의 ─────────────────────────────
    base_font = {'font_name': '맑은 고딕', 'border': 1, 'valign': 'vcenter'}
    
    styles = {
        'title': {**base_font, 'bold': True, 'font_size': 18, 'bg_color': '#1F4E79', 
                'font_color': 'white', 'border': 3, 'align': 'center'},
        'main_header': {**base_font, 'bold': True, 'font_size': 15, 'bg_color': '#2E75B6', 
                    'font_color': 'white', 'border': 2, 'align': 'center'},
        'section': {**base_font, 'bold': True, 'font_size': 15, 'bg_color': '#5B9BD5', 
                   'font_color': 'white', 'align': 'center'},
        'common': {**base_font, 'bold': True, 'font_size': 15, 'bg_color': '#70AD47', 
                  'font_color': 'white', 'align': 'center'},
        'common_sub': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#E2EFDA', 
                      'font_color': '#375623', 'align': 'center'},
        'sub_header': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#DEEBF7', 
                      'font_color': '#1F4E79', 'align': 'center'},
        'label': {**base_font, 'font_size': 12, 'align': 'left'},
        'value': {**base_font, 'bold': True, 'font_size': 12, 'align': 'center'},
        'number': {**base_font, 'bold': True, 'font_size': 12, 'num_format': '#,##0.0', 
                  'align': 'center'},
        'unit': {**base_font, 'font_size': 12, 'font_color': '#000000', 'align': 'center'},
        'combo': {**base_font, 'bold': True, 'font_size': 12, 'bg_color': '#FFF3CD', 
                 'font_color': '#856404', 'align': 'center'},
        'ok': {**base_font, 'bold': True, 'font_size': 13, 'bg_color': '#D5EDDA', 
              'font_color': '#155724', 'border': 2, 'align': 'center'},
        'ng': {**base_font, 'bold': True, 'font_size': 13, 'bg_color': '#F8D7DA', 
              'font_color': '#721C24', 'border': 2, 'align': 'center'},
        'final_ok': {**base_font, 'bold': True, 'font_size': 16, 'bg_color': '#D1F2EB', 
                    'font_color': '#0C6B40', 'border': 3, 'align': 'center'},
        'final_ng': {**base_font, 'bold': True, 'font_size': 16, 'bg_color': '#FADBD8', 
                    'font_color': '#A93226', 'border': 3, 'align': 'center'},
        'note': {**base_font, 'font_size': 11, 'align': 'left', 'valign': 'top', 
                'bg_color': '#F8F9FA', 'font_color': '#495057', 'text_wrap': True},
        'subscript': {**base_font, 'font_size': 10, 'align': 'left', 'font_script': 2}  # 아래첨자용
    }
    
    formats = {name: wb.add_format(props) for name, props in styles.items()}
    
    # 아래첨자 포맷 추가
    label_format = formats['label']
    subscript_format = wb.add_format({**styles['label'], 'font_script': 2})
    
    # ─── 컬럼 너비 설정 ─────────────────────────────
    col_widths = {
        'A': 24, 'B': 18, 'C': 18,  # 왼쪽 공통 데이터
        'D': 15, 'E': 14, 'F': 14,  # 왼쪽 철근 데이터
        'G': 8,                      # 구분선
        'H': 24, 'I': 18, 'J': 18,  # 오른쪽 공통 데이터
        'K': 15, 'L': 14, 'M': 14   # 오른쪽 철근 데이터
    }
    
    for col, width in col_widths.items():
        review_ws.set_column(f'{col}:{col}', width)
    
    # Rich string 작성 헬퍼 함수
    def write_with_subscript(row, col, text):
        """아래첨자가 필요한 레이블 작성"""
        if 'be' in text:
            review_ws.write_rich_string(row, col, '단위폭 b', subscript_format, 'e', label_format)
        elif 'dc\'' in text:
            review_ws.write_rich_string(row, col, '피복두께 d', subscript_format, 'c', label_format, '\'')
        elif 'dc' in text:
            review_ws.write_rich_string(row, col, '피복두께 d', subscript_format, 'c', label_format)
        elif 'fck' in text:
            review_ws.write_rich_string(row, col, '압축강도 f', subscript_format, 'ck', label_format)
        elif 'Ec' in text:
            review_ws.write_rich_string(row, col, '탄성계수 E', subscript_format, 'c', label_format)
        elif 'Es' in text:
            review_ws.write_rich_string(row, col, '탄성계수 E', subscript_format, 's', label_format)
        elif 'fy' in text:
            review_ws.write_rich_string(row, col, '항복강도 f', subscript_format, 'y', label_format)
        elif 'Pb' in text:
            review_ws.write_rich_string(row, col, '⚖️ 축력 P', subscript_format, 'b', label_format)
        elif 'Mb' in text:
            review_ws.write_rich_string(row, col, '📏 모멘트 M', subscript_format, 'b', label_format)
        elif 'eb' in text:
            review_ws.write_rich_string(row, col, '📐 편심 e', subscript_format, 'b', label_format)
        elif 'cb' in text:
            review_ws.write_rich_string(row, col, '🎯 중립축 깊이 c', subscript_format, 'b', label_format)
        else:
            review_ws.write(row, col, text, formats['label'])
    
    row = 0
    max_col = 12  # Column M
    
    # ─── 1. 메인 타이틀 ───────────────────────────────
    review_ws.merge_range(row, 0, row, max_col, '🏗️ 구조부재 강도 검토 보고서', formats['title'])
    review_ws.set_row(row, 40)
    row += 2
    
    # ─── 2. 공통 설계 조건 (좌우 배치) ─────────────────
    review_ws.merge_range(row, 0, row, max_col, '◈ 공통 설계 조건', formats['common'])
    review_ws.set_row(row, 24)
    row += 1
    
    # 좌측 섹션 데이터
    left_sections = [
        ('📐 단면 제원', [
            ('단위폭 be', In.be, 'mm'),
            ('단면 두께 h', In.height, 'mm'),
            ('공칭 철근간격 s', In.sb[0], 'mm'),
        ]),
        ('🏭 콘크리트 재료 특성', [
            ('압축강도 fck', In.fck, 'MPa'),
            ('탄성계수 Ec', In.Ec/1000, 'GPa'),
        ]),
        ('📋 설계 조건', [
            ('설계방법', In.Design_Method.split('(')[0].strip(), ''),
            ('설계기준', In.RC_Code, ''),
            ('기둥형식', In.Column_Type, ''),
        ])
    ]
    
    # 우측 섹션 데이터 (철근 배치 통합)
    right_sections = [
        ('🔩 철근 배치 (인장측)', [
            ('철근 직경 D', In.dia[0], 'mm'),
            ('피복두께 dc', In.dc[0], 'mm'),
            ('유효깊이 d', In.height - In.dc[0], 'mm'),
        ]),
        ('🔩 철근 배치 (압축측)', [
            ('철근 직경 D\'', In.dia1[0], 'mm'),
            ('피복두께 dc\'', In.dc1[0], 'mm'),
            ('압축철근 깊이 d\'', In.dc1[0], 'mm'),
        ])
    ]
    
    # 좌우 배치 함수 (최적화)
    def write_lr_sections(left_data, right_data):
        nonlocal row
        start_row = row
        max_row = start_row
        
        # 좌측 섹션
        left_row = start_row
        for section_title, items in left_data:
            review_ws.merge_range(left_row, 0, left_row, 2, section_title, formats['common_sub'])
            review_ws.set_row(left_row, 22)
            left_row += 1
            
            for label, value, unit in items:
                write_with_subscript(left_row, 0, label)
                fmt = formats['number'] if isinstance(value, (int, float)) and unit else formats['value']
                review_ws.write(left_row, 1, value, fmt)
                review_ws.write(left_row, 2, unit, formats['unit'])
                review_ws.set_row(left_row, 20)
                left_row += 1
            left_row += 1
        
        # 우측 섹션
        right_row = start_row
        for section_title, items in right_data:
            review_ws.merge_range(right_row, 7, right_row, 9, section_title, formats['common_sub'])
            right_row += 1
            
            for label, value, unit in items:
                write_with_subscript(right_row, 7, label)
                fmt = formats['number'] if isinstance(value, (int, float)) and unit else formats['value']
                review_ws.write(right_row, 8, value, fmt)
                review_ws.write(right_row, 9, unit, formats['unit'])
                right_row += 1
            right_row += 1
        
        row = max(left_row, right_row)
    
    write_lr_sections(left_sections, right_sections)
    
    # ─── 3. 철근별 상세 조건 ─────────────────────────
    review_ws.merge_range(row, 0, row, 5, '📊 이형철근 검토', formats['main_header'])
    review_ws.merge_range(row, 7, row, 12, '📊 중공철근 검토', formats['main_header'])
    review_ws.set_row(row, 28)
    row += 2
    
    # 재료 특성
    review_ws.merge_range(row, 0, row, 5, '◈ 철근 재료 특성', formats['section'])
    review_ws.merge_range(row, 7, row, 12, '◈ 철근 재료 특성', formats['section'])
    review_ws.set_row(row, 24)
    row += 1
    
    material_data = [
        ('항복강도 fy', In.fy, In.fy_hollow, 'MPa'),
        ('탄성계수 Es', In.Es/1000, In.Es_hollow/1000, 'GPa'),
    ]
    
    for label, vR, vF, unit in material_data:
        # 좌측 (이형철근)
        write_with_subscript(row, 0, label)
        review_ws.write(row, 1, vR, formats['number'])
        review_ws.write(row, 2, unit, formats['unit'])
        # 우측 (중공철근)
        write_with_subscript(row, 7, label)
        review_ws.write(row, 8, vF, formats['number'])
        review_ws.write(row, 9, unit, formats['unit'])
        review_ws.set_row(row, 20)
        row += 1
    row += 1
    
    # ─── 4. 평형상태 검토 ─────────────────────────────
    review_ws.merge_range(row, 0, row, 5, '◈ 평형상태 검토', formats['section'])
    review_ws.merge_range(row, 7, row, 12, '◈ 평형상태 검토', formats['section'])
    review_ws.set_row(row, 24)
    row += 1
    
    equilibrium_data = [
        ('⚖️ 축력 Pb', R.Pd[3], F.Pd[3], 'kN'),
        ('📏 모멘트 Mb', R.Md[3], F.Md[3], 'kN·m'),
        ('📐 편심 eb', R.e[3], F.e[3], 'mm'),
        ('🎯 중립축 깊이 cb', R.c[3], F.c[3], 'mm'),
    ]
    
    for label, vR, vF, unit in equilibrium_data:
        # 좌측 (이형철근)
        write_with_subscript(row, 0, label)
        review_ws.write(row, 1, vR, formats['number'])
        review_ws.write(row, 2, unit, formats['unit'])
        # 우측 (중공철근)
        write_with_subscript(row, 7, label)
        review_ws.write(row, 8, vF, formats['number'])
        review_ws.write(row, 9, unit, formats['unit'])
        review_ws.set_row(row, 20)
        row += 1
    row += 1
    
    # ─── 5. 기둥강도 검토 결과 ─────────────────────────
    review_ws.merge_range(row, 0, row, 5, '◈ 기둥강도 검토 결과', formats['section'])
    review_ws.merge_range(row, 7, row, 12, '◈ 기둥강도 검토 결과', formats['section'])
    review_ws.set_row(row, 24)
    row += 1
    
    # 헤더 작성 (아래첨자 적용)
    headers = ['하중조합', 'Pu / ϕPn', 'Mu / ϕMn', '편심 e', '안전률 SF', '판정']
    header_units = ['', '[kN]', '[kN·m]', '[mm]', '', '']
    
    for i, (hdr, unit) in enumerate(zip(headers, header_units)):
        if i == 1:  # Pu / ϕPn
            review_ws.write_rich_string(row, i, 'P', subscript_format, 'u', label_format, ' / ϕP', subscript_format, 'n', label_format, ' ' + unit, formats['sub_header'])
            review_ws.write_rich_string(row, i + 7, 'P', subscript_format, 'u', label_format, ' / ϕP', subscript_format, 'n', label_format, ' ' + unit, formats['sub_header'])
        elif i == 2:  # Mu / ϕMn
            review_ws.write_rich_string(row, i, 'M', subscript_format, 'u', label_format, ' / ϕM', subscript_format, 'n', label_format, ' ' + unit, formats['sub_header'])
            review_ws.write_rich_string(row, i + 7, 'M', subscript_format, 'u', label_format, ' / ϕM', subscript_format, 'n', label_format, ' ' + unit, formats['sub_header'])
        else:
            review_ws.write(row, i, hdr + ' ' + unit if unit else hdr, formats['sub_header'])
            review_ws.write(row, i + 7, hdr + ' ' + unit if unit else hdr, formats['sub_header'])
    review_ws.set_row(row, 22)
    row += 1
    
    # 결과 저장
    all_results = {'R': [], 'F': []}
    num_load_cases = len(In.Pu) if hasattr(In, 'Pu') and hasattr(In, 'Mu') and len(In.Pu) == len(In.Mu) else 0
    
    for i in range(num_load_cases):
        Pu, Mu = In.Pu[i], In.Mu[i]
        e = (Mu / Pu) * 1000 if Pu != 0 else 0
        
        sR = In.safe_RC[i]
        sF = In.safe_FRP[i]
        
        R_pass = sR >= 1.0
        F_pass = sF >= 1.0
        all_results['R'].append(R_pass)
        all_results['F'].append(F_pass)
        
        # 값/값 형식으로 표시
        Pu_str_R = f'{Pu:,.1f} / {In.Pd_RC[i]:,.1f}'
        Mu_str_R = f'{Mu:,.1f} / {In.Md_RC[i]:,.1f}'
        Pu_str_F = f'{Pu:,.1f} / {In.Pd_FRP[i]:,.1f}'
        Mu_str_F = f'{Mu:,.1f} / {In.Md_FRP[i]:,.1f}'
        
        # 좌측 (이형철근) 결과
        result_data = [
            (f'LC-{i+1}', formats['combo']),
            (Pu_str_R, formats['value']),
            (Mu_str_R, formats['value']),
            (e, formats['number']),
            (f'{sR:.2f}', formats['number']),
            ('PASS ✅' if R_pass else 'FAIL ❌', formats['ok'] if R_pass else formats['ng'])
        ]
        
        for j, (val, fmt) in enumerate(result_data):
            review_ws.write(row, j, val, fmt)
        
        # 우측 (중공철근) 결과
        result_data[1] = (Pu_str_F, formats['value'])
        result_data[2] = (Mu_str_F, formats['value'])
        result_data[4] = (f'{sF:.2f}', formats['number'])
        result_data[5] = ('PASS ✅' if F_pass else 'FAIL ❌', formats['ok'] if F_pass else formats['ng'])
        
        for j, (val, fmt) in enumerate(result_data):
            review_ws.write(row, j + 7, val, fmt)
            
        review_ws.set_row(row, 24)
        row += 1
    row += 1
    
    # ─── 6. 최종 종합 판정 ────────────────────────────
    review_ws.merge_range(row, 0, row, 5, '◈ 최종 종합 판정', formats['section'])
    review_ws.merge_range(row, 7, row, 12, '◈ 최종 종합 판정', formats['section'])
    review_ws.set_row(row, 24)
    row += 1
    
    # 최종 판정
    for key, col_start in [('R', 0), ('F', 7)]:
        final_pass = all(all_results[key]) if all_results[key] else False
        text = '🎉 전체 조건 만족 - 구조 안전' if final_pass else '⚠️ 일부 조건 불만족 - 보강 검토 필요'
        fmt = formats['final_ok'] if final_pass else formats['final_ng']
        review_ws.merge_range(row, col_start, row, col_start + 5, text, fmt)
    
    review_ws.set_row(row, 34)
    row += 2
    
    # ─── 7. 참고사항 ─────────────────────────────
    note_text = (
        "📋 검토 기준 및 참고사항\n\n"
        "🔍 안전률(SF) 판정 기준:\n"
        "  • SF ≥ 1.0 → PASS (구조적으로 안전)\n"
        "  • SF < 1.0 → FAIL (보강 검토 필요)\n"
        "  • SF는 작용하중(Pu, Mu)에 대응하는 P-M 선도상의 설계강도(ϕPn, ϕMn)를 이용하여 산정됩니다.\n"
        "  • 편심거리 e = Mu / Pu (단위: mm)\n\n"
        f"📖 설계 기준: {getattr(In, 'RC_Code', 'KDS 41 17 00 (2021)')} (콘크리트구조 설계기준)\n"
        "📊 상세 분석 데이터: '데이터' 시트 참조 (P-M Interaction Diagram 등)"
    )
    review_ws.merge_range(row, 0, row + 5, max_col, note_text, formats['note'])
    for i in range(6):
        review_ws.set_row(row + i, 20)
    
    return review_ws

