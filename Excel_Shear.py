import numpy as np

def create_shear_sheet(wb, In, R):
    check_type = In.check_type

    """
    Excel 전단설계 보고서 - 가독성 최적화 버전
    """

    shear_ws = wb.add_worksheet('전단설계 최적화 보고서')

    # ─── 최적화된 색상 팔레트 (가독성 및 대비도 개선) ─────────────────────────
    colors = {
        # 메인 컬러 (더 진하고 선명하게)
        'navy_deep': '#1e3a8a',      # 진한 파란색
        'navy_medium': '#3730a3',     # 중간 파란색
        'section_dark': '#1d4ed8',    # 섹션 헤더용 파란색
        'section_medium': '#2563eb',  # 중간 섹션 색상

        # 판정 기준 컬러 (더 선명하게)
        'criteria_blue': '#3b82f6',   # 파란색 기준
        'criteria_purple': '#7c3aed', # 보라색 기준  
        'criteria_rose': '#e11d48',   # 분홍색 기준

        # 테이블 컬러 (더 진하게)
        'table_header': '#059669',    # 녹색 헤더
        'table_data': '#f8fafc',      # 밝은 데이터 배경
        'table_alt': '#f1f5f9',       # 교대 배경

        # 상태 컬러 (더 선명하게)
        'success': '#059669',         # 성공 색상
        'success_bg': '#d1fae5',      # 성공 배경
        'danger': '#dc2626',          # 실패 색상  
        'danger_bg': '#fee2e2',       # 실패 배경

        # 케이스 상세 계산 색상 (대비도 향상)
        'case_title': '#cbd5e1',      # 케이스 제목 배경
        'case_result': '#f1f5f9',     # 결과 배경
        'step_header': '#dbeafe',     # 단계 헤더 배경
        'calc_block': '#ffffff',      # 계산 블록 배경
        'formula': '#f3f4f6',         # 수식 배경
        'warning': '#fef3c7',         # 경고 배경
        'sub_header': '#e2e8f0',      # 서브헤더 배경

        # 텍스트 컬러 (더 진하게)
        'text_dark': '#0f172a',       # 진한 텍스트
        'text_medium': '#374151',     # 중간 텍스트
        'text_light': '#6b7280',      # 밝은 텍스트
    }

    base_font = {'font_name': 'Noto Sans KR', 'border': 1, 'valign': 'vcenter', 'bold': True}

    styles = {
        # 메인 타이틀
        'main_title': {**base_font, 'font_size': 28, 'bg_color': colors['navy_deep'], 
                      'font_color': 'white', 'border': 2, 'align': 'center'},
        'sub_title': {**base_font, 'font_size': 16, 'bg_color': colors['navy_medium'], 
                     'font_color': 'white', 'align': 'center'},
        'section_header': {**base_font, 'font_size': 20, 'bg_color': colors['section_dark'], 
                          'font_color': 'white', 'border': 2, 'align': 'center'},

        # 판정 기준 박스들
        'criteria_no_shear': {**base_font, 'font_size': 14, 'bg_color': colors['criteria_blue'], 
                             'font_color': 'white', 'align': 'center', 'text_wrap': True, 'border': 1},
        'criteria_min_shear': {**base_font, 'font_size': 14, 'bg_color': colors['criteria_purple'], 
                              'font_color': 'white', 'align': 'center', 'text_wrap': True, 'border': 1},
        'criteria_design_shear': {**base_font, 'font_size': 14, 'bg_color': colors['criteria_rose'], 
                                 'font_color': 'white', 'align': 'center', 'text_wrap': True, 'border': 1},

        # 요약 테이블
        'summary_header': {**base_font, 'font_size': 15, 'bg_color': colors['table_header'], 
                          'font_color': 'white', 'border': 2, 'align': 'center'},
        'summary_data': {**base_font, 'font_size': 13, 'bg_color': colors['table_data'], 
                        'font_color': colors['text_dark'], 'align': 'center', 'border': 1},
        'summary_data_alt': {**base_font, 'font_size': 13, 'bg_color': colors['table_alt'], 
                            'font_color': colors['text_dark'], 'align': 'center', 'border': 1},

        # 판정결과 색상
        'category_no_shear': {**base_font, 'font_size': 13, 'bg_color': '#dbeafe', 
                             'font_color': colors['criteria_blue'], 'align': 'center', 'border': 1},
        'category_min_shear': {**base_font, 'font_size': 13, 'bg_color': '#e0e7ff', 
                              'font_color': colors['criteria_purple'], 'align': 'center', 'border': 1},
        'category_design_shear': {**base_font, 'font_size': 13, 'bg_color': '#fce7f3', 
                                 'font_color': colors['criteria_rose'], 'align': 'center', 'border': 1},

        # 최종 판정 상태
        'status_ok': {**base_font, 'font_size': 13, 'bg_color': colors['success_bg'], 
                     'font_color': colors['success'], 'border': 2, 'align': 'center'},
        'status_ng': {**base_font, 'font_size': 13, 'bg_color': colors['danger_bg'], 
                     'font_color': colors['danger'], 'border': 2, 'align': 'center'},

        # 케이스별 상세 계산
        # 케이스 제목 전용 (크기 및 배경색 개선)
        'case_title_main': {**base_font, 'font_size': 16, 'bg_color': '#3b82f6', 
                           'font_color': 'white', 'border': 2, 'align': 'center'},
        'case_title': {**base_font, 'font_size': 14, 'bg_color': colors['case_title'], 
                      'font_color': colors['text_dark'], 'border': 1, 'align': 'left'},
        'case_result': {**base_font, 'font_size': 13, 'bg_color': colors['case_result'], 
                       'font_color': colors['text_dark'], 'border': 1, 'align': 'left'},

        # 단계별 계산 테이블 헤더
        'calc_table_header': {**base_font, 'font_size': 13, 'bg_color': colors['table_header'], 
                             'font_color': 'white', 'border': 2, 'align': 'center'},
        
        # 단계 헤더 (더 밝은 배경)
        'step_header_bright': {**base_font, 'font_size': 14, 'bg_color': '#dbeafe', 
                              'font_color': colors['section_dark'], 'border': 1, 'align': 'left'},
        
        # 서브 헤더 (더 진하게)
        'sub_header_bold': {**base_font, 'font_size': 12, 'bg_color': '#f1f5f9', 
                           'font_color': colors['text_dark'], 'border': 1, 'align': 'center'},

        # 계산 블록 (진하게 처리)
        'calc_content': {**base_font, 'font_size': 12, 'bg_color': colors['calc_block'], 
                        'font_color': colors['text_dark'], 'align': 'left', 'text_wrap': True, 
                        'border': 1, 'bold': True},

        # 수식 스타일 (이탤릭 제거, 진하게 처리)
        'formula_wide': {**base_font, 'font_size': 12, 'bg_color': colors['formula'], 'align': 'left', 
                        'font_color': colors['text_dark'], 'border': 1, 'bold': True, 
                        'text_wrap': True},

        # 결과 표시 (진하게 처리)
        'result_wide': {**base_font, 'font_size': 12, 'bg_color': colors['table_data'], 
                       'font_color': colors['text_dark'], 'border': 1, 'align': 'center', 
                       'text_wrap': True, 'bold': True},

        # 최종 결과 박스
        'final_success': {**base_font, 'font_size': 15, 'bg_color': colors['success_bg'], 
                         'font_color': colors['success'], 'border': 2, 'align': 'center', 'text_wrap': True},
        'final_fail': {**base_font, 'font_size': 15, 'bg_color': colors['danger_bg'], 
                      'font_color': colors['danger'], 'border': 2, 'align': 'center', 'text_wrap': True},

        # 기타
        'criteria_selection_header': {**base_font, 'font_size': 18, 'bg_color': colors['section_dark'], 
                                     'font_color': 'white', 'border': 2, 'align': 'center'},
        'criteria_selected': {**base_font, 'font_size': 14, 'bg_color': '#dbeafe', 
                             'font_color': colors['navy_deep'], 'border': 1, 'align': 'center'},
        'criteria_unselected': {**base_font, 'font_size': 14, 'bg_color': '#f0f9ff', 
                               'font_color': colors['text_medium'], 'border': 1, 'align': 'center'},
        'case_separator_line': {**base_font, 'font_size': 8, 'bg_color': '#e2e8f0', 
                               'font_color': '#e2e8f0', 'border': 0, 'align': 'center'},
        'step_number': {**base_font, 'font_size': 12, 'bg_color': colors['table_data'], 
                       'font_color': colors['text_dark'], 'border': 1, 'align': 'center', 
                       'text_wrap': True, 'bold': True},
    }

    formats = {name: wb.add_format(props) for name, props in styles.items()}

    # ─── 컬럼 너비 설정 (더 넓은 공간 확보) ─────────────────────────────
    shear_ws.set_column('A:A', 2)     # 여백
    shear_ws.set_column('B:B', 6)    # 단계 번호
    shear_ws.set_column('C:P', 16)    # 항목명
    # shear_ws.set_column('D:I', 20)    # 수식/계산 (병합 활용)
    # shear_ws.set_column('J:L', 20)    # 결과 (병합 활용)
    # shear_ws.set_column('M:P', 20)    # 설명 (병합 활용)

    row = 0
    max_col = 15

    # ─── 계산 헬퍼 함수들 ─────────────────────────────
    def format_number(num, decimal_places=1):
        return f"{num:,.{decimal_places}f}"

    def format_N_to_kN(value, dp=2):
        return f"{value/1000:,.{dp}f}"

    def format_load_condition_compact(Pu, Vu):
        return f"Pu={format_number(Pu, 0)}, Vu={format_number(Vu)}"

    # ─── 설계 상수 및 계산 로직 ─────────────────────────────
    phi_v = 0.75
    lamda = 1.0
    fy_shear = 400
    bar_dia = 13
    legs = 2
    bar_area = np.pi * (bar_dia / 2)**2
    Av_stirrup = bar_area * legs

    bw, d, fck, Ag = In.be, In.depth, In.fck, R.Ag

    results = []

    # 각 하중 케이스별 계산 수행
    for i in range(len(In.Vu)):
        Vu = In.Vu[i]
        Pu = In.Pu_shear[i]

        p_factor = 1 + (Pu * 1000) / (14 * Ag) if Pu != 0 else 1.0
        Vc = (1/6) * p_factor * lamda * np.sqrt(fck) * bw * d
        phi_Vc = phi_v * Vc
        half_phi_Vc = 0.5 * phi_Vc

        # 전단철근 필요성 판정
        if check_type == '프리캐스트 (3단계)':
            if Vu * 1000 <= half_phi_Vc:
                shear_category = "전단철근 불필요"
                category_fmt = formats['category_no_shear']
            elif Vu * 1000 <= phi_Vc:
                shear_category = "최소전단철근"
                category_fmt = formats['category_min_shear']
            else:
                shear_category = "설계전단철근"
                category_fmt = formats['category_design_shear']
        else:
            if Vu * 1000 <= phi_Vc:
                shear_category = "전단철근 불필요"
                category_fmt = formats['category_no_shear']
            else:
                shear_category = "설계전단철근"
                category_fmt = formats['category_design_shear']

        # 최소 전단철근량 계산
        min_Av_s_1_val = 0.0625 * np.sqrt(fck)
        min_Av_s_2_val = 0.35
        min_Av_s_1 = min_Av_s_1_val * (bw / fy_shear)
        min_Av_s_2 = min_Av_s_2_val * (bw / fy_shear)

        min_Av_s_req = max(min_Av_s_1, min_Av_s_2)
        s_from_min_req = Av_stirrup / min_Av_s_req

        # 설계 전단철근량 계산
        Vs_req = (Vu * 1000 - phi_Vc) / phi_v if shear_category == "설계전단철근" else 0
        s_from_vs_req = (Av_stirrup * fy_shear * d) / Vs_req if Vs_req > 0 else float('inf')

        # 최대 간격 제한
        Vs_limit_d4 = (1/3) * np.sqrt(fck) * bw * d
        s_max_code = min(d / 4, 300) if Vs_req > Vs_limit_d4 else min(d / 2, 600)

        # 최종 간격 결정
        if shear_category == "전단철근 불필요":
            actual_s = s_max_code
            stirrups_needed = "전단철근 불필요"
        elif shear_category == "최소전단철근":
            s_calc = s_from_min_req
            actual_s = min(s_calc, s_max_code)
            actual_s = np.floor(actual_s / 5) * 5
            stirrups_needed = f"H{bar_dia}-{legs}leg @{actual_s:.0f}"
        else:
            s_calc = min(s_from_min_req, s_from_vs_req)
            actual_s = min(s_calc, s_max_code)
            actual_s = np.floor(actual_s / 5) * 5
            stirrups_needed = f"H{bar_dia}-{legs}leg @{actual_s:.0f}"

        # 제공 전단강도 계산
        if shear_category == "전단철근 불필요":
            phi_Vs = 0
        else:
            phi_Vs = (phi_v * Av_stirrup * fy_shear * d) / actual_s if actual_s > 0 else 0

        phi_Vn = phi_Vc + phi_Vs

        # 안전성 검토
        is_safe_strength = (phi_Vn >= Vu * 1000)
        Vs_max_limit = (2/3) * np.sqrt(fck) * bw * d
        Vs_provided = phi_Vs / phi_v if phi_Vs > 0 else 0
        is_safe_section = (Vs_provided <= Vs_max_limit)
        is_safe_total = is_safe_strength and is_safe_section

        stirrups_per_meter = 1000 / actual_s if actual_s > 0 and shear_category != "전단철근 불필요" else 0

        # 판정 결과
        if not is_safe_section:
            final_status = "❌ NG (단면 부족)"
            ng_reason = f"전단철근이 부담하는 강도(Vs = {format_N_to_kN(Vs_provided, 1)} kN)가 최대 허용치(Vs,max = {format_N_to_kN(Vs_max_limit, 1)} kN)를 초과하여 단면 파괴가 우려됩니다."
        elif not is_safe_strength:
            final_status = "❌ NG (강도 부족)"
            ng_reason = f"설계 전단강도(φVn = {format_N_to_kN(phi_Vn, 1)} kN)가 요구 전단강도(Vu = {format_number(Vu, 1)} kN)보다 작아 안전하지 않습니다."
        else:
            final_status = "✅ OK"
            ng_reason = ""

        results.append({
            'case': i + 1, 'Pu': Pu, 'Vu': Vu, 'shear_category': shear_category,
            'category_fmt': category_fmt, 'phi_Vn_kN': phi_Vn / 1000, 
            'is_safe': is_safe_total, 'is_safe_section': is_safe_section, 'actual_s': actual_s,
            'stirrups_needed': stirrups_needed, 'stirrups_per_meter': stirrups_per_meter,
            'p_factor': p_factor, 'Vc_N': Vc, 'phi_Vc_N': phi_Vc, 'half_phi_Vc_N': half_phi_Vc,
            'Vs_req_N': Vs_req, 'min_Av_s_req': min_Av_s_req, 's_from_min_req': s_from_min_req,
            's_from_vs_req': s_from_vs_req, 's_max_code': s_max_code,
            'Vs_limit_d4_N': Vs_limit_d4, 'phi_Vs_N': phi_Vs,
            'Vs_provided_N': Vs_provided, 'Vs_max_limit_N': Vs_max_limit,
            'final_status': final_status, 'ng_reason': ng_reason,
            'min_Av_s_1_val': min_Av_s_1_val, 'min_Av_s_2_val': min_Av_s_2_val
        })

    # ═══════════════════════════════════════════════════════════════
    # 1. 메인 타이틀
    # ═══════════════════════════════════════════════════════════════
    shear_ws.merge_range(row, 0, row, max_col, '🛡️ 전단설계 최적화 보고서', formats['main_title'])
    shear_ws.set_row(row, 50)
    row += 1

    shear_ws.merge_range(row, 0, row, max_col, 'KDS 14 20 콘크리트구조설계기준 적용', formats['sub_title'])
    shear_ws.set_row(row, 28)
    row += 2

    # ═══════════════════════════════════════════════════════════════
    # 2. 전단철근 판정 기준 선택
    # ═══════════════════════════════════════════════════════════════
    shear_ws.merge_range(row, 0, row, max_col, '📋 전단철근 판정 기준 선택', formats['criteria_selection_header'])
    shear_ws.set_row(row, 32)
    row += 1

    if check_type == '일반 (2단계)':
        option_text = "● 일반 (2단계)                    ○ 프리캐스트 (3단계)"
    else:
        option_text = "○ 일반 (2단계)                    ● 프리캐스트 (3단계)"

    shear_ws.merge_range(row, 0, row, max_col, option_text, formats['criteria_selected'])
    shear_ws.set_row(row, 22)
    row += 1

    # note_text = "※ Excel에서는 라디오 버튼 클릭이 지원되지 않습니다."
    # shear_ws.merge_range(row, 0, row, max_col, note_text, formats['criteria_unselected'])
    shear_ws.set_row(row, 20)
    row += 2

    # ═══════════════════════════════════════════════════════════════
    # 3. 전단철근 판정 기준 박스들
    # ═══════════════════════════════════════════════════════════════
    if check_type == '프리캐스트 (3단계)':
        box_configs = [
            (1, 5, '🔵 전단철근 불필요', 'Vu ≤ ½φVc', '이론적으로 전단철근 불필요', 'criteria_no_shear'),
            (6, 10, '🟡 최소전단철근', '½φVc < Vu ≤ φVc', '규정 최소량 적용', 'criteria_min_shear'),
            (11, 15, '🔴 설계전단철근', 'Vu > φVc', '계산에 의한 철근량', 'criteria_design_shear')
        ]
    else:
        box_configs = [
            (1, 7, '🔵 전단철근 불필요', 'Vu ≤ φVc', '최소철근 배근 또는 불필요', 'criteria_no_shear'),
            (9, 15, '🔴 설계전단철근', 'Vu > φVc', '계산에 의한 철근량', 'criteria_design_shear')
        ]

    for col_start, col_end, title, condition, description, style_name in box_configs:
        shear_ws.merge_range(row, col_start, row, col_end, title, formats[style_name])
        shear_ws.merge_range(row + 1, col_start, row + 1, col_end, condition, formats[style_name])
        shear_ws.merge_range(row + 2, col_start, row + 2, col_end, description, formats[style_name])

    for i in range(3):
        shear_ws.set_row(row + i, 25)
    row += 4

    # ═══════════════════════════════════════════════════════════════
    # 4. 전체 설계 결과 요약
    # ═══════════════════════════════════════════════════════════════
    shear_ws.merge_range(row, 0, row, max_col, '📊 전체 설계 결과 요약', formats['section_header'])
    shear_ws.set_row(row, 35)
    row += 2

    headers = ['Case', '하중조건 (kN)', '판정결과', '최적 설계', '1m당 개수', '설계강도 (kN)', '최종 판정']
    header_start_cols = [1, 3, 5, 7, 9, 11, 13]

    for i, (header, col) in enumerate(zip(headers, header_start_cols)):
        shear_ws.merge_range(row, col, row, col + 1, header, formats['summary_header'])
    shear_ws.set_row(row, 28)
    row += 1

    for i, r in enumerate(results):
        data_fmt = formats['summary_data'] if i % 2 == 0 else formats['summary_data_alt']

        shear_ws.merge_range(row, 1, row, 2, f"Case {r['case']}", data_fmt)
        shear_ws.merge_range(row, 3, row, 4, format_load_condition_compact(r['Pu'], r['Vu']), data_fmt)
        shear_ws.merge_range(row, 5, row, 6, r['shear_category'], r['category_fmt'])
        shear_ws.merge_range(row, 7, row, 8, r['stirrups_needed'], data_fmt)

        count_text = f"{r['stirrups_per_meter']:.1f}개" if r['stirrups_per_meter'] > 0 else "—"
        shear_ws.merge_range(row, 9, row, 10, count_text, data_fmt)
        shear_ws.merge_range(row, 11, row, 12, f"φVn = {format_number(r['phi_Vn_kN'])}", data_fmt)

        status_fmt = formats['status_ok'] if r['is_safe'] else formats['status_ng']
        shear_ws.merge_range(row, 13, row, 14, r['final_status'], status_fmt)

        shear_ws.set_row(row, 32)
        row += 1

    row += 2

    # ═══════════════════════════════════════════════════════════════
    # 5. 케이스별 상세 계산 과정 (개선된 레이아웃)
    # ═══════════════════════════════════════════════════════════════
    shear_ws.merge_range(row, 0, row, max_col, '⚙️ 케이스별 상세 계산 과정', formats['section_header'])
    shear_ws.set_row(row, 35)
    row += 2

    case_symbols = ["❶", "❷", "❸", "❹", "❺", "❻", "❼", "❽", "❾", "❿"]

    for i, r in enumerate(results):
        # ═══ 케이스 구분선 ═══
        if i > 0:
            separator_text = "━" * 50
            shear_ws.merge_range(row, 0, row, max_col, separator_text, formats['case_separator_line'])
            shear_ws.set_row(row, 10)
            row += 1

        # ═══ 케이스 제목 ═══
        case_title = f'{case_symbols[i]} Case {r["case"]} 검토'
        # shear_ws.write(row, 1, case_title, formats['case_title'])
        shear_ws.merge_range(row, 1, row, max_col, f'{case_title} : {format_load_condition_compact(r["Pu"], r["Vu"])} kN', formats['case_title'])
        shear_ws.set_row(row, 28)
        row += 1

        # ═══ 결과 요약 ═══
        shear_ws.write(row, 1, '요약', formats['case_result'])
        shear_ws.merge_range(row, 2, row, max_col, f'{r["shear_category"]} / {r["stirrups_needed"]}', formats['case_result'])
        shear_ws.set_row(row, 25)
        row += 1

        # ═══ 단계별 계산 과정 (개선된 표 형식) ═══
        shear_ws.merge_range(row, 1, row, max_col, '단계별 계산 과정', formats['calc_table_header'])
        shear_ws.set_row(row, 25)
        row += 1

        # 표 헤더 (병합 활용)
        shear_ws.write(row, 1, '단계', formats['calc_table_header'])
        shear_ws.write(row, 2, '항목', formats['calc_table_header'])
        shear_ws.merge_range(row, 3, row, 8, '수식 및 계산 과정', formats['calc_table_header'])
        shear_ws.merge_range(row, 9, row, 11, '결과', formats['calc_table_header'])
        shear_ws.merge_range(row, 12, row, 15, '설명', formats['calc_table_header'])
        shear_ws.set_row(row, 25)
        row += 1

        # 1단계: 축력 영향 계수
        shear_ws.write(row, 1, '1', formats['step_number'])
        shear_ws.write(row, 2, '축력 영향 계수', formats['calc_content'])
        formula_text = f'P증가 = 1 + (Pu / (14 × Ag))\n= 1 + ({format_number(r["Pu"]*1000, 0)} / (14 × {format_number(Ag, 0)}))'
        shear_ws.merge_range(row, 3, row, 8, formula_text, formats['formula_wide'])
        shear_ws.merge_range(row, 9, row, 11, f'{r["p_factor"]:.3f}', formats['result_wide'])
        shear_ws.merge_range(row, 12, row, 15, '축력(Pu)이 단면(Ag)에 미치는 영향을 보정하는 계수', formats['calc_content'])
        shear_ws.set_row(row, 50)
        row += 1

        # 2단계: 콘크리트 설계 전단강도
        shear_ws.write(row, 1, '2', formats['step_number'])
        shear_ws.write(row, 2, '콘크리트 전단강도', formats['calc_content'])
        formula_text = f'φVc = φv × (1/6 × P증가 × λ × √fck × bw × d)\n= {phi_v} × (1/6 × {r["p_factor"]:.3f} × {lamda} × {np.sqrt(fck):.2f} × {bw} × {d})'
        shear_ws.merge_range(row, 3, row, 8, formula_text, formats['formula_wide'])
        shear_ws.merge_range(row, 9, row, 11, f'{format_N_to_kN(r["phi_Vc_N"])} kN', formats['result_wide'])
        shear_ws.merge_range(row, 12, row, 15, '콘크리트 기본 전단강도에 강도감소계수와 축력 보정을 적용', formats['calc_content'])
        shear_ws.set_row(row, 50)
        row += 1

        # 3단계: 전단철근 필요성 판정
        shear_ws.write(row, 1, '3', formats['step_number'])
        shear_ws.write(row, 2, '전단철근 판정', formats['calc_content'])
        if check_type == '프리캐스트 (3단계)':
            if r['shear_category'] == "전단철근 불필요":
                judgement = f'Vu ≤ ½φVc\n{format_number(r["Vu"])} ≤ {format_N_to_kN(r["half_phi_Vc_N"])}'
            elif r['shear_category'] == "최소전단철근":
                judgement = f'½φVc < Vu ≤ φVc\n{format_N_to_kN(r["half_phi_Vc_N"])} < {format_number(r["Vu"])} ≤ {format_N_to_kN(r["phi_Vc_N"])}'
            else:
                judgement = f'Vu > φVc\n{format_number(r["Vu"])} > {format_N_to_kN(r["phi_Vc_N"])}'
        else:
            if r['shear_category'] == "전단철근 불필요":
                judgement = f'Vu ≤ φVc\n{format_number(r["Vu"])} ≤ {format_N_to_kN(r["phi_Vc_N"])}'
            else:
                judgement = f'Vu > φVc\n{format_number(r["Vu"])} > {format_N_to_kN(r["phi_Vc_N"])}'
        shear_ws.merge_range(row, 3, row, 8, judgement, formats['calc_content'])
        shear_ws.merge_range(row, 9, row, 11, r['shear_category'], r['category_fmt'])
        shear_ws.merge_range(row, 12, row, 15, f'{check_type} 기준으로 요구 전단력과 콘크리트 전단강도 비교', formats['calc_content'])
        shear_ws.set_row(row, 50)
        row += 1

        # 전단철근 필요 시 추가 계산
        if r['shear_category'] != "전단철근 불필요":
            step_count = 4
            
            # 전단철근 단면적
            shear_ws.write(row, 1, str(step_count), formats['step_number'])
            shear_ws.write(row, 2, '전단철근 단면적', formats['calc_content'])
            formula_text = f'Av = {legs} × (π × ({bar_dia}/2)²)\n= {legs} × (π × {bar_dia/2:.1f}²)'
            shear_ws.merge_range(row, 3, row, 8, formula_text, formats['formula_wide'])
            shear_ws.merge_range(row, 9, row, 11, f'{Av_stirrup:.1f} mm²', formats['result_wide'])
            shear_ws.merge_range(row, 12, row, 15, '스터럽 철근 직경과 다리 수로 단면적 계산', formats['calc_content'])
            shear_ws.set_row(row, 50)
            row += 1
            step_count += 1

            if r['shear_category'] == "설계전단철근":
                # 필요 전단강도
                shear_ws.write(row, 1, str(step_count), formats['step_number'])
                shear_ws.write(row, 2, '필요 전단강도', formats['calc_content'])
                formula_text = f'Vs,req = (Vu - φVc) / φv\n= ({format_number(r["Vu"])} - {format_N_to_kN(r["phi_Vc_N"])}) / {phi_v}'
                shear_ws.merge_range(row, 3, row, 8, formula_text, formats['formula_wide'])
                shear_ws.merge_range(row, 9, row, 11, f'{format_N_to_kN(r["Vs_req_N"])} kN', formats['result_wide'])
                shear_ws.merge_range(row, 12, row, 15, '요구 전단력에서 콘크리트 전단강도를 뺀 철근이 부담할 강도', formats['calc_content'])
                shear_ws.set_row(row, 50)
                row += 1
                step_count += 1

                # 강도 요구 간격
                shear_ws.write(row, 1, str(step_count), formats['step_number'])
                shear_ws.write(row, 2, '강도 요구 간격', formats['calc_content'])
                formula_text = f's강도 = (Av × fyt × d) / Vs,req\n= ({Av_stirrup} × {fy_shear} × {d}) / {format_number(r["Vs_req_N"]*1000)}'
                shear_ws.merge_range(row, 3, row, 8, formula_text, formats['formula_wide'])
                shear_ws.merge_range(row, 9, row, 11, f'{format_number(r["s_from_vs_req"])} mm', formats['result_wide'])
                shear_ws.merge_range(row, 12, row, 15, '철근 강도로부터 도출된 최대 허용 간격', formats['calc_content'])
                shear_ws.set_row(row, 50)
                row += 1
                step_count += 1

            # 최소 철근량 간격
            shear_ws.write(row, 1, str(step_count), formats['step_number'])
            shear_ws.write(row, 2, '최소 철근량 간격', formats['calc_content'])
            formula_text = f's최소 = Av / max(0.0625×√fck×bw/fyt, 0.35×bw/fyt)\n= {Av_stirrup} / {r["min_Av_s_req"]:.4f}'
            shear_ws.merge_range(row, 3, row, 8, formula_text, formats['formula_wide'])
            shear_ws.merge_range(row, 9, row, 11, f'{format_number(r["s_from_min_req"])} mm', formats['result_wide'])
            shear_ws.merge_range(row, 12, row, 15, 'KDS 규정 최소 전단철근량 기준을 만족하는 간격', formats['calc_content'])
            shear_ws.set_row(row, 50)
            row += 1
            step_count += 1

            # 최대 허용 간격
            shear_ws.write(row, 1, str(step_count), formats['step_number'])
            shear_ws.write(row, 2, '최대 허용 간격', formats['calc_content'])
            vs_condition = "Vs > (1/3)×√fck×bw×d" if r["Vs_req_N"] > r["Vs_limit_d4_N"] else "Vs ≤ (1/3)×√fck×bw×d"
            max_condition = "min(d/4, 300)" if r["Vs_req_N"] > r["Vs_limit_d4_N"] else "min(d/2, 600)"
            formula_text = f'{vs_condition}\n따라서 s최대 = {max_condition}'
            shear_ws.merge_range(row, 3, row, 8, formula_text, formats['formula_wide'])
            shear_ws.merge_range(row, 9, row, 11, f'{format_number(r["s_max_code"])} mm', formats['result_wide'])
            shear_ws.merge_range(row, 12, row, 15, '전단강도 크기에 따른 구조적 안전성을 위한 간격 제한', formats['calc_content'])
            shear_ws.set_row(row, 50)
            row += 1
            step_count += 1

            # 최종 간격 결정
            shear_ws.write(row, 1, str(step_count), formats['step_number'])
            shear_ws.write(row, 2, '최종 간격 결정', formats['calc_content'])
            if r['shear_category'] == "설계전단철근":
                formula_text = f'min(s강도, s최소, s최대)\n= min({format_number(r["s_from_vs_req"])}, {format_number(r["s_from_min_req"])}, {format_number(r["s_max_code"])})'
            else:
                formula_text = f'min(s최소, s최대)\n= min({format_number(r["s_from_min_req"])}, {format_number(r["s_max_code"])})'
            shear_ws.merge_range(row, 3, row, 8, formula_text, formats['formula_wide'])
            shear_ws.merge_range(row, 9, row, 11, f'{r["actual_s"]:.0f} mm', formats['result_wide'])
            shear_ws.merge_range(row, 12, row, 15, '각 조건의 최솟값을 택하여 시공성을 고려해 5mm 단위로 내림', formats['calc_content'])
            shear_ws.set_row(row, 50)
            row += 1
            step_count += 1
        else:
            step_count = 4

        # 단면 안전성 검토
        shear_ws.write(row, 1, str(step_count), formats['step_number'])
        shear_ws.write(row, 2, '단면 안전성 검토', formats['calc_content'])
        formula_text = f'Vs,배근 ≤ Vs,max = (2/3)×√fck×bw×d\n{format_N_to_kN(r["Vs_provided_N"],1)} ≤ {format_N_to_kN(r["Vs_max_limit_N"],1)}'
        shear_ws.merge_range(row, 3, row, 8, formula_text, formats['formula_wide'])

        # status_fmt = formats['status_ok'] if r['is_safe'] else formats['status_ng']
        # shear_ws.merge_range(row, 13, row, 14, r['final_status'], status_fmt)

        if r["is_safe_section"]:
            result_format = formats['status_ok']
            safety_result = "✅ 안전"
        else:
            result_format = formats['status_ng']
            safety_result = "❌ 위험"
        shear_ws.merge_range(row, 9, row, 11, safety_result, result_format)

        shear_ws.merge_range(row, 12, row, 15, '전단철근이 부담하는 강도가 최대 허용치를 초과하지 않는지 확인', formats['calc_content'])
        shear_ws.set_row(row, 50)
        row += 1
        step_count += 1

        # 최종 안전성 검토
        shear_ws.write(row, 1, str(step_count), formats['step_number'])
        shear_ws.write(row, 2, '최종 안전성 검토', formats['calc_content'])
        formula_text = f'φVn ≥ Vu\n{format_number(r["phi_Vn_kN"])} ≥ {format_number(r["Vu"])}'
        shear_ws.merge_range(row, 3, row, 8, formula_text, formats['formula_wide'])
        shear_ws.merge_range(row, 9, row, 11, r["final_status"], status_fmt)
        shear_ws.merge_range(row, 12, row, 15, '총 설계 전단강도가 요구 전단강도를 충족하는지 최종 확인', formats['calc_content'])
        shear_ws.set_row(row, 50)
        row += 1

        # ═══ 최종 결과 요약 ═══
        final_result_text = f'배근: {r["stirrups_needed"]} (1m당 {r["stirrups_per_meter"]:.1f}개)'
        if r['ng_reason']:
            final_result_text += f'\n판정 사유: {r["ng_reason"]}'
        final_result_fmt = formats['final_success'] if r['is_safe'] else formats['final_fail']
        shear_ws.merge_range(row, 1, row + 1, max_col, final_result_text, final_result_fmt)
        shear_ws.set_row(row, 30)
        shear_ws.set_row(row + 1, 30)
        row += 2

        if i < len(results) - 1:
            row += 1

    # ═══════════════════════════════════════════════════════════════
    # 6. 설계 기준 및 참고사항
    # ═══════════════════════════════════════════════════════════════
    row += 2
    shear_ws.merge_range(row, 0, row, max_col, '📋 설계 기준 및 참고사항', formats['section_header'])
    shear_ws.set_row(row, 35)
    row += 1

    reference_text = f'''📖 적용 기준: KDS 14 20 콘크리트구조설계기준

🎯 판정기준: {check_type}

🔧 조건: fyt = {fy_shear} MPa, Av = {Av_stirrup:.1f} mm², φv = {phi_v}, λ = {lamda}

🔍 판정: {check_type} 적용 기준'''

    if check_type == '프리캐스트 (3단계)':
        reference_text += '''

 • Vu ≤ ½φVc: 불필요

 • ½φVc < Vu ≤ φVc: 최소철근

 • Vu > φVc: 설계철근'''
    else:
        reference_text += '''

 • Vu ≤ φVc: 불필요

 • Vu > φVc: 설계철근'''

    reference_text += f'''

📊 최소 (Av/s): max(0.0625×√fck×bw/fyt, 0.35×bw/fyt)

⚡ s최대: Vs > (1/3)×√fck×bw×d → min(d/4, 300mm), 아니면 min(d/2, 600mm)

🛡️ Vs,max: (2/3)×√fck×bw×d

💡 P증가: 1 + Pu/(14×Ag)

🎯 간격: 5mm 단위 내림'''

    shear_ws.merge_range(row, 0, row + 15, max_col, reference_text.strip(), formats['calc_content'])
    for i in range(16):
        shear_ws.set_row(row + i, 20)
    row += 16

    # ═══════════════════════════════════════════════════════════════
    # 7. 설계 요약 및 권장사항
    # ═══════════════════════════════════════════════════════════════
    row += 2
    shear_ws.merge_range(row, 0, row, max_col, '💡 설계 요약 및 권장사항', formats['section_header'])
    shear_ws.set_row(row, 35)
    row += 1

    all_safe = all(r['is_safe'] for r in results)
    critical_cases = [r for r in results if not r['is_safe']]

    summary_text = f'''🔍 결과: {len(results)}개 중 {len([r for r in results if r['is_safe']])}개 안전

📊 배근:'''
    for r in results:
        summary_text += f'\n • Case {r["case"]}: {r["stirrups_needed"]} {"✅" if r["is_safe"] else "❌"}'

    if critical_cases:
        summary_text += f'''\n⚠️ 문제 케이스:'''
        for r in critical_cases:
            summary_text += f'\n • Case {r["case"]}: {r["ng_reason"]}'
        summary_text += f'''\n💡 개선: 단면 확대, 간격 조정, 고강도 콘크리트'''
    else:
        summary_text += f'''\n✅ 모든 케이스 안전

💡 권장: 배근도 준수, 정착장 확보, 품질 관리'''

    final_summary_fmt = formats['final_success'] if all_safe else formats['final_fail']
    shear_ws.merge_range(row, 0, row + 9, max_col, summary_text.strip(), final_summary_fmt)
    for i in range(10):
        shear_ws.set_row(row + i, 22)

    # ═══════════════════════════════════════════════════════════════
    # 8. 시트 설정
    # ═══════════════════════════════════════════════════════════════
    shear_ws.set_margins(left=0.5, right=0.5, top=0.75, bottom=0.75)
    shear_ws.set_header('&C&"Noto Sans KR,Bold"&18🛡️ 전단설계 최적화 보고서')
    shear_ws.set_footer(f'&L&D &T&C{check_type} 적용&R&"Noto Sans KR"&12KDS 14 20 기준')

    shear_ws.set_landscape()
    shear_ws.set_paper(9)
    shear_ws.fit_to_pages(1, 0)
    shear_ws.set_default_row(18)

    return shear_ws