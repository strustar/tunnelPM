import numpy as np

def create_shear_sheet(wb, In, R):
    """
    Excel 전단설계 보고서 - Streamlit 웹 버전과 동일 (Mm 개념 추가)
    """
    check_type = In.check_type
    shear_ws = wb.add_worksheet('전단 검토')

    # ─── 색상 팔레트 ─────────────────────────────
    colors = {
        'navy_deep': '#1e3a8a', 'navy_medium': '#3730a3',
        'section_dark': '#1d4ed8', 'section_medium': '#2563eb',
        'criteria_blue': '#3b82f6', 'criteria_purple': '#7c3aed', 'criteria_rose': '#e11d48',
        'table_header': '#059669', 'table_data': '#f8fafc', 'table_alt': '#f1f5f9',
        'success': '#059669', 'success_bg': '#d1fae5',
        'danger': '#dc2626', 'danger_bg': '#fee2e2',
        'case_title': '#cbd5e1', 'case_result': '#f1f5f9',
        'step_header': '#dbeafe', 'calc_block': '#ffffff',
        'formula': '#f3f4f6', 'warning': '#fef3c7', 'sub_header': '#e2e8f0',
        'text_dark': '#0f172a', 'text_medium': '#374151', 'text_light': '#6b7280',
    }

    base_font = {'font_name': 'Noto Sans KR', 'border': 1, 'valign': 'vcenter', 'bold': True}

    styles = {
        'main_title': {**base_font, 'font_size': 28, 'bg_color': colors['navy_deep'], 
                      'font_color': 'white', 'border': 2, 'align': 'center'},
        'sub_title': {**base_font, 'font_size': 16, 'bg_color': colors['navy_medium'], 
                     'font_color': 'white', 'align': 'center'},
        'section_header': {**base_font, 'font_size': 20, 'bg_color': colors['section_dark'], 
                          'font_color': 'white', 'border': 2, 'align': 'center'},
        'criteria_no_shear': {**base_font, 'font_size': 14, 'bg_color': colors['criteria_blue'], 
                             'font_color': 'white', 'align': 'center', 'text_wrap': True, 'border': 1},
        'criteria_min_shear': {**base_font, 'font_size': 14, 'bg_color': colors['criteria_purple'], 
                              'font_color': 'white', 'align': 'center', 'text_wrap': True, 'border': 1},
        'criteria_design_shear': {**base_font, 'font_size': 14, 'bg_color': colors['criteria_rose'], 
                                 'font_color': 'white', 'align': 'center', 'text_wrap': True, 'border': 1},
        'summary_header': {**base_font, 'font_size': 15, 'bg_color': colors['table_header'], 
                          'font_color': 'white', 'border': 2, 'align': 'center'},
        'summary_data': {**base_font, 'font_size': 13, 'bg_color': colors['table_data'], 
                        'font_color': colors['text_dark'], 'align': 'center', 'border': 1},
        'summary_data_alt': {**base_font, 'font_size': 13, 'bg_color': colors['table_alt'], 
                            'font_color': colors['text_dark'], 'align': 'center', 'border': 1},
        'category_no_shear': {**base_font, 'font_size': 13, 'bg_color': '#dbeafe', 
                             'font_color': colors['criteria_blue'], 'align': 'center', 'border': 1},
        'category_min_shear': {**base_font, 'font_size': 13, 'bg_color': '#e0e7ff', 
                              'font_color': colors['criteria_purple'], 'align': 'center', 'border': 1},
        'category_design_shear': {**base_font, 'font_size': 13, 'bg_color': '#fce7f3', 
                                 'font_color': colors['criteria_rose'], 'align': 'center', 'border': 1},
        'status_ok': {**base_font, 'font_size': 13, 'bg_color': colors['success_bg'], 
                     'font_color': colors['success'], 'border': 2, 'align': 'center'},
        'status_ng': {**base_font, 'font_size': 13, 'bg_color': colors['danger_bg'], 
                     'font_color': colors['danger'], 'border': 2, 'align': 'center'},
        'case_title_main': {**base_font, 'font_size': 16, 'bg_color': '#3b82f6', 
                           'font_color': 'white', 'border': 2, 'align': 'center'},
        'case_title': {**base_font, 'font_size': 14, 'bg_color': colors['case_title'], 
                      'font_color': colors['text_dark'], 'border': 1, 'align': 'left'},
        'case_result': {**base_font, 'font_size': 13, 'bg_color': colors['case_result'], 
                       'font_color': colors['text_dark'], 'border': 1, 'align': 'left'},
        'calc_table_header': {**base_font, 'font_size': 13, 'bg_color': colors['table_header'], 
                             'font_color': 'white', 'border': 2, 'align': 'center'},
        'step_header_bright': {**base_font, 'font_size': 14, 'bg_color': '#dbeafe', 
                              'font_color': colors['section_dark'], 'border': 1, 'align': 'left'},
        'sub_header_bold': {**base_font, 'font_size': 12, 'bg_color': '#f1f5f9', 
                           'font_color': colors['text_dark'], 'border': 1, 'align': 'center'},
        'calc_content': {**base_font, 'font_size': 12, 'bg_color': colors['calc_block'], 
                        'font_color': colors['text_dark'], 'align': 'left', 'text_wrap': True, 
                        'border': 1, 'bold': True},
        'formula_wide': {**base_font, 'font_size': 12, 'bg_color': colors['formula'], 'align': 'left', 
                        'font_color': colors['text_dark'], 'border': 1, 'bold': True, 
                        'text_wrap': True},
        'result_wide': {**base_font, 'font_size': 12, 'bg_color': colors['table_data'], 
                       'font_color': colors['text_dark'], 'border': 1, 'align': 'center', 
                       'text_wrap': True, 'bold': True},
        'final_success': {**base_font, 'font_size': 15, 'bg_color': colors['success_bg'], 
                         'font_color': colors['success'], 'border': 2, 'align': 'center', 'text_wrap': True},
        'final_fail': {**base_font, 'font_size': 15, 'bg_color': colors['danger_bg'], 
                      'font_color': colors['danger'], 'border': 2, 'align': 'center', 'text_wrap': True},
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

    # ─── 컬럼 너비 설정 ─────────────────────────────
    shear_ws.set_column('A:A', 5)
    shear_ws.set_column('B:B', 12)
    shear_ws.set_column('C:P', 16)
    shear_ws.set_column('I:J', 20)
    shear_ws.set_column('G:G', 26)
    
    row = 0
    max_col = 15

    # ─── 헬퍼 함수 ─────────────────────────────
    def format_number(num, decimal_places=1):
        return f"{num:,.{decimal_places}f}"

    def format_N_to_kN(value, dp=1):
        return f"{value/1000:,.{dp}f}"

    # ─── 설계 상수 ─────────────────────────────
    phi_v = 0.75
    lamda = 1.0
    bar_dia = 13
    legs = 2
    bar_area = np.pi * (bar_dia / 2)**2
    Av_stirrup = bar_area * legs

    bw, d, h, fck, Ag, fy_shear = In.be, In.depth, In.height, In.fck, R.Ag, In.fy_hollow
    Ast_tension = R.Ast_tension[0]/2
    Ast_compression = R.Ast_compression[0]/2
    As = Ast_tension + Ast_compression
    rho_w = As/(bw*d)

    results = []

    # ═══════════════════════════════════════════════════════════════
    # 각 하중 케이스별 계산 (Mm 개념 추가)
    # ═══════════════════════════════════════════════════════════════
    for i in range(len(In.Vu)):
        Vu = In.Vu[i]
        Pu = In.Pu[i]
        Mu = In.Mu[i]
        
        Nu = Pu * 1000
        Mu_Nmm = Mu * 1e6

        # Mm (수정 모멘트) 계산
        Mm = Mu_Nmm - Nu * (4 * h - d) / 8
        Mm_kNm = Mm / 1e6

        # Vc 계산 (Mm 값에 따라 식 선택)
        if Mm < 0:
            # 축력 고려식
            Vc = 0.29 * lamda * np.sqrt(fck) * bw * d * np.sqrt(1 + Nu / (3.5 * Ag))
            vc_method = "축력 고려식 (Mm < 0)"
        else:
            # 정밀식
            Vc = (1/6 * np.sqrt(fck) + 17.6 * rho_w * Vu*1000 * d / Mm) * bw * d
            vc_method = "정밀식 (Mm ≥ 0)"

        phi_Vc = phi_v * Vc
        half_phi_Vc = 0.5 * phi_Vc

        # 판정
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

        # 최소 전단철근량
        min_Av_s_1_val = 0.0625 * np.sqrt(fck)
        min_Av_s_2_val = 0.35
        min_Av_s_1 = min_Av_s_1_val * (bw / fy_shear)
        min_Av_s_2 = min_Av_s_2_val * (bw / fy_shear)
        min_Av_s_req = max(min_Av_s_1, min_Av_s_2)
        s_from_min_req = Av_stirrup / min_Av_s_req

        # 필요 전단철근량
        Vs_req = (Vu * 1000 - phi_Vc) / phi_v if shear_category == "설계전단철근" else 0
        s_from_vs_req = (Av_stirrup * fy_shear * d) / Vs_req if Vs_req > 0 else float('inf')
        
        # 최대 간격
        Vs_limit_d4 = (1/3) * np.sqrt(fck) * bw * d
        s_max_code = min(d / 4, 300) if Vs_req > Vs_limit_d4 else min(d / 2, 600)
        
        # 최종 간격
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

        # 최종 강도
        if shear_category == "전단철근 불필요":
            phi_Vs = 0
        else:
            phi_Vs = (phi_v * Av_stirrup * fy_shear * d) / actual_s if actual_s > 0 else 0
        
        phi_Vn = phi_Vc + phi_Vs
        is_safe_strength = (phi_Vn >= Vu * 1000)
        
        # 단면 안전성
        Vs_max_limit = (2/3) * np.sqrt(fck) * bw * d
        Vs_provided = phi_Vs / phi_v if phi_Vs > 0 else 0
        is_safe_section = (Vs_provided <= Vs_max_limit)
        is_safe_total = is_safe_strength and is_safe_section
        
        stirrups_per_meter = 1000 / actual_s if actual_s > 0 and shear_category != "전단철근 불필요" else 0
        
        # 최종 판정
        if not is_safe_section:
            final_status = "❌ NG (단면 부족)"
            ng_reason = f"Vs = {format_N_to_kN(Vs_provided)} kN > Vs,max = {format_N_to_kN(Vs_max_limit)} kN"
        elif not is_safe_strength:
            final_status = "❌ NG (강도 부족)"
            ng_reason = f"φVn = {format_N_to_kN(phi_Vn)} kN < Vu = {format_number(Vu)} kN"
        else:
            final_status = "✅ OK"
            ng_reason = ""

        results.append({
            'case': i + 1, 'Pu': Pu, 'Vu': Vu, 'Mu': Mu,
            'Mm_kNm': Mm_kNm, 'vc_method': vc_method,
            'shear_category': shear_category, 'category_fmt': category_fmt,
            'phi_Vn_kN': phi_Vn / 1000, 'is_safe': is_safe_total,
            'is_safe_section': is_safe_section, 'is_safe_strength': is_safe_strength,
            'actual_s': actual_s, 'stirrups_needed': stirrups_needed, 
            'stirrups_per_meter': stirrups_per_meter,
            'Vc_N': Vc, 'phi_Vc_N': phi_Vc, 'half_phi_Vc_N': half_phi_Vc,
            'Vs_req_N': Vs_req, 'min_Av_s_req': min_Av_s_req,
            's_from_min_req': s_from_min_req, 's_from_vs_req': s_from_vs_req,
            's_max_code': s_max_code, 'Vs_limit_d4_N': Vs_limit_d4,
            'phi_Vs_N': phi_Vs, 'Vs_provided_N': Vs_provided,
            'Vs_max_limit_N': Vs_max_limit, 'final_status': final_status,
            'ng_reason': ng_reason, 'min_Av_s_1_val': min_Av_s_1_val,
            'min_Av_s_2_val': min_Av_s_2_val
        })

    # ═══════════════════════════════════════════════════════════════
    # 1. 메인 타이틀
    # ═══════════════════════════════════════════════════════════════
    shear_ws.merge_range(row, 0, row, max_col, '🛡️ 전단설계 최적화 보고서', formats['main_title'])
    shear_ws.set_row(row, 50)
    row += 1

    shear_ws.merge_range(row, 0, row, max_col, 'KDS 14 20 콘크리트구조설계기준 적용 (축력 고려)', formats['sub_title'])
    shear_ws.set_row(row, 28)
    row += 2

    # ═══════════════════════════════════════════════════════════════
    # 2. 판정 기준 선택
    # ═══════════════════════════════════════════════════════════════
    shear_ws.merge_range(row, 0, row, max_col, '📋 전단철근 판정 기준 선택', formats['criteria_selection_header'])
    shear_ws.set_row(row, 32)
    row += 1

    option_text = "● 일반 (2단계)                    ○ 프리캐스트 (3단계)" if check_type == '일반 (2단계)' else "○ 일반 (2단계)                    ● 프리캐스트 (3단계)"
    shear_ws.merge_range(row, 0, row, max_col, option_text, formats['criteria_selected'])
    shear_ws.set_row(row, 22)
    row += 2

    # ═══════════════════════════════════════════════════════════════
    # 3. 판정 기준 박스
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
    # 4. 결과 요약
    # ═══════════════════════════════════════════════════════════════
    shear_ws.merge_range(row, 0, row, max_col, '📊 전단설계 결과 요약', formats['section_header'])
    shear_ws.set_row(row, 35)
    row += 2

    headers = ['Case', 'Vu (kN)', 'Pu (kN)', 'Mu (kN·m)', 'Mm (kN·m)', 'Vc 계산법', 'φVc (kN)', '판정', '배근', 'φVn (kN)', '최종']
    col_spans = [(1,1), (2,2), (3,3), (4,4), (5,5), (6,6), (7,7), (8,8), (9,9), (10,10), (11,11)]

    for header, (col_start, col_end) in zip(headers, col_spans):
        if col_start == col_end:
            shear_ws.write(row, col_start, header, formats['summary_header'])
        else:
            shear_ws.merge_range(row, col_start, row, col_end, header, formats['summary_header'])
    shear_ws.set_row(row, 28)
    row += 1

    for i, r in enumerate(results):
        data_fmt = formats['summary_data'] if i % 2 == 0 else formats['summary_data_alt']
        
        shear_ws.write(row, 1, f"Case {r['case']}", data_fmt)
        shear_ws.write(row, 2, f"{r['Vu']:.1f}", data_fmt)
        shear_ws.write(row, 3, f"{r['Pu']:.1f}", data_fmt)
        shear_ws.write(row, 4, f"{r['Mu']:.1f}", data_fmt)
        shear_ws.write(row, 5, f"{r['Mm_kNm']:.1f}", data_fmt)
        shear_ws.write(row, 6, r['vc_method'], data_fmt)
        shear_ws.write(row, 7, format_N_to_kN(r['phi_Vc_N']), data_fmt)
        shear_ws.write(row, 8, r['shear_category'], r['category_fmt'])
        shear_ws.write(row, 9, r['stirrups_needed'], data_fmt)
        shear_ws.write(row, 10, f"{r['phi_Vn_kN']:.1f}", data_fmt)
        
        status_fmt = formats['status_ok'] if r['is_safe'] else formats['status_ng']
        shear_ws.write(row, 11, r['final_status'], status_fmt)
        
        shear_ws.set_row(row, 30)
        row += 1

    row += 2

    # ═══════════════════════════════════════════════════════════════
    # 5. 상세 계산 과정
    # ═══════════════════════════════════════════════════════════════
    shear_ws.merge_range(row, 0, row, max_col, '📝 상세 계산 과정', formats['section_header'])
    shear_ws.set_row(row, 35)
    row += 2

    case_symbols = ["❶", "❷", "❸", "❹", "❺", "❻", "❼", "❽", "❾", "❿"]

    for idx, r in enumerate(results):
        # 케이스 구분선
        if idx > 0:
            shear_ws.merge_range(row, 0, row, max_col, "━" * 50, formats['case_separator_line'])
            shear_ws.set_row(row, 10)
            row += 1

        # 케이스 제목
        shear_ws.merge_range(row, 1, row, max_col, 
                            f'{case_symbols[idx]} Case {r["case"]} 상세 계산', 
                            formats['case_title_main'])
        shear_ws.set_row(row, 30)
        row += 1

        # 1단계: 설계 조건
        shear_ws.merge_range(row, 1, row, max_col, '1단계: 설계 조건 확인', formats['step_header_bright'])
        shear_ws.set_row(row, 25)
        row += 1

        condition_text = f'''■ 하중 조건
  Vu = {r['Vu']:.1f} kN, Pu = {r['Pu']:.1f} kN, Mu = {r['Mu']:.1f} kN·m

■ 부재 제원
  bw = {bw:,.0f} mm, d = {d:,.0f} mm, h = {h:,.0f} mm   (dc = {h-d:,.1f} mm)

■ 재료 특성
  fck = {fck:.0f} MPa, fys = {fy_shear:.0f} MPa, λ = {lamda:.1f}

■ 배근 정보
  인장측: Ast = {Ast_tension:,.1f} mm²
  압축측: Asc = {Ast_compression:,.1f} mm²
  전단철근: H{bar_dia}-{legs}leg (Av = {Av_stirrup:.1f} mm²)'''

        shear_ws.merge_range(row, 1, row + 11, max_col, condition_text, formats['calc_content'])
        for i in range(12):
            shear_ws.set_row(row + i, 22)
        row += 12

        # 2단계: Mm 계산
        shear_ws.merge_range(row, 1, row, max_col, '2단계: Mm (수정 모멘트) 계산', formats['step_header_bright'])
        shear_ws.set_row(row, 25)
        row += 1

        mm_text = f'''■ 일반식
  Mm = Mu - Pu × (4h - d) / 8

■ 값 대입 및 계산
  Mm = {r['Mu']:.1f} - {r['Pu']:.1f} × (4×{h}-{d}) / 8,000
     = {r['Mm_kNm']:.1f} kN·m'''

        shear_ws.merge_range(row, 1, row + 5, max_col, mm_text, formats['calc_content'])
        for i in range(6):
            shear_ws.set_row(row + i, 22)
        row += 6

        # 3단계: φVc 계산
        shear_ws.merge_range(row, 1, row, max_col, '3단계: 콘크리트 부담 전단강도 (φVc) 계산', formats['step_header_bright'])
        shear_ws.set_row(row, 25)
        row += 1

        if r['Mm_kNm'] < 0:
            vc_text = f'''■ φVc 산정식 선택
  Mm = {r['Mm_kNm']:.1f} kN·m < 0 → 축력 고려식 적용

■ 축력 고려식
  φVc = φ × 0.29 λ √fck × bw × d × √(1 + Nu/(3.5Ag))
  φVc = 0.75 × 0.29 × 1.0 × √{fck} × {bw} × {d} × √(1 + {r['Pu']*1000:,.0f}/(3.5×{Ag:,.0f}))
  φVc = {format_N_to_kN(r['phi_Vc_N'])} kN'''
        else:
            vc_text = f'''■ φVc 산정식 선택
  Mm = {r['Mm_kNm']:.1f} kN·m ≥ 0 → 정밀식 적용

■ 정밀식 (전단력과 휨 모멘트 고려)
  ρw = As / (bw × d) = {As:,.0f} / ({bw} × {d}) = {rho_w:.4f}
  
  φVc = φ × [(1/6)√fck + 17.6 ρw (Vu·d/Mu)] × bw × d
  φVc = 0.75 × [(1/6)√{fck} + 17.6×{rho_w:.4f}×({r['Vu']:.1f}×{d}/{r['Mm_kNm']:.1f}×1000)] × {bw} × {d}
  φVc = {format_N_to_kN(r['phi_Vc_N'])} kN'''

        shear_ws.merge_range(row, 1, row + 7, max_col, vc_text, formats['calc_content'])
        for i in range(8):
            shear_ws.set_row(row + i, 22)
        row += 8

        # 4단계: 전단철근 판정
        shear_ws.merge_range(row, 1, row, max_col, '4단계: 전단철근 판정', formats['step_header_bright'])
        shear_ws.set_row(row, 25)
        row += 1

        if check_type == '프리캐스트 (3단계)':
            if r['shear_category'] == "전단철근 불필요":
                judgement = f"Vu ≤ ½φVc : {r['Vu']:.1f} ≤ {format_N_to_kN(r['half_phi_Vc_N'])}"
            elif r['shear_category'] == "최소전단철근":
                judgement = f"½φVc < Vu ≤ φVc : {format_N_to_kN(r['half_phi_Vc_N'])} < {r['Vu']:.1f} ≤ {format_N_to_kN(r['phi_Vc_N'])}"
            else:
                judgement = f"Vu > φVc : {r['Vu']:.1f} > {format_N_to_kN(r['phi_Vc_N'])}"
        else:
            if r['shear_category'] == "전단철근 불필요":
                judgement = f"Vu ≤ φVc : {r['Vu']:.1f} ≤ {format_N_to_kN(r['phi_Vc_N'])}"
            else:
                judgement = f"Vu > φVc : {r['Vu']:.1f} > {format_N_to_kN(r['phi_Vc_N'])}"

        judgement_text = f'''■ 판정 조건
  {judgement}

■ 판정 결과
  {r['shear_category']}'''

        shear_ws.merge_range(row, 1, row + 4, max_col, judgement_text, formats['calc_content'])
        for i in range(5):
            shear_ws.set_row(row + i, 22)
        row += 5

        # 전단철근 필요 시 추가 계산
        if r['shear_category'] != "전단철근 불필요":
            # 5단계: 전단철근 설계 (세분화)
            shear_ws.merge_range(row, 1, row, max_col, '5단계: 전단철근 설계', formats['step_header_bright'])
            shear_ws.set_row(row, 25)
            row += 1

            # 5-1. 전단철근 단면적
            av_text = f'''■ 전단철근 단면적
  Av = {legs} × (π × {bar_dia}²/4) = {Av_stirrup:.1f} mm²'''
            shear_ws.merge_range(row, 1, row + 1, max_col, av_text, formats['calc_content'])
            for i in range(2):
                shear_ws.set_row(row + i, 22)
            row += 2

            if r['shear_category'] == "설계전단철근":
                # 5-2. 필요 전단강도
                vs_text = f'''■ 필요 전단강도
  Vs,req = (Vu - φVc) / φv
  Vs,req = ({r['Vu']:.1f} - {format_N_to_kN(r['phi_Vc_N'])}) / {phi_v}
  Vs,req = {format_N_to_kN(r['Vs_req_N'])} kN'''
                shear_ws.merge_range(row, 1, row + 3, max_col, vs_text, formats['calc_content'])
                for i in range(4):
                    shear_ws.set_row(row + i, 22)
                row += 4

                # 5-3. 강도 요구 간격
                s_strength_text = f'''■ 강도 요구 간격
  s강도 = (Av × fyt × d) / Vs,req
  s강도 = ({Av_stirrup:.1f} × {fy_shear} × {d}) / {r['Vs_req_N']*1000:.0f}
  s강도 = {r['s_from_vs_req']:.1f} mm'''
                shear_ws.merge_range(row, 1, row + 3, max_col, s_strength_text, formats['calc_content'])
                for i in range(4):
                    shear_ws.set_row(row + i, 22)
                row += 4

            # 5-4. 최소 철근량 간격
            s_min_text = f'''■ 최소 철근량 간격
  (Av/s)min = max(0.0625√fck×bw/fyt, 0.35×bw/fyt)
  (Av/s)min = max({r['min_Av_s_1_val']:.4f}×{bw}/{fy_shear}, {r['min_Av_s_2_val']:.4f}×{bw}/{fy_shear})
  (Av/s)min = {r['min_Av_s_req']:.4f}
  s최소 = {Av_stirrup:.1f} / {r['min_Av_s_req']:.4f} = {r['s_from_min_req']:.1f} mm'''
            shear_ws.merge_range(row, 1, row + 4, max_col, s_min_text, formats['calc_content'])
            for i in range(5):
                shear_ws.set_row(row + i, 22)
            row += 5

            # 5-5. 최대 허용 간격
            vs_condition = "Vs > (1/3)√fck×bw×d" if r['Vs_req_N'] > r['Vs_limit_d4_N'] else "Vs ≤ (1/3)√fck×bw×d"
            max_condition = "min(d/4, 300)" if r['Vs_req_N'] > r['Vs_limit_d4_N'] else "min(d/2, 600)"
            s_max_text = f'''■ 최대 허용 간격
  {vs_condition}
  s최대 = {max_condition} = {r['s_max_code']:.1f} mm'''
            shear_ws.merge_range(row, 1, row + 2, max_col, s_max_text, formats['calc_content'])
            for i in range(3):
                shear_ws.set_row(row + i, 22)
            row += 3

            # 5-6. 최종 간격 결정
            if r['shear_category'] == "설계전단철근":
                final_s_text = f'''■ 최종 간격 결정
  min(s강도, s최소, s최대) = min({r['s_from_vs_req']:.1f}, {r['s_from_min_req']:.1f}, {r['s_max_code']:.1f})
  s최종 = {r['actual_s']:.0f} mm (5mm 단위 내림)'''
            else:
                final_s_text = f'''■ 최종 간격 결정
  min(s최소, s최대) = min({r['s_from_min_req']:.1f}, {r['s_max_code']:.1f})
  s최종 = {r['actual_s']:.0f} mm (5mm 단위 내림)'''
            shear_ws.merge_range(row, 1, row + 2, max_col, final_s_text, formats['calc_content'])
            for i in range(3):
                shear_ws.set_row(row + i, 22)
            row += 3

            # 5-7. 단면 안전성 검토 (정밀식: 계산식 수치 대입 포함)
            section_text = f'''■ 단면 안전성 검토
  Vs,배근 ≤ Vs,max = (2/3)√fck × bw × d
  
  Vs,max = (2/3) × √{fck} × {bw} × {d}
  Vs,max = {format_N_to_kN(r['Vs_max_limit_N'])} kN
  
  Vs,배근 = φVs / φv = {format_N_to_kN(r['phi_Vs_N'])} / {phi_v}
  Vs,배근 = {format_N_to_kN(r['Vs_provided_N'])} kN

■ 판정
  {format_N_to_kN(r['Vs_provided_N'])} ≤ {format_N_to_kN(r['Vs_max_limit_N'])} → {"✅ 안전" if r['is_safe_section'] else "❌ 위험"}'''
            shear_ws.merge_range(row, 1, row + 9, max_col, section_text, formats['calc_content'])
            for i in range(10):
                shear_ws.set_row(row + i, 22)
            row += 10

            step_count = 6
        else:
            step_count = 5

        # 최종 안전성 검토 (이전 단계와 동일한 스타일)
        shear_ws.merge_range(row, 1, row, max_col, f'{step_count}단계: 최종 안전성 검토', formats['step_header_bright'])
        shear_ws.set_row(row, 25)
        row += 1

        final_text = f'''■ 설계 전단강도 검토
  φVn ≥ Vu
  {r['phi_Vn_kN']:.1f} kN {"≥" if r['is_safe_strength'] else "<"} {r['Vu']:.1f} kN → {"✅ 만족" if r['is_safe_strength'] else "❌ 불만족"}

■ 최종 배근
  {r['stirrups_needed']}
  (1m당 {r['stirrups_per_meter']:.1f}개)

■ 설계 전단강도
  φVn = φVc + φVs
  φVn = {format_N_to_kN(r['phi_Vc_N'])} + {format_N_to_kN(r['phi_Vs_N'])}
  φVn = {r['phi_Vn_kN']:.1f} kN

■ 최종 판정
  {r['final_status']}'''
        
        if r['ng_reason']:
            final_text += f"\n\n■ 판정 사유\n  {r['ng_reason']}"
            
        shear_ws.merge_range(row, 1, row + 13, max_col, final_text, formats['calc_content'])
        for i in range(14):
            shear_ws.set_row(row + i, 22)
        row += 14

        if idx < len(results) - 1:
            row += 1

    # ═══════════════════════════════════════════════════════════════
    # 6. 참고사항
    # ═══════════════════════════════════════════════════════════════
    row += 2
    shear_ws.merge_range(row, 0, row, max_col, '📋 설계 기준 및 참고사항', formats['section_header'])
    shear_ws.set_row(row, 35)
    row += 1

    reference_text = f'''📖 적용 기준: KDS 14 20 콘크리트구조설계기준

🎯 판정기준: {check_type}

🔧 설계조건
  • fyt = {fy_shear} MPa
  • Av = {Av_stirrup:.1f} mm² (H{bar_dia}-{legs}leg)
  • φv = {phi_v}, λ = {lamda}

📊 Mm (수정 모멘트) 개념
  • Mm = Mu - Pu × (4h - d) / 8
  • Mm < 0: 축력 고려식 → φVc = 0.29λ√fck × bw × d × √(1 + Nu/(3.5Ag))
  • Mm ≥ 0: 정밀식 → φVc = φ[(1/6)√fck + 17.6ρw(Vu·d/Mu)] × bw × d

🔍 최소 (Av/s): max(0.0625√fck×bw/fyt, 0.35×bw/fyt)

⚡ s최대
  • Vs > (1/3)√fck×bw×d → min(d/4, 300mm)
  • Vs ≤ (1/3)√fck×bw×d → min(d/2, 600mm)

🛡️ Vs,max: (2/3)√fck×bw×d

💡 간격: 5mm 단위 내림'''

    shear_ws.merge_range(row, 0, row + 18, max_col, reference_text.strip(), formats['calc_content'])
    for i in range(19):
        shear_ws.set_row(row + i, 20)
    row += 19

    # ═══════════════════════════════════════════════════════════════
    # 7. 설계 요약
    # ═══════════════════════════════════════════════════════════════
    row += 2
    shear_ws.merge_range(row, 0, row, max_col, '💡 설계 요약 및 권장사항', formats['section_header'])
    shear_ws.set_row(row, 35)
    row += 1

    all_safe = all(r['is_safe'] for r in results)
    critical_cases = [r for r in results if not r['is_safe']]

    summary_text = f'''🔍 전체 결과: {len(results)}개 케이스 중 {len([r for r in results if r['is_safe']])}개 안전

📊 배근 결과:'''
    for r in results:
        summary_text += f'\n  • Case {r["case"]}: {r["stirrups_needed"]} {"✅" if r["is_safe"] else "❌"}'

    if critical_cases:
        summary_text += f'''\n\n⚠️ 문제 케이스 ({len(critical_cases)}개):'''
        for r in critical_cases:
            summary_text += f'\n  • Case {r["case"]}: {r["ng_reason"]}'
        summary_text += f'''\n\n💡 개선 방안:
  • 단면 확대 검토
  • 철근 간격 조정
  • 고강도 콘크리트 사용'''
    else:
        summary_text += f'''\n\n✅ 모든 케이스 안전

💡 시공 권장사항:
  • 배근도 준수
  • 정착장 확보
  • 품질 관리 철저'''

    final_summary_fmt = formats['final_success'] if all_safe else formats['final_fail']
    shear_ws.merge_range(row, 0, row + 12, max_col, summary_text.strip(), final_summary_fmt)
    for i in range(13):
        shear_ws.set_row(row + i, 22)

    # ═══════════════════════════════════════════════════════════════
    # 8. 시트 설정
    # ═══════════════════════════════════════════════════════════════
    shear_ws.set_margins(left=0.5, right=0.5, top=0.75, bottom=0.75)
    shear_ws.set_header('&C&"Noto Sans KR,Bold"&18🛡️ 전단설계 최적화 보고서')
    shear_ws.set_footer(f'&L&D &T&C{check_type} 적용&R&"Noto Sans KR"&12KDS 14 20 기준 (축력 고려)')
    shear_ws.set_landscape()
    shear_ws.set_paper(9)
    shear_ws.fit_to_pages(1, 0)
    shear_ws.set_default_row(18)

    return shear_ws