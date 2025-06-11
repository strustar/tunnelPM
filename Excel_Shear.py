import xlsxwriter
import numpy as np

def create_shear_sheet(wb, In, R):
    """
    스트림릿 앱의 전단설계 검토 내용을 그대로 Excel 시트로 생성합니다.
    (설계 기준, 상세 계산 과정, 전단강도 검토 결과 포함) - 최적화 버전
    """
    shear_ws = wb.add_worksheet('전단 검토')
    # shear_ws.activate()

    # --- 1. 스타일 정의 (최적화) ---
    base_font = {'font_name': '맑은 고딕', 'border': 1, 'valign': 'vcenter', 'text_wrap': True}
    styles = {
        # 메인 헤더
        'title_main': {'bold': True, 'font_size': 18, 'bg_color': '#1E3C72', 'font_color': 'white', 'align': 'center'},
        'title_sub': {'bold': True, 'font_size': 12, 'bg_color': '#2E75B6', 'font_color': 'white', 'align': 'center'},
        
        # 판정 기준 섹션
        'criteria_header': {'bold': True, 'font_size': 14, 'bg_color': '#FFA500', 'font_color': 'black', 'align': 'center'},
        'criteria_blue_header': {'bold': True, 'font_size': 11, 'bg_color': '#E3F2FD', 'font_color': '#1565C0', 'align': 'center'},
        'criteria_orange_header': {'bold': True, 'font_size': 11, 'bg_color': '#FFF8E1', 'font_color': '#E65100', 'align': 'center'},
        'criteria_red_header': {'bold': True, 'font_size': 11, 'bg_color': '#FFEBEE', 'font_color': '#C62828', 'align': 'center'},
        'criteria_blue': {'font_size': 10, 'bg_color': '#F8FDFF', 'align': 'left'},
        'criteria_orange': {'font_size': 10, 'bg_color': '#FFFDF7', 'align': 'left'},
        'criteria_red': {'font_size': 10, 'bg_color': '#FFFAF9', 'align': 'left'},
        
        # 요약표
        'summary_header': {'bold': True, 'font_size': 14, 'bg_color': '#FFA500', 'font_color': 'black', 'align': 'center'},
        'table_header': {'bold': True, 'font_size': 10, 'bg_color': '#495057', 'font_color': 'white', 'align': 'center'},
        'table_cell': {'font_size': 9, 'align': 'center', 'num_format': '#,##0.0'},
        'table_text': {'font_size': 9, 'align': 'center'},
        
        # 상세 계산
        'detail_header': {'bold': True, 'font_size': 14, 'bg_color': '#FFA500', 'font_color': 'black', 'align': 'center'},
        'case_title': {'bold': True, 'font_size': 11, 'bg_color': '#343A40', 'font_color': 'white', 'align': 'center'},
        'case_result': {'bold': True, 'font_size': 10, 'bg_color': '#FFF3CD', 'font_color': '#856404', 'align': 'center'},
        'step_header': {'bold': True, 'font_size': 10, 'bg_color': '#2E75B6', 'font_color': 'white', 'align': 'left'},
        'formula_box': {'font_size': 9, 'bg_color': '#F8F9FA', 'align': 'center', 'italic': True},
        'calc_box': {'font_size': 9, 'bg_color': '#FFFFFF', 'align': 'left'},
        'result_box': {'bold': True, 'font_size': 10, 'bg_color': '#E8F5E8', 'align': 'center'},
        
        # 판정 결과
        'ok': {'bold': True, 'font_size': 10, 'bg_color': '#D4EDDA', 'font_color': '#155724', 'align': 'center'},
        'ng': {'bold': True, 'font_size': 10, 'bg_color': '#F8D7DA', 'font_color': '#721C24', 'align': 'center'},
        'warning': {'bold': True, 'font_size': 9, 'bg_color': '#FFF3CD', 'font_color': '#856404', 'align': 'left'},
    }
    
    formats = {name: wb.add_format({**base_font, **prop}) for name, prop in styles.items()}
    subscript_format = wb.add_format({'font_name': '맑은 고딕', 'font_script': 2, 'font_size': 8})

    # --- 2. 컬럼 너비 설정 ---
    col_widths = [2, 12, 8, 12, 8, 12, 8, 2, 12, 8, 12, 8, 12, 8, 2]
    for i, width in enumerate(col_widths):
        shear_ws.set_column(i, i, width)
    
    # --- 3. 헬퍼 함수 ---
    def write_rich(r, c, base, sub, unit="", fmt=formats['calc_box']):
        """아래첨자 텍스트 작성"""
        shear_ws.write_rich_string(r, c, fmt, base, subscript_format, sub, fmt, unit)

    # --- 4. 전단설계 계산 로직 ---
    phi_v = 0.75
    lamda = 1.0
    fy_shear = 400
    bar_dia = 13
    legs = 2
    bar_area = np.pi * (bar_dia / 2)**2
    Av_stirrup = bar_area * legs
    
    bw, d, fck, Ag = In.be, In.depth, In.fck, R.Ag
    
    # 각 케이스별 계산 결과 저장
    results = []
    for i in range(len(In.Vu)):
        Vu = In.Vu[i]
        Pu = In.Pu_shear[i]
        
        # 축력 영향 계수
        p_factor = 1 + (Pu * 1000) / (14 * Ag) if Pu != 0 else 1.0
        
        # 콘크리트 전단강도
        Vc = (1/6) * p_factor * lamda * np.sqrt(fck) * bw * d
        phi_Vc = phi_v * Vc
        half_phi_Vc = 0.5 * phi_Vc
        
        # 전단철근 판정
        if Vu * 1000 <= half_phi_Vc:
            category = "전단철근 불필요"
            category_short = "불필요"
            color_style = "blue"
        elif Vu * 1000 <= phi_Vc:
            category = "최소전단철근"
            category_short = "최소전단철근"
            color_style = "orange"
        else:
            category = "설계전단철근"
            category_short = "설계전단철근"
            color_style = "red"
        
        # 최소 전단철근량
        min_Av_s_1 = 0.0625 * np.sqrt(fck) * (bw / fy_shear)
        min_Av_s_2 = 0.35 * (bw / fy_shear)
        min_Av_s_req = max(min_Av_s_1, min_Av_s_2)
        s_from_min_req = Av_stirrup / min_Av_s_req
        
        # 필요 전단철근강도
        Vs_req = (Vu * 1000 - phi_Vc) / phi_v if category == "설계전단철근" else 0
        Vs_limit_d4 = (1/3) * np.sqrt(fck) * bw * d
        s_max_code = min(d / 4, 300) if Vs_req > Vs_limit_d4 else min(d / 2, 600)
        
        # 최종 간격 결정
        if category == "전단철근 불필요":
            actual_s = s_max_code
            stirrups = "불필요"
        else:
            if category == "설계전단철근":
                s_from_vs_req = (Av_stirrup * fy_shear * d) / Vs_req if Vs_req > 0 else float('inf')
                s_calc = min(s_from_min_req, s_from_vs_req)
            else:
                s_calc = s_from_min_req
            
            actual_s = min(s_calc, s_max_code)
            actual_s = np.floor(actual_s / 5) * 5
            stirrups = f"H{bar_dia}-{legs}leg @{actual_s:.0f}"
        
        # 제공 전단강도
        phi_Vs = (phi_v * Av_stirrup * fy_shear * d) / actual_s if actual_s > 0 and category != "전단철근 불필요" else 0
        phi_Vn = phi_Vc + phi_Vs
        safety_ratio = phi_Vn / (Vu * 1000) if Vu > 0 else float('inf')
        is_safe = phi_Vn >= Vu * 1000
        
        # 단면 검토
        Vs_max_limit = (2/3) * np.sqrt(fck) * bw * d
        Vs_provided = phi_Vs / phi_v if phi_Vs > 0 else 0
        is_safe_section = Vs_provided <= Vs_max_limit
        is_safe_total = is_safe and is_safe_section
        
        results.append({
            'case': i + 1, 'Pu': Pu, 'Vu': Vu, 'category': category, 'category_short': category_short,
            'color_style': color_style, 'p_factor': p_factor, 'Vc': Vc, 'phi_Vc': phi_Vc, 'half_phi_Vc': half_phi_Vc,
            'Vs_req': Vs_req, 'min_Av_s_req': min_Av_s_req, 's_from_min_req': s_from_min_req,
            's_max_code': s_max_code, 'actual_s': actual_s, 'stirrups': stirrups,
            'phi_Vs': phi_Vs, 'phi_Vn': phi_Vn, 'safety_ratio': safety_ratio,
            'is_safe': is_safe_total, 'is_safe_section': is_safe_section
        })

    # --- 5. 시트 내용 작성 ---
    row = 0
    
    # ─── 메인 타이틀 ───
    shear_ws.merge_range(row, 0, row, 14, '🛡️ 전단설계 최적화 보고서', formats['title_main'])
    shear_ws.set_row(row, 30)
    row += 1
    shear_ws.merge_range(row, 0, row, 14, 'KDS 14 20 콘크리트구조설계기준 적용', formats['title_sub'])
    shear_ws.set_row(row, 20)
    row += 2

    # ─── 전단철근 판정 기준 (3단계) ───
    shear_ws.merge_range(row, 0, row, 14, '📋 전단철근 판정 기준 (3단계)', formats['criteria_header'])
    shear_ws.set_row(row, 25)
    row += 1
    
    # 3단계 기준표 (개선된 레이아웃)
    criteria_data = [
        ['🔵 전단철근 불필요', 'Vu ≤ ½φVc', '이론적으로 전단철근 불필요', 'blue'],
        ['🟡 최소전단철근', '½φVc < Vu ≤ φVc', '규정 최소량 적용', 'orange'],
        ['🔴 설계전단철근', 'Vu > φVc', '계산에 의한 철근량', 'red']
    ]
    
    for i, (title, condition, description, color) in enumerate(criteria_data):
        col_start = i * 5 + 1
        # 제목
        shear_ws.merge_range(row, col_start, row, col_start + 3, title, formats[f'criteria_{color}_header'])
        # 조건
        shear_ws.merge_range(row + 1, col_start, row + 1, col_start + 3, condition, formats['formula_box'])
        # 설명
        shear_ws.merge_range(row + 2, col_start, row + 2, col_start + 3, description, formats[f'criteria_{color}'])
    
    shear_ws.set_row(row, 18)
    shear_ws.set_row(row + 1, 16)
    shear_ws.set_row(row + 2, 14)
    row += 4

    # ─── 전체 설계 결과 요약 ───
    shear_ws.merge_range(row, 0, row, 14, '📊 전체 설계 결과 요약', formats['summary_header'])
    shear_ws.set_row(row, 25)
    row += 1

    # 요약표 헤더
    summary_headers = ['Case', 'Pu(kN)', 'Vu(kN)', '판정결과', '최적설계', 'φVn(kN)', '안전율', '최종판정']
    col_positions = [2, 3, 4, 5, 6, 8, 10, 12]
    
    for i, (header, col) in enumerate(zip(summary_headers, col_positions)):
        if i == 4:  # 최적설계 컬럼은 2칸 병합
            shear_ws.merge_range(row, col, row, col + 1, header, formats['table_header'])
        else:
            shear_ws.write(row, col, header, formats['table_header'])
    shear_ws.set_row(row, 18)
    row += 1

    # 요약표 데이터
    for result in results:
        shear_ws.write(row, 2, f"Case {result['case']}", formats['table_text'])
        shear_ws.write(row, 3, f"{result['Pu']:.0f}", formats['table_cell'])
        shear_ws.write(row, 4, f"{result['Vu']:.1f}", formats['table_cell'])
        
        # 판정결과에 색상 적용
        color_format = formats[f'criteria_{result["color_style"]}_header']
        shear_ws.write(row, 5, result['category_short'], color_format)
        
        # 최적설계 (2칸 병합)
        shear_ws.merge_range(row, 6, row, 7, result['stirrups'], formats['table_text'])
        
        shear_ws.write(row, 8, f"{result['phi_Vn']/1000:.1f}", formats['table_cell'])
        shear_ws.write(row, 10, f"{result['safety_ratio']:.3f}", formats['table_cell'])
        
        # 최종판정
        final_result = "✅ 안전" if result['is_safe'] else "❌ NG"
        if not result['is_safe_section']:
            final_result += " (단면!)"
        result_format = formats['ok'] if result['is_safe'] else formats['ng']
        shear_ws.write(row, 12, final_result, result_format)
        
        shear_ws.set_row(row, 16)
        row += 1

    row += 2

    # ─── 케이스별 상세 계산 과정 ───
    shear_ws.merge_range(row, 0, row, 14, '⚙️ 케이스별 상세 계산 과정', formats['detail_header'])
    shear_ws.set_row(row, 25)
    row += 2

    num_symbols = ["❶", "❷", "❸", "❹", "❺", "❻", "❼", "❽", "❾", "❿"]
    
    for i, result in enumerate(results):
        # 케이스 헤더
        case_title = f"{num_symbols[i]} Case {result['case']} 검토 (Vu = {result['Vu']:.1f} kN)"
        shear_ws.merge_range(row, 1, row, 13, case_title, formats['case_title'])
        shear_ws.set_row(row, 20)
        row += 1
        
        # 결과 요약
        result_summary = f"결과: {result['category']} / {result['stirrups']}"
        shear_ws.merge_range(row, 1, row, 13, result_summary, formats['case_result'])
        shear_ws.set_row(row, 16)
        row += 1

        # 1단계: 축력 영향 계수
        shear_ws.merge_range(row, 1, row, 4, '1단계: 축력 영향 계수 (P증가)', formats['step_header'])
        shear_ws.set_row(row, 16)
        row += 1
        
        # 수식
        write_rich(row, 2, 'P', '증가', ' = 1 + Pu/(14×Ag)', formats['formula_box'])
        row += 1
        
        # 계산
        calc_text = f"P증가 = 1 + {result['Pu']*1000:,.0f} ÷ (14 × {Ag:,.0f}) = {result['p_factor']:.3f}"
        shear_ws.merge_range(row, 2, row, 12, calc_text, formats['calc_box'])
        row += 2

        # 2단계: 콘크리트 설계 전단강도
        shear_ws.merge_range(row, 1, row, 4, '2단계: 콘크리트 설계 전단강도 (φVc)', formats['step_header'])
        shear_ws.set_row(row, 16)
        row += 1
        
        # 수식 (여러 줄로 분할)
        write_rich(row, 2, 'φV', 'c', ' = φv × (1/6 × P증가 × λ × √fck × bw × d)', formats['formula_box'])
        row += 1
        
        # 계산
        calc_text = f"φVc = {phi_v} × (1/6 × {result['p_factor']:.3f} × {lamda} × √{fck} × {bw:,.0f} × {d:,.0f})"
        shear_ws.merge_range(row, 2, row, 12, calc_text, formats['calc_box'])
        row += 1
        calc_text = f"    = {result['phi_Vc']/1000:.1f} kN"
        shear_ws.merge_range(row, 2, row, 12, calc_text, formats['result_box'])
        row += 2

        # 3단계: 전단철근 필요성 판정
        shear_ws.merge_range(row, 1, row, 4, '3단계: 전단철근 필요성 판정', formats['step_header'])
        shear_ws.set_row(row, 16)
        row += 1
        
        # 비교값들
        comparison_text = f"Vu = {result['Vu']:.1f} kN,  φVc = {result['phi_Vc']/1000:.1f} kN,  ½φVc = {result['half_phi_Vc']/1000:.1f} kN"
        shear_ws.merge_range(row, 2, row, 12, comparison_text, formats['calc_box'])
        row += 1
        
        # 판정
        judgment_text = f"판정: {result['category']}"
        color_format = formats[f'criteria_{result["color_style"]}_header']
        shear_ws.merge_range(row, 2, row, 12, judgment_text, color_format)
        row += 2

        if result['category'] != "전단철근 불필요":
            # 4단계: 필요 전단철근량 및 간격 계산
            shear_ws.merge_range(row, 1, row, 4, '4단계: 필요 전단철근량 및 간격 계산', formats['step_header'])
            shear_ws.set_row(row, 16)
            row += 1
            
            calc_text = f"• 간격(최소철근량) = {Av_stirrup:.1f} ÷ {result['min_Av_s_req']:.4f} = {result['s_from_min_req']:.1f} mm"
            shear_ws.merge_range(row, 2, row, 12, calc_text, formats['calc_box'])
            row += 1
            
            calc_text = f"• 간격(최대허용) = {result['s_max_code']:.1f} mm"
            shear_ws.merge_range(row, 2, row, 12, calc_text, formats['calc_box'])
            row += 2

        # 5단계: 최종 배근 및 강도 검토
        shear_ws.merge_range(row, 1, row, 4, '5단계: 최종 배근 및 강도 검토', formats['step_header'])
        shear_ws.set_row(row, 16)
        row += 1
        
        # 최종 배근
        final_text = f"최종 배근: {result['stirrups']}"
        shear_ws.merge_range(row, 2, row, 6, final_text, formats['result_box'])
        row += 1
        
        # 설계 강도
        strength_text = f"설계 전단강도: φVn = {result['phi_Vc']/1000:.1f} + {result['phi_Vs']/1000:.1f} = {result['phi_Vn']/1000:.1f} kN"
        shear_ws.merge_range(row, 2, row, 12, strength_text, formats['calc_box'])
        row += 1
        
        # 안전성
        safety_symbol = '≥' if result['is_safe'] else '<'
        safety_text = f"안전성: φVn = {result['phi_Vn']/1000:.1f} kN {safety_symbol} Vu = {result['Vu']:.1f} kN"
        shear_ws.merge_range(row, 2, row, 12, safety_text, formats['calc_box'])
        row += 1
        
        # 안전율
        ratio_text = f"안전율: S.F = {result['phi_Vn']/1000:.1f} ÷ {result['Vu']:.1f} = {result['safety_ratio']:.3f}"
        shear_ws.merge_range(row, 2, row, 12, ratio_text, formats['calc_box'])
        row += 1
        
        # 최종 판정
        final_result = "✅ 최종 검토 결과: 안전" if result['is_safe'] else "❌ 최종 검토 결과: NG"
        result_format = formats['ok'] if result['is_safe'] else formats['ng']
        shear_ws.merge_range(row, 2, row, 12, final_result, result_format)
        shear_ws.set_row(row, 20)
        row += 1

        # 단면 검토 경고
        if not result['is_safe_section']:
            warning_text = "⚠️ 단면 검토 경고: 전단철근이 부담하는 강도가 최대 허용치를 초과했습니다. 단면 크기 상향이 필요합니다."
            shear_ws.merge_range(row, 2, row, 12, warning_text, formats['warning'])
            shear_ws.set_row(row, 18)
            row += 1

        row += 2  # 다음 케이스를 위한 간격

    return shear_ws