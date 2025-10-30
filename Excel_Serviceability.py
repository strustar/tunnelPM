import xlsxwriter
import math

def check_crack_section(P0_case, M0_case, In):
    """
    균열단면 여부를 판단하는 함수
    조건: M*y/I - P/A ≥ 0.63*√fck → 균열단면
    """
    # 단면 제원 추출
    b = float(getattr(In, 'be', 1000))  # mm
    h = float(getattr(In, 'height', 500))  # mm
    fck = float(getattr(In, 'fck', 24))  # MPa

    # 단면 성질 계산
    A = b * h  # 단면적 (mm²)
    I = b * h**3 / 12  # 단면2차모멘트 (mm⁴)
    y = h / 2  # 중립축에서 최외단까지 거리 (mm)

    # 단위 환산 (kN → N, kN·m → N·mm)
    P0_N = P0_case * 1000  # N
    M0_Nmm = M0_case * 1000 * 1000  # N·mm

    # 균열 판정 계산
    stress_term = (M0_Nmm * y / I) - (P0_N / A) if I > 0 and A > 0 else 0 # MPa 단위
    crack_limit = 0.63 * math.sqrt(fck)  # MPa

    is_cracked = stress_term >= crack_limit

    return is_cracked, stress_term, crack_limit, A, I, y

def _render_case_to_excel(ws, start_row, start_col, data, In, i, symbol, formats):
    """
    단일 케이스 분석 결과를 엑셀 시트에 상세하게 렌더링
    """
    row = start_row
    
    # 데이터 추출
    fs_case, x_case = data.fss[i], data.x[i]
    P0_case, M0_case = In.P0[i], In.M0[i]

    # 케이스 헤더
    ws.merge_range(row, start_col, row, start_col + 6, f"{symbol}번 검토", formats['case_header'])
    ws.set_row(row, 30)
    row += 1

    # --- Step 0: 균열단면 체크 ---
    ws.merge_range(row, start_col, row, start_col + 6, "🔍 Step 0: 균열단면 체크", formats['step_header'])
    row += 1
    
    is_cracked, stress_term, crack_limit, A, I, y = check_crack_section(P0_case, M0_case, In)
    
    # 균열 판정 계산 상세 표시
    ws.write(row, start_col, "균열 판정 계산:", formats['subheader'])
    row += 1
    
    ws.write(row, start_col, "공식", formats['label'])
    ws.merge_range(row, start_col + 1, row, start_col + 6, "M × y / I - P / A  vs  0.63 × √fck", formats['formula'])
    row += 1
    
    # 상세 계산 과정
    b = float(getattr(In, 'be', 1000))
    h = float(getattr(In, 'height', 500))
    fck = float(getattr(In, 'fck', 24))
    
    ws.write(row, start_col, "계산", formats['label'])
    calculation_text = f"{M0_case*1e6:,.0f} × {y:,.1f} / {I:,.0f} - {P0_case*1000:,.0f} / {A:.0f} = {stress_term:,.1f} MPa"
    ws.merge_range(row, start_col + 1, row, start_col + 6, calculation_text, formats['calculation'])
    row += 1
    
    ws.write(row, start_col, "한계값", formats['label'])
    limit_text = f"0.63 × √{fck:.1f} = 0.63 × {math.sqrt(fck):.1f} = {crack_limit:,.1f} MPa"
    ws.merge_range(row, start_col + 1, row, start_col + 6, limit_text, formats['calculation'])
    row += 1

    # 균열 판정 결과
    if not is_cracked:
        ws.merge_range(row, start_col, row + 1, start_col + 6, 
                      f"✅ 비균열 단면\n{stress_term:,.1f} MPa < {crack_limit:,.1f} MPa\n🎉 균열 검토 불필요", 
                      formats['no_crack_box'])
        ws.set_row(row, 50)
        row += 2
        
        # 최종 판정 (비균열)
        ws.merge_range(row, start_col, row, start_col + 6, "✅ 균열 검토 불필요 (비균열 단면)", formats['result_success'])
        row += 2
        return row - start_row

    # 균열 단면인 경우
    ws.merge_range(row, start_col, row + 1, start_col + 6, 
                  f"⚠️ 균열 단면\n{stress_term:,.1f} MPa ≥ {crack_limit:,.1f} MPa\n🔍 균열 검토 필요", 
                  formats['crack_box'])
    ws.set_row(row, 50)
    row += 2

    # 케이스 분류
    if P0_case == 0:
        case_title = f"🎯 Case Ⅰ: 특수한 경우\n순수 휨 (P₀ = {P0_case:,.1f} kN, M₀ = {M0_case:,.1f} kN·m)\n📊 보(Beam)에 해당 - 해석적 풀이 적용"
        box_format = formats['case_special_box']
    else:
        case_title = f"⚙️ Case Ⅱ: 일반적인 경우\n축력+휨 (P₀ = {P0_case:,.1f} kN, M₀ = {M0_case:,.1f} kN·m)\n🏛️ 기둥(Column)에 해당 - 수치해석 필요"
        box_format = formats['case_general_box']
    
    ws.merge_range(row, start_col, row + 2, start_col + 6, case_title, box_format)
    ws.set_row(row, 55)
    row += 3

    # --- A. 탄성 해석 과정 ---
    ws.merge_range(row, start_col, row, start_col + 6, "🔬 A. 탄성 해석 과정 (수치해석 접근)", formats['section_header'])
    row += 1

    if P0_case == 0:
        # 특수한 경우: 순수 휨
        ws.write(row, start_col, "Step 1: 연립 평형방정식 설정", formats['step_title'])
        row += 1
        ws.write(row, start_col, "축력 평형", formats['label'])
        ws.merge_range(row, start_col + 1, row, start_col + 6, "P₀ = C - T = 1/2 × fc × b × x - As × fs", formats['formula'])
        row += 1
        ws.write(row, start_col, "모멘트 평형", formats['label'])
        ws.merge_range(row, start_col + 1, row, start_col + 6, "M₀ = C × (h/2 - x/3) + T × (d - h/2)", formats['formula'])
        row += 1
        
        ws.write(row, start_col, "Step 2: 수치해석 결과", formats['step_title'])
        row += 1
        ws.write(row, start_col, "주의", formats['label'])
        ws.merge_range(row, start_col + 1, row, start_col + 6, "직접 풀이 불가능하여 반복계산 통해 도출", formats['explanation'])
        row += 1
        ws.write(row, start_col, "중립축", formats['label'])
        ws.merge_range(row, start_col + 1, row, start_col + 6, f"x = {x_case:,.1f} mm (수치해)", formats['result_value'])
        row += 1
        ws.write(row, start_col, "철근응력", formats['label'])
        ws.merge_range(row, start_col + 1, row, start_col + 6, f"fs = {fs_case:,.1f} MPa (수치해)", formats['result_value'])
        row += 1
        
    else:
        # 일반적인 경우: 축력+휨
        ws.write(row, start_col, "Step 1: 연립 평형방정식 설정", formats['step_title'])
        row += 1
        ws.write(row, start_col, "축력 평형", formats['label'])
        ws.merge_range(row, start_col + 1, row, start_col + 6, "P₀ = C - T = 1/2 × fc × b × x - As × fs", formats['formula'])
        row += 1
        ws.write(row, start_col, "모멘트 평형", formats['label'])
        ws.merge_range(row, start_col + 1, row, start_col + 6, "M₀ = C × (h/2 - x/3) + T × (d - h/2)", formats['formula'])
        row += 1
        
        ws.write(row, start_col, "Step 2: 수치해석 결과", formats['step_title'])
        row += 1
        ws.write(row, start_col, "특징", formats['label'])
        ws.merge_range(row, start_col + 1, row, start_col + 6, "비선형 연립방정식 → fsolve 등 반복계산 필요", formats['explanation'])
        row += 1
        ws.write(row, start_col, "중립축", formats['label'])
        ws.merge_range(row, start_col + 1, row, start_col + 6, f"x = {x_case:,.1f} mm (수치해)", formats['result_value'])
        row += 1
        ws.write(row, start_col, "철근응력", formats['label'])
        ws.merge_range(row, start_col + 1, row, start_col + 6, f"fs = {fs_case:,.1f} MPa (수치해)", formats['result_value'])
        row += 1

    # --- B. 휨균열 제어 검토 ---
    ws.merge_range(row, start_col, row, start_col + 6, "📏 B. 휨균열 제어 검토", formats['section_header'])
    row += 1

    # Step 1: 최외단 철근 응력 산정
    ws.write(row, start_col, "Step 1: 최외단 철근 응력 fst 산정", formats['step_title'])
    row += 1
    fst_case = fs_case  # 1단 배근 가정
    ws.write(row, start_col, "가정", formats['label'])
    ws.merge_range(row, start_col + 1, row, start_col + 6, "fst = fs × (h - dc - x) / (d - x) ≈ fs", formats['formula'])
    row += 1
    ws.write(row, start_col, "결과", formats['label'])
    ws.merge_range(row, start_col + 1, row, start_col + 6, f"fst = {fst_case:,.1f} MPa", formats['result_value'])
    row += 1

    # Step 2: 최대 허용 간격 산정
    ws.write(row, start_col, "Step 2: 최대 허용 간격 산정 [KDS 기준]", formats['step_title'])
    row += 1
    
    # 조건 1
    ws.write(row, start_col, "조건 1", formats['label'])
    ws.merge_range(row, start_col + 1, row, start_col + 6, "s₁ = 375 × (210 / fst) - 2.5 × Cc", formats['formula'])
    row += 1
    
    s_allowed_1 = 375 * (210 / fst_case) - 2.5 * In.Cc if fst_case > 0 else float('inf')
    ws.write(row, start_col, "계산", formats['label'])
    calc_text_1 = f"s₁ = 375 × (210 / {fst_case:,.1f}) - 2.5 × {In.Cc:,.1f} = {s_allowed_1:,.1f} mm"
    ws.merge_range(row, start_col + 1, row, start_col + 6, calc_text_1, formats['calculation'])
    row += 1
    
    # 조건 2
    ws.write(row, start_col, "조건 2", formats['label'])
    ws.merge_range(row, start_col + 1, row, start_col + 6, "s₂ = 300 × (210 / fst)", formats['formula'])
    row += 1
    
    s_allowed_2 = 300 * (210 / fst_case) if fst_case > 0 else float('inf')
    ws.write(row, start_col, "계산", formats['label'])
    calc_text_2 = f"s₂ = 300 × (210 / {fst_case:,.1f}) = {s_allowed_2:,.1f} mm"
    ws.merge_range(row, start_col + 1, row, start_col + 6, calc_text_2, formats['calculation'])
    row += 1
    
    # 최종 허용 간격
    s_allowed_final = min(s_allowed_1, s_allowed_2)
    ws.write(row, start_col, "최종 허용", formats['label'])
    ws.merge_range(row, start_col + 1, row, start_col + 6, f"sallow = min(s₁, s₂) = {s_allowed_final:,.1f} mm", formats['result_value_bold'])
    row += 1

    # Step 3: 최종 판정
    ws.write(row, start_col, "Step 3: 최종 판정", formats['step_title'])
    row += 1
    
    ws.write(row, start_col, "최종 허용 간격", formats['metric_label'])
    ws.merge_range(row, start_col + 1, row, start_col + 2, f"{s_allowed_final:,.1f} mm", formats['metric_value'])
    ws.merge_range(row, start_col + 3, row, start_col + 6, "Min(s₁, s₂)", formats['metric_note'])
    row += 1
    
    ws.write(row, start_col, "실제 배근 간격", formats['metric_label'])
    ws.merge_range(row, start_col + 1, row, start_col + 2, f"{In.sb[0]:,.1f} mm", formats['metric_value'])
    row += 1

    # 최종 판정 결과
    if In.sb[0] <= s_allowed_final:
        result_text = f"✅ O.K. (배근 간격 {In.sb[0]:.1f} mm ≤ 허용 간격 {s_allowed_final:,.1f} mm)"
        result_format = formats['result_success']
    else:
        result_text = f"❌ N.G. (배근 간격 {In.sb[0]:.1f} mm > 허용 간격 {s_allowed_final:,.1f} mm)"
        result_format = formats['result_error']
    
    ws.merge_range(row, start_col, row + 1, start_col + 6, result_text, result_format)
    ws.set_row(row, 40)
    row += 2
    
    return row - start_row

def create_serviceability_sheet(wb, In, R, F):
    """
    엑셀 워크북에 '사용성 검토' 시트를 생성하고 상세 해석 결과를 작성
    """
    ws = wb.add_worksheet('사용성 검토')
    base_font = {'font_name': 'Noto Sans KR', 'border': 1, 'valign': 'vcenter', 'bold': True}

    # --- 서식 정의 (글자 크기 최소 12pt 기준) ---
    formats = {
        'main_title': wb.add_format({**base_font,
            'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#FF8C00', 'font_color': '#000000', 'border': 1
        }),
        'case_header': wb.add_format({**base_font,
            'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#4472C4', 'font_color': 'white', 'border': 1
        }),
        'section_header': wb.add_format({**base_font,
            'bold': True, 'font_size': 13, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#D9D9D9', 'border': 1
        }),
        'step_header': wb.add_format({**base_font,
            'bold': True, 'font_size': 13, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#E7E6E6', 'border': 1
        }),
        'step_title': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'left', 'valign': 'vcenter',
            'fg_color': '#F2F2F2', 'border': 1, 'indent': 1
        }),
        'label': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#E7E6E6', 'border': 1
        }),
        'formula': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'left', 'valign': 'vcenter', 'border': 1,
            'font_name': 'Noto Sans KR', 'indent': 1
        }),
        'calculation': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'left', 'valign': 'vcenter', 'border': 1,
            'font_name': 'Noto Sans KR', 'indent': 1, 'fg_color': '#F0F8FF'
        }),
        'explanation': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'indent': 1
        }),
        'result_value': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'left', 'valign': 'vcenter', 'border': 1,
            'indent': 1, 'fg_color': '#F0F8F0'
        }),
        'result_value_bold': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'left', 'valign': 'vcenter',
            'border': 1, 'indent': 1, 'fg_color': '#F0F8F0'
        }),
        'metric_label': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'fg_color': '#E7E6E6'
        }),
        'metric_value': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'fg_color': '#FFFF99'
        }),
        'metric_note': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'left', 'valign': 'vcenter', 'border': 1,
            'italic': True, 'font_color': '#666666'
        }),
        'no_crack_box': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter',
            'font_color': 'green', 'border': 1, 'text_wrap': True
        }),
        'crack_box': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter',
             'font_color': 'purple', 'border': 1, 'text_wrap': True
        }),
        'case_special_box': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter',
            'font_color': 'magenta', 'border': 1, 'text_wrap': True
        }),
        'case_general_box': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter',
            'font_color': 'blue', 'border': 1, 'text_wrap': True
        }),
        'result_success': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#28a745', 'font_color': 'white', 'border': 1, 'text_wrap': True
        }),
        'result_error': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#dc3545', 'font_color': 'white', 'border': 1, 'text_wrap': True
        }),
        'subheader': wb.add_format({**base_font,
            'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#E7E6E6', 'border': 1
        }),
    }

    # --- 컬럼 너비 설정 ---
    ws.set_column('A:A', 22)    # 항목명
    ws.set_column('B:G', 18)    # 데이터 영역
    ws.set_column('H:H', 3)     # 구분선
    ws.set_column('I:I', 22)    # 항목명
    ws.set_column('J:O', 18)    # 데이터 영역

    # --- 메인 타이틀 ---
    ws.merge_range('A1:O1', '하중 케이스별 상세 균열 검토', formats['main_title'])
    ws.set_row(0, 40)
    
    ws.merge_range('A2:G2', '이형철근', formats['case_header'])
    ws.merge_range('I2:O2', '중공철근', formats['case_header'])

    # --- 케이스별 상세 결과 렌더링 ---
    num_symbols = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"]
    current_row = 2

    for i in range(len(In.P0)):
        # 좌측: R 데이터 검토
        rows_r = _render_case_to_excel(ws, current_row, 0, R, In, i, num_symbols[i], formats)
        
        # 우측: F 데이터 검토
        rows_f = _render_case_to_excel(ws, current_row, 8, F, In, i, num_symbols[i], formats)

        # 다음 케이스를 위한 행 위치 조정
        current_row += max(rows_r, rows_f) + 1

    return ws
