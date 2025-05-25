def create_additional_charts(chart_ws, wb):
    """추가 차트들 생성 (e vs c, e vs Pn, Pn vs phi 등)"""
    
    def create_secondary_chart(title, chart_type, x_col_offset, y_col_offset, data_start_col, chart_position):
        """보조 차트 생성"""
        chart = wb.add_chart({'type': chart_type, 'subtype': 'smooth'})
        
        # 이형철근 데이터
        chart.add_series({
            'name': f'이형철근 {title}',
            'categories': ['데이터', 2, data_start_col + x_col_offset, 100, data_start_col + x_col_offset],
            'values': ['데이터', 2, data_start_col + y_col_offset, 100, data_start_col + y_col_offset],
            'line': {'color': '#E74C3C', 'width': 3},
            'marker': {'type': 'circle', 'size': 4}
        })
        
        # 중공철근 데이터
        chart.add_series({
            'name': f'중공철근 {title}',
            'categories': ['데이터', 2, 13 + x_col_offset, 100, 13 + x_col_offset],
            'values': ['데이터', 2, 13 + y_col_offset, 100, 13 + y_col_offset],
            'line': {'color': '#3498DB', 'width': 3},
            'marker': {'type': 'circle', 'size': 4}
        })
        
        # X축, Y축 레이블 설정
        axis_labels = {
            'e vs c': ('편심거리 e [mm]', '중립축깊이 c [mm]'),
            'e vs Pn': ('편심거리 e [mm]', '공칭축력 Pn [kN]'),
            'Pn vs φ': ('공칭축력 Pn [kN]', '강도감소계수 φ'),
            'c vs φ': ('중립축깊이 c [mm]', '강도감소계수 φ')
        }
        
        x_label, y_label = axis_labels.get(title, ('X축', 'Y축'))
        
        chart.set_title({
            'name': title,
            'name_font': {'name': '맑은 고딕', 'size': 14, 'bold': True}
        })
        
        chart.set_x_axis({
            'name': x_label,
            'name_font': {'name': '맑은 고딕', 'size': 11, 'bold': True},
            'major_gridlines': {'visible': True},
            'num_font': {'name': '맑은 고딕', 'size': 10}
        })
        
        chart.set_y_axis({
            'name': y_label,
            'name_font': {'name': '맑은 고딕', 'size': 11, 'bold': True},
            'major_gridlines': {'visible': True},
            'num_font': {'name': '맑은 고딕', 'size': 10}
        })
        
        chart.set_legend({
            'position': 'top',
            'font': {'name': '맑은 고딕', 'size': 10}
        })
        
        # 차트 삽입
        chart_ws.insert_chart(chart_position[0], chart_position[1], chart, {
            'x_scale': 1.0, 'y_scale': 1.2
        })
        
        return chart
    
    # 추가 차트들 생성
    # e vs c (편심거리 vs 중립축깊이)
    create_secondary_chart('e vs c', 'scatter', 0, 1, 0, (35, 1))    # e[0], c[1]
    
    # e vs Pn (편심거리 vs 공칭축력)  
    create_secondary_chart('e vs Pn', 'scatter', 0, 2, 0, (35, 15))  # e[0], Pn[2]
    
    # Pn vs φ (공칭축력 vs 강도감소계수)
    create_secondary_chart('Pn vs φ', 'scatter', 2, 4, 0, (55, 1))  # Pn[2], φ[4]
    
    # c vs φ (중립축깊이 vs 강도감소계수)
    create_secondary_chart('c vs φ', 'scatter', 1, 4, 0, (55, 15))  # c[1], φ[4]


def create_pm_chart_excel(wb, In, R_data, F_data):
    """P-M 상관도 차트만 생성 (기존 데이터 시트 활용)"""
    
    # 차트 시트 생성
    chart_ws = wb.add_worksheet('P-M 상관도')
    chart_ws.activate()
    
    # ─── 1. 이형철근 P-M 상관도 ─────────────────
    def create_chart(title, data_start_col, chart_position):
        """P-M 상관도 차트 생성"""
        chart = wb.add_chart({'type': 'scatter', 'subtype': 'smooth'})
        
        # Pn-Mn 곡선 (빨간색 점선)
        chart.add_series({
            'name': 'Pn-Mn Diagram',
            'categories': ['데이터', 2, data_start_col + 3, 100, data_start_col + 3],  # Mₙ 열
            'values': ['데이터', 2, data_start_col + 2, 100, data_start_col + 2],      # Pₙ 열
            'line': {'color': '#E74C3C', 'width': 3, 'dash_type': 'dash'},
            'marker': {'type': 'none'}
        })
        
        # φPn-φMn 곡선 (파란색 실선)
        chart.add_series({
            'name': 'φPn-φMn Diagram',
            'categories': ['데이터', 2, data_start_col + 6, 100, data_start_col + 6],  # ϕMₙ 열
            'values': ['데이터', 2, data_start_col + 5, 100, data_start_col + 5],      # ϕPₙ 열
            'line': {'color': '#3498DB', 'width': 4},
            'marker': {'type': 'none'},
            'fill': {'color': '#E8F6F3', 'transparency': 70}
        })
        
        # 차트 설정
        chart.set_title({
            'name': f'{title} P-M 상관도',
            'name_font': {'name': '맑은 고딕', 'size': 16, 'bold': True}
        })
        
        chart.set_x_axis({
            'name': 'Mn or ϕMn [kN·m]',
            'name_font': {'name': '맑은 고딕', 'size': 12, 'bold': True},   # 축 이름            
            'major_gridlines': {'visible': True},
            'min': 0,
            'num_font': {'name': '맑은 고딕', 'size': 12, 'bold': True},       # 눈금 숫자 크기 설정
            'num_format': '0'  # ⬅ 여기를 추가하면 정수만 표시됨!
        })
        
        chart.set_y_axis({
            'name': 'Pn or ϕPn [kN]',
            'name_font': {'name': '맑은 고딕', 'size': 12, 'bold': True},
            'major_gridlines': {'visible': True},
            'num_font': {'name': '맑은 고딕', 'size': 12, 'bold': True},       # 눈금 숫자 크기 설정
            'num_format': '0'  # ⬅ 여기를 추가하면 정수만 표시됨!
        })
        
        chart.set_legend({
            'position': 'top',
            'font': {'name': '맑은 고딕', 'size': 11, 'bold': True}
        })
        
        # 차트 삽입
        chart_ws.insert_chart(chart_position[0], chart_position[1], chart, {
            'x_scale': 1.5, 'y_scale': 2.0
        })
        
        return chart
    
    # ─── 2. 차트 생성 및 배치 ─────────────────
    
    # 제목
    title_format = wb.add_format({
        'bold': True, 'font_size': 20, 'bg_color': '#1F4E79', 'font_color': 'white',
        'align': 'center', 'valign': 'vcenter', 'border': 2
    })
    chart_ws.merge_range(0, 0, 0, 25, '📊 P-M 상관도', title_format)
    chart_ws.set_row(0, 35)
    
    # 이형철근 차트 (왼쪽)
    create_chart('이형철근', 0, (2, 1))
    
    # 중공철근 차트 (오른쪽) 
    create_chart('중공철근', 13, (2, 15))  # 13번 컬럼부터 시작 (11컬럼 + 2간격)

    # 추가 차트들 생성
    create_additional_charts(chart_ws, wb)
    
    return chart_ws
