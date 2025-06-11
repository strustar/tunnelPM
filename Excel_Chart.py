def create_chart_excel(wb, In, R_data, F_data):
    """P-M 상관도 차트만 생성 (기존 데이터 시트 활용)"""
    
    # 차트 시트 생성
    chart_ws = wb.add_worksheet('P-M 상관도')
    # chart_ws.activate()
    
    # ─── 1. 설계점 색상 및 스타일 정의 ─────────────────
    design_point_colors = ['#E74C3C', '#F39C12', '#27AE60', '#8E44AD', '#3498DB', '#E67E22']  # 빨강, 주황, 초록, 보라, 파랑, 갈색
    design_point_symbols = ["①", "②", "③", "④", "⑤", "⑥"]
    
    # ─── 2. 설계점 데이터를 워크시트에 쓰기 ─────────────────
    def write_design_points_data():
        """설계점 데이터를 차트 시트에 쓰기"""
        # 설계점 데이터 시작 위치
        data_start_row = 50
        data_start_col = 0
        
        # 헤더 쓰기
        chart_ws.write(data_start_row, data_start_col, 'Case')
        chart_ws.write(data_start_row, data_start_col + 1, 'Pu')
        chart_ws.write(data_start_row, data_start_col + 2, 'Mu')
        
        # 데이터 쓰기
        for i in range(len(In.Pu)):
            row = data_start_row + 1 + i
            chart_ws.write(row, data_start_col, f'Case {i+1}')
            chart_ws.write(row, data_start_col + 1, In.Pu[i])
            chart_ws.write(row, data_start_col + 2, In.Mu[i])
        
        return data_start_row, data_start_col
    
    # 설계점 데이터 쓰기
    data_row, data_col = write_design_points_data()

    # ─── 3. P-M 상관도 차트 생성 함수 ─────────────────
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
            'name': 'ϕPn-ϕMn Diagram',
            'categories': ['데이터', 2, data_start_col + 6, 100, data_start_col + 6],  # ϕMₙ 열
            'values': ['데이터', 2, data_start_col + 5, 100, data_start_col + 5],      # ϕPₓ 열
            'line': {'color': '#3498DB', 'width': 4},
            'marker': {'type': 'none'},
            'fill': {'color': '#E8F6F3', 'transparency': 70}
        })
        
        # ─── 설계점 추가 (In.Pu, In.Mu) ─────────────────
        for i in range(len(In.Pu)):
            if i < len(design_point_colors):  # 색상 범위 내에서만
                # 각 설계점을 개별 시리즈로 추가
                point_name = f'Case {i+1} ({design_point_symbols[i]})'
                point_row = data_row + 1 + i
                
                chart.add_series({
                    'name': point_name,
                    'categories': ['P-M 상관도', point_row, data_col + 2, point_row, data_col + 2],  # Mu
                    'values': ['P-M 상관도', point_row, data_col + 1, point_row, data_col + 1],      # Pu
                    'line': {'none': True},    # 선 없음
                    'marker': {
                        'type': 'circle',
                        'size': 12,
                        'border': {'color': design_point_colors[i], 'width': 3},
                        'fill': {'color': design_point_colors[i], 'transparency': 20}
                    }
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
            'num_format': '0'  # 정수만 표시
        })
        
        chart.set_y_axis({
            'name': 'Pn or ϕPn [kN]',
            'name_font': {'name': '맑은 고딕', 'size': 12, 'bold': True},
            'major_gridlines': {'visible': True},
            'num_font': {'name': '맑은 고딕', 'size': 12, 'bold': True},       # 눈금 숫자 크기 설정
            'num_format': '0'  # 정수만 표시
        })
        
        chart.set_legend({
            'position': 'top',
            'font': {'name': '맑은 고딕', 'size': 10, 'bold': True}
        })
        
        # 차트 삽입
        chart_ws.insert_chart(chart_position[0], chart_position[1], chart, {
            'x_scale': 1.5, 'y_scale': 2.0
        })
        
        return chart
    
    # ─── 4. 설계점 데이터 테이블 생성 ─────────────────
    def create_design_points_table():
        """설계점 정보를 테이블로 표시"""
        # 스타일 정의
        header_format = wb.add_format({
            'bold': True, 'font_size': 12, 'bg_color': '#34495E', 'font_color': 'white',
            'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        
        cell_format = wb.add_format({
            'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'border': 1,
            'num_format': '#,##0.0'
        })
        
        text_format = wb.add_format({
            'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        
        # 테이블 시작 위치
        start_row = 2
        start_col = 30
        
        # 테이블 제목
        title_format = wb.add_format({
            'bold': True, 'font_size': 14, 'bg_color': '#2C3E50', 'font_color': 'white',
            'align': 'center', 'valign': 'vcenter', 'border': 2
        })
        
        chart_ws.merge_range(start_row, start_col, start_row, start_col + 3, 
                           '📋 설계점 정보', title_format)
        chart_ws.set_row(start_row, 25)
        
        # 헤더
        headers = ['Case', 'Pu (kN)', 'Mu (kN·m)', '마커']
        for i, header in enumerate(headers):
            chart_ws.write(start_row + 1, start_col + i, header, header_format)
        
        chart_ws.set_row(start_row + 1, 20)
        
        # 데이터 행
        for i in range(len(In.Pu)):
            if i < len(design_point_symbols):
                row = start_row + 2 + i
                
                # Case 번호
                chart_ws.write(row, start_col, f'Case {i+1}', text_format)
                
                # Pu 값
                chart_ws.write(row, start_col + 1, In.Pu[i], cell_format)
                
                # Mu 값
                chart_ws.write(row, start_col + 2, In.Mu[i], cell_format)
                
                # 마커 (색상 적용)
                marker_format = wb.add_format({
                    'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'border': 1,
                    'bg_color': design_point_colors[i], 'font_color': 'white', 'bold': True
                })
                chart_ws.write(row, start_col + 3, design_point_symbols[i], marker_format)
                
                chart_ws.set_row(row, 18)
        
        # 컬럼 너비 설정
        chart_ws.set_column(start_col, start_col, 10)      # Case
        chart_ws.set_column(start_col + 1, start_col + 1, 12)  # Pu
        chart_ws.set_column(start_col + 2, start_col + 2, 14)  # Mu
        chart_ws.set_column(start_col + 3, start_col + 3, 8)   # 마커
    
    # ─── 5. 차트 생성 및 배치 ─────────────────
    
    # 제목
    title_format = wb.add_format({
        'bold': True, 'font_size': 20, 'bg_color': '#1F4E79', 'font_color': 'white',
        'align': 'center', 'valign': 'vcenter', 'border': 2
    })
    chart_ws.merge_range(0, 0, 0, 35, '📊 P-M 상관도 (설계점 포함)', title_format)
    chart_ws.set_row(0, 35)
    
    # 이형철근 차트 (왼쪽)
    create_chart('이형철근', 0, (2, 1))
    
    # 중공철근 차트 (오른쪽) 
    create_chart('중공철근', 13, (2, 15))  # 13번 컬럼부터 시작 (11컬럼 + 2간격)
    
    # 설계점 정보 테이블 생성
    create_design_points_table()
    
    # ─── 6. 설명 텍스트 추가 ─────────────────
    note_format = wb.add_format({
        'font_size': 11, 'align': 'left', 'valign': 'vcenter', 'border': 1,
        'bg_color': '#F8F9FA', 'text_wrap': True
    })
    
    note_text = ("• 빨간 점선: 공칭강도 Pn-Mn 곡선\n"
                "• 파란 실선: 설계강도 ϕPn-ϕMn 곡선 (안전영역)\n"
                "• 원형 마커: 실제 설계점 (Pu, Mu)\n"
                "• 설계점이 파란 영역 내부에 있으면 안전합니다.")
    
    chart_ws.merge_range(34, 1, 38, 13, note_text, note_format)
    
    return chart_ws

def create_additional_analysis_charts(wb, In, R_data, F_data, chart_ws):
    """
    PM 상관도 기반 추가 분석 차트 생성 - 공학적 의미 있는 곡선들
    - 1차: 편심거리 e 기반 곡선 (e-φPn, e-φMn, e-εt)
    - 2차: 중립축 깊이 c 기반 곡선 (c-φPn, c-φMn, c-ft)
    - 3차: 응력·변형률 조합 (εt-ft, εc-φPn)
    - 4차: 무차원 비율 곡선 (상호작용도, 설계효율성)
    
    Parameters:
    - wb: xlsxwriter 워크북 객체
    - In: 입력 데이터
    - R_data, F_data: 철근 데이터
    - chart_ws: 'P-M 상관도' 워크시트 객체
    """
    
    # ─── 1. 차트 스타일 정의 ─────────────────
    chart_colors = {
        'rebar': '#E74C3C',      # 이형철근 - 빨간색
        'hollow': '#3498DB',     # 중공철근 - 파란색
        'design': '#27AE60',     # 설계강도 - 초록색
        'nominal': '#F39C12',    # 공칭강도 - 주황색
        'ratio': '#8E44AD',      # 비율 - 보라색
        'strain': '#E67E22',     # 변형률 - 갈색
        'stress': '#16A085',     # 응력 - 청록색
        'efficiency': '#D35400'   # 효율성 - 주황갈색
    }
    
    # ─── 2. 편심거리-축력 허용한계 곡선 (e-φPn) ─────────────────
    def create_eccentricity_axial_limit_chart():
        """편심거리에 따른 축력 허용한계 곡선 - 현장 편심 발생시 즉시 안전성 판단"""
        chart = wb.add_chart({'type': 'scatter', 'subtype': 'smooth_with_markers'})
        
        # 이형철근 축력 허용한계
        chart.add_series({
            'name': '이형철근 φPn (축력한계)',
            'categories': ['데이터', 2, 0, 100, 0],    # e 열 (편심거리)
            'values': ['데이터', 2, 5, 100, 5],        # φPn 열
            'line': {'color': chart_colors['rebar'], 'width': 4},
            'marker': {'type': 'circle', 'size': 6, 'fill': {'color': chart_colors['rebar']}}
        })
        
        # 중공철근 축력 허용한계
        chart.add_series({
            'name': '중공철근 φPn (축력한계)',
            'categories': ['데이터', 2, 13, 100, 13],  # e 열 (중공철근)
            'values': ['데이터', 2, 18, 100, 18],      # φPn 열 (중공철근)
            'line': {'color': chart_colors['hollow'], 'width': 4},
            'marker': {'type': 'square', 'size': 6, 'fill': {'color': chart_colors['hollow']}}
        })
        
        chart.set_title({
            'name': '편심거리-축력 허용한계 곡선 (안전성 즉시판단)',
            'name_font': {'name': '맑은 고딕', 'size': 14, 'bold': True}
        })
        
        chart.set_x_axis({
            'name': '편심거리 e = M/P [mm]',
            'name_font': {'name': '맑은 고딕', 'size': 11, 'bold': True},
            'major_gridlines': {'visible': True},
            'num_format': '#,##0'
        })
        
        chart.set_y_axis({
            'name': '축력 허용한계 φPn [kN]',
            'name_font': {'name': '맑은 고딕', 'size': 11, 'bold': True},
            'major_gridlines': {'visible': True},
            'num_format': '#,##0'
        })
        
        chart.set_legend({
            'position': 'top',
            'font': {'name': '맑은 고딕', 'size': 9}
        })
        
        chart_ws.insert_chart(44, 1, chart, {
            'x_scale': 1.2, 'y_scale': 1.5
        })
        
        return chart
    
    # ─── 3. 중립축-파괴양식 구분 곡선 (c-φMn) ─────────────────
    def create_neutral_axis_failure_mode_chart():
        """중립축 깊이에 따른 파괴양식 구분선"""
        chart = wb.add_chart({'type': 'scatter', 'subtype': 'smooth_with_markers'})
        
        # 이형철근 모멘트-중립축 관계
        chart.add_series({
            'name': '이형철근 φMn',
            'categories': ['데이터', 2, 1, 100, 1],    # c 열 (중립축 깊이)
            'values': ['데이터', 2, 6, 100, 6],        # φMn 열
            'line': {'color': chart_colors['rebar'], 'width': 3},
            'marker': {'type': 'circle', 'size': 5, 'fill': {'color': chart_colors['rebar']}}
        })
        
        # 중공철근 모멘트-중립축 관계
        chart.add_series({
            'name': '중공철근 φMn',
            'categories': ['데이터', 2, 14, 100, 14],  # c 열 (중공철근)
            'values': ['데이터', 2, 19, 100, 19],      # φMn 열 (중공철근)
            'line': {'color': chart_colors['hollow'], 'width': 3},
            'marker': {'type': 'square', 'size': 5, 'fill': {'color': chart_colors['hollow']}}
        })
        
        chart.set_title({
            'name': '중립축-파괴양식 구분 곡선 (압축↔인장 파괴영역)',
            'name_font': {'name': '맑은 고딕', 'size': 14, 'bold': True}
        })
        
        chart.set_x_axis({
            'name': '중립축 깊이 c [mm]',
            'name_font': {'name': '맑은 고딕', 'size': 11, 'bold': True},
            'major_gridlines': {'visible': True},
            'num_format': '#,##0'
        })
        
        chart.set_y_axis({
            'name': '모멘트 φMn [kN·m]',
            'name_font': {'name': '맑은 고딕', 'size': 11, 'bold': True},
            'major_gridlines': {'visible': True},
            'num_format': '#,##0'
        })
        
        chart.set_legend({
            'position': 'top',
            'font': {'name': '맑은 고딕', 'size': 9}
        })
        
        chart_ws.insert_chart(44, 15, chart, {
            'x_scale': 1.2, 'y_scale': 1.5
        })
        
        return chart
    
    # ─── 4. 인장변형률-응력 선도 (εt-ft) ─────────────────
    def create_tension_strain_stress_chart():
        """인장변형률-응력 선도: 균열검토/피로검토용"""
        chart = wb.add_chart({'type': 'scatter', 'subtype': 'smooth_with_markers'})
        
        # 이형철근 인장 변형률-응력 (데이터가 있는 경우만)
        chart.add_series({
            'name': '이형철근 εt-ft',
            'categories': ['데이터', 2, 7, 100, 7],    # εt 열
            'values': ['데이터', 2, 8, 100, 8],        # ft 열
            'line': {'color': chart_colors['strain'], 'width': 3},
            'marker': {'type': 'circle', 'size': 5, 'fill': {'color': chart_colors['strain']}}
        })
        
        # 중공철근 인장 변형률-응력 (데이터가 있는 경우만)
        chart.add_series({
            'name': '중공철근 εt-ft',
            'categories': ['데이터', 2, 20, 100, 20],  # εt 열 (중공철근)
            'values': ['데이터', 2, 21, 100, 21],      # ft 열 (중공철근)
            'line': {'color': chart_colors['stress'], 'width': 3},
            'marker': {'type': 'square', 'size': 5, 'fill': {'color': chart_colors['stress']}}
        })
        
        chart.set_title({
            'name': '인장변형률-응력 선도 (균열/피로 검토)',
            'name_font': {'name': '맑은 고딕', 'size': 14, 'bold': True}
        })
        
        chart.set_x_axis({
            'name': '인장변형률 εt',
            'name_font': {'name': '맑은 고딕', 'size': 11, 'bold': True},
            'major_gridlines': {'visible': True},
            'num_format': '0.0000'
        })
        
        chart.set_y_axis({
            'name': '인장응력 ft [MPa]',
            'name_font': {'name': '맑은 고딕', 'size': 11, 'bold': True},
            'major_gridlines': {'visible': True},
            'num_format': '#,##0'
        })
        
        chart.set_legend({
            'position': 'top',
            'font': {'name': '맑은 고딕', 'size': 9}
        })
        
        chart_ws.insert_chart(62, 1, chart, {
            'x_scale': 1.2, 'y_scale': 1.5
        })
        
        return chart
    
    # ─── 5. 압축변형률-축력 곡선 (εc-φPn) ─────────────────
    def create_compression_strain_axial_chart():
        """압축변형률에 따른 축력 저감 경향"""
        chart = wb.add_chart({'type': 'scatter', 'subtype': 'smooth_with_markers'})
        
        # 이형철근 압축변형률-축력 (데이터가 있는 경우만)
        chart.add_series({
            'name': '이형철근 εc-φPn',
            'categories': ['데이터', 2, 9, 100, 9],    # εc 열
            'values': ['데이터', 2, 5, 100, 5],        # φPn 열
            'line': {'color': chart_colors['rebar'], 'width': 3},
            'marker': {'type': 'circle', 'size': 5, 'fill': {'color': chart_colors['rebar']}}
        })
        
        # 중공철근 압축변형률-축력 (데이터가 있는 경우만)
        chart.add_series({
            'name': '중공철근 εc-φPn',
            'categories': ['데이터', 2, 22, 100, 22],  # εc 열 (중공철근)
            'values': ['데이터', 2, 18, 100, 18],      # φPn 열 (중공철근)
            'line': {'color': chart_colors['hollow'], 'width': 3},
            'marker': {'type': 'square', 'size': 5, 'fill': {'color': chart_colors['hollow']}}
        })
        
        chart.set_title({
            'name': '압축변형률-축력 곡선 (콘크리트 압축능력)',
            'name_font': {'name': '맑은 고딕', 'size': 14, 'bold': True}
        })
        
        chart.set_x_axis({
            'name': '압축변형률 εc',
            'name_font': {'name': '맑은 고딕', 'size': 11, 'bold': True},
            'major_gridlines': {'visible': True},
            'num_format': '0.0000'
        })
        
        chart.set_y_axis({
            'name': '축력 φPn [kN]',
            'name_font': {'name': '맑은 고딕', 'size': 11, 'bold': True},
            'major_gridlines': {'visible': True},
            'num_format': '#,##0'
        })
        
        chart.set_legend({
            'position': 'top',
            'font': {'name': '맑은 고딕', 'size': 9}
        })
        
        chart_ws.insert_chart(62, 15, chart, {
            'x_scale': 1.2, 'y_scale': 1.5
        })
        
        return chart
    
    # ─── 6. 강도감소계수 변화 곡선 (φ factor) ─────────────────
    def create_strength_reduction_factor_chart():
        """편심거리에 따른 강도감소계수 변화"""
        chart = wb.add_chart({'type': 'scatter', 'subtype': 'smooth_with_markers'})
        
        # 이형철근 φ 값
        chart.add_series({
            'name': '이형철근 φ',
            'categories': ['데이터', 2, 0, 100, 0],    # e 열
            'values': ['데이터', 2, 4, 100, 4],        # φ 열
            'line': {'color': chart_colors['ratio'], 'width': 3},
            'marker': {'type': 'circle', 'size': 5, 'fill': {'color': chart_colors['ratio']}}
        })
        
        # 중공철근 φ 값
        chart.add_series({
            'name': '중공철근 φ',
            'categories': ['데이터', 2, 13, 100, 13],  # e 열 (중공철근)
            'values': ['데이터', 2, 17, 100, 17],      # φ 열 (중공철근)
            'line': {'color': chart_colors['hollow'], 'width': 3},
            'marker': {'type': 'square', 'size': 5, 'fill': {'color': chart_colors['hollow']}}
        })
        
        chart.set_title({
            'name': '편심거리에 따른 강도감소계수 (φ) 변화',
            'name_font': {'name': '맑은 고딕', 'size': 14, 'bold': True}
        })
        
        chart.set_x_axis({
            'name': '편심거리 e [mm]',
            'name_font': {'name': '맑은 고딕', 'size': 11, 'bold': True},
            'major_gridlines': {'visible': True},
            'num_format': '#,##0'
        })
        
        chart.set_y_axis({
            'name': '강도감소계수 φ',
            'name_font': {'name': '맑은 고딕', 'size': 11, 'bold': True},
            'major_gridlines': {'visible': True},
            'num_format': '0.000',
            'min': 0.6,
            'max': 0.9
        })
        
        chart.set_legend({
            'position': 'top',
            'font': {'name': '맑은 고딕', 'size': 9}
        })
        
        chart_ws.insert_chart(80, 1, chart, {
            'x_scale': 1.2, 'y_scale': 1.5
        })
        
        return chart
    
    # ─── 7. 용량비 비교 곡선 ─────────────────
    def create_capacity_ratio_chart():
        """이형철근 vs 중공철근 성능 효율성 비교"""
        chart = wb.add_chart({'type': 'scatter', 'subtype': 'smooth_with_markers'})
        
        # 축력 성능 비교 (이형철근)
        chart.add_series({
            'name': '이형철근 축력 φPn',
            'categories': ['데이터', 2, 0, 50, 0],     # e 열 (적당한 범위)
            'values': ['데이터', 2, 5, 50, 5],         # φPn 열
            'line': {'color': chart_colors['rebar'], 'width': 3},
            'marker': {'type': 'circle', 'size': 6, 'fill': {'color': chart_colors['rebar']}}
        })
        
        # 축력 성능 비교 (중공철근)
        chart.add_series({
            'name': '중공철근 축력 φPn',
            'categories': ['데이터', 2, 13, 50, 13],   # e 열 (중공철근)
            'values': ['데이터', 2, 18, 50, 18],       # φPn 열 (중공철근)
            'line': {'color': chart_colors['hollow'], 'width': 3},
            'marker': {'type': 'square', 'size': 6, 'fill': {'color': chart_colors['hollow']}}
        })
        
        # 모멘트 성능 비교 (이형철근)
        chart.add_series({
            'name': '이형철근 모멘트 φMn',
            'categories': ['데이터', 2, 0, 50, 0],     # e 열
            'values': ['데이터', 2, 6, 50, 6],         # φMn 열
            'line': {'color': chart_colors['rebar'], 'width': 2, 'dash_type': 'dash'},
            'marker': {'type': 'none'}
        })
        
        # 모멘트 성능 비교 (중공철근)
        chart.add_series({
            'name': '중공철근 모멘트 φMn',
            'categories': ['데이터', 2, 13, 50, 13],   # e 열 (중공철근)
            'values': ['데이터', 2, 19, 50, 19],       # φMn 열 (중공철근)
            'line': {'color': chart_colors['hollow'], 'width': 2, 'dash_type': 'dash'},
            'marker': {'type': 'none'}
        })
        
        chart.set_title({
            'name': '성능 비교 곡선 (이형철근 vs 중공철근)',
            'name_font': {'name': '맑은 고딕', 'size': 14, 'bold': True}
        })
        
        chart.set_x_axis({
            'name': '편심거리 e [mm]',
            'name_font': {'name': '맑은 고딕', 'size': 11, 'bold': True},
            'major_gridlines': {'visible': True},
            'num_format': '#,##0'
        })
        
        chart.set_y_axis({
            'name': '설계강도 [kN 또는 kN·m]',
            'name_font': {'name': '맑은 고딕', 'size': 11, 'bold': True},
            'major_gridlines': {'visible': True},
            'num_format': '#,##0'
        })
        
        chart.set_legend({
            'position': 'top',
            'font': {'name': '맑은 고딕', 'size': 9}
        })
        
        chart_ws.insert_chart(80, 15, chart, {
            'x_scale': 1.2, 'y_scale': 1.5
        })
        
        return chart
    
    # ─── 8. 차트 생성 실행 ─────────────────
    
    # 제목 추가
    title_format = wb.add_format({
        'bold': True, 'font_size': 16, 'bg_color': '#34495E', 'font_color': 'white',
        'align': 'center', 'valign': 'vcenter', 'border': 2
    })
    
    chart_ws.merge_range(42, 0, 42, 35, '📈 공학적 성능 분석 곡선 (실무 설계용)', title_format)
    chart_ws.set_row(42, 25)
    
    # 각 차트 생성
    create_eccentricity_axial_limit_chart()          # e-φPn: 축력 허용한계
    create_neutral_axis_failure_mode_chart()         # c-φMn: 파괴양식 구분
    create_tension_strain_stress_chart()             # εt-ft: 균열/피로 검토
    create_compression_strain_axial_chart()          # εc-φPn: 압축능력 분석
    create_strength_reduction_factor_chart()         # φ: 강도감소계수 변화
    create_capacity_ratio_chart()                    # 성능비교: 이형철근 vs 중공철근
    
    # ─── 9. 공학적 분석 설명 텍스트 ─────────────────
    note_format = wb.add_format({
        'font_size': 10, 'align': 'left', 'valign': 'top', 'border': 1,
        'bg_color': '#F8F9FA', 'text_wrap': True
    })
    
    analysis_text = ("🔧 공학적 활용 포인트:\n"
                    "• e-φPn 곡선: 현장 편심 발생시 즉시 안전성 판단용\n"
                    "• c-φMn 곡선: 압축↔인장 파괴모드 전이점 파악\n"
                    "• εt-ft 선도: 균열제어·피로검토 (fy=400MPa, εt=0.004 기준)\n"
                    "• εc-φPn 곡선: 콘크리트 압축능력 한계 (εcu=0.003 기준)\n"
                    "• φ 변화: 편심증가→강도감소계수 상승 추세 확인\n"
                    "• 성능비교: 이형철근 vs 중공철근 직접 성능 대비")
    
    chart_ws.merge_range(102, 1, 107, 20, analysis_text, note_format)
    
    return chart_ws