def create_chart_excel(wb, In, R_data, F_data):
    """P-M 상관도 차트 생성 (데이터 시트+설계점 테이블 활용)"""
    
    # ─── 1. 차트 시트 및 스타일 정의 ─────────────────
    chart_ws = wb.add_worksheet('P-M 상관도')
    # 설계점 색상·심볼
    design_point_colors = ['#E74C3C','#F39C12','#27AE60','#8E44AD','#3498DB','#E67E22']
    design_point_symbols = ["①","②","③","④","⑤","⑥"]
    # 설계점 테이블 시작 위치
    dp_start_row, dp_start_col = 32, 16

    # 시트 제목
    title_fmt = wb.add_format({
        'bold': True,'font_size': 20,'bg_color': '#1F4E79','font_color': 'white',
        'align':'center','valign':'vcenter','border':2
    })
    chart_ws.merge_range(0, 0, 0, 28, '📊 P-M 상관도 (설계점 포함)', title_fmt)
    chart_ws.set_row(0, 35)

    # ─── 2. 차트 생성 내부 함수 ─────────────────
    def draw_pm(title, col_offset, insert_pos):
        chart = wb.add_chart({'type':'scatter','subtype':'smooth'})
        # Pn-Mn 곡선
        chart.add_series({
            'name': 'Pn-Mn Diagram',
            'categories': ['P-M 데이터', 2, col_offset+3, 100, col_offset+3],
            'values':     ['P-M 데이터', 2, col_offset+2, 100, col_offset+2],
            'line': {'color':'#E74C3C','width':3,'dash_type':'dash'},
            'marker': {'type':'none'}
        })
        # ϕPn-ϕMn 곡선
        chart.add_series({
            'name': 'ϕPn-ϕMn Diagram',
            'categories': ['P-M 데이터', 2, col_offset+6, 100, col_offset+6],
            'values':     ['P-M 데이터', 2, col_offset+5, 100, col_offset+5],
            'line': {'color':'#3498DB','width':4},
            'marker': {'type':'none'},
            'fill': {'color':'#E8F6F3','transparency':70}
        })
        # 설계점 포인트
        for i in range(min(len(In.Pu), len(design_point_colors))):
            r = dp_start_row + 2 + i
            chart.add_series({
                'name': f'Case {i+1} ({design_point_symbols[i]})',
                'categories': ['P-M 상관도', r, dp_start_col+2, r, dp_start_col+2],
                'values':     ['P-M 상관도', r, dp_start_col+1, r, dp_start_col+1],
                'line': {'none': True},
                'marker': {
                    'type': 'circle',
                    'size': 12,
                    'border': {'color': design_point_colors[i], 'width': 3},
                    'fill': {'color': design_point_colors[i], 'transparency':20}
                }
            })
        # 차트 설정
        chart.set_title({
            'name': f'{title} P-M 상관도',
            'name_font':{'name':'맑은 고딕','size':16,'bold':True}
        })
        chart.set_x_axis({
            'name':'Mn or ϕMn [kN·m]',
            'name_font':{'name':'맑은 고딕','size':12,'bold':True},
            'major_gridlines':{'visible':True},
            'min':0,'num_font':{'name':'맑은 고딕','size':12,'bold':True},
            'num_format':'0'
        })
        chart.set_y_axis({
            'name':'Pn or ϕPn [kN]',
            'name_font':{'name':'맑은 고딕','size':12,'bold':True},
            'major_gridlines':{'visible':True},
            'num_font':{'name':'맑은 고딕','size':12,'bold':True},
            'num_format':'0'
        })
        chart.set_legend({'position':'top','font':{'name':'맑은 고딕','size':10,'bold':True}})
        chart_ws.insert_chart(insert_pos[0], insert_pos[1], chart,
                              {'x_scale':1.5,'y_scale':2.0})

    # ─── 3. 차트 배치 ─────────────────
    draw_pm('이형철근', 0,  (2, 1))
    draw_pm('중공철근', 13, (2,15))

    # ─── 4. 설계점 테이블 생성 ─────────────────
    header_fmt = wb.add_format({
        'bold':True,'font_size':12,'bg_color':'#34495E','font_color':'white',
        'align':'center','valign':'vcenter','border':1
    })
    cell_fmt = wb.add_format({
        'font_size':11,'align':'center','valign':'vcenter','border':1,
        'num_format':'#,##0.0'
    })
    text_fmt = wb.add_format({
        'font_size':11,'align':'center','valign':'vcenter','border':1
    })
    title_fmt2 = wb.add_format({
        'bold':True,'font_size':14,'bg_color':'#2C3E50','font_color':'white',
        'align':'center','valign':'vcenter','border':2
    })
    # 테이블 제목
    chart_ws.merge_range(dp_start_row, dp_start_col, dp_start_row, dp_start_col+3,
                         '📋 설계점 정보', title_fmt2)
    chart_ws.set_row(dp_start_row, 25)
    # 헤더
    for j, h in enumerate(['Case','Pu (kN)','Mu (kN·m)','마커']):
        chart_ws.write(dp_start_row+1, dp_start_col+j, h, header_fmt)
    chart_ws.set_row(dp_start_row+1, 20)
    # 데이터
    for i in range(len(In.Pu)):
        r = dp_start_row + 2 + i
        chart_ws.write(r, dp_start_col,   f'Case {i+1}', text_fmt)
        chart_ws.write(r, dp_start_col+1, In.Pu[i],       cell_fmt)
        chart_ws.write(r, dp_start_col+2, In.Mu[i],       cell_fmt)
        mf = wb.add_format({
            'font_size':14,'align':'center','valign':'vcenter','border':1,
            'bg_color':design_point_colors[i],'font_color':'white','bold':True
        })
        chart_ws.write(r, dp_start_col+3, design_point_symbols[i], mf)
        chart_ws.set_row(r,18)
    # 컬럼 너비
    chart_ws.set_column(dp_start_col, dp_start_col,   10)
    chart_ws.set_column(dp_start_col+1, dp_start_col+1, 12)
    chart_ws.set_column(dp_start_col+2, dp_start_col+2, 14)
    chart_ws.set_column(dp_start_col+3, dp_start_col+3,  8)

    # ─── 5. 설명 텍스트 병합 ─────────────────
    note_fmt = wb.add_format({
        'font_size':11,'align':'left','valign':'vcenter','border':1,
        'bg_color':'#F8F9FA','text_wrap':True
    })
    note = (
        "• 빨간 점선: 공칭강도 Pn-Mn 곡선\n"
        "• 파란 실선: 설계강도 ϕPn-ϕMn 곡선 (안전영역)\n"
        "• 원형 마커: 실제 설계점 (Pu, Mu)\n"
        "• 설계점이 파란 영역 내부에 있으면 안전합니다."
    )
    chart_ws.merge_range(dp_start_row+1, 1,
                        dp_start_row+5, 13, note, note_fmt)

    return chart_ws
