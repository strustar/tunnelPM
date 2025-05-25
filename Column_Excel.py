def Excel(In, R, F):
    import os, time#, subprocess#, pythoncom
    # from win32com.client import GetActiveObject, DispatchEx
    import pandas as pd
    import numpy as np

    path = os.path.abspath("a.xlsx")

    # # ─── 1) 기존 엑셀 닫기 & 프로세스 종료 ─────────────
    # pythoncom.CoInitialize()
    # try: 
    #     excel = GetActiveObject("Excel.Application")
    # except: 
    #     excel = DispatchEx("Excel.Application")
    # excel.Visible = False
    # for wb in list(excel.Workbooks):
    #     if os.path.normcase(wb.FullName) == os.path.normcase(path):
    #         wb.Close(SaveChanges=False)
    #         break
    # excel.Quit()
    # pythoncom.CoUninitialize()
    # subprocess.call(["taskkill", "/F", "/IM", "EXCEL.EXE"],
    #                 stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    # time.sleep(1)

    # ─── 2) 데이터 처리 함수 ────────────────────────────
    def process_data(PM):
        df = pd.DataFrame({
            'e\n[mm]':     PM.Ze,
            'c\n[mm]':     PM.Zc,
            'Pₙ\n[kN]':    PM.ZPn,
            'Mₙ\n[kN·m]':  PM.ZMn,
            'ϕ':           PM.Zphi,
            'ϕPₙ\n[kN]':   PM.ZPd,
            'ϕMₙ\n[kN·m]': PM.ZMd,
            'εₜ':          PM.Zep_s[:,1],
            'fₜ\n[MPa]':   PM.Zfs[:,1],
            'εc':          PM.Zep_s[:,0],
            'fc\n[MPa]':   PM.Zfs[:,0],
        })
        # 정렬·중복 제거
        df = (df.sort_values('c\n[mm]', ascending=False)
                .drop_duplicates('c\n[mm]')
                .reset_index(drop=True))
        # Inf 표시
        df = df.replace(np.inf, '∞').replace(-np.inf, '-∞')
        return df

    # ─── 3) 포맷 적용 함수 ────────────────────────────
    def apply_formatting(wb, ws, df, start_col, PM, title):
        # 컬럼별 소수 자리
        dec_pl = {
            'e\n[mm]': 1, 'c\n[mm]': 1, 'Pₙ\n[kN]': 1, 'Mₙ\n[kN·m]': 1,
            'ϕ': 3, 'ϕPₙ\n[kN]': 1, 'ϕMₙ\n[kN·m]': 1,
            'εₜ': 4, 'fₜ\n[MPa]': 1, 'εc': 4, 'fc\n[MPa]': 1
        }

        # 포맷 정의
        fmt = {}
        fmt['title'] = wb.add_format({
            'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter',
            'bg_color': '#4472C4', 'font_color': 'white', 'border': 2
        })
        fmt['header'] = wb.add_format({
            'bold': True, 'font_color': '#4472C4', 'bg_color': '#DCE6F1',
            'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 12,
            'text_wrap': True  # 줄바꿈 활성화
        })
        
        for col, dp in dec_pl.items():
            fmt[col] = wb.add_format({
                'num_format': '#,##0.' + '0'*dp,
                'align': 'right', 'valign': 'vcenter', 'border': 1, 'font_size': 12, 'bold': True
            })
        
        fmt['negative'] = wb.add_format({'font_color': 'red'})
        fmt['light_red'] = wb.add_format({'bg_color': '#F8CBAD', 'border': 1})
        fmt['light_green'] = wb.add_format({'bg_color': '#D9EAD3', 'border': 1})
        fmt['dark_green'] = wb.add_format({
            'bg_color': '#D9EAD3', 'font_color': 'blue', 'border': 1, 'bold': True
        })

        # 제목 쓰기
        ws.merge_range(0, start_col, 0, start_col + len(df.columns) - 1, title, fmt['title'])
        
        # 헤더 쓰기 & 컬럼 너비 설정
        for i, col in enumerate(df.columns):
            col_idx = start_col + i
            ws.write(1, col_idx, col, fmt['header'])
            ws.set_column(col_idx, col_idx, 10, fmt[col])

        # 헤더 행 높이 설정 (2줄을 위해)
        ws.set_row(1, 33)

        # 조건부 서식 적용 (데이터는 2행부터 시작)
        data_start_row = 2
        nrows = len(df) + data_start_row - 1  # 데이터 마지막 행
        end_col = start_col + len(df.columns) - 1
        
        # (1) 음수 빨간 글씨
        for i, col in enumerate(df.columns):
            col_idx = start_col + i
            ws.conditional_format(data_start_row, col_idx, nrows, col_idx, {
                'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt['negative']
            })
        
        # (2) φ 범위행 연녹색(행 전체), φ 셀 진녹색
        phi_col_idx = start_col + df.columns.get_loc('ϕ')
        phi_col_letter = chr(65 + phi_col_idx)
        
        # φ 값이 0.65~0.85 범위인 행 전체를 연녹색
        ws.conditional_format(data_start_row, start_col, nrows, end_col, {
            'type': 'formula',
            'criteria': f'=AND(${phi_col_letter}3>0.65,${phi_col_letter}3<0.85)',
            'format': fmt['light_green']
        })
        
        # φ 셀만 진녹색
        ws.conditional_format(data_start_row, phi_col_idx, nrows, phi_col_idx, {
            'type': 'formula', 
            'criteria': f'=AND(${phi_col_letter}3>0.65,${phi_col_letter}3<0.85)',
            'format': fmt['dark_green']
        })
        
        # (3) special Pₙ 값 행 연붉은 배경
        Pn_col_idx = start_col + df.columns.get_loc('Pₙ\n[kN]')
        Pn_col_letter = chr(65 + Pn_col_idx)
        
        tol = 1e-3  # 오차 허용치
        for v in PM.Pn:
            formula = f'=ABS(${Pn_col_letter}3-{v})<{tol}'
            ws.conditional_format(data_start_row, start_col, nrows, end_col, {
                'type': 'formula',
                'criteria': formula,
                'format': fmt['light_red']
            })

    # ─── 4) 두 개 데이터셋 처리 ────────────────────────────
    data_sets = [
        (R, '이형철근'),
        (F, '중공철근')
    ]

    with pd.ExcelWriter(path, engine='xlsxwriter',
            engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:

        for idx, (PM, title) in enumerate(data_sets):
            df = process_data(PM)
            # if idx == 0:
            #     R_df = df
            # else:
            #     F_df = df
            
            # 좌우 배치를 위한 시작 컬럼 계산 (2칸 간격)
            start_col = idx * (len(df.columns) + 2)
            
            # 데이터 쓰기 (제목과 헤더를 위해 startrow=2)
            df.to_excel(writer, sheet_name='데이터', index=False, 
                        startrow=2, startcol=start_col, header=False)
            
            wb = writer.book
            if '데이터' not in writer.sheets:
                ws = wb.add_worksheet('데이터')
                writer.sheets['데이터'] = ws
            else:
                ws = writer.sheets['데이터']
            
            # 포맷팅 및 조건부 서식 적용
            apply_formatting(wb, ws, df, start_col, PM, title)

        import Column_Review
        Column_Review.create_review_sheet(wb, In, R, F)
        import Column_Chart
        Column_Chart.create_pm_chart_excel(wb, In, R, F)

    # ─── 5) 엑셀 파일 열기 ────────────────────────────
    # os.startfile(path)
    import streamlit as st
    with open(path, "rb") as file:
        st.download_button(
            label="📥 엑셀 파일 다운로드",
            data=file,
            file_name=path.split("/")[-1],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    