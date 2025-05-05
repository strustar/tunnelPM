def display_data_with_aggrid(headers, PM, height, width=840, column_widths=None, decimal_places=None):
    from st_aggrid import GridOptionsBuilder, AgGrid, JsCode
    import pandas as pd
    import numpy as np
    import streamlit as st

    # 4. 셀 클래스 함수 - 음수 검출용
    cell_class_js = JsCode(
        """
    function(params) {
        if (params.value && params.value.startsWith('-')) {
            return 'negative-value';
        }
        return '';
    }
    """
    )

    epsilon = 0.1  # 허용 오차
    # 조건 생성: 모든 PM.e[i] 값 기준 비교문 생성
    conditions = "\n".join(
        [
            f"""if (Math.abs(e_val - {round(v, 1)}) < {epsilon}) {{
            return {{
                backgroundColor: 'rgba(255, 0, 0, 0.3)'
            }};
        }}"""
            for v in PM.Pn
        ]
    )

    row_style_js = JsCode(
        f"""
    function(params) {{    
        var phi = Number(params.data['ϕ']);
        var e_val = Number(params.data['Pₙ [kN]'].replace(/,/g, ''));


        // 조건1: ϕ가 특정 범위일 때
        if (phi > 0.65 && phi < 0.85) {{
            return {{
                backgroundColor: 'rgba(0, 255, 0, 0.3)',
                color: 'black',
                fontWeight: 'bold'
            }};
        }}

        // 조건2: e_val 비교 (여러 개)
        {conditions}
        return null;
        }}
    """
    )

    phi_cell_style = JsCode(
        """
    function(params) {
    var v = Number(params.value);
    if (v > 0.65 && v < 0.85) {
        return {'color':'green', 'fontWeight':'bold'};
    }
    return null;
    }
    """
    )

    # 1. 데이터 준비 (이전과 동일)
    rows = []
    for n in range(len(PM.Ze)):
        rows.append(
            [
                PM.Ze[n],
                PM.Zc[n],
                PM.ZPn[n],
                PM.ZMn[n],
                PM.Zphi[n],
                PM.ZPd[n],
                PM.ZMd[n],
                PM.Zep_s[n, 1],
                PM.Zfs[n, 1],
                PM.Zep_s[n, 0],
                PM.Zfs[n, 0],
            ]
        )

    # 2. NumPy 배열로 변환
    arr = np.array(rows)

    # 3. Zc 기준으로 정렬 (Zc는 두 번째 열 → index 1)
    arr_sorted = arr[arr[:, 1].argsort()[::-1]]  # 내림차순 정렬

    # 4. 중복 제거 (Zc 기준으로 첫 등장만 남김)
    _, unique_indices = np.unique(arr_sorted[:, 1], return_index=True)
    data = arr_sorted[np.sort(unique_indices)]

    columns = [
        'e [mm]',
        'c [mm]',
        'Pₙ [kN]',
        'Mₙ [kN·m]',
        'ϕ',
        'ϕPₙ [kN]',
        'ϕMₙ [kN·m]',
        'εₜ',
        'fₜ [MPa]',
        'εc',
        'fc [MPa]',
    ]

    if decimal_places is None:
        decimal_places = {
            'e [mm]': 1,
            'c [mm]': 1,
            'Pₙ [kN]': 1,
            'Mₙ [kN·m]': 1,
            'ϕ': 3,
            'ϕPₙ [kN]': 1,
            'ϕMₙ [kN·m]': 1,
            'εₜ': 4,
            'fₜ [MPa]': 1,
            'εc': 4,
            'fc [MPa]': 1,
        }

    # 2. DataFrame 생성 - 문자열로 미리 변환하여 AgGrid가 직접 렌더링하지 않도록 함
    formatted_data = []
    for row in data:
        formatted_row = []
        for i, val in enumerate(row):
            col_name = columns[i] if i < len(columns) else ''
            precision = decimal_places.get(col_name, 1) if decimal_places else 1

            # 값 포맷팅 (파이썬에서 직접 처리)
            if val is None or pd.isna(val):
                formatted_row.append('')
            elif val == float('inf'):
                formatted_row.append('∞')
            elif val == float('-inf'):
                formatted_row.append('-∞')
            elif isinstance(val, (int, float)):
                formatted_row.append(f"{val:,.{precision}f}")
            else:
                formatted_row.append(str(val))
        formatted_data.append(formatted_row)

    # 포맷팅된 문자열 데이터로 DataFrame 생성
    df = pd.DataFrame(formatted_data, columns=columns)

    # 3. GridOptionsBuilder 설정
    gb = GridOptionsBuilder.from_dataframe(df)

    gb.configure_column(
        'ϕ',
        headerName='ϕ',  # 이미 설정된 헤더
        cellStyle=phi_cell_style,  # 여기에 함수 지정
        precision=3,
        width=65,
        suppressMenu=True,
        wrapHeaderText=True,
        autoHeaderHeight=True,
    )

    # 5. 컬럼 설정 - 각 컬럼마다 개별 설정
    for col in columns:
        # 컬럼 너비 결정 (기본값 설정)
        width_value = None
        if column_widths:
            width_value = column_widths.get(col)
        else:
            # 기본 너비 설정
            if 'e' in col:
                width_value = 90
            elif 'kN' in col:
                width_value = 85
            elif 'kN·m' in col:
                width_value = 90
            elif col == 'ϕ':
                width_value = 65
            elif 'ε' in col:
                width_value = 80
            elif 'MPa' in col:
                width_value = 70
            else:
                width_value = 65

        # 헤더 텍스트 설정 - 줄바꿈 처리
        header_text = col.replace(' [', '\n[')

        # 컬럼 설정
        gb.configure_column(
            col,
            headerName=header_text,
            width=width_value,
            maxWidth=width_value,  # 추가
            minWidth=width_value,  # 추가
            suppressMenu=True,
            wrapHeaderText=True,
            autoHeaderHeight=True,
            cellClass=cell_class_js,
        )

    # 6. 그리드 옵션 설정
    gb.configure_grid_options(
        domLayout='normal',
        rowHeight=40,
        headerHeight=60,
        suppressRowTransform=True,
        enableRangeSelection=True,
        # enableExcelExport=True,
        # enableCsvExport=False,  # csv 비활성화
        # suppressExcelExport=False,
    )

    # 7. 풀 데이터 모드인 경우 페이지네이션 적용
    if hasattr(PM, 'Ze') and len(PM.Ze) > 10:
        gb.configure_grid_options(pagination=True, paginationAutoPageSize=False, paginationPageSize=100)

    # 8. 그리드 옵션 빌드
    gridOptions = gb.build()
    gridOptions['getRowStyle'] = row_style_js
    # # row_style_js 함수 정의 후
    # gb.configure_grid_options(
    #     getRowStyle=row_style_js,
    #     # 다른 옵션들...
    # )

    # 9. AgGrid 출력
    return AgGrid(
        df,
        gridOptions=gridOptions,
        theme="streamlit",
        fit_columns_on_grid_load=False,
        update_mode="value_changed",
        allow_unsafe_jscode=True,
        enable_enterprise_modules=True,  # ⚠️ 필수: export 기능 활성화
        height=height,
        width=width,
        custom_css={
            ".ag-header-cell-menu-button": {"display": "none !important"},  # 점 세 개 숨김
            ".ag-header-cell": {
                "border": "1px solid gray",
                "background-color": "rgba(0, 0, 255, 0.2) !important",
                "white-space": "normal",
                "line-height": "1.8em",
            },
            ".ag-header-cell-label": {
                "justify-content": "center",
                "align-items": "center",
                "text-align": "center",
                "color": "#1976d2",
                "font-weight": "bold",
                "font-size": "16px",
            },
            ".ag-row:hover .ag-cell": {"background-color": "rgba(0,255,0,0.7) !important"},
            ".ag-cell": {
                "border": "1px solid gray",
                # "background-color": "rgba(245, 245, 245, 0.9) !important",
                "font-size": "16px",
                "text-align": "right",
                "padding-right": "10px",
                "font-weight": "bold",
                "display": "flex",  # 추가: 플렉스 박스 사용
                "align-items": "center",  # 추가: 세로 중앙 정렬
                "justify-content": "flex-end",  # 추가: 오른쪽 정렬 유지
            },
            ".negative-value": {"color": "red !important", "font-weight": "bold"},
        },
        key=f"pm_table_{id(PM)}",
    )
