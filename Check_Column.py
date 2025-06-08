import streamlit as st
import pandas as pd
import numpy as np

def check_column(In, R, F):
    """
    In, R, F 객체의 실제 데이터를 사용하여 
    스트림릿에서 실시간으로 계산되는 기둥 강도 검토 보고서 생성
    """

    # =================================================================
    # 스타일리시한 보고서 스타일 (글자 크기 2포인트 증가)
    # =================================================================
    st.markdown("""
    <style>
        /* 메인 컨테이너 */
        .main-container {
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            padding: 30px;
            border-radius: 15px;
            margin: 20px 0;
        }
        
        /* 메인 헤더 */
        .main-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            text-align: center;
            font-weight: bold;
            font-size: 2.7em;  /* 2.5em → 2.7em */
            padding: 30px;
            border-radius: 15px;
            margin-bottom: 35px;
            box-shadow: 0 10px 30px rgba(102, 126, 234, 0.4);
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }
        
        /* 공통 조건 컨테이너 */
        .common-conditions {
            background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
            color: white;
            padding: 25px;
            border-radius: 12px;
            margin-bottom: 30px;
            box-shadow: 0 8px 25px rgba(17, 153, 142, 0.3);
        }
        
        .common-header {
            text-align: center;
            font-weight: bold;
            font-size: 1.7em;  /* 1.5em → 1.7em */
            margin-bottom: 25px;
            text-shadow: 1px 1px 2px rgba(0,0,0,0.2);
        }
        
        /* 리포트 컨테이너 */
        .report-container {
            background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
            border: 2px solid #e9ecef;
            border-radius: 15px;
            padding: 30px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            height: 100%;
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }
        
        .report-container:hover {
            transform: translateY(-5px);
            box-shadow: 0 15px 40px rgba(0,0,0,0.15);
        }
        
        /* 섹션 헤더 */
        .section-header {
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
            color: white;
            padding: 18px 25px;  /* 15px → 18px, 20px → 25px */
            border-radius: 12px;
            margin-top: 25px;
            margin-bottom: 20px;
            font-size: 1.5em;  /* 1.3em → 1.5em */
            font-weight: bold;
            text-align: center;
            box-shadow: 0 5px 15px rgba(79, 172, 254, 0.3);
            text-shadow: 1px 1px 2px rgba(0,0,0,0.2);
        }
        
        /* 서브 섹션 헤더 */
        .sub-section-header {
            background: linear-gradient(135deg, #fa709a 0%, #fee140 100%);
            color: #2c3e50;
            padding: 12px 20px;  /* 10px → 12px, 15px → 20px */
            border-left: 6px solid #e74c3c;  /* 5px → 6px */
            margin-top: 20px;
            margin-bottom: 15px;
            font-weight: 700;
            font-size: 1.2em;  /* 1.0em → 1.2em */
            border-radius: 8px;
            box-shadow: 0 3px 10px rgba(250, 112, 154, 0.2);
        }
        
        /* 공통 조건 테이블 */
        .common-table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }
        
        .common-table td {
            padding: 12px 15px;  /* 10px → 12px */
            border: 1px solid #dee2e6;
            font-size: 1.1em;  /* 0.95em → 1.1em */
        }
        
        .common-table td:first-child {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            font-weight: 600;
            width: 50%;
        }
        
        .common-table td:last-child {
            text-align: right;
            font-weight: bold;
            color: #2c3e50;
            background-color: #f8f9fa;
        }
        
        /* 파라미터 테이블 */
        .param-table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 8px;
            overflow: hidden;
            margin-bottom: 20px;
            box-shadow: 0 3px 10px rgba(0,0,0,0.08);
        }
        
        .param-table td {
            padding: 12px 15px;  /* 9px → 12px */
            border: 1px solid #dee2e6;
            font-size: 1.1em;  /* 1em → 1.1em */
        }
        
        .param-table td:nth-child(1) {
            background: linear-gradient(135deg, #ffecd2 0%, #fcb69f 100%);
            font-weight: 600;
            width: 60%;  /* 55% → 60% */
            color: #8b4513;
        }
        
        .param-table td:nth-child(2) {
            text-align: right;
            font-weight: bold;
            color: #2c3e50;
            background-color: #ffffff;
        }
        
        /* 결과 테이블 */
        .results-table {
            width: 100%;
            border-collapse: collapse;
            text-align: center;
            background: white;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }
        
        .results-table th {
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
            color: white;
            padding: 15px 10px;  /* 12px → 15px */
            font-size: 1.0em;  /* 0.9em → 1.0em */
            font-weight: 600;
            text-shadow: 1px 1px 2px rgba(0,0,0,0.3);
        }
        
        .results-table td {
            padding: 12px 10px;  /* 10px → 12px */
            border: 1px solid #dee2e6;
            vertical-align: middle;
            font-size: 1.0em;  /* 0.9em → 1.0em */
        }
        
        .results-table .pass {
            background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
            color: #155724;
            font-weight: bold;
            box-shadow: inset 0 2px 4px rgba(21, 87, 36, 0.1);
        }
        
        .results-table .fail {
            background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%);
            color: #721c24;
            font-weight: bold;
            box-shadow: inset 0 2px 4px rgba(114, 28, 36, 0.1);
        }
        
        /* 최종 판정 */
        .final-verdict-container {
            padding: 20px;  /* 15px → 20px */
            border-radius: 12px;
            text-align: center;
            font-size: 1.3em;  /* 1.1em → 1.3em */
            font-weight: bold;
            margin-top: 20px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }
        
        .final-pass {
            background: linear-gradient(135deg, #d4edda 0%, #a3d977 100%);
            color: #1e8449;
            border: 2px solid #28a745;
        }
        
        .final-fail {
            background: linear-gradient(135deg, #fadbd8 0%, #ff6b6b 100%);
            color: #a93226;
            border: 2px solid #dc3545;
        }
        
        /* 아이콘 스타일 */
        .icon {
            font-size: 1.3em;  /* 1.1em → 1.3em */
            margin-right: 8px;
        }
        
        /* 반응형 */
        @media (max-width: 768px) {
            .main-header {
                font-size: 2.2em;  /* 모바일에서도 크게 */
            }
            .section-header {
                font-size: 1.3em;
            }
        }
    </style>
    """, unsafe_allow_html=True)

    # =================================================================
    # 데이터 처리 함수들
    # =================================================================
    def extract_pm_data(PM_obj, material_type, In):
        """R 또는 F 객체에서 PM 상관도 데이터 추출 (Excel 코드 기반)"""
        try:
            def safe_extract(attr_name, default=[]):
                values = getattr(PM_obj, attr_name, default)
                if hasattr(values, 'tolist'):
                    return values.tolist()
                elif isinstance(values, (list, tuple)):
                    return list(values)
                else:
                    return [float(values)] if values is not None else []
            
            pm_data = {
                'e_mm': safe_extract('Ze'),
                'c_mm': safe_extract('Zc'),
                'Pn_kN': safe_extract('ZPn'),
                'Mn_kNm': safe_extract('ZMn'),
                'phi': safe_extract('Zphi'),
                'phiPn_kN': safe_extract('ZPd'),
                'phiMn_kNm': safe_extract('ZMd')
            }
            
            # 평형점 데이터 추출 (Excel 코드 참조: index [3] 사용)
            try:
                Pd_values = safe_extract('Pd')
                Md_values = safe_extract('Md')
                e_values = safe_extract('e')
                c_values = safe_extract('c')
                
                balanced_data = {
                    'Pb_kN': float(Pd_values[3]) if len(Pd_values) > 3 else 0.0,
                    'Mb_kNm': float(Md_values[3]) if len(Md_values) > 3 else 0.0,
                    'eb_mm': float(e_values[3]) if len(e_values) > 3 else 0.0,
                    'cb_mm': float(c_values[3]) if len(c_values) > 3 else 0.0
                }
            except (IndexError, ValueError, TypeError):
                balanced_data = {'Pb_kN': 0.0, 'Mb_kNm': 0.0, 'eb_mm': 0.0, 'cb_mm': 0.0}
            
            # 재료 특성
            if material_type == '이형철근':
                material_props = {
                    'fy': float(getattr(In, 'fy', 400.0)),
                    'Es': float(getattr(In, 'Es', 200000.0)) / 1000
                }
            else:
                material_props = {
                    'fy': float(getattr(In, 'fy_hollow', 800.0)),
                    'Es': float(getattr(In, 'Es_hollow', 200000.0)) / 1000
                }
            
            return pm_data, balanced_data, material_props
            
        except Exception as e:
            st.error(f"데이터 추출 중 오류 발생: {e}")
            return {}, {'Pb_kN': 0.0, 'Mb_kNm': 0.0, 'eb_mm': 0.0, 'cb_mm': 0.0}, {'fy': 0.0, 'Es': 0.0}

    def calculate_strength_check(In, material_type):
        """하중조합별 강도 검토 계산 (Excel 코드 기반)"""
        try:
            Pu_values = getattr(In, 'Pu', [])
            Mu_values = getattr(In, 'Mu', [])
            
            if hasattr(Pu_values, 'tolist'):
                Pu_values = Pu_values.tolist()
            if hasattr(Mu_values, 'tolist'):
                Mu_values = Mu_values.tolist()
            
            if material_type == '이형철근':
                safety_factors = getattr(In, 'safe_RC', [])
                Pd_values = getattr(In, 'Pd_RC', [])
                Md_values = getattr(In, 'Md_RC', [])
            else:
                safety_factors = getattr(In, 'safe_FRP', [])
                Pd_values = getattr(In, 'Pd_FRP', [])
                Md_values = getattr(In, 'Md_FRP', [])
            
            if hasattr(safety_factors, 'tolist'):
                safety_factors = safety_factors.tolist()
            if hasattr(Pd_values, 'tolist'):
                Pd_values = Pd_values.tolist()
            if hasattr(Md_values, 'tolist'):
                Md_values = Md_values.tolist()
            
            if not Pu_values or not Mu_values:
                return []
            
            checks = []
            num_cases = min(len(Pu_values), len(Mu_values))
            
            for i in range(num_cases):
                try:
                    Pu = float(Pu_values[i])
                    Mu = float(Mu_values[i])
                    Pd = float(Pd_values[i]) if i < len(Pd_values) else 0.0
                    Md = float(Md_values[i]) if i < len(Md_values) else 0.0
                    sf = float(safety_factors[i]) if i < len(safety_factors) else 0.0
                    
                    e = (Mu / Pu) * 1000 if Pu != 0 else 0.0
                    verdict = 'PASS' if sf >= 1.0 else 'FAIL'
                    
                    checks.append({
                        'LC': f'LC-{i+1}',
                        'Pu/phiPn': f'{Pu:,.1f} / {Pd:,.1f}',
                        'Mu/phiMn': f'{Mu:,.1f} / {Md:,.1f}',
                        'e_mm': e,
                        'SF': sf,
                        'Verdict': verdict
                    })
                except (ValueError, TypeError, ZeroDivisionError):
                    checks.append({
                        'LC': f'LC-{i+1}',
                        'Pu/phiPn': 'ERROR',
                        'Mu/phiMn': 'ERROR',
                        'e_mm': 0.0,
                        'SF': 0.0,
                        'Verdict': 'FAIL'
                    })
            
            return checks
            
        except Exception as e:
            st.error(f"강도 검토 계산 중 오류 발생: {e}")
            return []

    # =================================================================
    # 공통 설계 조건 렌더링
    # =================================================================
    def render_common_conditions(In):
        """공통 설계 조건 섹션"""
        st.markdown('<div class="common-conditions">', unsafe_allow_html=True)
        st.markdown('<div class="common-header">🏗️ 공통 설계 조건</div>', unsafe_allow_html=True)
        
        col1, col2, col3, col4 = st.columns(4)
        
        # 단면 제원
        with col1:
            be_val = getattr(In, 'be', getattr(In, 'b', 1000))
            h_val = getattr(In, 'height', getattr(In, 'h', 300))
            sb_val = getattr(In, 'sb', [150.0])[0] if hasattr(In, 'sb') and len(getattr(In, 'sb', [])) > 0 else 150.0
            
            st.markdown('''
            <table class="common-table">
                <tr><td colspan="2" style="text-align: center; font-weight: bold; font-size: 1.1em;">📐 단면 제원</td></tr>
                <tr><td><span class="icon">📏</span>단위폭 b<sub>e</sub></td><td>{:,.1f} mm</td></tr>
                <tr><td><span class="icon">📏</span>단면 두께 h</td><td>{:,.1f} mm</td></tr>
                <tr><td><span class="icon">📐</span>공칭 철근간격 s</td><td>{:,.1f} mm</td></tr>
            </table>
            '''.format(be_val, h_val, sb_val), unsafe_allow_html=True)
        
        # 콘크리트 특성
        with col2:
            fck_val = getattr(In, 'fck', 40.0)
            Ec_val = getattr(In, 'Ec', 30000.0) / 1000
            
            st.markdown('''
            <table class="common-table">
                <tr><td colspan="2" style="text-align: center; font-weight: bold; font-size: 1.1em;">🏭 콘크리트 재료</td></tr>
                <tr><td><span class="icon">💪</span>압축강도 f<sub>ck</sub></td><td>{:,.1f} MPa</td></tr>
                <tr><td><span class="icon">⚡</span>탄성계수 E<sub>c</sub></td><td>{:,.1f} GPa</td></tr>
                <tr><td style="opacity: 0;"></td><td style="opacity: 0;"></td></tr>
            </table>
            '''.format(fck_val, Ec_val), unsafe_allow_html=True)
        
        # 설계 조건
        with col3:
            design_method = getattr(In, 'Design_Method', 'USD').split('(')[0].strip()
            rc_code = getattr(In, 'RC_Code', 'KDS-2021')
            column_type = getattr(In, 'Column_Type', 'Tied Column')
            
            st.markdown('''
            <table class="common-table">
                <tr><td colspan="2" style="text-align: center; font-weight: bold; font-size: 1.1em;">📋 설계 조건</td></tr>
                <tr><td><span class="icon">🔧</span>설계방법</td><td>{}</td></tr>
                <tr><td><span class="icon">📖</span>설계기준</td><td>{}</td></tr>
                <tr><td><span class="icon">🏛️</span>기둥형식</td><td>{}</td></tr>
            </table>
            '''.format(design_method, rc_code, column_type), unsafe_allow_html=True)
        
        # 철근 배치
        with col4:
            dia_val = getattr(In, 'dia', [22.0])[0] if hasattr(In, 'dia') and len(getattr(In, 'dia', [])) > 0 else 22.0
            dc_val = getattr(In, 'dc', [60.0])[0] if hasattr(In, 'dc') and len(getattr(In, 'dc', [])) > 0 else 60.0
            effective_depth = h_val - dc_val
            
            st.markdown('''
            <table class="common-table">
                <tr><td colspan="2" style="text-align: center; font-weight: bold; font-size: 1.1em;">🔩 철근 배치</td></tr>
                <tr><td><span class="icon">⭕</span>철근 직경 D</td><td>{:,.1f} mm</td></tr>
                <tr><td><span class="icon">🛡️</span>피복두께 d<sub>c</sub></td><td>{:,.1f} mm</td></tr>
                <tr><td><span class="icon">📊</span>유효깊이 d</td><td>{:,.1f} mm</td></tr>
            </table>
            '''.format(dia_val, dc_val, effective_depth), unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)

    # =================================================================
    # 개별 검토 섹션 렌더링
    # =================================================================
    def create_report_column(column_ui, title, In, PM_obj, material_type):
        """개별 검토 컬럼 생성"""
        with column_ui:
            st.markdown('<div class="report-container">', unsafe_allow_html=True)
            st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)

            pm_data, balanced_data, material_props = extract_pm_data(PM_obj, material_type, In)
            
            # 철근 재료 특성
            st.markdown('<div class="sub-section-header">🔧 철근 재료 특성</div>', unsafe_allow_html=True)
            st.markdown(f'''
            <table class="param-table">
                <tr><td><span class="icon">💪</span>항복강도 f<sub>y</sub></td><td>{material_props['fy']:,.1f} MPa</td></tr>
                <tr><td><span class="icon">⚡</span>탄성계수 E<sub>s</sub></td><td>{material_props['Es']:,.1f} GPa</td></tr>
            </table>''', unsafe_allow_html=True)

            # 평형상태 검토
            st.markdown('<div class="sub-section-header">⚖️ 평형상태(Balanced) 검토</div>', unsafe_allow_html=True)
            st.markdown(f'''
            <table class="param-table">
                <tr><td><span class="icon">⚖️</span>축력 P<sub>b</sub></td><td>{balanced_data.get('Pb_kN', 0):,.1f} kN</td></tr>
                <tr><td><span class="icon">📏</span>모멘트 M<sub>b</sub></td><td>{balanced_data.get('Mb_kNm', 0):,.1f} kN·m</td></tr>
                <tr><td><span class="icon">📐</span>편심 e<sub>b</sub></td><td>{balanced_data.get('eb_mm', 0):,.1f} mm</td></tr>
                <tr><td><span class="icon">🎯</span>중립축 깊이 c<sub>b</sub></td><td>{balanced_data.get('cb_mm', 0):,.1f} mm</td></tr>
            </table>''', unsafe_allow_html=True)

            # 기둥 강도 검토 결과
            st.markdown('<div class="sub-section-header">📊 기둥강도 검토 결과</div>', unsafe_allow_html=True)
            
            check_results = calculate_strength_check(In, material_type)
            
            if check_results:
                def render_html_table(results):
                    html = '''<table class="results-table">
                    <tr>
                        <th>하중조합</th>
                        <th>P<sub>u</sub> / φP<sub>n</sub> [kN]</th>
                        <th>M<sub>u</sub> / φM<sub>n</sub> [kN·m]</th>
                        <th>편심 e [mm]</th>
                        <th>안전률 SF</th>
                        <th>판정</th>
                    </tr>'''
                    
                    all_passed = True
                    for result in results:
                        verdict_class = "pass" if result['Verdict'] == 'PASS' else "fail"
                        if result['Verdict'] == 'FAIL':
                            all_passed = False
                        
                        html += f'''
                        <tr>
                            <td><b>{result['LC']}</b></td>
                            <td>{result['Pu/phiPn']}</td>
                            <td>{result['Mu/phiMn']}</td>
                            <td>{result['e_mm']:.1f}</td>
                            <td>{result['SF']:.2f}</td>
                            <td class="{verdict_class}">{result['Verdict']} {'✅' if result['Verdict'] == 'PASS' else '❌'}</td>
                        </tr>'''
                    html += '</table>'
                    return html, all_passed

                html_table, all_passed = render_html_table(check_results)
                st.markdown(html_table, unsafe_allow_html=True)

                # 최종 종합 판정
                st.markdown('<div class="sub-section-header">🎯 최종 종합 판정</div>', unsafe_allow_html=True)
                if all_passed:
                    st.markdown('<div class="final-verdict-container final-pass">🎉 전체 조건 만족 - 구조 안전</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="final-verdict-container final-fail">⚠️ 일부 조건 불만족 - 보강 검토 필요</div>', unsafe_allow_html=True)
            else:
                st.warning("⚠️ 검토 데이터를 계산할 수 없습니다.", icon="⚠️")
            
            st.markdown('</div>', unsafe_allow_html=True)

    # =================================================================
    # 메인 렌더링
    # =================================================================
    try:
        st.markdown('<div class="main-container">', unsafe_allow_html=True)
        
        # 메인 헤더
        st.markdown('<div class="main-header">🏗️ 기둥 강도 검토 보고서</div>', unsafe_allow_html=True)
        
        # 공통 설계 조건 (분리된 섹션)
        render_common_conditions(In)
        
        # 좌우 검토 섹션
        col1, col2 = st.columns(2, gap="large")
        create_report_column(col1, "📊 이형철근 검토", In, R, "이형철근")
        create_report_column(col2, "📊 중공철근 검토", In, F, "중공철근")
        
        st.markdown('</div>', unsafe_allow_html=True)
        
    except Exception as e:
        st.error(f"⚠️ 보고서 생성 중 오류 발생: {e}")
        
        # 디버깅 정보
        with st.expander("🔍 디버깅 정보 보기"):
            try:
                st.write("**In 객체 속성:**", [attr for attr in dir(In) if not attr.startswith('_')][:15])
                st.write("**R 객체 속성:**", [attr for attr in dir(R) if not attr.startswith('_')][:15])
                st.write("**F 객체 속성:**", [attr for attr in dir(F) if not attr.startswith('_')][:15])
            except:
                st.write("객체 정보를 가져올 수 없습니다.")

