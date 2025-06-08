import numpy as np
from scipy.optimize import fsolve

def calculate_steel_stress(In, R):
    """
    여러 하중 케이스(P0, M0)에 대해 반복적으로 철근 응력을 계산하고,
    그 결과를 리스트 형태로 반환하는 메인 함수입니다.

    Args:
        In (object): 단면과 재료의 공통 속성을 담고 있는 객체.
        R (object): 철근량(Ast_total) 등 추가 정보를 담고 있는 객체.

    Returns:
        list: 각 하중 케이스별 계산 결과 딕셔너리의 리스트.
            실패 시 빈 리스트 또는 실패 정보가 담긴 리스트를 반환할 수 있습니다.
    """
    
    # --- 1. 입력값 안전하게 추출 ---
    try:
        # 공통 속성 추출
        b = float(getattr(In, 'be', []))
        h = float(getattr(In, 'height', []))
        d = float(getattr(In, 'depth', []))
        Ec = float(getattr(In, 'Ec', []))
        Es = float(getattr(In, 'Es', []))
        fy = float(getattr(R, 'fy', []))
        As = float(getattr(R, 'Ast_total', []))
        
        # 하중 케이스 (리스트 형태) 추출
        P0_list = getattr(In, 'P0', [])
        M0_list = getattr(In, 'M0', [])

        # numpy 배열일 경우 리스트로 변환
        if hasattr(P0_list, 'tolist'): P0_list = P0_list.tolist()
        if hasattr(M0_list, 'tolist'): M0_list = M0_list.tolist()

        if len(P0_list) != len(M0_list):
            raise ValueError("P0와 M0의 하중 케이스 개수가 일치하지 않습니다.")

    except (TypeError, ValueError) as e:
        print(f"입력 데이터 처리 중 오류 발생: {e}")
        return []

    # --- 2. 단일 케이스 계산을 위한 내부 함수 ---
    def _solve_for_single_case(params, P0_kN, M0_kNm):
        """
        단일 하중 케이스에 대해 fsolve를 실행하는 내부 함수.
        이 함수는 반복문 안에서 호출됩니다.
        """
        # 평형 방정식 정의
        def equilibrium_equations(vars):
            x, ec = vars
            
            # 물리적으로 불가능한 해 방지
            if x <= 0 or x >= h or ec <= 0:
                return [1e9, 1e12]

            es = ec * (d - x) / x
            fs = np.clip(Es * es, -fy, fy) # 탄성-완전소성 모델
            
            C = 0.5 * b * x * (Ec * ec)
            T = As * fs
            
            # N, N-mm 단위로 계산
            force_eq = C - T - (P0_kN * 1000)
            moment_eq = C * (h/2 - x/3) + T * (d - h/2) - (M0_kNm * 1e6)
            
            return [force_eq, moment_eq]

        # 여러 초기 추정값으로 해찾기 시도
        initial_guesses = [[h/2, 0.001], [h/3, 0.0015], [h/4, 0.0008], [2*h/3, 0.002]]
        for guess in initial_guesses:
            solution, info, ier, msg = fsolve(equilibrium_equations, guess, full_output=True, xtol=1e-8)
            if ier == 1: # 성공적으로 해를 찾은 경우
                x_sol, ec_sol = solution
                if 0 < x_sol < h: # 물리적으로 타당한 해인지 검증
                    es_sol = ec_sol * (d - x_sol) / x_sol
                    fs_sol = np.clip(Es * es_sol, -fy, fy)
                    return {'success': True, 'fs': fs_sol, 'es': es_sol, 'ec': ec_sol, 'x': x_sol}
        
        # 모든 초기값으로 실패한 경우
        return {'success': False, 'message': '수렴하는 해를 찾지 못했습니다.'}

    # --- 3. 모든 하중 케이스에 대해 반복 계산 ---
    all_results = []
    for P0_val, M0_val in zip(P0_list, M0_list):
        # 현재 케이스의 파라미터로 내부 계산 함수 호출
        result = _solve_for_single_case(locals(), P0_val, M0_val)
        all_results.append(result)

    return all_results

