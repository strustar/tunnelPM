import numpy as np
import streamlit as st


class PM1:
    pass


# PM Diagram : RC(KDS-2021, KCI-2012) 🆚 FRP(AASHTO, ACI440)
def Cal(In, Reinforcement_Type):
    # Input Data
    RC_Code = In.RC_Code
    FRP_Code = In.FRP_Code
    Column_Type = In.Column_Type
    Section_Type = In.Section_Type
    PM_Type = In.PM_Type

    [be, height, D] = [In.be, In.height, In.D]
    [fck, fy, f_fu, Ec, Es, Ef, Es_hollow, fy_hollow] = [In.fck, In.fy, In.f_fu, In.Ec, In.Es, In.Ef, In.Es_hollow, In.fy_hollow]
    [Layer, dia, dc, nh, nb, nD, sb, dia1, dc1] = [In.Layer, In.dia, In.dc, In.nh, In.nb, In.nD, In.sb, In.dia1, In.dc1]
    # Input Data
    [ep_y, ep_fu, ep_y_hollow] = [fy / Es, f_fu / Ef, fy_hollow / Es_hollow]
    if 'hollow' in Reinforcement_Type:
        [Es, fy, ep_y] = [Es_hollow, fy_hollow, ep_y_hollow]
    if 'FRP' in Reinforcement_Type:
        [Es, fy, ep_y] = [Ef, f_fu, ep_fu]

    ###* Coefficient ###
    if 'KDS-2021' in RC_Code and 'RC' in Reinforcement_Type:
        [n, ep_co, ep_cu] = [2, 0.002, 0.0033]
        if fck > 40:
            n = 1.2 + 1.5 * ((100 - fck) / 60) ** 4
            ep_co = 0.002 + (fck - 40) / 1e5
            ep_cu = 0.0033 - (fck - 40) / 1e5
        if n >= 2:
            n = 2
        n = round(n * 100) / 100

        alpha = 1 - 1 / (1 + n) * (ep_co / ep_cu)
        temp = 1 / (1 + n) / (2 + n) * (ep_co / ep_cu) ** 2
        if fck <= 40:
            alpha = 0.8
        beta = 1 - (0.5 - temp) / alpha
        if fck <= 50:
            beta = 0.4

        [alpha, beta] = [round(alpha * 100) / 100, round(beta * 100) / 100]
        beta1 = 2 * beta
        eta = alpha / beta1
        eta = round(eta * 100) / 100
        if fck == 50:
            eta = 0.97
        if fck == 80:
            eta = 0.87
    else:
        [ep_cu, eta] = [0.003, 1.0]
        beta1 = 0.85 if fck <= 28 else 0.85 - 0.007 * (fck - 28)
        if beta1 < 0.65:
            beta1 = 0.65

    if 'Tied' in Column_Type:
        [alpha, phi0] = [0.80, 0.65]
    if 'Spiral' in Column_Type:
        [alpha, phi0] = [0.85, 0.70]
    ###* Coefficient ###

    ###* Preparation for Calculation ###
    [nst, nst1, ni, Ast, Ast1] = [np.zeros(Layer), np.zeros(Layer), np.zeros(Layer), [np.pi * d**2 / 4 for d in dia], [np.pi * d**2 / 4 for d in dia1]]
    if 'hollow' in Reinforcement_Type:
        Ast = [a / 2 for a in Ast]  ## / 2 : 중공철근 !!!
        Ast1 = [a / 2 for a in Ast1]  ## / 2 : 중공철근 !!!
    if 'Rectangle' in Section_Type:
        [hD, Ag, ni] = [height, be * height, nh]
        for L in range(Layer):
            nst[L] = be / sb[0]
            nst1[L] = be / sb[0]
            # for i in range(nh[L]):
            #     for j in range(nb[L]):
            #         if (i > 0 and i < nh[L] - 1) and (j > 0 and j < nb[L] - 1):
            #             continue
            #         nst[L] = nst[L] + 1

    Ast_total = np.multiply(nst, Ast) + np.multiply(nst1, Ast1)
    rho = np.sum(Ast_total) / Ag
    if 'Rectangle' in Section_Type:
        A1 = 0
        A2 = 0

    dsi, Asi = np.zeros((Layer, np.max(ni))), np.zeros((Layer, np.max(ni)))  # initial_rotation = 0 for Circle Section
    dsi[0, :2] = [dc1[0], height - dc1[0]]
    # nb = be / sb[0]
    Asi[0, :2] = [Ast1[0] * nst1[0], Ast[0] * nst[0]]

    # for L in range(Layer):
    #     for i in range(ni[L]):
    #         if 'Rectangle' in Section_Type:
    #             # dsi[L, i] = dc[L] + i * (height - 2 * dc[L]) / (ni[L] - 1)
    #             Asi[L, i] = 2 * Ast[L] / nst[L]
    #             Asi[L, 0] = nb[L] * Ast[L] / nst[L]
    #             Asi[L, ni[L] - 1] = Asi[L, 0]
    ###* Preparation for Calculation ###

    # st.write(dsi, Asi, nst)
    d = np.max(dsi)
    gamma = d / hD
    n = 7
    [cc, Pn, Mn, ee] = [np.zeros(n), np.zeros(n), np.zeros(n), np.zeros(n)]
    [ep_s, fs] = [np.zeros((n, 2)), np.zeros((n, 2))]

    ###* Calculation Point A(1-1) : Pure Comppression (e = 0, c = inf, Mn = 0) ###
    z1 = 1 - 1
    cc[z1] = np.inf
    Mn[z1] = 0
    Rein = 0 if 'FRP' in Reinforcement_Type else fy * np.sum(Ast_total)
    Pn[z1] = (eta * 0.85 * fck * (Ag - np.sum(Ast_total)) + Rein) / 1e3
    if ('FRP' in Reinforcement_Type) and ('ACI 440.1' in FRP_Code):
        Pn[z1] = 0.85 * fck * Ag / 1e3
    ee[z1] = Mn[z1] / Pn[z1] * 1e3
    ep_s[z1, 0:2] = ep_cu
    fs[z1, 0:2] = fy
    ###* Calculation Point A(1-1) : Pure Comppression (e = 0, c = inf, Mn = 0) ###

    ###* Calculation Point G(7-1) : Pure Tension(e = 0, c = inf, Mn = 0)
    z1 = 7 - 1
    cc[z1] = -np.inf
    Mn[z1] = 0
    Pn[z1] = -np.sum(Ast_total) * fy / 1e3
    ee[z1] = -0.000000001  # Mn[z1]/Pn[z1]*1e3
    ep_s[z1, 0:2] = -ep_y
    fs[z1, 0:2] = -fy
    ###* Calculation Point G(7-1) : Pure Tension(e = 0, c = inf, Mn = 0)

    ##* %%% Calculation Point x = c = 0(8-1) : for Only ACI 440.1
    if 'FRP' in Reinforcement_Type and 'ACI 440.1' in FRP_Code:  #! for ACI 440.1R**  Only Only
        if 'Rectangle' in Section_Type:
            M = (2 * gamma - 1) ** 2 / (2 * gamma) * (A1 + A2 / 3) * f_fu * hD / 1e6
        if 'Circle' in Section_Type:
            M = (2 * gamma - 1) ** 2 / (8 * gamma) * np.sum(Ast_total) * f_fu * hD / 1e6
        P = -np.sum(Ast_total) / (2 * gamma) * f_fu / 1e3
        [Pn8, Mn8, Pd8, Md8] = [P, M, 0.55 * P, 0.55 * M]  # c = 0, ep_s = -ep_y, fs = -fy
    ##* %%% Calculation Point x = c = 0(8-1) : for Only ACI 440.1

    ###* Calculation Point C(3-1), D(4-1) (Zero Tension, Balacne Point) and E(5-1) &&  Z = eps/ep_cu  0.1~9.?
    [ep_si, fsi, Fsi] = [np.zeros((Layer, np.max(ni))), np.zeros((Layer, np.max(ni))), np.zeros((Layer, np.max(ni)))]
    for zz in [1, 2]:
        if zz == 1:
            [zz1, zz2] = [3 - 1, 5 - 1]
        if zz == 2:
            [zz1, zz2] = [0, 90]
        [Zcc, ZMn, ZPn, Zee] = [np.zeros(zz2 + 1), np.zeros(zz2 + 1), np.zeros(zz2 + 1), np.zeros(zz2 + 1)]
        [Zep_s, Zfs] = [np.zeros((zz2 + 1, 2)), np.zeros((zz2 + 1, 2))]
        for z1 in range(zz1, zz2 + 1):
            if ('FRP' in Reinforcement_Type) and ('ACI 440.1' in FRP_Code):  #############! for ACI 440.1**
                xb = d * ep_cu / (ep_cu + ep_fu)  # (xb <= x <= d & d < x <=h)
                if zz == 1:
                    if z1 == 3 - 1:
                        x = d  # C x = d    Zero Tension (ep_s(end) = 0)
                    if z1 == 4 - 1:
                        x = xb  # D x = xb   Balance Point (ep_t = ep_fu)
                    if z1 == 5 - 1:
                        x = d * ep_cu / (ep_cu + 0.8 * ep_fu)  # E x = 0.8*xb (ep_t = 0.8*ep_fu)
                if zz == 2:
                    # x = xb + (z1 - 1)*(d - xb)/(zz2 - 1)
                    x = xb + z1 * (d - xb) / zz2

                [alp, c] = [x / d, x]
                if 'Rectangle' in Section_Type:
                    [P, M] = ACI440_Rectangle(alp, beta1, gamma, fck, ep_cu, Ef, A1, A2, be, height)
                if 'Circle' in Section_Type:
                    [P, M] = ACI440_Circle(alp, beta1, gamma, fck, ep_cu, Ef, Ast_total, D)

                for L in range(Layer):
                    for i in range(ni[L]):
                        ep_si[L, i] = ep_cu * (c - dsi[L, i]) / c
                        fsi[L, i] = Es * ep_si[L, i]
            else:  #############! for RC & AASHTO
                if zz == 1:
                    if z1 == 3 - 1:
                        eps = 0  # C   0% fy or ffu  Zero Tension (ep_s(end) = 0)
                    if z1 == 4 - 1:
                        eps = 1.00 * ep_y  # D 100% fy or ffu  Balance Point
                    if z1 == 5 - 1:
                        eps = 2.50 * ep_y if 'RC' in Reinforcement_Type else 0.8 * ep_y  # E 250% fy or 80% ffu
                if zz == 2:
                    Z = z1 / 10
                    eps = Z * ep_cu

                c = d * ep_cu / (ep_cu + eps)
                if 'Rectangle' in Section_Type:
                    bhD = [hD, be, height]
                [P, M] = RC_and_AASHTO(Section_Type, Reinforcement_Type, beta1, c, eta, fck, Layer, ni, ep_si, ep_cu, dsi, fsi, Es, fy, Fsi, Asi, *bhD)

            if zz == 1:
                cc[z1] = c
                Mn[z1] = M
                Pn[z1] = P
                ee[z1] = Mn[z1] / Pn[z1] * 1e3
                ep_s[z1, 0] = ep_si[0, 0]
                ep_s[z1, 1] = ep_si[0, -1]
                fs[z1, 0] = fsi[0, 0]
                fs[z1, 1] = fsi[0, -1]
            if zz == 2:
                Zcc[z1] = c
                ZMn[z1] = M
                ZPn[z1] = P
                Zee[z1] = ZMn[z1] / ZPn[z1] * 1e3
                Zep_s[z1, 0] = ep_si[0, 0]
                Zep_s[z1, 1] = ep_si[0, -1]
                Zfs[z1, 0] = fsi[0, 0]
                Zfs[z1, 1] = fsi[0, -1]
    ###* Calculation Point C(3-1), D(4-1) (Zero Tension, Balacne Point) and E(5-1) &&  Z = eps/ep_cu  0.1~9.?

    ###* Calculation Point B(2-1), F(6-1) : Minimum Eccentricity (e = min), Pure Moment(Pn = 0)
    [iter, tol, temp] = [1e5, 1e-1, np.zeros(1000)]
    for z1 in [2 - 1, 6 - 1]:
        if z1 == 2 - 1:
            Pnn = alpha * Pn[1 - 1]
        if z1 == 6 - 1:
            Pnn = 0
        if ('FRP' in Reinforcement_Type) and ('ACI 440.1' in FRP_Code):  #############! for ACI 440.1**
            if z1 == 2 - 1:
                if 'Rectangle' in Section_Type:
                    x = alpha / beta1 * height
                    alp = x / d
                    c = x
                    Cc = Pnn
                    Lc = (1 - alp * beta1 * gamma) * height / 2
                    M = Cc * Lc / 1e3
                if 'Circle' in Section_Type:
                    for k3 in [1, 2, 3]:
                        for k1 in np.arange(1, iter):
                            x = hD / 2 + (k1 - 1) * hD / iter
                            alp = x / d
                            c = x
                            [P, M] = ACI440_Circle(alp, beta1, gamma, fck, ep_cu, Ef, Ast_total, D)
                            if np.abs(Pnn - P) < tol:
                                break
                        if k1 == iter:
                            tol = 10 * tol  #! 수렴이 안될 경우
                        else:
                            break
            elif z1 == 6 - 1:
                if Pn[4 - 1] > 0:  #! Pn(4-1) : D, Balance Point > 0일때 구할수 없고(식이 없음), 직선 보간 한다
                    P = 0
                    # M = Pn8*(Mn[4-1] - Mn8)/(Pn8 + Pn[4-1])
                    M = Mn[4 - 1] - Pn[4 - 1] * (Mn[4 - 1] - Mn8) / (Pn[4 - 1] - Pn8)
                    c = cc[4 - 1] * 0.9
                    pass
                elif Pn[4 - 1] <= 0:  #! Pn(4-1) : D, Balance Point < 0일때 F(Pn=0)점을 계산으로 구할수 있다.
                    if 'Rectangle' in Section_Type:
                        for k3 in [1, 2, 3]:
                            for k1 in np.arange(1, iter):
                                x = hD / 100 + (k1 - 1) * hD / 2 / iter
                                alp = x / d
                                c = x
                                [P, M] = ACI440_Rectangle(alp, beta1, gamma, fck, ep_cu, Ef, A1, A2, be, height)
                                if np.abs(P - Pnn) <= tol:
                                    break
                            if k1 == iter:
                                tol = 10 * tol  #! 수렴이 안될 경우
                            else:
                                break
                    if 'Circle' in Section_Type:
                        k7 = 0
                        x = hD / 2
                        for k1 in np.arange(1, 1000):  #! 처음부터 고려한 경우  ##########################
                            alp = x / d
                            c = x
                            [P, M] = ACI440_Circle(alp, beta1, gamma, fck, ep_cu, Ef, Ast_total, D)
                            temp[k1] = Pnn - P
                            sgn1 = np.sign(temp[k1 - 1])
                            sgn2 = np.sign(temp[k1])
                            sgn = sgn1 * sgn2
                            if sgn == -1:
                                k7 = k7 + 1
                            x = x + sgn2 * 100 / 10**k7
                            if np.abs(temp[k1]) < 1e-1:
                                break
                        # tol = 1
                        # for k3 in [1, 2, 3]:          #! 처음부터 고려한 경우  ##########################
                        #     for k1 in np.arange(1, iter):
                        #         x = hD/100 + (k1 - 1)*hD/2/iter;  alp = x/d;  c = x
                        #         [P, M] = ACI440_Circle(alp, beta1, gamma, fck, ep_cu, Ef, Ast_total, D)
                        #         if np.abs(Pnn - P) < tol: break
                        #     if k1 == iter: tol = 10*tol   #! 수렴이 안될 경우
                        #     else: break

            for L in range(Layer):
                for i in np.arange(ni[L]):
                    ep_si[L, i] = ep_cu * (c - dsi[L, i]) / c
                    fsi[L, i] = Es * ep_si[L, i]
        else:  #############! for RC & AASHTO
            # 초기화
            k7 = 0
            c = hD / 2
            temp = np.zeros(1001)  # temp 배열 명시적 초기화
            temp[1] = float('inf')  # 첫 번째 값을 무한대로 설정하여 비교 가능하게 함

            # Pnn=0인 경우 특별 처리
            if abs(Pnn) < 1e-6:
                # 순수 휨 상태에서는 중립축 위치를 단면 높이의 일정 비율로 시작
                c = 0.4 * hD  # 일반적으로 RC 단면에서 순수 휨 상태의 중립축은 높이의 약 0.4배 근처

                # 더 세밀한 초기 스텝 크기
                step_size = 0.05 * hD
                direction = 1  # 초기 탐색 방향

                # 첫 번째 값 계산
                bhD = [hD, be, height] if 'Rectangle' in Section_Type else [hD, D]
                [P_init, M_init] = RC_and_AASHTO(Section_Type, Reinforcement_Type, beta1, c, eta, fck, Layer, ni, ep_si, ep_cu, dsi, fsi, Es, fy, Fsi, Asi, *bhD)
                temp[1] = P_init  # Pnn=0이므로 P_init 자체가 오차

                # 방향 결정
                if P_init > 0:
                    direction = -1  # P를 줄이는 방향으로 탐색
                else:
                    direction = 1  # P를 늘리는 방향으로 탐색

                for k1 in np.arange(2, 1000):
                    bhD = [hD, be, height] if 'Rectangle' in Section_Type else [hD, D]
                    [P, M] = RC_and_AASHTO(Section_Type, Reinforcement_Type, beta1, c, eta, fck, Layer, ni, ep_si, ep_cu, dsi, fsi, Es, fy, Fsi, Asi, *bhD)

                    temp[k1] = P  # Pnn=0이므로 P 자체가 오차

                    # 부호 변화 확인
                    sgn1 = np.sign(temp[k1 - 1])
                    sgn2 = np.sign(temp[k1])
                    sgn = sgn1 * sgn2

                    # 부호가 바뀌면 스텝 크기 감소 및 방향 전환
                    if sgn == -1:
                        k7 = k7 + 1
                        direction = -direction  # 방향 전환

                    # 중립축 위치 조정 (더 점진적으로)
                    c = c + direction * step_size / 10**k7

                    # 충분히 정확하면 종료
                    if abs(temp[k1]) <= 1e-1:
                        break

                    # 너무 많은 반복에도 수렴하지 않으면 가장 좋은 해 사용
                    if k1 > 500:
                        best_idx = np.argmin(np.abs(temp[2 : k1 + 1])) + 2
                        c_best = c - (k1 - best_idx) * direction * step_size / 10**k7
                        c = c_best
                        st.write(f"경고: 100회 반복 후 최적해 선택. P={temp[best_idx]}")
                        break

                    # 스텝 크기가 너무 작아지면 종료
                    if k7 > 6:
                        st.write(f"정보: 정밀도 한계에 도달. 마지막 P={temp[k1]}")
                        break
            else:
                # 기존 코드 (Pnn != 0인 경우)
                for k1 in np.arange(2, 1000):
                    bhD = [hD, be, height] if 'Rectangle' in Section_Type else [hD, D]
                    [P, M] = RC_and_AASHTO(Section_Type, Reinforcement_Type, beta1, c, eta, fck, Layer, ni, ep_si, ep_cu, dsi, fsi, Es, fy, Fsi, Asi, *bhD)

                    temp[k1] = Pnn - P
                    sgn1 = np.sign(temp[k1 - 1])
                    sgn2 = np.sign(temp[k1])
                    sgn = sgn1 * sgn2

                    if sgn == -1:
                        k7 = k7 + 1

                    c = c + sgn2 * 100 / 10**k7

                    if abs(temp[k1]) <= 1e-1:
                        break
        cc[z1] = c
        Mn[z1] = M
        Pn[z1] = Pnn
        ee[z1] = np.inf if z1 == 6 - 1 else Mn[z1] / Pn[z1] * 1e3
        ep_s[z1, 0] = ep_si[0, 0]
        ep_s[z1, 1] = ep_si[0, -1]
        fs[z1, 0] = fsi[0, 0]
        fs[z1, 1] = fsi[0, -1]
    ###* Calculation Point B(2-1), F(6-1) : Minimum Eccentricity (e = min), Pure Moment(Pn = 0)

    ###*%% Sorting
    Ze = np.concatenate((ee, Zee))
    Zc = np.concatenate((cc, Zcc))
    ZPn = np.concatenate((Pn, ZPn))
    ZMn = np.concatenate((Mn, ZMn))
    Zep_s = np.concatenate((ep_s, Zep_s))
    Zfs = np.concatenate((fs, Zfs))
    loc = np.argsort(ZPn)
    Ze = Ze[loc][::-1]
    Zc = Zc[loc][::-1]
    ZPn = ZPn[loc][::-1]
    ZMn = ZMn[loc][::-1]
    Zep_s = Zep_s[loc][::-1]
    Zfs = Zfs[loc][::-1]
    ###*%% Sorting

    ### 강도감소계수 (phi)
    for zz in [1, 2]:
        if zz == 1:
            n = len(Pn)
            phi = np.zeros(n)
            phi_Status = ['' for _ in range(n)]
        if zz == 2:
            n = len(ZPn)
            Zphi = np.zeros(n)
        for z1 in range(n - 1):
            if zz == 1:
                eps = -ep_s[z1, 1]  # 압축 (+)
            if zz == 2:
                eps = -Zep_s[z1, 1]
            if 'RC' in Reinforcement_Type:  #! RC : KDS-2021 & KCI-2012
                [ep_tccl, ep_ttcl] = [ep_y, 0.005]
                if fy >= 400:
                    ep_ttcl = 2.5 * ep_y
                ph = phi0 + (eps - ep_tccl) * (0.85 - phi0) / (ep_ttcl - ep_tccl)
                ph_Status = 'Transition Zone'
                if eps <= ep_tccl:
                    ph = phi0
                    ph_Status = 'Compression Controlled'
                if eps >= ep_ttcl:
                    ph = 0.85
                    ph_Status = 'Tension Controlled'

            if 'FRP' in Reinforcement_Type:  #! FRP : AASHTO & ACI  (ACI 440.1R**은 약간 다른데 거의 비슷해서 여기에 포함)
                if 'AASHTO' in FRP_Code:
                    a1 = 1.55
                    a2 = 1
                    a3 = 0.75
                elif 'ACI 440.1' in FRP_Code:
                    a1 = 1.05
                    a2 = 0.5
                    a3 = 0.65
                ph = a1 - a2 * eps / ep_fu
                ph_Status = 'Transition Zone'
                if eps <= 0.8 * ep_fu:
                    ph = a3
                    ph_Status = 'Compression Controlled'
                if eps >= ep_fu:
                    ph = 0.55
                    ph_Status = 'Tension Controlled'
            if zz == 1:
                phi[z1] = ph
                phi_Status[z1] = ph_Status
            if zz == 2:
                Zphi[z1] = ph
    if 'RC' in Reinforcement_Type:
        [phi[-1], Zphi[-1]] = [0.85, 0.85]
        phi_Status[-1] = 'Tension Controlled'
    if 'FRP' in Reinforcement_Type:
        [phi[-1], Zphi[-1]] = [0.55, 0.55]
        phi_Status[-1] = 'Tension Controlled'
    ### 강도감소계수 (phi)
    Pd = phi * Pn
    Md = phi * Mn
    ZPd = Zphi * ZPn
    ZMd = Zphi * ZMn

    PM = PM1()
    In.hD = hD
    In.nhD = nh
    if 'Circle' in Section_Type:
        In.nhD = nD
    PM.ep_y, PM.ep_fu, PM.ep_cu = ep_y, ep_fu, ep_cu
    PM.beta1, PM.eta, PM.alpha, PM.phi0 = beta1, eta, alpha, phi0

    if 'Rectangle' in Section_Type:
        [PM.A1, PM.A2] = [A1, A2]
    PM.Ag = Ag
    PM.Ast_total = Ast_total
    PM.nst = nst
    PM.dsi = dsi
    PM.rho = rho

    PM.e = ee
    PM.c = cc
    PM.Pn = Pn
    PM.Mn = Mn
    PM.ep_s = ep_s
    PM.fs = fs
    PM.phi_Status = phi_Status
    PM.phi = phi
    PM.Pd = Pd
    PM.Md = Md
    PM.Ze = Ze
    PM.Zc = Zc
    PM.ZPn = ZPn
    PM.ZMn = ZMn
    PM.Zep_s = Zep_s
    PM.Zfs = Zfs
    PM.Zphi = Zphi
    PM.ZPd = ZPd
    PM.ZMd = ZMd
    if 'FRP' in Reinforcement_Type and 'ACI 440.1' in FRP_Code:  #! for ACI 440.1R**  Only Only
        PM.Pn8 = Pn8
        PM.Mn8 = Mn8
        PM.Pd8 = Pd8
        PM.Md8 = Md8
    return PM


def ACI440_Circle(alp, beta1, gamma, fck, ep_cu, Ef, Ast_total, D):
    t = np.arccos(1 - 2 * alp * beta1 * gamma)
    if t < 0 or t > np.pi:
        st.wr('ACI440_Circle theta', t)
    Cc = 0.85 * fck * (t - np.sin(t) * np.cos(t)) * D**2 / 4
    Mc = 0.85 * fck * (np.sin(t)) ** 3 * D**3 / 12

    tt = (1 - 2 * alp * gamma) / (2 * gamma - 1)
    if tt < -1:
        tt = -1
    if tt > 1:
        tt = 1
    t = np.arccos(tt)
    if t < 0 or t > np.pi:
        st.wr('ACI440_Circle theta', t)
    if np.iscomplex(t) or np.abs(t - np.pi) < 1e-2:  # if np.abs(t - np.pi) < 1e-2:  T = 0;  Mt = 0
        [T, Mt] = [0, 0]
    else:
        T = (np.pi * np.cos(t) - t * np.cos(t) + np.sin(t)) / (np.pi * (1 + np.cos(t))) * (1 - alp) / alp * ep_cu * Ef * np.sum(Ast_total)
        Mt = (np.pi - t + np.sin(t) * np.cos(t)) / (2 * np.pi * (1 + np.cos(t))) * (1 - alp) / alp * (gamma - 1 / 2) * ep_cu * Ef * np.sum(Ast_total) * D

    [P, M] = [(Cc - T) / 1e3, (Mc + Mt) / 1e6]
    return P, M


def ACI440_Rectangle(alp, beta1, gamma, fck, ep_cu, Ef, A1, A2, b, h):
    Cc = 0.85 * alp * beta1 * gamma * fck * b * h
    Lc = (1 - alp * beta1 * gamma) * h / 2
    T1 = (1 - alp) / alp * ep_cu * Ef * A1
    L1 = (2 * gamma - 1) * h / 2
    T2 = gamma / (2 * gamma - 1) * (1 - alp) ** 2 / alp * ep_cu * Ef * A2
    L2 = (2 / 3 * (2 + alp) * gamma - 1) * h / 2

    [P, M] = [(Cc - T1 - T2) / 1e3, (Cc * Lc + T1 * L1 + T2 * L2) / 1e6]
    return P, M


def RC_and_AASHTO(Section_Type, Reinforcement_Type, beta1, c, eta, fck, Layer, ni, ep_si, ep_cu, dsi, fsi, Es, fy, Fsi, Asi, *bhD):
    a = beta1 * c
    if 'Rectangle' in Section_Type:
        [hD, b, h] = bhD
        Ac = a * b if a < h else h * b
        y_bar = (h - a) / 2

    [Cc, M] = [eta * (0.85 * fck) * Ac / 1e3, 0]
    for L in range(Layer):
        for i in range(ni[L]):
            if c == 0:
                continue
            ep_si[L, i] = ep_cu * (c - dsi[L, i]) / c
            fsi[L, i] = Es * ep_si[L, i]
            if fsi[L, i] >= fy:
                fsi[L, i] = fy
            if fsi[L, i] <= -fy:
                fsi[L, i] = -fy
            if 'RC' in Reinforcement_Type:
                if c >= dsi[L, i]:
                    Fsi[L, i] = Asi[L, i] * (fsi[L, i] - eta * 0.85 * fck) / 1e3  # 압축 철근  -0.85*fck
                elif c < dsi[L, i]:
                    Fsi[L, i] = Asi[L, i] * fsi[L, i] / 1e3  #! 인장 철근 (c로 판단) ccccc
            M = M + Fsi[L, i] * (hD / 2 - dsi[L, i])

    [P, M] = [np.sum(Fsi) + Cc, (M + Cc * y_bar) / 1e3]  # Mc =  Cc*y_bar
    return P, M
