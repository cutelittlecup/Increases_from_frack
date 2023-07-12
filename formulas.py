import math


def viscosity_liq_finder(Wc, viscosity_oil, viscosity_water, B0):
    fw = 0.3
    Exo = 1.0
    Exw = 1.0

    if Wc > 0 and Wc != 1:
        S0 = (1 / Wc - 1) * (viscosity_oil * B0 * fw) / (viscosity_water * 1.01)
        a = 0
        b = 1
        S1 = 0
        S = (a + b) / 2
        while not ((S - S1) / S <= 0.05 and abs(((1 - S) ** Exo) / (S ** Exw) - S0) < 0.01):
            if ((1 - S) ** Exo) / (S ** Exw) > S0:
                a = S
            else:
                b = S
            S1 = S
            S = (a + b) / 2
        Kr_o = (1 - S) ** Exo
        Kr_w = fw * S ** Exw
        viscosity_liq = (viscosity_water * viscosity_oil) / (Kr_o * viscosity_water + Kr_w * viscosity_oil)

    elif Wc == 1:
        viscosity_liq = viscosity_water
    else:
        viscosity_liq = 0
    return viscosity_liq


def formulas(objects_info, keys, text):
    # константы
    K_good = 1  # Кусп
    Re = 250  # Радиус контура питания, Re
    Kf_fact = 500000  # Kf, мД проницаемость пропанта ФАКТ
    Kf_plan = Kf_fact  # Kf, мД проницаемость пропанта ПЛАН
    wf_fact = 0.005  # ширина трещины ФАКТ wf, м
    wf_plan = wf_fact  # ширина трещины ПЛАН wf, м
    Ld = 1  # Ld, дол.ед (=1)
    coeff_anizotropii = 0.1  # Коэффициент анизотропии пласта
    method = 1

    for i in range(len(keys)):

        n_grp_povt = objects_info[keys[i]][32]

        if n_grp_povt < 4:
            for_djoshi = 1
            if objects_info[keys[i]][0] == 'ГС' or objects_info[keys[i]][0] == 1:
                if objects_info[keys[i]][1] > 1:
                    Mpr_stock = objects_info[keys[i]][2]
                    if Mpr_stock == 0 and objects_info[keys[i]][1] == 0 and n_grp_povt != 0:
                        for_djoshi = 0
                    if Mpr_stock == 0 and objects_info[keys[i]][1] != 0:
                        for_djoshi = 0

            if type(objects_info[keys[i]][23]) == float and (for_djoshi != 0) and (
                    keys[i] not in list(objects_info_problems.keys())):
                # print('')
                # print('расчёт для ' + keys[i])
                # print(objects_info[keys[i]])

                n_mgrp_fact = objects_info[keys[i]][1]  # Кол-во стадий МГРП ФАКТ
                # print('n_mgrp_fact: ' + str(n_mgrp_fact))
                Mpr_stock = objects_info[keys[i]][2]
                # print('Mpr_stock: ' + str(Mpr_stock))
                l_fact = objects_info[keys[i]][3]  # Длина ГС (расстояние между крайними портами) ФАКТ
                # print('l_fact: ' + str(l_fact))
                B0 = objects_info[keys[i]][4]  # Bo
                # print('B0: ' + str(B0))
                a_Xf = objects_info[keys[i]][5]  # a, Xf
                # print('a_Xf: ' + str(a_Xf))
                b_Xf = objects_info[keys[i]][6]  # b, Xf
                # print('b_Xf: ' + str(b_Xf))
                S_nns = objects_info[keys[i]][7]  # S ННС расч
                # print('S_nns: ' + str(S_nns))
                M_plan = objects_info[keys[i]][8]  # Мпр план
                # print('M_plan: ' + str(M_plan))
                P_nas = objects_info[keys[i]][9]  # Рнас
                # print('P_nas: ' + str(P_nas))
                viscosity_oil = objects_info[keys[i]][10]  # Вязкость нефти
                # print('viscosity_oil: ' + str(viscosity_oil))
                viscosity_water = objects_info[keys[i]][11]  # Вязкость воды
                # print('viscosity_water: ' + str(viscosity_water))
                nnt = objects_info[keys[i]][16]  # ННТ / Нэфф
                # print('nnt: ' + str(nnt))

                if objects_info[keys[i]][0] == "ГС" or objects_info[keys[i]][0] == 1 or objects_info[keys[i]][0] == 2:
                    if n_mgrp_fact > 1:
                        formula_type = 1  # формула Ли
                    else:
                        formula_type = 2  # формула Джоши
                else:
                    formula_type = 3  # формула Дюпюи

                if formula_type != 3:
                    rw = 0.07  # Радиус скважины, rw
                else:
                    rw = 0.108
                # print(formula_type)
                objects_info[keys[i]][0] = formula_type

                P_pl_now = objects_info[keys[i]][25]  # Pпл на последний месяц
                P_zab_now = objects_info[keys[i]][26]  # Pзаб на последний месяц
                B_tr_now = objects_info[keys[i]][24]  # % ТР на последний месяц

                viscosity_pot = viscosity_liq_finder(B_tr_now / 100, viscosity_oil, viscosity_water, B0)
                # Вязкость жидкости при расчёте потенциала
                # print('viscosity_pot: ' + str(viscosity_pot))

                # n_mgrp_plan = n_mgrp_fact  # Кол-во стадий МГРП ПЛАН
                n_mgrp_plan = 3

                l_plan = l_fact  # Длина ГС (расстояние между крайними портами) ПЛАН

                plan_Xf = a_Xf * math.log(M_plan) + b_Xf  # ПЛАН Xf
                plan_Xf = plan_Xf * 1.15
                try:
                    fact_Xf = a_Xf * math.log(Mpr_stock) + b_Xf  # ФАКТ Xf
                except ValueError:
                    fact_Xf = 0

                KP_X_plan = 2 * plan_Xf + Re  # Контур питания X ПЛАН
                KP_X_fact = fact_Xf + Re  # Контур питания X, факт

                KP_Y_plan = l_plan + Re * 2  # Контур питания Y ПЛАН
                KP_Y_fact = l_fact + Re * 2  # Контур питания Y, факт

                plan_x_xf = KP_X_plan / plan_Xf  # ПЛАН x/xf

                try:
                    fact_x_xf = KP_X_fact / fact_Xf  # ФАКТ x/xf
                except ZeroDivisionError:
                    fact_x_xf = 0

                if formula_type == 1:

                    Lf1_plan = (l_plan / n_mgrp_plan) / 2  # Lf1,  ПЛАН

                    if (l_fact / n_mgrp_fact) / 2 > Lf1_plan:
                        Lf1_fact = Lf1_plan  # Lf1 ФАКТ
                    else:
                        Lf1_fact = (l_fact / n_mgrp_fact) / 2

                    Lf2_plan = Lf1_plan  # Lf2,  ПЛАН
                    Lf2_fact = Lf1_fact  # Lf2 ФАКТ

                    Lf22_plan = (KP_Y_plan - l_plan) / 2 + Lf2_plan  # Lf22 ПЛАН
                    Lf22_fact = (KP_Y_fact - l_fact) / 2 + Lf2_fact  # Lf22 ФАКТ

                    L_star_plan = plan_x_xf * plan_Xf / 2  # L*, ПЛАН
                    L_star_fact = fact_x_xf * fact_Xf / 2  # L*, ФАКТ

                    c_plan = plan_Xf / nnt - 1 / 2 + math.log(nnt / 2 / rw) / 3.1415  # c ПЛАН
                    c_fact = fact_Xf / nnt - 1 / 2 + math.log(nnt / 2 / rw) / 3.1415  # c

                    Ra2_plan = 1 / (
                            nnt * plan_Xf * (
                            1 / Lf1_plan + 1 / Lf22_plan)) + c_plan / Kf_plan / wf_plan  # R_a2 ПЛАН
                    Ra2 = 1 / (nnt * fact_Xf * (1 / Lf1_fact + 1 / Lf22_fact)) + c_fact / Kf_fact / wf_fact  # R_a2

                    Rb2_plan = Ld * (Lf1_plan + Lf22_plan) / c_plan  # R_b2 ПЛАН
                    Rb2 = Ld * (Lf1_fact + Lf22_fact) / c_fact  # R_b2

                    Rd2_plan = (L_star_plan - plan_Xf) / (nnt * (Lf1_plan + Lf22_plan))  # R_d2 ПЛАН
                    Rd2 = (L_star_fact - fact_Xf) / (nnt * (Lf1_fact + Lf22_fact))  # R_d2

                    Ra_plan = 1 / (
                            nnt * plan_Xf * (1 / Lf1_plan + 1 / Lf2_plan)) + c_plan / Kf_plan / wf_plan  # R_a ПЛАН
                    Ra = 1 / (nnt * fact_Xf * (1 / Lf1_fact + 1 / Lf2_fact)) + c_fact / Kf_fact / wf_fact  # R_a

                    Rb_plan = Ld * (Lf1_plan + Lf2_plan) / c_plan  # R_b ПЛАН
                    Rb = Ld * (Lf1_fact + Lf2_fact) / c_fact  # R_b

                    Rd_plan = (L_star_plan - plan_Xf) / (nnt * (Lf1_plan + Lf2_plan))  # R_d ПЛАН
                    Rd = (L_star_fact - fact_Xf) / (nnt * (Lf1_fact + Lf2_fact))  # R_d

                # method = 1 - через эллипс
                # method = 2 - через прямоугольник и круг
                if formula_type == 2:
                    if method == 1:
                        Reh = ((KP_Y_fact / 2) * Re) ** 0.5  # Эффективный радиус ГС, Reh
                    else:
                        Reh = (((fact_Xf * (2 * Re)) + (math.pi * Re ** 2)) / math.pi) ** 0.5

                    a_dgoshi = l_fact / 2 * (0.5 + math.sqrt(0.25 + (2 * Reh / l_fact) ** 4)) ** 0.5  # a_джоши

                n_stad = n_mgrp_plan  # кол-во стадий

                Q_water = objects_info[keys[i]][31] + 1
                iterator = 3

                while objects_info[keys[i]][31] < Q_water and iterator <= len(objects_info[keys[i]][18]):

                    P_zab_for_k = objects_info[keys[i]][21][0:iterator]
                    P_pl_for_k = objects_info[keys[i]][20][0:iterator]
                    B_for_k = objects_info[keys[i]][19][0:iterator]
                    Q_w_for_k = objects_info[keys[i]][18][0:iterator]

                    K_array = list()
                    for p in range(len(Q_w_for_k)):

                        P_pl_2_months = P_pl_for_k[p]  # Pпл ЗАПУСКНЫЕ (2 мес)
                        P_zab_2_months = P_zab_for_k[p]  # Pзаб ЗАПУСКНЫЕ (2 мес)
                        B_2_months = B_for_k[p]  # % ЗАПУСКНЫЕ (2 мес)
                        Q_w_2_months = Q_w_for_k[p]  # Qж ЗАПУСКНЫЕ (2 мес)

                        d_p = P_pl_2_months - P_zab_2_months  # dp

                        viscosity_start = viscosity_liq_finder(B_2_months / 100, viscosity_oil, viscosity_water, B0)
                        # Вязкость жидкости на ЗАПУСКЕ
                        # print('viscosity_start: ' + str(viscosity_start))

                        if formula_type == 1:
                            try:
                                K = (Q_w_2_months / (1.7054 * 0.01 * d_p / viscosity_start / B0)) / (
                                        2 / (Ra2 / (1 + Ra2 * Rb2) + Rd2) + (n_mgrp_fact - 2) / (
                                            Ra / (1 + Ra * Rb) + Rd))
                            except ZeroDivisionError:
                                K = ''

                        elif formula_type == 2:

                            try:
                                K = (Q_w_2_months / (P_pl_2_months - P_zab_2_months) / nnt * viscosity_start * B0 *
                                     (math.log(
                                         (a_dgoshi + math.sqrt(a_dgoshi * a_dgoshi - (l_fact / 2) ** 2)) / (l_fact / 2))
                                      + coeff_anizotropii ** 0.5 * nnt / l_fact * math.log(
                                                 coeff_anizotropii ** 0.5 * nnt / (
                                                             coeff_anizotropii ** 0.5 + 1) / rw) + 0) * 18.41)
                            except ZeroDivisionError:
                                K = ''

                        elif formula_type == 3:
                            try:
                                if P_zab_2_months > P_nas and P_pl_2_months > P_nas:
                                    K = Q_w_2_months * 18.4 * B0 * viscosity_start * (
                                                math.log(Re / rw) - 0.75 + S_nns) / (
                                                nnt * (P_pl_2_months - P_zab_2_months))
                                else:
                                    if P_zab_2_months <= P_nas:
                                        K = Q_w_2_months * 18.4 * B0 * viscosity_start * (
                                                    math.log(Re / rw) - 0.75 + S_nns) / (
                                                    nnt * (P_pl_2_months - P_nas + P_nas / 1.8 * (
                                                    1 - 0.2 * P_zab_2_months / P_nas - 0.8 * (
                                                        P_zab_2_months / P_nas) ** 2)))
                                    else:
                                        K = Q_w_2_months * 18.4 * B0 * viscosity_start * (
                                                    math.log(Re / rw) - 0.75 + S_nns) / (
                                                    nnt * P_pl_2_months / 1.8 * (
                                                        1 - 0.2 * P_zab_2_months / P_pl_2_months - 0.8 * (
                                                        P_zab_2_months / P_pl_2_months) ** 2))

                            except ZeroDivisionError:
                                K = ''

                        if type(K) != str:
                            if K > 0:
                                K_array.append(K)

                    K_true = list()
                    for p in range(len(K_array) - 1, 0, -1):
                        a = 0
                        for s in range(p - 1, -1, -1):
                            if (K_array[p] - K_array[s]) / K_array[p] > 100:
                                a += 1
                        if a != len(K_array):
                            K_true.append(K_array[p])

                    K_array = K_true
                    if len(K_array) != 0:
                        K = sum(K_array) / len(K_array)
                    else:
                        K = ''

                    if type(K) != str:
                        if formula_type == 1:
                            try:
                                Q_water = 0.017054 * (P_pl_now - P_zab_now) / viscosity_pot / B0 * (
                                        2 / (1 / K * (Ra2_plan / (1 + Ra2_plan * Rb2_plan) + Rd2_plan)) + (n_stad - 2) /
                                        (1 / K * (Ra_plan / (1 + Ra_plan * Rb_plan) + Rd_plan))) * K_good
                            except ZeroDivisionError:
                                Q_water = 0

                        elif formula_type == 2:
                            try:
                                Q_water = (K_good * K * (P_pl_now - P_zab_now) * nnt) / (viscosity_pot * B0 * (
                                        math.log((a_dgoshi + math.sqrt(a_dgoshi * a_dgoshi - (l_fact / 2) ** 2)) / (
                                                l_fact / 2)) + coeff_anizotropii ** 0.5 * nnt / l_fact * math.log(
                                    coeff_anizotropii ** 0.5 * nnt / (coeff_anizotropii ** 0.5 + 1) / rw) + 0) * 18.41)
                            except ZeroDivisionError:
                                Q_water = 0

                        elif formula_type == 3:
                            try:
                                if P_zab_now > P_nas and P_pl_now > P_nas:
                                    Q_water = K * nnt / (
                                            18.4 * viscosity_pot * B0 * (math.log(Re / rw) - 0.75 + S_nns)) * (
                                                      P_pl_now - P_zab_now)
                                else:
                                    if P_zab_now <= P_nas:
                                        Q_water = K * nnt / (
                                                    18.4 * viscosity_pot * B0 * (math.log(Re / rw) - 0.75 + S_nns)) * \
                                                  (P_pl_now - P_nas + P_nas / 1.8 * (
                                                              1 - 0.2 * P_zab_now / P_nas - 0.8 * (
                                                                  P_zab_now / P_nas) ** 2))
                                    else:
                                        Q_water = K * nnt / (
                                                    18.4 * viscosity_pot * B0 * (math.log(Re / rw) - 0.75 + S_nns)) * \
                                                  (P_pl_now / 1.8 * (1 - 0.2 * P_zab_now / P_pl_now - 0.8 * (
                                                              P_zab_now / P_pl_now) ** 2)) * K_good

                            except ZeroDivisionError:
                                Q_water = 0
                    else:
                        Q_water = objects_info[keys[i]][31] + 1

                    iterator += 1

                if iterator > 6:
                    objects_info[keys[i]][37] = 'Обрискования по последнему фраку'

                objects_info[keys[i]][33] = K
                if type(K) == str:
                    objects_info[keys[i]][34] = ''
                    Q_water = ''
                else:
                    objects_info[keys[i]][34] = Q_water

                B_now = objects_info[keys[i]][24]
                summer = 0
                if B_now >= 50:
                    summer = 10
                elif 20 <= B_now < 50:
                    summer = 15
                elif B_now < 20:
                    summer = 25

                if (B_now + summer) > 100:
                    objects_info[keys[i]][35] = 100
                else:
                    objects_info[keys[i]][35] = round(B_now + summer, 2)

                ro_oil = float(objects_info[keys[i]][12])

                try:
                    objects_info[keys[i]][36] = round(Q_water * ro_oil * (100 - objects_info[keys[i]][35]) / 100, 2)
                except TypeError:
                    objects_info[keys[i]][36] = ''

    return objects_info

