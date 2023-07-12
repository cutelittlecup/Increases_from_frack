import pandas as pd


def data_from_df(array_with_file_names, text):
    L = pd.read_excel('L.xlsx')
    Frack = pd.read_excel('ФРАК.xls', header=1)
    PVT = pd.read_excel('PVT.xlsx')
    H = pd.read_excel('H.xlsx', header=1)
    New_strat = pd.read_excel('Новая стратегия.xls', header=1)
    TR = pd.read_excel('ТР для загрузки.xlsx')
    """
    L = pd.read_excel(array_with_file_names[0])
    Frack = pd.read_excel(array_with_file_names[2], header=1)
    PVT = pd.read_excel(array_with_file_names[3])
    H = pd.read_excel(array_with_file_names[1], header=1)
    New_strat = pd.read_excel(array_with_file_names[4], header=1)
    TR = pd.read_excel(array_with_file_names[5])
    """

    objects_info_zapasi = {}
    if array_with_file_names[6] != 0:
        Zapasi = pd.read_excel(array_with_file_names[6], header=11)
        Zapasi.columns = map(str.lower, Zapasi.columns)
        stock_zapasi = Zapasi['скважина'].tolist()
        oiz = Zapasi['оиз'].tolist()
        for i in range(len(stock_zapasi)):
            objects_info_zapasi[str(stock_zapasi[i])] = oiz[i]

    L.columns = map(str.lower, L.columns)
    Frack.columns = map(str.lower, Frack.columns)
    PVT.columns = map(str.lower, PVT.columns)
    H.columns = map(str.lower, H.columns)
    New_strat.columns = map(str.lower, New_strat.columns)
    TR.columns = map(str.lower, TR.columns)

    global objects_info
    objects_info = {}
    global objects_info_for_graphs
    objects_info_for_graphs = {}
    objects_dates_for_graphs = {}

    # удаляем строки, где нет заполненных данных по интересующим нас столбцам
    L = L.dropna(subset=['гс/ннс']).reset_index(drop=True)

    # заносим в массивы данные для объектов
    field = L['м-е'].tolist()
    formation = L['пласт'].tolist()
    stock = L['№ скважины'].tolist()
    type_stock = L['гс/ннс'].tolist()

    # создаём словари для графиков и основной словарь (заносим тип скважины)
    form_and_stock = list()
    field_and_form = list()
    for i in range(len(field)):
        try:
            stock[i] = int(stock[i])
        except ValueError:
            stock[i] = stock[i]
        field_and_form.append(str(field[i]) + '_' + str(formation[i]))
        form_and_stock.append(str(formation[i]) + '_' + str(stock[i]))
        objects_info[str(field[i]) + '_' + str(formation[i]) + '_' + str(stock[i])] = [type_stock[i]]  # тип скважины [0]
        objects_info_for_graphs[str(field[i]) + '_' + str(formation[i]) + '_' + str(stock[i])] = [stock[i]]
        objects_dates_for_graphs[str(field[i]) + '_' + str(formation[i]) + '_' + str(stock[i])] = [stock[i]]

    for i in range(len(stock)):
        stock[i] = str(stock[i])

    keys = list(objects_info.keys())
    try:
        Frack['месторождение'] = Frack['месторождение'].fillna(8220748620877487)
        field_Frack = Frack['месторождение'].tolist()
    except KeyError:
        Frack['unnamed: 0'] = Frack['unnamed: 0'].fillna(8220748620877487)  # i am nan
        field_Frack = Frack['unnamed: 0'].tolist()
    Frack['номер скважины'] = Frack['номер скважины'].fillna(8220748620877487)  # i am nan

    Frack['м пр'] = Frack['м пр'].fillna(0)
    number_Frack = Frack['номер скважины'].tolist()

    form_Frack = Frack['пласт'].tolist()
    Mpr_stock_Frack = Frack['м пр'].tolist()
    date_Frack = Frack['дата'].tolist()

    scep = list()
    new_stock_Frack = {}
    Mpr_Frack = {}
    for i in range(len(number_Frack)):
        try:
            number_Frack[i] = int(number_Frack[i])
        except ValueError:
            number_Frack[i] = number_Frack[i]
        if str(field_Frack[i]) + '_' + str(form_Frack[i]) + '_' + str(number_Frack[i]) not in list(new_stock_Frack.keys()):
            new_stock_Frack[str(field_Frack[i]) + '_' + str(form_Frack[i]) + '_' + str(number_Frack[i])] = []
            Mpr_Frack[str(field_Frack[i]) + '_' + str(form_Frack[i]) + '_' + str(number_Frack[i])] = []

        if number_Frack[i] == 8220748620877487 or field_Frack[i] == 8220748620877487:
            if number_Frack[i] == 8220748620877487:
                number_Frack[i] = number_Frack[i - 1]
            else:
                field_Frack[i] = field_Frack[i - 1]
            if str(field_Frack[i]) + '_' + str(form_Frack[i]) + '_' + str(number_Frack[i]) not in list(
                    new_stock_Frack.keys()):
                new_stock_Frack[str(field_Frack[i]) + '_' + str(form_Frack[i]) + '_' + str(number_Frack[i])] = [i]
                Mpr_Frack[str(field_Frack[i]) + '_' + str(form_Frack[i]) + '_' + str(number_Frack[i])] = []
        else:
            new_stock_Frack[str(field_Frack[i]) + '_' + str(form_Frack[i]) + '_' + str(number_Frack[i])].append(i)

        new_stock_Frack[str(field_Frack[i]) + '_' + str(form_Frack[i]) + '_' + str(number_Frack[i])].append(date_Frack[i])
        Mpr_Frack[str(field_Frack[i]) + '_' + str(form_Frack[i]) + '_' + str(number_Frack[i])].append(Mpr_stock_Frack[i])
        scep.append(str(field_Frack[i]) + '_' + str(form_Frack[i]) + '_' + str(number_Frack[i]))

    Frack['номер скважины'] = scep
    keys_new_stock_Frack = list(new_stock_Frack.keys())
    date_Frack_new = {}
    Mpr_Frack_new = {}
    for i in range(len(keys_new_stock_Frack)):
        if keys_new_stock_Frack[i] in keys:
            date_Frack_new[keys_new_stock_Frack[i]] = []
            Mpr_Frack_new[keys_new_stock_Frack[i]] = []
            info = new_stock_Frack[keys_new_stock_Frack[i]]
            info_pr = Mpr_Frack[keys_new_stock_Frack[i]]
            while len(info) != 0:
                counter_blocks = 0
                counter_steps = list()
                counter = 0
                for j in range(len(info)):
                    try:
                        int(info[j])
                        counter_blocks += 1
                        counter_steps.append(counter)
                        counter = 0
                    except TypeError:
                        counter += 1
                if counter_blocks > 1:
                    date_Frack_new[keys_new_stock_Frack[i]].append(info[1:counter_steps[1] + 1])
                    Mpr_Frack_new[keys_new_stock_Frack[i]].append(info_pr[0:counter_steps[1]])
                    del info[0:counter_steps[1] + 1]
                    del info_pr[0:counter_steps[1]]
                else:
                    date_Frack_new[keys_new_stock_Frack[i]].append(info[1:len(info)])
                    Mpr_Frack_new[keys_new_stock_Frack[i]].append(info_pr[0:len(info_pr)])
                    break

    date_Frack_new_keys = list(date_Frack_new.keys())

    TR = TR.fillna(0)
    field_TR = TR['месторождение'].tolist()
    form_TR = TR['объекты работы'].tolist()
    stock_TR = TR['№ скважины'].tolist()
    scepka_to_TR = list()
    scepka_TR = list()
    for i in range(len(form_TR)):
        try:
            stock_TR[i] = int(stock_TR[i])
        except ValueError:
            stock_TR[i] = stock_TR[i]
        scepka_to_TR.append(str(field_TR[i]) + '_' + str(form_TR[i]) + '_' + str(stock_TR[i]))
        if scepka_to_TR[i] not in scepka_TR:
            scepka_TR.append(scepka_to_TR[i])

    TR['месторождение'] = scepka_to_TR
    objects_info_problems = {}
    date_TR_new = {}
    for i in range(len(scepka_TR) - 1, -1, -1):
        df = TR[TR['месторождение'] == scepka_TR[i]].reset_index(drop=True)
        df = df[df['характер работы'] == 'НЕФ'].reset_index(drop=True)
        try:
            df_first = df.iloc[0]
            date_TR_new[scepka_TR[i]] = df_first['дата']
        except IndexError:
            if scepka_TR[i] not in list(objects_info_problems.keys()):
                objects_info_problems[scepka_TR[i]] = 'Нет данных в файле "ТР"'
            scepka_TR.pop(i)

    date_TR_new_keys = list(date_TR_new.keys())
    New_strat = New_strat[['месторождение', 'скважина', 'тип', 'объект разработки до гтм', 'внр.1']]
    New_strat = New_strat[New_strat['тип'] == 'ГРП'].reset_index(drop=True)
    field_New_strat = New_strat['месторождение'].tolist()
    stock_New_strat = New_strat['скважина'].tolist()
    form_New_strat = New_strat['объект разработки до гтм'].tolist()

    scepka_New_strat = list()
    scepka_to_New_strat = list()
    for i in range(len(form_New_strat)):
        try:
            stock_New_strat[i] = int(stock_New_strat[i])
        except ValueError:
            stock_New_strat[i] = stock_New_strat[i]
        a = str(field_New_strat[i]) + '_' + str(form_New_strat[i]) + '_' + str(stock_New_strat[i])
        scepka_to_New_strat.append(a)
        if a not in scepka_New_strat and a in keys:
            scepka_New_strat.append(a)

    New_strat['месторождение'] = scepka_to_New_strat
    refracks = {}
    for i in range(len(scepka_New_strat)):
        if scepka_New_strat[i] in date_TR_new_keys:
            refracks[scepka_New_strat[i]] = []
            df = New_strat[New_strat['месторождение'] == scepka_New_strat[i]]
            date_New_strat = df['внр.1'].tolist()
            A = []
            for j in range(len(date_New_strat) - 1, -1, -1):
                a = (date_TR_new[scepka_New_strat[i]] - date_New_strat[j]).days
                a = int(a / 31)
                A.append(a)
            A.reverse()
            counter = 0
            first_refrack = list()
            for j in range(len(A)):
                if -1 < A[j] <= 6:
                    first_refrack.append(date_New_strat[j])
                elif A[j] < -1:
                    counter += 1
            first_refrack_not_grp = list()
            if len(first_refrack) == 0:
                if counter != 0:
                    counter = counter - 1
                for j in range(len(A)):
                    if A[j] < -1:
                        first_refrack_not_grp.append(date_New_strat[j])
            if len(first_refrack) == 0 and len(first_refrack_not_grp) == 0:
                first_refrack.append('-')
            elif len(first_refrack) == 0 and len(first_refrack_not_grp) != 0:
                first_refrack.append(min(first_refrack_not_grp))
            refracks[scepka_New_strat[i]].append(first_refrack[len(first_refrack) - 1])
            refracks[scepka_New_strat[i]].append(counter)

    refracks_keys = list(refracks.keys())

    for i in range(len(date_TR_new_keys)):
        if date_TR_new_keys[i] in date_Frack_new_keys and date_TR_new_keys[i] not in refracks_keys:
            min_array = list()
            for j in range(len(date_Frack_new[date_TR_new_keys[i]])):
                min_array.append(min(date_Frack_new[date_TR_new_keys[i]][0]))
            if len(min_array) != 0:
                val_min, idx_min = min((val_min, idx_min) for (idx_min, val_min) in enumerate(min_array))
                if (date_TR_new[date_TR_new_keys[i]] - val_min).days / 31 <= 6:
                    refracks[date_TR_new_keys[i]] = [val_min, len(date_Frack_new[date_TR_new_keys[i]]) - 1]

    refracks_keys = list(refracks.keys())
    """
    for i in range(len(date_TR_new_keys)):
        if date_TR_new_keys[i] in date_Frack_new_keys and date_TR_new_keys[i] not in refracks_keys:
            if date_TR_new_keys[i] not in list(objects_info_problems.keys()):
                objects_info_problems[date_TR_new_keys[i]] = 'Нет данных в файле "Новая стратегия"'
    """

    grp_first = {}
    Mpr_last = {}
    for i in range(len(keys_new_stock_Frack)):
        if keys_new_stock_Frack[i] in refracks_keys:
            if refracks[keys_new_stock_Frack[i]][0] != '-':
                for j in range(len(date_Frack_new[keys_new_stock_Frack[i]])):
                    for k in range(len(date_Frack_new[keys_new_stock_Frack[i]][j]) - 1, -1, -1):
                        if (refracks[keys_new_stock_Frack[i]][0] - date_Frack_new[keys_new_stock_Frack[i]][j][
                            k]).days / 31 > 6:
                            date_Frack_new[keys_new_stock_Frack[i]][j].pop(k)
                            Mpr_Frack_new[keys_new_stock_Frack[i]][j].pop(k)
                min_array = list()
                for j in range(len(date_Frack_new[keys_new_stock_Frack[i]])):
                    if len(date_Frack_new[keys_new_stock_Frack[i]][j]) != 0:
                        min_array.append(min(date_Frack_new[keys_new_stock_Frack[i]][j]))
                    else:
                        if len(date_Frack_new[keys_new_stock_Frack[i]]) < 2:
                            print('')
                            break

                if len(min_array) != 0:
                    val_min, idx_min = min((val_min, idx_min) for (idx_min, val_min) in enumerate(min_array))

                    grp_first[keys_new_stock_Frack[i]] = len(date_Frack_new[keys_new_stock_Frack[i]][idx_min])
                    Mpr_last[keys_new_stock_Frack[i]] = sum(Mpr_Frack_new[keys_new_stock_Frack[i]][idx_min]) / len(
                        Mpr_Frack_new[keys_new_stock_Frack[i]][idx_min])

    try:
        l_L = L['l'].tolist()  # l - длина скважины
    except KeyError:
        x = L['координата x'].tolist()
        y = L['координата y'].tolist()
        x_tr = L['координата забоя х (по траектории)'].tolist()
        y_tr = L['координата забоя y (по траектории)'].tolist()
        l_L = list()
        for i in range(len(x)):
            try:
                l_L.append(((float(x[i]) - float(x_tr[i])) ** 2 + (float(y_tr[i]) - float(y[i])) ** 2) ** 0.5)
            except TypeError:
                if objects_info[keys[i]][0] == 'ГС':
                    if keys[i] not in list(objects_info_problems.keys()):
                        objects_info_problems[keys[i]] = 'Нет данных в файле L'

    grp_first_keys = list(grp_first.keys())
    for i in range(len(keys)):
        if keys[i] in grp_first_keys:
            objects_info[keys[i]].append(grp_first[keys[i]])  # стадии ГРП [1]
            objects_info[keys[i]].append(Mpr_last[keys[i]])  # масса пропанта [2]
        else:
            objects_info[keys[i]].append(0)
            objects_info[keys[i]].append(0)
            # if (keys[i] not in scepka_New_strat) and (keys[i] not in list(objects_info_problems.keys())):
            # objects_info_problems[keys[i]] = 'Нет данных по ГРП в файле "Новая стратегия"'
            if objects_info[keys[i]] == 'ГС':
                if (keys[i] not in date_Frack_new_keys) and (keys[i] in scepka_New_strat) and (
                        keys[i] not in list(objects_info_problems.keys())):
                    objects_info_problems[keys[i]] = 'Нет данных по 1 ГРП в файле "ФРАК"'

        objects_info[keys[i]].append(l_L[i])  # длина скважины [3]

    field_PVT = PVT['месторождение'].tolist()
    form_PVT = PVT['пласт'].tolist()
    B0_PVT = PVT['bo'].tolist()
    a_Xf_PVT = PVT['а xf'].tolist()
    b_Xf_PVT = PVT['b xf'].tolist()
    S_nns_PVT = PVT['sрасч ннс'].tolist()
    M_plan_PVT = PVT['мпр среднее на стадию'].tolist()
    P_nas_PVT = PVT['рb'].tolist()
    viscosity_oil_PVT = PVT['вязкость нефти'].tolist()
    viscosity_water_PVT = PVT['вязкость воды'].tolist()
    ro_PVT = PVT['ro'].tolist()
    K_por_PVT = PVT['пористость'].tolist()
    KIN_PVT = PVT['кин'].tolist()
    K_nn_PVT = PVT['кнн'].tolist()

    params_from_PVT = {}
    for i in range(len(field_PVT)):
        params_from_PVT[str(field_PVT[i]) + '_' + str(form_PVT[i])] = [B0_PVT[i], a_Xf_PVT[i], b_Xf_PVT[i], S_nns_PVT[i],
                                                                       M_plan_PVT[i], P_nas_PVT[i], viscosity_oil_PVT[i],
                                                                       viscosity_water_PVT[i], ro_PVT[i],
                                                                       K_por_PVT[i], KIN_PVT[i], K_nn_PVT[i]]

        # B0 [4]
        # a_xF [5]
        # b_Xf [6]
        # S_nns [7]
        # Масса пропанта плановая [8]
        # Pнас [9]
        # вязкость нефти [10]
        # вязкость воды [11]
        # плотность нефти [12]
        # Коэффициент пористости [13]
        # КИН [14]
        # Кнн [15]

    PVT_keys = list(params_from_PVT.keys())
    for i in range(len(field_and_form)):
        if field_and_form[i] in PVT_keys:
            for j in range(len(params_from_PVT[field_and_form[i]])):
                objects_info[keys[i]].append(params_from_PVT[field_and_form[i]][j])
        else:
            if keys[i] not in list(objects_info_problems.keys()):
                objects_info_problems[keys[i]] = 'Нет данных в файле "PVT"'
            for j in range(12):
                objects_info[keys[i]].append(0)

    field_H = H['м-е'].tolist()
    form_H = H['пласт'].tolist()
    stock_H = H['скважина - забой'].tolist()
    nnt_H = H['ннт'].tolist()
    onnt_H = H['оннт'].tolist()

    params_from_H = {}
    for i in range(len(field_H)):
        try:
            stock_H[i] = int(stock_H[i])
        except ValueError:
            stock_H[i] = stock_H[i]
        params_from_H[str(field_H[i]) + '_' + str(form_H[i]) + '_' + str(stock_H[i])] = [nnt_H[i], onnt_H[i]]

    H_keys = list(params_from_H.keys())
    for i in range(len(keys)):
        if keys[i] in H_keys:
            for j in range(len(params_from_H[keys[i]])):
                objects_info[keys[i]].append(params_from_H[keys[i]][j])
                # ННТ [16]
                # ОННТ [17]
        else:
            # если не прогружен файл с добычей
            objects_info_problems[keys[i]] = 'Нет данных в файле "H"'
            for j in range(2):
                objects_info[keys[i]].append(0)

    date = {}
    for i in range(len(scepka_New_strat)):
        df = New_strat[New_strat['месторождение'] == scepka_New_strat[i]].reset_index(drop=True)
        data = df['внр.1'].tolist()
        data1 = data[0]
        date[scepka_New_strat[i]] = [datetime.date(data1.year, data1.month, data1.day)]
        data2 = data[len(data) - 1]
        date[scepka_New_strat[i]].append(datetime.date(data2.year, data2.month, data2.day))

    date_keys = date.keys()
    info = {}
    info_for_graphs = {}
    dates_for_graphs = {}
    for i in range(len(scepka_TR)):
        df = TR[TR['месторождение'] == scepka_TR[i]].reset_index(drop=True)
        df = df[df['характер работы'] == 'НЕФ'].reset_index(drop=True)
        try:
            xr_df = df.loc[len(df) - 9:len(df)].reset_index(drop=True)
        except IndexError:
            xr_df = df
        xr_df_array = xr_df['состояние'].tolist()
        if 'ПЬЕЗ.' not in xr_df_array and 'ПЬЕЗ' not in xr_df_array:
            try:
                xr_df = df.loc[len(df) - 6:len(df)].reset_index(drop=True)
            except IndexError:
                xr_df = df
            xr_df_array = xr_df['состояние'].tolist()
        df['пластовое давление (тр), атм'] = df['пластовое давление (тр), атм'].fillna(0)
        plast_p = df['пластовое давление (тр), атм'].tolist()
        if plast_p[len(plast_p) - 1] == 0:
            for j in range(len(plast_p) - 2, -1, -1):
                if plast_p[j] != 0:
                    plast_p[len(plast_p) - 1] = plast_p[j]
                    break
        for j in range(len(plast_p) - 2, -1, -1):
            if plast_p[j] == 0:
                if plast_p[j + 1] != 0:
                    plast_p[j] = plast_p[j + 1]

        df['пластовое давление (тр), атм'] = plast_p

        Q_zhid_for_graphs = df['дебит жидкости (тр), м3/сут'].tolist()
        Q_nef_for_graphs = df['дебит нефти (тр), т/сут'].tolist()
        Wat_for_graphs = df['обводненность (тр), % (объём)'].tolist()
        Ppl_for_graphs = df['пластовое давление (тр), атм'].tolist()
        Pzab_for_graphs = df['забойное давление (тр), атм'].tolist()
        date_for_graphs = df['дата'].tolist()

        info_for_graphs[scepka_TR[i]] = [Q_zhid_for_graphs, Wat_for_graphs, Ppl_for_graphs, Pzab_for_graphs,
                                         Q_nef_for_graphs]
        dates_for_graphs[scepka_TR[i]] = [date_for_graphs]
        if (len(df) != 0) and (
                'РАБ.' in xr_df_array or 'РАБ' in xr_df_array or 'НАК.' in xr_df_array or 'НАК' in xr_df_array):
            df_last = df.loc[len(df) - 1]
            dates = df['дата'].tolist()
            k = 0
            if scepka_TR[i] not in date_keys:
                date_n_s = dates[0]
            else:
                date_n_s = date[scepka_TR[i]][0]
            for j in range(len(dates)):
                if date_n_s.year == dates[j].year and date_n_s.month == dates[j].month:
                    k = 1
                    df_new = df.iloc[j:len(df)].reset_index(drop=True)
                    break
                if k == 0 and j == len(dates) - 1:
                    df_new = df

            if scepka_TR[i] not in date_keys:
                date_n_s = dates[0]
            else:
                date_n_s = date[scepka_TR[i]][1]
            for j in range(len(dates)):
                if date_n_s.year == dates[j].year and date_n_s.month == dates[j].month:
                    k = 1
                    df_last_grp = df.iloc[j]
                    break
                if k == 0 and j == len(dates) - 1:
                    df_last_grp = df.iloc[0]

            info[scepka_TR[i]] = [df_new['дебит жидкости (тр), м3/сут'].tolist(),
                                  df_new['обводненность (тр), % (объём)'].tolist(),
                                  df_new['пластовое давление (тр), атм'].tolist(),
                                  df_new['забойное давление (тр), атм'].tolist(),
                                  df_new['дебит нефти (тр), т/сут'].tolist(),
                                  float(df_last['дебит жидкости (тр), м3/сут']),
                                  float(df_last['обводненность (тр), % (объём)']),
                                  float(df_last['пластовое давление (тр), атм']),
                                  float(df_last['забойное давление (тр), атм']),
                                  float(df_last['дебит нефти (тр), т/сут']),
                                  df_last['характер работы'],
                                  df_last['состояние'],
                                  datetime.date(df_last['дата'].year, df_last['дата'].month, df_last['дата'].day),
                                  float(df_last_grp['дебит жидкости (тр), м3/сут'])]
            # Дебит жидкости (ТР) [18]
            # Обводнённость (ТР) [19]
            # Пластовое давление (ТР) [20]
            # Забойное давление (ТР) [21]
            # Дебит нефти (ТР) [22]
            # Дебит жидкости (последний) [23]
            # Обводнённость (последняя) [24]
            # Пластовое давление (последнее) [25]
            # Забойное давление (последнее) [26]
            # Дебит нефти (последний) [27]
            # Характер работы [28]
            # Состояние [29]
            # Дата [30]
            # Дебит жидкости (последний грп) [31]
        else:
            sost = ''
            x_w = ''
            date_last = ''
            if (len(df) != 0) and (scepka_TR[i] not in list(objects_info_problems.keys())):
                df_last = df.loc[len(df) - 1]
                sost = df_last['состояние']
                x_w = df_last['характер работы']
                date_last = datetime.date(df_last['дата'].year, df_last['дата'].month, df_last['дата'].day)
                objects_info_problems[scepka_TR[i]] = 'Скважина в бездействующем фонде'
            if (len(df) == 0) and (scepka_TR[i] not in list(objects_info_problems.keys())):
                objects_info_problems[scepka_TR[i]] = 'Нет данных в файле "ТР"'
            info[scepka_TR[i]] = [[], [], [], [], [], '', '', '', '', '', x_w, sost, date_last, '']

    info_keys = list(info.keys())
    info_keys_for_graphs = list(info_for_graphs.keys())
    for i in range(len(keys)):
        if keys[i] in info_keys:
            for j in range(len(info[keys[i]])):
                objects_info[keys[i]].append(info[keys[i]][j])
        else:
            if keys[i] not in list(objects_info_problems.keys()):
                objects_info_problems[keys[i]] = 'Нет данных в файле "ТР"'
            for j in range(5):
                objects_info[keys[i]].append([])
            for j in range(9):
                objects_info[keys[i]].append('')

        if keys[i] in info_keys_for_graphs:
            objects_info_for_graphs[keys[i]].append(info_for_graphs[keys[i]])
            objects_dates_for_graphs[keys[i]].append(dates_for_graphs[keys[i]])

    for i in range(len(keys) - 1, -1, -1):
        if objects_info[keys[i]][12] == 0:
            del objects_info[keys[i]]
            del objects_info_for_graphs[keys[i]]
            del objects_dates_for_graphs[keys[i]]
            del objects_info_problems[keys[i]]
            form_and_stock.pop(i)
            field_and_form.pop(i)
            keys.pop(i)

    for i in range(len(objects_info)):
        if keys[i] in refracks_keys:
            objects_info[keys[i]].append(refracks[keys[i]][1])
        else:
            objects_info[keys[i]].append(0)
            # Количество повторных ГРП [32]
    for i in range(len(keys)):
        for j in range(6):
            objects_info[keys[i]].append('')  # Кпрон [33]
            # Qжид [34]

    return objects_info