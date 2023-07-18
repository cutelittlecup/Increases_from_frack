import pandas as pd
from excel_to_df import df_reader


class Well:
    def __init__(self, well_name, formation, field):
        self.well_name = well_name
        self.formation = formation
        self.field = field
        self.well_type = None
        self.propping_agent_mass = None
        self.frack_stages = None
        self.first_date_tr = None
        self.first_frack = None
        self.count_fracks = None
        self.l = None
        self.B0 = None
        self.a_Xf = None
        self.b_Xf = None
        self.S_nns = None
        self.M_plan = None
        self.P_nas = None
        self.viscosity_oil = None
        self.viscosity_water = None
        self.ro = None
        self.K_por = None
        self.KIN = None
        self.K_nn = None
        self.nnt = None
        self.onnt = None

    def find_l(self, x, x_tr, y, y_tr):
        self.l = ((x - x_tr) ** 2 + (y - y_tr) ** 2) ** 0.5


def data_from_df(files, text):
    files = ['L.xlsx', 'ФРАК.xls', 'PVT.xlsx', 'H.xlsx', 'Новая стратегия.xls', 'ТР для загрузки.xlsx']
    L, Frack, PVT, H, New_strat, TR = df_reader(files)

    objects_info = {}

    # удаляем строки, где нет заполненных данных по интересующим нас столбцам
    L = L.dropna(subset=['гс/ннс']).reset_index(drop=True)
    L[['м-е', 'пласт', '№ скважины']] = L[['м-е', 'пласт', '№ скважины']].astype(str)

    try:
        Frack['месторождение']
    except KeyError:
        Frack.rename(columns={'unnamed: 0': 'месторождение'}, inplace=True)

    Frack[['месторождение', 'номер скважины']] = Frack[['месторождение', 'номер скважины']].fillna(method='ffill')
    Frack['м пр'] = Frack['м пр'].fillna(0)
    Frack[['месторождение', 'пласт', 'номер скважины']] = Frack[['месторождение', 'пласт', 'номер скважины']].astype(
        str)
    Frack['сцепка'] = Frack['месторождение'] + '_' + Frack['пласт'] + '_' + Frack['номер скважины']

    TR = TR.fillna(0)
    TR[['месторождение', 'объекты работы', '№ скважины']] = TR[
        ['месторождение', 'объекты работы', '№ скважины']].astype(str)
    TR['сцепка'] = TR['месторождение'] + '_' + TR['объекты работы'] + '_' + TR['№ скважины']

    New_strat = New_strat[['месторождение', 'скважина', 'тип', 'объект разработки до гтм', 'внр.1']]
    New_strat = New_strat[New_strat['тип'] == 'ГРП'].reset_index(drop=True)
    New_strat[['месторождение', 'объект разработки до гтм', 'скважина']] = New_strat[
        ['месторождение', 'объект разработки до гтм', 'скважина']].astype(str)
    New_strat['сцепка'] = New_strat['месторождение'] + '_' + New_strat['объект разработки до гтм'] + '_' + New_strat[
        'скважина']

    L['сцепка'] = L['м-е'] + '_' + L['пласт'] + '_' + L['№ скважины']
    keys = L['сцепка'].unique()

    for key in keys:
        sub_l = L[L['сцепка'] == key]
        objects_info[key] = Well(sub_l['№ скважины'], sub_l['пласт'], sub_l['м-е'])
        objects_info[key].well_type = sub_l['гс/ннс'].item()

        try:
            objects_info[key].l = sub_l['l'].item()
        except KeyError:
            x = float(sub_l['координата x'])
            x_tr = float(sub_l['координата забоя х (по траектории)'])
            y = float(sub_l['координата y'])
            y_tr = float(sub_l['координата забоя y (по траектории)'])
            objects_info[key].find_l(x, x_tr, y, y_tr)

        # objects_info_for_graphs[key] = Well(row['№ скважины'], row['пласт'], row['м-е'])
        # objects_dates_for_graphs[key] = Well(row['№ скважины'], row['пласт'], row['м-е'])

        sub_pvt = PVT[(PVT['месторождение'] == sub_l['м-е']) & (PVT['пласт'] == sub_l['пласт'])]
        objects_info[key].B0 = sub_pvt['bo']
        objects_info[key].a_Xf = sub_pvt['а xf']
        objects_info[key].b_Xf = sub_pvt['b xf']
        objects_info[key].S_nns = sub_pvt['sрасч ннс']
        objects_info[key].M_plan = sub_pvt['мпр среднее на стадию']
        objects_info[key].P_nas = sub_pvt['рb']
        objects_info[key].viscosity_oil = sub_pvt['вязкость нефти']
        objects_info[key].viscosity_water = sub_pvt['вязкость воды']
        objects_info[key].ro = sub_pvt['ro']
        objects_info[key].K_por = sub_pvt['пористость']
        objects_info[key].KIN = sub_pvt['кин']
        objects_info[key].K_nn = sub_pvt['кнн']

        sub_h = H[H[key] == key]
        objects_info[key].nnt = sub_h['ннт']
        objects_info[key].onnt = sub_h['oннт']
        # тут есть какая-то связь с файлом с добычей

        sub_tr = TR[(TR['сцепка'] == key) & (TR['характер работы'] == 'НЕФ')].reset_index(drop=True)
        if len(sub_tr) != 0:
            objects_info[key].first_date_tr = sub_tr.iloc[0]['дата']

        sub_frack = Frack[Frack['сцепка'] == key]
        sub_frack['diff'] = ((objects_info[key].first_date_tr - sub_frack['дата']).dt.days / 31).astype(int)
        sub_frack.drop(index=sub_frack[sub_frack['diff'] > 6].index, inplace=True)
        if len(sub_frack) != 0:
            objects_info[key].propping_agent_mass = Frack['м пр'].summ() / len(Frack['м пр'])
            objects_info[key].frack_stages = Frack['дата'].count()

        sub_new_strat = New_strat[New_strat['сцепка'] == key].copy()
        if len(sub_new_strat) != 0:
            sub_new_strat['diff'] = ((objects_info[key].first_date_tr - sub_new_strat['внр.1']).dt.days / 31).astype(
                int)
            objects_info[key].count_fracks = (sub_new_strat['diff'] < -1).sum()
            try:
                frack_drilling = sub_new_strat.loc[(sub_new_strat['diff'] > -1) & (sub_new_strat['diff'] <= 6)].index[0]
                objects_info[key].first_frack = sub_new_strat.loc[frack_drilling, 'внр.1']
            except IndexError:
                try:
                    objects_info[key].count_fracks = objects_info[key].count_fracks - 1
                    frack = sub_new_strat.loc[(sub_new_strat['diff'] < -1)].index[0]
                    objects_info[key].first_frack = sub_new_strat.loc[frack, 'внр.1']
                except IndexError:
                    objects_info[key].first_frack = '-'

        # если скв из тех режимов есть во фраке, но нет в НС
        if objects_info[key].frack_stages is not None and objects_info[key].first_frack is None:
            stages_dates = objects_info[key].frack_stages_dates
            stages_dates['diff'] = ((objects_info[key].first_date_tr - stages_dates['дата']).dt.days / 31).astype(int)
            objects_info[key].count_fracks = 0
            try:
                objects_info[key].first_frack = stages_dates.loc[(stages_dates['diff'] > -1) & (stages_dates['diff'] <= 6)]['дата'].min()
            except IndexError:
                try:
                    objects_info[key].first_frack = stages_dates.loc[(stages_dates['diff'] < -1)]['дата'].min()
                    objects_info[key].count_fracks = 0
                except IndexError:
                    objects_info[key].first_frack = '-'

    # new_stock_Frack == [0, Timestamp('2017-05-24 00:00:00'), 1, Timestamp('2017-05-24 00:00:00'), ...]
    # Mpr_Frack == [10.0, 9.8, 10.0, 10.0, 10.0]

    # date_Frack_new == [[Timestamp('2017-05-24 00:00:00')], [Timestamp('2017-05-24 00:00:00')], ...]
    # Mpr_Frack_new == [[10.0], [9.8], [10.0], [10.0], [10.0]]

    for key in keys:
        sub_tr = TR[(TR['месторождение'] == key) & (TR['характер работы'] == 'НЕФ')].reset_index(drop=True)
        try:
            work_check = sub_tr.loc[len(sub_tr) - 9:len(sub_tr)].reset_index(drop=True)
        except IndexError:
            work_check = sub_tr
        if 'ПЬЕЗ.' not in work_check['состояние'] and 'ПЬЕЗ' not in work_check['состояние']:
            try:
                work_check = sub_tr.loc[len(sub_tr) - 6:len(sub_tr)].reset_index(drop=True)
            except IndexError:
                work_check = sub_tr
        df['пластовое давление (тр), атм'] = df['пластовое давление (тр), атм'].fillna(0)

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
        """
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
    

    objects_info_reserves = {}
    if files[6] != 0:
        reserves = pd.read_excel(files[6], header=11)
        reserves.columns = map(str.lower, reserves.columns)
        stock_reserves = reserves['скважина'].tolist()
        oiz = reserves['оиз'].tolist()
        for i in range(len(stock_reserves)):
            objects_info_reserves[str(stock_reserves[i])] = oiz[i]
    """
    return objects_info
