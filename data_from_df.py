import pandas as pd
from excel_to_df import df_reader
import formulas as f


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
        self.history = None
        self.oiz = None

    def find_l(self, x, x_tr, y, y_tr):
        self.l = ((x - x_tr) ** 2 + (y - y_tr) ** 2) ** 0.5


def data_from_df(files, text):

    files = ['L.xlsx', 'ФРАК.xls', 'PVT.xlsx', 'H.xlsx', 'Новая стратегия.xls', 'ТР для загрузки.xlsx', '']
    L, Frack, PVT, H, New_strat, TR, reserves = df_reader(files)

    objects_info = {}

    L = L.dropna(subset=['гс/ннс']).reset_index(drop=True)
    L[['м-е', 'пласт', '№ скважины']] = L[['м-е', 'пласт', '№ скважины']].astype(str)
    L['сцепка'] = L['м-е'] + '_' + L['пласт'] + '_' + L['№ скважины']

    H[['м-е', 'пласт', 'скважина - забой']] = H[['м-е', 'пласт', 'скважина - забой']].astype(str)
    H['сцепка'] = H['м-е'] + '_' + H['пласт'] + '_' + H['скважина - забой']

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

        sub_pvt = PVT[(PVT['месторождение'] == sub_l['м-е'].item()) & (PVT['пласт'] == sub_l['пласт'].item())]
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

        sub_h = H[H['сцепка'] == key]
        objects_info[key].nnt = sub_h['ннт']
        objects_info[key].onnt = sub_h['оннт']

        sub_tr = TR[(TR['сцепка'] == key) & (TR['характер работы'] == 'НЕФ')].reset_index(drop=True)
        if len(sub_tr) != 0:
            objects_info[key].first_date_tr = sub_tr.iloc[0]['дата']

            try:
                work_check = sub_tr.loc[len(sub_tr) - 9:len(sub_tr)].reset_index(drop=True)
            except IndexError:
                work_check = sub_tr
            if 'ПЬЕЗ.' not in work_check['состояние'] and 'ПЬЕЗ' not in work_check['состояние']:
                try:
                    work_check = sub_tr.loc[len(sub_tr) - 6:len(sub_tr)].reset_index(drop=True)
                except IndexError:
                    work_check = sub_tr
            sub_tr['пластовое давление (тр), атм'] = sub_tr['пластовое давление (тр), атм'].fillna(method='ffill').fillna(method='bfill')

            objects_info[key].history = sub_tr

            sub_frack = Frack[Frack['сцепка'] == key].copy()

            sub_frack['diff'] = ((objects_info[key].first_date_tr - sub_frack['дата']).dt.days / 31).astype(int)
            sub_frack.drop(index=sub_frack[sub_frack['diff'] > 6].index, inplace=True)
            if len(sub_frack) != 0:
                objects_info[key].propping_agent_mass = sub_frack['м пр'].sum() / len(sub_frack['м пр'])
                objects_info[key].frack_stages = len(sub_frack['дата'])

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
                objects_info[key].count_fracks = 0
                try:
                    objects_info[key].first_frack = sub_frack.loc[(sub_frack['diff'] > -1) & (sub_frack['diff'] <= 6)]['дата'].min()
                except IndexError:
                    try:
                        objects_info[key].first_frack = sub_frack.loc[(sub_frack['diff'] < -1)]['дата'].min()
                        objects_info[key].count_fracks = 0
                    except IndexError:
                        objects_info[key].first_frack = '-'

        if len(reserves) != 0:
            sub_reserves = reserves[reserves['скважина'] == key]
            if len(sub_reserves) != 0:
                objects_info[key].oiz = sub_reserves['оиз']

    #objects_info = f.formulas(objects_info, text)

    return objects_info
