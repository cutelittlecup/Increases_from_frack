import tkinter as tk
import pandas as pd


def to_excel(objects_info, text):
    objects_info = dict(sorted(objects_info.items(), key=lambda x: x[0]))
    keys = list(objects_info.keys())

    field_for_excel = list()
    form_for_excel = list()
    stock_for_excel = list()
    type_stock_for_excel = list()
    Q_water_now_for_excel = list()
    B_now_for_excel = list()
    Q_oil_now_for_excel = list()
    Q_water_for_excel = list()
    B_for_excel = list()
    Q_oil_for_excel = list()
    P_zab_now_for_excel = list()
    P_pl_now_for_excel = list()
    up_Q_oil_for_excel = list()
    K_pron_for_excel = list()
    risks = list()
    kandidat = list()
    x_w = list()
    sost = list()
    last_month = list()
    formula_types = list()
    grp_povt = list()
    n_stad_first = list()
    obrisk = list()

    for i in range(len(keys)):
        string = keys[i].split('_')
        field_for_excel.append(string[0])
        form_for_excel.append(string[1])
        stock_for_excel.append(string[2])
        if objects_info[keys[i]][0] == 3 or objects_info[keys[i]][0] == 'ННС':
            type_stock_for_excel.append('ННС')
        else:
            type_stock_for_excel.append('ГС')

        n_stad_first.append(objects_info[keys[i]][1])

        # try:
        # last_month.append(objects_info[keys[i]][30])
        # except AttributeError:
        #    last_month.append('')

        grp_povt.append(objects_info[keys[i]][32])

        x_w.append(objects_info[keys[i]][28])
        sost.append(objects_info[keys[i]][29])
        last_month.append(objects_info[keys[i]][30])

        for_djoshi = 1
        if objects_info[keys[i]][0] == 'ГС' or objects_info[keys[i]][0] == 1:
            if objects_info[keys[i]][1] > 1:
                Mpr_stock = objects_info[keys[i]][2]
                if Mpr_stock == 0 and objects_info[keys[i]][1] == 0 and grp_povt[i] != 0:
                    for_djoshi = 0
                if Mpr_stock == 0 and objects_info[keys[i]][1] != 0:
                    for_djoshi = 0

        if for_djoshi != 0 and (keys[i] not in list(objects_info_problems.keys())):

            if grp_povt[i] < 4:

                obrisk.append(objects_info[keys[i]][37])

                if type(objects_info[keys[i]][23]) == float:

                    if objects_info[keys[i]][0] == 1:
                        formula_types.append('Ли')
                    elif objects_info[keys[i]][0] == 2:
                        formula_types.append('Джоши')
                    elif objects_info[keys[i]][0] == 3:
                        formula_types.append('Дюпюи')
                    else:
                        formula_types.append('')

                    Q_water_now_for_excel.append(round(objects_info[keys[i]][23], 2))
                    B_now_for_excel.append(round(objects_info[keys[i]][24], 2))
                    Q_oil_now_for_excel.append(round(objects_info[keys[i]][27], 2))

                    try:
                        K_pron_for_excel.append(round(objects_info[keys[i]][33], 2))
                        Q_water_for_excel.append(round(objects_info[keys[i]][34], 2))
                    except TypeError:
                        K_pron_for_excel.append(objects_info[keys[i]][33])
                        Q_water_for_excel.append(objects_info[keys[i]][34])
                    try:
                        B_for_excel.append(round(objects_info[keys[i]][35], 2))
                    except TypeError:
                        B_for_excel.append(objects_info[keys[i]][35])

                    ro_oil = float(objects_info[keys[i]][12])

                    try:
                        Q_oil_for_excel.append(round(objects_info[keys[i]][36], 2))
                    except TypeError:
                        Q_oil_for_excel.append(objects_info[keys[i]][36])

                    try:
                        up = Q_oil_for_excel[i] - Q_oil_now_for_excel[i]
                        if up >= 0:
                            up_Q_oil_for_excel.append(round(up, 2))
                        else:
                            up_Q_oil_for_excel.append(round(0, 2))
                    except TypeError:
                        up_Q_oil_for_excel.append('')

                    # три типа рисков: по запасам, по пластовому давлению и по обводнённости
                    risk = ''
                    onnt = objects_info[keys[i]][17]
                    K_por = objects_info[keys[i]][13]
                    KIN = objects_info[keys[i]][14]
                    Knn = objects_info[keys[i]][15]
                    B0 = objects_info[keys[i]][4]
                    l_fact = objects_info[keys[i]][3]

                    if stock[i] in list(objects_info_zapasi.keys()):
                        oiz = objects_info_zapasi[stock[i]]
                        if oiz < 5000:
                            risk = risk + 'Риски по запасам'
                    else:
                        if stock_for_excel[i] == 'ГС':
                            oiz = (l_fact + 300) * 300 * onnt * K_por * Knn / B0 * ro_oil * KIN
                        else:
                            oiz = math.pi * 300 ** 2 * onnt * K_por * Knn * 1 / B0 * ro_oil
                        if oiz < 5000:
                            risk = risk + 'Риски по запасам'

                    P_pl_start_for_excel = objects_info_for_graphs[keys[i]][1][2][0]
                    P_pl_now_for_excel.append(round(objects_info[keys[i]][25], 2))
                    P_zab_now_for_excel.append(round(objects_info[keys[i]][26], 2))

                    if P_pl_now_for_excel[i] < P_pl_start_for_excel * 0.6:
                        if len(risk) != 0:
                            risk = risk + ', ' + 'риск по Pпл'
                        else:
                            risk = 'Риск по Pпл'

                    if B_now_for_excel[i] >= 70:
                        if len(risk) != 0:
                            risk = risk + ', ' + 'высокая обводнённость'
                        else:
                            risk = 'Высокая обводнённость'
                    elif Q_oil_now_for_excel[i] >= 15:
                        if len(risk) != 0:
                            risk = risk + ', ' + 'высокая база'
                        else:
                            risk = 'Высокая база'
                    elif Q_oil_now_for_excel[i] > 10:
                        if len(risk) != 0:
                            risk = risk + ', ' + 'по снижению базы'
                        else:
                            risk = 'По снижению базы'
                    risks.append(risk)
                    try:
                        if up_Q_oil_for_excel[i] >= 6:
                            if risk == '':
                                kandidat.append('Кандидат')
                            else:
                                kandidat.append('Кандидат с рисками')
                        else:
                            kandidat.append('')
                    except TypeError:
                        kandidat.append('')

                else:
                    Q_water_now_for_excel.append('')
                    B_now_for_excel.append('')
                    Q_oil_now_for_excel.append('')
                    Q_water_for_excel.append('')
                    B_for_excel.append('')
                    Q_oil_for_excel.append('')
                    K_pron_for_excel.append('')
                    P_zab_now_for_excel.append('')
                    P_pl_now_for_excel.append('')
                    up_Q_oil_for_excel.append('')
                    formula_types.append('')
                    if type(objects_info[keys[i]][30]) == str:
                        risks.append(objects_info[keys[i]][30])
                    kandidat.append('')
            else:
                Q_water_now_for_excel.append('')
                B_now_for_excel.append('')
                Q_oil_now_for_excel.append('')
                Q_water_for_excel.append('')
                B_for_excel.append('')
                Q_oil_for_excel.append('')
                K_pron_for_excel.append('')
                P_zab_now_for_excel.append('')
                P_pl_now_for_excel.append('')
                up_Q_oil_for_excel.append('')
                formula_types.append('')
                risks.append('Риски по эффективности проведения повторного ГРП')
                kandidat.append('')
                obrisk.append('')
        else:
            Q_water_now_for_excel.append('')
            B_now_for_excel.append('')
            Q_oil_now_for_excel.append('')
            Q_water_for_excel.append('')
            B_for_excel.append('')
            Q_oil_for_excel.append('')
            K_pron_for_excel.append('')
            P_zab_now_for_excel.append('')
            P_pl_now_for_excel.append('')
            up_Q_oil_for_excel.append('')
            formula_types.append('')
            if keys[i] in list(objects_info_problems.keys()):
                risks.append(objects_info_problems[keys[i]])
            else:
                risks.append('')
            kandidat.append('')
            obrisk.append('')

    df = pd.DataFrame({'м-е': field_for_excel, 'скв.': stock_for_excel, 'пласт': form_for_excel,
                       'тип скважины': type_stock_for_excel, 'N стадий при первичном ГРП': n_stad_first,
                       'кол-во повторных ГРП': grp_povt, 'характер работы': x_w, 'cостояние': sost,
                       'последний рабочий месяц': last_month, 'формула': formula_types,
                       'Qж.тек': Q_water_now_for_excel, '%.тек': B_now_for_excel, 'Qн.тек': Q_oil_now_for_excel,
                       'Qж.расч': Q_water_for_excel, '%.расч': B_for_excel, 'Qн.расч': Q_oil_for_excel,
                       'Кпрон.расч': K_pron_for_excel, 'Pзаб.расч': P_zab_now_for_excel,
                       'Pпл.расч': P_pl_now_for_excel, 'прирост Qн.расч': up_Q_oil_for_excel,
                       'риски/комментарии по кандидату': risks, 'Кандидат': kandidat, ' ': obrisk})

    root = tk.Tk()
    root.withdraw()
    name = filedialog.asksaveasfilename(filetypes=[('Excel file', '*.xlsx')])
    if name[len(name) - 5: len(name)] != '.xlsx':
        name = name + '.xlsx'
    if name == '.xlsx':
        name = 'Расчёт приростов.xlsx'

    root.destroy()

    writer = pd.ExcelWriter(name)
    df.to_excel(writer, sheet_name='ИТОГ', index=False, na_rep='NaN')

    # Auto-adjust columns' width
    for column in df:
        column_width = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets['ИТОГ'].set_column(col_idx, col_idx, column_width + 4)

    writer.close()
