import pandas as pd
import math
import datetime
from datetime import timedelta
import numpy as np
from tkinter import *
import tkinter as tk
import tkinter.font
from tkinter import filedialog
from tkinter import ttk
from functools import partial
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk


def load_files(pos_in_array, row_to_label):
    file = filedialog.askopenfilename()
    if len(file) > 0:
        try:
            r = frame_files.nametowidget("." + str(pos_in_array))
            r.grid_forget()
        except KeyError:
            print('')
        lbl = tk.Label(frame_files, text='✔︎ Файл загружен ︎', font=('Arial Bold', 8), justify='center', name=str(pos_in_array))
        lbl.grid(row=row_to_label, column=0, columnspan=25, pady=1, sticky='e')
        array_with_file_names[pos_in_array] = file


def load_files_zap(pos_in_array, row_to_label):
    file = filedialog.askopenfilename()
    if len(file) > 0:
        try:
            r = frame_files.nametowidget("." + str(pos_in_array))
            r.grid_forget()
        except KeyError:
            print('')
        global label_ok
        label_ok = tk.Label(frame_files, text='✔︎ Файл загружен ︎', font=('Arial Bold', 8), justify='center', name=str(pos_in_array))
        label_ok.grid(row=row_to_label, column=0, columnspan=25, pady=1, sticky='e')
        array_with_file_names[pos_in_array] = file


def load_zap():
    lbl_zap.grid(row=22, column=0, pady=1, padx=1)
    btn_zap.grid(row=22, column=1, pady=1)


def not_load_zap():
    array_with_file_names[6] = 0
    lbl_zap.grid_forget()
    btn_zap.grid_forget()
    label_ok.grid_forget()


def FullScreen(event, fullScreenState):
    """
    открывает окно в полноэкранном режиме

    :param event: событие нажатия на F11 или кнопку, вводящую в полноэкранный режим
    :param fullScreenState: параметр, отвечающий за вид открытого окна
    """

    fullScreenState = not fullScreenState
    window.attributes("-fullscreen", fullScreenState)


def quitFullScreen(event):
    """
    выводит окно из полноэкранного режима

    :param event: событие нажатия на esc или кнопку, выводящую из полноэкранного режима
    """

    fullScreenState = False
    window.attributes("-fullscreen", fullScreenState)


def manual():
    global win
    try:
        win.destroy()
    except:
        print('')
    win = tk.Toplevel(window)
    win.title('Мануал')
    win.geometry("640x700")
    win.resizable(width=False, height=False)

    text = ["   I) Для расчета необходимо загрузить файлы с входными данными по образцу:\n\n ", "", "",
            " 1. 'L' - тип ствола, длина ГС - выгружаются координаты с NGT, по формуле считается длина ГС\n ", "",
            " 2. 'H' - ННТ и ОННТ по скважинам, выгружаются с NGT\n ", "",
            " 3. 'PVT' - занести актуальные данные ГФХ по объекту по форме\n ", "",
            " 4. 'ФРАК' - данные по фрак-листам, выгружаются с NGT без изменений\n ", "",
            " 5. 'Новая стратегия' - данные по проведенным ГТМ на скважинах, выгружаются с NGT без изменений\n ", "",
            " 6. 'ТР для загрузки' - данные по работе скважины на дату расчета, выгружаются с NGT по шаблону\n\n ", "", "",
            "   II) после загрузки выходных данных нажать кнопку 'расчёт'\n\n ", "", "",
            "   III) после расчeта в окнах справа появятся месторождения, также можно выгрузить данные в EXEL.\n\n\n ", "", "", "",
            "Комментарии\n ", "",
            "Расчет производится по формулам:\n ", "",
            "ГС + МГРП(>1стадии) - (МОДЕЛЬ ЛИ)\n ", "",
            "ГС без ГРП/ГС + ГРП (1стадия) - (ДЖОШИ)\n ", "",
            "ННС без ГРП/с ГРП - (ДЮПЮИ)\n\n ", "", "",
            "В расчёт введены условия\n ", "",
            " - расчёт производится при условии работы скважины на текущий момент или при простое не более 6мес (НЕФ + НАК/РАБ) или ПЬЕЗ/ОСТ (не более 9мес))\n ", "",
            " - ограничение по количеству повторных фраков (не более 3 фраков на один объект)\n\n", "", "",
            "РИСКИ по расчету:\n ", "",
            " - по запасам: <5тыс.т. - 'риски по ОИЗ'\n ", "",
            " - по обводненности: >70%(остановочная) - 'риски по обводненности'\n ", "",
            " - по Рпл: Рпл.текущее<0.6*Рпл.нач. - 'риски по Рпл'\n\n ", "", "",
            "Кандидат оценивается по критериям прироста по Qн и наличию рисков:\n ", "",
            "'кандидат' - Qн(расчет) >=6т/сут\n ", "",
            "'кандидат с рисками' - Qн(расчет) >=6т/сут с рисками\n\n ", "",
            "Расчетное Рзаб - по умолчанию берется остановочное забойное давление, есть возможность расчета на целевое забойное давление,\nкоторое можно задать в окне модуля нажав на кнопку 'целевое Рзаб'"]

    editor = Text(win, wrap="word", height=len(text))
    for i in range(len(text)):
        editor.insert(float(i+1), text[i])
    editor.configure(state='disabled')

    editor.grid(row=0, rowspan=10, column=0, sticky='e')


def destroyer(win):
    win.destroy()


class Tree_maker:

    def __init__(self, column):
        self.objects = None
        self.indexes = None
        self.columns = (
            ('Показатель', 160),
            ('Значение', 110)
        )
        self.tree = ttk.Treeview(frame_params, columns=[x[0] for x in self.columns], show="headings", height=5, selectmode='browse')
        self.style = ttk.Style()
        self.style.configure("Treeview.Heading", font=('Arial Bold', 11))
        for col, width in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width, anchor=tk.CENTER)
        self.tree.grid(row=1, column=column, pady=5, sticky='n')
        #self.tree.bind('<Double-Button-1>', self.edit_cell)
    """
    def edit_cell(self, event):

        col = self.tree.identify_column(event.x)

        if col == '#2':
            if self.tree.identify_region(event.x, event.y) != 'heading':
                def ok(event):
                    row = int(self.tree.selection()[0][3]) - 1
                    self.tree.set(item, col, entry.get())
                    entry.destroy()
                    try:
                        if self.indexes[0] == 18:
                            objects_info[self.objects][self.indexes[row]] = [float(self.tree.set(item)['Значение'])]
                        else:
                            objects_info[self.objects][self.indexes[row]] = float(self.tree.set(item)['Значение'])

                        if self.indexes[0] == 23:
                            objects_info_for_graphs[self.objects][1][row][len(objects_info_for_graphs[self.objects][1][row]) - 1] = float(self.tree.set(item)['Значение'])
                    except ValueError:
                        print('Ввведите в ячейку число')

                item = self.tree.identify_row(event.y)
                x, y, width, height = self.tree.bbox(item, col)

                value = self.tree.set(item, col)
                entry = Entry(self.tree)
                entry.place(x=x, y=y, width=width, height=height, anchor='nw')

                entry.insert(0, value)
                entry.bind('<FocusOut>', lambda e: entry.destroy())
                entry.bind('<Return>', ok)
    """
    def clear(self):
        self.tree.delete(*self.tree.get_children())

    def insert(self, info):
        self.tree.insert("", END, values=info)

    def object_finder(self, objects, indexes):
        self.objects = objects
        self.indexes = indexes


def main(array_with_file_names):

    class Listbox_maker:
        def __init__(self, frame, row, array):
            self.frame = frame
            self.array = array
            self.entry = tk.Entry(frame_info, width=50)
            self.entry.grid(row=row, column=0)
            self.entry.bind('<KeyRelease>', self.Scankey)

            self.listbox_info = list()
            self.listbox = Listbox(self.frame, selectmode=SINGLE, width=50, exportselection=False)
            self.listbox.grid(row=row+1, column=0)

        def Scankey(self, event):
            val = event.widget.get()
            if val == '':
                data = self.array
            else:
                data = []
                for item in self.array:
                    if val.lower() in item.lower():
                        data.append(item)
            self.Update(data)

        def Update(self, data):
            self.listbox.delete(0, 'end')
            listbox_info = []
            for item in data:
                if item not in listbox_info:
                    listbox_info.append(self.array[i])
                    self.listbox.insert('end', item)

    class Info_for_listbox:

        def __init__(self, frame, data_dict, field_for_file, formation_for_file, stock_for_file, objects_info_for_graphs, objects_dates_for_graphs):

            self.fig = plt.Figure(figsize=(15, 8))
            self.plot1 = self.fig.add_subplot()

            self.canvas = FigureCanvasTkAgg(self.fig, frame_canvas)
            self.canvas.get_tk_widget().grid(row=0, column=0)
            self.toolbar = NavigationToolbar2Tk(self.canvas, frame_canvas, pack_toolbar=False)
            self.toolbar.grid(row=1, column=0, sticky='e')

            self.objects_info_for_graphs = objects_info_for_graphs
            self.objects_dates_for_graphs = objects_dates_for_graphs

            self.how_zab = IntVar()

            self.lbl = tk.Label(frame_info, text='Считать на:                                ︎',
                                   font=('Arial Bold', 12), justify='center')

            self.lbl_ok = tk.Label(frame_info, text='✔︎', font=('Arial Bold', 12), justify='center')

            self.rb_zab_stop = Radiobutton(frame_info, text='остановочное Pзаб', font=('Arial Bold', 12),
                                           variable=self.how_zab,
                                           value=0)

            self.rb_zab_cel = Radiobutton(frame_info, text='целевое Pзаб', font=('Arial Bold', 12),
                                          variable=self.how_zab,
                                          value=1, command=self.Pzab_entry)

            self.entry = tk.Entry(frame_info, width=20, justify='left')

            self.objects_info = data_dict
            self.field_for_file = field_for_file
            self.formation_for_file = formation_for_file
            self.stock_for_file = stock_for_file

            self.frame = frame
            self.data_dict = data_dict

            self.field = Listbox_maker(self.frame, 2, self.field_for_file)
            self.stock = Listbox_maker(self.frame, 10, self.stock_for_file)
            self.formation = Listbox_maker(self.frame, 6, self.formation_for_file)

            self.field.listbox.bind("<<ListboxSelect>>", self.choose)
            self.formation.listbox.bind("<<ListboxSelect>>", self.choose_form)
            self.stock.listbox.bind("<<ListboxSelect>>", self.choose_stock)

            for i in range(0, len(self.field_for_file)):
                if self.field_for_file[i] not in self.field.listbox_info:
                    self.field.listbox.insert(END, self.field_for_file[i])
                    self.field.listbox_info.append(self.field_for_file[i])

        def choose(self, event):

            self.lbl.grid_forget()
            self.lbl_ok.grid_forget()
            self.rb_zab_stop.grid_forget()
            self.rb_zab_cel.grid_forget()
            self.entry.grid_forget()

            self.choose_param_field = self.field.listbox.get(self.field.listbox.curselection()[0])
            self.formation.listbox.delete(0, 'end')
            self.formation.listbox_info = []
            self.stock.listbox.delete(0, 'end')
            for i in range(len(self.field_for_file)):
                if self.field_for_file[i] == self.choose_param_field:
                    if self.formation_for_file[i] not in self.formation.listbox_info:
                        self.formation.listbox.insert(END, self.formation_for_file[i])
                        self.formation.listbox_info.append(self.formation_for_file[i])

        def choose_form(self, event):

            self.lbl.grid_forget()
            self.lbl_ok.grid_forget()
            self.rb_zab_stop.grid_forget()
            self.rb_zab_cel.grid_forget()
            self.entry.grid_forget()

            self.choose_param_formation = self.formation.listbox.get(self.formation.listbox.curselection()[0])
            self.stock.listbox.delete(0, 'end')
            self.stock.listbox_info = []
            for i in range(0, len(self.formation_for_file)):
                if self.field_for_file[i] + self.formation_for_file[i] == self.choose_param_field + self.choose_param_formation:
                    if self.stock_for_file[i] not in self.stock.listbox_info:
                        self.stock.listbox.insert(END, self.stock_for_file[i])
                        self.stock.listbox_info.append(self.stock_for_file[i])

        def choose_stock(self, event):
            self.how_zab.set(0)
            self.lbl_ok.grid_forget()
            self.lbl.grid(row=18, column=0, pady=1, padx=10)
            self.rb_zab_stop.grid(row=18, column=0, pady=1, sticky='e')
            self.rb_zab_cel.grid(row=19, column=0, pady=1, sticky='e')
            self.entry.grid_forget()

            tree1 = Tree_maker(0)
            tree2 = Tree_maker(1)
            tree3 = Tree_maker(2)

            self.choose_param_stock = self.stock.listbox.get(self.stock.listbox.curselection()[0])
            self.choose_params = self.choose_param_field + '_' + self.choose_param_formation + '_' + \
                                 str(self.choose_param_stock)

            params3 = ['Кпрон', 'Qжид', 'Обводнённость', 'Дебит нефти']
            for i in range(len(params3)):
                try:
                    to_tree3 = [params3[i], round(self.objects_info[self.choose_params][i + 33], 2)]
                except TypeError:
                    to_tree3 = [params3[i], self.objects_info[self.choose_params][i + 33]]

                tree3.insert(to_tree3)
                tree3.object_finder(self.choose_params, [33, 34, 35, 36])

            params = ['Дебит жидкости', 'Обводнённость', 'Пластовое давление', 'Забойное давление', 'Дебит нефти']
            for i in range(len(params)):
                try:
                    to_tree1 = [params[i], round(sum(self.objects_info[self.choose_params][i + 18])/
                                                 len(self.objects_info[self.choose_params][i + 18]), 2)]
                except (TypeError, ZeroDivisionError):
                    to_tree1 = [params[i], self.objects_info[self.choose_params][i + 18]]
                try:
                    to_tree2 = [params[i], round(self.objects_info[self.choose_params][i + 23], 2)]
                except TypeError:
                    to_tree2 = [params[i], self.objects_info[self.choose_params][i + 23]]

                tree1.insert(to_tree1)
                tree1.object_finder(self.choose_params, [18, 19, 20, 21, 22])
                tree2.insert(to_tree2)
                tree2.object_finder(self.choose_params, [23, 24, 25, 26, 27])

            self.fig.clear()

            plot1 = self.fig.add_subplot()
            if len(self.objects_info_for_graphs[self.choose_params]) > 1 and type(objects_info[self.choose_params][34]) != str:

                x_array = self.objects_dates_for_graphs[self.choose_params][1][0]

                h = timedelta(days=30)

                plot1.minorticks_on()

                w = plot1.plot(x_array, self.objects_info_for_graphs[self.choose_params][1][0],  color='royalblue', label='Дебит жидкости')
                for_4 = [x / objects_info[self.choose_params][12] for x in
                         self.objects_info_for_graphs[self.choose_params][1][4]]
                o = plot1.plot(x_array, for_4, color='sienna', label='Дебит нефти')

                try:
                    for_rasc = objects_info[self.choose_params][36] / objects_info[self.choose_params][12]
                except TypeError:
                    for_rasc = 0

                o = plot1.plot(x_array, for_4, color='sienna', label='Дебит нефти')

                x = [x_array[len(x_array) - 1], x_array[len(x_array) - 1] + h]

                y = [objects_info[self.choose_params][34], objects_info[self.choose_params][34]]

                e = plot1.plot(x, y, color='royalblue', label='  Дебит жидкости (расч.)', linewidth=1)

                plot1.fill_between(x, objects_info[self.choose_params][34], np.zeros_like(objects_info[self.choose_params][34]),
                                   color='royalblue')

                y = [for_rasc, for_rasc]

                d = plot1.plot(x, y, color='sienna', label='  Дебит нефти (расч.)', linewidth=1)

                plot1.fill_between(x, for_rasc, np.zeros_like(for_rasc), color='sienna')

                plot2 = self.plot1.twinx()
                b = plot2.plot(x_array, self.objects_info_for_graphs[self.choose_params][1][1], marker='o',
                               color='cyan', label='Обводнённость')

                plot3 = self.plot1.twinx()
                plot3.spines.right.set_position(("axes", 1.09))

                p = plot3.plot(x_array, self.objects_info_for_graphs[self.choose_params][1][2], marker='o',
                               color='red', label='Пластовое давление')
                z = plot3.plot(x_array, self.objects_info_for_graphs[self.choose_params][1][3], marker='o',
                               color='powderblue', label='Забойное давление')

                plot1.fill_between(x_array, self.objects_info_for_graphs[self.choose_params][1][0],
                                   np.zeros_like(self.objects_info_for_graphs[self.choose_params][1][0]), color='lavender')
                plot1.fill_between(x_array, for_4,
                                   np.zeros_like(for_4), color='plum')


                # [Q_zhid_for_graphs, Wat_for_graphs, Ppl_for_graphs, Pzab_for_graphs, Q_nef_for_graphs]
                plot1.set_xlim(x_array[0] - h / 2, x_array[len(x_array) - 1] + h)
                plot2.set_xlim([x_array[0] - h / 2, x_array[len(x_array) - 1] + h])
                plot2.set_ylim(0, 100)
                plot3.set_ylim(0, max(max(self.objects_info_for_graphs[self.choose_params][1][2]), max(self.objects_info_for_graphs[self.choose_params][1][3])) +
                               max(max(self.objects_info_for_graphs[self.choose_params][1][2]), max(self.objects_info_for_graphs[self.choose_params][1][3]))/100)
                plot1.set_ylim(0, max(max(for_4), max(self.objects_info_for_graphs[self.choose_params][1][0]), for_rasc, objects_info[self.choose_params][34]) +
                               max(max(for_4), max(self.objects_info_for_graphs[self.choose_params][1][0]), for_rasc, objects_info[self.choose_params][34])/100)

                plot1.set_title('скважина ' + str(self.choose_param_stock))
                myl = w + o + p + z + b + d + e
                labs = [l.get_label() for l in myl]
                legend_obj = plot1.legend(myl, labs, loc='upper left', fontsize=8, frameon=True, framealpha=1, edgecolor='none')
                legend_obj.set_draggable(True)

                plot1.set_xlabel('Дата', fontsize=10)
                plot1.set_ylabel('Q, м3/сут', fontsize=10)
                plot2.set_ylabel('Обводнённость, %', fontsize=10)
                plot3.set_ylabel('Р, атм.', fontsize=10)

                self.fig.tight_layout()
                plot1.grid(which='both', linewidth=0.2)

            self.canvas.draw()

        def Pzab_entry(self):
            self.entry.grid(row=20, column=0, pady=1, sticky='e')
            self.entry.delete(0, END)
            self.entry.insert(0, round(objects_info[self.choose_params][26], 2))
            self.entry.bind('<FocusOut>', lambda e: entry.destroy())
            self.entry.bind('<Return>', self.to_object_info)

        def to_object_info(self, event):
            objects_info[self.choose_params][26] = float(self.entry.get())
            objects_info_for_graphs[self.choose_params][1][3][len(objects_info_for_graphs[self.choose_params][1][3]) - 1] = float(self.entry.get())
            self.lbl_ok.grid(row=20, column=0, pady=1, padx=2, sticky='nsew')

    def formulas(objects_info, keys):

        # константы
        K_good = 1                                                                                                # Кусп
        Re = 250                                                                            # Радиус контура питания, Re
        Kf_fact = 500000                                                            # Kf, мД проницаемость пропанта ФАКТ
        Kf_plan = Kf_fact                                                           # Kf, мД проницаемость пропанта ПЛАН
        wf_fact = 0.005                                                                      # ширина трещины ФАКТ wf, м
        wf_plan = wf_fact                                                                    # ширина трещины ПЛАН wf, м
        Ld = 1                                                                                         # Ld, дол.ед (=1)
        coeff_anizotropii = 0.1                                                         # Коэффициент анизотропии пласта
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

                if type(objects_info[keys[i]][23]) == float and (for_djoshi != 0) and (keys[i] not in list(objects_info_problems.keys())):
                    #print('')
                    #print('расчёт для ' + keys[i])
                    #print(objects_info[keys[i]])

                    n_mgrp_fact = objects_info[keys[i]][1]  # Кол-во стадий МГРП ФАКТ
                    #print('n_mgrp_fact: ' + str(n_mgrp_fact))
                    Mpr_stock = objects_info[keys[i]][2]
                    #print('Mpr_stock: ' + str(Mpr_stock))
                    l_fact = objects_info[keys[i]][3]  # Длина ГС (расстояние между крайними портами) ФАКТ
                    #print('l_fact: ' + str(l_fact))
                    B0 = objects_info[keys[i]][4]  # Bo
                    #print('B0: ' + str(B0))
                    a_Xf = objects_info[keys[i]][5]  # a, Xf
                    #print('a_Xf: ' + str(a_Xf))
                    b_Xf = objects_info[keys[i]][6]  # b, Xf
                    #print('b_Xf: ' + str(b_Xf))
                    S_nns = objects_info[keys[i]][7]  # S ННС расч
                    #print('S_nns: ' + str(S_nns))
                    M_plan = objects_info[keys[i]][8]  # Мпр план
                    #print('M_plan: ' + str(M_plan))
                    P_nas = objects_info[keys[i]][9]  # Рнас
                    #print('P_nas: ' + str(P_nas))
                    viscosity_oil = objects_info[keys[i]][10]  # Вязкость нефти
                    #print('viscosity_oil: ' + str(viscosity_oil))
                    viscosity_water = objects_info[keys[i]][11]  # Вязкость воды
                    #print('viscosity_water: ' + str(viscosity_water))
                    nnt = objects_info[keys[i]][16]  # ННТ / Нэфф
                    #print('nnt: ' + str(nnt))


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
                    #print(formula_type)
                    objects_info[keys[i]][0] = formula_type

                    P_pl_now = objects_info[keys[i]][25]  # Pпл на последний месяц
                    P_zab_now = objects_info[keys[i]][26]  # Pзаб на последний месяц
                    B_tr_now = objects_info[keys[i]][24]  # % ТР на последний месяц

                    viscosity_pot = viscosity_liq_finder(B_tr_now / 100, viscosity_oil, viscosity_water, B0)
                    # Вязкость жидкости при расчёте потенциала
                    #print('viscosity_pot: ' + str(viscosity_pot))

                    #n_mgrp_plan = n_mgrp_fact  # Кол-во стадий МГРП ПЛАН
                    n_mgrp_plan = 3

                    l_plan = l_fact  # Длина ГС (расстояние между крайними портами) ПЛАН

                    plan_Xf = a_Xf * math.log(M_plan) + b_Xf  # ПЛАН Xf
                    plan_Xf = plan_Xf*1.15
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

                            P_pl_2_months = P_pl_for_k[p]                                        # Pпл ЗАПУСКНЫЕ (2 мес)
                            P_zab_2_months = P_zab_for_k[p]                                     # Pзаб ЗАПУСКНЫЕ (2 мес)
                            B_2_months = B_for_k[p]                                                # % ЗАПУСКНЫЕ (2 мес)
                            Q_w_2_months = Q_w_for_k[p]                                           # Qж ЗАПУСКНЫЕ (2 мес)

                            d_p = P_pl_2_months - P_zab_2_months  # dp

                            viscosity_start = viscosity_liq_finder(B_2_months / 100, viscosity_oil, viscosity_water, B0)
                            # Вязкость жидкости на ЗАПУСКЕ
                            #print('viscosity_start: ' + str(viscosity_start))

                            if formula_type == 1:
                                try:
                                    K = (Q_w_2_months / (1.7054 * 0.01 * d_p / viscosity_start / B0)) / (
                                            2 / (Ra2 / (1 + Ra2 * Rb2) + Rd2) + (n_mgrp_fact - 2) / (Ra / (1 + Ra * Rb) + Rd))
                                except ZeroDivisionError:
                                    K = ''

                            elif formula_type == 2:

                                try:
                                    K = (Q_w_2_months / (P_pl_2_months - P_zab_2_months) / nnt * viscosity_start * B0 *
                                         (math.log((a_dgoshi + math.sqrt(a_dgoshi * a_dgoshi - (l_fact / 2) ** 2)) / (l_fact / 2))
                                          + coeff_anizotropii ** 0.5 * nnt / l_fact * math.log(
                                                     coeff_anizotropii ** 0.5 * nnt / (coeff_anizotropii ** 0.5 + 1) / rw) + 0) * 18.41)
                                except ZeroDivisionError:
                                    K = ''

                            elif formula_type == 3:
                                try:
                                    if P_zab_2_months > P_nas and P_pl_2_months > P_nas:
                                        K = Q_w_2_months * 18.4 * B0 * viscosity_start * (math.log(Re / rw) - 0.75 + S_nns) / (
                                                nnt * (P_pl_2_months - P_zab_2_months))
                                    else:
                                        if P_zab_2_months <= P_nas:
                                            K = Q_w_2_months * 18.4 * B0 * viscosity_start * (math.log(Re / rw) - 0.75 + S_nns) / (
                                                    nnt * (P_pl_2_months - P_nas + P_nas / 1.8 * (
                                                    1 - 0.2 * P_zab_2_months / P_nas - 0.8 * (P_zab_2_months / P_nas) ** 2)))
                                        else:
                                            K = Q_w_2_months * 18.4 * B0 * viscosity_start * (math.log(Re / rw) - 0.75 + S_nns) / (
                                                    nnt * P_pl_2_months / 1.8 * (1 - 0.2 * P_zab_2_months / P_pl_2_months - 0.8 * (
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
                                if (K_array[p] - K_array[s])/K_array[p] > 100:
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
                                            Q_water = K * nnt / (18.4 * viscosity_pot * B0 * (math.log(Re / rw) - 0.75 + S_nns)) * \
                                                      (P_pl_now - P_nas + P_nas / 1.8 * (1 - 0.2 * P_zab_now / P_nas - 0.8 * (P_zab_now / P_nas) ** 2))
                                        else:
                                            Q_water = K * nnt / (18.4 * viscosity_pot * B0 * (math.log(Re / rw) - 0.75 + S_nns)) * \
                                                      (P_pl_now / 1.8 * (1 - 0.2 * P_zab_now / P_pl_now - 0.8 * (P_zab_now / P_pl_now) ** 2)) * K_good

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

    def to_excel(objects_info):

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

            #try:
            #last_month.append(objects_info[keys[i]][30])
            #except AttributeError:
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

                        if P_pl_now_for_excel[i] < P_pl_start_for_excel*0.6:
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
        objects_info[str(field[i]) + '_' + str(formation[i]) + '_' + str(stock[i])] = [type_stock[i]]                # тип скважины [0]
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
            if str(field_Frack[i]) + '_' + str(form_Frack[i]) + '_' + str(number_Frack[i]) not in list(new_stock_Frack.keys()):
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
                a = int(a/31)
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
                if (date_TR_new[date_TR_new_keys[i]] - val_min).days/31 <= 6:
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
                        if (refracks[keys_new_stock_Frack[i]][0] - date_Frack_new[keys_new_stock_Frack[i]][j][k]).days / 31 > 6:
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
                    Mpr_last[keys_new_stock_Frack[i]] = sum(Mpr_Frack_new[keys_new_stock_Frack[i]][idx_min])/len(Mpr_Frack_new[keys_new_stock_Frack[i]][idx_min])

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
                l_L.append(((float(x[i])-float(x_tr[i]))**2+(float(y_tr[i])-float(y[i]))**2)**0.5)
            except TypeError:
                if objects_info[keys[i]][0] == 'ГС':
                    if keys[i] not in list(objects_info_problems.keys()):
                        objects_info_problems[keys[i]] = 'Нет данных в файле L'

    grp_first_keys = list(grp_first.keys())
    for i in range(len(keys)):
        if keys[i] in grp_first_keys:
            objects_info[keys[i]].append(grp_first[keys[i]])                                            # стадии ГРП [1]
            objects_info[keys[i]].append(Mpr_last[keys[i]])                                         # масса пропанта [2]
        else:
            objects_info[keys[i]].append(0)
            objects_info[keys[i]].append(0)
            #if (keys[i] not in scepka_New_strat) and (keys[i] not in list(objects_info_problems.keys())):
                #objects_info_problems[keys[i]] = 'Нет данных по ГРП в файле "Новая стратегия"'
            if objects_info[keys[i]] == 'ГС':
                if (keys[i] not in date_Frack_new_keys) and (keys[i] in scepka_New_strat) and (keys[i] not in list(objects_info_problems.keys())):
                    objects_info_problems[keys[i]] = 'Нет данных по 1 ГРП в файле "ФРАК"'

        objects_info[keys[i]].append(l_L[i])                                                        # длина скважины [3]

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
        if (len(df) != 0) and ('РАБ.' in xr_df_array or 'РАБ' in xr_df_array or 'НАК.' in xr_df_array or 'НАК' in xr_df_array):
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

    for i in range(len(keys)-1, -1, -1):
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
            objects_info[keys[i]].append('')                                                                # Кпрон [33]
                                                                                                             # Qжид [34]

    objects_info = formulas(objects_info, keys)
    Info_for_listbox(frame_info, objects_info, field, formation, stock, objects_info_for_graphs, objects_dates_for_graphs)

    btn = tk.Button(frame_info, text='Выгрузить в excel', font=('Arial Bold', 12), justify='center', width=20,
                    command=partial(to_excel, objects_info))
    btn.grid(row=31, column=0, pady=2)

    btn = tk.Button(frame_info, text='Повторить расчёт', font=('Arial Bold', 12), justify='center', width=20,
                    command=partial(formulas, objects_info, keys))
    btn.grid(row=32, column=0, pady=2)


# создание окна
window = tk.Tk()
window.title('Расчёт приростов')
fullScreenState = False
window.attributes("-fullscreen", fullScreenState)

w, h = window.winfo_screenwidth(), window.winfo_screenheight()
window.geometry("%dx%d" % (w - 100, h - 200))

window.bind("<F11>", FullScreen)
window.bind("<Escape>", quitFullScreen)

for c in range(35): window.columnconfigure(index=c, weight=1)
for r in range(35): window.rowconfigure(index=r, weight=1)

frame_files = LabelFrame(window)
frame_files.grid(row=0, column=0, rowspan=35, padx=10, pady=10, ipadx=20, ipady=10, sticky='nswe')
frame_info = LabelFrame(window)
frame_info.grid(row=0, column=1, rowspan=35, padx=10, pady=10, ipadx=20, ipady=10, sticky='nswe')
frame_canvas = LabelFrame(window)
frame_canvas.grid(row=0, column=2, rowspan=30, columnspan=33, padx=10, pady=10, ipadx=20, ipady=10, sticky='nswe')
frame_params = LabelFrame(window)
frame_params.grid(row=31, column=2, rowspan=5, columnspan=33, padx=10, pady=10, ipadx=20, ipady=10, sticky='nswe')

for c in range(2): frame_files.columnconfigure(index=c, weight=1)
for r in range(35): frame_files.rowconfigure(index=r, weight=1)
for c in range(1): frame_info.columnconfigure(index=c, weight=1)
for r in range(35): frame_info.rowconfigure(index=r, weight=1)
for c in range(3): frame_params.columnconfigure(index=c, weight=1)
for r in range(2): frame_params.rowconfigure(index=r, weight=1)
for c in range(1): frame_canvas.columnconfigure(index=c, weight=1)
for r in range(1): frame_canvas.rowconfigure(index=r, weight=1)

global array_with_file_names
array_with_file_names = [0, 0, 0, 0, 0, 0, 0]

main_menu = Menu()
main_menu.add_cascade(label="Мануал", command=manual)

lbl = tk.Label(frame_files, text='Загрузка файлов', font=('Arial Bold', 14), justify='center')
lbl.grid(row=0, column=0, columnspan=2, pady=1)

font = tkinter.font.Font(family="Arial Bold", size=12, underline=True)

lbl = tk.Label(frame_files, text='                     L                     ', font=font, justify='center')
lbl.grid(row=2, column=0,  pady=1, padx=1)
btn = tk.Button(frame_files, text='Загрузить', font=('Arial Bold', 12), justify='center', width=20,
                command=partial(load_files, 0, 3))
btn.grid(row=2, column=1, pady=1)
lbl = tk.Label(frame_files, text=' ', font=('Arial Bold', 8), justify='center')
lbl.grid(row=3, column=0, columnspan=2, pady=1, sticky='e')


lbl = tk.Label(frame_files, text='                     H                     ', font=font, justify='center')
lbl.grid(row=5, column=0, pady=1, padx=1)
btn = tk.Button(frame_files, text='Загрузить', font=('Arial Bold', 12), justify='center', width=20,
                command=partial(load_files, 1, 6))
btn.grid(row=5, column=1, pady=1)
lbl = tk.Label(frame_files, text=' ', font=('Arial Bold', 8), justify='center')
lbl.grid(row=6, column=0, columnspan=2, pady=1, sticky='e')


lbl = tk.Label(frame_files, text='                  ФРАК                 ', font=font, justify='center')
lbl.grid(row=8, column=0, pady=1, padx=1)
btn = tk.Button(frame_files, text='Загрузить', font=('Arial Bold', 12), justify='center', width=20,
                command=partial(load_files, 2, 9))
btn.grid(row=8, column=1, pady=1)
lbl = tk.Label(frame_files, text=' ', font=('Arial Bold', 8), justify='center')
lbl.grid(row=9, column=0, columnspan=2, pady=1, sticky='e')


lbl = tk.Label(frame_files, text='                   PVT                   ', font=font, justify='center')
lbl.grid(row=11, column=0, pady=1, padx=1)
btn = tk.Button(frame_files, text='Загрузить', font=('Arial Bold', 12), justify='center', width=20,
                command=partial(load_files, 3, 12))
btn.grid(row=11, column=1, pady=1)
lbl = tk.Label(frame_files, text=' ', font=('Arial Bold', 8), justify='center')
lbl.grid(row=12, column=0, columnspan=2, pady=1, sticky='e')


lbl = tk.Label(frame_files, text='          Новая стратегия       ', font=font, justify='center')
lbl.grid(row=14, column=0, pady=1, padx=1)
btn = tk.Button(frame_files, text='Загрузить', font=('Arial Bold', 12), justify='center', width=20,
                command=partial(load_files, 4, 15))
btn.grid(row=14, column=1, pady=1)
lbl = tk.Label(frame_files, text=' ', font=('Arial Bold', 8), justify='center')
lbl.grid(row=15, column=0, columnspan=2, pady=1, sticky='e')


lbl = tk.Label(frame_files, text='                      ТР                     ', font=font, justify='center')
lbl.grid(row=17, column=0, pady=1, padx=1)
btn = tk.Button(frame_files, text='Загрузить', font=('Arial Bold', 12), justify='center', width=20,
                command=partial(load_files, 5, 18))
btn.grid(row=17, column=1, pady=1)
lbl = tk.Label(frame_files, text=' ', font=('Arial Bold', 8), justify='center')
lbl.grid(row=18, column=0, columnspan=2, pady=1, sticky='e')

how_zap = IntVar()
how_zap.set(0)

lbl = tk.Label(frame_files, text='Загружать файл с запасами:',
               font=('Arial Bold', 12), justify='center')

lbl.grid(row=20, column=0, pady=1, padx=1, sticky='nsew')

rb_zap_no = Radiobutton(frame_files, text='нет', font=('Arial Bold', 12),
                               variable=how_zap,
                               value=0, command=not_load_zap)
rb_zap_no.grid(row=20, column=1, pady=1, padx=1, sticky='w')

rb_zap_yes = Radiobutton(frame_files, text='да', font=('Arial Bold', 12),
                              variable=how_zap,
                              value=1, command=load_zap)

rb_zap_yes.grid(row=21, column=1, pady=1, padx=1, sticky='w')

lbl_zap = tk.Label(frame_files, text='                   Запасы                 ', font=font, justify='center')
lbl_zap.grid(row=22, column=0, pady=1, padx=1)
btn_zap = tk.Button(frame_files, text='Загрузить', font=('Arial Bold', 12), justify='center', width=20,
                    command=partial(load_files_zap, 6, 23))
btn_zap.grid(row=22, column=1, pady=1)
lbl_zap.grid_forget()
btn_zap.grid_forget()

lbl = tk.Label(frame_files, text=' ', font=('Arial Bold', 20), justify='center')
lbl.grid(row=22, column=0, columnspan=2, pady=1)

lbl = tk.Label(frame_files, text=' ', font=('Arial Bold', 16), justify='center')
lbl.grid(row=23, column=0, padx=10, sticky='e')

btn = tk.Button(frame_files, text='Расчитать', font=('Arial Bold', 12), justify='center', width=20,
                command=partial(main, array_with_file_names))
btn.grid(row=32, column=0, columnspan=2, pady=1)

width = 50

lbl = tk.Label(frame_info, text='Поиск по объектам', font=('Arial Bold', 14), justify='center')
lbl.grid(row=0, column=0, columnspan=25, pady=10)

entry = tk.Entry(frame_info, width=width)
entry.grid(row=2, column=0, padx=10)
listbox = Listbox(frame_info, selectmode=SINGLE, width=width, exportselection=False)
listbox.grid(row=3, column=0,  padx=10)

entry = tk.Entry(frame_info, width=width)
entry.grid(row=6, column=0, padx=10)
listbox = Listbox(frame_info, selectmode=SINGLE, width=width, exportselection=False)
listbox.grid(row=7, column=0, padx=10)

entry = tk.Entry(frame_info, width=width)
entry.grid(row=10, column=0, padx=10)
listbox = Listbox(frame_info, selectmode=SINGLE, width=width, exportselection=False)
listbox.grid(row=11, column=0,  padx=10)

lbl = tk.Label(frame_info, text=' ', font=('Arial Bold', 20), justify='center')
lbl.grid(row=18, column=0, pady=1, padx=10)
lbl = tk.Label(frame_info, text=' ', font=('Arial Bold', 20), justify='center')
lbl.grid(row=18, column=0, pady=1, sticky='e')
lbl = tk.Label(frame_info, text=' ', font=('Arial Bold', 20), justify='center')
lbl.grid(row=19, column=0, pady=1, sticky='e')
lbl = tk.Label(frame_info, text=' ', font=('Arial Bold', 20), justify='center')
lbl.grid(row=20, column=0, pady=1, sticky='e')


lbl = tk.Label(frame_info, text=' ', font=('Arial Bold', 20), justify='center')
lbl.grid(row=31, column=0, pady=2)

lbl = tk.Label(frame_info, text=' ', font=('Arial Bold', 20), justify='center')
lbl.grid(row=32, column=0, pady=2)

lbl = tk.Label(frame_params, text='Запускные', font=('Arial Bold', 12), justify='center')
lbl.grid(row=0, column=0, pady=1)

lbl = tk.Label(frame_params, text='Текущие', font=('Arial Bold', 12), justify='center')
lbl.grid(row=0, column=1, pady=1)

lbl = tk.Label(frame_params, text='Расчёт', font=('Arial Bold', 12), justify='center')
lbl.grid(row=0, column=2, pady=1)

tree1 = Tree_maker(0)
tree2 = Tree_maker(1)
tree3 = Tree_maker(2)

fig = plt.Figure(figsize=(15, 8))
plot1 = fig.add_subplot()

canvas = FigureCanvasTkAgg(fig, frame_canvas)
canvas.get_tk_widget().grid(row=0, column=0)
toolbar = NavigationToolbar2Tk(canvas, frame_canvas, pack_toolbar=False)
toolbar.grid(row=1, column=0, sticky='e')
window.config(menu=main_menu)
window.mainloop()
