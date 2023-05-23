import pandas as pd
import math
import datetime
import numpy as np
from tkinter import *
import tkinter as tk
import tkinter.font
from tkinter import filedialog
from tkinter import ttk
from functools import partial
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure


def load_files(pos_in_array, row_to_label):
    file = filedialog.askopenfilename()
    try:
        r = frame_files.nametowidget("." + str(pos_in_array))
        r.grid_forget()
    except KeyError:
        print('')
    lbl = tk.Label(frame_files, text='✔︎ Файл загружен ︎', font=('Arial Bold', 8), justify='center', name=str(pos_in_array))
    lbl.grid(row=row_to_label, column=0, columnspan=25, pady=1, sticky='e')
    array_with_file_names[pos_in_array] = file


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
        self.tree.grid(row=1, column=column, pady=1, sticky='n')
        self.tree.bind('<Double-Button-1>', self.edit_cell)

    def edit_cell(self, event):

        col = self.tree.identify_column(event.x)

        if col == '#2':
            if self.tree.identify_region(event.x, event.y) != 'heading':
                def ok(event):
                    row = int(self.tree.selection()[0][3]) - 1
                    self.tree.set(item, col, entry.get())
                    entry.destroy()
                    try:
                        objects_info[self.objects][self.indexes[row]] = float(self.tree.set(item)['Значение'])
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
            self.stock = Listbox_maker(self.frame, 6, self.stock_for_file)
            self.formation = Listbox_maker(self.frame, 10, self.formation_for_file)

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
            self.stock.listbox.delete(0, 'end')
            self.stock.listbox_info = []
            self.formation.listbox.delete(0, 'end')
            for i in range(len(self.field_for_file)):
                if self.field_for_file[i] == self.choose_param_field:
                    if self.stock_for_file[i] not in self.stock.listbox_info:
                        self.stock.listbox.insert(END, self.stock_for_file[i])
                        self.stock.listbox_info.append(self.stock_for_file[i])

        def choose_stock(self, event):

            self.lbl.grid_forget()
            self.lbl_ok.grid_forget()
            self.rb_zab_stop.grid_forget()
            self.rb_zab_cel.grid_forget()
            self.entry.grid_forget()

            self.choose_param_stock = self.stock.listbox.get(self.stock.listbox.curselection()[0])
            self.formation.listbox.delete(0, 'end')
            self.formation.listbox_info = []
            for i in range(0, len(self.stock_for_file)):
                if self.field_for_file[i] + self.stock_for_file[i] == self.choose_param_field + self.choose_param_stock:
                    if self.formation_for_file[i] not in self.formation.listbox_info:
                        self.formation.listbox.insert(END, self.formation_for_file[i])
                        self.formation.listbox_info.append(self.formation_for_file[i])

        def choose_form(self, event):
            self.how_zab.set(0)

            self.lbl.grid(row=18, column=0, pady=1, padx=10)
            self.rb_zab_stop.grid(row=18, column=0, pady=1, sticky='e')
            self.rb_zab_cel.grid(row=19, column=0, pady=1, sticky='e')
            self.entry.grid_forget()

            tree1 = Tree_maker(0)
            tree2 = Tree_maker(1)
            tree3 = Tree_maker(2)

            self.choose_param_form = self.formation.listbox.get(self.formation.listbox.curselection()[0])
            self.choose_params = self.choose_param_field + '_' + self.choose_param_form + '_' + \
                                 str(self.choose_param_stock)

            params = ['Дебит жидкости', 'Обводнённость', 'Пластовое давление', 'Забойное давление', 'Дебит нефти']
            for i in range(len(params)):
                try:
                    to_tree1 = [params[i], round(self.objects_info[self.choose_params][i + 24], 2)]
                except TypeError:
                    to_tree1 = [params[i], self.objects_info[self.choose_params][i + 24]]
                try:
                    to_tree2 = [params[i], round(self.objects_info[self.choose_params][i + 19], 2)]
                except TypeError:
                    to_tree2 = [params[i], self.objects_info[self.choose_params][i + 19]]
                try:
                    to_tree3 = [params[i], round(self.objects_info[self.choose_params][i + 29], 2)]
                except TypeError:
                    to_tree3 = [params[i], self.objects_info[self.choose_params][i + 29]]

                tree2.insert(to_tree2)
                tree2.object_finder(self.choose_params, [19, 20, 21, 22, 23])
                tree1.insert(to_tree1)
                tree1.object_finder(self.choose_params, [24, 25, 26, 27, 28])
                tree3.insert(to_tree3)
                tree3.object_finder(self.choose_params, [28, 29, 30, 31, 32])

            self.fig.clear()

            plot1 = self.fig.add_subplot()

            if len(self.objects_dates_for_graphs[self.choose_params]) > 1:

                x_array = self.objects_dates_for_graphs[self.choose_params][1][0]
                plot1.minorticks_on()
                w = plot1.plot(x_array, self.objects_info_for_graphs[self.choose_params][1][0],  color='royalblue', label='Дебит жидкости')
                o = plot1.plot(x_array, self.objects_info_for_graphs[self.choose_params][1][1], color='sienna', label='Дебит нефти')

                p = plot1.plot(x_array, self.objects_info_for_graphs[self.choose_params][1][3], marker='o',
                                     color='red', label='Пластовое давление')
                z = plot1.plot(x_array, self.objects_info_for_graphs[self.choose_params][1][4], marker='o',
                                     color='blue', label='Забойное давление')
                plot2 = self.plot1.twinx()
                b = plot2.plot(x_array, self.objects_info_for_graphs[self.choose_params][1][2], marker='o',
                               color='cyan', label='Обводнённость')
                plot1.fill_between(x_array, self.objects_info_for_graphs[self.choose_params][1][0],
                                   np.zeros_like(self.objects_info_for_graphs[self.choose_params][1][0]), color='lavender')
                plot1.fill_between(x_array, self.objects_info_for_graphs[self.choose_params][1][1],
                                   np.zeros_like(self.objects_info_for_graphs[self.choose_params][1][1]), color='plum')

                plot1.set_title('скважина ' + str(self.choose_param_stock))
                myl = w + o + p + z + b
                labs = [l.get_label() for l in myl]
                plot1.legend(myl, labs, loc='upper left')

                plot1.set_xlabel('Дата')
                plot1.set_ylabel('Qж (по МЭР), м3/сут')
                plot2.set_ylabel('Обводнённость (по МЭР), %')
                self.fig.tight_layout()
                plot1.grid(which='both', linewidth=0.2)
                self.canvas.draw()

            print(self.choose_params)

        def Pzab_entry(self):
            self.entry.grid(row=20, column=0, pady=1, sticky='e')
            self.entry.delete(0, END)
            self.entry.insert(0, objects_info[self.choose_params][22])
            self.entry.bind('<FocusOut>', lambda e: entry.destroy())
            self.entry.bind('<Return>', self.to_object_info)

        def to_object_info(self, event):
            objects_info[self.choose_params][22] = float(self.entry.get())
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

            if objects_info[keys[i]][39] < 4:

                if type(objects_info[keys[i]][24]) == float:

                    n_mgrp_fact = objects_info[keys[i]][1]                                     # Кол-во стадий МГРП ФАКТ
                    Mpr_stock = objects_info[keys[i]][2]
                    l_fact = objects_info[keys[i]][3]                # Длина ГС (расстояние между крайними портами) ФАКТ
                    B0 = objects_info[keys[i]][4]                                                                   # Bo
                    a_Xf = objects_info[keys[i]][5]                                                              # a, Xf
                    b_Xf = objects_info[keys[i]][6]                                                              # b, Xf
                    S_nns = objects_info[keys[i]][7]                                                        # S ННС расч
                    M_plan = objects_info[keys[i]][8]                                                         # Мпр план
                    P_nas = objects_info[keys[i]][9]                                                              # Рнас
                    viscosity_oil = objects_info[keys[i]][10]                                           # Вязкость нефти
                    viscosity_water = objects_info[keys[i]][11]                                          # Вязкость воды
                    nnt = objects_info[keys[i]][17]                                                         # ННТ / Нэфф

                    if objects_info[keys[i]][0] == "ГС":
                        if n_mgrp_fact > 1:
                            formula_type = 1                     # формула Ли
                        else:
                            formula_type = 2                     # формула Джоши
                    else:
                        formula_type = 3                         # формула Дюпюи

                    if formula_type != 3:
                        rw = 0.07                                                                  # Радиус скважины, rw
                    else:
                        rw = 0.108

                    """
                    if formula_type == 3:
                        if AA10 > S_nns:
                            skin = AA10 + AA10 * 0.1            # Скин расчет
                        else:
                            skin = S_nns
                    else:
                        skin = 0
                    """
                    objects_info[keys[i]].append(formula_type)

                    P_pl_2_months = objects_info[keys[i]][26]                                    # Pпл ЗАПУСКНЫЕ (2 мес)
                    P_zab_2_months = objects_info[keys[i]][27]                                  # Pзаб ЗАПУСКНЫЕ (2 мес)
                    B_2_months = objects_info[keys[i]][25]                                         # % ЗАПУСКНЫЕ (2 мес)
                    Q_w_2_months = objects_info[keys[i]][24]                                      # Qж ЗАПУСКНЫЕ (2 мес)

                    d_p = P_pl_2_months - P_zab_2_months                                                            # dp

                    n_mgrp_plan = n_mgrp_fact                                                  # Кол-во стадий МГРП ПЛАН
                    l_plan = l_fact                                  # Длина ГС (расстояние между крайними портами) ПЛАН

                    plan_Xf = a_Xf * math.log(M_plan) + b_Xf                                                   # ПЛАН Xf

                    # Mpr_stock = 0
                    try:
                        fact_Xf = a_Xf * math.log(Mpr_stock) + b_Xf                                            # ФАКТ Xf
                    except ValueError:
                        fact_Xf = 0

                    velocity_start = viscosity_liq_finder(B_2_months / 100, viscosity_oil, viscosity_water, B0)
                                                                                          # Вязкость жидкости на ЗАПУСКЕ

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
                                nnt * plan_Xf * (1 / Lf1_plan + 1 / Lf22_plan)) + c_plan / Kf_plan / wf_plan  # R_a2 ПЛАН
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
                    if method == 1:
                        Reh = ((KP_Y_fact / 2) * Re) ** 0.5  # Эффективный радиус ГС, Reh
                    else:
                        Reh = (((fact_Xf * (2 * Re)) + (math.pi * Re ** 2)) / math.pi) ** 0.5

                    if formula_type == 2:
                        a_dgoshi = l_fact / 2 * (0.5 + math.sqrt(0.25 + (2 * Reh / l_fact) ** 4)) ** 0.5  # a_джоши

                    n_stad = n_mgrp_plan  # кол-во стадий

                    P_pl_tr = objects_info[keys[i]][21]                                                          # Pпл по ТР
                    P_zab_tr = objects_info[keys[i]][22]                                                           # Pзаб ТР
                    B_tr = objects_info[keys[i]][20]                                                                  # % ТР

                    velocity_pot = viscosity_liq_finder(B_tr / 100, viscosity_oil, viscosity_water, B0)
                                                                                  # Вязкость жидкости при расчёте потенциала

                    if formula_type == 1:

                        try:
                            K = (Q_w_2_months / (1.7054 * 0.01 * d_p / velocity_start / B0)) / (
                                        2 / (Ra2 / (1 + Ra2 * Rb2) + Rd2) + (n_mgrp_fact - 2) / (Ra / (1 + Ra * Rb) + Rd))

                            Q_water = 0.017054 * (P_pl_tr - P_zab_tr) / velocity_pot / B0 * (
                                        2 / (1 / K * (Ra2_plan / (1 + Ra2_plan * Rb2_plan) + Rd2_plan)) + (n_stad - 2) /
                                        (1 / K * (Ra_plan / (1 + Ra_plan * Rb_plan) + Rd_plan))) * K_good
                            print('                           ПОСЧИТАЛОСЬ')
                        except ZeroDivisionError:
                            K = ''
                            Q_water = ''

                        objects_info[keys[i]][37] = K
                        objects_info[keys[i]][38] = Q_water

                    elif formula_type == 2:

                        try:
                            K = (Q_w_2_months / (P_pl_2_months - P_zab_2_months) / nnt * velocity_start * B0 *
                                 (math.log((a_dgoshi + math.sqrt(a_dgoshi * a_dgoshi - (l_fact / 2) ** 2)) / (l_fact / 2))
                                  + coeff_anizotropii ** 0.5 * nnt / l_fact * math.log(
                                             coeff_anizotropii ** 0.5 * nnt / (coeff_anizotropii ** 0.5 + 1) / rw) + 0) * 18.41)

                            Q_water = (K_good * K * (P_pl_tr - P_zab_tr) * nnt) / (velocity_start * B0 * (
                                        math.log((a_dgoshi + math.sqrt(a_dgoshi * a_dgoshi - (l_fact / 2) ** 2)) / (l_fact / 2))
                                        + coeff_anizotropii ** 0.5 * nnt / l_fact * math.log(
                                    coeff_anizotropii ** 0.5 * nnt / (coeff_anizotropii ** 0.5 + 1) / rw) + 0) * 18.41)
                            print('                                            ПОСЧИТАЛОСЬ')
                        except ZeroDivisionError:
                            K = ''
                            Q_water = ''

                        objects_info[keys[i]][37] = K
                        objects_info[keys[i]][38] = Q_water

                    elif formula_type == 3:
                        try:
                            if P_zab_2_months > P_nas and P_pl_2_months > P_nas:
                                K = Q_w_2_months * 18.4 * B0 * velocity_start * (math.log(Re / rw) - 0.75 + S_nns) / (
                                        nnt * (P_pl_2_months - P_zab_2_months))
                            else:
                                if P_zab_2_months <= P_nas:
                                    K = Q_w_2_months * 18.4 * B0 * velocity_start * (math.log(Re / rw) - 0.75 + S_nns) / (
                                            nnt * (P_pl_2_months - P_nas + P_nas / 1.8 * (
                                            1 - 0.2 * P_zab_2_months / P_nas - 0.8 * (P_zab_2_months / P_nas) ** 2)))
                                else:
                                    K = Q_w_2_months * 18.4 * B0 * velocity_start * (math.log(Re / rw) - 0.75 + S_nns) / (
                                            nnt * P_pl_2_months / 1.8 * (1 - 0.2 * P_zab_2_months / P_pl_2_months - 0.8 * (
                                                P_zab_2_months / P_pl_2_months) ** 2))

                            """
                            if P_pl_2_months < P_nas:
                                K_dupui = Q_w * 18.4 * B0 * velocity_start * (math.log(Re / rw) - 0.75 + S_nns) / (
                                        nnt * P_pl_2_months / 1.8 * (1 - 0.2 * P_zab_2_months / P_pl_2_months - 0.8 * (P_zab_2_months / P_pl_2_months) ^ 2))
                            """
                            if P_zab_tr > P_nas and P_pl_tr > P_nas:
                                Q_water = K * nnt / (18.4 * velocity_pot * B0 * (math.log(Re / rw) - 0.75 + S_nns)) * (
                                            P_pl_tr - P_zab_tr)
                            else:
                                if P_zab_tr <= P_nas:
                                    Q_water = K * nnt / (18.4 * velocity_pot * B0 * (math.log(Re / rw) - 0.75 + S_nns)) * \
                                              (P_pl_tr - P_nas + P_nas / 1.8 * (
                                                          1 - 0.2 * P_zab_tr / P_nas - 0.8 * (P_zab_tr / P_nas) ** 2))
                                else:
                                    Q_water = K * nnt / (18.4 * velocity_pot * B0 * (math.log(Re / rw) - 0.75 + S_nns)) * \
                                              (P_pl_tr / 1.8 * (1 - 0.2 * P_zab_tr / P_pl_tr - 0.8 * (
                                                          P_zab_tr / P_pl_tr) ^ 2)) * K_good
                            """
                            if P_pl_tr < P_nas:
                                Q_water_dupui = K_dupui * nnt / (18.4 * velocity_pot * B0 * (math.log(Re / rw) - 0.75 + skin)) * (
                                        P_pl_tr / 1.8 * (1 - 0.2 * P_zab / P_pl_tr - 0.8 * (P_zab / P_pl_tr) ^ 2))
                            """
                        except ZeroDivisionError:
                            K = ''
                            Q_water = ''
                        objects_info[keys[i]][37] = K
                        objects_info[keys[i]][38] = Q_water
                else:
                    objects_info[keys[i]][37] = ''
                    objects_info[keys[i]][38] = ''
            else:
                objects_info[keys[i]][37] = ''
                objects_info[keys[i]][38] = ''

        return objects_info

    def viscosity_liq_finder(Wc, viscosity_oil, viscosity_water, B0):
        """
        print('')

        print(keys[i])

        print('обв.: ' + str(Wc))
        """

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
                """
                print('first: ' + str((S - S1)/S))
                """
                if ((1 - S) ** Exo) / (S ** Exw) > S0:
                    a = S
                else:
                    b = S
                S1 = S
                S = (a + b) / 2
                """
                print('a: ' + str(a))
                print('b: ' + str(b))
                print('S: ' + str(S))
                """
            Kr_o = (1 - S) ** Exo
            Kr_w = fw * S ** Exw
            viscosity_liq = (viscosity_water * viscosity_oil) / (Kr_o * viscosity_water + Kr_w * viscosity_oil)
            """
            print(viscosity_water)
            print(viscosity_oil)
            print(Kr_o)
            print(Kr_w)
            """

        elif Wc == 1:
            viscosity_liq = viscosity_water
        else:
            viscosity_liq = 0
        """
        print('вязкость: ' + str(viscosity_liq))
        """
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
        P_zab_for_excel = list()
        P_pl_for_excel = list()
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

        for i in range(len(keys)):
            string = keys[i].split('_')
            field_for_excel.append(string[0])
            form_for_excel.append(string[1])
            stock_for_excel.append(string[2])
            type_stock_for_excel.append(objects_info[keys[i]][0])

            n_stad_first.append(objects_info[keys[i]][1])

            x_w.append(objects_info[keys[i]][34])
            sost.append(objects_info[keys[i]][35])

            try:
                last_month.append(objects_info[keys[i]][36])
            except AttributeError:
                last_month.append('')

            grp_povt.append(objects_info[keys[i]][39])

            if grp_povt[i] < 4:
                if type(objects_info[keys[i]][24]) == float:

                    if objects_info[keys[i]][40] == 1:
                        formula_types.append('Ли')
                    elif objects_info[keys[i]][40] == 2:
                        formula_types.append('Джоши')
                    else:
                        formula_types.append('Дюпюи')

                    Q_water_now_for_excel.append(round(objects_info[keys[i]][19], 2))
                    B_now_for_excel.append(round(objects_info[keys[i]][20], 2))
                    Q_oil_now_for_excel.append(round(objects_info[keys[i]][23], 2))
                    try:
                        K_pron_for_excel.append(round(objects_info[keys[i]][37], 2))
                        Q_water_for_excel.append(round(objects_info[keys[i]][38], 2))
                    except TypeError:
                        K_pron_for_excel.append(objects_info[keys[i]][37])
                        Q_water_for_excel.append(objects_info[keys[i]][38])

                    summer = 0

                    if B_now_for_excel[i] >= 50:
                        summer = 10
                    elif 20 <= B_now_for_excel[i] < 50:
                        summer = 15
                    elif B_now_for_excel[i] < 20:
                        summer = 25

                    if (B_now_for_excel[i] + summer) > 100:
                        B_for_excel.append(100)
                    else:
                        B_for_excel.append(round(B_now_for_excel[i] + summer, 2))

                    ro_oil = float(objects_info[keys[i]][12])
                    try:
                        Q_oil_for_excel.append(round(Q_water_for_excel[i] * ro_oil * (100 - B_for_excel[i]) / 100, 2))
                    except TypeError:
                        Q_oil_for_excel.append(0)

                    try:
                        up = Q_oil_for_excel[i] - Q_oil_now_for_excel[i]
                        if up >= 0:
                            up_Q_oil_for_excel.append(round(up, 2))
                        else:
                            up_Q_oil_for_excel.append(round(0, 2))
                    except TypeError:
                        up_Q_oil_for_excel.append(round(0, 2))

                    # три типа рисков: по запасам, по пластовому давлению и по обводнённости
                    risk = ''
                    onnt = objects_info[keys[i]][18]
                    K_por = objects_info[keys[i]][14]
                    KIN = objects_info[keys[i]][15]
                    Knn = objects_info[keys[i]][16]
                    B0 = objects_info[keys[i]][4]
                    l_fact = objects_info[keys[i]][3]

                    if onnt > 0:
                        if stock_for_excel[i] == 'ГС':
                            oiz = (l_fact + 300) * 300 * onnt * K_por * Knn / B0 * ro_oil * KIN
                        else:
                            oiz = math.pi * 300 ** 2 * onnt * K_por * Knn * 1 / B0 * ro_oil
                    else:
                        risk = 'ОННТ !> 0'
                        oiz = 5001
                    if oiz < 5000:
                        risk = risk + 'Риски по запасам'

                    P_zab_for_excel.append(round(objects_info[keys[i]][22], 2))
                    P_pl_for_excel.append(round(objects_info[keys[i]][21], 2))
                    P_pl_2_month_for_excel = round(objects_info[keys[i]][26], 2)

                    if P_pl_for_excel[i] < P_pl_2_month_for_excel*0.6:
                        if len(risk) != 0:
                            risk = risk + ', ' + 'риск по Pпл'
                        else:
                            risk = 'Риск по Pпл'
                    risks.append(risk)
                    """
                    B_last = objects_info[keys[i]][28]
                    Q_oil_stop = objects_info[keys[i]][29]
                    """

                    if B_for_excel[i] >= 70:
                        if len(risk) != 0:
                            risk = risk + ', ' + 'высокая обводнённость'
                        else:
                            risk = 'Высокая обводнённость'
                    elif Q_oil_for_excel[i] >= 15:
                        if len(risk) != 0:
                            risk = risk + ', ' + 'высокая база'
                        else:
                            risk = 'Высокая база'
                    elif Q_oil_for_excel[i] > 10:
                        if len(risk) != 0:
                            risk = risk + ', ' + 'по снижению базы'
                        else:
                            risk = 'По снижению базы'

                    if up_Q_oil_for_excel[i] >= 6:
                        if risk == '':
                            kandidat.append('Кандидат')
                        else:
                            kandidat.append('Кандидат с рисками')
                    else:
                        kandidat.append('')
                else:
                    Q_water_now_for_excel.append('')
                    B_now_for_excel.append('')
                    Q_oil_now_for_excel.append('')
                    Q_water_for_excel.append('')
                    B_for_excel.append('')
                    Q_oil_for_excel.append('')
                    K_pron_for_excel.append('')
                    P_zab_for_excel.append('')
                    P_pl_for_excel.append('')
                    up_Q_oil_for_excel.append('')
                    formula_types.append('')
                    if type(objects_info[keys[i]][36]) != str:
                        risks.append('Скважина в бездействующем фонде')
                    else:
                        risks.append('Не найден в ТР')
                    kandidat.append('')
            else:
                Q_water_now_for_excel.append('')
                B_now_for_excel.append('')
                Q_oil_now_for_excel.append('')
                Q_water_for_excel.append('')
                B_for_excel.append('')
                Q_oil_for_excel.append('')
                K_pron_for_excel.append('')
                P_zab_for_excel.append('')
                P_pl_for_excel.append('')
                up_Q_oil_for_excel.append('')
                formula_types.append('')
                risks.append('Риски по эффективности проведения повторного ГРП')
                kandidat.append('')

        df = pd.DataFrame({'м-е': field_for_excel, 'скв.': stock_for_excel, 'пласт': form_for_excel,
                           'тип скважины': type_stock_for_excel, 'N стадий при первичном ГРП': n_stad_first,
                           'кол-во повторных ГРП': grp_povt, 'характер работы': x_w, 'cостояние': sost,
                           'последний рабочий месяц': last_month, 'формула': formula_types,
                           'Qж.тек': Q_water_now_for_excel, '%.тек': B_now_for_excel, 'Qн.тек': Q_oil_now_for_excel,
                           'Qж.расч': Q_water_for_excel, '%.расч': B_for_excel, 'Qн.расч': Q_oil_for_excel,
                           'Кпрон.расч': K_pron_for_excel, 'Pзаб.расч': P_zab_for_excel,
                           'Pпл.расч': P_pl_for_excel, 'прирост Qн.расч': up_Q_oil_for_excel,
                           'риски/комментарии по кандидату': risks, 'Кандидат': kandidat})

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
    global objects_dates_for_graphs
    objects_dates_for_graphs = {}

    # удаляем строки, где нет заполненных данных по интересующим нас столбцам
    L = L.dropna(subset=['гс/ннс']).reset_index(drop=True)

    # заносим в массивы данные для объектов
    field = L['м-е'].tolist()
    formation = L['пласт'].tolist()
    stock = L['№ скважины'].tolist()
    type_stock = L['гс/ннс'].tolist()

    for i in range(len(stock)):
        stock[i] = str(stock[i])

    # заносим значение типа скважины в словарь
    form_and_stock = list()
    field_and_form = list()
    for i in range(len(field)):
        field_and_form.append(field[i] + '_' + formation[i])
        form_and_stock.append(formation[i] + '_' + stock[i])
        objects_info[field[i] + '_' + formation[i] + '_' + stock[i]] = [type_stock[i]]                # тип скважины [0]
        objects_info_for_graphs[field[i] + '_' + formation[i] + '_' + stock[i]] = [stock[i]]
        objects_dates_for_graphs[field[i] + '_' + formation[i] + '_' + stock[i]] = [stock[i]]

    keys = list(objects_info.keys())

    TR = TR.dropna(axis=0, subset=['объекты работы'])
    form_TR = TR['объекты работы'].tolist()
    stock_TR = TR['№ скважины'].tolist()
    scepka_TR = list()
    form_TR_array = list()
    stock_TR_array = list()
    for i in range(len(form_TR)):
        if str(form_TR[i]) + '_' + str(stock_TR[i]) not in scepka_TR:
            scepka_TR.append(str(form_TR[i]) + '_' + str(stock_TR[i]))
            form_TR_array.append(form_TR[i])
            stock_TR_array.append(stock_TR[i])

    date_TR_new = {}

    for i in range(len(scepka_TR)):
        if scepka_TR[i] in form_and_stock:
            df = TR[TR['№ скважины'] == stock_TR_array[i]].reset_index(drop=True)
            df = df[df['объекты работы'] == form_TR_array[i]].reset_index(drop=True)
            df = df[df['характер работы'] == 'НЕФ'].reset_index(drop=True)
            df_first = df.iloc[0]
            date_TR_new[scepka_TR[i]] = df_first['дата']

    print('date_TR_new')
    print(date_TR_new)

    date_TR_new_keys = list(date_TR_new.keys())

    New_strat = New_strat[['скважина', 'тип', 'объект разработки до гтм', 'внр.1']]
    New_strat = New_strat[New_strat['тип'] == 'ГРП'].reset_index(drop=True)
    print(New_strat)

    stock_New_strat = New_strat['скважина'].tolist()
    form_New_strat = New_strat['объект разработки до гтм'].tolist()

    scepka_New_strat = list()
    form_New_strat_array = list()
    stock_New_strat_array = list()
    for i in range(len(form_New_strat)):
        if form_New_strat[i] + '_' + str(stock_New_strat[i]) not in scepka_New_strat:
            scepka_New_strat.append(form_New_strat[i] + '_' + str(stock_New_strat[i]))
            form_New_strat_array.append(form_New_strat[i])
            stock_New_strat_array.append(stock_New_strat[i])

    refracks = {}
    for i in range(len(scepka_New_strat)):
        if scepka_New_strat[i] in date_TR_new_keys:
            print(scepka_New_strat[i])
            refracks[scepka_New_strat[i]] = []
            print('я считаюсь')
            df = New_strat[New_strat['объект разработки до гтм'] == form_New_strat_array[i]]
            df = df[df['скважина'] == stock_New_strat_array[i]]
            print(df)
            date_New_strat = df['внр.1'].tolist()
            print('дата из df: ')
            print(date_New_strat)
            A = []
            for j in range(len(date_New_strat) - 1, -1, -1):
                a = (date_TR_new[scepka_New_strat[i]] - date_New_strat[j]).days
                a = int(a/31)
                A.append(a)
            A.reverse()
            print('A: ')
            print(A)
            counter = 0
            first_refrack = list()
            for j in range(len(A)):
                if -1 < A[j] <= 6:
                    first_refrack.append(date_New_strat[j])
                elif A[j] < -1:
                    counter += 1
            if len(first_refrack) == 0:
                counter = counter - 1
                first_refrack_not_grp = list()
                for j in range(len(A)):
                    if A[j] < -1:
                        first_refrack_not_grp.append(date_New_strat[j])
                    first_refrack.append(min(first_refrack_not_grp))
            if len(first_refrack) == 0:
                first_refrack.append('-')
            refracks[scepka_New_strat[i]].append(first_refrack[len(first_refrack) - 1])
            refracks[scepka_New_strat[i]].append(counter)
    print('refracks')
    print(refracks)
    refracks_keys = list(refracks.keys())

    Frack['номер скважины'] = Frack['номер скважины'].fillna(8220748620877487)  # i am nan
    Frack = Frack.fillna(0)
    number_Frack = Frack['номер скважины'].tolist()
    form_Frack = Frack['пласт'].tolist()
    Mpr_stock_Frack = Frack['м пр'].tolist()
    date_Frack = Frack['дата'].tolist()

    print(Frack)
    scep = list()
    new_stock_Frack = {}
    Mpr_Frack = {}
    print(Mpr_stock_Frack)
    for i in range(len(number_Frack)):
        if form_Frack[i] + '_' + str(number_Frack[i]) not in list(new_stock_Frack.keys()):
            new_stock_Frack[form_Frack[i] + '_' + str(number_Frack[i])] = []
            Mpr_Frack[form_Frack[i] + '_' + str(number_Frack[i])] = []
        # print(form_Frack[i] + '_' + str(number_Frack[i]))

        if number_Frack[i] == 8220748620877487:
            # print('я в if')
            number_Frack[i] = number_Frack[i - 1]
            # print(form_Frack[i] + '_' + str(number_Frack[i]))
            if form_Frack[i] + '_' + str(number_Frack[i]) not in list(new_stock_Frack.keys()):
                # print('я в if if')
                new_stock_Frack[form_Frack[i] + '_' + str(number_Frack[i])] = [i]
                Mpr_Frack[form_Frack[i] + '_' + str(number_Frack[i])] = []
        else:
            # print('я в else')
            new_stock_Frack[form_Frack[i] + '_' + str(number_Frack[i])].append(i)
        # print(new_stock_Frack[form_Frack[i] + '_' + str(number_Frack[i])])
        new_stock_Frack[form_Frack[i] + '_' + str(number_Frack[i])].append(date_Frack[i])
        Mpr_Frack[form_Frack[i] + '_' + str(number_Frack[i])].append(Mpr_stock_Frack[i])
        # print(new_stock_Frack[form_Frack[i] + '_' + str(number_Frack[i])])
        # print('')
        scep.append(form_Frack[i] + '_' + number_Frack[i])

    print(new_stock_Frack)
    print(Mpr_Frack)

    Frack['номер скважины'] = scep
    stock_number_Frack = Frack['номер скважины'].tolist()

    keys_new_stock_Frack = list(new_stock_Frack.keys())

    """
    for i in range(len(keys_new_stock_Frack) - 1, -1, -1):
        if len(Mpr_Frack[keys_new_stock_Frack[i]]) == 0:
            del Mpr_Frack[keys_new_stock_Frack[i]]
        if len(new_stock_Frack[keys_new_stock_Frack[i]]) == 0:
            del new_stock_Frack[keys_new_stock_Frack[i]]
            keys_new_stock_Frack.pop(i)
    """

    print(Frack)
    print('new_stock_Frack')
    print(new_stock_Frack)

    date_Frack_new = {}
    Mpr_Frack_new = {}
    for i in range(len(keys_new_stock_Frack)):
        if keys_new_stock_Frack[i] in form_and_stock:
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

    print('РЕЗУЛЬТАТ: ')
    print(date_Frack_new)
    print(Mpr_Frack_new)
    print(refracks)
    print('')
    grp_first = {}
    Mpr_last = {}
    for i in range(len(keys_new_stock_Frack)):
        if keys_new_stock_Frack[i] in refracks_keys:
            if len(refracks[keys_new_stock_Frack[i]]) != 0:
                print('я есть в рефраке: ' + keys_new_stock_Frack[i])
                print(date_Frack_new[keys_new_stock_Frack[i]])
                print(Mpr_Frack_new[keys_new_stock_Frack[i]])
                print(refracks[keys_new_stock_Frack[i]])
                print('')
                for j in range(len(date_Frack_new[keys_new_stock_Frack[i]])):
                    print('изначальные массивы')
                    print(date_Frack_new[keys_new_stock_Frack[i]][j])
                    print(Mpr_Frack_new[keys_new_stock_Frack[i]][j])
                    print('')
                    for k in range(len(date_Frack_new[keys_new_stock_Frack[i]][j]) - 1, -1, -1):
                        if (refracks[keys_new_stock_Frack[i]][0] - date_Frack_new[keys_new_stock_Frack[i]][j][k]).days / 31 > 6:
                            print('удаляю старые ГРП')
                            date_Frack_new[keys_new_stock_Frack[i]][j].pop(k)
                            Mpr_Frack_new[keys_new_stock_Frack[i]][j].pop(k)
                            print(date_Frack_new[keys_new_stock_Frack[i]][j])
                            print(Mpr_Frack_new[keys_new_stock_Frack[i]][j])
                            print('')
                    print('итог')
                    print(date_Frack_new[keys_new_stock_Frack[i]][j])
                    print(Mpr_Frack_new[keys_new_stock_Frack[i]][j])
                    print('')
                print('итог')
                print(date_Frack_new[keys_new_stock_Frack[i]])
                print(Mpr_Frack_new[keys_new_stock_Frack[i]])
                print('')
                print('словари с удалёнными данными')
                print(date_Frack_new[keys_new_stock_Frack[i]])
                print(Mpr_Frack_new[keys_new_stock_Frack[i]])
                min_array = list()
                for j in range(len(date_Frack_new[keys_new_stock_Frack[i]])):
                    print('я ищу 1 грп: ' + keys_new_stock_Frack[i])
                    print('среди: ')
                    print(date_Frack_new[keys_new_stock_Frack[i]][j])
                    print(Mpr_Frack_new[keys_new_stock_Frack[i]][j])
                    print('')
                    if len(date_Frack_new[keys_new_stock_Frack[i]][j]) != 0:
                        min_array.append(min(date_Frack_new[keys_new_stock_Frack[i]][j]))
                    else:
                        if len(date_Frack_new[keys_new_stock_Frack[i]]) < 2:
                            print('в файле ФРАК нет данных по ГРП для ' + keys_new_stock_Frack[i])
                            break

                if len(min_array) != 0:
                    val_min, idx_min = min((val_min, idx_min) for (idx_min, val_min) in enumerate(min_array))
                    val_max, idx_max = max((val_max, idx_max) for (idx_max, val_max) in enumerate(min_array))

                    print('минимум: ')
                    print(min_array)
                    print(val_min, idx_min)
                    print('максимум: ')
                    print(val_max, idx_max)
                    print('')
                    print('1 грп: ')
                    print(date_Frack_new[keys_new_stock_Frack[i]][idx_min])
                    print(Mpr_Frack_new[keys_new_stock_Frack[i]][idx_min])
                    print('')
                    print('последний грп: ')
                    print(date_Frack_new[keys_new_stock_Frack[i]][idx_max])
                    print(Mpr_Frack_new[keys_new_stock_Frack[i]][idx_max])
                    print('')

                    grp_first[keys_new_stock_Frack[i]] = len(date_Frack_new[keys_new_stock_Frack[i]][idx_min])
                    Mpr_last[keys_new_stock_Frack[i]] = sum(Mpr_Frack_new[keys_new_stock_Frack[i]][idx_max])/\
                                                        len(Mpr_Frack_new[keys_new_stock_Frack[i]][idx_max])
                    print('заполняем')
                    print(grp_first[keys_new_stock_Frack[i]])
                    print(Mpr_last[keys_new_stock_Frack[i]])
                    print('')
                else:
                    grp_first[keys_new_stock_Frack[i]] = 0
                    Mpr_last[keys_new_stock_Frack[i]] = 0
        else:

            grp_first[keys_new_stock_Frack[i]] = 0
            Mpr_last[keys_new_stock_Frack[i]] = 0

    l_L = L['l'].tolist()  # l - длина скважины

    for i in range(len(form_and_stock)):
        if form_and_stock[i] in keys_new_stock_Frack:
            objects_info[keys[i]].append(grp_first[form_and_stock[i]])                                   # число ГРП [1]
            objects_info[keys[i]].append(Mpr_last[form_and_stock[i]])                               # масса пропанта [2]
        else:
            objects_info[keys[i]].append(0)
            objects_info[keys[i]].append(0)
            # print('Нет данных по числу ГРП и массе пропанта (файл ФРАК) для ' + keys[i])
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
    up_B_PVT = PVT['рост % после грп'].tolist()
    K_por_PVT = PVT['пористость'].tolist()
    KIN_PVT = PVT['кин'].tolist()
    K_nn_PVT = PVT['кнн'].tolist()

    params_from_PVT = {}
    for i in range(len(field_PVT)):
        params_from_PVT[str(field_PVT[i]) + '_' + str(form_PVT[i])] = [B0_PVT[i], a_Xf_PVT[i], b_Xf_PVT[i], S_nns_PVT[i],
                                                                       M_plan_PVT[i], P_nas_PVT[i], viscosity_oil_PVT[i],
                                                                       viscosity_water_PVT[i], ro_PVT[i], up_B_PVT[i],
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
                                                                                     # рост обводнённости после ГРП [13]
                                                                                           # Коэффициент пористости [14]
                                                                                                              # КИН [15]
                                                                                                              # Кнн [16]

    PVT_keys = list(params_from_PVT.keys())
    for i in range(len(field_and_form)):
        if field_and_form[i] in PVT_keys:
            for j in range(len(params_from_PVT[field_and_form[i]])):
                objects_info[keys[i]].append(params_from_PVT[field_and_form[i]][j])
        else:
            # print('Нет данных по параметрам из файла PVT для ' + keys[i])
            for j in range(len(params_from_PVT[PVT_keys[0]])):
                objects_info[keys[i]].append(0)

    """
    print('')
    print('добавление данных из файла PVT')
    print('параметры в словаре: ГС/ННС, количество стадий, Мпр ФРАК, длина скважины, B0, a_Xf, b_Xf, S_ННС, Мпр PVT, Рнас')
    print(objects_info)
    """

    field_H = H['м-е'].tolist()
    form_H = H['пласт'].tolist()
    stock_H = H['скважина - забой'].tolist()
    nnt_H = H['ннт'].tolist()
    onnt_H = H['оннт'].tolist()

    params_from_H = {}
    for i in range(len(field_H)):
        params_from_H[str(field_H[i]) + '_' + str(form_H[i]) + '_' + str(stock_H[i])] = [nnt_H[i], onnt_H[i]]

    H_keys = list(params_from_H.keys())
    for i in range(len(keys)):
        if keys[i] in H_keys:
            for j in range(len(params_from_H[keys[i]])):
                objects_info[keys[i]].append(params_from_H[keys[i]][j])
                                                                                                              # ННТ [17]
                                                                                                             # ОННТ [18]
        else:
            # print('Нет данных по ННТ и ОННТ (файл H) для ' + keys[i])
            for j in range(len(params_from_H[H_keys[i]])):
                objects_info[keys[i]].append(0)

    """
    print('')
    print('добавление данных из файла H')
    print('параметры в словаре: ГС/ННС, количество стадий, длина скважины, B0, a_Xf, b_Xf, S_ННС, Мпр, Рнас, ННТ')
    print(objects_info)
    """

    date = {}
    for i in range(len(stock_New_strat_array)):

        df = New_strat[New_strat['объект разработки до гтм'] == form_New_strat_array[i]].reset_index(drop=True)
        df = df[df['скважина'] == stock_New_strat_array[i]].reset_index(drop=True)
        data = df['внр.1'].tolist()
        data = data[len(data) - 1]
        date[scepka_New_strat[i]] = datetime.date(data.year, data.month, data.day)

    date_keys = date.keys()
    info = {}
    print(date_keys)
    info_for_graphs = {}
    dates_for_graphs = {}
    for i in range(len(stock_TR_array)):
        TR = TR.fillna(0)
        df = TR[TR['№ скважины'] == stock_TR_array[i]].reset_index(drop=True)
        df = df[df['объекты работы'] == form_TR_array[i]].reset_index(drop=True)
        month = 6
        if scepka_TR[i] in date_keys:
            df = df[df['характер работы'] == 'НЕФ'].reset_index(drop=True)
            Q_zhid_for_graphs = df['дебит жидкости (тр), м3/сут'].tolist()
            Q_nef_for_graphs = df['дебит нефти (тр), т/сут'].tolist()
            Wat_for_graphs = df['обводненность (тр), % (объём)'].tolist()
            Ppl_for_graphs = df['пластовое давление (тр), атм'].tolist()
            Pzab_for_graphs = df['забойное давление (тр), атм'].tolist()
            date_for_graphs = df['дата'].tolist()
            info_for_graphs[scepka_TR[i]] = [Q_zhid_for_graphs, Q_nef_for_graphs, Wat_for_graphs, Ppl_for_graphs, Pzab_for_graphs]
            dates_for_graphs[scepka_TR[i]] = [date_for_graphs]
            try:
                xr_df = df.iloc[len(df) - 6:len(df)].reset_index(drop=True)
            except IndexError:
                xr_df = df
            xr_df_array = xr_df['состояние'].tolist()

            if (len(df) != 0) and ('РАБ.' in xr_df_array or 'НАК.' in xr_df_array or 'РАБ' in xr_df_array or 'НАК' in xr_df_array):
                # print('Я СЧИТАЮСЬ')
                try:
                    df_2_months = df.iloc[1]
                except IndexError:
                    df_2_months = df.iloc[0]
                df_last = df.iloc[len(df) - 1]
                dates = df['дата'].tolist()
                k = 0
                for j in range(len(dates)):
                    if date[scepka_TR[i]].year == dates[j].year and date[scepka_TR[i]].month == dates[j].month:
                        k = 1
                        try:
                            df = df.iloc[j:(j + month)].reset_index(drop=True)
                        except IndexError:
                            df = df.iloc[j:len(df)].reset_index(drop=True)

                        break
                    if k == 0 and j == len(dates) - 1:
                        try:
                            df = df.iloc[len(df) - month:len(df)].reset_index(drop=True)
                        except IndexError:
                            df = df

                info[scepka_TR[i]] = [df['дебит жидкости (тр), м3/сут'].sum() / len(df),
                                      df['обводненность (тр), % (объём)'].sum() / len(df),
                                      df['пластовое давление (тр), атм'].sum() / len(df),
                                      df['забойное давление (тр), атм'].sum() / len(df),
                                      df['дебит нефти (тр), т/сут'].sum() / len(df),
                                      float(df_2_months['дебит жидкости (тр), м3/сут']),
                                      float(df_2_months['обводненность (тр), % (объём)']),
                                      float(df_2_months['пластовое давление (тр), атм']),
                                      float(df_2_months['забойное давление (тр), атм']),
                                      float(df_2_months['дебит нефти (тр), т/сут']),
                                      float(df_last['дебит жидкости (тр), м3/сут']),
                                      float(df_last['обводненность (тр), % (объём)']),
                                      float(df_last['пластовое давление (тр), атм']),
                                      float(df_last['забойное давление (тр), атм']),
                                      float(df_last['дебит нефти (тр), т/сут']),
                                      df_last['характер работы'],
                                      df_last['состояние'],
                                      datetime.date(df_last['дата'].year, df_last['дата'].month, df_last['дата'].day)]
                                                                                              # Дебит жидкости (ТР) [19]
                                                                                               # Обводнённость (ТР) [20]
                                                                                          # Пластовое давление (ТР) [21]
                                                                                           # Забойное давление (ТР) [22]
                                                                                                 # Дебит нефти (ТР) [23]
                                                                                               # Дебит жидкости (2) [24]
                                                                                                # Обводнённость (2) [25]
                                                                                           # Пластовое давление (2) [26]
                                                                                            # Забойное давление (2) [27]
                                                                                                  # Дебит нефти (2) [28]
                                                                                       # Дебит жидкости (последний) [29]
                                                                                        # Обводнённость (последняя) [30]
                                                                                   # Пластовое давление (последнее) [31]
                                                                                    # Забойное давление (последнее) [32]
                                                                                          # Дебит нефти (последний) [33]
                                                                                                  # Характер работы [34]
                                                                                                        # Состояние [35]
                                                                                                             # Дата [36]

            else:
                # print('Я НЕ СЧИТАЮСЬ')
                df_last = df.iloc[len(df) - 1]
                info[scepka_TR[i]] = ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
                                      df_last['характер работы'], df_last['состояние'],
                                      datetime.date(df_last['дата'].year, df_last['дата'].month, df_last['дата'].day)]

        else:
            # print('МЕНЯ НЕТ')
            info[scepka_TR[i]] = ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']

    # print(info)
    info_keys = list(info.keys())
    print(objects_info_for_graphs)
    info_keys_for_graphs = list(info_for_graphs.keys())
    for i in range(len(form_and_stock)):
        if form_and_stock[i] in info_keys:
            for j in range(len(info[form_and_stock[i]])):
                objects_info[keys[i]].append(info[form_and_stock[i]][j])
        else:
            for j in range(18):
                objects_info[keys[i]].append('')

        if form_and_stock[i] in info_keys_for_graphs:
            print(info_for_graphs[form_and_stock[i]])
            print(form_and_stock[i])
            print(keys[i])
            objects_info_for_graphs[keys[i]].append(info_for_graphs[form_and_stock[i]])
            objects_dates_for_graphs[keys[i]].append(dates_for_graphs[form_and_stock[i]])

    print(objects_info)
    for i in range(len(keys)-1, -1, -1):
        if objects_info[keys[i]][12] == 0:
            del objects_info[keys[i]]
            del objects_info_for_graphs[keys[i]]
            del objects_dates_for_graphs[keys[i]]
            form_and_stock.pop(i)
            field_and_form.pop(i)
            keys.pop(i)

    print(objects_info)
    for i in range(len(keys)):
        for j in range(2):
            objects_info[keys[i]].append(0)                                                                 # Кпрон [37]
                                                                                                             # Qжид [38]
    for i in range(len(objects_info)):
        if form_and_stock[i] in refracks_keys:
            objects_info[keys[i]].append(refracks[form_and_stock[i]][1])
        else:
            objects_info[keys[i]].append(0)
                                                                                         # Количество повторных ГРП [39]
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
array_with_file_names = [0, 0, 0, 0, 0, 0]


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


lbl = tk.Label(frame_files, text='                    Запасы                  ', font=font, justify='center')
lbl.grid(row=20, column=0, pady=1, padx=1)
btn = tk.Button(frame_files, text='Загрузить', font=('Arial Bold', 12), justify='center', width=20)
btn.grid(row=20, column=1, pady=1)
lbl = tk.Label(frame_files, text=' ', font=('Arial Bold', 8), justify='center')
lbl.grid(row=21, column=0, columnspan=2, pady=1, sticky='e')

lbl = tk.Label(frame_files, text=' ', font=('Arial Bold', 16), justify='center')
lbl.grid(row=25, column=0, padx=10, sticky='e')

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

window.mainloop()
