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
import threading

from text import text_create, clear_text, text_append, colored_text_append
import formulas as f
import data_from_df as dfd
import df_to_excel as dte


class Widgets:
    """Класс для создания сцепки виджеты по сцепкам Entry-Listbox"""

    def __init__(self, frame, header, widget_width, row_number, column, columnspan):
        """
        :param frame: фрейм, в котором будут отстраиваться виджеты
        :param header: заголовок для пары виджетов (указывает на данные, которая будет заноситься в Listbox)
        :param widget_width: ширина виджетов
        :param row_number: номер строки, в который виджет будет вставлен
        :param column: номер столбца, в который виджет будет вставлен
        :param columnspan: число столбцов, объединённых под виджет
        """
        self.items = None
        self.choice = None
        self.frame = frame
        self.width = widget_width
        self.row = row_number
        self.column = column
        self.columnspan = columnspan
        self.header = header

        self.create_label()
        self.listbox = self.create_listbox()
        self.entry = self.create_entry()
        self.entry.bind("<KeyRelease>", self.filter_listbox)

    def filter_listbox(self, event):
        """
        Связывание Listbox с Entry (поиск нужного объекта в Listbox)
        :param event: ввод символов в Entry
        :return: -
        """
        self.update_choice()
        self.listbox.delete(0, tk.END)
        filtered_items = [item for item in self.items if self.entry.get().lower() in item.lower()]
        for item in filtered_items:
            self.listbox.insert(tk.END, str(item))
            if item == self.choice:
                index = self.listbox.get(0, "end").index(item)
                self.listbox.selection_set(index)

    def lock_listbox(self):
        self.clear_all()
        self.listbox.configure(state='disabled', bg='lightgray')
        self.entry.configure(state='disabled', bg='gray')

    def unlock_listbox(self):
        self.listbox.configure(state='normal', bg='white')
        self.entry.configure(state='normal', bg='white')

    def update_choice(self):
        """
        Обновление выбранного элемента при изменении Entry
        :return: -
        """
        selected_index = self.listbox.curselection()
        if len(selected_index) != 0:
            self.choice = self.listbox.get(selected_index[0])

    def create_label(self):
        """
        Создание строки заголовка для пары виджетов
        :return: -
        """
        label = tk.Label(self.frame, text=self.header, font=font, justify='left')
        label.grid(row=self.row, column=self.column, columnspan=self.columnspan, sticky='we', padx=2)

    def create_entry(self):
        """
        Создание виджета Entry
        :return: виджет Entry
        """
        entry = tk.Entry(self.frame, width=self.width)
        entry.grid(row=self.row + 1, column=self.column, columnspan=self.columnspan, sticky='we', padx=2)
        return entry

    def create_listbox(self):
        """
        Создание виджета Listbox
        :return: виджет Listbox
        """
        listbox = tk.Listbox(self.frame, width=self.width, exportselection=False, height=8)
        listbox.grid(row=self.row + 2, column=self.column, columnspan=self.columnspan, sticky='we', padx=2)
        return listbox

    def clear_all(self):
        self.listbox.delete(0, tk.END)
        self.entry.delete(0, tk.END)


class Info:
    """
    Класс для создания блоков с виджетами для плоских файлов и эмпирики
    """
    def __init__(self):
        """
        :param is_base: подаётся df, соответветствующий ПФ (1 - да, 0 - нет)
        """
        self.items = None
        self.dictionary = None

        self.field = None
        self.kp = None
        self.formation = None
        self.well = None
        self.gtm = None

        self.widgets()

        self.check_type = ['field', 'formation', 'well']
        self.list_type = [self.field, self.formation, self.well]

    def widgets(self):
        """
        Создание групп виджетов (4 пары Listbox-Entry и 1 Combobox для вида ГТМ) под плоские файлы и эмпирику
        :return: -
        """
        row_number = 2
        self.field = Widgets(frame_info, 'Месторождение', 20, row_number, 0, 2)
        self.formation = Widgets(frame_info, 'Пласт', 20, row_number + 3, 0, 2)
        self.well = Widgets(frame_info, 'Скважина', 20, row_number + 6, 0, 2)

    def info_append(self, number_list, selection, select_array):
        """
        Внесение данных в Listbox, соответствующих сцепкам
        :param number_list: номер Listbox'a
        :param selection: последнее выбранное значение в Listbox'aх
        :param select_array: массив с выбранными значениями из всех виджетов (если значение не выбрано, стоит '')
        :return: -
        """
        self.items_finder(number_list, selection, select_array)
        self.list_type[number_list + 1].items = self.items
        listbox = self.list_type[number_list + 1].listbox

        listbox.delete(0, tk.END)
        for item in self.items:
            listbox.insert(tk.END, str(item))

        if number_list != len(self.check_type) - 2:
            for i in range(number_list + 2, len(self.check_type) - 1):
                self.list_type[i].listbox.delete(0, tk.END)

    def items_finder(self, number_list, selection, select_array):
        """
        Поиск данных для занесения внутрь Listbox'ов по выбранным ранее значениям, указанным в select_array
        :param number_list: номер Listbox'a
        :param selection: последнее выбранное значение в Listbox'aх
        :param select_array: массив с выбранными значениями из всех виджетов (если значение не выбрано, стоит ''
        :return: -
        """

        self.items = []
        for key in self.dictionary:
            key = self.dictionary[key]
            checker = 0
            for j in range(0, select_array.index(selection) + 1):
                method = getattr(key, self.check_type[j])
                if method == select_array[j]:
                    checker += 1

            if number_list == -1:
                check = select_array.index(selection)
            else:
                check = select_array.index(selection) + 1
            if checker == check and getattr(key, self.check_type[number_list + 1]) not in self.items:
                self.items.append(getattr(key, self.check_type[number_list + 1]))

    def clear(self):
        for listbox in range(len(self.list_type) - 1):
            self.list_type[listbox].clear_all()
        self.gtm.delete(0, tk.END)


class Update:
    """
    Класс для обновления информации в виджетах
    """
    def __init__(self, dictionary, widgets):
        """
        :param dictionary: словарь с объектами класса Field для считывания данных
        :param widgets: набор виджетов из класса Widget
        """
        self.graph_window = None

        self.dictionary = dictionary
        self.selected_values = ''
        self.old_select = ['', '', '', '', '']
        self.widgets = widgets
        self.widgets.dictionary = self.dictionary
        self.widgets.info_append(-1, self.selected_values, self.old_select)

        self.lists = [self.widgets.field.listbox, self.widgets.formation.listbox, self.widgets.well.listbox]

        for i, listbox in enumerate(self.lists):
            listbox.bind('<<ListboxSelect>>', lambda event, index=i: self.update_listbox_select(event, index))

    def update_values(self, event):
        try:
            select = []
            for listbox in self.lists:
                selected_indices = listbox.curselection()
                if selected_indices:
                    for index in selected_indices:
                        select.append(listbox.get(index))
                else:
                    select.append('')

            gtm = self.widgets.gtm.get()
            select.append(gtm)
            self.old_select = select
            self.selected_values = ', '.join(select)
            if self.selected_values in self.dictionary:
                graph.on_var_changed(self.dictionary[self.selected_values])
        except Exception as ex:
            colored_text_append(text, str(ex))
            text_append(text, 'Ошибка при обновлении графика')

    def update_listbox_select(self, event, list_number):
        """
        Добавление информации в i + 1 виджет для выбора
        :param event: выбор элемента из i виджета
        :param list_number: номер i виджета
        :return: -
        """
        try:
            # Получаем выбранные значения из всех Listbox'ов
            select = []
            for listbox in self.lists:
                selected_indices = listbox.curselection()
                if selected_indices:
                    for index in selected_indices:
                        select.append(listbox.get(index))
                else:
                    select.append('')

            difference = [x for x in select if x not in self.old_select]
            if len(difference) != 0:
                difference = difference[0]
            else:
                for i in range(list_number + 1, len(self.lists)):
                    select[i] = ''
                try:
                    difference = select[list_number]
                except IndexError:
                    difference = select[len(select) - 1]

            self.old_select = select
            self.selected_values = ', '.join(select)
            if list_number != len(self.lists) - 1:
                self.widgets.info_append(list_number, str(difference), self.old_select)

        except Exception as ex:
            colored_text_append(text, str(ex))
            text_append(text, 'Ошибка при заполнении виджетов')


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
        files[pos_in_array] = file


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
        files[pos_in_array] = file


def load_zap():
    lbl_zap.grid(row=22, column=0, pady=1, padx=1)
    btn_zap.grid(row=22, column=1, pady=1)


def not_load_zap():
    files[6] = 0
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


"""
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
    
    def clear(self):
        self.tree.delete(*self.tree.get_children())

    def insert(self, info):
        self.tree.insert("", END, values=info)

    def object_finder(self, objects, indexes):
        self.objects = objects
        self.indexes = indexes

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

"""


def runner():
    """
    Функция основной работы программы
    :return: -
    """
    btn_start.grid_forget()

    clear_text(text)
    #widgets.clear()

    objects_info = dfd.data_from_df(files, text)
    #objects_info = f.formulas(objects_info, keys, text)

    btn_excel.configure(command=lambda: dte.to_excel(objects_info, text))
    btn_excel.grid(row=31, column=0, pady=2)
    #btn_repeat.configure(command=lambda: f.formulas(objects_info, keys, text))
    #btn_repeat.grid(row=32, column=0, pady=2)

    #Info_for_listbox(frame_info, objects_info, field, formation, stock, objects_info_for_graphs, objects_dates_for_graphs)

    btn_start.grid(row=8, column=0, columnspan=2, pady=1)


def start():
    """
    Создание и запуск фонового потока для выполнения функции
    :return: -
    """
    thread = threading.Thread(target=runner)
    thread.start()


# создание окна
window = tk.Tk()
window.title('Расcчёт приростов')
fullScreenState = False
window.attributes("-fullscreen", fullScreenState)

window.geometry("%dx%d" % (1250, 700))

window.bind("<F11>", FullScreen)
window.bind("<Escape>", quitFullScreen)

for c in range(5): window.columnconfigure(index=c, weight=1)
for r in range(35): window.rowconfigure(index=r, weight=1)

frame_files = LabelFrame(window)
frame_files.grid(row=0, column=0, rowspan=20, padx=10, pady=10, ipadx=10, ipady=10, sticky='nswe')
frame_info = LabelFrame(window)
frame_info.grid(row=0, column=1, columnspan=2, rowspan=10, padx=10, pady=10, ipadx=10, ipady=10, sticky='nswe')
frame_text = LabelFrame(window)
frame_text.grid(row=0, column=3, padx=10, pady=10, ipadx=20, ipady=10, sticky='nswe')

for c in range(2): frame_files.columnconfigure(index=c, weight=1)
for r in range(35): frame_files.rowconfigure(index=r, weight=1)
for c in range(2): frame_info.columnconfigure(index=c, weight=1)
for r in range(35): frame_info.rowconfigure(index=r, weight=1)
for c in range(1): frame_text.columnconfigure(index=c, weight=1)
for r in range(35): frame_text.rowconfigure(index=r, weight=1)

global files
files = [0, 0, 0, 0, 0, 0, 0]


head = tkinter.font.Font(family="Arial Bold", size=14)
font = tkinter.font.Font(family="Arial Bold", size=12)

"""                                     З А Г Р У З К А   Ф А Й Л О В                                                """
tk.Label(frame_files, text='Загрузка файлов', font=head, justify='center').grid(row=0, column=0, columnspan=2, pady=1)

tk.Label(frame_files, text='L', font=font, justify='center').grid(row=2, column=0,  pady=1, padx=1)
tk.Button(frame_files, text='Загрузить', font=font, justify='center', width=20, command=partial(load_files, 0, 3)).grid(row=2, column=1, pady=1)
tk.Label(frame_files, text=' ', font=font, justify='center').grid(row=3, column=0, columnspan=2, pady=1, sticky='e')

tk.Label(frame_files, text='H', font=font, justify='center').grid(row=5, column=0, pady=1, padx=1)
tk.Button(frame_files, text='Загрузить', font=font, justify='center', width=20, command=partial(load_files, 1, 6)).grid(row=5, column=1, pady=1)
tk.Label(frame_files, text=' ', font=font, justify='center').grid(row=6, column=0, columnspan=2, pady=1, sticky='e')

tk.Label(frame_files, text='ФРАК', font=font, justify='center').grid(row=8, column=0, pady=1, padx=1)
tk.Button(frame_files, text='Загрузить', font=font, justify='center', width=20, command=partial(load_files, 2, 9)).grid(row=8, column=1, pady=1)
tk.Label(frame_files, text=' ', font=font, justify='center').grid(row=9, column=0, columnspan=2, pady=1, sticky='e')

tk.Label(frame_files, text='PVT', font=font, justify='center').grid(row=11, column=0, pady=1, padx=1)
tk.Button(frame_files, text='Загрузить', font=font, justify='center', width=20, command=partial(load_files, 3, 12)).grid(row=11, column=1, pady=1)
tk.Label(frame_files, text=' ', font=font, justify='center').grid(row=12, column=0, columnspan=2, pady=1, sticky='e')


tk.Label(frame_files, text='Новая стратегия', font=font, justify='center').grid(row=14, column=0, pady=1, padx=1)
tk.Button(frame_files, text='Загрузить', font=font, justify='center', width=20, command=partial(load_files, 4, 15)).grid(row=14, column=1, pady=1)
tk.Label(frame_files, text=' ', font=font, justify='center').grid(row=15, column=0, columnspan=2, pady=1, sticky='e')


tk.Label(frame_files, text='ТР', font=font, justify='center').grid(row=17, column=0, pady=1, padx=1)
tk.Button(frame_files, text='Загрузить', font=font, justify='center', width=20, command=partial(load_files, 5, 18)).grid(row=17, column=1, pady=1)
tk.Label(frame_files, text=' ', font=font, justify='center').grid(row=18, column=0, columnspan=2, pady=1, sticky='e')

how_zap = IntVar()
how_zap.set(0)

tk.Label(frame_files, text='Загружать файл с запасами:', font=font, justify='center').grid(row=20, column=0, pady=1, padx=1, sticky='nsew')

rb_zap_no = Radiobutton(frame_files, text='нет', font=font, variable=how_zap, value=0, command=not_load_zap)
rb_zap_no.grid(row=20, column=1, pady=1, padx=1, sticky='w')

rb_zap_yes = Radiobutton(frame_files, text='да', font=font, variable=how_zap, value=1, command=load_zap)
rb_zap_yes.grid(row=21, column=1, pady=1, padx=1, sticky='w')

lbl_zap = tk.Label(frame_files, text='Запасы', font=font, justify='center')
lbl_zap.grid(row=22, column=0, pady=1, padx=1)
btn_zap = tk.Button(frame_files, text='Загрузить', font=font, justify='center', width=20, command=partial(load_files_zap, 6, 23))
btn_zap.grid(row=22, column=1, pady=1)
lbl_zap.grid_forget()
btn_zap.grid_forget()

tk.Label(frame_files, text=' ', font=head, justify='center').grid(row=22, column=0, columnspan=2, pady=1)
btn_start = tk.Button(frame_files, text='Рассчитать', font=font, justify='center', width=20, command=start)
btn_start.grid(row=32, column=0, columnspan=2, pady=1)

"""                                                  П О И С К                                                       """
tk.Label(frame_info, text='Поиск по объектам', font=head, justify='center').grid(row=0, column=0, columnspan=25, pady=10)

widgets = Info()

btn_repeat = tk.Button(frame_info, text='Повторить расчёт', font=font, justify='center', width=20)
btn_excel = tk.Button(frame_info, text='Выгрузить в excel', font=font, justify='center', width=20)


"""                                                L O G - F R A M E                                                 """
main_menu = Menu()
main_menu.add_cascade(label="Мануал", command=manual)
window.config(menu=main_menu)
text = text_create(frame_text)

window.mainloop()
