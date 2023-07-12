import tkinter as tk


def text_create(frame):
    """
    Создание виджета для понимания хода работы программы
    :param frame: блок, в котором он будет находиться
    :return: виджет
    """
    text = tk.Text(frame, width=40, height=45, wrap="word")
    text.grid(row=0, column=0, sticky='nswe')
    sb = tk.Scrollbar(frame)
    text.config(yscrollcommand=sb.set)
    sb.grid(row=0, column=0, sticky='nse')
    sb.config(command=text.yview)
    return text


def text_append(text, string):
    """
    Добавление строки в виджет Text для понимания хода работы программы
    :param text: виджет для записи информации о ходе работы программы
    :return: -
    """
    text.configure(state='normal')
    text.insert('end', string + '\n')
    text.update()
    text.see('end')
    text.configure(state='disabled')


def colored_text_append(text, string):
    """
    Добавление строки красного цвета в виджет Text для выделения проблем в ПФ
    :param text: виджет для записи информации о ходе работы программы
    :return: -
    """
    text.tag_config("colored", foreground="red")
    text.configure(state='normal')
    text.insert('end', string + '\n', "colored")
    text.update()
    text.see('end')
    text.configure(state='disabled')


def clear_text(text):
    """
    Удаление информации из виджета Text
    :param text: виджет для записи информации о ходе работы программы
    :return: -
    """
    text.configure(state='normal')
    text.delete("1.0", "end")
    text.configure(state='disabled')