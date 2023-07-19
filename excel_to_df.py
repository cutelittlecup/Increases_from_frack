import pandas as pd
import queue
import threading


def df_reader(files):
    """
    Считывание информации из файлов и преобразование её в df
    :param files: массив с названиями файлов для считывания
    :return: массив с тремя датафреймами, соответствующими считываемым файлам
    """

    def dataframe_reader(ind):
        """
        Считывание эксель файла и преобразование информации в df
        :param ind: индекс файла в массиве с названиями для отделения эмпирики от ПР
        :return: датафрейм
        """
        print(len(files), ind, files[ind])
        file = files[ind]
        if len(file) > 0:
            if ind in [1, 3, 4]:
                df = pd.read_excel(file, header=1)
            elif ind == 6:
                df = pd.read_excel(file, header = 11)
            else:
                df = pd.read_excel(file)
            df.columns = map(str.lower, df.columns)
        else:
            df = ''
        return df

    que1 = queue.Queue()
    que2 = queue.Queue()
    que3 = queue.Queue()
    que4 = queue.Queue()
    que5 = queue.Queue()
    que6 = queue.Queue()
    que7 = queue.Queue()

    t1 = threading.Thread(target=lambda q, arg: q.put(dataframe_reader(arg)), args=(que1, 0))
    t2 = threading.Thread(target=lambda q, arg: q.put(dataframe_reader(arg)), args=(que2, 1))

    t3 = threading.Thread(target=lambda q, arg: q.put(dataframe_reader(arg)), args=(que3, 2))
    t4 = threading.Thread(target=lambda q, arg: q.put(dataframe_reader(arg)), args=(que4, 3))

    t5 = threading.Thread(target=lambda q, arg: q.put(dataframe_reader(arg)), args=(que5, 4))
    t6 = threading.Thread(target=lambda q, arg: q.put(dataframe_reader(arg)), args=(que6, 5))
    t7 = threading.Thread(target=lambda q, arg: q.put(dataframe_reader(arg)), args=(que6, 6))

    t1.start()
    t2.start()
    t3.start()
    t4.start()
    t5.start()
    t6.start()
    t7.start()

    t1.join()
    t2.join()
    t3.join()
    t4.join()
    t5.join()
    t6.join()
    t7.join()

    while not que1.empty():
        L = que1.get()

    while not que2.empty():
        Frack = que2.get()

    while not que3.empty():
        PVT = que3.get()

    while not que4.empty():
        H = que4.get()

    while not que5.empty():
        New_strat = que5.get()

    while not que6.empty():
        TR = que6.get()

    while not que7.empty():
        reserves = que6.get()

    return L, Frack, PVT, H, New_strat, TR, reserves
