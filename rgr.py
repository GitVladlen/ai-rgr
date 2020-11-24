# -*- coding: utf-8 -*-

import tkinter as tk
import datetime
import pyodbc


def read_file():
    mdb = r".\КПИ_БД_ОШІ.mdb"
    drv = r"{Microsoft Access Driver (*.mdb, *.accdb)}"
    pwd = r"pw"

    con = pyodbc.connect(f"DRIVER={drv};DBQ={mdb};PWD={pwd}")
    cur = con.cursor()

    query = r"select DATEandTIME, PATNAME, STRESS, attr1, attr3, attr4, attr9, attr17, attr21, attr32, attr37 from EXP_ATTRIBUTES"
    raw_rows = cur.execute(query).fetchall()
    cur.close()
    con.close()

    index, stop = 0, len(raw_rows)
    rows = []
    print(len(raw_rows))
    while True:
        bad_row = False
        row = raw_rows[index:index + 3]
        for i, r in enumerate(row):
            for j, o in enumerate(r):
                if o == None:
                    row[i][j] = 0
                    # bad_row = True

        if not bad_row:
            for i in row:
                rows.append(i)
        if index + 3 == stop:
            break
        index = index + 3
    year_15, year_17 = [], []
    sep_date = datetime.datetime(2016, 12, 31)
    for row in rows:
        if row[0] > sep_date:
            year_17.append(row)
        else:
            year_15.append(row)
    return year_15, year_17


def calc_type(a0, a1, a3):
    h0 = 0.1 * abs(a0)
    if (a1 - a0) > h0 and (a1 - a3) > h0: return 1
    if (a0 - a1) > h0 and (a3 - a1) > h0: return 2
    if (a1 - a0) > h0 or a3 - a1 > h0 or a3 - a0 > h0:  return 3
    if a0 - a1 > h0 or a1 - a3 > h0 or a0 - a3 > h0:  return 4
    return 5


def calc_types(dataset):
    types = []
    index, stop = 0, len(dataset)
    while True:
        chunk = dataset[index:index + 3]
        repack = [(chunk[0][i], chunk[1][i], chunk[2][i]) for i in range(3, len(chunk[0]))]
        types.append([])
        for a0, a1, a3 in repack:
            types[-1].append(calc_type(a0, a1, a3))
        if index + 3 == stop:
            break
        index = index + 3
    return types


def make_table1(dataset):
    result = []
    for i in range(len(dataset[0])):
        lists = [data[i] for data in dataset]
        res = [0, 0, 0, 0, 0]
        for i in lists:
            res[i - 1] = res[i - 1] + 1
        result.append(res.copy())
    for res in result:
        sum_res = sum(res)
        for i in range(len(res)):
            res[i] = round(res[i] * 100 / sum_res, 3)
    return result


def calc_normals(dataset, types):
    normals = []
    for data in dataset:
        maxs = data.index(max(data))
        normals.append(maxs + 1)
    status = []
    len_normal = len(normals)
    for i, t in enumerate(types):
        similar = sum([i == j for i, j in zip(t, normals)])
        if similar == len_normal:
            status.append(0)
        elif similar == len_normal - 1:
            status.append(1)
        else:
            status.append(2)
    return status


def make_table2(dataset):
    res = [0, 0, 0]
    for data in dataset:
        res[data] = res[data] + 1
    sum_res = len(dataset)
    pr = [round(i * 100 / sum_res, 3) for i in res]
    return pr, res


if __name__ == '__main__':
    year_15, year_17 = read_file()
    types_15 = calc_types(year_15)
    types_17 = calc_types(year_17)
    table_15 = make_table1(types_15)
    status_15 = calc_normals(table_15, types_15)
    status_17 = calc_normals(table_15, types_17)
    table2_1 = make_table2(status_15)
    table2_2 = make_table2(status_17)

    # GUI
    root = tk.Tk()
    root.option_add('*Font', '12')
    root.option_add('*Font', 'Times')

    root.configure(bg='white')
    root.configure(padx=5, pady=5)

    panel_table_1 = tk.PanedWindow(root, bg='white')

    label1 = tk.Label(panel_table_1, bg='white', padx=3, pady=8, text="Груповий розподіл типів реакцій показників на навантаження")
    label1.pack()

    # Table 1
    table1 = tk.Frame(panel_table_1, bg='white')
    list_label1 = ["ЧСС", "Середня симетрія Т", "СКО симетрії Т", "SDNN",
                   "Індекс напруги", "Зсув ST, мв.", "Інтервал P-Q(R), мс.", "Площі T/R"]
    for i in range(len(table_15[0])):
        tk.Label(table1, bg='white', text=f"Тип {i + 1}", padx=5, pady=5).grid(row=0, column=i + 1)
    for i, row in enumerate(table_15):
        maxs = max(row)
        tk.Label(table1, bg='white', text=f"{list_label1[i]}", borderwidth=1, relief="solid", padx=5, pady=5).grid(row=i + 1, column=0, sticky="nsew")
        for j, leb in enumerate(row):
            # l = tk.Label(table1, bg='white', text=f"{leb}%", borderwidth=1, relief="solid", padx=5, pady=5)
            l = tk.Label(table1, bg='white', text="{:.1f} %".format(leb), borderwidth=1, relief="solid", padx=5, pady=5)
            if leb == maxs:
                l.configure(bg="lightgreen")
            l.grid(row=i + 1, column=j + 1, sticky="nsew")
    table1.pack()

    panel_table_1.pack(padx=5, pady=5, side=tk.LEFT)

    # Table 2
    panel_table_2 = tk.PanedWindow(root, bg='white')

    label2 = tk.Label(panel_table_2, padx=3, pady=8, bg='white',
                      text="\n\nРезультати порівняння індивідкальних характеристик\n" + " студентів (2017 рік) з домінантними реакціями")
    label2.pack()

    table2 = tk.Frame(panel_table_2, bg='white')
    list_label2 = ["Задовільний", "Умовно задовільний", "Незадовільний"]
    for i, text in enumerate(["Стан", "Кількість студентів", "%"]):
        tk.Label(table2, bg='white', text=text, padx=5, pady=5).grid(row=0, column=i)
    for i, row in enumerate(table2_2[1]):
        tk.Label(table2, bg='white', text=list_label2[i], borderwidth=1, relief="solid", padx=5, pady=5).grid(row=i + 1, column=0, sticky="nsew")
        l = tk.Label(table2, bg='white', text=f"{row}", borderwidth=1, relief="solid", padx=5, pady=5)
        l.grid(row=i + 1, column=1, sticky="nsew")
    for i, row in enumerate(table2_2[0]):
        tk.Label(table2, bg='white', text=list_label2[i], borderwidth=1, relief="solid", padx=5, pady=5).grid(row=i + 1, column=0, sticky="nsew")
        l = tk.Label(table2, bg='white', text=f"{row}", borderwidth=1, relief="solid", padx=5, pady=5)
        l.grid(row=i + 1, column=2, sticky="nsew")

    table2.pack()

    panel_table_2.pack(padx=5, pady=5, side=tk.RIGHT)

    root.mainloop()