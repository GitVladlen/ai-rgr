# -*- coding: utf-8 -*-
import datetime
import pyodbc
from tabulate import tabulate


def read_file():
    connection_string = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\КПИ_БД_ОШІ.mdb;"

    connection = pyodbc.connect(connection_string)
    cursor = connection.cursor()

    query = r"select DATEandTIME, PATNAME, STRESS, attr1, attr3, attr4, attr9, attr17, attr21, attr32, attr37 from EXP_ATTRIBUTES"
    raw_rows = cursor.execute(query).fetchall()
    cursor.close()
    connection.close()

    index, stop = 0, len(raw_rows)
    rows = []
    while True:
        row = raw_rows[index:index + 3]
        for i, r in enumerate(row):
            for j, o in enumerate(r):
                if o == None:
                    row[i][j] = 0
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


def calc_type(x1, x2, x3):
    h0 = 0.1
    hi = abs(h0 * x1)
    if (x2 - x1) > hi and (x2 - x3) > hi:
        return 1
    elif (x1 - x2) > hi and (x3 - x2) > hi:
        return 2
    elif (x2 - x1) > hi or (x3 - x2) > hi or (x3 - x1) > hi:
        return 3
    elif (x1 - x2) > hi or (x2 - x3) > hi or (x1 - x3) > hi:
        return 4
    else:
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
            res[i] = round(res[i] * 100 / sum_res, 2)
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


def get_student_names(dataset):
    names = []

    index, stop = 0, len(dataset)
    while True:
        names.append(dataset[index][1])

        if index + 3 == stop:
            break
        index = index + 3
    return names


def make_table3(names, types_percents, all_students_types, states):
    dominant_types = []
    for data in types_percents:
        maxs = data.index(max(data))
        dominant_types.append(maxs + 1)

    table_headers = ["Ім'я студента"]

    raw_headers = [
        "ЧСС",
        "Середня\nсиметрія Т",
        "СКО\nсиметрії Т",
        "SDNN",
        "Індекс\nнапруги",
        "Зсув ST,\nмв.",
        "Інтервал\nP-Q(R), мс.",
        "Площі\nT/R"
    ]

    for raw_header, dominant_type in zip(raw_headers, dominant_types):
        table_headers.append(raw_header + f"\n(Тип {dominant_type})")

    table_headers.append("Стан")

    state_names = ["Задовільний", "Умовно задовільний", "Незадовільний"]

    students_data = []
    for name, student_types, state in zip(names, all_students_types, states):
        fixed_name = '\n'.join(name.split(' '))
        student_data = [fixed_name]
        for dominant_type, student_type in zip(dominant_types, student_types):
            if dominant_type == student_type:
                student_data.append(f"{student_type} (+)")
            else:
                student_data.append(f"{student_type} (-)")
        student_data.append(state_names[state])

        students_data.append(student_data)

    return table_headers, students_data


if __name__ == '__main__':
    year_15, year_17 = read_file()
    types_15 = calc_types(year_15)
    types_17 = calc_types(year_17)
    table_15 = make_table1(types_15)
    status_15 = calc_normals(table_15, types_15)
    status_17 = calc_normals(table_15, types_17)
    table2_1 = make_table2(status_15)
    table2_2 = make_table2(status_17)

    # Table 1
    table_1_data = []

    table_1_headers = [
        "Показники",
        "Тип 1\n(Максимум)",
        "Тип 2\n(Мінімум)",
        "Тип 3\n(Зростання)",
        "Тип 4\n(Спадання)",
        "Тип 5\n(Постійний)"
    ]

    table_1_labels = [
        "ЧСС",
        "Середня симетрія Т",
        "СКО симетрії Т",
        "SDNN",
        "Індекс напруги",
        "Зсув ST, мв.",
        "Інтервал P-Q(R), мс.",
        "Площі T/R"
    ]

    for i, row in enumerate(table_15):
        tab_row = [table_1_labels[i]]
        max_el = max(row)
        for row_el in row:

            tab_row.append(f"[ {row_el} % ]" if row_el == max_el else f"{row_el} %")
        table_1_data.append(tab_row)

    table_1_string = tabulate(table_1_data, headers=table_1_headers, tablefmt='fancy_grid', stralign='center')
    print(table_1_string)

    # Table 2
    table_2_data = []

    table_2_headers = [
        "Стан",
        "Кількість студентів",
        "%"
    ]

    table_2_labels = [
        "Задовільний",
        "Умовно задовільний",
        "Незадовільний"
    ]

    for i, label in enumerate(table_2_labels):
        row = [label, table2_2[1][i], table2_2[0][i]]
        table_2_data.append(row)

    table_2_string = tabulate(table_2_data, headers=table_2_headers, tablefmt='fancy_grid', stralign='center')
    print(table_2_string)

    names = get_student_names(year_17)
    table_headers, students_data = make_table3(names, table_15, types_17, status_17)

    table_3_string = tabulate(students_data, headers=table_headers, tablefmt='fancy_grid', stralign='center')
    print(table_3_string)
