import json
import openpyxl


def open_base1():
    with open(r'C:\Users\Админ\Desktop\IT\Python\MyHomeWork_Python\HomeWork8\base.json', 'r', encoding='utf-8') as fh:
        db = json.load(fh)
    db['staffs'] = {int(k): v for k, v in db['staffs'].items()}
    db['departments'] = {
        int(k): v for k, v in db['departments'].items()}
    db['positions'] = {
        int(k): v for k, v in db['positions'].items()}
    return db


def save_base(db):
    with open(r"C:\Users\Админ\Desktop\IT\Python\MyHomeWork_Python\HomeWork8\base.json", 'w', encoding='utf-8') as file:
        file.write(json.dumps(db, ensure_ascii=False))


def get_department_id(db, department):
    if department in [i[0] for i in db['departments'].values()]:
        for k, v in db['departments'].items():
            if department == v[0]:
                return k
    else:
        return None


def add_department(db, department):
    department_id = get_department_id(db, department)
    if department_id is None:
        if len(db['departments']) == 0:
            department_id = 1
        else:
            department_id = max(db['departments'].keys())+1
        db['departments'][department_id] = [department]
    return department_id


def get_positionst_id(db, position):
    if position in [n[0] for n in db['positions'].values()]:
        for k, v in db['positions'].items():
            if position == v[0]:
                return k
    else:
        return None


def add_position(db, position):
    position_id = get_positionst_id(db, position)
    if position_id is None:
        if len(db['positions']) == 0:
            position_id = 1
        else:
            position_id = max(db['positions'].keys()) + 1
        db['positions'][position_id] = [position]
    return position_id


def add_staff(db, fio, tel, department, position):
    department_id = add_department(db, department)
    position_id = add_position(db, position)
    if len(db['staffs']) == 0:
        staff_id = 1
    else:
        staff_id = max(db['staffs'].keys()) + 1
    db['staffs'][staff_id] = [fio, tel, department_id, position_id]


def get_staff(db, id):
    if id in db['staffs']:
        return [
            id,
            db['staffs'][id][0],
            db['staffs'][id][1],
            db['departments'][db['staffs'][id][2]][0],
            db["positions"][db['staffs'][id][3]][0]
        ]
    else:
        return None


def remove_staff(db, id):
    db['staffs'].pop(id, None)


def all_staffs(db):
    return [get_staff(db, key) for key in db['staffs'].keys()]


def export_to_xlsx(db):
    a = [get_staff(db, key) for key in db['staffs'].keys()]
    book = openpyxl.Workbook()
    sheet = book.active
    sheet['A1'] = 'ID'
    sheet['B1'] = 'ФИО'
    sheet['C1'] = 'ТЕЛЕФОН'
    sheet['D1'] = 'ОТДЕЛ'
    sheet['E1'] = 'ДОЛЖНОСТЬ'
    columns1 = 'ABCDE'
    columns2 = 2
    for i in a:
        a = 0
        for j in i:
            sheet[f'{columns1[a]}{str(columns2)}'] = j
            a += 1
        columns2 += 1
    book.save('my_book.xlsx')
    book.close()


# db = open_base1()
# print(export_to_xlsx(db))
