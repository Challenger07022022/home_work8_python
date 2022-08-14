from easygui import *
import json


def menu(choices):
    msg = 'Выберите действие:'
    return choicebox(msg, 'Информационная система', choices)


def input_data():
    msg = 'Введите данные о сотруднике:'
    title = 'Ввод данных'
    field_names = ['ФИО', 'Телефон', 'Отдел', 'Должность']
    field_values = []
    field_values = multenterbox(msg, title, field_names)
    return field_values


def show_data(db, msg):
    sp = ['id, ФИО, Телефон, Отдел, Должность', ''] + \
        [f'{i[0]}, {i[1]}, {i[2]}, {i[3]}, {i[4]}' for i in db]
    return choicebox(msg, '', sp)


def get_remove_id(db):
    remove_item = show_data(db, 'Выберите какую запись нужно удалить')
    if remove_item is not None:
        if remove_item[0].isdigit():
            if ynbox('Вы действительно хотите удалить запись о сотруднике?\n' + remove_item):
                id = int(remove_item.split(",")[0])
                return id
    return None


def msg_save():
    msgbox('База данных успешно сохранена')


def open_base2():
    msg1 = 'База не найдена!'
    msg2 = ['Указать путь к существующей базе', 'Создать новую базу']
    action = choicebox(msg1, 'Выберите действие', msg2)
    if action == 'Создать новую базу':
        base = {'staffs': {}, 'departments': {}, 'positions': {}}
        return base
    elif action == 'Указать путь к существующей базе':
        load = fileopenbox(title='Выберите базу для загрузки',
                           default='*', filetypes=["*.json"])
        with open(load, 'r', encoding='utf-8') as fh:
            base = json.load(fh)
        base['staffs'] = {int(k): v for k, v in base['staffs'].items()}
        base['departments'] = {
            int(k): v for k, v in base['departments'].items()}
        base['positions'] = {
            int(k): v for k, v in base['positions'].items()}
        return base
    else:
        exit()

def msg_export():
    msgbox('База данных успешно экспортирована в .xlsx')