import user_interface as ui
import main as m

def start():
    sp = [
        'Показать весь справочник',
        'Добавить сотрудника',
        'Удалить запись о сотруднике',
        'Экспорт базы данных',
        'Сохранить',
        'Выход']

    try:
        db = m.open_base1()
    except:
        db = ui.open_base2()

    while True:
        a = ui.menu(sp)
        match a:
            case 'Показать весь справочник':
                ui.show_data(m.all_staffs(db), 'Справочник:')
            case 'Добавить сотрудника':
                new_data = ui.input_data()
                if new_data is not None:
                    m.add_staff(db, new_data[0], new_data[1],
                                new_data[2], new_data[3])
            case 'Удалить запись о сотруднике':
                id = ui.get_remove_id(m.all_staffs(db))
                m.remove_staff(db, id)
            case 'Экспорт базы данных':
                m.export_to_xlsx(db)
                ui.msg_export()
            case 'Сохранить':
                m.save_base(db)
                ui.msg_save()
            case 'Выход':
                exit()


start()
