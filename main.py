import pymysql.cursors
import xlrd
import datetime
from sys import exit
from prettytable import PrettyTable


conn = pymysql.connect(
    host='127.0.0.1',
    port=3306,
    user='root',
    passwd='137842',
    db='db38',
    cursorclass=pymysql.cursors.DictCursor
)
cursor = conn.cursor()
fname = "Urovni2_1 (1) (1).xls"
# открыть файл
book = xlrd.open_workbook(fname)
# Открыть лист
sh = book.sheets()[0]


def mainMenu():
    commands = '\n1.Создание БД для импорта из Excel\n' \
               '2.Импортирование данных\n' \
               '3.Вывод данных по выбираемой дате или либо одной дате\n' \
               '4.Выход\n'
    print(commands)
    cmd = input('Введите команду: ')

    if cmd == '1':
        creating()
    elif cmd == '2':
        inserting()
    elif cmd == '3':
        choosing()
    elif cmd == '4':
        exit()

    else:
        print("Нет такой команды")
        mainMenu()


def creating():
    try:
        cursor.execute(
            'CREATE TABLE IF NOT EXISTS levels (object_id INT, post_code INT, parameter_code INT, data_time DATE, water_level INT)')
        print('База данных создана')
    except Exception as ex:
        print(ex)
    mainMenu()


def inserting():
    query = '''INSERT INTO levels(object_id, post_code, parameter_code, data_time, water_level)VALUES (%s, %s, %s, %s, %s)'''
    for r in range(1, sh.nrows):
        object_id = sh.cell(r, 0).value
        post_code = sh.cell(r, 1).value
        parameter_code = sh.cell(r, 2).value
        data_time = datetime.date.fromordinal(int(sh.cell(r, 3).value) + 693594).isoformat()
        water_level = sh.cell(r, 4).value

        values = (object_id, post_code, parameter_code, data_time, water_level)
        cursor.execute(query, values)

    cursor.close()
    conn.commit()
    conn.close()

    columns = str(sh.ncols)
    rows = str(sh.nrows)
    print("I just imported " + columns + " columns and " + rows + " rows")
    mainMenu()


def choosing():
    cmd_output = '\n1.Вывод по дате и коду гидрологического поста\n' \
                 '2.Вывод по диапазону дат и коду гидрологического поста\n' \
                 '3.Выход\n'

    print(cmd_output)
    cmd_input = input("Введите команду: ")

    if cmd_input == '1':
        output = input("Введите дату (в формате YYYY-MM-DD): ")
        code = int(input("Введите код гидрологического поста: "))
        try:
            cursor.execute("SELECT * FROM levels WHERE data_time = %s AND post_code = %s", [output, code])
            sort()
        except Exception as ex:
            print(ex)
        choosing()

    if cmd_input == '2':
        code = int(input("Введите код гидрологического поста: "))
        start_date = input("Введите начальную дату (в формате YYYY-MM-DD): ")
        end_date = input("Введите конечную дату (в формате YYYY-MM-DD): ")

        try:
            cursor.execute("SELECT * FROM levels WHERE data_time > %s AND data_time < %s AND post_code = %s", [start_date, end_date, code])
            sort()
        except Exception as ex:
            print(ex)
        choosing()

    if cmd_input == '3':
        mainMenu()

def sort():
    result = cursor.fetchall()
    newTable = PrettyTable(["ID", "Код_поста", "Код_параметра", "Дата", "Уровень_воды"])
    for row in result:
        newTable.add_row(row.values())
    print(newTable)

mainMenu()
