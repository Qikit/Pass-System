import eel
import openpyxl

eel.init("C:/Users/zaks-/PycharmProjects/pythonProject1/program/web")



@eel.expose
def get_code(code):
    book = openpyxl.load_workbook("C:/Users/zaks-/PycharmProjects/pythonProject1/classroom.xlsx")
    sheet = book.active
    pinCodes = []
    fio = []
    no = 0

    for i in range(2, int(sheet.max_row) + 1):
        pinCodes.append(sheet['E' + str(i)].value)
        fio.append(sheet['C' + str(i)].value)

    # print("Добро пожаловать в систему пропусков!")

    for e in range(len(pinCodes)):
        if pinCodes[e] == code:
            return "\nКод верный! Пропуск разрешен.\nФИО ученика: " + str(fio[e]) + "\n"
            no += 1
            break
    if no == 0:
        return "\nКод неверный! Пропуск запрещен.\n"
    print("Работа завершена.")


eel.start("main.html", size=(700, 700))
