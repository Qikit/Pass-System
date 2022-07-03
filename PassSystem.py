##

import openpyxl

book = openpyxl.load_workbook("classroom.xlsx")
sheet = book.active
pinCodes = []
fio = []
pin = 1
no = 0

print("Добро пожаловать в систему пропусков!")
for i in range(2, int(sheet.max_row) + 1):
    pinCodes.append(sheet['E' + str(i)].value)
    fio.append(sheet['C' + str(i)].value)

while pin != 0:
    pin = int(input("Введите код ученика: "))
    for e in range(len(pinCodes)):
        if pinCodes[e] == pin:
            print("\nКод верный! Пропуск разрешен.\nФИО ученика: " + str(fio[e]) + "\n")
            no += 1
            break
    if no == 0:
        print("\nКод неверный! Пропуск запрещен.\n")
    no=0
print("Работа завершена.")

input("\nНажмите Enter для выхода")
