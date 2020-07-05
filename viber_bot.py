import pyautogui, openpyxl, shutil, os
from openpyxl import Workbook
from datetime import datetime


wbkName = '//----------/------/------/tablica.xlsx'  #excel sheet for numbers + enforcement case number
wbk = openpyxl.load_workbook(wbkName, keep_vba=True)
dummy = 0
counter = 1

for wks in wbk.worksheets:
    for row in wks:
        number = str(row[0].value)
        case = str(row[1].value)

        pyautogui.click(x=250, y=101, interval=1)
        pyautogui.click(x=143, y=330, interval=1)
        pyautogui.click(x=279, y=155, clicks=20)
        pyautogui.click(x=126, y=159, interval=1)
        pyautogui.write(number)
        pyautogui.click(x=204, y=506, interval=1)
        if dummy == 0:
            pyautogui.hotkey('altleft','shiftleft')
            dummy = 1
        pyautogui.click(x=740, y=834, interval=1)
        pyautogui.write("Hello!") #the message
        pyautogui.keyDown('shift')
        pyautogui.press('~')
        pyautogui.keyUp('shift')
        pyautogui.write("the case number is %s." % (case))
        pyautogui.click(x=1566, y=830, interval=1)
        print("Успешно изпратено съобщение на запис " + str(counter) + " !") #success!
        counter = counter + 1

wbk.save(wbkName)
wbk.close()


#КОПИРАНЕ НА ФАЙЛА В ПАПКА АРХИВ
src = '//-------/----/-----/tablica.xlsx'
dst = '//-------/----/-----/Архив/tablica.xlsx'
shutil.copy(src, dst)


#ПРЕИМЕНУВАНЕ НА АРХИВ ФАЙЛА С ДАТА И ЧАС НА ИЗПРАЩАНЕТО
now = datetime.now()
today = now.strftime("//-------/----/-----/%d-%b-%Y %H-%M-%S.xlsx")
def main():
    os.rename("//-------/----/-----/Архив/tablica.xlsx",str(today))

main()


#ИЗЧИСТВАНЕ НА ОРИГИНАЛНИЯ ФАЙЛ, ЗА ДА Е ГОТОВ ЗА РАБОТА
def deleteData():
    wbkName = '//-------/----/-----/tablica.xlsx'
    wbk = openpyxl.load_workbook(wbkName, keep_vba=True)

    for wks in wbk.worksheets:
        for row in wks["A1:E300"]:
            for cell in row:
                cell.value = None

    wbk.save(wbkName)
    wbk.close()

deleteData()


print("Край на телефоните!")
