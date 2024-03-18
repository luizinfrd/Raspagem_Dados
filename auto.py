import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('Pasta1.xlsx')
sheet = workbook['clientes']

for linha in sheet.iter_rows(min_row = 3):
    pyautogui.click(142, 235)
    pyautogui.write(linha[0].value)

    pyautogui.click(148, 253)
    pyautogui.write(linha[1].value)

    pyautogui.click(179, 273)
    pyautogui.write(linha[2].value)

    pyautogui.click(172, 290)
    pyautogui.write(linha[3].value)

    pyautogui.click(157, 310)
    pyautogui.write(linha[4].value)

    pyautogui.click(141, 326)
    pyautogui.write(linha[5].value)


