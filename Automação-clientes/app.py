import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

for linha in vendas_sheet.iter_rows(min_row=2):
    # nome
    pyautogui.click(621,409,duration=1.5)
    pyautogui.write(linha[0].value)
    # produto
    pyautogui.click(620,436,duration=1.5)
    pyautogui.write(linha[1].value)
    # quantidade
    pyautogui.click(639,462,duration=1.5)
    pyautogui.write(str(linha[2].value))
    # categoria
    pyautogui.click(696,487,duration=1.5)
    pyautogui.write(linha[3].value)
    pyautogui.click(579,518,duration=1.5)
    pyautogui.click(584,475,duration=1.5)
    
    