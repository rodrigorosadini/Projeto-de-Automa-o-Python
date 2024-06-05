import openpyxl
import pyautogui
import os

# Caminho absoluto do arquivo
caminho_absoluto = os.path.join(os.getcwd(), 'vendaProdutos.xlsx')
print("Tentando abrir o arquivo:", caminho_absoluto)

workbook = openpyxl.load_workbook(caminho_absoluto)

vendasSheet = workbook['vendas']

for linha in vendasSheet.iter_rows(min_row=2):
    # nome
    pyautogui.click(1100,203, duration=1)
    pyautogui.write(linha[0].value)
    # produto
    pyautogui.click(1103,227, duration=1)
    pyautogui.write(linha[1].value)
    # quantidade
    pyautogui.click(1116,263, duration=1)
    pyautogui.write(str(linha[2].value))
    # categoria
    pyautogui.click(1179,281, duration=1)
    pyautogui.write(linha[3].value)
    # botão salvar
    pyautogui.click(1051,312, duration=1)
    # botão okay
    pyautogui.click(649, 427, duration=1)