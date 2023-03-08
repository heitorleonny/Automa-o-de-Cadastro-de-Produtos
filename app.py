import pyautogui
import openpyxl
import pyperclip

# 2 - abrir a planilha
workbook = openpyxl.load_workbook(
    r'C:\Users\heitor\Desktop\projeto2-cadastro_automatico\produtos.xlsx')
sheet_produtos = workbook['produtos']

# 3 - copiar o dado e colar

for linha in sheet_produtos.iter_rows(min_row=2, max_row=501):
    produto = linha[0].value
    fornecedor = linha[1].value
    categoria = linha[2].value
    quantidade = linha[3].value
    valor_unitario = linha[4].value
    notificar_venda = linha[5].value

    pyautogui.click(791, 249, duration=3)
    pyautogui.write(produto)
    pyautogui.click(814, 346, duration=1)
    pyautogui.write(fornecedor)
    pyautogui.click(815, 440, duration=1)
    pyperclip.copy(categoria)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(808, 537, duration=1)
    pyperclip.copy(valor_unitario)
    pyautogui.hotkey('ctrl', 'v')

    if notificar_venda == 'Sim':
        pyautogui.click(733, 633, duration=1)
    else:
        pyautogui.click(840, 636, duration=1)

    pyautogui.click(834, 703, duration=1)
    pyautogui.click(1198, 169, duration=3)
