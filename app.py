import openpyxl
import pyautogui
import time

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

pyautogui.click(722, 739, duration=0.3)
time.sleep(2)
pyautogui.click(316, 64, duration=0.4)
pyautogui.write('https://form-teste-pyautogui.netlify.app/')
pyautogui.press('enter')
time.sleep(2)

# Definir as coordenadas dos campos do formulário
customer_coord = (110, 178)
product_coord = (338, 176)
amount_coord = (598, 177)
category_coord = (809, 178)
button_coord = (992, 176)


for linha in vendas_sheet.iter_rows(min_row=2):
    name = linha[0].value
    product = linha[1].value
    amount = linha[2].value
    category = linha[3].value

    # Preencher campo name
    pyautogui.click(customer_coord, duration=0.2)
    pyautogui.write(str(name), interval=0.1)

    # Preencher campo product
    pyautogui.click(product_coord, duration=0.2)
    pyautogui.write(str(product), interval=0.1)

    # Preencher campo amount
    pyautogui.click(amount_coord, duration=0.2)
    pyautogui.write(str(amount), interval=0.1)

    # Preencher campo de category
    pyautogui.click(category_coord, duration=0.2)
    pyautogui.write(str(category), interval=0.1)

    # Clicar no botão enviar
    pyautogui.click(button_coord, duration=0.2)

    time.sleep(1)

print("Formulários preenchidos com sucesso!")
