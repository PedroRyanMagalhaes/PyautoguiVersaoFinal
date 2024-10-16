import openpyxl
import pyautogui as pa
import time
import pyperclip
import pytesseract
import cv2
import numpy as np
from openpyxl.styles import PatternFill


wb = openpyxl.load_workbook('20Setembro.xlsx')
sheet = wb.active

def obterTelefone(linha):
    return sheet[f'E{linha}'].value

def pintarDeVerde(linha):
    verde = PatternFill(start_color="00FF00", end_color="00ff00", fill_type="solid")
    for cell in sheet[linha]:
        cell.fill = verde

def pintarDeVermelho(linha):
    vermelho = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    for cell in sheet[linha]:
        cell.fill = vermelho

# Mês esperado vai ser inserido aqui 
numeroEsperado = "Cisne -"

comecoLinha = 4
finalLinha = 5

# Muda a tela
pa.hotkey('alt', 'tab')
time.sleep(0.5)

with open ('telefone.txt', 'w') as f:
    for linha in range(comecoLinha, finalLinha + 1):
        telefone = obterTelefone(linha)
        f.write(f"linha {linha}, Telefone {telefone}\n")

for linha in range(comecoLinha, finalLinha + 1):
    telefone = obterTelefone(linha)  
    pyperclip.copy(telefone)  

    # Clica no lugar para dar 'ctrl + a'
    pa.click(278, 77)
    pa.hotkey('ctrl', 'a')
    time.sleep(0.5)

    # Cola o telefone
    pa.hotkey('ctrl', 'v')
    pa.press('enter')
    time.sleep(2)

    # Clicar nos 3 pontinhos na versão final 
    pa.click(454, 259)
    time.sleep(2)

    # Clicar em detalhes 
    pa.click(636, 403)
    time.sleep(2)

    # Clicar nos três pontinhos novamente 
    pa.click(1776,793)
    time.sleep(2)

    # Clicar em faturas
    pa.click(1589, 1018)
    time.sleep(2)

    # Ler o número que está visível na tela
    screenshot_filename = f'screenshot_linha_{linha}.png'  # Nomeia o screenshot com base na linha
    screenshot = pa.screenshot(region=(959, 486, 205, 60))
    screenshot.save(screenshot_filename)
    print(f"Screenshot salva: {screenshot_filename}")

    # Ler a imagem
    img = cv2.imread(screenshot_filename)

    # Tratar imagem
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

    # Extrair texto da imagem
    extracted_text = pytesseract.image_to_string(img)

    # Verifica se é o mês correto
    if numeroEsperado in extracted_text:
        print(f"O número esperado {numeroEsperado} e eu achei esse {extracted_text}")
        pintarDeVerde(linha)  # Pinta a linha correspondente ao telefone correto
    else:   
        print(f"Não achei o número esperado {numeroEsperado} e eu achei esse {extracted_text}")
        pintarDeVermelho(linha) 

wb.save("20Setembro_1.xlsx")
