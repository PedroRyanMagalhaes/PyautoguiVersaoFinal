#falta implantar ler faturamento 
#


import openpyxl
import pyautogui as pa
import time
import pyperclip
import pytesseract
import cv2
from openpyxl.styles import PatternFill

# Carrega a planilha
wb = openpyxl.load_workbook('20Setembro.xlsx')
sheet = wb.active

def obterTelefone(linha):
    return sheet[f'E{linha}'].value

def pintarDeVerde(linha):
    verde = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    for cell in sheet[linha]:
        cell.fill = verde

def pintarDeVermelho(linha):
    vermelho = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    for cell in sheet[linha]:
        cell.fill = vermelho

def pintarDeLaranja(linha):
    laranja = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
    for cell in sheet[linha]:
        cell.fill = laranja

# Função para verificar o mês
def verificarMes(linha, mesEsperado):
    try:
        screenshot_mes = f'printMes{linha}.png'
        screenshot = pa.screenshot(region=(431, 883, 70, 40))  # Ajuste a região conforme necessário
        screenshot.save(screenshot_mes)

        img = cv2.imread(screenshot_mes)

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

        extracted_text = pytesseract.image_to_string(thresh)
        print(f"Texto extraído para o pagamento na linha {linha}: {extracted_text}")

        return mesEsperado in extracted_text

    except Exception as e:
        print(f"Ocorreu um erro ao verificar o mês na linha {linha}: {str(e)}")
        return False

# Função para verificar o mês em uma nova coordenada
def verificarMesAlternativo(linha, mesEsperado):
    try:
        screenshot_mes = f'printMesAlternativo{linha}.png'
        screenshot = pa.screenshot(region=(433, 971, 70, 50))  # Nova região
        screenshot.save(screenshot_mes)

        img = cv2.imread(screenshot_mes)

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

        extracted_text = pytesseract.image_to_string(thresh)
        print(f"Texto extraído para o pagamento na linha {linha} (alternativo): {extracted_text}")

        return mesEsperado in extracted_text

    except Exception as e:
        print(f"Ocorreu um erro ao verificar o mês alternativo na linha {linha}: {str(e)}")
        return False

# Função para verificar se está "pago" ou "atrasada"
def verificarStatusPagamento(linha):
    try:
        screenshot_status = f'printPagamento{linha}.png'
        screenshot = pa.screenshot(region=(1044, 884, 70, 40))  # Ajuste a região conforme necessário
        screenshot.save(screenshot_status)

        img = cv2.imread(screenshot_status)

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

        extracted_text = pytesseract.image_to_string(thresh)
        print(f"Texto extraído para o pagamento na linha {linha}: {extracted_text}")

        if "paga" in extracted_text.lower():
            print(f"Status 'paga' encontrado na linha {linha}.")
            return "pago"
        elif "atrasada" in extracted_text.lower():
            print(f"Status 'atrasada' encontrado na linha {linha}.")
            return "atrasada"
        else:
            print(f"Status desconhecido na linha {linha}: {extracted_text}")
            return None

    except Exception as e:
        print(f"Ocorreu um erro ao verificar o status de pagamento na linha {linha}: {str(e)}")
        return "erro"

# Função para verificar o status de pagamento em uma nova coordenada
def verificarStatusPagamentoAlternativo(linha):
    try:
        screenshot_status = f'printPagamentoAlternativo{linha}.png'
        screenshot = pa.screenshot(region=(1040,975, 100, 50))  # Ajuste a nova região conforme necessário
        screenshot.save(screenshot_status)

        img = cv2.imread(screenshot_status)

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

        extracted_text = pytesseract.image_to_string(thresh)
        print(f"Texto extraído para o pagamento na linha {linha} (alternativo): {extracted_text}")

        if "paga" in extracted_text.lower():
            print(f"Status 'paga' encontrado na linha {linha}.")
            return "pago"
        elif "atrasada" in extracted_text.lower():
            print(f"Status 'atrasada' encontrado na linha {linha}.")
            return "atrasada"
        else:
            print(f"Status desconhecido na linha {linha}: {extracted_text}")
            return None

    except Exception as e:
        print(f"Ocorreu um erro ao verificar o status de pagamento alternativo na linha {linha}: {str(e)}")
        return "erro"

# Mês esperado
mesEsperado = "09"

comecoLinha = 4
finalLinha = 6

# Muda a tela
pa.hotkey('alt', 'tab')
time.sleep(0.5)

with open('telefone.txt', 'w') as f:
    for linha in range(comecoLinha, finalLinha + 1):
        telefone = obterTelefone(linha)
        f.write(f"linha {linha}, Telefone {telefone}\n")

for linha in range(comecoLinha, finalLinha + 1):
    telefone = obterTelefone(linha)  
    pyperclip.copy(telefone)  

    # Clica no lugar para dar 'ctrl + a'
    pa.click(587, 365)
    pa.hotkey('ctrl', 'a')
    time.sleep(2)

    # Cola o telefone
    pa.hotkey('ctrl', 'v')
    pa.press('enter')
    time.sleep(3)

    # Clicar nos 3 pontinhos
    pa.click(635, 433)
    time.sleep(2)

    # Clicar em detalhes 
    pa.click(451, 485)
    time.sleep(3)

    # Clicar nos três pontinhos novamente
    pa.click(899, 651)
    time.sleep(2)

    # Clicar em faturas
    pa.click(470, 659)
    time.sleep(1)

    # Verificar o mês capturado na tela
    if verificarMes(linha, mesEsperado):
        # Verificar o status de pagamento
        status = verificarStatusPagamento(linha)
        if status == "pago":
            pintarDeVerde(linha)
        elif status == "atrasada":
            pintarDeVermelho(linha)
    else:
        # Se o mês não for o esperado, tenta verificar em uma nova coordenada
        if verificarMesAlternativo(linha, mesEsperado):
            # Verificar o status de pagamento em nova coordenada
            status = verificarStatusPagamentoAlternativo(linha)
            if status == "pago":
                pintarDeVerde(linha)
            elif status == "atrasada":
                pintarDeVermelho(linha)
        else:
            # Se o mês ainda não for o esperado, pinta a linha de laranja
            pintarDeLaranja(linha)
    
    # voltar
    pa.hotkey('alt', 'left')
    time.sleep(0.5)  
    #voltar
    pa.hotkey('alt', 'left')
    time.sleep(0.5)  

# Salvar o arquivo Excel atualizado
wb.save("20SetembroAtualizada.xlsx")
