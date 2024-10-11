import openpyxl
import pyautogui as pa
import time
import pyperclip
import pytesseract
import cv2

# Testar carregamento da planilha
print("Carregando a planilha...")
wb = openpyxl.load_workbook('20Setembro.xlsx')
sheet = wb.active
print("Planilha carregada com sucesso.")

def obterTelefone(linha):
    telefone = sheet[f'E{linha}'].value
    print(f"Obtendo telefone da linha {linha}: {telefone}")
    return telefone

# Testar o início e fim do intervalo de linhas
comecoLinha = 4
finalLinha = 100
print(f"Processando linhas de {comecoLinha} até {finalLinha}.")

# Testar a parte do loop onde manipula os telefones
for linha in range(comecoLinha, finalLinha + 1):
    telefone = obterTelefone(linha)  
    print(f"Copiando telefone para clipboard: {telefone}")
    pyperclip.copy(telefone)

    # Aqui você pode adicionar outros prints para testar etapas do código
    print("Simulando clique para buscar dados na tela.")
    pa.hotkey('alt', 'tab')  # Mudando de janela (testar se isso acontece corretamente)
    time.sleep(0.5)
    pa.click(278, 77)  # Clique no campo de busca
    pa.hotkey('ctrl', 'a')
    pa.hotkey('ctrl', 'v')
    pa.press('enter')
    print("Telefone colado e enter pressionado.")
    time.sleep(2)

    # Simular clique nos 3 pontinhos
    print("Clicando nos três pontinhos.")
    pa.click(454, 259)
    time.sleep(2)

    # Simular clique em detalhes
    print("Clicando em detalhes.")
    pa.click(636, 403)
    time.sleep(2)

    # Capturar e processar a imagem
    screenshot_filename = f'screenshot_linha_{linha}.png'
    screenshot = pa.screenshot(region=(959, 486, 205, 60))
    screenshot.save(screenshot_filename)
    print(f"Screenshot salva: {screenshot_filename}")

    # Ler a imagem usando o Tesseract
    img = cv2.imread(screenshot_filename)
    extracted_text = pytesseract.image_to_string(img)
    print(f"Texto extraído da imagem: {extracted_text}")

print("Teste concluído.")
