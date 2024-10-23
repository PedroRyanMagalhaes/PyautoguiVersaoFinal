

import openpyxl
import pyautogui as pa  
import time
import pyperclip
import pytesseract
import cv2
from openpyxl.styles import PatternFill


# Carrega a planilha
wb = openpyxl.load_workbook('out1.21.xlsx')
sheet = wb.active

def verificarStatus(linha):
    status = verificarStatusPagamento(linha) or verificarStatusPagamentoAlternativo(linha)
    return status

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
        screenshot_mes = f'{linha}printMes.png'
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
        screenshot_mes = f'{linha}printMesAlternativo.png'
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
        screenshot_status = f'{linha}printPagamento.png'
        screenshot = pa.screenshot(region=(1044, 884, 100, 40))  # Ajuste a região conforme necessário
        screenshot.save(screenshot_status)

        img = cv2.imread(screenshot_status)

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

        extracted_text = pytesseract.image_to_string(thresh)
        print(f"Texto extraído para o pagamento na linha {linha}: {extracted_text}")

        if "paga" in extracted_text.lower():
            print(f"Status 'paga' encontrado na linha {linha}.")
            return "paga"
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
        screenshot_status = f'{linha}printPagamentoAlternativo.png'
        screenshot = pa.screenshot(region=(1040,975, 100, 50))  # Ajuste a nova região conforme necessário
        screenshot.save(screenshot_status)

        img = cv2.imread(screenshot_status)

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

        extracted_text = pytesseract.image_to_string(thresh)
        print(f"Texto extraído para o pagamento na linha {linha} (alternativo): {extracted_text}")

        if "paga" in extracted_text.lower():
            print(f"Status 'paga' encontrado na linha {linha}.")
            return "paga"
        elif "atrasada" in extracted_text.lower():
            print(f"Status 'atrasada' encontrado na linha {linha}.")
            return "atrasada"
        else:
            print(f"Status desconhecido na linha {linha}: {extracted_text}")
            return None

    except Exception as e:
        print(f"Ocorreu um erro ao verificar o status de pagamento alternativo na linha {linha}: {str(e)}")
        return "erro"

def verificarFaturamento(linha):
    try:
        screenshot_faturamento = f'{linha}printFaturamento.png'
        screenshot = pa.screenshot(region=(408,640,400, 50))  # Ajuste essa coordenada para a região correta
        screenshot.save(screenshot_faturamento)

        img = cv2.imread(screenshot_faturamento)

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

        extracted_text = pytesseract.image_to_string(thresh)
        print(f"Texto extraído para faturamento na linha {linha}: {extracted_text}")

        if "faturamento" in extracted_text.lower():
            print(f"'Faturamento' encontrado na linha {linha}.")
            return True
        else:
            print(f"'Faturamento' não encontrado na linha {linha}.")
            return False

    except Exception as e:
        print(f"Ocorreu um erro ao verificar o 'faturamento' na linha {linha}: {str(e)}")
        return False

# Função para verificar "faturamento" em uma coordenada alternativa
def verificarFaturamentoAlternativo(linha):
    try:
        screenshot_faturamento_alt = f'{linha}printFaturamentoAlternativo.png'
        screenshot = pa.screenshot(region=(472,544, 400, 50))  # Ajuste essa coordenada para a nova região
        screenshot.save(screenshot_faturamento_alt)

        img = cv2.imread(screenshot_faturamento_alt)

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

        extracted_text = pytesseract.image_to_string(thresh)
        print(f"Texto extraído para faturamento na linha {linha} (alternativo): {extracted_text}")

        if "faturamento" in extracted_text.lower():
            print(f"'Faturamento' encontrado na linha {linha}.")
            return True
        else:
            print(f"'Faturamento' não encontrado na linha {linha}.")
            return False

    except Exception as e:
        print(f"Ocorreu um erro ao verificar o 'faturamento' alternativo na linha {linha}: {str(e)}")
        return False
    
contador = 1

def clicarTresPontos(imagem_tres_pontos, regiao=None):
    global contador
    # Aguarda um momento para garantir que a tela esteja pronta
    time.sleep(1)

    # Captura a tela ou a região especificada
    screenshot = pa.screenshot(region=regiao)

    nome_arquivo = f"{contador}trespontos.png"
    screenshot.save(nome_arquivo)
    print(f"Captura de tela salva como {nome_arquivo}")

    # Verifica se a imagem dos três pontinhos está presente na captura
    posicao = pa.locateOnScreen(imagem_tres_pontos, confidence=0.5, region=regiao)

    if posicao is not None:
        # Clica no centro da imagem encontrada
        pa.click(pa.center(posicao))
        print("Clicou nos três pontinhos.")
        contador +=1
        return True
    else:
        print("Imagem dos três pontinhos não encontrada.")
        return False
    
def clicarfaturas(imagem_faturas, regiao=None):
    global contador
    
    # Aguarda um momento para garantir que a tela esteja pronta
    time.sleep(1)

    # Captura a tela ou a região especificada
    screenshot = pa.screenshot(region=regiao)

    nome_arquivo = f"{contador}faturas.png"
    screenshot.save(nome_arquivo)
    print(f"Captura de tela salva como {nome_arquivo}")

    # Verifica se a imagem dos faturas está presente na captura
    posicao = pa.locateOnScreen(imagem_faturas, confidence=0.5, region=regiao)

    if posicao is not None:
        # Clica no centro da imagem encontrada
        pa.click(pa.center(posicao))
        print("Clicou em faturas.")
        contador += 1
        return True
    else:
        print("Imagem dos faturas não encontrada.")
        return False


# Mês esperado
mesEsperado = "10"

comecoLinha = 2
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
    time.sleep(5)

  # Verificar "faturamento" após clicar em detalhes
    if verificarFaturamento(linha):
        imagem_tres_pontos = 'imagemtrespontos.jpg'  # Substitua pelo caminho real da sua imagem
        regiao = (829, 612, 200, 200)  # Substitua pelas coordenadas (x, y, largura, altura) região que deseja capturar
    
    
        if clicarTresPontos(imagem_tres_pontos, regiao):
            time.sleep(2)  # Aguardar um tempo após clicar nos três pontos

        # Clicar em faturas
        imagem_faturas = 'imagemfaturas.jpg'  # Clicar em faturas
        regiao = (393, 623, 200, 200)

        if clicarfaturas(imagem_faturas, regiao):
            time.sleep(2)
    else:
        pa.scroll(-500)  # Scroll para baixo
        time.sleep(1)

    # Se "faturamento" não for encontrado, verificar na coordenada alternativa
    if verificarFaturamentoAlternativo(linha):
        # Se encontrado na coordenada alternativa, realizar scroll e cliques adicionais

            pa.click(897, 554)  # Clique em nova coordenada
            time.sleep(2)

            pa.click(478, 592)  # Outro clique
            time.sleep(2)

            pa.scroll(500)  # Scroll para cima
            time.sleep(1)
        

    else:
        # Se "faturamento" não for encontrado em nenhuma das coordenadas, pular linha
            print(f"'Faturamento' não encontrado na linha {linha}. Pulando para a próxima.")
         # Pula para a próxima linha se nenhuma verificação passar

# Verificar o mês capturado na tela
    if verificarMes(linha, mesEsperado):
    # Verificar o status de pagamento
        status = verificarStatus(linha)
        print(f"Status para a linha {linha}: {status}")
        if status == "paga":
            pintarDeVerde(linha)
        elif status == "atrasada":
            pintarDeVermelho(linha)
        else:
            pintarDeLaranja(linha)
    else:
    # Se o mês não for o esperado, tenta verificar em uma nova coordenada
        if verificarMesAlternativo(linha, mesEsperado):
        # Verificar o status de pagamento na coordenada alternativa
            status = verificarStatusPagamentoAlternativo(linha)
            print(f"Status alternativo para a linha {linha}: {status}")
            if status == "paga":
                pintarDeVerde(linha)
            elif status == "atrasada":
                pintarDeVermelho(linha)
        else:
        # Se o mês ainda não for o esperado, pinta a linha de laranja
            pintarDeLaranja(linha)
        

# Voltar para a tela anterior duas vezes
    pa.hotkey('alt', 'left')
    time.sleep(0.5)
    pa.hotkey('alt', 'left')
    time.sleep(3)

# Salvar o arquivo Excel atualizado
wb.save("out1.21Atualizada.xlsx")

