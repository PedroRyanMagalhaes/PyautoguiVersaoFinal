#alfa rodar estorno

import openpyxl
import datetime
import pyautogui as pa  
import time
import pyperclip
import pytesseract
import cv2
from openpyxl.styles import PatternFill
import os
from openpyxl import load_workbook
import re
from difflib import get_close_matches

# Carrega a planilha
wb = openpyxl.load_workbook('estornos.xlsx')
sheet = wb.active

import pyautogui as pa
import time

import pyautogui as pa
import time
import os

def esperar_carregamento(imagem_carregando, confiança=0.8,pasta_salvamento='capturas', max_tentativas=100):

    # Cria a pasta para salvar as capturas, se não existir

    tentativas = 0

    while tentativas < max_tentativas:
        try:
            posicao = pa.locateOnScreen(imagem_carregando, confidence=confiança)


            if posicao is None:
                break
            else:
                time.sleep(0.2)  # Verifica a cada 100 milissegundos
                tentativas += 1  # Incrementa o contador de tentativas
        except Exception as e:
            print(f"Ocorreu um erro: {e}")
            break  # Em caso de erro, saia do loop
    
    if tentativas >= max_tentativas:
        print("Número máximo de tentativas alcançado. Saindo do loop.")


def criarPastaParaLinha(linha):
    pasta_ = f'prints{linha}'
    if not os.path.exists(pasta_):
        os.makedirs(pasta_)
    return pasta_

def verificarStatus(linha):
    status = verificarStatusPagamento(linha) or verificarStatusPagamentoAlternativo(linha)
    return status

def obterTelefone(linha):
    return sheet[f'B{linha}'].value

def pintarDeVerde(linha):
    verde = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
    for cell in sheet[linha]:
        cell.fill = verde
    

def pintarDeVermelho(linha):
    vermelho = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    for cell in sheet[linha]:
        cell.fill = vermelho
    

def pintarDeAmarelo(linha):
    amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for cell in sheet[linha]:
        cell.fill = amarelo

# Função para verificar o mês
def verificarMes(linha, mesEsperado):
    try:
        pasta_ = criarPastaParaLinha(linha)
        screenshot_mes = os.path.join(pasta_, f'{linha}_printMes.png')
        screenshot = pa.screenshot(region=(431, 883, 70, 60))  # Ajuste a região conforme necessário
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
        pasta_ = criarPastaParaLinha(linha)
        screenshot_mes = os.path.join(pasta_, f'{linha}_printMesAlternativo.png')
        screenshot = pa.screenshot(region=(433, 971, 70, 80))  # Nova região
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
        pasta_ = criarPastaParaLinha(linha)
        screenshot_status = os.path.join(pasta_, f'{linha}_printPagamento.png')
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
        pasta_ = criarPastaParaLinha(linha)
        screenshot_status = os.path.join(pasta_, f'{linha}_printPagamentAlternativo.png')
        screenshot = pa.screenshot(region=(1040,975, 120, 50))  # Ajuste a nova região conforme necessário
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
        pasta_ = criarPastaParaLinha(linha)
        screenshot_faturamento = os.path.join(pasta_, f'{linha}_printFaturamento.png')
        screenshot = pa.screenshot(region=(408,640,400,80))  # Ajuste essa coordenada para a região correta
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
        pasta_ = criarPastaParaLinha(linha)
        screenshot_faturamento_alt = os.path.join(pasta_, f'{linha}_printFaturamentoAlt.png')
        screenshot = pa.screenshot(region=(386,479, 400, 250))  # Ajuste essa coordenada para a nova região
        screenshot.save(screenshot_faturamento_alt)

        img = cv2.imread(screenshot_faturamento_alt)

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

        extracted_text = pytesseract.image_to_string(thresh)
        print(f"Texto extraído para faturamento na linha {linha} (alternativo): {extracted_text}")

        if "faturamento" in extracted_text.lower():
            print(f"'Faturamento' encontrado na linha Alt {linha}.")
            return True
        else:
            return False

    except Exception as e:
        print(f"Ocorreu um erro ao verificar o 'faturamento' alternativo na linha {linha}: {str(e)}")
        return False
    
def verificarSaldo(linha):
    try:
        pasta_ = criarPastaParaLinha(linha)
        screenshot_saldo = os.path.join(pasta_, f'{linha}_printSaldo.png')
        screenshot = pa.screenshot(region=(400,640,400,100))  # Ajuste essa coordenada para a região correta
        screenshot.save(screenshot_saldo)

        img = cv2.imread(screenshot_saldo)

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

        extracted_text = pytesseract.image_to_string(thresh)
        print(f"Texto extraído para saldo na linha {linha}: {extracted_text}")

        if "saldo" in extracted_text.lower():
            print(f"'Saldo' encontrado na linha {linha}.")
            return True
        else:
            return False

    except Exception as e:
        print(f"Ocorreu um erro ao verificar o 'saldo' na linha {linha}: {str(e)}")
        return False

    
contador = 1

def clicarTresPontos(imagem_tres_pontos, regiao=None):
    
    # Aguarda um momento para garantir que a tela esteja pronta
    time.sleep(1)

    pasta_ = criarPastaParaLinha(linha)

    # Captura a tela ou a região especificada
    screenshot = pa.screenshot(region=regiao)

    nome_arquivo = os.path.join(pasta_, f"{linha}trespontos.png")

    screenshot.save(nome_arquivo)

    posicao = None

    # Tenta localizar a imagem na tela
    try:
        posicao = pa.locateOnScreen(imagem_tres_pontos, confidence=0.5, region=regiao)
    except Exception as e:
        print(f"Erro ao tentar localizar a imagem: {e}")
        return False  # Retorna False se ocorrer qualquer erro

    if posicao is not None:
        # Clica no centro da imagem encontrada
        pa.click(pa.center(posicao))
        print("Clicou nos três pontinhos.")
        return True
    else:
        print("Imagem dos três pontinhos não encontrada.")
        return False
   


    
def clicarfaturas(imagem_faturas, regiao=None):
    
    
    # Aguarda um momento para garantir que a tela esteja pronta
    time.sleep(1)

    pasta_ = criarPastaParaLinha(linha)

    # Captura a tela ou a região especificada
    screenshot = pa.screenshot(region=regiao)

    nome_arquivo = os.path.join(pasta_, f"{linha}faturas.png")

    screenshot.save(nome_arquivo)

    # Verifica se a imagem dos faturas está presente na captura
    posicao = pa.locateOnScreen(imagem_faturas, confidence=0.5, region=regiao)
    
    if posicao is not None:
        # Clica no centro da imagem encontrada
        pa.click(pa.center(posicao))
        print("Clicou em faturas.")
        return True
    else:
        print("Imagem dos faturas não encontrada.")
        return False

def verificarLinhaNaoLocalizada(linha):
        pasta_ = criarPastaParaLinha(linha)
        screenshot_linhanaolocalizada = os.path.join(pasta_, f'{linha}_LinhaNaoLocalizada.png')
        screenshot = pa.screenshot(region=(774,419,400,400))  # Ajuste essa coordenada para a região correta
        screenshot.save(screenshot_linhanaolocalizada)

        img = cv2.imread(screenshot_linhanaolocalizada)

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

        extracted_text = pytesseract.image_to_string(thresh)
        print(f"Texto extraído na linha {linha}: {extracted_text}")

        if ("linha não localizada" in extracted_text.lower() or 
            "localizada" in extracted_text.lower() or 
            "venda avulsa" in extracted_text.lower()):
             print(f"Essa linha é não localizada")
        else:
            return False



# Função para verificar status "Suspenso" ou "Pago" e atualizar a planilha na coluna especificada


import os
import cv2
import pytesseract
import pyautogui as pa

def verificarSuspensoEPago(linha):
    pasta_ = criarPastaParaLinha(linha)
    screenshot_status = os.path.join(pasta_, f'{linha}_status.png')
    screenshot = pa.screenshot(region=(968, 330, 150, 60))  # Ajuste essa coordenada para a região correta
    screenshot.save(screenshot_status)

    # Carregar e processar a imagem
    img = cv2.imread(screenshot_status)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)  # Suavização para reduzir ruído
    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)  # Binarização adaptativa

    # Configuração do pytesseract para melhorar a precisão
    config = r'--psm 6 -c tessedit_char_whitelist=abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'

    # Extrair o texto e converter para minúsculas
    extracted_text = pytesseract.image_to_string(thresh, config=config).lower().strip()
    print(f"Texto extraído na linha {linha}: {extracted_text}")

    # Verificar diretamente o conteúdo extraído
    if "suspenso" in extracted_text:
        print("Essa linha é suspensa.")
        sheet[f'L{linha}'] = "Suspenso"  # Atualiza a coluna L da linha correspondente
        return True
    elif "pago" in extracted_text:
        print("Essa linha é paga.")
        sheet[f'L{linha}'] = "Pago"  # Atualiza a coluna L da linha correspondente
        return True
    else:
        print("Status não reconhecido.")
        return False



def verificarData(linha, region, coluna):
    try:
        # Cria a pasta para salvar a imagem com o nome da linha
        pasta_ = criarPastaParaLinha(linha)
        screenshot_data = os.path.join(pasta_, f'{linha}_printdata.png')
        
        # Tira um print na região especificada
        screenshot = pa.screenshot(region=region)
        screenshot.save(screenshot_data)

        # Carrega a imagem e faz o processamento
        img = cv2.imread(screenshot_data)

        # Pré-processamento da imagem para melhorar a detecção de texto
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (3, 3), 0)  # Suavizar para reduzir ruído
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)  # Binarização adaptativa

        # Ajuste das configurações do pytesseract para melhorar a detecção
        config = r'--psm 6 -c tessedit_char_whitelist=0123456789/'

        # Extrai o texto da imagem
        extracted_text = pytesseract.image_to_string(thresh, config=config)
        print(f"data extraído na linha {linha}: {extracted_text}")

        # Procura qualquer data no formato "dd/mm/yyyy"
        date_pattern = r'\b\d{2}/\d{2}/\d{4}\b'
        matched_dates = re.findall(date_pattern, extracted_text)

        if matched_dates:
            # Pega a primeira data encontrada e registra na coluna especificada
            data_encontrada = matched_dates[0]
            sheet[f'{coluna}{linha}'] = data_encontrada
            print(f"Data encontrada e registrada na linha {linha}, coluna {coluna}: {data_encontrada}")
            return True  # Indica que uma data foi encontrada

        # Se nenhuma data for encontrada na região especificada
        print(f"Nenhuma data encontrada na linha {linha} na região {region}.")
        return False

    except Exception as e:
        print(f"Ocorreu um erro ao verificar o mês na linha {linha}: {str(e)}")
        return False





# Mês esperado
mesEsperado = "10"


comecoLinha = 4
finalLinha = 4

horario_inicial = datetime.datetime.now().strftime("%H:%M:%S")


# Muda a tela
pa.hotkey('alt', 'tab')
time.sleep(0.3)

with open('telefone.txt', 'w') as f:
    for linha in range(comecoLinha, finalLinha + 1):
        telefone = obterTelefone(linha)
        f.write(f"linha {linha}, Telefone {telefone}\n")



for linha in range(comecoLinha, finalLinha + 1):
    telefone = obterTelefone(linha)  
    pyperclip.copy(telefone)  


    # Clica no lugar para dar 'ctrl + a'
    esperar_carregamento('assets/PCcarregando.png',0.5)
    pa.click(681,294)
    pa.hotkey('ctrl', 'a')
    time.sleep(0.3)

    # Cola o telefone
    pa.hotkey('ctrl', 'v')
    pa.press('enter')
    time.sleep(0.3)
    

    esperar_carregamento('assets/PCcarregando.png',0.5)
    

    # Clicar nos 3 pontinhos
    time.sleep(0.5)
    resultado = clicarTresPontos(imagem_tres_pontos='assets/PCtrespontos.PNG', regiao=(674,334, 100, 100))
    
    if resultado:
        time.sleep(0.3)
    else:
        pa.click(567,604)
        resultado = clicarTresPontos(imagem_tres_pontos='assets/PCtrespontos.PNG', regiao=(674,334, 100, 100))
        if resultado:
            time.sleep(0.3)
        else:
            verificarLinhaNaoLocalizada(linha)
            pintarDeVermelho(linha)
            pa.hotkey('alt', 'left')
            time.sleep(0.5)
            print(f"Processamento da linha {linha} = linha nao localizada.")
            time.sleep(0.5)
            continue
            
            

    # Clicar em detalhes 
    pa.click(549,385)
    time.sleep(0.5)
    esperar_carregamento('assets/carregando.jpg',0.5)
    time.sleep(1.5)

    if verificarSuspensoEPago(linha):
        print ("Encontrei")
        time.sleep(1)
        

    
    if verificarSaldo(linha):
            pintarDeVermelho(linha)
            print(f'Saldo encontrado na linha {linha}. Linha pintada de vermelho.')
            pa.hotkey('alt', 'left')  # Voltar com Alt + Left
            #time.sleep(2)
            esperar_carregamento('assets/carregando.jpg',0.5)
            continue  # Volta para o início do for
    
        


  # Verificar "faturamento" após clicar em detalhes
    if verificarFaturamento(linha):
        imagem_tres_pontos = 'assets/imagemtrespontos.jpg'  # Substitua pelo caminho real da sua imagem
        regiao = (829, 612, 200, 280)  # Substitua pelas coordenadas (x, y, largura, altura) região que deseja capturar
    
    
        if clicarTresPontos(imagem_tres_pontos, regiao):
            time.sleep(0.5)  # Aguardar um tempo após clicar nos três pontos

        # Clicar em faturas
        imagem_faturas = 'assets/imagemfaturas.jpg'  # Clicar em faturas
        regiao = (393, 623, 200, 200)

        if clicarfaturas(imagem_faturas, regiao):
            time.sleep(0.5)
            pa.scroll(-500)
        
         
    elif verificarSaldo(linha):
        pa.hotkey('alt', 'left')
        pintarDeVermelho(linha)
        continue

    else:
        pa.scroll(-500)  # Scroll para baixo
        time.sleep(0.5)

    # Se "faturamento" não for encontrado, verificar na coordenada alternativa
    if verificarFaturamentoAlternativo(linha):
        # Se encontrado na coordenada alternativa, realizar scroll e cliques adicionais

            imagem_tres_pontos = 'assets/imagemtrespontos.jpg'  # Clique em nova coordenada
            regiao = (832, 483, 250, 280)

            if clicarTresPontos(imagem_tres_pontos, regiao):
                time.sleep(0.5)

            imagem_faturas = 'assets/imagemfaturas.jpg'  # Outro clique
            regiao = (392, 493, 250, 380)

            if clicarfaturas(imagem_faturas, regiao):
                time.sleep(0.5)

            pa.scroll(-500)  # Scroll para cima
            time.sleep(0.5)

    elif verificarSaldo(linha):
        pa.hotkey('alt', 'left')
        pintarDeVermelho(linha)
        continue

    else:
        # Se "faturamento" não for encontrado em nenhuma das coordenadas, pular linha
            print(f"'Faturamento' não encontrado na linha alternativa {linha}.")
         # Pula para a próxima linha se nenhuma verificação passar


    if verificarData(linha,(518,671,150,40),'M'):
        #verificarStatusPagamentoo()
        time.sleep (1)
        #if verificarData(linha,(518,748,150,60),'N'):
            #verificarStatusPagamentoo2()
         #   time.sleep(0.5)
        #else:
         #   print ("nao achei data de segunda")
        #if verificarData(linha, (520,825,150,60),'P'):
            #verificarStatusPagamentoo3()
         #   time.sleep(0.5)
        #else:
         #   print ("nao achei data de terceira")
    else:
        print ("nao achei nada de data")
        

    pa.hotkey('alt', 'left')
    time.sleep(0.5)
    pa.hotkey('alt', 'left')
    #time.sleep(7)
    #esperar_carregamento('assets/carregando.jpg',0.5)
    wb.save("ESTORNONOVO.xlsx")

print (f"Começou às {horario_inicial}")
horario_final = datetime.datetime.now().strftime("%H:%M:%S")

print(f"Processo finalizado para todas as linhas às {horario_final}")









    #if verificarMes(linha, mesEsperado):
    # Verificar o status de pagamento
        #status = verificarStatus(linha)
        #print(f"Status para a linha {linha}: {status}")
        #if status == "paga":
        #    pintarDeVerde(linha)
        #elif status == "atrasada":
        #    pintarDeAmarelo(linha)
        #else:
        #    pintarDeVermelho(linha)
        #    print ("Pintei de vermelho, verificar depois por favor")
    #else:
    # Se o mês não for o esperado, tenta verificar em uma nova coordenada
       # if verificarMesAlternativo(linha, mesEsperado):
        # Verificar o status de pagamento na coordenada alternativa
         #   status = verificarStatusPagamentoAlternativo(linha)
         #   print(f"Status alternativo para a linha {linha}: {status}")
         #   if status == "paga":
         #       pintarDeVerde(linha)
          #  elif status == "atrasada":
          #      pintarDeAmarelo(linha)
        #else:
        ## Se o mês ainda não for o esperado, pinta a linha de laranja
         #   pintarDeVermelho(linha)
          #  print ("Pintei de vermelho, verificar depois por favor")
        

# Voltar para a tela anterior duas vezes
    #pa.hotkey('alt', 'left')
    #time.sleep(0.5)
    #pa.hotkey('alt', 'left')
    ##time.sleep(7)
    #esperar_carregamento('assets/carregando.jpg',0.5)
    #wb.save("ESTORNOATUALIZADA.xlsx")

#print (f"Começou às {horario_inicial}")
#horario_final = datetime.datetime.now().strftime("%H:%M:%S")

#print(f"Processo finalizado para todas as linhas às {horario_final}")