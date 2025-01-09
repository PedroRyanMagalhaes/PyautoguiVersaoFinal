import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time

# Ler a planilha com pandas
planilha = pd.read_excel("NOV2.xlsx")

# Configurar o driver do Chrome corretamente
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# Iterar pelas linhas da planilha
for index, row in planilha.iterrows():
    telefone = row['Telefone']  # Certifique-se de que "E" é o nome da coluna correta

    # Acessar o site
    driver.get("https://vendasapp.claro.com.br/SVCv2/posicionamento/resultado-pesquisa")

    time.sleep(10)

    # Encontrar o campo de pesquisa e inserir o número
    campo_pesquisa = driver.find_element(By.ID, "filtro")
    campo_pesquisa.send_keys(str(telefone))

    # Enviar o formulário
    campo_pesquisa.submit()

    # Esperar 3 segundos para carregar a página
    time.sleep(3)

    # Mensagem de status
    print(f"Pesquisa realizada para o telefone: {telefone}")

# Fechar o navegador
driver.quit()
