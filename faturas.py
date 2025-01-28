from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, date
import os
from openpyxl import Workbook
import requests
import time


# Abrindo o navegador e acessando o site
navegador = webdriver.Chrome()
navegador.get("https://rpachallengeocr.azurewebsites.net/")

# Criando pasta de saída dos arquivos
pasta_saida = "C:/Users/Vinicius/Documents/faturas"
os.makedirs(pasta_saida, exist_ok=True)

# Criando a planilha
planilha_faturas = Workbook()
sheet = planilha_faturas.active
sheet.title = "Faturas"
sheet.append(["Número da Fatura", "Data da Fatura", "URL da Fatura"])

# Capturando a data atual para comparação
data_atual = date.today()

# Loop para navegar pelas páginas e não tornar os elementos obsoletos
while True:
    try:
        # Localizando as linhas da tabela
        linhas = WebDriverWait(navegador, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//tbody/tr[@role="row"]'))
        )

        # Iterando sobre cada linha
        for linha in linhas:
            try:
                # Capturando dados da linha
                data_fatura = linha.find_element(By.XPATH, './td[3]').text
                data_fatura_convertida = datetime.strptime(data_fatura, "%d-%m-%Y").date()

                # Verificando se a data da fatura é igual ou maior que a data atual
                if data_fatura_convertida >= data_atual:
                    numero_fatura = linha.find_element(By.XPATH, './td[1]').text
                    url_fatura = linha.find_element(By.XPATH, './td[4]/a').get_attribute('href')

                    # Ajustando nome do arquivo e download de cada um deles
                    nome_arquivo = f"Fatura_{numero_fatura}_data_{data_fatura_convertida}.jpg"
                    arquivo_saida = os.path.join(pasta_saida, nome_arquivo)

                    download_imagem = requests.get(url_fatura)
                    with open(arquivo_saida, "wb") as arquivo:
                        arquivo.write(download_imagem.content)

                    # Salvando os dados na planilha
                    sheet.append([numero_fatura, data_fatura_convertida, url_fatura])
                    print(f"Fatura salva: {numero_fatura}, Data: {data_fatura}, URL: {url_fatura}")

            except Exception as e:
                print(f"Erro ao processar linha: {e}")

        # Verificando se existe o botão "next" habilitado e clicando
        try:
            botao_next = WebDriverWait(navegador, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//a[contains(@class, "paginate_button") and contains(@class, "next")]'))
            )
            # Se no nome da classe houver "disabled", ele para o processo.
            if "disabled" in botao_next.get_attribute("class"):
                print("Fim do processo")
                break
            botao_next.click()
        except Exception as e:
            print(f"Erro ao localizar ou clicar no botão Next: {e}")
            break

    except Exception as e:
        print(f"Erro ao processar a tabela: {e}")
        break

# Salvando a planilha
caminho_planilha_faturas = os.path.join(pasta_saida, "Faturas.xlsx")
planilha_faturas.save(caminho_planilha_faturas)
print(f"Planilha salva em: {caminho_planilha_faturas}")

# Fechando o navegador
navegador.quit()
