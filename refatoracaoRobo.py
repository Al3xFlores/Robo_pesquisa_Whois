from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import time

# Define o caminho do arquivo Excel
caminho_arquivo = 'C:/Users/alexm/Documents/Alex/Robos_Py/pesquisa.xlsx'

# Carrega o arquivo Excel
df = pd.read_excel(caminho_arquivo)

# Inicializa o driver do Selenium
driver = webdriver.Edge('C:/Users/alexm/Documents/Alex/Robos_Py/msedgedriver.exe')  # Substitua pelo caminho do seu msedgedriver.exe

# Abre o site específico
driver.get('https://registro.br/tecnologia/ferramentas/whois/')  # Substitua pela URL do site desejado
time.sleep(2)

# Localiza o input onde a informação será inserida
input_pesquisa = driver.find_element('xpath', '//*[@id="whois-field"]')

# Percorre as linhas do arquivo Excel
for index, row in df.iterrows():
    # Obtém a informação da coluna desejada
    informacao = row['IP']
    
    # Insere a informação no input
    input_pesquisa.clear()
    input_pesquisa.send_keys(informacao)
    
    # Clica no botão de pesquisa
    botao_pesquisa = driver.find_element('xpath', '//*[@id="app"]/div/main/div/section/div[1]/form/fieldset/button')
    botao_pesquisa.click()
    
    time.sleep(2)  # Aguarda 5 segundos
    
    try:
        # Verifica se ocorreu uma exceção indicando que o elemento resultante não foi encontrado
        informacao_resultante = driver.find_element('xpath', '//*[@id="app"]/div/main/div/section/div[2]/div[2]/div/div/div[1]/div/table/tbody/tr[4]/td/span').text
        
        # Copia a informação resultante para a coluna ao lado
        df.at[index, 'coluna_destino'] = informacao_resultante
    
    except NoSuchElementException:
        # Se a exceção foi capturada, significa que ocorreu um erro na pesquisa
        # Adicione a lógica para tratar a situação de erro, por exemplo, definir o valor da coluna_destino como 'Erro na pesquisa'
        df.at[index, 'coluna_destino'] = 'Erro na pesquisa'

# Fecha o driver do Selenium
driver.quit()

# Salva o resultado no mesmo arquivo Excel
df.to_excel(caminho_arquivo, index=False)
