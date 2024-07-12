import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import datetime
import time
import re

# Definir os códigos dos produtos
codigos = [ 1500981518,1512965671,1502840900,1525437968,1545628364,1545512765,1500952265,1501694721]

# Definir os CEPs por estado
estados_cidades_ceps = {
    "SP": "01002903",
    "RJ": "20020022",
    "BA":"40020015",
    "PR": "80010020",
    "DF": "70040252",
    "MG": "30110000",
    "PE": "50010040",
    "GO": "74565920",
    "TO": "77001074",
    "PA": "66010095"
}

# Configurar o driver do Selenium
def setup_driver():
    options = Options()
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--no-sandbox")
    options.add_argument('log-level=3')
    driver = webdriver.Chrome(options=options)
    return driver

# Função para calcular dias úteis
def calcular_dias_uteis(data_inicial, prazo_str):
    # Extrair a data do prazo_str usando regex
    match = re.search(r'(\d{2}) de (\w+)', prazo_str)
    if not match:
        print(f"Formato de data inválido: {prazo_str}")
        return None
    
    dia = int(match.group(1))
    mes = match.group(2).lower()
    ano = data_inicial.year
    
    # Dicionário para converter mês em número
    meses = {
        'janeiro': 1, 'fevereiro': 2, 'março': 3, 'abril': 4,
        'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8,
        'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12
    }
    
    if mes not in meses:
        print(f"Mês inválido: {mes}")
        return None
    
    prazo_data = datetime.datetime(ano, meses[mes], dia)
    dias = pd.date_range(start=data_inicial, end=prazo_data)
    dias_uteis = dias[~dias.weekday.isin([5, 6])]
    return len(dias_uteis)

# Tabela para armazenar os dados
dados_frete = []

# Iterar pelos códigos dos produtos
for codigo in codigos:
    url = f"https://www.casasbahia.com.br/cozinha-compacta-madesa-emilly-top-com-armario-e-balcao-rustic-preto/p/{codigo}"
    driver = setup_driver()
    driver.get(url)
    wait = WebDriverWait(driver, 15)

    # Iterar pelos CEPs de cada estado
    for estado, cep in estados_cidades_ceps.items():
        try:
            cep_input = wait.until(EC.element_to_be_clickable((By.ID, 'frete')))
            cep_input.clear()
            cep_input.send_keys(cep)
            cep_input.send_keys(Keys.RETURN)
           
        except TimeoutException:
            print(f"Erro: Não foi possível interagir com o campo de CEP para o código {codigo} e estado {estado}.")
            continue

        try:
            freight_info = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="freight-price-value"]'))).text
            prazo_info_element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="cep-box"]/div[3]/div/div/div/div/div/span')))
            prazo_info = prazo_info_element.text if prazo_info_element else ""
            
            data_inicial = datetime.datetime.today()
            dias_uteis = calcular_dias_uteis(data_inicial,prazo_info)+1 if prazo_info else ""
            
            dados_frete.append({
                'Código': codigo,
                'Estado': estado,
                'CEP': cep,
                'Frete': freight_info,
                'Dias Úteis': dias_uteis
            })
        except TimeoutException:
            print(f"Erro: Não foi possível encontrar as informações de frete para o código {codigo} e estado {estado}.")
        except Exception as e:
            print(f"Erro ao processar o código {codigo} e estado {estado}: {e}")
    
    driver.quit()

# Criar um DataFrame a partir dos dados coletados
df_frete = pd.DataFrame(dados_frete)
print(df_frete)

# Salvar os dados em um arquivo Excel
output_path = 'valores_frete_semelhantes.xlsx'
df_frete.to_excel(output_path, index=False)
print(f"Dados de frete salvos em {output_path}.")
