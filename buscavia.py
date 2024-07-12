import os
import locale
import datetime

def buscar_planilha_mes_atual(diretorio):
    # Definir o locale para português do Brasil
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

    # Obtenha o mês e o ano atuais
    hoje = datetime.datetime.now()
    nomeAbrevi = hoje.strftime("%b")  # Nome abreviado do mês, e.g., "Abr"

    # Crie o nome esperado da planilha
    nome_planilha = f"SEMELHANTES VIA - {nomeAbrevi}.xlsx"

    # Caminho completo do arquivo
    caminho_completo = os.path.join(diretorio, nome_planilha)
    
    # Verifique se o arquivo existe
    if os.path.isfile(caminho_completo):
        return caminho_completo
    else:
        print(f"Arquivo não encontrado: {caminho_completo}")
        return None
