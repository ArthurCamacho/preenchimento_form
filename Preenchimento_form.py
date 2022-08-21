from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from tkinter import messagebox


def preenche_campos():
    """Inicia o navegador
       Navega para o site do desafio
       Clica no botão "START"
       Faz o preenchimento dos campos com base na planilha excel"""

    # Clica no botão de "START"
    navegador.find_element(By.CSS_SELECTOR,
                           'body > app-root > div.body.row1.scroll-y > app-rpa1 > div > '
                           'div.instructions.col.s3.m3.l3.uiColorSecondary > div:nth-child(7) > button').click()

    # Faz o preenchimento dos campos e ao final clica no botão para avançar para o próximo registro
    for individuo in lista_final.values:
        navegador.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelLastName"]').send_keys(individuo[1])

        navegador.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelEmail"]').send_keys(individuo[5])

        navegador.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelAddress"]').send_keys(individuo[4])

        navegador.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelPhone"]').send_keys(individuo[6])

        navegador.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelCompanyName"]').send_keys(individuo[2])

        navegador.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelFirstName"]').send_keys(individuo[0])

        navegador.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelRole"]').send_keys(individuo[3])

        navegador.find_element(By.CSS_SELECTOR,
                               'body > app-root > div.body.row1.scroll-y > app-rpa1 > div > '
                               'div.inputFields.col.s6.m6.l6 > form > input').click()


def exibe_resultados():
    """Captura as mensagens com as estatísticas do preenchimento
       Fecha o navegador
       Exibe as estatísticas na tela"""

    # Armazena as mensagens em uma lista, sendo o indíce 0 o título e o indíce 1 o corpo da mensagem
    info = [
        navegador.find_element(By.CSS_SELECTOR,
                               'body > app-root > div.body.row1.scroll-y > app-rpa1 > div > '
                               'div.congratulations.col.s8.m8.l8 > div.message1.teal-text.text-darken-2').text,
        navegador.find_element(By.CSS_SELECTOR,
                               'body > app-root > div.body.row1.scroll-y > app-rpa1 > div > '
                               'div.congratulations.col.s8.m8.l8 > div.message2').text
    ]
    # Fecha o navegador
    navegador.quit()
    # Exibe a mensagem na tela
    messagebox.showinfo(info[0], info[1])


# Faz a leitura da planilha
lista_final = pd.read_excel('Dados.xlsx')
# Inicia e navega para o site do desafio
navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
navegador.get('https://www.rpachallenge.com')
preenche_campos()
exibe_resultados()
