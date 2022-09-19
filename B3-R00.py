from math import floor
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import PySimpleGUI as sg
Contador = 0
df = []
sg.theme('Dark Blue 3')  # please make your windows colorful
layout = [
            [sg.Text('Por favor faça seu login caixa:')],
            [sg.Text('Matricula', size=(15, 1)), sg.InputText()],
            [sg.Text('Senha', size=(15, 1)), sg.InputText(password_char='*')],
            [sg.Text('WebDriver', size=(15, 1)), sg.InputText(), sg.FileBrowse()],
            [sg.Text('Planilha', size=(15, 1)), sg.InputText(), sg.FileBrowse()],
            [sg.Submit(), sg.Cancel()]
            ]

window = sg.Window('Automatização de aplicações', layout)
event, values = window.read()
matricula, senha, web_driver, planilha = values[0], values [1], values[2], values[3]
window.close()
navegador = webdriver.Chrome(executable_path = web_driver)
Planilhacaixa = pd.read_excel(planilha, "Lista",skiprows = [0])
print(Planilhacaixa)
    # link
navegador.get("https://siart.caixa/siart/SiartController/menu/login")
#acessar sistema
navegador.find_element(By.NAME, "login").send_keys(matricula)
navegador.find_element(By.NAME, "password").send_keys(senha + Keys.ENTER)
#link da consulta
WebDriverWait(navegador,60).until(EC.invisibility_of_element_located((By.NAME, "login")))
for i, cont in enumerate(Planilhacaixa['Conta']):
  navegador.get("https://siart.caixa/siart/SiartController/menu/consulta")
  dados = cont.split(".")
  Conta = str(dados[2]).zfill(14).replace("-",'')
  Operacao = str(dados[1]).zfill(4)
  agencia = str(dados[0]).zfill(4)
  COD = str(Planilhacaixa.loc[i,'Fundo']).zfill(4)
  ValorApl = Planilhacaixa.loc[i, 'SL Conta']
  ValorApl = f'{ValorApl:.2f}'
  print(agencia, Operacao, Conta, COD, ValorApl)
  #dados planilha agencia conta e tipo
  navegador.find_element(By.NAME, "agencia").send_keys(agencia)
  navegador.find_element(By.NAME, "conta").send_keys(Conta)
  TipoConta = navegador.find_element(By.NAME, "operacao")
  drop = Select(TipoConta)
  drop.select_by_value(Operacao)
  navegador.find_element(By.XPATH, "/html/body/table[2]/tbody/tr/td[2]/form/table/tbody/tr[7]/td[4]/input").click()
  WebDriverWait(navegador,60).until(EC.invisibility_of_element_located((By.NAME, "agencia")))
  #clicar em consultar
  try: #tentar selecionar o tipo de fundo
    navegador.get("https://siart.caixa/siart/SiartController/aplic/aplic_msg?tipoFundo=rede")
    if COD == "5930":
      COD = "5930-CAIXA FIC GIRO IMEDIATO RF REF DI LP    "
    if COD == "0088":
      COD = "0088-CAIXA FACIL RENDA FIXA SIMPLES          "
    if COD == "5901":
      COD = "5901-CAIXA FIC GIRO EMPRESAS RF REF DI LP    "
    if COD == "5948":
      COD = "5948-CAIXA FIC GIRO MPE RF REF DI LP         "
    TipoFundo = navegador.find_element(By.NAME, "SelProduto")
    drop = Select(TipoFundo)
    drop.select_by_visible_text(COD)
    navegador.find_element(By.XPATH, "/html/body/table[2]/tbody/tr/td[2]/form/table[3]/tbody/tr/td/div/input").click()
    #clicar em confirmar
    #valor da aplicação
    navegador.find_element(By.NAME, "valorAplicacao").send_keys(ValorApl)
    navegador.find_element(By.XPATH, "/html/body/table[2]/tbody/tr/td[2]/form/table[3]/tbody/tr/td/div/input").click()
  except: #se não conseguir selecionar abrir janela
    Contador += 1
    df.append([Planilhacaixa.loc[i,'Nome'],Planilhacaixa.loc[i,'CPF/CNPJ'], Conta, COD, ValorApl])
    continue

#apos finalizar tudo
resumo = pd.DataFrame(df, columns=["Nomes","CPF/CNPJ", "Conta", "Fundo", "Valor Apl"])
resumo.to_excel("resumo.xlsx", index=False)
if Contador == 0:
  sg.theme('Dark Blue 3') 
  layout = [
  [sg.Text(f'Concluido com Sucesso')],
  [sg.Button('OK')]
  ]
  window = sg.Window('Automação para o doidão', layout)
  event = window.read()
else:
  sg.theme('Dark Blue 3') 
  layout = [
  [sg.Text(f'Aconteceram {Contador} erros durante a aplicação.\nAbra o resume.xlsx junto ao executável para saber mais informações')],
  [sg.Button('OK')]
  ]
  window = sg.Window('Automação para o doidão', layout)
  event = window.read()  
window.close()
navegador.close()

