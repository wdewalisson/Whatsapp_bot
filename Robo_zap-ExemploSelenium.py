import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import urllib.parse
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

clientes = pd.read_excel("Base_Clientes\Clientes.xlsx")
# usuario = input("Qual seu nome?\n")
# artigo = ""

# user_gender = input("Por favor, insira o seu gênero (M para Masculino, F para Feminino): ")
# if user_gender.upper() == "M":
#     artigo = "o"
    
# elif user_gender.upper() == "F":
#     artigo = "a"
# else:
#     artigo = "o"  

usuario = "Rosivania"
artigo = "a"

navegador = webdriver.Chrome()
navegador.get("https://web.whatsapp.com/")

# Configure o tempo de espera máximo (em segundos)
wait = WebDriverWait(navegador, 10000)

# Aguarde até que o elemento com o ID "side" seja visível
element = wait.until(EC.visibility_of_element_located((By.ID, "side")))

for i, mensagem in enumerate(clientes['PROTOCOLO']):
    protocolo = clientes.loc[i, "PROTOCOLO"]
   
    endereco = clientes.loc[i, "ENDERECO_CLIENTE"]
    
    servico_do_cliente = clientes.loc[i, "SERVICO"]
   
    nome_do_cliente = clientes.loc[i, "NOME_CLIENTE"]
   
    
    
    texto = urllib.parse.quote(f"Olá! Tudo bem? Sou {artigo} {usuario}, da *Algar Telecom*. Estou entrando em contato para realizar o agendamento da *{servico_do_cliente}* do cliente *{nome_do_cliente}*, protocolo {protocolo}. \n\n*Para quando podemos agendar este serviço?* \n\n\nAh, e para garantir que o nosso técnico irá encontrar o endereço, poderia confirmar se está correto? *{endereco}*")

    
    telefone_cliente = clientes.loc[i, "TELEFONE_CLIENTE"]

    link = f"https://web.whatsapp.com/send?phone=55{telefone_cliente}&text={texto}"
    navegador.get(link)
    element = wait.until(EC.visibility_of_element_located((By.ID, "side")))
        
    element2 = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p')))
    element2.send_keys(Keys.ENTER)
    time.sleep(10)
