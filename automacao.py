from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import os
import glob
import pandas as pd 
from time import sleep
import psycopg2 
from sqlalchemy import create_engine
from login import credenciais
from login import banco

servico = Service(ChromeDriverManager().install())
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)

nome_usuario = credenciais.get('NOME_USUARIO')
senha = credenciais.get('SENHA')
host = banco.get('HOST')
database = banco.get('DATABASE')
user = banco.get('USER')
password = banco.get('PASSWORD')

navegador = webdriver.Chrome(service=servico, options=chrome_options)

navegador.get('https://limesurvey.funai.gov.br/index.php?r=admin')

navegador.find_element("xpath", '//*[@id="details-button"]').click()

navegador.find_element("xpath", '//*[@id="proceed-link"]').click()

navegador.find_element("xpath", '//*[@id="user"]').send_keys(nome_usuario)

navegador.find_element("xpath", '//*[@id="password"]').send_keys(senha)

navegador.find_element("xpath", '//*[@id="loginform"]/div[2]/div/p/button').click()

navegador.find_element("xpath", '//*[@id="panel-2"]').click()

navegador.find_element("xpath", '//*[@id="Survey_searched_value"]').send_keys('Projetos de Etnodesenvolvimento - 2022')

navegador.find_element("xpath", '//*[@id="yw0"]/input[2]').click()

navegador.find_element("xpath", '//*[@id="survey-grid"]/table/tbody/tr').click()

navegador.find_element("xpath", '//*[@id="surveybarid"]/div/div[1]/div[3]/button').click()

navegador.find_element("xpath", '//*[@id="surveybarid"]/div/div[1]/div[3]/ul/li[1]/a').click()

navegador.implicitly_wait(10)

navegador.find_element("xpath", '//*[@id="browsermenubarid"]/div/div/div[1]/button').click()

navegador.find_element("xpath", '//*[@id="browsermenubarid"]/div/div/div[1]/ul/li[1]/a').click()

navegador.implicitly_wait(3)

navegador.find_element("xpath", '//*[@id="panel-1"]/div[2]/div/div[1]/div/div[2]').click()

navegador.find_element("xpath", '//*[@id="panel-5"]/div[2]/div[1]/div/label[1]').click()

navegador.find_element("xpath", '//*[@id="export-button"]').click()

sleep(10)

lista_download = glob.glob("C:/Users/davi.costa/Downloads/*")
arquivo_baixado = max(lista_download, key=os.path.getmtime)

print(arquivo_baixado)

resultado = pd.read_excel(arquivo_baixado) 
resultado.dropna(subset=["Data de envio"], inplace=True)

print(len(resultado))

resultado.drop(columns = resultado.iloc[:,2:7], inplace=True) # apaga colunas de um determinado intervalo pelo método iloc
resultado = resultado.iloc[:, :-3]

print(resultado)

resultado.drop(columns = ["Telefone", "Email", "URL de referência","CTL 1 [Outros]", "CTL 7", "CTL 8", "CTL 9", "CTL 10", "CTL 11", "CTL 12", "Outros públicos beneficiados indiretamente", "Apresente a memória de cálculo referente ao subelemento acima"], inplace=True)

resultado.set_index('ID da resposta',inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 1º subelemento de despesa selecionado "])#transforma em número e preenche colunas vazias com zero, podemos também criar uma função para fazer isso com as outras colunas
resultado["Valor solicitado para o 1º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 1º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 1º subelemento de despesa selecionado "].fillna(0, inplace=True)
print(sum(resultado["Valor solicitado para o 1º subelemento de despesa selecionado "]))

pd.to_numeric(resultado["Valor solicitado para o 2º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 2º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 2º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 2º subelemento de despesa selecionado "].fillna(0, inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 3º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 3º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 3º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 3º subelemento de despesa selecionado "].fillna(0, inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 4º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 4º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 4º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 4º subelemento de despesa selecionado "].fillna(0, inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 5º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 5º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 5º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 5º subelemento de despesa selecionado "].fillna(0, inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 6º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 6º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 6º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 6º subelemento de despesa selecionado "].fillna(0, inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 7º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 7º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 7º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 7º subelemento de despesa selecionado "].fillna(0, inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 8º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 8º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 8º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 8º subelemento de despesa selecionado "].fillna(0, inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 9º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 9º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 9º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 9º subelemento de despesa selecionado "].fillna(0, inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 10º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 10º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 10º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 10º subelemento de despesa selecionado "].fillna(0, inplace=True)

resultado.to_excel('Z:/CGETNO/CGETNO/Estágio/Davi/BI/dragolatocoloni4.xlsx', engine='xlsxwriter')
os.remove(arquivo_baixado)

novos_nomes = []
cols = []

for x in range(len(resultado.columns)):
    cols.append(x)

resultado.columns = cols

navegador.close()

conexao = psycopg2.connect(host=host, database=database, user=user, password=password)

engine = create_engine('postgresql://postgres:1234@localhost:5432/PAT 2022')
resultado.to_sql('Projetos 2022', engine) 

cursor = conexao.cursor()
cursor.execute('SELECT "Projetos 2022"."ID da resposta", nomes."terrai_nom" \
from public."Projetos 2022", public."nomes"  \
WHERE "Projetos 2022"."149" = nomes."terrai_cod"  \
OR "Projetos 2022"."150" = nomes."terrai_cod" \
OR "Projetos 2022"."151" = nomes."terrai_cod" \
OR "Projetos 2022"."152" = nomes."terrai_cod" \
OR "Projetos 2022"."153" = nomes."terrai_cod" \
OR "Projetos 2022"."154" = nomes."terrai_cod" \
OR "Projetos 2022"."155" = nomes."terrai_cod" \
OR "Projetos 2022"."156" = nomes."terrai_cod" \
OR "Projetos 2022"."157" = nomes."terrai_cod" \
OR "Projetos 2022"."158" = nomes."terrai_cod" \
OR "Projetos 2022"."159" = nomes."terrai_cod" \
OR "Projetos 2022"."160" = nomes."terrai_cod" \
OR "Projetos 2022"."161" = nomes."terrai_cod" \
OR "Projetos 2022"."162" = nomes."terrai_cod" \
OR "Projetos 2022"."163" = nomes."terrai_cod" \
OR "Projetos 2022"."164" = nomes."terrai_cod" \
OR "Projetos 2022"."165" = nomes."terrai_cod" \
OR "Projetos 2022"."166" = nomes."terrai_cod" \
OR "Projetos 2022"."167" = nomes."terrai_cod" \
OR "Projetos 2022"."168" = nomes."terrai_cod" \
OR "Projetos 2022"."169" = nomes."terrai_cod" \
OR "Projetos 2022"."170" = nomes."terrai_cod" \
OR "Projetos 2022"."171" = nomes."terrai_cod" \
OR "Projetos 2022"."172" = nomes."terrai_cod" \
OR "Projetos 2022"."173" = nomes."terrai_cod" \
ORDER BY "ID da resposta";')

dados = cursor.fetchall()
df_bd = pd.DataFrame(dados, columns=['ID', 'Terras Indígenas'])

df_bd.set_index('ID',inplace=True)

df_bd.to_excel('Z:/CGETNO/CGETNO/Estágio/Davi/terras_por_projeto.xlsx') #nome do arquivo
print('DataFrame convertido em Excel com sucesso.')


print('fim')

