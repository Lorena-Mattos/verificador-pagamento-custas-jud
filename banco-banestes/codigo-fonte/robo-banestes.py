from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import pyscreenshot
import xlsxwriter
import time

# Pasta onde os arquivos serão encontrados

diretorio_pdf = r'C:\Users\lorena.machado\Documents\robo-banestes\pdf'
diretorio_codigos = r'C:\Users\lorena.machado\Documents\robo-banestes\ids\ids.txt'
diretorio_excel = r'C:\Users\lorena.machado\Documents\robo-banestes\relatorio'

# Criar excel para registro de downloads
data = (f'{datetime.today():%d-%m-%Y}')
hora = (f'{datetime.now():%H-%M}')
nome_arquivo = (f'{data}_{hora}.xlsx')
excel = xlsxwriter.Workbook(f'{diretorio_excel}\\Relatorio_{nome_arquivo}')
planilha = excel.add_worksheet()
planilha.write(0, 0, 'Código')
planilha.write(0, 1, 'Codigo encontrado?')

# Abrir o Chrome no endereço dos comprovantes
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()
link = "http://depositojudicial.banestes.com.br/DepositoJudicial/impressaoGuia/listImpressaoTEDInput.jsf"
driver.get(link)
time.sleep(2)

# Aqui é onde começa a consulta dos ids, bem como a quebra de captcha para cada id consultado em loop
ct_linha = 0

with open(diretorio_codigos) as arquivo:
    linhas = arquivo.readlines()
for linha in linhas:
    codigo = linha.strip()
    campo = driver.find_element(By.ID, 'frm:nrIdDeposito')
    campo.clear()
    # id no campo de pesquisa
    campo.send_keys(codigo)
    driver.find_element(By.ID, 'frm:cbConfirmar').click()
    time.sleep(3)

    ct_linha += 1

    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'errorMessages')))
        element = driver.find_element(By.ID, 'frm:nrIdDeposito')
        element.clear()
        print(f'O {codigo} não tem pagamento')
        time.sleep(3)

    except:
        driver.switch_to.window(driver.window_handles[1])
        # prints windows id
        print(driver.window_handles)
        print('Segunda janela')
        print(f'O {codigo} tem pagamento')
        # switch window in 7 seconds
        time.sleep(1)
        # print (lado direito = 0 (sempre diminuir), baixo = 102 (sempre diminuir), lado esquerdo = 1850 e cima = 1040
        # (lado esquerdo e cima) sempre aumentar)
        image = pyscreenshot.grab(bbox=(0, 102, 1850, 1040))
        image.save((f'{diretorio_pdf}\\{codigo}.png'))
        time.sleep(3)
        # planilha
        planilha.write(ct_linha, 0, codigo)
        planilha.write(ct_linha, 1, 'Sim')
        # switch to new window
        driver.switch_to.window(driver.window_handles[0])
        # prints windows id
        print(driver.window_handles)
        print('Primeira janela')
        time.sleep(2)

    else:
        # Quando não tem o comprovante
        planilha.write(ct_linha, 0, codigo)
        planilha.write(ct_linha, 1, 'Não')
        print('Nada foi encontrado')
        continue

# Salvar excel e fechar chrome
excel.close()
driver.close()
