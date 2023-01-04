from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from anticaptchaofficial.recaptchav2proxyless import *
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import pyscreenshot
import xlsxwriter
import time

# Pasta onde os arquivos serão encontrados

diretorio_pdf = r'C:\Users\SEU_USUARIO\Documents\verificador-pagamento-custas-jud\banco-do-brasil\pdf'
diretorio_codigos = r'C:\Users\SEU_USUARIO\Documents\verificador-pagamento-custas-jud\banco-do-brasil\ids\ids.txt'
diretorio_excel = r'C:\Users\SEU_USUARIO\Documents\verificador-pagamento-custas-jud\banco-do-brasil\relatorio'

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
link = "https://www63.bb.com.br/portalbb/djo/id/comprovante/consultaDepositoJudicial,802,4647,4650,0,1.bbx"
driver.get(link)


# Essa parte é responsável pela quebra de captcha

def chave_captcha():
    driver.find_element(By.CLASS_NAME, 'g-recaptcha').get_attribute('data-sitekey')


def solver():
    solver = recaptchaV2Proxyless()
    solver.set_verbose(1)
    solver.set_key('chave_api_anticaptcha')
    solver.set_website_url(link)
    solver.set_website_key(chave_captcha)


def resposta():
    chave_captcha()
    solver = recaptchaV2Proxyless()
    resposta = solver.solve_and_return_solution()
    driver.execute_script(f"document.getElementById('g-recaptcha-response').innerHTML = '{resposta}'")


# Aqui é onde começa a consulta dos ids, bem como a quebra de captcha para cada id consultado em loop
ct_linha = 0

with open(diretorio_codigos) as arquivo:
    linhas = arquivo.readlines()
for linha in linhas:
    codigo = linha.strip()
    campo = driver.find_element(By.ID, 'formulario:numPreDeposito')
    campo.clear()
    # id no campo de pesquisa
    campo.send_keys(codigo)
    # quebrando captcha
    chave_captcha()
    solver()
    resposta()
    driver.find_element(By.ID, 'formulario:btnContinuar').click()
    time.sleep(2)
    ct_linha += 1

    # Verificar se existe mensagem de erro do comprovante
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[1]/div/div/ul/li')))
        element = driver.find_element(By.ID, 'formulario:numPreDeposito')
        element.clear()
        print(f'O {codigo} não tem pagamento')

    except:
        # Quando tem o comprovante
        # planilha
        planilha.write(ct_linha, 0, codigo)
        planilha.write(ct_linha, 1, 'Sim')
        time.sleep(2)
        # clicando na bolinha
        driver.find_element(By.CSS_SELECTOR, "input[type='radio'][value='0']").click()
        # botão visualizar
        driver.find_element(By.ID, 'formulario:btnVisualizar').click()
        # zoom
        driver.execute_script("document.body.style.zoom='77%'")
        # print (lado direito = 210, cima = 300 , lado esquerdo = 1500 e cima = 1035)
        image = pyscreenshot.grab(bbox=(210, 300, 1500, 1035))
        image.save((f'{diretorio_pdf}\\{codigo}.png'))
        time.sleep(3)
        driver.execute_script("document.body.style.zoom='100%'")
        # botão de retornar
        driver.find_element(By.ID, 'formulario:botaoRetornar').click()
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div/form/div[3]/div/input[2]').click()
        print(f'Pagamento efetuado no {codigo}, comprovante salvo')

    else:
        # Quando não tem o comprovante
        planilha.write(ct_linha, 0, codigo)
        planilha.write(ct_linha, 1, 'Não')
        continue
# Salvar excel e fechar chrome
excel.close()
driver.close()
