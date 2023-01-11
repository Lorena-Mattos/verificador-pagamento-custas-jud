from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import xlsxwriter
import time

# Pasta onde os arquivos serão encontrados

diretorio_pdf = r'C:\Users\lorena.machado\Documents\verificador-pagamento-custas-jud\banco-caixa-economica\pdf'
diretorio_codigos = r'C:\Users\lorena.machado\Documents\verificador-pagamento-custas-jud\banco-caixa-economica\ids\ids.txt'
diretorio_excel = r'C:\Users\lorena.machado\Documents\verificador-pagamento-custas-jud\banco-caixa-economica\relatorio'

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
link = "https://depositojudicial.caixa.gov.br/sigsj_internet/impressao-de-documentos/guias-depositos/#"
driver.get(link)
time.sleep(2)
# Aqui é onde começa a consulta dos ids, bem como a quebra de captcha para cada id consultado em loop
ct_linha = 0

with open(diretorio_codigos) as arquivo:
	linhas = arquivo.readlines()
for linha in linhas:
	codigo = linha.strip()
	campo = driver.find_element(By.ID, 'j_id5:filtroView:j_id6:idDeposito')
	campo.clear()
	# id no campo de pesquisa
	campo.send_keys(codigo)
	driver.find_element(By.ID, 'j_id5:filtroView:j_id6:btnConsultar').click()
	time.sleep(3)

	ct_linha += 1

	try:
		element = driver.find_element(By.ID, 'j_id5:filtroView:j_id6:_fecharButton').click,
		time.sleep(1)
		element = WebDriverWait(driver, 10).until(
			EC.presence_of_element_located((By.XPATH, 
											'/html/body/div[2]/div[2]/div/div[2]/table/tbody/tr[2]/td/div/div[3]/table[1]/tbody/tr/td[2]/table[11]/tbody/tr/td[2]')))
		element = driver.find_element(By.CLASS_NAME, 'icon icon-closethick').click
		element = driver.find_element(By.ID, 'j_id5:filtroView:j_id6:idDeposito')
		element.clear()
		print(f'O {codigo} não tem pagamento efetuado')
		time.sleep(3)

	except:
		driver.find_element(By.ID, 'j_id5:filtroView:j_id6:_fecharButton').click,
		time.sleep(1)
		
		driver.find_element(By.XPATH, 
							'/html/body/div[2]/div[2]/div/div[2]/table/tbody/tr[2]/td/div/div[3]/table[1]/tbody/tr/td[2]/table[11]/tbody/tr/td[2]/text()')
		print(f'O {codigo} tem pagamento efetuado')
		time.sleep(2)
		
		driver.find_element(By.XPATH, 
							'/html/body/div[2]/div[2]/div/div[2]/table/tbody/tr[2]/td/div/div[1]/table/tbody/tr/td[3]/span/a[1]/img').click
		driver.find_element(By.CLASS_NAME, 'icon icon-closethick').click
		# planilha
		planilha.write(ct_linha, 0, codigo)
		planilha.write(ct_linha, 1, 'Sim')

	else:
		# Quando não tem o comprovante
		planilha.write(ct_linha, 0, codigo)
		planilha.write(ct_linha, 1, 'Não')
		print('Nada foi encontrado')
		continue

# Salvar excel e fechar chrome
excel.close()
driver.close()