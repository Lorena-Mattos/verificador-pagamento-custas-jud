from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from datetime import date
import xlsxwriter
import pdfkit
import time


#Esse robô só funciona com versão do Office 2010 ou superior.

#Pasta onde os arquivos serão encontrados
diretorio_pdf = r'C:\Users\SEU_USUARIO\Documents\verificador-pagamento-custas-jud\banco-brb\pdf'
diretorio_codigos = r'C:\Users\SEU_USUARIO\Documents\verificador-pagamento-custas-jud\banco-brb\ids\ids.txt'
diretorio_excel = r'C:\Users\SEU_USUARIO\Documents\verificador-pagamento-custas-jud\banco-brb\relatorio'

# Criar excel para registro de downloads
data = (f'{datetime.today():%d-%m-%Y}')
hora = (f'{datetime.now():%H-%M}')
nome_arquivo = f'{data}_{hora}.xlsx' 
excel = xlsxwriter.Workbook(f'{diretorio_excel}\\Relatorio_{nome_arquivo}')
planilha = excel.add_worksheet()
planilha.write(0, 0, 'Código')
planilha.write(0, 1, 'Codigo encontrado?')

#Abrir o Chrome no endereço dos comprovantes
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()
driver.get("https://sdj.brb.com.br/depositos-judiciais/comprovantes")

#Loop pelos códigos do arquivo txt e busca no site
ct_linha = 0
with open(diretorio_codigos) as arquivo:
    linhas = arquivo.readlines()
    for linha in linhas:     
        codigo = linha.strip()        
        campo = driver.find_element(By.XPATH, '//*[@id="idGuia"]')
        campo.clear()
        campo.send_keys(codigo)
        
        # Clicar no Confirmar
        bt_confirmar = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/form/panel/div/div/div[2]/div[2]/div/button')))
        bt_confirmar.click()

        #Verificar se existe mensagem de erro do comprovante
        ct_linha+=1
                      
        try:
            element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/form/panel/div/div/div[2]/msg/div/div/div')))
            time.sleep(3)
        except:
            #Quando tem o comprovante  
            planilha.write(ct_linha, 0, codigo)
            planilha.write(ct_linha, 1, 'Sim')            
            time.sleep(2)            
        else:
            #Quando não tem o comprovante
            planilha.write(ct_linha, 0, codigo)
            planilha.write(ct_linha, 1, 'Não')
            
            time.sleep(2)            
            continue
            
        #Continua o processo para quando tem comprovante
        
        datahora = driver.find_element("xpath", '/html/body/div[2]/div/div/form/panel/div/div/div[2]/div[1]/div[1]/div[2]/span').text
        identificador = driver.find_element("xpath", '/html/body/div[2]/div/div/form/panel/div/div/div[2]/div[1]/div[1]/div[1]/span').text
        data = driver.find_element("xpath", '/html/body/div[2]/div/div/form/panel/div/div/div[2]/div[1]/div[1]/div[2]/span').text
        modalidade = driver.find_element("xpath",'/html/body/div[2]/div/div/form/panel/div/div/div[2]/div[1]/div[3]/div/span').text
        processo = driver.find_element("xpath", '/html/body/div[2]/div/div/form/panel/div/div/div[2]/div[1]/div[4]/div/span').text
        vara = driver.find_element("xpath", '/html/body/div[2]/div/div/form/panel/div/div/div[2]/div[1]/div[5]/div/span').text
        autor = driver.find_element("xpath", '/html/body/div[2]/div/div/form/panel/div/div/div[2]/div[1]/div[6]/div/span').text
        reu = driver.find_element("xpath", '/html/body/div[2]/div/div/form/panel/div/div/div[2]/div[1]/div[7]/div/span').text
        valor_boleto = driver.find_element("xpath", '/html/body/div[2]/div/div/form/panel/div/div/div[2]/div[1]/div[9]/div/span').text
        valor_pago = driver.find_element("xpath", '/html/body/div[2]/div/div/form/panel/div/div/div[2]/div[1]/div[10]/div/span').text
        conta = driver.find_element("xpath", '/html/body/div[2]/div/div/form/panel/div/div/div[2]/div[1]/div[11]/div/span').text
        data_pagamento = driver.find_element("xpath", '/html/body/div[2]/div/div/form/panel/div/div/div[2]/div[1]/div[12]/div/span').text
        data_credito = driver.find_element("xpath", '/html/body/div[2]/div/div/form/panel/div/div/div[2]/div[1]/div[13]/div/span').text

        nome_pdf = f'{diretorio_pdf}\\{codigo}.pdf'

        html = """\
        <html>
          <head>
            <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
          </head>
          <body style=font-family:Lato,sans-serif">
            <span style="font-family:Lato,sans-serif; font-size: 12px">
              <strong>{datahora}</strong>
            </span>
            </h4>
            <h1>
              <span style="font-size:25px; font-family:Lato,sans-serif; font-weight:300">Comprovante de Dep&oacute;sito Judicial</span>
            </h1>
            <p>&nbsp;</p>
            <h2>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: bold"> Identificador do dep&oacute;sito: </span>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: 300"> {identificador} </span>
            </h2>
            <h2>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: bold">Data:</span>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: 300"> {data} </span>
            </h2>
            <h2>&nbsp;</h2>
            <h2>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: bold"> Modalidade do dep&oacute;sito: </span>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: 300"> {modalidade} </span>
            </h2>
            <h2>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: bold">Processo:</span>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: 300"> {processo} </span>
            </h2>
            <h2>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: bold">Vara:</span>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: 300"> {vara} </span>
            </h2>
            <h2>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: bold">Autor:</span>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: 300"> {autor} </span>
            </h2>
            <h2>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: bold">R&eacute;u:</span>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: 300"> {reu} </span>
            </h2>
            <p>&nbsp;</p>
            <h1 style="text-align:center">
              <strong>
                <span style="font-size:25px; font-family:Lato,sans-serif">DEP&Oacute;SITO EFETIVADO COM SUCESSO &nbsp;</span>
              </strong>
            </h1>
            <h2>&nbsp;</h2>
            <h2>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: bold">Valor do Boleto:</span>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: 300"> {valor_boleto} </span>
            </h2>
            <h2>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: bold">Valor Pago:</span>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: 300"> {valor_pago} </span>
            </h2>
            <h2>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: bold">Conta Judicial:</span>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: 300"> {conta} </span>
            </h2>
            <h2>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: bold">Data do Pagamento:</span>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: 300"> {data_pagamento} </span>
            </h2>
            <h2>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: bold">Data do Cr&eacute;dito:&nbsp;</span>
              <span style="font-size:18px; font-family:Lato,sans-serif; font-weight: 300"> {data_credito} </span>
            </h2>
            <p>&nbsp;</p>
            <button style="padding:6px 12px; font-size:14px; display:inline-block; vertical-align:middle; cursor:pointer;
            background-image:none; text-decoration:none; background-color: white; border: 1px; border-style: solid; border-color: black; ">
              <span style="font-size:18px; font-weight:300; font-family:Lato,sans-serif">Nova consulta</span>
            </button>
          </body>
        </html>    
        """.format(identificador=identificador, datahora=datahora, data=data, modalidade=modalidade, processo=processo,
                  vara=vara, autor=autor, reu=reu, valor_boleto=valor_boleto, valor_pago=valor_pago, conta=conta, 
                  data_pagamento=data_pagamento, data_credito=data_credito)


        wkhtml_path = pdfkit.configuration(wkhtmltopdf = "C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe")  #by using configuration you can add path value.

        pdfkit.from_string(html, nome_pdf, configuration = wkhtml_path)
        
        time.sleep(3)
        
        # Clicar no Nova Consulta
        bt_nova_consulta_tem = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/form/panel/div/div/div[2]/div[2]/div/button[2]')))
        bt_nova_consulta_tem.click()        
        #Fim Loop
        
#Salvar excel e fechar chrome          
excel.close()
driver.close()
