
"""
pip install selenium
pip install webdriver-manager
pip install reportlab
pip install pandas
pip install openpyxl
"""

# Importa as bibliotecas
import time
import os
import pandas as pd
import logging
import pyautogui
import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph
from reportlab.lib import colors


# Funções
def apaga_conversa():
    try:
        tres_pontos = driver.find_element(By.CLASS_NAME, '_3vsRF')
        tres_pontos.find_element(By.TAG_NAME, 'span').click()
        time.sleep(2)
        menus = driver.find_elements(By.CLASS_NAME, 'iWqod')
    except:
        raise ValueError("error")

    for menu in menus:
        try:
            menu.text
        except:
            continue

        if menu.text == 'Apagar conversa':
            try:
                menu.click()
                time.sleep(2)
            except:
                break

            # Clica na mensagem que pode aparecer antes
            try:
                driver.find_element(By.CLASS_NAME, 'oixtjehm').click()
                time.sleep(2)
            except:
                pass
            
            # Clica para apagar a conversa
            try:
                driver.find_element(By.CLASS_NAME, 'oixtjehm').click()
                time.sleep(3)
                break
            except:
                break

def enviar_pdf():

    # captura todas as tabelas na página
    try:
        tabelas = driver2.find_elements(By.TAG_NAME, 'table')
        time.sleep(1)
    except:
        logging.info('Erro ao capturar as tabelas')
        driver2.quit()
        raise ValueError("error")


    # Cria o pdf
    doc = SimpleDocTemplate("Consulta.pdf", pagesize=A4)
    elements = []

    # Título da Página
    try:
        dados = [['Pesquisa Cadastral Simplificada']]
        t = Table(dados)
        t.setStyle(TableStyle([('FONTSIZE', (0, 0), (0, 0), 20)]))
        elements.append(t)
        elements.append(Spacer(width=0, height=50))
    except:
        logging.info('Erro ao gerar o título do pdf')
        driver2.quit()
        raise ValueError("error")


    # Dados iniciais
    try:
        rows = tabelas[0].find_elements(By.TAG_NAME, "tr")
        cols = rows[0].find_elements(By.TAG_NAME, "td")
        nome_cpf = cols[1].text
    except:
        logging.info('Erro ao capturar as linhas e colunas da tabela')
        driver2.quit()
        raise ValueError("error")

    try:
        cols = rows[2].find_elements(By.TAG_NAME, "td")
        cpf = cols[1].text
    except:
        logging.info('Erro ao capturar o cpf da tabela')
        driver2.quit()
        raise ValueError("error")

    try:
        cols = rows[6].find_elements(By.TAG_NAME, "td")
        data_hora = cols[1].text
    except:
        logging.info('Erro ao capturar a data e hora da tabela')
        driver2.quit()
        raise ValueError("error")

    try:
        dados = [['Nome do Cliente:', nome_cpf], ['CPF:', cpf], ['Data / Hora:', data_hora]]
        t = Table(dados)
        t.setStyle(TableStyle([('FONTSIZE', (0, 0), (-1, -1), 12)]))
        elements.append(t)
        elements.append(Spacer(width=0, height=40))
    except:
        logging.info('Erro ao gerar o cabeçalho do pdf')
        driver2.quit()
        raise ValueError("error")

    # Itera as outras tabelas da segunda até a penúltima
    for i in range(1,len(tabelas) - 1):
        lista_tabela = []
        j = 0
        
        rows = tabelas[i].find_elements(By.TAG_NAME, "tr")
        for row in rows:

            # Transforma a primeira linha da tabela no título
            if j == 0:
                j += 1
                lista_auxiliar = []
                lista_titulo = []
                
                c = row.find_element(By.TAG_NAME, "td")
                try:
                    lista_titulo.append(c.text)
                    lista_auxiliar.append(lista_titulo)
                    t = Table(lista_auxiliar)
                    t.setStyle(TableStyle([('FONTSIZE', (0, 0), (0, 0), 12)]))
                    elements.append(t)
                    elements.append(Spacer(width=0, height=10))
                except:
                    logging.info('Erro ao iterar os títulos das tabelas')
                    pass
                
            # Para o restante da tabela 
            else:
                lista_celula = []    
                cols = row.find_elements(By.TAG_NAME, "td")
                
                for col in cols:
                    try:
                        lista_celula.append(Paragraph(col.text))
                    except:
                        logging.info('Erro ao iterar os dados das tabelas')
                        continue
                
                lista_tabela.append(lista_celula)
            
        try:
            t = Table(lista_tabela)
        
            t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 1), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 1), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, 0), 'LEFT'),
                            ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 10),
                            ('FONTSIZE', (0, 1), (-1, -1), 8),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
            
            elements.append(t)
            elements.append(Spacer(width=0, height=30))
        except:
            logging.info('Erro ao formatar as tabelas')
            continue

    try:
        doc.build(elements)
    except:
        logging.info('Erro ao gerar o pdf')
        driver2.quit()
        raise ValueError("error")

    # Fecha a segunda janela
    driver2.quit()
    

    try:
        diretorio_atual = os.getcwd()
        arquivo = diretorio_atual + '\Consulta.pdf'
    except:
        logging.info('Erro ao pegar o caminho do arquivo')
        raise ValueError("error")


    # Clica para abrir o arquivo
    try:
        driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div/div/div/div').click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div/div/span/div/ul/div/div[4]/li/div').click()
        time.sleep(4)
    except:
        try:
            driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div').click()
            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[5]/button').click()
            time.sleep(4)
        except:
            logging.info('Erro ao clicar em anexo')
            raise ValueError("error")
    
    # Envia o arquivo
    try:
        pyautogui.typewrite(arquivo)
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(4)
        driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div[3]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div').click()
        time.sleep(10)
    except:
        logging.info('Erro ao enviar o arquivo')
        raise ValueError("error")

    # Remove o arquivo
    try:
        os.remove(arquivo)
    except:
        logging.info('Erro ao apagar o arquivo')
        pass
    


# Cria o log
logging.basicConfig(level=logging.INFO, filename='app.log', format='%(asctime)s - %(levelname)s - %(message)s')
logging.info('Inicio do programa')


# Abre a planilha com os dados de login
try:
    df = pd.read_excel('login.xlsx') 
except:
    logging.info('Erro ao abrir o arquivo login.xlsx')


# Início do programa
whatsapp = 'https://web.whatsapp.com/'
caminho_pasta_atual = os.getcwd()

options = Options()
options.add_argument("--profile-directory=Default")
options.add_argument(f"--user-data-dir={caminho_pasta_atual}/cookies")


#driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()

driver.get(whatsapp)
time.sleep(5)


while True:

    try:
        app = driver.find_element(By.ID, 'app')
        side = app.find_element(By.ID, 'side')
        time.sleep(5)
    except:
        time.sleep(5)
    
    # Pega as últimas conversas
    try:
        conversas = side.find_elements(By.CLASS_NAME, 'aprpv14t')  
    except:
        logging.info('Erro ao pegar as conversas')
        time.sleep(10)
        continue

    i = 0

    # Itera as conversas encontradas
    for conversa in conversas:

        try:  
            conversa.text 
        except:
            logging.info('Erro na conversa')
            continue 


        # Verifica se as conversas tem menos de 24h
        if len(conversa.text) == 5 and conversa.text[0].isnumeric():

            # Clica no conversa
            try:
                conversa.click()
                time.sleep(2)
            except:
                logging.info('Erro ao clicar na conversa')
                time.sleep(10)
                continue

     
            # Captura o usuario da conversa   
            try:
                element = WebDriverWait(app, 10).until(EC.presence_of_element_located((By.ID, "main")))
                main = app.find_element(By.ID, 'main')
                nome = main.find_element(By.CSS_SELECTOR, '.ggj6brxn.gfz4du6o.r7fjleex.g0rxnol2.lhj4utae.le5p0ye3.l7jjieqr._11JPr').text
   
            except:
                logging.info('Erro ao pegar nome do usuário')
                time.sleep(10)
                continue
            

            # Busca as mensagens
            try:
                mensagens_class = main.find_elements(By.CSS_SELECTOR, '.cm280p3y.to2l77zo.n1yiu2zv.c6f98ldp.ooty25bp.oq31bsqd')
            except:
                logging.info('Erro ao pegar as divs')
                time.sleep(10)
                continue

            for mensagem_class in mensagens_class:
                try:
                    mensagem_div = mensagem_class.find_element(By.CSS_SELECTOR, '._11JPr.selectable-text.copyable-text')
                    mensagem = mensagem_div.find_element(By.TAG_NAME, 'span')
                    numero = mensagem.text.replace('.', '').replace('-', '').replace('/', '').replace(' ', '') 
                except:
                    continue
        
        
                # Verifica se o cpf foi digitado errado
                if (len(numero) == 10 and numero.isnumeric()) or (len(numero) == 12 and numero.isnumeric()):
                    
                    # Verifica se o nome do contato está salvo
                    try:
                        if nome[-4:-1].isnumeric():
                            resultado = 'Identifiquei que você não tem autorização para acesso a consulta. Solicite acesso ao administrador!'
                        else:
                            resultado = 'Parece que você digitou o CPF errado!'
                    except:
                        break

                    try:
                        driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(resultado)
                        time.sleep(1)
                        driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button').click()
                        time.sleep(4)
                        logging.info('Mensagem enviada')
                        apaga_conversa()
                    except:
                        logging.info('Erro ao enviar a mensagem')
                        break

                # Verifica se o número é um CPF
                elif len(numero) == 11 and numero.isnumeric():


                    # Verifica se o nome do contato está salvo
                    if nome[-4:-1].isnumeric():
                        resultado = 'Identifiquei que você não tem autorização para acesso a consulta. Solicite acesso ao administrador!'

                        try:
                            driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(resultado)
                            time.sleep(1)
                            driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button').click()
                            time.sleep(4)
                            apaga_conversa()
                        except:
                            logging.info('Erro ao enviar mensagem de contato não salvo')
                            continue

                    else:
                        caixa = 'https://caixaaqui.caixa.gov.br/caixaaqui/CaixaAquiController/index'

                        service2 = Service(ChromeDriverManager().install())
                        driver2 = webdriver.Chrome(service=service2)
                        driver2.maximize_window()
                        driver2.get(caixa)
                        time.sleep(10)

                        # Faz login na Caixa
                        try:
                            driver2.find_element(By.XPATH, '//*[@id="convenio"]').send_keys(df['Valores'][0])
                            time.sleep(1)
                            driver2.find_element(By.XPATH, '//*[@id="login"]').send_keys(df['Valores'][1])
                            time.sleep(1)
                            driver2.find_element(By.XPATH, '//*[@id="password"]').send_keys(df['Valores'][2])
                            time.sleep(1)
                            driver2.find_element(By.XPATH, '//*[@id="btLogin"]/input').click()
                            time.sleep(5) 

                        except:
                            time.sleep(2)
                            pass

                        # Navega no site
                        try:
                            driver2.find_element(By.XPATH, '//*[@id="menu-principal"]/tbody/tr[1]/td/a').click()
                            time.sleep(2)

                            driver2.find_element(By.XPATH, '//*[@id="menu-principal"]/tbody/tr[1]/td/a').click()
                            time.sleep(2)

                            driver2.find_element(By.XPATH, '//*[@id="menu-principal"]/tbody/tr[2]/td/a').click()
                            time.sleep(2)
                            
                        except:
                            logging.info('Erro ao navegar no site')
                            driver2.quit()
                            time.sleep(1)
                            break

                        # Escreve o CPF e busca no site
                        try:      
                            driver2.find_element(By.XPATH, '//*[@id="cpf"]').send_keys(numero)
                            time.sleep(1)

                            botao = driver2.find_elements(By.CLASS_NAME, 'btn-azul')
                            botao[1].click()
                            time.sleep(10)
                            
                        except:
                            logging.info('Erro ao buscar cpf no site')
                            driver2.quit()
                            time.sleep(1)
                            break

                        # Variáveis de controle
                        alert_texto = ''
                        pc = ''

                        # Verifica se tem alert
                        try:
                            # Muda o foco para o alerta
                            alert = Alert(driver2)
                            
                            # Obtem o texto do alerta
                            alert_texto = alert.text
                            
                            # Aceite o alerta
                            alert.accept()
                            time.sleep(1)

                        except:
                            # Verifica se gerou um link
                            try:
                                pesquisa_cadastral = driver2.find_element(By.ID, 'pesquisa-cadastral')
                                link_pesquisa_cadastral = pesquisa_cadastral.find_element(By.TAG_NAME, 'a')
                                pc = link_pesquisa_cadastral.text
                                
                            except:
                                try:
                                    cpf_cliente = driver2.find_element(By.ID, 'tdcpf').text
                                    nome_cliente = driver2.find_element(By.ID, 'nome').text
                                    regularidade = driver2.find_element(By.ID, 'regularidade').text
                                    pesquisa_cadastral = driver2.find_element(By.ID, 'pesquisa-cadastral')
                                    mensagem_avaliacao = driver2.find_element(By.ID, 'mensagem-avaliacao-risco').text
                                    pc = pesquisa_cadastral.text
                                    
                                except:
                                    logging.info('Erro no servidor da caixa')
                                    pc = 'Erro no servidor da Caixa!'

                        # Verifica o texto do alert
                        if alert_texto != '':

                            if alert_texto == 'CPF do cliente inválido.':
                                resultado = 'CPF inválido!'

                            elif alert_texto == 'For input string: " "':

                                try:
                                    driver2.get('https://caixaaqui.caixa.gov.br/caixaaqui/CaixaAquiController/consulta_cadastral/consulta_cadastral1')
                                    time.sleep(5)
                                    driver2.find_element(By.XPATH, '//*[@id="dataCpf"]').send_keys(numero)
                                    driver2.find_element(By.XPATH, '//*[@id="spanCPF"]/a').click()
                                    time.sleep(2)
                                except:
                                    logging.info('Erro ao pegar dados do alert')
                                    driver2.quit()
                                    time.sleep(1)
                                    break

                                try:
                                    enviar_pdf()
                                    apaga_conversa()

                                except:
                                    logging.info('Erro ao enviar o pdf do alert')
                                    break

                                continue
                                        
                            else:
                                resultado = 'Erro com o CPF'
                                

                        elif pc == 'Nada consta':
                            resultado = f"""Nome do Cliente: {nome_cliente}
                            Situação CPF na Receita: {regularidade}
                            Pesquisa Cadastral: {pc}
                            {mensagem_avaliacao}"""

                        elif pc == 'Constam Ocorrências':
                            link_pesquisa_cadastral.click()
                            time.sleep(5)

                            try:
                                enviar_pdf()
                                apaga_conversa()
                            except:
                                logging.info('Erro ao salvar o dado do alert no banco')
                                break

                            continue

                        else:
                            resultado = 'Erro no servidor da Caixa!'

                        driver2.quit()

                        try:
                            driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(resultado)
                            time.sleep(1)
                            driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button').click()
                            time.sleep(4)
                            apaga_conversa()
                            logging.info('Mensagem enviada')
                        except:
                            logging.info('Erro ao enviar a mensagem')
                            break

    time.sleep(5)
    