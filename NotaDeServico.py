import os
import time
import re
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from openpyxl import load_workbook
#Dados para acesso-----
from dados import *
#----------------------

os.environ['TF_CPP_MIN_LOG_LEVEL'] = '2'

def menu():
    print("Digite um número de 1 a 8 para qual contrato inicilizar:")
    for chave, valor in contratos.items():
        print(f"\n|{chave}| = {valor}")
def escolhaContrato(opcao):
    match int(opcao):
        case 1:
            #rio piracicaba
            print("\nIniciando Contrato 30000916")
            diretorio = diretorio0916
            emailGestor = emailGestor916
        case 2:
            #sao goncalo rio abaixo
            print("\nIniciando Contrato 30001168")
            diretorio = diretorio1168
            emailGestor = emailGestor1168
        case 3:
            print("\nIniciando Contrato 30001169")
            diretorio = diretorio1169
            emailGestor = emailGestor1169
        case 4:
            print("\nIniciando Contrato 30001170")
            diretorio = diretorio1170
            emailGestor = emailGestor1170
        case 5:
            print("\nIniciando Contrato 30001203")
            diretorio = diretorio1203
            emailGestor = emailGestor1203
        case 6:
            print("\nIniciando Contrato 30001315")
            diretorio = diretorio1315
            emailGestor = emailGestor1315
        case 7:
            print("\nIniciando Contrato 30001388")
            diretorio = diretorio1388
            emailGestor = emailGestor1388
        case 8:
            print("\nIniciando Contrato 30001395")
            diretorio = diretorio1395
            emailGestor = emailGestor1395
        case _:
            print("\nOpção inválida.")
            quit
def Login():
    link = "https://vale.virtual360.io/"
    nav.get(link)
    nav.maximize_window()
    btnStart = nav.find_element(By.ID, "login_portal")
    btnStart.click()
    inputEmail = nav.find_element(By.ID, "user_login")
    inputEmail.send_keys(email)
    inputToken = nav.find_element(By.ID, "user_password")
    inputToken.send_keys(token)
    btnLogin = nav.find_element(By.NAME, "commit")
    btnLogin.click()
def JoinNotaFiscalDeServicos():
    nav.find_element(By.TAG_NAME, "body").send_keys(Keys.CONTROL + 't')
    nav.get("https://vale.virtual360.io/nf/tax_documents/service_invoice/new")
def InsertArchives(xmlCaminho, pdfCaminho):
    xmlUp = WebDriverWait(nav, 10).until(
        EC.presence_of_element_located((By.ID, "tax_document_document_xml"))
    )
    xmlUp.send_keys(xmlCaminho)
    time.sleep(7)
    pdfUp = WebDriverWait(nav, 10).until(
        EC.presence_of_element_located((By.ID, "tax_document_document_pdf"))
    )
    pdfUp.send_keys(pdfCaminho)
    time.sleep(10)
    inputGestor = WebDriverWait(nav, 10).until(
        EC.presence_of_element_located((By.ID, "tax_document_requester_area"))
    )
    inputGestor.send_keys(emailGestor)
    time.sleep(1)
def SubmitNota():
    btnIngressarNota = nav.find_element(By.NAME, "status_id")
    btnIngressarNota.click()
    time.sleep(3)
    try:
        WebDriverWait(nav, 10).until(EC.alert_is_present())
        alerta = nav.switch_to.alert
        alerta.accept()
    except Exception as e:
        print(f"Erro ao lidar com o alerta: {e}")
        quit
    time.sleep(1)
def BuscarArquivos(archive_num):
    arquivoEncontrado = []
    if not os.path.exists(diretorio):
        raise FileNotFoundError(f"Arquivo não encontrado")
    for arquivo in os.listdir(diretorio):
        if archive_num in arquivo:
            arquivoEncontrado.append(arquivo)
    if len(arquivoEncontrado) == 2:
        for arquivo in arquivoEncontrado:
            caminho_completo = os.path.join(diretorio, arquivo)
            if arquivo.endswith('xml'):
                xmlCaminho = caminho_completo
            elif arquivo.endswith('pdf'):
                pdfCaminho = caminho_completo
        return xmlCaminho, pdfCaminho
    else:
        raise FileNotFoundError("Não há arquivos suficientes com esse nome no diretório.")
def DadosExcel():
    IndexProtocolo = nav.find_element(By.CLASS_NAME, 'v-h4')
    ProtocoloCompleto = IndexProtocolo.text
    match = re.search(r'#(\d+)', ProtocoloCompleto)
    if match:
        numeroP = match.group(1)
        NumeroProtocolo = numeroP
    else:
        print("Nenhum número encontrado no texto.")
    NumeroNota = archive_num
    print(f'numero protocolo: {NumeroProtocolo}')
    print(f'numero Nota: {NumeroNota}')
    time.sleep(2)
    try:
        wb = load_workbook('relatorio.xlsx')
        ws = wb.active
    except Exception as e:
        print(f'Ocorreu um erro ao inicializar a planilha {e}') 
    while True:
        print('\nDigite a celula para registrar o protocolo (Exemplo: A1')
        celulaP = input()
        try:
            if ws[celulaP].value is None:
                ws[celulaP].value = NumeroProtocolo
                print('Protocolo Computado')
                break
            else:
                print('Esta celula já contém dados. Tente novamente.')
        except KeyError:
            print("Célula inválida. Por favor, digite uma célula válida.")
    while True:
        print('\nDigite a celula para registrar a Nota (Exemplo: B1')
        celulaN = input()
        try:
            if ws[celulaN].value is None:
                ws[celulaN].value = NumeroNota
                print('Nota Computada')
                break
            else:
                print('Esta celula já contém dados. Tente novamente.')
        except KeyError:
            print("Célula inválida. Por favor, digite uma célula válida.")
    wb.save('relatorio.xlsx')
if __name__ == "__main__":
    try:
        menu()
        contratoEscolhido = input("\n")
        escolhaContrato(contratoEscolhido)
        nav = webdriver.Chrome()
        Login()
        while True:
            print('\nNúmero do Arquivo: ')
            archive_num = input().strip()
            # archive_num += 1
            try:
                xmlCaminho, pdfCaminho = BuscarArquivos(archive_num)
                JoinNotaFiscalDeServicos()
                InsertArchives(xmlCaminho, pdfCaminho)
                SubmitNota()
                print("\nFormulário enviado com sucesso!")
                DadosExcel()
                print('\nDados Computados no Excel')
                time.sleep(2)
            except FileNotFoundError as e:
                print(f"Erro: {e}")
                # print(f"Arquivo {archive_num} não encontrado: {e}")
                # archive_num = int(input("Digite o próximo número do arquivo: "))
                # continue
            print("\nDeseja processar outro arquivo? (s/n): ")
            continuar = input().strip().lower()
            if continuar != 's':
                print("Encerrando o programa...")
                break
    except Exception as e:
        print(f'ERRO {e}')
    finally:
        nav.quit()