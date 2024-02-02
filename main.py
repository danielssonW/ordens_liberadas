import win32com.client
import re
import os
from win32com.client import Dispatch

outlook_app = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
outlook = win32com.client.Dispatch("Outlook.Application")
inbox = outlook_app.GetDefaultFolder(6)
emails = inbox.Items

def iniciar():
    ordens_emails = pegar_ordens_arquivo()
    verificar_orden_no_job(ordens_emails)

def pegar_ordens_arquivo():
    with open("ordens.txt", "r", encoding='utf-8') as arquivo:
        linhas = arquivo.readlines()

        return [linha.strip().split("/") for linha in linhas]

def verificar_orden_no_job(ordem_email):
    emails = inbox.Folders("ordens").Items
    for email in emails:
        if email.SenderEmailAddress == "weg@weg.net":
            arquivo = baixar_job(email)
            job = extrair_job()

    iterar_comparar(job, ordem_email)

def baixar_job(email):
    for anexo in email.Attachments:
        dir_path = os.path.dirname(os.path.realpath(__file__))
        caminho_completo = os.path.join(dir_path, anexo.FileName)
        anexo.SaveAsFile(caminho_completo)

        return caminho_completo
    
def extrair_job():
    with open("RELATORIO_ORDENS_LIBERADAS.txt", "r") as arquivo:
        linhas = arquivo.readlines()
        linhas = [linha.strip().split(",") for linha in linhas]
        
        linhas_divididas = [item.split("|") for linha in linhas for item in linha]
        return linhas_divididas

def iterar_comparar(job, email_arquivo):
    for pessoa in email_arquivo:
        ordem_pessoa = pessoa[0].strip()
        email_pessoa = pessoa[1].strip()
        
        for ordem_array in job:
            material_job = None

            if len(ordem_array) > 3:
                material_job = ordem_array[5].strip()
            if material_job == ordem_pessoa:
                enviar_email(ordem_pessoa, email_pessoa)
                print(ordem_pessoa, email_pessoa) 

def enviar_email(material_ordem, email):
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = f'{material_ordem} LIBERADO/A'
    mail.Body = 'Aviso de ordem liberada'

    mail.Send()

pasta_destino = "Q:\GROUPS\BR_SC_JGS_WM_DEPARTAMENTO_CALDEIRARIA\DEPARTAMENTO DE CALDEIRARIA\02 - DOCUMENTOS\16 - RELATORIOS BI\Códigos Python\Ordens Liberadas"
iniciar()
