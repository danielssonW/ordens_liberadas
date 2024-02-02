import win32com.client
import re
import os
from win32com.client import Dispatch
import time

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
    emails = inbox.Items
    for email in emails:
        if email.SenderEmailAddress == "weg@weg.net":
            arquivo = baixar_job(email)
            job = extrair_job()

    iterar_comparar(job, ordem_email)

def baixar_job(email):
    for anexo in email.Attachments:
        dir_path = os.path.dirname(os.path.realpath(__file__))
        caminho_completo = os.path.join(dir_path, anexo.FileName)
        if not os.path.exists(dir_path):
            os.makedirs(dir_path)

        anexo.SaveAsFile(caminho_completo)

    return caminho_completo
    
def extrair_job():
    caminho_arquivo = os.path.join(os.path.dirname(os.path.realpath(__file__)), "RELATORIO_ORDENS_LIBERADAS.txt")
    with open(caminho_arquivo, "r") as arquivo:
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
                deletar_linhas_com_valor(ordem_pessoa)

def enviar_email(material_ordem, email):
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = f'{material_ordem} LIBERADO/A'
    mail.Body = 'Aviso de ordem liberada'

    mail.Send()

def deletar_linhas_com_valor(valor_a_deletar):
    with open("ordens.txt", "r", encoding='utf-8') as arquivo_original:
        linhas = arquivo_original.readlines()

    linhas_filtradas = [linha for linha in linhas if valor_a_deletar not in linha]

    with open("ordens.txt", 'w') as arquivo_modificado:
        arquivo_modificado.writelines(linhas_filtradas)

iniciar()
