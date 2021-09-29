import os
import win32com.client as win32
import os.path
import pyautogui  # para importar biblioteca
import time  # para importar biblioteca
import pyperclip  # para importar biblioteca
import getpass
import smtplib
import sys
import calendar
from pathlib import Path
import datetime
import os.path


def get_last_business_day():
    today = datetime.date.today()
    delta = max(1, (today.weekday() + 6) % 7 - 3)
    return today - datetime.timedelta(days=delta)


get_last_business_day().strftime("%Y%m%d")
get_last_business_day = (get_last_business_day().strftime("%Y%m%d"))

if (os.path.isfile(r'C:\Users\gterr\Desktop\convenio\teste3' + str(get_last_business_day) + '.txt')):
    arquivo = 'true'
else:
    arquivo = 'false'

if arquivo == 'true':
    def get_last_business_day():
        today = datetime.date.today()
        delta = max(1, (today.weekday() + 6) % 7 - 3)
        return today - datetime.timedelta(days=delta)


    get_last_business_day().strftime("%Y%m%d")
    get_last_business_day = (get_last_business_day().strftime("%Y%m%d"))

    os.rename(r'C:\Users\gterr\Desktop\convenio\teste3' + str(get_last_business_day) + '.txt',
              r'C:\Users\gterr\Desktop\convenio2\gterribele@hotmail.com.txt')

    if (os.path.isfile(r'C:\Users\gterr\Desktop\convenio2\gterribele@hotmail.com.txt')):
        arquivo = 'true'
    else:
        arquivo = 'false'

    if arquivo == 'true':
        outlook = win32.Dispatch('outlook.application')
        folder = Path(r"C:\Users\gterr\Desktop\convenio2")

        # leitura documentos em anexo
        for attachment in folder.iterdir():
            # envio de e-mail
            mail = outlook.CreateItem(0)
            mail.SentOnBehalfOfName = 'gterribele@hotmail.com'
            mail.HTMLBody = """
            <p>Olá,</p>
            <p>Segue documentos em anexo!</p>
            <p>Atenciosamente.</p>
            <p>João</p>"""
            mail.Attachments.Add(str(attachment))
            # coloca o nome do arquivo, sem a extnsão, como endereço do email antes do "@"
            mail.To = f'{attachment.stem}'
            # mail.Subject = str(attachment)
            mail.subject = ('arquivo de retorno')
            mail.display()
            time.sleep(5)
            mail.Send()
            print("email enviado com sucesso!")
    else:
        print('arquivo não existe')
else:
    print('arquivo não existe')


def get_last_business_day():
    today = datetime.date.today()
    delta = max(1, (today.weekday() + 6) % 7 - 3)
    return today - datetime.timedelta(days=delta)


get_last_business_day().strftime("%Y%m%d")
get_last_business_day = (get_last_business_day().strftime("%Y%m%d"))

if (os.path.isfile(r'C:\Users\gterr\Desktop\convenio\teste4' + str(get_last_business_day) + '.txt')):
    arquivo = 'true'
else:
    arquivo = 'false'

if arquivo == 'true':
    def get_last_business_day():
        today = datetime.date.today()
        delta = max(1, (today.weekday() + 6) % 7 - 3)
        return today - datetime.timedelta(days=delta)


    get_last_business_day().strftime("%Y%m%d")
    get_last_business_day = (get_last_business_day().strftime("%Y%m%d"))

    os.rename(r'C:\Users\gterr\Desktop\convenio\teste4' + str(get_last_business_day) + '.txt',
              r'C:\Users\gterr\Desktop\convenio2\gterribele@gmail.com.txt')

    if (os.path.isfile(r'C:\Users\gterr\Desktop\convenio2\gterribele@hotmail.com.txt')):
        arquivo = 'true'
    else:
        arquivo = 'false'

    if arquivo == 'true':
        outlook = win32.Dispatch('outlook.application')
        folder = Path(r"C:\Users\gterr\Desktop\convenio2")

        # leitura documentos em anexo
        for attachment in folder.iterdir():
            # envio de e-mail
            mail = outlook.CreateItem(0)
            mail.SentOnBehalfOfName = 'gterribele@hotmail.com'
            mail.HTMLBody = """
            <p>Olá,</p>
            <p>Segue documentos em anexo!</p>
            <p>Atenciosamente.</p>
            <p>João</p>"""
            mail.Attachments.Add(str(attachment))
            # coloca o nome do arquivo, sem a extnsão, como endereço do email antes do "@"
            mail.To = f'{attachment.stem}'
            # mail.Subject = str(attachment)
            mail.subject = ('arquivo de retorno')
            mail.display()
            time.sleep(5)
            mail.Send()
            print("email enviado com sucesso!")
    else:
        print('arquivo não existe')
else:
    print('arquivo não existe')