import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header

# Configurar informações de login do Gmail
email = 'seuemail@gmail.com'
senha = 'senhaXdeYapp'

# Configurar servidor SMTP do Gmail
smtp_server = 'smtp.gmail.com'
smtp_port = 587

tabela = pd.read_excel('arquivo.xlsx')

# Loop para enviar e-mails para cada linha no DataFrame
for i, email_destinatario in enumerate(tabela["email_destinatario"]):
    assunto = tabela.loc[i, "assunto"]
    mensagem = tabela.loc[i, "mensagem"]
    destinatarios_e_pdfs = tabela.loc[i, "destinatarios_e_pdfs"]

    # Certifique-se de que destinatarios_e_pdfs seja uma lista
    destinatarios_e_pdfs = destinatarios_e_pdfs.split(';')  # Supondo que os nomes dos arquivos estão separados por ';'

    # Iniciar uma nova instância de mensagem para cada e-mail
    msg = MIMEMultipart()
    msg['From'] = email
    msg['To'] = email_destinatario
    msg['Subject'] = Header(assunto, 'utf-8')

    # Adicionar corpo da mensagem
    msg.attach(MIMEText(mensagem, 'plain'))

    # Loop para adicionar anexos
    for pdf_anexo in destinatarios_e_pdfs:
        with open(pdf_anexo.strip(), 'rb') as anexo:
            pdf = MIMEApplication(anexo.read(), _subtype='pdf')
            pdf.add_header('content-disposition', 'attachment', filename=pdf_anexo.strip())
            msg.attach(pdf)

    # Iniciar conexão SMTP
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(email, senha)

    # Enviar o email
    server.send_message(msg)

    # Encerrar conexão SMTP
    server.quit()

    print('Email enviado')