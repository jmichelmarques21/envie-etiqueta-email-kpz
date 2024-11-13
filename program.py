import os
import shutil
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr
from datetime import datetime

# Configurações de e-mail
SMTP_SERVER = 'smtp.skymail.net.br'
SMTP_PORT = 587
EMAIL_FROM = 'jean.marques@kapazi.com.br'
PASSWORD = '6nB3iD5z'
EMAILS_TO = ['alessandro.gouveia@kapazi.com.br', 'saviane.dantas@kapazi.com.br', 'e-commerce@kapazi.com.br']

# Pasta onde as etiquetas são salvas
PASTA_ETIQUETAS = 'etiquetas'
PASTA_ENVIADAS = 'etiquetas/enviadas'

# Cria a pasta 'enviadas' se não existir
os.makedirs(PASTA_ENVIADAS, exist_ok=True)

def listar_arquivos_pdf():
    arquivos_pdf = [f for f in os.listdir(PASTA_ETIQUETAS) if f.endswith('.pdf')]
    print(f"Arquivos encontrados para envio: {arquivos_pdf}")
    return arquivos_pdf

def criar_mensagem_email(nome_arquivo):
    numero_pedido, portal = nome_arquivo.split('-')[0], nome_arquivo.split('-')[1].replace('.pdf', '')
    hora_atual = datetime.now().hour
    saudacao = "Bom dia" if hora_atual < 12 else "Boa tarde"
    return f"{saudacao}! Segue etiqueta de envio para o pedido {numero_pedido} da {portal}."

def enviar_email_com_anexo(destinatarios, arquivo_pdf):
    nome_arquivo = os.path.basename(arquivo_pdf)
    msg_conteudo = criar_mensagem_email(nome_arquivo)
    
    msg = MIMEMultipart()
    msg['From'] = formataddr(('Jean Michel Marques', EMAIL_FROM))
    msg['To'] = ', '.join(destinatarios)
    msg['Subject'] = f'Etiqueta {nome_arquivo.split("-")[0]} {nome_arquivo.split("-")[1].replace(".pdf", "")}'

    msg.attach(MIMEText(msg_conteudo, 'plain'))

    with open(arquivo_pdf, 'rb') as f:
        pdf_anexo = MIMEApplication(f.read(), _subtype='pdf')
        pdf_anexo.add_header('Content-Disposition', 'attachment', filename=nome_arquivo)
        msg.attach(pdf_anexo)

    try:
        print(f"Tentando enviar o e-mail para {destinatarios} com o arquivo {nome_arquivo}...")
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            print("Conectando ao servidor SMTP...")
            server.starttls()
            print("Iniciando login...")
            server.login(EMAIL_FROM, PASSWORD)
            print("Enviando e-mail...")
            server.sendmail(EMAIL_FROM, destinatarios, msg.as_string())
            print("E-mail enviado.")
        print(f"Etiqueta '{nome_arquivo}' enviada para: {', '.join(destinatarios)}")
        return True
    except smtplib.SMTPException as e:
        print(f"Erro de SMTP ao enviar a etiqueta '{nome_arquivo}': {e}")
    except Exception as e:
        print(f"Erro inesperado ao enviar a etiqueta '{nome_arquivo}': {e}")
    return False

def mover_para_enviadas(arquivo_pdf):
    nome_arquivo = os.path.basename(arquivo_pdf)
    destino = os.path.join(PASTA_ENVIADAS, nome_arquivo)
    try:
        shutil.move(arquivo_pdf, destino)
        print(f"Arquivo '{nome_arquivo}' movido para a pasta 'enviadas'.")
    except Exception as e:
        print(f"Erro ao mover '{nome_arquivo}' para 'enviadas': {e}")

arquivos_pdf = listar_arquivos_pdf()

if arquivos_pdf:
    for arquivo in arquivos_pdf:
        caminho_arquivo = os.path.join(PASTA_ETIQUETAS, arquivo)
        print(f"Iniciando o envio da etiqueta: {caminho_arquivo}")

        if enviar_email_com_anexo(EMAILS_TO, caminho_arquivo):
            mover_para_enviadas(caminho_arquivo)
        else:
            print(f"Falha no envio da etiqueta '{arquivo}', pulando...")

    print("Processo concluído: Todos os e-mails foram processados.")
else:
    print("Nenhuma etiqueta encontrada.")
