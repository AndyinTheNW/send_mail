# Importando bibliotecas necessárias
import os
import shutil
import smtplib
import xml.etree.ElementTree as ET
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from time import sleep

# Configurações de email e diretórios
CONFIG = {
    "smtp_host": "smtp.office365.com",
    "smtp_port": 587,
    "email_from": "destino@outlook.com",
    "password": "senha",
    "email_to": "DestinoEmpresaXML@outlook.com",
    "dir_path": os.path.join(os.path.expanduser("~"), "pyfiles"),
    "image_path": os.path.join(os.path.expanduser("~"), "pyfiles", "car.jpeg"),
}

# Configurando o logging para registrar informações e erros
logging.basicConfig(
    filename="mailtoclient.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# Função para extrair texto de um elemento XML com um caminho específico
def extract_from_xml(xml_file_path, xpath):
    tree = ET.parse(xml_file_path)
    root = tree.getroot()

    namespaces = {"ns": "http://www.portalfiscal.inf.br/nfe"}
    element = root.find(xpath, namespaces)
    
    return element.text if element else None

# Função para buscar CFOP e NCM de um arquivo XML
def find_cfop_ncm_from_xml(xml_file_path):
    cfop = extract_from_xml(xml_file_path, ".//ns:NFe/ns:infNFe/ns:det/ns:prod/ns:CFOP")
    ncm = extract_from_xml(xml_file_path, ".//ns:NFe/ns:infNFe/ns:det/ns:prod/ns:NCM")
    
    return cfop, ncm

# Função para buscar o email do cliente de um arquivo XML
def find_client_email_from_xml(xml_file_path):
    return extract_from_xml(xml_file_path, ".//ns:NFe/ns:infNFe/ns:entrega/ns:email")

# Função para enviar email com anexos
def send_email(email_to, subject, body, attachments=[]):
    server = smtplib.SMTP(CONFIG["smtp_host"], CONFIG["smtp_port"])
    server.starttls()
    server.login(CONFIG["email_from"], CONFIG["password"])
    
# Configuração do email
    msg = MIMEMultipart()
    msg["From"] = CONFIG["email_from"]
    msg["To"] = email_to
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))
    
# Laço para anexar arquivos ao email
    for file_path in attachments:
        with open(file_path, "rb") as file:
            mime_type = "xml" if file_path.endswith(".xml") else "pdf"
            attach = MIMEApplication(file.read(), _subtype=mime_type)
            attach.add_header(
                "Content-Disposition",
                "attachment",
                filename=os.path.basename(file_path)
            )
            msg.attach(attach)

    server.sendmail(CONFIG["email_from"], email_to, msg.as_string())
    server.quit()

# Função para configurar e criar diretórios para arquivos enviados
def setup_directories():
    all_sent_dir = os.path.join(CONFIG["dir_path"], "enviados")
    sent_to_client_dir = os.path.join(all_sent_dir, "enviados_cliente")

    for dir in [all_sent_dir, sent_to_client_dir]:
        if not os.path.exists(dir):
            os.makedirs(dir)

    return all_sent_dir, sent_to_client_dir

# Função principal que controla a lógica do programa
def main():
    logging.info("Iniciando o programa")

    try:
        all_sent_dir, sent_to_client_dir = setup_directories()

        # Loop infinito para verificar constantemente novos arquivos XML
        while True:
            xml_files = [f for f in os.listdir(CONFIG["dir_path"]) if f.endswith(".xml")]

            for xml_file in xml_files:
                xml_file_path = os.path.join(CONFIG["dir_path"], xml_file)
                cfop, ncm = find_cfop_ncm_from_xml(xml_file_path)

                # Condição específica para identificar os XMLs de interesse
                if cfop == "6102" and ncm and ncm.startswith("8703"):
                    client_email = find_client_email_from_xml(xml_file_path)
                    if client_email:
                        send_email(client_email, "Subject", "Body", [xml_file_path])
                        shutil.move(xml_file_path, os.path.join(sent_to_client_dir, xml_file))

            # Pausa o programa por 5 segundos antes de verificar novamente os arquivos
            sleep(5)

    except Exception as e:
        logging.error("Ocorreu um erro: %s", e)
    finally:
        logging.info("Fim do programa")

# Inicializa o programa quando executado como script
if __name__ == "__main__":
    main()
