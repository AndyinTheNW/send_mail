"""
WORK IN PROGRESS
"""

import os
import smtplib
import shutil
import logging
import xml.etree.ElementTree as ET
from datetime import datetime as dt
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from time import sleep

logging.basicConfig(
    filename="mailtoclient.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logging.info("Starting the program")

SMTP_HOST = "smtp.office365.com"
SMTP_PORT = 587
EMAIL_FROM = "@outlook.com"
PASSWORD = "password"
EMAIL_TO = "@outlook.com"
BASE_PATH = os.path.join(os.path.expanduser("~"), "pyfiles")
IMAGE_PATH = os.path.join(BASE_PATH, "car.jpeg")



def get_files_from_directory(directory, file_ext):
    return [f for f in os.listdir(directory) if f.endswith(file_ext)]


def initialize_directories():
    all_send_directory = os.path.join(BASE_PATH, "send")
    send_costumer_directory = os.path.join(all_send_directory, "send_costumer")
    for dir in [all_send_directory, send_costumer_directory]:
        if not os.path.exists(dir):
            os.makedirs(dir)
    return all_send_directory, send_costumer_directory


def extract_xml_data(path_file):
    tree = ET.parse(path_file)
    root = tree.getroot()
    namespaces = {"ns": "http://www.portalfiscal.inf.br/nfe"}
    infNFe_element = root.find(".//ns:infNFe_", namespaces)
    prod_element = root.find(".//ns:prod", namespaces)
    email_element = root.find(".//ns:email", namespaces)

    cfop = (
        prod_element.findtext("ns:CFOP", namespaces=namespaces)
        if prod_element
        else None
    )
    ncm = (
        prod_element.findtext("ns:NCM", namespaces=namespaces) if prod_element else None
    )
    costumer_email = email_element.text if email_element else None

    return cfop, ncm, costumer_email


def send_email(file, recipients, subject, body):
    server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
    server.starttls()
    server.login(EMAIL_FROM, PASSWORD)

    for recipient in recipients:
        msg = MIMEMultipart()
        msg["From"] = EMAIL_FROM
        msg["To"] = recipient
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "html" if "<html>" in body else "plain"))

        for attachment_path, subtype in [
            (file, "xml"),
            (file.replace(".xml", ".pdf"), "pdf"),
        ]:
            if os.path.exists(attachment_path):
                with open(attachment_path, "rb") as f:
                    attachment = MIMEApplication(f.read(), _subtype=subtype)
                    attachment.add_header(
                        "Content-Disposition",
                        "attachment",
                        filename=os.path.basename(attachment_path),
                    )
                    msg.attach(attachment)

        server.sendmail(EMAIL_FROM, recipient, msg.as_string())
    server.quit()


def main():
    try:
        all_send_directory, send_costumer_directory = initialize_directories()
        while True:
            for file in get_files_from_directory(BASE_PATH, ".xml"):
                path_file = os.path.join(BASE_PATH, file)
                cfop, ncm, costumer_email = extract_xml_data(path_file)

                recipients = [EMAIL_TO]
                if costumer_email:
                    recipients.append(costumer_email)

                subject = "Company files XML and PDF"
                body = "Hello Company,\n\nPlease find attached the XML file and the PDF file.\n\nYours sincerely,\nAnderson"
                if costumer_email:
                    body = f"""
                    <html>
                    <body>
                        <p>Hello, costumer! I hope this e-mail finds you well</p>
                        <p>Please find attached the files relating to the purchase of your newest car! Congratulations!</p>
                        <img src="cid:cars_image" alt="cars image">
                        <p>Yours sincerely,</p>
                        <p>Anderson</p>
                    </body>
                    </html>
                    """
                send_email(path_file, recipients, subject, body)
                shutil.move(path_file, all_send_directory)
                pdf_path = path_file.replace(".xml", ".pdf")
                if os.path.exists(pdf_path):
                    shutil.move(pdf_path, all_send_directory)
            sleep(5)

    except Exception as e:
        logging.error("An error occurred: %s", e)

        SMTP_HOST = "smtp-mail.outlook.com"
        ADMIN_EMAILS = ["@bravocorp.com.br"]
        subject = "PyFiles application error"
        body = f"Warning: an error has occurred in the Pyfiles - GWM application. Please correct as soon as possible. Error:\n{e}"
        for email_admin in ADMIN_EMAILS:
            send_email("", [email_admin], subject, body)
    finally:
        logging.info("End of the program")


if __name__ == "__main__":
    main()
