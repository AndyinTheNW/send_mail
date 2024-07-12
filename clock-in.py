from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
import schedule
from datetime import datetime, timedelta
import time
import logging
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from smtplib import SMTPAuthenticationError
import random

from constants import *

logging.basicConfig(
    filename="clockIn.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%d-%m-%Y %H:%M:%S",
)

options = webdriver.ChromeOptions()
options.add_argument("--headless")

driver = webdriver.Chrome(
    service=ChromeService(ChromeDriverManager().install()), options=options
)

def generate_random_time(base_time, range_minutes, range_seconds=300):
    """
    Gera um horário aleatório em torno do horário base dentro do intervalo especificado de minutos e segundos.
    """
    total_seconds = range_minutes * 60 + range_seconds
    delta = timedelta(seconds=random.randint(-total_seconds, total_seconds))
    return (base_time + delta).strftime("%H:%M:%S")

def clock_in(event):
    try:
        logging.info(f"{event} às: %s", datetime.now())
        print(f"{event} às:", datetime.now())

        driver.get(SITE)
        wait = WebDriverWait(driver, 10)

        email_field = wait.until(EC.element_to_be_clickable((By.ID, ID_EMAIL)))
        email_field.clear()
        email_field.send_keys(EMAIL)

        password_field = driver.find_element(By.ID, ID_SENHA)
        password_field.clear()
        password_field.send_keys(SENHA)

        driver.find_element(By.ID, LOGIN).click()

        wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_CLOCK_IN))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_CONFIRM))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_OK))).click()
        time.sleep(5)
        # Voltar para a página de login
        driver.get(SITE)

        logging.info(f"Sucesso ao {event} às: %s", datetime.now())
        print(f"Sucesso ao {event} às:", datetime.now())

        send_email_running(
            f"Informando que {event} às {datetime.now().strftime('%H:%M:%S')} foi realizado com sucesso!"
        )

    except Exception as error:
        logging.error(f"Falha ao {event}: {error}")
        print(f"Falha ao {event}:", error)
        send_email_error(f"Erro Crítico: {event} falhou às {datetime.now()}")

def send_email(subject, body):
    print("Enviando e-mail...\n De", EMAIL_NOTIF, "\nPara:", EMAIL)
    try:
        msg = MIMEMultipart()
        msg["From"] = EMAIL_NOTIF
        msg["To"] = EMAIL
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.starttls()
            server.login(EMAIL_NOTIF, SENHA_EMAIL)
            server.sendmail(EMAIL_NOTIF, EMAIL_NOTIF, msg.as_string())
    except SMTPAuthenticationError:
        print(
            "Falha na autenticação SMTP. Por favor, verifique suas credenciais de e-mail ou configurações de segurança."
        )

def send_email_running(message):
    subject = "Notificação do Software de Registro de Ponto"
    body = f"Olá, Andy!\n{message}\n\nAtenciosamente,\nSeu Software de Registro de Ponto. \n\n\n\n"
    send_email(subject, body)

def send_email_error(message):
    subject = "Alerta de Erro do Software de Registro de Ponto"
    body = f"Olá, Andy!\n\nEsta é uma mensagem automática para alertá-lo sobre um erro no Software de Registro de Ponto. Aqui estão os detalhes:\n\n{message}\n\nPor favor, verifique o software o mais rápido possível.\n\nAtenciosamente,\nSeu Software de Registro de Ponto. \n\n\n\n"
    send_email(subject, body)

def schedule_clock_in():
    initial_time = datetime.now().replace(hour=8, minute=40, second=0, microsecond=0)
    lunch_time = initial_time + timedelta(hours=4)
    back_from_lunch = initial_time + timedelta(hours=5)
    leaving_time = initial_time + timedelta(hours=9)

    # Gera horários aleatórios dentro de um intervalo de 5 minutos (300 segundos) em torno de cada evento
    random_initial_time = generate_random_time(initial_time, 5)
    random_lunch_time = generate_random_time(lunch_time, 5)
    random_back_from_lunch = generate_random_time(back_from_lunch, 5)
    random_leaving_time = generate_random_time(leaving_time, 5)

    logging.info("Iniciando o software às: %s", datetime.now().strftime("%H:%M:%S"))
    print("Iniciando o software às:", datetime.now().strftime("%H:%M:%S"))

    send_email_running(
        "Vamos trabalhar! O software de registro de ponto foi iniciado e funcionará até o final do seu expediente. \n\nAqui está sua programação para hoje:\nRegistrar ponto às: "
        + random_initial_time
        + "\nHora do almoço às: "
        + random_lunch_time
        + "\nVoltar do almoço às: "
        + random_back_from_lunch
        + "\nSaída do trabalho às: "
        + random_leaving_time
        + "\nTenha um bom dia e um bom trabalho"
    )

    schedule.every().day.at(random_initial_time).do(clock_in, event="Registrado ponto")
    schedule.every().day.at(random_lunch_time).do(clock_in, event="Hora do almoço")
    schedule.every().day.at(random_back_from_lunch).do(clock_in, event="Voltou do almoço")
    schedule.every().day.at(random_leaving_time).do(clock_in, event="Saindo do trabalho")

    logging.info("Todas as marcações de ponto agendadas com sucesso.")
    print("Todas as marcações de ponto agendadas com sucesso.")

    while True:
        schedule.run_pending()
        time.sleep(1)

        if datetime.now().strftime("%H:%M:%S") > random_leaving_time:
            logging.info("O software terminou de rodar hoje.")
            print("O software terminou de rodar hoje.")
            send_email_running(
                "O software de registro de ponto terminou de rodar hoje. Obrigado por usá-lo! Suas marcações de ponto foram registradas com sucesso desde o início do seu expediente até o final. \n \n Tenha um ótimo dia e até amanhã!\n"
            )
            time.sleep(5)
            break

if __name__ == "__main__":
    schedule_clock_in()
