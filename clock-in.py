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

from constants import *


logging.basicConfig(
    filename="clockIn.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%d-%m-%Y %H:%M:%S",
)

options = webdriver.ChromeOptions()
# options.add_argument("--headless")

driver = webdriver.Chrome(
    service=ChromeService(ChromeDriverManager().install()), options=options
)


def clock_in(event):
    try:
        logging.info(f"{event} at: %s", datetime.now())
        print(f"{event} at:", datetime.now())

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
        # back to the login page
        driver.get(SITE)

        logging.info(f"Successfully {event} at: %s", datetime.now())
        print(f"Successfully {event} at:", datetime.now())

        send_email_running(
            f"Just to let you know that {event} at {datetime.now().strftime('%H:%M')} was successful!"
        )

    except Exception as error:
        logging.error(f"Failed to {event}: {error}")
        print(f"Failed to {event}:", error)
        send_email_error(f"Critical Error: {event} Failed at {datetime.now()}")


def send_email(subject, body):
    print("Sending email...\n From", EMAIL_NOTIF, "\nTo:", EMAIL)
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
            "SMTP Authentication failed. Please check your email credentials or security settings."
        )


def send_email_running(message):
    subject = "Clock-in Software Notification"
    body = f"Hey, Andy!\n{message}\n\nBest,\nYour Clock-in Software. \n\n\n\n"
    send_email(subject, body)


def send_email_error(message):
    subject = "Clock-in Software Error Alert"
    body = f"Hey, Andy!\n\nThis is an automated message to alert you about an error in the Clock-in Software. Here are the details:\n\n{message}\n\nPlease check the software as soon as possible.\n\nBest,\nYour Clock-in Software. \n\n\n\n"
    send_email(subject, body)


def schedule_clock_in():

    initial_time = datetime.now().replace(hour=8, minute=40, second=0, microsecond=0)
    lunch_time = (initial_time + timedelta(hours=4)).strftime("%H:%M")
    back_from_lunch = (initial_time + timedelta(hours=5)).strftime("%H:%M")
    leaving_time = (initial_time + timedelta(hours=9)).strftime("%H:%M")

    logging.info("Starting the software at: %s", datetime.now().strftime("%H:%M"))
    print("Starting the software at:", datetime.now().strftime("%H:%M"))

    send_email_running(
        "Let's get to work! The Clock-in software has been started, and it will run until the end of your workday. \n\nHere is your schedule for today:\nClock in at: "
        + initial_time.strftime("%H:%M")
        + "\nLunch time at: "
        + lunch_time
        + "\nBack from lunch at: "
        + back_from_lunch
        + "\nLeaving work at: "
        + leaving_time
        + "\nEnjoy your day and have a good work"
    )

    schedule.every().day.at(initial_time.strftime("%H:%M")).do(
        clock_in, event="Clocked in"
    )
    schedule.every().day.at(lunch_time).do(clock_in, event="Lunch time")
    schedule.every().day.at(back_from_lunch).do(clock_in, event="Back from lunch")
    schedule.every().day.at(leaving_time).do(clock_in, event="Leaving work")

    logging.info("Successfully set all scheduled clock-ins.")
    print("Successfully set all scheduled clock-ins.")

    while True:
        schedule.run_pending()
        time.sleep(1)

        if datetime.now().strftime("%H:%M") > leaving_time:
            logging.info("The software has finished running for today.")
            print("The software has finished running for today.")
            send_email_running(
                "The Clock-in software has finished running for today. Thank you for using it! Your clock-ins have been successfully registered from the start of your workday until the end. \n \n Have a great day and see you tomorrow!\n"
            )
            time.sleep(5)

            break


if __name__ == "__main__":
    schedule_clock_in()