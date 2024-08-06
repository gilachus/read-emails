# Script to login in saleforce
# pip install selenium webdriver_manager

import time
# import pickle
# import sys
import winsound

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import imaplib
import email
from email.header import decode_header
from datetime import datetime, timedelta
import pytz
import os

chrome_options = Options()
# chrome_options.add_argument('--headless')
# chrome_options.add_argument('--no-sandbox')
# chrome_options.add_argument('--disable-dev-shm-usage')
# chrome_options.add_argument('--verbose')


url = "https://data-ability-5274.lightning.force.com/lightning/o/Case/list?filterName=AllOpenCases"
usuario = "jesusgilberdugo+test-agnh@force.com"
contrasena = "$Test741456"

# def goforit(url):
#     try:
#         service = Service(executable_path="chromedriver.exe")
#         driver = webdriver.Chrome(service=service, options=chrome_options)
#         driver.get(url)
#         time.sleep(5)
#         username = driver.find_element(By.ID, 'username')
#         username.send_keys(usuario)

#         password = driver.find_element(By.ID, 'password')
#         password.send_keys(contrasena)
#         driver.find_element(By.ID, 'Login').click()
#         time.sleep(5)
#         newContact = WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located(
#                 (By.XPATH, "//button[@class='slds-button slds-button_brand']"))
#         )
#         print("cargó")
#         newContact.click()
#         time.sleep(5)
#         driver.quit()
#     except:
#         driver.quit()
#         winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
#         print("expiraron las cookies o algo falló")

def find_email():
    # imap_server = "outlook.office365.com"
    imap_server = os.environ.get("SERVER", "")
    imap_port = os.environ.get("PORT", "")
    email_address = os.environ.get("EMAIL", "")
    password = os.environ.get("PASSWORD", "")
    local_tz = pytz.timezone("America/Bogota")
    try:
        print("Connecting to the server...")
        imap = imaplib.IMAP4_SSL(host=imap_server, port=imap_port)
        # print(imap)
        print("Try to login...")
        imap.login(email_address, password)
    except Exception as e:
        winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
        print(f"Error connecting to the server: {e}")
        return
    
    processed_ids = set()

    while True:
        try:
            now = datetime.now(local_tz)
            start_of_today = now.replace(hour=0, minute=0, second=0, microsecond=0)
            start_of_today_utc = start_of_today.astimezone(pytz.utc)
            date_today_utc = start_of_today_utc.strftime("%d-%b-%Y")
            search_criteria = f'(SUBJECT "prospecto" SINCE {date_today_utc})'
            print("Searching for emails...")
            imap.select("inbox")
            # _, messages = imap.search(None, "FROM", "jesusgilberdugo@gmail.com")
            _, messages = imap.search(None, search_criteria)
            for msg_id in messages[0].split():
                if msg_id in processed_ids:
                    continue
                
                _, msg_data = imap.fetch(msg_id, "(RFC822)")
                msg = email.message_from_bytes(msg_data[0][1])
                print(f"msg id {msg_id}")
                print(f"From {msg.get('From')}")
                print(f"To {msg.get('To')}")
                print(f"BCC {msg.get('BCC')}")
                print(f"Subject {msg.get('Subject')}")
                print(f"Date {msg.get('Date')}")
                print("Content")
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True)
                        print(body)
                processed_ids.add(msg_id)
        except Exception as e:
            winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
            print(f"Error fetching emails: {e}")
        
        time.sleep(60)

if __name__ == '__main__':
    find_email()
