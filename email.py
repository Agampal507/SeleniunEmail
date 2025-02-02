from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


service = Service(executable_path="C:\\Users\\Lenovo\\Desktop\\selenium\\chromedriver-win64\\chromedriver.exe")
driver = webdriver.Chrome(service=service)
driver.maximize_window()


df = pd.read_excel("C:\\Users\\Lenovo\\Desktop\\selenium\\credentials.xlsx")
email = df.iloc[0]['Email']
password = df.iloc[0]['Password']

driver.get("https://mail.google.com/")

email_input = driver.find_element(By.ID, "identifierId")
email_input.send_keys(email)
driver.find_element(By.ID, "identifierNext").click()
time.sleep(3)


password_input = driver.find_element(By.NAME, "password")
password_input.send_keys(password)
driver.find_element(By.ID, "passwordNext").click()
time.sleep(5)

if "inbox" in driver.current_url:
    df.loc[0, 'Result'] = "Success"
else:
    df.loc[0, 'Result'] = "Failure"

df.to_excel("C:\\Users\\Lenovo\\Desktop\\selenium\\credentials.xlsx", index=False)
driver.find_element(By.XPATH, "//div[text()='Compose']").click()
time.sleep(2)

to_field = driver.find_element(By.NAME, "to")
subject_field = driver.find_element(By.NAME, "subjectbox")
subject_field.send_keys("Automation Test Email")

body_field = driver.find_element(By.XPATH, "//div[@aria-label='Message Body']")
body_field.send_keys("Hello, \nThis is an automated email sent using Selenium.\nRegards.")

send_button = driver.find_element(By.XPATH, "//div[text()='Send']")
send_button.click()
time.sleep(3)

driver.find_element(By.XPATH, "//a[contains(@aria-label, 'Google Account')]").click()
time.sleep(2)
driver.find_element(By.XPATH, "//a[text()='Sign out']").click()
time.sleep(3)

driver.quit()
