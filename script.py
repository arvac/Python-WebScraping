import sys
# Utilizar la variable en el script de Python
#print("Variable recibida desde VBA:", variable)

from logging import error
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
from tkinter import *
from tkinter import messagebox
# Obtener la variable pasada desde VBA
variable = sys.argv[1]
#  credentials

    #uname = e1.get()
uname = variable.strip()
username = "carmijo"
password = "Turion.4"
codigo = "p00095"

    # OR options.add_argument("--disable-gpu")
driver = webdriver.Chrome(executable_path="C:\chromedriver.exe" )
    # head to github login page
driver.get("https://jde.favoritafruit.com:8092/jde/E1Menu.maf?selectJPD920=SOP_ANLPD&envRadioGroup=&jdeowpBackButtonProtect=PROTECTED")
driver.implicitly_wait(0.5)
    # find username/email field and send the username itself to the input field
    #driver.find_element_by_id("login_field").send_keys(username)
driver.find_element("name", "User").send_keys(username)

    # find password input field and insert password as well
    #driver.find_element_by_id("password").send_keys(password)
driver.find_element("name", "Password").send_keys(password)
driver.find_element("name", "Password").send_keys(Keys.ENTER)
time.sleep(1)

    # click login button
driver.find_element("name" , "SUBMIT_BUTTON").click()
time.sleep(1)
driver.find_element("id", "drop_mainmenu").click()
time.sleep(1)
driver.find_element("id", "TE_FAST_PATH_BOX").send_keys("p00095")
time.sleep(1)
driver.find_element("id" , "fastPathButton").click()
time.sleep(1)

driver.switch_to.frame(8)
driver.find_element(By.NAME, "qbe0_1.6").click()
driver.find_element(By.NAME, "qbe0_1.6").send_keys("*" + uname + "*")
driver.find_element(By.ID, "hc_Find").click()
time.sleep(7)
driver.find_element(By.ID, "selectAll0_1").click()

driver.find_element(By.ID, "hc_Delete").click()
assert driver.switch_to.alert.text == "¿Está seguro que desea eliminar el dato seleccionado?"
driver.switch_to.alert.accept()
    #driver.find_element("name", "a").send_keys(Keys.ENTER)
    #driver.find_element("id", "drop_mainmenu").send_keys(Keys.ENTER)
    # wait the ready state to be complete
WebDriverWait(driver=driver, timeout=10).until(
        lambda x: x.execute_script("return document.readyState === 'complete'")
    )
error_message = "Incorrect username or password."
    # get the errors (if there are)
errors = driver.find_elements_by_class_name("flash-error")
    # print the errors optionally
    # for e in errors:
    #     print(e.text)
    # if we find that error message within errors, then login is failed
if any(error_message in e.text for e in errors):
        print("[!] Login failed")
else:
        print("[+] Login successful")
