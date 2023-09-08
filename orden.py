import openpyxl
from logging import error
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import sys
#  credentials
# Cargar el archivo de Excel
numero  = sys.argv[1]
orden = sys.argv[2]
num = numero
ord = orden
username = "carmijo"
password = "Turion.4"
codigo = "p00095"
 
    # initialize the Chrome driver
    #driver = webdriver.Chrome(r"chromedriver")
"""options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument('window-size=1200x1600')
"""
    #options.add_argument("disable-gpu")
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
driver.find_element("id", "TE_FAST_PATH_BOX").send_keys("p43025")
time.sleep(1)
driver.find_element("id" , "fastPathButton").click()
time.sleep(1)

driver.switch_to.frame(8)
time.sleep(1)
time.sleep(1)
driver.find_element(By.ID, "C0_10").send_keys(num)  
time.sleep(2)
driver.find_element(By.ID, "C0_12").click()

driver.find_element(By.ID, "C0_12").send_keys(ord)  
time.sleep(2)
driver.find_element(By.ID, "hc_Find").click()
time.sleep(1)
driver.find_element(By.ID, "selectAll0_1").click() 
time.sleep(2)
driver.find_element(By.ID, "C0_53").click()
time.sleep(2)
driver.find_element("xpath","//nobr[contains(text(), 'Actualizar estado')]").click() 

"""
assert driver.switch_to.alert.text == "¿Está seguro que desea eliminar el dato seleccionado?"
driver.switch_to.alert.accept()
#driver.find_element("name", "a").send_keys(Keys.ENTER)
#driver.find_element("id", "drop_mainmenu").send_keys(Keys.ENTER)
"""
for _ in range(15):
  #tiempo de espera hasta iniciar de nuevo 
  
  time.sleep(2)
  driver.find_element(By.ID, "hc_Find").click()
  time.sleep(1)
  driver.find_element(By.ID, "selectAll0_1").click()
  time.sleep(1)
  driver.find_element(By.ID, "hc_Select").click()
  time.sleep(3)
  driver.find_element(By.ID, "divC0_30").click()
  time.sleep(1)
  driver.find_element(By.ID, "hc_OK").click()
  time.sleep(3)
  driver.find_element(By.ID, "divC0_30").click()
  time.sleep(1)
  time.sleep(1)
  driver.find_element(By.ID, "hc_OK").click()
  time.sleep(1)
  time.sleep(1)
  driver.find_element(By.ID, "divC0_30").click()
  time.sleep(1)
  driver.find_element(By.ID, "hc_Select").click()
  time.sleep(1)
  driver.find_element(By.ID, "hc_OK").click()
  time.sleep(1)
  driver.find_element(By.ID, "hc_Select").click()
  time.sleep(1)
  driver.find_element(By.ID, "divC0_30").click()
  time.sleep(1)
  driver.find_element(By.ID, "hc_Select").click()
  time.sleep(1)
  driver.find_element(By.ID, "hc_OK").click()
"""
driver.find_element(By.ID, "PO8T0").send_keys("110046")
driver.find_element(By.ID, "PO2T0").click()
driver.find_element(By.ID, "PO2T0").send_keys("l12089")
driver.find_element(By.ID, "PO3T0").click()
driver.find_element(By.ID, "PO3T0").send_keys("f21001a1")
driver.find_element(By.ID, "PO4T0").click()
driver.find_element(By.ID, "PO4T0").send_keys("  ")
driver.find_element(By.ID, "PO7T0").click()
driver.find_element(By.ID, "PO7T0").send_keys("a")
driver.find_element(By.ID, "hc_Select").click()
driver.find_element(By.ID, "hc_OK").click()
"""
''' 
for row in sheet.iter_rows(min_row=2, values_only=True):
    item = row[0]
    bodega = row[1]
    ubicacion = row[2]
    lote = row[3]

    driver.find_element(By.ID, "PO8T0").click()
    driver.find_element(By.ID, "PO8T0").clear()
    driver.find_element(By.ID, "PO8T0").send_keys(item)
    driver.find_element(By.ID, "PO2T0").click()
    driver.find_element(By.ID, "PO2T0").send_keys(bodega)
    driver.find_element(By.ID, "PO3T0").click()
    driver.find_element(By.ID, "PO3T0").clear()
    #para campo vació y completos
    #driver.find_element(By.ID, "PO3T0").send_keys(" ")
    driver.find_element(By.ID, "PO3T0").send_keys(ubicacion)
    
    driver.find_element(By.ID, "PO4T0").click()
    driver.find_element(By.ID, "PO4T0").clear()
   # driver.find_element(By.ID, "PO4T0").send_keys("")
    driver.find_element(By.ID, "PO4T0").send_keys(lote)
    
    driver.find_element(By.ID, "PO7T0").click()
    driver.find_element(By.ID, "PO7T0").send_keys("A")
    driver.find_element(By.ID, "hc_Select").click()
    driver.find_element(By.ID, "hc_OK").click()

    driver.find_element(By.ID, "hc_Select").click()
    driver.find_element(By.ID, "divC0_30").click()
'''    



# close the driver
#driver.close()