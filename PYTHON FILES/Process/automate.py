from selenium import webdriver
from selenium.webdriver.common.by import By
from db import clientes
from folders import getFielFiles, getImssFile
from utility import workaroundWrite
import pyautogui
import time

link = "https://idse.imss.gob.mx/imss/"
sleepShort = 2
sleepLarge = 3

def downloadClientes():

    for cliente in clientes:
        driver = loginCliente(cliente)
        driver = requestEmitions(driver)
        driver = checkForButtons(driver) # May be possible, there are generate link btns
        downloadFiles(driver)


def requestEmitions(driver):

    driver.find_elements(By.TAG_NAME, "H3")[1].find_element(By.XPATH, ".//*").click()
    time.sleep(sleepLarge)
    driver.find_element(By.NAME, "consultarEmisiones").click()
    time.sleep(sleepLarge)
    driver.find_element(By.XPATH, "//*[text()='Descargar emisión actual']").click()
    time.sleep(sleepLarge)
    return driver


def checkForButtons(driver):

    while (len(driver.find_elements(By.XPATH, "//*[text()='Iniciar descarga']")) > 0):
        driver.find_elements(By.XPATH, "//*[text()='Iniciar descarga']")[0].click()
        time.sleep(sleepLarge)

    return driver


def downloadFiles(driver):

    for file in driver.find_elements(By.XPATH, "//*[text()='Archivo generado']"):
        file.find_element(By.XPATH, 'following-sibling::*[1]').click()
        time.sleep(sleepShort)

        while(driver.find_element(By.XPATH, "//*[text()='Mensaje del sistema']").is_displayed()):   
            time.sleep(sleepShort)
        # GO TO DOWNLOADS, CHECK IF THAT DOWNLOAD IS IN FILE LIST
            # IF IS IN FILE LIST ERASE IT
            # ELSE COPY IT TO DESCARGAS/RFC FOLDER AND ERASE IT FROM DOWNLOADS
    # DELETE OR CLOSE WEB DRIVER


def loginCliente(cliente):

    driver = openDriver()
    dotpfx = None
    dotcer = None
    dotkey = None

    if (cliente['categoria'] == "FIEL"):
        dotcer, dotkey = getFielFiles(cliente['rfc'])
    elif (cliente['categoria'] == "IMSS"):
        dotpfx = getImssFile(cliente['rfc'])
        
    driver = login(cliente=cliente, dotcer=dotcer, dotkey=dotkey, dotpfx=dotpfx, driver=driver)
    
    return driver


def openDriver():

    driver = webdriver.Edge()
    driver.maximize_window()
    driver.get(link)
    return driver


def login(cliente=None, dotcer=None, dotkey=None, dotpfx=None, driver=None):

    time.sleep(sleepLarge)
    driver.find_element(By.XPATH, "//*[text()='Cerrar']").click()
    time.sleep(sleepShort)
    driver.find_element(By.ID, "certificado").find_element(By.XPATH, "..").click()
    time.sleep(sleepShort)
    if ( dotcer != None ):
        workaroundWrite(dotcer)
        time.sleep(sleepShort)
        pyautogui.press("enter")
        time.sleep(sleepShort)
        driver.find_element(By.ID, "llave").find_element(By.XPATH, "..").click()
        time.sleep(sleepShort)
        workaroundWrite(dotkey)
        time.sleep(sleepShort)
        pyautogui.press("enter")
        time.sleep(sleepShort)
    else:
        workaroundWrite(dotpfx)
        time.sleep(sleepShort)
        pyautogui.press("enter")
        time.sleep(sleepShort)

    driver.find_element(By.ID, "idUsuario").find_element(By.XPATH, "..").click()
    time.sleep(sleepShort)
    workaroundWrite(cliente['usuario'])
    time.sleep(sleepShort)
    driver.find_element(By.ID, "password").find_element(By.XPATH, "..").click()
    time.sleep(sleepShort)
    workaroundWrite(cliente['contraseña'])
    time.sleep(sleepShort)
    driver.find_element(By.XPATH, "//*[text()='Iniciar sesión']").click()
    time.sleep(sleepLarge)

    return driver




