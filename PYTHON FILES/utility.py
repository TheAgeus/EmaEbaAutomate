from datetime import datetime
import pyperclip
import pyautogui
import time
from selenium.webdriver.common.by import By



def clickByText(texto, driver, ofsetX, ofsetY):
    while(True):
            # Le damos un pequeño espacio para que cargue la página de login
            sleep(5)

            # Verificar si ya estamos en la página de login, el tamaño de ese arreglo debe ser distinto de 0 para saber que ya estás allí con el mensaje de "Importante"
            elemento = driver.find_elements(By.XPATH, f"//*[text()='{texto}']")
            if(len(elemento) == 0):
                # Todavía no carga
                # Le damos otros 5 segundos, si no, refrescamos
                sleep(5)
                if(len(driver.find_elements(By.XPATH, f"//*[text()='{texto}']")) == 0):
                    pyautogui.press("F5")
            # Si ya encontramos el mensaje entonces podemos romper este ciclo de espera
            else: 
                pyautogui.moveTo(elemento[0].location['x'] + 20 + ofsetX, elemento[0].location['y'] + 140 + ofsetY)
                pyautogui.leftClick()
                break


def clickElement(elemento):
    pyautogui.moveTo(elemento[0].location['x'] + 20, elemento[0].location['y'] + 140)
    pyautogui.leftClick()

def clickById(id, driver):
    while(True):
            # Le damos un pequeño espacio para que cargue la página de login
            sleep(5)

            # Verificar si ya estamos en la página de login, el tamaño de ese arreglo debe ser distinto de 0 para saber que ya estás allí con el mensaje de "Importante"
            elemento = driver.find_elements(By.ID, id)
            if(len(elemento) == 0):
                # Todavía no carga
                # Le damos otros 5 segundos, si no, refrescamos
                sleep(5)
                if(len(driver.find_elements(By.ID, id)) == 0):
                    pyautogui.press("F5")
            # Si ya encontramos el mensaje entonces podemos romper este ciclo de espera
            else: 
                pyautogui.moveTo(elemento[0].location['x'] + 20, elemento[0].location['y'] + 140)
                pyautogui.leftClick()
                break



def click(img, timeOut): 
    res = pyautogui.locateOnScreen(img)
    count = 0
    while(res == None):
        res = pyautogui.locateOnScreen(img)
        count = count + 1
        if(count == timeOut):
            return False
    try:
        pyautogui.moveTo(pyautogui.locateOnScreen(img))
        pyautogui.leftClick()
        pyautogui.leftClick()
        return True
    except:
        return False




def sleep(value):
    time.sleep(value)




def tab(times=1):
    for i in range(times):
                                    pyautogui.press("tab")




def enter():
    pyautogui.press("enter")




def workaroundWrite(text):
    pyperclip.copy(text)
    pyautogui.hotkey('ctrl', 'v')
    pyperclip.copy('')




def monthName(monthNumber):
    if monthNumber == 1:
        return "ENERO"
    elif monthNumber == 2:
        return "FEBRERO"
    elif monthNumber == 3:
        return "MARZO"
    elif monthNumber == 4:
        return "ABRIL"
    elif monthNumber == 5:
        return "MAYO"
    elif monthNumber == 6:
        return "JUNIO"
    elif monthNumber == 7:
        return "JULIO"
    elif monthNumber == 8:
        return "AGOSTO"
    elif monthNumber == 9:
        return "SEPTIEMBRE"
    elif monthNumber == 10:
        return "OCTUBRE"
    elif monthNumber == 11:
        return "NOVIEMBRE"
    elif monthNumber == 12:
        return "DICIEMBRE"
    else:
        return "Número de mes no válido"




def getLastMonth():
   # Get the current date
    current_date = datetime.now()

    # Calculate the last month
    if current_date.month == 1:
        last_month = 12
    else:
        last_month = current_date.month - 1
    return last_month
