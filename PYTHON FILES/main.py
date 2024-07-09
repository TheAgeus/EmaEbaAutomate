from datetime import datetime
from utility import getLastMonth, monthName, workaroundWrite, sleep, tab, enter, click, clickById, clickByText, clickElement
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import openpyxl
import pyperclip
import pyautogui
import time
import os
import shutil

elegirArchivo = "C:\\Users\\EMA Y EBA\\Desktop\\EMA Y EBA\\IMGS\\ELEGIR ARCHIVO.png"
microsoftEdge = "C:\\Users\\EMA Y EBA\\Desktop\\EMA Y EBA\\IMGS\\MicrosoftEdge.png"
mensajeSistema = "C:\\Users\\EMA Y EBA\\Desktop\\EMA Y EBA\\IMGS\\MensajeSistema.png"
entiendo = "C:\\Users\\EMA Y EBA\\Desktop\\EMA Y EBA\\IMGS\\entiendo.png"


def esperarSeVayaImg(imgPath, timeOut):
    try:
        isThere = pyautogui.locateOnScreen(imgPath)
    except: return None
    myCounter = 0

    while(isThere != None):
        time.sleep(1)
        isThere = pyautogui.locateOnScreen(imgPath)
        myCounter = myCounter + 1
        if(myCounter == timeOut): return None

    return isThere



def waitFor(imgPath, timeOut):

    isThere = pyautogui.locateOnScreen(imgPath)
    myCounter = 0

    while(isThere == None):
        time.sleep(1)
        isThere = pyautogui.locateOnScreen(imgPath)
        myCounter = myCounter + 1
        if(myCounter == timeOut): return None

    return isThere




def esperarPorUnElementoByText(texto, driver):
    while(True):
        # Le damos un pequeño espacio para que cargue la página de login
        sleep(5)

        # Verificar si ya estamos en la página de login, el tamaño de ese arreglo debe ser distinto de 0 para saber que ya estás allí con el mensaje de "Importante"
        if(len(driver.find_elements(By.XPATH, f"//*[text()='{texto}']")) == 0):
            # Todavía no carga
            # Le damos otros 5 segundos, si no, refrescamos
            sleep(5)
            if(len(driver.find_elements(By.XPATH, f"//*[text()='{texto}']")) == 0):
                pyautogui.press("F5")
        # Si ya encontramos el mensaje entonces podemos romper este ciclo de espera
        else: break




def loginConFiel(rfc, usuario, contrasenia):

    # Abrimos una instancia de webdriver
    driver = webdriver.Edge()
    driver.maximize_window()
    driver.get("https://idse.imss.gob.mx/imss/")    

    # Obtenemos las rutas de los archivos cer y key, aqui solo se va a necesitar esos, ya que en el excel esta el usuario y la contraseña
    files = os.listdir("C:\\Users\\EMA Y EBA\\Desktop\\EMA Y EBA\FIEL\\" + rfc)
    dotcer = "C:\\Users\\EMA Y EBA\\Desktop\\EMA Y EBA\FIEL\\" + rfc + "\\" + [f for f in files if f.endswith(".cer")][0]
    dotkey = "C:\\Users\\EMA Y EBA\\Desktop\\EMA Y EBA\FIEL\\" + rfc + "\\" + [f for f in files if f.endswith(".key")][0]

    
    # Proceso de espera para saber si ya está disponible la página de inicio
    esperarPorUnElementoByText('Importante', driver)

    # Ya que sabemos que sí existe, cerramos el span de personalizar la experiencia web
    isThere = waitFor(entiendo, 3)
    pyautogui.moveTo(isThere)
    pyautogui.leftClick()
    sleep(1)
    pyautogui.leftClick()
    
    # Log in
    sleep(3)
    clickById("certificado", driver)
    time.sleep(1)
    workaroundWrite(dotcer)
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    workaroundWrite(dotcer)
    time.sleep(1)
    pyautogui.press("enter")

    clickById("llave", driver)
    pyautogui.leftClick()
    time.sleep(1)
    pyautogui.press("Enter")
    time.sleep(1)
    workaroundWrite(dotkey)
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)

    pyautogui.press("Tab")
    time.sleep(1)
    workaroundWrite(usuario)
    time.sleep(1)

    pyautogui.press("Tab")
    time.sleep(1)
    workaroundWrite(contrasenia)
    time.sleep(1)
    pyautogui.press("tab")
    time.sleep(1)

    driver.find_element(By.ID, "botonFirma").click()
    time.sleep(10)

    return driver





def loginConIMSS(rfc, usuario, contrasenia):

    # Abrimos una instancia de webdriver
    driver = webdriver.Edge()
    driver.maximize_window()
    driver.get("https://idse.imss.gob.mx/imss/")

    # Obtenemos las rutas de los archivos pfx, aqui solo se va a necesitar ese, ya que en el excel esta el usuario y la contraseña
    files = os.listdir("C:\\Users\\EMA Y EBA\\Desktop\\EMA Y EBA\IMSS\\" + rfc)
    dotpfx = "C:\\Users\\EMA Y EBA\\Desktop\\EMA Y EBA\IMSS\\" + rfc + "\\" + [f for f in files if f.endswith(".pfx")][0]

    sleep(5)

    # Proceso de espera para saber si ya está disponible la página de inicio
    esperarPorUnElementoByText('Importante', driver)

    # Ya que sabemos que sí existe, cerramos el span de personalizar la experiencia web
    isThere = waitFor(entiendo, 3)
    pyautogui.moveTo(isThere)
    pyautogui.leftClick()
    sleep(1)
    pyautogui.leftClick()

    # Log in
    sleep(3)
    clickById("certificado", driver)
    time.sleep(1)
    workaroundWrite(dotpfx)
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    workaroundWrite(dotpfx)
    time.sleep(1)
    pyautogui.press("enter")

    tab()
    time.sleep(1)
    workaroundWrite(usuario)
    time.sleep(1)
    pyautogui.press("tab")
    time.sleep(1)

    time.sleep(1)
    workaroundWrite(contrasenia)
    time.sleep(1)
    pyautogui.press("tab")
    time.sleep(1)

    driver.find_element(By.ID, "botonFirma").click()
    time.sleep(10)

    return driver



def checkForError(driver):
    # Puede que aquí nos salga un error de sesión inválida
    if(len(driver.find_elements(By.CLASS_NAME, "alert-danger")) > 0):
        return {
            "Tipo" : "Error",
            "Mensaje" : driver.find_elements(By.CLASS_NAME, "alert-danger")[0].text
        }


def goToDescargas(driver):

    # Esperar a que el mensaje de "Procesando por el sistema" o algo así se vaya
    esperarSeVayaImg(mensajeSistema, 20)

    # Esperar a que este elemento exista, esto quiere decir que ya puedo picarle a emisión
    esperarPorUnElementoByText("Consulta y descarga de emisión mensual y bimestral en formato SUA, Visor, PDF y Excel.", driver)

    # Localizarlo y darle click con offset de -50
    clickByText("Consulta y descarga de emisión mensual y bimestral en formato SUA, Visor, PDF y Excel.", driver, 0, -50)

    # Esperar un poquito
    sleep(5)

    # Puede que aquí nos salga un error de sesión inválida
    if(len(driver.find_elements(By.CLASS_NAME, "alert-danger")) > 0):
        return {
            "Tipo" : "Error",
            "Mensaje" : driver.find_elements(By.CLASS_NAME, "alert-danger")[0].text
        }


    # Verificar que sí estoy en el lugar correcto
    if (driver.find_element(By.TAG_NAME, "h1").text == "Emisión"):
        
        # Intentar llegar a "Consultar"
        clickByText("Salir", driver, 0, -75)
        sleep(5)


    # Puede que aquí nos salga un error de sesión inválida
    if(len(driver.find_elements(By.CLASS_NAME, "alert-danger")) > 0):
        return {
            "Tipo" : "Error",
            "Mensaje" : driver.find_elements(By.CLASS_NAME, "alert-danger")[0].text
        }



    # Verificar que sí estoy en el lugar correcto
    if (driver.find_element(By.TAG_NAME, "h1").text == "Emisión anticipada"):

        # Intentar llegar a "Consultar"
        clickByText("Descargar emisión actual", driver, 0, 0)
        sleep(5)

    # Sigue saliendo ese mensaje
    esperarSeVayaImg(mensajeSistema, 20)

    # Puede que aquí nos salga un error de sesión inválida
    if(len(driver.find_elements(By.CLASS_NAME, "alert-danger")) > 0):
        return {
            "Tipo" : "Error",
            "Mensaje" : driver.find_elements(By.CLASS_NAME, "alert-danger")[0].text
        }


    # Verificar que sí estoy en el lugar correcto                   
    if (driver.find_element(By.TAG_NAME, "h1").text == "Descarga emisión"):

        rowTables = driver.find_elements(By.TAG_NAME, "tr")

        # Puede que aquí estén los botones para generar los archivos
        buttons = driver.find_elements(By.XPATH, "//*[text()='Iniciar descarga']")
        while (len(driver.find_elements(By.XPATH, "//*[text()='Iniciar descarga']")) > 0):
            driver.find_elements(By.XPATH, "//*[text()='Iniciar descarga']")[0].click()
            esperarSeVayaImg(mensajeSistema, 200)

        # Intentar llegar a los links de descarga
        rowTables = driver.find_elements(By.TAG_NAME, "tr")
        pyautogui.scroll(-20)
        for rowTable in rowTables:
            try:
                rowTable.find_elements(By.TAG_NAME, "td")[-1].find_element(By.TAG_NAME, "a").click()
                
                sleep(5)
                if(len(driver.find_elements(By.CLASS_NAME, "alert-danger")) > 0):
                    return {
                        "Tipo" : "Error",
                        "Mensaje" : driver.find_elements(By.CLASS_NAME, "alert-danger")[0].text
                    }
                esperarSeVayaImg(mensajeSistema, 3000)
            except: None

    return "Done"


# Este método es para guardar los archivos en runtime, cuando se está ejecutando el programa
def guardar(rfc, descargasPath):

    # Todos los archivos que están en la carpeta de descargas
    files = [[descargasPath + "\\" + f, f] for f in os.listdir(descargasPath)]
    
    # Ruta base para intentar conseguir la ruta completa de donde deben ser colocados los archivos
    basePath = "\\\\192.168.1.77\\planeación fiscal mx\\"

    # Obtener lista de todos los folders que se encuentren de las carpetas de planeación fiscal, solo si el rfc es igual al que se pasa por parámetro, es decir que solo se va a regresar el folder del rfc que se pasa por parametro y que este en la carpeta de planeacion fiscal
    rfc_folder = [f for f in os.listdir(basePath) if f.split("_")[0] == rfc][0]

    #folder = [(basePath + f) for f in rfc_folder if os.path.isdir(basePath + f)][0]
    rfc_folder = basePath + rfc_folder

    # Por cada archivo que está en la carpeta de descargas
    for f in files:

        # Si hay otro archivo descargado que no tiene el _, pasamos de él
        if (not "_" in f[1]):
            continue

        # Si los archivos descargados son los que deberán de ser, el año y el mes están detrás del caracter _ ESTO PUEDE QUE TENGA ERROR SI NO ENCUENTRA EL _
        fecha = f[1].split("_")[0]

        # Si solo estan descargados los archivos que tienen que estar descargados, entonces el año debe de estar en los primeros 4 characters
        year = fecha[:4]

        # Si solo estan descargados los archivos que tienen que estar descargados, entonces el mes debe de estar n la posición 5
        month = fecha[4:]

        # Si no exste la carpeta en nuestro control local, la creamos, RFC/AÑO/MES
        if (not os.path.isdir("C:\\Users\\EMA Y EBA\\Desktop\\EMA Y EBA\\DESCARGAS\\" + rfc + "\\" + str(year) + "\\" + month)):
            os.makedirs("C:\\Users\\EMA Y EBA\\Desktop\\EMA Y EBA\\DESCARGAS\\" + rfc + "\\" + str(year) + "\\" + month)
        
        # Como estamos en cierto archivo, creamos la ruta completa para poder colocarlorfc_folder
        destinationPath = "C:\\Users\\EMA Y EBA\\Desktop\\EMA Y EBA\\DESCARGAS\\" + rfc + "\\" + str(year) + "\\" + month + "\\" + f[1]
        
        # Ahora para guardarlo en la carpeta de planeación fiscal, igual, si no existe esa ruta, la creamos
        if(not os.path.isdir(rfc_folder + "\\EMISIONES\\" + str(year) + "\\" + month)):
            os.makedirs(rfc_folder + "\\EMISIONES\\" + str(year) + "\\" + month)
        
        # De igual manera, creamos la ruta completa para que el archivo se guarde en la carpeta de planeación fiscal
        destinationFiscal = rfc_folder + "\\EMISIONES\\" + str(year) + "\\" + month + "\\" + f[1]

        # Aquí entonces ya intentamos copiarlos
        try:
            shutil.copy(f[0], destinationFiscal)
            shutil.copy(f[0], destinationPath)
        except: None

    print("Archivos guardados")



def acomodarDescargasLocales():
    descargasPath = "C:\\Users\\EMA Y EBA\\Desktop\\EMA Y EBA\\DESCARGAS\\"
    basePath = "\\\\192.168.1.77\\planeación fiscal mx\\"

    listaDeRfcs = os.listdir(descargasPath)

    for folder in listaDeRfcs:
        
        foldersAnios = os.listdir(descargasPath + folder)

        for folderAnio in foldersAnios:

            foldersMonths = os.listdir( descargasPath + folder + "\\" + folderAnio)

            for folderMonth in foldersMonths:

                # Esta es la carpeta del mes en el control local, es decir, no en planeación fiscal
                controlLocalFolder = descargasPath + folder + "\\" + folderAnio + "\\" + folderMonth

                # Obtener la carpeta de planeación fiscal de ese rfc
                rfc_folder = [basePath + f for f in os.listdir(basePath) if f.split("_")[0] == folder][0] + "\\EMISIONES\\" + folderAnio + "\\" + folderMonth

                # Ahora, si no existe, crearla
                if( not os.path.isdir(rfc_folder) ):
                    os.makedirs(rfc_folder)

                # Copiar los archivos nuevos de mi control local
                for archivo in os.listdir(controlLocalFolder):
                    sorucePath = controlLocalFolder + "\\" + archivo
                    destinyPath = rfc_folder  + "\\" + archivo
                    try:
                        shutil.copy(sorucePath, destinyPath)
                        print(f"{sorucePath} \n Copiado a \n {destinyPath}")
                    except: None


def borrarDescargas():
    print("Borrando descargas\n")
    path = "C:\\Users\\EMA Y EBA\\Downloads\\"     
    filesInDownloads = [path + f for f in os.listdir(path)]
    for file in filesInDownloads:
        print(f"{file} borrado\n")
        os.remove(file)

def borrarCarpeta(ruta):    
    try:
        filesInDownloads = [ruta + "\\" + f for f in os.listdir(ruta)]
        
        for file in filesInDownloads:
            print(f"{file} borrado\n")
            os.remove(file)
    except: None

# PROCEDIMIENTO PRINCIPAL O ENTRY POINT
def main():

    

    # Prellenado
    #acomodarDescargasLocales()

    # Archivo de excel que tiene los rfcs y las contraseñas
    controlExcelPath = "C:\\Users\\EMA Y EBA\\Desktop\\EMA Y EBA\\CONTROL\\Control.xlsx"

    # Cargamos el archivo de excel
    workbook = openpyxl.load_workbook(controlExcelPath)

    # Seleccionamos la primera hoja
    firstSheet = workbook.active

    # Obtenemos el maximo d filas que existen
    lastRow = firstSheet.max_row

    # Por cada fila desde la fila 1
    for i in range(1, lastRow):
        
        # Borrar descargas
        borrarDescargas()

        # Obtener los datos de esa fila
        categoria = firstSheet["A" + str(i + 1)].value
        rfc = firstSheet["B" + str(i + 1)].value
        usuario = firstSheet["C" + str(i + 1)].value
        contrasenia = firstSheet["D" + str(i + 1)].value

        # Aqui veo si es que existe la carpeta del mes anterior de este rfc, si existe con los cuatro archivos entonces significa que
        # ya no tiene sentido ir a descargarlos otra vez
        ruta = f"C:\\Users\\EMA Y EBA\\Desktop\\EMA Y EBA\\DESCARGAS\\{rfc}\\{str(datetime.now().year)}\\{str(getLastMonth())}"
        
        if(os.path.isdir(ruta)):
            files = os.listdir(ruta)
            if(len(files) > 0):
                continue


        # Borrar
        # borrarCarpeta(ruta)

        # El try es porque puede que no exista esa ruta
        try:
            files = os.listdir(ruta)
            
        except:
            os.makedirs(ruta)
            files = os.listdir(ruta)


        # Aqui se obtiene la carpeta donde se supone que debe ir lo que descargo
        planeacionFolders = os.listdir("\\\\192.168.1.77\\planeación fiscal mx")
        planeacionFolder = [f for f in planeacionFolders if f.split("_")[0] == rfc][0]
        descargasFolder = "\\\\192.168.1.77\\planeación fiscal mx\\" + planeacionFolder + "\\EMISIONES\\" + str(datetime.now().year) + "\\" + str(getLastMonth())
        
        # Borrar
        borrarCarpeta(descargasFolder)

        # Ahora se viene la parte donde vamos a descargar esos archivos
        # Si en el excel aparece como FIEL, significa que no tiene certificados .pfx, entonces se va a loguear con la fiel
        # Intento de login
        if (categoria == "FIEL"):
            driver = loginConFiel(rfc, usuario, contrasenia)
            if(len(driver.find_elements(By.CLASS_NAME, "alert-danger")) > 0):
                continue
        elif (categoria == "IMSS"):
            driver = loginConIMSS(rfc, usuario, contrasenia)
            if(len(driver.find_elements(By.CLASS_NAME, "alert-danger")) > 0):
                continue

        # Puede ser que me salga "Renovación de certificado"
        time.sleep(3)
        try:
            elem = driver.find_elements(By.XPATH, "//*[text()='Renovación de certificado']")
            if(len(elem) > 0):
                driver.close()
                continue
        except:
            None

        # Tratar de llegar a los links de descarga
        resultado = goToDescargas(driver)

        # Si llega a fallar, que continúe con el que sigue para que no se pierda tiempo
        try:
            if (resultado['Tipo'] == "Error"): continue
        except:
            None

        try:
            driver.close()
        except: None

        # Método para acomodar lso archivos descargados en tiempo de ejecución
        guardar(rfc, "C:\\Users\\EMA Y EBA\\Downloads")

main()