# -*- coding: utf-8 -*-
import msvcrt
from re import L
import shutil
import sys
import time
import os
from datetime import datetime
import pyexcel
import xlwt
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, ElementNotVisibleException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager


demora = 3
#Nombre asignado al excel temporal en ejecucion
fnametemp = "temp_" + time.strftime("%d%m%Y%H%M%S") + ".xls"

#Datos de entrada para el programa

#Escoger los archivos con los que se trabajara
inputFile=input("\n\nIngrese archivo(s) de Excel separado por comas: ")
listFile=listaArchivos=inputFile.split(",")
os.system ("cls")
#Escoger si queremos buscar todas las actuaciones (1) o solamente en los ultimos 4 dias (2)
inicioBusqueda=input("\n\n1.Inicio\n2.Final\nDonde comenzará la busqueda en la pagina: ")
os.system ("cls")
#Escoger desde que linea se quiere comenzar a escribir en el excel
filaExcel=input("\n\nEn que fila comenzará a introducir datos: ")
os.system ("cls")



class extractor(object):
    def __init__(self):



        #Escoger Chrome como Navegador
            op = webdriver.ChromeOptions()
            op.add_argument("--headless")
            op.add_argument("--disable-gpu")
            op.add_argument("--ignore-certificate-errors")
            op.add_argument(" --disable-extensions ")
            op.add_argument(" - -no-sandbox ")
            op.add_argument(" --disable-dev-shm-usage ")
            op.add_experimental_option('excludeSwitches', ['enable-logging'])
            prefs = {"profile.managed_default_content_settings.images":2, 
            "profile.default_content_setting_values.notifications":2, 
            "profile.managed_default_content_settings.stylesheets":2, 
            "perfil.managed_default_content_settings.cookies":2, 
            "profile.managed_default_content_settings.javascript":1, 
            "profile.managed_default_content_settings.plugins":1, 
            "profile.managed_default_content_settings.popups":2, 
            "perfil.managed_default_content_settings.geolocation":2, 
            "perfil.managed_default_content_settings.media_stream":2, 
            } 
            op.add_experimental_option("prefs",prefs) 
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=op)
        
        #Escoger preferencias del WebDriver
            self.base_url = "https://procesos.ramajudicial.gov.co/procesoscs/"
            self.delay = 5
            self.driver.wait = WebDriverWait(self.driver, self.delay)
            self.driver.set_window_size(1024, 768)
            self.load_page()
    
    #Cargar la pagina solicitada
    def load_page(self):
        self.driver.get(self.base_url)

        def page_loaded(driver):
            path = '//select[@id="ddlCiudad"]'
            return driver.find_element(By.XPATH, path)

        wait = WebDriverWait(self.driver, self.delay)
        try:
            wait.until(page_loaded)
        except TimeoutException:
            print('line: 38 error: No se cargo la pagina, TimeoutException')

    #Encontrar TextBox de Ciudad
    def scrape_ciudad(self, laciudad):
        exito = True
        select = Select(self.driver.find_element(By.XPATH, '//*[@id="ddlCiudad"]'))
        select.select_by_visible_text(laciudad)
        wait = WebDriverWait(self.driver, self.delay)
        try:
            wait.until(lambda driver: driver.find_element(By.ID,'miVentana').is_displayed() == False)
        except TimeoutException:
            print('line: 50 error: No se termino la carga del Ajax de Ciudad, TimeoutException')
            exito = False
        return exito

    #Saber si la Entidad esta activa
    def entidad_activa(self, locator, entidad, entidad_txt):
        activa = True
        cod_enti = entidad[5:9]
        options = [x for x in locator.find_elements(By.TAG_NAME,'option')]
        for v in options:
            valorOp = v.get_attribute("value")
            if len(valorOp) > 1:
                ordStr = valorOp
                ordStr_a = ordStr.split("-")
                if ordStr_a[1] == "False" and ordStr_a[2] == cod_enti:
                    if entidad_txt in v.get_attribute("text"):
                        activa = False
                        break
        return activa

    #Encontrar TextBox de Entidad
    def scrape_entidad(self, laentidad, elradicado, datos):
        exito = True
        if WaitForElement(self, "//option[contains(.,'Seleccione la Corporación/Especialidad')]"):
            enti_drop = self.driver.find_element(By.CSS_SELECTOR,"select#ddlEntidadEspecialidad")
            # si la entidad esta activa
            if self.entidad_activa(enti_drop, elradicado, laentidad):
                select_dropdown_option_entidad(self.driver, enti_drop, laentidad)
                if self.driver.find_element(By.ID,"msjError").is_displayed():
                    # hay mensaje de error guardarlo y comparalo para superar el problema de ciudades
                    #print(self.driver.find_element(By.ID,"msjError").text)
                    exito = False
                    try:
                        self.driver.find_element(By.CSS_SELECTOR,
                            "div.inisideModal > table > tbody > tr > td > input[type=\"button\"]").click()
                    except ElementNotVisibleException:
                        print(
                            'line: 81 error: No esta el boton de cerrar la ventana de error para entidad, ElementNotVisibleException')
            else:
                #print("inactiva")
                datos.append("Inactiva")
                exito = False
        else:
            exito = False
            # guardar este error
        return exito

    #Encontrar TextBox de Radicado
    def scrape_radicado(self, elradicado, datos, control):
        exito = True
        mns_error = ""
        try:
            self.driver.find_element(By.XPATH,"//input[@maxlength='23']").clear()
            self.driver.find_element(By.XPATH,"//input[@maxlength='23']").send_keys(elradicado)
        except ElementNotVisibleException:
            exito = False
            print('line: 102 error: Element is not currently visible and so may not be interacted with, ElementNotVisibleException')
        
        #mover_slider(self)
        elema = self.driver.find_element(By.XPATH,"//div[@id='divNumRadicacion']/table/tbody/tr[3]/td/input")

        self.driver.execute_script('arguments[0].removeAttribute("disabled")', elema)
        self.driver.find_element(By.XPATH,"//div[@id='divNumRadicacion']/table/tbody/tr[3]/td/input").click()
        wait = WebDriverWait(self.driver, 60)
        try:
            wait.until(lambda driver: driver.find_element(By.ID,'miVentana').is_displayed() == False)
        except TimeoutException:
            print(
                'line: 98 error: No termino la carga del Ajax de Radicado: ' + elradicado + ', TimeoutException')
            exito = False
        if exists_by_xpath(self.driver, ".//*[@id='modalError' and contains(@style,'display: block')]"):
            exito = False
            mns_error = self.driver.find_element(By.XPATH,'//*[@id="msjError"]').text
            # si el error indica bloqueo comprobar
            if control == 1:
                datos.append(mns_error)
            try:
                self.driver.find_element(By.XPATH,'//*[@id="modalError"]/div/table/tbody/tr/td/input').click()
            except ElementNotVisibleException:
                print(
                    'line: 109 error: No se encuentra el boton de Cerrar de la ventana de error en radicado, ElementNotVisibleException')
        return exito

    def extraer_datos_actuaciones(self, datos, actos):

        if WaitForElement(self, "//*[@id='lblFechaSistema']"):
            fecharesult = self.driver.find_element(By.XPATH,"//*[@id='lblFechaSistema']").text
            if bool(fecharesult and fecharesult.strip()):
                datos.append(self.driver.find_element(By.XPATH,".//*[@id='lblFechaSistema']").text)
                datos.append(i[0])
                datos.append(i[1])
                datos.append(i[2])
                datos.append(self.driver.find_element(By.XPATH,".//*[@id='lblJuzgadoActual']").text)
                datos.append(self.driver.find_element(By.XPATH,".//*[@id='lblPonente']").text)
                datos.append(self.driver.find_element(By.XPATH,".//*[@id='lblTipo']").text)
                datos.append(self.driver.find_element(By.XPATH,".//*[@id='lblClase']").text)
                datos.append(self.driver.find_element(By.XPATH,".//*[@id='lblRecurso']").text)
                datos.append(self.driver.find_element(By.XPATH,".//*[@id='lblUbicacion']").text)
                datos.append(self.driver.find_element(By.XPATH,".//*[@id='lblNomDemandante']").text)
                datos.append(self.driver.find_element(By.XPATH,".//*[@id='lblNomDemandado']").text)
                datos.append(self.driver.find_element(By.XPATH,".//*[@id='lblContenido']").text)

                table_actos = self.driver.find_element(By.XPATH,
                    "//*[@id='divActuacionesDetalle']/table/tbody/tr[2]/td/table/tbody")
                
                #Si la busqueda comienza desde abajo(final)
                if inicioBusqueda == "2": 
     
                    
                        allrows = table_actos.find_elements(By.TAG_NAME,"tr")[1:7]

                        
                        for tr in allrows:
                                lista_td = []
                                lista_td.append(i[2])
                                allcols = tr.find_elements(By.TAG_NAME,"td")
                                fecha_str=allcols[0].text
                                print(fecha_str)
                                if fecha_str!="--" and dife_fecha(fecha_str).days <= 4:
                                    for j in range(len(allcols)):
                                        lista_td.append(allcols[j].text)
                                    actos.append(lista_td)
                else:

                    allrows = table_actos.find_elements(By.TAG_NAME,"tr")[1:]
                    for tr in allrows:
                        lista_td = []
                        lista_td.append(i[2])
                        allcols = tr.find_elements(By.TAG_NAME,"td")
                        for j in range(len(allcols)):
                            lista_td.append(allcols[j].text)
                            
                        actos.append(lista_td)
                        print(actos)



            else:
                # la fecha de resultados esta vacia o no aparecen los resultados
                datos.append("ERROR : No aparece resultado o fecha vacia")
                datos.append(i[0])
                datos.append(i[1])
                datos.append(i[2])
        else:
            # no aparece el localizador de fecha de resultado
            datos.append("ERROR : No se encuentra el localizador de fecha de resultado")
            datos.append(i[0])
            datos.append(i[1])
            datos.append(i[2])
        self.load_page()

    def final(self):
        self.driver.quit()

def WaitForElement(self, path):
    limit = demora
    inc = 1
    c = 0
    while c < limit:
        try:
            self.driver.find_element(By.XPATH,path)
            return 1
        except:
            time.sleep(inc)
            c+=inc
    return 0

def exists_by_xpath(driver, xpath):
    try:
        driver.find_element(By.XPATH,xpath)
    except NoSuchElementException:
        return False
    return True

def mover_slider(self):
    slider = self.driver.find_element(By.XPATH,"//*[@id='sliderBehaviorNumeroProceso_handleImage']")
    action = ActionChains(self.driver)
    # action.move_to_element_with_offset(slider,60,0)
    action.drag_and_drop_by_offset(slider, 60, 0)
    # action.click()
    action.perform()

def select_dropdown_option_entidad(self, select_locator, option_text):
    for option in select_locator.find_elements(By.TAG_NAME,'option'):
        if option_text in option.text:
            option.click()
            break

def crear_xls(wb):
    data = {'INPUT': ['CIUDAD', 'ENTIDAD/ESPECIALIDAD', 'RADICADO', 'EXITOSO'],
            'DATOS DEL PROCESO': ['FECHA CONSULTA', 'CIUDAD', 'ENTIDAD', 'RADICADO', 'DESPACHO', 'PONENTE', 'TIPO',
                                  'CLASE', 'RECURSO', 'UBICACION', 'DEMANDANTE(S)', 'CONTENIDO'],
            'ACTUACIONES DEL PROCESO': ['RADICADO', 'FECHA ACTUACION', 'ACTUACION', 'ANOTACION', 'FECHA INICIA TERMINO',
                                        'FECHA FIN TERMINO',
                                        'FECHA REGISTRO']}
    for key, nomHoja in enumerate(data):
        ws = wb.add_sheet(nomHoja)
        for clave, valor in enumerate(data[nomHoja]):
            ws.write(0, clave, valor)
    wb.save(fnametemp)

def dife_fecha(fecha):
    hoy = datetime.now()
    dia, mes, ano = fecha.split()
    if (mes == 'Jan') or (mes == 'Ene'):
        mes = '01'
    if (mes == 'Feb'):
        mes = '02'
    if (mes == 'Mar'):
        mes = '03'
    if (mes == 'Apr') or (mes == 'Abr'):
        mes = '04'
    if (mes == 'May'):
        mes = '05'
    if (mes == 'Jun'):
        mes = '06'
    if (mes == 'Jul'):
        mes = '07'
    if (mes == 'Aug') or (mes == 'Ago'):
        mes = '08'
    if (mes == 'Sep'):
        mes = '09'
    if (mes == 'Oct'):
        mes = '10'
    if (mes == 'Nov'):
        mes = '11'
    if (mes == 'Dic') or (mes == 'Dec'):
        mes = '12'
    fecha_str = dia + ' ' + mes + ' ' + ano
    dt_obj = datetime.strptime(fecha_str, '%d %m %Y')
    return (hoy - dt_obj)

def escribir_xls(datosPro, actosPro):
    wb = pyexcel.get_book(file_name=fnametemp)
    wb.sheet_by_name('INPUT').row += i
    wb.sheet_by_name('DATOS DEL PROCESO').row += datosPro
    for dats in actosPro:
        wb.sheet_by_name('ACTUACIONES DEL PROCESO').row += dats

    wb.save_as(fnametemp)

def File_Existence(filepath):
    try:
        f = open(filepath)
    except IOError:
        return False
    return True

def terminar():
    # cambiar el nombre del temp.xls y eliminarlo
    fname = excelFile.split(".")[0] + "_" + time.strftime("%d%m%Y%H%M%S") + ".xls"
    shutil.move(fnametemp, fname)
    print("Creado el archivo: " + fname)

#CLASE PRINCIPAL
if __name__ == "__main__":
    
    #Hacer esto con todos los archivos ingresados
    for file in listFile:
        excelFile="C:\\Users\\RUBEN\\Desktop\\"+file+".xlsx"
        # revisar la existencia del archivo de datos .xlsx
        if not (File_Existence(excelFile)):
            print("Revisar la existencia del archivo de datos: " + excelFile)
            msvcrt.getch()
            sys.exit(1)

        #Indicar desde que fila en el excel se comenzara a escribir
        if (int(filaExcel) <= 0):
            fila_inicio = 1
        else:
            fila_inicio = int(filaExcel)

        #Seleccionar si queremos buscar todas las actuaciones (1) o solamente en los ultimos 4 dias (2)
        if (inicioBusqueda == "1") or (inicioBusqueda == "2"):

            w = extractor()
            print('Ejecutando ...'+file)

            t0 = time.time()
            # extraer los datos del archivo de entrada
            my_array = pyexcel.get_array(file_name=excelFile, start_row=fila_inicio)
            # crear el archivo de output.xlsx
            wout = xlwt.Workbook()
            crear_xls(wout)
            num_rad = 1
            for i in my_array:
                datosPro = []
                actosPro = []
                # Ciudad: i[0] - Entidad: i[1] - Radicado: i[2]
                if w.scrape_ciudad(i[0]):
                    if w.scrape_entidad(i[1], i[2], i):
                        if w.scrape_radicado(i[2], i, 1):
                            # extraer los datos y actuaciones
                            w.extraer_datos_actuaciones(datosPro, actosPro)
                        else:
                            w.load_page()
                            if w.scrape_ciudad(i[0]):
                                if w.scrape_entidad(i[1], i[2], i):
                                    if w.scrape_radicado(i[2], i, 2):
                                        # extraer los datos y actuaciones
                                        print(i[2])
                                        print(i)
                                        w.extraer_datos_actuaciones(datosPro, actosPro)
                # escribir los resultados en el archivo
                escribir_xls(datosPro, actosPro)
                num_rad += 1
            w.final()
            terminar()
            t1 = time.time() - t0
            print("Finalizado:", file)
            print("\ntiempo transcurrido %.2f s" % (t1))
        else:
            print("Usar\n1.Inicio = para correr sin condicional de dias\n2.Final para condicional de dias")
            msvcrt.getch()
            sys.exit(1)
        print("\n----------------------\n")  
    print("Presiona una tecla para cerrar")
    msvcrt.getch()
    sys.exit(1)
