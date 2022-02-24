# -*- coding: utf-8 -*-
import msvcrt
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


demora = 3
fnametemp = "temp_" + time.strftime("%d%m%Y%H%M%S") + ".xls"

navegador=input("Seleccione el Navegador:\n1. Firefox\n2. Chrome\n3.Phantom\nEscoja una opcion: ")
excelFile1=input("\nIngrese el archivo de Excel: ")
excelFile="C:\\Users\\RUBEN\\Desktop\\"+excelFile1+".xlsx"
inicioBusqueda=input("\n1.Inicio\n2.Final\nDonde comenzará la busqueda en la pagina: ")
filaExcel=input("\nEn que fila comenzará a introducir datos: ")



class extractor(object):
    def __init__(self):
        
        
        self.base_url = "https://procesos.ramajudicial.gov.co/procesoscs/"
        self.delay = 5
        if navegador == "1":
            self.driver = webdriver.Firefox()
        elif navegador == "2":
            #path_to_chromedriver = ('./chromedriver')
            #self.driver = webdriver.Chrome(executable_path=r"chromedriver.exe")
           # ser = Service("C:\\Users\\Guayara\\Desktop\\Extractor\\chromedriver.exe")
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
            self.driver= webdriver.Chrome("C:\\Users\\RUBEN\\Documents\\EXTRAERDIEGO\\chromedriver.exe", options=op)
            self.driver.get('https://something.com/login')
        elif navegador == "3":
            self.driver = webdriver.PhantomJS('phantomjs.exe')
        self.driver.wait = WebDriverWait(self.driver, self.delay)
        self.driver.set_window_size(1024, 768)
        self.load_page()

    def load_page(self):
        self.driver.get(self.base_url)
        #self.driver.find_element_by_link_text("https://procesos.ramajudicial.gov.co/procesoscs/").click()

        def page_loaded(driver):
            path = '//select[@id="ddlCiudad"]'
            return driver.find_element(By.XPATH, path)

        wait = WebDriverWait(self.driver, self.delay)
        try:
            wait.until(page_loaded)
        except TimeoutException:
            log_errors.write('line: 38 error: No se cargo la pagina, TimeoutException')

    def scrape_ciudad(self, laciudad):
        exito = True
        # dropdown = self.driver.find_element(By.XPATH,'//*[@id="ddlCiudad"]')
        # select_dropdown_option(self.driver, dropdown, laciudad)
        select = Select(self.driver.find_element(By.XPATH, '//*[@id="ddlCiudad"]'))
        select.select_by_visible_text(laciudad)
        wait = WebDriverWait(self.driver, self.delay)
        try:
            wait.until(lambda driver: driver.find_element(By.ID,'miVentana').is_displayed() == False)
        except TimeoutException:
            log_errors.write('line: 50 error: No se termino la carga del Ajax de Ciudad, TimeoutException')
            exito = False
        return exito

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
                        log_errors.write(
                            'line: 81 error: No esta el boton de cerrar la ventana de error para entidad, ElementNotVisibleException')
            else:
                #print("inactiva")
                datos.append("Inactiva")
                exito = False
        else:
            exito = False
            # guardar este error
        return exito

    def scrape_radicado(self, elradicado, datos, control):
        exito = True
        mns_error = ""
        try:
            self.driver.find_element(By.XPATH,"//input[@maxlength='23']").clear()
            self.driver.find_element(By.XPATH,"//input[@maxlength='23']").send_keys(elradicado)
        except ElementNotVisibleException:
            exito = False
            log_errors.write('line: 102 error: Element is not currently visible and so may not be interacted with, ElementNotVisibleException')
        
        #mover_slider(self)
        elema = self.driver.find_element(By.XPATH,"//div[@id='divNumRadicacion']/table/tbody/tr[3]/td/input")

        self.driver.execute_script('arguments[0].removeAttribute("disabled")', elema)
        #self.driver.find_element(By.ID,"btnConsultarNum").click()
        self.driver.find_element(By.XPATH,"//div[@id='divNumRadicacion']/table/tbody/tr[3]/td/input").click()
        wait = WebDriverWait(self.driver, 60)
        try:
            wait.until(lambda driver: driver.find_element(By.ID,'miVentana').is_displayed() == False)
        except TimeoutException:
            log_errors.write(
                'line: 98 error: No termino la carga del Ajax de Radicado: ' + elradicado + ', TimeoutException')
            exito = False
        if exists_by_xpath(self.driver, ".//*[@id='modalError' and contains(@style,'display: block')]"):
            exito = False
            mns_error = self.driver.find_element(By.XPATH,'//*[@id="msjError"]').text
            # si el error indica bloqueo comprobar
            if control == 1:
                #print(mns_error)
                datos.append(mns_error)
            try:
                self.driver.find_element(By.XPATH,'//*[@id="modalError"]/div/table/tbody/tr/td/input').click()
            except ElementNotVisibleException:
                log_errors.write(
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
                #table_actos = self.driver.find_element(By.XPATH,
                #    ".//*[@id='divActuaciones']/div[3]/table[2]/tbody/tr[2]/td/table[1]/tbody")
                table_actos = self.driver.find_element(By.XPATH,
                    "//*[@id='divActuacionesDetalle']/table/tbody/tr[2]/td/table/tbody")
                allrows = table_actos.find_elements(By.TAG_NAME,"tr")[1:]
                for tr in allrows:
                    lista_td = []
                    lista_td.append(i[2])
                    allcols = tr.find_elements(By.TAG_NAME,"td")
                    for j in range(len(allcols)):
                        lista_td.append(allcols[j].text)
                       # print(allcols[j].text)
                    actos.append(lista_td)
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
        #self.driver.find_element(By.ID,"btnNuevaConsultaNum").click()
        self.load_page()
        # try:
        #    self.driver.find_element_by_link_text("INICIO").click()
        #    #self.driver.find_element_by_link_text("https://procesos.ramajudicial.gov.co/procesoscs/").click()
        # except ElementNotVisibleException:
        #                 log_errors.write('line: 196 error: Hay un problema con la recarga al INICIO, ElementNotVisibleException')

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
    # slider = self.driver.find_element(By.XPATH,"//*[@id='sliderBehaviorNumeroProceso_railElement']")
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
        #Si la busqueda comienza desde abajo(final)
        if inicioBusqueda == "2":
            fecha_str = dats[6]
            if fecha_str!="--" and dife_fecha(fecha_str).days <= 4:
                # comprobar el numero de dias para poder guardar el dato
                wb.sheet_by_name('ACTUACIONES DEL PROCESO').row += dats
        else:
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

if __name__ == "__main__":
    # revisar que se cumplan los parametros
   # if len(sys.argv) < 5:
    #    print("\tUSO: python extractor.py [nombre_de_archivo.xlsx] inicial  'X' 'y' -->"
     #   "Ejecuta el programa sin condicional de 4 dias\n\t      python extractor.py [nombre_de_archivo.xlsx] final 'X' 'y' -->"
      # "Ejecuta el programa con condicional de 4 dias")
       # print("\nReemplazar 'X' por F --> Firefox, C --> Chrome o P --> PhantomJS")
        #print("\nReemplazar 'y' por el número de renglón del archivo de entrada para comenzar")
        #sys.exit(1)
    # revisar la existencia del archivo de datos .xlsx
    
    if not (File_Existence(excelFile)):
        print("Revisar la existencia del archivo de datos: " + excelFile)
        msvcrt.getch()
        sys.exit(1)
    if (int(filaExcel) <= 0):
        fila_inicio = 1
    else:
        fila_inicio = int(filaExcel)
    if (inicioBusqueda == "1") or (inicioBusqueda == "2"):
        # si los parametros corresponden entonces proceder con el programa
        pathlog = './'
        log_errors = open(pathlog + 'log_errors.txt', mode='w')
        w = extractor()
        #print('> ' + str(datetime.today()))
        print('Ejecutando ...')

        t0 = time.time()
        # extraer los datos del archivo de entrada
        my_array = pyexcel.get_array(file_name=excelFile, start_row=fila_inicio)
        # crear el archivo de output.xlsx
        wout = xlwt.Workbook()
        crear_xls(wout)
        num_rad = 1
        for i in my_array:
            #print(num_rad, " --> ", i)
            datosPro = []
            actosPro = []
            # print("Ciudad: ", i[0]," Entidad: ", i[1], " Radicado: ", i[2])
            if w.scrape_ciudad(i[0]):
                if w.scrape_entidad(i[1], i[2], i):
                    if w.scrape_radicado(i[2], i, 1):
                        #print("listo")
                        # extraer los datos y actuaciones
                        w.extraer_datos_actuaciones(datosPro, actosPro)
                    else:
                        #w.reload_page()
                        w.load_page()
                        if w.scrape_ciudad(i[0]):
                            if w.scrape_entidad(i[1], i[2], i):
                                if w.scrape_radicado(i[2], i, 2):
                                    #print("listo 2")
                                    # extraer los datos y actuaciones
                                    w.extraer_datos_actuaciones(datosPro, actosPro)
            # escribir los resultados en el archivo
            escribir_xls(datosPro, actosPro)
            num_rad += 1
        w.final()
        terminar()
        t1 = time.time() - t0
        print("\ntiempo transcurrido %.2f s" % (t1))
    else:
        print("Usar\n1.Inicio = para correr sin condicional de dias\n2.Final para condicional de dias")
        msvcrt.getch()
        sys.exit(1)
