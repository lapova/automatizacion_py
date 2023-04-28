import unittest
from time import sleep
import os
# from tkinter import messagebox
from pyunitreport import HTMLTestRunner
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium import webdriver

#Hola Laura uwuwuwuu

class ExtractInfoCosmo(unittest.TestCase):
    #Abre navegador para ejecutar casos de prueba
    def setUp(self):
        
        #Ruta para descargar archivo
        chromeOptions = Options()
        chromeOptions.add_experimental_option("prefs", {
        "download.default_directory" : "C:\\Users\\Temp Tech\\OneDrive - CESDE\\Desktop\\sisi"
        })
        
        archivos = os.listdir('C:\\Users\\Temp Tech\\OneDrive - CESDE\\Desktop\\sisi')

        for archivo in archivos:
            archivo = archivo.lower()
            if 'preinscritos' in archivo:
                os.remove('C:\\Users\\Temp Tech\\OneDrive - CESDE\\Desktop\\sisi\\' + archivo)

        #Abrir navegador
        self.driver = webdriver.Chrome(executable_path = r'./chromedriver.exe', chrome_options=chromeOptions)
        driver = self.driver
        driver.maximize_window()
        #Navegar a Idaptive
        driver.get('https://comfama.my.idaptive.app/my?customerId=AAZ0175#/MyApps')

        sleep(3)

        #Inicio sesión en Idaptive
        # Encontrar el campo de usuario y enviar credenciales
        try: 
            username_field = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[1]/span/div/img[1]')
        except:
            # messagebox.showwarning("Idaptive", "Inicia la sesión")
            # username_field = WebDriverWait(driver, 300).until(
            #     EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[1]/span/div/img[1]'))
            # )
            username_field = driver.find_element(By.NAME, 'username')
            username_field.send_keys('LauraPoV')
        
            boton_username = driver.find_element(By.XPATH, '//*[@id="usernameForm"]/div[2]/button')
            boton_username.click()

        sleep(2)
   
        # Encontrar el campo de contraseña y enviar credenciales
        password_field = driver.find_element(By.NAME, 'answer')
        password_field.send_keys('Cesde2022*')
        
        boton_password = driver.find_element(By.XPATH, '//*[@id="passwordForm"]/div[3]/button')
        boton_password.click()

        # Presionar la tecla Enter para enviar el formulario
        password_field.send_keys(Keys.RETURN)

        username_field = WebDriverWait(driver, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[1]/span/div/img[1]'))
        )
        sleep(2)

        #Abrir una nueva pestaña
        driver.execute_script("window.open('');")

        #Cambiar a la nueva pestaña
        driver.switch_to.window(driver.window_handles[-1])

        #Pegar URL de q10 en nueva pestaña
        driver.get('https://comfama.my.idaptive.app/uprest/HandleAppClick?appkey=378c4e08-3645-440f-b5f9-78eb109ebb28&antixss=unLu92acHl4SZAfe7f6D6w__&cbecurversion=23.3.2')

        sleep(2)

    #Extraer informe de preinscritos
    def test_informe(self):
        driver = self.driver
        informe = driver.find_element(By.XPATH, '/html/body/nav/div[2]/ul[1]/li[4]/a')
        driver.execute_script("arguments[0].click();", informe)

        sleep(2)
        
        preinscritos = driver.find_element(By.XPATH, '/html/body/div[5]/div/div[1]/div[2]/div[1]/div[2]/div[8]/div[2]/div[2]/div[2]/a')
        driver.execute_script("arguments[0].click();", preinscritos)

        sleep(2)

        fecha = driver.find_element(By.XPATH, '/html/body/div[5]/div/div[1]/div[2]/div[2]/div/div/div/form/div[1]/div[3]/div/a/div[1]')
        fecha.click()

        sleep(2)

        rango_personalizado = driver.find_element(By.XPATH, '/html/body/div[10]/div[3]/ul/li[7]')
        rango_personalizado.click()

        sleep(2)

        mes = driver.find_element(By.XPATH, '/html/body/div[10]/div[2]/div/table/thead/tr[1]/th[2]/select[1]')
        mes = Select(mes)
        mes.select_by_visible_text('Julio')

        sleep(3)

        anio = driver.find_element(By.XPATH, '/html/body/div[10]/div[2]/div/table/thead/tr[1]/th[2]/select[2]')
        anio = Select(anio)
        anio.select_by_visible_text('2022')

        dia = driver.find_element(By.XPATH, '/html/body/div[10]/div[2]/div/table/tbody/tr[4]/td[3]')
        dia.click()

        boton_aceptar = driver.find_element(By.XPATH, '/html/body/div[10]/div[3]/div/button[1]')
        boton_aceptar.click()

        sleep(3)

        boton_exportar = driver.find_element(By.XPATH, '/html/body/div[5]/div/div[1]/div[2]/div[2]/div/div/div/form/div[4]/button')
        boton_exportar.click()

        sleep(8)

        archivos = os.listdir('C:\\Users\\Temp Tech\\OneDrive - CESDE\\Desktop\\sisi')

        for archivo in archivos:
            if '.xlsx' in archivo:
                if archivo.lower() != archivo:
                    os.rename('C:\\Users\\Temp Tech\\OneDrive - CESDE\\Desktop\\sisi\\' + archivo, 'C:\\Users\\Temp Tech\\OneDrive - CESDE\\Desktop\\sisi\\' + archivo.lower())

        os.chdir('C:\\Users\\Temp Tech\\OneDrive - CESDE\\Desktop\\sisi')
        os.system('python cosmo.py')
 
    #Cerrar navegador al terminar casos de prueba
    def tearDown(self):
        self.driver.quit()

if __name__ == "__main__":
    unittest.main(verbosity = 2, testRunner = HTMLTestRunner(output = 'reportes', report_name = 'informes-q10-report'))


