# Libreria para tiempo de espera
from time import sleep

# Librerias Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys

# Libreria para hora actual de la prueba
import datetime

# Libreria para escribir en excel
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

# Libreria para hora actual de la prueba
import datetime


book = openpyxl.load_workbook('Caidas.xlsx')
sheet = book.active


a = 1
Pruebas_Realizar = int(input("Pruebas que se desean realizar: "))
Espera_Entre_Pruebas = int(input("Tiempo que se desea esperar entre pruebas: "))
Pruebas_Exitosas = 0
Pruebas_Fallidas = 0

Comillas = '"'

while a <= Pruebas_Realizar:

    print("-------------------------------------------------------------------------------------------------------"
          "------------------------------------------------------------------------------------------------------------"
          "---------------------------------------------------------------------------------------------------------")

    print(f'Prueba #{a}')

    # Obteniendo hora de la prueba
    hora = datetime.datetime.now()

    print(hora)



    opts = Options()
    opts.add_argument("user-agent= Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Ubuntu Chromium/71.0.3578.80 Chrome/71.0.3578.80 Safari/537.36")
    driver = webdriver.Chrome('./chromedriver.exe', options=opts)


    driver.set_page_load_timeout(30)

    try:
        driver.get('https://ods.org.hn/')
        demanda_diurna = driver.find_element(By.XPATH,'/html/body/div[1]/div/section[2]/div/div/div[3]'
                                                      '/div/div/div[2]/div/div[1]/div[1]/div/h4').text

        if demanda_diurna == "Demanda Maxima Diurna":
            Pruebas_Exitosas += 1
            print("Carga exitosa")
            sheet[f'A{a+1}'] = f'Prueba #{a}'
            sheet[f'B{a+1}'] = "Exitosa"
            sheet[f'C{a+1}'] = hora
            book.save('Caidas.xlsx')

        else:
            Pruebas_Fallidas += 1
            print("Carga fallida")
            sheet[f'A{a+1}'] = f'Prueba #{a}'
            sheet[f'B{a+1}'] = "Fallida"
            sheet[f'C{a+1}'] = hora
            book.save('Caidas.xlsx')


    except BaseException:
        Pruebas_Fallidas += 1
        print("Carga fallida")
        sheet[f'A{a + 1}'] = f'Prueba #{a}'
        sheet[f'B{a + 1}'] = "Fallida"
        sheet[f'C{a + 1}'] = hora
        book.save('Caidas.xlsx')


    # Sumandole uno al contador para que el While le queden menos ciclos
    a += 1

    # Cerrar el navegador
    driver.close()
    sleep(Espera_Entre_Pruebas)


# Colocando el total de pruebas exitosas y fallidas
sheet[f'A{a + 2}'] = "Exitosas"
sheet[f'A{a + 3}'] = "Fallidas"

sheet[f'B{a + 2}'] = Pruebas_Exitosas
sheet[f'B{a + 3}'] = Pruebas_Fallidas
book.save('Caidas.xlsx')


print("-------------------------------------------------------------------------------------------------------"
          "------------------------------------------------------------------------------------------------------------"
          "---------------------------------------------------------------------------------------------------------")

print(f'| Total de pruebas realizas: {a - 1} | Total de pruebas exitosas: {Pruebas_Exitosas} | '
          f'Total de pruebas fallidas: {Pruebas_Fallidas} |')
