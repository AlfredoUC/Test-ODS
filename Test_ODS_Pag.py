# Libreria para escribir en excel
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font


# Libreria para test
import unittest

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


# Normalizar todas la palabras a minusculas
def normalizar(s):
    remplazar = (
        ("á", "a"),
        ("é", "e"),
        ("í", "i"),
        ("ó", "o"),
        ("ú", "u"),
        ("A", "a"),
        ("B", "b"),
        ("C", "c"),
        ("D", "d"),
        ("E", "e"),
        ("F", "f"),
        ("H", "h"),
        ("G", "g"),
        ("I", "i"),
        ("J", "j"),
        ("K", "k"),
        ("L", "l"),
        ("M", "m"),
        ("N", "n"),
        ("O", "o"),
        ("P", "p"),
        ("Q", "q"),
        ("R", "r"),
        ("S", "s"),
        ("T", "t"),
        ("U", "u"),
        ("V", "v"),
        ("W", "w"),
        ("X", "x"),
        ("Y", "y"),
        ("Z", "z"),
        ("´", "")
    )
    for a, b in remplazar:
        s = s.replace(a, b).replace(a.upper(), b.upper())
    return s

class Mis_test(unittest.TestCase):
    def test_abrir_inicio_ODS(self):
        a = 1

        print(f'Prueba #{a}')

        # Obteniendo hora de la prueba
        hora = datetime.datetime.now()
        print(f'Hora y fecha de la prueba: {hora}')

        opts = Options()
        opts.add_argument("user-agent= Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Ubuntu Chromium/71.0.3578.80 Chrome/71.0.3578.80 Safari/537.36")
        driver = webdriver.Chrome('./chromedriver.exe', options=opts)

        # Tiempo maximo que se espera a que cargue la pagina antes de dar como fallida la carga

        driver.set_page_load_timeout(30)

        try:
            # Tiempo de espera para mientras carga de la pagina antes de mostrar error
            driver.get('https://ods.org.hn/')

            # Luego de cargar comprobar si el elemento o esta disponible
            demanda_diurna = driver.find_element(By.XPATH,
                                                 '/html/body/div[1]/div/section[2]/div/div/div[3]/div/div/div[2]/div/div[1]/div[1]/div/h4').text

            driver.close()

        except BaseException:
            demanda_diurna = "No se logro cargar la pagina"

        self.assertEqual(demanda_diurna, "Demanda Maxima Diurna")

    def test_busqueda_ODS_Titulo(self):
        a = 0
        titulo = None
        n_titulo = 0
        n_descripcion = 2
        hora = datetime.datetime.now()

        # Llamando al documento
        book = openpyxl.load_workbook('Titulos.xlsx')
        sheet = book.active


        # Palabra que se desea buscar
        buscar_palabra = normalizar(input("Coloque la palabra que desea buscar: "))
        print(f'Hora y fecha de la prueba: {hora}')


        # Requerimientos del navegador para evitar el BAN
        opts = Options()
        opts.add_argument("user-agent= Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Ubuntu Chromium/71.0.3578.80 Chrome/71.0.3578.80 Safari/537.36")
        driver = webdriver.Chrome('./chromedriver.exe', options=opts)

        driver.set_page_load_timeout(20)
        driver.get('https://ods.org.hn/index.php/documento2/centro-control-ods')

        # Buscar palabra
        campo_busqueda = driver.find_element(By.XPATH, '//*[@id="mod-search-searchword"]').send_keys(buscar_palabra,
                                                                                                     Keys.ENTER)

        # Mostrar toda la informacion relacionada con la palabra buscada
        lista_desplegable = driver.find_element(By.XPATH, '//*[@id="limit"]').send_keys("Todos", Keys.ENTER)

        # Busqueda exacta de la palabra en las etiquetas
        frase_exacta = driver.find_element(By.XPATH, '//*[@id="searchphraseexact-lbl"]').click()
        etiqueta = driver.find_element(By.XPATH, '//*[@id="searchForm"]/fieldset[2]/label[5]').click()
        btn_buscar = driver.find_element(By.XPATH, '//*[@id="searchForm"]/div[1]/div[2]/button').click()

        while 1 == 1:
            # Refrescando Hora
            hora = datetime.datetime.now()

            # Variables que debe aumentar
            a += 1
            n_titulo += 1

            # Separador
            print("--------------------------------------------------------------------------------------------------"
                  "-------------------------------------------------------------------------------------------------"
                  "-----------------------------------------------------------------------------------------------")



            try:
                titulo = driver.find_element(By.XPATH, f'//*[@id="sp-component"]/div/div[2]/dl/dt[{n_titulo}]/a').text
                titulo = normalizar(titulo)

                print(f'Titulo #{a}')
                print(f'Hora de la busqueda del titulo: {hora}')

            except BaseException:
                # Total encontrados
                print(f' SE ENCONTRO UN TOTAL DE: {n_titulo - 1} ARTICULOS')

                break
                
            sleep(0)

            if buscar_palabra in titulo:
                print("Contiene la palabra buscada")

                # Guardamos datos titulo correcto
                sheet[f'A{a + 1}'] = f'Titulo #{a}'
                sheet[f'B{a + 1}'] = titulo
                sheet[f'C{a + 1}'] = "Correcto"
                sheet[f'D{a + 1}'] = hora
                book.save('Titulos.xlsx')



            else:

                print("No contiene la palabra buscada")
                print(f"Error en el titulo #{n_titulo}")

                # Guardamos datos titulo fallidos
                sheet[f'A{a + 1}'] = f'Titulo #{a}'
                sheet[f'B{a + 1}'] = titulo
                sheet[f'C{a + 1}'] = "Fallido"
                sheet[f'D{a + 1}'] = hora
                book.save('Titulos.xlsx')
                self.assertRegex(titulo, buscar_palabra)
            n_descripcion += 3

        # Separador
        print("-------------------------------------------------------------------------------------------------------"
              "-------------------------------------------------------------------------------------------------------"
              "------------------------------------------------------------------------------------")



        self.assertRegex(titulo, buscar_palabra)

    def test_busqueda_ODS_Articulo(self):
        a = 0
        text_articulo = None
        n_articulo = 2
        n_descripcion = 2
        hora = datetime.datetime.now()

        # Llamando al documento
        book = openpyxl.load_workbook('Articulos.xlsx')
        sheet = book.active

        # Palabra que se desea buscar
        buscar_palabra = normalizar(input("Coloque la palabra que desea buscar: "))
        print(f'Hora y fecha de la prueba: {hora}')



        opts = Options()
        opts.add_argument("user-agent= Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Ubuntu Chromium/71.0.3578.80 Chrome/71.0.3578.80 Safari/537.36")
        driver = webdriver.Chrome('./chromedriver.exe', options=opts)

        driver.set_page_load_timeout(20)
        driver.get('https://ods.org.hn/index.php/documento2/centro-control-ods')

        # Buscar palabra
        campo_busqueda = driver.find_element(By.XPATH, '//*[@id="mod-search-searchword"]').send_keys(buscar_palabra,
                                                                                                     Keys.ENTER)

        # Mostrar toda la informacion relacionada con la palabra buscada
        lista_desplegable = driver.find_element(By.XPATH, '//*[@id="limit"]').send_keys("Todos", Keys.ENTER)

        # Busqueda exacta de la palabra en las etiquetas
        frase_exacta = driver.find_element(By.XPATH, '//*[@id="searchphraseexact-lbl"]').click()
        articulo = driver.find_element(By.XPATH, '//*[@id="searchForm"]/fieldset[2]/label[1]').click()
        btn_buscar = driver.find_element(By.XPATH, '//*[@id="searchForm"]/div[1]/div[2]/button').click()

        while 1 == 1:
            n_articulo += 3
            a += 1

            # Refrescando Hora
            hora = datetime.datetime.now()

            # Separador
            print("--------------------------------------------------------------------------------------------------"
                  "-------------------------------------------------------------------------------------------------"
                  "-----------------------------------------------------------------------------------------------")


            try:
                text_articulo = driver.find_element(By.XPATH, f'//*[@id="sp-component"]/div/div[2]/dl/dd[{n_articulo}]').text
                text_articulo = normalizar(text_articulo)

                print(f'Articulo #{a}')
                print(f'Hora de la busqueda: {hora}')
                print(text_articulo)

            except BaseException:
                print(f' SE ENCONTRO UN TOTAL DE: {n_articulo - 3} ARTICULOS')
                break
                
            # Subir o bajar velocidad
            sleep(0)
            
            
            if buscar_palabra in text_articulo:
                print("Contiene la palabra buscada")

                # Guardamos datos titulo correcto
                sheet[f'A{a + 1}'] = f'Articulo #{a}'
                sheet[f'B{a + 1}'] = text_articulo
                sheet[f'C{a + 1}'] = "Correcto"
                sheet[f'D{a + 1}'] = hora
                book.save('Articulos.xlsx')






            else:

                print("No contiene la palabra buscada")
                print(f"Error en la descripcion #{a}")

                # Guardamos datos titulo fallidos
                sheet[f'A{a + 1}'] = f'Articulo #{a}'
                sheet[f'B{a + 1}'] = text_articulo
                sheet[f'C{a + 1}'] = "Fallido"
                sheet[f'D{a + 1}'] = hora
                book.save('Articulos.xlsx')

                # Terminamos test como fallido
                self.assertRegex(text_articulo, buscar_palabra)


        self.assertRegex(text_articulo, buscar_palabra)













