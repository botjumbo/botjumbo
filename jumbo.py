import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoSuchWindowException
from datetime import datetime
import openpyxl
import xlwings as xw
import re
import time
fecha_actual = datetime.now().strftime('%d-%m-%Y')


urls = [
    "https://www.jumbo.com.ar/pan-de-viena-la-panerita-x-6-un/p",
    "https://www.jumbo.com.ar/galletitas-minipolvorita-vainilla-frutilla-152g/p",
    "https://www.jumbo.com.ar/galletitas-de-agua-traviata-303-gr/p",
    "https://www.jumbo.com.ar/harina-integral-100pureza-1-kg/p",
    "https://www.jumbo.com.ar/almidon-de-maiz-sin-tacc-maizena-x-500-gr/p",
    "https://www.jumbo.com.ar/arroz-gallo-parboil-selec-x500g/p",
    "https://www.jumbo.com.ar/fideos-lucchetti-tallarin-n5-x500g/p",
    "https://www.jumbo.com.ar/asado-del-centro-2/p",
    "https://www.jumbo.com.ar/carnaza-comun-2/p",
    "https://www.jumbo.com.ar/osobuco-2/p",
    "https://www.jumbo.com.ar/paleta-churr-de-nov-envasado-al-vacio/p",
    "https://www.jumbo.com.ar/carne-vacuna-picada-enfriada-e/p",
    "https://www.jumbo.com.ar/milanesa-nalga-2/p",
    "https://www.jumbo.com.ar/higado-congelado/p",
    "https://www.jumbo.com.ar/pechito-de-cerdo-fresco-3/p",
    "https://www.jumbo.com.ar/pollo-con-menudos-congelado/p",
    "https://www.jumbo.com.ar/milanesas-de-merluza-cuisine-y-co-rebozadas-500-gr/p",
    "https://www.jumbo.com.ar/mortadela-paladini-familiar-500-gr/p",
    "https://www.jumbo.com.ar/bife-de-paleta-de-cerdo-fresco-2/p",
    "https://www.jumbo.com.ar/salchichon-calchaqui-2/p",
    "https://www.jumbo.com.ar/salame-milan-tripack-por-kg-minimo-800-gr-mayorista/p",
    "https://www.jumbo.com.ar/aceite-canuelas-de-girasol-1-5-l/p",
    "https://www.jumbo.com.ar/manteca-ls-bienestar-animal-200-g/p",
    "https://www.jumbo.com.ar/leche-la-serenisima-liviana-bot-1l/p",
    "https://www.jumbo.com.ar/leche-en-polvo-nutrifuerza-la-lechera-800-gr/p",
    "https://www.jumbo.com.ar/queso-crema-clasico-tregar-280g/p",
    "https://www.jumbo.com.ar/queso-cremoso-punta-del-agua-horma-x-kg-2/p",
    "https://www.jumbo.com.ar/queso-reggianito-rallado-la-serenisima-175gr/p",
    "https://www.jumbo.com.ar/manteca-primer-premio-2/p",
    "https://www.jumbo.com.ar/yogur-entero-la-serenisima-clasico-frutilla-900g/p",
    "https://www.jumbo.com.ar/dulce-de-leche-la-serenisima-colonial-400-g/p",
    "https://www.jumbo.com.ar/huevos-blancos-maxima-30-u/p",
    "https://www.jumbo.com.ar/manzana-royal-gala-por-kg/p",
    "https://www.jumbo.com.ar/mandarina-okitsu-por-kg/p",
    "https://www.jumbo.com.ar/naranja-jugo-especial-por-kg/p",
    "https://www.jumbo.com.ar/banana-ecuador-por-kg/p",
    "https://www.jumbo.com.ar/pera-por-kg/p",
    "https://www.jumbo.com.ar/batata-por-kg/p",
    "https://www.jumbo.com.ar/papa-fraccionada-por-kg/p",
    "https://www.jumbo.com.ar/acelga-green-life-550-gr/p",
    "https://www.jumbo.com.ar/cebolla-superior-por-kg/p",
    "https://www.jumbo.com.ar/choclo-x-unidad/p",
    "https://www.jumbo.com.ar/lechuga-capuchina-por-kg/p",
    "https://www.jumbo.com.ar/tomate-redondo-grande-por-kg/p",
    "https://www.jumbo.com.ar/zanahoria-organica-fraccionada/p",
    "https://www.jumbo.com.ar/zapallo-coreano-por-kg/p",
    "https://www.jumbo.com.ar/tomate-perita-en-lata-arcor-400-gr/p",
    "https://www.jumbo.com.ar/arvejas-inalpa-secas-remojadas-x300gr-2/p",
    "https://www.jumbo.com.ar/lentejas-cuisine-co-300-gr/p",
    "https://www.jumbo.com.ar/azucar-chango-1-kg/p",
    "https://www.jumbo.com.ar/dulce-de-batata-el-guri-bar-1-kg/p",
    "https://www.jumbo.com.ar/mermelada-light-durazno-alco-390-gr/p",
    "https://www.jumbo.com.ar/sal-fina-dos-anclas-500-gr-3/p",
    "https://www.jumbo.com.ar/mayonesa-clasica-hellmanns-237-gr/p",
    "https://www.jumbo.com.ar/vinagre-de-manzana-menoyo-500-ml/p",
    "https://www.jumbo.com.ar/caldo-knorr-de-verduras-6-cubos-2/p",
    "https://www.jumbo.com.ar/gaseosa-coca-cola-sabor-original-2-25-l/p",
    "https://www.jumbo.com.ar/jugo-en-polvo-clight-naranja-dulce-7-5-g/p",
    "https://www.jumbo.com.ar/soda-sifon-saldan-2-l/p",
    "https://www.jumbo.com.ar/cerveza-quilmes-clasica-retornable-1-l/p",
    "https://www.jumbo.com.ar/vino-tinto-vinas-de-balbo-borgona-1-125-cc/p",
    "https://www.jumbo.com.ar/cafe-dolca-suave-nescafe-170-gr-3/p",
    "https://www.jumbo.com.ar/yerba-mate-suave-playadito-500-gr/p",
    "https://www.jumbo.com.ar/te-naturalidad-intacta-la-virginia-50-saquitos/p"
]

titulos_precios = []
workbook = openpyxl.load_workbook("DatosJumbo.xlsx")
sheet = workbook["Diaria"]

fila_vacia = 1
columna = 2 
fila_titulos = 1
columna_titulos = 2

# Especifica la ruta del perfil de Chrome personalizado
profile_directory = 'C:\\Users\\"TuUsuario\\AppData\\Local\\Google\\Chrome\\User Data\\'

# Configurar el driver de Selenium con el perfil personalizado
options = webdriver.ChromeOptions()
options.add_argument(f"user-data-dir={profile_directory}")
service = Service()
driver = webdriver.Chrome(service=service, options=options)

# Iterar sobre las URLs
for url in urls:
    try:
    #Averiguar si hay stock
        driver.get(url)
        time.sleep(5)
        titulo = driver.find_element(By.XPATH, "//h1[@class='vtex-store-components-3-x-productNameContainer mv0 t-heading-4']/span[@class='vtex-store-components-3-x-productBrand ']").text.strip()
        stock  = driver.find_element(By.XPATH, "//div[contains(@class, 'vtex-flex-layout-0-x-flexColChild')]//p[contains(@class, 'vtex-outOfStockFlag__text')]").text.strip() 
        if stock == "Producto sin stock":
            titulos_precios.append((titulo,0)) #si no hay stock agrego el precio como 0
            print(titulo,0)
    except NoSuchElementException:
        try:
            # Obtener el titulo
            titulo = driver.find_element(By.XPATH, "//h1[@class='vtex-store-components-3-x-productNameContainer mv0 t-heading-4']/span[@class='vtex-store-components-3-x-productBrand ']").text.strip()
        except NoSuchElementException:
            print("No se encontro el Titulo en la página")
            titulos_precios.append(("Producto X", 0))
            continue
        try:
            #Obtener el precio
            #precio de lista
            precios = driver.find_elements(By.XPATH, "//div[contains(@class, 'vtex-flex-layout-0-x-flexColChild--separator') and .//div[contains(text(), 'espacio')]]//span[contains(@class, 'jumboargentinaio-store-theme-1QiyQadHj-1_x9js9EXUYK')]")
            #precio con descuentos
            #precios = driver.find_elements(By.XPATH, "//div[contains(@class, 'vtex-flex-layout-0-x-flexColChild--separator') and .//div[contains(text(), 'espacio')]]//div[contains(@class, 'jumboargentinaio-store-theme-1dCOMij_MzTzZOCohX1K7w')]")

            for precio in precios:
                precio_texto = precio.get_attribute("innerText")

        except (NoSuchElementException,NoSuchWindowException) as e:
                
                titulos_precios.append((titulo,0))
                print(titulo, 0)
                continue

        if precio_texto:
            #Si encuentra precios por el primer XPATH de precios
            try:

                precio = re.search(r'(?<=\$)[\d,.]+', precio_texto)
                precio = precio.group().replace('.','').replace(',','.')
                precio = float(precio)

        
                print(titulo,precio)
                titulos_precios.append((titulo,precio))
                continue
            except Exception as e:
                titulos_precios.append((titulo,0))
                print(titulo, 0)
                continue




driver.quit()



while sheet[f"A{fila_vacia}"].value is not None:
    fila_vacia += 1

#agregar fecha
sheet[f"A{fila_vacia}"] = fecha_actual


for _, precio in titulos_precios:
    sheet.cell(row=fila_vacia, column=columna).value = precio
    columna += 1  

for  titulo, _ in titulos_precios:

    sheet.cell(row=fila_titulos, column=columna_titulos).value = titulo
    columna_titulos += 1  



# Guardar el libro de trabajo datos
workbook.save("DatosJumbo.xlsx") 
############################



#https://www.jumbo.com.ar/queso-crema-la-paulina-alioli-250g/p
#https://www.jumbo.com.ar/margarina-vegetal-danica-200g/p
#https://www.jumbo.com.ar/vino-toro-clasico-tinto-1125cc/p
#https://www.jumbo.com.ar/osobuco-de-novillo/p
