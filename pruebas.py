import openpyxl
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
import time

# Abre el archivo Excel y selecciona la hoja
workbook = openpyxl.load_workbook('registros.xlsx')
sheet = workbook['Sheet']

# Inicializa el controlador del navegador (asegúrate de tener el controlador correspondiente instalado)
driver = webdriver.Chrome()
driver.maximize_window()
driver.get('https://localhost:7244/usuarios/registrar') 

for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
    
    curp = row[0].value
    nombre = row[1].value
    primer_apellido = row[2].value
    segundo_apellido = row[3].value
    fecha_nacimiento = row[4].value
    genero = row[5].value
    pais = row[6].value  
    tipo_documento = row[7].value
    codigo_postal = row[8].value
    colonia = row[9].value
    calle = row[10].value
    numero_exterior = row[11].value
    entre_calles = row[12].value
    correo_electronico = row[13].value
    confirmar_correo = row[14].value
    contrasena = row[15].value
    lada = row[16].value
    numero_celular = row[17].value


if curp is None:
    time.sleep(7)
    checkbox = driver.find_element(By.CSS_SELECTOR, "#form > div:nth-child(2) > div.w-100.w-sm-75.mx-auto > div.steps > div.step-1 > div > div:nth-child(4) > div > div > div.col-sm-4.offset-0.offset-sm-1.text-center > p:nth-child(2) > span > span.k-switch-track.k-rounded-full")
    checkbox.click()
    # Completa los campos del formulario
    driver.find_element(By.CSS_SELECTOR, "#nombre").send_keys(nombre)
    driver.find_element(By.CSS_SELECTOR, "#primerApellido").send_keys(primer_apellido)
    driver.find_element(By.CSS_SELECTOR, "#segundoApellido").send_keys(segundo_apellido)
    driver.find_element(By.CSS_SELECTOR, "#fechaNacimiento").send_keys(fecha_nacimiento)

    # Selecciona el género en un menú desplegable
    genero_dropdown = Select(driver.find_element(By.CSS_SELECTOR, "#genero"))
    genero_dropdown.select_by_index(genero)

    pais_dropdown = Select(driver.find_element(By.CSS_SELECTOR, "#paisId"))
    pais_dropdown.select_by_index(pais)

    tipo_documento_dropdown = Select(driver.find_element(By.CSS_SELECTOR, "#tipoDocumento"))
    tipo_documento_dropdown.select_by_index(tipo_documento)

    

    cp = driver.find_element(By.CSS_SELECTOR, "#codigoPostal").send_keys(codigo_postal)
    div = driver.find_element(By.CSS_SELECTOR,"#infoDomicilio > div > div")
    time.sleep(3)
    div.click()
    time.sleep(15)
    driver.find_element(By.CSS_SELECTOR, "#SelectColoniaId > div").click()
    time.sleep(3)
    driver.find_element(By.CSS_SELECTOR, "#calle").send_keys(calle)
    driver.find_element(By.CSS_SELECTOR, "#numeroExterior").send_keys(numero_exterior)
    driver.find_element(By.CSS_SELECTOR, "#entreCalles").send_keys(entre_calles)
    driver.find_element(By.CSS_SELECTOR, "#siguienteBtn").click()
    time.sleep(2)

    # Completa los campos de correo electrónico y número de celular
    driver.find_element(By.CSS_SELECTOR, "#correo").send_keys(correo_electronico)
    driver.find_element(By.CSS_SELECTOR, "#confirmarCorreo").send_keys(confirmar_correo)

    # Encuentra el elemento select por su nombre
    lada_select = Select(driver.find_element(By.CSS_SELECTOR, "#lada"))
    lada_select.select_by_index(lada)

    driver.find_element(By.CSS_SELECTOR, "#celular").send_keys(numero_celular)
    time.sleep(3)
    driver.find_element(By.CSS_SELECTOR, "#siguienteBtn").click()
    time.sleep(2)
    # Agrega los campos de Contraseña y Confirmar Contraseña
    driver.find_element(By.CSS_SELECTOR, "#contrasena").send_keys(contrasena)
    driver.find_element(By.CSS_SELECTOR, "#confirmarContrasena").send_keys(contrasena)
    time.sleep(5)

    # Marca el checkbox "Acepto Términos y Condiciones"
    driver.find_element(By.CSS_SELECTOR, '#form > div:nth-child(2) > div.w-100.w-sm-75.mx-auto > div.steps > div.step-3 > div > div.col-sm-12 > div > p > span > span.k-switch-thumb-wrap > span').click()
    time.sleep(2)
    # Envía el formulario
    driver.find_element(By.CSS_SELECTOR, "#finalizarBtn").click()
    time.sleep(15)
    driver.find_element(By.CSS_SELECTOR, "body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled").click()
    time.sleep(7)
    driver.find_element(By.CSS_SELECTOR, "#loginForm > div > div > div:nth-child(1) > div > div.col-sm-9 > div > div.col-sm-12 > p > a").click()

    # Espera a que la página se cargue antes de continuar al siguiente registro
    driver.implicitly_wait(10)  # Espera hasta 10 segundos para la próxima operación
else:
    driver.find_element(By.CSS_SELECTOR, "#curp").send_keys(curp)
    #  driver.find_element(By.XPATH,"/html/body/div[2]/div[3]/div[1]/div/div/span").click()
    time.sleep(5)
    driver.find_element(By.CSS_SELECTOR, "#buscarBtn").click()
    time.sleep(15)
    cp = driver.find_element(By.CSS_SELECTOR, "#codigoPostal").send_keys(codigo_postal)
    div = driver.find_element(By.CSS_SELECTOR,"#infoDomicilio > div > div")
    time.sleep(3)
    div.click()
    time.sleep(15)
    driver.find_element(By.CSS_SELECTOR, "#SelectColoniaId > div").click()
    time.sleep(3)
    driver.find_element(By.CSS_SELECTOR, "#calle").send_keys(calle)
    driver.find_element(By.CSS_SELECTOR, "#numeroExterior").send_keys(numero_exterior)
    driver.find_element(By.CSS_SELECTOR, "#entreCalles").send_keys(entre_calles)
    driver.find_element(By.CSS_SELECTOR, "#siguienteBtn").click()
    time.sleep(2)

    # Completa los campos de correo electrónico y número de celular
    driver.find_element(By.CSS_SELECTOR, "#correo").send_keys(correo_electronico)
    driver.find_element(By.CSS_SELECTOR, "#confirmarCorreo").send_keys(confirmar_correo)

    # Encuentra el elemento select por su nombre
    lada_select = Select(driver.find_element(By.CSS_SELECTOR, "#lada"))
    lada_select.select_by_index(lada)

    driver.find_element(By.CSS_SELECTOR, "#celular").send_keys(numero_celular)
    time.sleep(3)
    driver.find_element(By.CSS_SELECTOR, "#siguienteBtn").click()
    time.sleep(2)
    # Agrega los campos de Contraseña y Confirmar Contraseña
    driver.find_element(By.CSS_SELECTOR, "#contrasena").send_keys(contrasena)
    driver.find_element(By.CSS_SELECTOR, "#confirmarContrasena").send_keys(contrasena)
    time.sleep(5)

    # Marca el checkbox "Acepto Términos y Condiciones"
    driver.find_element(By.CSS_SELECTOR, '#form > div:nth-child(2) > div.w-100.w-sm-75.mx-auto > div.steps > div.step-3 > div > div.col-sm-12 > div > p > span > span.k-switch-thumb-wrap > span').click()
    time.sleep(2)
    # Envía el formulario
    driver.find_element(By.CSS_SELECTOR, "#finalizarBtn").click()
    time.sleep(15)
    driver.find_element(By.CSS_SELECTOR, "body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled").click()
    time.sleep(7)
    driver.find_element(By.CSS_SELECTOR, "#loginForm > div > div > div:nth-child(1) > div > div.col-sm-9 > div > div.col-sm-12 > p > a").click()

    # Espera a que la página se cargue antes de continuar al siguiente registro
    driver.implicitly_wait(10)

# Cierra el navegador cuando hayas terminado
    driver.quit()
