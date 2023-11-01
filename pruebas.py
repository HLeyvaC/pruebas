import openpyxl
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By

# Abre el archivo Excel y selecciona la hoja
workbook = openpyxl.load_workbook('registros.xlsx')
sheet = workbook['Sheet']

# Inicializa el controlador del navegador (asegúrate de tener el controlador correspondiente instalado)
driver = webdriver.Chrome()
driver.get('https://localhost:7244/usuarios/registrar')  # Reemplaza 'URL_DEL_FORMULARIO' con la URL real del formulario web

for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
    nombre = row[0].value
    primer_apellido = row[1].value
    segundo_apellido = row[2].value
    fecha_nacimiento = row[3].value
    genero = row[4].value
    tipo_documento = row[5].value
    codigo_postal = row[6].value
    colonia = row[7].value
    calle = row[8].value
    numero_exterior = row[9].value
    numero_interior = row[10].value
    entre_calles = row[11].value
    correo_electronico = row[12].value
    confirmar_correo = row[13].value
    contrasena = row[14].value
    lada = row[15].value
    numero_celular = row[16].value

    
    # Ejecuta JavaScript para modificar el contenido de la etiqueta
    checkbox = driver.find_element(By.CSS_SELECTOR, "#form > div:nth-child(2) > div.w-100.w-sm-75.mx-auto > div.steps > div.step-1 > div > div:nth-child(4) > div > div > div.col-sm-4.offset-0.offset-sm-1.text-center > p:nth-child(2) > span > span.k-switch-track.k-rounded-full")
    checkbox.click()

    # Completa los campos del formulario
    driver.find_element_by_name('nombre').send_keys(nombre)
    driver.find_element_by_name('primerApellido').send_keys(primer_apellido)
    driver.find_element_by_name('segundoApellido').send_keys(segundo_apellido)
    driver.find_element_by_name('fechaNacimiento').send_keys(fecha_nacimiento)

    # Selecciona el género en un menú desplegable
    genero_dropdown = Select(driver.find_element_by_name('genero'))
    genero_dropdown.select_by_visible_text(genero)

    # Completa los campos de dirección
    driver.find_element_by_name('CodigoPostal').send_keys(codigo_postal)
    driver.find_element_by_name('coloniaID').send_keys(colonia)
    driver.find_element_by_name('calle').send_keys(calle)
    driver.find_element_by_name('numeroExterior').send_keys(numero_exterior)
    driver.find_element_by_name('numeroInterior').send_keys(numero_interior)
    driver.find_element_by_name('entreCalles').send_keys(entre_calles)

    # Completa los campos de correo electrónico y número de celular
    driver.find_element_by_name('correo').send_keys(correo_electronico)
    driver.find_element_by_name('confirmarCorreo').send_keys(confirmar_correo)

    # Encuentra el elemento select por su nombre
    lada_select = Select(driver.find_element_by_name('lada'))
    lada_select.select_by_value(str(lada))

    driver.find_element_by_name('celular').send_keys(numero_celular)

    # Agrega los campos de Contraseña y Confirmar Contraseña
    driver.find_element_by_name('contrasena').send_keys(contrasena)
    driver.find_element_by_name('ConfirmarContrasena').send_keys(contrasena)

    # Marca el checkbox "Acepto Términos y Condiciones"
    driver.find_element_by_name('acepto').click()

    # Envía el formulario
    driver.find_element_by_name('finalizarBtn').click()

    # Espera a que la página se cargue antes de continuar al siguiente registro
    driver.implicitly_wait(10)  # Espera hasta 10 segundos para la próxima operación

# Cierra el navegador cuando hayas terminado
driver.quit()
