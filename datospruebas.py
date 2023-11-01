import openpyxl
from random import randint
from faker import Faker

# Crea un objeto Faker para generar datos ficticios
faker = Faker()

# Crea un nuevo libro de Excel y selecciona una hoja
workbook = openpyxl.Workbook()
sheet = workbook.active

# Agrega encabezados para tus columnas
headers = ["Nombre", "Primer Apellido", "Segundo Apellido", "Fecha de Nacimiento", "Género", "Tipo de Documento",
           "Código Postal", "Colonia", "Calle", "Número Exterior", "Número Interior", "Entre Calles",
           "Correo Electrónico", "Confirmar Correo Electrónico", "Contraseña", "Lada", "Número de Celular"]

sheet.append(headers)

# Genera 10 registros ficticios y agrégalos al archivo Excel
for _ in range(10):
    nombre = faker.first_name()
    primer_apellido = faker.last_name()
    segundo_apellido = faker.last_name()
    fecha_nacimiento = faker.date_of_birth(minimum_age=18)
    genero = faker.random_int(min=1,max=3)
    tipo_documento = 'Acta de Nacimiento'
    codigo_postal = faker.random_element(elements=('85210', '83280','85150'))
    colonia = faker.random_int(min=1,max=6)
    calle = faker.street_name()
    numero_exterior = str(randint(1, 100))
    numero_interior = str(randint(1, 20)) if randint(1, 10) % 2 == 0 else ""
    entre_calles = faker.street_name() + " y " + faker.street_name()
    correo_electronico = faker.email()
    confirmar_correo = correo_electronico
    contrasena = faker.password()
    lada = faker.random_int(min=1, max=21)
    numero_celular = faker.random_int(min=1000000000, max=9999999999)

    numero_celular = str(randint(1000000000, 9999999999))
    

    # Agrega los datos generados a la fila
    row = [nombre, primer_apellido, segundo_apellido, fecha_nacimiento, genero, tipo_documento,
           codigo_postal, colonia, calle, numero_exterior, numero_interior, entre_calles,
           correo_electronico, confirmar_correo, contrasena, lada, numero_celular]
    sheet.append(row)

# Guarda el archivo Excel
workbook.save('registros.xlsx')
