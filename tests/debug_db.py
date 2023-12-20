import json
from datetime import datetime

# Convertir la base de datos de JSON a un diccionario de Python
input_database = r"E:\Python\Emes\despachos\data\emes-despachos-19122023.json"
output_database = r"E:\Python\Emes\despachos\data\depurated_database.json"

# Cargar la base de datos desde un archivo JSON con el códec adecuado
with open(input_database, encoding='utf-8') as file:
    data = json.load(file)

# Crear una lista para almacenar las claves de los elementos que se deben eliminar
items_to_delete = []

fecha_referencia = datetime.strptime("2023-11-10", "%Y-%m-%d")

# Recorrer los elementos de la clave "empaques"
for key, value in data['empaques'].items():
    if not value:
        items_to_delete.append(key)
        continue

    factura = str(value['factura'])
    estado = value['estado']
    fecha_actual = datetime.strptime(value['fecha'], "%Y-%m-%d %H:%M")

    # Paso 1: Quitar elementos cuya factura empiece por 5
    if int(factura[0]) <= 5:
        items_to_delete.append(key)

    # Paso 2: Quitar elementos con estado -1
    if estado == -1:
        items_to_delete.append(key)

    # Verificar si la fecha actual es anterior a la fecha de referencia
    if fecha_actual < fecha_referencia:
        items_to_delete.append(key)

    # Paso 3: Quitar elementos con valor no dict o empty string
    if not isinstance(value, dict) or value == "":
        items_to_delete.append(key)

keys = data["empaques"].keys()

# Eliminar los elementos marcados para eliminación
for uid in items_to_delete:
    if uid in keys:
        del data['empaques'][uid]

empaques_ordenados = {
    uid: empaque
    for uid, empaque in sorted(data['empaques'].items(), key=lambda x: x[1]['fecha'])
}

sorted_data = {'empaques': empaques_ordenados}

# # Guardar el diccionario resultante en un nuevo archivo JSON
with open(output_database, 'w', encoding='utf-8') as file:
    json.dump(sorted_data, file, indent=2)
