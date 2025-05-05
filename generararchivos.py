import random
import string

def generar_cadena_con_longitud(min_total_len=200, id_len=7):
    separador = 1  # la coma
    min_dato_len = min_total_len - id_len - separador
    return ''.join(random.choices(string.ascii_letters + string.digits, k=min_dato_len))

# Ruta del archivo de salida
archivo_salida = "archivo_5_millones.txt"

# Generar y guardar los 5 millones de registros
with open(archivo_salida, "w") as f:
    for i in range(1, 5_000_001):
        id_val = f"{i:07d}"
        dato = generar_cadena_con_longitud()
        f.write(f"{id_val},{dato}\n")

print("Archivo generado exitosamente: archivo_5_millones.txt con líneas de mínimo 200 caracteres.")
