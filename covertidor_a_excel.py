############################################################################################################################################################
#
#                                                   convertidor_a_excel.py
#       
#   Realizado por: ING. David Pangay
#   Fecha:         12/11/2024
#
#   Convierte archivo plano (csv o txt) a Excel
#
#   Parametros de entrada:      P1:  nombre_archivo_origen (obligatorio)
#                               P2:  nombre_archivo_destino (opcional)
#                          En caso de no recibir el parametro 2 toma el nombre del parametro 1
#   Parametros de salida:     Archivo convertido en Excel
#
#   Ruta: /home/ProcesosSistemas/fuentes/
#   Ejecucion: python3 convertidor_a_excel.py "archivo.csv || .txt"
############################################################################################################################################################


#Importamos bibliotecas:

#pandas permite leer, manipular, y guardar datos en estructuras como DataFrames
import pandas as pd  
#sys para acceder a argumentos de la lÃ­nea de comandos
import sys  
#os para manejar archivos y rutas
import os 
#datetime Para obtener la fecha actual y formatearla.
import datetime 


#Convierte un archivo CSV o TXT a formato Excel (.xlsx)
def convertir_a_excel(archivo_entrada, archivo_salida):
#Detecta el separador en el archivo de entrada (CSV o TXT)
    with open(archivo_entrada, 'r', encoding='utf-8') as f:
        primera_linea = f.readline()  

#Define los posibles separadores que podrían usarse en el archivo
    posibles_separadores = [',', '\t', ';', '|']
    separador = None  

#lee los separadores posibles
    for sep in posibles_separadores:
        if sep in primera_linea:  
            separador = sep
            break 

#Verifica si se encontró un separador válido
    if separador is None:
        print(f"Error: No se pudo detectar un separador en el archivo {archivo_entrada}.")
        return

#Lee el archivo CSV o TXT con el separador detectado
    try:
        df = pd.read_csv(archivo_entrada, sep=separador)  
    except Exception as e:
        print(f"Error al leer el archivo {archivo_entrada}: {e}")
        return

#Guarda el DataFrame como un archivo Excel
    try:
        df.to_excel(archivo_salida, index=False)  
    except Exception as e:
        print(f"Error al guardar el archivo Excel {archivo_salida}: {e}")
        return  

#Genera el nombre de salida con fecha y contador en caso de que ya exista el archivo
def generar_nombre_salida(archivo_entrada, archivo_salida=None):
    fecha_actual = datetime.datetime.now()
    fecha_formateada = fecha_actual.strftime("%Y-%m-%d")

#Si no se proporciona el archivo de salida, usa el archivo de entrada como base
    if not archivo_salida:
        base_nombre = os.path.splitext(archivo_entrada)[0]
        archivo_salida = f"{base_nombre}-({fecha_formateada}).xlsx"
    
#Verifica si el archivo ya existe
    base_nombre, ext = os.path.splitext(archivo_salida)
    if os.path.exists(archivo_salida):
        contador = 1
        nuevo_nombre = f"{base_nombre}({contador}){ext}"
        while os.path.exists(nuevo_nombre):
            contador += 1
            nuevo_nombre = f"{base_nombre}({contador}){ext}"
        return nuevo_nombre  

    return archivo_salida

#Ejecucion del script cuando es llamado
def main():
    if len(sys.argv) < 2:
        print("Uso: python convertidor_a_excel.py <archivo_entrada.csv> o <archivo_entrada.txt> [archivo_salida.xlsx]")
        sys.exit(1)

    archivo_entrada = sys.argv[1]  # El primer argumento es el archivo de entrada

#Verifica si el archivo de entrada existe
    if not os.path.isfile(archivo_entrada):
        print(f"Error: El archivo {archivo_entrada} no se encuentra en el directorio actual.")
        sys.exit(1)

#Si no se pasa un archivo de salida, se genera uno automáticamente
    archivo_salida = sys.argv[2] if len(sys.argv) > 2 else None

#Genera el nombre de salida (con fecha y contador si es necesario)
    archivo_salida = generar_nombre_salida(archivo_entrada, archivo_salida)

#Llama a la función para convertir el archivo a formato Excel
    convertir_a_excel(archivo_entrada, archivo_salida)

#Mensaje
    print(f"¡Proceso Exitoso!")
    print(f"El archivo {archivo_entrada} fue convertido y guardado como {archivo_salida}")

#Ejecuta la función principal si el script se ejecuta directamente
if __name__ == "__main__":
    main()

#Fin del proceso