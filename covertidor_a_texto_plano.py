############################################################################################################################################################
#
#                                                   convertidor_a_texto_plano.py
#       
#   Realizado por: ING. David Pangay
#   Fecha:         12/11/2024
#
#   Convierte archivo Excel a texto plano (csv o txt)
#
#   Parametros de entrada:      P1:  nombre_archivo_origen (obligatorio)
#                               P2:  nombre_archivo_destino (opcional) 
#						   		P3:  la extension a convertir
#						   		P4:  separador 
#                          En caso de no recibir el parametro 2 toma el nombre del parametro 1 
#						   En caso de no recibir el parametro 4 el archivo se guarda con el separador ","
#							
#   Parametros de salida:     Archivo convertido en Texto Plano
#
#   Ruta: /home/ProcesosSistemas/fuentes/
#   Ejecucion: python3 convertidor_a_texto_plano.py "convertidor_a_texto_plano.py" .csv || .txt ","
#   Ejecucion: python3 convertidor_a_texto_plano.py "convertidor_a_texto_plano.py" .csv || .txt \t (en opcion de tabulacion, no usar comillas"
############################################################################################################################################################

#Importamos bibliotecas:

#pandas permite leer, manipular, y guardar datos en estructuras como DataFrames
import pandas as pd  
#sys para acceder a argumentos de la li­nea de comandos
import sys  
#os para manejar archivos y rutas
import os 
#datetime Para obtener la fecha actual y formatearla.
import datetime 

#Funcion para convertir un archivo Excel a formato CSV o TXT
def convertir_a_txt_o_csv(archivo_entrada, archivo_salida, separador):
    try:
#Usa openpyxl como motor para leer archivos .xlsx
        df = pd.read_excel(archivo_entrada, engine="openpyxl")
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        sys.exit(1)

#Verifica que el separador sea un solo caracter
    if len(separador) != 1:
        print(f"Error: El separador debe ser un unico caracter.")
        sys.exit(1)

#Obtener la extension de salida (csv o txt)
    ext_salida = archivo_salida.split('.')[-1].lower()

    if ext_salida == 'csv':
        try:
#Guarda como CSV con el separador que se especifique
            df.to_csv(archivo_salida, index=False, sep=separador)
        except Exception as e:
            print(f"Error al guardar el archivo CSV: {e}")
            sys.exit(1)
    elif ext_salida == 'txt':
        try:
#Guarda como TXT con el separador que se especifique
            df.to_csv(archivo_salida, index=False, sep=separador)
        except Exception as e:
            print(f"Error al guardar el archivo TXT: {e}")
            sys.exit(1)
    else:
        print("Error: Solo se admiten formatos .csv y .txt para la salida.")
        sys.exit(1)

#Funcion para generar el nombre de salida
def generar_nombre_salida(archivo_entrada, extension_salida):
    fecha_actual = datetime.datetime.now()
    fecha_formateada = fecha_actual.strftime("%Y-%m-%d")

    if extension_salida not in ['.csv', '.txt']:
        print("Error: La extension de salida debe ser .csv o .txt.")
        sys.exit(1)

    base_nombre = os.path.splitext(archivo_entrada)[0]
    nombre_final = f"{base_nombre}_({fecha_formateada}){extension_salida}"

#Verifica si el archivo ya existe y agrega un contador
    contador = 1
    while os.path.exists(nombre_final):
        nombre_final = f"{base_nombre}_({fecha_formateada})({contador}){extension_salida}"
        contador += 1

    return nombre_final

#Funcion principal que se ejecuta cuando el script es llamado
def main():
#Verifica si hay suficientes argumentos
    if len(sys.argv) < 3:
        print("Uso: python convertidor_a_texto_plano.py <archivo_entrada.xlsx> <.csv o .txt> [separador]")
        sys.exit(1)

    archivo_entrada = sys.argv[1]
    extension_salida = sys.argv[2]

#Obtiene el separador, si no se proporciona, usa coma "," por defecto
    separador = sys.argv[3] if len(sys.argv) > 3 else ','

    print(f"Buscando el archivo: {archivo_entrada}")
    
#Verifica si el archivo existe
    if not os.path.isfile(archivo_entrada):
        print(f"Error: El archivo {archivo_entrada} no existe.")
        sys.exit(1)

#Genera el nombre del archivo de salida con fecha
    archivo_salida = generar_nombre_salida(archivo_entrada, extension_salida)

#Convierte el archivo a formato CSV o TXT con el separador especificado
    convertir_a_txt_o_csv(archivo_entrada, archivo_salida, separador)

#Mensaje
    print(f"¡Proceso Exitoso!")
    print(f"Archivo convertido y guardado como {archivo_salida}")

#Ejecuta la función principal si el script se ejecuta directamente
if __name__ == "__main__":
    main()

#Fin del proceso