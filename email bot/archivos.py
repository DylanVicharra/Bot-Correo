import os
import json
from datetime import datetime, timedelta
from pathlib import Path

src_path = Path(__file__).parent
main_path = src_path.parents[1]
data_path = src_path.parent / 'data'

def archivo_fijo():
    ruta = data_path / 'Resale certificate Camlem 2021.pdf'
    return ruta

def read_json():
    with open(data_path / 'correo.json', encoding='utf-8') as archivo_json:
        return json.load(archivo_json)

def download_path():
    for root, _, _ in os.walk(main_path):
        if('download' in root.lower()):
            return root

def obtencion_archivos(listado_fechas):
    listado_carpeta = os.listdir(download_path())
    # Declaro una variable para guardar cada una de las carpetas que cumplen las condiciones de fecha
    carpetas = {}
    for carpeta in listado_carpeta:
        for fecha in listado_fechas:
            if fecha in carpeta:
                carpetas[f'{download_path()}\\{carpeta}'] = os.listdir(f'{download_path()}\\{carpeta}')

    return carpetas

def listar_fechas(fecha_inicio, fecha_fin):
    fecha_inicio = datetime.strptime(fecha_inicio, "%Y-%m-%d")
    fecha_fin = datetime.strptime(fecha_fin,"%Y-%m-%d")

    if fecha_inicio < fecha_fin:
        lista_fechas = [(fecha_inicio + timedelta(days=d)).strftime("%Y-%m-%d") for d in range((fecha_fin - fecha_inicio).days + 1)]
        return lista_fechas
    else: 
        print("La fecha de inicio es una fecha mayor a la fecha final. Ingrese un intervalo de fechas correcta.")
        return False

def log_error(carpeta, archivo):
    if os.path.exists(f'{src_path.parent}\\log.txt') == False:
        with open(f'{src_path.parent}\\log.txt', 'w') as error:
            error.write(f'*{datetime.today()} ---- Ha ocurrido un error con el archivo {archivo} en la carpeta {carpeta}' + "\n")
    else: 
        with open(f'{src_path.parent}\\log.txt', 'a') as error:
            error.write(f'*{datetime.today()} ---- Ha ocurrido un error con el archivo {archivo} en la carpeta {carpeta}' + "\n")