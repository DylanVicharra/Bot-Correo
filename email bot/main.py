import os
from ntpath import basename
import archivos as ar
import mail as mail

try:
    # Elimino el anterior log_error
    ar.eliminar_log_error()
    # Verifico y creo carpetas, ademas de los archivos
    ar.verificacion_carpetas()
    ar.verificacion_archivo()
    # Declaro todos los parametros necesario para el envio del correo
    datos = ar.read_json()
    password = datos["password"]
    remitente = datos["usuario"]
    destinatario = "salestax@apple.com"
    asunto = ""
except Exception as ex:
    print(f'{ex}' + '\n' + "Finalizando..." + '\n')
    exit(1)


def main():
    
    # Se pide la variables de busqueda 
    os.system('cls')
    print("                 ============ BOT - EMAIL TAX ============                 ")
    print("Aclaracion el nombre de la carpeta tiene que ser ingresada de en el siguiente formato 'AAAA-MM-DD-Nºcarpeta'")

    try:
        nombre_carpeta = input("Ingrese un nombre de carpeta: ")
    except:
        print("Se han ingresado datos erroneos" + "\n" + "Finalizando...")
        exit(1)

    rutas = ar.obtencion_archivos(nombre_carpeta)

    if rutas:
        # Inicio la conexion con el servidor outlook
        servidor = mail.ingreso_servidor(remitente, password)

        for ruta in rutas:
            if rutas[ruta]:
                for archivo in rutas[ruta]:
                    try:
                        correo = mail.crear_correo(remitente, destinatario, asunto=str(archivo).replace("-invoice duplicate.pdf", "").strip())
                        correo = mail.adjuntar_archivo([f'{ruta}\\{archivo}', ar.archivo_fijo()], correo)
                        
                        mail.envio_correo(servidor, remitente, destinatario, correo)
                        print(f"Se ha enviado correctamente el archivo {archivo}")
                    except:
                        ar.log_error(basename(ruta), archivo)
                        print(f"Archivo no enviado: {archivo} - Carpeta: {basename(ruta)}")
            else:
                print("La carpeta no contiene archivos")
    
        mail.finalizar_servidor(servidor)
    else:
        print("La carpeta no existe")

    print("Finalizando BOT - EMAIL TAX ...")

if __name__ == "__main__":
    main()