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
    
    # Inicio la conexion con el servidor outlook
    servidor = mail.ingreso_servidor(remitente, password)

    # Se pide la variables de busqueda 
    os.system('cls')
    print("                 ============ BOT - EMAIL TAX ============                 ")
    print("Aclaracion las fechas tienen que ser ingresadas en el siguiente formato AAAA-MM-DD")

    try:
        inicio = input("Ingrese una fecha de inicio: ")
        fin = input("Ingrese una fecha de fin: ")

        lista_fechas = ar.listar_fechas(inicio,fin)
    except:
        print("Se han ingresado datos erroneos" + "\n" + "Finalizando...")
        exit(1)

    
    if lista_fechas:
        rutas = ar.obtencion_archivos(lista_fechas)

        for ruta in rutas:
            if rutas[ruta]:
                for archivo in rutas[ruta]:
                    try:
                        correo = mail.crear_correo(remitente, destinatario, asunto)
                        correo = mail.adjuntar_archivo([f'{ruta}\\{archivo}', ar.archivo_fijo()], correo)
                        
                        mail.envio_correo(servidor, remitente, destinatario, correo)
                        print(f"Se ha enviado correctamente el archivo {archivo}")
                    except:
                        ar.log_error(basename(ruta), archivo)
                        print(f"Archivo no enviado: {archivo} - Carpeta: {basename(ruta)}")
    
    
    mail.finalizar_servidor(servidor)
    
    print("Finalizando BOT - EMAIL TAX ...")

if __name__ == "__main__":
    main()