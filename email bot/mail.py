import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from ntpath import basename


def adjuntar_archivo(lista_rutas, correo):
    for ruta in lista_rutas:

        with open(ruta, "rb") as pdf:
            archivo_pdf = MIMEBase('application','pdf')
            archivo_pdf.set_payload(pdf.read())
            pdf.close()

        encoders.encode_base64(archivo_pdf)
        archivo_pdf.add_header(
            "Content-Disposition",
            f"attachment; filename= {basename(ruta)}",
        )
        correo.attach(archivo_pdf)

    return correo

def ingreso_servidor(usuario, password):
    context = ssl.create_default_context()

    try:
        server = smtplib.SMTP(host='smtp.office365.com', port='587')
        server.starttls(context=context)
        server.login(usuario, password)
        return server
    except Exception as e:
        print(e)
        print("Ha ocurrido un error con la conexion del servidor")


def envio_correo(server, remitente, destinatario, correo):
    try:
        server.sendmail(remitente,destinatario,correo.as_string())
    except Exception as e:
        print(e)
        raise Exception ('Ha ocurrido un error con el envio del correo')
        
def finalizar_servidor(server):
    server.quit()

def crear_correo(remitente, destinatario, asunto):
    correo = MIMEMultipart()
    correo['From'] = remitente
    correo['To'] = destinatario
    correo['Subject'] = asunto

    return correo