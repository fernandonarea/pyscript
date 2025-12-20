import pandas as pd
import smtplib
import time
import os
import logging
import glob
from dotenv import load_dotenv
from logs import logging_config
from jinja2 import Environment, FileSystemLoader
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Cargando variables de entorno y configuración de los LOGS
load_dotenv()
logging_config()

file_loader = FileSystemLoader(".")
env = Environment(loader=file_loader)

SERVER_TYPE = os.getenv("SERVER_TYPE")
SERVER_PORT = os.getenv("SERVER_PORT")
REMITENTE = os.getenv("EMAIL_ADDRESS")
PASS = os.getenv("APP_PASSWORD")
ORIGEN_DATOS = os.getenv("CARPETA_ORIGEN")
EMAIL_CC = ["fernando@gmail.com", "Correo@.com", ""]

# Obtención y lectura de archivo .xlsx más reciente
try:
    logging.info("== Proceso inciado ==")
    archivos = glob.glob(os.path.join(ORIGEN_DATOS, "*.xlsx"))
    archivo_reciente = max(archivos, key=os.path.getctime)
    logging.info(f"Archivo {archivo_reciente} obtenido")

    df = pd.read_excel(archivo_reciente)
    logging.info(f"Archivo {archivo_reciente} leido exitosamente")
except Exception as e:
    logging.error("Error en la obtencion y lectura del archivo: {e}")
    exit(1)

# Obtención de datos y envío de mensaje
try:
    # Conexion al servidor de envio de correos
    servidor = smtplib.SMTP(SERVER_TYPE, SERVER_PORT)
    servidor.starttls()
    servidor.login(REMITENTE, PASS)
    logging.info("Conexion exitosa con el servidor")

    # Lectura de datos del .xlsx
    for index, row in df.iterrows():
        try:
            usuario = row["Usuario"]
            incidencias = row["Incidencias"]
            campana = row["Campaña"]
            nombre_archivo = row["Nombre Archivo"]
            destinatario = row.get("Usuario", "")

            if not destinatario:
                logging.error(
                    f"Fila {index + 2}: No se encontró el dato de destinatario"
                )

            logging.info("Construyendo correo")
            
            # Creación del mensaje
            msg = MIMEMultipart()
            msg["From"] = usuario
            msg["To"] = destinatario
            msg["Subject"] = nombre_archivo
            msg["Cc"] = EMAIL_CC

            # Obteniendo el template HTML
            template = env.get_template("email_templates.html")

            # Llenado del template HTML
            mensaje_html = template.render(
                usuario=usuario, incidencias=incidencias, nombre_archiv=nombre_archivo
            )

            msg.attach(MIMEText(mensaje_html, "html"))

            logging.info(f"Enviando Correo a {destinatario}")
            servidor.send_message(msg)
            logging.info(f"Correo enviado a {destinatario} con exito")

        except Exception as e:
            logging.error(
                f"Error al enviar correo a {row['Usuario']}: ",
                e,"Saltando al siguiente usuario",
            )
            continue

        time.sleep(3)
    servidor.quit()
    logging.info("== Proceso Terminado con exito")

except smtplib.SMTPAuthenticationError:
    logging.error("Gmail rechazó tu clave de aplicacion")
    logging.warning("Probablemente se necesite una clave de aplicacion")
except Exception as e:
    logging.error("Error general del script: {e}")