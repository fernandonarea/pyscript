import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time
import os
from dotenv import load_dotenv
load_dotenv()
import logging
from logs import logging_config

logging_config()


# Configuración del servidor SMTP
SMTP_SERVER = os.getenv('SERVER_TYPE')
SMTP_PORT = os.getenv('SERVER_PORT')
REMITENTE = os.getenv('EMAIL_ADDRESS')
PASS = os.getenv('APP_PASSWORD')


try:
    logging.info('=== Proceso Iniciado ===')
    df = pd.read_excel('reportes.xlsx')
    print('Archivo Excel leido correctamente.')
    logging.info("Archivo Excel leido correctamente.")
except Exception as e:
    print("Error al leer el archivo Excel. Revisa logs.log para más detalles.")
    logging.error(f"Error al leer el archivo Excel: {e}")
    exit()

# Conexión y envío de correos
try:
    logging.info("Iniciando conexion al servidor SMTP")
    servidor = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    servidor.starttls()
    servidor.login(REMITENTE, PASS)
    logging.info("Conexion exitosa al servidor SMTP.")
    print('Conexion exitosa al servidor SMTP.')

    for index, row in df.iterrows():
        try:
            usuario = row['Usuario']
            campana = row['Campaña']
            incidencias = row['Incidencias']
            nombre_archivo = row['Nombre Archivo']
            destinatario = row.get('Usuario', '')
            logging.info(f"Leyendo filas del Excel.")
           
            if not destinatario:
                logging.warning(f"Fila {index + 2}: No se encontro correo electronico. Saltando...")
                print(f"Fila {index + 2}: No se encontro correo electronico. Saltando...")
                continue
            
            if index == 1:
                raise Exception("Segunda fila de datos, simulando error de envio.")
            
            # Creando mensaje
            msg = MIMEMultipart()
            msg['From'] = REMITENTE
            msg['To'] = destinatario
            msg['Subject'] = f"{nombre_archivo}"

            cuerpo1 = f"""
            <html>
            <body>
                <h1>Primer mensaje</h1>
                <p>Hola {usuario},</p>
                <p>Te informamos que en la campaña <b>{campana}</b> se han registrado <b>{incidencias}</b> incidencias.</p>
                <p>Por favor, revisa el archivo adjunto para más detalles.</p>
                <br>
                <p>Saludos cordiales,</p>
                <p>Equipo de Soporte</p>
            </body>
            </html>
            """

            cuerpo2 = f"""
            <html>
            <body>
                <h1>Segundo mensaje</h1>
                <p>Hola {usuario},</p>
                <p>Te informamos que en la campaña <b>{campana}</b> se han registrado <b>{incidencias}</b> incidencias.</p>
                <p>Por favor, revisa el archivo adjunto para más detalles.</p>
                <br>
                <p>Saludos cordiales,</p>
                <p>Equipo de Soporte</p>
            </body>
            </html>
            """

            if campana.lower() == 'campaña a':
                msg.attach(MIMEText(cuerpo1, 'html'))
            elif campana.lower() == 'campaña b':
                msg.attach(MIMEText(cuerpo2, 'html'))

            logging.info(f"Enviando correo a {destinatario}")

            servidor.send_message(msg)

            logging.info(f"Correo enviado a {destinatario} exitosamente.")
            print(f"Correo enviado a {destinatario} exitosamente.")
        except Exception as e:
            logging.error(f"Error al enviar correo a {row['Usuario']}: {e}")
            print(f"Error al enviar correo a {row['Usuario']}. Revisar logs para más detalles.")
            continue

        time.sleep(5)

    servidor.quit()
    logging.info("=== Proceso terminado con exito. ===")
    print('=== Proceso Terminado ===')

except smtplib.SMTPAuthenticationError:
    logging.error("Error de Autenticacion: Gmail rechazo tu clave.")
    print("Error de Autenticacion: Gmail rechazo tu clave. Revisa logs.log para más detalles.")
    logging.error("Solucion: Probablemente necesites una 'Clave de Aplicacion'.")
    print("Solucion: Probablemente necesites una 'Clave de Aplicacion'. Revisa logs.log para más detalles.")
except Exception as e:
    logging.error(f"\nError general: {e}")
    print(f"\nError general. Revisa logs.log para mas detalles.")