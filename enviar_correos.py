import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time

# Configuración del servidor SMTP
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
REMITENTE = 'fernandonarea6@gmail.com'
PASS = 'clave-ultrasecreta-de-la-app'

try:
    df = pd.read_excel('reportes.xlsx')
    print("Archivo Excel leído correctamente.")
except Exception as e:
    print(f"Error al leer el archivo Excel: {e}")
    exit()

# Conexión y envío de correos
try:
    servidor = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    # servidor.set_debuglevel(1)
    servidor.starttls()
    servidor.login(REMITENTE, PASS)
    print("Conexión exitosa al servidor SMTP.")

    for index, column in df.iterrows():
        usuario = column['Usuario']
        campana = column['Campaña ']
        incidencias = column['Incidencias']
        nombre_archivo = column['Nombre Archivo']
        destinatario = column.get('Usuario', '')
        if not destinatario:
            print(f"Fila {index + 2}: No se encontró correo electrónico. Saltando...")
            continue
        # Creando mensaje
        msg = MIMEMultipart()
        msg['From'] = REMITENTE
        msg['To'] = destinatario
        msg['Subject'] = f"{nombre_archivo}"
        msg['']

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

        # Enviando el correo
        try:
            servidor.send_message(msg)
            print(f"Correo enviado a {destinatario} exitosamente.")
        except Exception as e:
            print(f"Error al enviar correo a {destinatario}: {e}")

        time.sleep(5)

    servidor.quit()
    print("Proceso terminado con exito.")

except smtplib.SMTPAuthenticationError:
    print("\nError de Autenticación: Outlook rechazó tu contraseña.")
    print("Solución: Probablemente necesites una 'Contraseña de Aplicación'.")
except Exception as e:
    print(f"\nError general: {e}")