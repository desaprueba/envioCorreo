import smtplib
from email.message import EmailMessage
import pandas as pd
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import configparser


def send_email(to, subject, body):
    config = configparser.ConfigParser()
    # Leer el archivo de configuración
    config.read('configuracion.config', encoding='utf-8')
    # Obtener correo y app password de la cuenta
    from_address = config.get('Configuracion', 'correo_remitente')
    clave = config.get('Configuracion', 'pass')     
    message = body
    
    email = EmailMessage()
    email["From"] = from_address
    email["To"] = to
    email["Subject"] = "Recordatorio de vencimiento de AoC"
    email.set_content(message)

    try:
        smtp = smtplib.SMTP_SSL("smtp.gmail.com")
        smtp.login(from_address, clave)
        smtp.quit()
        print(f"Correo electrónico enviado a: {to}")
    except smtplib.SMTPException as e:
        print(f"Error al enviar el correo electrónico a {to}: {str(e)}")

def main():
    try:
        # Autenticación con Google Drive API
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('credenciales2.json', scope)  # Credenciales autenticación
        client = gspread.authorize(credentials)

        # Abrir el archivo Excel en Google Drive
        # Se tuvo que habilitar Google Sheets API
        spreadsheet = client.open("vencimientos")  # Nombre archivo en drive
        sheet = spreadsheet.sheet1  # Hoja de cálculo 1

        # Obtener los datos de la hoja de cálculo
        vencimientos = sheet.get_all_values()[1:]  # Ignorar la primera fila (encabezados)
        fecha_actual = datetime.now()
        fecha_max_notif = (fecha_actual + timedelta(days=7)).strftime("%Y-%m-%d") #  Calcula fecha máxima para envio de notificación

        # Procesar los vencimientos y enviar correos electrónicos
        for vencimiento in vencimientos:
            nombre_tercero = vencimiento[0]
            servicio = vencimiento[1]
            clasificacion = vencimiento[2]
            fecha_expiracion = vencimiento[3]
            destinatario = vencimiento[4]
            
            subject = "Recordatorio de vencimiento de AoC"
            body = f"""
                Estimado {nombre_tercero},

                Le escribimos para recordarle que su Attestation of Compliance (AoC) para el servicio de {servicio}
                está próximo a vencer. La fecha de expiración es el {fecha_expiracion}.

                Por favor, tome las medidas necesarias para renovar su AoC a tiempo.

                Saludos cordiales,
                Meli
            """            
            if fecha_expiracion <= fecha_max_notif: #Envia correo a los que estén próximos a vencer (7 días)
                print(f"El certificado de {nombre_tercero} vence pronto: {fecha_expiracion} ")
                send_email(destinatario, subject, body)
                            
    except Exception as e:
        print(f"Error durante la ejecución del programa: {str(e)}")

if __name__ == "__main__":
    main()
