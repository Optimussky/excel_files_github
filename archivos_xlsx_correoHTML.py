# pip install openpyxl
# Configurar mail: https://www.letscodemore.com/blog/smtplib-smtpauthenticationerror-username-and-password-not-accepted/
# Enviar correos electrónicos Youtube: https://www.youtube.com/watch?v=mRXR8eO9igQ
import smtplib, ssl
import getpass
# Leer excel
import openpyxl as op
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# leer el archivo
book = op.load_workbook('plantilla.xlsx', data_only=True)
# fijar la hoja
hoja = book.active

celdas = hoja['A2' : 'D5']

lista_empleados = []

for fila in celdas:
    empleado = [celda.value for celda in fila]
    lista_empleados.append(empleado)

#print(lista_empleados)

username = input("Ingrese su email: ")
password =  getpass.getpass("Ingrese su password: ")

# Crear conexión segura
context = ssl.create_default_context()

with smtplib.SMTP_SSL("smtp.gmail.com", 465,context=context) as server:
    server.login(username,password)
    print("Inicio de sesión")
    for empleado in lista_empleados:
        SUBJECT = f"Constancia de pago de {empleado[0]}"
        destinatario = empleado[3]
        html = f"""
        <html>
        <body>
        <h1>Constancia de pago </h1>
        <p>
        Hola {empleado[0]}, este mes ganaste {empleado[2]}
        </p>
        </body>
        </html>
        
        """
        BODY = MIMEMultipart("alternative") # estandar
        BODY = '\r\n'.join(['To: %s' % destinatario,
        'From: %s' % username,
        'Subject: %s' % SUBJECT,
        '', html])
        server.sendmail(username,[destinatario],BODY)
        print("Mensaje Enviado ...")


