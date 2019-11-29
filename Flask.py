from flask import Flask, render_template, request, flash
import xlwt
from xlwt import Workbook
import pandas as pd
import datetime
import xlsxwriter
import os
import fpdf
from fpdf import FPDF
import sys
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from pymongo import MongoClient
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

app = Flask(__name__)

client = MongoClient("mongodb://127.0.0.1:27017")
db = client.Comandos
todos = db.ComandosEnviados

usuario = ""
correo = ""
comando = ""

@app.errorhandler(404)
def not_found(error):
    return render_template("Error.html")

@app.route("/", methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        print("Ciente conectado con exito")
    elif request.form.get('Ingresar') == 'Ingresar':
        global usuario
        global correo
        usuario = request.form['Usuario']
        correo = request.form['Correo']
        print("Se ha detectado una nueva conexion.\nUsuario: "+usuario+"\nCorreo: "+correo)
        return render_template("index.html")
    return render_template("User.html")

@app.route("/Menu", methods=['GET', 'POST'])
def Menu():
    return render_template("index.html")

@app.route("/Comandos", methods=['GET', 'POST'])
def Comandos():
    if request.form.get('Enviar') == 'Enviar':
        global comando
        comando = request.form['comando']
        now = datetime.datetime.now()
        newDirName = now.strftime("%H:%M:%S %d-%m-%Y")
        database_entry={'Usuario:':usuario, 'Comado:':comando, 'Hora y Fecha:':newDirName, 'Ip Del Servidor:':request.remote_addr}
        todos.insert_one(database_entry)
        os.system(comando+" > Respuesta.txt")
        with open ("Respuesta.txt", "r") as myfile:
            data=myfile.read().replace('\n', ' ')
            myfile.close()
        if not data:
            data="Comando invalido, intenta con uno nuevo!"
        else: data=data
        return render_template("Comandos.html", data=data)
    elif request.form.get('Excel') == 'Excel':
        datos = pd.read_csv('Respuesta.txt',error_bad_lines=False,engine="python",index_col=False,header=None)
        datos.to_excel("Excel.xlsx", index=False, header=False)
        Fuente = request.form.get("Fuente")
        Tamaño = request.form.get("Tamaño")
        Color = request.form.get("Color")
        if Color == "Negro":
            Colorfont = "000000"
        elif Color == "Rojo":
            Colorfont = "DB0000"
        elif Color == "Verde":
            Colorfont = "0B6623"
        file = 'Excel.xlsx'
        wb = load_workbook(filename=file)
        ws = wb['Sheet1']
        colorfuente = Font(name=Fuente, color=Colorfont, size=Tamaño)
        for cell in ws["A"]:
            cell.font = colorfuente
        wb.save(filename=file)
        print("\nEl archivo excel se ha creado con exito!")
        # Iniciamos los parámetros del script
        remitente = 'solariumdelvallebr@gmail.com'
        destinatarios = correo
        asunto = '[FLASK]Resultado del comando '+comando+' en formato Excel'
        cuerpo = 'EL siguiente archivo contiene el resultado del comando solicitado!'
        ruta_adjunto = 'Excel.xlsx'
        nombre_adjunto = 'Excel.xlsx'
        mensaje = MIMEMultipart()
        mensaje['From'] = remitente
        mensaje['To'] = destinatarios
        mensaje['Subject'] = asunto
        mensaje.attach(MIMEText(cuerpo, 'plain'))
        archivo_adjunto = open(ruta_adjunto, 'rb')
        adjunto_MIME = MIMEBase('application', 'octet-stream')
        adjunto_MIME.set_payload((archivo_adjunto).read())
        encoders.encode_base64(adjunto_MIME)
        adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % nombre_adjunto)
        mensaje.attach(adjunto_MIME)
        sesion_smtp = smtplib.SMTP('smtp.gmail.com', 587)
        sesion_smtp.starttls()
        sesion_smtp.login('solariumdelvallebr@gmail.com','Solariumdelvalle')
        texto = mensaje.as_string()
        sesion_smtp.sendmail(remitente, destinatarios, texto)
        sesion_smtp.quit()
        return render_template("ExcelE.html", datos=datos)
    elif request.form.get('Pdf') == 'Pdf':
        with open ("Respuesta.txt", "r", encoding="utf-8") as myfile:
            data=myfile.read().replace('\n', ' ')
            myfile.close()
        Fuente = request.form.get("Fuente")
        Tamaño = request.form.get("Tamaño")
        Orientacion = request.form.get("Orientacion")
        Color = request.form.get("Color")
        if Color == "Negro":
            R = 0
            G = 0
            B = 0
        elif Color == "Rojo":
            R = 255
            G = 0
            B = 0
        elif Color == "Verde":
            R = 0
            G = 255
            B = 0
        pdf=FPDF()
        pdf.add_page()
        pdf.set_font(Fuente,"",int(Tamaño))
        pdf.set_text_color(R,G,B)
        pdf.set_margins(10,10,10)
        pdf.multi_cell(0, 6, data, 0, Orientacion)
        pdf.ln(h="")
        pdf.output("PDF.pdf","f")
        print("\nEl archivo pdf se ha creado con exito!")
        remitente = 'solariumdelvallebr@gmail.com'
        destinatarios = correo
        asunto = '[FLASK]Resultado del comando '+comando+' en formato PDF'
        cuerpo = 'EL siguiente archivo contiene el resultado del comando solicitado!'
        ruta_adjunto = 'PDF.pdf'
        nombre_adjunto = 'PDF.pdf'
        mensaje = MIMEMultipart()
        mensaje['From'] = remitente
        mensaje['To'] = destinatarios
        mensaje['Subject'] = asunto
        mensaje.attach(MIMEText(cuerpo, 'plain'))
        archivo_adjunto = open(ruta_adjunto, 'rb')
        adjunto_MIME = MIMEBase('application', 'octet-stream')
        adjunto_MIME.set_payload((archivo_adjunto).read())
        encoders.encode_base64(adjunto_MIME)
        adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % nombre_adjunto)
        mensaje.attach(adjunto_MIME)
        sesion_smtp = smtplib.SMTP('smtp.gmail.com', 587)
        sesion_smtp.starttls()
        sesion_smtp.login('solariumdelvallebr@gmail.com','Solariumdelvalle')
        texto = mensaje.as_string()
        sesion_smtp.sendmail(remitente, destinatarios, texto)
        sesion_smtp.quit()
        return render_template("PdfE.html", data=data)
    return render_template("Comandos.html")

@app.route("/Registro", methods=['GET', 'POST'])
def Registro():
    registro = open("Registros.txt", "w")
    for doc in todos.find():
        registro.writelines(str(doc)+"\n")
    registro.close()
    with open ("Registros.txt", "r", encoding="utf-8") as myfile:
        data=myfile.readlines()
        myfile.close()
    return render_template("Registro.html", data=data)

if __name__ == "__main__":
    app.run(debug=True,host='0.0.0.0')