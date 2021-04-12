import os
import logging
import principal
from flask import Flask, flash, request, redirect, url_for, send_from_directory, render_template
from werkzeug.utils import secure_filename

from werkzeug.middleware.shared_data import SharedDataMiddleware
#
UPLOAD_FOLDER = ''

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


logging.basicConfig(level=logging.DEBUG)
logging.debug("Log habilitad!")


def allowed_file(filename, ALLOWED_EXTENSIONS):

    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def manejoDeArchivo(archivo):
    f = open(archivo, 'r')
    f.seek(0)
    logging.debug(f.readline())

#Convierte una cadena de caracteres (los nombres de columnas excel) a su valor numerico
def cadenaANum(cadena):
    listNum = []

    # recorro la cadena
    for i in cadena:
        # paso acada caracter a mayuscula y obtengo su ascii
        test = i.upper()
        ascci = ord(test)

        if(ascci < 65 or ascci > 90):
            print("Numero fuera de rango")
            return
        # del ascci lo paso a su valor numerico    
        num = ascci - 65 + 1
        # y lo añado a una lista
        listNum.append(num)

    # Luego paso esta lista (que se encuentra en base 26) a base 10
    suma = 0
    mul = pow(26,len(listNum)-1)
    for i in listNum:

        suma += i * mul
        mul /= 26

    # y finalmente devuelvo el valor final
    return suma - 1


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'excel' not in request.files:
            flash('No file apart')
            return redirect(request.url)
        

        excel = request.files['excel']
        # if user does not select file, browser also
        # submit an empty part without filename
            
        if excel.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if excel and allowed_file(excel.filename, {"xlsx","xls"}):
            filenameXLSX = "datos.xlsx"
            pathXLSX = os.path.join(app.config['UPLOAD_FOLDER'], filenameXLSX)
            excel.save(pathXLSX)

            #nombreHojaExcel = request.form['nomb-hoja']
            principal.ejecutarCodigo(pathXLSX)
            nombreOUT = 'output.xlsx'
            return redirect(url_for('uploaded_file',
                                    filename=nombreOUT))



    return render_template("index.html")



app.add_url_rule('/uploads/<filename>', 'uploaded_file',
                 build_only=True)
app.wsgi_app = SharedDataMiddleware(app.wsgi_app, {
    '/uploads':  app.config['UPLOAD_FOLDER']
})
