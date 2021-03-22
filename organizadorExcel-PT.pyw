from tkinter import filedialog
from tkinter import *
from functools import partial
import os
import main


global pathDATOS
###
#

global existe_path_archivo_excel
existe_path_archivo_excel = False

global se_puede_ejecutar
se_puede_ejecutar = False


def ejecutarPrograma():
    if(se_puede_ejecutar):
        print("Programa ejecutado correctamente")
        main.ejecutarCodigo(pathDATOS)
    else:
        print("No se ejecutó el programa")


def comprobarEjecutar(mi_frame,botonEjecutar):
    global se_puede_ejecutar
    if(existe_path_archivo_excel):
        botonEjecutar.configure(bg="pale green")
        se_puede_ejecutar = True
    else:
        botonEjecutar.configure(bg="red")
        se_puede_ejecutar = False
        print("comprobando falso " + str(existe_path_archivo_excel))


def abrir_archivo_excel(boton, botonEjecutar, mi_frame):
    global existe_path_archivo_excel
    archivo_abierto = filedialog.askopenfilename(initialdir = os.getcwd(), title = "Seleccione archivo", filetypes = (("excel", "*.xlsx"),("all files")))
    if(archivo_abierto != ""):
        boton.configure(bg="pale green")
        existe_path_archivo_excel = True
        print(archivo_abierto)
    else:
        boton.configure(bg="red")
        existe_path_archivo_excel = False
    
    global pathDATOS
    pathDATOS = archivo_abierto
    comprobarEjecutar(mi_frame, botonEjecutar)



# creo la interfaz
raiz = Tk()
mi_Frame = Frame(raiz)
mi_Frame.pack()

# creo los elementos
texto2 = Label(mi_Frame, text="Seleccione su archivo excel modificado")
boton2 = Button(mi_Frame,text = "Abrir archivo excel")
botonEjecutar = Button(mi_Frame, bg = "red", text = "Ejecutar", command = ejecutarPrograma)

boton2.configure(command = partial(abrir_archivo_excel,boton2,botonEjecutar,mi_Frame)) 
# los posiciono
texto2.grid(row=2, column=0)
boton2.grid(row=3, column=0)
botonEjecutar.grid(row = 4,column = 1)

raiz.mainloop()


