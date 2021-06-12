# LISTADOS URL FILTERING FORTIGATE
"""
Programa que genera el script para un listado de url filtering de un Fortigate,
el archivo fuente es un documento de excel, generando un archivo de texto para
el script, tambien verifica el listado, para filtrar y retirar las url's que
no funcionan o con un status diferente a 200, tambien genera el archivo en excel
con el listado de url's y con su respectivo estatus.

Tambien por medio de Selenium, ejecuta el upload del script directamente
al fortigate, por medio del chromedrive.exe.
"""
# En este módulo se define todo el manejo de la parte gráfica del proyecto

from tkinter import *
from tkinter import filedialog as FileDialog
from io import open
from funciones.funciones import *       # Paquete y módulo con las respectivas funciones logicas del programa
import openpyxl                         # Libreria para tratamiento de documentos excel
import sys



ruta_excel =""
ruta_s =""

root = Tk()
root.title("Converter Url2ScriptForti")

# Funciones

def abrir():
    global ruta_excel

    ruta_excel = FileDialog.askopenfilename(
        title="Abrir el archivo",
        initialdir='.',
        filetypes=(("Ficheros de excel", "*.xlsx"),)                                            )
    abrir_excel(ruta_excel,texto)     #Función para cargar el archivo de excel
    mensaje.set("Se cargo exitosamente el archivo de excel")
    entry1.config(state="normal")
    entry2.config(state="normal")
    entry3.config(state="normal")
    entry4.config(state="normal")
    entry5.config(state="normal")

def guardar_script():
    global ruta_script,estatus_script
    estatus_script = 0
    if estatus_run == 0:
        MessageBox.showwarning("Reintentar", "Primero debe ejecutar el Botón RUN, luego intente guardar el script")
    elif estatus_run == 1:
        fichero = FileDialog.asksaveasfile(title="Guardar el Script", mode="a+", defaultextension=".txt")
        try:
            ruta_script = fichero.name
            fichero.close()
            f_script = crear_script(name_vdom.get(), lista_perfiles, ruta_script)
            mensaje.set("Se Guardo exitosamente el Script ")
            estatus_script = 1
        except:
            pass

def guardar_excel():
    if estatus_run == 0:
        MessageBox.showwarning("Reintentar", "Primero debe ejecutar el Botón RUN, luego intente guardar el excel")
    elif estatus_run == 1:
        fichero = FileDialog.asksaveasfile(title="Guardar el Excel", mode="a+", defaultextension=".xlsx")
        try:
            ruta = fichero.name
        except:
            pass
        fichero.close()
        sys.stdout.flush()
        ruta1 = sys.path
        f_excel = openpyxl.load_workbook(ruta1[0] + "/Url Filter Estatus.xlsx")
        f_excel.save(ruta)
        mensaje.set("Se Guardo exitosamente el Excel ")

def run():
    global doc,name_hoja, ruta_excel,estatus_run;
    estatus_run = 0
    try:
        doc = openpyxl.load_workbook(ruta_excel)
        hoja = doc.get_sheet_by_name(str(name_hoja.get()))  # Se carga la hoja donde esta el listado
        if check_vdom.get() == 1:
            if name_vdom.get() == 'None' or name_vdom.get() == "":
                error = 1
                raise ValueError("El nombre del vdom esta vacio, debe escribir el nombre del vdom")
        if columna_excel.get() == "":
            error = 2
            raise ValueError("El nombre de la columna esta vacio, debe escribir el nombre de la columna")
    except ValueError:
        if error == 1:
            MessageBox.showwarning("Nombre VDOM Vacio", "Escriba el nombre del vdom")
        elif error == 2:
            MessageBox.showwarning("Nombre de Columna Vacio", "Escriba la letra de la columna donde esta el listado")
    except Exception:
        MessageBox.showwarning("Reintentar", "El nombre de la hoja no existe en el archivo excel")
    else:
        mensaje.set("Se verificaron el listado de Url's exitosamente ")
        i,k = lista_hoja(doc,hoja,columna_excel.get())
        if i == 0:
            MessageBox.showwarning("Columna Vacia", "La Columna indicada esta vaciada, verifiquela y corriga la columna")
        else:
            estatus_run = 1
            MessageBox.showinfo("Finalizado exitosamente", "Número de Url's: {}\n Número de Url's buenas: {}".format(i,k))  # (titulo, información)

def enviar():
    if estatus_script == 0:
        MessageBox.showwarning("Estatus Script", "Primero Guarde el Script antes de Enviar")
    else:
        if ip_fortigate == "":
            MessageBox.showwarning("Dirección IP Fortigate", "Escriba la Dirección IP del Fortigate\nEjemplo: https://192.168.1.99:20443")
        elif user_fortigate == "":
            MessageBox.showwarning("User Fortigate", "Escriba el Usuario correspondiente.")
        else:
            send_forti(ip_fortigate.get(),user_fortigate.get(),pass_fortigate.get(),ruta_script)

def c_vdom():
    global name_vdom
    if check_vdom.get() == 1:
        label6 = Label(frame1, text="Nombre del VDOM:")
        label6.grid(row=4, column=1, sticky='e', padx=3, pady=3)

        entry6 = Entry(frame1, textvariable=name_vdom)
        entry6.grid(row=4, column=2, sticky='e', padx=3, pady=3)
        entry6.config(justify="left", state="normal")
    else:
        name_vdom.set("")
        label6 = Label(frame1, text="\t\t    ")
        label6.grid(row=4, column=1, sticky='e', padx=3, pady=3)

        label7 = Label(frame1, text="\t\t\t")
        label7.grid(row=4, column=2, sticky='e', padx=3, pady=3)

def guardar_perfil1():
    global lista_perfiles
    lista_perfiles = []
    if nombre_perfil.get() == "":
        MessageBox.showwarning("Nombre Pérfil Vacio","Escriba un nombre de Pérfil")
    else:
        lista_perfiles.append(nombre_perfil.get())
        num_perfiles.set(1)

def guardar_perfil2():
    global lista_perfiles
    if nombre_perfil.get() == "":
        MessageBox.showwarning("Lista de perfiles vacio", "Escriba la lista de Pérfiles")
    else:
        lista_perfiles = nombre_perfil.get().split(',')
        texto.config(state="normal")
        texto.insert('insert', "\nLista de pérfiles:\n{}\n".format(lista_perfiles))
        texto.config(state="disable")
def perfiles():
    global lista_perfiles, num_perfiles, nombre_perfil
    lista_perfiles= []
    if opcion_perfiles.get() == 1:
        label8 = Label(frame1, text="   \t\t\t\t\n\t\tNombre del Pérfil:\t\n\t\t\n")
        label8.grid(row=6, column=0, sticky='e', padx=3, pady=3)

        entry7 = Entry(frame1, textvariable=nombre_perfil)
        entry7.grid(row=6, column=1, sticky='e', padx=3, pady=3)
        entry7.config(justify="left", state="normal")

        boton1 = Button(frame1, text="Guardar perfil", width=21, height=1, command=guardar_perfil1)
        boton1.grid(row=6, column=2, sticky='e', padx=3, pady=3)

    elif opcion_perfiles.get() == 2:
        label8 = Label(frame1, text="Ingresa una lista con los nombres \nde los Pérfiles, no dejar\n espacios despúes de las comas:\n Ejemplo: bajo,medio,vip")
        label8.grid(row=6, column=0, sticky='e', padx=3, pady=3)

        entry7 = Entry(frame1, textvariable=nombre_perfil)
        entry7.grid(row=6, column=1, sticky='e', padx=3, pady=3)
        entry7.config(justify="left", state="normal")

        boton1 = Button(frame1, text="Guardar", width=21, height=1, command=guardar_perfil2)
        boton1.grid(row=6, column=2, sticky='e', padx=3, pady=3)

# Variables
check_vdom = IntVar()           #Verifica si requiere vdom, 1= con vdom, 0 = sin vdom
name_vdom = StringVar()         #Guarda el nombre del vdom
name_hoja = StringVar()         #Nombre de la hoja del excel
columna_excel = StringVar()     #Columna de la hoja de excel, donde se encuentra el listado de url's
opcion_perfiles = IntVar()      # Opción que me indica número de perfiles
lista_perfiles=[]               #Lista donde se guardan los nombre de los pérfiles
num_perfiles = StringVar()      #Número de pérfiles
nombre_perfil = StringVar()     #Nombre del pérfil
ip_fortigate = StringVar()      #Dirección ip del Fortigate
user_fortigate = StringVar()    #Usuario de administrador del Fortigate
pass_fortigate = StringVar()    #Password de administrador del Fortigate
estatus_run = 0
estatus_script = 0



# Menu Superior
menubar = Menu(root)
filemenu = Menu(menubar,tearoff=0)
filemenu.add_command(label="Abrir lista en excel", command=abrir)
filemenu.add_command(label="Guardar script Fortigate", command=guardar_script)
filemenu.add_command(label="Guardar resultado excel", command=guardar_excel)
filemenu.add_separator()
filemenu.add_command(label="Salir", command=root.quit)
menubar.add_cascade(menu=filemenu, label="Archivo")


# Frame con los botones
frame = Frame(root)
frame.pack(side="left", fill="y", expand=0)
frame.config(bd=3, bg="#4e6af3")
Button(frame, text="Cargar Archivo excel", width=21,height=1,command=abrir).pack(anchor="w", padx=5,pady=5)

# Frame para la entrada de datos
frame1 = Frame(frame)
frame1.pack()

label = Label(frame1, text="Ingrese los siguientes Datos:")
label.grid(row=0, column=0, sticky='w', padx=3, pady=3)

label1= Label(frame1, text="Nombre hoja de excel")
label1.grid(row=1, column=0, sticky='e',padx=3, pady=3)

entry1 = Entry(frame1, textvariable=name_hoja)
entry1.grid(row=1, column=1, sticky='e', padx=3, pady=3)
entry1.config(justify="left", state="disable")

label2= Label(frame1, text="Columna de la hoja de excel")
label2.grid(row=1, column=2, sticky='e',padx=3, pady=3)

entry2 = Entry(frame1, textvariable=columna_excel)
entry2.grid(row=1, column=3, sticky='e', padx=3, pady=3)
entry2.config(justify="left", state="disable")

label3= Label(frame1, text="Dirección IP Fortigate:")
label3.grid(row=2, column=0, sticky='e',padx=3, pady=3)

entry3 = Entry(frame1, textvariable=ip_fortigate)
entry3.grid(row=2, column=1, sticky='e', padx=3, pady=3)
entry3.config(justify="left", state="disable")

label4= Label(frame1, text="User:")
label4.grid(row=3, column=0, sticky='e',padx=3, pady=3)

entry4 = Entry(frame1, textvariable=user_fortigate)
entry4.grid(row=3, column=1, sticky='e', padx=3, pady=3)
entry4.config(justify="left", state="disable")

label5= Label(frame1, text="Password:")
label5.grid(row=3, column=2, sticky='e',padx=3, pady=3)

entry5 = Entry(frame1, textvariable=pass_fortigate)
entry5.grid(row=3, column=3, sticky='e', padx=3, pady=3)
entry5.config(justify="left", show="*", state="disable")

check = Checkbutton(frame1, text="VDOM?", variable=check_vdom, onvalue=1, offvalue=0, command=c_vdom)
check.grid(row=4, column=0, sticky='e', padx=3, pady=3)

radio1 = Radiobutton(frame1, text="Un solo Pérfil", variable=opcion_perfiles, value=1, command=perfiles)
radio1.grid(row=5, column=0, sticky='e', padx=3, pady=3)
radio2 = Radiobutton(frame1, text="Más de un solo Pérfil", variable=opcion_perfiles, value=2, command=perfiles)
radio2.grid(row=5, column=1, sticky='e', padx=3, pady=3)

Button(frame, text="RUN", width=21, height=1, command=run).pack(anchor="w",padx=5,pady=10)
Button(frame, text="Guardar el script Fortigate", width=21, height=1, command=guardar_script).pack(anchor="w",padx=5,pady=5)
Button(frame, text="Guardar el resultado en excel", width=21, height=1, command=guardar_excel).pack(anchor="w",padx=5,pady=5)
Button(frame, text="ENVIAR", width=21, height=1, command=enviar).pack(anchor="w",padx=5,pady=10)

# Caja de texto Central
texto = Text(root)
texto.pack(fill="both",expand=1)
texto.config(bd=0, padx=5, pady=5, font=("Consolas",12), state="disable")

# Monitor Inferior
mensaje = StringVar()
mensaje.set("Bienvenidos a Url2ScriptForti")
monitor = Label(frame, textvar=mensaje, justify="left")
monitor.pack(side="bottom",anchor="w")

root.config(menu=menubar)

root.mainloop()