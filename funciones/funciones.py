# En este módulo se definen las funciones logicas del programa

import openpyxl  # Libreria para tratamiento de documentos excel
import urllib.request  # Libreria para request a paginas web
import socket
import time
import sys
from io import StringIO
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from tkinter import messagebox as MessageBox


global lista, lista1, req_url, i, lista_st,doc;
code = 300
lista_st = []


# Función para verificar el estatus de una url
def get_code(url):
    global code, req_url
    code = 260
    try:

        req = urllib.request.Request(url)  # Crea el request
        req.add_header('User-Agent',
                       'User_Agent')  # Configurar la cabecera del request, para que en ciertas paginas acepten el request
        req_url = urllib.request.urlopen(req, timeout=1.8)
        url1 = req_url.geturl()
        code = req_url.getcode()
        req_url.close()

    except urllib.error.HTTPError as e:
        try:
            code = e.code
        except:
            pass
    except urllib.error.URLError as e:
        # e.reason()
        code = 600

        try:
            print(str(e.reason()))
        except:
            pass
        req_url.close()
    except socket.timeout:
        # write more crap to the logs
        code = 408
        req_url.close()

    else:
        req_url.close()
        pass


# Función para filtrar y depurar el listado de url's
def filtro_url():
    global req_url, lista1, k
    j = 0
    k = i
    for url in lista1:
        url1 = str(url)
        http1 = str(url[:7])
        http2 = str(url[:8])
        print(url)
        if http1 == "http://":
            get_code(url1)
            if code < 200 or code >= 404:
                del lista[j]
                j -= 1
                k -= 1
        elif http2 == "https://":
            get_code(url1)
            if code < 200 or code >= 404:
                del lista[j]
                j -= 1
                k -= 1
        else:
            get_code("http://" + url)
            if code < 200 or code >= 404:
                del lista[j]
                j -= 1
                k -= 1
            else:

                lista[j] = "http://" + url1
        if code == 408:  # Serie de IF anidados para crear la lista con los codigos de estatus de las url's
            print(str(code) + ": Timeout de 1.4 seg\n")
            lista_st.append("Timeout de 1.4 seg")
        elif code == 600:
            print(str(code) + ":Error de Url, no existe\n")
            lista_st.append("Error de Url, no existe")
        else:
            print(str(code) + "\n")
            lista_st.append(code)

        j += 1

    print("Número de url's buenas:", k)

# Función para abrir y cargar el archivo de excel

def abrir_excel(ruta,texto):
    global doc

    texto.config(state="normal")
    doc = openpyxl.load_workbook(ruta)  # Se carga el archivo en excel con su respectiva extensión .xlsx
    texto.insert('insert',"A continuacion se imprimiran las hojas existentes en el documento de Excel:\n")
    for nombres in doc.get_sheet_names():
        texto.insert('insert',str(nombres)+"\n")
    texto.config(state="disable")

# Función para cargar la hoja y crear la lista
def lista_hoja(doc, hoja,columna):
    global lista, lista1, i, k

    num_col = columna
    # source = doc.active
    # hoja2 = doc.copy_worksheet(source)															#Se hace una copia de la hoja de excel donde esta el listado de url's
    # hoja2.title = "Estatus Url"																	#Se cambia el nombre de la hoja de excel
    hoja2 = doc.create_sheet()
    hoja2.title = "Estatus Url"

    # For loop anidados para extraer la columna con el listado de url's
    i = 0  # La variable i es para contar el número de url's
    lista = []  # La variable lista se guarda todas las url's extraidas del documento de excel
    lista1 = []
    for fila in hoja.rows:  # Primero for loop para recorrer fila por fila
        for columna in fila:  # Segundo for loop para recorrer por columna
            if num_col == columna.coordinate[
                0]:  # Condicional if para solo agregar los valores de la columna escogida
                if columna.value == None:
                    break
                lista.append(columna.value)  # Se va creando la lista de todas las url's
                lista1.append(columna.value)
                i += 1  # Contador de filas o de número de url's
    if i != 0:
        filtro_url()  # Llamamos la funcion para que filtre el listado de las url no existentes o que fallan
        # Codigo para crear una columna de excel donde se ubicaran el codigo de estatus del listado de Url's
        for k1 in range(0, i + 1):
            if k1 == 0:
                hoja2.append(["Listado de URL's", "Estatus de la URL"])
            elif k1 > 0:
                hoja2.append([str(lista1[k1 - 1]), str(lista_st[k1 - 1])])
        doc.save("Url Filter Estatus.xlsx")
    else:
        k = i
    return i,k

# Código para crear el archivo de texto, el cual contendra el script con el listado de url's para agregar al Fortigate
def crear_script(vdom1, lista_perfiles,ruta):
    try:
        vdom = vdom1
        f = open(ruta, 'w', encoding='cp1252')  # Abre o crea el archivo de texto llamado script_url_filter.txt
        if vdom != "":
            f.write("config vdom\n")
            f.write("edit " + vdom + "\n")
        f.write("config webfilter urlfilter\n")
        f.write("    ")
        f.write("edit 126\n")
        f.write("        ")
        f.write('set name ')
        f.write('"')
        f.write('"Lista126"')
        f.write('"\n')
        f.write("        ")
        f.write("config entries\n")

        # For loop para crear el script en archivo de texto con el listado del url filter para Fortigate
        for url1 in lista:
            f.write("            "), f.write("edit 0\n")
            f.write("                "), f.write('set url "'), f.write(url1), f.write('"\n')
            f.write("                "), f.write("set action block\n")
            for j in url1:
                if j == '*':
                    f.write("                "), f.write("set type wildcard\n")
            f.write("            "), f.write("next\n")
        f.write("        "), f.write("end\n")
        f.write("    "), f.write("next\n")
        f.write("end\n")
        # For para adicionar el istado de url's a los perfiles web
        if lista_perfiles == []:
            pass
        else:
            f.write("config webfilter profile\n")
            for wf in lista_perfiles:
                f.write('\tedit "{}"\n'.format(wf))
                f.write('\t\tconfig web\n')
                f.write('\t\t\tset urlfilter-table 126\n')
                f.write('\t\tend\n')
                f.write('\tnext\n')
        if vdom != "":
            f.write('end\n')
            f.write('end')
        else:
            f.write('end')

    finally:
        f.close()

# pagina con timeout www.tecnocam.com.ve/gymnetwork
# pagina no existe http://prod.msocdn14.com
# pagina not found 404  http://agent.office.net

def send_forti(ip_forti,user_forti,pass_forti,ruta_script):
    sys.stdout.flush()
    script = ruta_script  # Ruta donde esta ubicado elscript Script_Url.txt
    url = ip_forti
    user = user_forti
    pass1 = pass_forti
    driver = webdriver.Chrome()
    driver.get(url + "/ng/system/advanced")
    time.sleep(1.0)
    try:
        elem1 = driver.find_element_by_link_text("CONFIGURACIÓN AVANZADA")
    except NoSuchElementException:
        pass
    else:
        elem1.click()  # Si aparece mensaje de seguridad por certificado no valido
        time.sleep(1)
        elem1 = driver.find_element_by_link_text("Acceder a {} (sitio no seguro)".format(url[8:]))
        elem1.click()
    time.sleep(1.0)
    username = driver.find_element_by_name("username")
    password = driver.find_element_by_name("secretkey")

    username.send_keys(user)
    time.sleep(0.5)
    password.send_keys(pass1)
    time.sleep(0.5)
    driver.find_element_by_name("login_button").click()
    time.sleep(4)
    driver.find_element_by_css_selector('input[type="file"]').send_keys(r"{}".format(script))
    time.sleep(4)
    driver.get(url + "/login")
    time.sleep(2)
    driver.quit()  # Cierra todas las ventanas y todos los procesos del webdriver