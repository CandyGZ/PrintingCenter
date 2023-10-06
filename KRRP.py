import pandas as pd
import os
import tkinter as tk
import subprocess
import platform

ventana_principal = None
df = None  # DataFrame global para usar en las funciones


def archivo_mas_reciente(ruta):
    # Obtener la lista de archivos en la carpeta
    archivos = os.listdir(ruta)

    # Filtrar solo los archivos (excluir subdirectorios)
    archivos = [
        archivo for archivo in archivos if os.path.isfile(os.path.join(ruta, archivo))
    ]

    # Si no se encontraron archivos, retornar None
    if not archivos:
        return None

    # Obtener la ruta completa de cada archivo y su fecha de modificación
    archivos_con_fechas = [
        (archivo, os.path.getmtime(os.path.join(ruta, archivo))) for archivo in archivos
    ]

    # Encontrar el archivo más reciente basado en la fecha de modificación
    archivo_mas_reciente = max(archivos_con_fechas, key=lambda x: x[1])
    print(archivo_mas_reciente)
    nombre_de_archivo = os.path.join(ruta, archivo_mas_reciente[0])
    print("archivo mas reciente:", archivo_mas_reciente)
    # Retornar la ruta completa del archivo más reciente
    return nombre_de_archivo


# Especifica la ruta de la carpeta que deseas buscar
ruta_carpeta = "C:\\DataScience\\Python\\Proyectos\\Konika\\statusResults"

# Llama a la función para obtener el archivo más reciente
archivo_reciente = archivo_mas_reciente(ruta_carpeta)
print("archivo reciente:", archivo_reciente)
inicio_shortName = archivo_reciente.find("statusReport_")
shortName = archivo_reciente[inicio_shortName + len("statusReport_") : -5]

if archivo_reciente:
    print(f"El archivo más reciente en la carpeta es: {shortName}")

else:
    print("La carpeta está vacía o no contiene archivos.")


def no():
    global ventana_principal
    ruta_carpeta = "C:\\DataScience\\Python\\Proyectos\\Konika\\statusResults"
    if platform.system() == "Darwin":  # macOS
        subprocess.Popen(["open", ruta_carpeta])
    elif platform.system() == "Windows":  # Windows
        os.startfile(ruta_carpeta)
    ventana_principal.destroy()


def yes():
    global df
    if platform.system() == "Darwin":  # macOS
        subprocess.Popen(["open", archivo_reciente])
    elif platform.system() == "Windows":  # Windows
        os.startfile(archivo_reciente)

    # Carga los datos en un DataFrame y asigna a df
    df = pd.read_excel(archivo_reciente, sheet_name="status report")
    # print("DataFrame df:", df)  # Agregar esta línea para verificar df

    mostrar_ventana_emergente_2()


def mostrar_ventana_emergente_2():
    ventana_emergente = tk.Tk()
    ancho_pantalla = ventana_emergente.winfo_screenwidth()
    alto_pantalla = ventana_emergente.winfo_screenheight()
    ejeX = ancho_pantalla - ventana_emergente.winfo_reqwidth()
    ventana_emergente.geometry("+%d-600" % ejeX)
    ejeY = alto_pantalla / 2
    ventana_emergente.geometry("200x220+1500+%d" % ejeY)

    # Agregar el texto
    texto = f"Last 3 images printed:"
    label = tk.Label(ventana_emergente, text=texto)
    label.pack(padx=10, pady=10)

    # Agregar las entradas
    label_1 = tk.Label(ventana_emergente, text="1:")
    entry_1 = tk.Entry(ventana_emergente)

    # Agregar las entradas a la ventana
    label_1.pack(padx=10, pady=10)
    entry_1.pack(padx=10, pady=10)

    # Agregar el botón de cierre
    boton_send = tk.Button(ventana_emergente, text="Send", command=buscar_productos)
    boton_send.pack(side="bottom", padx=10, pady=10)

    # Agregar el código para obtener los datos de las entradas
    boton_send.config(command=lambda: obtener_datos(entry_1.get()))

    ventana_emergente.mainloop()


def obtener_datos(dato_1):
    print("de obtener datos: ", dato_1)
    buscar_productos(dato_1)


def buscar_productos(dato_1):
    # global df
    # print("de buscar_productos: ", dato_1, dato_2, dato_3)
    print("de buscar_productos", df)
    texto_busqueda = dato_1.strip()
    print(texto_busqueda)
    posicion = df[df["ProductID"] == texto_busqueda].index
    print(posicion)
    if (
        len(posicion) > 0
    ):  # si la búsqueda si encontró el texto de búsqueda en productID
        # Convierte la variable posicion a lista
        posicion_list = list(posicion)
        print(posicion_list)
        # Usa el método loc para obtener los elementos del df a partir de la posición encontrada
        elementos = df.iloc[posicion_list[0] :]
        print(elementos)
        # Guarda los elementos en un documento
        elementos.to_excel(
            "C:\DataScience\Python\Proyectos\Konika\statusResults\Delivery\Deliverables.xlsx",
            index=False,
        )
    else:
        print("No se encontró el resultado")


def mostrar_ventana_emergente(shortName):
    global ventana_principal
    ventana_principal = tk.Tk()
    ventana_principal.title("KRRP")
    ancho_pantalla = ventana_principal.winfo_screenwidth()
    alto_pantalla = ventana_principal.winfo_screenheight()
    ejeX = ancho_pantalla - ventana_principal.winfo_reqwidth()
    ventana_principal.geometry("+%d-600" % ejeX)
    ejeY = alto_pantalla / 2
    ventana_principal.geometry("200x100+1500+%d" % ejeY)

    texto = f"Do you need to PREP for: \n {shortName} \n ?"
    label = tk.Label(ventana_principal, text=texto)
    label.pack(padx=10, pady=10)

    boton_yes = tk.Button(ventana_principal, text="YES", command=yes)
    boton_yes.pack(side="left", padx=20)

    boton_no = tk.Button(ventana_principal, text="NO", command=no)
    boton_no.pack(side="right", padx=20)

    ventana_principal.mainloop()


# Llama a la función principal con shortName
mostrar_ventana_emergente(shortName)
