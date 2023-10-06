# User Manual for the Printing Control System   (Go down for Spanish)

## General Description
The printing control system is a tool designed to manage printing jobs in a printing center. It allows you to find the most recent file with information on pending jobs, enter the last job completed, and generate a report of pending jobs for the next day's work.

## Steps to Use the System

### 1. Program Startup
- Run the program.
- The system will automatically search for the most recent file with information on pending jobs in the specified folder. This file contains information about the jobs awaiting printing.

### 2. Viewing the Last Job
- Once the most recent file is found, the program will display the file name. This should be the last document you worked with during the day (where additional documents were added for printing).

### 3. Confirming the Last Job
- The program will ask if you want to confirm that this is the file you are working with today.
- If it's correct, select "YES," the file will be opened, and you will proceed to step 4.
- If it's not correct, select "NO," your file explorer will open, allowing you to manually open the document of the day.

### 4. Entering Completed Jobs (Optional)
- If you selected "YES," the program will allow you to enter the last job completed. This is done in a pop-up window.
- In the pop-up window, enter the full name of the file exactly as it appears in ProductID.
- Click the "Send" button to confirm the information.

### 5. Report Generation
- After confirming the last job, the system will generate a report of pending jobs that can be used, for example, the next day.
- This report will be saved in a file called "Deliverables.xlsx" in the specified folder.
- The report will include an updated list of pending jobs, excluding jobs that have already been completed.

## Important Notes
- Make sure that the pending job files are located in the specified folder.
- If the system does not find any files in the folder, it will display a message that the folder is empty.
- It is important to correctly enter the name of the last job completed in the pop-up window.
- The system will save the report in Excel format (.xlsx) without including indexes in the file.

That's it! With these steps, you can effectively use the printing control system to manage pending jobs at your printing center. If you need something similar adapted for your printing center, please feel free to reach out.

Pandas: Used to read and manipulate Excel data.
The archivo_mas_reciente() function uses Pandas to read the Excel file and get the modification date of each file.
The buscar_productos() function uses Pandas to search for a product in the Excel file and get the corresponding data.
Tkinter: Used to create a pop-up window.
The mostrar_ventana_emergente() function uses Tkinter to create a pop-up window with two buttons, one to answer "Yes" and one to answer "No".
subprocess: Used to open the Excel file or the folder containing the Excel files.
The no() function uses subprocess to open the folder containing the Excel files.
The yes() function uses subprocess to open the most recent Excel file.
os: Used to get the path to the most recent Excel file.
The archivo_mas_reciente() function uses os to get the path to the most recent Excel file.
platform: Used to determine the user's operating system.
The no() function uses platform to determine the user's operating system and open the folder containing the Excel files or the most recent Excel file, depending on the operating system.


***************************************************************************************************

# Manual de Uso del Sistema de control de Impresiones   

## Descripción General
El sistema de impresión es una herramienta diseñada para gestionar los trabajos de impresión en un centro de impresión. Permite encontrar el archivo más reciente con información de trabajos pendientes, ingresar el último trabajo realizado y generar un reporte de trabajos pendientes para trabajar al día siguiente.

## Pasos para Usar el Sistema

### 1. Inicio del Programa
- Ejecuta el programa.
- El sistema buscará automáticamente el archivo más reciente con información de trabajos pendientes en la carpeta especificada. Este archivo contiene la información sobre los trabajos que están pendientes de impresión.

### 2. Visualización del Último Trabajo
- Una vez encontrado el archivo más reciente, el programa mostrará el nombre del archivo. Este debería ser el último documento con el que se ha trabajado durante el día (donde se han agregado más documentos a imprimir).

### 3. Confirmar el Último Trabajo
- El programa preguntará si deseas confirmar que este es el archivo donde están trabajando el día de hoy.
- Si es correcto, selecciona "YES" (SÍ), se abrirá el archivo y pasarás al paso 4.
- Si no es correcto, selecciona "NO", se abrirá tu explorador de archivos y te permitirá abrir el documento del día de forma manual.

### 4. Ingreso de Trabajos Realizados (Opcional)
- Si seleccionaste "YES", el programa te permitirá ingresar en la celda el último trabajos realizado. Esto se hace en una ventana emergente.
- En la ventana emergente, introduce el nombre completo del archivo tal como aparece en ProductID.
- Haz clic en el botón "Send" (Enviar) para confirmar la información.

### 5. Generación de Reporte
- Después de confirmar el último trabajo, el sistema generará un reporte de trabajos pendientes que se podrá usar por ejemplo al día siguiente.
- Este reporte se guardará en un archivo llamado "Deliverables.xlsx" en la carpeta especificada.
- El reporte incluirá la lista de trabajos pendientes actualizada, excluyendo los trabajos que ya han sido completados.

## Notas Importantes
- Asegúrate de que los archivos de trabajo pendiente se encuentren en la carpeta especificada.
- Si el sistema no encuentra ningún archivo en la carpeta, mostrará un mensaje de que la carpeta está vacía.
- Es importante ingresar correctamente el nombre del último trabajo realizado en la ventana emergente.
- El sistema guardará el reporte en formato Excel (.xlsx) sin incluir los índices en el archivo.

¡Listo! Con estos pasos, podrás utilizar eficazmente el sistema de impresión para gestionar los trabajos pendientes en tu centro de impresión.
Si necesitas adaptar algo parecido para tu centro de impresión quedo a tus órdenes.

Pandas: se utilizó para leer y manipular los datos del archivo Excel. La función archivo_mas_reciente() utiliza Pandas para leer el archivo Excel y obtener la fecha de modificación de cada archivo. La función buscar_productos() utiliza Pandas para buscar un producto en el archivo Excel y obtener los datos correspondientes.
Tkinter: se utilizó para crear una ventana emergente. La función mostrar_ventana_emergente() utiliza Tkinter para crear una ventana emergente con dos botones, uno para responder "Sí" y otro para responder "No".
subprocess: se utilizó para abrir el archivo Excel o la carpeta que contiene los archivos Excel. La función no() utiliza subprocess para abrir la carpeta que contiene los archivos Excel. La función yes() utiliza subprocess para abrir el archivo Excel más reciente.
os: se utilizó para obtener la ruta del archivo Excel más reciente. La función archivo_mas_reciente() utiliza os para obtener la ruta del archivo Excel más reciente.
platform: se utilizó para determinar el sistema operativo del usuario. La función no() utiliza platform para determinar el sistema operativo del usuario y abrir la carpeta que contiene los archivos Excel o el archivo Excel más reciente, según el sistema operativo.
