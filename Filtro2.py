import os
import csv
import tkinter as tk
from tkinter import filedialog
from collections import defaultdict
from datetime import datetime, timedelta

def browse_files():
    # Abre una ventana para que el usuario seleccione los archivos a buscar.
    files = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])
    return list(files)

def browse_folder():
    # Abre una ventana para que el usuario seleccione la carpeta donde se guardará el archivo resultados.csv.
    folder = filedialog.askdirectory()
    return folder

def calcular_fase_recordatorio(fecha_entrada):
    # Calcula la fase de recordatorio según la fecha de entrada
    hoy = datetime.today()
    
    try:
        # Ajuste para el formato de fecha "día, mes, año"
        fecha_entrada_dt = datetime.strptime(fecha_entrada, "%d/%m/%Y")
    except ValueError:
        print(f"Error en el formato de fecha: {fecha_entrada}")
        return "Fecha inválida"

    dias_desde_entrada = (hoy - fecha_entrada_dt).days
    
    if dias_desde_entrada <= 7:
        return "Primera fase: Invitación (primeros 7 días)"
    elif dias_desde_entrada > 7 and dias_desde_entrada < 30:
        return "Segunda fase: Reminder antes de que se acabe el mes"
    elif dias_desde_entrada >= 30 and dias_desde_entrada < 53:  # 7 días antes de que termine el primer mes
        return "Tercera fase: Se acabó el mes, recordar"
    elif dias_desde_entrada >= 53 and dias_desde_entrada < 60:  # 7 días después de que termine el mes
        return "Cuarta fase: Reminder antes de que se acabe el segundo mes"
    elif dias_desde_entrada >= 60 and dias_desde_entrada < 83:  # 7 días antes del segundo mes
        return "Quinta fase: Comunicación con su N+1"
    else:
        return "Quinta fase: Comunicación con su N+1"

def search_csv(files, folder):
    # Diccionario para almacenar qué debe cada persona junto con su fecha de entrada, correo, job function y fase de recordatorio
    pendientes = defaultdict(lambda: {'cursos': [], 'fecha': '', 'email': '', 'job_function': '', 'fase_recordatorio': ''})
    
    # Itera sobre cada archivo seleccionado
    for file in files:
        if file.endswith('.csv'):
            with open(file, newline='') as csv_file:
                reader = csv.DictReader(csv_file)  # Leer CSV como diccionario para acceder por nombre de columna
                curso_name = os.path.basename(file).replace('.csv', '')  # Usar el nombre del archivo como nombre del curso
                
                for row in reader:
                    if row.get('Completion Status', row.get('declaration_status', '')).lower() in ['not completed', 'open']:
                        # Verificar si existe la columna 'Name' o 'employee_name' y asignar
                        empleado = row.get('Name', row.get('employee_name', ''))  # Usar 'Name' o 'employee_name'
                        
                        if not empleado:  # Verifica si no se pudo encontrar el nombre
                            print(f"Advertencia: Nombre no encontrado para el archivo {file}.")
                            continue

                        # Obtener la fecha de entrada, email y job function
                        fecha_entrada = row.get('Entry Date', '')  # Cambiar por el nombre real de la columna
                        
                        # Buscar el email en 'Email' o 'employee_email'
                        email = row.get('Email', row.get('employee_email', ''))  # Captura 'Email' o 'employee_email'
                        
                        # Buscar el job function en 'Job Function' o 'job_function'
                        job_function = row.get('Job Function', row.get('job_function', ''))  # Captura 'Job Function' o 'job_function'

                        if not fecha_entrada:
                            print(f"Advertencia: Fecha de entrada no encontrada para el empleado {empleado} en el archivo {file}.")
                        
                        # Si la fecha de entrada no está ya guardada, se guarda
                        if pendientes[empleado]['fecha'] == '':
                            pendientes[empleado]['fecha'] = fecha_entrada
                        
                        # Si el correo no está ya guardado, se guarda
                        if pendientes[empleado]['email'] == '':
                            pendientes[empleado]['email'] = email
                        
                        # Si el job function no está ya guardado, se guarda
                        if pendientes[empleado]['job_function'] == '':
                            pendientes[empleado]['job_function'] = job_function
                        
                        # Calcular la fase de recordatorio y guardarla
                        if fecha_entrada:
                            pendientes[empleado]['fase_recordatorio'] = calcular_fase_recordatorio(fecha_entrada)

                        pendientes[empleado]['cursos'].append(curso_name)

    # Guardar los resultados en un archivo CSV
    with open(os.path.join(folder, 'resultados.csv'), 'w', newline='') as results_file:
        writer = csv.writer(results_file)
        writer.writerow(['Empleado', 'Email', 'Fecha de Entrada', 'Job Function', 'Cursos pendientes', 'Fase de Recordatorio'])
        
        # Escribir los resultados
        for empleado, info in pendientes.items():
            writer.writerow([empleado, info['email'], info['fecha'], info['job_function'], ', '.join(info['cursos']), info['fase_recordatorio']])

def button1_clicked():
    search_label.config(text="Iniciando búsqueda de pendientes...")
    files = browse_files()
    folder = browse_folder()
    search_csv(files, folder)
    search_label.config(text='Búsqueda completada. Resultados guardados en resultados.csv')

def search():
    # Crea la ventana principal
    window = tk.Tk()
    window.title("Opciones")
    window.geometry('250x100')
    option_label = tk.Label(window, text='Búsqueda de pendientes: ', padx=5, pady=5)
    option_label.pack()

    # Crea los botones y los agrega a la ventana
    button1 = tk.Button(window, text="Iniciar búsqueda", command=button1_clicked, border=3, foreground="gray18").pack()

# Crear la interfaz de usuario
root = tk.Tk()
root.title('Búsqueda de MOOCs y ECOI pendientes en archivos CSV')
root.geometry('400x200')

# Etiqueta de resultados
search_label = tk.Label(root, text='')
search_label.pack()

# Botón de búsqueda
search_button = tk.Button(root, text='Buscar pendientes', command=search, border=3, foreground="gray18").pack()

root.mainloop()
