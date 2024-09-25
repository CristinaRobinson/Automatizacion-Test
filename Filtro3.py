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
        return f"Primera fase: Invitación (primeros 7 días) ({dias_desde_entrada} días desde la entrada)"
    elif dias_desde_entrada > 7 and dias_desde_entrada < 30:
        dias_desde_ultimo_recordatorio = dias_desde_entrada - 7  # Días desde que terminó la primera fase
        return f"Segunda fase: Reminder antes de que se acabe el mes ({dias_desde_ultimo_recordatorio} días desde el último recordatorio)"
    elif dias_desde_entrada >= 30 and dias_desde_entrada < 53:  # 7 días antes de que termine el primer mes
        dias_desde_ultimo_recordatorio = dias_desde_entrada - 30  # Días desde el último recordatorio
        return f"Tercera fase: Se acabó el mes, recordar ({dias_desde_ultimo_recordatorio} días desde el último recordatorio)"
    elif dias_desde_entrada >= 53 and dias_desde_entrada < 60:  # 7 días después de que termine el mes
        dias_desde_ultimo_recordatorio = dias_desde_entrada - 53  # Días desde el último recordatorio
        return f"Cuarta fase: Reminder antes de que se acabe el segundo mes ({dias_desde_ultimo_recordatorio} días desde el último recordatorio)"
    elif dias_desde_entrada >= 60 and dias_desde_entrada < 83:  # 7 días antes del segundo mes
        dias_desde_ultimo_recordatorio = dias_desde_entrada - 60  # Días desde el último recordatorio
        return f"Quinta fase: Comunicación con su N+1 ({dias_desde_ultimo_recordatorio} días desde el último recordatorio)"
    else:
        dias_desde_ultimo_recordatorio = dias_desde_entrada - 83  # Días desde la cuarta fase
        return f"Quinta fase: Comunicación con su N+1 ({dias_desde_ultimo_recordatorio} días desde el último recordatorio)"

def search_csv(files, folder):
    # Diccionario para almacenar qué debe cada persona junto con su fecha de entrada, correo, job function, work location, business group y fase de recordatorio
    pendientes = defaultdict(lambda: {'cursos': [], 'fecha': '', 'email': '', 'job_function': '', 'work_location': '', 'business_group': '', 'fase_recordatorio': '', 'declaration_status': '', 'declaration_expired': ''})
    
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

                        # Obtener la fecha de entrada, email, job function, work location, business group, y datos de ECOI
                        fecha_entrada = row.get('Entry Date', '')  # Cambiar por el nombre real de la columna
                        email = row.get('Email', row.get('employee_email', ''))  # Captura 'Email' o 'employee_email'
                        job_function = row.get('Job Function', row.get('job_function', ''))  # Captura 'Job Function' o 'job_function'
                        work_location = row.get('Work Location', row.get('work_location', ''))  # Captura 'Work Location' o 'work_location'
                        business_group = row.get('Business Group', row.get('bg', ''))  # Captura 'Business Group' o 'bg'

                        # Para ECOI, capturar los campos específicos 'declaration_status' y 'declaration_expired'
                        declaration_status = row.get('declaration_status', '')
                        declaration_expired = row.get('declaration_expired', '')

                        # Si la fecha de entrada no está ya guardada, se guarda
                        if pendientes[empleado]['fecha'] == '':
                            pendientes[empleado]['fecha'] = fecha_entrada
                        
                        # Si el correo no está ya guardado, se guarda
                        if pendientes[empleado]['email'] == '':
                            pendientes[empleado]['email'] = email
                        
                        # Si el job function no está ya guardado, se guarda
                        if pendientes[empleado]['job_function'] == '':
                            pendientes[empleado]['job_function'] = job_function
                        
                        # Si el work location no está ya guardado, se guarda
                        if pendientes[empleado]['work_location'] == '':
                            pendientes[empleado]['work_location'] = work_location
                        
                        # Si el business group no está ya guardado, se guarda
                        if pendientes[empleado]['business_group'] == '':
                            pendientes[empleado]['business_group'] = business_group
                        
                        # Guardar declaration_status y declaration_expired solo si está en estado "expired"
                        if declaration_status:
                            pendientes[empleado]['declaration_status'] = declaration_status
                        if declaration_expired.lower() == 'expired':  # Solo guardar si está expirado
                            pendientes[empleado]['declaration_expired'] = 'expired'

                        # Calcular la fase de recordatorio y guardarla
                        if fecha_entrada:
                            pendientes[empleado]['fase_recordatorio'] = calcular_fase_recordatorio(fecha_entrada)

                        pendientes[empleado]['cursos'].append(curso_name)

    # Guardar los resultados en un archivo CSV
    with open(os.path.join(folder, 'resultados.csv'), 'w', newline='') as results_file:
        writer = csv.writer(results_file)
        writer.writerow(['Empleado', 'Email', 'Fecha de Entrada', 'Job Function', 'Work Location', 'Business Group', 'Cursos pendientes', 'Fase de Recordatorio', 'Declaration Status', 'Declaration Expired'])
        
        # Escribir los resultados
        for empleado, info in pendientes.items():
            writer.writerow([empleado, info['email'], info['fecha'], info['job_function'], info['work_location'], info['business_group'], ', '.join(info['cursos']), info['fase_recordatorio'], info['declaration_status'], info['declaration_expired']])

def button1_clicked():
    search_label.config(text="Iniciando búsqueda de pendientes...")
    files = browse_files()
    folder = browse_folder()
    search_csv(files, folder)
    search_label.config(text='Búsqueda completada. Resultados guardados en resultados.csv')

def search():
    # Crea una nueva ventana
    window = tk.Toplevel(root)
    window.title("Opciones")
    window.geometry('350x150')
    window.configure(bg="#F1F1F1")  # Fondo moderno en gris claro
    
    # Agregamos una etiqueta de bienvenida con estilo
    option_label = tk.Label(window, text='Búsqueda de pendientes:', font=('Helvetica', 14), bg="#F1F1F1", fg="#333")
    option_label.pack(pady=10)
    
    # Botón de búsqueda con estilo moderno y bordes redondeados
    button1 = tk.Button(window, text="Iniciar búsqueda", command=button1_clicked, 
                        font=('Helvetica', 12, 'bold'), bg="#6C63FF", fg="white", 
                        activebackground="#3B3B98", activeforeground="white", padx=10, pady=5,
                        bd=0, relief="flat")  # Sin bordes
    button1.pack(pady=10)

    # Aplicamos un efecto "hover" al botón
    def on_enter(e):
        button1.config(bg='#5A55FF')  # Cambia a un color más claro al pasar el cursor

    def on_leave(e):
        button1.config(bg='#6C63FF')  # Regresa al color original al quitar el cursor

    button1.bind("<Enter>", on_enter)
    button1.bind("<Leave>", on_leave)

# Crear la interfaz de usuario
root = tk.Tk()
root.title('MOOCs y ECOI Tracking')
root.geometry('400x300')
root.configure(bg="#2C3E50")  # Fondo de la ventana principal en azul oscuro

# Agregamos un título moderno
title_label = tk.Label(root, text="Seguimiento de MOOCs y ECOI", font=('Helvetica', 18, 'bold'), bg="#2C3E50", fg="white")
title_label.pack(pady=20)

# Etiqueta de estado de búsqueda
search_label = tk.Label(root, text='', bg="#2C3E50", fg="white")
search_label.pack()

# Botón de búsqueda principal con bordes redondeados y efecto hover
search_button = tk.Button(root, text='Buscar pendientes', command=search, 
                          font=('Helvetica', 14), bg="#3498DB", fg="white", 
                          activebackground="#2980B9", activeforeground="white", padx=10, pady=10,
                          bd=0, relief="flat")  # Sin bordes
search_button.pack(pady=20)

# Aplicamos un efecto "hover" al botón de búsqueda principal
def on_enter_search(e):
    search_button.config(bg='#58A6FF')  # Cambia a un color más claro al pasar el cursor

def on_leave_search(e):
    search_button.config(bg='#3498DB')  # Regresa al color original al quitar el cursor

search_button.bind("<Enter>", on_enter_search)
search_button.bind("<Leave>", on_leave_search)

# Ejecución de la ventana principal
root.mainloop()



-----//////









