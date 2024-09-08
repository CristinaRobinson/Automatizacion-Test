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



-----


Ahora recibió una modificación de estructura bajo su responsabilidad hace un mes.
Justo ahora tiene la misma responsabilidad al mismo tiempo que cuando estaba en VW. 
Recibió en su estructura y su responsabilidad
Tiene otras dos responsabilidades
Tiene una menor estructura en SLP

Desde 2019 tiene bastante
¿Cuál fue tu anterior empleo y educación?
Trabajo en baleo, cuando termino su carrera y continuo con external tractor por un año
Se movió a queretaro por 6 año y 1 mes
Quality engineer customer service

Se quedo en otra planta 1 año y 6 meses
Ahora recibió una llamada que es posible tener un quality manager pero tiene

Bagela le dijo que ahora mismo no tiene contacto para SLP

¿Cuál es tu relación de trabajo con Joaquin MaRTINEZ?
Ahora es suN+2 Quality director
Mora es su N+1

Antes fue su N+1

¿Hace cuanto es tu manager?
Con Joaquin 1 año y 1 mes o 2 meses


¿Has trabajado con ellos un rato, los conoces bien?
Si, no es relevante compararlos con los demás, pero trabajaron con el en diferentes… es mas blanco y negro. Es la admin que mas le ha dejado mas enseñanza para que sea profesionalmente, te protege, tiene coach y es lo que buscas de un N+1
En general es un buen manager y es el tipo de persona que te hace ser mejor, PEDRO en general es un buen líder.

¿Como ha sido tu relación de trabajo con Gabriel estrada?
-	En el pasado solo tuvieron una comunicación y la última comunicación es referente a Joaquín y otros colegas que ha tenido.
-	De que te hablo ella de Joaquín.
-	De tareas, la discusión con Gabriela estrada fue referente a la Srta. Abigail Gallo de preferencias a su puesto de trabajo y si tenia alguna relación con Joaquín Martínez extra oficial

¿Gabriela te dijo por que te estaba potenciando de esta relación?
Alguien mas empezó un canal de investigación y fui llamado con las siguiente preguntas si esta persona recibió cierta preferencia en el trabajo? No.
Tienes confirmación de que etsa persona tiene algo de relación y fue favorecida, tu te has visto afectado por esta potencial relación? Si, no.
Tipos de lideres no deberíamos discutir este tema.
La sugerencia es levantar un ticket de ICB

¿Habias visto un favoritismo o te viste afectado?
No, cuando comente el tema de Abigail Gallo, fue que algunos mas del equipo es que había sido favorecida ella con el tema de carga laboral a través de diferencia de complejidad en su carga laboral. Cuando el conteste dije obviamente no es cierto, tenia una matriz, tengo el archivo de matriz de complejidad de clientes y evaluamos ocurrencia, y cantidad de cuentas VW y audi son 7 cuentas.
En esa parte esta dividida la complejidad de las cuentas y te puedo mostrar, se ve minimizada la carga de trabajo. Pero el resto de las cuentas tienen menos complejidad.
Si bien había una carga diferente de trabajo era por la naturaleza de las responsabilidades que tenia cada quien.

¿Supiste si mas personas del equipo fueron entrevistadas?
Al ser confidencial, mi interpretación es que alguien mas al menos fue entrevistada. No tengo confirmación pero lo describió globalmente. Este favoritismo ha afectado al resto del equipo.
-	Clau: ¿Como si ella estuviera segura de que estaba sucediendo?
Si, porque en mi cabeza decía cuantos mas de mi equipo, no puede ser del otro lado tiene que ser de mi equipo, yo solo tengo otro costumer Diego SILVA levanto la mano por diferencia de cargas.

-	Clau: Tenemos entendido que Gabriela después de esa reunión te mando el link, es correcto?
Al cierre de la entrevista? Fue mas de una hora y dije bueno voy a levantar el ticket, en toda la entrevista dimos vueltas, estamos favorecidos. 
Después hubo una llamada o recordatorio de ¿ya levantaste el tcicket? Y dije no

Gabriela estaba insistiendo si había algún beneficio de Joaquín sobre esta mujer, el resto del team sienten que valen menos que los otros, fue incomodo recibir preguntas una y otra vez, con el objetivo de detener la conversación pedro dijo al final, voy a reportar esto con la intención de que no quiere ser mas cuestionado, después recibió una pregunta que Gabriela mando un mensaje si el reporto el problema.

¿Puedes mandarnos este email sobre lo que Gabriela te mando?

¿Tienes algún sentimiento sobre porque ella te preguntaba por que hacer eso?
Gabriela empezó la entrevista diciendo que ella recibió información de alguien en el equipo y era su responsabilidad hacer esa entrevista para confirmar las allegations. 

¿Tienes idea de por que ella te estaba acusando en alguna dirección, suena como si te estuviera empujando?
En su cabeza solo decía que las personas que habían hablado previamente econ ella habían hecho el reporte. Había algo mas que le habían dicho, al yo decir que tenia una matriz de complejidad. Estaabn divididos en base a lo que estaban viviendo y en base a las funciones. Para mi no hay favoritismo. Había algo mas que habían mencionado. Regreso con Gabriela

El piensa que fue porque una persona o mas personas se acercaron a Gabriela para reportar la situación.

¿Piensas que fue algo mas personal de parte de Gabriel o fue porque recibió la información de ciertas personas?
No estoy seguro

¿Consideras que lo de Gabriela fue mas en contra de Joaquín?
No puedo confirmarlo, no conozco a Gabriela

¿Tienes también la invitación?
-	Calendario
-	Revisión
-	Cuando te escribió diciéndote que tienes que levantar el ticket

¿Para cuándo tu recibiste aquella invitación tus sabias que iba a ser para una entrevista por parte de Gabriel?
No recuerda. Tal cual así fue llamado, tema de favoritismo.

Pedro dice que el no esta segura si fue algo personal, pero Gabriel insistió cuestionando una y otra vez lo mismo, Gabriela recibió la queja de alguien más, Pedro no esta seguro si hay algo personal porque el No sabia si fue al segunda vez que Gabriela discutió con él. Sabe que Gabriela insistió, pero no personal

¿Trabajaste con Junior García?
Si, tenía la misma posición que yo, pidió un cambio a puebla como Plant Manager, pero me quede con el trabajando en SLP cada vez.

La transferencia de tesla va a ser para el siguiente periodo en el siguiente periodo, el se quedo en MTY como quality manager, viaje a monterrey una vez.

¿Sabes si el tenia buena o mala relación con Joaquin?
Nunca escuche ningún problema entre Joaquin. MTY tiene la misma posición que SLP como Quality Manager.

¿Alguna relación que hayas escuchado entre Junior Garcia y Gabriel?
Vi a Junior solo 2 veces, la primera vez fue en el mismo periodo cuando inicie la entrevista como Quality Manager, es una competición externa.
Nunca escuche nada de Junior García y Gabriela.

¿Había alguna relación entre Gabriela y Junior García?
-	Fue gracioso para mi escuchar, que Junior García y Gabriela, la ves que estuvo en FRAMES En MTY, no tengo idea de Junior García y Gabriela.
-	Le dije cara a cara voy a levantar el ticket, pero no lo hice tuve mas confianza de ir con Joaquín y decirle está pasando esto.

¿A quien crees que tenemos que entrevistar?
-	Frances MORA
-	Gabriela
No trate de entender quien levanto la mano, me quede callado como 2 meses y después tuve la comunicación de decirle a Joaquín lo que pasó. 

¿Ramses MORA también esta entero de que Gabriel te entrevisto?
Yo creo que no, no tengo la evidencia. Yo creo que el fue primero antes que yo.

Convendría hablar con él para saber si participo en una entrevista con Gabriela.
Yo considero que él fue primero, salió el, entre yo, creo que el fue la primera persona.

Dio un correo como Compliance Interview. Pero ella no hace investigaciones, se dedica al Development.
Realmente no debí haber tomado esa serie de preguntas con Gabriela, no era la persona indicada.
Empezamos con Abigail Gallo y terminamos con el tema de Compliance.

La posición de Gabriela es para desarrollo, es cierto que las personas que son potenciales para desarrollo tengan las habilidades y no haya un conflicto de interés, lo que no hay sentido es presionar a hacer una investigación

Es importante saber el motivo principal de esta platica contigo.

Ahora entiendo que Abigail estaba como “” y después caimos en el tema de Joaquín, una renuncia de Abigail, ,

Clau:
Estoy sorprendida que Pedro fue presionado por Gabriel.

¿Qué piensas del problema con Gabriela?
No se mas de lo que te he dado. La dificultad es de que Ana es una president muy dura.
Fue hace una semana el viernes que 
La historia viene de nuestro presidente Ana y el VP de ingeniería Everson y Ana, es difícil describir. La escucho y me pregunto por que? 
No conozco a Pedro, nunca conocí a Joaquin, lo que le intentaba explicar a Ana, que no es de confiar pero no confía en HR.
Tenemks que tener eso en mente cuando tengamos que entender esto.

Cuando inicio esta conversación tuve que ser muy políticamente correcta en ser objetiva y no estar a la defensiva.
Cuando iniciamos la llamada, solo para que sepas tengo una investigación de una investigación
Pregunte, ¿es sobre mi?
Dije, no creo que es sobre ti, no se de que es.

Siempre esta sospechando y muy sospechosa de HR, no cree en el mundo donde los empleados tienen el derecho de ir con HR
Habia un email de Gabriela a Pedro, siguiendo nuestra discusión: [link speak up]
Cuando lo enseño estaba molesta, y dije tal vez pedro tuvo una reunión con Gabriel.
Sharon, mira, un mensaje de seguimiento y dije, no hay nada que lo implique que esta siendo forzado.

No vi nada que dijera algo de Gabriela inapropiado, explique objetivamente.
Tiene que haber una investigación antes de que sepamos que hay un problema.

Tiene una relación con el Quality Manager y creo que no le gusta Joaquin y no quiere que sea promovido.
Esto es preocupante, esta podría ser la razón.

SHARON:

Le dije Ana, lo tomo en serio pero para poder hacer algo necesito hablar con Everson y pedro, tal vez. No hagas nada no le hagas saber a Everson que hablo contigo.

Fue hace 3 semanas en un lunes y no escuche de everson y ana. Fue hace una semana everson la contacto por una razón y Gabriel y el estaban en MTY, después le dije a el, ana me mencioo que hay un problema que te preocupa.
No te preocupes Gabriela esta aquí, le enseña el clossur document de que el caso esta cerrado.
¿Qué esta pasando?
Nada, hay que verlo mañana.

No escuche nada y después creo que fue el marte-meircoles tengo un email de everson
Gabriela no ___ clossure document.
Esa noche respondi su mensaje y dije, tienes razón ella no dio los documentos.
Tomo el asunto en serio voy a tener en cuenta el asunto para tener la investicacion.
Ana responde: Gracias Sharon pero esto no es lo que pide, te pide que investigues a Gabriela, es lo que necesita.
Después me manda un mensaje de teams, tenemos información relevante que tenemos que darle a los investigadores.

Con Ana todo importa, tuve que sentarme varios minutos decidiendo como quería responer, porque como HR Manager tengo que ser tomada en serio.
Ana tengo toda la intención de manejar esto de manera profesional.

Ana: Si es genial es bueno pero tenemos información relevante ahora.

Ana esta haciendo suposiciones de lo que piensan.

Everson y yo tenemos que hablar con los investigadores y tener la información relevante y pone.
Gabriela le esta poniendo trampa para quitarlo de encima.

Ana, entiendo tu preocupación mañana en la mañana cuando vaya a la oficina voy a contactar a la fuente correcta, vino el viernes y ella había hablado con Everson.
Everson estaba completamente mal, siguiéndome, acosándome
Le digo, que necesitas?
“Tu, Everson iniciaste todo esto”

Ahora tiene miedo y esta paniqueado y dice okay no necesitamos ningún investigador.

No se cual es el problema de Ana contra Gabriela.
Ana siempre tiene sospechas, tenemos que tener cuidado charon, se que piensas que tu equipo es bueno pero nos van a traicionar nos van a reportar
Y dijo, tu no entiendes a los empleados mexicanos y nos van a traicionar.

Ana se ha comportado poco profesional en una reunión con el presidente. También estaba negativa.

Everson quiere mandarme el email que Gabriela le mando a pedro con el link. Le dije, sabes que? Tu y Ana solicitaron una investigación no tienen que mandarme nada a mí.

Esta mentalmente en un lugar extraño. También conoce a Ana y dijo, tienes razón.
Esto es lo que te puedo decir de ellos dos, no conozco a Gabriela ni Pedro.

Junior Garcia ya había renunciado hace 2-3 meses. 
Everson es el que le comento a Ana.

Sharon en junio 18 estaba en MTY y el viernes del 21 y te dije estaba en una conferencia con Gabriela, Mariana y Sergio y escuche de lejos que dijeron que Everson quiere forzarnos con la promoción, tuvimos que quejarnos.

Hay varias cosas en los correos que no tiene relación
Sergio le dijo a Joaquín que la investigación estaba terminada y que no había más problemas.
Everson dijo creo que Ana está haciendo un cochinero aquí, esta haciendo tres historias distintas.
Gabriel no mando el documento que está cerrado.

Sergio cerro el caso y no estaba cerrado.
Hay una investigación abierta de Joaquín.
Gabriela  __?

Tal vez Joaquin estaba saliendo con una subordinada.

Primero dice que la investigación fue cerrado
Pero después dice que la investigación no es necesitada.

“No fue un reporte oficial de compliance”

Sergio sabe que sus maneras no fueron las correctas.

Joaquín no esta siendo bueno conmigo (con quién)?

3 diferentes emails, diciendo que no se necesita investigación, no estaba al tanto. NO ES CORRECTO.
No se siguió un protocolo formal.







Joaquín:
He trabajado en forma pro 7 años, trabaje como Quality Manager en una de las plantas.
Después tuve la promoción para se Manager. Estaba viendo 5 plantas y varios suppliers.
Solicite ser movido a MTY.
Cuando recibí ser movido a MTY la rechace porque no estaba de acuerdo con la letra.

Trabaje después en MTY 3 meses, al final una de las plantas que es una de las mas criticas en FRAMES PLANTS
Al final del año fue muy complicado, sin embargo, tuvimos éxito.

Ramses es el senior quality manager.
Ahora estoy tomando de Quality directos de Mexico Region.

¿Entiendo estas de VP hace como 1 mes?


¿Cuál era u relación con PEDRO Delgado?
Soy N+2 de PEDRO Delgado cuando era senior quality manager el era uno de mis quality managers.

¿Nos darías información de la meeting que tuvo PEDRO?

En cuanto al favoritismo:
Pedro explica que Abigail estaba respondando a Pedro, no a el.
Gabriela estaba questionando si pusieron esta situacion 

Vieron que esta persona tenia potencial y que estaba haciendo un buen trabajo.
Le pregunto a PEDRO por que recibiste etse tipo de emails diciendo que ella esta esperando a crear el ticket, pedro dijo que después en esta reunión pedro dijo que hará un ticket pero necesitaba tiempo q ue con este comentario se escapo de la conversación incomoda.
Al final PEDRO No abrió el ticket, pedro recibió el mismo día que paso el fase to fase de Gabriela diciendo Pedro, como comentamos hoy, te necesitamos aquí, te dejo aquí el link para el speak up para que puedas hacer el reporte, dime si necesitas apoyo.

Pedro no hizo el ticket, en Julio 11 pedro recibió otro email.
Hola pedro buenas tardes, solo quiero un follow up de este item, al final hiciste el reporte? Y si si tuviste notificación de los next steps?

No he abierto el reporte, confirmare después de JULY 24 después hable con Everson mi N+1 porque PEDRO dice no se cuales son las consecuencias si te digo estas cosas Joaquín porque no se si vas a decir algo, no se si te vas a quedar callado. Por eso Joaquín reporto la situación a Everson.

¿Cómo ha sido tu relación con Gabriela, has trabajado con ella en los últimos 7 años?
Si, la conozco desde hace 4 años estaba participando en las potencial reviews?.

Toda la situación durante lo de monterrey fue con una persona de HR, no se si detrás de eso estaba Gabriel, pero no tuve ninguna conversación con Gabriela en Monterrey.

¿Tus interacciones fueron relacionadas a tu desarrollo en los últimos años?
No, no identifique ninguna interacción con Gabriela. Nos hablamos normal.

¿Pedro menciono o especulo por que Gabriela quería que hiciera eso o  tuviste algún pensamiento de cual podría ser la razón?
Si tuve algunos pensamientos, pero no tuve hechos para afirmar.
Después del 30 de Julio Everson presiono a RH para que le dieran la carta a Joaquin.
El 11 de Julio fui a la oficina de Sergio y dijo que hay una investigación en contra de ti, tienes que esperar.
Pedro contesta el viernes, 12 de julio que no se ha levantado nada.

Comparti la información con Everson de lo que estaba pasando.
Abigail la persona que se supone que debe tener favoritismos, estaba en un potencial numero 2 de 
En forvia si tienes potencial potencial, tienes potencial number 1, 2 o 3
Numero 2 es vas conseguir 
Cuando fue nominada como potencial número 2, no fue solo ella fueron 3 personas, Abigail renuncio a la planta porque un quality manager llego, no se sintió cómoda y renuncio, en ese momento ella tenia Costumer Service Relationship. Se empieza a investigar porque renuncio, el nuevo QM para mi Abigail no tiene el potencial que debería
Para mi esto fue lo que hizo a Abigail renunciar.
¿Quién es el de Quality Manager que no le gusto su actitud?
Ramses

Puntos:
•	Abigail nunca pidió ser Quality Manager, pidió renunciar porque vive en Querétaro, dice que renunciara, les avisa.
•	Y le decimos: Abigail espera y después te puedes mover.
•	Ella quiere vivir en Querétaro

Ella se quería ir antes de que Ramses tomara la posición. Ella dijo que se quería ir aproximadamente en febrero aproximadamente ella se iría de la posición.
 June 19 ella renuncio.

¿Cuál fue la conexión, ella estaba feliz o triste de irse, cual fue que crees que fue la conexión?
-	Cuando una persona con potencial 2 renuncia tenemos que investigar.
-	Pienso que como era mejor
-	Gabriela es la que esta en el desarrollo interno.
-	Pienso que en lugar de investiga porque se fue Abigail, mejor cuestionan las personas que intervenimos en ponerle el potencial.
-	Hay algo que sucedió durante la renuncia de ella, cuando ella se fue de incapacidad, tenía que recibir una operación y Ramsés estaba en la posición, se fue un miércoles, la operaron un jueves y tuvimos una situación con el cliente y le empezaron a llamar. No es posible que me marquen y estoy recién operada, tomen este mensaje como mi renuncia.
-	La conversación fue acalorada, una vez que regrese revisamos que paso.
-	Cuando llega no la dejan llegar a su lugar, sino que vaya a RH y le dan su carta de renuncia, ella firma y se va, después Ramsés le comenta a Joaquín que ella hizo una demanda porque tenia que meter comprobante de gastos todo se lo descontaron.
Se fue, metió demanda y empieza a ver esto de manera complicada.
Favoritismo:
¿Consideras que tenias favoritismo sobre Abigail?
No, porque ella no me reportaba directo a mí, le reportaba a PEDRO. Estaban sobre la línea jerárquica de una manera.
Si había interacciones con Abigail, no creo que sea favoritismo porque tenemos datos y análisis donde, al contrario, era quien tenía más carga.

¿Había alguna relación fuera del entorno laboral entre tu y Abigail?
No.

¿Junior García estaba involucrado en esto?
No, el fue un Quality Manager en MTY, después se mueve, el se fue de la compañía 3-4 meses atrás.
 



Hay una renuncia, parte de la responsabilidad de HR es porque la gente se va de la compañía.
Hay un cuestionario
Es parte de la responsabilidad de Gabriela.
Si ella recibe información de que “Joaquín no es justo con los demás siento que tiene una preferencia” esa seria la respuesta de porque Gabriela continuo con otra discusión, pero sería la razón para ir con una entrevista con pedro.
















