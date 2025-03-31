import os
import shutil
import time
import win32com.client
from datetime import datetime

#Direcotorio del exe

dir_script = "Escribir directorio"

# Directorio de origen
dir_origen = ""

# Directorio de destino
dir_destino = ""


# Crea una instancia del programador de tareas
task_scheduler = win32com.client.Dispatch('Schedule.Service')

# Conecta el programador de tareas local
task_scheduler.Connect()

# Crea una nueva tarea
task = task_scheduler.NewTask(0)

# Establece el nombre y la descripción de la tarea
task.RegistrationInfo.Description = "Tarea para mover archivos png"

# Establece el tipo de desencadenador
trigger = task.Triggers.Create(8)  # Desencadenador "Al inicio del equipo"

# Establece la acción a realizar
action = task.Actions.Create(0)  # Acción "Iniciar un programa"
action.Path = dir_script # Ruta al ejecutable de Python
#action.Arguments = __file__  # Argumentos del programa (nombre del archivo actual)

# Guarda la tarea
task_folder = task_scheduler.GetFolder("\\")
task_folder.RegisterTaskDefinition(
    "Mover archivos png",  # Nombre de la tarea
    task,  # Objeto de la tarea
    6,  # Constante para crear la tarea
    "",  # Nombre de usuario para ejecutar la tarea (opcional)
    "",  # Contraseña para ejecutar la tarea (opcional)
    1  # Constante para habilitar la tarea
)



# Función para mover y organizar archivos png por fecha
def mover_archivos_png():
    # Comprueba si hay archivos en el directorio de origen
    archivos = [f for f in os.listdir(dir_origen) if f.endswith('.png')]
    if archivos:
        # Crea una carpeta para el día actual si no existe
        hoy = datetime.today().date()
        carpeta_destino = os.path.join(dir_destino, str(hoy))
        if not os.path.exists(carpeta_destino):
            os.mkdir(carpeta_destino)
        # Mueve los archivos a la carpeta de destino
        for archivo in archivos:
            fecha_mod = os.path.getmtime(os.path.join(dir_origen, archivo))
            fecha_mod = datetime.fromtimestamp(fecha_mod).date()
            carpeta_fecha = os.path.join(dir_destino, str(fecha_mod))
            if not os.path.exists(carpeta_fecha):
                os.mkdir(carpeta_fecha)
            shutil.move(os.path.join(dir_origen, archivo), carpeta_fecha)
        print(f"Se movieron {len(archivos)} archivos.")
    else:
        print("No hay archivos para mover.")
    # Espera 2 segundos antes de volver a comprobar
    time.sleep(2)
    # Vuelve a llamar a la función
    mover_archivos_png()

# Llama a la función para mover y organizar archivos png
mover_archivos_png()
