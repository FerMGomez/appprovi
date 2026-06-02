import os
import shutil

# 1. Nombre de la carpeta principal donde se consolidará todo el proyecto
PROYECTO_RAIZ = 'app_Provision'

# 2. Definir la estructura de subcarpetas interna
ESTRUCTURA = ['inputs', 'master_data', 'db', 'outputs', 'utilidades', 'eliminar', 'templates']

# 3. Mapeo exhaustivo de archivos y dependencias a sus destinos finales
MAPEO = {
    # Scripts Principales y Automatización (Van en la raíz de app_Provision)
    '2026_Provision.py': '.',
    'Ejecutar_Provision.bat': '.',
    'app.py': '.', # app.py ahora es el punto de entrada para el EXE
    'requirements.txt': '.',
    'RESUMEN_PROYECTO.txt': '.',
    'Provision.html': 'templates',

    # Entradas mensuales (Lo que cambias mes a mes)
    'Viajes trafico.xlsx': 'inputs',
    'billing_consolidated.parquet': 'inputs',
    'Tarifario_macro.xlsm': 'inputs',
    
    # Datos maestros (Archivos fijos)
    'ZCUST.xlsx': 'master_data',
    'Descripcion de rutas.xlsx': 'master_data',
    'Tarifas_Alcance.xlsx': 'master_data',
    
    # Herramientas de soporte
    'actualizar_tarifario.py': 'utilidades',
    'crear_maestro_clientes.py': 'utilidades',
    
    # Archivos obsoletos o de prueba
    'pruebas.py': 'eliminar',
    '1_script_ok.py': 'eliminar',
    'pruebas_backup.py': 'eliminar',
    'temp_debug.py': 'eliminar',
    'provision.py': 'eliminar', 
    'Modificaciones.txt': 'eliminar',
    'Consultas.txt': 'eliminar'
}

def organizar():
    print(f"--- Iniciando Consolidación Completa en: {PROYECTO_RAIZ} ---")
    
    # Crear carpeta principal si no existe
    if not os.path.exists(PROYECTO_RAIZ):
        os.makedirs(PROYECTO_RAIZ)
        print(f"Directorio raíz creado: {PROYECTO_RAIZ}")

    # Crear subcarpetas internas dentro de la raíz del proyecto
    for carpeta in ESTRUCTURA:
        path = os.path.join(PROYECTO_RAIZ, carpeta)
        if not os.path.exists(path):
            os.makedirs(path)
            print(f"Subcarpeta creada: {path}")
            
    # Mover archivos a sus posiciones finales dentro de app_Provision
    for archivo, destino in MAPEO.items():
        if os.path.exists(archivo):
            try:
                target_path = os.path.join(PROYECTO_RAIZ, destino, archivo)
                # Evitar mover el archivo sobre sí mismo si ya se está en la carpeta correcta
                if os.path.abspath(archivo) != os.path.abspath(target_path):
                    shutil.move(archivo, target_path)
                    print(f"Consolidado: {archivo} -> {os.path.join(PROYECTO_RAIZ, destino)}/")
            except Exception as e:
                print(f"Error moviendo {archivo}: {e}")
                
    # Intentar mover este propio script de organización al final
    nombre_script = os.path.basename(__file__)
    if os.path.exists(nombre_script):
        try:
            shutil.move(nombre_script, os.path.join(PROYECTO_RAIZ, nombre_script))
            print(f"Script de organización auto-movido a {PROYECTO_RAIZ}/")
        except:
            print(f"\n[AVISO] No se pudo mover '{nombre_script}' automáticamente. Muévelo manualmente.")

    print(f"\n--- ¡Listo! Todo el proyecto ha sido consolidado en '{PROYECTO_RAIZ}' ---")
    print(f"Ahora puedes entrar en la carpeta y usar 'Ejecutar_Provision.bat'.")

if __name__ == "__main__":
    organizar()