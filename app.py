from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import pandas as pd
import sqlite3
import os
import sys
import webbrowser

# Importar el motor de cálculo
import importlib.util
spec = importlib.util.spec_from_file_location("motor", os.path.join(os.path.dirname(__file__), "2026_Provision.py"))
motor = importlib.util.module_from_spec(spec)
spec.loader.exec_module(motor)

app = Flask(__name__)
app.secret_key = "provision_engine_2026_secret"

# --- GESTIÓN DE RUTAS PARA PORTABILIDAD ---
if getattr(sys, 'frozen', False):
    # Rutas internas (empaquetadas en el EXE)
    INTERNAL_DIR = sys._MEIPASS
    # Rutas externas (al lado del EXE)
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Definir la base como la carpeta donde está este archivo app.py
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    INTERNAL_DIR = BASE_DIR

app.template_folder = os.path.join(INTERNAL_DIR, 'templates')
os.makedirs(os.path.join(BASE_DIR, 'db'), exist_ok=True)

DB_PATH = os.path.join(BASE_DIR, 'db', 'provision_master.db')
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'inputs')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Información detallada de las bases permanentes
DB_INFO = {
    'maestro_clientes': {'name': 'Maestro de Clientes', 'desc': 'Vincula Ship-to con nombres y zonas.'},
    'tarifario': {'name': 'Tarifario General', 'desc': 'Precios por Carrier, Zona y Unidad.'},
    'rutas': {'name': 'Descripción de Rutas', 'desc': 'Asignación de rutas logísticas por zona.'},
    'tarifas_alcance': {'name': 'Configuración de Alcance', 'desc': 'Tarifas fijas de alcance/retiro.'}
}

def get_db_connection():
    return sqlite3.connect(DB_PATH)

def get_current_status():
    status = {}
    try:
        with get_db_connection() as conn:
            for table in DB_INFO.keys():
                try:
                    count = pd.read_sql(f"SELECT COUNT(*) as c FROM {table}", conn).iloc[0]['c']
                    status[table] = f"Cargado ({count} registros)"
                except:
                    status[table] = "No cargado"
    except:
        status = {k: "Error de conexión" for k in DB_INFO.keys()}
    return status

@app.route('/')
def index():
    status = get_current_status()
    return render_template('Provision.html', status=status, db_info=DB_INFO, results=None)

@app.route('/download/<filename>')
def download_file(filename):
    path = os.path.join(motor.DIR_OUTPUTS, filename)
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    else:
        flash("El archivo solicitado no existe.")
        return redirect(url_for('index'))

@app.route('/process', methods=['POST'])
def process():
    mes = request.form.get('mes')
    
    try:
        success, result = motor.ejecutar_calculo_provision(mes)
        if success:
            # Leer el excel generado para mostrarlo en el HTML
            xl = pd.ExcelFile(result)
            sheets_data = {}
            for sheet_name in xl.sheet_names:
                df = pd.read_excel(result, sheet_name=sheet_name)
                # Convertimos a HTML limitando a 50 filas para que la página sea fluida
                sheets_data[sheet_name] = df.head(50).to_html(
                    classes='table table-striped table-hover table-sm table-bordered',
                    index=False,
                    justify='left'
                )
            
            filename = os.path.basename(result)
            return render_template('Provision.html', 
                                 status=get_current_status(), 
                                 db_info=DB_INFO, 
                                 results=sheets_data, 
                                 download_url=url_for('download_file', filename=filename))
        else:
            flash(f"Error: {result}")
    except Exception as e:
        flash(f"Error en el proceso: {str(e)}")
    
    return redirect(url_for('index'))

if __name__ == '__main__':
    # Abre el navegador automáticamente al iniciar
    webbrowser.open("http://127.0.0.1:5000")
    app.run(debug=True, port=5000)