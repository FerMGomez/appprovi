import pandas as pd
import numpy as np
import re
import os
import warnings
import importlib.util
import sqlite3
import sys

# Suprimir warnings
warnings.filterwarnings('ignore', category=FutureWarning)

# --- CONFIGURACIÓN GLOBAL ---
DEBUG = False

# --- CONFIGURACIÓN DE RUTAS PORTÁTILES ---
# Detectar si estamos corriendo como ejecutable o como script
if getattr(sys, 'frozen', False):
    # Si es un EXE, BASE_DIR es la carpeta donde está el EXE
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Si es script, BASE_DIR es la carpeta del script
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DIR_INPUTS = os.path.join(BASE_DIR, 'inputs')
DIR_MASTER = os.path.join(BASE_DIR, 'master_data')
DIR_OUTPUTS = os.path.join(BASE_DIR, 'outputs')
DIR_DB = os.path.join(BASE_DIR, 'db')

if not os.path.exists(DIR_DB): os.makedirs(DIR_DB)
if not os.path.exists(DIR_OUTPUTS): os.makedirs(DIR_OUTPUTS)

DB_PATH = os.path.join(DIR_DB, 'provision_master.db')

# Archivos de actualización mensual (en la carpeta inputs)
viajes_file = os.path.join(DIR_INPUTS, 'Viajes trafico.xlsx')
billing_parquet_file = os.path.join(DIR_INPUTS, 'billing_consolidated.parquet')
tarifario_file = os.path.join(DIR_INPUTS, 'Tarifario_macro.xlsm')

# Archivos maestros permanentes
maestro_file = os.path.join(DIR_MASTER, 'ZCUST.xlsx')
rutas_file = os.path.join(DIR_MASTER, 'Descripcion de rutas.xlsx')
alcance_file = os.path.join(DIR_MASTER, 'Tarifas_Alcance.xlsx')

RE_RE_VIAJE = re.compile(r'^\d+_[A-Z]{3}_\d+_[A-Z]$')
RE_UNICO = re.compile(r'^\d+_[A-Z]{3}_\d+$')
RE_ALCANCE_EXPO = re.compile(r'^\d+_EX_\d+$')
RE_AR = re.compile(r'^AR[A-Z0-9]{8}$')

# --- Funciones de Utilidad ---
def normalize_column(df, column):
    if column in df.columns:
        df[column] = df[column].where(df[column].notna(), '').astype(str).str.strip()
    return df

def normalize_columns(df, columns):
    """Aplica la normalización a una lista de columnas en un DataFrame."""
    for col in columns:
        df = normalize_column(df, col)
    return df

def normalizar_transportista(valor):
    """Normaliza nombres de transportistas y unifica aliases frecuentes."""
    base = str(valor).strip().upper()
    if not base or base == 'NAN':
        return ''

    base = ' '.join(base.split())
    aliases = {
        'LOGISCHER NEA S.A.': 'LOGISCHER NEA SA',
        'LOS AMIGOS': 'TTES. LOS AMIGOS S.A.',
        'LOS AMIGOS S.A.': 'TTES. LOS AMIGOS S.A.',
        'TTES LOS AMIGOS S.A.': 'TTES. LOS AMIGOS S.A.',
        'TTES. LOS AMIGOS SA': 'TTES. LOS AMIGOS S.A.',
        'TRANSPORTES LOS AMIGOS S.A.': 'TTES. LOS AMIGOS S.A.',
        'GIAMPAOLETTI': 'GIAMPAOLETTI BUOSI S.A.',
        'LOGICARGO': 'LOGICARGO SRL',
        'GIAMPAOLETTI BUOSI SA': 'GIAMPAOLETTI BUOSI S.A.',
        'DORINKA SRL': 'DORINKA S.R.L.',
        'DORINKA': 'DORINKA S.R.L.',
        'DIA ARGENTINA SA': 'DIA ARGENTINA S.A.',
        'DIA ARGENTINA': 'DIA ARGENTINA S.A.',
        'LOGICARGO SRL': 'LOGICARGO SRL',
    }
    return aliases.get(base, base)


def coalesce_columns(df, target, candidates):
    """Consolida múltiples variantes de una columna en una sola columna canónica."""
    disponibles = [col for col in candidates if col in df.columns]
    if not disponibles:
        return df

    if target not in df.columns:
        df[target] = df[disponibles[0]]
        disponibles = disponibles[1:]

    target_vacio = df[target].isna() | df[target].astype(str).str.strip().eq('')
    for col in disponibles:
        df.loc[target_vacio, target] = df.loc[target_vacio, col]
        target_vacio = df[target].isna() | df[target].astype(str).str.strip().eq('')

    return df


def preparar_columnas_transporte_viajes(df):
    """Normaliza las columnas de transporte del archivo de viajes.

    Soporta el formato nuevo con TRANSPORTE_ALCANCE y TRANSPORTE_XD,
    manteniendo compatibilidad con archivos viejos que solo traen TRANSPORTE.
    """
    df = df.copy()

    # 1. Separar transportistas por barra / (vuelto a la lógica original estable)
    if 'TRANSPORTE' in df.columns:
        split_trans = df['TRANSPORTE'].fillna('').astype(str).str.split('/')
        df['TRANSPORTE_ALCANCE'] = split_trans.str[0].str.strip().apply(normalizar_transportista)

        # El segundo transportista es el de distribución (XD)
        parsed_xd = split_trans.apply(lambda x: x[1].strip() if len(x) > 1 else '')
        override_xd = df['TRANSPORTE_XD'].fillna('').astype(str).str.strip() if 'TRANSPORTE_XD' in df.columns else pd.Series('', index=df.index)
        
        df['TRANSPORTE_XD'] = np.where(override_xd != '', override_xd, parsed_xd)
        df['TRANSPORTE_XD'] = df['TRANSPORTE_XD'].apply(normalizar_transportista)
    
    return df


def clasificar_t_viaje_vectorizado(df):
    """Clasifica T_VIAJE de forma vectorizada para evitar apply por fila."""
    id_viaje = df['ID_VIAJES'].where(df['ID_VIAJES'].notna(), '').astype(str).str.strip()
    tipo_viaje = df['TIPO DE VIAJE'].where(df['TIPO DE VIAJE'].notna(), '').astype(str).str.strip()

    resultado = pd.Series('Otro', index=df.index, dtype='object')

    mask_alcance_d = tipo_viaje.eq('Alcance - Distribución')
    mask_re_viaje = id_viaje.str.match(RE_RE_VIAJE) & tipo_viaje.eq('Simple')
    mask_unico = id_viaje.str.match(RE_UNICO) & tipo_viaje.eq('Simple')
    mask_alcance_expo = id_viaje.str.match(RE_ALCANCE_EXPO) & tipo_viaje.eq('Exportación')
    mask_alcance = id_viaje.str.match(RE_AR) & tipo_viaje.isin(['Alcance', 'Exportación'])
    mask_alcance_simple = id_viaje.str.match(RE_AR) & tipo_viaje.eq('Simple')
    mask_retiro = id_viaje.str.match(RE_AR) & tipo_viaje.eq('Retiro')
    mask_distribucion = id_viaje.str.match(RE_UNICO) & tipo_viaje.eq('Distribución - Troncal')

    resultado.loc[mask_alcance_d] = 'Alcance_D'
    resultado.loc[mask_re_viaje] = 'Re_Viaje'
    resultado.loc[mask_unico] = 'Unico'
    resultado.loc[mask_alcance_expo] = 'Alcance_Expo'
    resultado.loc[mask_alcance] = 'Alcance'
    resultado.loc[mask_alcance_simple] = 'Alcance'
    resultado.loc[mask_retiro] = 'Retiro'
    resultado.loc[mask_distribucion] = 'Distribución'

    return resultado


def preparar_tarifario_auxiliares(tarifario):
    """Precalcula estructuras auxiliares del tarifario para reutilizarlas."""
    id_vars = ['CARRIER', 'TRANSPORT ZONE']
    value_vars = [col for col in tarifario.columns if col not in id_vars]

    tarifario_long = tarifario.melt(
        id_vars=id_vars,
        value_vars=value_vars,
        var_name='UNIDAD_NORM',
        value_name='Tarifa_Base'
    )
    tarifario_long['TRANSPORTE_NORM'] = tarifario_long['CARRIER'].apply(normalizar_transportista)
    tarifario_long['ZONE_KEY'] = tarifario_long['TRANSPORT ZONE'].astype(str).str.strip().str.upper()
    tarifario_long['UNIDAD_NORM'] = tarifario_long['UNIDAD_NORM'].astype(str).str.strip().str.upper()

    tarifario_dist = tarifario.rename(
        columns={'CARRIER': 'TRANSPORTE_NORM', 'TRANSPORT ZONE': 'ZONE_KEY'}
    ).copy()
    tarifario_dist['TRANSPORTE_NORM'] = tarifario_dist['TRANSPORTE_NORM'].apply(normalizar_transportista)
    tarifario_dist['ZONE_KEY'] = tarifario_dist['ZONE_KEY'].astype(str).str.strip().str.upper()

    zona_interior_ref = 'AR00BA1001'
    tarifario_ref = tarifario_long[tarifario_long['ZONE_KEY'] == zona_interior_ref][
        ['TRANSPORTE_NORM', 'UNIDAD_NORM', 'Tarifa_Base']
    ].copy()
    tarifario_ref.rename(columns={'Tarifa_Base': 'Tarifa_Ref'}, inplace=True)
    tarifario_ref.drop_duplicates(subset=['TRANSPORTE_NORM', 'UNIDAD_NORM'], inplace=True)

    return {
        'long': tarifario_long,
        'dist': tarifario_dist,
        'ref': tarifario_ref,
    }


def asignar_tarifa_vectorizado(df_viajes, tarifario_aux):
    """Asigna tarifas de forma vectorizada usando auxiliares precalculados."""
    df = df_viajes.copy()

    # Claves de cruce
    df['TRANSPORTE_NORM'] = df['TRANSPORTE'].apply(normalizar_transportista)
    df['UNIDAD_NORM'] = df['UNIDAD'].astype(str).str.strip().str.upper()

    df['ZONE_KEY'] = np.where(
        df['T_VIAJE'].isin(['Alcance', 'Retiro']),
        df['ID_VIAJES'].astype(str).str.strip().str.upper(),
        df['TRANSPORT ZONE'].astype(str).str.strip().str.upper()
    )

    # 1. Tarifa Base (viajes directos)
    df = pd.merge(
        df, 
        tarifario_aux['long'], 
        on=['TRANSPORTE_NORM', 'ZONE_KEY', 'UNIDAD_NORM'], 
        how='left'
    )

    # 2. Lógica Distribución (si aplica)
    # El cálculo por peso se activa por clasificación (Distribución/Alcance_D) o por transportista (Andreani/Logicargo)
    carriers_aforo = ['ANDREANI LOGISTICA S.A.', 'LOGICARGO SRL']
    mask_dist = df['T_VIAJE'].isin(['Distribución', 'Alcance_D']) | \
                df['TRANSPORTE_NORM'].isin(carriers_aforo)

    if mask_dist.any() and 'Gross weight' in df.columns and not df[mask_dist]['Gross weight'].isna().all():
        # Procesamos solo las filas de distribución
        df_dist_rows = df[mask_dist].drop(columns=['Tarifa_Base'], errors='ignore')
        df_dist_rows = pd.merge(df_dist_rows, tarifario_aux['dist'], on=['TRANSPORTE_NORM', 'ZONE_KEY'], how='left')
        
        peso = df_dist_rows['Gross weight'].fillna(0)
        conds = [
            (df_dist_rows['TRANSPORTE_NORM'].isin(['ANDREANI LOGISTICA S.A.', 'LOGICARGO SRL'])) & (peso <= 900),
            (df_dist_rows['TRANSPORTE_NORM'].isin(['ANDREANI LOGISTICA S.A.', 'LOGICARGO SRL'])) & (peso > 900),
            (df_dist_rows['TRANSPORTE_NORM'].isin(['GIAMPAOLETTI BUOSI S.A.', 'LOGISCHER NEA SA']))
        ]
        choices = [
            df_dist_rows['Aforo x 900KG'],
            df_dist_rows['Aforo x 900KG'] + (peso - 900) * df_dist_rows['X KG'],
            df_dist_rows['Aforo x 900KG'] + (peso * df_dist_rows['X KG'])
        ]
        df.loc[mask_dist, 'Tarifa_Base'] = np.select(conds, choices, default=np.nan)

    # 3. Penalidades
    ZONAS_AMBA = ['AR00BA1001', 'AR00BA1002', 'AR00BA1003', 'AR00BA1004']
    presentismo = df['PRESENTISMO'].astype(str).str.strip().str.upper()
    es_ff = presentismo == 'FALSO FLETE'
    es_ns = presentismo == 'NO SHOW'
    es_penalidad = es_ff | es_ns

    df['Tarifa'] = df['Tarifa_Base'].fillna(0)

    if es_penalidad.any():
        df = pd.merge(df, tarifario_aux['ref'], on=['TRANSPORTE_NORM', 'UNIDAD_NORM'], how='left')
        es_interior_pen = es_penalidad & ~df['ZONE_KEY'].isin(ZONAS_AMBA)
        df.loc[es_interior_pen, 'Tarifa'] = df['Tarifa_Ref']
        df.drop(columns=['Tarifa_Ref'], inplace=True, errors='ignore')

        df.loc[es_ff, 'Tarifa'] *= 0.5
        df.loc[es_ns, 'Tarifa'] *= -0.5

    return df


# --- Carga de Datos (Modificada) ---
def load_billing(files, columns, month_to_filter):
    """Carga y pre-procesa los archivos de facturación (billing).

    Concatena múltiples archivos Excel, filtra por mes si se especifica,
    excluye clientes específicos y normaliza columnas clave.
    """
    df_list = []
    for f in files:
        df = pd.read_excel(f, engine='openpyxl')
        df['source_file'] = os.path.basename(f)
        print(f"  Archivo {os.path.basename(f)}: {len(df)} filas cargadas")
        # Mostrar rango de fechas en este archivo
        if 'Billing date' in df.columns:
            df['Billing date'] = pd.to_datetime(df['Billing date'])
            print(f"    Fechas: {df['Billing date'].min()} a {df['Billing date'].max()}")
        df_list.append(df)
    
    billing = pd.concat(df_list, ignore_index=True)
    print(f"  Total después de concatenar: {len(billing)} filas")
    cols_to_keep = list(set(columns + ['source_file']))
    billing = billing[[col for col in cols_to_keep if col in billing.columns]]

    # Filtrar filas donde 'Accounting document number' está vacío
    if 'Accounting document number' in billing.columns:
        billing.dropna(subset=['Accounting document number'], inplace=True)
        billing = billing[billing['Accounting document number'].astype(str).str.strip() != ''].copy()

    billing['Billing date'] = pd.to_datetime(billing['Billing date'])
    if month_to_filter is not None:
        print(f"Filtrando facturación para el mes: {month_to_filter}")
        billing = billing[billing['Billing date'].dt.month == month_to_filter].copy()

    ship_to_exclude = ["3000021508", "1000078395", "1000078394", "1000078144"]
    billing['Ship to party'] = billing['Ship to party'].astype(str)
    billing = billing[~billing['Ship to party'].isin(ship_to_exclude)]

    billing = normalize_columns(billing, ['Reference Document number', 'Ship to party'])
    return billing

def load_customer_master(path):
    """Carga y pre-procesa el archivo maestro de clientes (ZCUST).

    Lee el archivo, selecciona y renombra columnas relevantes, filtra clientes,
    elimina duplicados y crea un campo combinado de nombre y ciudad.
    """
    df = pd.read_excel(path, engine='openpyxl')
    
    required_columns = ['Sold To Num', 'Ship To', 'Customer Number', 'City', 'TranspZone']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Faltan las siguientes columnas en el archivo '{path}': {', '.join(missing_columns)}")

    df = df[required_columns].copy()
    df['Sold To Num'] = df['Sold To Num'].astype(str)
    df = df[df['Sold To Num'].str.startswith('1')].copy()
    df.drop_duplicates(subset=['Ship To'], inplace=True)
    df['Customer_Name_City'] = df['Customer Number'] + ' (' + df['City'] + ')'
    
    # Normalizar columnas clave a string antes de renombrar
    df['Ship To'] = df['Ship To'].astype(str).str.strip()
    
    df.rename(columns={
        'Sold To Num': 'CLIENTE_MC_NUM', 
        'Customer Number': 'CLIENTE_MC_NAME',
        'TranspZone': 'TRANSPORT ZONE'
    }, inplace=True)
    return df

def load_rutas(path):
    """Carga y pre-procesa el archivo de descripción de rutas.

    Lee la tabla desde B2:U206, selecciona columnas R (Transport Zone) y U (Ruta),
    normaliza y elimina duplicados.
    """
    # Columnas: A=0, B=1, ..., R=17, U=20
    df = pd.read_excel(path, sheet_name='OBD', usecols=[17, 20], skiprows=1, engine='openpyxl', header=None)
    df.columns = ['TRANSPORT ZONE', 'Ruta']
    df = normalize_columns(df, ['TRANSPORT ZONE', 'Ruta'])
    df.drop_duplicates(subset=['TRANSPORT ZONE'], inplace=True)
    return df

def load_tarifario(path):
    """Carga y pre-procesa el archivo de tarifario.

    Lee la hoja "Tarifario" de un archivo Excel, omite las primeras filas,
    elimina duplicados y normaliza las columnas de transportista y zona de transporte.
    """
    df = pd.read_excel(path, sheet_name="Tarifario", skiprows=11, engine='openpyxl')
    df.drop_duplicates(subset=['CARRIER', 'TRANSPORT ZONE'], inplace=True)
    df = normalize_columns(df, ['CARRIER', 'TRANSPORT ZONE'])
    return df

# --- GESTIÓN DE BASE DE DATOS ROBUSTA ---
def actualizar_base_de_datos(db_path, maestro_file, tarifario_file, rutas_file, alcance_file, dir_master):
    """Carga los archivos maestros a SQLite si hay cambios o es la primera vez."""
    conn = sqlite3.connect(db_path)
    print("Sincronizando Base de Datos de Maestros...")
    
    if os.path.exists(maestro_file):
        load_customer_master(maestro_file).to_sql('maestro_clientes', conn, if_exists='replace', index=False)
        print("  - Maestro de clientes [OK]")
    
    # Prioridad: inputs (aumentos) > master_data (base)
    path_tarifario = tarifario_file
    if not os.path.exists(path_tarifario):
        path_tarifario = os.path.join(dir_master, 'Tarifario_macro.xlsm')

    if os.path.exists(path_tarifario):
        load_tarifario(path_tarifario).to_sql('tarifario', conn, if_exists='replace', index=False)
        print(f"  - Tarifario cargado desde {os.path.basename(os.path.dirname(path_tarifario))} [OK]")
    else:
        # Si no hay archivo físico, confiamos en que la tabla ya exista en la DB
        pass
        
    if os.path.exists(rutas_file):
        load_rutas(rutas_file).to_sql('rutas', conn, if_exists='replace', index=False)
        print("  - Descripcion de rutas [OK]")

    if os.path.exists(alcance_file):
        pd.read_excel(alcance_file).to_sql('tarifas_alcance', conn, if_exists='replace', index=False)
        print("  - Tarifas de alcance [OK]")
        
    conn.close()

def preparar_billing_consolidado(billing_df, customer_master_df, rutas_df):
    """Normaliza y enriquece el billing consolidado.

    Si el parquet ya trae columnas enriquecidas, las reutiliza. Si no,
    completa los datos desde maestro de clientes y descripción de rutas.
    """
    billing = billing_df.copy()

    columns_to_normalize = [
        'Ship to party',
        'Reference Document number',
        'Billing document',
        'Accounting document number',
        'LR number',
        'Reference',
        'TRANSPORT ZONE',
        'Ruta',
        'Ruta_SAP',
        'RUTA SAP',
    ]
    billing = normalize_columns(billing, [col for col in columns_to_normalize if col in billing.columns])

    if 'Billing date' in billing.columns:
        billing['Billing date'] = pd.to_datetime(billing['Billing date'], errors='coerce')

    if 'source_file' not in billing.columns:
        billing['source_file'] = 'billing_consolidated.parquet'

    required_customer_cols = ['CLIENTE_MC_NUM', 'CLIENTE_MC_NAME', 'Customer_Name_City', 'TRANSPORT ZONE']
    missing_customer_cols = [col for col in required_customer_cols if col not in billing.columns]
    if missing_customer_cols:
        rows_before_merge = len(billing)
        billing = pd.merge(
            billing,
            customer_master_df,
            left_on='Ship to party',
            right_on='Ship To',
            how='left'
        )
        rows_after_merge = len(billing)
        if rows_after_merge != rows_before_merge:
            print(
                f"Advertencia: el merge de facturación con maestro de clientes cambió "
                f"el número de filas de {rows_before_merge} a {rows_after_merge}"
            )
        billing.drop(columns=['Ship To'], inplace=True, errors='ignore')

    billing = coalesce_columns(billing, 'TRANSPORT ZONE', ['TRANSPORT ZONE', 'TRANSPORT ZONE_x', 'TRANSPORT ZONE_y'])
    billing = coalesce_columns(billing, 'Ruta', ['Ruta', 'Ruta_x', 'Ruta_y', 'Ruta_SAP', 'RUTA SAP'])
    billing = coalesce_columns(billing, 'Ruta_SAP', ['Ruta_SAP', 'Ruta_SAP_x', 'Ruta_SAP_y', 'RUTA SAP', 'Ruta'])

    if 'Ruta' not in billing.columns:
        if 'Ruta_SAP' in billing.columns:
            billing['Ruta'] = billing['Ruta_SAP']
        elif 'RUTA SAP' in billing.columns:
            billing['Ruta'] = billing['RUTA SAP']
        else:
            billing = pd.merge(billing, rutas_df, on='TRANSPORT ZONE', how='left')

    if 'Ruta_SAP' not in billing.columns:
        if 'RUTA SAP' in billing.columns:
            billing['Ruta_SAP'] = billing['RUTA SAP']
        elif 'Ruta' in billing.columns:
            billing['Ruta_SAP'] = billing['Ruta']

    if 'Ruta' in billing.columns:
        billing['Ruta'] = billing['Ruta'].fillna('')
    if 'Ruta_SAP' in billing.columns:
        billing['Ruta_SAP'] = billing['Ruta_SAP'].fillna('')

    return billing


def filtrar_billing_operativo(billing_df):
    """Aplica filtros de negocio sobre billing consolidado."""
    billing = billing_df.copy()

    if 'Accounting document number' in billing.columns:
        accounting_doc = billing['Accounting document number'].where(
            billing['Accounting document number'].notna(), ''
        ).astype(str).str.strip()
        billing = billing[accounting_doc != ''].copy()

    # Excluir solo clientes no operativos requeridos por negocio para resumenes:
    # - Rofina (en algunos archivos aparece como Reckitt Gran Barata)
    # - Reckitt Samples
    if 'Ship to party name' in billing.columns:
        ship_to_name = billing['Ship to party name'].where(
            billing['Ship to party name'].notna(), ''
        ).astype(str).str.strip().str.upper()
        mask_excluir = ship_to_name.str.contains(
            r'ROFINA|GRAN BARATA|SAMPLES?',
            regex=True,
            na=False
        )
        billing = billing[~mask_excluir].copy()

    return billing


def agregar_source_suffix(df, source_map):
    billing = df.copy()
    if 'source_file' not in billing.columns:
        billing['source_file'] = 'billing_consolidated.parquet'
    billing['source_suffix'] = billing['source_file'].map(source_map).fillna('Otro')
    return billing

def reset_index_if_needed(df):
    if isinstance(df.index, pd.RangeIndex):
        return df
    return df.reset_index()

def debug_print(*args, **kwargs):
    if DEBUG:
        print(*args, **kwargs)

def ejecutar_calculo_provision(mes_analisis_input):
    """Función principal para ser llamada desde la App Web o Script."""
    MES_DE_ANALISIS = int(mes_analisis_input)
    
    print(f"Iniciando procesamiento Mes {MES_DE_ANALISIS}...")
    try:
        # 0. Verificación preventiva de archivos de entrada (Inputs)
        missing_inputs = []
        if not os.path.exists(viajes_file): missing_inputs.append(f"  - {viajes_file}")
        if not os.path.exists(billing_parquet_file): missing_inputs.append(f"  - {billing_parquet_file}")
        
        if missing_inputs:
            print("\n" + "="*70)
            print("ERROR: FALTAN ARCHIVOS MENSUALES EN LA CARPETA 'inputs'")
            print("="*70)
            for m in missing_inputs:
                print(m)
            print("\nPor favor, coloca los archivos del mes y vuelve a ejecutar.")
            print("="*70 + "\n")
            return False, "Faltan archivos de entrada en /inputs"

        # 1. Sincronizar Base de Datos pasando las rutas
        actualizar_base_de_datos(DB_PATH, maestro_file, tarifario_file, rutas_file, alcance_file, DIR_MASTER)
        conn = sqlite3.connect(DB_PATH)

        # Verificación de tablas críticas
        tablas_requeridas = ['tarifario', 'maestro_clientes', 'rutas', 'tarifas_alcance']
        for t in tablas_requeridas:
            res = conn.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{t}'").fetchone()
            if not res:
                print(f"\n[ERROR] La tabla '{t}' no existe. Asegúrate de que los archivos en 'master_data' existan.")
                conn.close()
                return False, f"La tabla {t} no pudo cargarse."

        print(f"Cargando viajes: {viajes_file}")
        df_viajes_trafico = pd.read_excel(viajes_file, sheet_name='Hoja1', engine='openpyxl')
        df_viajes_trafico['ID_FILA'] = range(len(df_viajes_trafico))

        # Cargar maestros desde la DB (mucho más rápido)
        df_alcance_config = pd.read_sql("SELECT * FROM tarifas_alcance", conn)
        tarifario = pd.read_sql("SELECT * FROM tarifario", conn)
        tarifario_aux = preparar_tarifario_auxiliares(tarifario)
        customer_master_df = pd.read_sql("SELECT * FROM maestro_clientes", conn)
        df_rutas = pd.read_sql("SELECT * FROM rutas", conn)

        # cargar archivo parquet consolidado de facturación
        print(f"  Cargando archivo parquet: {billing_parquet_file}")
        try:
            billing_b_full = pd.read_parquet(billing_parquet_file)
        except Exception as e:
            print(f"Error al leer archivo parquet: {e}")
            return False, str(e)

        # Usar SOLO datos de 2026 o posteriores
        billing_b_full = billing_b_full[billing_b_full['Billing date'].dt.year >= 2026].copy()
        billing_b_full = preparar_billing_consolidado(billing_b_full, customer_master_df, df_rutas)
        billing_b_full = filtrar_billing_operativo(billing_b_full)

        # Crear la version filtrada DESPUÉS del merge
        billing_b_filtered = billing_b_full[billing_b_full['Billing date'].dt.month == MES_DE_ANALISIS].copy()
        if billing_b_filtered.empty:
            fecha_max = billing_b_full['Billing date'].max()
            print(
                f"Advertencia: no hay facturación para el mes {MES_DE_ANALISIS} en el parquet. "
                f"Última fecha disponible: {fecha_max}"
            )

        billing_match_source = billing_b_full.copy()
        billing_report_source = billing_b_filtered.copy() if not billing_b_filtered.empty else billing_b_full.copy()

        # --- Lógica de procesamiento de viajes (V_BASE) ---
        transportes_correctos = [
            'ANDREANI LOGISTICA S.A.', 'LOGICARGO SRL', 'CELSUR LOGISTICA S.A.', 
            'DISTRI 10 S.R.L.', 'GIAMPAOLETTI BUOSI S.A.', 'I-FLOW S.A.', 
            'LOGISCHER NEA SA', 'TTES. LOS AMIGOS S.A.', 'WAL-MART ARGENTINA SRL',
            'DORINKA S.R.L.', 'DIA ARGENTINA S.A.'
        ]
        tipos_simples = ['Simple', 'Dos Puntos', 'Exportación', 'Alcance', 'Retiro']
        tipos_distribucion = ['Distribución - Troncal', 'Alcance - Distribución']
        # Normalizar columnas de transporte del archivo de viajes.
        df_viajes_trafico = preparar_columnas_transporte_viajes(df_viajes_trafico)
        transporte_base = df_viajes_trafico['TRANSPORTE_ALCANCE'].apply(normalizar_transportista)
        df = df_viajes_trafico[transporte_base.isin(transportes_correctos)].copy()
        df['ID_VIAJES'] = df['N DE VIAJE'].astype(str).str.split(',').str[0].str.split('&').str[0].str.strip()
        df['T_VIAJE'] = clasificar_t_viaje_vectorizado(df)

        columnas_finales = ['ID_FILA', 'FECHA CTA', 'N DE VIAJE', 'ID_VIAJES', 'T_VIAJE', 'UNIDAD', 'CLIENTE', 'OBD', 'LOCALIDAD', 'POS','TIPO DE VIAJE', 'PRESENTISMO']
        df_final = df[[c for c in columnas_finales if c in df.columns]].copy()
        # agregar las columnas de transportistas desde el dataframe original
        df_final['TRANSPORTE_ALCANCE'] = df['TRANSPORTE_ALCANCE']
        df_final['TRANSPORTE_XD'] = df['TRANSPORTE_XD']

        # Normalizar OBD para merges
        df_final['OBD'] = df_final['OBD'].astype(str).str.split(',').str[0].str.strip()

        # comprobación inicial de duplicados en df_final
        dup_inicio = df_final['ID_FILA'].duplicated().sum()
        if dup_inicio:
            print(f"Advertencia: df_final contiene {dup_inicio} filas duplicadas en ID_FILA antes de empezar la lógica unificada")

        # 1. Validar los transportistas para cada función (Alcance y Distribución)
        print("Determinando transportistas de Alcance y Distribución...")
        mask_alc_dist = df_final['TIPO DE VIAJE'] == 'Alcance - Distribución'
        if mask_alc_dist.any():
            print(f"Detectados {mask_alc_dist.sum()} viajes de tipo 'Alcance - Distribución'.")

        mask_dist = df_final['TIPO DE VIAJE'].isin(['Distribución - Troncal', 'Alcance - Distribución'])
        mask_xd_vacio = mask_dist & df_final['TRANSPORTE_XD'].eq('')
        if mask_xd_vacio.any():
            print(f"Advertencia: {mask_xd_vacio.sum()} viajes de distribución no tienen TRANSPORTE_XD informado.")

        # 2. Calcular Tarifa de Alcance
        print("Calculando tarifas de Alcance...")
        # ajustamos nombres para coincidir con nuestras columnas nuevas
        df_alcance_config.rename(columns={'TRANSPORTISTA': 'TRANSPORTE_ALCANCE', 'ALCANCE': 'Tarifa_Alcance'}, inplace=True, errors='ignore')

        # Para buscar la tarifa, los viajes 'Alcance - Distribución' deben mapearse a 'Alcance', 
        # que es como probablemente figuran en el archivo Tarifas_Alcance.xlsx.
        # Creamos una clave temporal para el merge.
        df_final['TIPO_VIAJE_ALCANCE_KEY'] = df_final['TIPO DE VIAJE'].replace({'Alcance - Distribución': 'Alcance'})

        # Hacemos el merge usando la clave temporal
        df_final = pd.merge(
            df_final, 
            df_alcance_config[['TRANSPORTE_ALCANCE', 'TIPO DE VIAJE', 'Tarifa_Alcance']],
            left_on=['TRANSPORTE_ALCANCE', 'TIPO_VIAJE_ALCANCE_KEY'], 
            right_on=['TRANSPORTE_ALCANCE', 'TIPO DE VIAJE'],
            how='left',
            suffixes=('', '_y') # Sufijo para la columna 'TIPO DE VIAJE' del df_alcance
        )

        # Limpiamos columnas duplicadas o innecesarias del merge
        df_final.drop(columns=[col for col in df_final.columns if col.endswith('_y')], inplace=True)
        df_final.drop(columns=['TIPO_VIAJE_ALCANCE_KEY'], inplace=True)

        # Llenamos con 0 los viajes que no encontraron tarifa de alcance y nos aseguramos de que la columna exista.
        if 'Tarifa_Alcance' not in df_final.columns:
            df_final['Tarifa_Alcance'] = 0.0
        df_final['Tarifa_Alcance'] = df_final['Tarifa_Alcance'].fillna(0)

        # 3. Calcular Tarifa de Distribución
        print("Calculando tarifas de Distribución...")
        # Asegurarnos de que no haya filas duplicadas antes del merge de master
        rows_before = len(billing_b_full)
        # (la cadena de merge con master ya se hizo arriba, esta es solo una comprobación)
        rows_after = len(billing_b_full)
        if rows_after != rows_before:
            print(f"Advertencia: el dataframe de facturación cambió de {rows_before} a {rows_after} filas después del merge de clientes")

        debug_print(f"DEBUG - billing_b_full antes de agg: {len(billing_b_full)} filas")
        debug_print(f"DEBUG - 'Reference Document number' en billing_b_full: {billing_b_full['Reference Document number'].nunique()} únicos")
        debug_print(f"DEBUG - Ejemplos Reference Document number: {billing_b_full['Reference Document number'].unique()[:5]}")
        transport_zone_col = next((col for col in ['TRANSPORT ZONE', 'TRANSPORT ZONE_x', 'TRANSPORT ZONE_y'] if col in billing_match_source.columns), None)
        ruta_col = next((col for col in ['Ruta', 'Ruta_x', 'Ruta_y', 'Ruta_SAP', 'RUTA SAP'] if col in billing_match_source.columns), None)

        if transport_zone_col is None:
            raise KeyError('No se encontró ninguna columna de TRANSPORT ZONE en billing_match_source')

        agg_dict_direct = {
            transport_zone_col: 'first',
            'Gross weight': 'sum',
            'Billed quantity (Base UoM)': 'sum'
        }
        if ruta_col is not None:
            agg_dict_direct[ruta_col] = 'first'

        billing_direct_agg = billing_match_source.groupby('Reference Document number', as_index=False).agg(agg_dict_direct)
        billing_direct_agg.rename(
            columns={
                transport_zone_col: 'TRANSPORT ZONE',
                'Billed quantity (Base UoM)': 'Cajas',
                'Reference Document number': 'OBD',
                **({ruta_col: 'Ruta'} if ruta_col is not None else {}),
            },
            inplace=True,
        )

        if 'Ruta' not in billing_direct_agg.columns:
            billing_direct_agg['Ruta'] = ''

    # Lookup para viajes simples usando OBD
        billing_direct_agg['OBD'] = billing_direct_agg['OBD'].astype(str).str.strip()
        billing_direct_lookup = billing_direct_agg.set_index('OBD')

        # Agregado para distribución: una fila por viaje y entrega
        agg_dict_dist = {
            transport_zone_col: 'first',
            'Gross weight': 'sum',
            'Billed quantity (Base UoM)': 'sum'
        }
        if ruta_col is not None:
            agg_dict_dist[ruta_col] = 'first'
        if 'Customer_Name_City' in billing_match_source.columns:
            agg_dict_dist['Customer_Name_City'] = 'first'

        billing_dist_agg = billing_match_source.groupby(['LR number', 'Ship to party'], as_index=False).agg(agg_dict_dist)
        billing_dist_agg.rename(
            columns={
                'LR number': 'ID_VIAJES',
                transport_zone_col: 'TRANSPORT ZONE',
                'Billed quantity (Base UoM)': 'Cajas',
                **({ruta_col: 'Ruta'} if ruta_col is not None else {}),
            },
            inplace=True,
        )

        if 'Ruta' not in billing_dist_agg.columns:
            billing_dist_agg['Ruta'] = ''

        billing_dist_agg['ID_VIAJES'] = billing_dist_agg['ID_VIAJES'].astype(str).str.strip()
        billing_dist_agg['Ship to party'] = billing_dist_agg['Ship to party'].astype(str).str.strip()

        debug_print(f"DEBUG - billing_direct_agg: {len(billing_direct_agg)} filas agregadas")
        debug_print(f"DEBUG - billing_dist_agg: {len(billing_dist_agg)} filas agregadas")
        debug_print(f"DEBUG - Gross weight notna en agg dist: {billing_dist_agg['Gross weight'].notna().sum()}")
        df_dist_detalle = df_final[df_final['TIPO DE VIAJE'].isin(['Distribución - Troncal', 'Alcance - Distribución'])].copy()

        if not df_dist_detalle.empty:
            debug_print(f"DEPURACIÓN - ID_VIAJES en df_dist_detalle: {sorted(df_dist_detalle['ID_VIAJES'].astype(str).unique())[:10]}")
            debug_print(f"DEPURACIÓN - ID_VIAJES en billing_dist_agg: {sorted(billing_dist_agg['ID_VIAJES'].astype(str).unique())[:10]}")
            debug_print("Buscando coincidencias entre ID_VIAJES y LR number...")
            coincidencias = df_dist_detalle['ID_VIAJES'].astype(str).isin(billing_dist_agg['ID_VIAJES']).sum()
            debug_print(f"Coincidencias encontradas: {coincidencias} / {len(df_dist_detalle)}")

            df_dist_detalle['ID_VIAJES'] = df_dist_detalle['ID_VIAJES'].astype(str).str.strip()
            df_dist_detalle = pd.merge(
                df_dist_detalle,
                billing_dist_agg,
                on='ID_VIAJES',
                how='left',
                suffixes=('', '_billing')
            )
            debug_print(f"DEPURACIÓN - df_dist_detalle después del cruce con billing: {df_dist_detalle['TRANSPORT ZONE'].notna().sum()} / {len(df_dist_detalle)} con TRANSPORT ZONE notna.")
            debug_print(f"DEPURACIÓN - Gross weight notna después cruce: {df_dist_detalle['Gross weight'].notna().sum()} / {len(df_dist_detalle)}")
            debug_print("Ejemplos después cruce:")
            debug_print(df_dist_detalle[['ID_VIAJES', 'Ship to party', 'TRANSPORT ZONE', 'Gross weight']].head(3))
            # Si no hay zona, asignar una por defecto para que la tarifación funcione
            df_dist_detalle['TRANSPORT ZONE'] = df_dist_detalle['TRANSPORT ZONE'].fillna('AR00BA1001')
            debug_print(f"Después de fillna zona: {df_dist_detalle['TRANSPORT ZONE'].notna().sum()} / {len(df_dist_detalle)}")
            # Para distribución el cálculo siempre usa el transportista informado en TRANSPORTE_XD.
            df_dist_detalle['TRANSPORTE'] = df_dist_detalle['TRANSPORTE_XD']
            df_dist_detalle_tarifado = asignar_tarifa_vectorizado(df_dist_detalle, tarifario_aux)
            tarifas_dist_sumadas = df_dist_detalle_tarifado.groupby('ID_FILA')['Tarifa'].sum().reset_index().rename(columns={'Tarifa': 'Tarifa_Distribucion'})
            before_merge = len(df_final)
            df_final = pd.merge(df_final, tarifas_dist_sumadas, on='ID_FILA', how='left')
            after_merge = len(df_final)
            if after_merge > before_merge:
                dup = df_final['ID_FILA'].duplicated().sum()
                print(f"Advertencia: la fusión de tarifas de distribución aumentó filas de {before_merge} a {after_merge} ({dup} duplicadas)")
        else:
            df_final['Tarifa_Distribucion'] = 0.0

        df_final['Tarifa_Distribucion'] = df_final['Tarifa_Distribucion'].fillna(0)

        # 4. Calcular Tarifa para viajes Directos/Simples
        print("Calculando tarifas para viajes Directos...")
        df_directos = df_final[df_final['TIPO DE VIAJE'].isin(tipos_simples + ['Alcance - Distribución'])].copy()
        if not df_directos.empty:
            # Regla de negocio: para viajes simples se usa solo una OBD por viaje.
            # La columna OBD ya fue normalizada arriba tomando el primer valor informado.
            df_directos['TRANSPORT ZONE'] = df_directos['OBD'].astype(str).map(billing_direct_lookup['TRANSPORT ZONE'])
            
            # Fallback: Si no encontró zona por OBD, buscar por ID_VIAJES (LR Number) en el agregado de distribución
            lr_to_zone = billing_dist_agg.drop_duplicates('ID_VIAJES').set_index('ID_VIAJES')['TRANSPORT ZONE']
            df_directos['TRANSPORT ZONE'] = df_directos['TRANSPORT ZONE'].fillna(df_directos['ID_VIAJES'].map(lr_to_zone))
            
            # Para transportistas que cobran por peso (Andreani/Logicargo), recuperamos los kilos de SAP
            # Intentamos primero por entrega (OBD) y luego por consolidado de viaje (LR)
            obd_to_weight = billing_direct_agg.set_index('OBD')['Gross weight']
            df_directos['Gross weight'] = df_directos['OBD'].astype(str).map(obd_to_weight)

            lr_to_weight = billing_dist_agg.groupby('ID_VIAJES')['Gross weight'].sum()
            df_directos['Gross weight'] = df_directos['Gross weight'].fillna(df_directos['ID_VIAJES'].astype(str).map(lr_to_weight))
            
            df_directos['Ruta'] = df_directos['OBD'].map(billing_direct_lookup['Ruta'])
            # Re-incluyo la lógica para buscar zonas de transporte faltantes
            zonas_conocidas = df_directos[df_directos['TRANSPORT ZONE'].notna() & (df_directos['CLIENTE'].notna())].drop_duplicates(subset=['CLIENTE'])
            mapeo_zonas = pd.Series(zonas_conocidas['TRANSPORT ZONE'].values, index=zonas_conocidas['CLIENTE']).to_dict()
            df_directos['TRANSPORT ZONE'] = df_directos['TRANSPORT ZONE'].fillna(df_directos['CLIENTE'].map(mapeo_zonas))
            # Para viajes simples el match se hace con TRANSPORTE_ALCANCE + TRANSPORT ZONE + UNIDAD.
            df_directos['TRANSPORTE'] = df_directos['TRANSPORTE_ALCANCE']
            df_directos_tarifado = asignar_tarifa_vectorizado(df_directos, tarifario_aux)
            # evitar duplicados si la tarifación produce múltiples filas por ID_FILA
            tarifas_directos_sumadas = df_directos_tarifado.groupby('ID_FILA')['Tarifa'].first().reset_index().rename(columns={'Tarifa':'Tarifa_Directo'})
            before_merge = len(df_final)
            df_final = pd.merge(df_final, tarifas_directos_sumadas, on='ID_FILA', how='left')
            after_merge = len(df_final)
            if after_merge > before_merge:
                dup = df_final['ID_FILA'].duplicated().sum()
                print(f"Advertencia: la fusión de tarifas directas aumentó filas de {before_merge} a {after_merge} ({dup} duplicadas)")
        else:
            df_final['Tarifa_Directo'] = 0.0

        df_final['Tarifa_Directo'] = df_final['Tarifa_Directo'].fillna(0)


        # 5. Calcular Tarifa Total
        print("Calculando Tarifa Total...")

        # Asegurar que las columnas base existan y estén llenas
        for col in ['Tarifa_Alcance', 'Tarifa_Distribucion', 'Tarifa_Directo']:
            if col not in df_final.columns:
                df_final[col] = 0
            df_final[col] = df_final[col].fillna(0)

        # Unificar costos de Alcance y Directo en la columna 'Tarifa_Alcance', según el pedido
        df_final['Tarifa_Alcance'] = df_final['Tarifa_Alcance'] + df_final['Tarifa_Directo']

        # El total es la suma de los dos componentes principales
        df_final['Tarifa_Total'] = df_final['Tarifa_Alcance'] + df_final['Tarifa_Distribucion']

        # Limpieza de columnas temporales, manteniendo las de costos y transportistas para el output
        # no borramos TRANSPORTE_ALCANCE ni TRANSPORTE_XD; estarán en V_BASE
        # sólo eliminamos la columna auxiliar de tarifa directa si existe (la usamos para sumar más arriba)
        df_final.drop(columns=['Tarifa_Directo'], errors='ignore', inplace=True)

        # --- FIN NUEVA LÓGICA ---

        # --- Preparación de hojas de salida (Directos, Distribucion) ---
        # Las variables df_directos_tarifado y df_dist_detalle_tarifado se generan en el nuevo bloque unificado.
        # Si no se crearon (p.ej. no hay viajes de ese tipo), se inicializan como dataframes vacíos para evitar errores.
        if 'df_directos_tarifado' not in locals():
            df_directos_tarifado = pd.DataFrame()
        if 'df_dist_detalle_tarifado' not in locals():
            df_dist_detalle_tarifado = pd.DataFrame()

        df_directos_final = df_directos_tarifado
        df_distribucion_final = pd.DataFrame()
        if not df_dist_detalle_tarifado.empty:
            # Asegurarse de que las columnas existan para evitar errores en el groupby
            for col in ['Cajas', 'Aforo x 900KG', 'X KG']:
                if col not in df_dist_detalle_tarifado.columns:
                    df_dist_detalle_tarifado[col] = np.nan
            
            # Usar solo columnas que seguramente existan en df_dist_detalle_tarifado
            groupby_cols = [col for col in ['ID_VIAJES', 'TRANSPORTE', 'Ship to party', 'Customer_Name_City', 'TRANSPORT ZONE', 'Ruta'] 
                            if col in df_dist_detalle_tarifado.columns]
            
            if groupby_cols:
                df_distribucion_final = df_dist_detalle_tarifado.groupby(groupby_cols).agg({
                    'Gross weight': 'sum',
                    'Cajas': 'sum',
                    'Aforo x 900KG': 'first',
                    'X KG': 'first',
                    'Tarifa': 'sum'
                }).reset_index()

        # --- NUEVA LÓGICA PARA RESUMEN CLIENTES ---
        print(f"Generando Resumen de Clientes")
        # Usar billing_b_full (toda la facturación disponible) para los resúmenes, ya que cubre los viajes del mes

        # Si existe la columna 'source_file' en billing_b_full, usarla; si no, asignar un valor por defecto
        if 'source_file' not in billing_b_full.columns:
            billing_b_full['source_file'] = 'Parquet'

        # Crear el mapeo de source_suffix, manejando tanto archivos individuales como consolidados
        source_map = {
            'AR02_IcADR_F.XLSX': '2',
            'AR06_IcADR_F.XLSX': '6',
            'FACTURAS_2025_Q3.xlsx': 'Q3',
            'Parquet': 'Consolidado',  # Para archivos consolidados de Mac
            'billing_consolidated.parquet': 'Consolidado'
        }
        billing_b_full = agregar_source_suffix(billing_b_full, source_map)
        billing_report_source = agregar_source_suffix(billing_report_source, source_map)

        column_rename_map = {
            'Reference Document number_2': 'OBD_ARA2',
            'Reference Document number_6': 'OBD_ARA6',
            'Reference Document number_Q3': 'OBD_Q3',
            'Reference Document number_Consolidado': 'OBD_Consolidado',
            'Billed quantity (Base UoM)_2': 'CAR_ARA2',
            'Billed quantity (Base UoM)_6': 'CAR_ARA6',
            'Billed quantity (Base UoM)_Q3': 'CAR_Q3',
            'Billed quantity (Base UoM)_Consolidado': 'CAR_Consolidado'
        }

        # Crear el pivot table solo si hay más de una source_suffix única
        unique_sources = billing_report_source['source_suffix'].nunique()
        if unique_sources > 1:
            pivot_inicial = pd.pivot_table(billing_report_source, index=['CLIENTE_MC_NUM', 'CLIENTE_MC_NAME'], columns='source_suffix', values=['Reference Document number', 'Billed quantity (Base UoM)'], aggfunc={'Reference Document number': pd.Series.nunique, 'Billed quantity (Base UoM)': 'sum'})
            pivot_inicial.columns = [f'{val}_{col}' for val, col in pivot_inicial.columns]
            
            # Crear column_rename_map solo para las columnas que existan
            available_renames = {k: v for k, v in column_rename_map.items() if k in pivot_inicial.columns}
            pivot_inicial.rename(columns=available_renames, inplace=True)
        else:
            # Si solo hay una fuente, crear un pivot simple
            pivot_inicial = billing_report_source.groupby(['CLIENTE_MC_NUM', 'CLIENTE_MC_NAME']).agg({
                'Reference Document number': pd.Series.nunique,
                'Billed quantity (Base UoM)': 'sum'
            }).reset_index()
            pivot_inicial.columns = ['CLIENTE_MC_NUM', 'CLIENTE_MC_NAME', 'Total_OBD', 'Total_Cajas']

        if 'Total_OBD' not in pivot_inicial.columns:
            pivot_inicial = pivot_inicial.fillna(0)
            obd_cols = [col for col in pivot_inicial.columns if 'OBD' in col]
            car_cols = [col for col in pivot_inicial.columns if 'CAR' in col]
            if obd_cols:
                pivot_inicial['Total_OBD'] = pivot_inicial[obd_cols].sum(axis=1)
            if car_cols:
                pivot_inicial['Total_Cajas'] = pivot_inicial[car_cols].sum(axis=1)
        else:
            # Renombrar si viene del pivotar simple
            if 'Total_OBD' in pivot_inicial.columns:
                pass  # Ya está bien llamado

        # Agregar desglose por Entidad (CORE/VESTACY) si existe la columna en los datos
        if 'Entidad' in billing_report_source.columns:
            pivot_entidad = pd.pivot_table(
                billing_report_source,
                index=['CLIENTE_MC_NUM', 'CLIENTE_MC_NAME'],
                columns='Entidad',
                values='Billed quantity (Base UoM)',
                aggfunc='sum'
            ).fillna(0)
            pivot_entidad.columns = [f'CAR_{col}' for col in pivot_entidad.columns]
            # Solo nos interesan CORE y VESTACY como columnas adicionales según el pedido
            cols_desglose = [c for c in ['CAR_CORE', 'CAR_VESTACY'] if c in pivot_entidad.columns]
            if cols_desglose:
                pivot_inicial = reset_index_if_needed(pivot_inicial)
                pivot_inicial = pd.merge(pivot_inicial, pivot_entidad.reset_index()[['CLIENTE_MC_NUM', 'CLIENTE_MC_NAME'] + cols_desglose], on=['CLIENTE_MC_NUM', 'CLIENTE_MC_NAME'], how='left').fillna(0)

        df_costos = pd.merge(billing_report_source, df_final, left_on='LR number', right_on='ID_VIAJES', how='left')

        # Separación explícita por regla de negocio:
        # simples: una sola OBD por viaje
        # distribución: costo por viaje y Ship to
        df_dist = df_costos[df_costos['TIPO DE VIAJE'].isin(tipos_distribucion)].copy()
        df_simple = df_costos[df_costos['TIPO DE VIAJE'].isin(tipos_simples)].copy()

        # Para viajes simples, el costo total del viaje se asigna a la entrega.
        df_simple['Costo_Total_Entrega'] = df_simple['Tarifa_Total']

        # Para viajes con distribución, el costo es la tarifa de la entrega individual + el costo de alcance prorrateado.
        if not df_dist.empty and not df_distribucion_final.empty:
            # df_dist ya debe tener 'Ship to party'
            if 'Ship to party' not in df_dist.columns:
                # Si falta, intentamos recuperarla desde billing_b_full
                df_dist = pd.merge(df_dist, billing_b_full[['LR number', 'Ship to party']].drop_duplicates('LR number').rename(columns={'LR number': 'LR number_x'}),
                                left_on='LR number_x', right_on='LR number_x', how='left')
            
            df_dist['Ship to party'] = df_dist['Ship to party'].astype(str)
            
            df_distribucion_merge_cols = [col for col in ['ID_VIAJES', 'Ship to party', 'Tarifa'] if col in df_distribucion_final.columns]
            df_dist = pd.merge(df_dist, df_distribucion_final[df_distribucion_merge_cols],
                            on=['ID_VIAJES', 'Ship to party'], how='left', suffixes=('', '_dist'))
            df_dist.rename(columns={'Tarifa': 'Costo_Por_Entrega', 'Tarifa_dist': 'Tarifa'}, inplace=True)

            # 2. Calcular el costo de alcance proporcional.
            entregas_por_viaje = df_dist.groupby('ID_VIAJES')['Ship to party'].nunique().reset_index().rename(columns={'Ship to party': 'N_Entregas'})
            df_dist = pd.merge(df_dist, entregas_por_viaje, on='ID_VIAJES', how='left')
            
            # 'Tarifa_Alcance' (costo fijo) se divide por el número de entregas.
            df_dist['Alcance_Proporcional'] = df_dist['Tarifa_Alcance'].fillna(0) / df_dist['N_Entregas']
            
            # 3. Calcular el costo total para la entrega.
            df_dist['Costo_Total_Entrega'] = df_dist['Costo_Por_Entrega'].fillna(0) + df_dist['Alcance_Proporcional'].fillna(0)
        else:
            # Si no hay viajes de distribución, asignar 0 para evitar errores
            df_dist['Costo_Total_Entrega'] = 0

        # --- DE-DUPLICACIÓN ANTES DEL RESUMEN FINAL ---
        # Nos aseguramos de contar el costo de cada entrega/viaje una sola vez, eliminando
        # las filas duplicadas que vienen de las líneas de producto en la facturación.
        if not df_dist.empty:
            df_dist.drop_duplicates(subset=['ID_VIAJES', 'Ship to party'] if 'Ship to party' in df_dist.columns else ['ID_VIAJES'], inplace=True)
        if not df_simple.empty:
            df_simple.drop_duplicates(subset=['ID_VIAJES'], inplace=True)

        dfs_to_concat = [df for df in [df_simple, df_dist] if not df.empty]
        df_costos_final = pd.concat(dfs_to_concat, ignore_index=True) if dfs_to_concat else pd.DataFrame()
        # verificar duplicados tras concatenar simples y distribución
        if all(col in df_costos_final.columns for col in ['ID_VIAJES', 'Ship to party']):
            dup_costos = df_costos_final[['ID_VIAJES', 'Ship to party']].duplicated().sum()
            if dup_costos:
                print(f"Advertencia: df_costos_final tiene {dup_costos} filas duplicadas en combinación ID_VIAJES/Ship to party")

        if all(col in df_costos_final.columns for col in ['CLIENTE_MC_NUM', 'CLIENTE_MC_NAME', 'Costo_Total_Entrega']):
            costo_total_por_cliente = df_costos_final.groupby(['CLIENTE_MC_NUM', 'CLIENTE_MC_NAME'])['Costo_Total_Entrega'].sum().reset_index().rename(columns={'Costo_Total_Entrega': 'Tarifa_Total_Cliente'})
        else:
            costo_total_por_cliente = pd.DataFrame(columns=['CLIENTE_MC_NUM', 'CLIENTE_MC_NAME', 'Tarifa_Total_Cliente'])

        df_reporte_final = pd.merge(reset_index_if_needed(pivot_inicial), costo_total_por_cliente, on=['CLIENTE_MC_NUM', 'CLIENTE_MC_NAME'], how='left')
        df_reporte_final['Costo_por_Caja'] = (df_reporte_final['Tarifa_Total_Cliente'] / df_reporte_final['Total_Cajas']).fillna(0)

        # Recalculo de costos por entidad/fuente (ARA2, ARA6, CORE, VESTACY)
        for suffix in ['ARA2', 'ARA6', 'CORE', 'VESTACY']:
            car_col = f'CAR_{suffix}'
            costo_col = f'COSTO_{suffix}'
            if car_col in df_reporte_final.columns:
                # % de cajas de la entidad * Tarifa Total del Cliente
                df_reporte_final[costo_col] = (df_reporte_final[car_col] / df_reporte_final['Total_Cajas'].replace(0, np.nan)) * df_reporte_final['Tarifa_Total_Cliente']
                df_reporte_final[costo_col] = df_reporte_final[costo_col].fillna(0)

        # --- NUEVA LÓGICA PARA RESUMEN SHIP-TO ---
        print(f"Generando Resumen de Ship-To")

        # 1. Crear el pivot de cantidades a nivel de Ship To
        if unique_sources > 1:
            pivot_shipto = pd.pivot_table(
                billing_report_source,
                index=['Ship to party', 'Ship to party name'],
                columns='source_suffix',
                values=['Reference Document number', 'Billed quantity (Base UoM)'],
                aggfunc={'Reference Document number': pd.Series.nunique, 'Billed quantity (Base UoM)': 'sum'}
            )
            pivot_shipto.columns = [f'{val}_{col}' for val, col in pivot_shipto.columns]
            
            shipto_renames = {k: v for k, v in column_rename_map.items() if k in pivot_shipto.columns}
            pivot_shipto.rename(columns=shipto_renames, inplace=True)
            pivot_shipto = pivot_shipto.fillna(0)
            
            # 2. Calcular totales de OBD y Cajas para Ship To
            obd_cols_st = [col for col in pivot_shipto.columns if 'OBD' in col]
            car_cols_st = [col for col in pivot_shipto.columns if 'CAR' in col]
            if obd_cols_st:
                pivot_shipto['Total_OBD'] = pivot_shipto[obd_cols_st].sum(axis=1)
            if car_cols_st:
                pivot_shipto['Total_Cajas'] = pivot_shipto[car_cols_st].sum(axis=1)
        else:
            # Si solo hay una fuente, crear un pivot simple
            pivot_shipto = billing_report_source.groupby(['Ship to party', 'Ship to party name']).agg({
                'Reference Document number': pd.Series.nunique,
                'Billed quantity (Base UoM)': 'sum'
            }).reset_index()
            pivot_shipto.columns = ['Ship to party', 'Ship to party name', 'Total_OBD', 'Total_Cajas']

        # Agregar desglose por Entidad para Ship To
        if 'Entidad' in billing_report_source.columns:
            pivot_entidad_st = pd.pivot_table(
                billing_report_source,
                index=['Ship to party', 'Ship to party name'],
                columns=['Entidad'],
                values='Billed quantity (Base UoM)',
                aggfunc='sum'
            ).fillna(0)
            pivot_entidad_st.columns = [f'CAR_{col}' for col in pivot_entidad_st.columns]
            cols_desglose_st = [c for c in ['CAR_CORE', 'CAR_VESTACY'] if c in pivot_entidad_st.columns]
            if cols_desglose_st:
                pivot_shipto = reset_index_if_needed(pivot_shipto)
                pivot_shipto = pd.merge(pivot_shipto, pivot_entidad_st.reset_index()[['Ship to party', 'Ship to party name'] + cols_desglose_st], on=['Ship to party', 'Ship to party name'], how='left').fillna(0)

        # Agregar desglose por Planta para Ship To
        if 'Plant' in billing_report_source.columns:
            pivot_plant_st = pd.pivot_table(
                billing_report_source,
                index=['Ship to party', 'Ship to party name'],
                columns=['Plant'],
                values='Billed quantity (Base UoM)',
                aggfunc='sum'
            ).fillna(0)
            # Renombrar columnas a CAR_Planta (ej: CAR_AR01)
            pivot_plant_st.columns = [f'CAR_{col}' for col in pivot_plant_st.columns]
            
            pivot_shipto = reset_index_if_needed(pivot_shipto)
            pivot_shipto = pd.merge(pivot_shipto, pivot_plant_st.reset_index(), on=['Ship to party', 'Ship to party name'], how='left').fillna(0)

        # 3. Calcular el costo total por Ship To desde df_costos_final
        # Aseguramos que las columnas para el groupby existan
        groupby_cols_shipto = ['Ship to party', 'Ship to party name']
        if all(col in df_costos_final.columns for col in groupby_cols_shipto):
            costo_total_por_shipto = df_costos_final.groupby(groupby_cols_shipto)['Costo_Total_Entrega'].sum().reset_index().rename(columns={'Costo_Total_Entrega': 'Tarifa_Total_Shipto'})
        else:
            costo_total_por_shipto = pd.DataFrame(columns=groupby_cols_shipto + ['Tarifa_Total_Shipto'])


        # 4. Unir la información de costos con la de cantidades
        df_reporte_shipto_final = pd.merge(
            reset_index_if_needed(pivot_shipto),
            costo_total_por_shipto,
            on=['Ship to party', 'Ship to party name'],
            how='left'
        )
        df_reporte_shipto_final['Tarifa_Total_Shipto'] = df_reporte_shipto_final['Tarifa_Total_Shipto'].fillna(0)

        # 5. Calcular costo por caja a nivel Ship To
        df_reporte_shipto_final['Costo_por_Caja'] = (df_reporte_shipto_final['Tarifa_Total_Shipto'] / df_reporte_shipto_final['Total_Cajas']).replace([np.inf, -np.inf], 0).fillna(0)

        # Recalculo de costos por entidad/fuente/planta a nivel Ship To
        # Identificamos dinámicamente todas las columnas de cajas (CAR_...) para prorratear el costo
        car_cols_to_process = [c for c in df_reporte_shipto_final.columns if c.startswith('CAR_') and c not in ['CAR_Consolidado', 'CAR_Q3']]
        
        for car_col in car_cols_to_process:
            suffix = car_col.replace('CAR_', '')
            costo_col = f'COSTO_{suffix}'
            
            # % de cajas de la entidad/planta * Tarifa Total del Shipto
            df_reporte_shipto_final[costo_col] = (df_reporte_shipto_final[car_col] / df_reporte_shipto_final['Total_Cajas'].replace(0, np.nan)) * df_reporte_shipto_final['Tarifa_Total_Shipto']
            df_reporte_shipto_final[costo_col] = df_reporte_shipto_final[costo_col].fillna(0)


        # --- NUEVA LÓGICA PARA RESUMEN RUTAS ---
        print(f"Generando Resumen de Rutas")

        # 1. Crear el pivot de cantidades a nivel de Ruta
        if unique_sources > 1:
            pivot_ruta = pd.pivot_table(
                billing_report_source,
                index=['Ruta'],
                columns='source_suffix',
                values=['Reference Document number', 'Billed quantity (Base UoM)'],
                aggfunc={'Reference Document number': pd.Series.nunique, 'Billed quantity (Base UoM)': 'sum'}
            )
            pivot_ruta.columns = [f'{val}_{col}' for val, col in pivot_ruta.columns]
            
            ruta_renames = {k: v for k, v in column_rename_map.items() if k in pivot_ruta.columns}
            pivot_ruta.rename(columns=ruta_renames, inplace=True)
            pivot_ruta = pivot_ruta.fillna(0)
            
            # 2. Calcular totales de OBD y Cajas para Ruta
            obd_cols_rt = [col for col in pivot_ruta.columns if 'OBD' in col]
            car_cols_rt = [col for col in pivot_ruta.columns if 'CAR' in col]
            if obd_cols_rt:
                pivot_ruta['Total_OBD'] = pivot_ruta[obd_cols_rt].sum(axis=1)
            if car_cols_rt:
                pivot_ruta['Total_Cajas'] = pivot_ruta[car_cols_rt].sum(axis=1)
        else:
            # Si solo hay una fuente, crear un pivot simple
            pivot_ruta = billing_report_source.groupby('Ruta').agg({
                'Reference Document number': pd.Series.nunique,
                'Billed quantity (Base UoM)': 'sum'
            }).reset_index()
            pivot_ruta.columns = ['Ruta', 'Total_OBD', 'Total_Cajas']

        # 3. Calcular el costo total por Ruta desde df_costos_final
        if 'Ruta' in df_costos_final.columns:
            costo_total_por_ruta = df_costos_final.groupby('Ruta')['Costo_Total_Entrega'].sum().reset_index().rename(columns={'Costo_Total_Entrega': 'Tarifa_Total_Ruta'})
        else:
            costo_total_por_ruta = pd.DataFrame(columns=['Ruta', 'Tarifa_Total_Ruta'])

        # 4. Unir la información de costos con la de cantidades
        df_reporte_ruta_final = pd.merge(
            reset_index_if_needed(pivot_ruta),
            costo_total_por_ruta,
            on='Ruta',
            how='left'
        )
        df_reporte_ruta_final['Tarifa_Total_Ruta'] = df_reporte_ruta_final['Tarifa_Total_Ruta'].fillna(0)

        # 5. Calcular costo por caja a nivel Ruta
        df_reporte_ruta_final['Costo_por_Caja'] = (df_reporte_ruta_final['Tarifa_Total_Ruta'] / df_reporte_ruta_final['Total_Cajas']).replace([np.inf, -np.inf], 0).fillna(0)


        # --- Guardar en Excel ---
        output_filename = os.path.join(DIR_OUTPUTS, f'Provision_Calculada_Mes_{MES_DE_ANALISIS}.xlsx')
        print(f"Guardando resultados en: {output_filename}")
        excel_engine = 'xlsxwriter' if importlib.util.find_spec('xlsxwriter') else 'openpyxl'
        df_directos_final = df_directos_tarifado if 'df_directos_tarifado' in locals() else pd.DataFrame()
        with pd.ExcelWriter(output_filename, engine=excel_engine) as writer:
            df_final.to_excel(writer, sheet_name='V_BASE', index=False)
            df_directos_final.to_excel(writer, sheet_name='Directos', index=False)
            df_distribucion_final.to_excel(writer, sheet_name='Distribucion', index=False)
            df_reporte_shipto_final.to_excel(writer, sheet_name='Resumen Ship To', index=False)
            df_reporte_ruta_final.to_excel(writer, sheet_name='Resumen Rutas', index=False)

            # (Aquí se puede añadir el formato de moneda si se desea)

        print("Procesamiento completado.")

        # --- DIAGNÓSTICO DE TARIFAS ---
        print("\n" + "="*50)
        print("DIAGNÓSTICO DE TARIFAS:")
        print("="*50)
        print(f"Total de viajes procesados: {len(df_final)}")
        print(f"Viajes con Tarifa_Alcance > 0: {len(df_final[df_final['Tarifa_Alcance'] > 0])}")
        print(f"Viajes con Tarifa_Distribucion > 0: {len(df_final[df_final['Tarifa_Distribucion'] > 0])}")
        print(f"Viajes con Tarifa_Total > 0: {len(df_final[df_final['Tarifa_Total'] > 0])}")
        print(f"Viajes con Tarifa_Total == 0: {len(df_final[df_final['Tarifa_Total'] == 0])}")
        print(f"Viajes con Tarifa_Total < 0: {len(df_final[df_final['Tarifa_Total'] < 0])}")

        # Mostrar motivos de tarifa
        if 'Motivo_Tarifa' in df_final.columns:
            motivos = df_final['Motivo_Tarifa'].value_counts()
            print(f"\nMotivos de tarifa encontrados:")
            for motivo, count in motivos.items():
                print(f"  {motivo}: {count}")

        print("="*50)
        return True, output_filename

    except FileNotFoundError as e:
        return False, f"No se encontró el archivo: {getattr(e, 'filename', 'desconocido')}"
    except Exception as e:
        return False, f"Error: {str(e)}"

if __name__ == "__main__":
    try:
        mes_input = input("Ingrese el número del mes a procesar (1-12): ")
        if mes_input.isdigit() and 1 <= int(mes_input) <= 12:
            success, result = ejecutar_calculo_provision(int(mes_input))
            if success:
                print("\n" + "="*60)
                print(f"¡PROCESO FINALIZADO CON ÉXITO!")
                print(f"Archivo generado en: {os.path.abspath(result)}")
                print("="*60 + "\n")
            else:
                print(f"\n[ERROR] El proceso falló: {result}")
        else:
            print("Error: Debe ingresar un número de mes válido entre 1 y 12.")
    except KeyboardInterrupt:
        print("\nProceso cancelado por el usuario.")