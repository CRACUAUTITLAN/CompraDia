import streamlit as st
import pandas as pd
import io
import datetime
from dateutil.relativedelta import relativedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# Configuración de la página
st.set_page_config(page_title="Analizador BI - Grupo Andrade", layout="wide")
st.title("🔧 Herramienta BI: Compras e Inventarios (Fase 4 - HITS Automáticos)")

# --- CONFIGURACIÓN GOOGLE DRIVE ---
try:
    if "gcp_service_account" in st.secrets and "general" in st.secrets:
        gcp_creds = dict(st.secrets["gcp_service_account"])
        PARENT_FOLDER_ID = st.secrets["general"]["drive_folder_id"] # Carpeta de SALIDA (Shared Drive)
        MASTER_SALES_ID = st.secrets["general"].get("master_sales_id") # Carpeta de LECTURA (Tu Drive Personal)
        
        creds = service_account.Credentials.from_service_account_info(
            gcp_creds, scopes=['https://www.googleapis.com/auth/drive']
        )
        drive_service = build('drive', 'v3', credentials=creds)
        st.success("✅ Robot Conectado (Lectura Histórica + Escritura Reportes).")
    else:
        st.error("❌ Faltan secretos. Revisa la configuración.")
        st.stop()
except Exception as e:
    st.error(f"⚠️ Error de conexión: {e}")
    st.stop()

# --- FUNCIONES DRIVE (LECTURA Y ESCRITURA) ---

def buscar_o_crear_carpeta(nombre_carpeta, parent_id):
    try:
        # Busca carpeta de salida (Shared Drive)
        query = f"mimeType='application/vnd.google-apps.folder' and name='{nombre_carpeta}' and '{parent_id}' in parents and trashed=false"
        results = drive_service.files().list(
            q=query, fields="files(id, name)", supportsAllDrives=True, includeItemsFromAllDrives=True
        ).execute()
        files = results.get('files', [])
        
        if files:
            return files[0]['id']
        else:
            metadata = {'name': nombre_carpeta, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
            folder = drive_service.files().create(
                body=metadata, fields='id', supportsAllDrives=True
            ).execute()
            return folder.get('id')
    except Exception as e:
        return None

def subir_excel_a_drive(buffer, nombre_archivo):
    try:
        fecha_hoy = datetime.datetime.now()
        anio = str(fecha_hoy.year)
        meses_es = {1:"01_Enero", 2:"02_Febrero", 3:"03_Marzo", 4:"04_Abril", 5:"05_Mayo", 6:"06_Junio", 7:"07_Julio", 8:"08_Agosto", 9:"09_Septiembre", 10:"10_Octubre", 11:"11_Noviembre", 12:"12_Diciembre"}
        mes_carpeta = meses_es[fecha_hoy.month]

        # Estructura de carpetas en Shared Drive
        id_anio = buscar_o_crear_carpeta(anio, PARENT_FOLDER_ID)
        if not id_anio: return None
        id_mes = buscar_o_crear_carpeta(mes_carpeta, id_anio)
        if not id_mes: return None
        
        media = MediaIoBaseUpload(buffer, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        file_metadata = {'name': nombre_archivo, 'parents': [id_mes]}
        
        # Subida con soporte para Shared Drives
        archivo = drive_service.files().create(
            body=file_metadata, media_body=media, fields='id, webViewLink', supportsAllDrives=True
        ).execute()
        return archivo.get('webViewLink')
    except Exception as e:
        st.error(f"Error subiendo: {e}")
        return None

def descargar_archivo_drive(file_id):
    """Descarga un archivo de Drive a memoria para que Pandas lo lea."""
    try:
        request = drive_service.files().get_media(fileId=file_id)
        file = io.BytesIO()
        downloader = MediaIoBaseDownload(file, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
        file.seek(0)
        return file
    except Exception as e:
        st.error(f"Error descargando ID {file_id}: {e}")
        return None

def buscar_archivos_ventas(agencia, anios):
    """
    Busca en la carpeta DATA_MASTER_VENTAS (Tu Drive Personal)
    """
    archivos_encontrados = []
    if not MASTER_SALES_ID:
        st.warning("⚠️ No configuraste 'master_sales_id'.")
        return []

    for anio in anios:
        # Filtro: Nombre contiene AGENCIA + AÑO + "MASTER"
        # Importante: supportsAllDrives=True sirve aunque sea 'Mi Unidad' para evitar errores de API
        query = f"name contains '{agencia}' and name contains '{anio}' and name contains 'MASTER' and '{MASTER_SALES_ID}' in parents and trashed=false"
        
        results = drive_service.files().list(
            q=query, fields="files(id, name)", supportsAllDrives=True, includeItemsFromAllDrives=True
        ).execute()
        files = results.get('files', [])
        archivos_encontrados.extend(files)
        
    return archivos_encontrados

# --- LOGICA BI (HITS HISTORICOS) ---

def calcular_hits_historicos(agencia):
    """
    Calcula HITS mensuales (meses completos cerrados).
    Ej: Si hoy es Feb 2026, rango = Feb 2025 al 31 Ene 2026.
    """
    st.info(f"🔄 Buscando históricos para {agencia} en Drive...")
    
    # 1. Definir Fechas (Meses Completos)
    hoy = datetime.datetime.now()
    # Primer día del mes actual (Límite superior, no inclusivo)
    fecha_fin = hoy.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    # Un año antes exacto (Inicio)
    fecha_inicio = fecha_fin - relativedelta(years=1)
    
    st.write(f"📅 Rango de análisis para HITS: **{fecha_inicio.strftime('%d-%b-%Y')}** al **{fecha_fin.strftime('%d-%b-%Y')}**")
    
    # Años involucrados para buscar archivos
    anios_necesarios = list(set([fecha_inicio.year, fecha_fin.year])) # Ej: [2025, 2026]
    
    # 2. Buscar Archivos
    files_metadata = buscar_archivos_ventas(agencia.upper(), anios_necesarios)
    
    if not files_metadata:
        st.warning(f"⚠️ No se encontraron archivos 'MASTER' para {agencia} ({anios_necesarios}) en la carpeta configurada.")
        return None

    dfs = []
    # 3. Descargar y Leer
    progress_bar = st.progress(0)
    for i, file_meta in enumerate(files_metadata):
        st.caption(f"Leyendo: {file_meta['name']}...")
        content = descargar_archivo_drive(file_meta['id'])
        if content:
            try:
                engine = 'xlrd' if 'xls' in file_meta['name'] and 'xlsx' not in file_meta['name'] else 'openpyxl'
                df_temp = pd.read_excel(content, engine=engine)
                df_temp.columns = df_temp.columns.str.upper().str.strip()
                dfs.append(df_temp)
            except Exception as e:
                st.error(f"Error leyendo {file_meta['name']}: {e}")
        progress_bar.progress((i + 1) / len(files_metadata))
    
    if not dfs: return None
    
    # 4. Concatenar
    df_total = pd.concat(dfs, ignore_index=True)
    
    # 5. Filtrar por Fechas
    if 'FECHA' not in df_total.columns:
        st.warning("⚠️ El archivo histórico no tiene columna 'FECHA'.")
        return None
        
    df_total['FECHA'] = pd.to_datetime(df_total['FECHA'], errors='coerce')
    
    # Filtro estricto: >= Inicio Y < Fin (Fin es 1ro del mes actual, así que toma hasta el último del mes anterior)
    mask = (df_total['FECHA'] >= fecha_inicio) & (df_total['FECHA'] < fecha_fin)
    df_filtrado = df_total.loc[mask].copy()
    
    if df_filtrado.empty:
        st.warning("⚠️ Archivos encontrados, pero no hay ventas en el rango de fechas seleccionado.")
        return None

    # 6. Cálculo Matemático de HITS
    if 'NP' not in df_filtrado.columns or 'CANTIDAD' not in df_filtrado.columns:
        st.error("Las bases históricas no tienen columna 'NP' o 'CANTIDAD'.")
        return None
        
    df_filtrado['NP'] = df_filtrado['NP'].astype(str).str.strip()
    df_filtrado['CANTIDAD'] = pd.to_numeric(df_filtrado['CANTIDAD'], errors='coerce').fillna(0)
    
    # Agrupar
    resumen = df_filtrado.groupby('NP').agg(
        total_eventos=('CANTIDAD', 'count'),
        eventos_negativos=('CANTIDAD', lambda x: (x < 0).sum())
    ).reset_index()
    
    # FÓRMULA MAESTRA: HITS = Eventos - (Devoluciones * 2)
    resumen['HITS_CALCULADO'] = resumen['total_eventos'] - (resumen['eventos_negativos'] * 2)
    resumen['HITS_CALCULADO'] = resumen['HITS_CALCULADO'].clip(lower=0) # Evitar negativos
    
    st.success(f"✅ HITS calculados para {agencia}: {len(resumen)} productos.")
    return resumen[['NP', 'HITS_CALCULADO']]

# --- FUNCIONES PANDAS PROCESAMIENTO NORMAL ---

COLS_CUAUTITLAN_ORDEN = [
    "N° PARTE", "SUGERIDO DIA", "POR FINCAR", "(Consumo Mensual / 2) - Inv Tran",
    "EXISTENCIA", "FECHA DE ULTIMA COMPRA", "PROMEDIO CUAUTITLAN", "HITS", 
    "CONSUMO MENSUAL", "2", "INVENTARIO TULTITLAN", "PROMEDIO TULTITLAN", 
    "HITS_FORANEO", "TRASPASO TULTI A CUATI", "NUEVO TRASPASO", "CANTIDAD A TRASPASAR", 
    "Fec ult Comp TULTI", "TRANSITO", "INV. TOTAL", "MESES VENTA ACTUAL", 
    "MESES VENTA SUGERIDO", "Line Value", "Cycle Count", "Status", "Ship Multiple", 
    "Last 12 Month Demand", "Current Month Demand", "Job Quantity", "Full Bin", 
    "Bin Location", "Dealer On Hand", "Stock on Order", "Stock On Back Order", "Reason Code"
]

COLS_TULTITLAN_ORDEN = [
    "N° DE PARTE", "SUGERIDO DIA", "POR FINCAR", "(Consumo Mensual / 2) - Inv Tran",
    "EXISTENCIA", "FECHA DE ULTIMA COMPRA", "PROMEDIO TULTITLAN", "HITS", 
    "CONSUMO MENSUAL", "2", "INVENTARIO CUAUTITLAN", "PROMEDIO CUAUTITLAN", 
    "HITS_FORANEO", "TRASPASO CUAUT A TULTI", "NUEVO TRASPASO", "CANTIDAD A TRASPASAR", 
    "Fec ult Comp CUAUTI", "TRANSITO", "INV. TOTAL", "MESES VENTA ACTUAL", 
    "MESES VENTA SUGERIDO", "Line Value", "Cycle Count", "Status", "Ship Multiple", 
    "Last 12 Month Demand", "Current Month Demand", "Job Quantity", "Full Bin", 
    "Bin Location", "Dealer On Hand", "Stock on Order", "Stock On Back Order", "Reason Code"
]

def limpiar_inventario(archivo, nombre_sucursal):
    try:
        engine = 'xlrd' if archivo.name.endswith('.xls') else 'openpyxl'
        df = pd.read_excel(archivo, header=None, engine=engine)
        col_indices = [0, 1, 2, 4, 8, 9, 10, 11]
        col_names = ["N° PARTE", "DESCR", "CLASIF", "PRECIO UNITARIO", "EXIST", "FEC INGRESO", "FEC ULT COMP", "FEC ULT VTA"]
        df_clean = df.iloc[:, col_indices].copy()
        df_clean.columns = col_names
        df_clean = df_clean.dropna(subset=["FEC INGRESO"])
        df_clean["N° PARTE"] = df_clean["N° PARTE"].astype(str).str.strip()
        return df_clean
    except Exception as e:
        return None

def cargar_base_sugerido(archivo):
    try:
        df = pd.read_excel(archivo)
        df.columns = df.columns.str.strip()
        df["N° PARTE"] = df["N° PARTE"].astype(str).str.strip()
        if "Last 12 Month Demand" in df.columns:
            df["Last 12 Month Demand"] = pd.to_numeric(df["Last 12 Month Demand"], errors='coerce').fillna(0)
            df["CONSUMO MENSUAL"] = df["Last 12 Month Demand"] / 12
        else:
            df["CONSUMO MENSUAL"] = 0
        df["2"] = df["CONSUMO MENSUAL"] / 2
        return df
    except Exception as e:
        return None

def procesar_transito(archivo):
    try:
        df = pd.read_excel(archivo)
        df.columns = df.columns.str.strip() 
        cols_necesarias = ["N° PARTE", "TRANSITO"]
        for col in cols_necesarias:
            if col not in df.columns: df[col] = 0
        df_resumen = df[cols_necesarias].copy()
        df_resumen["N° PARTE"] = df_resumen["N° PARTE"].astype(str).str.strip()
        return df_resumen.groupby("N° PARTE", as_index=False)["TRANSITO"].sum()
    except Exception:
        return None

def procesar_traspasos(archivo, filtro):
    try:
        engine = 'xlrd' if archivo.name.endswith('.xls') else 'openpyxl'
        df = pd.read_excel(archivo, header=None, engine=engine)
        df_filtrado = df[df[0].astype(str).str.strip() == filtro].copy()
        if df_filtrado.empty: return pd.DataFrame(columns=["N° PARTE", "CANTIDAD_TRASPASO"])
        df_resumen = df_filtrado[[2, 4]].copy()
        df_resumen.columns = ["N° PARTE", "CANTIDAD_TRASPASO"]
        df_resumen["N° PARTE"] = df_resumen["N° PARTE"].astype(str).str.strip()
        df_resumen["CANTIDAD_TRASPASO"] = pd.to_numeric(df_resumen["CANTIDAD_TRASPASO"], errors='coerce').fillna(0).abs()
        return df_resumen.groupby("N° PARTE", as_index=False)["CANTIDAD_TRASPASO"].sum()
    except Exception:
        return None

def completar_y_ordenar(df, lista_columnas_deseadas):
    for col in lista_columnas_deseadas:
        if col not in df.columns: df[col] = 0 
    df = df[lista_columnas_deseadas].fillna(0)
    return df

# --- INTERFAZ GRAFICA ---

st.info("📂 Los archivos se subirán a Google Drive (Unidad Compartida). Se leerá historial de 'Mi Unidad'.")

# PASO 1
st.header("Paso 1: Bases Iniciales (Sugeridos)")
c1, c2 = st.columns(2)
file_sugerido_cuauti = c1.file_uploader("📂 Sugerido Cuautitlán", type=["xlsx"], key="sug_cuauti")
file_sugerido_tulti = c2.file_uploader("📂 Sugerido Tultitlán", type=["xlsx"], key="sug_tulti")

st.markdown("---")

# PASO 2
st.header("Paso 2: Reportes de Tránsito")
c3, c4 = st.columns(2)
file_transito_cuauti = c3.file_uploader("🚢 Tránsito Cuautitlán", type=["xlsx"], key="trans_cuauti")
file_transito_tulti = c4.file_uploader("🚢 Tránsito Tultitlán", type=["xlsx"], key="trans_tulti")

st.markdown("---")

# PASO 3
st.header("Paso 3: Reportes de Situación (Traspasos)")
c5, c6 = st.columns(2)
file_situacion_para_cuauti = c5.file_uploader("🚛 Situación Cuautitlán (Busca TRASUCTU)", type=["xlsx", "xls"], key="sit_cuauti")
file_situacion_para_tulti = c6.file_uploader("🚛 Situación Tultitlán (Busca TRASUCCU)", type=["xlsx", "xls"], key="sit_tulti")

st.markdown("---")

# PASO 4
st.header("Paso 4: Inventarios (Almacén)")
c7, c8 = st.columns(2)
file_inv_cuauti = c7.file_uploader("📦 Inventario Cuautitlán", type=["xlsx", "xls"], key="inv_cuauti")
file_inv_tulti = c8.file_uploader("📦 Inventario Tultitlán", type=["xlsx", "xls"], key="inv_tulti")

if st.button("🚀 Procesar Todo, Calcular HITS y Subir"):
    if file_sugerido_cuauti and file_sugerido_tulti and file_inv_cuauti and file_inv_tulti:
        
        # --- A. CALCULO DE HITS HISTORICOS (LO PRIMERO) ---
        # El programa busca en TU carpeta personal los Excels maestros
        df_hits_cuauti = calcular_hits_historicos("CUAUTITLAN")
        df_hits_tulti = calcular_hits_historicos("TULTITLAN")

        with st.spinner('Procesando bases locales...'):
            # --- B. PROCESAMIENTO NORMAL ---
            df_base_cuauti = cargar_base_sugerido(file_sugerido_cuauti)
            df_base_tulti = cargar_base_sugerido(file_sugerido_tulti)
            df_inv_cuauti_clean = limpiar_inventario(file_inv_cuauti, "Cuautitlán")
            df_inv_tulti_clean = limpiar_inventario(file_inv_tulti, "Tultitlán")
            
            df_trans_cuauti = procesar_transito(file_transito_cuauti) if file_transito_cuauti else pd.DataFrame(columns=["N° PARTE", "TRANSITO"])
            df_trans_tulti = procesar_transito(file_transito_tulti) if file_transito_tulti else pd.DataFrame(columns=["N° PARTE", "TRANSITO"])

            df_traspasos_a_cuauti = procesar_traspasos(file_situacion_para_cuauti, "TRASUCTU") if file_situacion_para_cuauti else pd.DataFrame(columns=["N° PARTE", "CANTIDAD_TRASPASO"])
            df_traspasos_a_tulti = procesar_traspasos(file_situacion_para_tulti, "TRASUCCU") if file_situacion_para_tulti else pd.DataFrame(columns=["N° PARTE", "CANTIDAD_TRASPASO"])

            if (df_base_cuauti is not None and df_base_tulti is not None):
                
                # === 1. HOJA DIA CUAUTITLAN ===
                df_final_cuauti = df_base_cuauti.copy()
                
                # Inyectar HITS calculados
                if df_hits_cuauti is not None:
                    df_hits_cuauti.rename(columns={'NP': 'N° PARTE', 'HITS_CALCULADO': 'HITS'}, inplace=True)
                    if 'HITS' in df_final_cuauti.columns: del df_final_cuauti['HITS']
                    df_final_cuauti = pd.merge(df_final_cuauti, df_hits_cuauti, on='N° PARTE', how='left')

                # Merges...
                df_final_cuauti = pd.merge(df_final_cuauti, df_inv_cuauti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                df_final_cuauti.rename(columns={'EXIST': 'EXISTENCIA', 'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'}, inplace=True)
                df_final_cuauti = pd.merge(df_final_cuauti, df_inv_tulti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                df_final_cuauti.rename(columns={'EXIST': 'INVENTARIO TULTITLAN', 'FEC ULT COMP': 'Fec ult Comp TULTI'}, inplace=True)
                df_final_cuauti = pd.merge(df_final_cuauti, df_trans_cuauti, on='N° PARTE', how='left')
                df_final_cuauti = pd.merge(df_final_cuauti, df_traspasos_a_cuauti, on='N° PARTE', how='left')
                df_final_cuauti.rename(columns={'CANTIDAD_TRASPASO': 'TRASPASO TULTI A CUATI'}, inplace=True)
                
                # HITS Foraneos (Tulti)
                if df_hits_tulti is not None:
                    temp_hits = df_hits_tulti.copy()
                    temp_hits.rename(columns={'NP': 'N° PARTE', 'HITS_CALCULADO': 'HITS_FORANEO'}, inplace=True)
                    df_final_cuauti = pd.merge(df_final_cuauti, temp_hits, on='N° PARTE', how='left')
                
                df_final_cuauti = completar_y_ordenar(df_final_cuauti, COLS_CUAUTITLAN_ORDEN)
                df_export_cuauti = df_final_cuauti.copy()
                df_export_cuauti.rename(columns={'HITS_FORANEO': 'HITS'}, inplace=True)

                # === 2. HOJA DIA TULTITLAN ===
                df_final_tulti = df_base_tulti.copy()
                
                if df_hits_tulti is not None:
                    temp_hits_t = df_hits_tulti.copy()
                    temp_hits_t.rename(columns={'NP': 'N° PARTE', 'HITS_CALCULADO': 'HITS'}, inplace=True)
                    if 'HITS' in df_final_tulti.columns: del df_final_tulti['HITS']
                    df_final_tulti = pd.merge(df_final_tulti, temp_hits_t, on='N° PARTE', how='left')

                df_final_tulti = pd.merge(df_final_tulti, df_inv_tulti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                df_final_tulti.rename(columns={'EXIST': 'EXISTENCIA', 'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'}, inplace=True)
                df_final_tulti = pd.merge(df_final_tulti, df_inv_cuauti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                df_final_tulti.rename(columns={'EXIST': 'INVENTARIO CUAUTITLAN', 'FEC ULT COMP': 'Fec ult Comp CUAUTI'}, inplace=True)
                df_final_tulti = pd.merge(df_final_tulti, df_trans_tulti, on='N° PARTE', how='left')
                df_final_tulti = pd.merge(df_final_tulti, df_traspasos_a_tulti, on='N° PARTE', how='left')
                df_final_tulti.rename(columns={'CANTIDAD_TRASPASO': 'TRASPASO CUAUT A TULTI'}, inplace=True)
                
                # HITS Foraneos (Cuauti)
                if df_hits_cuauti is not None:
                    temp_hits_c = df_hits_cuauti.copy()
                    temp_hits_c.rename(columns={'NP': 'N° PARTE', 'HITS_CALCULADO': 'HITS_FORANEO'}, inplace=True)
                    df_final_tulti = pd.merge(df_final_tulti, temp_hits_c, on='N° PARTE', how='left')

                df_final_tulti.rename(columns={'N° PARTE': 'N° DE PARTE'}, inplace=True)
                df_final_tulti = completar_y_ordenar(df_final_tulti, COLS_TULTITLAN_ORDEN)
                df_export_tulti = df_final_tulti.copy()
                df_export_tulti.rename(columns={'HITS_FORANEO': 'HITS'}, inplace=True)

                # === 3. SUBIDA ===
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_export_cuauti.to_excel(writer, sheet_name='DIA CUAUTITLAN', index=False)
                    df_export_tulti.to_excel(writer, sheet_name='DIA TULTITLAN', index=False)
                buffer.seek(0)
                
                fecha_hoy_str = datetime.datetime.now().strftime("%d_%m_%Y")
                hora_str = datetime.datetime.now().strftime("%H%M") 
                nombre_archivo_final = f"Analisis_Compras_{fecha_hoy_str}_{hora_str}.xlsx"
                link = subir_excel_a_drive(buffer, nombre_archivo_final)
                
                if link:
                    st.success(f"✅ ¡Proceso Completo! Archivo: {nombre_archivo_final}")
                    st.markdown(f"### [📂 Ver en Drive]({link})")
                    st.dataframe(df_final_cuauti.head())
    else:
        st.warning("⚠️ Carga al menos Sugeridos e Inventarios.")
