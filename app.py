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
st.title("🔧 Herramienta BI: Compras e Inventarios (Fase 4 - HITS Cruzados)")

# --- CONFIGURACIÓN GOOGLE DRIVE ---
try:
    if "gcp_service_account" in st.secrets and "general" in st.secrets:
        gcp_creds = dict(st.secrets["gcp_service_account"])
        PARENT_FOLDER_ID = st.secrets["general"]["drive_folder_id"] # Salida (Shared Drive)
        MASTER_SALES_ID = st.secrets["general"].get("master_sales_id") # Lectura (Tu Drive)
        
        creds = service_account.Credentials.from_service_account_info(
            gcp_creds, scopes=['https://www.googleapis.com/auth/drive']
        )
        drive_service = build('drive', 'v3', credentials=creds)
        st.success("✅ Robot Conectado.")
    else:
        st.error("❌ Faltan secretos.")
        st.stop()
except Exception as e:
    st.error(f"⚠️ Error de conexión: {e}")
    st.stop()

# --- FUNCIONES DRIVE ---

def buscar_o_crear_carpeta(nombre_carpeta, parent_id):
    try:
        query = f"mimeType='application/vnd.google-apps.folder' and name='{nombre_carpeta}' and '{parent_id}' in parents and trashed=false"
        results = drive_service.files().list(
            q=query, fields="files(id, name)", supportsAllDrives=True, includeItemsFromAllDrives=True
        ).execute()
        files = results.get('files', [])
        if files: return files[0]['id']
        else:
            metadata = {'name': nombre_carpeta, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
            folder = drive_service.files().create(body=metadata, fields='id', supportsAllDrives=True).execute()
            return folder.get('id')
    except Exception: return None

def subir_excel_a_drive(buffer, nombre_archivo):
    try:
        fecha_hoy = datetime.datetime.now()
        anio = str(fecha_hoy.year)
        meses_es = {1:"01_Enero", 2:"02_Febrero", 3:"03_Marzo", 4:"04_Abril", 5:"05_Mayo", 6:"06_Junio", 7:"07_Julio", 8:"08_Agosto", 9:"09_Septiembre", 10:"10_Octubre", 11:"11_Noviembre", 12:"12_Diciembre"}
        mes_carpeta = meses_es[fecha_hoy.month]

        id_anio = buscar_o_crear_carpeta(anio, PARENT_FOLDER_ID)
        if not id_anio: return None
        id_mes = buscar_o_crear_carpeta(mes_carpeta, id_anio)
        if not id_mes: return None
        
        media = MediaIoBaseUpload(buffer, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        file_metadata = {'name': nombre_archivo, 'parents': [id_mes]}
        archivo = drive_service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()
        return archivo.get('webViewLink')
    except Exception as e:
        st.error(f"Error subiendo: {e}")
        return None

def descargar_archivo_drive(file_id):
    try:
        request = drive_service.files().get_media(fileId=file_id)
        file = io.BytesIO()
        downloader = MediaIoBaseDownload(file, request)
        done = False
        while done is False: status, done = downloader.next_chunk()
        file.seek(0)
        return file
    except Exception as e: return None

def buscar_archivos_ventas(agencia, anios):
    archivos_encontrados = []
    if not MASTER_SALES_ID:
        st.warning("⚠️ No configuraste 'master_sales_id'.")
        return []

    for anio in anios:
        # Busca por NOMBRE DEL ARCHIVO (Ej: CUAUTITLAN_Ventas_2025_MASTER)
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
    Calcula HITS con la fórmula: Total Eventos - (Negativos * 2).
    """
    st.info(f"🔄 Consultando Drive para ventas de: {agencia} ...")
    
    # 1. Fechas (12 meses cerrados anteriores)
    hoy = datetime.datetime.now()
    fecha_fin = hoy.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    fecha_inicio = fecha_fin - relativedelta(years=1)
    
    st.caption(f"📅 Rango Histórico: {fecha_inicio.strftime('%d-%b-%Y')} al {fecha_fin.strftime('%d-%b-%Y')}")
    
    anios = list(set([fecha_inicio.year, fecha_fin.year]))
    
    # 2. Buscar Archivos por Nombre
    files_metadata = buscar_archivos_ventas(agencia.upper(), anios)
    
    if not files_metadata:
        st.warning(f"⚠️ No se encontraron archivos MASTER para {agencia} en {anios}.")
        return None

    dfs = []
    # 3. Descargar
    bar = st.progress(0)
    for i, file_meta in enumerate(files_metadata):
        content = descargar_archivo_drive(file_meta['id'])
        if content:
            try:
                engine = 'xlrd' if 'xls' in file_meta['name'] and 'xlsx' not in file_meta['name'] else 'openpyxl'
                df_temp = pd.read_excel(content, engine=engine)
                df_temp.columns = df_temp.columns.str.upper().str.strip()
                dfs.append(df_temp)
            except Exception: pass
        bar.progress((i + 1) / len(files_metadata))
    
    if not dfs: return None
    
    # 4. Procesar
    df_total = pd.concat(dfs, ignore_index=True)
    if 'FECHA' not in df_total.columns: return None
        
    df_total['FECHA'] = pd.to_datetime(df_total['FECHA'], errors='coerce')
    mask = (df_total['FECHA'] >= fecha_inicio) & (df_total['FECHA'] < fecha_fin)
    df_filtrado = df_total.loc[mask].copy()
    
    if df_filtrado.empty: return None

    # 5. Calculo Matemático Correcto
    if 'NP' not in df_filtrado.columns or 'CANTIDAD' not in df_filtrado.columns: return None
        
    df_filtrado['NP'] = df_filtrado['NP'].astype(str).str.strip()
    df_filtrado['CANTIDAD'] = pd.to_numeric(df_filtrado['CANTIDAD'], errors='coerce').fillna(0)
    
    # Agrupar por NP
    resumen = df_filtrado.groupby('NP').agg(
        total_eventos=('CANTIDAD', 'count'),  # Cuenta todas las filas (9)
        eventos_negativos=('CANTIDAD', lambda x: (x < 0).sum()) # Cuenta las negativas (1)
    ).reset_index()
    
    # Formula: 9 - (1*2) = 7
    resumen['HITS_CALCULADO'] = resumen['total_eventos'] - (resumen['eventos_negativos'] * 2)
    resumen['HITS_CALCULADO'] = resumen['HITS_CALCULADO'].clip(lower=0) 
    
    st.success(f"✅ HITS {agencia} calculados ({len(resumen)} productos).")
    return resumen[['NP', 'HITS_CALCULADO']]

# --- PROCESAMIENTO PANDAS ---

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
    except Exception: return None

def cargar_base_sugerido(archivo):
    try:
        df = pd.read_excel(archivo)
        df.columns = df.columns.str.strip()
        df["N° PARTE"] = df["N° PARTE"].astype(str).str.strip()
        if "Last 12 Month Demand" in df.columns:
            df["Last 12 Month Demand"] = pd.to_numeric(df["Last 12 Month Demand"], errors='coerce').fillna(0)
            df["CONSUMO MENSUAL"] = df["Last 12 Month Demand"] / 12
        else: df["CONSUMO MENSUAL"] = 0
        df["2"] = df["CONSUMO MENSUAL"] / 2
        return df
    except Exception: return None

def procesar_transito(archivo):
    try:
        df = pd.read_excel(archivo)
        df.columns = df.columns.str.strip() 
        cols = ["N° PARTE", "TRANSITO"]
        for c in cols: 
            if c not in df.columns: df[c] = 0
        df = df[cols].copy()
        df["N° PARTE"] = df["N° PARTE"].astype(str).str.strip()
        return df.groupby("N° PARTE", as_index=False)["TRANSITO"].sum()
    except Exception: return None

def procesar_traspasos(archivo, filtro):
    try:
        engine = 'xlrd' if archivo.name.endswith('.xls') else 'openpyxl'
        df = pd.read_excel(archivo, header=None, engine=engine)
        df = df[df[0].astype(str).str.strip() == filtro].copy()
        if df.empty: return pd.DataFrame(columns=["N° PARTE", "CANTIDAD_TRASPASO"])
        df = df[[2, 4]].copy()
        df.columns = ["N° PARTE", "CANTIDAD_TRASPASO"]
        df["N° PARTE"] = df["N° PARTE"].astype(str).str.strip()
        df["CANTIDAD_TRASPASO"] = pd.to_numeric(df["CANTIDAD_TRASPASO"], errors='coerce').fillna(0).abs()
        return df.groupby("N° PARTE", as_index=False)["CANTIDAD_TRASPASO"].sum()
    except Exception: return None

def completar_y_ordenar(df, lista_columnas_deseadas):
    for col in lista_columnas_deseadas:
        if col not in df.columns: df[col] = 0 
    df = df[lista_columnas_deseadas].fillna(0)
    return df

# --- INTERFAZ GRAFICA ---

st.info("📂 Los archivos se subirán a Google Drive (Unidad Compartida). Se leerá historial de 'Mi Unidad'.")

# PASOS 1, 2, 3, 4
st.header("Paso 1: Bases Iniciales (Sugeridos)")
c1, c2 = st.columns(2)
file_sug_cuauti = c1.file_uploader("📂 Sugerido Cuauti", type=["xlsx"], key="sc")
file_sug_tulti = c2.file_uploader("📂 Sugerido Tulti", type=["xlsx"], key="st")

st.markdown("---")
st.header("Paso 2: Tránsito")
c3, c4 = st.columns(2)
file_trans_cuauti = c3.file_uploader("🚢 Tránsito Cuauti", type=["xlsx"], key="tc")
file_trans_tulti = c4.file_uploader("🚢 Tránsito Tulti", type=["xlsx"], key="tt")

st.markdown("---")
st.header("Paso 3: Traspasos (Situación)")
c5, c6 = st.columns(2)
file_sit_cuauti = c5.file_uploader("🚛 Sit. Cuauti (TRASUCTU)", type=["xlsx", "xls"], key="sic")
file_sit_tulti = c6.file_uploader("🚛 Sit. Tulti (TRASUCCU)", type=["xlsx", "xls"], key="sit")

st.markdown("---")
st.header("Paso 4: Inventarios")
c7, c8 = st.columns(2)
file_inv_cuauti = c7.file_uploader("📦 Inv. Cuauti", type=["xlsx", "xls"], key="ic")
file_inv_tulti = c8.file_uploader("📦 Inv. Tulti", type=["xlsx", "xls"], key="it")

if st.button("🚀 PROCESAR TODO"):
    if file_sug_cuauti and file_sug_tulti and file_inv_cuauti and file_inv_tulti:
        
        # --- A. HITS HISTORICOS (CRUCE DE INFORMACION) ---
        # 1. Calculamos Hits de Cuautitlan (Usando archivos CUAUTITLAN_..._MASTER)
        df_hits_origen_cuauti = calcular_hits_historicos("CUAUTITLAN")
        
        # 2. Calculamos Hits de Tultitlan (Usando archivos TULTITLAN_..._MASTER)
        df_hits_origen_tulti = calcular_hits_historicos("TULTITLAN")

        with st.spinner('Procesando bases locales...'):
            # --- B. CARGA NORMAL ---
            base_c = cargar_base_sugerido(file_sug_cuauti)
            base_t = cargar_base_sugerido(file_sug_tulti)
            inv_c = limpiar_inventario(file_inv_cuauti, "Cuauti")
            inv_t = limpiar_inventario(file_inv_tulti, "Tulti")
            
            trans_c = procesar_transito(file_trans_cuauti) if file_trans_cuauti else pd.DataFrame(columns=["N° PARTE", "TRANSITO"])
            trans_t = procesar_transito(file_trans_tulti) if file_trans_tulti else pd.DataFrame(columns=["N° PARTE", "TRANSITO"])

            trasp_c = procesar_traspasos(file_sit_cuauti, "TRASUCTU") if file_sit_cuauti else pd.DataFrame(columns=["N° PARTE", "CANTIDAD_TRASPASO"])
            trasp_t = procesar_traspasos(file_sit_tulti, "TRASUCCU") if file_sit_tulti else pd.DataFrame(columns=["N° PARTE", "CANTIDAD_TRASPASO"])

            if base_c is not None and base_t is not None:
                
                # === 1. HOJA DIA CUAUTITLAN ===
                final_c = base_c.copy()
                
                # COL H: HITS LOCALES (De ventas Cuautitlan)
                if df_hits_origen_cuauti is not None:
                    hits_local = df_hits_origen_cuauti.copy()
                    hits_local.rename(columns={'NP': 'N° PARTE', 'HITS_CALCULADO': 'HITS'}, inplace=True)
                    if 'HITS' in final_c.columns: del final_c['HITS']
                    final_c = pd.merge(final_c, hits_local, on='N° PARTE', how='left')

                # COL M: HITS FORANEOS (De ventas Tultitlan)
                if df_hits_origen_tulti is not None:
                    hits_foraneo = df_hits_origen_tulti.copy()
                    hits_foraneo.rename(columns={'NP': 'N° PARTE', 'HITS_CALCULADO': 'HITS_FORANEO'}, inplace=True)
                    final_c = pd.merge(final_c, hits_foraneo, on='N° PARTE', how='left')

                # Merges de Inventario, Transito, Traspaso
                final_c = pd.merge(final_c, inv_c[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                final_c.rename(columns={'EXIST': 'EXISTENCIA', 'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'}, inplace=True)
                final_c = pd.merge(final_c, inv_t[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                final_c.rename(columns={'EXIST': 'INVENTARIO TULTITLAN', 'FEC ULT COMP': 'Fec ult Comp TULTI'}, inplace=True)
                final_c = pd.merge(final_c, trans_c, on='N° PARTE', how='left')
                final_c = pd.merge(final_c, trasp_c, on='N° PARTE', how='left')
                final_c.rename(columns={'CANTIDAD_TRASPASO': 'TRASPASO TULTI A CUATI'}, inplace=True)
                
                final_c = completar_y_ordenar(final_c, COLS_CUAUTITLAN_ORDEN)
                export_c = final_c.copy()
                export_c.rename(columns={'HITS_FORANEO': 'HITS'}, inplace=True) # Renombrar Col M para Excel

                # === 2. HOJA DIA TULTITLAN ===
                final_t = base_t.copy()

                # COL H: HITS LOCALES (De ventas Tultitlan)
                if df_hits_origen_tulti is not None:
                    hits_local_t = df_hits_origen_tulti.copy()
                    hits_local_t.rename(columns={'NP': 'N° PARTE', 'HITS_CALCULADO': 'HITS'}, inplace=True)
                    if 'HITS' in final_t.columns: del final_t['HITS']
                    final_t = pd.merge(final_t, hits_local_t, on='N° PARTE', how='left')

                # COL M: HITS FORANEOS (De ventas Cuautitlan)
                if df_hits_origen_cuauti is not None:
                    hits_foraneo_t = df_hits_origen_cuauti.copy()
                    hits_foraneo_t.rename(columns={'NP': 'N° PARTE', 'HITS_CALCULADO': 'HITS_FORANEO'}, inplace=True)
                    final_t = pd.merge(final_t, hits_foraneo_t, on='N° PARTE', how='left')

                final_t = pd.merge(final_t, inv_t[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                final_t.rename(columns={'EXIST': 'EXISTENCIA', 'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'}, inplace=True)
                final_t = pd.merge(final_t, inv_c[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                final_t.rename(columns={'EXIST': 'INVENTARIO CUAUTITLAN', 'FEC ULT COMP': 'Fec ult Comp CUAUTI'}, inplace=True)
                final_t = pd.merge(final_t, trans_t, on='N° PARTE', how='left')
                final_t = pd.merge(final_t, trasp_t, on='N° PARTE', how='left')
                final_t.rename(columns={'CANTIDAD_TRASPASO': 'TRASPASO CUAUT A TULTI'}, inplace=True)
                
                final_t.rename(columns={'N° PARTE': 'N° DE PARTE'}, inplace=True)
                final_t = completar_y_ordenar(final_t, COLS_TULTITLAN_ORDEN)
                export_t = final_t.copy()
                export_t.rename(columns={'HITS_FORANEO': 'HITS'}, inplace=True) # Renombrar Col M para Excel

                # === 3. SUBIDA ===
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    export_c.to_excel(writer, sheet_name='DIA CUAUTITLAN', index=False)
                    export_t.to_excel(writer, sheet_name='DIA TULTITLAN', index=False)
                buffer.seek(0)
                
                fecha_str = datetime.datetime.now().strftime("%d_%m_%Y_%H%M")
                name_file = f"Analisis_Compras_{fecha_str}.xlsx"
                link = subir_excel_a_drive(buffer, name_file)
                
                if link:
                    st.success(f"✅ Archivo Creado: {name_file}")
                    st.markdown(f"### [📂 Ver en Drive]({link})")
                    st.dataframe(final_c.head())
    else:
        st.warning("⚠️ Faltan archivos.")
