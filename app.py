import streamlit as st
import pandas as pd
import io
import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# Configuración de la página
st.set_page_config(page_title="Analizador de Compras - Grupo Andrade", layout="wide")
st.title("🔧 Herramienta de Análisis de Inventarios y Compras (Fase 3 - Ajuste Tránsito)")

# --- CONFIGURACIÓN GOOGLE DRIVE ---
try:
    if "gcp_service_account" in st.secrets and "general" in st.secrets:
        gcp_creds = dict(st.secrets["gcp_service_account"])
        PARENT_FOLDER_ID = st.secrets["general"]["drive_folder_id"]
        
        creds = service_account.Credentials.from_service_account_info(
            gcp_creds, scopes=['https://www.googleapis.com/auth/drive']
        )
        drive_service = build('drive', 'v3', credentials=creds)
        st.success("✅ Robot de Google Drive conectado (Unidad Compartida).")
    else:
        st.error("❌ Faltan secretos. Revisa la configuración .streamlit/secrets.toml")
        st.stop()
except Exception as e:
    st.error(f"⚠️ Error de conexión con Google: {e}")
    st.stop()

# --- FUNCIONES DRIVE ---

def buscar_o_crear_carpeta(nombre_carpeta, parent_id):
    try:
        query = f"mimeType='application/vnd.google-apps.folder' and name='{nombre_carpeta}' and '{parent_id}' in parents and trashed=false"
        results = drive_service.files().list(
            q=query, fields="files(id, name)", supportsAllDrives=True, includeItemsFromAllDrives=True
        ).execute()
        files = results.get('files', [])

        if files:
            return files[0]['id']
        else:
            metadata = {'name': nombre_carpeta, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
            folder = drive_service.files().create(body=metadata, fields='id', supportsAllDrives=True).execute()
            return folder.get('id')
    except Exception as e:
        st.error(f"❌ Error carpetas Drive: {e}")
        return None

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
        
        archivo_nuevo = drive_service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()
        return archivo_nuevo.get('webViewLink')
    except Exception as e:
        st.error(f"Error subiendo a Drive: {e}")
        return None

# --- FUNCIONES DE PROCESAMIENTO (PANDAS) ---

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
        if archivo.name.endswith('.xls'):
            df = pd.read_excel(archivo, header=None, engine='xlrd')
        else:
            df = pd.read_excel(archivo, header=None, engine='openpyxl')
        
        col_indices = [0, 1, 2, 4, 8, 9, 10, 11]
        col_names = ["N° PARTE", "DESCR", "CLASIF", "PRECIO UNITARIO", "EXIST", "FEC INGRESO", "FEC ULT COMP", "FEC ULT VTA"]
        
        df_clean = df.iloc[:, col_indices].copy()
        df_clean.columns = col_names
        df_clean = df_clean.dropna(subset=["FEC INGRESO"])
        df_clean["N° PARTE"] = df_clean["N° PARTE"].astype(str).str.strip()
        return df_clean
    except Exception as e:
        st.error(f"Error inventario {nombre_sucursal}: {e}")
        return None

def cargar_base_sugerido(archivo):
    try:
        df = pd.read_excel(archivo)
        df.columns = df.columns.str.strip()
        
        if "N° PARTE" not in df.columns:
            st.error(f"❌ Error en {archivo.name}: No se encuentra la columna 'N° PARTE'.")
            return None
        
        df["N° PARTE"] = df["N° PARTE"].astype(str).str.strip()
        
        if "Last 12 Month Demand" in df.columns:
            df["Last 12 Month Demand"] = pd.to_numeric(df["Last 12 Month Demand"], errors='coerce').fillna(0)
            df["CONSUMO MENSUAL"] = df["Last 12 Month Demand"] / 12
        else:
            df["CONSUMO MENSUAL"] = 0
            
        df["2"] = df["CONSUMO MENSUAL"] / 2
            
        return df
    except Exception as e:
        st.error(f"Error al leer sugerido: {e}")
        return None

def procesar_transito(archivo, nombre_sucursal):
    """
    Procesa archivos de Transito. 
    Espera columnas: N° PARTE, TRANSITO
    (Ignora INV. TOTAL por ahora)
    """
    try:
        df = pd.read_excel(archivo)
        df.columns = df.columns.str.strip() 
        
        # Ajuste: Solo buscamos N° PARTE y TRANSITO
        cols_necesarias = ["N° PARTE", "TRANSITO"]
        
        for col in cols_necesarias:
            if col not in df.columns:
                st.warning(f"⚠️ Falta la columna '{col}' en Tránsito {nombre_sucursal}.")
                df[col] = 0
                
        df_resumen = df[cols_necesarias].copy()
        df_resumen["N° PARTE"] = df_resumen["N° PARTE"].astype(str).str.strip()
        
        # Agrupamos por si hay duplicados
        df_agrupado = df_resumen.groupby("N° PARTE", as_index=False)["TRANSITO"].sum()
        
        return df_agrupado
        
    except Exception as e:
        st.error(f"Error procesando tránsito {nombre_sucursal}: {e}")
        return None

def procesar_traspasos(archivo, filtro_nomenclatura, nombre_proceso):
    try:
        if archivo.name.endswith('.xls'):
            df = pd.read_excel(archivo, header=None, engine='xlrd')
        else:
            df = pd.read_excel(archivo, header=None, engine='openpyxl')
            
        df_filtrado = df[df[0].astype(str).str.strip() == filtro_nomenclatura].copy()
        
        if df_filtrado.empty:
            return pd.DataFrame(columns=["N° PARTE", "CANTIDAD_TRASPASO"])
            
        df_resumen = df_filtrado[[2, 4]].copy()
        df_resumen.columns = ["N° PARTE", "CANTIDAD_TRASPASO"]
        
        df_resumen["N° PARTE"] = df_resumen["N° PARTE"].astype(str).str.strip()
        df_resumen["CANTIDAD_TRASPASO"] = pd.to_numeric(df_resumen["CANTIDAD_TRASPASO"], errors='coerce').fillna(0).abs()
        
        df_agrupado = df_resumen.groupby("N° PARTE", as_index=False)["CANTIDAD_TRASPASO"].sum()
        
        return df_agrupado
    except Exception as e:
        st.error(f"Error procesando traspasos {nombre_proceso}: {e}")
        return None

def completar_y_ordenar(df, lista_columnas_deseadas):
    for col in lista_columnas_deseadas:
        if col not in df.columns:
            # Rellenamos con 0 las columnas faltantes (incluyendo INV. TOTAL)
            df[col] = 0 
    
    # Reordenamos
    df = df[lista_columnas_deseadas]
    
    # LIMPIEZA FINAL: Todo lo que sea NaN o vacío se vuelve 0
    df = df.fillna(0)
    
    return df

# --- INTERFAZ GRAFICA ---

st.info("📂 Los archivos se subirán a Google Drive (Unidad Compartida).")

# PASO 1: SUGERIDOS
st.header("Paso 1: Bases Iniciales (Sugeridos)")
c1, c2 = st.columns(2)
file_sugerido_cuauti = c1.file_uploader("📂 Sugerido Cuautitlán", type=["xlsx"], key="sug_cuauti")
file_sugerido_tulti = c2.file_uploader("📂 Sugerido Tultitlán", type=["xlsx"], key="sug_tulti")

st.markdown("---")

# PASO 2: TRANSITO (Solo Columna TRANSITO)
st.header("Paso 2: Reportes de Tránsito")
st.markdown("Sube los archivos que contienen: N° PARTE y TRANSITO")
c3, c4 = st.columns(2)
file_transito_cuauti = c3.file_uploader("🚢 Tránsito Cuautitlán", type=["xlsx"], key="trans_cuauti")
file_transito_tulti = c4.file_uploader("🚢 Tránsito Tultitlán", type=["xlsx"], key="trans_tulti")

st.markdown("---")

# PASO 3: TRASPASOS
st.header("Paso 3: Reportes de Situación (Traspasos)")
c5, c6 = st.columns(2)
file_situacion_para_cuauti = c5.file_uploader("🚛 Situación Cuautitlán (Busca TRASUCTU)", type=["xlsx", "xls"], key="sit_cuauti")
file_situacion_para_tulti = c6.file_uploader("🚛 Situación Tultitlán (Busca TRASUCCU)", type=["xlsx", "xls"], key="sit_tulti")

st.markdown("---")

# PASO 4: INVENTARIOS
st.header("Paso 4: Inventarios (Almacén)")
c7, c8 = st.columns(2)
file_inv_cuauti = c7.file_uploader("📦 Inventario Cuautitlán", type=["xlsx", "xls"], key="inv_cuauti")
file_inv_tulti = c8.file_uploader("📦 Inventario Tultitlán", type=["xlsx", "xls"], key="inv_tulti")

if st.button("🚀 Procesar Todo y Subir a Drive"):
    if file_sugerido_cuauti and file_sugerido_tulti and file_inv_cuauti and file_inv_tulti:
        with st.spinner('Procesando bases...'):
            
            # A. CARGAR
            df_base_cuauti = cargar_base_sugerido(file_sugerido_cuauti)
            df_base_tulti = cargar_base_sugerido(file_sugerido_tulti)
            
            # B. INVENTARIOS
            df_inv_cuauti_clean = limpiar_inventario(file_inv_cuauti, "Cuautitlán")
            df_inv_tulti_clean = limpiar_inventario(file_inv_tulti, "Tultitlán")
            
            # C. TRANSITOS (Paso 2 - Modificado para solo leer TRANSITO)
            if file_transito_cuauti:
                df_trans_cuauti = procesar_transito(file_transito_cuauti, "Cuautitlán")
            else:
                df_trans_cuauti = pd.DataFrame(columns=["N° PARTE", "TRANSITO"])

            if file_transito_tulti:
                df_trans_tulti = procesar_transito(file_transito_tulti, "Tultitlán")
            else:
                df_trans_tulti = pd.DataFrame(columns=["N° PARTE", "TRANSITO"])

            # D. TRASPASOS (Paso 3)
            if file_situacion_para_cuauti:
                df_traspasos_a_cuauti = procesar_traspasos(file_situacion_para_cuauti, "TRASUCTU", "Sit. Cuautitlán")
            else:
                df_traspasos_a_cuauti = pd.DataFrame(columns=["N° PARTE", "CANTIDAD_TRASPASO"])

            if file_situacion_para_tulti:
                df_traspasos_a_tulti = procesar_traspasos(file_situacion_para_tulti, "TRASUCCU", "Sit. Tultitlán")
            else:
                df_traspasos_a_tulti = pd.DataFrame(columns=["N° PARTE", "CANTIDAD_TRASPASO"])

            # --- ARMADO FINAL ---
            
            if (df_base_cuauti is not None and df_base_tulti is not None and 
                df_inv_cuauti_clean is not None and df_inv_tulti_clean is not None):
                
                # === 1. HOJA DIA CUAUTITLAN ===
                df_final_cuauti = df_base_cuauti.copy()
                
                # Merge Inventarios
                df_final_cuauti = pd.merge(df_final_cuauti, df_inv_cuauti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                df_final_cuauti.rename(columns={'EXIST': 'EXISTENCIA', 'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'}, inplace=True)
                df_final_cuauti = pd.merge(df_final_cuauti, df_inv_tulti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                df_final_cuauti.rename(columns={'EXIST': 'INVENTARIO TULTITLAN', 'FEC ULT COMP': 'Fec ult Comp TULTI'}, inplace=True)
                
                # Merge Transito (Solo TRANSITO)
                df_final_cuauti = pd.merge(df_final_cuauti, df_trans_cuauti, on='N° PARTE', how='left')
                
                # Merge Traspasos
                df_final_cuauti = pd.merge(df_final_cuauti, df_traspasos_a_cuauti, on='N° PARTE', how='left')
                df_final_cuauti.rename(columns={'CANTIDAD_TRASPASO': 'TRASPASO TULTI A CUATI'}, inplace=True)
                
                # Completar y Llenar Ceros (Aqui INV. TOTAL se llena con 0)
                df_final_cuauti = completar_y_ordenar(df_final_cuauti, COLS_CUAUTITLAN_ORDEN)
                df_export_cuauti = df_final_cuauti.copy()
                df_export_cuauti.rename(columns={'HITS_FORANEO': 'HITS'}, inplace=True)

                # === 2. HOJA DIA TULTITLAN ===
                df_final_tulti = df_base_tulti.copy()
                
                # Merge Inventarios
                df_final_tulti = pd.merge(df_final_tulti, df_inv_tulti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                df_final_tulti.rename(columns={'EXIST': 'EXISTENCIA', 'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'}, inplace=True)
                df_final_tulti = pd.merge(df_final_tulti, df_inv_cuauti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                df_final_tulti.rename(columns={'EXIST': 'INVENTARIO CUAUTITLAN', 'FEC ULT COMP': 'Fec ult Comp CUAUTI'}, inplace=True)
                
                # Merge Transito
                df_final_tulti = pd.merge(df_final_tulti, df_trans_tulti, on='N° PARTE', how='left')
                
                # Merge Traspasos
                df_final_tulti = pd.merge(df_final_tulti, df_traspasos_a_tulti, on='N° PARTE', how='left')
                df_final_tulti.rename(columns={'CANTIDAD_TRASPASO': 'TRASPASO CUAUT A TULTI'}, inplace=True)
                
                # Ajustes finales
                df_final_tulti.rename(columns={'N° PARTE': 'N° DE PARTE'}, inplace=True)
                df_final_tulti = completar_y_ordenar(df_final_tulti, COLS_TULTITLAN_ORDEN)
                df_export_tulti = df_final_tulti.copy()
                df_export_tulti.rename(columns={'HITS_FORANEO': 'HITS'}, inplace=True)

                # === 3. SUBIDA A DRIVE ===
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_export_cuauti.to_excel(writer, sheet_name='DIA CUAUTITLAN', index=False)
                    df_export_tulti.to_excel(writer, sheet_name='DIA TULTITLAN', index=False)
                buffer.seek(0)
                
                fecha_hoy_str = datetime.datetime.now().strftime("%d_%m_%Y")
                hora_str = datetime.datetime.now().strftime("%H%M") 
                nombre_archivo_final = f"Analisis_Compras_{fecha_hoy_str}_{hora_str}.xlsx"
                
                link_drive = subir_excel_a_drive(buffer, nombre_archivo_final)
                
                if link_drive:
                    st.success(f"✅ ¡Proceso Completo! Archivo: {nombre_archivo_final}")
                    st.markdown(f"### [📂 Ver archivo en Google Drive]({link_drive})")
                    st.balloons()
                    
                    st.markdown("#### Vista Previa: DIA CUAUTITLAN")
                    st.dataframe(df_final_cuauti.head())
                else:
                    st.error("❌ Falló la subida a Drive.")
    else:
        st.warning("⚠️ Debes cargar al menos los Sugeridos y los Inventarios.")
