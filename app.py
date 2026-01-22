import streamlit as st
import pandas as pd
import io
import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# Configuración de la página
st.set_page_config(page_title="Analizador de Compras - Grupo Andrade", layout="wide")
st.title("🔧 Herramienta de Análisis de Inventarios y Compras (Auto-Drive)")

# --- CONFIGURACIÓN GOOGLE DRIVE ---
try:
    # Intentamos leer los secretos
    if "gcp_service_account" in st.secrets and "general" in st.secrets:
        gcp_creds = dict(st.secrets["gcp_service_account"])
        PARENT_FOLDER_ID = st.secrets["general"]["drive_folder_id"]
        
        creds = service_account.Credentials.from_service_account_info(
            gcp_creds, scopes=['https://www.googleapis.com/auth/drive']
        )
        drive_service = build('drive', 'v3', credentials=creds)
        st.success("✅ Conexión con Google Drive configurada correctamente.")
    else:
        st.error("❌ Faltan los secretos 'gcp_service_account' o 'general' en la configuración.")
        st.stop()
except Exception as e:
    st.error(f"⚠️ Error crítico configurando Google Drive: {e}")
    st.stop()

# --- FUNCIONES DRIVE ---

def buscar_o_crear_carpeta(nombre_carpeta, parent_id):
    """Busca una carpeta dentro de otra. Si no existe, la crea."""
    try:
        query = f"mimeType='application/vnd.google-apps.folder' and name='{nombre_carpeta}' and '{parent_id}' in parents and trashed=false"
        results = drive_service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get('files', [])

        if files:
            return files[0]['id']
        else:
            metadata = {
                'name': nombre_carpeta,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [parent_id]
            }
            folder = drive_service.files().create(body=metadata, fields='id').execute()
            return folder.get('id')
    except Exception as e:
        st.error(f"Error al buscar/crear carpeta '{nombre_carpeta}': {e}")
        return None

def subir_excel_a_drive(buffer, nombre_archivo):
    try:
        fecha_hoy = datetime.datetime.now()
        anio = str(fecha_hoy.year)
        
        meses_es = {
            1: "01_Enero", 2: "02_Febrero", 3: "03_Marzo", 4: "04_Abril",
            5: "05_Mayo", 6: "06_Junio", 7: "07_Julio", 8: "08_Agosto",
            9: "09_Septiembre", 10: "10_Octubre", 11: "11_Noviembre", 12: "12_Diciembre"
        }
        mes_carpeta = meses_es[fecha_hoy.month]

        # 1. Carpeta AÑO
        id_anio = buscar_o_crear_carpeta(anio, PARENT_FOLDER_ID)
        if not id_anio: return None # Si falla, abortamos
        
        # 2. Carpeta MES
        id_mes = buscar_o_crear_carpeta(mes_carpeta, id_anio)
        if not id_mes: return None

        # 3. Subir Archivo
        media = MediaIoBaseUpload(buffer, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        file_metadata = {
            'name': nombre_archivo,
            'parents': [id_mes]
        }
        
        archivo_nuevo = drive_service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
        return archivo_nuevo.get('webViewLink')

    except Exception as e:
        st.error(f"Error subiendo a Drive: {e}")
        return None

# --- CONSTANTES Y FUNCIONES DE PANDAS ---

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
        st.error(f"Error crítico al procesar inventario de {nombre_sucursal}: {e}")
        return None

def cargar_base_sugerido(archivo):
    try:
        df = pd.read_excel(archivo)
        df.columns = df.columns.str.strip()
        if "N° PARTE" not in df.columns:
            st.error(f"❌ Error en archivo {archivo.name}: No se encuentra 'N° PARTE'.")
            return None
        df["N° PARTE"] = df["N° PARTE"].astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Error al leer sugerido: {e}")
        return None

def completar_y_ordenar(df, lista_columnas_deseadas):
    for col in lista_columnas_deseadas:
        if col not in df.columns:
            df[col] = "" 
    return df[lista_columnas_deseadas]

# --- INTERFAZ ---

st.info("📂 Los archivos generados se guardarán automáticamente en Google Drive.")
st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    st.subheader("Para Cuautitlán")
    file_sugerido_cuauti = st.file_uploader("📂 Sugerido Cuautitlán (.xlsx)", type=["xlsx"], key="sug_cuauti")
with col2:
    st.subheader("Para Tultitlán")
    file_sugerido_tulti = st.file_uploader("📂 Sugerido Tultitlán (.xlsx)", type=["xlsx"], key="sug_tulti")

st.markdown("---")
col3, col4 = st.columns(2)
with col3:
    st.subheader("Inventario Cuautitlán")
    file_inv_cuauti = st.file_uploader("📦 Inventario Cuautitlán", type=["xlsx", "xls"], key="inv_cuauti")
with col4:
    st.subheader("Inventario Tultitlán")
    file_inv_tulti = st.file_uploader("📦 Inventario Tultitlán", type=["xlsx", "xls"], key="inv_tulti")

if st.button("Procesar y Guardar en Drive"):
    if file_sugerido_cuauti and file_sugerido_tulti and file_inv_cuauti and file_inv_tulti:
        with st.spinner('Analizando datos y conectando con Google Drive...'):
            
            # 1. CARGA Y LIMPIEZA
            df_base_cuauti = cargar_base_sugerido(file_sugerido_cuauti)
            df_base_tulti = cargar_base_sugerido(file_sugerido_tulti)
            
            if df_base_cuauti is not None and df_base_tulti is not None:
                df_inv_cuauti_clean = limpiar_inventario(file_inv_cuauti, "Cuautitlán")
                df_inv_tulti_clean = limpiar_inventario(file_inv_tulti, "Tultitlán")
                
                if df_inv_cuauti_clean is not None and df_inv_tulti_clean is not None:
                    
                    # 2. LOGICA DE NEGOCIO
                    # --- CUAUTITLAN ---
                    df_final_cuauti = df_base_cuauti.copy()
                    df_final_cuauti = pd.merge(df_final_cuauti, df_inv_cuauti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                    df_final_cuauti.rename(columns={'EXIST': 'EXISTENCIA', 'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'}, inplace=True)
                    df_final_cuauti = pd.merge(df_final_cuauti, df_inv_tulti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                    df_final_cuauti.rename(columns={'EXIST': 'INVENTARIO TULTITLAN', 'FEC ULT COMP': 'Fec ult Comp TULTI'}, inplace=True)
                    df_final_cuauti = completar_y_ordenar(df_final_cuauti, COLS_CUAUTITLAN_ORDEN)
                    df_export_cuauti = df_final_cuauti.copy()
                    df_export_cuauti.rename(columns={'HITS_FORANEO': 'HITS'}, inplace=True)

                    # --- TULTITLAN ---
                    df_final_tulti = df_base_tulti.copy()
                    df_final_tulti = pd.merge(df_final_tulti, df_inv_tulti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                    df_final_tulti.rename(columns={'EXIST': 'EXISTENCIA', 'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'}, inplace=True)
                    df_final_tulti = pd.merge(df_final_tulti, df_inv_cuauti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                    df_final_tulti.rename(columns={'EXIST': 'INVENTARIO CUAUTITLAN', 'FEC ULT COMP': 'Fec ult Comp CUAUTI'}, inplace=True)
                    df_final_tulti.rename(columns={'N° PARTE': 'N° DE PARTE'}, inplace=True)
                    df_final_tulti = completar_y_ordenar(df_final_tulti, COLS_TULTITLAN_ORDEN)
                    df_export_tulti = df_final_tulti.copy()
                    df_export_tulti.rename(columns={'HITS_FORANEO': 'HITS'}, inplace=True)

                    # 3. GENERAR EXCEL EN MEMORIA
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        df_export_cuauti.to_excel(writer, sheet_name='DIA CUAUTITLAN', index=False)
                        df_export_tulti.to_excel(writer, sheet_name='DIA TULTITLAN', index=False)
                    buffer.seek(0)
                    
                    # 4. SUBIR A DRIVE
                    fecha_hoy_str = datetime.datetime.now().strftime("%d_%m_%Y")
                    nombre_archivo_final = f"Analisis_Compras_{fecha_hoy_str}.xlsx"
                    
                    link_drive = subir_excel_a_drive(buffer, nombre_archivo_final)
                    
                    if link_drive:
                        st.success(f"✅ ¡Éxito! Archivo guardado correctamente en la carpeta del mes.")
                        st.markdown(f"### [📂 Abrir archivo en Google Drive]({link_drive})")
                        st.balloons()
                    else:
                        st.error("❌ El archivo se procesó, pero hubo un error al subirlo a Drive (revisa los mensajes de arriba).")
    else:
        st.warning("⚠️ Faltan archivos.")
