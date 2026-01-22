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
    if "gcp_service_account" in st.secrets and "general" in st.secrets:
        gcp_creds = dict(st.secrets["gcp_service_account"])
        PARENT_FOLDER_ID = st.secrets["general"]["drive_folder_id"]
        
        creds = service_account.Credentials.from_service_account_info(
            gcp_creds, scopes=['https://www.googleapis.com/auth/drive']
        )
        drive_service = build('drive', 'v3', credentials=creds)
        st.success("✅ Robot de Google Drive conectado.")
    else:
        st.error("❌ Faltan secretos. Revisa la configuración .streamlit/secrets.toml")
        st.stop()
except Exception as e:
    st.error(f"⚠️ Error de conexión con Google: {e}")
    st.stop()

# --- FUNCIONES DRIVE ---

def buscar_o_crear_carpeta(nombre_carpeta, parent_id):
    try:
        # Buscamos si ya existe la carpeta
        query = f"mimeType='application/vnd.google-apps.folder' and name='{nombre_carpeta}' and '{parent_id}' in parents and trashed=false"
        results = drive_service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get('files', [])

        if files:
            return files[0]['id']
        else:
            # Si no existe, la creamos
            metadata = {
                'name': nombre_carpeta,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [parent_id]
            }
            folder = drive_service.files().create(body=metadata, fields='id').execute()
            return folder.get('id')
    except Exception as e:
        # Si falla aquí, es 99% seguro que es por PERMISOS (Error 404)
        st.error(f"❌ Error al acceder a la carpeta. ¿Compartiste la carpeta '{parent_id}' con el correo del robot?")
        st.error(f"Detalle técnico: {e}")
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
        if not id_anio: return None 
        
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

# --- FUNCIONES PANDAS ---

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
        # Detectar extensión para usar xlrd si es necesario
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
        st.error(f"Error procesando inventario de {nombre_sucursal}. ¿Está instalado xlrd? Error: {e}")
        return None

def cargar_base_sugerido(archivo):
    try:
        df = pd.read_excel(archivo)
        df.columns = df.columns.str.strip()
        if "N° PARTE" not in df.columns:
            st.error(f"❌ Error en {archivo.name}: No se encuentra la columna 'N° PARTE'.")
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

st.info("📂 Los archivos se subirán automáticamente a Google Drive.")
st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    st.subheader("Para Cuautitlán")
    file_sugerido_cuauti = st.file_uploader("📂 Sugerido Cuautitlán", type=["xlsx"], key="sug_cuauti")
with col2:
    st.subheader("Para Tultitlán")
    file_sugerido_tulti = st.file_uploader("📂 Sugerido Tultitlán", type=["xlsx"], key="sug_tulti")

st.markdown("---")
col3, col4 = st.columns(2)
with col3:
    st.subheader("Inventario Cuautitlán")
    file_inv_cuauti = st.file_uploader("📦 Inventario Cuautitlán", type=["xlsx", "xls"], key="inv_cuauti")
with col4:
    st.subheader("Inventario Tultitlán")
    file_inv_tulti = st.file_uploader("📦 Inventario Tultitlán", type=["xlsx", "xls"], key="inv_tulti")

if st.button("Procesar y Subir a Drive"):
    if file_sugerido_cuauti and file_sugerido_tulti and file_inv_cuauti and file_inv_tulti:
        with st.spinner('Procesando datos y conectando con Google Drive...'):
            
            # 1. CARGA
            df_base_cuauti = cargar_base_sugerido(file_sugerido_cuauti)
            df_base_tulti = cargar_base_sugerido(file_sugerido_tulti)
            
            if df_base_cuauti is not None and df_base_tulti is not None:
                df_inv_cuauti_clean = limpiar_inventario(file_inv_cuauti, "Cuautitlán")
                df_inv_tulti_clean = limpiar_inventario(file_inv_tulti, "Tultitlán")
                
                if df_inv_cuauti_clean is not None and df_inv_tulti_clean is not None:
                    
                    # 2. LOGICA CUAUTITLAN
                    df_final_cuauti = df_base_cuauti.copy()
                    df_final_cuauti = pd.merge(df_final_cuauti, df_inv_cuauti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                    df_final_cuauti.rename(columns={'EXIST': 'EXISTENCIA', 'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'}, inplace=True)
                    df_final_cuauti = pd.merge(df_final_cuauti, df_inv_tulti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                    df_final_cuauti.rename(columns={'EXIST': 'INVENTARIO TULTITLAN', 'FEC ULT COMP': 'Fec ult Comp TULTI'}, inplace=True)
                    df_final_cuauti = completar_y_ordenar(df_final_cuauti, COLS_CUAUTITLAN_ORDEN)
                    
                    # PREPARAR EXPORTACION (Renombrando HITS)
                    df_export_cuauti = df_final_cuauti.copy()
                    df_export_cuauti.rename(columns={'HITS_FORANEO': 'HITS'}, inplace=True)

                    # 3. LOGICA TULTITLAN
                    df_final_tulti = df_base_tulti.copy()
                    df_final_tulti = pd.merge(df_final_tulti, df_inv_tulti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                    df_final_tulti.rename(columns={'EXIST': 'EXISTENCIA', 'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'}, inplace=True)
                    df_final_tulti = pd.merge(df_final_tulti, df_inv_cuauti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                    df_final_tulti.rename(columns={'EXIST': 'INVENTARIO CUAUTITLAN', 'FEC ULT COMP': 'Fec ult Comp CUAUTI'}, inplace=True)
                    df_final_tulti.rename(columns={'N° PARTE': 'N° DE PARTE'}, inplace=True)
                    df_final_tulti = completar_y_ordenar(df_final_tulti, COLS_TULTITLAN_ORDEN)
                    
                    # PREPARAR EXPORTACION (Renombrando HITS)
                    df_export_tulti = df_final_tulti.copy()
                    df_export_tulti.rename(columns={'HITS_FORANEO': 'HITS'}, inplace=True)

                    # 4. GENERAR EXCEL
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        df_export_cuauti.to_excel(writer, sheet_name='DIA CUAUTITLAN', index=False)
                        df_export_tulti.to_excel(writer, sheet_name='DIA TULTITLAN', index=False)
                    buffer.seek(0)
                    
                    # 5. SUBIR A DRIVE
                    fecha_hoy_str = datetime.datetime.now().strftime("%d_%m_%Y")
                    # Agregamos hora para evitar nombres duplicados si pruebas mucho
                    hora_str = datetime.datetime.now().strftime("%H%M") 
                    nombre_archivo_final = f"Analisis_Compras_{fecha_hoy_str}_{hora_str}.xlsx"
                    
                    link_drive = subir_excel_a_drive(buffer, nombre_archivo_final)
                    
                    if link_drive:
                        st.success(f"✅ ¡Éxito! Archivo guardado: {nombre_archivo_final}")
                        st.markdown(f"### [📂 Ver archivo en Google Drive]({link_drive})")
                        st.balloons()
                        
                        # MOSTRAR VISTA PREVIA (Usamos la versión SIN duplicados para que no falle)
                        st.markdown("#### Vista Previa (Cuautitlán)")
                        st.dataframe(df_final_cuauti.head())
                    else:
                        st.error("❌ Falló la subida a Drive. Revisa los permisos.")
    else:
        st.warning("⚠️ Faltan archivos por cargar.")
