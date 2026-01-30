import streamlit as st
import pandas as pd
import io
import datetime
from dateutil.relativedelta import relativedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# Configuración de la página
st.set_page_config(page_title="CRA INT - Compras Día", layout="wide")
st.title("💎 CRA INTERNATIONAL: COMPRAS DÍA")

# --- CONFIGURACIÓN GOOGLE DRIVE ---
try:
    if "gcp_service_account" in st.secrets and "general" in st.secrets:
        gcp_creds = dict(st.secrets["gcp_service_account"])
        PARENT_FOLDER_ID = st.secrets["general"]["drive_folder_id"] # Salida
        MASTER_SALES_ID = st.secrets["general"].get("master_sales_id") # Lectura
        
        creds = service_account.Credentials.from_service_account_info(
            gcp_creds, scopes=['https://www.googleapis.com/auth/drive']
        )
        drive_service = build('drive', 'v3', credentials=creds)
    else:
        st.error("❌ Faltan secretos. Revisa la configuración.")
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
        return []

    for anio in anios:
        query = f"name contains '{agencia}' and name contains '{anio}' and name contains 'MASTER' and '{MASTER_SALES_ID}' in parents and trashed=false"
        results = drive_service.files().list(
            q=query, fields="files(id, name)", supportsAllDrives=True, includeItemsFromAllDrives=True
        ).execute()
        files = results.get('files', [])
        archivos_encontrados.extend(files)
    return archivos_encontrados

# --- LOGICA BI (HISTÓRICO) ---

def mapear_mes_a_numero(mes_texto):
    if not isinstance(mes_texto, str): return 0
    mes = mes_texto.upper().strip()
    diccionario = {
        'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6,
        'JULIO': 7, 'AGOSTO': 8, 'SEPTIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11, 'DICIEMBRE': 12,
        'ENE': 1, 'FEB': 2, 'MAR': 3, 'ABR': 4, 'MAY': 5, 'JUN': 6,
        'JUL': 7, 'AGO': 8, 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DIC': 12
    }
    return diccionario.get(mes, 0)

def obtener_dataframe_ventas(agencia):
    hoy = datetime.datetime.now()
    periodo_fin = hoy.year * 100 + hoy.month
    fecha_inicio = hoy - relativedelta(years=1)
    periodo_inicio = fecha_inicio.year * 100 + fecha_inicio.month
    
    anios_drive = list(set([fecha_inicio.year, hoy.year]))
    files_metadata = buscar_archivos_ventas(agencia.upper(), anios_drive)
    
    if not files_metadata: return None, periodo_inicio, periodo_fin

    dfs = []
    for file_meta in files_metadata:
        content = descargar_archivo_drive(file_meta['id'])
        if content:
            try:
                engine = 'xlrd' if 'xls' in file_meta['name'] and 'xlsx' not in file_meta['name'] else 'openpyxl'
                df_temp = pd.read_excel(content, engine=engine)
                df_temp.columns = df_temp.columns.str.upper().str.strip()
                rename_map = {}
                for col in df_temp.columns:
                    if 'AÑO' in col or 'ANIO' in col: rename_map[col] = 'AÑO'
                    if 'MES' in col and 'PROMEDIO' not in col: rename_map[col] = 'MES'
                df_temp.rename(columns=rename_map, inplace=True)
                if 'AÑO' in df_temp.columns and 'MES' in df_temp.columns:
                    dfs.append(df_temp)
            except Exception: pass
            
    if not dfs: return None, periodo_inicio, periodo_fin
    
    df_total = pd.concat(dfs, ignore_index=True)
    return df_total, periodo_inicio, periodo_fin

def calcular_bi_historico(agencia, debug_np=None):
    df_total, p_inicio, p_fin = obtener_dataframe_ventas(agencia)
    if df_total is None:
        if debug_np: st.error(f"No hay datos para {agencia}.")
        return None

    df_total['AÑO'] = pd.to_numeric(df_total['AÑO'], errors='coerce').fillna(0).astype(int)
    df_total['MES_NUM'] = df_total['MES'].astype(str).apply(mapear_mes_a_numero)
    df_total['PERIODO'] = (df_total['AÑO'] * 100) + df_total['MES_NUM']
    
    if debug_np:
        st.markdown(f"### 🕵️ MODO DETECTIVE: {debug_np} en {agencia}")
        st.write(f"**Periodo:** {p_inicio} al {p_fin}")

    mask = (df_total['PERIODO'] >= p_inicio) & (df_total['PERIODO'] < p_fin)
    df_filtrado = df_total.loc[mask].copy()

    if df_filtrado.empty: return None

    if 'NP' not in df_filtrado.columns or 'CANTIDAD' not in df_filtrado.columns: return None
        
    df_filtrado['NP'] = df_filtrado['NP'].astype(str).str.strip()
    df_filtrado['CANTIDAD'] = pd.to_numeric(df_filtrado['CANTIDAD'], errors='coerce').fillna(0)
    
    resumen = df_filtrado.groupby('NP').agg(
        total_eventos=('CANTIDAD', 'count'),
        eventos_negativos=('CANTIDAD', lambda x: (x < 0).sum()),
        suma_cantidad=('CANTIDAD', 'sum')
    ).reset_index()
    
    resumen['HITS_CALCULADO'] = resumen['total_eventos'] - (resumen['eventos_negativos'] * 2)
    resumen['HITS_CALCULADO'] = resumen['HITS_CALCULADO'].clip(lower=0) 
    resumen['PROMEDIO_CALCULADO'] = resumen['suma_cantidad'] / 12

    if debug_np:
        row = resumen[resumen['NP'] == str(debug_np).strip()]
        if not row.empty:
            st.success(f"**HITS:** {row.iloc[0]['HITS_CALCULADO']} | **PROMEDIO:** {row.iloc[0]['PROMEDIO_CALCULADO']}")

    return resumen[['NP', 'HITS_CALCULADO', 'PROMEDIO_CALCULADO']]

# --- FUNCIONES PANDAS ---
COLS_CUAUTITLAN_ORDEN = ["N° PARTE", "SUGERIDO DIA", "POR FINCAR", "(Consumo Mensual / 2) - Inv Tran", "EXISTENCIA", "FECHA DE ULTIMA COMPRA", "PROMEDIO CUAUTITLAN", "HITS", "CONSUMO MENSUAL", "2", "INVENTARIO TULTITLAN", "PROMEDIO TULTITLAN", "HITS_FORANEO", "TRASPASO TULTI A CUATI", "NUEVO TRASPASO", "CANTIDAD A TRASPASAR", "Fec ult Comp TULTI", "TRANSITO", "INV. TOTAL", "MESES VENTA ACTUAL", "MESES VENTA SUGERIDO", "Line Value", "Cycle Count", "Status", "Ship Multiple", "Last 12 Month Demand", "Current Month Demand", "Job Quantity", "Full Bin", "Bin Location", "Dealer On Hand", "Stock on Order", "Stock On Back Order", "Reason Code"]
COLS_TULTITLAN_ORDEN = ["N° DE PARTE", "SUGERIDO DIA", "POR FINCAR", "(Consumo Mensual / 2) - Inv Tran", "EXISTENCIA", "FECHA DE ULTIMA COMPRA", "PROMEDIO TULTITLAN", "HITS", "CONSUMO MENSUAL", "2", "INVENTARIO CUAUTITLAN", "PROMEDIO CUAUTITLAN", "HITS_FORANEO", "TRASPASO CUAUT A TULTI", "NUEVO TRASPASO", "CANTIDAD A TRASPASAR", "Fec ult Comp CUAUTI", "TRANSITO", "INV. TOTAL", "MESES VENTA ACTUAL", "MESES VENTA SUGERIDO", "Line Value", "Cycle Count", "Status", "Ship Multiple", "Last 12 Month Demand", "Current Month Demand", "Job Quantity", "Full Bin", "Bin Location", "Dealer On Hand", "Stock on Order", "Stock On Back Order", "Reason Code"]

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

# --- DISEÑO Y FORMULAS DE EXCEL (ELEGANTES) ---
def formatear_excel_final(writer, df, sheet_name):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # 1. ESTILOS CORPORATIVOS CRA
    # Estilo Base: Azul Marino Profundo (Corporativo)
    fmt_header_base = workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        'bg_color': '#10345C', 'font_color': 'white', 'border': 1
    })
    
    # Estilo Local (Sucursal Actual): Azul Acero (Serio)
    fmt_header_local = workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        'bg_color': '#4B8BBE', 'font_color': 'white', 'border': 1
    })
    
    # Estilo Foraneo (Otra Sucursal): Terracota Oscuro (Diferenciación elegante)
    fmt_header_foraneo = workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        'bg_color': '#A64d4d', 'font_color': 'white', 'border': 1
    })
    
    # Estilo Input (Usuario): Beige muy claro
    fmt_header_input = workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        'bg_color': '#F2F2F2', 'font_color': 'black', 'border': 1
    })

    # Estilo Celdas (Todo centrado y bordes finos grises)
    cell_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#D3D3D3'})
    
    # 2. APLICAR FORMATO A ENCABEZADOS SEGÚN LA LÓGICA DE NEGOCIO
    for col_num, value in enumerate(df.columns.values):
        col_name = str(value).upper()
        
        # Determinar estilo por nombre de columna
        style = fmt_header_base # Default
        
        # Lógica de Colores:
        if col_num < 3: # Primeras 3 columnas (Prioridad)
            style = fmt_header_base
        elif "NUEVO TRASPASO" in col_name or "CANTIDAD A TRASPASAR" in col_name:
            style = fmt_header_input
        elif "CUAUTITLAN" in sheet_name:
            # En hoja Cuautitlán: Local es Cuauti (Azul), Foráneo es Tulti (Rojo)
            if "CUAUTI" in col_name or col_name == "HITS" or col_name == "EXISTENCIA" or col_name == "CONSUMO MENSUAL":
                style = fmt_header_local
            elif "TULTI" in col_name or "FORANEO" in col_name:
                style = fmt_header_foraneo
        elif "TULTITLAN" in sheet_name:
            # En hoja Tultitlán: Local es Tulti (Azul), Foráneo es Cuauti (Rojo)
            if "TULTI" in col_name or col_name == "HITS" or col_name == "EXISTENCIA" or col_name == "CONSUMO MENSUAL":
                style = fmt_header_local
            elif "CUAUTI" in col_name or "FORANEO" in col_name:
                style = fmt_header_foraneo

        worksheet.write(0, col_num, value, style)
    
    # Ajuste de anchos
    worksheet.set_column('A:A', 20) 
    worksheet.set_column('B:Z', 14, cell_fmt) 
    
    # 3. INSERTAR FÓRMULAS
    start_row = 1
    for i in range(len(df)):
        row = start_row + i
        excel_row = row + 1 
        
        # Mapeo de columnas por índice (A=0, B=1, C=2...)
        # N° PARTE=0, SUGERIDO=1, POR FINCAR=2, (Cons/2)-Inv=3, EXISTENCIA=4
        # FEC ULT=5, PROM LOC=6, HITS=7, CONS MEN=8, "2"=9
        # INV FOR=10, PROM FOR=11, HITS FOR=12, TRASP FOR=13, NUEVO TRASP=14
        # CANT TRASP=15, FEC FOR=16, TRANSITO=17, INV TOTAL=18
        # MESES ACT=19, MESES SUG=20
        
        # C: POR FINCAR
        # Formula: =IF(((B+S+P)/I)>1.5, MIN(I,B), IF((R+(E/I))>3, 0, B-P))
        f_por_fincar = f'=IFERROR(IF(((B{excel_row}+S{excel_row}+P{excel_row})/I{excel_row})>1.5, MIN(I{excel_row},B{excel_row}), IF((R{excel_row}+(E{excel_row}/I{excel_row}))>3, 0, B{excel_row}-P{excel_row})), 0)'
        worksheet.write_formula(row, 2, f_por_fincar, cell_fmt)
        
        # D: (Consumo/2)-Inv = J-S (Col "2" - Inv Total)
        f_formula_d = f'=J{excel_row}-S{excel_row}'
        worksheet.write_formula(row, 3, f_formula_d, cell_fmt)
        
        # O: Validacion SI/NO
        worksheet.data_validation(row, 14, row, 14, {'validate': 'list', 'source': ['SI', 'NO']})
        
        # S: INV TOTAL = R+E+P (Transito + Existencia + Nuevo Traspaso)
        f_inv_total = f'=R{excel_row}+E{excel_row}+P{excel_row}'
        worksheet.write_formula(row, 18, f_inv_total, cell_fmt)
        
        # T: MESES ACT = S/I
        f_meses_act = f'=IFERROR(S{excel_row}/I{excel_row}, 0)'
        worksheet.write_formula(row, 19, f_meses_act, cell_fmt)
        
        # U: MESES SUG = (C+S+P)/I
        f_meses_sug = f'=IFERROR((C{excel_row}+S{excel_row}+P{excel_row})/I{excel_row}, 0)'
        worksheet.write_formula(row, 20, f_meses_sug, cell_fmt)

# --- INTERFAZ GRAFICA ---

st.info("📂 Los archivos se subirán a Google Drive (Unidad Compartida).")

with st.expander("🕵️ MODO DETECTIVE (Revisar HITS y PROMEDIOS)"):
    col_deb1, col_deb2 = st.columns(2)
    np_investigar = col_deb1.text_input("N° PARTE:", "")
    agencia_inv = col_deb2.selectbox("Agencia:", ["CUAUTITLAN", "TULTITLAN"])
    if st.button("🔍 Investigar"):
        if np_investigar: calcular_bi_historico(agencia_inv, debug_np=np_investigar)

st.markdown("---")

# PASOS
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

if st.button("🚀 PROCESAR Y GENERAR REPORTE"):
    if file_sug_cuauti and file_sug_tulti and file_inv_cuauti and file_inv_tulti:
        
        # A. BI HISTORICO
        df_bi_cuauti = calcular_bi_historico("CUAUTITLAN")
        df_bi_tulti = calcular_bi_historico("TULTITLAN")

        with st.spinner('Procesando bases y generando diseño corporativo...'):
            base_c = cargar_base_sugerido(file_sug_cuauti)
            base_t = cargar_base_sugerido(file_sug_tulti)
            inv_c = limpiar_inventario(file_inv_cuauti, "Cuauti")
            inv_t = limpiar_inventario(file_inv_tulti, "Tulti")
            trans_c = procesar_transito(file_trans_cuauti) if file_trans_cuauti else pd.DataFrame(columns=["N° PARTE", "TRANSITO"])
            trans_t = procesar_transito(file_trans_tulti) if file_trans_tulti else pd.DataFrame(columns=["N° PARTE", "TRANSITO"])
            trasp_c = procesar_traspasos(file_sit_cuauti, "TRASUCTU") if file_sit_cuauti else pd.DataFrame(columns=["N° PARTE", "CANTIDAD_TRASPASO"])
            trasp_t = procesar_traspasos(file_sit_tulti, "TRASUCCU") if file_sit_tulti else pd.DataFrame(columns=["N° PARTE", "CANTIDAD_TRASPASO"])

            if base_c is not None and base_t is not None:
                # === HOJA DIA CUAUTITLAN ===
                final_c = base_c.copy()
                
                if df_bi_cuauti is not None:
                    bi_local = df_bi_cuauti.copy()
                    bi_local.rename(columns={'NP': 'N° PARTE', 'HITS_CALCULADO': 'HITS', 'PROMEDIO_CALCULADO': 'PROMEDIO CUAUTITLAN'}, inplace=True)
                    if 'HITS' in final_c.columns: del final_c['HITS']
                    if 'PROMEDIO CUAUTITLAN' in final_c.columns: del final_c['PROMEDIO CUAUTITLAN']
                    final_c = pd.merge(final_c, bi_local, on='N° PARTE', how='left')

                if df_bi_tulti is not None:
                    bi_foraneo = df_bi_tulti.copy()
                    bi_foraneo.rename(columns={'NP': 'N° PARTE', 'HITS_CALCULADO': 'HITS_FORANEO', 'PROMEDIO_CALCULADO': 'PROMEDIO TULTITLAN'}, inplace=True)
                    if 'PROMEDIO TULTITLAN' in final_c.columns: del final_c['PROMEDIO TULTITLAN']
                    final_c = pd.merge(final_c, bi_foraneo, on='N° PARTE', how='left')

                final_c = pd.merge(final_c, inv_c[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                final_c.rename(columns={'EXIST': 'EXISTENCIA', 'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'}, inplace=True)
                final_c = pd.merge(final_c, inv_t[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                final_c.rename(columns={'EXIST': 'INVENTARIO TULTITLAN', 'FEC ULT COMP': 'Fec ult Comp TULTI'}, inplace=True)
                final_c = pd.merge(final_c, trans_c, on='N° PARTE', how='left')
                final_c = pd.merge(final_c, trasp_c, on='N° PARTE', how='left')
                final_c.rename(columns={'CANTIDAD_TRASPASO': 'TRASPASO TULTI A CUATI'}, inplace=True)
                final_c = completar_y_ordenar(final_c, COLS_CUAUTITLAN_ORDEN)
                export_c = final_c.copy()
                export_c.rename(columns={'HITS_FORANEO': 'HITS'}, inplace=True)

                # === HOJA DIA TULTITLAN ===
                final_t = base_t.copy()

                if df_bi_tulti is not None:
                    bi_local_t = df_bi_tulti.copy()
                    bi_local_t.rename(columns={'NP': 'N° PARTE', 'HITS_CALCULADO': 'HITS', 'PROMEDIO_CALCULADO': 'PROMEDIO TULTITLAN'}, inplace=True)
                    if 'HITS' in final_t.columns: del final_t['HITS']
                    if 'PROMEDIO TULTITLAN' in final_t.columns: del final_t['PROMEDIO TULTITLAN']
                    final_t = pd.merge(final_t, bi_local_t, on='N° PARTE', how='left')

                if df_bi_cuauti is not None:
                    bi_foraneo_t = df_bi_cuauti.copy()
                    bi_foraneo_t.rename(columns={'NP': 'N° PARTE', 'HITS_CALCULADO': 'HITS_FORANEO', 'PROMEDIO_CALCULADO': 'PROMEDIO CUAUTITLAN'}, inplace=True)
                    if 'PROMEDIO CUAUTITLAN' in final_t.columns: del final_t['PROMEDIO CUAUTITLAN']
                    final_t = pd.merge(final_t, bi_foraneo_t, on='N° PARTE', how='left')

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
                export_t.rename(columns={'HITS_FORANEO': 'HITS'}, inplace=True)

                # === SUBIDA CON DISEÑO Y FORMULAS ===
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    # En lugar de to_excel directo, escribimos los datos y luego aplicamos el formato
                    export_c.to_excel(writer, sheet_name='DIA CUAUTITLAN', index=False)
                    formatear_excel_final(writer, export_c, 'DIA CUAUTITLAN')
                    
                    export_t.to_excel(writer, sheet_name='DIA TULTITLAN', index=False)
                    formatear_excel_final(writer, export_t, 'DIA TULTITLAN')
                    
                buffer.seek(0)
                
                fecha_str = datetime.datetime.now().strftime("%d_%m_%Y_%H%M")
                name_file = f"Analisis_Compras_{fecha_str}.xlsx"
                link = subir_excel_a_drive(buffer, name_file)
                
                if link:
                    st.success(f"✅ Archivo Maestro Creado: {name_file}")
                    st.markdown(f"### [📂 Abrir en Google Drive]({link})")
                    st.balloons()
    else:
        st.warning("⚠️ Faltan archivos.")
