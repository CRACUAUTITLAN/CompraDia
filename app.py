import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Analizador de Compras - Grupo Andrade", layout="wide")

st.title("🔧 Herramienta de Análisis de Inventarios y Compras")
st.markdown("### Fase 1: Carga, Relación y Estructura Completa")

# --- LISTAS DE COLUMNAS (USAMOS NOMBRES UNICOS INTERNAMENTE) ---
# Nota: Usamos HITS_FORANEO temporalmente para evitar errores visuales en la app.
# Al descargar el Excel, se cambiará automáticamente a HITS.

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

# --- FUNCIONES ---

def limpiar_inventario(archivo, nombre_sucursal):
    try:
        # Detectamos extensión para usar el motor correcto
        if archivo.name.endswith('.xls'):
            df = pd.read_excel(archivo, header=None, engine='xlrd')
        else:
            df = pd.read_excel(archivo, header=None, engine='openpyxl')
        
        # Selección por posición fija (Indices)
        col_indices = [0, 1, 2, 4, 8, 9, 10, 11]
        col_names = [
            "N° PARTE", "DESCR", "CLASIF", "PRECIO UNITARIO", 
            "EXIST", "FEC INGRESO", "FEC ULT COMP", "FEC ULT VTA"
        ]
        
        df_clean = df.iloc[:, col_indices].copy()
        df_clean.columns = col_names
        # Eliminamos filas basura que no tengan fecha de ingreso
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
            st.error(f"❌ Error en archivo {archivo.name}: No se encuentra la columna 'N° PARTE'.")
            return None
            
        df["N° PARTE"] = df["N° PARTE"].astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Error al leer el archivo de sugerido: {e}")
        return None

def completar_y_ordenar(df, lista_columnas_deseadas):
    for col in lista_columnas_deseadas:
        if col not in df.columns:
            df[col] = "" 
    return df[lista_columnas_deseadas]

# --- INTERFAZ ---

st.markdown("---")
st.header("Paso 1: Bases Iniciales (Sugeridos)")
col1, col2 = st.columns(2)
with col1:
    st.subheader("Para Cuautitlán")
    file_sugerido_cuauti = st.file_uploader("📂 Sugerido Cuautitlán (.xlsx)", type=["xlsx"], key="sug_cuauti")
with col2:
    st.subheader("Para Tultitlán")
    file_sugerido_tulti = st.file_uploader("📂 Sugerido Tultitlán (.xlsx)", type=["xlsx"], key="sug_tulti")

st.markdown("---")
st.header("Paso 2: Subida de Inventarios")
col3, col4 = st.columns(2)
with col3:
    st.subheader("Inventario Cuautitlán")
    file_inv_cuauti = st.file_uploader("📦 Inventario Cuautitlán", type=["xlsx", "xls"], key="inv_cuauti")
with col4:
    st.subheader("Inventario Tultitlán")
    file_inv_tulti = st.file_uploader("📦 Inventario Tultitlán", type=["xlsx", "xls"], key="inv_tulti")

# --- PROCESAMIENTO ---

if st.button("Generar Reporte Final"):
    if file_sugerido_cuauti and file_sugerido_tulti and file_inv_cuauti and file_inv_tulti:
        with st.spinner('Procesando y estructurando reporte...'):
            
            # 1. Cargar
            df_base_cuauti = cargar_base_sugerido(file_sugerido_cuauti)
            df_base_tulti = cargar_base_sugerido(file_sugerido_tulti)
            
            if df_base_cuauti is not None and df_base_tulti is not None:
                
                # 2. Limpiar Inventarios
                df_inv_cuauti_clean = limpiar_inventario(file_inv_cuauti, "Cuautitlán")
                df_inv_tulti_clean = limpiar_inventario(file_inv_tulti, "Tultitlán")
                
                if df_inv_cuauti_clean is not None and df_inv_tulti_clean is not None:
                    
                    # --- PROCESO CUAUTITLAN ---
                    df_final_cuauti = df_base_cuauti.copy()
                    
                    # Merge Local
                    df_final_cuauti = pd.merge(df_final_cuauti, df_inv_cuauti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                    df_final_cuauti.rename(columns={'EXIST': 'EXISTENCIA', 'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'}, inplace=True)
                    
                    # Merge Foráneo
                    df_final_cuauti = pd.merge(df_final_cuauti, df_inv_tulti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                    df_final_cuauti.rename(columns={'EXIST': 'INVENTARIO TULTITLAN', 'FEC ULT COMP': 'Fec ult Comp TULTI'}, inplace=True)
                    
                    # Completar y Ordenar (Mantiene HITS_FORANEO)
                    df_final_cuauti = completar_y_ordenar(df_final_cuauti, COLS_CUAUTITLAN_ORDEN)


                    # --- PROCESO TULTITLAN ---
                    df_final_tulti = df_base_tulti.copy()
                    
                    # Merge Local
                    df_final_tulti = pd.merge(df_final_tulti, df_inv_tulti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                    df_final_tulti.rename(columns={'EXIST': 'EXISTENCIA', 'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'}, inplace=True)
                    
                    # Merge Foráneo
                    df_final_tulti = pd.merge(df_final_tulti, df_inv_cuauti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                    df_final_tulti.rename(columns={'EXIST': 'INVENTARIO CUAUTITLAN', 'FEC ULT COMP': 'Fec ult Comp CUAUTI'}, inplace=True)
                    
                    # Ajuste nombre columna
                    df_final_tulti.rename(columns={'N° PARTE': 'N° DE PARTE'}, inplace=True)
                    
                    # Completar y Ordenar (Mantiene HITS_FORANEO)
                    df_final_tulti = completar_y_ordenar(df_final_tulti, COLS_TULTITLAN_ORDEN)

                    # --- EXPORTAR ---
                    # Aquí hacemos el truco: Creamos copias para el Excel donde SÍ duplicamos el nombre HITS
                    df_export_cuauti = df_final_cuauti.copy()
                    df_export_cuauti.rename(columns={'HITS_FORANEO': 'HITS'}, inplace=True)
                    
                    df_export_tulti = df_final_tulti.copy()
                    df_export_tulti.rename(columns={'HITS_FORANEO': 'HITS'}, inplace=True)

                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        df_export_cuauti.to_excel(writer, sheet_name='DIA CUAUTITLAN', index=False)
                        df_export_tulti.to_excel(writer, sheet_name='DIA TULTITLAN', index=False)
                        
                    buffer.seek(0)
                    st.success("✅ Reporte generado correctamente.")
                    st.download_button(label="📥 Descargar Excel Completo", data=buffer, file_name="Analisis_Compras_Fase1_Completo.xlsx", mime="application/vnd.ms-excel")
                    
                    # --- VISTA PREVIA ---
                    # Mostramos la versión CON nombres únicos para que Streamlit no falle
                    st.markdown("#### Vista Previa (Cuautitlán)")
                    st.write("Nota: En pantalla verás 'HITS_FORANEO', pero en tu Excel saldrá como 'HITS'.")
                    st.dataframe(df_final_cuauti.head())

    else:
        st.warning("⚠️ Faltan archivos.")
