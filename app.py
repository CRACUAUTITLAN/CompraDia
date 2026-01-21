import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Analizador de Compras - Grupo Andrade", layout="wide")

st.title("🔧 Herramienta de Análisis de Inventarios y Compras")
st.markdown("### Fase 1: Carga y Relación de Inventarios (Por Sucursal)")

# --- FUNCIONES DE LIMPIEZA ---
def limpiar_inventario(archivo, nombre_sucursal):
    """
    Lee el archivo raw de inventario (.xls o .xlsx), selecciona columnas por posición 
    y limpia basándose en la columna J.
    """
    try:
        # Detectamos extensión para usar el motor correcto
        if archivo.name.endswith('.xls'):
            df = pd.read_excel(archivo, header=None, engine='xlrd')
        else:
            df = pd.read_excel(archivo, header=None, engine='openpyxl')
        
        # Mapeo de columnas basado en índices
        col_indices = [0, 1, 2, 4, 8, 9, 10, 11]
        col_names = [
            "N° PARTE", "DESCR", "CLASIF", "PRECIO UNITARIO", 
            "EXIST", "FEC INGRESO", "FEC ULT COMP", "FEC ULT VTA"
        ]
        
        df_clean = df.iloc[:, col_indices].copy()
        df_clean.columns = col_names
        
        # Limpieza clave
        df_clean = df_clean.dropna(subset=["FEC INGRESO"])
        
        # Convertimos N° PARTE a string
        df_clean["N° PARTE"] = df_clean["N° PARTE"].astype(str).str.strip()
        
        return df_clean

    except Exception as e:
        st.error(f"Error crítico al procesar inventario de {nombre_sucursal}: {e}")
        return None

def cargar_base_sugerido(archivo):
    """Carga la base simple de N° PARTE y SUGERIDO DIA con limpieza de encabezados"""
    try:
        df = pd.read_excel(archivo)
        
        # --- LIMPIEZA DE ENCABEZADOS ---
        # Quitamos espacios en blanco al principio y final de los nombres de columnas
        df.columns = df.columns.str.strip()
        
        # Verificación de columna clave
        if "N° PARTE" not in df.columns:
            st.error(f"❌ Error en archivo {archivo.name}: No se encuentra la columna 'N° PARTE'.")
            st.write("Las columnas encontradas son:", list(df.columns))
            st.warning("Por favor revisa que esté escrita exactamente así: N° PARTE (cuidado con los espacios o el símbolo °)")
            return None
            
        # Aseguramos formato texto
        df["N° PARTE"] = df["N° PARTE"].astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Error al leer el archivo de sugerido: {e}")
        return None

# --- INTERFAZ DE CARGA ---

st.markdown("---")
st.header("Paso 1: Bases Iniciales (Sugeridos)")
st.info("Sube aquí los archivos con las columnas 'N° PARTE' y 'SUGERIDO DIA'.")

col1, col2 = st.columns(2)
with col1:
    st.subheader("Para Cuautitlán")
    file_sugerido_cuauti = st.file_uploader("📂 Sugerido Cuautitlán (.xlsx)", type=["xlsx"], key="sug_cuauti")

with col2:
    st.subheader("Para Tultitlán")
    file_sugerido_tulti = st.file_uploader("📂 Sugerido Tultitlán (.xlsx)", type=["xlsx"], key="sug_tulti")

st.markdown("---")
st.header("Paso 2: Subida de Inventarios (Almacén)")
st.info("Sube aquí los reportes de inventario completos (.xls o .xlsx).")

col3, col4 = st.columns(2)
with col3:
    st.subheader("Inventario Cuautitlán")
    file_inv_cuauti = st.file_uploader("📦 Inventario Cuautitlán", type=["xlsx", "xls"], key="inv_cuauti")

with col4:
    st.subheader("Inventario Tultitlán")
    file_inv_tulti = st.file_uploader("📦 Inventario Tultitlán", type=["xlsx", "xls"], key="inv_tulti")

# --- PROCESAMIENTO ---

if st.button("Generar Reporte Fase 1"):
    if file_sugerido_cuauti and file_sugerido_tulti and file_inv_cuauti and file_inv_tulti:
        
        with st.spinner('Procesando bases de datos...'):
            # 1. Cargar Bases Sugeridos
            df_base_cuauti = cargar_base_sugerido(file_sugerido_cuauti)
            df_base_tulti = cargar_base_sugerido(file_sugerido_tulti)
            
            # Solo seguimos si las bases de sugerido se cargaron bien
            if df_base_cuauti is not None and df_base_tulti is not None:
                
                # 2. Limpiar Inventarios
                df_inv_cuauti_clean = limpiar_inventario(file_inv_cuauti, "Cuautitlán")
                df_inv_tulti_clean = limpiar_inventario(file_inv_tulti, "Tultitlán")
                
                if df_inv_cuauti_clean is not None and df_inv_tulti_clean is not None:
                    
                    # --- CRUCE CUAUTITLAN ---
                    df_final_cuauti = df_base_cuauti.copy()
                    
                    # Local
                    df_final_cuauti = pd.merge(df_final_cuauti, df_inv_cuauti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                    df_final_cuauti.rename(columns={'EXIST': 'EXISTENCIA', 'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'}, inplace=True)
                    
                    # Foráneo
                    df_final_cuauti = pd.merge(df_final_cuauti, df_inv_tulti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                    df_final_cuauti.rename(columns={'EXIST': 'INVENTARIO TULTITLAN', 'FEC ULT COMP': 'Fec ult Comp TULTI'}, inplace=True)

                    # --- CRUCE TULTITLAN ---
                    df_final_tulti = df_base_tulti.copy()
                    
                    # Local
                    df_final_tulti = pd.merge(df_final_tulti, df_inv_tulti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                    df_final_tulti.rename(columns={'EXIST': 'EXISTENCIA', 'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'}, inplace=True)
                    
                    # Foráneo
                    df_final_tulti = pd.merge(df_final_tulti, df_inv_cuauti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], on='N° PARTE', how='left')
                    df_final_tulti.rename(columns={'EXIST': 'INVENTARIO CUAUTITLAN', 'FEC ULT COMP': 'Fec ult Comp CUAUTI'}, inplace=True)

                    # --- EXPORTAR ---
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        df_final_cuauti.to_excel(writer, sheet_name='DIA CUAUTITLAN', index=False)
                        df_final_tulti.to_excel(writer, sheet_name='DIA TULTITLAN', index=False)
                        
                    buffer.seek(0)
                    st.success("✅ ¡Todo listo! Descarga tu archivo.")
                    st.download_button(label="📥 Descargar Excel Final", data=buffer, file_name="Analisis_Compras_Fase1.xlsx", mime="application/vnd.ms-excel")
                    
                    st.markdown("#### Vista Previa: DIA CUAUTITLAN")
                    st.dataframe(df_final_cuauti.head())

    else:
        st.warning("⚠️ Por favor sube los 4 archivos.")
