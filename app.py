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
    Lee el archivo raw de inventario, selecciona columnas específicas por posición 
    y limpia basándose en la columna J (Fecha de Ingreso).
    """
    try:
        # Leemos sin encabezados porque el formato es crudo
        df = pd.read_excel(archivo, header=None)
        
        # Mapeo de columnas basado en índices (A=0, B=1, etc.)
        # A: N° PARTE (0), B: DESCR (1), C: CLASIF (2), E: PRECIO UNIT(4)
        # I: EXIST (8), J: FEC INGRESO (9), K: FEC ULT COMP (10), L: FEC ULT VTA (11)
        col_indices = [0, 1, 2, 4, 8, 9, 10, 11]
        col_names = [
            "N° PARTE", "DESCR", "CLASIF", "PRECIO UNITARIO", 
            "EXIST", "FEC INGRESO", "FEC ULT COMP", "FEC ULT VTA"
        ]
        
        # Seleccionamos solo las columnas de interés
        df_clean = df.iloc[:, col_indices].copy()
        df_clean.columns = col_names
        
        # --- LIMPIEZA CLAVE ---
        # Eliminamos filas donde 'FEC INGRESO' sea nulo
        df_clean = df_clean.dropna(subset=["FEC INGRESO"])
        
        # Convertimos N° PARTE a string y quitamos espacios
        df_clean["N° PARTE"] = df_clean["N° PARTE"].astype(str).str.strip()
        
        return df_clean

    except Exception as e:
        st.error(f"Error al procesar el inventario de {nombre_sucursal}: {e}")
        return None

def cargar_base_sugerido(archivo):
    """Carga la base simple de N° PARTE y SUGERIDO DIA"""
    try:
        df = pd.read_excel(archivo)
        # Aseguramos que N° PARTE sea texto para el cruce
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
    file_sugerido_cuauti = st.file_uploader("📂 Base Sugerido Cuautitlán", type=["xlsx"], key="sug_cuauti")

with col2:
    st.subheader("Para Tultitlán")
    file_sugerido_tulti = st.file_uploader("📂 Base Sugerido Tultitlán", type=["xlsx"], key="sug_tulti")

st.markdown("---")
st.header("Paso 2: Subida de Inventarios (Almacén)")
st.info("Sube aquí los reportes de inventario completos.")

col3, col4 = st.columns(2)
with col3:
    st.subheader("Inventario Cuautitlán")
    file_inv_cuauti = st.file_uploader("📦 Inventario Cuautitlán (Raw)", type=["xlsx", "xls"], key="inv_cuauti")

with col4:
    st.subheader("Inventario Tultitlán")
    file_inv_tulti = st.file_uploader("📦 Inventario Tultitlán (Raw)", type=["xlsx", "xls"], key="inv_tulti")

# --- PROCESAMIENTO ---

if st.button("Generar Reporte Fase 1"):
    # Verificamos que los 4 archivos estén cargados
    if file_sugerido_cuauti and file_sugerido_tulti and file_inv_cuauti and file_inv_tulti:
        
        with st.spinner('Procesando bases de datos...'):
            # 1. Cargar Bases Sugeridos
            df_base_cuauti = cargar_base_sugerido(file_sugerido_cuauti)
            df_base_tulti = cargar_base_sugerido(file_sugerido_tulti)
            
            # 2. Limpiar Inventarios
            df_inv_cuauti_clean = limpiar_inventario(file_inv_cuauti, "Cuautitlán")
            df_inv_tulti_clean = limpiar_inventario(file_inv_tulti, "Tultitlán")
            
            if (df_base_cuauti is not None and df_base_tulti is not None and 
                df_inv_cuauti_clean is not None and df_inv_tulti_clean is not None):
                
                # ---------------------------------------------------------
                # LOGICA PARA HOJA: DIA CUAUTITLAN
                # Usamos la base sugerido de Cuauti + Inv Cuauti (Local) + Inv Tulti (Foráneo)
                # ---------------------------------------------------------
                df_final_cuauti = df_base_cuauti.copy()
                
                # A. Datos Locales (Cuauti) -> EXISTENCIA, ULT COMPRA
                df_final_cuauti = pd.merge(
                    df_final_cuauti, 
                    df_inv_cuauti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']], 
                    on='N° PARTE', 
                    how='left'
                )
                df_final_cuauti.rename(columns={
                    'EXIST': 'EXISTENCIA',
                    'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'
                }, inplace=True)
                
                # B. Datos Foráneos (Tulti) -> INV TULTI, ULT COMP TULTI
                df_final_cuauti = pd.merge(
                    df_final_cuauti,
                    df_inv_tulti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']],
                    on='N° PARTE',
                    how='left'
                )
                df_final_cuauti.rename(columns={
                    'EXIST': 'INVENTARIO TULTITLAN',
                    'FEC ULT COMP': 'Fec ult Comp TULTI'
                }, inplace=True)

                # ---------------------------------------------------------
                # LOGICA PARA HOJA: DIA TULTITLAN
                # Usamos la base sugerido de Tulti + Inv Tulti (Local) + Inv Cuauti (Foráneo)
                # ---------------------------------------------------------
                df_final_tulti = df_base_tulti.copy()
                
                # A. Datos Locales (Tulti) -> EXISTENCIA, ULT COMPRA
                df_final_tulti = pd.merge(
                    df_final_tulti,
                    df_inv_tulti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']],
                    on='N° PARTE',
                    how='left'
                )
                df_final_tulti.rename(columns={
                    'EXIST': 'EXISTENCIA',
                    'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'
                }, inplace=True)
                
                # B. Datos Foráneos (Cuauti) -> INV CUAUTI, ULT COMP CUAUTI
                df_final_tulti = pd.merge(
                    df_final_tulti,
                    df_inv_cuauti_clean[['N° PARTE', 'EXIST', 'FEC ULT COMP']],
                    on='N° PARTE',
                    how='left'
                )
                df_final_tulti.rename(columns={
                    'EXIST': 'INVENTARIO CUAUTITLAN',
                    'FEC ULT COMP': 'Fec ult Comp CUAUTI'
                }, inplace=True)

                # ---------------------------------------------------------
                # EXPORTACIÓN
                # ---------------------------------------------------------
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_final_cuauti.to_excel(writer, sheet_name='DIA CUAUTITLAN', index=False)
                    df_final_tulti.to_excel(writer, sheet_name='DIA TULTITLAN', index=False)
                    
                buffer.seek(0)
                
                st.success("✅ ¡Archivos procesados correctamente!")
                
                st.download_button(
                    label="📥 Descargar Excel Final",
                    data=buffer,
                    file_name="Analisis_Compras_Fase1.xlsx",
                    mime="application/vnd.ms-excel"
                )
                
                # Vistas previas rápidas
                st.markdown("#### Vista Previa: Resultado Cuautitlán")
                st.dataframe(df_final_cuauti.head())
                
                st.markdown("#### Vista Previa: Resultado Tultitlán")
                st.dataframe(df_final_tulti.head())

    else:
        st.warning("⚠️ Faltan archivos. Asegúrate de cargar los 2 sugeridos y los 2 inventarios.")
