import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Analizador de Compras - Grupo Andrade", layout="wide")

st.title("🔧 Herramienta de Análisis de Inventarios y Compras")
st.markdown("### Fase 1: Carga y Relación de Inventarios (Cuautitlán y Tultitlán)")

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
        # "Siempre tendremos valor en la columna J (FEC INGRESO)"
        # Eliminamos filas donde 'FEC INGRESO' sea nulo
        df_clean = df_clean.dropna(subset=["FEC INGRESO"])
        
        # Convertimos N° PARTE a string y quitamos espacios para asegurar cruces exactos
        df_clean["N° PARTE"] = df_clean["N° PARTE"].astype(str).str.strip()
        
        st.success(f"✅ Inventario {nombre_sucursal} cargado: {df_clean.shape[0]} registros encontrados.")
        return df_clean

    except Exception as e:
        st.error(f"Error al procesar el inventario de {nombre_sucursal}: {e}")
        return None

# --- INTERFAZ DE CARGA ---

st.sidebar.header("1. Carga de Archivos")

# 1. Base Inicial (Sugerido)
file_sugerido = st.sidebar.file_uploader("Cargar Base Inicial (N° Parte y Sugerido)", type=["xlsx"])

st.markdown("---")
st.subheader("Paso 1: SUBIDA INVENTARIOS")

col1, col2 = st.columns(2)
with col1:
    file_inv_cuauti = st.file_uploader("Cargar Inventario Cuautitlán", type=["xlsx", "xls"])
with col2:
    file_inv_tulti = st.file_uploader("Cargar Inventario Tultitlán", type=["xlsx", "xls"])

# --- PROCESAMIENTO ---

if st.button("Generar Reporte Fase 1"):
    if file_sugerido and file_inv_cuauti and file_inv_tulti:
        
        # 1. Cargar Base Sugerido
        df_base = pd.read_excel(file_sugerido)
        # Aseguramos que la columna clave sea string y limpia
        # Asumimos que las columnas se llaman "N° PARTE" y "SUGERIDO DIA"
        df_base["N° PARTE"] = df_base["N° PARTE"].astype(str).str.strip()
        
        # 2. Limpiar Inventarios
        df_inv_cuauti = limpiar_inventario(file_inv_cuauti, "Cuautitlán")
        df_inv_tulti = limpiar_inventario(file_inv_tulti, "Tultitlán")
        
        if df_inv_cuauti is not None and df_inv_tulti is not None:
            
            # --- LOGICA PARA HOJA: DIA CUAUTITLAN ---
            # Base: Sugerido
            df_dia_cuauti = df_base.copy()
            
            # Relación 1: Datos locales (Cuautitlán)
            # Traemos EXIST y FEC ULT COMP de Cuauti
            df_dia_cuauti = pd.merge(
                df_dia_cuauti, 
                df_inv_cuauti[['N° PARTE', 'EXIST', 'FEC ULT COMP']], 
                on='N° PARTE', 
                how='left'
            )
            # Renombramos a lo que pide el usuario
            df_dia_cuauti.rename(columns={
                'EXIST': 'EXISTENCIA',
                'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'
            }, inplace=True)
            
            # Relación 2: Datos foráneos (Tultitlán)
            # Traemos EXIST y FEC ULT COMP de Tulti
            df_dia_cuauti = pd.merge(
                df_dia_cuauti,
                df_inv_tulti[['N° PARTE', 'EXIST', 'FEC ULT COMP']],
                on='N° PARTE',
                how='left'
            )
            # Renombramos
            df_dia_cuauti.rename(columns={
                'EXIST': 'INVENTARIO TULTITLAN',
                'FEC ULT COMP': 'Fec ult Comp TULTI'
            }, inplace=True)

            # --- LOGICA PARA HOJA: DIA TULTITLAN ---
            # Base: Sugerido
            df_dia_tulti = df_base.copy()
            
            # Relación 1: Datos locales (Tultitlán)
            df_dia_tulti = pd.merge(
                df_dia_tulti,
                df_inv_tulti[['N° PARTE', 'EXIST', 'FEC ULT COMP']],
                on='N° PARTE',
                how='left'
            )
            df_dia_tulti.rename(columns={
                'EXIST': 'EXISTENCIA',
                'FEC ULT COMP': 'FECHA DE ULTIMA COMPRA'
            }, inplace=True)
            
            # Relación 2: Datos foráneos (Cuautitlán)
            df_dia_tulti = pd.merge(
                df_dia_tulti,
                df_inv_cuauti[['N° PARTE', 'EXIST', 'FEC ULT COMP']],
                on='N° PARTE',
                how='left'
            )
            df_dia_tulti.rename(columns={
                'EXIST': 'INVENTARIO CUAUTITLAN',
                'FEC ULT COMP': 'Fec ult Comp CUAUTI'
            }, inplace=True)

            # --- EXPORTACIÓN ---
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Hoja Cuautitlán
                df_dia_cuauti.to_excel(writer, sheet_name='DIA CUAUTITLAN', index=False)
                # Hoja Tultitlán
                df_dia_tulti.to_excel(writer, sheet_name='DIA TULTITLAN', index=False)
                
                # Opcional: Si quieres ver como quedaron los inventarios limpios, descomenta esto:
                # df_inv_cuauti.to_excel(writer, sheet_name='Base Limpia Cuauti', index=False)
                # df_inv_tulti.to_excel(writer, sheet_name='Base Limpia Tulti', index=False)
                
            buffer.seek(0)
            
            st.success("¡Cruce de bases realizado con éxito!")
            st.download_button(
                label="📥 Descargar Excel Final (Fase 1)",
                data=buffer,
                file_name="Analisis_Compras_Fase1.xlsx",
                mime="application/vnd.ms-excel"
            )
            
            # Vista previa en pantalla
            st.markdown("#### Vista Previa: DIA CUAUTITLAN")
            st.dataframe(df_dia_cuauti.head())

    else:
        st.warning("⚠️ Por favor carga los 3 archivos para continuar.")
