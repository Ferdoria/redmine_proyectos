import streamlit as st
import pandas as pd
import plotly.express as px

# Forzar modo ancho en toda la app
st.set_page_config(layout="wide")

# Asegúrate de tener instalada la biblioteca openpyxl
# Si no la tienes, ejecuta en tu terminal: pip install openpyxl

# Título de la aplicación
st.title('Dashboard de Migración de Objetos')
st.markdown('---')

# --- Sidebar para carga de archivos ---
st.sidebar.header('📁 Carga de Archivo')
uploaded_file = st.sidebar.file_uploader(
    "Carga el archivo 'Listado de Objetos a migrar_al_11_09.xlsx'", 
    type=['xlsx']
)

# Verificar si se cargó el archivo antes de continuar
if uploaded_file is not None:
    try:
        # Cargar los datos de ambas hojas del mismo archivo
        # Especificar keep_default_na=False para preservar valores 'N/A' como texto
        df_dia_a_dia = pd.read_excel(uploaded_file, sheet_name='Dia a Dia', keep_default_na=False, na_values=[''])
        df_incidentes = pd.read_excel(uploaded_file, sheet_name='Incidentes', keep_default_na=False, na_values=[''])

        # Verificar si los DataFrames están vacíos para evitar FutureWarning
        dataframes_to_concat = []
        if not df_dia_a_dia.empty:
            dataframes_to_concat.append(df_dia_a_dia)
        if not df_incidentes.empty:
            dataframes_to_concat.append(df_incidentes)
        
        # Concatenar los DataFrames solo si hay datos
        if dataframes_to_concat:
            df = pd.concat(dataframes_to_concat, ignore_index=True)
        else:
            st.error("Ambas hojas del archivo están vacías.")
            st.stop()

        # Convertir FECHA XPZ a string si existe
        if 'FECHA XPZ' in df.columns:
            df['FECHA XPZ'] = df['FECHA XPZ'].astype(str)

        # Convertir FECHA XPZ GX8 a string si existe
        if 'FECHA XPZ GX8' in df.columns:
            df['FECHA XPZ GX8'] = df['FECHA XPZ GX8'].astype(str)

        # Convertir FECHA OBJETO a string si existe
        if 'FECHA OBJETO' in df.columns:
            df['FECHA OBJETO'] = df['FECHA OBJETO'].astype(str)

        # Limpiar valores NaN en todas las columnas numéricas o de fecha que puedan causar problemas
        # EXCEPTO la columna RESPONSABLE MIGRACION que se procesará específicamente después
        for column in df.columns:
            if column == 'RESPONSABLE MIGRACION':
                continue  # Saltar esta columna para procesarla específicamente después
            if df[column].dtype == 'object':
                df[column] = df[column].fillna('').astype(str)
            elif df[column].dtype in ['int64', 'float64']:
                df[column] = df[column].fillna(0)
            else:
                df[column] = df[column].astype(str)

        # Renombrar columnas para mayor claridad
        df.rename(columns={
            'RESPONSABLE MIGRACION': 'Responsable_Migracion',
            'COMPILADO?': 'Compilado',
            'TESTEADO': 'Testeado',
            'PROYECTO': 'Proyecto'
        }, inplace=True)
        
        # Limpiar espacios en los nombres de las columnas
        df.columns = df.columns.str.strip()
        
        # --- Limpieza de la columna Responsable_Migracion ---
        # Primero convertir a string manteniendo los valores N/A originales
        df['Responsable_Migracion'] = df['Responsable_Migracion'].astype(str)
        
        # Limpiar espacios pero preservar N/A
        df['Responsable_Migracion'] = df['Responsable_Migracion'].str.strip()
        
        # Solo reemplazar valores que representan ausencia de datos
        # Ahora que preservamos N/A del Excel, solo reemplazamos celdas realmente vacías
        df['Responsable_Migracion'] = df['Responsable_Migracion'].replace(['nan', 'NaN', 'None', '', 'nat', '0'], 'Sin Asignar')
        # Los N/A del Excel ahora se mantienen como 'N/A'
        df['Responsable_Migracion'] = df['Responsable_Migracion'].fillna('Sin Asignar')
        
        # Convertir las columnas relevantes a tipo string para evitar errores
        df['Compilado'] = df['Compilado'].fillna('NO').astype(str)
        df['Testeado'] = df['Testeado'].fillna('NO').astype(str)
        df['Responsable_Migracion'] = df['Responsable_Migracion'].astype(str)
        df['Proyecto'] = df['Proyecto'].fillna('Sin Proyecto').astype(str)
        
        # Asegurar que todas las columnas sean compatibles con Arrow
        # EXCEPTO Responsable_Migracion que ya fue procesada específicamente
        for col in df.columns:
            if col == 'Responsable_Migracion':
                continue  # Saltar esta columna ya que fue procesada específicamente
            if df[col].dtype == 'object':
                df[col] = df[col].fillna('').astype(str)
            elif df[col].dtype in ['float64', 'int64']:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # --- Sidebar para mostrar información del archivo cargado ---
        st.sidebar.success(f"✅ Archivo cargado exitosamente")
        st.sidebar.info(f"📊 Total de registros: {len(df)}")
        
        # --- Sidebar para filtros ---
        st.sidebar.header('🔍 Filtros')
        
        # Seleccionar proyectos
        proyectos = df['Proyecto'].dropna().unique()
        selected_proyectos = st.sidebar.multiselect('Selecciona Proyectos', sorted(proyectos))
        
        # Filtrar el DataFrame
        if selected_proyectos:
            filtered_df = df[df['Proyecto'].isin(selected_proyectos)]
        else:
            filtered_df = df.copy()
        
        # --- Sección de Visualización ---
        st.header('📈 Dashboard de Análisis')

        # KPI's principales
        total_objetos = len(filtered_df)
        objetos_compilados = filtered_df['Compilado'].str.contains('SI', na=False).sum()
        #objetos_testeados = filtered_df['Testeado'].str.contains('SI', na=False).sum()

        # Calcular pendientes a compilar (distinto a SI y N/A)
        objetos_pendientes_compilar = len(filtered_df[
            (~filtered_df['Compilado'].str.contains('SI', na=False)) &
            (~filtered_df['Compilado'].str.contains('N/A', na=False))
        ])

        # Calcular XPZ enviados y pendientes de envío
        if 'XPZ enviado' in filtered_df.columns:
            total_xpz_enviados = filtered_df['XPZ enviado'].str.contains('SI', na=False).sum()
            xpz_pend_envio = objetos_compilados - total_xpz_enviados
        else:
            total_xpz_enviados = 0
            xpz_pend_envio = 0

        # Mostrar KPIs ordenados
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Total de Objetos", total_objetos)
        col2.metric("Objetos Compilados", objetos_compilados)
        col3.metric("⏳ Pendientes a Compilar", objetos_pendientes_compilar)
        #col4.metric("Objetos Testeados", objetos_testeados)
        col4.metric("XPZ Enviados", total_xpz_enviados)
        col5.metric("XPZ Pend. Envio", xpz_pend_envio)
       
        
        st.markdown('---')
        
        # Gráfico por Responsable de Migración
        st.subheader('Asignaciones y Estado por Responsable de Migración')

        # Agrupar por Responsable_Migracion y calcular métricas
        resumen_responsable = filtered_df.groupby('Responsable_Migracion').agg(
            Asignaciones=('Responsable_Migracion', 'count'),
            Compilados=('Compilado', lambda x: (x.str.contains('SI', na=False)).sum()),
            XPZ_Enviados=('XPZ enviado', lambda x: (x.str.contains('SI', na=False)).sum() if 'XPZ enviado' in filtered_df.columns else 0)
        ).reset_index()

        # Calcular XPZ Pendientes de Envío por responsable
        resumen_responsable['XPZ_Pend_Envio'] = resumen_responsable['Compilados'] - resumen_responsable['XPZ_Enviados']

        # Ordenar para que los valores especiales aparezcan primero
        resumen_responsable['orden'] = resumen_responsable['Responsable_Migracion'].apply(
            lambda x: 0 if x == 'Sin Asignar' else (1 if x == 'N/A' else 2)
        )
        resumen_responsable = resumen_responsable.sort_values(['orden', 'Asignaciones'], ascending=[True, False])

        # Mostrar estadísticas adicionales
        total_sin_asignar = resumen_responsable[resumen_responsable['Responsable_Migracion'] == 'Sin Asignar']['Asignaciones'].sum()
        total_na = resumen_responsable[resumen_responsable['Responsable_Migracion'] == 'N/A']['Asignaciones'].sum()
        total_asignados = len(filtered_df) - total_sin_asignar - total_na

        col_info1, col_info2, col_info3 = st.columns(3)
        col_info1.metric("📝 Sin Asignar", total_sin_asignar)
        col_info2.metric("❌ N/A (No Aplica)", total_na)
        col_info3.metric("👤 Con Responsable", total_asignados)

        # Mostrar tabla resumen por responsable
        st.dataframe(resumen_responsable[['Responsable_Migracion', 'Asignaciones', 'Compilados', 'XPZ_Pend_Envio']], use_container_width=True)

        # Crear gráfico de barras apiladas por responsable
        fig_responsable = px.bar(
            resumen_responsable,
            x='Responsable_Migracion',
            y=['Asignaciones', 'Compilados', 'XPZ_Pend_Envio'],
            title='Asignaciones, Compilados y XPZ Pend. Envio por Responsable de Migración',
            labels={'value': 'Cantidad', 'variable': 'Estado', 'Responsable_Migracion': 'Responsable'},
        )
        fig_responsable.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_responsable, use_container_width=True)
        
        # ---        
        st.header('📊 Resumen por Proyecto XPZ Pendientes de envío')
        
        # Crear resumen agrupado por proyecto
        resumen_proyectos = []
        
        for proyecto in filtered_df['Proyecto'].unique():
            # Filtrar datos por proyecto, excluyendo los que tienen Compilado = N/A
            df_proyecto = filtered_df[
            (filtered_df['Proyecto'] == proyecto) & 
            (~filtered_df['Compilado'].str.contains('N/A', na=False))
            ]
            
            # Solo continuar si el proyecto tiene registros válidos (sin N/A)
            if len(df_proyecto) == 0:
                continue
            
            # Calcular métricas por proyecto (ya sin los N/A)
            total_objetos_proyecto = len(df_proyecto)
            objetos_compilados_proyecto = df_proyecto['Compilado'].str.contains('SI', na=False).sum()
            
            # Solo agregar al resumen si hay al menos un objeto compilado (SI)
            if objetos_compilados_proyecto == 0:
                continue
            
            # Calcular XPZ enviados por proyecto
            if 'XPZ enviado' in df_proyecto.columns:
                xpz_enviados_proyecto = df_proyecto['XPZ enviado'].str.contains('SI', na=False).sum()
            else:
                xpz_enviados_proyecto = 0
            
            resumen_proyectos.append({
            'Proyecto': proyecto,
            'Total Objetos': total_objetos_proyecto,
            'Objetos Compilados': objetos_compilados_proyecto,
            'XPZ Enviados': xpz_enviados_proyecto
            })
        
        # Convertir a DataFrame
        df_resumen = pd.DataFrame(resumen_proyectos)
        
        # Filtrar solo proyectos donde XPZ Enviados < Objetos Compilados (pendientes de envío)
        df_resumen_pendientes = df_resumen[df_resumen['XPZ Enviados'] < df_resumen['Objetos Compilados']]
        
        # Ordenar por total de objetos descendente
        df_resumen_pendientes = df_resumen_pendientes.sort_values('Total Objetos', ascending=False).reset_index(drop=True)
        
        # Renumerar la primera columna (índice) para mostrar el orden
        df_resumen_pendientes.index = df_resumen_pendientes.index + 1
        df_resumen_pendientes.index.name = 'N°'
        
        # Mostrar la tabla resumen solo con pendientes
        st.dataframe(df_resumen_pendientes, use_container_width=True)
        # ---
        st.header('📋 Detalle de Objetos')
        st.dataframe(filtered_df)

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo. Asegúrate de que el archivo subido es válido y contiene las hojas 'Dia a Dia' e 'Incidentes'. Error: {e}")

else:
    st.info('Por favor, sube el archivo XLSX para visualizar el dashboard.')