import streamlit as st
import pandas as pd
import plotly.express as px
import re
from io import BytesIO
import subprocess
import sys

# Función para instalar dependencias faltantes
def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

try:
    import openpyxl
except ImportError:
    st.warning("openpyxl no está instalado. Instalando...")
    install("openpyxl")
    import openpyxl
    st.success("openpyxl instalado correctamente!")

# Configurar la página
st.set_page_config(page_title="Dashboard Proyectos", page_icon="📊", layout="wide")
st.title("📊 Dashboard de Proyectos - Core Bancario")

# Funciones de procesamiento de datos
def limpiar_espacios_guion(nombre):
    """Elimina los espacios alrededor del guion en una cadena."""
    if isinstance(nombre, str):
        return re.sub(r"\s*-\s*", "-", nombre.lstrip())
    return nombre

def extraer_codigos(nombre):
    """Extrae Codigo_Proyecto y Codigo_Estabilizacion del nombre."""
    codigo_proyecto = None
    codigo_estabilizacion = None
    if isinstance(nombre, str):
        nombre_sin_espacios_iniciales = nombre.lstrip()
        match_estabilizacion = re.search(r"^(E\d+)", nombre_sin_espacios_iniciales)
        if match_estabilizacion:
            codigo_estabilizacion = match_estabilizacion.group(1)
            nombre_restante = nombre_sin_espacios_iniciales[len(match_estabilizacion.group(1)):].lstrip("- ").lstrip()
        else:
            nombre_restante = nombre_sin_espacios_iniciales

        match_proyecto = re.search(r"([PMANAI]\d+/\d+)", nombre_restante)
        if match_proyecto:
            codigo_proyecto = match_proyecto.group(1)

    return pd.Series([codigo_proyecto, codigo_estabilizacion])

def clasificar(nombre):
    if isinstance(nombre, str):
        nombre_sin_espacios = nombre.lstrip()
        if nombre_sin_espacios.startswith("E"):
            return "Estabilización"
        elif nombre_sin_espacios.startswith("I"):
            return "Incidente"
        elif nombre_sin_espacios.startswith("P"):
            return "Proyecto"
        elif nombre_sin_espacios.startswith("M"):
            return "Mantenimiento"
        elif nombre_sin_espacios.startswith("A"):
            return "Auditoria"
        elif nombre_sin_espacios.startswith("N"):
            return "Normativo"
        else:
            return "Otro"
    return "Otro"

# Carga de datos
uploaded_file = st.file_uploader("Sube el archivo Excel de proyectos", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Leer el archivo Excel
        df = pd.read_excel(uploaded_file, header=3)
        
        # Procesamiento de datos
        df["Nombre"] = df["Nombre"].apply(limpiar_espacios_guion)
        df[['codigo_proyecto', 'codigo_estabilizacion']] = df['Nombre'].apply(extraer_codigos)
        df["tipo"] = df["Nombre"].apply(clasificar)
        
        # Renombrar columnas para consistencia
        df = df.rename(columns={
            "Nombre": "nombre",
            "Estado Actual": "estado_actual",
            "Jefatura": "jefatura",
            "Asignatario predeterminado": "asignatario",
            "Fecha de inicio": "fecha_inicio",
            "Fecha de fin": "fecha_fin",
            "Actualizado por última vez": "actualizado",
            "Etiquetas": "etiquetas",
            "Gestor del proyecto": "gestor",
            "Propietario del proyecto": "propietario",
            "Gerencia/Unidad": "gerencia",
            "Fecha Pasaje a Producción": "fecha_pasaje_prod",
            "Estabilización": "estabilizacion",
            "Autor": "autor"
        })
        
        # Formatear fechas
        df['fecha_inicio'] = pd.to_datetime(df['fecha_inicio'], errors='coerce').dt.strftime('%Y-%m-%d')
        df['fecha_fin'] = pd.to_datetime(df['fecha_fin'], errors='coerce').dt.strftime('%Y-%m-%d')
        df['actualizado'] = pd.to_datetime(df['actualizado'], errors='coerce')
        df['fecha_pasaje_prod'] = pd.to_datetime(df['fecha_pasaje_prod'], errors='coerce')
        
        # Mostrar estadísticas básicas
        st.success(f"Datos cargados correctamente. Total de registros: {len(df)}")
        
        # Inicializar df_filtrado
        df_filtrado = pd.DataFrame()

        # --- Barra Lateral para Filtros ---
        with st.sidebar:
            st.header("🔍 Filtros")

            # Filtro por jefatura
            st.subheader("Jefatura")
            jefaturas_unicas = sorted(df['jefatura'].dropna().unique().tolist())
            default_jefaturas = [j for j in jefaturas_unicas if "Core Bancario" in j]
            selected_jefaturas = st.multiselect("Selecciona jefaturas:", options=jefaturas_unicas, default=default_jefaturas)

        # Aplicar los filtros
        df_filtrado = df[df['jefatura'].isin(selected_jefaturas)]

        ############# Indicadores Clave en el Contenido Principal ############# 
        selected_indicator = st.session_state.get("selected_indicator", None)

        def display_key_indicator(label, value, key):
            button_label = f"**{label}**\n({value})"
            if st.button(button_label, key=key, use_container_width=True):
                st.session_state["selected_indicator"] = label

        col_indicador1_1, col_indicador1_2, col_indicador1_3, col_indicador1_4 = st.columns(4) 
        with col_indicador1_1:
            display_key_indicator("Total Proyectos", len(df_filtrado) if not df_filtrado.empty else 0, "total_proyectos_button")
        with col_indicador1_2:
            display_key_indicator("Finalizados", df_filtrado[df_filtrado['estado_actual'] == 'Finalizado'].shape[0] if not df_filtrado.empty else 0, "finalizado_button")
        with col_indicador1_3:
            display_key_indicator("En Estabilización", df_filtrado[df_filtrado['estado_actual'] == 'Estabilización'].shape[0] if not df_filtrado.empty else 0, "estabilizacion_button")
        with col_indicador1_4:
            display_key_indicator("Para Comité", df_filtrado[df_filtrado['estado_actual'].str.contains('PROD-Para Comité de Pasajes', na=False)].shape[0] if not df_filtrado.empty else 0, "comite_button")

        col_indicador2_1, col_indicador2_2 = st.columns(2)
        with col_indicador2_1:
            display_key_indicator("Análisis Tec (DESA)", df_filtrado[df_filtrado['estado_actual'].str.contains('DESA-Análisis Técnico', na=False)].shape[0] if not df_filtrado.empty else 0, "analisis_button")
        with col_indicador2_2:
            display_key_indicator("En Curso (DESA)", df_filtrado[df_filtrado['estado_actual'].str.contains('DESA-En Curso', na=False)].shape[0] if not df_filtrado.empty else 0, "en_curso_button")

        col_indicador3_1, col_indicador3_2 = st.columns(2)
        with col_indicador3_1:
            display_key_indicator("En QA", df_filtrado[df_filtrado['estado_actual'].str.contains('QA-En Pruebas QA', na=False)].shape[0] if not df_filtrado.empty else 0, "qa_button")
        with col_indicador3_2:
            display_key_indicator("En UAT", df_filtrado[df_filtrado['estado_actual'].str.contains('QA-En Pruebas UAT', na=False)].shape[0] if not df_filtrado.empty else 0, "uat_button")

        col_indicador4_1, col_indicador4_2, col_indicador4_3, col_indicador4_4 = st.columns(4)
        with col_indicador4_1:
            display_key_indicator("PMO-Detenido", df_filtrado[df_filtrado['estado_actual'].str.contains('PMO-Detenido', na=False)].shape[0] if not df_filtrado.empty else 0, "pmo_button")
        with col_indicador4_2:
            display_key_indicator("PMO-No iniciado", df_filtrado[df_filtrado['estado_actual'].str.contains('PMO-No iniciado', na=False)].shape[0] if not df_filtrado.empty else 0, "pmo_no_iniciado_button")
        with col_indicador4_3:
            display_key_indicator("PMO-Relevamiento PMO", df_filtrado[df_filtrado['estado_actual'].str.contains('PMO-Relevamiento PMO', na=False)].shape[0] if not df_filtrado.empty else 0, "pmo_relevamiento_button")
        with col_indicador4_4:
            display_key_indicator("PMO-Pend. Validación técnica", df_filtrado[df_filtrado['estado_actual'].str.contains('PMO-Pend. Validación técnica', na=False)].shape[0] if not df_filtrado.empty else 0, "pmo_pend_validacion_button")

        st.markdown("---")

        col_info1, col_info2, col_info3, col_info4, col_info5 = st.columns(5)
        with col_info1:
            display_key_indicator("Sin Gestor", df_filtrado[df_filtrado['gestor'].isna()].shape[0] if not df_filtrado.empty else 0, "sin_gestor_button")
        with col_info2:
            display_key_indicator("Sin Fecha Inicio", df_filtrado[(df_filtrado['fecha_inicio'].isna()) & (~df_filtrado['estado_actual'].isin(['Estabilización', 'Finalizado', 'PMO-Detenido', 'PMO-No iniciado']))].shape[0] if not df_filtrado.empty else 0, "sin_fecha_inicio_button")
        with col_info3:
            display_key_indicator("En Prod y Sin Fecha Pasaje", df_filtrado[(df_filtrado['fecha_pasaje_prod'].isna()) & (df_filtrado['estado_actual'].isin(['Finalizado', 'Estabilización']))].shape[0] if not df_filtrado.empty else 0, "sin_fecha_pasaje_prod_button")      
        with col_info4:
            display_key_indicator("Sin Fecha Fin", df_filtrado[(df_filtrado['fecha_fin'].isna()) & (~df_filtrado['estado_actual'].isin(['Estabilización', 'Finalizado', 'PMO-Detenido', 'PMO-No iniciado']))].shape[0] if not df_filtrado.empty else 0, "sin_fecha_fin_estado_button")
        with col_info5:
            display_key_indicator("Sin Asignatario", df_filtrado[df_filtrado['asignatario'].isna()].shape[0] if not df_filtrado.empty else 0, "sin_asignatario_button")

        ############## Contenedor Principal para Detalles #############
        st.subheader("Detalles de Indicadores Clave")
        if st.session_state.get("selected_indicator") == "Total Proyectos":
            st.subheader("Detalles de Total Proyectos")
            st.dataframe(df_filtrado, use_container_width=True)
        elif st.session_state.get("selected_indicator") == "En Estabilización":
            st.subheader("Detalles de En Estabilización")
            st.dataframe(df_filtrado[df_filtrado['estado_actual'] == 'Estabilización'] if not df_filtrado.empty else pd.DataFrame(), use_container_width=True)
        elif st.session_state.get("selected_indicator") == "Análisis Tec (DESA)":
            st.subheader("Detalles de Análisis Tec (DESA)")
            st.dataframe(df_filtrado[df_filtrado['estado_actual'].str.contains('DESA-Análisis Técnico', na=False)] if not df_filtrado.empty else pd.DataFrame(), use_container_width=True)
        elif st.session_state.get("selected_indicator") == "En Curso (DESA)":
            st.subheader("Detalles de En Curso (DESA)")
            st.dataframe(df_filtrado[df_filtrado['estado_actual'].str.contains('DESA-En Curso', na=False)] if not df_filtrado.empty else pd.DataFrame(), use_container_width=True)
        elif st.session_state.get("selected_indicator") == "En QA":
            st.subheader("Detalles de En QA")
            st.dataframe(df_filtrado[df_filtrado['estado_actual'].str.contains('QA-En Pruebas QA', na=False)] if not df_filtrado.empty else pd.DataFrame(), use_container_width=True)
        elif st.session_state.get("selected_indicator") == "En UAT":
            st.subheader("Detalles de En UAT")
            st.dataframe(df_filtrado[df_filtrado['estado_actual'].str.contains('QA-En Pruebas UAT', na=False)] if not df_filtrado.empty else pd.DataFrame(), use_container_width=True)
        elif st.session_state.get("selected_indicator") == "Para Comité":
            st.subheader("Detalles de Para Comité")
            st.dataframe(df_filtrado[df_filtrado['estado_actual'].str.contains('PROD-Para Comité de Pasajes', na=False)] if not df_filtrado.empty else pd.DataFrame(), use_container_width=True)
        elif st.session_state.get("selected_indicator") == "Finalizados":
            st.subheader("Detalles de Finalizados")
            st.dataframe(df_filtrado[df_filtrado['estado_actual'] == 'Finalizado'] if not df_filtrado.empty else pd.DataFrame(), use_container_width=True)  
        elif st.session_state.get("selected_indicator") == "PMO-Detenido":
            st.subheader("Detalles de PMO-Detenido")
            st.dataframe(df_filtrado[df_filtrado['estado_actual'].str.contains('PMO-Detenido', na=False)] if not df_filtrado.empty else pd.DataFrame(), use_container_width=True)
        elif st.session_state.get("selected_indicator") == "PMO-No iniciado":   
            st.subheader("Detalles de PMO-No iniciado")
            st.dataframe(df_filtrado[df_filtrado['estado_actual'].str.contains('PMO-No iniciado', na=False)] if not df_filtrado.empty else pd.DataFrame(), use_container_width=True)
        elif st.session_state.get("selected_indicator") == "PMO-Relevamiento PMO":
            st.subheader("Detalles de PMO-Relevamiento PMO")
            st.dataframe(df_filtrado[df_filtrado['estado_actual'].str.contains('PMO-Relevamiento PMO', na=False)] if not df_filtrado.empty else pd.DataFrame(), use_container_width=True)
        elif st.session_state.get("selected_indicator") == "PMO-Pend. Validación técnica":
            st.subheader("Detalles de PMO-Pend. Validación técnica")
            st.dataframe(df_filtrado[df_filtrado['estado_actual'].str.contains('PMO-Pend. Validación técnica', na=False)] if not df_filtrado.empty else pd.DataFrame(), use_container_width=True)
        elif st.session_state.get("selected_indicator") == "Sin Gestor":
            st.subheader("Detalles de Proyectos Sin Gestor")
            st.dataframe(df_filtrado[df_filtrado['gestor'].isna()] if not df_filtrado.empty else pd.DataFrame(), use_container_width=True)  
        elif st.session_state.get("selected_indicator") == "Sin Fecha Inicio":
            st.subheader("Detalles de Proyectos Sin Fecha Inicio")
            st.dataframe(df_filtrado[(df_filtrado['fecha_inicio'].isna()) & (~df_filtrado['estado_actual'].isin(['Estabilización', 'Finalizado', 'PMO-Detenido', 'PMO-No iniciado']))] if not df_filtrado.empty else pd.DataFrame(), use_container_width=True)      
        elif st.session_state.get("selected_indicator") == "En Prod y Sin Fecha Pasaje":    
            st.subheader("Detalles de Proyectos en Prod y Sin Fecha Pasaje")
            st.dataframe(df_filtrado[(df_filtrado['fecha_pasaje_prod'].isna()) & (df_filtrado['estado_actual'].isin(['Finalizado', 'Estabilización']))] if not df_filtrado.empty else pd.DataFrame(), use_container_width=True) 
        elif st.session_state.get("selected_indicator") == "Sin Fecha Fin":
            st.subheader("Detalles de Proyectos Sin Fecha Fin")
            st.dataframe(df_filtrado[(df_filtrado['fecha_fin'].isna()) & (~df_filtrado['estado_actual'].isin(['Estabilización', 'Finalizado', 'PMO-Detenido', 'PMO-No iniciado']))] if not df_filtrado.empty else pd.DataFrame(), use_container_width=True) 
        elif st.session_state.get("selected_indicator") == "Sin Asignatario":
            st.subheader("Detalles de Proyectos Sin Asignatario")
            st.dataframe(df_filtrado[df_filtrado['asignatario'].isna()] if not df_filtrado.empty else pd.DataFrame(), use_container_width=True)
        elif st.session_state.get("selected_indicator"):
            st.info(f"Selecciona un indicador para ver sus detalles.")
        else:
            st.info("Selecciona un indicador para ver sus detalles.")

        ############# Gráficos de Distribución #############
        st.markdown("### 📊 Distribuciones")
        if not df_filtrado.empty:
            st.markdown("#### 📌 Proyectos por Tipo")
            tipos = df_filtrado['tipo'].value_counts().reset_index()
            tipos.columns = ['Tipo', 'Cantidad']
            total_proyectos = tipos['Cantidad'].sum()
            tipos['Porcentaje'] = (tipos['Cantidad'] / total_proyectos) * 100

            fig_tipos = px.bar(tipos, y='Tipo', x='Porcentaje',
                                labels={'Porcentaje': '% del Total', 'Tipo': 'Tipo de Proyecto'},
                                color='Cantidad',
                                color_continuous_scale=px.colors.sequential.Viridis,
                                orientation='h',
                                text=tipos['Cantidad'])

            fig_tipos.update_traces(textposition='outside')
            fig_tipos.update_layout(xaxis_ticksuffix='%')
            st.plotly_chart(fig_tipos, use_container_width=True, key="proyectos_por_tipo")
        else:
            st.info("No hay datos disponibles para mostrar los gráficos de distribución.")

        # Tabs
        tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
            "Por Estado", "Por Asignatario", "Por Jefatura", "Por Etiquetas",
            "Por Gestor", "Proyectos con Estabilizaciones", "Proyectos Pre-Migración-NBT", "Implementados"
        ])

        with tab1:
            if not df_filtrado.empty:
                estado_count = df_filtrado['estado_actual'].value_counts().reset_index()
                estado_count.columns = ['Estado', 'Cantidad']
                fig_estado = px.bar(estado_count, x='Estado', y='Cantidad',
                                             title="Proyectos por Estado",
                                             labels={'Cantidad': 'Número de Proyectos', 'Estado': 'Estado Actual'},
                                             color='Cantidad',
                                             color_continuous_scale=px.colors.sequential.Plasma)
                st.plotly_chart(fig_estado, use_container_width=True, key="proyectos_por_estado_tab")
            else:
                st.info("No hay datos disponibles para mostrar proyectos por estado.")

        with tab2:
            st.markdown("#### 👤 Distribución por Asignatario")
            if not df_filtrado.empty:
                asignatarios = df_filtrado['asignatario'].value_counts(dropna=False).reset_index()
                asignatarios.columns = ['Asignatario', 'Cantidad']
                asignatarios['Asignatario'] = asignatarios['Asignatario'].fillna('Sin Asignar')

                fig_asignatario_tab = px.bar(asignatarios, y='Asignatario', x='Cantidad',
                                                    labels={'Cantidad': 'Número de Proyectos', 'Asignatario': 'Asignatario'},
                                                    color='Cantidad',
                                                    color_continuous_scale=px.colors.sequential.Viridis,
                                                    orientation='h')
                st.plotly_chart(fig_asignatario_tab, use_container_width=True, key="distribucion_asignatario_tab")
            else:
                st.info("No hay datos disponibles para mostrar la distribución por asignatario.")

        with tab3:
            if not df_filtrado.empty:
                jefaturas = df_filtrado['jefatura'].value_counts().reset_index()
                jefaturas.columns = ['Jefatura', 'Cantidad']
                fig_jefatura = px.bar(jefaturas, x='Jefatura', y='Cantidad',
                                             title="Proyectos por Jefatura",
                                             labels={'Cantidad': 'Número de Proyectos', 'Jefatura': 'Jefatura'},
                                             color='Cantidad',
                                             color_continuous_scale=px.colors.sequential.Viridis)
                st.plotly_chart(fig_jefatura, use_container_width=True, key="proyectos_por_jefatura_tab")
            else:
                st.info("No hay datos disponibles para mostrar proyectos por jefatura.")

        with tab4:
            if not df_filtrado.empty:
                etiquetas = df_filtrado['etiquetas'].dropna().str.split(", ").explode().value_counts().reset_index()
                etiquetas.columns = ['Etiqueta', 'Cantidad']
                fig_etiquetas = px.bar(etiquetas, x='Etiqueta', y='Cantidad',
                                              title="Proyectos por Etiqueta",
                                              labels={'Cantidad': 'Número de Proyectos', 'Etiqueta': 'Etiqueta'},
                                              color='Cantidad',
                                              color_continuous_scale=px.colors.sequential.Viridis)
                st.plotly_chart(fig_etiquetas, use_container_width=True, key="proyectos_por_etiqueta_tab")
            else:
                st.info("No hay datos disponibles para mostrar proyectos por etiqueta.")

        with tab5:
            st.markdown("#### 👤 Distribución por Gestor")
            if not df_filtrado.empty:
                gestores = df_filtrado['gestor'].value_counts(dropna=False).reset_index()
                gestores.columns = ['Gestor', 'Cantidad']
                gestores['Gestor'] = gestores['Gestor'].fillna('Sin asignar')

                fig_gestores = px.bar(gestores, y='Gestor', x='Cantidad',
                                             title="Distribución por Gestor",
                                             labels={'Cantidad': 'Número de Proyectos', 'Gestor': 'Gestor'},
                                             color='Cantidad',
                                             color_continuous_scale=px.colors.sequential.Viridis,
                                             orientation='h')
                st.plotly_chart(fig_gestores, use_container_width=True, key="distribucion_gestor_tab")
            else:
                st.info("No hay datos disponibles para mostrar la distribución por gestor.")
                
        with tab6:
            if not df_filtrado.empty:
                df_estabilizaciones = df_filtrado[df_filtrado['codigo_estabilizacion'].astype(str).str.startswith('E', na=False)].copy()
                df_estabilizaciones['asignatario'] = df_estabilizaciones['asignatario'].fillna('Sin Asignar')
                if not df_estabilizaciones.empty:
                    asignatarios_conteo = df_estabilizaciones['asignatario'].value_counts().reset_index()
                    asignatarios_conteo.columns = ['Asignatario', 'Cantidad de Estabilizaciones']
                    asignatarios_conteo = asignatarios_conteo.sort_values(by='Cantidad de Estabilizaciones', ascending=False).head(10)
                    st.subheader("Top 10 Asignatarios con Mayor Cantidad de Estabilizaciones Asignadas")
                    fig_asignatarios = px.bar(asignatarios_conteo, x='Asignatario', y='Cantidad de Estabilizaciones',
                                              title="Top 10 Asignatarios por Cantidad de Estabilizaciones",
                                              labels={'Cantidad de Estabilizaciones': 'Número de Estabilizaciones', 'Asignatario': 'Asignatario'},
                                              color='Cantidad de Estabilizaciones',
                                              color_continuous_scale=px.colors.sequential.Plasma)
                    st.plotly_chart(fig_asignatarios, use_container_width=True, key="top_10_asignatarios_estabs")
                else:
                    st.info("No se encontraron registros de estabilizaciones con asignatario.")

                st.markdown("---")

                codigos_proyecto_base = df_filtrado['codigo_proyecto'].unique()

                if len(codigos_proyecto_base) > 0:
                    estabilizaciones_por_proyecto = {}
                    nombres_proyectos = {}

                    for codigo_base in codigos_proyecto_base:
                        nombre_proyecto = df_filtrado[df_filtrado['codigo_proyecto'] == codigo_base]['nombre'].iloc[0]
                        nombres_proyectos[codigo_base] = nombre_proyecto

                        estabilizaciones = df_filtrado[df_filtrado['codigo_proyecto'] == codigo_base]
                        cantidad_estabilizaciones = estabilizaciones['codigo_estabilizacion'].astype(str).str.startswith('E', na=False).sum()
                        estabilizaciones_por_proyecto[codigo_base] = cantidad_estabilizaciones

                    resumen_estabilizaciones = pd.DataFrame(list(estabilizaciones_por_proyecto.items()), columns=['codigo_proyecto', 'cantidad_estabilizaciones'])
                    resumen_estabilizaciones['nombre'] = resumen_estabilizaciones['codigo_proyecto'].map(nombres_proyectos)
                    resumen_estabilizaciones_con_estabs = resumen_estabilizaciones[resumen_estabilizaciones['cantidad_estabilizaciones'] > 0].sort_values(by='cantidad_estabilizaciones', ascending=False)

                    st.subheader("Detalle de Estabilizaciones por Proyecto Base")
                    st.dataframe(resumen_estabilizaciones_con_estabs[['nombre', 'cantidad_estabilizaciones']], use_container_width=True)

                else:
                    st.info("No se encontraron códigos de proyecto para el detalle.")

            else:
                st.info("No hay datos disponibles para mostrar proyectos con estabilizaciones.")
                
        with tab7:        
            if not df_filtrado.empty:
                proyectos_pre_migracion = df_filtrado[df_filtrado['etiquetas'].str.contains('Pre-Migración-NBT', na=False, regex=False)].copy()
                if not proyectos_pre_migracion.empty:
                    st.subheader("Listado de Proyectos Pre-Migración-NBT")

                    st.subheader("Filtrar Proyectos Pre-Migración-NBT por Estado")
                    estados_unicos_pre_migracion = proyectos_pre_migracion['estado_actual'].unique()
                    estado_seleccionado = st.selectbox("Selecciona un Estado", ["Todos"] + list(estados_unicos_pre_migracion), key="selector_estado_pre_migracion")

                    if estado_seleccionado == "Todos":
                        st.dataframe(proyectos_pre_migracion, use_container_width=True)
                    else:
                        proyectos_filtrados_estado = proyectos_pre_migracion[proyectos_pre_migracion['estado_actual'] == estado_seleccionado]
                        st.dataframe(proyectos_filtrados_estado, use_container_width=True)

                    estado_actual_counts = proyectos_pre_migracion['estado_actual'].value_counts().reset_index()
                    estado_actual_counts.columns = ['Estado Actual', 'Cantidad']
                    fig_pre_migracion_estados = px.bar(estado_actual_counts, x='Estado Actual', y='Cantidad',
                                                                     title="Total por Estado Actual (Proyectos Pre-Migración-NBT)",
                                                                     labels={'Cantidad': 'Número de Proyectos', 'Estado Actual': 'Estado'},
                                                                     color='Cantidad',
                                                                     color_continuous_scale=px.colors.sequential.Viridis)
                    st.plotly_chart(fig_pre_migracion_estados, use_container_width=True, key="estados_pre_migracion_tab")
                else:
                    st.info("No se encontraron proyectos con la etiqueta 'Pre-Migración-NBT'.")
            else:
                st.info("No hay datos disponibles para mostrar proyectos de pre-migración NBT.")

        with tab8:        
            if not df_filtrado.empty:  
                st.markdown("#### 📅 Total Implementado por Mes")
                df_implementado = df_filtrado[df_filtrado['estado_actual'].isin(['Estabilización', 'Finalizado'])].copy()
                if not df_implementado.empty:
                    df_implementado['mes_implementacion'] = df_implementado['fecha_pasaje_prod'].dt.to_period('M')
                    implementados_por_mes = df_implementado['mes_implementacion'].value_counts().sort_index().reset_index()
                    implementados_por_mes.columns = ['mes_period', 'Cantidad']
                    implementados_por_mes['Mes'] = implementados_por_mes['mes_period'].astype(str)
                    implementados_por_mes['Año'] = implementados_por_mes['mes_period'].dt.year.astype(str)

                    anos_unicos = sorted(implementados_por_mes['Año'].unique(), reverse=True)
                    ano_seleccionado = st.selectbox("Selecciona el año:", anos_unicos, key="filtro_ano_implementado")

                    implementados_filtrado_ano = implementados_por_mes[implementados_por_mes['Año'] == ano_seleccionado]

                    fig_implementados_mes = px.bar(implementados_filtrado_ano, x='Mes', y='Cantidad',
                                                    title=f'Proyectos Implementados por Mes ({ano_seleccionado})',
                                                    labels={'Cantidad': 'Número de Proyectos', 'Mes': 'Mes'},
                                                    color='Cantidad',
                                                    color_continuous_scale=px.colors.sequential.Viridis)
                    st.plotly_chart(fig_implementados_mes, use_container_width=True, key="implementados_por_mes")
                else:
                    st.info("No hay proyectos finalizados o en estabilización para mostrar el gráfico por mes.")  

                st.markdown("#### ⚖️ Implementados por Tipo vs Pendientes")
                df_implementados = df_filtrado[df_filtrado['estado_actual'].isin(['Estabilización', 'Finalizado'])].copy()
                tipos_implementados = df_implementados['tipo'].value_counts().reset_index()
                tipos_implementados.columns = ['Tipo', 'Implementados']

                total_por_tipo = df_filtrado['tipo'].value_counts().reset_index()
                total_por_tipo.columns = ['Tipo', 'Total']

                merged_df = pd.merge(total_por_tipo, tipos_implementados, on='Tipo', how='left').fillna(0)
                merged_df['Pendientes'] = merged_df['Total'] - merged_df['Implementados']

                df_plot = merged_df.melt(id_vars=['Tipo'], value_vars=['Implementados', 'Pendientes'], var_name='Estado', value_name='Cantidad')

                fig_implementado_tipo = px.bar(df_plot, x='Tipo', y='Cantidad', color='Estado',
                                             labels={'Cantidad': 'Número de Proyectos', 'Tipo': 'Tipo de Proyecto', 'Estado': 'Estado'},
                                             color_discrete_sequence=px.colors.qualitative.Vivid)
                st.plotly_chart(fig_implementado_tipo, use_container_width=True, key="implementado_por_tipo")
            else:
                st.info("No hay datos disponibles para mostrar proyectos Implementados.")

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
else:
    st.info("Por favor, sube un archivo Excel para comenzar.")