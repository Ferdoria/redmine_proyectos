import streamlit as st
import pandas as pd
import plotly.express as px
import re
from io import BytesIO
import subprocess
import sys
#FD
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
st.set_page_config(page_title="Dashboard Proyectos Agosto 2025", page_icon="📊", layout="wide")

st.title("📊 Dashboard de Proyectos (agosto 25) - Core Bancario")

# Texto explicativo sobre los colores (compatibles con modo oscuro)
st.markdown(
    """
    <div style="display:flex;gap:16px;flex-wrap:wrap;">
        <div style="background-color:#2c7873;color:#fff;padding:8px 16px;border-radius:6px;display:inline-block;min-width:220px;">
            <b>Color verde:</b> proyectos en estado <b>Estabilización</b> o <b>Finalizado</b>.
        </div>
        <div style="background:transparent;color:#ffd700;padding:8px 16px;border-radius:6px;display:inline-block;min-width:220px;border:1.5px solid #ffd700;">
            <b>Color amarillo:</b> proyectos cuya etiqueta contiene <b>Post/Agos/25</b>.
        </div>
        <div style="background:transparent;color:#2980b9;padding:8px 16px;border-radius:6px;display:inline-block;min-width:220px;border:1.5px solid #2980b9;">
            <b>Color azul:</b> Despues Freeze<br>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

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
uploaded_file = st.sidebar.file_uploader("Sube el archivo Excel de proyectos", type=["xlsx"])

if uploaded_file is not None:
    # Función para mostrar barra de progreso en % Realizado
    def barra_porcentaje(val):
        try:
            pct = float(val)
        except:
            return val
        pct = max(0, min(100, pct))
        color = '#2c7873' if pct == 100 else '#2980b9'
        return f'<div style="background:#e0e0e0;border-radius:4px;position:relative;height:22px;width:100%;"><div style="background:{color};width:{pct}%;height:100%;border-radius:4px;"></div><span style="position:absolute;left:50%;top:0;transform:translateX(-50%);color:#222;font-weight:bold;">{pct:.0f}%</span></div>'
    # Función para resaltar filas:
    # - Verde suave de fondo si estado_actual es 'Estabilización' o 'Finalizado'
    # - Texto dorado si etiquetas contiene 'Post/Agos/25' (aunque el fondo sea verde)
    def highlight_filas(row):
        estado = str(row.get("estado_actual", "")).strip().lower()
        tiene_post_agos = "post/agos/25" in str(row.get("etiquetas", "")).lower()
        nombre = str(row.get("nombre", ""))
        codigos_azules = [
            "M022/24", "M030/25", "M018/25", "M048/25", "M034/25",
            "M041/25", "M136/24", "M043/25", "M034/24"
        ]
        color = ""
        # Fondo verde si corresponde
        if estado in ["finalizado", "estabilización"]:
            color = "background-color: #2c7873; color: #fff;"
        # Si tiene Post/Agos/25, forzar color de texto dorado
        if tiene_post_agos:
            color = color.replace('color: #fff;', 'color: #ffd700;') if 'color: #fff;' in color else color + ' color: #ffd700;'
        # Si el nombre contiene alguno de los códigos, forzar color azul
        if any(codigo in nombre for codigo in codigos_azules):
            # Si ya hay color de fondo, solo cambia el color de letra
            if 'color:' in color:
                color = re.sub(r'color: #[0-9a-fA-F]{3,6};?', 'color: #2980b9;', color)
            else:
                color += ' color: #2980b9;'
        return [color] * len(row)

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
    # Filtrar solo proyectos cuyas etiquetas contengan '/Agos/25'
    df = df[df['etiquetas'].str.contains('/Agos/25', na=False, regex=False)]

    # Tabs principales: Datos Completos, Datos Agrupados y Agrupados por Estados
    main_tab1, main_tab2, main_tab3, tab_gerencia, tab_asignatario = st.tabs(["Datos Completos", "Gráficos", "Agrupados por Estados", "Agrupados por Gerencia/Unidad", "Por Asignatarios"])
    # Nuevo tab: Agrupados por Asignatario
    with tab_asignatario:
        st.write("Proyectos agrupados por Asignatario:")

        def extraer_asignatario(valor):
            if pd.isna(valor) or str(valor).strip() == '':
                return "(Sin asignatario)"
            return str(valor).strip()

        # Buscar el nombre real de la columna 'asignatario' ignorando mayúsculas, minúsculas y espacios
        col_asignatario = None
        for col in df.columns:
            if col.replace(' ', '').lower() in ["asignatariopredeterminado", "asignatario"]:
                col_asignatario = col
                break
        if col_asignatario is None:
            st.error("No se encontró la columna 'Asignatario' en los datos.")
        else:
            # Filtrar por jefatura que contenga 'core bancario' o 'normativo'
            df_asignatario = df[df['jefatura'].str.lower().str.contains('core bancario|normativo', na=False)].copy()
            df_asignatario['Asignatario'] = df_asignatario[col_asignatario].apply(extraer_asignatario)
            asignatarios = df_asignatario['Asignatario'].dropna().unique()
            asignatarios = sorted(asignatarios)

            # Calcular columnas a mostrar para este tab (evitar columnas inexistentes)
            columnas_ocultas_asignatario = ['proyecto matriz', 'autor', 'codigo_proyecto', 'estabilizacion', 'codigo_estabilizacion', 'Grupo Jefatura']
            columnas_a_mostrar_asignatario = [col for col in df_asignatario.columns if col.lower() not in [c.lower() for c in columnas_ocultas_asignatario]]

            for asign in asignatarios:
                n = df_asignatario[df_asignatario['Asignatario'] == asign].shape[0]
                color = "#51748b"  # azul para agrupación
                st.markdown(
                    f'<div style="background-color:{color};padding:10px 16px;border-radius:6px;margin-bottom:0px;font-weight:bold;font-size:1.1em;">{asign} <span style="float:right">{n}</span></div>',
                    unsafe_allow_html=True
                )
                if n > 0:
                    st.dataframe(
                        df_asignatario[df_asignatario['Asignatario'] == asign][columnas_a_mostrar_asignatario].style.apply(highlight_filas, axis=1),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.info("No hay proyectos para este asignatario.")

    with main_tab1:
        st.subheader("Datos completos Core Bancario/Normativo")
        # --- Bloque resumen ---
        total_proyectos = len(df)
        total_finalizados = df['estado_actual'].astype(str).str.lower().eq('finalizado').sum()
        total_estabilizacion = df['estado_actual'].astype(str).str.lower().eq('estabilización').sum()
        total_implementados = total_finalizados + total_estabilizacion
        total_pres_agos = df['etiquetas'].astype(str).str.lower().str.contains('pres/agos/25').sum()
        total_post_agos = df['etiquetas'].astype(str).str.lower().str.contains('post/agos/25').sum()

        resumen_data = {
            'Total Proyectos': [total_proyectos],           
            'Presentación Agosto/25': [total_pres_agos],
            'Posterior Presentación Agosto/25': [total_post_agos],
            'Implementados': [total_implementados]
        }
        resumen_df = pd.DataFrame(resumen_data)
        #st.markdown('<div style="background:#f5f5f5;padding:8px 16px;border-radius:6px;display:inline-block;font-weight:bold;margin-bottom:10px;">Resumen general de la planilla</div>', unsafe_allow_html=True)
        st.dataframe(resumen_df, use_container_width=True, hide_index=True)
        columnas_ocultas = ['proyecto matriz', 'autor', 'codigo_proyecto', 'estabilizacion', 'codigo_estabilizacion']
        columnas_a_mostrar = [col for col in df.columns if col.lower() not in columnas_ocultas]
        # Filtrar solo jefatura Core Bancario y Normativo
        df_core = df[df['jefatura'].str.lower().str.contains('core bancario|normativo', na=False)].copy()
        codigos_azules = [
            "M022/24", "M030/25", "M018/25", "M048/25", "M034/25",
            "M041/25", "M136/24", "M043/25", "M034/24"
        ]
        # Bloque Antes del Freeze (excluyendo Finalizado y Estabilización)
        estados_excluir = ['finalizado', 'estabilización']
        df_antes = df_core[
            (~df_core['nombre'].astype(str).apply(lambda x: any(c in x for c in codigos_azules))) &
            (~df_core['estado_actual'].astype(str).str.lower().isin(estados_excluir))
        ]
        st.markdown('<div style="background:#2c7873;color:#fff;padding:8px 16px;border-radius:6px;display:inline-block;font-weight:bold;">Antes del Freeze</div>', unsafe_allow_html=True)
        st.success(f"Total de registros: {len(df_antes[columnas_a_mostrar])}")
        st.dataframe(
            df_antes[columnas_a_mostrar].style.apply(highlight_filas, axis=1),
            use_container_width=True,
            height=(35 * len(df_antes[columnas_a_mostrar]) + 40),
            hide_index=True
        )
        # Bloque Después del Freeze (excluyendo Finalizado y Estabilización)
        df_despues = df_core[
            (df_core['nombre'].astype(str).apply(lambda x: any(c in x for c in codigos_azules))) &
            (~df_core['estado_actual'].astype(str).str.lower().isin(estados_excluir))
        ]
        st.markdown('<div style="background:#2980b9;color:#fff;padding:8px 16px;border-radius:6px;display:inline-block;font-weight:bold;">Después del Freeze</div>', unsafe_allow_html=True)
        st.success(f"Total de registros: {len(df_despues[columnas_a_mostrar])}")
        st.dataframe(
            df_despues[columnas_a_mostrar].style.apply(highlight_filas, axis=1),
            use_container_width=True,
            height=(35 * len(df_despues[columnas_a_mostrar]) + 40),
            hide_index=True
        )

         # --- Tabla solo implementados al final ---
        # Definir df_core y columnas_a_mostrar si no existen
        df_core_impl = df[df['jefatura'].str.lower().str.contains('core bancario|normativo', na=False)].copy()
        columnas_ocultas_impl = ['proyecto matriz', 'autor', 'codigo_proyecto', 'estabilizacion', 'codigo_estabilizacion']
        columnas_a_mostrar_impl = [col for col in df_core_impl.columns if col.lower() not in columnas_ocultas_impl]
        df_implementados = df_core_impl[df_core_impl['estado_actual'].astype(str).str.lower().isin(['finalizado', 'estabilización'])]
        if not df_implementados.empty:
            st.markdown('<div style="background:#2c7873;color:#fff;padding:8px 16px;border-radius:6px;display:inline-block;font-weight:bold;margin-top:24px;">Implementados (Finalizado o Estabilización)</div>', unsafe_allow_html=True)
            st.dataframe(
                df_implementados[columnas_a_mostrar_impl].style.apply(highlight_filas, axis=1),
                use_container_width=True,
                height=(35 * len(df_implementados[columnas_a_mostrar_impl]) + 40),
                hide_index=True
            )
        

    with main_tab2:
        st.subheader("Gráficos estadísticos de proyectos (solo Jefatura Core Bancario y Normativo)")

        # Filtrar solo jefatura Core Bancario y Normativo
        df_graf = df[df['jefatura'].str.lower().str.contains('core bancario|normativo', na=False)].copy()

        # Gráfico 1: Proyectos por Estado Actual en orden personalizado
        orden_estados = [
            "PMO-Detenido",
            "PMO-No iniciado",
            "PMO-Relevamiento PMO",
            "PMO-Pend. Validación técnica",
            "DESA-Listo p/ Análisis Técnico",
            "DESA-Análisis Técnico",
            "DESA-Pendiente Desarrollo",
            "DESA-En Curso",
            "QA-En Pruebas QA",
            "QA-En Pruebas Detenidas",
            "QA-En Pruebas UAT",
            "PROD-Para Comité de Pasajes",
            "Estabilización",
            "Finalizado"
        ]
        if 'estado_actual' in df_graf.columns:
            df_graf['estado_actual'] = pd.Categorical(df_graf['estado_actual'], categories=orden_estados, ordered=True)
            conteo_estados_ordenado = df_graf['estado_actual'].value_counts().reindex(orden_estados).fillna(0)
            fig2 = px.bar(
                conteo_estados_ordenado,
                x=conteo_estados_ordenado.index,
                y=conteo_estados_ordenado.values,
                labels={'x': 'Estado Actual', 'y': 'Cantidad'},
                title='Proyectos por Estado Actual (orden personalizado)',
                color=conteo_estados_ordenado.index,
                color_discrete_sequence=["#2980b9", "#2c7873", "#ffd700", "#444444"]*4
            )
            st.plotly_chart(fig2, use_container_width=True, key="fig2_estado_actual")

        # Gráfico 2: Proyectos por etiquetas Pres/Agos/25 y Post/Agos/25 (colores similares a Gráfico 1)
        def etiqueta_tipo(etiquetas):
            etiquetas = str(etiquetas).lower()
            if 'post/agos/25' in etiquetas:
                return 'Post/Agos/25'
            elif 'pres/agos/25' in etiquetas:
                return 'Pres/Agos/25'
            else:
                return 'Otro'
        df_graf['Tipo_Etiqueta'] = df_graf['etiquetas'].apply(etiqueta_tipo)
        conteo_etiqueta = df_graf['Tipo_Etiqueta'].value_counts()

        # Usar colores personalizados: implementados siempre #2c7873
        color_map = {
            'Post/Agos/25': '#ffd700',
            'Pres/Agos/25': '#2980b9',
            'Otro': '#2c7873'
        }
        # Si la categoría representa implementados, forzar color #2c7873
        # (Aquí solo aplica si quieres que 'Post/Agos/25' o 'Pres/Agos/25' sean implementados, ajusta según tu lógica)
        # Si quieres que todos sean #2c7873 cuando sean implementados, deberías hacerlo en el gráfico de implementados vs no implementados (fig4 y fig5)

        fig3 = px.pie(
            conteo_etiqueta,
            names=conteo_etiqueta.index,
            values=conteo_etiqueta.values,
            title='Distribución de proyectos por Pres/Agos/25 y Post/Agos/25',
            color=conteo_etiqueta.index,
            color_discrete_map=color_map,
            hole=0.3
        )
        fig3.update_traces(textinfo='label+percent')

        # Gráfico 3b: Torta de implementados/no implementados por tipo de etiqueta (con porcentajes)
        df_graf['implementado'] = df_graf['estado_actual'].apply(lambda x: 'Implementado' if str(x).strip().lower() in ['estabilización', 'finalizado'] else 'No implementado')
        df_etiqueta_impl = df_graf.groupby(['Tipo_Etiqueta', 'implementado']).size().reset_index(name='Cantidad')
        fig3b = px.sunburst(
            df_etiqueta_impl,
            path=['Tipo_Etiqueta', 'implementado'],
            values='Cantidad',
            color='implementado',
            color_discrete_map={
                'Implementado': '#2c7873',
                'No implementado': '#ffd700'
            },
            title='Implementados vs No implementados (pres/agos/25 y post/agos/25)'
        )
        fig3b.update_traces(textinfo='label+percent entry')

        # Mostrar ambos gráficos lado a lado
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(fig3, use_container_width=True, key="fig3_etiqueta")
        with col2:
            st.plotly_chart(fig3b, use_container_width=True, key="fig3b_sunburst")

        # Gráfico 4: Comparativa de implementados (Estabilización/Finalizado) vs no implementados
        def estado_implementado(estado):
            estado = str(estado).strip().lower()
            if estado in ['estabilización', 'finalizado']:
                return 'Implementado'
            else:
                return 'No implementado'
        df_graf['implementado'] = df_graf['estado_actual'].apply(estado_implementado)
        conteo_impl = df_graf['implementado'].value_counts()
        fig4 = px.pie(
            conteo_impl,
            names=conteo_impl.index,
            values=conteo_impl.values,
            title='Proyectos implementados vs no implementados',
            color_discrete_map={
                'Implementado': '#2c7873',
                'No implementado': '#ffd700'
            }
        )
        # Forzar color verde en el bloque de 'Implementado'
        fig4.update_traces(marker=dict(colors=[
            '#2c7873' if n == 'Implementado' else '#ffd700' for n in conteo_impl.index
        ]))

        # Gráfico 5: Estados agrupados (Implementado vs No implementado) por Gerencia Principal
        fig5 = None
        # Usar la misma lógica de identificación de Gerencia/Unidad que en el tab de agrupados por Gerencia/Unidad
        def extraer_gerencia(valor):
            if pd.isna(valor):
                return "(Sin dato)"
            return str(valor).split('>')[0].strip()

        col_gerencia = None
        for col in df_graf.columns:
            if col.replace(' ', '').lower() in ["gerencia/unidad", "gerenciaunidad"]:
                col_gerencia = col
                break
        if col_gerencia is not None:
            df_graf = df_graf[df_graf[col_gerencia].notna() & (df_graf[col_gerencia].astype(str).str.strip() != '')].copy()
            df_graf['Gerencia_Principal'] = df_graf[col_gerencia].apply(extraer_gerencia)
            df_grouped = df_graf[df_graf['Gerencia_Principal'].notna() & (df_graf['Gerencia_Principal'].astype(str).str.strip() != '')]
            df_grouped = df_grouped.groupby(['Gerencia_Principal', 'implementado']).size().reset_index(name='Cantidad')
            if not df_grouped.empty:
                fig5 = px.bar(
                    df_grouped,
                    x='Gerencia_Principal',
                    y='Cantidad',
                    color='implementado',
                    barmode='group',
                    title='Implementados vs No implementados por Gerencia Principal',
                    color_discrete_map={
                        'Implementado': '#2c7873',
                        'No implementado': '#ffd700'
                    }
                )
        # Mostrar gráfico 4 y 5 lado a lado
        col3, col4 = st.columns(2)
        with col3:
            st.plotly_chart(fig4, use_container_width=True, key="fig4_impl_col")
        with col4:
            if fig5 is not None:
                st.plotly_chart(fig5, use_container_width=True, key="fig5_gerencia")
            else:
                st.info("No hay datos suficientes para mostrar el gráfico por Gerencia Principal.")

    with main_tab3:
        st.subheader("Agrupados por Estados (solo Core)")
        # Agrupación y visualización por estado para Core
        def agrupar_jefatura(jef):
            if pd.isna(jef):
                return 'Sin Jefatura'
            jef = str(jef)
            if 'core bancario' in jef.lower() or 'normativo' in jef.lower():
                return 'Core'
            else:
                return 'Canales'
        df_agrupado = df.copy()
        df_agrupado['Grupo Jefatura'] = df_agrupado['jefatura'].apply(agrupar_jefatura)
        columnas_ocultas_agrupado = ['proyecto matriz', 'autor', 'codigo_proyecto', 'estabilizacion', 'codigo_estabilizacion']
        columnas_a_mostrar_agrupado = [col for col in df_agrupado.columns if col.lower() not in columnas_ocultas_agrupado]
        df_core = df_agrupado[df_agrupado['Grupo Jefatura'] == 'Core'].copy()

        # Orden deseado de estados
        orden_estados = [
            "PMO-Detenido",
            "PMO-No iniciado",
            "PMO-Relevamiento PMO",
            "PMO-Pend. Validación técnica",
            "DESA-Listo p/ Análisis Técnico",
            "DESA-Análisis Técnico",
            "DESA-Pendiente Desarrollo",
            "DESA-En Curso",
            "QA-En Pruebas QA",
            "QA-En Pruebas Detenidas",
            "QA-En Pruebas UAT",
            "PROD-Para Comité de Pasajes",
            "Estabilización",
            "Finalizado"
        ]
        # Colores suaves para los grupos
        def obtener_color_estado(estado):
            # Colores intensos y contrastantes para dark mode, verde más suave
            if estado.startswith("PMO"):
                return "#51748b"  
            elif estado.startswith("DESA"):
                return "#27a9ae"  
            elif estado.startswith("QA"):
                return "#ffd90066"  
            else:
                return "#2c7873"  

        # Convertir la columna a categoría para ordenar
        df_core['estado_actual'] = pd.Categorical(df_core['estado_actual'], categories=orden_estados, ordered=True)
        # Filtrar solo los estados presentes y en el orden deseado
        estados_core = [estado for estado in orden_estados if estado in df_core['estado_actual'].cat.categories and (df_core['estado_actual'] == estado).any()]

        st.write("Estados de proyectos Core:")
        for estado in estados_core:
            n = df_core[df_core['estado_actual'] == estado].shape[0]
            color = obtener_color_estado(estado)
            st.markdown(
                f'<div style="background-color:{color};padding:10px 16px;border-radius:6px;margin-bottom:0px;font-weight:bold;font-size:1.1em;">{estado} <span style="float:right">{n}</span></div>',
                unsafe_allow_html=True
            )
            if n > 0:
                st.dataframe(
                    df_core[df_core['estado_actual'] == estado][columnas_a_mostrar_agrupado].style.apply(highlight_filas, axis=1),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("No hay proyectos en este estado.")

    # Nuevo tab: Agrupados por Gerencia/Unidad
    with tab_gerencia:
        st.write("Proyectos agrupados por Gerencia/Unidad (primer nivel):")

        def extraer_gerencia(valor):
            if pd.isna(valor):
                return "(Sin dato)"
            return str(valor).split('>')[0].strip()

        # Buscar el nombre real de la columna 'Gerencia/Unidad' ignorando mayúsculas, minúsculas y espacios
        col_gerencia = None
        for col in df.columns:
            if col.replace(' ', '').lower() in ["gerencia/unidad", "gerenciaunidad"]:
                col_gerencia = col
                break
        if col_gerencia is None:
            st.error("No se encontró la columna 'Gerencia/Unidad' en los datos.")
        else:
            # Filtrar por jefatura que contenga 'core bancario' o 'normativo'
            df_jefatura = df[df['jefatura'].str.lower().str.contains('core bancario|normativo', na=False)].copy()
            df_jefatura['Gerencia_Principal'] = df_jefatura[col_gerencia].apply(extraer_gerencia)
            gerencias = df_jefatura['Gerencia_Principal'].dropna().unique()
            gerencias = sorted(gerencias)

            # Calcular columnas a mostrar para este tab (evitar columnas inexistentes)
            columnas_ocultas_gerencia = ['proyecto matriz', 'autor', 'codigo_proyecto', 'estabilizacion', 'codigo_estabilizacion', 'Grupo Jefatura']
            columnas_a_mostrar_gerencia = [col for col in df_jefatura.columns if col.lower() not in [c.lower() for c in columnas_ocultas_gerencia]]

            for ger in gerencias:
                n = df_jefatura[df_jefatura['Gerencia_Principal'] == ger].shape[0]
                color = "#b96329"  # azul fuerte
                st.markdown(
                    f'<div style="background-color:{color};padding:10px 16px;border-radius:6px;margin-bottom:0px;font-weight:bold;font-size:1.1em;">{ger} <span style="float:right">{n}</span></div>',
                    unsafe_allow_html=True
                )
                if n > 0:
                    st.dataframe(
                        df_jefatura[df_jefatura['Gerencia_Principal'] == ger][columnas_a_mostrar_gerencia].style.apply(highlight_filas, axis=1),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.info("No hay proyectos en esta gerencia/unidad.")
    # ...existing code...
else:
    st.info("Por favor, sube un archivo Excel para comenzar.")