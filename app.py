import os
import streamlit as st
import pandas as pd
import plotly.express as px

# Configuración de página
st.set_page_config(page_title="Dashboard Representantes", layout="wide")

# ---------------------------------------------------------------------------
# CONFIGURACIÓN DE FUENTES DE DATOS
# Nombres y hojas REALES de los archivos del repositorio.
# Se asume que los .xlsx están en la misma carpeta que este app.py.
#
# NOTA: el archivo "Cuadro_40_-_Representante__Proclamados_Parciales_2014..."
# (hoja 'repre2014_parciales') NO se incluye a propósito: tiene una estructura
# distinta (5 columnas, sin 'TOTAL DE VOTOS' ni 'OBSERVACIÓN') y corresponde a
# proclamaciones parciales. Si en algún momento quieres sumarlo, requiere un
# bloque de carga aparte, no el mismo esquema de 7 columnas.
# ---------------------------------------------------------------------------
FUENTES = {
    2014: ("Cuadro_13_-_Representantes_Proclamados_2014_DESBLOQ.xlsx", "repre2014"),
    2019: ("Representantes_Proclamados_2019_DESBLOQ.xlsx",            "repre2019"),
    2024: ("Cuadro_13_Representantes_Proclamados_2024_DESBLOQ.xlsx",  "repre2024"),
}

# Columnas REALES de las hojas de Excel (7). El 'Año' se agrega en código,
# por eso NO va en esta lista de validación.
COLUMNAS_FUENTE = ['Provincia', 'Distrito', 'Corregimiento',
                   'Candidato', 'Partido', 'Votos', 'Observacion']

# ---------------------------------------------------------------------------
# NORMALIZACIÓN DE PROVINCIAS / COMARCAS
# La misma entidad aparece escrita distinto entre años (tildes, abreviaturas).
# Sin unificarla, una misma comarca se cuenta como dos en el filtro y los
# gráficos. La clave es la grafía ORIGINAL (en mayúsculas, espacios colapsados)
# y el valor es la forma canónica. Los archivos .xlsx NO se modifican: el
# mapeo vive aquí, versionado y auditable.
#
# Para agregar equivalencias nuevas, añade una línea: 'GRAFÍA ORIGINAL': 'CANÓNICA'.
# ---------------------------------------------------------------------------
NORMALIZACION_PROVINCIAS = {
    'COMARCA EMBERÁ':           'COMARCA EMBERÁ WOUNAÁN',   # 2014 -> forma completa
    'COMARCA NGÄBE BUGLÉ':      'COMARCA NGÄBÉ BUGLÉ',      # 2014 (sin tilde en NGÄBÉ)
    'COMARCA K. DE MADUNGANDÍ': 'COMARCA KUNA DE MADUGANDÍ',# 2014 (abrev. + "MADUNGANDÍ")
    'COMARCA K. DE WARGANDÍ':   'COMARCA KUNA DE WARGANDÍ', # 2014 (abreviatura "K.")
}

# Universo de provincias y comarcas válidas (en forma canónica). Sirve para
# avisar si en una actualización aparece una grafía que el diccionario aún no
# contempla, en lugar de que pase desapercibida.
PROVINCIAS_ESPERADAS = {
    'BOCAS DEL TORO', 'CHIRIQUÍ', 'COCLÉ', 'COLÓN', 'DARIÉN', 'HERRERA',
    'LOS SANTOS', 'PANAMÁ', 'PANAMÁ OESTE', 'VERAGUAS',
    'COMARCA EMBERÁ WOUNAÁN', 'COMARCA NGÄBÉ BUGLÉ', 'COMARCA KUNA YALA',
    'COMARCA KUNA DE MADUGANDÍ', 'COMARCA KUNA DE WARGANDÍ',
    'COMARCA NASO TJËR DI',
}


def normaliza_provincia(p):
    """Devuelve la forma canónica de la provincia/comarca.
    Pasa a mayúsculas y colapsa espacios; si la grafía está en el diccionario,
    aplica la equivalencia. Si no, devuelve la versión saneada (no la rompe)."""
    if pd.isna(p):
        return p
    clave = ' '.join(str(p).upper().split())
    return NORMALIZACION_PROVINCIAS.get(clave, clave)


def _firmas_archivos():
    """Devuelve la fecha de modificación de cada archivo.
    Sirve como clave de caché: si actualizas un .xlsx, cambia la firma
    y @st.cache_data recarga automáticamente (sin datos previos)."""
    firmas = {}
    for año, (ruta, _) in FUENTES.items():
        firmas[año] = os.path.getmtime(ruta) if os.path.exists(ruta) else None
    return firmas


@st.cache_data
def load_data(_firmas):
    # El argumento _firmas fuerza la recarga cuando cambian los archivos.
    frames = []
    for año, (ruta, hoja) in FUENTES.items():
        if not os.path.exists(ruta):
            st.error(f"No se encontró el archivo: {ruta}")
            st.stop()

        df = pd.read_excel(ruta, sheet_name=hoja)

        # Validar PRIMERO contra las columnas reales del Excel (7), ANTES de
        # agregar 'Año'. Así detectamos un cambio de estructura de verdad y el
        # renombrado posicional no termina desalineando las columnas.
        if df.shape[1] != len(COLUMNAS_FUENTE):
            st.error(
                f"La hoja '{hoja}' de {ruta} tiene {df.shape[1]} columnas "
                f"y se esperaban {len(COLUMNAS_FUENTE)}. Revisa la estructura."
            )
            st.stop()

        df.columns = COLUMNAS_FUENTE
        df['Año'] = año

        # --- LIMPIEZA DE FILAS BASURA ---
        # Algunas hojas (p. ej. 2019) traen al final una fila vacía y una fila
        # de pie con "Fuente: Actas de proclamación...". Si no se eliminan,
        # inflan el conteo de representantes, crean una "provincia" falsa en el
        # filtro y un partido fantasma "NAN" en el gráfico de balance de poder.
        df = df[df['Candidato'].notna() & df['Provincia'].notna()]
        df = df[~df['Provincia'].astype(str).str.strip().str.upper()
                .str.startswith(('FUENTE', 'TOTAL'))]

        frames.append(df)

    df_master = pd.concat(frames, ignore_index=True)

    # Votos a numérico por seguridad
    df_master['Votos'] = pd.to_numeric(df_master['Votos'], errors='coerce').fillna(0)

    # Unificar grafías de provincia/comarca entre años (ver diccionario arriba)
    df_master['Provincia'] = df_master['Provincia'].apply(normaliza_provincia)

    # Estandarización de partidos
    def clean_party(p):
        if pd.isna(p):
            return p
        p = str(p).upper().strip()
        if '/' in p:
            p = p.split('/')[0].strip()
        # Libre postulación: unificar 'LP' y 'LIBRE POSTULACIÓN'
        if p in ('LP', 'LIBRE POSTULACIÓN') or 'LIBRE POSTULAC' in p:
            return 'LIBRE POSTULACIÓN'
        # OJO: revisa esta regla. PP = Partido Popular, no necesariamente
        # Panameñista. Como tú trabajas en el TE, decide el mapeo correcto.
        # if 'PANAMEÑISTA' in p or p == 'PP':
        #     return 'PANAMEÑISTA'
        return p

    df_master['Partido'] = df_master['Partido'].apply(clean_party)
    return df_master


df = load_data(_firmas_archivos())

# Aviso si aparece una provincia/comarca que el diccionario aún no contempla
# (p. ej. una grafía nueva en una actualización de los cuadros). No detiene la
# app: solo te alerta para que agregues la equivalencia en NORMALIZACION_PROVINCIAS.
_desconocidas = sorted(set(df['Provincia'].dropna().unique()) - PROVINCIAS_ESPERADAS)
if _desconocidas:
    st.warning(
        "Provincias/comarcas no reconocidas (revisa el diccionario "
        "NORMALIZACION_PROVINCIAS): " + ", ".join(_desconocidas)
    )

# Título
st.title("EVOLUCIÓN DE LOS PROCESOS ELECTORALES - REPRESENTANTES 2014 - 2019 - 2024")

# Sidebar
st.sidebar.header("Filtros")
if st.sidebar.button("🔄 Recargar datos (limpiar caché)"):
    st.cache_data.clear()
    st.rerun()

years = st.sidebar.multiselect("Año", [2014, 2019, 2024], default=[2014, 2019, 2024])
provincias = st.sidebar.multiselect("Provincia", sorted(df['Provincia'].dropna().unique()))

mask = (df['Año'].isin(years))
if provincias:
    mask &= (df['Provincia'].isin(provincias))
df_filt = df[mask]

# Tabs
tab1, tab2, tab3 = st.tabs(["Métricas Globales", "Balance de Poder", "Detalle por Corregimiento"])

with tab1:
    st.subheader("Resumen por Año Electoral")
    st.markdown("---")

    # Obtener los datos agrupados
    resumen_anual = df_filt.groupby('Año').agg(
        Total_Representantes=('Candidato', 'count'),
        Fuerzas_Politicas=('Partido', 'nunique'),
        Votos_Validos=('Votos', 'sum')
    ).sort_index()

    # Definimos colores para cada año para dar el toque llamativo
    colores = {2014: "#1f77b4", 2019: "#ff7f0e", 2024: "#2ca02c"}
    
    # Crear columnas para la "infografía"
    cols = st.columns(len(resumen_anual))
    
    for i, (año, row) in enumerate(resumen_anual.iterrows()):
        color = colores.get(año, "#333333")
        with cols[i]:
            # Encabezado del año con color
            st.markdown(f"<h2 style='color: {color};'>{año}</h2>", unsafe_allow_html=True)
            
            # Métricas con estilo personalizado
            st.markdown(f"**Representantes:** <span style='color: {color}; font-size: 20px;'>{int(row['Total_Representantes']):,}</span>", unsafe_allow_html=True)
            st.markdown(f"**Fuerzas Políticas:** <span style='color: {color}; font-size: 20px;'>{int(row['Fuerzas_Politicas']):,}</span>", unsafe_allow_html=True)
            st.markdown(f"**Votos Válidos:** <span style='color: {color}; font-size: 20px;'>{int(row['Votos_Validos']):,}</span>", unsafe_allow_html=True)

with tab2:
    st.subheader("Distribución de Representantes por Partido")
    
    # Agrupamos los datos por año y partido
    dist_partido = df_filt.groupby(['Año', 'Partido']).size().reset_index(name='Conteo')
    
    # Calculamos el total de representantes por año para el cálculo de porcentajes
    # Plotly Express hace esto automáticamente en el tooltip
    
    # Creamos una gráfica de anillo (pie con hole) por cada año seleccionado
    # Usamos facetas (facet_col) para que cada año tenga su propia gráfica de anillo
    fig2 = px.pie(dist_partido, 
                  values='Conteo', 
                  names='Partido', 
                  facet_col='Año',
                  hole=0.4,
                  title="Balance de Poder: Representantes por Fuerza Política",
                  labels={'Conteo': 'Representantes'})
    
    # Ajustamos la configuración para mostrar valor absoluto y porcentaje
    fig2.update_traces(textinfo='percent+value')
    
    st.plotly_chart(fig2, use_container_width=True)

with tab3:
    st.subheader("Buscador de Resultados")
    search = st.text_input("Buscar por candidato o corregimiento")
    if search:
        df_search = df_filt[
            df_filt['Candidato'].astype(str).str.contains(search, case=False, na=False) |
            df_filt['Corregimiento'].astype(str).str.contains(search, case=False, na=False)
        ]
        st.dataframe(df_search)
    else:
        st.dataframe(df_filt)
