import streamlit as st
import pandas as pd
import openpyxl
import io

st.set_page_config(
    page_title="Informe Epidemiológico · Laboratorio",
    page_icon="🦠",
    layout="centered"
)

st.title("🦠 Informe Epidemiológico")
st.caption("Sube el Excel con la hoja 'Gral' y las hojas 'Centinela' y 'Urgencia'. "
           "La app calcula y rellena las tablas automáticamente.")
st.divider()

# ── Constantes ────────────────────────────────────────────
VIRUS_MAP = {
    'VIRUS RESPIRATORIO SINCICIAL': 'VRS',
    'ADENOVIRUS': 'Adenovirus',
    'PARAINFLUENZA': 'Parainfluenza',
    'INFLUENZA A': 'Influenza A',
    'INFLUENZA B': 'Influenza B',
    'METAPNEUMOVIRUS': 'Metapneumovirus',
    'RHINOVIRUS': 'Rhinovirus',
    'SARS COV 2 (COVID-19)': 'Covid-19',
}
VIRUS_ORDER = list(VIRUS_MAP.values())  # display order

AGE_RANGES = ['Menor 1 año', '1-4 años', '5-14 años', '15-54 años', '55-64 años', '65 y más']
ALL_AGES   = ['Total'] + AGE_RANGES

# Filas en openpyxl (1-indexed) para la primera tabla
VIRUS_ROWS_TABLE1 = {
    'VRS':             12,
    'Adenovirus':      13,
    'Parainfluenza':   14,
    'Influenza A':     15,
    'Influenza B':     16,
    'Metapneumovirus': 17,
    'Rhinovirus':      18,
    'Covid-19':        19,
    'Negativos':       20,
    '_Total':          21,
}
TABLE2_OFFSET = 24  # segunda tabla empieza 24 filas más abajo

# Columnas para M y F por rango etario
AGE_COLS = {
    'Total':        (2,  3),
    'Menor 1 año':  (4,  5),
    '1-4 años':     (6,  7),
    '5-14 años':    (8,  9),
    '15-54 años':  (10, 11),
    '55-64 años':  (12, 13),
    '65 y más':    (14, 15),
}

CENTINELA_PROC = ['CESFAM CERRO ALTO']
URGENCIA_PROC  = ['URGENCIA PEDIATRIA', 'URGENCIA ADULTO', 'INGRESO MATER']


# ── Lógica de cálculo ─────────────────────────────────────
def compute_tables(df_gral, procedencias):
    """Retorna (table1, table2): {virus: {age: {M:int, F:int}}}"""
    mask = df_gral['Procedencia'].str.strip().str.upper().isin(
        [p.upper() for p in procedencias]
    )
    df = df_gral[mask].copy()
    if df.empty:
        return None, None

    # Detectar columna de virus dinámicamente
    virus_col = next((c for c in df.columns if 'tru' in c.lower() and 'staci' in c.lower()), None)
    if not virus_col:
        virus_col = 'Prestación Estructura'

    df['_fecha_str'] = df['Fecha'].astype(str)
    df['_pos'] = df['Valor'].astype(str).str.strip().str.upper() == 'POSITIVO'

    # Identificar episodios con coinfección (mismo paciente + mismo timestamp, ≥2 positivos)
    coinf_keys = set()
    for (doc, fecha), grp in df.groupby(['Documento', '_fecha_str']):
        if grp['_pos'].sum() >= 2:
            coinf_keys.add((doc, fecha))

    def new_table():
        rows = VIRUS_ORDER + ['Negativos']
        return {v: {a: {'M': 0, 'F': 0} for a in ALL_AGES} for v in rows}

    t1, t2 = new_table(), new_table()

    def add(table, virus, age, sex):
        if virus in table:
            table[virus]['Total'][sex] += 1
            if age in table[virus]:
                table[virus][age][sex] += 1

    for (doc, fecha), grp in df.groupby(['Documento', '_fecha_str']):
        pos_rows = grp[grp['_pos']]
        sex = 'M' if 'MASCUL' in str(grp.iloc[0]['Sexo']).upper() else 'F'
        age = str(grp.iloc[0].get('Rango Etario', 'Sin dato')).strip()

        if len(pos_rows) == 0:
            add(t1, 'Negativos', age, sex)
        elif len(pos_rows) == 1:
            raw = str(pos_rows.iloc[0][virus_col]).strip().upper()
            v   = VIRUS_MAP.get(raw)
            if v:
                add(t1, v, age, sex)
        else:
            for _, r in pos_rows.iterrows():
                raw = str(r[virus_col]).strip().upper()
                v   = VIRUS_MAP.get(raw)
                if v:
                    add(t2, v, age, sex)

    return t1, t2


def total_row(table):
    """Suma total de todas las filas del virus (incluye Negativos)."""
    tots = {a: {'M': 0, 'F': 0} for a in ALL_AGES}
    for v in VIRUS_ORDER + ['Negativos']:
        for age in ALL_AGES:
            tots[age]['M'] += table.get(v, {}).get(age, {}).get('M', 0)
            tots[age]['F'] += table.get(v, {}).get(age, {}).get('F', 0)
    return tots


# ── Escritura en plantilla ────────────────────────────────
def set_val(ws, row, col, value):
    ws.cell(row=row, column=col).value = value if value else None


def fill_sheet(ws, table1, table2):
    all_virus = VIRUS_ORDER + ['Negativos']

    for virus in all_virus:
        row1 = VIRUS_ROWS_TABLE1.get(virus)
        row2 = row1 + TABLE2_OFFSET if row1 else None
        if not row1:
            continue
        for age, (col_m, col_f) in AGE_COLS.items():
            d1 = (table1 or {}).get(virus, {}).get(age, {'M': 0, 'F': 0})
            set_val(ws, row1, col_m, d1['M'])
            set_val(ws, row1, col_f, d1['F'])
            if table2 and row2:
                d2 = table2.get(virus, {}).get(age, {'M': 0, 'F': 0})
                set_val(ws, row2, col_m, d2['M'])
                set_val(ws, row2, col_f, d2['F'])

    # Fila Total
    tot1_row = VIRUS_ROWS_TABLE1['_Total']
    tot2_row = tot1_row + TABLE2_OFFSET
    if table1:
        tots1 = total_row(table1)
        for age, (col_m, col_f) in AGE_COLS.items():
            set_val(ws, tot1_row, col_m, tots1[age]['M'])
            set_val(ws, tot1_row, col_f, tots1[age]['F'])
    if table2:
        tots2 = total_row(table2)
        for age, (col_m, col_f) in AGE_COLS.items():
            set_val(ws, tot2_row, col_m, tots2[age]['M'])
            set_val(ws, tot2_row, col_f, tots2[age]['F'])


# ── Vista previa ──────────────────────────────────────────
def preview_table(table, label):
    if not table:
        st.info(f"Sin datos para {label}.")
        return
    rows = []
    for v in VIRUS_ORDER + ['Negativos']:
        d = table.get(v, {}).get('Total', {'M': 0, 'F': 0})
        rows.append({'Virus': v, 'Masc.': d['M'], 'Fem.': d['F'],
                     'Total': d['M'] + d['F']})
    df_prev = pd.DataFrame(rows)
    df_prev = df_prev[df_prev['Total'] > 0]
    if df_prev.empty:
        st.write("(sin casos)")
    else:
        st.dataframe(df_prev, use_container_width=True, hide_index=True)


# ── Interfaz ──────────────────────────────────────────────
archivo = st.file_uploader(
    "Selecciona el archivo Excel (debe tener hojas: Gral, Centinela, Urgencia)",
    type=["xlsx", "xls"],
    help="El archivo de salida de la primera herramienta, con las tres hojas."
)

if archivo:
    with st.spinner("Leyendo archivo..."):
        try:
            todas_hojas = pd.read_excel(archivo, sheet_name=None, header=0)
        except Exception as e:
            st.error(f"No se pudo leer el archivo: {e}")
            st.stop()

    hojas = {k.strip(): v for k, v in todas_hojas.items()}
    requeridas = ['Gral', 'Centinela', 'Urgencia']
    faltantes = [h for h in requeridas if h not in hojas]
    if faltantes:
        st.error(f"El archivo no tiene la(s) hoja(s): {', '.join(faltantes)}. "
                 "Verifica los nombres de las hojas.")
        st.stop()

    df_gral = hojas['Gral']

    with st.spinner("Calculando tablas..."):
        t1_cent, t2_cent = compute_tables(df_gral, CENTINELA_PROC)
        t1_urg,  t2_urg  = compute_tables(df_gral, URGENCIA_PROC)

    st.divider()
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Centinela")
        proc_cent = df_gral[
            df_gral['Procedencia'].str.strip().str.upper()
            .isin([p.upper() for p in CENTINELA_PROC])
        ]
        st.metric("Casos analizados", len(proc_cent))
        st.write("**Infecciones únicas**")
        preview_table(t1_cent, "Centinela")
        st.write("**Coinfecciones**")
        preview_table(t2_cent, "Centinela coinfecciones")

    with col2:
        st.subheader("Urgencia")
        proc_urg = df_gral[
            df_gral['Procedencia'].str.strip().str.upper()
            .isin([p.upper() for p in URGENCIA_PROC])
        ]
        st.metric("Casos analizados", len(proc_urg))
        st.write("**Infecciones únicas**")
        preview_table(t1_urg, "Urgencia")
        st.write("**Coinfecciones**")
        preview_table(t2_urg, "Urgencia coinfecciones")

    # Generar Excel con openpyxl
    with st.spinner("Generando Excel..."):
        archivo.seek(0)
        wb = openpyxl.load_workbook(archivo)

        ws_cent = wb['Centinela']
        ws_urg  = wb['Urgencia']

        fill_sheet(ws_cent, t1_cent, t2_cent)
        fill_sheet(ws_urg,  t1_urg,  t2_urg)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

    nombre_salida = archivo.name.rsplit('.', 1)[0] + '_informe.xlsx'
    st.divider()
    st.download_button(
        label="⬇️ Descargar informe completo",
        data=output,
        file_name=nombre_salida,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )
