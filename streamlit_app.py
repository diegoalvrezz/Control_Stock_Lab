import streamlit as st
import pandas as pd
import numpy as np
import datetime
import shutil
import os
from io import BytesIO
import itertools

st.set_page_config(page_title="Control de Stock con Lotes", layout="centered")

STOCK_FILE = "Stock_Original.xlsx"
VERSIONS_DIR = "versions"
ORIGINAL_FILE = os.path.join(VERSIONS_DIR, "Stock_Original.xlsx")

os.makedirs(VERSIONS_DIR, exist_ok=True)

def init_original():
    """Copia STOCK_FILE en versions/Stock_Original.xlsx si no existe."""
    if not os.path.exists(ORIGINAL_FILE):
        if os.path.exists(STOCK_FILE):
            shutil.copy(STOCK_FILE, ORIGINAL_FILE)
        else:
            st.error(f"No se encontró {STOCK_FILE}. Sube el archivo o revisa la ruta.")

init_original()

def load_data():
    """Lee todas las hojas de STOCK_FILE y elimina la columna 'Restantes' si existe."""
    try:
        data = pd.read_excel(STOCK_FILE, sheet_name=None, engine="openpyxl")
        for sheet, df_sheet in data.items():
            if "Restantes" in df_sheet.columns:
                df_sheet.drop(columns=["Restantes"], inplace=True, errors="ignore")
        return data
    except FileNotFoundError:
        st.error("❌ No se encontró el archivo principal.")
        return None
    except Exception as e:
        st.error(f"❌ Error al cargar la base de datos: {e}")
        return None

data_dict = load_data()

def enforce_types(df: pd.DataFrame):
    """Fuerza tipos en las columnas habituales."""
    if "Ref. Saturno" in df.columns:
        df["Ref. Saturno"] = pd.to_numeric(df["Ref. Saturno"], errors="coerce").fillna(0).astype(int)
    if "Ref. Fisher" in df.columns:
        df["Ref. Fisher"] = df["Ref. Fisher"].astype(str)
    if "Nombre producto" in df.columns:
        df["Nombre producto"] = df["Nombre producto"].astype(str)
    if "Tª" in df.columns:
        df["Tª"] = df["Tª"].astype(str)
    if "Uds." in df.columns:
        df["Uds."] = pd.to_numeric(df["Uds."], errors="coerce").fillna(0).astype(int)
    if "NºLote" in df.columns:
        df["NºLote"] = pd.to_numeric(df["NºLote"], errors="coerce").fillna(0).astype(int)
    for col in ["Caducidad", "Fecha Pedida", "Fecha Llegada"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    if "Sitio almacenaje" in df.columns:
        df["Sitio almacenaje"] = df["Sitio almacenaje"].astype(str)
    if "Stock" in df.columns:
        df["Stock"] = pd.to_numeric(df["Stock"], errors="coerce").fillna(0).astype(int)
    return df

def crear_nueva_version_filename():
    fecha_hora = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return os.path.join(VERSIONS_DIR, f"Stock_{fecha_hora}.xlsx")

def generar_excel_en_memoria(df_act: pd.DataFrame, sheet_nm="Hoja1"):
    """Genera un Excel en memoria para descargar."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_act.to_excel(writer, index=False, sheet_name=sheet_nm)
    output.seek(0)
    return output.getvalue()

# -------------------------------------------------------------------------
# DICCIONARIO DE LOTES (definición de grupos)
# -------------------------------------------------------------------------
LOTS_DATA = {
    "FOCUS": {
        "Panel Oncomine Focus Library Assay Chef Ready": [
            "Primers DNA", "Primers RNA", "Reagents DL8", "Chef supplies (plásticos)", "Placas", "Solutions DL8"
        ],
        "Ion 510/520/530 kit-Chef (TEMPLADO)": [
            "Chef Reagents", "Chef Solutions", "Chef supplies (plásticos)", "Solutions Reagent S5", "Botellas S5"
        ],
        "Recover All TM Multi-Sample RNA/DNA Isolation workflow-Kit": [
            "Kit extracción DNA/RNA", "RecoverAll TM kit (Dnase, protease,…)", "H2O RNA free",
            "Tubos fondo cónico", "Superscript VILO cDNA Syntheis Kit", "Qubit 1x dsDNA HS Assay kit (100 reactions)"
        ],
        "Chip secuenciación liberación de protones 6 millones de lecturas": []
    },
    "OCA": {
        "Panel OCA Library Assay Chef Ready": [
            "Primers DNA", "Primers RNA", "Reagents DL8", "Chef supplies (plásticos)", "Placas", "Solutions DL8"
        ],
        "kit-Chef (TEMPLADO)": [
            "Ion 540 TM Chef Reagents", "Chef Solutions", "Chef supplies (plásticos)",
            "Solutions Reagent S5", "Botellas S5"
        ],
        "Chip secuenciación liberación de protones 6 millones de lecturas": [
            "Ion 540 TM Chip Kit"
        ],
        "Recover All TM Multi-Sample RNA/DNA Isolation workflow-Kit": [
            "Kit extracción DNA/RNA", "RecoverAll TM kit (Dnase, protease,…)", "H2O RNA free", "Tubos fondo cónico"
        ]
    },
    "OCA PLUS": {
        "Panel OCA-PLUS Library Assay Chef Ready": [
            "Primers DNA", "Uracil-DNA Glycosylase heat-labile", "Reagents DL8",
            "Chef supplies (plásticos)", "Placas", "Solutions DL8"
        ],
        "kit-Chef (TEMPLADO)": [
            "Ion 550 TM Chef Reagents", "Chef Solutions", "Chef Supplies (plásticos)",
            "Solutions Reagent S5", "Botellas S5", "Chip secuenciación Ion 550 TM Chip Kit"
        ],
        "Recover All TM Multi-Sample RNA/DNA Isolation workflow-Kit": [
            "Kit extracción DNA/RNA", "RecoverAll TM kit (Dnase, protease,…)", "H2O RNA free", "Tubos fondo cónico"
        ]
    }
}

# Agrupación por "Ref. Saturno"
panel_order = ["FOCUS", "OCA", "OCA PLUS"]
colors = [
    "#FED7D7", "#FEE2E2", "#FFEDD5", "#FEF9C3", "#D9F99D",
    "#CFFAFE", "#E0E7FF", "#FBCFE8", "#F9A8D4", "#E9D5FF",
    "#FFD700", "#F0FFF0", "#D1FAE5", "#BAFEE2", "#A7F3D0", "#FFEC99"
]

def build_group_info_by_ref(df: pd.DataFrame, panel_default=None):
    """
    Agrupa los registros según "Ref. Saturno" y asigna:
      - GroupID igual a "Ref. Saturno"
      - GroupCount: tamaño del grupo
      - ColorGroup: color asignado a ese grupo
      - EsTitulo: se marca como título la fila cuyo "Nombre producto" coincida con
        alguno de los títulos definidos en LOTS_DATA para el panel; si no se encuentra,
        se marca la primera fila del grupo.
      - MultiSort y NotTitulo para ordenar.
    """
    df = df.copy()
    df["GroupID"] = df["Ref. Saturno"]
    group_sizes = df.groupby("GroupID").size().to_dict()
    df["GroupCount"] = df["GroupID"].apply(lambda x: group_sizes.get(x, 0))
    
    unique_ids = sorted(df["GroupID"].unique())
    group_color_mapping = {}
    color_cycle_local = itertools.cycle(colors)
    for gid in unique_ids:
        group_color_mapping[gid] = next(color_cycle_local)
    df["ColorGroup"] = df["GroupID"].apply(lambda x: group_color_mapping.get(x, "#FFFFFF"))
    
    group_titles = []
    if panel_default in LOTS_DATA:
        group_titles = [t.strip().lower() for t in LOTS_DATA[panel_default].keys()]
    df["EsTitulo"] = False
    for gid, group_df in df.groupby("GroupID"):
        mask = group_df["Nombre producto"].str.strip().str.lower().isin(group_titles)
        if mask.any():
            idxs = group_df[mask].index
            df.loc[idxs, "EsTitulo"] = True
        else:
            first_idx = group_df.index[0]
            df.at[first_idx, "EsTitulo"] = True

    df["MultiSort"] = df["GroupCount"].apply(lambda x: 0 if x > 1 else 1)
    df["NotTitulo"] = df["EsTitulo"].apply(lambda x: 0 if x else 1)
    return df

def calc_alarma(row):
    """Devuelve ícono según Stock y Fecha Pedida."""
    s = row.get("Stock", 0)
    fp = row.get("Fecha Pedida", None)
    if s == 0 and pd.isna(fp):
        return "🔴"
    elif s == 0 and not pd.isna(fp):
        return "🟨"
    return ""

def style_lote(row):
    """Aplica estilo según 'ColorGroup'; si EsTitulo es True, negrita en 'Nombre producto'."""
    bg = row.get("ColorGroup", "")
    es_titulo = row.get("EsTitulo", False)
    styles = [f"background-color:{bg}"] * len(row)
    if es_titulo and "Nombre producto" in row.index:
        idx = row.index.get_loc("Nombre producto")
        styles[idx] += "; font-weight:bold"
    return styles

st.markdown("""
    <style>
    .big-select select {
        font-size: 18px;
        height: auto;
    }
    </style>
    """, unsafe_allow_html=True)

# -------------------------------------------------------------------------
# DEFINICIÓN GLOBAL DE LA HOJA A EDITAR (en la barra lateral)
# -------------------------------------------------------------------------
hojas_principales = list(data_dict.keys())
sheet_name = st.sidebar.selectbox("Selecciona la hoja a editar:", hojas_principales, key="sheet_name")

# -------------------------------------------------------------------------
# CONTROL DE VERSIONES (INTOCABLE)
# -------------------------------------------------------------------------
with st.sidebar.expander("Ver / Gestionar versiones guardadas", expanded=False):
    if data_dict:
        files = sorted(os.listdir(VERSIONS_DIR))
        versions_no_original = [f for f in files if f != "Stock_Original.xlsx"]
        if versions_no_original:
            version_sel = st.selectbox("Selecciona versión:", versions_no_original)
            if version_sel:
                file_path = os.path.join(VERSIONS_DIR, version_sel)
                if os.path.isfile(file_path):
                    with open(file_path, "rb") as excel_file:
                        excel_bytes = excel_file.read()
                    st.download_button(
                        label=f"Descargar {version_sel}",
                        data=excel_bytes,
                        file_name=version_sel,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                if st.button("Eliminar esta versión"):
                    try:
                        os.remove(file_path)
                        st.warning(f"Versión '{version_sel}' eliminada.")
                        st.experimental_rerun()
                    except Exception as e:
                        st.error(f"Error al eliminar versión: {e}")
        else:
            st.write("No hay versiones guardadas (excepto la original).")
        if st.button("Eliminar TODAS las versiones (excepto original)"):
            for f in versions_no_original:
                try:
                    os.remove(os.path.join(VERSIONS_DIR, f))
                except:
                    pass
            st.info("Todas las versiones (excepto la original) eliminadas.")
            st.experimental_rerun()
    else:
        st.error("No hay data_dict. Verifica Stock_Original.xlsx.")

# Botón de Limpiar Base de Datos (resetea al original)
with st.sidebar:
    if st.button("Limpiar Base de Datos"):
        original_path = os.path.join(VERSIONS_DIR, "Stock_Original.xlsx")
        if os.path.exists(original_path):
            try:
                shutil.copy(original_path, STOCK_FILE)
                st.success("✅ Base de datos restaurada a la versión original.")
                st.experimental_rerun()
            except Exception as e:
                st.error(f"Error restaurando la base de datos: {e}")
        else:
            st.error("❌ No se encontró la copia original en 'versions/Stock_Original.xlsx'.")

# -------------------------------------------------------------------------
# SECCIÓN: Recepción de lote completo
# -------------------------------------------------------------------------
with st.expander("Recepción de lote completo", expanded=False):
    st.subheader("Confirmar recepción de lote")
    if sheet_name in LOTS_DATA:
        lot_titles = list(LOTS_DATA[sheet_name].keys())
    else:
        lot_titles = []
    selected_lot = st.selectbox("Seleccione el título del lote", lot_titles, key="selected_lot")
    if selected_lot:
        df_current = enforce_types(data_dict[sheet_name])
        st.write("Reactivos del lote:")
        st.dataframe(df_current[["Nombre producto", "Ref. Fisher"]])
        row_lot = df_current[df_current["Nombre producto"].str.lower() == selected_lot.lower()]
        if not row_lot.empty:
            lot_ref = row_lot.iloc[0]["Ref. Saturno"]
            df_lote = df_current[df_current["Ref. Saturno"] == lot_ref].copy()
            st.write("Edite la información común del lote:")
            tipo_almacenaje = st.selectbox("Tipo de Almacenaje", ["Congelador 1", "Congelador 2", "Frigorífico", "Tª Ambiente"], key="tipo_recepcion")
            subopcion_recep = ""
            if tipo_almacenaje == "Congelador 1":
                subopcion_recep = st.selectbox("Cajón (1 Arriba,8 Abajo)", [f"Cajón {i}" for i in range(1,9)], key="cajon_recep")
            elif tipo_almacenaje == "Congelador 2":
                subopcion_recep = st.selectbox("Cajón (1 Arriba,6 Abajo)", [f"Cajón {i}" for i in range(1,7)], key="cajon_recep")
            elif tipo_almacenaje == "Frigorífico":
                subopcion_recep = st.selectbox("Balda", [f"Balda {i}" for i in range(1,8)] + ["Puerta"], key="balda_recep")
            elif tipo_almacenaje == "Tª Ambiente":
                subopcion_recep = st.text_input("Comentario (opcional)", key="coment_recep")
            sitio_recep = f"{tipo_almacenaje} - {subopcion_recep}" if subopcion_recep else tipo_almacenaje
            cols_edit = ["NºLote", "Fecha Llegada", "Caducidad", "Sitio almacenaje"]
            df_lote["Sitio almacenaje"] = sitio_recep
            # Usamos st.data_editor (Streamlit 1.18+)
            df_edit = st.data_editor(df_lote[cols_edit], num_rows="dynamic", key="edicion_lote")
            if st.button("Guardar Recepción del Lote"):
                for idx in df_lote.index:
                    for col in cols_edit:
                        data_dict[sheet_name].at[idx, col] = df_edit.at[idx, col]
                new_file = crear_nueva_version_filename()
                with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
                    for sht, df_sht in data_dict.items():
                        temp = df_sht.drop(columns=["ColorGroup","EsTitulo","GroupCount","MultiSort","NotTitulo","GroupID"], errors="ignore")
                        temp.to_excel(writer, sheet_name=sht, index=False)
                with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
                    for sht, df_sht in data_dict.items():
                        temp = df_sht.drop(columns=["ColorGroup","EsTitulo","GroupCount","MultiSort","NotTitulo","GroupID"], errors="ignore")
                        temp.to_excel(writer, sheet_name=sht, index=False)
                st.success("Recepción del lote actualizada correctamente.")
                st.experimental_rerun()
        else:
            st.warning("No se encontró un lote con ese título en la hoja actual.")

# -------------------------------------------------------------------------
# SECCIÓN: Edición Individual y Guardado
# -------------------------------------------------------------------------
st.title("Control de Stock: Edición Individual")
st.markdown("---")
st.header("Edición en Hoja Principal y Guardado")
df_main_original = data_dict[sheet_name].copy()
df_main_original = enforce_types(df_main_original)
df_for_style = df_main_original.copy()
df_for_style["Alarma"] = df_for_style.apply(calc_alarma, axis=1)
df_for_style = build_group_info_by_ref(df_for_style, panel_default=sheet_name)
df_for_style.sort_values(by=["MultiSort", "GroupID", "NotTitulo"], inplace=True)
df_for_style.reset_index(drop=True, inplace=True)
styled_df = df_for_style.style.apply(style_lote, axis=1)
all_cols = df_for_style.columns.tolist()
cols_to_hide = ["ColorGroup", "EsTitulo", "GroupCount", "MultiSort", "NotTitulo", "GroupID"]
final_cols = [c for c in all_cols if c not in cols_to_hide]
table_html = styled_df.to_html(columns=final_cols)
df_main = df_for_style.copy()
df_main.drop(columns=cols_to_hide, inplace=True, errors="ignore")
st.write("#### Vista de la Hoja (con columna 'Alarma' y sin columnas internas)")
st.write(table_html, unsafe_allow_html=True)

# NUEVA SECCIÓN: Botón para resetear la fila actual (Revertir cambios)
if st.button("Resetear fila (Revertir cambios)"):
    # Se lee la hoja original desde ORIGINAL_FILE para obtener los valores originales
    original_df = pd.read_excel(ORIGINAL_FILE, sheet_name=sheet_name, engine="openpyxl")
    original_df = enforce_types(original_df)
    current_row = df_main.loc[row_index]
    # Usamos "Ref. Saturno" y "Nombre producto" (en minúsculas) para identificar la fila
    key = (current_row["Ref. Saturno"], current_row["Nombre producto"].strip().lower())
    original_match = original_df[(original_df["Ref. Saturno"] == key[0]) & (original_df["Nombre producto"].str.lower() == key[1])]
    if not original_match.empty:
        original_values = original_match.iloc[0]
        for col in df_main.columns:
            data_dict[sheet_name].at[row_index, col] = original_values[col]
        st.success("Fila reseteada a los valores originales.")
        st.experimental_rerun()
    else:
        st.error("No se encontró la fila original para resetear.")

if "Nombre producto" in df_main.columns and "Ref. Fisher" in df_main.columns:
    display_series = df_main.apply(lambda r: f"{r['Nombre producto']} ({r['Ref. Fisher']})", axis=1)
else:
    display_series = df_main.iloc[:, 0].astype(str)
reactivo_sel = st.selectbox("Selecciona Reactivo a Modificar:", display_series.unique(), key="react_modif")
row_index = display_series[display_series == reactivo_sel].index[0]

def get_val(col, default=None):
    return df_main.at[row_index, col] if col in df_main.columns else default

lote_actual = get_val("NºLote", 0)
caducidad_actual = get_val("Caducidad", None)
fecha_pedida_actual = get_val("Fecha Pedida", None)
fecha_llegada_actual = get_val("Fecha Llegada", None)
sitio_almacenaje_actual = get_val("Sitio almacenaje", "")
uds_actual = get_val("Uds.", 0)
stock_actual = get_val("Stock", 0)

colA, colB, colC, colD = st.columns([1,1,1,1])
with colA:
    lote_nuevo = st.number_input("Nº de Lote", value=int(lote_actual), step=1)
    caducidad_nueva = st.date_input("Caducidad", value=caducidad_actual if pd.notna(caducidad_actual) else None)
with colB:
    fp_date = st.date_input("Fecha Pedida (fecha)",
                            value=fecha_pedida_actual.date() if pd.notna(fecha_pedida_actual) else None,
                            key="fp_date_main")
    fp_time = st.time_input("Hora Pedida",
                            value=fecha_pedida_actual.time() if pd.notna(fecha_pedida_actual) else datetime.time(0,0),
                            key="fp_time_main")
with colC:
    fl_date = st.date_input("Fecha Llegada (fecha)",
                            value=fecha_llegada_actual.date() if pd.notna(fecha_llegada_actual) else None,
                            key="fl_date_main")
    fl_time = st.time_input("Hora Llegada",
                            value=fecha_llegada_actual.time() if pd.notna(fecha_llegada_actual) else datetime.time(0,0),
                            key="fl_time_main")
with colD:
    st.write("")
    st.write("")
    if st.button("Refrescar Página"):
        st.rerun()
fecha_pedida_nueva = None
if fp_date is not None:
    dt_ped = datetime.datetime.combine(fp_date, fp_time)
    fecha_pedida_nueva = pd.to_datetime(dt_ped)
fecha_llegada_nueva = None
if fl_date is not None:
    dt_lleg = datetime.datetime.combine(fl_date, fl_time)
    fecha_llegada_nueva = pd.to_datetime(dt_lleg)
st.write("Sitio de Almacenaje")
opciones_sitio = ["Congelador 1", "Congelador 2", "Frigorífico", "Tª Ambiente"]
sitio_principal = sitio_almacenaje_actual.split(" - ")[0] if " - " in sitio_almacenaje_actual else sitio_almacenaje_actual
if sitio_principal not in opciones_sitio:
    sitio_principal = opciones_sitio[0]
sitio_top = st.selectbox("Tipo Almacenaje", opciones_sitio, index=opciones_sitio.index(sitio_principal))
subopcion = ""
if sitio_top == "Congelador 1":
    cajones = [f"Cajón {i}" for i in range(1,9)]
    subopcion = st.selectbox("Cajón (1 Arriba,8 Abajo)", cajones)
elif sitio_top == "Congelador 2":
    cajones = [f"Cajón {i}" for i in range(1,7)]
    subopcion = st.selectbox("Cajón (1 Arriba,6 Abajo)", cajones)
elif sitio_top == "Frigorífico":
    baldas = [f"Balda {i}" for i in range(1,8)] + ["Puerta"]
    subopcion = st.selectbox("Baldas (1 Arriba, 7 Abajo)", baldas)
elif sitio_top == "Tª Ambiente":
    comentario = st.text_input("Comentario (opcional)")
    subopcion = comentario.strip()
if subopcion:
    sitio_almacenaje_nuevo = f"{sitio_top} - {subopcion}"
else:
    sitio_almacenaje_nuevo = sitio_top

group_order_selected = None
if pd.notna(fecha_pedida_nueva):
    group_id = df_for_style.at[row_index, "GroupID"]
    group_reactivos = df_for_style[df_for_style["GroupID"] == group_id]
    if not group_reactivos.empty:
        if group_reactivos["EsTitulo"].any():
            lot_name = group_reactivos[group_reactivos["EsTitulo"]==True]["Nombre producto"].iloc[0]
        else:
            lot_name = f"Ref. Saturno {group_id}"
        group_reactivos_reset = group_reactivos.reset_index()
        options = group_reactivos_reset.apply(lambda r: f"{r['index']} - {r['Nombre producto']} ({r['Ref. Fisher']})", axis=1).tolist()
        st.markdown('<div class="big-select">', unsafe_allow_html=True)
        group_order_selected = st.multiselect(f"¿Quieres pedir también los siguientes reactivos del lote **{lot_name}**?", options, default=options)
        st.markdown('</div>', unsafe_allow_html=True)

if st.button("Guardar Cambios"):
    if pd.notna(fecha_llegada_nueva):
        fecha_pedida_nueva = pd.NaT
    if "Stock" in df_main.columns:
        if fecha_llegada_nueva != fecha_llegada_actual and pd.notna(fecha_llegada_nueva):
            df_main.at[row_index, "Stock"] = stock_actual + uds_actual
            st.info(f"Sumadas {uds_actual} uds al stock => {stock_actual + uds_actual}")
    if "NºLote" in df_main.columns:
        df_main.at[row_index, "NºLote"] = int(lote_nuevo)
    if "Caducidad" in df_main.columns:
        if pd.notna(caducidad_nueva):
            df_main.at[row_index, "Caducidad"] = pd.to_datetime(caducidad_nueva)
        else:
            df_main.at[row_index, "Caducidad"] = pd.NaT
    if "Fecha Pedida" in df_main.columns:
        if pd.notna(fecha_pedida_nueva):
            df_main.at[row_index, "Fecha Pedida"] = pd.to_datetime(fecha_pedida_nueva)
        else:
            df_main.at[row_index, "Fecha Pedida"] = pd.NaT
    if "Fecha Llegada" in df_main.columns:
        if pd.notna(fecha_llegada_nueva):
            df_main.at[row_index, "Fecha Llegada"] = pd.to_datetime(fecha_llegada_nueva)
        else:
            df_main.at[row_index, "Fecha Llegada"] = pd.NaT
    if "Sitio almacenaje" in df_main.columns:
        df_main.at[row_index, "Sitio almacenaje"] = sitio_almacenaje_nuevo
    if pd.notna(fecha_pedida_nueva) and group_order_selected:
        for label in group_order_selected:
            try:
                i_val = int(label.split(" - ")[0])
                df_main.at[i_val, "Fecha Pedida"] = fecha_pedida_nueva
            except Exception as e:
                st.error(f"Error actualizando índice {label} (Fecha Pedida): {e}")
    data_dict[sheet_name] = df_main
    new_file = crear_nueva_version_filename()
    with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
        for sht, df_sht in data_dict.items():
            cols_internos = ["ColorGroup","EsTitulo","GroupCount","MultiSort","NotTitulo","GroupID"]
            temp = df_sht.drop(columns=cols_internos, errors="ignore")
            temp.to_excel(writer, sheet_name=sht, index=False)
    with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
        for sht, df_sht in data_dict.items():
            cols_internos = ["ColorGroup","EsTitulo","GroupCount","MultiSort","NotTitulo","GroupID"]
            temp = df_sht.drop(columns=cols_internos, errors="ignore")
            temp.to_excel(writer, sheet_name=sht, index=False)
    st.success(f"✅ Cambios guardados en '{new_file}' y '{STOCK_FILE}'.")
    excel_bytes = generar_excel_en_memoria(df_main, sheet_nm=sheet_name)
    st.download_button(
        label="Descargar Excel modificado",
        data=excel_bytes,
        file_name="Reporte_Stock.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.rerun()
