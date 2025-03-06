import streamlit as st
import pandas as pd
import numpy as np
import datetime
import shutil
import os
from io import BytesIO

# -------------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA
# -------------------------------------------------------------------------
st.set_page_config(page_title="Control de Stock con Lotes", layout="centered")

STOCK_FILE = "Stock_Original.xlsx"  # Archivo principal de trabajo
VERSIONS_DIR = "versions"
ORIGINAL_FILE = os.path.join(VERSIONS_DIR, "Stock_Original.xlsx")

os.makedirs(VERSIONS_DIR, exist_ok=True)

def init_original():
    """Si no existe 'versions/Stock_Original.xlsx', lo creamos a partir de 'Stock_Original.xlsx'."""
    if not os.path.exists(ORIGINAL_FILE):
        if os.path.exists(STOCK_FILE):
            shutil.copy(STOCK_FILE, ORIGINAL_FILE)
            print("Creada versión original en:", ORIGINAL_FILE)
        else:
            st.error(f"No se encontró {STOCK_FILE}. Asegúrate de subirlo.")

init_original()

def load_data():
    """Carga todas las hojas de STOCK_FILE en un dict {nombre_hoja: DataFrame}."""
    try:
        return pd.read_excel(STOCK_FILE, sheet_name=None, engine="openpyxl")
    except FileNotFoundError:
        st.error(f"❌ No se encontró {STOCK_FILE}.")
        return None
    except Exception as e:
        st.error(f"❌ Error al cargar la base de datos: {e}")
        return None

data_dict = load_data()

# -------------------------------------------------------------------------
# FUNCIÓN PARA CONVERSIONES DE TIPOS
# -------------------------------------------------------------------------
def enforce_types(df: pd.DataFrame):
    """Aplica los tipos correctos a las columnas."""
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
    """Generar un Excel en memoria (bytes) para descargar."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_act.to_excel(writer, index=False, sheet_name=sheet_nm)
    output.seek(0)
    return output.getvalue()

# -------------------------------------------------------------------------
# ESTRUCTURA DE LOTES (diccionario). Solo se usará para colorear / agrupar
# -------------------------------------------------------------------------
LOTS_DATA = {
    "FOCUS": {
        "Panel Oncomine Focus Library Assay Chef Ready": [
            "Primers DNA",
            "Primers RNA",
            "Reagents DL8",
            "Chef supplies (plásticos)",
            "Placas",
            "Solutions DL8"
        ],
        "Ion 510/520/530 kit-Chef (TEMPLADO)": [
            "Chef Reagents",
            "Chef Solutions",
            "Chef supplies (plásticos)",
            "Solutions Reagent S5",
            "Botellas S5"
        ],
        "Recover All TM Multi-Sample RNA/DNA Isolation workflow-Kit": [
            "Kit extracción DNA/RNA",
            "RecoverAll TM kit (Dnase, protease,…)",
            "H2O RNA free",
            "Tubos fondo cónico",
            "Superscript VILO cDNA Syntheis Kit",
            "Qubit 1x dsDNA HS Assay kit (100 reactions)"
        ]
    },
    "OCA": {
        "Panel OCA Library Assay Chef Ready": [
            "Primers DNA",
            "Primers RNA",
            "Reagents DL8",
            "Chef supplies (plásticos)",
            "Placas",
            "Solutions DL8"
        ],
        "kit-Chef (TEMPLADO)": [
            "Ion 540 TM Chef Reagents",
            "Chef Solutions",
            "Chef supplies (plásticos)",
            "Solutions Reagent S5",
            "Botellas S5"
        ],
        "Chip secuenciación liberación de protones 6 millones de lecturas": [
            "Ion 540 TM Chip Kit"
        ],
        "Recover All TM Multi-Sample RNA/DNA Isolation workflow-Kit": [
            "Kit extracción DNA/RNA",
            "RecoverAll TM kit (Dnase, protease,…)",
            "H2O RNA free",
            "Tubos fondo cónico"
        ]
    },
    "OCA PLUS": {
        "Panel OCA-PLUS Library Assay Chef Ready": [
            "Primers DNA",
            "Uracil-DNA Glycosylase heat-labile",
            "Reagents DL8",
            "Chef supplies (plásticos)",
            "Placas",
            "Solutions DL8"
        ],
        "kit-Chef (TEMPLADO)": [
            "Ion 550 TM Chef Reagents",
            "Chef Solutions",
            "Chef Supplies (plásticos)",
            "Solutions Reagent S5",
            "Botellas S5",
            "Chip secuenciación Ion 550 TM Chip Kit"
        ],
        "Recover All TM Multi-Sample RNA/DNA Isolation workflow-Kit": [
            "Kit extracción DNA/RNA",
            "RecoverAll TM kit (Dnase, protease,…)",
            "H2O RNA free",
            "Tubos fondo cónico"
        ]
    }
}

# -------------------------------------------------------------------------
# MAPEO SUBLOTE -> COLOR
# -------------------------------------------------------------------------
import itertools
color_cycle = itertools.cycle([
    "#FFC0CB", "#B0E0E6", "#FFB6C1", "#E6E6FA", "#FFE4E1",
    "#F0E68C", "#F5DEB3", "#D3D3D3", "#C8A2C8", "#90EE90"
])
# Creamos un dict: { (panel, sublote_name): color_str }
sub_lot_to_color = {}
for panel, subdict in LOTS_DATA.items():
    for sublot_name in subdict.keys():
        c = next(color_cycle)
        sub_lot_to_color[(panel, sublot_name)] = c


def find_sub_lot(nombre_prod: str):
    """
    Devuelve (panel, sublote) si:
    - nombre_prod coincide EXACTAMENTE con el sublote
      * Ej: "Panel Oncomine Focus Library Assay Chef Ready"
    - nombre_prod está en la lista de reactivos del sublote
      * Ej: "Primers DNA"
    En caso de no encontrar, retorna None.
    """
    for panel_name, subdict in LOTS_DATA.items():
        for sublot_name, react_list in subdict.items():
            # Si el nombre_prod coincide con el sublot_name => "fila de cabecera" del lote
            if nombre_prod == sublot_name:
                return (panel_name, sublot_name, True)  # True => es sublote principal
            # O si nombre_prod está en la lista => reactivo
            if nombre_prod in react_list:
                return (panel_name, sublot_name, False)
    return None  # no encontrado

def style_lotes(row):
    """Devuelve un dict de estilos para cada celda de la fila en la visualización."""
    nombre_prod = row.get("Nombre producto", "")
    info = find_sub_lot(nombre_prod)
    if info:
        panel_name, sublot_name, es_sublote = info
        color = sub_lot_to_color[(panel_name, sublot_name)]
        styles = [f"background-color: {color}"] * len(row)
        # Si es sublote principal => marcamos la celda "Nombre producto" en negrita
        if es_sublote:
            # localizamos índice de la columna "Nombre producto" para poner bold
            idx_col = row.index.get_loc("Nombre producto") if "Nombre producto" in row.index else None
            if idx_col is not None:
                styles[idx_col] += "; font-weight: bold"
        return styles
    else:
        return ["" for _ in row]

# -------------------------------------------------------------------------
# BARRA LATERAL: Expander con 3 apartados
# -------------------------------------------------------------------------
with st.sidebar:
    # 1) Ver / Gestionar versiones guardadas
    with st.expander("🔎 Ver / Gestionar versiones guardadas", expanded=False):
        if data_dict:
            files = sorted(os.listdir(VERSIONS_DIR))
            versions_no_original = [f for f in files if f != "Stock_Original.xlsx"]
            if versions_no_original:
                version_sel = st.selectbox("Selecciona una versión:", versions_no_original)
                confirm_delete = False

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
                    if st.checkbox(f"Confirmar eliminación de '{version_sel}'"):
                        confirm_delete = True

                    if st.button("Eliminar esta versión"):
                        if confirm_delete:
                            try:
                                os.remove(file_path)
                                st.warning(f"Versión '{version_sel}' eliminada.")
                                st.rerun()
                            except:
                                st.error("Error al intentar eliminar la versión.")
                        else:
                            st.error("Marca la casilla de confirmación para eliminar la versión.")
            else:
                st.write("No hay versiones guardadas (excepto la original).")

            if st.button("Eliminar TODAS las versiones (excepto original)"):
                for f in versions_no_original:
                    try:
                        os.remove(os.path.join(VERSIONS_DIR, f))
                    except:
                        pass
                st.info("Todas las versiones (excepto la original) han sido eliminadas.")
                st.rerun()

            if st.button("Eliminar TODAS las versiones excepto la última y la original"):
                if len(versions_no_original) > 1:
                    sorted_vers = sorted(versions_no_original)
                    last_version = sorted_vers[-1]  # la última alfabéticamente
                    for f in versions_no_original:
                        if f != last_version:
                            try:
                                os.remove(os.path.join(VERSIONS_DIR, f))
                            except:
                                pass
                    st.info(f"Se han eliminado todas las versiones excepto: {last_version} y Stock_Original.xlsx")
                    st.rerun()
                else:
                    st.write("Solo hay una versión o ninguna versión, no se elimina nada más.")

            if st.button("Limpiar Base de Datos"):
                st.write("¿Seguro que quieres limpiar la base de datos?")
                if st.checkbox("Sí, confirmar limpieza."):
                    original_path = os.path.join(VERSIONS_DIR, "Stock_Original.xlsx")
                    if os.path.exists(original_path):
                        shutil.copy(original_path, STOCK_FILE)
                        st.success("✅ Base de datos restaurada al estado original.")
                        st.rerun()
                    else:
                        st.error("❌ No se encontró la copia original en 'versions/Stock_Original.xlsx'.")
        else:
            st.error("No hay data_dict. Asegúrate de que existe Stock_Original.xlsx.")
            st.stop()

    # 2) Alarmas
    with st.expander("⚠️ Alarmas", expanded=False):
        st.write("MOSTRAR ALARMAS como icono en la tabla (columna 'Alarma'):")
        st.write("Se pintará 🔴 si Stock=0 y Fecha Pedida=None, 🟠 si Stock=0 y Fecha Pedida!=None.")
        # (Explicado, no se colorean las filas)
        # No mostramos nada más aquí, ya que no coloreamos filas. 
        # La lógica de la columna 'Alarma' la haremos en la tabla principal.

    # 3) Reactivo Agotado
    with st.expander("Reactivo Agotado (Consumido en Lab)", expanded=False):
        st.write("Selecciona la hoja y el reactivo, y cuántas unidades restar del stock sin necesidad de guardar cambios.")
        if data_dict:
            hojas_agotado = list(data_dict.keys())
            hoja_sel_consumo = st.selectbox("Hoja para consumir reactivo:", hojas_agotado, key="consumo_hoja_sel")
            df_agotado = data_dict[hoja_sel_consumo].copy()
            df_agotado = enforce_types(df_agotado)

            if "Nombre producto" in df_agotado.columns and "Ref. Fisher" in df_agotado.columns:
                disp_series_consumo = df_agotado.apply(lambda r: f"{r['Nombre producto']} ({r['Ref. Fisher']})", axis=1)
            else:
                disp_series_consumo = df_agotado.iloc[:, 0].astype(str)

            reactivo_consumir = st.selectbox("Reactivo a consumir:", disp_series_consumo.unique(), key="select_reactivo_cons")
            idx_cons = disp_series_consumo[disp_series_consumo == reactivo_consumir].index[0]
            stock_cons_actual = df_agotado.at[idx_cons, "Stock"] if "Stock" in df_agotado.columns else 0

            uds_consumidas = st.number_input("Uds. consumidas", min_value=0, step=1, key="uds_cons_laboratorio")

            if st.button("Registrar Consumo en Lab"):
                nuevo_stock = max(0, stock_cons_actual - uds_consumidas)
                df_agotado.at[idx_cons, "Stock"] = nuevo_stock
                st.warning(f"Se han consumido {uds_consumidas} uds. Stock final => {nuevo_stock}")

                # Actualizamos en data_dict sin guardar en Excel
                data_dict[hoja_sel_consumo] = df_agotado
                st.success("No se ha creado versión nueva. Los datos se mantienen en memoria hasta 'Guardar Cambios'.")

        else:
            st.error("No hay data_dict. Asegúrate de que existe Stock_Original.xlsx.")
            st.stop()

# -------------------------------------------------------------------------
#  CUERPO PRINCIPAL
# -------------------------------------------------------------------------
st.title("📦 Control de Stock: Alarmas + Lotes agrupados + Reactivo Agotado")

if not data_dict:
    st.error("No se pudo cargar la base de datos. Verifica Stock_Original.xlsx")
    st.stop()

st.markdown("---")
st.header("Edición en Hoja Principal y Guardado")

hojas_principales = list(data_dict.keys())
sheet_name = st.selectbox("Selecciona la hoja a editar:", hojas_principales, key="sheet_principal_sel")
df_main = data_dict[sheet_name].copy()
df_main = enforce_types(df_main)

# 1) Añadimos una columna "Alarma" con un icono: 🔴 o 🟠
def calcular_alarma(row):
    stock_val = row.get("Stock", None)
    fecha_ped = row.get("Fecha Pedida", None)
    if pd.isna(stock_val) or stock_val != 0:
        return ""
    # stock=0
    if pd.isna(fecha_ped):
        return "🔴"  # sin pedir => ROJA
    else:
        return "🟠"  # pedida => NARANJA

if "Alarma" not in df_main.columns:
    # Creamos la col "Alarma"
    df_main["Alarma"] = df_main.apply(calcular_alarma, axis=1)
else:
    # Si ya existe, la recalculamos
    df_main["Alarma"] = df_main.apply(calcular_alarma, axis=1)

# 2) Generamos el estilo de sub-lotes (colores)
def style_sub_lotes(row):
    """Devuelve lista de estilos CSS, coloreando según sub-lote y marcando en negrita si es el sub-lote principal."""
    nombre_prod = row.get("Nombre producto", "")
    info = find_sub_lot(nombre_prod)
    if info:
        panel_name, sublot_name, es_sublote = info
        color = sub_lot_to_color[(panel_name, sublot_name)]
        styles = [f"background-color: {color}"] * len(row)
        if es_sublote:
            idx_col = row.index.get_loc("Nombre producto") if "Nombre producto" in row.index else None
            if idx_col is not None:
                styles[idx_col] += "; font-weight: bold"
        return styles
    else:
        return ["" for _ in row]

def find_sub_lot(nombre_prod: str):
    """Retorna (panel, sublot, es_principal) si 'nombre_prod' coincide con el sub-lote o con uno de sus reactivos."""
    for panel, subdict in LOTS_DATA.items():
        for sublot_name, reactives in subdict.items():
            if nombre_prod == sublot_name:
                return (panel, sublot_name, True)
            if nombre_prod in reactives:
                return (panel, sublot_name, False)
    return None

# 3) Renderizamos la tabla con st.dataframe (o st.write) y estilo
styled_df = df_main.style.apply(style_sub_lotes, axis=1)
st.write("#### Vista de la Hoja (con columna 'Alarma' y color por Lote)")
st.write(styled_df.to_html(), unsafe_allow_html=True)

# Seleccionar reactivo a modificar
if "Nombre producto" in df_main.columns and "Ref. Fisher" in df_main.columns:
    display_series = df_main.apply(lambda row: f"{row['Nombre producto']} ({row['Ref. Fisher']})", axis=1)
else:
    display_series = df_main.iloc[:, 0].astype(str)

reactivo = st.selectbox("Selecciona Reactivo a Modificar:", display_series.unique(), key="reactivo_modif")
row_index = display_series[display_series == reactivo].index[0]

def get_val(col, default=None):
    return df_main.at[row_index, col] if col in df_main.columns else default

# Cargamos valores
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
    fecha_pedida_date = st.date_input("Fecha Pedida (fecha)",
                                      value=fecha_pedida_actual.date() if pd.notna(fecha_pedida_actual) else None,
                                      key="fp_date_main")
    fecha_pedida_time = st.time_input("Hora Pedida",
                                      value=fecha_pedida_actual.time() if pd.notna(fecha_pedida_actual) else datetime.time(0, 0),
                                      key="fp_time_main")

with colC:
    fecha_llegada_date = st.date_input("Fecha Llegada (fecha)",
                                       value=fecha_llegada_actual.date() if pd.notna(fecha_llegada_actual) else None,
                                       key="fl_date_main")
    fecha_llegada_time = st.time_input("Hora Llegada",
                                       value=fecha_llegada_actual.time() if pd.notna(fecha_llegada_actual) else datetime.time(0, 0),
                                       key="fl_time_main")

with colD:
    st.write("")
    st.write("")
    if st.button("Refrescar Página"):
        st.rerun()

fecha_pedida_nueva = None
if fecha_pedida_date is not None:
    fecha_pedida_nueva = datetime.datetime.combine(fecha_pedida_date, fecha_pedida_time)

fecha_llegada_nueva = None
if fecha_llegada_date is not None:
    fecha_llegada_nueva = datetime.datetime.combine(fecha_llegada_date, fecha_llegada_time)

st.write("Sitio de Almacenaje")
opciones_sitio = ["Congelador 1", "Congelador 2", "Frigorífico", "Tª Ambiente"]
sitio_principal = sitio_almacenaje_actual.split(" - ")[0] if " - " in sitio_almacenaje_actual else sitio_almacenaje_actual
if sitio_principal not in opciones_sitio:
    sitio_principal = opciones_sitio[0]
sitio_top = st.selectbox("Tipo Almacenaje", opciones_sitio, index=opciones_sitio.index(sitio_principal))

subopcion = ""
if sitio_top == "Congelador 1":
    cajones = [f"Cajón {i}" for i in range(1, 9)]
    subopcion = st.selectbox("Cajón (1 Arriba, 8 Abajo)", cajones)
elif sitio_top == "Congelador 2":
    cajones = [f"Cajón {i}" for i in range(1, 7)]
    subopcion = st.selectbox("Cajón (1 Arriba, 6 Abajo)", cajones)
elif sitio_top == "Frigorífico":
    baldas = [f"Balda {i}" for i in range(1, 8)] + ["Puerta"]
    subopcion = st.selectbox("Baldas (1 Arriba, 7 Abajo)", baldas)
elif sitio_top == "Tª Ambiente":
    comentario = st.text_input("Comentario (opcional)")
    subopcion = comentario.strip()

if subopcion:
    sitio_almacenaje_nuevo = f"{sitio_top} - {subopcion}"
else:
    sitio_almacenaje_nuevo = sitio_top

# Botón Guardar Cambios
if st.button("Guardar Cambios"):
    # Si se introduce Fecha Llegada, borramos Fecha Pedida => "ya llegó"
    if pd.notna(fecha_llegada_nueva):
        fecha_pedida_nueva = pd.NaT

    # Sumar Stock si la fecha de llegada cambió
    if "Stock" in df_main.columns:
        if fecha_llegada_nueva != fecha_llegada_actual and pd.notna(fecha_llegada_nueva):
            df_main.at[row_index, "Stock"] = stock_actual + uds_actual
            st.info(f"Sumadas {uds_actual} uds al stock. Nuevo stock => {stock_actual + uds_actual}")

    # Guardar ediciones en la fila
    if "NºLote" in df_main.columns:
        df_main.at[row_index, "NºLote"] = int(lote_nuevo)
    if "Caducidad" in df_main.columns:
        df_main.at[row_index, "Caducidad"] = caducidad_nueva
    if "Fecha Pedida" in df_main.columns:
        df_main.at[row_index, "Fecha Pedida"] = fecha_pedida_nueva
    if "Fecha Llegada" in df_main.columns:
        df_main.at[row_index, "Fecha Llegada"] = fecha_llegada_nueva
    if "Sitio almacenaje" in df_main.columns:
        df_main.at[row_index, "Sitio almacenaje"] = sitio_almacenaje_nuevo

    data_dict[sheet_name] = df_main

    # Creamos la versión
    new_file = crear_nueva_version_filename()
    with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
        for sht, df_sheet in data_dict.items():
            df_sheet.to_excel(writer, sheet_name=sht, index=False)

    # Guardamos en STOCK_FILE
    with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
        for sht, df_sheet in data_dict.items():
            df_sheet.to_excel(writer, sheet_name=sht, index=False)

    st.success(f"✅ Cambios guardados en '{new_file}' y '{STOCK_FILE}'.")

    excel_bytes = generar_excel_en_memoria(df_main, sheet_nm=sheet_name)
    st.download_button(
        label="Descargar Excel modificado",
        data=excel_bytes,
        file_name="Reporte_Stock.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.rerun()
