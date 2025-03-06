import streamlit as st
import pandas as pd
import numpy as np
import datetime
import shutil
import os
from io import BytesIO

st.set_page_config(page_title="Control de Stock con Lotes", layout="centered")

STOCK_FILE = "Stock_Original.xlsx"
VERSIONS_DIR = "versions"
ORIGINAL_FILE = os.path.join(VERSIONS_DIR, "Stock_Original.xlsx")

os.makedirs(VERSIONS_DIR, exist_ok=True)

def init_original():
    """Si no existe 'versions/Stock_Original.xlsx', lo creamos a partir de 'Stock_Original.xlsx'."""
    if not os.path.exists(ORIGINAL_FILE):
        if os.path.exists(STOCK_FILE):
            shutil.copy(STOCK_FILE, ORIGINAL_FILE)
        else:
            st.error(f"No se encontró {STOCK_FILE}. Asegúrate de subirlo.")

init_original()

def load_data():
    """Carga todas las hojas de STOCK_FILE en un dict {nombre_hoja: DataFrame}."""
    try:
        return pd.read_excel(STOCK_FILE, sheet_name=None, engine="openpyxl")
    except FileNotFoundError:
        st.error("❌ No se encontró el archivo principal.")
        return None
    except Exception as e:
        st.error(f"❌ Error al cargar la base de datos: {e}")
        return None

data_dict = load_data()

def enforce_types(df: pd.DataFrame):
    """Forzamos tipos en las columnas."""
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
    """Generamos un Excel en memoria para descargar."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_act.to_excel(writer, index=False, sheet_name=sheet_nm)
    output.seek(0)
    return output.getvalue()

# Diccionario de paneles, sub-lotes y reactivos
LOTS_DATA = {
    "FOCUS": {
        "Panel Oncomine Focus Library Assay Chef Ready": [
            "Primers DNA","Primers RNA","Reagents DL8","Chef supplies (plásticos)","Placas","Solutions DL8"
        ],
        "Ion 510/520/530 kit-Chef (TEMPLADO)": [
            "Chef Reagents","Chef Solutions","Chef supplies (plásticos)","Solutions Reagent S5","Botellas S5"
        ],
        "Recover All TM Multi-Sample RNA/DNA Isolation workflow-Kit": [
            "Kit extracción DNA/RNA","RecoverAll TM kit (Dnase, protease,…)","H2O RNA free",
            "Tubos fondo cónico","Superscript VILO cDNA Syntheis Kit","Qubit 1x dsDNA HS Assay kit (100 reactions)"
        ],
    },
    "OCA": {
        "Panel OCA Library Assay Chef Ready": [
            "Primers DNA","Primers RNA","Reagents DL8","Chef supplies (plásticos)","Placas","Solutions DL8"
        ],
        "kit-Chef (TEMPLADO)": [
            "Ion 540 TM Chef Reagents","Chef Solutions","Chef supplies (plásticos)","Solutions Reagent S5","Botellas S5"
        ],
        "Chip secuenciación liberación de protones 6 millones de lecturas": [
            "Ion 540 TM Chip Kit"
        ],
        "Recover All TM Multi-Sample RNA/DNA Isolation workflow-Kit": [
            "Kit extracción DNA/RNA","RecoverAll TM kit (Dnase, protease,…)","H2O RNA free","Tubos fondo cónico"
        ]
    },
    "OCA PLUS": {
        "Panel OCA-PLUS Library Assay Chef Ready": [
            "Primers DNA","Uracil-DNA Glycosylase heat-labile","Reagents DL8","Chef supplies (plásticos)","Placas","Solutions DL8"
        ],
        "kit-Chef (TEMPLADO)": [
            "Ion 550 TM Chef Reagents","Chef Solutions","Chef Supplies (plásticos)",
            "Solutions Reagent S5","Botellas S5","Chip secuenciación Ion 550 TM Chip Kit"
        ],
        "Recover All TM Multi-Sample RNA/DNA Isolation workflow-Kit": [
            "Kit extracción DNA/RNA","RecoverAll TM kit (Dnase, protease,…)","H2O RNA free","Tubos fondo cónico"
        ]
    }
}

### Paso 1: Definir un orden para los paneles y sublotes
panel_order = ["FOCUS", "OCA", "OCA PLUS"]
# Asignamos manualmente un índice a cada panel
panel_index = { panel_name: i for i, panel_name in enumerate(panel_order) }

# Asignamos, para cada panel, un dict sublote->(orden, color)
# Ejemplo de colores asignados manualmente
import itertools
colors = ["#FED7D7","#FEE2E2","#FFEDD5","#FEF9C3",
          "#D9F99D","#CFFAFE","#E0E7FF","#FBCFE8","#F9A8D4","#E9D5FF",
          "#FFD700","#F0FFF0","#D1FAE5","#BAFEE2","#A7F3D0"]
color_cycle = itertools.cycle(colors)

# Construimos una estructura con (panel, sub-lote) -> (panelIdx, subloteIdx, color)
sub_lot_metadata = {}
for p in panel_order:
    subdict = LOTS_DATA.get(p,{})
    sublots = list(subdict.keys())
    for j, subl in enumerate(sublots):
        sub_lot_metadata[(p, subl)] = (panel_index[p], j, next(color_cycle))

### Función para encontrar sub-lote (y panel) de una fila
def find_sub_lot(nombre_prod: str):
    """
    Retorna (panel_name, sublote_name, esPrincipal)
    - esPrincipal=True si 'nombre_prod' coincide con la cabecera (ej: "Panel OCA Library...")
    - esPrincipal=False si 'nombre_prod' coincide con uno de los reactivos
    Devuelve None si no está en la estructura.
    """
    for p in panel_order:
        subdict = LOTS_DATA[p]
        for sublot_name, reactives in subdict.items():
            if nombre_prod == sublot_name:
                return (p, sublot_name, True)
            if nombre_prod in reactives:
                return (p, sublot_name, False)
    return None

### Asignar metadatos de panel/sublote a cada fila
def categorize_rows(df: pd.DataFrame):
    """Agrega 4 columnas: 'PanelIdx','SubLoteIdx','EsPrincipal','Color' para ordenar y colorear."""
    df = df.copy()
    df["PanelIdx"] = 999  # default
    df["SubLoteIdx"] = 999
    df["EsPrincipal"] = False
    df["Color"] = ""

    for i, row in df.iterrows():
        nombre_prod = row.get("Nombre producto","")
        info = find_sub_lot(nombre_prod)
        if info:
            p, sl, esP = info  # panel, sublote, esPrincipal
            (p_idx, s_idx, color) = sub_lot_metadata[(p, sl)]
            df.at[i, "PanelIdx"] = p_idx
            df.at[i, "SubLoteIdx"] = s_idx
            df.at[i, "EsPrincipal"] = esP
            df.at[i, "Color"] = color
        else:
            # no coincide con ningún sub-lote => panel=999, sublote=999, color=''
            pass

    return df

### Creamos la columna Alarma con íconos
def calcular_alarma(row):
    stock_val = row.get("Stock", None)
    fecha_ped = row.get("Fecha Pedida", None)
    if pd.isna(stock_val) or stock_val != 0:
        return ""
    # stock=0
    if pd.isna(fecha_ped):
        return "🔴"  # sin pedir
    else:
        return "🟨"  # pedido

### CSS styling para cada fila, en base a 'Color' y 'EsPrincipal'
def style_lotes_icons(row):
    bg_color = row.get("Color","")
    es_principal = row.get("EsPrincipal", False)
    styles = [f"background-color: {bg_color}"] * len(row)
    if es_principal:
        # localizamos índice 'Nombre producto'
        if "Nombre producto" in row.index:
            idx_np = row.index.get_loc("Nombre producto")
            styles[idx_np] += "; font-weight: bold"
    return styles

# -------------------------------------------------------------------------
# BARRA LATERAL
# -------------------------------------------------------------------------
with st.sidebar:
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
                    last_version = sorted_vers[-1]
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

    with st.expander("⚠️ Alarmas", expanded=False):
        st.write("Ahora no coloreamos filas por alarmas.")
        st.write("En la **columna 'Alarma'** verás:")
        st.write("- `'🔴'` (rojo) si `Stock=0` y `Fecha Pedida` es nula")
        st.write("- `'🟨'` (amarillo) si `Stock=0` y `Fecha Pedida` no es nula")

    with st.expander("Reactivo Agotado (Consumido en Lab)", expanded=False):
        if data_dict:
            st.write("Selecciona hoja y reactivo; descuenta stock sin crear versión.")
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
                data_dict[hoja_sel_consumo] = df_agotado
                st.success("No se ha creado versión nueva. Los datos se mantienen en memoria.")
        else:
            st.error("No hay data_dict. Asegúrate de que existe Stock_Original.xlsx.")
            st.stop()


# CUERPO PRINCIPAL
st.title("📦 Control de Stock: Lotes agrupados + Alarmas con íconos + Reactivo Agotado")

if not data_dict:
    st.error("No se pudo cargar la base de datos.")
    st.stop()

st.markdown("---")
st.header("Edición en Hoja Principal y Guardado")

hojas_principales = list(data_dict.keys())
sheet_name = st.selectbox("Selecciona la hoja a editar:", hojas_principales, key="sheet_principal_sel")
df_main = data_dict[sheet_name].copy()
df_main = enforce_types(df_main)

# (1) Creamos / recalculamos la columna "Alarma" con íconos
def calc_alarma_icon(row):
    stock_val = row.get("Stock",None)
    fecha_ped = row.get("Fecha Pedida",None)
    if (stock_val == 0) and pd.isna(fecha_ped):
        return "🔴"
    elif (stock_val == 0) and not pd.isna(fecha_ped):
        return "🟨"
    return ""

df_main["Alarma"] = df_main.apply(calc_alarma_icon, axis=1)

# (2) Añadimos columnas "PanelIdx","SubLoteIdx","EsPrincipal","Color" para ordenar y colorear
def build_lote_info(df):
    df = df.copy()
    df["PanelIdx"] = 999
    df["SubLoteIdx"] = 999
    df["EsPrincipal"] = False
    df["Color"] = ""
    for i, row in df.iterrows():
        name_prod = row.get("Nombre producto","")
        sub_info = find_sub_lot(name_prod)
        if sub_info:
            p, sl, is_main = sub_info
            (p_idx, s_idx, c) = sub_lot_metadata.get((p, sl),(999,999,""))
            df.at[i,"PanelIdx"] = p_idx
            df.at[i,"SubLoteIdx"] = s_idx
            df.at[i,"EsPrincipal"] = is_main
            df.at[i,"Color"] = c
    return df

def find_sub_lot(nombre_prod:str):
    """Devuelve (panelName, subLoteName, esPrincipal)."""
    for p in panel_order:  # ["FOCUS", "OCA", "OCA PLUS"]
        if p not in LOTS_DATA: continue
        subdict = LOTS_DATA[p]
        for sublot_name, reactivos in subdict.items():
            if nombre_prod == sublot_name:
                return (p, sublot_name, True)
            if nombre_prod in reactivos:
                return (p, sublot_name, False)
    return None

# sub_lot_metadata = dict((panel, sublote) -> (panel_idx, sublote_idx, color))
panel_order = ["FOCUS","OCA","OCA PLUS"]
panel_index = {p:i for i,p in enumerate(panel_order)}

import itertools
colors = ["#FED7D7","#FEE2E2","#FFEDD5","#FEF9C3",
          "#D9F99D","#CFFAFE","#E0E7FF","#FBCFE8","#F9A8D4","#E9D5FF",
          "#FFD700","#F0FFF0","#D1FAE5","#BAFEE2","#A7F3D0","#CCFF66"]
color_cycle = itertools.cycle(colors)

sub_lot_metadata = {}
for p in panel_order:
    subdict = LOTS_DATA.get(p,{})
    sublots = list(subdict.keys())
    for j, subl in enumerate(sublots):
        sub_lot_metadata[(p, subl)] = (panel_index[p], j, next(color_cycle))

df_main = build_lote_info(df_main)

# (3) Reordenamos el DataFrame por PanelIdx y SubLoteIdx, de modo que aparezca todo junto
df_main.sort_values(by=["PanelIdx","SubLoteIdx","EsPrincipal"],
                    ascending=[True,True,False], inplace=True)
df_main.reset_index(drop=True, inplace=True)

# (4) Creamos un estilo que use la columna "Color" y ponga en negrita la fila principal
def style_sub_lote(row):
    bg = row.get("Color","")
    is_main = row.get("EsPrincipal",False)
    # generamos un array de estilos
    styles = [f"background-color: {bg}"]* len(row)
    if is_main:
        # localizamos "Nombre producto" => negrita
        if "Nombre producto" in row.index:
            idx_np = row.index.get_loc("Nombre producto")
            styles[idx_np] += "; font-weight: bold"
    return styles

styled_df = df_main.style.apply(style_sub_lote, axis=1)

# (5) Mostramos la tabla
st.write("#### Vista de la Hoja (con columna 'Alarma' y color por Lote)")
st.write(styled_df.to_html(), unsafe_allow_html=True)

# Procedemos a la edición
if "Nombre producto" in df_main.columns and "Ref. Fisher" in df_main.columns:
    display_series = df_main.apply(lambda row: f"{row['Nombre producto']} ({row['Ref. Fisher']})", axis=1)
else:
    display_series = df_main.iloc[:, 0].astype(str)

reactivo = st.selectbox("Selecciona Reactivo a Modificar:", display_series.unique(), key="reactivo_modif")
row_index = display_series[display_series == reactivo].index[0]

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

if st.button("Guardar Cambios"):
    # Borrar fecha pedida si llega => stock sumado
    if pd.notna(fecha_llegada_nueva):
        fecha_pedida_nueva = pd.NaT

    if "Stock" in df_main.columns:
        if fecha_llegada_nueva != fecha_llegada_actual and pd.notna(fecha_llegada_nueva):
            df_main.at[row_index, "Stock"] = stock_actual + uds_actual
            st.info(f"Sumadas {uds_actual} uds al stock. Nuevo stock => {stock_actual + uds_actual}")

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

    new_file = crear_nueva_version_filename()
    with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
        for sht, df_sheet in data_dict.items():
            df_sheet.to_excel(writer, sheet_name=sht, index=False)

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
