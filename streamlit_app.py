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
    """Si no existe 'versions/Stock_Original.xlsx', lo creamos a partir de STOCK_FILE."""
    if not os.path.exists(ORIGINAL_FILE):
        if os.path.exists(STOCK_FILE):
            shutil.copy(STOCK_FILE, ORIGINAL_FILE)
        else:
            st.error(f"No se encontró {STOCK_FILE}. Asegúrate de subirlo.")

init_original()

def load_data():
    """Lee todas las hojas en un dict {nombre_hoja: DataFrame} y quita 'Restantes'."""
    try:
        data = pd.read_excel(STOCK_FILE, sheet_name=None, engine="openpyxl")
        # Borramos 'Restantes' si existe
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
    """Fuerza tipos en las columnas más comunes."""
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

# ============= LÓGICA DE LOTES =============
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
            "Primers DNA","Uracil-DNA Glycosylase heat-labile","Reagents DL8","Chef supplies (plásticos)",
            "Placas","Solutions DL8"
        ],
        "kit-Chef (TEMPLADO)": [
            "Ion 550 TM Chef Reagents","Chef Solutions","Chef Supplies (plásticos)","Solutions Reagent S5",
            "Botellas S5","Chip secuenciación Ion 550 TM Chip Kit"
        ],
        "Recover All TM Multi-Sample RNA/DNA Isolation workflow-Kit": [
            "Kit extracción DNA/RNA","RecoverAll TM kit (Dnase, protease,…)","H2O RNA free","Tubos fondo cónico"
        ]
    }
}

panel_order = ["FOCUS","OCA","OCA PLUS"]
panel_index = {p:i for i,p in enumerate(panel_order)}

colors = [
    "#FED7D7","#FEE2E2","#FFEDD5","#FEF9C3","#D9F99D",
    "#CFFAFE","#E0E7FF","#FBCFE8","#F9A8D4","#E9D5FF",
    "#FFD700","#F0FFF0","#D1FAE5","#BAFEE2","#A7F3D0","#FFEC99"
]
color_cycle = itertools.cycle(colors)

sub_lot_metadata = {}
for p in panel_order:
    subdict = LOTS_DATA.get(p,{})
    sublots = list(subdict.keys())
    for j, subl in enumerate(sublots):
        sub_lot_metadata[(p, subl)] = (panel_index[p], j, next(color_cycle))

def find_sub_lot(nombre_prod:str):
    """Retorna (panelName, subloteName, esPrincipal)."""
    for p in panel_order:
        subdict = LOTS_DATA.get(p,{})
        for sublote_name, reactivos in subdict.items():
            if nombre_prod == sublote_name:
                return (p, sublote_name, True)
            if nombre_prod in reactivos:
                return (p, sublote_name, False)
    return None

def build_lote_info(df: pd.DataFrame):
    """Agrega col: PanelIdx, SubLoteIdx, EsPrincipal, Color para agrupar/colorear. 
       Se ocultarán después."""
    df = df.copy()
    df["PanelIdx"] = 999
    df["SubLoteIdx"] = 999
    df["EsPrincipal"] = False
    df["Color"] = ""
    for i, row in df.iterrows():
        nombre_prod = row.get("Nombre producto","")
        info = find_sub_lot(nombre_prod)
        if info:
            p, sl, is_main = info
            (p_idx, s_idx, c) = sub_lot_metadata.get((p, sl),(999,999,""))
            df.at[i,"PanelIdx"] = p_idx
            df.at[i,"SubLoteIdx"] = s_idx
            df.at[i,"EsPrincipal"] = is_main
            df.at[i,"Color"] = c
    return df

def calc_alarma(row):
    """'Alarma': '🔴' si Stock=0 y Fecha Pedida=None, '🟨' si Stock=0 y FechaPed !=None."""
    stock_val = row.get("Stock",0)
    fecha_ped = row.get("Fecha Pedida",None)
    if stock_val == 0 and pd.isna(fecha_ped):
        return "🔴"
    elif stock_val == 0 and not pd.isna(fecha_ped):
        return "🟨"
    return ""

def style_lote(row):
    """Aplica color a toda la fila, y si es sub-lote principal (EsPrincipal=True) pone 'Nombre producto' en negrita."""
    bg = row.get("Color","")
    es_main = row.get("EsPrincipal",False)
    styles = [f"background-color:{bg}"]*len(row)
    if es_main:
        if "Nombre producto" in row.index:
            idx = row.index.get_loc("Nombre producto")
            styles[idx]+= "; font-weight:bold"
    return styles


# BARRA LATERAL
with st.sidebar:
    with st.expander("🔎 Ver / Gestionar versiones guardadas", expanded=False):
        if data_dict:
            files = sorted(os.listdir(VERSIONS_DIR))
            versions_no_original = [f for f in files if f != "Stock_Original.xlsx"]
            if versions_no_original:
                version_sel = st.selectbox("Selecciona versión:", versions_no_original)
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
            st.error("No hay data_dict. Revisa Stock_Original.xlsx.")
            st.stop()

    with st.expander("⚠️ Alarmas", expanded=False):
        st.write("Col 'Alarma':")
        st.write("- '🔴' => Stock=0 y FechaPed=None")
        st.write("- '🟨' => Stock=0 y FechaPed!=None")

    with st.expander("Reactivo Agotado (Consumido en Lab)", expanded=False):
        if data_dict:
            st.write("Selecciona la hoja y reactivo para consumir stock sin crear versión.")
            hojas_agotado = list(data_dict.keys())
            hoja_sel_consumo = st.selectbox("Hoja a consumir:", hojas_agotado, key="cons_hoja_sel")
            df_agotado = data_dict[hoja_sel_consumo].copy()
            df_agotado = enforce_types(df_agotado)

            if "Nombre producto" in df_agotado.columns and "Ref. Fisher" in df_agotado.columns:
                disp_consumo = df_agotado.apply(lambda r: f"{r['Nombre producto']} ({r['Ref. Fisher']})", axis=1)
            else:
                disp_consumo = df_agotado.iloc[:,0].astype(str)

            reactivo_consumir = st.selectbox("Reactivo:", disp_consumo.unique(), key="cons_react_sel")
            idx_c = disp_consumo[disp_consumo==reactivo_consumir].index[0]
            stock_c = df_agotado.at[idx_c,"Stock"] if "Stock" in df_agotado.columns else 0

            uds_consumidas = st.number_input("Uds. consumidas", min_value=0, step=1)
            if st.button("Registrar Consumo en Lab"):
                nuevo_stock = max(0, stock_c - uds_consumidas)
                df_agotado.at[idx_c,"Stock"] = nuevo_stock
                st.warning(f"Consumidas {uds_consumidas} uds. Stock final => {nuevo_stock}")
                data_dict[hoja_sel_consumo] = df_agotado
                st.success("No se genera versión, cambios solo en memoria.")
        else:
            st.error("No hay data_dict. Revisa Stock_Original.xlsx.")
            st.stop()

# CUERPO
st.title("📦 Control de Stock: Lotes agrupados + Alarmas con íconos (Ocultando columnas internas)")

if not data_dict:
    st.error("No se pudo cargar la base de datos.")
    st.stop()

st.markdown("---")
st.header("Edición en Hoja Principal y Guardado")

hojas_principales = list(data_dict.keys())
sheet_name = st.selectbox("Selecciona la hoja a editar:", hojas_principales, key="main_sheet_sel")
df_main_original = data_dict[sheet_name].copy()
df_main_original = enforce_types(df_main_original)

# ---- A) df_for_style con col extra y ordenado
df_for_style = df_main_original.copy()
df_for_style["Alarma"] = df_for_style.apply(calc_alarma, axis=1)
df_for_style = build_lote_info(df_for_style)
df_for_style.sort_values(by=["PanelIdx","SubLoteIdx","EsPrincipal"], ascending=[True,True,False], inplace=True)
df_for_style.reset_index(drop=True, inplace=True)

styled_df = df_for_style.style.apply(style_lote, axis=1)

# B) Ocultamos las col 'PanelIdx','SubLoteIdx','EsPrincipal','Color' en la vista
all_cols = df_for_style.columns.tolist()
cols_to_hide = ["PanelIdx","SubLoteIdx","EsPrincipal","Color"]
final_cols = [c for c in all_cols if c not in cols_to_hide]
table_html = styled_df.to_html(columns=final_cols)

# C) df_main final (el DataFrame real sin col internas)
df_main = df_for_style.copy()
df_main.drop(columns=cols_to_hide, inplace=True, errors="ignore")

# D) Mostramos la tabla
st.write("#### Vista de la Hoja (col 'Alarma', sin 'Restantes', ni columnas internas)")
st.write(table_html, unsafe_allow_html=True)

# E) Seleccionar Reactivo a Modificar
if "Nombre producto" in df_main.columns and "Ref. Fisher" in df_main.columns:
    display_series = df_main.apply(lambda r: f"{r['Nombre producto']} ({r['Ref. Fisher']})", axis=1)
else:
    display_series = df_main.iloc[:,0].astype(str)

reactivo_sel = st.selectbox("Selecciona Reactivo a Modificar:", display_series.unique(), key="react_modif")
row_index = display_series[display_series==reactivo_sel].index[0]

def get_val(col, default=None):
    return df_main.at[row_index, col] if col in df_main.columns else default

lote_actual = get_val("NºLote", 0)
caducidad_actual = get_val("Caducidad", None)
fecha_pedida_actual = get_val("Fecha Pedida", None)
fecha_llegada_actual = get_val("Fecha Llegada", None)
sitio_almacenaje_actual = get_val("Sitio almacenaje","")
uds_actual = get_val("Uds.",0)
stock_actual = get_val("Stock",0)

cA,cB,cC,cD = st.columns([1,1,1,1])
with cA:
    lote_nuevo = st.number_input("Nº de Lote", value=int(lote_actual), step=1)
    caducidad_nueva = st.date_input("Caducidad", value=caducidad_actual if pd.notna(caducidad_actual) else None)
with cB:
    fp_date = st.date_input("Fecha Pedida (fecha)",
                            value=fecha_pedida_actual.date() if pd.notna(fecha_pedida_actual) else None,
                            key="fp_date_main")
    fp_time = st.time_input("Hora Pedida",
                            value=fecha_pedida_actual.time() if pd.notna(fecha_pedida_actual) else datetime.time(0,0),
                            key="fp_time_main")
with cC:
    fl_date = st.date_input("Fecha Llegada (fecha)",
                            value=fecha_llegada_actual.date() if pd.notna(fecha_llegada_actual) else None,
                            key="fl_date_main")
    fl_time = st.time_input("Hora Llegada",
                            value=fecha_llegada_actual.time() if pd.notna(fecha_llegada_actual) else datetime.time(0,0),
                            key="fl_time_main")
with cD:
    st.write("")
    st.write("")
    if st.button("Refrescar Página"):
        st.rerun()

# Convertir a Timestamp para evitar FutureWarning
fecha_pedida_nueva = None
if fp_date is not None:
    dt_ped = datetime.datetime.combine(fp_date, fp_time)
    fecha_pedida_nueva = pd.to_datetime(dt_ped)  # casting

fecha_llegada_nueva = None
if fl_date is not None:
    dt_lleg = datetime.datetime.combine(fl_date, fl_time)
    fecha_llegada_nueva = pd.to_datetime(dt_lleg)

st.write("Sitio de Almacenaje")
opciones_sitio = ["Congelador 1","Congelador 2","Frigorífico","Tª Ambiente"]
sitio_principal = sitio_almacenaje_actual.split(" - ")[0] if " - " in sitio_almacenaje_actual else sitio_almacenaje_actual
if sitio_principal not in opciones_sitio:
    sitio_principal = opciones_sitio[0]
sitio_top = st.selectbox("Tipo Almacenaje", opciones_sitio, index=opciones_sitio.index(sitio_principal))

subopc = ""
if sitio_top=="Congelador 1":
    caj = [f"Cajón {i}" for i in range(1,9)]
    subopc = st.selectbox("Cajón (1 Arriba,8 Abajo)", caj)
elif sitio_top=="Congelador 2":
    caj = [f"Cajón {i}" for i in range(1,7)]
    subopc = st.selectbox("Cajón (1 Arriba,6 Abajo)", caj)
elif sitio_top=="Frigorífico":
    bal = [f"Balda {i}" for i in range(1,8)] + ["Puerta"]
    subopc = st.selectbox("Baldas (1 Arriba, 7 Abajo)", bal)
elif sitio_top=="Tª Ambiente":
    com = st.text_input("Comentario (opcional)")
    subopc = com.strip()

if subopc:
    sitio_almacenaje_nuevo = f"{sitio_top} - {subopc}"
else:
    sitio_almacenaje_nuevo = sitio_top

if st.button("Guardar Cambios"):
    # Si se introduce fecha llegada => borrar pedida
    if pd.notna(fecha_llegada_nueva):
        fecha_pedida_nueva = pd.NaT

    # Sumar stock si llega
    if "Stock" in df_main.columns:
        if fecha_llegada_nueva != fecha_llegada_actual and pd.notna(fecha_llegada_nueva):
            df_main.at[row_index,"Stock"] = stock_actual + uds_actual
            st.info(f"Sumadas {uds_actual} uds al stock => {stock_actual+uds_actual}")

    # Asignar con casting para evitar FutureWarning
    if "NºLote" in df_main.columns:
        df_main.at[row_index,"NºLote"] = int(lote_nuevo)

    if "Caducidad" in df_main.columns:
        if pd.notna(caducidad_nueva):
            df_main.at[row_index,"Caducidad"] = pd.to_datetime(caducidad_nueva)
        else:
            df_main.at[row_index,"Caducidad"] = pd.NaT

    if "Fecha Pedida" in df_main.columns:
        if pd.notna(fecha_pedida_nueva):
            df_main.at[row_index,"Fecha Pedida"] = pd.to_datetime(fecha_pedida_nueva)
        else:
            df_main.at[row_index,"Fecha Pedida"] = pd.NaT

    if "Fecha Llegada" in df_main.columns:
        if pd.notna(fecha_llegada_nueva):
            df_main.at[row_index,"Fecha Llegada"] = pd.to_datetime(fecha_llegada_nueva)
        else:
            df_main.at[row_index,"Fecha Llegada"] = pd.NaT

    if "Sitio almacenaje" in df_main.columns:
        df_main.at[row_index,"Sitio almacenaje"] = sitio_almacenaje_nuevo

    data_dict[sheet_name] = df_main

    new_file = crear_nueva_version_filename()
    # Guardamos cada hoja en new_file sin col internas
    with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
        for sht, df_sht in data_dict.items():
            # Repetimos la operación para cada sheet
            # Por si se hubieran creado col internas en otros sitios
            if all(col in df_sht.columns for col in ["PanelIdx","SubLoteIdx","EsPrincipal","Color"]):
                dtemp = df_sht.drop(columns=["PanelIdx","SubLoteIdx","EsPrincipal","Color"], errors="ignore")
            else:
                dtemp = df_sht
            dtemp.to_excel(writer, sheet_name=sht, index=False)

    with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
        for sht, df_sht in data_dict.items():
            if all(col in df_sht.columns for col in ["PanelIdx","SubLoteIdx","EsPrincipal","Color"]):
                dtemp = df_sht.drop(columns=["PanelIdx","SubLoteIdx","EsPrincipal","Color"], errors="ignore")
            else:
                dtemp = df_sht
            dtemp.to_excel(writer, sheet_name=sht, index=False)

    st.success(f"✅ Cambios guardados en '{new_file}' y '{STOCK_FILE}'.")

    excel_bytes = generar_excel_en_memoria(df_main, sheet_nm=sheet_name)
    st.download_button(
        label="Descargar Excel modificado",
        data=excel_bytes,
        file_name="Reporte_Stock.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.rerun()
