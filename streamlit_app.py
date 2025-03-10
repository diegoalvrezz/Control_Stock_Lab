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

# ---------------------- DICCIONARIO DE LOTES ----------------------
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
        "Chip secuenciación liberación de protones 6 millones de lecturas": []
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

# Paleta de colores
color_list = [
    "#FED7D7","#FEE2E2","#FFEDD5","#FEF9C3","#D9F99D",
    "#CFFAFE","#E0E7FF","#FBCFE8","#F9A8D4","#E9D5FF",
    "#FFD700","#F0FFF0","#D1FAE5","#BAFEE2","#A7F3D0","#FFEC99"
]

# ----------------- Buscar sub-lote en LOTS_DATA -----------------
def find_sub_lot(nombre_prod: str, panel_name: str):
    """
    Si 'nombre_prod' coincide con un sub-lote principal o uno de sus reactivos
    en LOTS_DATA[panel_name], devuelve (sub_lote, esPrincipal).
    De lo contrario, None.
    Comparación case-insensitive.
    """
    if panel_name not in LOTS_DATA:
        return None
    subdict = LOTS_DATA[panel_name]
    np_lower = nombre_prod.strip().lower()
    for sub_lote, reactivos in subdict.items():
        if np_lower == sub_lote.strip().lower():
            return (sub_lote, True)
        for r in reactivos:
            if np_lower == r.strip().lower():
                return (sub_lote, False)
    return None

# ----------------- Lógica grouping final -----------------
def group_rows(df: pd.DataFrame, panel_name: str):
    """
    - Si 'Nombre producto' coincide con un sub-lote principal o reactivo => agrupa con ese sub-lote.
    - Si no, agrupa por la "Ref. Saturno" => se usará la Ref. Saturno como 'sub-lote' y la
      fila 'principal' es la primera encontrada de ese Saturno. 
    Devolvemos col extra: SubLoteName, EsPrincipal, ColorGroup
    """
    df = df.copy()
    df["SubLoteName"] = None
    df["EsPrincipal"] = False
    df["ColorGroup"] = "#FFFFFF"

    # 1) Asignar sub-lote según LOTS_DATA
    # 2) Si no, usar "Ref. Saturno"
    # 3) Asignar color

    # Recolectar sub-lotes que vayamos asignando
    assigned_sub_lotes = {}
    color_cycle = itertools.cycle(color_list)

    # Pre-computar si "Ref. Saturno" es 0 o no, para saber si es un "grupo" distinto
    # (Puede haber un Saturno=0 que se repite, p.ej. sin panel => se agrupan juntos.)
    for i, row in df.iterrows():
        np_name = str(row.get("Nombre producto","")).strip()
        sub_info = find_sub_lot(np_name, panel_name)
        if sub_info is not None:
            # (sub_lote, is_main)
            sub_lote, is_main = sub_info
            df.at[i,"SubLoteName"] = sub_lote
            df.at[i,"EsPrincipal"] = is_main
        else:
            # agrupa por 'Ref. Saturno' => 'sub-lote' = "RefSat_{xxx}"
            ref_sat = row.get("Ref. Saturno", 0)
            df.at[i,"SubLoteName"] = f"RefSat_{ref_sat}"
            # esPrincipal => si es la 1ra fila en la que aparece esa RefSat
            # lo marcamos en un dict
            # y la 1ra vez => is_main = True
            # resto => is_main=False
            if f"RefSat_{ref_sat}" not in assigned_sub_lotes:
                assigned_sub_lotes[f"RefSat_{ref_sat}"] = True
                df.at[i,"EsPrincipal"] = True
            else:
                df.at[i,"EsPrincipal"] = False

    # Recolectar lista sub-lotes usados
    unique_sub_lotes = df["SubLoteName"].unique().tolist()
    color_map = {}
    for sl in sorted(unique_sub_lotes):
        color_map[sl] = next(color_cycle)

    # Asignar color
    for i, row in df.iterrows():
        sl = row.get("SubLoteName")
        df.at[i,"ColorGroup"] = color_map.get(sl,"#FFFFFF")

    return df


def calc_alarma(row):
    """Col 'Alarma': '🔴' si Stock=0 y Fecha Pedida=None, '🟨' si Stock=0 y FechaPed!=None."""
    s = row.get("Stock", 0)
    fp = row.get("Fecha Pedida", None)
    if s == 0 and pd.isna(fp):
        return "🔴"
    elif s == 0 and not pd.isna(fp):
        return "🟨"
    return ""

def style_lote(row):
    """Aplica color a toda la fila, y si EsPrincipal=True, pone 'Nombre producto' en negrita."""
    bg = row.get("ColorGroup", "")
    es_main = row.get("EsPrincipal", False)
    styles = [f"background-color:{bg}"] * len(row)
    if es_main:
        # Poner en negrita 'Nombre producto'
        if "Nombre producto" in row.index:
            idx = row.index.get_loc("Nombre producto")
            styles[idx] += "; font-weight:bold"
    return styles

# BARRA LATERAL
with st.sidebar:
    st.write("## Opciones Base de Datos")
    if data_dict:
        files = sorted(os.listdir(VERSIONS_DIR))
        versions_no_original = [f for f in files if f != "Stock_Original.xlsx"]
        if versions_no_original:
            version_sel = st.selectbox("Selecciona versión:", versions_no_original)
            if version_sel:
                file_path = os.path.join(VERSIONS_DIR, version_sel)
                with open(file_path,"rb") as f:
                    st.download_button("Descargar "+version_sel, data=f, file_name=version_sel)
                # etc. ... (botones eliminar y limpiar)

    st.write("---")
    # Reactivo Agotado
    st.write("### Reactivo Agotado (sin crear versión)")
    if data_dict:
        hoja_cons = st.selectbox("Hoja para consumir Reactivo:", list(data_dict.keys()), key="hoja_consume")
        df_c = data_dict[hoja_cons].copy()
        df_c = enforce_types(df_c)

        if "Nombre producto" in df_c.columns and "Ref. Fisher" in df_c.columns:
            ds = df_c.apply(lambda r: f"{r['Nombre producto']} ({r['Ref. Fisher']})", axis=1)
        else:
            ds = df_c.iloc[:,0].astype(str)

        sel_reac = st.selectbox("Reactivo a consumir:", ds.unique())
        idx_reac = ds[ds==sel_reac].index[0]
        stock_val = df_c.at[idx_reac,"Stock"] if "Stock" in df_c.columns else 0
        cantidad = st.number_input("Cantidad consumida", min_value=0, step=1)

        if st.button("Consumir Stock"):
            nuevo_stk = max(0, stock_val - cantidad)
            df_c.at[idx_reac,"Stock"] = nuevo_stk
            data_dict[hoja_cons] = df_c
            st.success(f"Consumidas {cantidad} uds. Queda => {nuevo_stk} en stock.")
    else:
        st.error("No se pudieron cargar datos.")


# CUERPO
st.title("📦 Control de Stock con Grupos por LOTS_DATA o Ref. Saturno")
if not data_dict:
    st.error("No hay datos.")
    st.stop()

st.header("Edición en Hoja Principal y Guardado")

hoja = st.selectbox("Selecciona la Hoja:", list(data_dict.keys()))
df_main0 = data_dict[hoja].copy()
df_main0 = enforce_types(df_main0)

# Crea col 'Alarma'
df_main0["Alarma"] = df_main0.apply(calc_alarma, axis=1)
# Agrupa => sub-lote (LOTS_DATA) si coincide, si no => agrupa por 'Ref. Saturno'
df_main0 = group_rows(df_main0, panel_name=hoja)

# Ordenar => primero sub-lotes con la misma SubLoteName,
# dentro => la fila principal primero (EsPrincipal=True), resto después
df_main0["SortKey"] = df_main0["SubLoteName"].astype(str) + df_main0["EsPrincipal"].apply(lambda b: "0" if b else "1")
df_main0.sort_values("SortKey", inplace=True)
df_main0.reset_index(drop=True,inplace=True)

styled = df_main0.style.apply(style_lote, axis=1)

# Ocultar col internas => SubLoteName, EsPrincipal, ColorGroup, SortKey
ocultas = ["SubLoteName","EsPrincipal","ColorGroup","SortKey"]
final_cols = [c for c in df_main0.columns if c not in ocultas]
html_table = styled.to_html(columns=final_cols)

# Creamos df_main final
df_main = df_main0.copy()
df_main.drop(columns=ocultas, inplace=True, errors="ignore")

st.write("### Vista de la Hoja (col 'Alarma', grupos, sin col internas)")
st.write(html_table, unsafe_allow_html=True)

# Selección de Reactivo a Modificar
if "Nombre producto" in df_main.columns and "Ref. Fisher" in df_main.columns:
    ds2 = df_main.apply(lambda r: f"{r['Nombre producto']} ({r['Ref. Fisher']})", axis=1)
else:
    ds2 = df_main.iloc[:,0].astype(str)

sel_react = st.selectbox("Selecciona Reactivo:", ds2.unique())
idx_r = ds2[ds2==sel_react].index[0]

def gval(col, default=None):
    return df_main.at[idx_r,col] if col in df_main.columns else default

lote_val = gval("NºLote",0)
caduc_val = gval("Caducidad",None)
fped_val = gval("Fecha Pedida",None)
flleg_val = gval("Fecha Llegada",None)
sitio_val = gval("Sitio almacenaje","")
uds_val = gval("Uds.",0)
stock_val = gval("Stock",0)

c1,c2,c3,c4 = st.columns(4)
with c1:
    lote_new = st.number_input("Nº Lote", value=int(lote_val), step=1)
    cad_new = st.date_input("Caducidad", value=caduc_val if pd.notna(caduc_val) else None)
with c2:
    fped_date = st.date_input("Fecha Pedida (fecha)",
                              value=fped_val.date() if pd.notna(fped_val) else None,
                              key="fped_date_main")
    fped_time = st.time_input("Hora Pedida",
                              value=fped_val.time() if pd.notna(fped_val) else datetime.time(0,0),
                              key="fped_time_main")
with c3:
    flleg_date = st.date_input("Fecha Llegada (fecha)",
                               value=flleg_val.date() if pd.notna(flleg_val) else None,
                               key="flleg_date_main")
    flleg_time = st.time_input("Hora Llegada",
                               value=flleg_val.time() if pd.notna(flleg_val) else datetime.time(0,0),
                               key="flleg_time_main")
with c4:
    if st.button("Refrescar"):
        st.experimental_rerun()

fped_new = None
if fped_date is not None:
    dt_ped = datetime.datetime.combine(fped_date, fped_time)
    fped_new = pd.to_datetime(dt_ped)

flleg_new = None
if flleg_date is not None:
    dt_lleg = datetime.datetime.combine(flleg_date, flleg_time)
    flleg_new = pd.to_datetime(dt_lleg)

st.write("Sitio almacenaje")
opciones_sitio = ["Congelador 1","Congelador 2","Frigorífico","Tª Ambiente"]
sitio_p = sitio_val.split(" - ")[0] if " - " in sitio_val else sitio_val
if sitio_p not in opciones_sitio:
    sitio_p = opciones_sitio[0]
sel_top = st.selectbox("Almacén Principal", opciones_sitio, index=opciones_sitio.index(sitio_p))

subopc = ""
if sel_top=="Congelador 1":
    cajs = [f"Cajón {i}" for i in range(1,9)]
    subopc = st.selectbox("Cajón (1arriba,8abajo)", cajs)
elif sel_top=="Congelador 2":
    cajs = [f"Cajón {i}" for i in range(1,7)]
    subopc = st.selectbox("Cajón (1arriba,6abajo)", cajs)
elif sel_top=="Frigorífico":
    blds = [f"Balda {i}" for i in range(1,8)] + ["Puerta"]
    subopc = st.selectbox("Baldas(1arriba,7abajo)", blds)
elif sel_top=="Tª Ambiente":
    com2 = st.text_input("Comentario (opt)")
    subopc = com2.strip()

if subopc:
    sitio_new = f"{sel_top} - {subopc}"
else:
    sitio_new = sel_top

if st.button("Guardar Cambios"):
    # si llega => borramos pedida
    if pd.notna(flleg_new):
        fped_new = pd.NaT

    if "Stock" in df_main.columns:
        if flleg_new!=flleg_val and pd.notna(flleg_new):
            df_main.at[idx_r,"Stock"] = stock_val + uds_val
            st.info(f"Añadidas {uds_val} uds. => {stock_val + uds_val}")

    # Asignar
    if "NºLote" in df_main.columns:
        df_main.at[idx_r,"NºLote"] = int(lote_new)
    if "Caducidad" in df_main.columns:
        if pd.notna(cad_new):
            df_main.at[idx_r,"Caducidad"] = pd.to_datetime(cad_new)
        else:
            df_main.at[idx_r,"Caducidad"] = pd.NaT
    if "Fecha Pedida" in df_main.columns:
        if pd.notna(fped_new):
            df_main.at[idx_r,"Fecha Pedida"] = pd.to_datetime(fped_new)
        else:
            df_main.at[idx_r,"Fecha Pedida"] = pd.NaT
    if "Fecha Llegada" in df_main.columns:
        if pd.notna(flleg_new):
            df_main.at[idx_r,"Fecha Llegada"] = pd.to_datetime(flleg_new)
        else:
            df_main.at[idx_r,"Fecha Llegada"] = pd.NaT
    if "Sitio almacenaje" in df_main.columns:
        df_main.at[idx_r,"Sitio almacenaje"] = sitio_new

    data_dict[hoja] = df_main

    # Guardar
    new_file = crear_nueva_version_filename()
    with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
        for sht, dataf in data_dict.items():
            ocultar = ["SubLoteName","EsPrincipal","ColorGroup","SortKey"]
            df_save = dataf.drop(columns=ocultar, errors="ignore")
            df_save.to_excel(writer, sheet_name=sht, index=False)

    with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
        for sht, dataf in data_dict.items():
            ocultar = ["SubLoteName","EsPrincipal","ColorGroup","SortKey"]
            df_save = dataf.drop(columns=ocultar, errors="ignore")
            df_save.to_excel(writer, sheet_name=sht, index=False)

    st.success(f"Guardado en {new_file} y {STOCK_FILE}.")

    excel_bytes = generar_excel_en_memoria(df_main, sheet_nm=hoja)
    st.download_button("Descargar Excel modificado", excel_bytes, "Reporte_Stock.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.experimental_rerun()
