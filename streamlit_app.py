import streamlit as st
import streamlit_authenticator as stauth
import pandas as pd
import numpy as np
import datetime
import shutil
import os
from io import BytesIO
import itertools
import openpyxl

st.set_page_config(page_title="Control de Stock con Lotes", layout="centered")

# ---------------------------
# Autenticación (estructura actualizada)
# ---------------------------
credentials = {
    "usernames": {
        "user1": {
            "email": "user1@example.com",
            "name": "admin",
            "password": "$2b$12$j2s41NdHSUTSL.1xEM/GyeKX7dzMZTpyLnq7p/g/j2aldw.KC5FxS"
        },
        "user2": {
            "email": "user2@example.com",
            "name": "Usuario Dos",
            "password": "$2b$12$F9F3nZL9eFQKyF2.0tKbEe2KKFZQ3LCO6X5FA5u2Lz8mL3yh5Ew0a"
        }
    }
}

cookie_key = "mi_cookie_secreta"
signature_key = "mi_signature_secreta"

authenticator = stauth.Authenticate(
    credentials,
    cookie_name=cookie_key,
    key=signature_key,
    cookie_expiry_days=1
)

authenticator.login(location="main")

if st.session_state["authentication_status"]:
    st.success(f"Bienvenido, {st.session_state['name']}!")
elif st.session_state["authentication_status"] is False:
    st.error("Usuario o contraseña incorrectos.")
    st.stop()
elif st.session_state["authentication_status"] is None:
    st.warning("Por favor, ingresa tus credenciales.")
    st.stop()

if st.button("Cerrar sesión"):
    authenticator.logout()
    st.rerun()


# -------------------------------------------------------------------------
# EXCEL A (Stock_Original)
# -------------------------------------------------------------------------
STOCK_FILE = "Stock_Original.xlsx"
VERSIONS_DIR = "versions"
ORIGINAL_FILE = os.path.join(VERSIONS_DIR, "Stock_Original.xlsx")

os.makedirs(VERSIONS_DIR, exist_ok=True)

def init_original():
    if not os.path.exists(ORIGINAL_FILE):
        if os.path.exists(STOCK_FILE):
            shutil.copy(STOCK_FILE, ORIGINAL_FILE)
        else:
            st.error(f"No se encontró {STOCK_FILE}.")

init_original()

def load_data_a():
    try:
        data = pd.read_excel(STOCK_FILE, sheet_name=None, engine="openpyxl")
        for sheet, df_sheet in data.items():
            if "Restantes" in df_sheet.columns:
                df_sheet.drop(columns=["Restantes"], inplace=True, errors="ignore")
        return data
    except FileNotFoundError:
        st.error("No se encontró Stock_Original.xlsx.")
        return {}
    except Exception as e:
        st.error(f"Error al cargar Stock_Original: {e}")
        return {}

def crear_nueva_version_filename():
    fh = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return os.path.join(VERSIONS_DIR, f"Stock_{fh}.xlsx")


# -------------------------------------------------------------------------
# EXCEL B (Stock_Historico)
# -------------------------------------------------------------------------
STOCK_FILE_B = "Stock_Historico.xlsx"
VERSIONS_DIR_B = "versions_b"
ORIGINAL_FILE_B = os.path.join(VERSIONS_DIR_B, "Stock_Historico_Original.xlsx")

os.makedirs(VERSIONS_DIR_B, exist_ok=True)

def init_original_b():
    if not os.path.exists(ORIGINAL_FILE_B):
        if os.path.exists(STOCK_FILE_B):
            shutil.copy(STOCK_FILE_B, ORIGINAL_FILE_B)
        else:
            df_empty = pd.DataFrame(columns=[
                "Ref. Saturno","Ref. Fisher","Nombre producto","NºLote","Caducidad",
                "Fecha Pedida","Fecha Llegada","Sitio almacenaje","Uds.","Stock"
            ])
            with pd.ExcelWriter(STOCK_FILE_B, engine="openpyxl") as writer:
                df_empty.to_excel(writer, sheet_name="Hoja1", index=False)
            shutil.copy(STOCK_FILE_B, ORIGINAL_FILE_B)

init_original_b()

def crear_nueva_version_filename_b():
    fh = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return os.path.join(VERSIONS_DIR_B, f"StockB_{fh}.xlsx")


# -------------------------------------------------------------------------
# FUNCIONES COMUNES
# -------------------------------------------------------------------------
def generar_excel_en_memoria(df_act: pd.DataFrame, sheet_nm="Hoja1"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_act.to_excel(writer, sheet_name=sheet_nm, index=False)
    output.seek(0)
    return output.getvalue()

def enforce_types(df: pd.DataFrame):
    # Ref. Saturno como int
    if "Ref. Saturno" in df.columns:
        df["Ref. Saturno"] = pd.to_numeric(df["Ref. Saturno"], errors="coerce").fillna(0).astype(int)
    # Ref. Fisher como str
    if "Ref. Fisher" in df.columns:
        df["Ref. Fisher"] = df["Ref. Fisher"].astype(str)
    # Nombre producto como str
    if "Nombre producto" in df.columns:
        df["Nombre producto"] = df["Nombre producto"].astype(str)
    # Tª como str
    if "Tª" in df.columns:
        df["Tª"] = df["Tª"].astype(str)
    # Uds. como int
    if "Uds." in df.columns:
        df["Uds."] = pd.to_numeric(df["Uds."], errors="coerce").fillna(0).astype(int)
    # NºLote como str (para admitir letras + números)
    if "NºLote" in df.columns:
        df["NºLote"] = df["NºLote"].astype(str).fillna("")
    # Fechas como datetime
    for col in ["Caducidad","Fecha Pedida","Fecha Llegada"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    # Sitio almacenaje como str
    if "Sitio almacenaje" in df.columns:
        df["Sitio almacenaje"] = df["Sitio almacenaje"].astype(str)
    # Stock como int
    if "Stock" in df.columns:
        df["Stock"] = pd.to_numeric(df["Stock"], errors="coerce").fillna(0).astype(int)
    return df

def load_data_b():
    if not os.path.exists(STOCK_FILE_B):
        return {}
    try:
        data_b = pd.read_excel(STOCK_FILE_B, sheet_name=None, engine="openpyxl")
        for sheetb, dfb in data_b.items():
            if "Restantes" in dfb.columns:
                dfb.drop(columns=["Restantes"], inplace=True, errors="ignore")
        return data_b
    except:
        return {}

# -------------------------------------------------------------------------
# USAR st.session_state
# -------------------------------------------------------------------------
if "data_dict" not in st.session_state:
    st.session_state["data_dict"] = load_data_a()

if "data_dict_b" not in st.session_state:
    st.session_state["data_dict_b"] = load_data_b()

data_dict = st.session_state["data_dict"]
data_dict_b = st.session_state["data_dict_b"]

# -------------------------------------------------------------------------
# LÓGICA DE LOTES Y ESTILOS
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

panel_order = ["FOCUS","OCA","OCA PLUS"]

colors = [
    "#FED7D7", "#FEE2E2", "#FFEDD5", "#FEF9C3", "#D9F99D",
    "#CFFAFE", "#E0E7FF", "#FBCFE8", "#F9A8D4", "#E9D5FF",
    "#FFD700", "#F0FFF0", "#D1FAE5", "#BAFEE2", "#A7F3D0", "#FFEC99"
]

def build_group_info_by_ref(df: pd.DataFrame, panel_default=None):
    df = df.copy()
    df["GroupID"] = df["Ref. Saturno"]
    group_sizes = df.groupby("GroupID").size().to_dict()
    df["GroupCount"] = df["GroupID"].apply(lambda x: group_sizes.get(x,0))

    unique_ids = sorted(df["GroupID"].unique())
    color_cycle_local = itertools.cycle(colors)
    group_color_map = {}
    for gid in unique_ids:
        group_color_map[gid] = next(color_cycle_local)
    df["ColorGroup"] = df["GroupID"].apply(lambda x: group_color_map.get(x,"#FFFFFF"))

    group_titles = []
    if panel_default in LOTS_DATA:
        group_titles = [t.strip().lower() for t in LOTS_DATA[panel_default].keys()]
    df["EsTitulo"] = False
    for gid, group_df in df.groupby("GroupID"):
        mask = group_df["Nombre producto"].str.strip().str.lower().isin(group_titles)
        if mask.any():
            idxs = group_df[mask].index
            df.loc[idxs,"EsTitulo"] = True
        else:
            first_idx = group_df.index[0]
            df.at[first_idx,"EsTitulo"] = True

    df["MultiSort"] = df["GroupCount"].apply(lambda x: 0 if x>1 else 1)
    df["NotTitulo"] = df["EsTitulo"].apply(lambda x: 0 if x else 1)
    return df

def calc_alarma(row):
    s = row.get("Stock",0)
    fp = row.get("Fecha Pedida",None)
    if s==0 and pd.isna(fp):
        return "🔴"
    elif s==0 and not pd.isna(fp):
        return "🟨"
    return ""

def style_lote(row):
    bg = row.get("ColorGroup","")
    es_titulo = row.get("EsTitulo",False)
    styles = [f"background-color:{bg}"]*len(row)
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
# REACTIVO AGOTADO (Consumido en Lab)
# -------------------------------------------------------------------------
st.title("📦 Reactivo Agotado (Consumido en Lab)")

if not st.session_state["data_dict"]:
    st.error("No se pudo cargar la base de datos (A).")
    st.stop()

with st.expander("Consumir Reactivo en Lab", expanded=False):
    hojas_a = list(st.session_state["data_dict"].keys())
    hoja_sel = st.selectbox("Hoja A a consumir:", hojas_a)
    
    df_a = st.session_state["data_dict"][hoja_sel].copy()
    df_a = enforce_types(df_a)

    # Nombre producto
    if "Nombre producto" not in df_a.columns:
        st.error("No existe columna 'Nombre producto' en esta hoja A.")
        st.stop()
    nombres_unicos = sorted(df_a["Nombre producto"].dropna().unique())
    nombre_sel = st.selectbox("Nombre producto en A:", nombres_unicos)

    # Tomamos la primera fila que coincida
    df_cand = df_a[df_a["Nombre producto"]==nombre_sel]
    if df_cand.empty:
        st.warning("No se encontró ese nombre en esta hoja A.")
    else:
        idx_c = df_cand.index[0]
        stock_c = df_a.at[idx_c,"Stock"] if "Stock" in df_a.columns else 0
        
        uds_consumir = st.number_input("Uds. a consumir en A:", min_value=0, step=1)
        if st.button("Consumir en Lab (memoria)"):
            nuevo_stock = max(0, stock_c - uds_consumir)
            df_a.at[idx_c,"Stock"] = nuevo_stock
            if nuevo_stock==0:
                # vaciar las columnas => NºLote, Caducidad, Fecha Pedida, Fecha Llegada, Sitio almacenaje
                for col_vaciar in ["NºLote","Caducidad","Fecha Pedida","Fecha Llegada","Sitio almacenaje"]:
                    if col_vaciar in df_a.columns:
                        if col_vaciar in ["Caducidad","Fecha Pedida","Fecha Llegada"]:
                            df_a.at[idx_c, col_vaciar] = pd.NaT
                        else:
                            df_a.at[idx_c, col_vaciar] = ""
            st.session_state["data_dict"][hoja_sel] = df_a
            st.warning(f"Consumidas {uds_consumir} uds. Stock final => {nuevo_stock}. (Sólo en memoria).")

    st.write("**Eliminar en B** => introduce el Lote exacto. Si coincide Nombre+Lote, se borra de B.")
    # Lote (B) como texto: puede tener letras + números
    lote_b = st.text_input("Nº de Lote (en B) - admite letras y números:", value="")

    if st.button("Guardar Cambios en Consumo Lab"):
        # 1) Escribimos la versión de A que está en st.session_state
        new_file = crear_nueva_version_filename()
        with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
            for sht, df_sht in st.session_state["data_dict"].items():
                cols_int = ["ColorGroup","EsTitulo","GroupCount","MultiSort","NotTitulo","GroupID"]
                temp = df_sht.drop(columns=cols_int, errors="ignore")
                temp.to_excel(writer, sheet_name=sht, index=False)
        with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
            for sht, df_sht in st.session_state["data_dict"].items():
                cols_int = ["ColorGroup","EsTitulo","GroupCount","MultiSort","NotTitulo","GroupID"]
                temp = df_sht.drop(columns=cols_int, errors="ignore")
                temp.to_excel(writer, sheet_name=sht, index=False)

        # 2) Eliminar en B si coincide
        if hoja_sel in st.session_state["data_dict_b"]:
            df_b_hoja = st.session_state["data_dict_b"][hoja_sel].copy()
            if "Nombre producto" in df_b_hoja.columns and "NºLote" in df_b_hoja.columns:
                df_b_hoja = df_b_hoja[~(
                    (df_b_hoja["Nombre producto"] == nombre_sel) &
                    (df_b_hoja["NºLote"] == lote_b)
                )]
                st.session_state["data_dict_b"][hoja_sel] = df_b_hoja

                new_file_b = crear_nueva_version_filename_b()
                with pd.ExcelWriter(new_file_b, engine="openpyxl") as writer_b:
                    for sht_b, df_sht_b in st.session_state["data_dict_b"].items():
                        df_sht_b.to_excel(writer_b, sheet_name=sht_b, index=False)

                with pd.ExcelWriter(STOCK_FILE_B, engine="openpyxl") as writer_b:
                    for sht_b, df_sht_b in st.session_state["data_dict_b"].items():
                        df_sht_b.to_excel(writer_b, sheet_name=sht_b, index=False)

        st.success("✅ Cambios guardados en Hoja A y B (si coincidía).")
        st.rerun()


# -------------------------------------------------------------------------
# CUERPO PRINCIPAL => Edición en Hoja Principal (A)
# -------------------------------------------------------------------------
st.header("Edición en Hoja Principal y Guardado (Excel A)")

if not st.session_state["data_dict"]:
    st.error("No se pudo cargar la base de datos (A).")
    st.stop()

hojas_principales = list(st.session_state["data_dict"].keys())
sheet_name = st.selectbox("Selecciona la hoja a editar:", hojas_principales, key="main_sheet_sel")
df_main_original = st.session_state["data_dict"][sheet_name].copy()
df_main_original = enforce_types(df_main_original)

df_for_style = df_main_original.copy()
df_for_style["Alarma"] = df_for_style.apply(calc_alarma, axis=1)
df_for_style = build_group_info_by_ref(df_for_style, panel_default=sheet_name)

df_for_style.sort_values(by=["MultiSort","GroupID","NotTitulo"], inplace=True)
df_for_style.reset_index(drop=True, inplace=True)
styled_df = df_for_style.style.apply(style_lote, axis=1)

all_cols = df_for_style.columns.tolist()
cols_to_hide = ["ColorGroup","EsTitulo","GroupCount","MultiSort","NotTitulo","GroupID"]
final_cols = [c for c in all_cols if c not in cols_to_hide]

table_html = styled_df.to_html(columns=final_cols)
df_main = df_for_style.copy()
df_main.drop(columns=cols_to_hide, inplace=True, errors="ignore")

st.write("#### Vista de la Hoja (con columna 'Alarma' y sin columnas internas)")
st.write(table_html, unsafe_allow_html=True)

if "Nombre producto" in df_main.columns and "Ref. Fisher" in df_main.columns:
    display_series = df_main.apply(lambda r: f"{r['Nombre producto']} ({r['Ref. Fisher']})", axis=1)
else:
    display_series = df_main.iloc[:,0].astype(str)

reactivo_sel = st.selectbox("Selecciona Reactivo a Modificar:", display_series.unique(), key="react_modif")
row_index = display_series[display_series == reactivo_sel].index[0]

def get_val(col, default=None):
    return df_main.at[row_index, col] if col in df_main.columns else default

lote_actual = get_val("NºLote","")
caducidad_actual = get_val("Caducidad",None)
fecha_pedida_actual = get_val("Fecha Pedida",None)
fecha_llegada_actual = get_val("Fecha Llegada",None)
sitio_almacenaje_actual = get_val("Sitio almacenaje","")
uds_actual = get_val("Uds.",0)
stock_actual = get_val("Stock",0)

colA, colB, colC, colD = st.columns([1,1,1,1])
with colA:
    # NºLote como texto (puede contener letras y números)
    lote_new = st.text_input("Nº de Lote", value=str(lote_actual))
    cad_new = st.date_input("Caducidad", value=caducidad_actual if pd.notna(caducidad_actual) else None)
with colB:
    fped_date = st.date_input("Fecha Pedida (fecha)",
                              value=fecha_pedida_actual.date() if pd.notna(fecha_pedida_actual) else None,
                              key="fped_date_main")
    fped_time = st.time_input("Hora Pedida",
                              value=fecha_pedida_actual.time() if pd.notna(fecha_pedida_actual) else datetime.time(0,0),
                              key="fped_time_main")
with colC:
    flleg_date = st.date_input("Fecha Llegada (fecha)",
                               value=fecha_llegada_actual.date() if pd.notna(fecha_llegada_actual) else None,
                               key="flleg_date_main")
    flleg_time = st.time_input("Hora Llegada",
                               value=fecha_llegada_actual.time() if pd.notna(fecha_llegada_actual) else datetime.time(0,0),
                               key="flleg_time_main")
with colD:
    st.write("")
    st.write("")
    if st.button("Refrescar Página"):
        st.rerun()

fped_new = None
if fped_date is not None:
    dt_ped = datetime.datetime.combine(fped_date, fped_time)
    fped_new = pd.to_datetime(dt_ped)
flleg_new = None
if flleg_date is not None:
    dt_lleg = datetime.datetime.combine(flleg_date, flleg_time)
    flleg_new = pd.to_datetime(dt_lleg)

st.write("Sitio de Almacenaje")
opciones_sitio = ["Congelador 1","Congelador 2","Frigorífico","Tª Ambiente"]
sitio_p = sitio_almacenaje_actual.split(" - ")[0] if " - " in sitio_almacenaje_actual else sitio_almacenaje_actual
if sitio_p not in opciones_sitio:
    sitio_p = opciones_sitio[0]
sel_top = st.selectbox("Almacén Principal", opciones_sitio, index=opciones_sitio.index(sitio_p))
subopc=""
if sel_top=="Congelador 1":
    cajs=[f"Cajón {i}" for i in range(1,9)]
    subopc= st.selectbox("Cajón (1 Arriba,8 Abajo)", cajs)
elif sel_top=="Congelador 2":
    cajs=[f"Cajón {i}" for i in range(1,7)]
    subopc= st.selectbox("Cajón (1 Arriba,6 Abajo)", cajs)
elif sel_top=="Frigorífico":
    blds=[f"Balda {i}" for i in range(1,8)] + ["Puerta"]
    subopc= st.selectbox("Baldas (1 Arriba,7 Abajo)", blds)
elif sel_top=="Tª Ambiente":
    com2= st.text_input("Comentario (opt)")
    subopc= com2.strip()
if subopc:
    sitio_new = f"{sel_top} - {subopc}"
else:
    sitio_new = sel_top

group_order_selected = None
if pd.notna(fped_new):
    group_id = df_for_style.at[row_index,"GroupID"]
    group_reactivos = df_for_style[df_for_style["GroupID"]==group_id]
    if not group_reactivos.empty:
        if group_reactivos["EsTitulo"].any():
            lot_name = group_reactivos[group_reactivos["EsTitulo"]==True]["Nombre producto"].iloc[0]
        else:
            lot_name = f"Ref. Saturno {group_id}"
        group_reactivos_reset = group_reactivos.reset_index()
        options = group_reactivos_reset.apply(lambda r: f"{r['index']} - {r['Nombre producto']} ({r['Ref. Fisher']})", axis=1).tolist()
        st.markdown('<div class="big-select">', unsafe_allow_html=True)
        group_order_selected = st.multiselect(
            f"¿Pedir también los siguientes reactivos del lote **{lot_name}**?",
            options,
            default=options
        )
        st.markdown('</div>', unsafe_allow_html=True)

# Botón para Guardar Cambios en Hoja A
if st.button("Guardar Cambios en Hoja A"):
    if pd.notna(flleg_new):
        fped_new = pd.NaT

    if "Stock" in df_main.columns:
        # Si el usuario modificó la fecha de llegada (flleg_new) o cambió el Lote (lote_new),
        # sumamos uds_actual al stock_actual
        if (
            (flleg_new != fecha_llegada_actual and pd.notna(flleg_new))
            or
            (lote_new != lote_actual and lote_new.strip() != "")
        ):
            df_main.at[row_index, "Stock"] = stock_actual + uds_actual
            st.info(f"Añadidas {uds_actual} uds => stock={stock_actual + uds_actual}")

    # Guardar los cambios en df_main
    if "NºLote" in df_main.columns:
        df_main.at[row_index,"NºLote"] = lote_new  # ahora es str
    if "Caducidad" in df_main.columns:
        df_main.at[row_index,"Caducidad"] = cad_new if pd.notna(cad_new) else pd.NaT
    if "Fecha Pedida" in df_main.columns:
        df_main.at[row_index,"Fecha Pedida"] = fped_new
    if "Fecha Llegada" in df_main.columns:
        df_main.at[row_index,"Fecha Llegada"] = flleg_new
    if "Sitio almacenaje" in df_main.columns:
        df_main.at[row_index,"Sitio almacenaje"] = sitio_new

    if pd.notna(fped_new) and group_order_selected:
        for label in group_order_selected:
            try:
                i_val = int(label.split(" - ")[0])
                df_main.at[i_val,"Fecha Pedida"] = fped_new
            except Exception as e:
                st.error(f"Error actualizando índice {label}: {e}")

    st.session_state["data_dict"][sheet_name] = df_main

    new_file = crear_nueva_version_filename()
    with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
        for sht, df_sht in st.session_state["data_dict"].items():
            ocultar=["ColorGroup","EsTitulo","GroupCount","MultiSort","NotTitulo","GroupID"]
            tmp = df_sht.drop(columns=ocultar, errors="ignore")
            tmp.to_excel(writer, sheet_name=sht, index=False)

    with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
        for sht, df_sht in st.session_state["data_dict"].items():
            ocultar=["ColorGroup","EsTitulo","GroupCount","MultiSort","NotTitulo","GroupID"]
            tmp = df_sht.drop(columns=ocultar, errors="ignore")
            tmp.to_excel(writer, sheet_name=sht, index=False)

    # Insertar en B
    if sheet_name not in st.session_state["data_dict_b"]:
        st.session_state["data_dict_b"][sheet_name] = pd.DataFrame()

    df_b_sh = st.session_state["data_dict_b"][sheet_name].copy()
    nueva_fila = {
        "Ref. Saturno": df_main.at[row_index,"Ref. Saturno"] if "Ref. Saturno" in df_main.columns else 0,
        "Ref. Fisher": df_main.at[row_index,"Ref. Fisher"] if "Ref. Fisher" in df_main.columns else "",
        "Nombre producto": df_main.at[row_index,"Nombre producto"] if "Nombre producto" in df_main.columns else "",
        "NºLote": df_main.at[row_index,"NºLote"],
        "Caducidad": df_main.at[row_index,"Caducidad"],
        "Fecha Pedida": df_main.at[row_index,"Fecha Pedida"],
        "Fecha Llegada": df_main.at[row_index,"Fecha Llegada"],
        "Sitio almacenaje": df_main.at[row_index,"Sitio almacenaje"],
        "Uds.": df_main.at[row_index,"Uds."] if "Uds." in df_main.columns else 0,
        "Stock": df_main.at[row_index,"Stock"] if "Stock" in df_main.columns else 0,
        "Fecha Registro B": datetime.datetime.now()
    }
    df_b_sh = pd.concat([df_b_sh, pd.DataFrame([nueva_fila])], ignore_index=True)
    st.session_state["data_dict_b"][sheet_name] = df_b_sh

    new_file_b = crear_nueva_version_filename_b()
    with pd.ExcelWriter(new_file_b, engine="openpyxl") as writerB:
        for shtB, df_shtB in st.session_state["data_dict_b"].items():
            df_shtB.to_excel(writerB, sheet_name=shtB, index=False)
    with pd.ExcelWriter(STOCK_FILE_B, engine="openpyxl") as writerB:
        for shtB, df_shtB in st.session_state["data_dict_b"].items():
            df_shtB.to_excel(writerB, sheet_name=shtB, index=False)

    st.success(f"✅ Cambios guardados en '{new_file}' y '{STOCK_FILE}' (A). También en '{new_file_b}' y '{STOCK_FILE_B}' (B).")
    excel_bytes = generar_excel_en_memoria(df_main, sheet_nm=sheet_name)
    st.download_button("Descargar Excel A modificado", excel_bytes, "Reporte_Stock.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.rerun()
