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
import time
import pytz

STOCK_FILE_B = "Stock_Historico.xlsx"
VERSIONS_DIR_B = "versions_b"
ORIGINAL_FILE_B = os.path.join(VERSIONS_DIR_B, "Stock_Historico_Original.xlsx")

st.set_page_config(page_title="Control de Stock con Lotes", layout="centered")
st.title("🔬 Control Stock Lab. Patología Molécular")

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

import glob

def obtener_ultima_version():
    archivos = glob.glob(f"{VERSIONS_DIR}/**/*.xlsx", recursive=True)
    archivos = [f for f in archivos if "Stock_Original.xlsx" not in f]
    if archivos:
        ultima_version = max(archivos, key=os.path.getctime)
        return ultima_version
    else:
        return STOCK_FILE  # si no hay versiones, devuelve la original

# Cargar automáticamente última versión guardada al inicio
archivo_a_cargar = obtener_ultima_version()

# Ahora cargamos los datos en session_state con manejo de errores
try:
    st.session_state["data_dict"] = pd.read_excel(archivo_a_cargar, sheet_name=None, engine="openpyxl")
except Exception as e:
    st.error(f"Error al cargar el archivo inicial ({archivo_a_cargar}): {e}")
    st.session_state["data_dict"] = {}

os.makedirs(VERSIONS_DIR, exist_ok=True)
os.makedirs(VERSIONS_DIR_B, exist_ok=True)

import calendar

def obtener_subcarpeta_versiones():
    zona_local = pytz.timezone('Europe/Madrid')
    ahora = datetime.datetime.now(zona_local)
    nombre_subcarpeta = ahora.strftime("%Y_%m_%B")  # Ej: 2025_03_Marzo
    ruta_subcarpeta = os.path.join(VERSIONS_DIR, nombre_subcarpeta)
    os.makedirs(ruta_subcarpeta, exist_ok=True)
    return ruta_subcarpeta

def crear_nueva_version_filename():
    ruta_subcarpeta = obtener_subcarpeta_versiones()
    zona_local = pytz.timezone('Europe/Madrid')
    fh = datetime.datetime.now(zona_local).strftime("%Y-%m-%d_%H-%M-%S")
    return os.path.join(ruta_subcarpeta, f"Stock_{fh}.xlsx")

# Explorador visual y subida manual en sidebar
with st.sidebar.expander("📂 Gestor avanzado de versiones", expanded=False):
    subcarpetas = sorted(
        [d for d in os.listdir(VERSIONS_DIR) if os.path.isdir(os.path.join(VERSIONS_DIR, d))],
        reverse=True
    )

    if subcarpetas:
        mes_elegido = st.selectbox("📅 Selecciona el mes para explorar versiones:", subcarpetas)
        ruta_actual = os.path.join(VERSIONS_DIR, mes_elegido)

        st.write(f"**Versiones guardadas en {ruta_actual}:**")
        archivos_versiones = sorted(os.listdir(ruta_actual), reverse=True)

        if archivos_versiones:
            import datetime
            versiones_df = pd.DataFrame({
                "Archivo": archivos_versiones,
                "Fecha creación": [
                    datetime.datetime.fromtimestamp(os.path.getctime(os.path.join(ruta_actual, f))
                ).strftime('%d/%m/%Y %H:%M:%S') for f in archivos_versiones]
            })
            st.dataframe(versiones_df)

            version_gestion = st.selectbox("Seleccione versión para gestionar:", archivos_versiones)
            ruta_version = os.path.join(ruta_actual, version_gestion)

            col_down, col_del = st.columns(2)
            with col_down:
                with open(ruta_version, "rb") as version_file:
                    st.download_button(
                        "Descargar versión",
                        data=version_file,
                        file_name=version_gestion,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            with col_del:
                confirm_eliminar = st.text_input("Escribe ELIMINAR para borrar", key="confirm_del_avanzado")
                if st.button("Eliminar versión seleccionada"):
                    if confirm_eliminar == "ELIMINAR":
                        os.remove(ruta_version)
                        st.success("Versión eliminada correctamente.")
                        time.sleep(1.5)
                        st.rerun()
                    else:
                        st.error("Debes escribir ELIMINAR para confirmar.")
        else:
            st.info("No hay versiones guardadas en esta subcarpeta.")

    else:
        st.info("Actualmente no existen subcarpetas de versiones en el directorio.")

    st.divider()
    st.write("⚠️ **Eliminar TODAS las versiones excepto la original:**")
    confirm_all_del = st.text_input("Escribe ELIMINAR TODO para confirmar", key="confirm_all_del_a")

    if st.button("🗑️ Eliminar todas las versiones (excepto original)"):
        if confirm_all_del == "ELIMINAR TODO":
            for subdir, dirs, files in os.walk(VERSIONS_DIR):
                for file in files:
                    ruta_archivo = os.path.join(subdir, file)
                    if file != "Stock_Original.xlsx":
                        os.remove(ruta_archivo)
            st.success("Todas las versiones eliminadas correctamente excepto la original.")
            time.sleep(2)
            st.rerun()
        else:
            st.error("Debes escribir 'ELIMINAR TODO' para confirmar.")

    archivo_subido = st.file_uploader("Selecciona archivo A (.xlsx)", type=["xlsx"], key="uploader_a")
    if archivo_subido:
        if 'uploaded_file_name_a' not in st.session_state or archivo_subido.name != st.session_state['uploaded_file_name_a']:
            st.session_state['uploaded_file_name_a'] = archivo_subido.name
            ruta_actual = obtener_subcarpeta_versiones()
            nombre_archivo_subido = f"Subido_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            ruta_guardado = os.path.join(ruta_actual, nombre_archivo_subido)

            with open(ruta_guardado, "wb") as out_file:
                shutil.copyfileobj(archivo_subido, out_file)

            try:
                data_subida = pd.read_excel(ruta_guardado, sheet_name=None, engine="openpyxl")
                st.session_state["data_dict"] = data_subida
                with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
                    for sheet_name, df_sheet in data_subida.items():
                        df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                st.success(f"✅ Archivo '{nombre_archivo_subido}' importado correctamente en la base de datos A.")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error al procesar el archivo A: {e}")


VERSIONS_DIR_B = "versions_b"

def obtener_subcarpeta_versiones_b():
    zona_local = pytz.timezone('Europe/Madrid')
    ahora = datetime.datetime.now(zona_local)
    nombre_subcarpeta_b = ahora.strftime("%Y_%m_%B")
    ruta_subcarpeta_b = os.path.join(VERSIONS_DIR_B, nombre_subcarpeta_b)
    os.makedirs(ruta_subcarpeta_b, exist_ok=True)
    return ruta_subcarpeta_b

def crear_nueva_version_filename_b():
    ruta_subcarpeta_b = obtener_subcarpeta_versiones_b()
    zona_local = pytz.timezone('Europe/Madrid')
    fh = datetime.datetime.now(zona_local).strftime("%Y-%m-%d_%H-%M-%S")
    return os.path.join(ruta_subcarpeta_b, f"StockB_{fh}.xlsx")

# Explorador visual y subida manual para versiones B
with st.sidebar.expander("🗃️ Gestor avanzado versiones B (Histórico)", expanded=False):
    subcarpetas_b = sorted(
        [d for d in os.listdir(VERSIONS_DIR_B) if os.path.isdir(os.path.join(VERSIONS_DIR_B, d))],
        reverse=True
    )

    if not subcarpetas_b:
        st.info("Aún no hay subcarpetas para versiones B.")
    else:
        mes_elegido_b = st.selectbox("📅 Selecciona el mes (Base B):", subcarpetas_b)
        ruta_actual_b = os.path.join(VERSIONS_DIR_B, mes_elegido_b)
        st.write(f"**Versiones guardadas en {ruta_actual_b}:**")
        archivos_versiones_b = sorted(os.listdir(ruta_actual_b), reverse=True)

        if archivos_versiones_b:
            import datetime
            versiones_b_df = pd.DataFrame({
                "Archivo": archivos_versiones_b,
                "Fecha creación": [
                    datetime.datetime.fromtimestamp(os.path.getctime(os.path.join(ruta_actual_b, f))
                ).strftime('%d/%m/%Y %H:%M:%S') for f in archivos_versiones_b]
            })
            st.dataframe(versiones_b_df)

            version_gestion_b = st.selectbox("Seleccione versión B para gestionar:", archivos_versiones_b)
            ruta_version_b = os.path.join(ruta_actual_b, version_gestion_b)

            col_down_b, col_del_b = st.columns(2)
            with col_down_b:
                with open(ruta_version_b, "rb") as version_file_b:
                    st.download_button(
                        "Descargar versión B",
                        data=version_file_b,
                        file_name=version_gestion_b,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            with col_del_b:
                confirm_eliminar_b = st.text_input("Escribe ELIMINAR para borrar versión B", key="confirm_del_avanzado_b")
                if st.button("Eliminar versión B seleccionada"):
                    if confirm_eliminar_b == "ELIMINAR":
                        os.remove(ruta_version_b)
                        st.success("Versión B eliminada correctamente.")
                        time.sleep(1.5)
                        st.rerun()
                    else:
                        st.error("Debes escribir ELIMINAR para confirmar.")
        else:
            st.info("No hay versiones guardadas en esta subcarpeta B.")

    st.divider()
    st.write("⚠️ **Eliminar TODAS las versiones B excepto la original:**")
    confirm_all_del_b = st.text_input("Escribe ELIMINAR TODO para confirmar", key="confirm_all_del_b")

    if st.button("🗑️ Eliminar todas las versiones B (excepto original)"):
        if confirm_all_del_b == "ELIMINAR TODO":
            for subdir, dirs, files in os.walk(VERSIONS_DIR_B):
                for file in files:
                    ruta_archivo = os.path.join(subdir, file)
                    if file != "Stock_Historico_Original.xlsx":
                        os.remove(ruta_archivo)
            st.success("Todas las versiones B eliminadas correctamente excepto la original.")
            time.sleep(2)
            st.rerun()
        else:
            st.error("Debes escribir 'ELIMINAR TODO' para confirmar.")

    archivo_subido_b = st.file_uploader("Selecciona archivo B (.xlsx)", type=["xlsx"], key="uploader_b")

    if archivo_subido_b:
        if 'uploaded_file_name_b' not in st.session_state or archivo_subido_b.name != st.session_state['uploaded_file_name_b']:
            st.session_state['uploaded_file_name_b'] = archivo_subido_b.name
            ruta_actual_b = obtener_subcarpeta_versiones_b()
            nombre_archivo_subido_b = f"SubidoB_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            ruta_guardado_b = os.path.join(ruta_actual_b, nombre_archivo_subido_b)

            with open(ruta_guardado_b, "wb") as out_file_b:
                shutil.copyfileobj(archivo_subido_b, out_file_b)

            try:
                data_subida_b = pd.read_excel(ruta_guardado_b, sheet_name=None, engine="openpyxl")
                st.session_state["data_dict_b"] = data_subida_b
                with pd.ExcelWriter(STOCK_FILE_B, engine="openpyxl") as writer_b:
                    for sheet_name_b, df_sheet_b in data_subida_b.items():
                        df_sheet_b.to_excel(writer_b, sheet_name=sheet_name_b, index=False)
                st.success(f"✅ Archivo B '{nombre_archivo_subido_b}' importado correctamente en la base de datos B.")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error al procesar el archivo B: {e}")


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

# -------------------------------------------------------------------------
# EXCEL B (Stock_Historico)
# -------------------------------------------------------------------------
STOCK_FILE_B = "Stock_Historico.xlsx"

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

def generar_excel_en_memoria(df_act: pd.DataFrame, sheet_nm="Hoja1"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_act.to_excel(writer, sheet_name=sheet_nm, index=False)
    output.seek(0)
    return output.getvalue()

def enforce_types(df: pd.DataFrame):
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
        df["NºLote"] = df["NºLote"].astype(str).fillna("")
    for col in ["Caducidad","Fecha Pedida","Fecha Llegada"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    if "Sitio almacenaje" in df.columns:
        df["Sitio almacenaje"] = df["Sitio almacenaje"].astype(str)
    if "Caducidad" in df.columns:
        df["Caducidad"] = pd.to_datetime(df["Caducidad"], errors="coerce")
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

if "data_dict" not in st.session_state:
    st.session_state["data_dict"] = load_data_a()

if "data_dict_b" not in st.session_state:
    st.session_state["data_dict_b"] = load_data_b()

data_dict = st.session_state["data_dict"]
data_dict_b = st.session_state["data_dict_b"]

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
    df["GroupCount"] = df["GroupID"].apply(lambda x: group_sizes.get(x, 0))

    unique_ids = sorted(df["GroupID"].unique())
    color_cycle_local = itertools.cycle(colors)
    group_color_map = {}
    for gid in unique_ids:
        group_color_map[gid] = next(color_cycle_local)
    df["ColorGroup"] = df["GroupID"].apply(lambda x: group_color_map.get(x, "#FFFFFF"))

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
    s = row.get("Stock", 0)
    fp = row.get("Fecha Pedida", None)
    if s == 0 and pd.isna(fp):
        return "🔴"
    elif s == 0 and not pd.isna(fp):
        return "🟨"
    return ""

def style_lote(row):
    bg = row.get("ColorGroup", "")
    es_titulo = row.get("EsTitulo", False)
    styles = [f"background-color:{bg}"] * len(row)
    if es_titulo and "Nombre producto" in row.index:
        idx = row.index.get_loc("Nombre producto")
        styles[idx] += "; font-weight:bold"
    return styles

st.markdown(
    """
    <style>
    .big-select select {
        font-size: 18px;
        height: auto;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown("### Información")
st.write("← Recuerde que en la barra lateral puede gestionar las versiones. Despliegue para consultarlo.")
st.divider()

# -------------------------------------------------------------------------
# CUERPO PRINCIPAL => Edición en Hoja Principal (A)
# -------------------------------------------------------------------------
st.header("Gestión del Stock")

if not st.session_state["data_dict"]:
    st.error("No se pudo cargar la base de datos (A).")
    st.stop()

hojas_principales = list(st.session_state["data_dict"].keys())
sheet_name = st.selectbox("Seleccione el panel:", hojas_principales, key="main_sheet_sel")

df_main_original = st.session_state["data_dict"][sheet_name].copy()
df_main_original = enforce_types(df_main_original)

df_for_style = df_main_original.copy()
df_for_style["Alarma"] = df_for_style.apply(calc_alarma, axis=1)
df_for_style = build_group_info_by_ref(df_for_style, panel_default=sheet_name)
df_for_style.sort_values(by=["MultiSort", "GroupID", "NotTitulo"], inplace=True)
df_for_style.reset_index(drop=True, inplace=True)
styled_df = df_for_style.style.apply(style_lote, axis=1)

all_cols = df_for_style.columns.tolist()
cols_to_hide = ["ColorGroup","EsTitulo","GroupCount","MultiSort","Notitulo","GroupID"]
final_cols = [c for c in all_cols if c not in cols_to_hide]

table_html = styled_df.to_html(columns=final_cols)
df_main = df_for_style.copy()
df_main.drop(columns=cols_to_hide, inplace=True, errors="ignore")

st.write(f"#### Stock del Panel {sheet_name}")
st.write(table_html, unsafe_allow_html=True)

if "Nombre producto" in df_main.columns and "Ref. Fisher" in df_main.columns:
    display_series = df_main.apply(lambda r: f"{r['Nombre producto']} ({r['Ref. Fisher']})", axis=1)
else:
    display_series = df_main.iloc[:, 0].astype(str)

reactivo_sel = st.selectbox("Seleccione Reactivo a Modificar:", display_series.unique(), key="react_modif")
row_index = display_series[display_series == reactivo_sel].index[0]

st.write("**Recuerde que no es necesario ingresar la fecha pedida si se está ingresando la fecha llegada**")

def get_val(col, default=None):
    return df_main.at[row_index, col] if col in df_main.columns else default

lote_actual = get_val("NºLote", "")
caducidad_actual = get_val("Caducidad", None)
fecha_pedida_actual = get_val("Fecha Pedida", None)
fecha_llegada_actual = get_val("Fecha Llegada", None)
sitio_almacenaje_actual = get_val("Sitio almacenaje", "")
uds_actual = get_val("Uds.", 0)
stock_actual = get_val("Stock", 0)

colA, colB, colC, colD = st.columns([1, 1, 1, 1])

with colA:
    lote_new = st.text_input("Nº de Lote", value=str(lote_actual))
    cad_new = st.date_input("Caducidad", value=caducidad_actual if pd.notna(caducidad_actual) else None)

with colB:
    fped_date = st.date_input(
        "Fecha Pedida",
        value=fecha_pedida_actual.date() if pd.notna(fecha_pedida_actual) else None,
        key="fped_date_main",
    )
    fped_time = st.time_input(
        "Hora Pedida (opcional)",
        value=fecha_pedida_actual.time() if pd.notna(fecha_pedida_actual) else datetime.time(0,0),
        key="fped_time_main",
    )

with colC:
    flleg_date = st.date_input(
        "Fecha Llegada",
        value=fecha_llegada_actual.date() if pd.notna(fecha_llegada_actual) else None,
        key="flleg_date_main",
    )
    flleg_time = st.time_input(
        "Hora Llegada (opcional)",
        value=fecha_llegada_actual.time() if pd.notna(fecha_llegada_actual) else datetime.time(0,0),
        key="flleg_time_main",
    )

with colD:
    st.write("")
    st.write("")
    if st.button("Refrescar Página"):
        st.rerun()

comentario_actual = ""
if "Comentario" in df_main.columns:
    comentario_actual = str(df_main.at[row_index, "Comentario"])
comentario_nuevo = st.text_area("Comentario (opcional)", value=comentario_actual, key="comentario_input_key")

# Para usar zona horaria local y guardar como string:
zone = pytz.timezone("Europe/Madrid")

# Procesar Fecha Pedida
fped_new = None
if fped_date is not None:
    dt_ped = datetime.datetime.combine(fped_date, fped_time)
    dt_ped_local = zone.localize(dt_ped)  # localizamos con DST
    fped_new_str = dt_ped_local.strftime("%Y-%m-%d %H:%M:%S")
else:
    fped_new_str = None

# Procesar Fecha Llegada
flleg_new_str = None
if flleg_date is not None:
    dt_lleg = datetime.datetime.combine(flleg_date, flleg_time)
    dt_lleg_local = zone.localize(dt_lleg)
    flleg_new_str = dt_lleg_local.strftime("%Y-%m-%d %H:%M:%S")

# Preparar la lógica de "group_order_selected"
group_id = df_for_style.at[row_index, "GroupID"]
group_reactivos = df_for_style[df_for_style["GroupID"] == group_id]
group_order_selected = None
if fped_new_str is not None:
    if not group_reactivos.empty:
        if group_reactivos["EsTitulo"].any():
            lot_name = group_reactivos[group_reactivos["EsTitulo"]==True]["Nombre producto"].iloc[0]
        else:
            lot_name = f"Ref. Saturno {group_id}"
        group_reactivos_reset = group_reactivos.reset_index()
        options = group_reactivos_reset.apply(
            lambda r: f"{r['index']} - {r['Nombre producto']} ({r['Ref. Fisher']})", axis=1
        ).tolist()
        st.markdown('<div class="big-select">', unsafe_allow_html=True)
        group_order_selected = st.multiselect(
            f"¿Pedir también los siguientes reactivos del lote **{lot_name}**?",
            options, default=options
        )
        st.markdown('</div>', unsafe_allow_html=True)

if st.button("Guardar Cambios en Hoja Stock"):
    # No forzamos a None la fecha pedida cuando hay fecha llegada
    # (a menos que tu lógica lo exija; aquí lo evitamos para retener las horas exactas).
    # Evitamos poner: if pd.notna(flleg_new_str): fped_new_str = None

    if "Stock" in df_main.columns:
        # Si el usuario modificó la fecha de llegada (o cambió el lote),
        # sumamos uds_actual al stock_actual
        if (
            (flleg_new_str != fecha_llegada_actual and flleg_new_str is not None)
            or
            (lote_new != lote_actual and lote_new.strip() != "")
        ):
            df_main.at[row_index, "Stock"] = stock_actual + uds_actual
            st.info(f"Añadidas {uds_actual} uds => stock={stock_actual + uds_actual}")

    # Actualizar Lote, Caducidad, etc. en df_main
    if "NºLote" in df_main.columns:
        df_main.at[row_index, "NºLote"] = str(lote_new)
    if "Caducidad" in df_main.columns:
        df_main.at[row_index, "Caducidad"] = cad_new if pd.notna(cad_new) else pd.NaT

    # Guardar fecha pedida como string
    if "Fecha Pedida" in df_main.columns:
        df_main.at[row_index, "Fecha Pedida"] = fped_new_str

    # Guardar fecha llegada como string
    if "Fecha Llegada" in df_main.columns:
        df_main.at[row_index, "Fecha Llegada"] = flleg_new_str

    if "Sitio almacenaje" in df_main.columns:
        df_main.at[row_index, "Sitio almacenaje"] = sitio_new

    if "Comentario" not in df_main.columns:
        df_main["Comentario"] = ""
    df_main.at[row_index, "Comentario"] = comentario_nuevo

    # Actualizamos "Fecha Pedida" para todos los reactivos seleccionados del grupo
    if fped_new_str is not None:
        if not group_order_selected:
            group_order_selected = options  # si el usuario no seleccionó nada, usamos todo
        for label in group_order_selected:
            try:
                i_val = int(label.split(" - ")[0])
                df_main.at[i_val, "Fecha Pedida"] = fped_new_str
            except Exception as e:
                st.error(f"Error actualizando índice {label}: {e}")

    # Guardamos df_main en st.session_state
    st.session_state["data_dict"][sheet_name] = df_main

    # Crear nueva versión A
    new_file = crear_nueva_version_filename()
    with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
        for sht, df_sht in st.session_state["data_dict"].items():
            ocultar_cols = ["ColorGroup", "EsTitulo", "GroupCount", "MultiSort", "NotTitulo", "GroupID", "Alarma"]
            tmp = df_sht.drop(columns=ocultar_cols, errors="ignore")
            tmp.to_excel(writer, sheet_name=sht, index=False)

    # Sobrescribir STOCK_FILE
    with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
        for sht, df_sht in st.session_state["data_dict"].items():
            ocultar_cols = ["ColorGroup", "EsTitulo", "GroupCount", "MultiSort", "NotTitulo", "GroupID", "Alarma"]
            tmp = df_sht.drop(columns=ocultar_cols, errors="ignore")
            tmp.to_excel(writer, sheet_name=sht, index=False)

    # Insertar registro en B (misma hoja)
    if sheet_name not in st.session_state["data_dict_b"]:
        st.session_state["data_dict_b"][sheet_name] = pd.DataFrame()

    df_b_sh = st.session_state["data_dict_b"][sheet_name].copy()
    nueva_fila = {
        "Ref. Saturno": df_main.at[row_index, "Ref. Saturno"] if "Ref. Saturno" in df_main.columns else 0,
        "Ref. Fisher": df_main.at[row_index, "Ref. Fisher"] if "Ref. Fisher" in df_main.columns else "",
        "Nombre producto": df_main.at[row_index, "Nombre producto"] if "Nombre producto" in df_main.columns else "",
        "NºLote": df_main.at[row_index, "NºLote"],
        "Caducidad": df_main.at[row_index, "Caducidad"],
        "Fecha Pedida": df_main.at[row_index, "Fecha Pedida"],
        "Fecha Llegada": df_main.at[row_index, "Fecha Llegada"],
        "Sitio almacenaje": df_main.at[row_index, "Sitio almacenaje"],
        "Uds.": df_main.at[row_index, "Uds."] if "Uds." in df_main.columns else 0,
        "Stock": df_main.at[row_index, "Stock"] if "Stock" in df_main.columns else 0,
        "Comentario": df_main.at[row_index, "Comentario"] if "Comentario" in df_main.columns else "",
        "Fecha Registro B": datetime.datetime.now()
    }
    df_b_sh = pd.concat([df_b_sh, pd.DataFrame([nueva_fila])], ignore_index=True)
    st.session_state["data_dict_b"][sheet_name] = df_b_sh

    # Crear nueva versión B
    new_file_b = crear_nueva_version_filename_b()
    with pd.ExcelWriter(new_file_b, engine="openpyxl") as writerB:
        for shtB, df_shtB in st.session_state["data_dict_b"].items():
            df_shtB.to_excel(writerB, sheet_name=shtB, index=False)

    # Sobrescribir STOCK_FILE_B
    with pd.ExcelWriter(STOCK_FILE_B, engine="openpyxl") as writerB:
        for shtB, df_shtB in st.session_state["data_dict_b"].items():
            df_shtB.to_excel(writerB, sheet_name=shtB, index=False)

    st.success(
        f"✅ Cambios guardados en '{new_file}' y '{STOCK_FILE}' (A). "
        f"También en '{new_file_b}' y '{STOCK_FILE_B}' (B)."
    )
    excel_bytes = generar_excel_en_memoria(df_main, sheet_nm=sheet_name)
    st.download_button(
        "Descargar Excel A modificado",
        excel_bytes,
        "Reporte_Stock.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    time.sleep(2)
    st.rerun()


st.divider()
st.divider()

# -------------------------------------------------------------------------
# AGRUPAR EN TABS: Ver Base de Datos Historial (B), Filtrar Reactivos, e Informar Reactivo Agotado
# -------------------------------------------------------------------------
tabs = st.tabs([
    "Ver Base de Datos Historial (B)",
    "Filtrar Reactivos Limitantes/Compartidos",
    "Informar Reactivo Agotado"
])

with tabs[0]:
    st.write("### Vista de la Base de Datos Historial (B)")
    if st.session_state["data_dict_b"]:
        hojas_b = list(st.session_state["data_dict_b"].keys())
        hoja_b_sel = st.selectbox("Seleccione hoja en B (vista):", hojas_b, key="vista_tab")
        df_b_vista = st.session_state["data_dict_b"][hoja_b_sel].copy()
        if "Nombre producto" in df_b_vista.columns and "NºLote" in df_b_vista.columns:
            df_b_vista.sort_values(by=["Nombre producto","NºLote"], inplace=True, ignore_index=True)
        st.dataframe(df_b_vista)
        excel_b_mem = generar_excel_en_memoria(df_b_vista, sheet_nm=hoja_b_sel)
        st.download_button(
            label="Descargar hoja de Excel B (vista)",
            data=excel_b_mem,
            file_name="Hoja_Historico_B_vista.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.write("No hay datos en la Base Historial (B).")

with tabs[1]:
    st.write("### Filtrar Reactivos Limitantes/Compartidos")

    limitantes_set = {
        "A42006","A42007","A27762","A34018","A33638","A33639","A27758","A27765","A4517",
        "A3410","A34537","A45617","A34540","A36410","A29025","A29027","A29026","A27754",
        "11754050","11766050"
    }

    if not st.session_state["data_dict_b"]:
        st.warning("No hay datos en base B. Verifica que Stock_Historico.xlsx tenga contenido.")
        st.stop()

    all_rows_b = []
    for sheet_b, df_b_sht in st.session_state["data_dict_b"].items():
        temp_df = df_b_sht.copy()
        temp_df["(Hoja B)"] = sheet_b
        all_rows_b.append(temp_df)

    df_b_combined = pd.concat(all_rows_b, ignore_index=True)
    df_b_combined = enforce_types(df_b_combined)

    limitantes_list = []
    compartidos_list = []
    for idx, row in df_b_combined.iterrows():
        ref = str(row.get("Ref. Fisher", "")).strip()
        nom = str(row.get("Nombre producto", "")).strip()
        hoja = str(row.get("(Hoja B)", "")).strip()

        if not ref and not nom:
            continue
        if ref.upper() == "A":
            continue

        if ref in limitantes_set:
            limitantes_list.append((ref, nom, hoja))
        else:
            compartidos_list.append((ref, nom))

    limitantes_unique = set(limitantes_list)
    compartidos_unique = set(compartidos_list)
    limitantes_list = list(limitantes_unique)
    compartidos_list = list(compartidos_unique)

    grupo_elegido = st.radio(
        "¿Qué grupo de reactivos quiere filtrar?",
        ("limitante", "compartido"),
        key="grupo_filtrar"
    )

    if grupo_elegido == "limitante":
        op_list = limitantes_list
    else:
        op_list = compartidos_list

    if not op_list:
        st.warning(f"No se encontraron reactivos en la categoría '{grupo_elegido}' dentro de la base B.")
        st.stop()

    def display_label_limit(tup):
        return f"{tup[0]} - {tup[1]} ({tup[2]})"

    def display_label_comp(tup):
        return f"{tup[0]} - {tup[1]}"

    if grupo_elegido == "limitante":
        seleccion = st.selectbox(
            "Seleccione Reactivo (Limitante)",
            [display_label_limit(x) for x in op_list],
            key="select_b_filtrado_tab"
        )
    else:
        seleccion = st.selectbox(
            "Seleccione Reactivo (Compartido)",
            [display_label_comp(x) for x in op_list],
            key="select_b_filtrado_tab"
        )

    if st.button("Buscar en Base Historial", key="buscar_filtrado"):
        if grupo_elegido == "limitante":
            i_sel = [display_label_limit(x) for x in op_list].index(seleccion)
            ref_sel, nom_sel, hoja_selB = op_list[i_sel]
        else:
            i_sel = [display_label_comp(x) for x in op_list].index(seleccion)
            ref_sel, nom_sel = op_list[i_sel]
            hoja_selB = None

        df_filtrado = df_b_combined[df_b_combined["Ref. Fisher"].astype(str).str.strip() == ref_sel].copy()

        if "Caducidad" in df_filtrado.columns:
            df_filtrado = df_filtrado.dropna(subset=["Caducidad"])

        if df_filtrado.empty:
            st.warning("No se encontraron reactivos (en B) con esos parámetros.")
        else:
            if "Caducidad" in df_filtrado.columns:
                df_filtrado.sort_values(by="Caducidad", inplace=True, ignore_index=True)

            if ref_sel == "A27754":
                st.info("Nota: Esta referencia se comparte entre OCA y FOCUS.")

            st.dataframe(df_filtrado)
            excel_filtro = generar_excel_en_memoria(df_filtrado, "Filtro_B")
            st.download_button(
                label="Descargar resultados filtrados en Excel",
                data=excel_filtro,
                file_name="Filtro_B.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

with tabs[2]:
    st.write("### Informar Reactivo Agotado")
    if not st.session_state["data_dict"]:
        st.error("No se pudo cargar la base de datos (A).")
        st.stop()

    hojas_a = list(st.session_state["data_dict"].keys())
    hoja_sel = st.selectbox("Hoja A a consumir:", hojas_a, key="agotado_hoja")
    df_a = st.session_state["data_dict"][hoja_sel].copy()
    df_a = enforce_types(df_a)

    if "Nombre producto" not in df_a.columns:
        st.error("No existe columna 'Nombre producto' en esta hoja A.")
        st.stop()

    df_a["nombre_ref"] = df_a["Nombre producto"].astype(str) + " (" + df_a["Ref. Fisher"].astype(str) + ")"
    nombre_ref_unicos = sorted(df_a["nombre_ref"].dropna().unique())
    nombre_ref_sel = st.selectbox("Nombre producto en A (Ref. Fisher):", nombre_ref_unicos, key="agotado_nombre")

    nombre_sel = nombre_ref_sel.rsplit(" (", 1)[0].strip()
    ref_sel = nombre_ref_sel.rsplit(" (", 1)[1].replace(")", "").strip()

    df_cand = df_a[
        (df_a["Nombre producto"] == nombre_sel) &
        (df_a["Ref. Fisher"] == ref_sel)
    ]

    if df_cand.empty:
        st.warning("No se encontró ese nombre en esta hoja A.")
    else:
        idx_c = df_cand.index[0]
        stock_c = df_a.at[idx_c, "Stock"] if "Stock" in df_a.columns else 0
        uds_consumir = st.number_input("Uds. a consumir en A:", min_value=0, step=1, key="agotado_uds")
        if st.button("Consumir en Lab (memoria)", key="agotado_consumir"):
            nuevo_stock = max(0, stock_c - uds_consumir)
            df_a.at[idx_c, "Stock"] = nuevo_stock
            if nuevo_stock == 0:
                for col_vaciar in ["NºLote","Caducidad","Fecha Pedida","Fecha Llegada","Sitio almacenaje"]:
                    if col_vaciar in df_a.columns:
                        if col_vaciar in ["Caducidad","Fecha Pedida","Fecha Llegada"]:
                            df_a.at[idx_c, col_vaciar] = pd.NaT
                        else:
                            df_a.at[idx_c, col_vaciar] = ""
            st.session_state["data_dict"][hoja_sel] = df_a
            st.warning(f"Consumidas {uds_consumir} uds. Stock final => {nuevo_stock}. (Sólo en memoria).")

    st.write("**Eliminar en B** => introduce el Lote exacto. Si coincide Nombre+Lote, se borra de B.")
    lote_b = st.number_input("Nº de Lote (en B)", min_value=0, step=1, key="agotado_lote")

    if st.button("Guardar Cambios en Consumo Lab", key="agotado_guardar"):
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

        if hoja_sel in st.session_state["data_dict_b"]:
            df_b_hoja = st.session_state["data_dict_b"][hoja_sel].copy()
            if "Nombre producto" in df_b_hoja.columns and "NºLote" in df_b_hoja.columns:
                df_b_hoja = df_b_hoja[~(
                    (df_b_hoja["Nombre producto"] == nombre_sel) &
                    (df_b_hoja["NºLote"] == str(lote_b))
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
        time.sleep(2)
        st.rerun()
