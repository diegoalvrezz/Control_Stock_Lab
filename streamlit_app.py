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

st.set_page_config(page_title="Control de Stock con Lotes", layout="centered")
st.title("🔬 Control Stock Lab. Patología Molécular")

def calc_alarma(row):
    """
    Devuelve un string con un ícono de alerta:
    - '🔴' si el Stock es 0 y no hay Fecha Pedida.
    - '🟨' si el Stock es 0 y sí hay Fecha Pedida.
    - '' (vacío) si no hay que alertar.
    """
    stock_val = row.get("Stock", 0)
    fecha_pedida = row.get("Fecha Pedida", None)

    if stock_val == 0 and pd.isna(fecha_pedida):
        return "🔴"
    elif stock_val == 0 and not pd.isna(fecha_pedida):
        return "🟨"
    return ""


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
    """
    TODO: Ajustar la lógica de lectura del Excel principal (Stock_Original).
    Retorna un dict con { hoja_name: DataFrame }.
    """
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

def load_data_b():
    """
    TODO: Ajustar la lógica de lectura del Excel histórico (Stock_Historico).
    Retorna un dict con { hoja_name: DataFrame }.
    """
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
# Inicializamos en session_state
# -------------------------------------------------------------------------
if "data_dict" not in st.session_state:
    st.session_state["data_dict"] = load_data_a()

if "data_dict_b" not in st.session_state:
    st.session_state["data_dict_b"] = load_data_b()

# -------------------------------------------------------------------------
# Funciones para crear subdirectorios Año/Mes/Día y listarlos
# -------------------------------------------------------------------------
def create_dated_version_filename(base_dir, prefix="Stock"):
    """
    Crea subdirectorios en base a Año/Mes/Día dentro de `base_dir`,
    y devuelve la ruta completa para guardar un nuevo archivo Excel.

    Ejemplo de ruta generada:
       versions/2025/03/14/Stock_2025-03-14_14-09-50.xlsx
    """
    now = datetime.datetime.now()
    year_str = str(now.year)
    month_str = now.strftime("%m")  # '01'..'12'
    day_str = now.strftime("%d")    # '01'..'31'

    year_dir = os.path.join(base_dir, year_str)
    month_dir = os.path.join(year_dir, month_str)
    day_dir = os.path.join(month_dir, day_str)

    os.makedirs(day_dir, exist_ok=True)

    timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"{prefix}_{timestamp}.xlsx"
    return os.path.join(day_dir, filename)

def list_files_by_date(base_dir):
    """
    Recorre recursivamente el directorio base_dir organizado en subcarpetas Año/Mes/Día,
    y retorna una estructura anidada, por ejemplo:
    {
      '2025': {
         '03': {
            '14': [ 'Stock_2025-03-14_14-09-50.xlsx', ... ],
            ...
         },
         ...
      },
    }
    """
    date_structure = {}
    if not os.path.exists(base_dir):
        return date_structure

    for year in sorted(os.listdir(base_dir)):
        year_path = os.path.join(base_dir, year)
        if not os.path.isdir(year_path):
            continue
        if year not in date_structure:
            date_structure[year] = {}

        for month in sorted(os.listdir(year_path)):
            month_path = os.path.join(year_path, month)
            if not os.path.isdir(month_path):
                continue
            if month not in date_structure[year]:
                date_structure[year][month] = {}

            for day in sorted(os.listdir(month_path)):
                day_path = os.path.join(month_path, day)
                if not os.path.isdir(day_path):
                    continue
                if day not in date_structure[year][month]:
                    date_structure[year][month][day] = []

                for fname in sorted(os.listdir(day_path)):
                    fpath = os.path.join(day_path, fname)
                    if os.path.isfile(fpath) and fname.endswith(".xlsx"):
                        date_structure[year][month][day].append(fname)

    return date_structure

# -------------------------------------------------------------------------
# Otras funciones comunes
# -------------------------------------------------------------------------
def generar_excel_en_memoria(df_act: pd.DataFrame, sheet_nm="Hoja1"):
    """
    Crea un Excel en memoria con df_act, para descargarlo con st.download_button
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_act.to_excel(writer, sheet_name=sheet_nm, index=False)
    output.seek(0)
    return output.getvalue()

def enforce_types(df: pd.DataFrame):
    # Ajusta según tus columnas
    if "Ref. Saturno" in df.columns:
        df["Ref. Saturno"] = pd.to_numeric(df["Ref. Saturno"], errors="coerce").fillna(0).astype(int)
    if "Ref. Fisher" in df.columns:
        df["Ref. Fisher"] = df["Ref. Fisher"].astype(str)
    if "Nombre producto" in df.columns:
        df["Nombre producto"] = df["Nombre producto"].astype(str)
    if "NºLote" in df.columns:
        df["NºLote"] = df["NºLote"].astype(str).fillna("")
    for col in ["Caducidad","Fecha Pedida","Fecha Llegada"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    if "Sitio almacenaje" in df.columns:
        df["Sitio almacenaje"] = df["Sitio almacenaje"].astype(str)
    if "Stock" in df.columns:
        df["Stock"] = pd.to_numeric(df["Stock"], errors="coerce").fillna(0).astype(int)
    return df

# -------------------------------------------------------------------------
# Barra lateral => Ver/Gestionar versiones Stock (A)
# -------------------------------------------------------------------------
with st.sidebar.expander("🔎 Ver / Gestionar versiones Stock (A)", expanded=False):
    if st.session_state["data_dict"]:
        structureA = list_files_by_date(VERSIONS_DIR)
        if not structureA:
            st.write("No hay versiones guardadas en subcarpetas (excepto la original).")
        else:
            yearsA = sorted(structureA.keys())
            selected_year_A = st.selectbox("Año disponible (A)", yearsA, key="yearA")
            monthsA = sorted(structureA[selected_year_A].keys())
            selected_month_A = st.selectbox("Mes disponible (A)", monthsA, key="monthA")
            daysA = sorted(structureA[selected_year_A][selected_month_A].keys())
            selected_day_A = st.selectbox("Día disponible (A)", daysA, key="dayA")
            files_day_A = structureA[selected_year_A][selected_month_A][selected_day_A]
            if files_day_A:
                version_selA = st.selectbox("Seleccione versión A:", files_day_A, key="fileA")
                if version_selA:
                    file_pathA = os.path.join(VERSIONS_DIR, selected_year_A, selected_month_A, selected_day_A, version_selA)
                    if os.path.isfile(file_pathA):
                        with open(file_pathA, "rb") as excel_file:
                            excel_bytes = excel_file.read()
                        st.download_button(
                            label=f"Descargar {version_selA}",
                            data=excel_bytes,
                            file_name=version_selA,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        confirm_text_A = st.text_input(
                            f"Para confirmar la eliminación de '{version_selA}', escribe 'ELIMINAR'",
                            key="confirm_elim_A"
                        )
                        if st.button("Eliminar esta versión A"):
                            if confirm_text_A == "ELIMINAR":
                                try:
                                    os.remove(file_pathA)
                                    st.warning(f"Versión '{version_selA}' eliminada.")
                                    time.sleep(2)
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Error al intentar eliminar la versión: {e}")
                            else:
                                st.error("Confirme la eliminación escribiendo 'ELIMINAR'.")
            else:
                st.write("No hay versiones guardadas en este día.")

        if st.button("Limpiar Base de Datos A"):
            original_path = os.path.join(VERSIONS_DIR, "Stock_Original.xlsx")
            if os.path.exists(original_path):
                shutil.copy(original_path, STOCK_FILE)
                st.success("Base de datos A restaurada al estado original.")
                st.session_state["data_dict"] = load_data_a()
                time.sleep(2)
                st.rerun()
            else:
                st.error("No se encontró la copia original de A.")
    else:
        st.error("No hay data_dict. Verifica Stock_Original.xlsx.")
        st.stop()

# -------------------------------------------------------------------------
# Barra lateral => Ver/Gestionar versiones Historial (B)
# -------------------------------------------------------------------------
with st.sidebar.expander("🔎 Ver / Gestionar versiones Historial (B)", expanded=False):
    if st.session_state["data_dict_b"]:
        structureB = list_files_by_date(VERSIONS_DIR_B)
        if not structureB:
            st.write("No hay versiones guardadas en subcarpetas (excepto la original).")
        else:
            yearsB = sorted(structureB.keys())
            selected_year_B = st.selectbox("Año disponible (B)", yearsB, key="yearB")
            monthsB = sorted(structureB[selected_year_B].keys())
            selected_month_B = st.selectbox("Mes disponible (B)", monthsB, key="monthB")
            daysB = sorted(structureB[selected_year_B][selected_month_B].keys())
            selected_day_B = st.selectbox("Día disponible (B)", daysB, key="dayB")
            files_day_B = structureB[selected_year_B][selected_month_B][selected_day_B]
            if files_day_B:
                version_selB = st.selectbox("Seleccione versión B:", files_day_B, key="fileB")
                if version_selB:
                    file_pathB = os.path.join(VERSIONS_DIR_B, selected_year_B, selected_month_B, selected_day_B, version_selB)
                    if os.path.isfile(file_pathB):
                        with open(file_pathB, "rb") as excel_file_b:
                            excel_bytes_b = excel_file_b.read()
                        st.download_button(
                            label=f"Descargar {version_selB}",
                            data=excel_bytes_b,
                            file_name=version_selB,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        confirm_text_B = st.text_input(
                            f"Para confirmar la eliminación de '{version_selB}', escribe 'ELIMINAR'",
                            key="confirm_elim_B"
                        )
                        if st.button("Eliminar esta versión B"):
                            if confirm_text_B == "ELIMINAR":
                                try:
                                    os.remove(file_pathB)
                                    st.warning(f"Versión '{version_selB}' eliminada de B.")
                                    time.sleep(2)
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Error al intentar eliminar la versión: {e}")
                            else:
                                st.error("Confirme la eliminación escribiendo 'ELIMINAR'.")
            else:
                st.write("No hay versiones guardadas en este día.")

        if st.button("Limpiar Base de Datos B"):
            original_path_b = os.path.join(VERSIONS_DIR_B, "Stock_Historico_Original.xlsx")
            if os.path.exists(original_path_b):
                shutil.copy(original_path_b, STOCK_FILE_B)
                st.success("Base de datos B restaurada al estado original.")
                st.session_state["data_dict_b"] = load_data_b()
                time.sleep(2)
                st.rerun()
            else:
                st.error("No se encontró la copia original de B.")
    else:
        st.write("No hay data_dict_b. Verifica Stock_Historico.xlsx.")

st.markdown("### Información")
st.write("← Recuerde que en la barra lateral puede gestionar las versiones. Despliegue para consultarlo.")
st.divider()

# ==============================================================================
# EJEMPLO DE FUNCIÓN: GUARDAR NUEVA VERSIÓN A
# ==============================================================================
def guardar_nueva_version_A():
    """
    Ejemplo a modo de demostración. 
    Imagina que en el botón 'Guardar Cambios' de tu app,
    llamas a esta función para crear la ruta y guardar el Excel.
    """
    new_file = create_dated_version_filename(VERSIONS_DIR, prefix="Stock")
    # TODO: Guardar tu DF principal en new_file
    # with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
    #     df_main.to_excel(writer, index=False, sheet_name="HojaA")
    st.success(f"Archivo guardado en {new_file}")

# ==============================================================================
# EJEMPLO DE FUNCIÓN: GUARDAR NUEVA VERSIÓN B
# ==============================================================================
def guardar_nueva_version_B():
    """
    Similar para la base de datos histórica B.
    """
    new_file_b = create_dated_version_filename(VERSIONS_DIR_B, prefix="StockB")
    # TODO: Guardar tu DF histórico en new_file_b
    st.success(f"Archivo histórico guardado en {new_file_b}")


# -------------------------------------------------------------------------
# Aquí iría el resto de tu lógica principal de la app, por ejemplo:
# "Gestión del Stock", "Informar Reactivo Agotado", etc.
# -------------------------------------------------------------------------

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
df_for_style.sort_values(by=["MultiSort","GroupID","NotTitulo"], inplace=True)
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
    display_series = df_main.iloc[:,0].astype(str)
reactivo_sel = st.selectbox("Seleccione Reactivo a Modificar:", display_series.unique(), key="react_modif")
row_index = display_series[display_series == reactivo_sel].index[0]
st.write("**Recuerde que no es necesario ingresar la fecha pedida si se está ingresando la fecha llegada**")
def get_val(col, default=None):
    return df_main.at[row_index, col] if col in df_main.columns else default
lote_actual = get_val("NºLote",0)
caducidad_actual = get_val("Caducidad",None)
fecha_pedida_actual = get_val("Fecha Pedida",None)
fecha_llegada_actual = get_val("Fecha Llegada",None)
sitio_almacenaje_actual = get_val("Sitio almacenaje","")
uds_actual = get_val("Uds.",0)
stock_actual = get_val("Stock",0)
colA, colB, colC, colD = st.columns([1,1,1,1])
with colA:
    lote_new = st.text_input("Nº de Lote", value=str(lote_actual))
    cad_new = st.date_input("Caducidad", value=caducidad_actual if pd.notna(caducidad_actual) else None)
with colB:
    fped_date = st.date_input("Fecha Pedida",
                              value=fecha_pedida_actual.date() if pd.notna(fecha_pedida_actual) else None,
                              key="fped_date_main")
    fped_time = st.time_input("Hora Pedidab (opcional)",
                              value=fecha_pedida_actual.time() if pd.notna(fecha_pedida_actual) else datetime.time(0,0),
                              key="fped_time_main")
with colC:
    flleg_date = st.date_input("Fecha Llegada",
                               value=fecha_llegada_actual.date() if pd.notna(fecha_llegada_actual) else None,
                               key="flleg_date_main")
    flleg_time = st.time_input("Hora Llegada (opcional)",
                               value=fecha_llegada_actual.time() if pd.notna(fecha_llegada_actual) else datetime.time(0,0),
                               key="flleg_time_main")
with colD:
    st.write("")
    st.write("")
    if st.button("Refrescar Página"):
        
        st.rerun()
# ---------------------------
# AÑADIMOS AQUÍ UN CAMPO PARA "Comentario"
# ---------------------------
comentario_actual = ""
if "Comentario" in df_main.columns:
    comentario_actual = str(df_main.at[row_index, "Comentario"])
comentario_nuevo = st.text_area(
    label="Comentario (opcional)",
    value=comentario_actual,
    key="comentario_input_key"
)
fped_new = None
if fped_date is not None:
    dt_ped = datetime.datetime.combine(fped_date, fped_time)
    fped_new = pd.to_datetime(dt_ped)
flleg_new = None

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

if flleg_date is not None:
    dt_lleg = datetime.datetime.combine(flleg_date, flleg_time)
    flleg_new = pd.to_datetime(dt_lleg)
st.write("#### Lugar de Almacenaje")
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

if st.button("Guardar Cambios en Hoja Stock"):
    if pd.notna(flleg_new):
        fped_new = pd.NaT

    if "Stock" in df_main.columns:
        # Si el usuario modificó la fecha de llegada (o cambió el lote), sumamos uds_actual al stock_actual
        if (
            (flleg_new != fecha_llegada_actual and pd.notna(flleg_new))
            or
            (lote_new != lote_actual and lote_new.strip() != "")
        ):
            df_main.at[row_index, "Stock"] = stock_actual + uds_actual
            st.info(f"Añadidas {uds_actual} uds => stock={stock_actual + uds_actual}")

    if "NºLote" in df_main.columns:
        df_main.at[row_index, "NºLote"] = str(lote_new)
    if "Caducidad" in df_main.columns:
        df_main.at[row_index, "Caducidad"] = cad_new if pd.notna(cad_new) else pd.NaT
    if "Fecha Pedida" in df_main.columns:
        df_main.at[row_index, "Fecha Pedida"] = fped_new
    if "Fecha Llegada" in df_main.columns:
        df_main.at[row_index, "Fecha Llegada"] = flleg_new
    if "Sitio almacenaje" in df_main.columns:
        df_main.at[row_index, "Sitio almacenaje"] = sitio_new
    if "Comentario" not in df_main.columns:
        df_main["Comentario"] = ""
    df_main.at[row_index, "Comentario"] = comentario_nuevo

    # Actualizamos "Fecha Pedida" para todos los reactivos seleccionados.
    # Si el multiselect quedó vacío, usamos todas las opciones generadas (es decir, actualizamos todo el grupo).
    if pd.notna(fped_new):
        if not group_order_selected:
            # group_order_selected se obtiene anteriormente en el multiselect.
            # En caso de que esté vacío, usamos todas las opciones (almacenadas en 'options')
            group_order_selected = options  # 'options' se definió al crear el multiselect.
        for label in group_order_selected:
            try:
                i_val = int(label.split(" - ")[0])
                df_main.at[i_val, "Fecha Pedida"] = fped_new
            except Exception as e:
                st.error(f"Error actualizando índice {label}: {e}")

    st.session_state["data_dict"][sheet_name] = df_main

    new_file = crear_nueva_version_filename()
    with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
        for sht, df_sht in st.session_state["data_dict"].items():
            ocultar = ["ColorGroup", "EsTitulo", "GroupCount", "MultiSort", "NotTitulo", "GroupID", "Alarma"]
            tmp = df_sht.drop(columns=ocultar, errors="ignore")
            tmp.to_excel(writer, sheet_name=sht, index=False)
    with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
        for sht, df_sht in st.session_state["data_dict"].items():
            ocultar = ["ColorGroup", "EsTitulo", "GroupCount", "MultiSort", "NotTitulo", "GroupID", "Alarma"]
            tmp = df_sht.drop(columns=ocultar, errors="ignore")
            tmp.to_excel(writer, sheet_name=sht, index=False)

    # Insertar en B
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
    
    time.sleep(2)
    st.rerun()


st.divider()
st.divider()
# -------------------------------------------------------------------------
# AGRUPAR EN TABS: Ver Base de Datos Historial (B), Filtrar Reactivos, e Informar Reactivo Agotado
# -------------------------------------------------------------------------
tabs = st.tabs(["Ver Base de Datos Historial (B)", "Filtrar Reactivos Limitantes/Compartidos", "Informar Reactivo Agotado"])

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

    # 1) Set de referencias \"limitantes\"
    limitantes_set = {
        "A42006","A42007","A27762","A34018","A33638","A33639","A27758","A27765","A4517",
        "A3410","A34537","A45617","A34540","A36410","A29025","A29027","A29026","A27754",
        "11754050","11766050"
    }

    if not st.session_state["data_dict_b"]:
        st.warning("No hay datos en base B. Verifica que Stock_Historico.xlsx tenga contenido.")
        st.stop()

    # Concatenar todas las hojas en un DF combinado
    all_rows_b = []
    for sheet_b, df_b_sht in st.session_state["data_dict_b"].items():
        temp_df = df_b_sht.copy()
        temp_df["(Hoja B)"] = sheet_b
        all_rows_b.append(temp_df)

    df_b_combined = pd.concat(all_rows_b, ignore_index=True)
    df_b_combined = enforce_types(df_b_combined)

    # 2) Recorremos df_b_combined para extraer datos
    #    y creamos dos estructuras distintas:
    #    - limitantes_list => (ref, nom, hoja)
    #    - compartidos_list => (ref, nom)
    limitantes_list = []
    compartidos_list = []

    for idx, row in df_b_combined.iterrows():
        ref = str(row.get("Ref. Fisher", "")).strip()
        nom = str(row.get("Nombre producto", "")).strip()
        hoja = str(row.get("(Hoja B)", "")).strip()

        # Para no mezclar datos sin ref ni nom
        if not ref and not nom:
            continue
        if ref.upper() == "A":
            continue  # Evitar refs \"A\"

        if ref in limitantes_set:
            # Limitantes: guardamos la hoja para que aparezca en la etiqueta
            limitantes_list.append((ref, nom, hoja))
        else:
            # Compartidos: ignoramos la hoja => unificamos
            compartidos_list.append((ref, nom))

    # 3) Para quitar duplicados
    limitantes_unique = set(limitantes_list)    # (ref, nom, hoja)
    compartidos_unique = set(compartidos_list)  # (ref, nom)

    # 4) Convertimos a lista para la interfaz
    limitantes_list = list(limitantes_unique)
    compartidos_list = list(compartidos_unique)

    # 5) El usuario elige limitante o compartido
    grupo_elegido = st.radio(
        "¿Qué grupo de reactivos quiere filtrar?",
        ("limitante", "compartido"),
        key="grupo_filtrar"
    )

    if grupo_elegido == "limitante":
        op_list = limitantes_list  # (ref, nom, hoja)
    else:
        op_list = compartidos_list  # (ref, nom)

    if not op_list:
        st.warning(f"No se encontraron reactivos en la categoría '{grupo_elegido}' dentro de la base B.")
        st.stop()

    # 6) Creamos funciones para mostrar etiqueta en el selectbox
    #    NOTA: sólo limitantes muestran la hoja en la etiqueta
    def display_label_limit(tup):
        # tup => (ref, nom, hoja)
        return f"{tup[0]} - {tup[1]} ({tup[2]})"

    def display_label_comp(tup):
        # tup => (ref, nom)
        return f"{tup[0]} - {tup[1]}"

    if grupo_elegido == "limitante":
        # con hoja
        seleccion = st.selectbox(
            "Seleccione Reactivo (Limitante)",
            [display_label_limit(x) for x in op_list],
            key="select_b_filtrado_tab"
        )
    else:
        # sin hoja
        seleccion = st.selectbox(
            "Seleccione Reactivo (Compartido)",
            [display_label_comp(x) for x in op_list],
            key="select_b_filtrado_tab"
        )

    # 7) Botón para buscar
    if st.button("Buscar en Base Historial", key="buscar_filtrado"):
        if grupo_elegido == "limitante":
            # Extraemos (ref, nom, hoja)
            i_sel = [display_label_limit(x) for x in op_list].index(seleccion)
            ref_sel, nom_sel, hoja_selB = op_list[i_sel]
        else:
            # Extraemos (ref, nom)
            i_sel = [display_label_comp(x) for x in op_list].index(seleccion)
            ref_sel, nom_sel = op_list[i_sel]
            hoja_selB = None

        # Filtrar en df_b_combined por la ref
        df_filtrado = df_b_combined[df_b_combined["Ref. Fisher"].astype(str).str.strip() == ref_sel].copy()

        # (Opcional) Eliminar filas sin Caducidad
        if "Caducidad" in df_filtrado.columns:
            df_filtrado = df_filtrado.dropna(subset=["Caducidad"])

        if df_filtrado.empty:
            st.warning("No se encontraron reactivos (en B) con esos parámetros.")
        else:
            # Ordenar por Caducidad asc
            if "Caducidad" in df_filtrado.columns:
                df_filtrado.sort_values(by="Caducidad", inplace=True, ignore_index=True)

            # Mensaje especial
            if ref_sel == "A27754":
                st.info("Nota: Esta referencia se comparte entre OCA y FOCUS.")

            # Mostrar DataFrame
            st.dataframe(df_filtrado)

            # Botón de descarga
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
    nombres_unicos = sorted(df_a["Nombre producto"].dropna().unique())
    nombre_sel = st.selectbox("Nombre producto en A:", nombres_unicos, key="agotado_nombre")
    df_cand = df_a[df_a["Nombre producto"]==nombre_sel]
    if df_cand.empty:
        st.warning("No se encontró ese nombre en esta hoja A.")
    else:
        idx_c = df_cand.index[0]
        stock_c = df_a.at[idx_c,"Stock"] if "Stock" in df_a.columns else 0
        uds_consumir = st.number_input("Uds. a consumir en A:", min_value=0, step=1, key="agotado_uds")
        if st.button("Consumir en Lab (memoria)", key="agotado_consumir"):
            nuevo_stock = max(0, stock_c - uds_consumir)
            df_a.at[idx_c,"Stock"] = nuevo_stock
            if nuevo_stock==0:
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



