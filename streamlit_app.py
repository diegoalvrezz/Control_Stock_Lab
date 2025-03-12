import streamlit as st
import streamlit_authenticator as stauth
import pandas as pd
import numpy as np
import datetime
import shutil
import os
from io import BytesIO
import itertools

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
# EXCEL A (Stock_Original) - Rutas y funciones
# -------------------------------------------------------------------------
STOCK_FILE = "Stock_Original.xlsx"   # Excel A
VERSIONS_DIR = "versions"
ORIGINAL_FILE = os.path.join(VERSIONS_DIR, "Stock_Original.xlsx")

os.makedirs(VERSIONS_DIR, exist_ok=True)

def init_original():
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

def crear_nueva_version_filename():
    fecha_hora = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return os.path.join(VERSIONS_DIR, f"Stock_{fecha_hora}.xlsx")

# -------------------------------------------------------------------------
# EXCEL B (Stock_Historico) - Rutas y funciones
# -------------------------------------------------------------------------
STOCK_FILE_B = "Stock_Historico.xlsx"  # Excel B (histórico)
VERSIONS_DIR_B = "versions_b"
ORIGINAL_FILE_B = os.path.join(VERSIONS_DIR_B, "Stock_Historico_Original.xlsx")

os.makedirs(VERSIONS_DIR_B, exist_ok=True)

def init_original_b():
    """Inicia un archivo B vacío si no existe la copia original."""
    if not os.path.exists(ORIGINAL_FILE_B):
        if os.path.exists(STOCK_FILE_B):
            shutil.copy(STOCK_FILE_B, ORIGINAL_FILE_B)
        else:
            # Creamos un excel vacío con 1 hoja "Hoja1"
            df_empty = pd.DataFrame(columns=[
                "Ref. Saturno","Ref. Fisher","Nombre producto","NºLote","Caducidad",
                "Fecha Pedida","Fecha Llegada","Sitio almacenaje","Uds.","Stock"
            ])
            with pd.ExcelWriter(STOCK_FILE_B, engine="openpyxl") as writer:
                df_empty.to_excel(writer, sheet_name="Hoja1", index=False)
            shutil.copy(STOCK_FILE_B, ORIGINAL_FILE_B)

init_original_b()

def crear_nueva_version_filename_b():
    fecha_hora = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return os.path.join(VERSIONS_DIR_B, f"StockB_{fecha_hora}.xlsx")

# -------------------------------------------------------------------------
# FUNCIONES COMUNES
# -------------------------------------------------------------------------
import openpyxl

def generar_excel_en_memoria(df_act: pd.DataFrame, sheet_nm="Hoja1"):
    """Genera un Excel en memoria para descargar."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_act.to_excel(writer, index=False, sheet_name=sheet_nm)
    output.seek(0)
    return output.getvalue()

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

# -------------------------------------------------------------------------
# CARGA LOS DICCIONARIOS DE A y B
# -------------------------------------------------------------------------
def load_data_b():
    """Lee todas las hojas de STOCK_FILE_B (histórico)."""
    if not os.path.exists(STOCK_FILE_B):
        return {}
    try:
        data_b = pd.read_excel(STOCK_FILE_B, sheet_name=None, engine="openpyxl")
        return data_b
    except:
        return {}

data_dict_b = load_data_b()
if data_dict_b is None:
    data_dict_b = {}

# -------------------------------------------------------------------------
# LÓGICA DE LOTES Y ESTILOS (del original)
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

panel_order = ["FOCUS", "OCA", "OCA PLUS"]

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

st.markdown("""
    <style>
    .big-select select {
        font-size: 18px;
        height: auto;
    }
    </style>
    """, unsafe_allow_html=True)

# -------------------------------------------------------------------------
# BARRA LATERAL => GESTIONAR VERSIONES DE A (tal cual tu código)
# -------------------------------------------------------------------------
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
                            st.error("Marca la casilla para confirmar la eliminación.")
            else:
                st.write("No hay versiones guardadas (excepto la original).")

            if st.button("Eliminar TODAS las versiones (excepto original)"):
                for f in versions_no_original:
                    try:
                        os.remove(os.path.join(VERSIONS_DIR, f))
                    except:
                        pass
                st.info("Todas las versiones (excepto la original) eliminadas.")
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

            # --- Botón de Limpiar Base de Datos ---
            limpiar_confirmado = st.checkbox("Confirmar limpieza de la base de datos", key="confirmar_limpiar")
            if st.button("Limpiar Base de Datos") and limpiar_confirmado:
                original_path = os.path.join(VERSIONS_DIR, "Stock_Original.xlsx")
                if os.path.exists(original_path):
                    shutil.copy(original_path, STOCK_FILE)
                    st.success("✅ Base de datos restaurada al estado original.")
                    st.rerun()
                else:
                    st.error("❌ No se encontró la copia original en 'versions/Stock_Original.xlsx'.")
        else:
            st.error("No hay data_dict. Verifica Stock_Original.xlsx.")
            st.stop()

# -------------------------------------------------------------------------
# BARRA LATERAL => GESTIONAR VERSIONES DE B (Histórico)
# -------------------------------------------------------------------------
    with st.expander("🔎 Ver / Gestionar versiones B (Histórico)", expanded=False):
        files_b = sorted(os.listdir(VERSIONS_DIR_B))
        versions_no_original_b = [f for f in files_b if f != "Stock_Historico_Original.xlsx"]
        if versions_no_original_b:
            version_sel_b = st.selectbox("Selecciona versión B:", versions_no_original_b)
            confirm_delete_b = False

            if version_sel_b:
                file_path_b = os.path.join(VERSIONS_DIR_B, version_sel_b)
                if os.path.isfile(file_path_b):
                    with open(file_path_b, "rb") as excel_file_b:
                        excel_bytes_b = excel_file_b.read()
                    st.download_button(
                        label=f"Descargar {version_sel_b}",
                        data=excel_bytes_b,
                        file_name=version_sel_b,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                if st.checkbox(f"Confirmar eliminación de '{version_sel_b}' (B)"):
                    confirm_delete_b = True

                if st.button("Eliminar esta versión B"):
                    if confirm_delete_b:
                        try:
                            os.remove(file_path_b)
                            st.warning(f"Versión '{version_sel_b}' eliminada de B.")
                            st.rerun()
                        except:
                            st.error("Error al intentar eliminar la versión.")
                    else:
                        st.error("Marca la casilla para confirmar la eliminación.")
        else:
            st.write("No hay versiones guardadas de B (excepto la original).")

        if st.button("Eliminar TODAS las versiones B (excepto original)"):
            for f in versions_no_original_b:
                try:
                    os.remove(os.path.join(VERSIONS_DIR_B, f))
                except:
                    pass
            st.info("Todas las versiones de B (excepto la original) eliminadas.")
            st.rerun()

        if st.button("Eliminar TODAS las versiones B excepto la última y la original"):
            if len(versions_no_original_b) > 1:
                sorted_vers_b = sorted(versions_no_original_b)
                last_version_b = sorted_vers_b[-1]
                for f in versions_no_original_b:
                    if f != last_version_b:
                        try:
                            os.remove(os.path.join(VERSIONS_DIR_B, f))
                        except:
                            pass
                st.info(f"Se han eliminado todas las versiones excepto: {last_version_b} y Stock_Historico_Original.xlsx")
                st.rerun()
            else:
                st.write("Solo hay una versión o ninguna versión, no se elimina nada más.")

        # Limpiar Base de Datos B
        if st.button("Limpiar Base de Datos B"):
            if os.path.exists(ORIGINAL_FILE_B):
                shutil.copy(ORIGINAL_FILE_B, STOCK_FILE_B)
                st.success("Base de datos B restaurada al estado original.")
                st.rerun()
            else:
                st.error("No se encontró la copia original de B.")


    # ---------------------------------------------------------------------
    # EXPANDER PARA VER BASE DE DATOS B (Histórico)
    # ---------------------------------------------------------------------
    with st.expander("Ver Base de Datos Histórica (Excel B)", expanded=False):
        if data_dict_b:
            hojas_b = list(data_dict_b.keys())
            hoja_b_sel = st.selectbox("Selecciona hoja en B:", hojas_b)
            df_b_vista = data_dict_b[hoja_b_sel].copy()

            # Orden por Nombre producto + NºLote (si existen)
            if "Nombre producto" in df_b_vista.columns and "NºLote" in df_b_vista.columns:
                df_b_vista.sort_values(by=["Nombre producto","NºLote"], inplace=True, ignore_index=True)

            st.write("Vista de B (Histórico):")
            st.dataframe(df_b_vista)

            excel_b_mem = generar_excel_en_memoria(df_b_vista, sheet_nm=hoja_b_sel)
            st.download_button(
                label="Descargar hoja de Excel B",
                data=excel_b_mem,
                file_name="Hoja_Historico_B.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.write("No se encontró data_dict_b o está vacío.")


    # ---------------------------------------------------------------------
    # REACTIVO AGOTADO (Consumido en Lab) => Actualizar Stock en A y Eliminar en B
    # ---------------------------------------------------------------------
    with st.expander("Reactivo Agotado (Consumido en Lab)", expanded=False):
        if data_dict:
            st.write("Selecciona hoja y reactivo (con lote) para consumir stock y guardar versión en A, y eliminar fila en B.")
            hojas_agotado = list(data_dict.keys())
            hoja_sel_consumo = st.selectbox("Hoja a consumir (A):", hojas_agotado, key="cons_hoja_sel")

            df_agotado = data_dict[hoja_sel_consumo].copy()
            df_agotado = enforce_types(df_agotado)

            # Construimos un display que muestre "Nombre producto (NºLote)" para distinguir
            if "Nombre producto" in df_agotado.columns and "NºLote" in df_agotado.columns:
                disp_consumo = df_agotado.apply(
                    lambda r: f"{r['Nombre producto']} (Lote:{r['NºLote']})", axis=1
                )
            else:
                # fallback
                disp_consumo = df_agotado.iloc[:, 0].astype(str)

            reactivo_consumir = st.selectbox("Reactivo+Lote:", disp_consumo.unique(), key="cons_react_sel")
            idx_c = disp_consumo[disp_consumo == reactivo_consumir].index[0]

            stock_c = df_agotado.at[idx_c, "Stock"] if "Stock" in df_agotado.columns else 0
            nombre_c = df_agotado.at[idx_c, "Nombre producto"] if "Nombre producto" in df_agotado.columns else ""
            lote_c = df_agotado.at[idx_c, "NºLote"] if "NºLote" in df_agotado.columns else 0

            uds_consumidas = st.number_input("Uds. consumidas", min_value=0, step=1, key="uds_consumidas")
            if st.button("Registrar Consumo en Lab"):
                nuevo_stock = max(0, stock_c - uds_consumidas)
                df_agotado.at[idx_c, "Stock"] = nuevo_stock
                st.warning(f"Consumidas {uds_consumidas} uds. Stock final => {nuevo_stock}")
                data_dict[hoja_sel_consumo] = df_agotado

            if st.button("Guardar Cambios en Consumo Lab"):
                # 1) Guardamos en A, creando nueva versión
                new_file = crear_nueva_version_filename()
                with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
                    for sht, df_sht in data_dict.items():
                        cols_internos = ["ColorGroup", "EsTitulo", "GroupCount", "MultiSort", "NotTitulo", "GroupID"]
                        temp = df_sht.drop(columns=cols_internos, errors="ignore")
                        temp.to_excel(writer, sheet_name=sht, index=False)
                with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
                    for sht, df_sht in data_dict.items():
                        cols_internos = ["ColorGroup", "EsTitulo", "GroupCount", "MultiSort", "NotTitulo", "GroupID"]
                        temp = df_sht.drop(columns=cols_internos, errors="ignore")
                        temp.to_excel(writer, sheet_name=sht, index=False)

                # 2) Eliminamos la fila en B (mismo hoja_sel_consumo si existe en B) que coincida en (Nombre producto, NºLote)
                if hoja_sel_consumo in data_dict_b:
                    df_b_hoja = data_dict_b[hoja_sel_consumo].copy()
                    # Eliminamos filas donde coincidan Nombre producto y NºLote
                    if "Nombre producto" in df_b_hoja.columns and "NºLote" in df_b_hoja.columns:
                        df_b_hoja = df_b_hoja[~((df_b_hoja["Nombre producto"] == nombre_c) &
                                               (df_b_hoja["NºLote"] == lote_c))]
                        data_dict_b[hoja_sel_consumo] = df_b_hoja
                        # Guardamos version B
                        new_file_b = crear_nueva_version_filename_b()
                        with pd.ExcelWriter(new_file_b, engine="openpyxl") as writer_b:
                            for shtb, df_shtb in data_dict_b.items():
                                df_shtb.to_excel(writer_b, sheet_name=shtb, index=False)
                        with pd.ExcelWriter(STOCK_FILE_B, engine="openpyxl") as writer_b:
                            for shtb, df_shtb in data_dict_b.items():
                                df_shtb.to_excel(writer_b, sheet_name=shtb, index=False)

                st.success(f"✅ Cambios guardados en '{new_file}' y '{STOCK_FILE}'. Fila eliminada en B si coincidía.")
                excel_bytes = generar_excel_en_memoria(df_agotado, sheet_nm=hoja_sel_consumo)
                st.download_button(
                    label="Descargar Excel modificado (A)",
                    data=excel_bytes,
                    file_name="Reporte_Stock.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.rerun()
        else:
            st.error("No hay data_dict. Revisa Stock_Original.xlsx.")
            st.stop()

# -------------------------------------------------------------------------
# CUERPO PRINCIPAL => Edición en Hoja Principal (A)
# -------------------------------------------------------------------------
st.title("📦 Control de Stock Secuenciación - (Excel A + Histórico B)")

if not data_dict:
    st.error("No se pudo cargar la base de datos.")
    st.stop()

st.markdown("---")
st.header("Edición en Hoja Principal y Guardado (Excel A)")

hojas_principales = list(data_dict.keys())
sheet_name = st.selectbox("Selecciona la hoja a editar:", hojas_principales, key="main_sheet_sel")
df_main_original = data_dict[sheet_name].copy()
df_main_original = enforce_types(df_main_original)

# 1) Crear df para estilo: calcular alarma y agrupar por Ref. Saturno
df_for_style = df_main_original.copy()
df_for_style["Alarma"] = df_for_style.apply(calc_alarma, axis=1)
df_for_style = build_group_info_by_ref(df_for_style, panel_default=sheet_name)

# 2) Ordenar
df_for_style.sort_values(by=["MultiSort", "GroupID", "NotTitulo"], inplace=True)
df_for_style.reset_index(drop=True, inplace=True)
styled_df = df_for_style.style.apply(style_lote, axis=1)

all_cols = df_for_style.columns.tolist()
cols_to_hide = ["ColorGroup", "EsTitulo", "GroupCount", "MultiSort", "NotTitulo", "GroupID"]
final_cols = [c for c in all_cols if c not in cols_to_hide]

table_html = styled_df.to_html(columns=final_cols)

# 3) Crear df_main final sin columnas internas
df_main = df_for_style.copy()
df_main.drop(columns=cols_to_hide, inplace=True, errors="ignore")

st.write("#### Vista de la Hoja (con columna 'Alarma' y sin columnas internas)")
st.write(table_html, unsafe_allow_html=True)

# 4) Seleccionar Reactivo a Modificar
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

colA, colB, colC, colD = st.columns([1, 1, 1, 1])
with colA:
    lote_new = st.number_input("Nº de Lote", value=int(lote_actual), step=1)
    cad_new = st.date_input("Caducidad", value=caducidad_actual if pd.notna(caducidad_actual) else None)
with colB:
    fped_date = st.date_input("Fecha Pedida (fecha)",
                              value=fecha_pedida_actual.date() if pd.notna(fecha_pedida_actual) else None,
                              key="fped_date_main")
    fped_time = st.time_input("Hora Pedida",
                              value=fecha_pedida_actual.time() if pd.notna(fecha_pedida_actual) else datetime.time(0, 0),
                              key="fped_time_main")
with colC:
    flleg_date = st.date_input("Fecha Llegada (fecha)",
                               value=fecha_llegada_actual.date() if pd.notna(fecha_llegada_actual) else None,
                               key="flleg_date_main")
    flleg_time = st.time_input("Hora Llegada",
                               value=fecha_llegada_actual.time() if pd.notna(fecha_llegada_actual) else datetime.time(0, 0),
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
opciones_sitio = ["Congelador 1", "Congelador 2", "Frigorífico", "Tª Ambiente"]
sitio_p = sitio_almacenaje_actual.split(" - ")[0] if " - " in sitio_almacenaje_actual else sitio_almacenaje_actual
if sitio_p not in opciones_sitio:
    sitio_p = opciones_sitio[0]
sel_top = st.selectbox("Almacén Principal", opciones_sitio, index=opciones_sitio.index(sitio_p))
subopc = ""
if sel_top == "Congelador 1":
    cajs = [f"Cajón {i}" for i in range(1, 9)]
    subopc = st.selectbox("Cajón (1 Arriba,8 Abajo)", cajs)
elif sel_top == "Congelador 2":
    cajs = [f"Cajón {i}" for i in range(1, 7)]
    subopc = st.selectbox("Cajón (1 Arriba,6 Abajo)", cajs)
elif sel_top == "Frigorífico":
    blds = [f"Balda {i}" for i in range(1, 8)] + ["Puerta"]
    subopc = st.selectbox("Baldas (1 Arriba, 7 Abajo)", blds)
elif sel_top == "Tª Ambiente":
    com2 = st.text_input("Comentario (opt)")
    subopc = com2.strip()
if subopc:
    sitio_new = f"{sel_top} - {subopc}"
else:
    sitio_new = sel_top

# NUEVA SECCIÓN: Si se ingresó Fecha Pedida, preguntar por el pedido del grupo completo
group_order_selected = None
if pd.notna(fped_new):
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
        group_order_selected = st.multiselect(
            f"¿Quieres pedir también los siguientes reactivos del lote **{lot_name}**?",
            options,
            default=options
        )
        st.markdown('</div>', unsafe_allow_html=True)

# -------------------------------------------------------------------------
# Botón para Guardar Cambios (actualiza fila en A) + EJEMPLO Insertar en B
# -------------------------------------------------------------------------
if st.button("Guardar Cambios"):
    # Si se ingresó Fecha Llegada, forzamos Fecha Pedida a NaT
    if pd.notna(flleg_new):
        fped_new = pd.NaT

    if "Stock" in df_main.columns:
        if flleg_new != fecha_llegada_actual and pd.notna(flleg_new):
            df_main.at[row_index, "Stock"] = stock_actual + uds_actual
            st.info(f"Añadidas {uds_actual} uds al stock => {stock_actual + uds_actual}")

    if "NºLote" in df_main.columns:
        df_main.at[row_index, "NºLote"] = int(lote_new)
    if "Caducidad" in df_main.columns:
        df_main.at[row_index, "Caducidad"] = cad_new if pd.notna(cad_new) else pd.NaT
    if "Fecha Pedida" in df_main.columns:
        df_main.at[row_index, "Fecha Pedida"] = fped_new
    if "Fecha Llegada" in df_main.columns:
        df_main.at[row_index, "Fecha Llegada"] = flleg_new
    if "Sitio almacenaje" in df_main.columns:
        df_main.at[row_index, "Sitio almacenaje"] = sitio_new

    # Actualización en grupo: actualizar "Fecha Pedida" para cada fila seleccionada en el multiselect.
    if pd.notna(fped_new) and group_order_selected:
        for label in group_order_selected:
            try:
                i_val = int(label.split(" - ")[0])
                df_main.at[i_val, "Fecha Pedida"] = fped_new
            except Exception as e:
                st.error(f"Error actualizando índice {label}: {e}")

    data_dict[sheet_name] = df_main
    new_file = crear_nueva_version_filename()
    # Guardamos en nueva versión A
    with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
        for sht, df_sht in data_dict.items():
            ocultar = ["ColorGroup", "EsTitulo", "GroupCount", "MultiSort", "NotTitulo", "GroupID"]
            temp = df_sht.drop(columns=ocultar, errors="ignore")
            temp.to_excel(writer, sheet_name=sht, index=False)
    # Guardamos en el principal A
    with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
        for sht, df_sht in data_dict.items():
            ocultar = ["ColorGroup", "EsTitulo", "GroupCount", "MultiSort", "NotTitulo", "GroupID"]
            temp = df_sht.drop(columns=ocultar, errors="ignore")
            temp.to_excel(writer, sheet_name=sht, index=False)

    # ===== EJEMPLO: Insertar la misma entrada en B (histórico)
    #       Ajusta si quieres que siempre se cree una nueva fila en B, o condicional
    if sheet_name not in data_dict_b:
        data_dict_b[sheet_name] = pd.DataFrame()

    df_b_sheet = data_dict_b[sheet_name].copy()
    nueva_fila_b = {
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
        "Fecha Registro B": datetime.datetime.now()  # extra
    }
    df_b_sheet = pd.concat([df_b_sheet, pd.DataFrame([nueva_fila_b])], ignore_index=True)
    data_dict_b[sheet_name] = df_b_sheet

    new_file_b = crear_nueva_version_filename_b()
    # Guardamos la nueva versión B
    with pd.ExcelWriter(new_file_b, engine="openpyxl") as writerB:
        for sht_b, df_sht_b in data_dict_b.items():
            df_sht_b.to_excel(writerB, sheet_name=sht_b, index=False)
    # Guardamos en el principal B
    with pd.ExcelWriter(STOCK_FILE_B, engine="openpyxl") as writerB:
        for sht_b, df_sht_b in data_dict_b.items():
            df_sht_b.to_excel(writerB, sheet_name=sht_b, index=False)
    # ===== FIN EJEMPLO

    st.success(f"✅ Cambios guardados en '{new_file}' y '{STOCK_FILE}' (Excel A). También en '{new_file_b}' y '{STOCK_FILE_B}' (Excel B).")
    excel_bytes = generar_excel_en_memoria(df_main, sheet_nm=sheet_name)
    st.download_button(
        label="Descargar Excel A modificado",
        data=excel_bytes,
        file_name="Reporte_Stock.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.rerun()
