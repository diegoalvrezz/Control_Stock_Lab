import streamlit as st
import pandas as pd
import datetime
import shutil
import os
from io import BytesIO

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Control de Stock", layout="centered")

STOCK_FILE = "Stock_Original.xlsx"  # Archivo principal de trabajo
VERSIONS_DIR = "versions"
ORIGINAL_FILE = os.path.join(VERSIONS_DIR, "Stock_Original.xlsx")

os.makedirs(VERSIONS_DIR, exist_ok=True)

def init_original():
    """Si no existe 'versions/Stock_Original.xlsx', lo creamos tomando 'Stock_Original.xlsx'."""
    if not os.path.exists(ORIGINAL_FILE):
        if os.path.exists(STOCK_FILE):
            shutil.copy(STOCK_FILE, ORIGINAL_FILE)
            print("Creada versión original en:", ORIGINAL_FILE)
        else:
            st.error(f"No se encontró {STOCK_FILE}. Asegúrate de subirlo.")

init_original()

def load_data():
    """Carga todas las hojas de STOCK_FILE en un diccionario {nombre_hoja: DataFrame}."""
    try:
        return pd.read_excel(STOCK_FILE, sheet_name=None, engine="openpyxl")
    except FileNotFoundError:
        st.error(f"❌ No se encontró {STOCK_FILE}.")
        return None
    except Exception as e:
        st.error(f"❌ Error al cargar la base de datos: {e}")
        return None

data_dict = load_data()

# -----------------------------------------------------------------------------
# FUNCIÓN PARA CONVERSIONES DE TIPOS
# -----------------------------------------------------------------------------
def enforce_types(df: pd.DataFrame):
    # Ref. Saturno -> int
    if "Ref. Saturno" in df.columns:
        df["Ref. Saturno"] = pd.to_numeric(df["Ref. Saturno"], errors="coerce").fillna(0).astype(int)
    # Ref. Fisher -> str
    if "Ref. Fisher" in df.columns:
        df["Ref. Fisher"] = df["Ref. Fisher"].astype(str)
    # Nombre producto -> str
    if "Nombre producto" in df.columns:
        df["Nombre producto"] = df["Nombre producto"].astype(str)
    # Tª -> str
    if "Tª" in df.columns:
        df["Tª"] = df["Tª"].astype(str)
    # Uds. -> int
    if "Uds." in df.columns:
        df["Uds."] = pd.to_numeric(df["Uds."], errors="coerce").fillna(0).astype(int)
    # NºLote -> int
    if "NºLote" in df.columns:
        df["NºLote"] = pd.to_numeric(df["NºLote"], errors="coerce").fillna(0).astype(int)
    # Fechas -> datetime
    for col in ["Caducidad", "Fecha Pedida", "Fecha Llegada"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    # Restantes -> int
    if "Restantes" in df.columns:
        df["Restantes"] = pd.to_numeric(df["Restantes"], errors="coerce").fillna(0).astype(int)
    # Sitio almacenaje -> str
    if "Sitio almacenaje" in df.columns:
        df["Sitio almacenaje"] = df["Sitio almacenaje"].astype(str)
    # Stock -> int (nueva col)
    if "Stock" in df.columns:
        df["Stock"] = pd.to_numeric(df["Stock"], errors="coerce").fillna(0).astype(int)
    return df

# -----------------------------------------------------------------------------
# Funciones auxiliares para versión y para generar Excel en memoria
# -----------------------------------------------------------------------------
def crear_nueva_version_filename():
    fecha_hora = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return os.path.join(VERSIONS_DIR, f"Stock_{fecha_hora}.xlsx")

def generar_excel_en_memoria(df_act: pd.DataFrame, sheet_nm="Hoja1"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_act.to_excel(writer, index=False, sheet_name=sheet_nm)
    output.seek(0)
    return output.getvalue()

# -----------------------------------------------------------------------------
# LAYOUT PRINCIPAL
# -----------------------------------------------------------------------------

st.title("📦 Control de Stock del Hospital")

# Podemos agrupar controles en sidebars o expanders para mejor estética

with st.sidebar:
    st.markdown("## Opciones de la Base de Datos")

    # Expander para ver y manejar las versiones
    with st.expander("🔎 Ver / Gestionar versiones guardadas"):
        files = sorted(os.listdir(VERSIONS_DIR))
        # Excluimos la original, por si no queremos mostrarla
        versions_no_original = [f for f in files if f != "Stock_Original.xlsx"]
        if versions_no_original:
            version_sel = st.selectbox("Selecciona una versión:", versions_no_original)
            if version_sel:
                # Descargar
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
                        st.rerun()
                    except:
                        st.error("Error al intentar eliminar la versión.")
        else:
            st.write("No hay versiones guardadas (excepto la original).")

        # Botón para eliminar TODAS las versiones excepto la original
        if st.button("Eliminar TODAS las versiones (excepto original)"):
            for f in versions_no_original:
                try:
                    os.remove(os.path.join(VERSIONS_DIR, f))
                except:
                    pass
            st.info("Todas las versiones (excepto la original) han sido eliminadas.")
            st.rerun()

    # Botón para limpiar la base de datos
    if st.button("Limpiar Base de Datos"):
        if os.path.exists(ORIGINAL_FILE):
            shutil.copy(ORIGINAL_FILE, STOCK_FILE)
            st.success("✅ Base de datos restaurada al estado original.")
            st.rerun()
        else:
            st.error("❌ No se encontró la copia original en 'versions/Stock_Original.xlsx'.")

# -----------------------------------------------------------------------------
# MOSTRAR / EDITAR DATOS
# -----------------------------------------------------------------------------
if data_dict:
    sheet_name = st.selectbox("Selecciona la categoría de stock:", list(data_dict.keys()))
    df = data_dict[sheet_name].copy()
    df = enforce_types(df)

    st.markdown(f"### Hoja seleccionada: **{sheet_name}**")
    st.dataframe(df)  # Ahora usamos st.dataframe para tener scroll

    if "Nombre producto" in df.columns and "Ref. Fisher" in df.columns:
        display_series = df.apply(lambda row: f"{row['Nombre producto']} ({row['Ref. Fisher']})", axis=1)
    else:
        display_series = df.iloc[:, 0].astype(str)

    reactivo = st.selectbox("Selecciona Reactivo a Modificar:", display_series.unique())
    row_index = display_series[display_series == reactivo].index[0]

    # -------------------------------------------------------------------------
    # Cargar valores
    def get_val(col, default=None):
        return df.at[row_index, col] if col in df.columns else default

    lote_actual = get_val("NºLote", 0)
    caducidad_actual = get_val("Caducidad", None)
    fecha_pedida_actual = get_val("Fecha Pedida", None)
    fecha_llegada_actual = get_val("Fecha Llegada", None)
    sitio_almacenaje_actual = get_val("Sitio almacenaje", "")
    uds_actual = get_val("Uds.", 0)
    stock_actual = get_val("Stock", 0)

    st.markdown("#### Modificar Atributos de Reactivo")
    col1, col2 = st.columns(2)
    with col1:
        lote_nuevo = st.number_input("Nº de Lote", value=int(lote_actual), step=1)
        caducidad_nueva = st.date_input("Caducidad", value=caducidad_actual if pd.notna(caducidad_actual) else None)
        fecha_pedida_nueva = st.date_input("Fecha Pedida", value=fecha_pedida_actual if pd.notna(fecha_pedida_actual) else None)
    with col2:
        fecha_llegada_nueva = st.date_input("Fecha Llegada", value=fecha_llegada_actual if pd.notna(fecha_llegada_actual) else None)

        # Sitio Almacenaje
        opciones_sitio = ["Congelador 1", "Congelador 2", "Frigorífico", "Tª Ambiente"]
        sitio_principal = sitio_almacenaje_actual.split(" - ")[0] if " - " in sitio_almacenaje_actual else sitio_almacenaje_actual
        if sitio_principal not in opciones_sitio:
            sitio_principal = opciones_sitio[0]
        sitio_top = st.selectbox("Sitio de Almacenaje", opciones_sitio, index=opciones_sitio.index(sitio_principal))

        subopcion = ""
        if sitio_top == "Congelador 1":
            cajones = [f"Cajón {i}" for i in range(1, 9)]
            subopcion = st.selectbox("Cajón", cajones)
        elif sitio_top == "Congelador 2":
            cajones = [f"Cajón {i}" for i in range(1, 7)]
            subopcion = st.selectbox("Cajón", cajones)
        elif sitio_top == "Frigorífico":
            baldas = [f"Balda {i}" for i in range(1, 7)] + ["Puerta"]
            subopcion = st.selectbox("Baldas", baldas)

        if subopcion:
            sitio_almacenaje_nuevo = f"{sitio_top} - {subopcion}"
        else:
            sitio_almacenaje_nuevo = sitio_top

    # -------------------------------------------------------------------------
    # Botón GUARDAR CAMBIOS
    # -------------------------------------------------------------------------
    if st.button("Guardar Cambios"):
        # EJEMPLO: Sumar Stock si la fecha de llegada cambió
        if "Stock" in df.columns:
            if (fecha_llegada_nueva != fecha_llegada_actual) and pd.notna(fecha_llegada_nueva):
                df.at[row_index, "Stock"] = stock_actual + uds_actual
                st.info(f"Sumadas {uds_actual} uds al stock. Nuevo stock => {stock_actual + uds_actual}")

        # Crear versión
        new_file = crear_nueva_version_filename()

        # Actualizar df
        if "NºLote" in df.columns:
            df.at[row_index, "NºLote"] = int(lote_nuevo)
        if "Caducidad" in df.columns:
            df.at[row_index, "Caducidad"] = pd.to_datetime(caducidad_nueva)
        if "Fecha Pedida" in df.columns:
            df.at[row_index, "Fecha Pedida"] = pd.to_datetime(fecha_pedida_nueva)
        if "Fecha Llegada" in df.columns:
            df.at[row_index, "Fecha Llegada"] = pd.to_datetime(fecha_llegada_nueva)
        if "Sitio almacenaje" in df.columns:
            df.at[row_index, "Sitio almacenaje"] = sitio_almacenaje_nuevo

        # Guardar versión en disco
        with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
            for sht, df_sheet in data_dict.items():
                if sht == sheet_name:
                    df.to_excel(writer, sheet_name=sht, index=False)
                else:
                    df_sheet.to_excel(writer, sheet_name=sht, index=False)

        # Guardar en STOCK_FILE
        with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
            for sht, df_sheet in data_dict.items():
                if sht == sheet_name:
                    df.to_excel(writer, sheet_name=sht, index=False)
                else:
                    df_sheet.to_excel(writer, sheet_name=sht, index=False)

        st.success(f"✅ Cambios guardados en '{new_file}' y '{STOCK_FILE}'.")

        # Descarga del Excel actualizado en memoria
        excel_bytes = generar_excel_en_memoria(df, sheet_nm=sheet_name)
        st.download_button(
            label="Descargar Excel modificado",
            data=excel_bytes,
            file_name="Reporte_Stock.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # st.experimental_rerun()

    # -------------------------------------------------------------------------
    # Sección "Reactivo Agotado"
    # -------------------------------------------------------------------------
    st.markdown("---")
    st.markdown("### Reactivo Agotado")
    st.write("Si un reactivo se acaba, indica cuál y cuántas unidades salieron del stock.")

    if "Stock" in df.columns:
        # Elegir el reactivo
        reactivo_agotado = st.selectbox("Selecciona Reactivo a Consumir:", display_series.unique())
        row_idx_agotado = display_series[display_series == reactivo_agotado].index[0]
        stock_actual_agotado = df.at[row_idx_agotado, "Stock"] if not pd.isna(df.at[row_idx_agotado, "Stock"]) else 0
        uds_consumidas = st.number_input("Unidades consumidas", min_value=0, step=1)

        if st.button("Registrar Consumo"):
            # Resta al stock sin caer por debajo de 0
            nuevo_stock = max(0, stock_actual_agotado - uds_consumidas)
            df.at[row_idx_agotado, "Stock"] = nuevo_stock
            st.warning(f"Se han consumido {uds_consumidas} uds. Stock final => {nuevo_stock}")

            # Guardamos en disco (y en versión) rápido
            new_file2 = crear_nueva_version_filename()
            # Actualizamos DF en el DataDict
            data_dict[sheet_name] = df

            with pd.ExcelWriter(new_file2, engine="openpyxl") as writer:
                for sht, df_sheet in data_dict.items():
                    df_sheet.to_excel(writer, sheet_name=sht, index=False)

            with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
                for sht, df_sheet in data_dict.items():
                    df_sheet.to_excel(writer, sheet_name=sht, index=False)

            st.success(f"✅ Stock actualizado. Cambios guardados en '{new_file2}' y '{STOCK_FILE}'.")

            # Si quieres permitir descarga inmediata:
            excel_bytes_2 = generar_excel_en_memoria(df, sheet_nm=sheet_name)
            st.download_button(
                label="Descargar Excel tras el consumo",
                data=excel_bytes_2,
                file_name="Reporte_Stock_Agotado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # st.experimental_rerun()
    else:
        st.info("No se encontró la columna 'Stock' en esta hoja. Agrega la columna primero.")
