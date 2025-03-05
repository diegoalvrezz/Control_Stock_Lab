import streamlit as st
import pandas as pd
import datetime
import shutil
import os
from io import BytesIO

STOCK_FILE = "Stock_Original.xlsx"  # Archivo principal de trabajo
VERSIONS_DIR = "versions"
ORIGINAL_FILE = os.path.join(VERSIONS_DIR, "Stock_Original.xlsx")

os.makedirs(VERSIONS_DIR, exist_ok=True)

def init_original():
    """Si no existe 'versions/Stock_Original.xlsx', lo creamos tomando el 'Stock_Original.xlsx'."""
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

# -------------------------------------------------------------------------------------
# BOTONES AUXILIARES
# -------------------------------------------------------------------------------------

# 1) Botón para ver versiones guardadas
if st.button("Ver versiones guardadas"):
    files = os.listdir(VERSIONS_DIR)
    if files:
        st.write("### Archivos en la carpeta 'versions':")
        for f in files:
            file_path = os.path.join(VERSIONS_DIR, f)
            if os.path.isfile(file_path):
                # Opción para descargar cada archivo
                with open(file_path, "rb") as excel_file:
                    excel_bytes = excel_file.read()
                st.download_button(
                    label=f"Descargar {f}",
                    data=excel_bytes,
                    file_name=f,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.write("No hay versiones guardadas aún.")

# 2) Botón para limpiar base de datos
if st.button("Limpiar Base de Datos"):
    if os.path.exists(ORIGINAL_FILE):
        shutil.copy(ORIGINAL_FILE, STOCK_FILE)
        st.success("✅ Base de datos restaurada al estado original.")
        st.rerun()
    else:
        st.error("❌ No se encontró la copia original en 'versions/Stock_Original.xlsx'.")

# -------------------------------------------------------------------------------------
# SI HAY DATOS, PROCEDEMOS
# -------------------------------------------------------------------------------------
if data_dict:
    st.title("📦 Control de Stock del Hospital")

    # Seleccionar la hoja
    sheet_name = st.selectbox("Selecciona la categoría de stock:", list(data_dict.keys()))
    df = data_dict[sheet_name].copy()

    # Conversiones de tipos, según tus requisitos
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
        return df

    df = enforce_types(df)

    # Muestra la tabla (sin PyArrow) con st.write
    st.write(f"📋 Mostrando datos de: **{sheet_name}**")
    st.write(df)

    # Crear columna de exhibición => "Nombre producto (Ref. Fisher)"
    if "Nombre producto" in df.columns and "Ref. Fisher" in df.columns:
        display_series = df.apply(
            lambda row: f"{row['Nombre producto']} ({row['Ref. Fisher']})",
            axis=1
        )
    else:
        display_series = df.iloc[:, 0].astype(str)

    reactivo = st.selectbox("Selecciona el reactivo a modificar:", display_series.unique())
    row_index = display_series[display_series == reactivo].index[0]

    # Cargar valores actuales
    def get_val(col, default=None):
        return df.at[row_index, col] if col in df.columns else default

    lote_actual = get_val("NºLote", 0)
    caducidad_actual = get_val("Caducidad", None)
    fecha_pedida_actual = get_val("Fecha Pedida", None)
    fecha_llegada_actual = get_val("Fecha Llegada", None)
    sitio_almacenaje_actual = get_val("Sitio almacenaje", "")

    st.subheader("✏️ Modificar Reactivo")

    # Nº Lote
    lote_nuevo = st.number_input("Nº de Lote",
        value=int(lote_actual) if pd.notna(lote_actual) else 0,
        step=1
    )
    # Caducidad
    caducidad_nueva = st.date_input("Caducidad",
        value=caducidad_actual if pd.notna(caducidad_actual) else None
    )
    # Fecha Pedida
    fecha_pedida_nueva = st.date_input("Fecha Pedida",
        value=fecha_pedida_actual if pd.notna(fecha_pedida_actual) else None
    )
    # Fecha Llegada
    fecha_llegada_nueva = st.date_input("Fecha Llegada",
        value=fecha_llegada_actual if pd.notna(fecha_llegada_actual) else None
    )

    # Manejo Sitio Almacenaje
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
    else:
        subopcion = ""

    if subopcion:
        sitio_almacenaje_nuevo = f"{sitio_top} - {subopcion}"
    else:
        sitio_almacenaje_nuevo = sitio_top

    # -------------------------------------------------------------------------
    # Función para generar un archivo con fecha/hora en "versions"
    # -------------------------------------------------------------------------
    def crear_nueva_version_filename():
        fecha_hora = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        return os.path.join(VERSIONS_DIR, f"Stock_{fecha_hora}.xlsx")

    # -------------------------------------------------------------------------
    # Función para generar Excel en memoria y retornarlo como bytes
    # -------------------------------------------------------------------------
    from io import BytesIO

    def generar_excel_en_memoria(df_act: pd.DataFrame, sheet_nm="Hoja1"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_act.to_excel(writer, index=False, sheet_name=sheet_nm)
        output.seek(0)
        return output.getvalue()

    # -------------------------------------------------------------------------
    # BOTÓN: GUARDAR CAMBIOS
    # -------------------------------------------------------------------------
    if st.button("Guardar Cambios"):
        # 1) Creamos un nuevo archivo en versions
        new_file = crear_nueva_version_filename()

        # 2) Actualizamos df en memoria
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

        # 3) Guardar la versión con fecha/hora en disco
        with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
            for sht, df_sheet in data_dict.items():
                if sht == sheet_name:
                    df.to_excel(writer, sheet_name=sht, index=False)
                else:
                    df_sheet.to_excel(writer, sheet_name=sht, index=False)

        # 4) Guardar TAMBIÉN en nuestro archivo de trabajo (STOCK_FILE)
        with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
            for sht, df_sheet in data_dict.items():
                if sht == sheet_name:
                    df.to_excel(writer, sheet_name=sht, index=False)
                else:
                    df_sheet.to_excel(writer, sheet_name=sht, index=False)

        st.success(f"✅ Cambios guardados en '{new_file}' y '{STOCK_FILE}'.")

        # 5) Generar Excel actualizado en memoria para descargar
        excel_bytes = generar_excel_en_memoria(df, sheet_nm=sheet_name)
        st.download_button(
            label="Descargar Excel con la tabla modificada",
            data=excel_bytes,
            file_name="Reporte_Stock.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Si deseas recargar la app en este punto, quita el comentario:
        # st.rerun()
