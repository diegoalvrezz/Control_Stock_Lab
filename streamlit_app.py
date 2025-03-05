import streamlit as st
import pandas as pd
import datetime
import shutil
import os

# -------------------------------------------------------------------------------------
# NOMBRE ÚNICO DEL ARCHIVO SUBIDO POR EL USUARIO
# -------------------------------------------------------------------------------------
STOCK_FILE = "Stock_Original.xlsx"  # El único archivo que el usuario sube
VERSIONS_DIR = "versions"               # Carpeta donde almacenamos versiones
ORIGINAL_FILE = os.path.join(VERSIONS_DIR, "Stock_Original.xlsx")

os.makedirs(VERSIONS_DIR, exist_ok=True)

# -------------------------------------------------------------------------------------
# FUNCIÓN PARA CREAR LA VERSIÓN ORIGINAL AL INICIO
# -------------------------------------------------------------------------------------
def init_original():
    """
    Si no existe versions/Stock_Original.xlsx, lo creamos tomando Stock_Modificadov1.xlsx
    Así solo se necesita subir un archivo, y en 'versions' guardamos la copia original.
    """
    if not os.path.exists(ORIGINAL_FILE):
        if os.path.exists(STOCK_FILE):
            shutil.copy(STOCK_FILE, ORIGINAL_FILE)
            print("Se creó el archivo original en 'versions/Stock_Original.xlsx'")
        else:
            st.error(f"No se encontró {STOCK_FILE}. Asegúrate de subirlo.")

init_original()

# -------------------------------------------------------------------------------------
# FUNCIÓN PARA CARGAR TODAS LAS HOJAS DEL ARCHIVO PRINCIPAL
# -------------------------------------------------------------------------------------
def load_data():
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
# BOTÓN PARA VER VERSIONES GUARDADAS
# -------------------------------------------------------------------------------------
if st.button("Ver versiones guardadas"):
    files = os.listdir(VERSIONS_DIR)
    if files:
        st.write("### Archivos en la carpeta 'versions':")
        for f in files:
            st.write(f"- {f}")
    else:
        st.write("No hay versiones guardadas aún.")

# -------------------------------------------------------------------------------------
# BOTÓN PARA LIMPIAR BASE DE DATOS (RESTARAURAR DESDE ORIGINAL)
# -------------------------------------------------------------------------------------
if st.button("Limpiar Base de Datos"):
    if os.path.exists(ORIGINAL_FILE):
        shutil.copy(ORIGINAL_FILE, STOCK_FILE)
        st.success("✅ Base de datos restaurada al estado original.")
        st.rerun()
    else:
        st.error("❌ No se encontró la copia original en 'versions/Stock_Original.xlsx'.")

# -------------------------------------------------------------------------------------
# SI HAY DATOS, TRABAJAMOS CON ELLOS
# -------------------------------------------------------------------------------------
if data_dict:
    st.title("📦 Control de Stock del Hospital")

    # Elegir la hoja
    sheet_name = st.selectbox("Selecciona la categoría de stock:", list(data_dict.keys()))
    df = data_dict[sheet_name].copy()

    # ==============================
    # CONVERSIONES A TIPOS (TU REQUISITO)
    # ==============================
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

    # Caducidad, Fecha Pedida, Fecha Llegada -> datetime
    for col in ["Caducidad", "Fecha Pedida", "Fecha Llegada"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Restantes -> int
    if "Restantes" in df.columns:
        df["Restantes"] = pd.to_numeric(df["Restantes"], errors="coerce").fillna(0).astype(int)

    # Sitio almacenaje -> str
    if "Sitio almacenaje" in df.columns:
        df["Sitio almacenaje"] = df["Sitio almacenaje"].astype(str)

    # ==============================
    # MOSTRAR TABLA (SIN PyArrow) USANDO st.write
    # ==============================
    st.write(f"📋 Mostrando datos de: **{sheet_name}**")
    st.write(df)

    # ==============================
    # CREAR COL. EXHIBICIÓN => Nombre producto + (Ref. Fisher)
    # ==============================
    if "Nombre producto" in df.columns and "Ref. Fisher" in df.columns:
        display_series = df.apply(
            lambda row: f"{row['Nombre producto']} ({row['Ref. Fisher']})",
            axis=1
        )
    else:
        display_series = df.iloc[:, 0].astype(str)

    reactivo = st.selectbox("Selecciona el reactivo a modificar:", display_series.unique())
    row_index = display_series[display_series == reactivo].index[0]

    # ==============================
    # CARGAR VALORES ACTUALES
    # ==============================
    def get_val(col, default=None):
        return df.at[row_index, col] if col in df.columns else default

    lote_actual = get_val("NºLote", 0)
    caducidad_actual = get_val("Caducidad", None)
    fecha_pedida_actual = get_val("Fecha Pedida", None)
    fecha_llegada_actual = get_val("Fecha Llegada", None)
    sitio_almacenaje_actual = get_val("Sitio almacenaje", "")

    st.subheader("✏️ Modificar Reactivo")

    # Nº Lote
    lote_nuevo = st.number_input("Nº de Lote", value=int(lote_actual) if pd.notna(lote_actual) else 0, step=1)
    # Caducidad
    caducidad_nueva = st.date_input("Caducidad", value=caducidad_actual if pd.notna(caducidad_actual) else None)
    # Fecha Pedida
    fecha_pedida_nueva = st.date_input("Fecha Pedida", value=fecha_pedida_actual if pd.notna(fecha_pedida_actual) else None)
    # Fecha Llegada
    fecha_llegada_nueva = st.date_input("Fecha Llegada", value=fecha_llegada_actual if pd.notna(fecha_llegada_actual) else None)

    # Sitio Almacenaje principal
    opciones_sitio = ["Congelador 1", "Congelador 2", "Frigorífico", "Tª Ambiente"]
    # Intentar extraer la parte principal
    sitio_principal = sitio_almacenaje_actual.split(" - ")[0] if " - " in sitio_almacenaje_actual else sitio_almacenaje_actual
    if sitio_principal not in opciones_sitio:
        sitio_principal = opciones_sitio[0]
    sitio_top = st.selectbox("Sitio de Almacenaje", opciones_sitio, index=opciones_sitio.index(sitio_principal))

    # Subopción
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

    # ---------------------------------------------------------------------------------
    # GUARDAR CAMBIOS
    # ---------------------------------------------------------------------------------
    def crear_nueva_version():
        """Crea un archivo en la carpeta 'versions' con fecha/hora."""
        fecha_hora = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        return os.path.join(VERSIONS_DIR, f"Stock_{fecha_hora}.xlsx")

    if st.button("Guardar Cambios"):
        # 1) Creamos un nuevo archivo con fecha/hora en 'versions'
        new_file = crear_nueva_version()

        # 2) Actualizamos df
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

        # 3) Guardamos la versión con fecha/hora
        with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
            for sht, df_sheet in data_dict.items():
                if sht == sheet_name:
                    df.to_excel(writer, sheet_name=sht, index=False)
                else:
                    df_sheet.to_excel(writer, sheet_name=sht, index=False)

        # 4) Guardamos TAMBIÉN en Stock_Modificadov1.xlsx (archivo de trabajo)
        with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
            for sht, df_sheet in data_dict.items():
                if sht == sheet_name:
                    df.to_excel(writer, sheet_name=sht, index=False)
                else:
                    df_sheet.to_excel(writer, sheet_name=sht, index=False)

        st.success(f"✅ Cambios guardados en '{new_file}' y '{STOCK_FILE}'.")
        st.rerun()
