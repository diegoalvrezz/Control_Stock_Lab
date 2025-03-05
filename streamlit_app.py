import streamlit as st
import pandas as pd
import datetime
import shutil
import os

# -------------------------------------------------------------------------------------
# CONFIGURACIÓN DE LA APP
# -------------------------------------------------------------------------------------
st.title("📦 Control de Stock del Hospital")

file_path = "Stock_Modificadov1.xlsx"               # Archivo actual de trabajo
backup_folder = "backups"
original_file = "Stock_Modificadov1_ORIGINAL.xlsx"  # Archivo original, para restaurar

os.makedirs(backup_folder, exist_ok=True)  # Crear carpeta de backups si no existe

# Verificar que openpyxl está instalado
try:
    import openpyxl
except ImportError:
    st.error("❌ Falta la librería 'openpyxl'. Instálala con 'pip install openpyxl'.")

# -------------------------------------------------------------------------------------
# FUNCIÓN PARA CARGAR DATOS DESDE EXCEL
# -------------------------------------------------------------------------------------
def load_data():
    try:
        return pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
    except FileNotFoundError:
        st.error("❌ No se encontró el archivo de la base de datos.")
        return None
    except Exception as e:
        st.error(f"❌ Error al cargar la base de datos: {e}")
        return None

data_dict = load_data()

# -------------------------------------------------------------------------------------
# BOTÓN PARA LIMPIAR LA BASE DE DATOS (RESTAURAR ARCHIVO ORIGINAL)
# -------------------------------------------------------------------------------------
if st.button("Limpiar Base de Datos"):
    if os.path.exists(original_file):
        shutil.copy(original_file, file_path)
        st.success("✅ Base de datos restaurada al estado original.")
        st.rerun()  # O st.experimental_rerun() en versiones antiguas
    else:
        st.error("❌ No se encontró el archivo original para restaurar.")

# -------------------------------------------------------------------------------------
# SI EXISTE LA BASE DE DATOS, SELECCIONAR HOJA
# -------------------------------------------------------------------------------------
if data_dict:
    sheet_name = st.selectbox("Selecciona la categoría de stock:", list(data_dict.keys()))
    df = data_dict[sheet_name].copy()

    # ---------------------------------------------------------------------------------
    # CONVERSIONES A TIPOS (según lo indicado)
    # ---------------------------------------------------------------------------------
    # Ref. Saturno -> int
    if "Ref. Saturno" in df.columns:
        df["Ref. Saturno"] = pd.to_numeric(df["Ref. Saturno"], errors="coerce").fillna(0).astype(int)

    # Ref. Fisher -> str (tiene letras y números)
    if "Ref. Fisher" in df.columns:
        df["Ref. Fisher"] = df["Ref. Fisher"].astype(str)

    # Nombre producto -> str
    if "Nombre producto" in df.columns:
        df["Nombre producto"] = df["Nombre producto"].astype(str)

    # Tª -> str (tiene letras y números)
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

    # ---------------------------------------------------------------------------------
    # MOSTRAR TABLA (SIN PyArrow) USANDO st.write
    # ---------------------------------------------------------------------------------
    st.write(f"📋 Mostrando datos de: **{sheet_name}**")
    st.write(df)

    # ---------------------------------------------------------------------------------
    # CREAR COLUMNA DE EXHIBICIÓN PARA ELEGIR REACTIVO:
    # Nombre producto (Ref. Fisher)
    # ---------------------------------------------------------------------------------
    if "Nombre producto" in df.columns and "Ref. Fisher" in df.columns:
        display_series = df.apply(
            lambda row: f"{row['Nombre producto']} ({row['Ref. Fisher']})",
            axis=1
        )
    else:
        # Si faltan columnas, usar la primera como fallback
        display_series = df.iloc[:, 0].astype(str)

    # ---------------------------------------------------------------------------------
    # SELECCIONAR REACTIVO DESDE ESA COLUMNA "display_series"
    # ---------------------------------------------------------------------------------
    reactivo = st.selectbox(
        "Selecciona el reactivo a modificar:",
        display_series.unique()
    )
    # Obtener el row_index real
    row_index = display_series[display_series == reactivo].index[0]

    # ---------------------------------------------------------------------------------
    # CARGAMOS VALORES ACTUALES DE LAS COLUMNAS PRINCIPALES (Evita KeyError si no existen)
    # ---------------------------------------------------------------------------------
    lote_actual = df.at[row_index, "NºLote"] if "NºLote" in df.columns else 0
    caducidad_actual = df.at[row_index, "Caducidad"] if "Caducidad" in df.columns else None
    fecha_pedida_actual = df.at[row_index, "Fecha Pedida"] if "Fecha Pedida" in df.columns else None
    fecha_llegada_actual = df.at[row_index, "Fecha Llegada"] if "Fecha Llegada" in df.columns else None
    sitio_almacenaje_actual = df.at[row_index, "Sitio almacenaje"] if "Sitio almacenaje" in df.columns else ""

    # ---------------------------------------------------------------------------------
    # FORMULARIO PARA MODIFICAR CADA DATO
    # ---------------------------------------------------------------------------------
    st.subheader("✏️ Modificar Reactivo")

    # Nº Lote
    lote_nuevo = st.number_input(
        "Nº de Lote", 
        value=int(lote_actual) if pd.notna(lote_actual) else 0, 
        step=1
    )
    # Caducidad
    caducidad_nueva = st.date_input(
        "Caducidad", 
        value=caducidad_actual if pd.notna(caducidad_actual) else None
    )
    # Fecha Pedida
    fecha_pedida_nueva = st.date_input(
        "Fecha Pedida", 
        value=fecha_pedida_actual if pd.notna(fecha_pedida_actual) else None
    )
    # Fecha Llegada
    fecha_llegada_nueva = st.date_input(
        "Fecha Llegada", 
        value=fecha_llegada_actual if pd.notna(fecha_llegada_actual) else None
    )

    # ---------------------------------------------------------------------------------
    # SITIO ALMACENAJE (PRIMER SELECTBOX)
    # ---------------------------------------------------------------------------------
    opciones_sitio = ["Congelador 1", "Congelador 2", "Frigorífico", "Tª Ambiente"]
    # Intentar extraer la parte principal (p.e. "Congelador 1") de "Congelador 1 - Cajón 3"
    try:
        sitio_principal = sitio_almacenaje_actual.split(" - ")[0]
        index_inicial = opciones_sitio.index(sitio_principal) if sitio_principal in opciones_sitio else 0
    except:
        index_inicial = 0
    sitio_top = st.selectbox("Sitio de Almacenaje", opciones_sitio, index=index_inicial)

    # ---------------------------------------------------------------------------------
    # SITIO ALMACENAJE (SUBSELECTBOX) SEGÚN ELECCIÓN
    # ---------------------------------------------------------------------------------
    subopcion = ""
    if sitio_top == "Congelador 1":
        cajones = [f"Cajón {i}" for i in range(1, 9)]  # 1..8
        subopcion = st.selectbox("Cajón", cajones)
    elif sitio_top == "Congelador 2":
        cajones = [f"Cajón {i}" for i in range(1, 7)]  # 1..6
        subopcion = st.selectbox("Cajón", cajones)
    elif sitio_top == "Frigorífico":
        # Balda 1..6 + Puerta
        baldas = [f"Balda {i}" for i in range(1, 7)] + ["Puerta"]
        subopcion = st.selectbox("Baldas", baldas)
    else:
        # Tª Ambiente => sin subopción
        subopcion = ""

    # Unimos sitio principal con la subopción
    if subopcion:
        sitio_almacenaje_nuevo = f"{sitio_top} - {subopcion}"
    else:
        sitio_almacenaje_nuevo = sitio_top

    # ---------------------------------------------------------------------------------
    # FUNCIÓN PARA GUARDAR COPIA DE SEGURIDAD
    # ---------------------------------------------------------------------------------
    def guardar_copia_seguridad():
        fecha_hora = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        backup_file = os.path.join(backup_folder, f"Stock_{fecha_hora}.xlsx")
        shutil.copy(file_path, backup_file)
        st.success(f"✅ Copia de seguridad guardada: {backup_file}")

    # ---------------------------------------------------------------------------------
    # BOTÓN: GUARDAR CAMBIOS
    # ---------------------------------------------------------------------------------
    if st.button("Guardar Cambios"):
        guardar_copia_seguridad()

        # Actualizamos en df:
        if "NºLote" in df.columns:
            df.at[row_index, "NºLote"] = int(lote_nuevo)
        if "Caducidad" in df.columns:
            # Asignar datetime, evitando advertencia de strings en columna datetime
            df.at[row_index, "Caducidad"] = pd.to_datetime(caducidad_nueva)
        if "Fecha Pedida" in df.columns:
            df.at[row_index, "Fecha Pedida"] = pd.to_datetime(fecha_pedida_nueva)
        if "Fecha Llegada" in df.columns:
            df.at[row_index, "Fecha Llegada"] = pd.to_datetime(fecha_llegada_nueva)
        if "Sitio almacenaje" in df.columns:
            df.at[row_index, "Sitio almacenaje"] = sitio_almacenaje_nuevo

        # Guardar en Excel
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            for sheet, data_sheet in data_dict.items():
                if sheet == sheet_name:
                    df.to_excel(writer, sheet_name=sheet, index=False)
                else:
                    data_sheet.to_excel(writer, sheet_name=sheet, index=False)

        st.success("✅ Datos actualizados correctamente.")
        st.write("🔍 Vista previa de los datos tras actualizar:")
        st.write(df)

        # Esperar un momento antes de recargar (opcional)
        import time
        time.sleep(2)

        st.rerun()
