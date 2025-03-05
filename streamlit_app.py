import streamlit as st
import pandas as pd
import datetime
import shutil
import os

# Configuración de la aplicación
st.title("📦 Control de Stock del Hospital")

# Ruta del archivo principal
file_path = "Stock_Modificadov1.xlsx"
backup_folder = "backups"
os.makedirs(backup_folder, exist_ok=True)  # Crear carpeta de backups si no existe

# Verificar que openpyxl está instalado
try:
    import openpyxl
except ImportError:
    st.error("❌ Falta la librería 'openpyxl'. Instálala con 'pip install openpyxl'.")

# Función para cargar los datos desde Excel
def load_data():
    try:
        # Carga TODAS las hojas en un diccionario
        return pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
    except FileNotFoundError:
        st.error("❌ No se encontró el archivo de la base de datos. Asegúrate de que 'Stock_Modificadov1.xlsx' está en el directorio.")
        return None
    except Exception as e:
        st.error(f"❌ Error al cargar la base de datos: {e}")
        return None

data = load_data()

if data:
    # Seleccionar la hoja a visualizar
    sheet_name = st.selectbox("Selecciona la categoría de stock:", list(data.keys()))
    df = data[sheet_name].copy()  # Copia para no pisar el original

    # ==============================
    # CONVERSIONES A TIPOS SOLICITADOS
    # ==============================

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

    # Caducidad -> fecha
    if "Caducidad" in df.columns:
        df["Caducidad"] = pd.to_datetime(df["Caducidad"], errors="coerce")

    # Fecha Pedida -> fecha
    if "Fecha Pedida" in df.columns:
        df["Fecha Pedida"] = pd.to_datetime(df["Fecha Pedida"], errors="coerce")

    # Fecha Llegada -> fecha
    if "Fecha Llegada" in df.columns:
        df["Fecha Llegada"] = pd.to_datetime(df["Fecha Llegada"], errors="coerce")

    # Restantes -> int
    if "Restantes" in df.columns:
        df["Restantes"] = pd.to_numeric(df["Restantes"], errors="coerce").fillna(0).astype(int)

    # Sitio almacenaje -> str
    if "Sitio almacenaje" in df.columns:
        df["Sitio almacenaje"] = df["Sitio almacenaje"].astype(str)

    # ==============================
    # MOSTRAR RESULTADO SIN USAR DATAFRAME (para evitar PyArrow)
    # ==============================
    st.write(f"📋 Mostrando datos de: **{sheet_name}**")
    st.write("🔎 Tipos de datos actuales:")
    st.write(df.dtypes)   # Verificamos tipos
    st.write("📋 Vista de la tabla con `st.write` en lugar de `st.dataframe`:")
    st.write(df)

    # ==============================
    # SELECCIONAR REACTIVO
    # ==============================
    reactivo = st.selectbox("Selecciona el reactivo a modificar:", df.iloc[:, 0].dropna().unique())

    # Obtener la fila del reactivo seleccionado
    row_index = df[df.iloc[:, 0] == reactivo].index[0]

    # Formulario para actualizar datos
    st.subheader("✏️ Modificar Reactivo")

    # Cargamos valores existentes (usamos get para evitar errores si no existe la columna)
    lote_actual = df.at[row_index, "NºLote"] if "NºLote" in df.columns else 0
    caducidad_actual = df.at[row_index, "Caducidad"] if "Caducidad" in df.columns else None
    fecha_pedida_actual = df.at[row_index, "Fecha Pedida"] if "Fecha Pedida" in df.columns else None
    fecha_llegada_actual = df.at[row_index, "Fecha Llegada"] if "Fecha Llegada" in df.columns else None
    sitio_almacenaje_actual = df.at[row_index, "Sitio almacenaje"] if "Sitio almacenaje" in df.columns else ""

    lote = st.number_input("Nº de Lote", value=int(lote_actual) if pd.notna(lote_actual) else 0, step=1)
    caducidad_val = st.date_input("Caducidad", value=caducidad_actual if pd.notna(caducidad_actual) else None)
    fecha_pedida_val = st.date_input("Fecha Pedida", value=fecha_pedida_actual if pd.notna(fecha_pedida_actual) else None)
    fecha_llegada_val = st.date_input("Fecha Llegada", value=fecha_llegada_actual if pd.notna(fecha_llegada_actual) else None)
    sitio_almacenaje_val = st.text_input("Sitio de Almacenaje", value=str(sitio_almacenaje_actual) if pd.notna(sitio_almacenaje_actual) else "")

    # Función para hacer copias de seguridad cada vez que se haga un cambio
    def guardar_copia_seguridad():
        fecha_hora = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        backup_file = os.path.join(backup_folder, f"Stock_{fecha_hora}.xlsx")
        shutil.copy(file_path, backup_file)
        st.success(f"✅ Copia de seguridad guardada: {backup_file}")

    # Guardar cambios
    if st.button("Guardar Cambios"):
        guardar_copia_seguridad()  # Hacer una copia antes de modificar

        # Actualizar los valores en el DataFrame
        if "NºLote" in df.columns:
            df.at[row_index, "NºLote"] = int(lote)
        if "Caducidad" in df.columns:
            df.at[row_index, "Caducidad"] = caducidad_val
        if "Fecha Pedida" in df.columns:
            df.at[row_index, "Fecha Pedida"] = fecha_pedida_val
        if "Fecha Llegada" in df.columns:
            df.at[row_index, "Fecha Llegada"] = fecha_llegada_val
        if "Sitio almacenaje" in df.columns:
            df.at[row_index, "Sitio almacenaje"] = str(sitio_almacenaje_val)

        # Guardar los cambios en Excel
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            for sheet, data_sheet in data.items():
                if sheet == sheet_name:
                    # Aseguramos que la df local actualizada (df) sea la que se escribe
                    df.to_excel(writer, sheet_name=sheet, index=False)
                else:
                    data_sheet.to_excel(writer, sheet_name=sheet, index=False)

        st.success("✅ Datos actualizados correctamente. Recarga la página para ver cambios.")
