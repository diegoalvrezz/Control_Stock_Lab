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
        return pd.read_excel(file_path, sheet_name=None, engine="openpyxl")  # Cargar todas las hojas en un diccionario
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
    df = data[sheet_name]
    
    # Mostrar los datos actuales
    st.write(f"📋 Mostrando datos de: **{sheet_name}**")
    st.dataframe(df)
    
    # Seleccionar reactivo a modificar
    reactivo = st.selectbox("Selecciona el reactivo a modificar:", df.iloc[:, 0].dropna().tolist())
    
    # Obtener la fila del reactivo seleccionado
    row_index = df[df.iloc[:, 0] == reactivo].index[0]
    
    # Formulario para actualizar datos
    st.subheader("✏️ Modificar Reactivo")
    lote = st.number_input("Nº de Lote", value=int(df.at[row_index, "NºLote"]) if pd.notna(df.at[row_index, "NºLote"]) else 0, step=1)
    caducidad = st.date_input("Caducidad", value=pd.to_datetime(df.at[row_index, "Caducidad"], errors='coerce') if pd.notna(df.at[row_index, "Caducidad"]) else None)
    fecha_pedida = st.date_input("Fecha Pedida", value=pd.to_datetime(df.at[row_index, "Fecha Pedida"], errors='coerce') if pd.notna(df.at[row_index, "Fecha Pedida"]) else None)
    fecha_llegada = st.date_input("Fecha Llegada", value=pd.to_datetime(df.at[row_index, "Fecha Llegada"], errors='coerce') if pd.notna(df.at[row_index, "Fecha Llegada"]) else None)
    sitio_almacenaje = st.text_input("Sitio de Almacenaje", value=df.at[row_index, "Sitio almacenaje"] if pd.notna(df.at[row_index, "Sitio almacenaje"]) else "")
    
    # Asegurar tipos correctos antes de guardar
    df["NºLote"] = pd.to_numeric(df["NºLote"], errors="coerce").astype("Int64")
    df["Caducidad"] = pd.to_datetime(df["Caducidad"], errors="coerce")
    df["Fecha Pedida"] = pd.to_datetime(df["Fecha Pedida"], errors="coerce")
    df["Fecha Llegada"] = pd.to_datetime(df["Fecha Llegada"], errors="coerce")
    df["Sitio almacenaje"] = df["Sitio almacenaje"].astype(str)
    
    # Función para hacer copias de seguridad cada vez que se haga un cambio
    def guardar_copia_seguridad():
        fecha_hora = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        backup_file = os.path.join(backup_folder, f"Stock_{fecha_hora}.xlsx")
        shutil.copy(file_path, backup_file)
        st.success(f"✅ Copia de seguridad guardada: {backup_file}")
    
    # Guardar cambios
    if st.button("Guardar Cambios"):
        guardar_copia_seguridad()  # Hacer una copia antes de modificar
        
        # Actualizar los valores en la base de datos
        df.at[row_index, "NºLote"] = int(lote) if pd.notna(lote) else None
        df.at[row_index, "Caducidad"] = caducidad.strftime("%Y-%m-%d") if pd.notna(caducidad) else None
        df.at[row_index, "Fecha Pedida"] = fecha_pedida.strftime("%Y-%m-%d") if pd.notna(fecha_pedida) else None
        df.at[row_index, "Fecha Llegada"] = fecha_llegada.strftime("%Y-%m-%d") if pd.notna(fecha_llegada) else None
        df.at[row_index, "Sitio almacenaje"] = str(sitio_almacenaje) if pd.notna(sitio_almacenaje) else ""
        
        # Guardar los cambios en Excel
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            for sheet, data in data.items():
                if sheet == sheet_name:
                    df.to_excel(writer, sheet_name=sheet, index=False)
                else:
                    data.to_excel(writer, sheet_name=sheet, index=False)
        
        st.success("✅ Datos actualizados correctamente")
        st.rerun()  # Recargar la app para mostrar los cambios
