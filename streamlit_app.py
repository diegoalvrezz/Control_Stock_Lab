import streamlit as st
import pandas as pd
import numpy as np
import datetime
import shutil
import os
from io import BytesIO

# -------------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA
# -------------------------------------------------------------------------
st.set_page_config(page_title="Control de Stock con Alarmas", layout="centered")

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

# -------------------------------------------------------------------------
# FUNCIÓN PARA CONVERSIONES DE TIPOS (eliminamos 'Restantes')
# -------------------------------------------------------------------------
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
        df["NºLote"] = pd.to_numeric(df["NºLote"], errors="coerce").fillna(0).astype(int)
    for col in ["Caducidad", "Fecha Pedida", "Fecha Llegada"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    if "Sitio almacenaje" in df.columns:
        df["Sitio almacenaje"] = df["Sitio almacenaje"].astype(str)
    if "Stock" in df.columns:
        df["Stock"] = pd.to_numeric(df["Stock"], errors="coerce").fillna(0).astype(int)
    return df

def crear_nueva_version_filename():
    fecha_hora = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return os.path.join(VERSIONS_DIR, f"Stock_{fecha_hora}.xlsx")

def generar_excel_en_memoria(df_act: pd.DataFrame, sheet_nm="Hoja1"):
    """Generar un Excel en memoria (bytes) para descargar."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_act.to_excel(writer, index=False, sheet_name=sheet_nm)
    output.seek(0)
    return output.getvalue()

# -------------------------------------------------------------------------
# LÓGICA PARA RESALTAR FILAS ROJAS O NARANJAS (opacidad reducida)
# -------------------------------------------------------------------------
def highlight_row(row):
    """
    - ALARMA ROJA => (stock=0) y (Fecha Pedida es None)
    - ALARMA NARANJA => (stock=0) y (Fecha Pedida != None)
    """
    stock_val = row.get("Stock", None)
    fecha_pedida_val = row.get("Fecha Pedida", None)

    if pd.isna(stock_val):
        return [""] * len(row)

    if stock_val == 0:
        if pd.isna(fecha_pedida_val):
            # color rojo translúcido
            return ['background-color: rgba(255, 0, 0, 0.2); color: black'] * len(row)
        else:
            # color naranja translúcido
            return ['background-color: rgba(255, 165, 0, 0.2); color: black'] * len(row)

    return [""] * len(row)

# -------------------------------------------------------------------------
# INTERFAZ
# -------------------------------------------------------------------------
st.title("📦 Control de Stock con Alarmas (Fechas + Hora)")

# SIDEBAR: Opciones de la Base de Datos
with st.sidebar:
    st.markdown("## Opciones de la Base de Datos")

    with st.expander("🔎 Ver / Gestionar versiones guardadas"):
        files = sorted(os.listdir(VERSIONS_DIR))
        versions_no_original = [f for f in files if f != "Stock_Original.xlsx"]
        if versions_no_original:
            version_sel = st.selectbox("Selecciona una versión:", versions_no_original)
            # Añadimos un campo de confirmación de acción
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
                # Mensaje de confirmación con un "checkbox"
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
                        st.error("Marca la casilla de confirmación para eliminar la versión.")
        else:
            st.write("No hay versiones guardadas (excepto la original).")

        if st.button("Eliminar TODAS las versiones (excepto original)"):
            for f in versions_no_original:
                try:
                    os.remove(os.path.join(VERSIONS_DIR, f))
                except:
                    pass
            st.info("Todas las versiones (excepto la original) han sido eliminadas.")
            st.rerun()

        # Botón extra: eliminar todas las versiones excepto la última y la original
        if st.button("Eliminar TODAS las versiones excepto la última y la original"):
            if len(versions_no_original) > 1:
                sorted_vers = sorted(versions_no_original)
                last_version = sorted_vers[-1]  # la última alfabéticamente
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

    # Botón para limpiar base
    if st.button("Limpiar Base de Datos"):
        st.write("¿Seguro que quieres limpiar la base de datos?")
        if st.checkbox("Sí, confirmar limpieza."):
            if os.path.exists(ORIGINAL_FILE):
                shutil.copy(ORIGINAL_FILE, STOCK_FILE)
                st.success("✅ Base de datos restaurada al estado original.")
                st.rerun()
            else:
                st.error("❌ No se encontró la copia original en 'versions/Stock_Original.xlsx'.")

    st.markdown("---")
    # Expander Alarmas
    with st.expander("⚠️ Alarmas"):
        """
        Alarma:
        - Roja => Stock=0 y Fecha Pedida=None
        - Naranja => Stock=0 y Fecha Pedida!=None
        """
        if data_dict:
            for nombre_hoja, df_hoja in data_dict.items():
                df_hoja = enforce_types(df_hoja)
                if "Stock" in df_hoja.columns:
                    df_cero = df_hoja[df_hoja["Stock"] == 0]
                    if not df_cero.empty:
                        st.markdown(f"**Hoja: {nombre_hoja}**")
                        for idx, fila in df_cero.iterrows():
                            fecha_ped = fila.get("Fecha Pedida", None)
                            producto = fila.get("Nombre producto", f"Fila {idx}")
                            fisher = fila.get("Ref. Fisher", "")
                            if pd.isna(fecha_ped):
                                # Alarma Roja
                                st.error(f"[{producto} ({fisher})] => Stock=0 => ALARMA ROJA (No pedido)")
                            else:
                                # Alarma Naranja
                                st.warning(f"[{producto} ({fisher})] => Stock=0 => ALARMA NARANJA (Pedido)")
        else:
            st.info("No se han cargado datos o no existe la base de datos.")

    st.markdown("---")
    # Botón Reactivo Agotado
    st.markdown("### Reactivo Agotado en el Laboratorio")
    st.write("Selecciona la hoja y el reactivo, y cuántas unidades restar del stock.")
    if data_dict:
        hojas_opciones = list(data_dict.keys())
        hoja_para_consumo = st.selectbox("Selecciona la categoría de stock:", hojas_opciones, key="hoja_agotado")
        df_temp = data_dict[hoja_para_consumo].copy()
        df_temp = enforce_types(df_temp)

        if "Nombre producto" in df_temp.columns and "Ref. Fisher" in df_temp.columns:
            display_series_temp = df_temp.apply(
                lambda row: f"{row['Nombre producto']} ({row['Ref. Fisher']})", axis=1
            )
        else:
            display_series_temp = df_temp.iloc[:, 0].astype(str)

        reactivo_agotado = st.selectbox("Reactivo a Consumir:", display_series_temp.unique(), key="reactivo_agotado")
        row_idx_agotado = display_series_temp[display_series_temp == reactivo_agotado].index[0]
        stock_actual_agotado = df_temp.at[row_idx_agotado, "Stock"] if "Stock" in df_temp.columns else 0

        uds_consumidas = st.number_input("Unidades consumidas en Lab", min_value=0, step=1, key="uds_consumidas")

        if st.button("Registrar Consumo", key="consumir_lab"):
            nuevo_stock = max(0, stock_actual_agotado - uds_consumidas)
            df_temp.at[row_idx_agotado, "Stock"] = nuevo_stock
            st.warning(f"Se han consumido {uds_consumidas} uds. Stock final => {nuevo_stock}")

            new_file2 = crear_nueva_version_filename()
            data_dict[hoja_para_consumo] = df_temp  # actualizamos en memoria

            with pd.ExcelWriter(new_file2, engine="openpyxl") as writer:
                for sht, df_sht in data_dict.items():
                    df_sht.to_excel(writer, sheet_name=sht, index=False)

            with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
                for sht, df_sht in data_dict.items():
                    df_sht.to_excel(writer, sheet_name=sht, index=False)

            st.success(f"✅ Stock actualizado. Guardado en '{new_file2}' y '{STOCK_FILE}'.")

            excel_bytes_2 = generar_excel_en_memoria(df_temp, sheet_nm=hoja_para_consumo)
            st.download_button(
                label="Descargar Excel tras el consumo",
                data=excel_bytes_2,
                file_name="Reporte_Stock_Agotado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.rerun()
    else:
        st.info("No se han cargado datos o no existe la base de datos.")


# -------------------------------------------------------------------------
# CUERPO PRINCIPAL
# -------------------------------------------------------------------------
if data_dict:
    sheet_name = st.selectbox("Selecciona la categoría de stock (principal):", list(data_dict.keys()))
    df = data_dict[sheet_name].copy()
    df = enforce_types(df)

    st.markdown(f"### Hoja seleccionada: **{sheet_name}**")

    # Reemplazar NaT / NaN por "-"
    df_display = df.copy()
    # fillna("-") => convertimos tanto nan como NaT en "-"
    df_display = df_display.fillna("-")

    styled_df = df_display.style.apply(highlight_row, axis=1)
    st.write(styled_df.to_html(), unsafe_allow_html=True)

    # Selecciona reactivo a modificar
    if "Nombre producto" in df.columns and "Ref. Fisher" in df.columns:
        display_series = df.apply(lambda row: f"{row['Nombre producto']} ({row['Ref. Fisher']})", axis=1)
    else:
        display_series = df.iloc[:, 0].astype(str)

    reactivo = st.selectbox("Selecciona Reactivo a Modificar:", display_series.unique())
    row_index = display_series[display_series == reactivo].index[0]

    def get_val(col, default=None):
        return df.at[row_index, col] if col in df.columns else default

    # Cargar valores
    lote_actual = get_val("NºLote", 0)
    caducidad_actual = get_val("Caducidad", None)
    fecha_pedida_actual = get_val("Fecha Pedida", None)
    fecha_llegada_actual = get_val("Fecha Llegada", None)
    sitio_almacenaje_actual = get_val("Sitio almacenaje", "")
    uds_actual = get_val("Uds.", 0)
    stock_actual = get_val("Stock", 0)

    st.markdown("#### Modificar Atributos de Reactivo")

    colA, colB, colC, colD = st.columns([1,1,1,1])
    with colA:
        lote_nuevo = st.number_input("Nº de Lote", value=int(lote_actual), step=1)
        caducidad_nueva = st.date_input("Caducidad", value=caducidad_actual if pd.notna(caducidad_actual) else None)

    with colB:
        # Fecha pedida => date + time
        fecha_pedida_date = st.date_input("Fecha Pedida (fecha)", value=fecha_pedida_actual.date() if pd.notna(fecha_pedida_actual) else None, key="fp_date")
        fecha_pedida_time = st.time_input("Hora Pedida", value=fecha_pedida_actual.time() if pd.notna(fecha_pedida_actual) else datetime.time(0, 0), key="fp_time")

    with colC:
        # Fecha llegada => date + time
        fecha_llegada_date = st.date_input("Fecha Llegada (fecha)", value=fecha_llegada_actual.date() if pd.notna(fecha_llegada_actual) else None, key="fl_date")
        fecha_llegada_time = st.time_input("Hora Llegada", value=fecha_llegada_actual.time() if pd.notna(fecha_llegada_actual) else datetime.time(0, 0), key="fl_time")

    with colD:
        st.write("")  # Espacio
        st.write("")
        # Botón para refrescar
        if st.button("Refrescar Página"):
            st.rerun()

    # Unimos date+time a un Timestamp
    fecha_pedida_nueva = None
    if fecha_pedida_date is not None:
        fecha_pedida_nueva = datetime.datetime.combine(fecha_pedida_date, fecha_pedida_time)

    fecha_llegada_nueva = None
    if fecha_llegada_date is not None:
        fecha_llegada_nueva = datetime.datetime.combine(fecha_llegada_date, fecha_llegada_time)

    st.write("Sitio de Almacenaje")
    opciones_sitio = ["Congelador 1", "Congelador 2", "Frigorífico", "Tª Ambiente"]
    sitio_principal = sitio_almacenaje_actual.split(" - ")[0] if " - " in sitio_almacenaje_actual else sitio_almacenaje_actual
    if sitio_principal not in opciones_sitio:
        sitio_principal = opciones_sitio[0]
    sitio_top = st.selectbox("Tipo de Almacenaje", opciones_sitio, index=opciones_sitio.index(sitio_principal))

    subopcion = ""
    if sitio_top == "Congelador 1":
        cajones = [f"Cajón {i}" for i in range(1, 9)]
        subopcion = st.selectbox("Cajón (1 Arriba, 8 Abajo)", cajones)
    elif sitio_top == "Congelador 2":
        cajones = [f"Cajón {i}" for i in range(1, 7)]
        subopcion = st.selectbox("Cajón (1 Arriba, 6 Abajo)", cajones)
    elif sitio_top == "Frigorífico":
        baldas = [f"Balda {i}" for i in range(1, 8)] + ["Puerta"]
        subopcion = st.selectbox("Baldas (1 Arriba, 7 Abajo)", baldas)
    elif sitio_top == "Tª Ambiente":
        comentario = st.text_input("Comentario (opcional)")
        subopcion = comentario.strip()

    if subopcion:
        sitio_almacenaje_nuevo = f"{sitio_top} - {subopcion}"
    else:
        sitio_almacenaje_nuevo = sitio_top

    # Botón Guardar Cambios
    if st.button("Guardar Cambios"):
        # Si se introduce Fecha Llegada, borramos Fecha Pedida => "ya llegó"
        if pd.notna(fecha_llegada_nueva):
            fecha_pedida_nueva = pd.NaT  # la pedida se vacía

        # Sumar Stock si la fecha de llegada cambió
        if "Stock" in df.columns:
            if fecha_llegada_nueva != fecha_llegada_actual and pd.notna(fecha_llegada_nueva):
                df.at[row_index, "Stock"] = stock_actual + uds_actual
                st.info(f"Sumadas {uds_actual} uds al stock. Nuevo stock => {stock_actual + uds_actual}")

        new_file = crear_nueva_version_filename()

        # Actualizar df en data_dict
        if "NºLote" in df.columns:
            df.at[row_index, "NºLote"] = int(lote_nuevo)
        if "Caducidad" in df.columns:
            df.at[row_index, "Caducidad"] = pd.to_datetime(caducidad_nueva)
        if "Fecha Pedida" in df.columns:
            df.at[row_index, "Fecha Pedida"] = fecha_pedida_nueva
        if "Fecha Llegada" in df.columns:
            df.at[row_index, "Fecha Llegada"] = fecha_llegada_nueva
        if "Sitio almacenaje" in df.columns:
            df.at[row_index, "Sitio almacenaje"] = sitio_almacenaje_nuevo

        data_dict[sheet_name] = df

        # Guardar versión
        with pd.ExcelWriter(new_file, engine="openpyxl") as writer:
            for sht, df_sheet in data_dict.items():
                df_sheet.to_excel(writer, sheet_name=sht, index=False)

        # Guardar en STOCK_FILE
        with pd.ExcelWriter(STOCK_FILE, engine="openpyxl") as writer:
            for sht, df_sheet in data_dict.items():
                df_sheet.to_excel(writer, sheet_name=sht, index=False)

        st.success(f"✅ Cambios guardados en '{new_file}' y '{STOCK_FILE}'.")

        excel_bytes = generar_excel_en_memoria(df, sheet_nm=sheet_name)
        st.download_button(
            label="Descargar Excel modificado",
            data=excel_bytes,
            file_name="Reporte_Stock.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.rerun()
