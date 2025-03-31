
# Aplicación de Control de Stock - Laboratorio de Patología Molecular

Esta aplicación ha sido desarrollada para facilitar la gestión de inventario de reactivos en un entorno hospitalario, permitiendo un control detallado de cada producto en función de su lote, fechas asociadas, ubicación de almacenamiento y estado del stock.

## Funcionalidades principales

- Interfaz web intuitiva construida con Streamlit, orientada a usuarios con conocimientos informáticos básicos.
- Gestión de stock por paneles técnicos (FOCUS, OCA, OCA PLUS).
- Control de lotes con campos como número de lote, caducidad, fechas de pedido y llegada, ubicación y unidades disponibles.
- Registro automático de cada modificación en una base de datos histórica (Base B) para trazabilidad.
- Posibilidad de consultar versiones anteriores, descargar hojas de trabajo y eliminar versiones antiguas.
- Filtrado avanzado de reactivos limitantes y compartidos.
- Funcionalidad para registrar consumo y actualizar stock en tiempo real.

## Consideraciones de uso

- La aplicación está diseñada para uso exclusivo **local** y no sube ninguna información a servidores externos.
- Para conservar la trazabilidad, cada modificación genera automáticamente una copia de seguridad en formato Excel.
- Para asegurar la confidencialidad, los archivos se suben y almacenan únicamente de forma local e independiente.

## Estructura del repositorio

- app.py: Código principal de la aplicación.
- versions/: Carpeta local donde se almacenan las versiones de la base de datos A.
- versions_b/: Carpeta local donde se almacenan las versiones históricas (base B).
- plantilla_base_datos.xlsx: Plantilla genérica de la base de datos sin datos sensibles.

## Mecanismo de guardado y versiones

Cada vez que se realiza una modificación en la base de datos A, se genera automáticamente una nueva versión del archivo Excel con marca de tiempo y se guarda en la carpeta versions/. Paralelamente, si se registra una modificación en el historial (base B), también se guarda una copia en versions_b/, manteniendo un registro completo de todos los cambios para asegurar la trazabilidad.

Los archivos se organizan en subcarpetas mensuales (YYYY_MM_Mes) para facilitar su consulta y gestión.

## Requisitos

- Python 3.9 o superior
- Paquetes: streamlit, pandas, openpyxl, streamlit_authenticator, pytz, entre otros (ver requirements.txt)

## Manual de uso

Para instrucciones detalladas sobre el funcionamiento, instalación y buenas prácticas, consulte el manual de usuario incluido en el repositorio:

**[Manual_Control_Stock.pdf](./Manual_Control_Stock.pdf)**
