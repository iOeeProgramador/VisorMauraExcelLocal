# Procesador Excel ZIP

Aplicación web desarrollada en **Streamlit** para procesar un archivo ZIP que contiene varios libros de Excel. Esta herramienta consolida datos de órdenes, inventario, estado, precios y gestión en un único archivo Excel para su análisis.

## 🔧 ¿Qué hace esta aplicación?

- Carga un archivo `.zip` con los siguientes archivos Excel:
  - `ORDENES.xlsx`
  - `INVENTARIO.xlsx`
  - `ESTADO.xlsx`
  - `PRECIOS.xlsx`
  - `GESTION.xlsx`
- Realiza joins inteligentes entre los archivos.
- Calcula días restantes según fechas de orden.
- Muestra un resumen por responsable de gestión.
- Permite descargar el resultado combinado como `DatosCombinados.xlsx`.

## 🚀 Cómo usar

```bash
# Clona el repositorio
git clone https://github.com/tu-usuario/procesador-excel-zip.git
cd procesador-excel-zip

# Instala las dependencias
pip install -r requirements.txt

# Ejecuta la app
streamlit run app.py
