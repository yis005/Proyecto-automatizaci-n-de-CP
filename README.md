# Proyecto-automatizaci-n-de-CP
Script en Python que recorre carpetas con PDFs DIAN, extrae facturas (campo 40) y valores (campo 52), suma montos de facturas repetidas y exporta un Excel con datos consolidados, también permitirá enviar en automático via Email correo a los clientes que aún no han notificado CP.

📑 Proyecto: Extracción de información requerida desde PDFs DIAN
📌 Descripción
Este proyecto en Python automatiza la extracción de datos desde certificados PDF emitidos por la DIAN.
El script recorre carpetas organizadas por año y cliente, identifica las facturas (campo 40. No. Factura) y sus valores (campo 52. Valor Total), consolida las líneas repetidas sumando los montos, y exporta los resultados a un archivo Excel.
Además, incluye información general del documento como el número de formulario y el código de barras, que se repiten en cada fila para mantener el contexto.
Luego de consolidar la información, permite enviar en automático correos a los cleintes que aún no han notificado los CP

📂 Estructura de carpetas
Los PDFs se organizan en carpetas por año y cliente, por ejemplo:
CP/
 ├── 2025/
 │    ├── CI clienteA/
 │    │    ├── factura1.pdf
 │    │    └── factura2.pdf
 │    └── CI clienteB/
 │         └── factura3.pdf
 └── 2026/
      ├── CI pablito/
      │    ├── archivo1.pdf
      │    └── archivo2.pdf
      └── CI clienteC/
           └── archivo4.pdf


👉 El script puede procesar:
- Una carpeta específica de cliente (CP/2026/CI pablito)
- O todo un año completo (CP/2026)

⚙️ Instalación
- Clona este repositorio:
git clone https://github.com/tuusuario/tu-repo.git
- Instala las dependencias:
pip install pdfplumber pandas



🚀 Uso
Ejecuta el script indicando la carpeta base y el nombre del archivo Excel de salida:
python procesar_pdfs.py


Por defecto:
- Carpeta base: documentos_pdf
- Archivo de salida: resultado.xlsx
Puedes personalizarlo:
procesar_pdfs(base_dir="CP/2026/CI pablito", salida="facturas_cliente.xlsx")

🛠️ Funcionalidades
- Recorre automáticamente todas las subcarpetas.
- Extrae facturas y valores totales.
- Suma valores cuando una factura aparece repetida.
- Exporta resultados a Excel con columnas adicionales: número de formulario y código de barras.
- Permite organizar documentos por año y cliente.
