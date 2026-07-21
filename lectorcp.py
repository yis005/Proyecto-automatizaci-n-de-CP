import os
import re
import pdfplumber
import pandas as pd
from collections import defaultdict
import shutil

def extraer_datos_pdf(pdf_path, debug=False):
    facturas = []
    valores_52 = []
    formulario = None
    razon_social = None
    texto_completo = ""

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            texto = page.extract_text() or ""
            texto_completo += texto + "\n"
            words = page.extract_words()

            # Buscar razón social (campo 11)
            patron_razon = r'11\.\s*Raz[oó]n\s*social\s*(?:\d+\s*){0,3}([A-Z0-9\s\.\-&]+)'
            coincidencia = re.search(patron_razon, texto)
            if coincidencia:
                razon_social = coincidencia.group(1).strip()
                # Limpiar caracteres extraños y espacios
                razon_social = re.sub(r'\s+', '_', razon_social)
                razon_social = re.sub(r'[^A-Za-z0-9_\.\-&]', '', razon_social)

            # Buscar número de formulario
            for w in words:
                if re.match(r'^0006\d+', w["text"]):
                    formulario = w["text"]
                    break
                elif re.match(r'^\d{14}$', w["text"]) and not formulario:
                    formulario = w["text"]

            # Buscar valores del campo 52
            etiquetas_52 = [w for w in words if "52." in w["text"]]
            for etq in etiquetas_52:
                for cand in words:
                    if (cand["top"] > etq["top"] and
                        abs(cand["x0"] - etq["x0"]) < 15 and
                        re.match(r'^[\d.]+$', cand["text"])):
                        if cand["text"] != "52.":
                            valores_52.append(cand["text"])
                        break

        # Si no encontró valores con palabras, usar regex
        if not valores_52:
            patron = r'52\.\s*Valor\s*total.*?(\d+\.\d+)'
            valores_52 = re.findall(patron, texto_completo, re.IGNORECASE | re.DOTALL)

        # Buscar facturas válidas
        facturas = re.findall(r'(?:CM-\d+|190-\d+)', texto_completo)

    # Convertir valores a float
    valores_limpios = []
    for v in valores_52:
        if re.match(r'^\d+\.\d+$', v):
            try:
                valores_limpios.append(float(v))
            except ValueError:
                pass

    if debug:
        print(f"📄 {os.path.basename(pdf_path)}")
        print(f"   CP: {formulario}")
        print(f"   Cliente (Razón social): {razon_social}")
        print(f"   Facturas: {facturas}")
        print(f"   Valores: {valores_limpios}")

    return facturas, valores_limpios, formulario, razon_social


def procesar_pdfs(base_dir, salida, carpeta_procesados, debug=True):
    filas = []
    os.makedirs(carpeta_procesados, exist_ok=True)

    for root, dirs, files in os.walk(base_dir):
        for file in files:
            if file.lower().endswith(".pdf"):
                pdf_path = os.path.join(root, file)
                try:
                    facturas, valores, formulario, razon_social = extraer_datos_pdf(pdf_path, debug=debug)

                    agrupados = defaultdict(float)
                    for i in range(min(len(facturas), len(valores))):
                        agrupados[facturas[i]] += valores[i]

                    for factura, total in agrupados.items():
                        filas.append({
                            "Archivo": file,
                            "Número Formulario (CP)": formulario if formulario else "",
                            "Factura": factura,
                            "Cliente (Razón social)": razon_social if razon_social else "",
                            "Valor Total": total
                        })

                    # Crear carpeta por cliente (solo nombre limpio)
                    carpeta_cliente = razon_social if razon_social else "sin_cliente"
                    destino_cliente = os.path.join(carpeta_procesados, carpeta_cliente)
                    os.makedirs(destino_cliente, exist_ok=True)

                    destino = os.path.join(destino_cliente, file)
                    shutil.move(pdf_path, destino)
                    if debug:
                        print(f"   ✅ Movido a: {destino}")

                except Exception as e:
                    print(f"x Error en {file}: {e}")

    df = pd.DataFrame(filas)
    df.to_excel(salida, index=False)
    print(f"\n Archivo Excel generado: {salida}")


if __name__ == "__main__":
    procesar_pdfs(
        base_dir="C:/Users/LENOVO/OneDrive/Desktop/CP",
        salida="C:/Users/LENOVO/OneDrive/Desktop/CP/resultado.xlsx",
        carpeta_procesados="C:/Users/LENOVO/OneDrive/Desktop/CP/procesados",
        debug=True
    )
