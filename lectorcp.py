import os
import re
import pdfplumber
import pandas as pd
from collections import defaultdict

def extraer_datos_pdf(pdf_path, debug=False):
    """
    Extrae del PDF:
        - Número de CP (formulario)
        - Lista de facturas (CM-xxxxx, 190-xxxxx)
        - Valores del campo 52 (valor total)
    """
    facturas = []
    valores_52 = []
    formulario = None
    texto_completo = ""

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            texto = page.extract_text() or ""
            texto_completo += texto + "\n"
            words = page.extract_words()

            # Buscar el total de cada fv
            etiquetas_52 = [w for w in words if "52." in w["text"]]
            for etq in etiquetas_52:
                for cand in words:
                    if (cand["top"] > etq["top"] and
                        abs(cand["x0"] - etq["x0"]) < 15 and
                        re.match(r'^[\d.]+$', cand["text"])):
                        if cand["text"] != "52.":
                            valores_52.append(cand["text"])
                        break

            # númeor CP (me está trayendo la resolución)
            for w in words:
                if re.match(r'^0006\d+', w["text"]):
                    formulario = w["text"]
                    break
                elif re.match(r'^\d{14}$', w["text"]) and not formulario:
                    formulario = w["text"]

        # Si no encontró valores con palabras, usar regex
        if not valores_52:
            patron = r'52\.\s*Valor\s*total.*?(\d+\.\d+)'
            valores_52 = re.findall(patron, texto_completo, re.IGNORECASE | re.DOTALL)

        # Buscar solo facturas válidas (CM-xxxxx o 190-xxxxx)
        facturas = re.findall(r'(?:CM-\d+|190-\d+)', texto_completo)

    # Convertir valores a float, para que no se quiebbbre
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
        print(f"   Facturas: {facturas}")
        print(f"   Valores: {valores_limpios}")

    return facturas, valores_limpios, formulario

def procesar_pdfs(base_dir, salida, debug=True):
    filas = []

    for root, dirs, files in os.walk(base_dir):
        for file in files:
            if file.lower().endswith(".pdf"):
                pdf_path = os.path.join(root, file)
                try:
                    facturas, valores, formulario = extraer_datos_pdf(pdf_path, debug=debug)

                    # Agrupar valores por factura
                    agrupados = defaultdict(float)
                    for i in range(min(len(facturas), len(valores))):
                        agrupados[facturas[i]] += valores[i]

                    # Crear filas únicas por factura
                    for factura, total in agrupados.items():
                        filas.append({
                            "Archivo": file,
                            "Número Formulario (CP)": formulario if formulario else "",
                            "Factura": factura,
                            "Valor Total": total
                        })

                except Exception as e:
                    print(f"x Error en {file}: {e}")

    df = pd.DataFrame(filas)
    df.to_excel(salida, index=False)
    print(f"\n Archivo Excel generado: {salida}")

if __name__ == "__main__":
    procesar_pdfs(
        base_dir="C:/Users/LENOVO/OneDrive/Desktop/CP",
        salida="C:/Users/LENOVO/OneDrive/Desktop/CP/resultado.xlsx",
        debug=True
    )