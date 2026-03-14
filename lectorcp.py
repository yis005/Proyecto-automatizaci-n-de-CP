import os
import re
import pdfplumber
import pandas as pd

def extraer_datos_pdf(pdf_path, debug=False):
    """
    Extrae del PDF:
        - Número de factura (patrón CM- seguido de dígitos)
        - Lista de valores del campo 52 (valor total de cada ítem)
        - Número de formulario (prioriza 0006..., luego 14 dígitos)
    """
    factura = None
    valores_52 = []
    formulario = None
    texto_completo = ""

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            texto = page.extract_text()
            if texto:
                texto_completo += texto + "\n"

            # Extraer palabras especificas
            words = page.extract_words()

            # Buscar casillas "52." en esta página
            etiquetas_52 = [w for w in words if "52." in w["text"]]

            # buscar el número de abajo de cada etiqueta
            for etq in etiquetas_52:
                for cand in words:
                    # Misma columna debajo de la etiqueta
                    if (cand["top"] > etq["top"] and
                        abs(cand["x0"] - etq["x0"]) < 15 and
                        re.match(r'^[\d.]+$', cand["text"])):
                        # no suma el 52.
                        if cand["text"] != "52.":
                            valores_52.append(cand["text"])
                        break  

            # Buscar el numero de codigo de barras (aún no me funciona)
            for w in words:
                if re.match(r'^0006\d+', w["text"]):
                    formulario = w["text"]
                    break
                elif re.match(r'^\d{14}$', w["text"]) and not formulario:
                    formulario = w["text"]

        # Si no se encontraron valores por palabras especificas, intentar con regex
        if not valores_52:
            # Patrón más flexible: "52. Valor total" seguido de un número con decimales
            patron = r'52\.\s*Valor\s*total.*?(\d+\.\d+)'
            valores_52 = re.findall(patron, texto_completo, re.IGNORECASE | re.DOTALL)

        # Buscar la palabra factura en todo el PDF
        match_factura = re.search(r'CM-\d+', texto_completo)
        if match_factura:
            factura = match_factura.group(0)

    if not factura:
        factura = os.path.splitext(os.path.basename(pdf_path))[0]

    # Depurar
    if debug:
        print("\n--- TEXTO EXTRAÍDO (primeros 2000 caracteres) ---")
        print(texto_completo[:2000])
        print("----------------------")

    return factura, valores_52, formulario

def procesar_pdfs(base_dir="C:/Users/LENOVO/OneDrive/Desktop/CP",
                  salida="C:/Users/LENOVO/OneDrive/Desktop/CP/resultado.xlsx",
                  debug=True):
    resultados = []

    for root, dirs, files in os.walk(base_dir):
        for file in files:
            if file.lower().endswith(".pdf"):
                pdf_path = os.path.join(root, file)
                try:
                    factura, valores, formulario = extraer_datos_pdf(pdf_path, debug=debug)

                    # Sumar los valores numéricos, menos el 52.
                    suma_total = 0.0
                    valores_limpios = []
                    for v in valores:
                        # Eliminar cualquier texto que no sea número (52.)
                        if re.match(r'^\d+\.\d+$', v):  # Asegura formato decimal
                            try:
                                suma_total += float(v)
                                valores_limpios.append(v)
                            except ValueError:
                                pass

                    resultados.append({
                        "Archivo": file,
                        "Factura": factura,
                        "Valor Total Sumado": suma_total,
                        "Número Formulario": formulario if formulario else "",
                        "Valores Encontrados": ", ".join(valores_limpios)  # Solo valores válidos
                    })

                    print(f"✅ {file}: Factura={factura}, Suma={suma_total}, Formulario={formulario}")
                    print(f"   Valores encontrados: {valores_limpios}")

                except Exception as e:
                    print(f"❌ Error en {file}: {e}")

    # Crear Datos y guardar
    df = pd.DataFrame(resultados)
    # quitar la columna de resultados en el excel, no es necesaria, solo la tengo para ver los resultados en el script
    df.drop(columns=["Valores Encontrados"], inplace=True)
    df.to_excel(salida, index=False)
    print(f"\n📁 Archivo Excel generado: {salida}")

if __name__ == "__main__":
    procesar_pdfs()
