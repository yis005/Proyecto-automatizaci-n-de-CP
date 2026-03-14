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

            # Extraer palabras con coordenadas
            words = page.extract_words()

            # Buscar etiquetas "52." en esta página
            etiquetas_52 = [w for w in words if "52." in w["text"]]

            # Para cada etiqueta, buscar el número asociado (misma columna, debajo)
            for etq in etiquetas_52:
                for cand in words:
                    # Misma columna con tolerancia, debajo de la etiqueta
                    if (cand["top"] > etq["top"] and
                        abs(cand["x0"] - etq["x0"]) < 15 and
                        re.match(r'^[\d.]+$', cand["text"])):
                        # Evitar agregar la propia etiqueta "52." como valor
                        if cand["text"] != "52.":
                            valores_52.append(cand["text"])
                        break  # Solo el primer número después de la etiqueta

            # Buscar número de formulario (priorizar 0006...)
            for w in words:
                if re.match(r'^0006\d+', w["text"]):
                    formulario = w["text"]
                    break
                elif re.match(r'^\d{14}$', w["text"]) and not formulario:
                    formulario = w["text"]

        # Si no se encontraron valores por coordenadas, intentar con regex
        if not valores_52:
            # Patrón más flexible: "52. Valor total" seguido de un número con decimales
            patron = r'52\.\s*Valor\s*total.*?(\d+\.\d+)'
            valores_52 = re.findall(patron, texto_completo, re.IGNORECASE | re.DOTALL)

        # Buscar factura en todo el texto
        match_factura = re.search(r'CM-\d+', texto_completo)
        if match_factura:
            factura = match_factura.group(0)

    if not factura:
        factura = os.path.splitext(os.path.basename(pdf_path))[0]

    # Depuración opcional
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

                    # Sumar los valores numéricos, ignorando los que no son válidos
                    suma_total = 0.0
                    valores_limpios = []
                    for v in valores:
                        # Eliminar cualquier texto que no sea número (como "52.")
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

    # Crear DataFrame y guardar
    df = pd.DataFrame(resultados)
    # Opcional: eliminar columna de depuración si no la quieres en el Excel
    # df.drop(columns=["Valores Encontrados"], inplace=True)
    df.to_excel(salida, index=False)
    print(f"\n📁 Archivo Excel generado: {salida}")

if __name__ == "__main__":
    procesar_pdfs()