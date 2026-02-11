import os
import pdfplumber
import pandas as pd
import re
from collections import defaultdict

#Ojo aquí está la ruta de las carpetas a recorrer
def procesar_pdfs(base_dir="C:\Users\LENOVO\OneDrive\Desktop\CP", salida="resultado.xlsx"):
    resultados = defaultdict(lambda: {"total": 0.0, "formulario": None, "codigo_barras": None})

#Recorrer carpetas y subcarpetas 
    for root, dirs, files in os.walk(base_dir):
        for file in files:
            if file.endswith(".pdf"): #ojo con esto nos aseguramos que solo recorra archivos .PDF
                pdf_path = os.path.join(root, file) #Ojo esto es para obtener la ruta completa de los archivos
                with pdfplumber.open(pdf_path) as pdf:
                    text = ""
                    for page in pdf.pages:
                        text += page.extract_text() or ""
                    
                    # Extraer número de formulario
                    formulario = re.search(r"Número de formulario.*?(\d+)", text)
                    # Extraer código de barras (secuencia larga de dígitos)
                    codigo_barras = re.search(r"\b\d{15,}\b", text)

                    facturas = re.findall(r"40\. No\. Factura\s*\n?[A-Za-z\-]*([0-9]+)", text)
                    valores = re.findall(r"52\. Valor total\s*\n?([\d,.]+)", text) # Buscar facturas y valores
                   

                    for factura, valor in zip(facturas, valores):
                        valor_num = float(valor.replace(",", "").strip())
                        resultados[factura]["total"] += valor_num
                        resultados[factura]["formulario"] = formulario.group(1) if formulario else None
                        resultados[factura]["codigo_barras"] = codigo_barras.group(0) if codigo_barras else None

    # Convertir resultados a DataFrame
    data = [
        {
            "Factura": factura,
            "Valor Total": info["total"],
            "Número Formulario": info["formulario"],
            "Código Barras": info["codigo_barras"]
        }
        for factura, info in resultados.items()
    ]
    df = pd.DataFrame(data)

    # Exportar a Excel
    df.to_excel(salida, index=False)
    print(f"✅ Archivo Excel generado: {salida}")

# Ejecutar
if __name__ == "__main__":
    procesar_pdfs()