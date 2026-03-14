import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configuración del servidor con gmail
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_USER = "tu_correo@gmail.com"
EMAIL_PASS = "232323"  #tokrn de envio

def enviar_correo(destinatario, cliente, factura):
    msg = MIMEMultipart()
    msg["From"] = EMAIL_USER
    msg["To"] = destinatario
    msg["Subject"] = f"Documentos faltantes - Factura {factura}"

    cuerpo = f"""
    Estimado/a {cliente},

    Hemos detectado que el Certificado al Proveedor (CP) correspondiente a la factura {factura}
    aún no ha sido recibido.

    Por favor, envíe el documento a la mayor brevedad.

    Saludos,
    Yisleza Vargas
    Analista de facturación
    Corrumed SAS
    """
    msg.attach(MIMEText(cuerpo, "plain"))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASS)
        server.sendmail(EMAIL_USER, destinatario, msg.as_string())
        print(f"✅ Correo enviado a {destinatario} por factura {factura}")

def procesar_excel_y_enviar(ruta_excel="C:/Users/LENOVO/OneDrive/Desktop/CP/resultado.xlsx"):
    # Leer el Excel con columnas: Factura, Cliente, Correo, CP
    df = pd.read_excel(ruta_excel)

    # Filtrar pendientes (CP vacío)
    pendientes = df[df["CP"].isna() | (df["CP"].astype(str).str.strip() == "")]

    # Guardar en nueva hoja "Pendientes"
    with pd.ExcelWriter(ruta_excel, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
        pendientes.to_excel(writer, sheet_name="Pendientes", index=False)

    print(f"\n✅ Hoja 'Pendientes' creada en {ruta_excel}")

    # Enviar correos a los pendientes
    for _, row in pendientes.iterrows():
        enviar_correo(row["Correo"], row["Cliente"], row["Factura"])

if __name__ == "__main__":
    procesar_excel_y_enviar()