import pandas as pd
import smtplib
from email.mime.text import MIMEText

# Ruta del archivo Excel
archivo_excel = r"C:\Users\LENOVO\OneDrive\Desktop\CP\2026\CP 2026.xlsx"

# Leer las hojas
df_cp = pd.read_excel(archivo_excel, sheet_name="2025")       # hoja con facturas
df_correos = pd.read_excel(archivo_excel, sheet_name="correos")  # hoja con correos

# Filtrar facturas sin CP
pendientes = df_cp[df_cp["Archivo"].isna() | (df_cp["Archivo"] == "")]

# Unir con correos
pendientes = pendientes.merge(df_correos, left_on="Cliente", right_on="Cliente", how="left")

# Guardar en nueva hoja del mismo archivo
with pd.ExcelWriter(archivo_excel, mode="a", if_sheet_exists="replace") as writer:
    pendientes.to_excel(writer, sheet_name="Pendientes", index=False)

print("✅ Hoja 'Pendientes' creada en el Excel")

# --- Enviar correos automáticamente ---
def enviar_correos(pendientes, servicio="gmail"):
    if servicio == "gmail":
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        user = "holguinyirleza@gmail.com"      # 👉 tu correo Gmail
        password = "123789"          # 👉 tu contraseña o token de aplicación
    elif servicio == "outlook":
        smtp_server = "smtp.office365.com"
        smtp_port = 587
        user = "yisleza.vargas@corrumed.co"    # 👉 tu correo Outlook
        password = "123789"          # 👉 tu contraseña o token de aplicación
    else:
        raise ValueError("Servicio no soportado. Usa 'gmail' o 'outlook'.")

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(user, password)

    for cliente, grupo in pendientes.groupby("Cliente"):
        email = grupo["correo"].iloc[0]
        if pd.isna(email):
            continue

        cuerpo = f"Estimado {cliente},\n\nTiene las siguientes facturas pendientes de CP:\n"
        for _, row in grupo.iterrows():
            cuerpo += f"- Factura {row['Factura']} (Valor: {row['Valor Total']})\n"

        cuerpo += "\nPor favor revisar.\n\nSaludos,\nCorrumed"

        msg = MIMEText(cuerpo)
        msg["Subject"] = f"Facturas pendientes de CP - {cliente}"
        msg["From"] = user
        msg["To"] = email

        server.sendmail(user, [email], msg.as_string())
        print(f"📧 Correo enviado a {cliente} ({email})")

    server.quit()

# Ejecutar envío (elige servicio: "gmail" o "outlook")
enviar_correos(pendientes, servicio="gmail")
# enviar_correos(pendientes, servicio="outlook")
