import requests
import pandas as pd

# Configuración
SMARTSHEET_API_TOKEN = "XhRRIgrs2G0Z8mz7Pz7EyJ4AfqsWhehSxxVZX"
SHEET_ID = "3252174586859396"  # El que usa el Data Shuttle como destino
EXCEL_PATH = "SFH_ScopeChanges_SS_DS.csv"  # O .xlsx

# 1. Leer columnas desde Excel
df = pd.read_csv(EXCEL_PATH)  # usa read_excel() si es .xlsx
excel_columns = list(df.columns)

# 2. Leer columnas desde Smartsheet
headers = {
    "Authorization": f"Bearer {SMARTSHEET_API_TOKEN}"
}
url = f"https://api.smartsheet.com/2.0/sheets/{SHEET_ID}"
response = requests.get(url, headers=headers)
sheet_data = response.json()

smartsheet_columns = [col["title"] for col in sheet_data["columns"]]

# 3. Comparar
excel_normalized = [col.strip().lower() for col in excel_columns]
smartsheet_normalized = [col.strip().lower() for col in smartsheet_columns]

faltan_en_smartsheet = [col for col in excel_columns if col.lower() not in smartsheet_normalized]
ya_mapeadas = [col for col in excel_columns if col.lower() in smartsheet_normalized]
sobrantes_en_smartsheet = [col for col in smartsheet_columns if col.lower() not in excel_normalized]

# 4. Resultados
print("📌 Columnas que faltan en Smartsheet:")
for col in faltan_en_smartsheet:
    print("-", col)

print("\n✅ Columnas que ya están mapeadas:")
for col in ya_mapeadas:
    print("-", col)

print("\n⚠️ Columnas en Smartsheet que no están en Excel:")
for col in sobrantes_en_smartsheet:
    print("-", col)
