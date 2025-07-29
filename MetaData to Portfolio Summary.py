import smartsheet
from smartsheet.models import CellLink, Cell, Row

API_TOKEN = 'XhRRIgrs2G0Z8mz7Pz7EyJ4AfqsWhehSxxVZX'  # 🔐 Coloca tu token
SOURCE_SHEET_ID = 4013098424815492
TARGET_SHEET_ID = 6381413683122052

SOURCE_START_ROW = 170  # índice base 0, así que fila 171 es index 170
SOURCE_END_ROW = 187    # inclusive

TARGET_ROW_NUMBER = 64  # fila destino en el target sheet

# Orden de las columnas destino
target_column_names = [
    "Subcontractor Tier 2 - #1",
    "Original Subcontract Amount Sub 1",
    "Total Change Orders Sub 1",
    "Total Subcontract Amount (w-Change Orders) Sub 1",
    "Total Invoiced Sub 1",
    "Total Paid Sub 1",
    "Subcontractor Tier 2 - #2",
    "Original Subcontract Amount Sub 2",
    "Total Change Orders Sub 2",
    "Total Subcontract Amount (w-Change Orders) Sub 2",
    "Total Invoiced Sub 2",
    "Total Paid Sub 2",
    "Subcontractor Tier 2 - #3",
    "Original Subcontract Amount Sub 3",
    "Total Change Orders Sub 3",
    "Total Subcontract Amount (w-Change Orders) Sub 3",
    "Total Invoiced Sub 3",
    "Total Paid Sub 3"
]

# Inicializa cliente
smartsheet_client = smartsheet.Smartsheet(API_TOKEN)

# Valida conexión
account_info = smartsheet_client.Users.get_current_user()
print("🔑 Conectado como:", account_info.email)

# Carga las hojas
source_sheet = smartsheet_client.Sheets.get_sheet(SOURCE_SHEET_ID)
target_sheet = smartsheet_client.Sheets.get_sheet(TARGET_SHEET_ID)

# Obtiene columna del source llamada "Text/Number"
source_column = next(col for col in source_sheet.columns if col.title == "Text/Number")

# Fila destino
target_row = next(row for row in target_sheet.rows if row.row_number == TARGET_ROW_NUMBER)
target_row_id = target_row.id

# Obtiene columnas destino en orden
target_columns = [col for name in target_column_names
                  for col in target_sheet.columns if col.title == name]

# 🔍 DEBUG: verifica si alguna columna no se encontró
print("\n🔎 Columnas encontradas en target:")
for col in target_columns:
    print("-", col.title)

faltantes = [name for name in target_column_names if name not in [col.title for col in target_columns]]
if faltantes:
    print("\n❌ No se encontraron estas columnas:")
    for f in faltantes:
        print("-", f)

# Verificación de longitud
expected_links = SOURCE_END_ROW - SOURCE_START_ROW + 1
if expected_links != len(target_columns):
    raise ValueError(f"\n🚫 El número de celdas verticales ({expected_links}) no coincide con columnas destino ({len(target_columns)}).")

# Construye los CellLinks
linked_cells = []
for i in range(len(target_columns)):
    source_row = source_sheet.rows[SOURCE_START_ROW + i]

    cell_link = CellLink()
    cell_link.sheet_id = SOURCE_SHEET_ID
    cell_link.row_id = source_row.id
    cell_link.column_id = source_column.id

    cell = Cell()
    cell.column_id = target_columns[i].id
    cell.value = ""  # ⚠️ necesario aunque se use link_in_from_cell
    cell.link_in_from_cell = cell_link
    linked_cells.append(cell)

# Actualiza la fila en el target
updated_row = Row()
updated_row.id = target_row_id
updated_row.cells = linked_cells

response = smartsheet_client.Sheets.update_rows(TARGET_SHEET_ID, [updated_row])
print("\n✅ Link creado con éxito. Status:", response.message)
