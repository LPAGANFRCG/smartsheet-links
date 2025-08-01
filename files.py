import smartsheet

# Autenticación
API_TOKEN = 'XhRRIgrs2G0Z8mz7Pz7EyJ4AfqsWhehSxxVZX'
smartsheet_client = smartsheet.Smartsheet(API_TOKEN)

# Parámetros
project_number = "PR-R3-23922"
sheet_folder_id = 3878298838165380

# Sheets fuente (con contenido ya dentro)
sheet_sources = {
    "11. MIT Sub Pmt Breakdown": 8748458094579588,
    "12. Material Take-Off": 7042029629427588,
}

# Clonar los sheets desde otros sheets existentes (NO templates)
for name_suffix, source_sheet_id in sheet_sources.items():
    full_name = f"{project_number} - {name_suffix}"
    print(f"🔄 Clonando sheet '{full_name}'...")
    try:
        copied_sheet = smartsheet_client.Sheets.copy_sheet(
            source_sheet_id,
            smartsheet.models.ContainerDestination({
                'destination_type': 'folder',
                'destination_id': sheet_folder_id,
                'new_name': full_name
            })
        ).data
        print(f"✅ Sheet copiado: {copied_sheet.name}")
    except Exception as e:
        print(f"❌ Error copiando sheet '{full_name}': {e}")
