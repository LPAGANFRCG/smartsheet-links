import smartsheet

# Reemplaza con tu token de acceso personal de Smartsheet
SMARTSHEET_ACCESS_TOKEN = "XhRRIgrs2G0Z8mz7Pz7EyJ4AfqsWhehSxxVZX"

# Inicializa el cliente de la API
ss_client = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)

def update_sheet_with_formulas(sheet_id, start_row, end_row, case_id_column_name="Project Name"):
    """
    Actualiza una hoja de Smartsheet con fórmulas para un rango de filas específico
    en múltiples columnas.

    Args:
        sheet_id (int): El ID de la hoja de Smartsheet a actualizar.
        start_row (int): El número de la fila inicial (visible en Smartsheet, no basado en 0).
        end_row (int): El número de la fila final (visible en Smartsheet).
        case_id_column_name (str): El nombre de la columna que contiene los Case IDs.
    """
    
    formulas_by_column = {
        "Subcontractor Demolition": '=SUMIFS({%s - 03. SOW Current - Sub 100%}, {%s - 03. SOW Current - Coverage}, "CE-DEMO")',
        "Subcontractor Site": '=SUMIFS({%s - 03. SOW Current - Sub 100%}, {%s - 03. SOW Current - Coverage}, "CE-SITE") + SUMIFS({%s - 03. SOW Current - Sub 100%}, {%s - 03. SOW Current - Coverage}, "CE-ELEV")',
        "Subcontractor Scope Reduction": '=SUMIFS({%s - 03. SOW Current - Sub 100%}, {%s - 03. SOW Current - Coverage}, "SR")',
        "Subcontractor Cistern/Panels": '=SUMIFS({%s - 03. SOW Current - Sub 100%}, {%s - 03. SOW Current - Coverage}, "CE-WSS") + SUMIFS({%s - 03. SOW Current - Sub 100%}, {%s - 03. SOW Current - Coverage}, "CE-PVS")'
    }

    try:
        sheet = ss_client.Sheets.get_sheet(sheet_id)
        
        column_map = {col.title: col.id for col in sheet.columns}

        missing_columns = [name for name in formulas_by_column.keys() if name not in column_map]
        if missing_columns:
            print(f"❌ Error: Las siguientes columnas no se encontraron en la hoja: {', '.join(missing_columns)}")
            return
            
        case_id_col_id = column_map.get(case_id_column_name)
        if not case_id_col_id:
            print(f"❌ Error: No se encontró la columna de Case IDs '{case_id_column_name}'.")
            return

        rows_to_update = []
        
        start_index = start_row - 1
        end_index = end_row - 1

        for i in range(start_index, end_index + 1):
            if i >= len(sheet.rows):
                print(f"⚠️ Advertencia: El índice de fila {i+1} está fuera de los límites de la hoja. Se detuvo la actualización.")
                break
            
            row = sheet.rows[i]
            row_id = row.id
            
            case_id = None
            for cell in row.cells:
                if cell.column_id == case_id_col_id:
                    case_id = cell.value
                    break
            
            if not case_id:
                print(f"⚠️ Advertencia: No se encontró un Case ID para la fila {i+1}. Se saltará esta fila.")
                continue
            
            cells_for_row = []
            for col_name, formula_template in formulas_by_column.items():
                final_formula = formula_template.replace("%s", case_id)
                
                cell = smartsheet.models.Cell({
                    "column_id": column_map[col_name],
                    "formula": final_formula
                })
                cells_for_row.append(cell)
            
            row_to_update = smartsheet.models.Row({
                "id": row_id,
                "cells": cells_for_row
            })
            rows_to_update.append(row_to_update)
            
        if rows_to_update:
            response = ss_client.Sheets.update_rows(sheet_id, rows_to_update)
            print(f"✅ Se han actualizado {len(response.data)} filas con fórmulas en la hoja con ID {sheet_id}.")
        else:
            print("❌ No se encontraron filas para actualizar.")
            
    except smartsheet.exceptions.ApiError as e:
        print(f"❌ Error al interactuar con la API de Smartsheet: {e}")
        print(e.error.result_code)
        print(e.error.message)
    except Exception as e:
        print(f"❌ Ha ocurrido un error inesperado: {e}")

if __name__ == "__main__":
    target_sheet_id = 6190865895608196
    start_row = 105
    end_row = 154

    update_sheet_with_formulas(target_sheet_id, start_row, end_row)