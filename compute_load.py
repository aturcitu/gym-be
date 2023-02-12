
from datetime import datetime
from gymdef import load_worksheet, log_info
from gspread.cell import Cell

#--- LOAD EXCEL INFO ---#
wks, sh = load_worksheet()
if wks:
    wks_dic = wks.get_all_records()
    # Get max load for each Exercise ID
    exer_records = {}
    for wks_row in wks_dic:
        row_id = wks_row['id']
        row_load = wks_row['carga']
        try:
            if row_load > exer_records[row_id]:
                exer_records[row_id] = row_load
        except KeyError:
            exer_records[row_id] = row_load
    # Compute relative load (normlizing to max -> 100)
    exer_rel_load = []
    for wks_row in wks_dic:
        row_id = wks_row['id']  
        row_load = wks_row['carga']
        try:
            rel_load = int((row_load/exer_records[row_id])*100)
        except:
            rel_load = 0
        exer_rel_load.append(rel_load)
    # Update values
    log_info("All relative load have been computed", tabs=1)
    cells = []
    wks_col = [item for item in wks.row_values(1) if item].index('cargaRelativa')+1
    for wks_row, cell_data in enumerate(exer_rel_load):
        cells.append(Cell(row=wks_row + 2, col=wks_col, value = cell_data)) 
        
    wks.update_cells(cells, value_input_option='USER_ENTERED')
    log_info("Worksheed updated successfully", tabs=1)

else:
    log_info("Fail", tabs=1)
