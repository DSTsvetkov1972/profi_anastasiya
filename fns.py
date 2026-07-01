from openpyxl import load_workbook
from openpyxl.styles import Alignment

def format_res_file(res_file):
    # Загружаем файл
    wb = load_workbook(res_file)
    ws = wb.active

    for row in range(1, ws.max_row + 1):
        if 'Пробная' in str(ws.cell(column=1, row=row).value):
            print(row)
            # Объединяем ячейки от (row, start_col) до (row, end_col)
            ws.merge_cells(
                start_row=row, start_column=1,
                end_row=row, end_column=ws.max_column)

            ws.merge_cells(
                start_row=row+1, start_column=1,
                end_row=row+3, end_column=1)
            
            ws.merge_cells(
                start_row=row+1, start_column=2,
                end_row=row+3, end_column=2)
            
                        
            ws.merge_cells(
                start_row=row+1, start_column=4,
                end_row=row+3, end_column=4)
            
                        
            ws.merge_cells(
                start_row=row+1, start_column=14,
                end_row=row+3, end_column=14)
            

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    wb.save(res_file)