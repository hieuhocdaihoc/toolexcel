from flask import Flask, request, send_file
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
import io

app = Flask(__name__)

@app.route('/api/format', methods=['POST'])
def format_excel():
    if 'file' not in request.files:
        return "No file uploaded", 400
        
    file = request.files['file']
    
    # Load workbook from memory (data_only=True để bỏ công thức ẩn)
    wb = openpyxl.load_workbook(file, data_only=True)
    
    header_fill = PatternFill(start_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
                         
    fill_red = PatternFill(start_color="FF0000", fill_type="solid")
    fill_green = PatternFill(start_color="00B050", fill_type="solid")
    fill_orange = PatternFill(start_color="FF9900", fill_type="solid")
    fill_yellow = PatternFill(start_color="FFFF00", fill_type="solid")
    
    white_font = Font(color="FFFFFF", bold=True)
    black_font = Font(color="000000", bold=True)
    
    for ws in wb.worksheets:
        max_cols = ws.max_column
        max_rows = ws.max_row
        status_col_idx = None
        
        for row in range(1, max_rows + 1):
            for col in range(1, max_cols + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = thin_border
                cell.alignment = Alignment(vertical='center')
                
                if row == 1:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    if cell.value and "TRẠNG THÁI" in str(cell.value).upper():
                        status_col_idx = col

        if status_col_idx:
            for row in range(2, max_rows + 1):
                cell = ws.cell(row=row, column=status_col_idx)
                val = str(cell.value or "").strip().upper()
                
                if "HẸN LỊCH" in val:
                    cell.fill = fill_red
                    cell.font = white_font
                elif "CHỐT" in val:
                    cell.fill = fill_green
                    cell.font = white_font
                elif "KH BONG" in val or "HỦY LỊCH" in val:
                    cell.fill = fill_orange
                    cell.font = black_font
                elif "TIỀM NĂNG" in val:
                    cell.fill = fill_yellow
                    cell.font = black_font
                    
        if max_cols > 0 and max_rows > 0:
            last_col_letter = openpyxl.utils.get_column_letter(max_cols)
            ws.auto_filter.ref = f"A1:{last_col_letter}{max_rows}"
        
        for col in ws.columns:
            max_length = 0
            column_letter = col[0].column_letter
            for cell in col:
                try:
                    if cell.value is not None:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[column_letter].width = min(max_length + 2, 45)
            
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    original_name = file.filename
    if original_name.endswith('.xlsx'):
        out_name = original_name.replace('.xlsx', '_Formatted.xlsx')
    else:
        out_name = "KetQua_Formatted.xlsx"
        
    return send_file(output, download_name=out_name, as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# Handler cho Vercel
def handler(request, response):
    pass
