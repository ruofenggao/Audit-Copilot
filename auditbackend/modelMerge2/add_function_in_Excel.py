import xlrd
from xlutils.copy import copy
from xlwt import Formula, XFStyle, Borders, Font

def create_style(xf_obj, rb):
    style = XFStyle()

    # 字体样式
    font_obj = rb.font_list[xf_obj.font_index]
    font_style = Font()
    font_style.name = font_obj.name
    font_style.bold = font_obj.bold
    font_style.italic = font_obj.italic
    font_style.height = font_obj.height
    style.font = font_style

    # 边框样式
    borders_obj = xf_obj.border
    borders_style = Borders()
    borders_style.left = borders_obj.left_line_style
    borders_style.right = borders_obj.right_line_style
    borders_style.top = borders_obj.top_line_style
    borders_style.bottom = borders_obj.bottom_line_style
    style.borders = borders_style

    return style

def write_formulas(sheet, ws, start_row, end_row, column_number, rb):
    for row_index in range(start_row, end_row - 1):
        xf_index = sheet.cell_xf_index(row_index, column_number)
        xf_obj = rb.xf_list[xf_index]

        style = create_style(xf_obj, rb)

        formula = f'E{row_index + 1}*F{row_index + 1}'
        ws.write(row_index, column_number, Formula(formula), style)
        print(f"Processed row {row_index + 1}")

def process_workbook(file_path, sheet_name):
    try:
        rb = xlrd.open_workbook(file_path, formatting_info=True)
        sheet = rb.sheet_by_name(sheet_name)

        wb = copy(rb)
        ws_index = rb.sheet_names().index(sheet_name)
        ws = wb.get_sheet(ws_index)

        write_formulas(sheet, ws, 3, sheet.nrows, 7, rb)

        wb.save(file_path)
        print("Workbook processed and saved successfully.")
    except Exception as e:
        print(f"Error: {e}")

# 使用函数
original_file_path = '/Users/ruofenggao/Desktop/副本2023年西安电信5G五期配套承载网STN接入层扩容工程（中通服-3）_23SN001515007_对照.xls'
sheet_name = '对照表三甲'
process_workbook(original_file_path, sheet_name)

