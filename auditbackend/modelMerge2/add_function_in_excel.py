import tkinter as tk
from tkinter import filedialog
import os
import xlrd
from xlutils.copy import copy
from xlwt import XFStyle, Font, Borders, Formula

# 根据单元格格式对象和读取的工作簿对象创建样式
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

# 在指定工作表中写入公式
def write_formulas(sheet, ws, formulas, rb):
    for row_index in range(sheet.nrows):
        for col_index, formula in formulas.items():
            xf_index = sheet.cell_xf_index(row_index, col_index)
            xf_obj = rb.xf_list[xf_index]
            style = create_style(xf_obj, rb)
            ws.write(row_index, col_index, Formula(formula.format(row_index + 1)), style)

# 处理单个工作簿的基类
class SheetProcessor:
    def __init__(self, sheet_name):
        self.sheet_name = sheet_name
        self.formulas = {}  # 默认没有特殊公式

    def process_sheet(self, rb, wb):
        sheet = rb.sheet_by_name(self.sheet_name)
        ws_index = rb.sheet_names().index(self.sheet_name)
        ws = wb.get_sheet(ws_index)
        write_formulas(sheet, ws, self.formulas, rb)

# 对照表一的处理类
class SheetProcessorFor对照表一(SheetProcessor):
    def __init__(self):
        super().__init__('对照表一')
        self.formulas_last_row = {8: 'D{}*E{}', 19: 'L{}*O{}'}  # 特殊公式：第8列和第19列最后一行的计算
    def calculate_last_row(self, sheet, ws, formulas, rb):
        # 计算最后一行的特殊公式
        for col_index, formula in formulas.items():
            xf_index = sheet.cell_xf_index(sheet.nrows - 1, col_index)
            xf_obj = rb.xf_list[xf_index]
            style = create_style(xf_obj, rb)
            ws.write(sheet.nrows - 1, col_index, Formula(formula.format(sheet.nrows)), style)

# 对照表二的处理类
class SheetProcessorFor对照表二(SheetProcessor):
    def __init__(self):
        super().__init__('对照表二')
        # 添加对照表二特殊公式...

# 对照表三甲的处理类
class SheetProcessorFor对照表三甲(SheetProcessor):
    def __init__(self):
        super().__init__('对照表三甲')
        self.formulas = {7: 'E{}*F{}', 18: 'L{}*O{}'}  # 覆盖默认公式
        self.formulas_last_row = {8: 'D{}*E{}', 19: 'L{}*O{}'}  # 特殊公式：第8列和第19列最后一行的计算# 计算特殊公式

    def calculate_last_row(self, sheet, ws, formulas, rb):
        # 计算最后一行的特殊公式
        for col_index, formula in formulas.items():
            xf_index = sheet.cell_xf_index(sheet.nrows - 1, col_index)
            xf_obj = rb.xf_list[xf_index]
            style = create_style(xf_obj, rb)
            ws.write(sheet.nrows - 1, col_index, Formula(formula.format(sheet.nrows)), style)

# GUI界面
class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.folder_selected = None
        self.center_window(400, 300)  # 调整窗口大小
        self.create_widgets()
        self.pack(expand=True)
        self.info_label = tk.Label(self, text="")
        self.info_label.pack(side="top", fill="x", expand=True)

    def center_window(self, width, height):
        # 获取屏幕尺寸以计算布局参数，使窗口居中
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()

        # 计算X和Y坐标
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2

        self.master.geometry(f'{width}x{height}+{x}+{y}')

    def create_widgets(self):
        # 新增的标签，用于显示选定的文件夹路径
        self.folder_path_label = tk.Label(self, text="未选择文件夹")  # 创建 Label 对象
        self.folder_path_label.pack(side="top", fill="x", expand=True)  # 使用 pack 方法

        self.select_folder_button = tk.Button(self)
        self.select_folder_button["text"] = "选择文件夹"
        self.select_folder_button["command"] = self.select_folder
        self.select_folder_button.pack(side="top", fill="x", expand=True, padx=20, pady=10)

        self.start_button = tk.Button(self)
        self.start_button["text"] = "开始处理"
        self.start_button["command"] = self.process_files
        self.start_button.pack(side="top", fill="x", expand=True, padx=20, pady=10)

        self.quit = tk.Button(self, text="退出", fg="red", command=self.master.destroy)
        self.quit.pack(side="bottom", fill="x", expand=True, padx=20, pady=10)

    def select_folder(self):
        self.folder_selected = filedialog.askdirectory()
        if self.folder_selected:
            self.folder_path_label["text"] = self.folder_selected
            self.info_label["text"] = "文件夹已选择"
        else:
            self.folder_path_label["text"] = "未选择文件夹"
            self.info_label["text"] = ""

    # 在 process_files 方法中更新 Label 文本
    def process_files(self):
        if hasattr(self, 'folder_selected') and self.folder_selected:
            for file in os.listdir(self.folder_selected):
                if file.endswith(".xls") or file.endswith(".xlsx"):
                    file_path = os.path.join(self.folder_selected, file)
                    sheet_processors = [
                        SheetProcessorFor对照表一(),
                        SheetProcessorFor对照表二(),
                        SheetProcessorFor对照表三甲()
                        # 添加其他 sheet 处理类...
                    ]
                    try:
                        process_workbook(file_path, sheet_processors)
                        self.info_label["text"] = f"{file} 处理完成"
                    except Exception as e:
                        self.info_label["text"] = f"{file} 处理失败: {e}"
            print("Processing completed.")
        else:
            self.info_label["text"] = "未选择文件夹"

def process_workbook(file_path, sheet_processors):
    try:
        rb = xlrd.open_workbook(file_path, formatting_info=True)
        wb = copy(rb)

        for processor in sheet_processors:
            processor.process_sheet(rb, wb)

        wb.save(file_path)
        print(f"Workbook {file_path} processed and saved successfully.")
    except Exception as e:
        print(f"Error processing {file_path}: {e}")

# 运行程序
root = tk.Tk()
app = Application(master=root)
app.mainloop()
