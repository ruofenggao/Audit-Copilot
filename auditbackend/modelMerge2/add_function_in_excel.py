from PyQt5 import QtWidgets, QtCore
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


# 对照表二的处理类
class SheetProcessorFor对照表二(SheetProcessor):
    def __init__(self):
        super().__init__('对照表二')
        # 添加对照表二特殊公式...

# 对照表三甲的处理类
class SheetProcessorFor对照表三甲(SheetProcessor):
    def __init__(self):
        super().__init__('对照表三甲')

    def process_sheet(self, rb, wb):
        sheet = rb.sheet_by_name(self.sheet_name)
        ws_index = rb.sheet_names().index(self.sheet_name)
        ws = wb.get_sheet(ws_index)

        # 在第8列写入公式，并在最后一行写入总和
        self.write_formulas(sheet, ws, 3, sheet.nrows, 7, rb)

    def write_formulas(self, sheet, ws, start_row, end_row, column_number, rb):
        for row_index in range(start_row, end_row):
            xf_index = sheet.cell_xf_index(row_index, column_number)
            xf_obj = rb.xf_list[xf_index]
            style = create_style(xf_obj, rb)

            # 写入第8列的计算公式
            formula = f'E{row_index + 1}*F{row_index + 1}'
            ws.write(row_index, column_number, Formula(formula), style)

        # 在倒数第二行后一行（即最后一行）写入总和公式
        # 注意：计算总和时不包括最后一行，以避免循环引用
        sum_formula = f'SUM(H4:H{end_row - 1})'
        ws.write(end_row - 1, column_number, Formula(sum_formula), style)





# GUI界面
class Application(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setGeometry(300, 300, 400, 300)
        self.setWindowTitle('文件处理')

        self.folder_path_label = QtWidgets.QLabel("未选择文件夹", self)
        self.folder_path_label.setAlignment(QtCore.Qt.AlignCenter)

        self.select_folder_button = QtWidgets.QPushButton("选择文件夹", self)
        self.select_folder_button.clicked.connect(self.select_folder)

        self.start_button = QtWidgets.QPushButton("开始处理", self)
        self.start_button.clicked.connect(self.process_files)

        self.info_label = QtWidgets.QLabel("", self)
        self.info_label.setAlignment(QtCore.Qt.AlignCenter)

        self.quit_button = QtWidgets.QPushButton("退出", self)
        self.quit_button.clicked.connect(QtWidgets.qApp.quit)

        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.folder_path_label)
        layout.addWidget(self.select_folder_button)
        layout.addWidget(self.start_button)
        layout.addWidget(self.info_label)
        layout.addWidget(self.quit_button)

        self.setLayout(layout)

    def select_folder(self):
        default_path = os.path.expanduser('~')  # 获取用户的主目录作为默认路径
        self.folder_selected = QtWidgets.QFileDialog.getExistingDirectory(self, "选择文件夹", default_path)
        if self.folder_selected:
            self.folder_path_label.setText(self.folder_selected)
            self.info_label.setText("文件夹已选择")
        else:
            self.folder_path_label.setText("未选择文件夹")
            self.info_label.setText("")

    def process_files(self):
        if hasattr(self, 'folder_selected') and self.folder_selected:
            for file in os.listdir(self.folder_selected):
                if file.endswith(".xls") or file.endswith(".xlsx"):
                    file_path = os.path.join(self.folder_selected, file)
                    sheet_processors = [
                        SheetProcessorFor对照表一(),
                        SheetProcessorFor对照表二(),
                        SheetProcessorFor对照表三甲()
                        # ... [添加其他 sheet 处理类] ...
                    ]
                    try:
                        process_workbook(file_path, sheet_processors)
                        self.info_label.setText(f"{file} 处理完成")
                    except Exception as e:
                        self.info_label.setText(f"{file} 处理失败: {e}")
            print("Processing completed.")
        else:
            self.info_label.setText("未选择文件夹")

# ... [其他函数和类定义，例如 SheetProcessorFor对照表一 等] ...

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

def main():
    app = QtWidgets.QApplication([])
    ex = Application()
    ex.show()
    app.exec_()

if __name__ == '__main__':
    main()