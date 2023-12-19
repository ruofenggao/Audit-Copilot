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

    def write_formula_and_sum(self, sheet, ws, start_row, end_row, column_number, formula_pattern, include_sum, rb):
        for row_index in range(start_row, end_row - 1):
            xf_index = sheet.cell_xf_index(row_index, column_number)
            xf_obj = rb.xf_list[xf_index]
            style = create_style(xf_obj, rb)  # 在循环内定义style

            formula = formula_pattern.format(row_index + 1, row_index + 1)
            ws.write(row_index, column_number, Formula(formula), style)

        # 如果需要，添加总和公式
        if include_sum and end_row - start_row > 0:
            xf_index = sheet.cell_xf_index(end_row - 1, column_number)  # 重新获取style
            xf_obj = rb.xf_list[xf_index]
            style = create_style(xf_obj, rb)
            sum_formula = f'SUM({chr(65 + column_number)}4:{chr(65 + column_number)}{end_row - 1})'
            ws.write(end_row - 1, column_number, Formula(sum_formula), style)


# 对照表四甲（设备）的处理类暂时没有
class SheetProcessorFor对照表四甲设备(SheetProcessor):
    def __init__(self):
        super().__init__('对照表四甲（设备）')

    def process_sheet(self, rb, wb):
        sheet = rb.sheet_by_name(self.sheet_name)
        ws_index = rb.sheet_names().index(self.sheet_name)
        ws = wb.get_sheet(ws_index)


# 对照表四甲材料的处理类
class SheetProcessorFor对照表四甲材料(SheetProcessor):
    def __init__(self):
        super().__init__('对照表四甲（材料）')

    def process_sheet(self, rb, wb):
        sheet = rb.sheet_by_name(self.sheet_name)
        ws_index = rb.sheet_names().index(self.sheet_name)
        ws = wb.get_sheet(ws_index)


# 对照表三丙的处理类
class SheetProcessorFor对照表三丙(SheetProcessor):
    def __init__(self):
        super().__init__('对照表三丙')

    def process_sheet(self, rb, wb):
        sheet = rb.sheet_by_name(self.sheet_name)
        ws_index = rb.sheet_names().index(self.sheet_name)
        ws = wb.get_sheet(ws_index)


# 对照表三乙的处理类
class SheetProcessorFor对照表三乙(SheetProcessor):
    def __init__(self):
        super().__init__('对照表三乙')

    def process_sheet(self, rb, wb):
        sheet = rb.sheet_by_name(self.sheet_name)
        ws_index = rb.sheet_names().index(self.sheet_name)
        ws = wb.get_sheet(ws_index)

        # 处理 U 列 (索引为 20)，T列*J列，不包含总和
        self.write_formula_and_sum(sheet, ws, 4, sheet.nrows, 20, 'T{}*J{}', True, rb)
        # 处理 J 列 (索引为 9)，H列*I列，不包含总和
        self.write_formula_and_sum(sheet, ws, 4, sheet.nrows, 9, 'H{}*I{}', True, rb)
        # 处理 I 列 (索引为 8)，E列*G列，不包含总和
        self.write_formula_and_sum(sheet, ws, 4, sheet.nrows, 8, 'E{}*G{}', False, rb)
        # 处理 T 列 (索引为 19)，R列*S列，不包含总和
        self.write_formula_and_sum(sheet, ws, 4, sheet.nrows, 19, 'R{}*S{}', True, rb)
        # 处理 S 列 (索引为 18)，Q列*O列，不包含总和
        self.write_formula_and_sum(sheet, ws, 4, sheet.nrows, 18, 'Q{}*O{}', False, rb)


# 对照表三甲的处理类
class SheetProcessorFor对照表三甲(SheetProcessor):
    def __init__(self):
        super().__init__('对照表三甲')

    def process_sheet(self, rb, wb):
        sheet = rb.sheet_by_name(self.sheet_name)
        ws_index = rb.sheet_names().index(self.sheet_name)
        ws = wb.get_sheet(ws_index)

        # 处理 S 列 (索引为 18)，N列*O列，包含总和
        self.write_formula_and_sum(sheet, ws, 3, sheet.nrows, 18, 'Q{}-H{}', True, rb)
        # 处理 T 列 (索引为 19)，N列*O列，包含总和
        self.write_formula_and_sum(sheet, ws, 3, sheet.nrows, 19, 'R{}-I{}', True, rb)
        # 处理 Q 列 (索引为 16)，N列*O列，包含总和
        self.write_formula_and_sum(sheet, ws, 3, sheet.nrows, 16, 'N{}*O{}', True, rb)
        # 处理 R 列 (索引为 17)，N列*O列，包含总和
        self.write_formula_and_sum(sheet, ws, 3, sheet.nrows, 17, 'N{}*P{}', True, rb)
        # 处理 I 列 (索引为 8)，E列*G列，包含总和
        self.write_formula_and_sum(sheet, ws, 3, sheet.nrows, 8, 'E{}*G{}', True, rb)
        # 处理 H 列 (索引为 7)，E列*G列，包含总和
        self.write_formula_and_sum(sheet, ws, 3, sheet.nrows, 7, 'E{}*F{}', True, rb)


# 对照表一的处理类
class SheetProcessorFor对照表一(SheetProcessor):
    def __init__(self):
        super().__init__('对照表一')

    def process_sheet(self, rb, wb):
        sheet = rb.sheet_by_name(self.sheet_name)
        ws_index = rb.sheet_names().index(self.sheet_name)
        ws = wb.get_sheet(ws_index)


# 对照表二的处理类
class SheetProcessorFor对照表二(SheetProcessor):
    def __init__(self):
        super().__init__('对照表二')

    def process_sheet(self, rb, wb):
        sheet = rb.sheet_by_name(self.sheet_name)
        ws_index = rb.sheet_names().index(self.sheet_name)
        ws = wb.get_sheet(ws_index)


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
            processed_files = []  # 用于跟踪所有处理过的文件

            for file in os.listdir(self.folder_selected):
                if file.endswith(".xls") or file.endswith(".xlsx"):
                    file_path = os.path.join(self.folder_selected, file)
                    sheet_processors = [
                        SheetProcessorFor对照表三丙(),
                        SheetProcessorFor对照表三乙(),
                        SheetProcessorFor对照表三甲(),
                        SheetProcessorFor对照表四甲材料(),
                        SheetProcessorFor对照表四甲设备(),
                        SheetProcessorFor对照表一(),
                        SheetProcessorFor对照表二()
                        # ... [添加其他 sheet 处理类] ...
                    ]
                    try:
                        process_workbook(file_path, sheet_processors)
                        processed_files.append(file)  # 添加文件到已处理列表
                    except Exception as e:
                        self.info_label.setText(f"{file} 处理失败: {e}")

            if processed_files:
                processed_files_str = "\n".join(processed_files)
                self.info_label.setText(f"以下文件已处理完成:\n{processed_files_str}")
            else:
                self.info_label.setText("未找到可处理的 Excel 文件")

        else:
            self.info_label.setText("未选择文件夹")

    # ... [其他函数和类定义，例如 SheetProcessorFor对照表一 等] ...


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
