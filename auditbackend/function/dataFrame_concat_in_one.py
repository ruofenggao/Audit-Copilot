import os
import pandas as pd
import config

class ConcatInOne:
    @staticmethod
    def _validate_excel_path(path):
        if not isinstance(path, str):
            raise TypeError("Path should be a string.")
        if not os.path.exists(path):
            raise FileNotFoundError(f"File '{path}' not found.")
        if not path.endswith('.xls') and not path.endswith('.xlsx'):
            raise ValueError(f"File '{path}' is not a recognized Excel format.")

    @staticmethod
    def merge_sheets_in_excel():
        excel_path = config.EXCEL_PATH  # 假设 EXCEL_PATH 是在 config.py 中定义的变量
        ConcatInOne._validate_excel_path(excel_path)

        try:
            xls = pd.ExcelFile(excel_path)
            dfs = [xls.parse(sheet) for sheet in xls.sheet_names]
            merged_df = pd.concat(dfs, ignore_index=True)
            return [merged_df]
        except Exception as e:
            raise ValueError(f"Error merging sheets in '{excel_path}': {e}")

    @staticmethod
    def merge_multiple_excels():
        excel_paths = config.EXCEL_PATHS_FOR_CONCAT  # 假设 EXCEL_PATHS_FOR_CONCAT 是在 config.py 中定义的变量
        if not isinstance(excel_paths, list):
            raise TypeError("Expected a list of Excel paths.")

        for path in excel_paths:
            ConcatInOne._validate_excel_path(path)

        try:
            dfs = [pd.read_excel(path) for path in excel_paths]
            merged_df = pd.concat(dfs, ignore_index=True)
            return [merged_df]
        except Exception as e:
            raise ValueError(f"Error merging excels: {e}")
    @staticmethod
    def merge():
        if config.CONCAT_MODE == 'sheets':
            return ConcatInOne.merge_sheets_in_excel()
        elif config.CONCAT_MODE == 'excels':
            return ConcatInOne.merge_multiple_excels()
        else:
            raise ValueError("Invalid CONCAT_MODE in config. Expected 'sheets' or 'excels'.")