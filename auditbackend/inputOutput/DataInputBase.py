import warnings
import pandas as pd
from tqdm import tqdm

warnings.simplefilter("ignore")

class DataInputBase:
    def __init__(self, mode, *sources):
        self.mode = mode
        self.sources = sources

    @staticmethod
    def _read_excel_with_progress(filepath):
        try:
            with tqdm(total=100, desc=f"Reading {filepath}") as pbar:
                df = pd.read_excel(filepath, engine='openpyxl')
                pbar.update(100)
            return [df]
        except Exception as e:
            print(f"Error reading {filepath}: {e}")
            return []

    @staticmethod
    def _read_excel_with_progress_for_all_sheets(filepath):
        try:
            with tqdm(total=100, desc=f"Reading all sheets from {filepath}") as pbar:
                all_sheets = pd.read_excel(filepath, sheet_name=None, engine='openpyxl')
                pbar.update(100)
            return list(all_sheets.values())
        except Exception as e:
            print(f"Error reading {filepath}: {e}")
            return []

    @staticmethod
    def _clean_dataframe(df):
        # 将浮点数列转换为整数，然后再转换为字符串，并过滤小数点
        for col in df.select_dtypes(include=['float64']).columns:
            # 创建一个mask，标记哪些值是NaN
            nan_mask = df[col].isna()

            # 对非NaN的值进行处理，先转换为int，然后转换为str
            df.loc[~nan_mask, col] = df.loc[~nan_mask, col].astype(int).astype(str)

            # 对NaN的值进行处理
            df.loc[nan_mask, col] = pd.NA

            # 去掉末尾的'.0'，实际上由于之前已经将非NaN值转换为int，这一步可能已经不再需要
            df[col] = df[col].str.replace('.0$', '', regex=True)
        return df

    def load_dataframes(self):
        dfs = []
        if self.mode == 'single':
            for fp in self.sources:
                raw_dfs = self._read_excel_with_progress(fp)
                cleaned_dfs = [self._clean_dataframe(df) for df in raw_dfs]
                dfs.extend(cleaned_dfs)
        elif self.mode == 'multi':
            for filepath in self.sources:
                raw_dfs = self._read_excel_with_progress_for_all_sheets(filepath)
                cleaned_dfs = [self._clean_dataframe(df) for df in raw_dfs]
                dfs.extend(cleaned_dfs)
        else:
            raise ValueError(f"Unsupported mode: {self.mode}")
        return dfs
