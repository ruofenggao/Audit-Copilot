import pandas as pd

class DataOutputBase:

    @staticmethod
    def _save_large_dataframe_to_multiple_sheets(df, output_path, max_rows=1048576):
        with pd.ExcelWriter(output_path) as writer:
            num_sheets = len(df) // max_rows + 1
            for i in range(num_sheets):
                start_row = i * max_rows
                end_row = (i + 1) * max_rows
                subset_df = df.iloc[start_row:end_row]
                sheet_name = f"Sheet_{i + 1}"
                subset_df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Saved rows {start_row}-{end_row} to {sheet_name} in {output_path}")

    @staticmethod
    def save_dataframes(dataframes, mode="excel", output_path=None):
        if not isinstance(dataframes, list):
            dataframes = [dataframes]

        if mode == "excel":
            DataOutputBase._save_to_excel(dataframes, output_path=output_path)
        elif mode == "database":
            print("Database mode is currently not supported.")
            # Placeholder for future database integration
            pass

    @staticmethod
    def _save_to_excel(dataframes, output_path=None):
        if output_path is None:
            output_path = [f"./output_{idx}.xlsx" for idx in range(len(dataframes))]
        elif not isinstance(output_path, list):
            output_path = [output_path]

        for idx, df in enumerate(dataframes):
            if len(df) <= 1048576:
                df.to_excel(output_path[idx], index=False)
                print(f"Saved DataFrame to {output_path[idx]}")
            else:
                # Handle large DataFrame
                DataOutputBase._save_large_dataframe_to_multiple_sheets(df, output_path[idx])

    @staticmethod
    def save_dataframes_with_prefix(dataframes, prefix, mode="excel", output_path=None):
        print(f"Starting the {mode} output process...")
        if output_path is None:
            output_path = [f"{prefix}{idx}.xlsx" for idx in range(len(dataframes))]
        DataOutputBase.save_dataframes(dataframes, mode=mode, output_path=output_path)
