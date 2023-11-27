import pandas as pd
from itertools import product


class DataFrameMerger:

    @staticmethod
    def _validate_args(df1, df2, columns_df1, columns_df2):
        if not isinstance(df1, pd.DataFrame) or not isinstance(df2, pd.DataFrame):
            raise TypeError("Both df1 and df2 should be of type pandas.DataFrame.")
        for column in columns_df1:
            if column not in df1.columns:
                raise ValueError(f"Column '{column}' not found in df1.")
        for column in columns_df2:
            if column not in df2.columns:
                raise ValueError(f"Column '{column}' not found in df2.")

    @staticmethod
    def _generate_rows_with_split_values(df, columns):
        for _, row in df.iterrows():
            split_values = {column: str(row[column]).split(',') for column in columns}
            other_values = {col: row[col] for col in df.columns if col not in columns}
            for values in product(*split_values.values()):
                combined = {**dict(zip(columns, values)), **other_values}
                yield combined

    @staticmethod
    def merge_on_columns(df1_list, df2_list, columns_df1, columns_df2):
        print("Starting the merging process...")

        if not isinstance(df1_list, list):
            df1_list = [df1_list]
        if not isinstance(df2_list, list):
            df2_list = [df2_list]

        result_dfs = []

        for df1 in df1_list:
            for df2 in df2_list:
            # Use generator to get the expanded rows for df1 and df2
            # print("\nLeft DataFrame (df1) content:")
            # print(df1.head())
            # print("\nRight DataFrame (df2) content:")
            # print(df2.head())

                df1_expanded_rows = list(DataFrameMerger._generate_rows_with_split_values(df1, columns_df1))
                df2_expanded_rows = list(DataFrameMerger._generate_rows_with_split_values(df2, columns_df2))

                # Convert expanded rows into DataFrames
                df1_expanded = pd.DataFrame(df1_expanded_rows)
                df2_expanded = pd.DataFrame(df2_expanded_rows)

                DataFrameMerger._validate_args(df1_expanded, df2_expanded, columns_df1, columns_df2)

                # Rename conflicting columns in df2, if there are any same column names in df1 and df2
                conflicting_columns = set(columns_df1).intersection(set(columns_df2))
                if conflicting_columns:
                    rename_dict = {col: col + "_from_df2" for col in conflicting_columns}
                    df2_expanded.rename(columns=rename_dict, inplace=True)
                    columns_df2 = [rename_dict.get(col, col) for col in columns_df2]

                # Perform the merge operation
                merged_df = pd.merge(df1_expanded, df2_expanded, left_on=columns_df1,
                                     right_on=columns_df2, how='left', suffixes=('', '_from_df2'))

                matched_records = merged_df[columns_df2[0]].notna().sum()
                print(f"Found {matched_records} matching records in expanded df2 for expanded df1.")

                result_dfs.append(merged_df)

        print("Merging process completed!")
        return result_dfs
