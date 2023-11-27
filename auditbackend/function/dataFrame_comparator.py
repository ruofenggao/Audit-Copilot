from itertools import product

import pandas as pd


class DataFrameComparator:

    @staticmethod
    def _validate_input(df, column_names):
        if not isinstance(df, pd.DataFrame):
            raise TypeError(f"Expected a DataFrame, but received {type(df)}")
        for column_name in column_names:
            if column_name not in df.columns:
                raise ValueError(f"Column '{column_name}' not found in the DataFrame.")

    @staticmethod
    def _split_and_explode(df, columns):
        for _, row in df.iterrows():
            split_values = {column: str(row[column]).split(',') for column in columns}
            other_values = {col: row[col] for col in df.columns if col not in columns}
            for values in product(*split_values.values()):
                combined = {**dict(zip(columns, values)), **other_values}
                yield combined

    @staticmethod
    def check_inclusion(target_dfs, compare_dfs, target_columns, compare_columns, filter_result='all'):
        print("Starting the inclusion check...")

        if not isinstance(target_dfs, (list, pd.DataFrame)):
            raise TypeError("Expected a DataFrame or a list of DataFrames for target_dfs.")
        if not isinstance(compare_dfs, (list, pd.DataFrame)):
            raise TypeError("Expected a DataFrame or a list of DataFrames for compare_dfs.")

        target_columns = [target_columns] if isinstance(target_columns, str) else target_columns
        compare_columns = [compare_columns] if isinstance(compare_columns, str) else compare_columns
        target_dfs = [target_dfs] if isinstance(target_dfs, pd.DataFrame) else target_dfs
        compare_dfs = [compare_dfs] if isinstance(compare_dfs, pd.DataFrame) else compare_dfs

        for df in target_dfs:
            DataFrameComparator._validate_input(df, target_columns)
        for df in compare_dfs:
            DataFrameComparator._validate_input(df, compare_columns)

        results = []

        total_set = set()
        for df in compare_dfs:
            df = pd.DataFrame(DataFrameComparator._split_and_explode(df, compare_columns))
            combined_series = df[compare_columns].astype(str).apply(lambda x: '|'.join(x), axis=1)
            total_set.update(combined_series.str.strip().str.lower())

        replace_dict = {True: 'true', False: 'false'}
        for target_df in target_dfs:
            target_df = pd.DataFrame(DataFrameComparator._split_and_explode(target_df, target_columns))

            def check_row(row):
                combined_val = '|'.join(row)
                return replace_dict[combined_val.lower() in total_set]

            target_df['Check'] = target_df[target_columns].astype(str).apply(check_row, axis=1)

            true_count_original = len(target_df[target_df['Check'] == 'true'])
            false_count_original = len(target_df[target_df['Check'] == 'false'])
            print(
                f"For the current target DataFrame, there are {true_count_original} 'true' and {false_count_original} 'false' records.")

            # 根据 filter_result 过滤数据
            if filter_result == 'true':
                target_df = target_df[target_df['Check'] == 'true']
                print("Filtered results to show only 'true' records.")
            elif filter_result == 'false':
                target_df = target_df[target_df['Check'] == 'false']
                print("Filtered results to show only 'false' records.")
            elif filter_result == 'all':
                print("Showing all records without filtering.")
            else:
                raise ValueError("Invalid value for filter_result. Accepted values are: 'all', 'true', 'false'.")

            results.append(target_df)

        print("Inclusion check completed!")
        return results
