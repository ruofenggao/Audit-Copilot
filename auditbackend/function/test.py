import pandas as pd

from auditbackend.function.dataFrame_comparator import DataFrameComparator

compare_dfs = [pd.DataFrame({"营销库TAC码": ["86336306", "86014306", "86337407", "86087506", "86921106", "86314906", "86891706", "86520106"]})]
target_dfs = [pd.DataFrame({"TAC": ["86336306"]})]

result_dfs = DataFrameComparator.check_inclusion(target_dfs, 'TAC', compare_dfs, '营销库TAC码')

print(result_dfs[0])