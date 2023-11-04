import pandas as pd

from FanZhongDuan.function.DataFrameComparator import DataFrameComparator
from FanZhongDuan.function.DataFrameMerger import DataFrameMerger

# 使用你的 DataFrameMerger 和 DataFrameComparator 定义

# 创建模拟数据
data1 = {
    'A': ['1,2,3', '4,5', '6,7,8,9'],
    'B': ['apple,banana', 'cherry', 'date,fig,grape'],
    'Other1': ['info1', 'info2', 'info3']
}

data2 = {
    'X': ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'],
    'Y': ['apple', 'banana', 'apple', 'cherry', 'cherry', 'date', 'fig', 'grape', 'kiwi', 'lemon'],
    'Other2': ['data1', 'data2', 'data3', 'data4', 'data5', 'data6', 'data7', 'data8', 'data9', 'data10']
}

df1 = pd.DataFrame(data1)
df2 = pd.DataFrame(data2)

# 使用 DataFrameMerger 进行测试
results_merger = DataFrameMerger.merge_on_columns(df1, df2, ['A', 'B'], ['X', 'Y'])
print(results_merger[0])

# 使用 DataFrameComparator 进行测试
results_comparator = DataFrameComparator.check_inclusion(df1, df2, ['A', 'B'], ['X', 'Y'])
print(results_comparator[0])

# 这应该给我们DataFrameMerger和DataFrameComparator的输出，并允许我们检查它们的行为是否相同。
