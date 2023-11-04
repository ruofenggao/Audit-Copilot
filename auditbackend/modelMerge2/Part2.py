import time
from auditbackend.function.DataFrameComparator import DataFrameComparator
from auditbackend.inputOutput.DataOutputBase import DataOutputBase
from auditbackend.inputOutput.DataInputBase import DataInputBase
import config

start_time = time.time()

# 使用DataInputBase加载所有目标数据
processor = DataInputBase(config.AUDIT_INPUT_MODE, *config.AUDIT_DATA_PATHS)
target_dfs = processor.load_dataframes()

# 使用DataInputBase加载对比数据
compare_processor = DataInputBase(config.COMPARE_INPUT_MODE, *config.COMPARE_DATA_PATHS)
compare_dfs = compare_processor.load_dataframes()

# 使用DataFrameComparator的check_inclusion方法处理数据，增加FILTER_RESULT参数
result_dfs = DataFrameComparator.check_inclusion(target_dfs, compare_dfs, config.TARGET_COLUMNS, config.COMPARE_COLUMNS,
                                                 config.FILTER_RESULT)

# 使用DataOutputBase保存处理后的数据
DataOutputBase.save_dataframes_with_prefix(result_dfs, config.OUTPUT_PATH_PREFIX, config.OUTPUT_TYPE)

end_time = time.time()
print(f"Elapsed time: {end_time - start_time:.2f} seconds")
