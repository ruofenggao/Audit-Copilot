
import time
from auditbackend.function.dataFrame_merger import DataFrameMerger
from auditbackend.inputOutput.data_output_base import DataOutputBase
from auditbackend.inputOutput.data_input_base import DataInputBase
import config

start_time = time.time()

# 使用DataInputBase加载所有目标数据
processor = DataInputBase(config.AUDIT_INPUT_MODE, *config.AUDIT_DATA_PATHS)
target_dfs = processor.load_dataframes()

# 使用DataInputBase加载对比数据
compare_processor = DataInputBase(config.COMPARE_INPUT_MODE, *config.COMPARE_DATA_PATHS)
compare_dfs = compare_processor.load_dataframes()

# 使用DataFrameComparator的check_inclusion方法处理数据
result_dfs = DataFrameMerger.merge_on_columns(target_dfs, compare_dfs, config.TARGET_COLUMNS, config.COMPARE_COLUMNS)

# 使用DataOutputBase保存处理后的数据
DataOutputBase.save_dataframes_with_prefix(result_dfs, config.OUTPUT_PATH_PREFIX, mode=config.OUTPUT_TYPE)

end_time = time.time()
print(f"Elapsed time: {end_time - start_time:.2f} seconds")
