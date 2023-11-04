# config.py

# 文件路径配置

# 数据文件路径
AUDIT_DATA_PATHS = ['/Users/ruofenggao/Desktop/重点数据大全/根据集团匹配出的问题数据DM/集团DM在库已激活.xlsx', ]
COMPARE_DATA_PATHS = ['/Users/ruofenggao/Desktop/重点数据大全/1008集团出出来的数据/模型10-终端库存合规性.xlsx']

# 输入类型
AUDIT_INPUT_MODE = 'single'
COMPARE_INPUT_MODE = 'single'

# 输出数据的类型
INPUT_TYPE = 'excel'
OUTPUT_TYPE = 'excel'

# 对比的列名 - 现在它们是列表
TARGET_COLUMNS = ['IMEI']
COMPARE_COLUMNS = ['串号_SERBUN']

FILTER_RESULT = 'all'  # 'none' | 'true' | 'f

# 如果有更多的列要比较，只需在列表中添加，如
# TARGET_COLUMNS = ['ACC_NBR', 'COL2', 'COL3']
# COMPARE_COLUMNS = ['BUND_ACC_NBR', 'BUND_COL2', 'BUND_COL3']

# 输出文件的前缀
OUTPUT_PATH_PREFIX = "/Users/ruofenggao/Desktop/outputExcel/temp_"
