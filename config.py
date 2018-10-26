
# 这是输出配置类，给使用者配置参数
# 用py文件是为了方便写备注

class Config():
    # 文件路径
    dir_path = "G:/huangwen/code/excel/el"
    # 输出路径
    putout = "G:/huangwen/code/excel/el"
    # 输出的文件名
    file_name = "file.xlsx"
    # 主键(不可为空，如果为空，该行数据无效，用于去除无效数据)
    primary_key = "id"
    # 排序键，用来做排序的键, 留 "" 表示不做排序处理
    order_key = "zzWF"
    # 升降序标志
    order = False



# 实例化一个单例
config = Config()