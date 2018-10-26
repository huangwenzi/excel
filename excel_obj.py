
# 这是数据对象

class Excel_obj():

    # 初始化对应的数据
    def __init__(self, config):
        # 文件路径
        self.dir_path = config.dir_path
        # 输出路径
        self.putout = config.putout
        # 输出的文件名
        self.file_name = config.file_name
        # 主键(不可为空，如果为空，该行数据无效，用于去除无效数据)
        self.primary_key = config.primary_key
        # 排序键，用来做排序的键
        self.order_key = config.order_key
        # 升降序标志
        self.order = config.order



        # 下面这些是在执行过程中会赋值的，放出来有个了解
        # 目录下的excel文件名
        self.name_list = []
        # 数据行的字典行总表
        self.data_list = []
        # 属性列表
        self.attr_list = []