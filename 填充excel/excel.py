
# excel
# 对象类


# 这是数据对象

class Excel():

    # 初始化对应的数据
    def __init__(self, config):
        # 文件路径
        self.dir_path = config.dir_path
        # 输出路径
        self.putout = config.putout
        # 输出的文件名
        # self.file_name = config.file_name
        # 起始行数，字段名在的行数
        self.begin_row = config.begin_row
        # 参考键, 用来做是否换一个对象的参考, 就如果下一行的这个字段变了，就是换对象
        self.consult_key = config.consult_key
        # 需要填充的键, 这里的键, 如果不存在将填0
        self.fill_key = config.fill_key
        # 保留键，数据参考上一行有效数据的值填入, 剩下的就留空
        self.Retain_key = config.Retain_key
        # 増位键, 这里的键值会比上一行的多一，或少一,剩下的就留空
        self.add_key = config.add_key


        # 下面这些是在执行过程中会赋值的，放出来有个了解
        # 目录下的excel文件名
        self.name_list = []
        # 数据行的字典行总表
        self.data_list = []
        # 属性列表
        self.attr_list = []

