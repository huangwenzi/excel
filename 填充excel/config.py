
# 这是输出配置类，给使用者配置参数
# 用py文件是为了方便写备注
# 因为必需要有一个示范，所以第一个必须要完整，作为参考

import toml

class Config():
    # # 文件路径
    # dir_path = "./el/1.xlsx"
    # # 输出路径
    # putout = "./el1/1.xlsx"
    # # 起始行数，字段名在的行数
    # begin_row = 4
    # # 输出的文件名
    # # file_name = "file.xlsx"
    # # 参考键, 用来做是否换一个对象的参考, 就如果下一行的这个字段变了，就是换对象
    # consult_key = "名字"
    # # 需要填充的键, 这里的键, 如果不存在将填0
    # fill_key = ["语文", "数学"]
    # # 保留键，数据参考上一行有效数据的值填入, 剩下的就留空
    # Retain_key = ["名字"]
    # # 増位键, 这里的键值会比上一行的多一，或少一,剩下的就留空
    # add_key = ["学期"]

    # 读取目录下的配置表格
    def __init__(self):
        # 把toml文件转化为字典数据
        f = open("config.toml", errors='ignore', encoding = "utf8")
        cfg = toml.load(f)["config"]

        # 开始进行配置处理
        self.dir_path = cfg["dir_path"]
        self.putout = cfg["putout"]
        self.begin_row = cfg["begin_row"]
        self.consult_key = cfg["consult_key"]
        self.fill_key = cfg["fill_key"]
        self.Retain_key = cfg["Retain_key"]
        self.add_key = cfg["add_key"]
        

# 实例化一个单例
config = Config()