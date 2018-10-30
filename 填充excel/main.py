
# 程序主入口
from excel import Excel
from operation import operation
from config import Config

def run():
    # name = input("请输入作为配置表的文件名")
    cfg = Config()
    # cfg.read(name)
    # 根据默认配置生产数据对象
    data = Excel(cfg)

    # 先检查输入
    if operation.read_data(data) == False:
        return

    print("处理完毕")
    input("回车退出程序")

# 程序运行
run()