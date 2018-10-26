
# 这是用于批量处理excel表格的程序
import os

from excel import excel
from excel_obj import Excel_obj as e_obj
from config import config

# 执行函数
def run():
    # 根据默认配置生产数据对象
    data = e_obj(config)

    # 先检查输入
    if excel.check_dir(data) == False:
        return

    # 对目标文件进行读取数据操作
    excel.read_data(data)

    # 对文件排序
    excel.sort_data(data)

    # 保存文件
    excel.save(data)

    print("处理完毕")

# 程序运行
run()
#input()


