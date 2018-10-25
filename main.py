
# 这是用于批量处理excel表格的程序
import os

from excel import Excel

# 执行函数
def run():
    # dir_name = input("请输入要处理的文件夹路径")
    dir_name = "G:/huangwen/code/excel"

    # 为了循环使用，不做单例
    excel = Excel()

    # 先检查输入
    if excel.check_dir(dir_name) == False:
        return

    # 对目标文件进行读取数据操作
    excel.read_data()

    print("处理完毕")

# 程序运行
run()
#input()


