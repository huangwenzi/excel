
# 这个是执行excel功能的类

import os
from openpyxl import Workbook  # 新建时导入这个
from openpyxl import load_workbook  # 读取时导入这个


class Excel():

    # 检查目录
    # dir_name : 路径名
    # 返回目录下的excel格式的文件名列表
    def check_dir(self, dir_name):
        # 检查路径是否存在
        if os.path.exists(dir_name) == False:
            print("路径不存在")
            return False

        # 获取目录下的文件名
        file_name_list = os.listdir(dir_name)
        print("目录下有文件:")
        print(file_name_list)

        # 查找符合的文件
        name_list = []
        for name in file_name_list:
            if name.find(".xlsx") != -1:
                name_list.append(name)
        print("符合条件的文件有:")
        print(name_list)

        self.dir_name = dir_name
        self.name_list = name_list

        return True


    # 读取全部文件数据到内存中，建立一个列表保存
    # 并对数据进行总的合并
    def read_data(self):
        # 检查文件列表
        if hasattr(self, "name_list") == False:
            print("name_list 不存在")
            return False

        name_list = self.name_list

        # 开始读取数据
        excel_list = []
        for index in range(0, len(name_list)):
            workbook = load_workbook(
                self.dir_name + "/" + name_list[index])
            excel_list.append(workbook)

        # 目前只对每一个表格的第一个sheet做处理
        sheet_list = []
        for index in range(0, len(name_list)):
            tmp = excel_list[index].get_sheet_names()
            sheet_list.append(excel_list[index][tmp[0]])

        # 获取每个列表的属性名列表
        keys_list = []
        for index in range(0, len(name_list)):
            attr_count = sheet_list[index].max_column
            # print(attr_count)
            tmp = []
            for index_1 in range(0, attr_count):
                tmp.append(sheet_list[index].cell(row=1,column=index_1+1).value)

            keys_list.append(tmp)
            # print(tmp)

        # 需要把每一行数据都转换成字典对象的形式保存才好操作
        dict_list = []
        # 这层循环是表格数的
        for index in range(0, len(name_list)):
            row_list = []
            row_count = sheet_list[index].max_row
            attr_count = len(keys_list[index])
            # 这层循环是行数的
            for index_1 in range(1, row_count):
                row_dict = {}
                # 这层循环是列数的
                for index_2 in range(0, attr_count):
                    row_dict[keys_list[index][index_2]] = sheet_list[index].cell(row=index_1+1,column=index_2+1).value
                row_list.append(row_dict)
            dict_list.append(row_list)

        # 属性列表合并去重，获得组合在一起的包含全部的属性表
        tmp_attr = []
        sum_attr = []
        for index in range(0, len(keys_list)):
            tmp_attr += keys_list[index]
        for index in range(0, len(tmp_attr)):
            if tmp_attr[index] not in sum_attr:
                sum_attr.append(tmp_attr[index])

        print("去重前的属性表")
        print(tmp_attr)
        print("去重后的属性表")
        print(sum_attr)

        # 创建新的excel文档
        excel = Workbook()
        excel.create_sheet('sheet1', index=0)
        sheet1 = excel['sheet1']

        # 先写入第一行
        for index in range(0, len(sum_attr)):
            sheet1.cell(row=1,column=index+1).value = sum_attr[index]

        # 根据去重后的属性表进行数据的输入
        # 这层循环是整合的表数
        row_count = 2   # 当前应该写入的行数
        for index in range(0, len(dict_list)):
            tmp_dict = dict_list[index]
            # 这层循环是表里的行数
            for index_1 in range(0, len(dict_list[index])):
                tmp_row = tmp_dict[index_1]
                tmp_keys = tmp_row.keys()
                # 这层循环是属性数
                for index_3 in range(0, len(sum_attr)):
                    if sum_attr[index_3] not in tmp_keys:
                        continue
                    sheet1.cell(row=row_count,column=index_3+1).value = tmp_row[sum_attr[index_3]]

                row_count += 1

        # 保存excel表
        excel.save("test.xlsx")
                

        