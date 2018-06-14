# -*- coding: utf-8 -*-

import xlrd, xlwt
import json
import codecs
from collections import OrderedDict


class Xls(object):
    """
    操作xls文件，方便读写。
    """
    def __init__(self):
        # 创建对象
        self.wt = xlwt.Workbook()
        self.title_key = True
        # 创建默认页面
        self.table = self.wt.add_sheet('sheet1', cell_overwrite_ok=True)
        # 记录程序行数
        self.y = 0

    def create_xls(self, sheet_name='sheet1'):
        """
        创建自定义xls对象，当多个页面时定义，返回页面对象，可对页面对象进行读写
        :param sheet_name: sheet名字，初始化定义
        :return: table对象（页面对象）
        """
        table = self.wt.add_sheet(sheet_name.decode('utf-8'), cell_overwrite_ok=True)
        return table

    def write_title(self, title_list=None, table=None):
        """
        保存标题，默认写第一行
        :param title_list: 标题列表
        :param table: 页面对象，不写默认使用sheet1
        :return: 无返回
        """
        if table is None:
            table = self.table
        # 设置全局标题，应用于json数据，保证数据与标题对应
        self.title_data_ = title_list
        for i, title in enumerate(title_list):
            table.write(0, i, title)
        # 设置数据长度，检验用的
        self.title_len = dict(zip(self.title_data_, [1 for i in range(len(self.title_data_))]))
        self.y += 1

    def write_data(self, y=None, dict_data=None, list_data=None, dict_list=None, table=None, title=None):
        """保存正文
        :param dict_list:
        :param list_data:
        :param table:
        :param dict_data:
        :type y: object
        """
        if y is None:
            y = self.y
        if table is None:
            table = self.table
        if title is None:
            title = self.title_data_
        if dict_data is not None:
            self.__dict_data(dict_data=dict_data, y=y, table=table, title=title)
        elif list_data is not None:
            self.__list_data(list_data, y, table)
        elif dict_list is not None:
            self.__dict_list(dict_list, y, table)
        else:
            raise TypeError(u'没有数据进入！！')
        if y is None:
            self.y += 1

    def __dict_data(self, dict_data, y, table, title):
        """保存字典类型数据"""
        for i, line in enumerate(title):
            try:
                table.write(y, i, dict_data[line])
            except KeyError:
                print(line + u'字段有不存在的情况')

    def __list_data(self, list_data, y, table):
        """保存列表类型数据"""
        for i, data in enumerate(list_data):
            table.write(y, i, data)

    def __dict_list(self, dict_list, y, table):
        """保存字典嵌套列表类型数据, 列表长度必须固定"""
        length = 0
        list_len = []
        list_name = []
        for line in self.title_data_:
            if type(dict_list[line]) is list:
                list_len.append(len(dict_list[line]))
                list_name.append(line)
                for j, data in enumerate(dict_list[line]):
                    table.write(y, length + j, data)
                length += len(dict_list[line])
            else:
                try:
                    table.write(y, length, dict_list[line])
                    length += 1
                except KeyError:
                    print(line + u'字段有不存在的情况')

        if self.title_key:
            if len(list_name) != 0:
                x = 0
                for title in self.title_data_:
                    if title in list_name:
                        k = len(dict_list[title])
                        if self.title_len[title] > k:
                            table.write_merge(0, 0, x, x + list_len[k], title)
                            x += k
                            self.title_len[title] = k
                    else:
                        table.write(0, x, title)
                        x += 1
                self.title_key = False

    def read(self, fileName, file=False):
        """读文件"""
        wb = xlrd.open_workbook(fileName)

        convert_list = []
        sh = wb.sheet_by_index(0)
        title = sh.row_values(0)
        for rownum in range(1, sh.nrows):
            rowvalue = sh.row_values(rownum)
            single = OrderedDict()
            for colnum in range(0, len(rowvalue)):
                single[title[colnum]] = rowvalue[colnum]
            convert_list.append(single)

        if file:
            j = json.dumps(convert_list)
            print j

            with codecs.open(fileName, "w", "utf-8") as f:
                f.write(j)
        else:
            return

    def save_filed(self, filed_name):
        """保存文件"""
        self.wt.save(filed_name + '.xls')


if __name__ == '__main__':
    Xls().read(u'测试程序.xls')
