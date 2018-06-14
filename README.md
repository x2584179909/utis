# 方便的操作xls文件

### 目前支持格式：列表嵌套字符串、字典、字符串

### 主要用于对数据进行处理后检验（其实可以随便用）还有一个对mongodb的数据特性进行统计的脚本，过两天放上来 

下面是演示程序
```
# coding: utf-8

import json
from xls import Xls


class xls_test(Xls):

    def __init__(self):
        Xls.__init__(self)
        # 创建xls文件对象,可有可无
        self.data_xls = self.create_xls('处理后')
        self.write_title(title_list=['brandName_en', 'brand_name', 'status', 'company_test', 'site', 'key'],
                         table=self.data_xls)
        pass

    def read_dict_data(self):
        item_list = []
        with open(u'演示.json') as f:
            for item in f.readlines():
                item_list.append(json.loads(item))
        
        return item_list
    
    def save_data(self, data_list):
        for item in data_list:
            self.write_data(dict_data=item)
    
    def run(self):
        item_list = self.read_dict_data()
        self.save_data(item_list)
        self.save_filed(u'文件名')  # 文件名不用加后缀}
        
                
