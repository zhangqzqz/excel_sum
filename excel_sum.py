# -*- coding:utf-8 -*-

import openpyxl 
import os
import time
import re
from functools import reduce
from shutil import copyfile
import logging
# logging
logging.basicConfig(format="%(levelname)s\t%(asctime)s\t%(message)s",filename="excel_sum.log")
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# 读表格
def rd_excel(file_name,row_nums):
    # df = pd.read_excel(f"input_excel\{file_name}")
    # print(df'3')
    dic = {}
    row_value_list =[]
    data = openpyxl.load_workbook(f"input_excel\{file_name}").active
    for row_num in row_nums:
        if row_num!='':
            row_value_list = [ [row_num+str(i.row),i.value] for i in (data[row_num]) ]
            logger.debug(f"{row_value_list}")
            for j in row_value_list:
                k,v = j
                if isinstance(v,int) == False and v!=None:
                    
                    get_num = [int(i) for i in re.sub("\D",",",v).split(',') if i!='']
                    if '是' in v: 
                        v = 1
                        logger.debug(f"{file_name} {k} {j[1]} to {v}")
                    elif '否' in v or v == '-':
                        v = 0
                        logger.debug(f"{file_name} {k} {j[1]} to {v}")
                    elif len(get_num)!=0:
                        v = sum(get_num)
                        logger.debug(f"{file_name} {k} {j[1]} to  {v}")
                    else:
                        
                        logger.debug(f"{file_name} {k} {j[1]}   {v}")
                if v != None:
                    dic[k] = v
            
    return dic

# 汇总数据
def sum_dict(a,b):
    temp = dict()
    # python3,dict_keys类似set； | 并集
    
    for key in a.keys()| b.keys():
        try:
            temp[key] = sum([d.get(key, 0) for d in (a, b)])
        except:
            logger.debug(f"{key}")

    return temp

# 写表格
def wr_excel(file_name,sum_dict):
    wb = openpyxl.load_workbook(file_name)
    data = wb.active
    for item in sum_dict.items():
        data[item[0]] = item[1]
    wb.save(file_name)
        



print("\nINFO:开始读取表格信息：")
row_nums = input("请输入您要汇总的列号(不区分大小写)，多列请用空格隔开，例如，列：D 或 多列：D F \n").upper().split(' ')

all_list = []
mytime = time.strftime("%Y%m%d%H%M%S",time.localtime())
output_file = f'output_excel\output{mytime}.xlsx'
for files in os.walk('input_excel'):
    copyfile(f'input_excel\{files[2][-1]}',output_file)
    for file_name in files[2]:
        if file_name[0:2]!='~$':
            all_list.append(rd_excel(file_name,row_nums))

print("\nINFO:开始计算指定数据的汇总值：")
sum_dict = (reduce(sum_dict,all_list))

print("\nINFO:开始将汇总结果写入汇总表格：")

wr_excel(output_file,sum_dict)

print(f"\nINFO:汇总表格生成完成，路径为\n{output_file}")
time.sleep(5)


