# -*- coding: utf-8 -*-
"""
Created on Wed Mar 22 17:02:41 2023

@author: Jinliang
"""
import math
import pandas as pd

# 需要保留两位有效数字的列数据
round_two_list = ["应收权益1-本利和", "房屋总价2", "抵债金额（总）", "剩余购房款（不含首付）", 
                  "乙方1产品1剩余本金", "乙方1产品1剩余收益", "乙方1产品1转让本金", "乙方1产品1转让收益"]

def handle_round_two(l):
    """
    一个list中的float，全部保留2位有效数字
        全部换成str会怎么样...
    """
    return [('%.2f'%x) for x in l]  # 强转float的时候，会丢失.00这种2位小数，只能存成str，后期要运算再变回来
    
def load_excel(excel_path):
    """
    读取excel，做成字典
    {'column' : [item, item, ...], ...}
    """
    # 身份证需要按str读取，其余按照excel内自己的格式
    df = pd.read_excel(excel_path, converters={"身份证号码1": str, '购房人身份证1': str})  # 又学到了，converter
    columns = list(df.columns)
    
    dic = {}
    for col in columns:
        col_list = list(df[col])
        if col in round_two_list:
            col_list = handle_round_two(col_list)
        dic[col] = col_list
        
    return dic
