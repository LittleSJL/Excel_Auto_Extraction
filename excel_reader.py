# -*- coding: utf-8 -*-
"""
Created on Wed Mar 22 17:02:41 2023

@author: Jinliang

用途：此reader用于读取excel中的内容信息，需要做好一切的预处理工作
"""
import math
import pandas as pd

# 需要保留两位有效数字的列数据
round_two_list = ["应收权益1-本利和", "房屋总价2", "抵债金额（总）", "剩余购房款（不含首付）",
                  "对应首付金额（元）", "乙方1产品1剩余本金", "乙方1产品1剩余收益",
                  "乙方1产品1转让本金", "乙方1产品1转让收益", '房源建面']

def isNan(item):
    """
    判断一个对象是否为nan
        nan一定是float，如果不是float，则一定不是nan
        float的情况下，接着使用isnan判断是否为nan
    """
    if type(item) == float and math.isnan(item):
        return True
    return False

def preprocess_round_two(l):
    """
    处理float的金额数字
        一个list中的float，全部保留2位有效数字（返回str格式）
        针对所有的float，nan转换为0.0
    """
    new_l = []
    for num in l:
        if isNan(num):
            num = 0.0
        new_l.append(('%.2f'%num)) # 强转float时，会丢失.00这种2位小数，只能存成str，后期要运算再变回来
    return new_l
    
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
            col_list = preprocess_round_two(col_list)
        dic[col] = col_list
    
    dic = generate_rest_money(dic)
    return dic
    
    
def generate_rest_money(dic):
    """
    生成【剩余应付房款】
        【剩余应付房款】 = 【对应首付金额（元）】+【剩余购房款（不含首付）】
    """
    value_list = []  # 记录剩余应付房款
    
    total_num = len(dic.get('乙方1'))
    for number in range(total_num):
        # 无内容的记录，直接跳过
        name = dic.get('乙方1')[number]
        if isNan(name):
            value_list.append(math.nan)
        else:
            item1 = dic.get('对应首付金额（元）')[number]
            if isNan(float(item1)):
                item1 = 0.0
            item2 = dic.get('剩余购房款（不含首付）')[number]
            if isNan(float(item2)):
                item2 = 0.0
            value = float(item1) + float(item2)
            value_list.append(value)
    
    value_list = preprocess_round_two(value_list)
    dic['剩余应付房款'] = value_list
     
    return dic