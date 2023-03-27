# -*- coding: utf-8 -*-
"""
Created on Sun Mar 26 23:00:48 2023

@author: Jinliang
"""
import argparse
from auto_extraction import replace_value_in_str
from excel_reader import load_excel, isNan
import pandas as pd

def load_excel_mode(excel_mode_path):
    df = pd.read_excel(excel_mode_path)
    return df

def extract_excel_to_confirmation(excel_mode_path, excel_path):
    excel_dic = load_excel(excel_path)
    total_num = len(excel_dic.get('乙方1'))
    count = 0
    for number in range(total_num):
        # 无内容的记录，直接跳过
        name = excel_dic.get('乙方1')[number]
        if isNan(name):
            continue
        count += 1
        
        # 添加模板
        excel_mode = load_excel_mode(excel_mode_path)
        
        new_item_list = []
        for item in list(excel_mode['客户信息表']):
            if isNan(item) or '##' not in item:
                new_item_list.append(item)
            else:
                new_para = replace_value_in_str(item, excel_dic, number)
                # 这个Excel有mode的说法吗...
                new_item_list.append(new_para)
        
        excel_mode['客户信息表'] = new_item_list
        
        ## 张爱军-客户信息填写确认单碧海云天-40栋-402（excel_mode）
        ## 乙方+项目名+房号
        name = excel_dic.get('乙方1')[number]
        project = excel_dic.get('项目名称')[number]
        room_number = excel_dic.get('房号')[number]
        file_name = name + '-客户信息填写确认单-' + project + '-' + room_number
        save_path = 'output/excel_confirmation/' + file_name + '.xlsx'
        
        excel_mode.to_excel(save_path, index=False)
        
    print('共生成确认单文件数目：', count)

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('excel_path', type=str) # 需要提取的excel文件
    args = parser.parse_args()
    
    print("从【" + args.excel_path + "】中抽取信息，自动填入excel模板中...")
    
    excel_mode_path = 'data/excel_mode/excel_mode.xlsx'
    extract_excel_to_confirmation(excel_mode_path, args.excel_path)
    
    print("全部记录生成完毕，结果已写入output/excel_confirmation文件夹")
            



