# -*- coding: utf-8 -*-
"""
Created on Wed Mar 22 16:26:34 2023

@author: Jinliang
"""

"""
word文档中的内容
    paragraph(段落)
    table(表格)
    character(字符)

目标
    从word中分析表格，并将表格信息结构化
"""

from docx import Document

def load_word(word_path):
    document = Document(word_path) ## 存储在document对象中
    
    return document

    ## 表格读取
    tables = document.tables
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                print(cell.text)
    
def save_word(document, save_path):
    document.save(save_path)
    




