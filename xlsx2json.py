#coding=utf-8

import os, re
import json
import datetime
from openpyxl import load_workbook

'''
    数组只有一个值时，以逗号（半角）结尾
    > 2,
    对象
    > name:dong, age:18
'''

def parse_cell_value(value):
    if not value:
        return ''
    #int类型
    if isinstance(value, int):
        return value
    #float类型
    if isinstance(value, float):
        return value
    #日期类型
    if isinstance(value, datetime.datetime):
        return value.ctime()
    value = value.replace(' ', '')
    value_list = value.split(',')
    #对象类型
    if value.find(':') != -1:
        obj = {}
        for i in value_list:
            i = i.split(':')
            obj[i[0]] = parse_cell_value(i[1])
        return obj
    #数组类型
    if len(value_list) > 1:
        parsed_list = []
        for i in value_list:
            if i:
                i = parse_cell_value(i)
                parsed_list.append(i)
        return parsed_list
    #布尔类型
    if re.match('true', value, re.IGNORECASE):
        return True
    if re.match('false', value, re.IGNORECASE):
        return False
    #递归判断
    if re.match(r'^\d+(\.\d+)?$', value):
        if value.find('.') != -1:
            return float(value)
        return int(value)
    #string类型
    return value
    
def get_workbooks(dir_name = 'excel'):
    excel_path = os.path.join(os.getcwd(), dir_name)
#     print os.listdir(excel_path)
    file_list = os.listdir(excel_path)
    if not file_list:
        print 'no excel file !'
        return
    workbooks = []
    for i in file_list:
        file_path = os.path.join(excel_path, i)
#         print file_path
        wb = load_workbook(file_path)
        workbooks.append(wb)
    return workbooks

def save_json(file_name, json_data, dir_name = 'json'):
    json_path = os.path.join(os.getcwd(), dir_name)
    file_path = os.path.join(json_path, file_name + '.json')
    with open(file_path, 'w') as f:
        json.dump(json_data, f, indent = 4)

def xlsx2json(head_row = 2):
    workbooks = get_workbooks()
    for wb in workbooks:
        for sheet in wb:
#             print sheet.title
            head = sheet.rows[head_row - 1]
#             print head
            json_list = []
            for row in sheet.rows[head_row:]:
                row_dic = {}
                for head_cell, cell in zip(head, row):
#                     print head_cell.value, cell.value, type(cell.value)
                    row_dic[head_cell.value] = parse_cell_value(cell.value)
                json_list.append(row_dic)
            save_json(sheet.title, json_list)

if __name__ == '__main__':
    xlsx2json()