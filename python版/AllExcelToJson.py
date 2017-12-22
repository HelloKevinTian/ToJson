# -*- coding: utf-8 -*-

import os, sys, platform
import json, xlrd, codecs
from collections import OrderedDict
from PythonTools import print_red
from PythonTools import print_green
from PythonTools import print_blue

# 配置路径
path = os.path.split(os.path.realpath(__file__))[0]
path_input = os.path.realpath(path + '/Table')
path_output = os.path.realpath(path + '/Json')
path_record = path_output + '/all_json.json'

#
lan_support = ['String_CN', 'String_EN']
compress = False

# 忽略列表
ignore_files = []
for line in open(path_input + '/_ignore.txt'):
    ignore_files.append(line.rstrip('\n'))

# 值类型转换
def to_value(value):
    return value

def to_none(value):
    return None

def to_bool(value):
    return bool(value)

def to_num(value):
    # print 'to_num: ' + str(value)
    return int(value)

def to_str(value):
    # print 'to_str: ' + str(value)
    # return str(value)
    return to_arr(value) if value[0] == '[' and value[-1] == ']' and value.count('[') == 1 and value.count(']') == 1 else value

def to_arr(value):
    return value[1:-1].split(',')

def get_value(cell):
    # print str(cell.ctype) + '   ' + str(xlrd.XL_CELL_NUMBER) + '  ' + str(xlrd.XL_CELL_TEXT) + ' ' + str(cell.value)
    return value_func[cell.ctype](cell.value)

value_func = {}
value_func[xlrd.XL_CELL_EMPTY] = to_none
value_func[xlrd.XL_CELL_TEXT] = to_str
value_func[xlrd.XL_CELL_NUMBER] = to_num
value_func[xlrd.XL_CELL_DATE] = to_value
value_func[xlrd.XL_CELL_BOOLEAN] = to_bool
value_func[xlrd.XL_CELL_ERROR] = to_value
value_func[xlrd.XL_CELL_BLANK] = to_value

# 检查是否有重复的健
def check_key(obj, key, excel):
    result = True

    if key == '' or key == 'None':
        print_red('[ERROR] ' + excel + ' 有为空的键')
        result = False
    if key in obj:
        print_red('[ERROR] ' + excel + ' 有重复为 ' + key + ' 的键')
        result = False

    return result

#by kevin
def get_val(type_row, index, cell):
    if type_row[index] == 'string' or type_row[index] == 'str':
        try:
            return str(int(cell.value))
        except ValueError:
            return cell.value
    elif type_row[index] == 'int' or type_row[index] == 'number':
        try:
            return int(cell.value)
        except ValueError:
            return (0 if cell.value == '' else cell.value)
    elif type_row[index] == 'float':
        return float(cell.value)
    elif type_row[index] == 'bool':
        return bool(cell.value)
    else:
        return cell.value

# 解析Excel表数据
def parse_excel(excel, entities):
    workbook = xlrd.open_workbook(excel)
    sheet = workbook.sheets()[0]
    # print 'total rows: ' + str(sheet.nrows)
    # print 'total cols: ' + str(sheet.ncols)

    attribute_row = [] #['item_type', 'item_id', 'item_number']
    for col in range(2, sheet.ncols):
        value = get_value(sheet.cell(0, col))
        result = check_key(attribute_row, str(value), excel)

        if result == True:
            attribute_row.append(str(value))
    # print (attribute_row)

    type_row = [] #['string', 'string', 'int']
    for col in range(2, sheet.ncols):
        value = get_value(sheet.cell(1, col))
        type_row.append(str(value))
    # print (type_row)

    attribute_col = [] # ['1', '2', '3']
    for row in range(3, sheet.nrows):
        value = get_value(sheet.cell(row, 1))
        result = check_key(attribute_col, str(value), excel)

        if result == True:
            attribute_col.append(str(value))
    # print (attribute_col)

    for row in range(3, sheet.nrows):
        entity = OrderedDict() #OrderedDict([('item_type', 1), ('item_id', 0), ('item_number', 100)])
        for col in range(2, sheet.ncols):
            index = col - 2
            if index < len(attribute_row):
                #by kevin
                entity[attribute_row[index]] = get_val(type_row, index, sheet.cell(row, col))
                #entity[attribute_row[index]] = get_value(sheet.cell(row, col))

        # print (entity)

        result = check_key(entities, attribute_col[row - 3], excel)
        if result == True:
            entities[attribute_col[row - 3]] = entity

    return entities

# 解析多语言表数据
def parse_localization(excel, entities):
    workbook = xlrd.open_workbook(excel)
    sheet = workbook.sheets()[0]

    lans = []
    for col in range(2, sheet.ncols):
        value = get_value(sheet.cell(0, col))
        check_key(lans, str(value), excel)
        lans.append(str(value))

    keys = []
    for row in range(3, sheet.nrows):
        value = get_value(sheet.cell(row, 1))
        check_key(keys, str(value), excel)
        keys.append(str(value))

    for col in range(2, sheet.ncols):
        lan = lans[col - 2]
        for row in range(3, sheet.nrows):
            if not lan in entities:
                entities[lan] = OrderedDict()

            key = keys[row - 3]
            check_key(entities[lan], key, excel)
            entities[lan][key] = get_value(sheet.cell(row, col))

    return entities

# 解析key:value表结构
def parse_one_to_one(excel, entities):
    workbook = xlrd.open_workbook(excel)
    sheet = workbook.sheets()[0]

    keys = []
    for row in range(2, sheet.nrows):
        key = get_value(sheet.cell(row, 0))
        value = get_value(sheet.cell(row, 1))
        entities[key] = value

    return entities

# 写入Json
def write_json(dir, entities, separator):
    if not os.path.exists(os.path.dirname(dir)):
        os.mkdir(os.path.dirname(dir))
    filed = codecs.open(dir, "w", "utf-8")
    if separator:
        filed.write("%s" % json.dumps(entities, ensure_ascii=False, separators=(',',':')))
    else:
        filed.write("%s" % json.dumps(entities, ensure_ascii=False, indent=4))
    filed.close()
    print_green('>>>>> ' + os.path.realpath(dir))

# 筛选合法文件
def filter_files(files, extname):
    filter = []
    for file in files:
        name, ext = os.path.splitext(file)
        if ext == extname and name.find('~$') == -1 and not file in ignore_files:
            filter.append(file) 
    return filter

# 遍历文件夹
def export_all():
    for root, dirs, files in os.walk(path_input):
        if root == path_input:
            for xlsx in filter_files(files, '.xlsx'):
                name, ext = os.path.splitext(xlsx)
                if name == 'error_code':
                    entities = parse_one_to_one(os.path.join(root, xlsx), OrderedDict())
                else:
                    entities = parse_excel(os.path.join(root, xlsx), OrderedDict())
                write_json(path_output + '/' + name + '.json', entities, compress)
        elif os.path.basename(root) == 'Localization':
            entities = OrderedDict()
            exportdir = path_output + root[len(path_input) : ]
            for xlsx in filter_files(files, '.xlsx'):
                entities = parse_localization(os.path.join(root, xlsx), entities)
            for language in entities:
                if language in lan_support:
                    write_json(exportdir + '/' + language + '.json', entities[language], compress)
        else:
            entities = OrderedDict()
            exportdir = path_output + root[len(path_input) : ]
            for xlsx in filter_files(files, '.xlsx'):
                entities = parse_excel(os.path.join(root, xlsx), entities)
            if len(filter_files(files, '.xlsx')) > 0:
                write_json(exportdir + '.json', entities, compress)

# 统计所有的Json
def record_all():
    if os.path.exists(path_record):
        os.remove(path_record)

    entities = []
    for root, dirs, files in os.walk(path_output):
        for config in filter_files(files, '.json'):
            configPath = os.path.join(root, config)
            relativeKey = path_output + '/'
            relativeDir = os.path.realpath(path_output + '/..') + '/'

            entity = OrderedDict()
            entity['key'] = configPath[len(relativeKey): ].replace('\\', '/')
            entity['path'] = configPath[len(relativeDir): ].replace('\\', '/')
            entities.append(entity)
    write_json(path_record, entities, compress)

# argv
if len(sys.argv) > 1 and sys.argv[1] == '-compress':
    compress = True
else:
    print_blue('use param -compress to compress json')
print 'Ignore Files: ' + str(ignore_files)

# Main
export_all()
# record_all()
print_green('[Finish]')