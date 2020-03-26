import xlrd
import copy
import sys

# 配置信息
# 平台文件路径
path=r"C:\Users\admin\Downloads\包.xlsx"
# 零部件名
target=r'BAFFLE ASM-RAD AIR UPR'
# 平台中表名，注意当name设置为None时，为全文件查找。
name=None
# JCCC文件路径
path_from=r"C:\Users\admin\Downloads\JC.xls"
# 记录文件路径
LOG_PATH=r"C:\log.txt"



def target_index(sheet, target, target_num, max_rows):
    ret = []
    for i in range(max_rows):
        if target == sheet.cell_value(i, target_num):
            ret.append(i)
    return ret


def value_value(sheet, value_num, idx_list):
    values = []
    for i in idx_list:
        v = sheet.cell_value(i, value_num)
        values.append(v)
    return values


def get_max_two(values):
    prob = []
    v_set_list = list(set(values))
    for i in v_set_list:
        prob.append(values.count(i))
    print(v_set_list)
    print(prob)
    bak_prob = copy.deepcopy(prob)
    max1 = max(bak_prob)
    if len(prob) > 1:
        bak_prob.remove(max1)
        max2 = max(bak_prob)
        return v_set_list[prob.index(max1)], v_set_list[prob.index(max2)], max1, max2
    else:
        return v_set_list[prob.index(max1)], v_set_list[prob.index(max1)], max1, max1


def in_one_sheet(sheet, target):
    target_num = sheet.row_values(0).index('FNA DESC.')
    if not target_num:
        return None
    value_num = sheet.row_values(0).index('包装顺序')
    max_rows = sheet.nrows
    t_idx_list = target_index(sheet, target, target_num, max_rows)
    value = value_value(sheet, value_num, t_idx_list)
    return value

def loop_sheet(workbook,target):
    names = workbook.sheet_names()
    values=[]
    for name in names:
        sheet=workbook.sheet_by_name(name)
        v=in_one_sheet(sheet,target)
        if v:
            values+=v
    return values

def get_fna_names(path_from):
    ret=[]
    workbook_from = xlrd.open_workbook(path_from)
    sheet = workbook_from.sheet_by_name('maintainWmsStandardBomDetailPag')
    target_num = sheet.row_values(0).index('FNA DESC.')
    if not target_num:
        return None
    max_rows = sheet.nrows
    for i in range(max_rows):
        ret.append(sheet.cell_value(i,target_num))
    return ret

def main(path_from,path,name,target):
    fna_names=get_fna_names(path_from)

    workbook = xlrd.open_workbook(path)

    if name:
        sheet = workbook.sheet_by_name(name)
        values = in_one_sheet(sheet,target)
        print("In Specified Platform:"+str(values))
        if not values:
            values = loop_sheet(workbook,target)
        end=get_max_two(values)
        print("=====Result======")
        print(end)
        print("=====Result======")
    else:
        for target in fna_names:
            values=loop_sheet(workbook,target)
            if not values:
                with open(LOG_PATH, 'a+') as f:
                    f.write("Not Found,Not Found,\n")
            else:
                end=get_max_two(values)
                with open(LOG_PATH,'a+') as f:
                    f.write("{},{},\n".format(end[0],end[1]))
# 主函数
main(path_from,path,name,target)
