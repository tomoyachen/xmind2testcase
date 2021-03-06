#!/usr/bin/env python
# _*_ coding:utf-8 _*_

import xlsxwriter
import logging
import os
from xmind2testcase.utils import get_xmind_testcase_list, get_absolute_path


"""
Convert XMind fie to iwork testcase excel file 
"""


def xmind_to_iwork_excel_file(xmind_file):
    """Convert XMind file to a iwork excel file"""
    xmind_file = get_absolute_path(xmind_file)
    logging.info('Start converting XMind file(%s) to iwork file...', xmind_file)
    testcases = get_xmind_testcase_list(xmind_file)

    fileheader = ["用例概要*", "用例描述", "测试步骤", "测试数据", "预期结果"]

    iwork_file = xmind_file[:-6] + '_iwork' + '.xlsx'

    if os.path.exists(iwork_file):
        logging.info('The eiwork excel file already exists, return it directly: %s', iwork_file)
        return iwork_file


    workbook = xlsxwriter.Workbook(iwork_file)

    # 红色字体
    style1 = workbook.add_format({'font_color':'#FF0000'})

    # 自动换行
    style2 = workbook.add_format({'text_wrap': 1, 'valign':'top'})

    sheet1 = workbook.add_worksheet('README')
    sheet1.write(0, 0, '测试用例内容请至第二页查看')  # 第0行第0列写入内容
    sheet1.write(1, 0, '确认数量正确、内容正确后，可将此文件直接导入iWork系统', style1)  # 第1行第0列写入内容

    sheet2 = workbook.add_worksheet('测试用例')
    sheet2.set_column("A:E", 30)


    # 用例title
    sheet2.write(0, 0, fileheader[0])
    sheet2.write(0, 1, fileheader[1])
    sheet2.write(0, 2, fileheader[2])
    sheet2.write(0, 3, fileheader[3])
    sheet2.write(0, 4, fileheader[4])

    #第二行开始写入用例
    case_index = 1
    for testcase in testcases:
        # row = gen_a_testcase_row(testcase)
        row_list = gen_a_testcase_row_list(testcase)
        # print("row_list >> ", row_list)
        for row in row_list:
            # for i in range(0,len(row)):
            #     sheet2.write(case_index, i, row[i])
            sheet2.write(case_index, 0, row[0], style2)
            sheet2.write(case_index, 1, row[1], style2)
            sheet2.write(case_index, 2, row[2], style2)
            sheet2.write(case_index, 3, row[3], style2)
            sheet2.write(case_index, 4, row[4], style2)
            # sheet2.set_row(case_index, len(row_list)*10)
            case_index = case_index + 1


    workbook.close()
    logging.info('Convert XMind file(%s) to a iwork excel file(%s) successfully!', xmind_file, iwork_file)

    return iwork_file


def gen_a_testcase_row(testcase_dict):
    case_module = gen_case_module(testcase_dict['suite'])
    case_title = testcase_dict['name']
    case_precontion = testcase_dict['preconditions']
    case_step, case_expected_result = gen_case_step_and_expected_result(testcase_dict['steps'])
    case_keyword = '功能测试'
    case_priority = gen_case_priority(testcase_dict['importance'])
    case_type = gen_case_type(testcase_dict['execution_type'])
    case_apply_phase = '迭代测试'
    row = [case_module, case_title, case_precontion, case_step, case_expected_result, case_keyword, case_priority, case_type, case_apply_phase]
    return row

#iwork
def gen_a_testcase_row_list(testcase_dict):
    case_module = gen_case_module(testcase_dict['suite'])
    case_title = testcase_dict['name']

    case_precontion = testcase_dict['preconditions']
    case_step_and_expected_result_dict = gen_case_step_and_expected_result_dict(testcase_dict['steps'])
    # print("case_step_and_expected_result_dict >> ", case_step_and_expected_result_dict)
    case_keyword = '功能测试'
    case_priority = gen_case_priority(testcase_dict['importance'])
    case_type = gen_case_type(testcase_dict['execution_type'])
    case_apply_phase = '迭代测试'

    #用例描述
    case_depict = ""
    case_depict += "所属模块: " + case_module + "\n"
    case_depict += "前置条件: " + case_precontion + "\n"
    case_depict += "关键词: " + case_keyword + "\n"
    case_depict += "优先级: " + case_priority + "\n"
    case_depict += "用例类型: " + case_type + "\n"
    case_depict += "适用阶段: " + case_apply_phase + "\n"

    # 列内容
    row_list = []
    row = ""
    row_index = 1
    if not case_step_and_expected_result_dict:
        row = [case_title, case_depict, "", "", ""]
        row_list.append(row)
        return row_list
    # print("aaa", case_step_and_expected_result_dict.items())
    for step, expected in case_step_and_expected_result_dict.items():
        # 是否首行
        if row_index > 1:
            case_title = ""
            case_depict = ""
        else:
            case_title = "[" + case_module + "]" + " " + case_title

        #拼接
        if step and expected:
            row = [case_title, case_depict, step, "", expected]
        elif step:
            row = [case_title, case_depict, step, "", ""]
        else:
            row = [case_title, case_depict, "", "", ""]
        row_list.append(row)
        row_index = row_index + 1

    # print("row_list by function >>", row_list)
    return row_list


def gen_case_module(module_name):
    if module_name:
        module_name = module_name.replace('（', '(')
        module_name = module_name.replace('）', ')')
    else:
        module_name = '/'
    return module_name


def gen_case_step_and_expected_result(steps):
    case_step = ''
    case_expected_result = ''

    for step_dict in steps:
        case_step += str(step_dict['step_number']) + '. ' + step_dict['actions'].replace('\n', '').strip() + '\n'
        case_expected_result += str(step_dict['step_number']) + '. ' + \
            step_dict['expectedresults'].replace('\n', '').strip() + '\n' \
            if step_dict.get('expectedresults', '') else ''

    return case_step, case_expected_result

#步骤与预期结果的字典
def gen_case_step_and_expected_result_dict(steps):
    total_dict = {}
    for step_dict in steps:
        # print(step_dict)
        # total_dict[step_dict['actions'].replace('\n', '').strip()] = step_dict['expectedresults'].replace('\n', '').strip()
        # 因为更换成预期结果也有值所以用下面的语句
        total_dict[step_dict['actions'].replace('\n', '').strip()] = step_dict['expectedresults'].strip()
    return total_dict



def gen_case_priority(priority):
    mapping = {1: '高', 2: '中', 3: '低'}
    if priority in mapping.keys():
        return mapping[priority]
    else:
        return '中'


def gen_case_type(case_type):
    mapping = {1: '手动', 2: '自动'}
    if case_type in mapping.keys():
        return mapping[case_type]
    else:
        return '手动'


if __name__ == '__main__':
    xmind_file = '../docs/zentao_testcase_template.xmind'
    iwork_excel_file = xmind_to_iwork_excel_file(xmind_file)
    print('Conver the xmind file to a iwork excel file succssfully: %s', iwork_excel_file)