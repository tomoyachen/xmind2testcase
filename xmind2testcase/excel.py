#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import xlwt
import logging
import os
from xmind2testcase.utils import get_xmind_testcase_list, get_absolute_path

"""
Convert XMind fie to Zentao testcase csv file 

Zentao official document about import CSV testcase file: https://www.zentao.net/book/zentaopmshelp/243.mhtml 
"""


def xmind_to_excel_file(xmind_file):
    """Convert XMind file to a excel csv file"""
    xmind_file = get_absolute_path(xmind_file)
    logging.info('Start converting XMind file(%s) to excel file...', xmind_file)
    testcases = get_xmind_testcase_list(xmind_file)

    fileheader = ["所属模块", "用例标题", "前置条件", "步骤", "预期", "关键词", "优先级", "用例类型", "适用阶段"]


    wbk = xlwt.Workbook()
    sheet1 = wbk.add_sheet('测试用例', cell_overwrite_ok=False)

    # 自动换行
    style1 = xlwt.easyxf('align: wrap on, vert top')
    sheet1.col(0).width = 256*30
    sheet1.col(1).width = 256*40
    sheet1.col(2).width = 256*30
    sheet1.col(3).width = 256*40
    sheet1.col(4).width = 256*40

    # 用例title
    for i in range(0, len(fileheader)):
        sheet1.write(0, i, fileheader[i])

    #第二行开始写入用例
    case_index = 1
    for testcase in testcases:
        # row = gen_a_testcase_row(testcase)
        row = gen_a_testcase_row(testcase)
        # print("row_list >> ", row_list)
        for i in range(0,len(row)):
            sheet1.write(case_index, i, row[i], style1)
        case_index = case_index + 1

    excel_file = xmind_file[:-5] + 'xls'
    if os.path.exists(excel_file):
        logging.info('The excel file already exists, return it directly: %s', excel_file)
        return excel_file

    if excel_file:
        wbk.save(excel_file)
        logging.info('Convert XMind file(%s) to a iwork excel file(%s) successfully!', xmind_file, excel_file)

    return excel_file


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
            step_dict['expectedresults'].replace('', '').strip() + '\n' \
            if step_dict.get('expectedresults', '') else ''

    return case_step, case_expected_result


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
    excel_file = xmind_to_excel_file(xmind_file)
    print('Conver the xmind file to a excel csv file succssfully: %s', excel_file)