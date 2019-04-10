#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import csv
import logging
import os
from xmind2testcase.utils import get_xmind_testcase_list, get_absolute_path

"""
Convert XMind fie to iwork testcase csv file 
"""


def xmind_to_iwork_csv_file(xmind_file):
    """Convert XMind file to a iwork csv file"""
    xmind_file = get_absolute_path(xmind_file)
    logging.info('Start converting XMind file(%s) to iwork file...', xmind_file)
    testcases = get_xmind_testcase_list(xmind_file)

    fileheader = ["用例概要*", "用例描述", "测试步骤", "测试数据", "预期结果"]
    iwork_testcase_rows = [fileheader]
    for testcase in testcases:
        # row = gen_a_testcase_row(testcase)
        row_list = gen_a_testcase_row_list(testcase)
        # print("row_list >> ", row_list)
        for row in row_list:
            iwork_testcase_rows.append(row)

    iwork_file = xmind_file[:-6] + '_iwork' + '.csv'
    if os.path.exists(iwork_file):
        logging.info('The eiwork csv file already exists, return it directly: %s', iwork_file)
        return iwork_file

    with open(iwork_file, 'w', encoding='gbk', newline="") as f:
        writer = csv.writer(f)
        writer.writerows(iwork_testcase_rows)
        logging.info('Convert XMind file(%s) to a iwork csv file(%s) successfully!', xmind_file, iwork_file)

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
        print(step_dict)
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
    iwork_csv_file = xmind_to_iwork_csv_file(xmind_file)
    print('Conver the xmind file to a iwork csv file succssfully: %s', iwork_csv_file)