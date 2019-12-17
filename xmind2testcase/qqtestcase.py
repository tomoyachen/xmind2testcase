#!/usr/bin/env python
# _*_ coding:utf-8 _*_

import xlsxwriter
import logging
import os
from xmind2testcase.utils import get_xmind_testcase_list2, get_absolute_path


"""
Convert XMind fie to qq testcase excel file 
"""


def xmind_to_qqtestcase_file(xmind_file):
    """Convert XMind file to a qqtestcase file"""
    xmind_file = get_absolute_path(xmind_file)
    logging.info('Start converting XMind file(%s) to qqtestcase file...', xmind_file)
    testsutie_dict = get_xmind_testcase_list2(xmind_file)



    fileheader = ["编号", "功能模块", "测试点", "前置条件", "操作步骤", "预期结果"]

    qqtestcase_file = xmind_file[:-6] + '_qq' + '.xlsx'

    if os.path.exists(qqtestcase_file):
        logging.info('The qqtestcase file already exists, return it directly: %s', qqtestcase_file)
        return qqtestcase_file

    workbook = xlsxwriter.Workbook(qqtestcase_file)

    # 自动换行
    style_text_wrap = workbook.add_format({'text_wrap': 1, 'valign':'top'})
    #
    # sheet1 = workbook.add_worksheet('README')
    # sheet1.write(0, 0, '测试用例内容请至第二页查看')  # 第0行第0列写入内容
    # sheet1.write(1, 0, '确认数量正确、内容正确后，可将此文件直接导入iWork系统', style1)  # 第1行第0列写入内容

    smoke_case_dict = {}
    for product in testsutie_dict:

        smoke_case_dict[product] = []

        sheet = workbook.add_worksheet(product) #sheet名
        sheet.set_column("A:B", 20)
        sheet.set_column("C:F", 30)

        # 用例title
        sheet.write(0, 0, fileheader[0])
        sheet.write(0, 1, fileheader[1])
        sheet.write(0, 2, fileheader[2])
        sheet.write(0, 3, fileheader[3])
        sheet.write(0, 4, fileheader[4])
        sheet.write(0, 5, fileheader[5])

        #第二行开始写入用例
        case_index = 1
        case_no = 0
        for testcase in testsutie_dict[product]:
            row_dict = gen_a_testcase_row_dict(testcase) #包含用例信息
            row_list = row_dict["case_row_list"] #取用例列
            # print("row_dict", row_dict)

            #####################
            smoke_case = {}
            if row_dict["case_priority"] == "高":
                smoke_case["module"] = row_dict["case_module"]
                smoke_case["name"] = row_dict["case_title"]
                smoke_case["precontion"] = row_dict["case_precontion"]
                smoke_case["case"] = []

            ##########################

            for row in row_list:
                if len(row[1]) > 0:
                    case_no += 1
                    sheet.write(case_index, 0, "No." + str(case_no), style_text_wrap)
                else:
                    sheet.write(case_index, 0, "", style_text_wrap)

                sheet.write(case_index, 1, row[0], style_text_wrap)
                sheet.write(case_index, 2, row[1], style_text_wrap)
                sheet.write(case_index, 3, row[2], style_text_wrap)
                sheet.write(case_index, 4, row[3], style_text_wrap)
                sheet.write(case_index, 5, row[4], style_text_wrap)
                case_index = case_index + 1

                #给用例加操作步骤 预期结果
                if len(smoke_case) > 0:
                    smoke_case["case"].append([row[3], row[4]])

            if len(smoke_case) > 0:
                smoke_case_dict[product].append(smoke_case)
    print(smoke_case_dict)

    #写入冒烟测试
    #############
    if len(smoke_case_dict) > 0:
        sheet2 = workbook.add_worksheet("冒烟用例")  # sheet名
        sheet2.set_column("A:A", 20)
        sheet2.set_column("B:B", 80)

        _case_index = 0
        #分解sheet
        for product in smoke_case_dict:
            if len(smoke_case_dict[product]) > 0:
                # 用例title
                sheet2.write(_case_index, 0, "编号", style_text_wrap)
                sheet2.write(_case_index, 1, "%s 冒烟用例" % product, style_text_wrap)
                _case_index += 1
                _smoke_case_list = smoke_case_dict[product]
                _case_no = 0
                for _smoke_case in _smoke_case_list:
                    _smoke_case_string = "测试点：\n"
                    _smoke_case_string += _smoke_case["name"] + "\n\n"
                    # print("_smoke_case", _smoke_case)
                    #前置条件
                    if len(_smoke_case["precontion"]) > 0 and _smoke_case["precontion"] != '无':
                        _smoke_case_string += "前置条件：\n"
                        _smoke_case_string += _smoke_case["precontion"] + "\n\n"
                    _step_index = 0
                    for _step in _smoke_case["case"]:
                        if len(_step[0]) > 0 and len(_step[1]):
                            _step_index += 1
                            _smoke_case_string += "操作步骤" + str(_step_index) + "：\n"
                            _smoke_case_string += _step[0] + "\n\n"
                            _smoke_case_string += "预期结果" + str(_step_index) + "：\n"
                            _smoke_case_string += _step[1] + "\n\n"
                        elif len(_step[0]) > 0:
                            _step_index += 1
                            _smoke_case_string += "操作步骤" + str(_step_index) + "：\n"
                            _smoke_case_string += _step[0] + "\n\n"

                    # print(_smoke_case_string)
                    _case_no += 1
                    sheet2.write(_case_index, 0, "No." + str(_case_no), style_text_wrap)
                    sheet2.write(_case_index, 1, _smoke_case_string, style_text_wrap)
                    _case_index += 1
                _case_index += 1


    ###################
    workbook.close()
    logging.info('Convert XMind file(%s) to a qqtestcase file(%s) successfully!', xmind_file, qqtestcase_file)

    return qqtestcase_file

#qqtestcase
def gen_a_testcase_row_dict(testcase_dict):
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

    #用例信息
    row_list_dict = {
        "case_module": case_module,
        "case_title": case_title,
        "case_precontion": case_precontion,
        "case_priority": case_priority,
        "case_row_list": []

    }

    # 列内容
    row_list = []
    row = ""
    row_index = 1
    if not case_step_and_expected_result_dict:
        row = [case_module, case_title, case_precontion, "", ""]
        row_list.append(row)
    # print("aaa", case_step_and_expected_result_dict.items())
    else:
        for step, expected in case_step_and_expected_result_dict.items():
            # 是否首行
            if row_index > 1:
                case_module = ""
                case_title = ""
                case_depict = ""
                case_precontion = ""
            else:
                pass
                # case_title = "[" + case_module + "]" + " " + case_title

            #拼接
            if step and expected:
                row = [case_module, case_title, case_precontion, step, expected] #预期结果一对多来自parser.py文件
            elif step:
                row = [case_module, case_title, case_precontion, step, ""]
            else:
                row = [case_module, case_title, case_precontion, "", ""]
            row_list.append(row)
            row_index = row_index + 1

    row_list_dict["case_row_list"].extend(row_list)
    return row_list_dict


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
    qqtestcase_file = xmind_to_qqtestcase_file(xmind_file)
    print('Conver the xmind file to a qqtestcase file succssfully: %s', qqtestcase_file)