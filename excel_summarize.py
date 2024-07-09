# -*- coding: UTF-8 -*-

import os
import pprint
import sys

import openpyxl

in_path = "/Users/burt/Downloads/test/test"
out_filename = "out.xlsx"


def title_match(title_list, key):
    for title in title_list:
        if title is not None and key in str(title):
            return True
    return False


def read_excel(in_filename):
    book = openpyxl.load_workbook(in_filename, data_only=True)

    name_list = ["姓名", "身份证", "应付工资", "失业保险", "养老保险", "医疗保险", "请假工资", "个税", "实发工资",
                 "请假工资", "手机号码"]

    data_map_list = []
    for sheet in book.worksheets:
        print(in_filename, sheet.title, sheet.dimensions, sheet.max_column, sheet.max_row)
        for row in sheet.iter_rows():
            data_map = {}
            for cell in row[0: 40]:
                title_list = list(map(lambda x: sheet.cell(row=x, column=cell.column).value, range(2, 5)))
                if set(title_list) & set(name_list):
                    title = (set(title_list) & set(name_list)).pop()
                    data_map[title] = cell.value
                if title_match(title_list, "公积金"):
                    data_map["公积金"] = cell.value
                if title_match(title_list, "项目"):
                    data_map["项目"] = cell.value
                if title_match(title_list, "身份证"):
                    data_map["身份证"] = cell.value
                if title_match(title_list, "应付工资"):
                    data_map["应付工资"] = cell.value
            if data_map["身份证"] is not None and len(data_map["身份证"]) == 18:
                data_map_list.append(data_map)

    # pprint.pprint(data_map_list, width=400)
    return data_map_list


def save_excel(excel_data_list):
    out_wb = openpyxl.Workbook()
    # out_book = openpyxl.load_workbook(out_filename, data_only=True)
    out_ws = out_wb.worksheets[0]
    out_ws.append(["序号", "姓名", "身份证号码", "手机号码", "应付工资", "子女教育", "继续教育", "房贷利息", "住房租金",
                   "赡养老人", "大病医疗", "养老保险", "医疗保险", "失业保险", "公积金", "个税", "实发工资", "签名",
                   "项目"])
    start = 2
    for excel_data in excel_data_list:
        real = excel_data["应付工资"] if "请假工资" not in excel_data or excel_data["请假工资"] is None \
            else excel_data["应付工资"] - excel_data["请假工资"]
        out_ws.cell(row=start, column=2).value = excel_data["姓名"]
        out_ws.cell(row=start, column=3).value = excel_data["身份证"]
        if "手机号码" in excel_data:
            out_ws.cell(row=start, column=4).value = excel_data["手机号码"]
        out_ws.cell(row=start, column=5).value = real
        out_ws.cell(row=start, column=12).value = excel_data["养老保险"]
        out_ws.cell(row=start, column=13).value = excel_data["医疗保险"]
        out_ws.cell(row=start, column=14).value = excel_data["失业保险"]
        if "公积金" in excel_data:
            out_ws.cell(row=start, column=15).value = excel_data["公积金"]
        out_ws.cell(row=start, column=16).value = excel_data["个税"]
        out_ws.cell(row=start, column=17).value = excel_data["实发工资"]
        if "项目" in excel_data:
            out_ws.cell(row=start, column=19).value = excel_data["项目"]
        start += 1

    out_wb.save(out_filename)


if __name__ == '__main__':
    path = os.path.dirname(os.path.realpath(sys.executable))
    os.chdir(path)
    print(os.getcwd())

    files = list()
    for item in os.scandir():
        if (item.is_file() and item.name.endswith(".xlsx") and "~$" not in item.name
                and item.name != out_filename and ".~" not in item.name):
            files.append(item.path)
    files.sort()
    print(files)

    data_list = list()
    for file in files:
        data = read_excel(file)
        data_list += data

    save_excel(data_list)

    print("===== 完成 =====")
