import openpyxl
import json

from config.config import EXCEL_SHEET


def excel_handle(file_path):
    #加载工作簿
    workbook = openpyxl.load_workbook(file_path)

    #加载工作表
    worksheet = workbook[EXCEL_SHEET]

    #处理数据
    #zip() 函数可以将多个列表进行打包为元组，方便后续处理，例如：两个列表的值一一对应可以组合为字典
    data = []
    keys = [cell.value for cell in worksheet[2]]
    for row in worksheet.iter_rows(min_row=3, values_only=True):  # 遍历每一行数据

        case = dict(zip(keys, row))

        #处理数据，将str转化为dict
        for k in keys:
            if isinstance(case[k], str) and case[k].startswith('{'):
                case[k] = json.loads(case[k])
        # print(case["is_true"])
        # 判断是否执行用例
        if case["is_true"]:
            data.append(case)

    #关闭工作簿
    workbook.close()

    return data


if __name__ == "__main__":
    all_data = excel_handle("../data/测试用例.xlsx")
    for data in all_data:
        print(data)