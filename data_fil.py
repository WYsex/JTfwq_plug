import xlrd,os
from collections import Counter
import logging


def Excel_p(bale_path, imp_path):
    # print(imp_path)
    bale_path = bale_path
    imp_path = imp_path  # 导入信息路径
    outer_pack = xlrd.open_workbook(bale_path)
    imp_t = xlrd.open_workbook(imp_path)
    out_sheet = outer_pack.sheets()[0]  # 读取外包装excel sheet页数，从0开始
    sheets = imp_t.sheet_names()
    print(sheets)

    for index, page in enumerate(sheets):  # 遍历当前excel的所有sheet页
        print(page)
        try:
            imp_sheet = imp_t.sheets()[index]  # 读取要导入信息的excel sheet页数，从0开始
            out_subno = out_sheet.col_values(0)[1:]  # 读取商品条码，跳过标题
            imp_subno = imp_sheet.col_values(0)[1:]
            Imp_subno = [str(int(sub)) for sub in imp_subno]  # 对导入的信息进行取整 并转换为文本类型
            a_lis = []
            [a_lis.append(i) for i in out_subno if i in Imp_subno]
            print("大小包装:", a_lis)
            imp_dic = dict(Counter(imp_subno))
            dup_value = {key: value for key, value in imp_dic.items() if value > 1}
            print("重复条码:", dup_value)
        except:
            print("当前表格有问题，请检查!")


def whil_Pro():
    try:
        bale_path = r'F:\河南交投项目\服务区导入信息\商品外包装信息.xlsx'  # 填写实际大小包装路径
        abs_path = r'F:\河南交投项目\期初库存\未整理好得'
        imp_path1 = input("请输入文件名:")
        imp_path=os.path.join(abs_path,imp_path1)
        # print(imp_path)
        Excel_p(bale_path=bale_path, imp_path=imp_path)
    except:
        print("请检查文件路径或文件格式")


if __name__ == '__main__':
    while True:
        print("开始按0,退出按1")
        log = int(input("请输入数字:"))
        if log == 0:
            whil_Pro()
        else:
            print("程序即将结束!!!")
            break
