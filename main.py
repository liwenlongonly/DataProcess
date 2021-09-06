# -*- coding: utf-8 -*-
# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
import pandas as pd
import sys
import os
import time
import pathlib

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    argv = sys.argv[1:]
    filterStr = "米色"
    data1 = None
    data2 = None
    outputPath = "output"

    for item in argv:
        if item.endswith(".csv"):
            data1 = pd.read_csv(item, encoding='gbk',
                                usecols=["订单编号", "购买数量", "商品属性"])
            data1["订单编号"] = data1["订单编号"].str.replace('=', '')
            data1["订单编号"] = data1["订单编号"].str.replace('"', '')
        elif item.endswith(".xlsx"):
            data2 = pd.read_excel(item, usecols=["订单编号", "收货人姓名", "联系手机", "收货地址 "])
            data2["订单编号"] = data2["订单编号"].astype(str)
            # 过滤收件人是null的行
            data2 = data2[data2['收货人姓名'].notnull()]
        else:
            filterStr = item

    if data1 is not None and data2 is not None:
        # 根据订单号合并表格
        data = pd.merge(data2, data1, on="订单编号")

        # 根据过滤条件过滤
        data3 = data[data['商品属性'].str.contains(filterStr)]
        print(data3)

        # 判断当前文件有没有输出文件夹，没有创建一个
        absolutePath = pathlib.Path(__file__).parent.absolute()
        path = pathlib.Path("{}/{}".format(absolutePath, outputPath))
        if not path.exists():
            path.mkdir()
        # 数据写入文件
        now = time.strftime("%Y-%m-%d %H:%M:%S")
        filePath = "{}".format(path.absolute()) + '/ExportOrderList_{}'.format(now) + ".xlsx"
        data3.to_excel(filePath, index=False, encoding='gbk', engine='openpyxl')
        print("请查看文件:{}".format(filePath))
        os.system("open {}".format(path.absolute()))
        print("process finish, good luck!")
    else:
        print("参数输入错误, 请检查参数！")
