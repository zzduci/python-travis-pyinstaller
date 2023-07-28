#你的python代码
# -*- coding:utf-8 -*-
import pandas as pd
from math import ceil
import os


def account(adress, weight):
    if adress == "湖北省":
        if weight <= 3:
            totel = 4
        elif (weight >= 3) and (weight <= 30):
            totel = 4+ ceil((weight - 1)) * 1
        else:
            totel = ceil(weight) * 1
        return totel

    elif adress in ["江苏省", "浙江省", "河南省", "江西省", "湖南省", "安徽省"]:
        if weight <= 3:
            totel = 4
        elif (weight >= 3) and (weight <= 30):
            totel = 4.5 + ceil((weight - 1)) * 1
        else:
            totel = ceil(weight) * 2
        return totel

    elif adress in ["广东省"]:
        if weight <= 3:
            totel = 4
        elif (weight >= 3) and (weight <= 30):
            totel = 4.5 + ceil((weight - 1)) * 1.30
        else:
            totel = ceil(weight) * 2
        return totel

    elif adress in ["上海"]:
        if weight <= 3:
            totel = 4
        elif (weight >= 3) and (weight <= 30):
            totel = 4.5 + ceil((weight - 1)) * 1.5
        else:
            totel = ceil(weight) * 3
        return totel

    elif adress in ["北京"]:
        if weight <= 3:
            totel = 4
        elif (weight >= 3) and (weight <= 30):
            totel = 4.5 + ceil((weight - 1)) * 2.2
        else:
            totel = 9 + ceil(weight - 1) * 6
        return totel

    elif adress in ["天津","河北省","山东省","福建省"]:
        if weight <= 3:
            totel = 4
        elif (weight >= 3) and (weight <= 30):
            totel = 5 + ceil((weight - 1)) * 1.5
        else:
            totel = 130 + ceil(weight - 1) * 12
        return totel

    elif adress in ["广西省"]:
        if weight <= 3:
            totel = 4
        elif (weight >= 3) and (weight <= 30):
            totel = 5 + ceil((weight - 1)) * 2.5
        else:
            totel = 130 + ceil(weight - 1) * 12
        return totel

    elif adress in ["陕西省"]:
        if weight <= 3:
            totel = 4
        elif (weight >= 3) and (weight <= 30):
            totel = 6 + ceil((weight - 1)) * 2.5
        else:
            totel = 130 + ceil(weight - 1) * 12
        return totel

    elif adress in ["重庆","四川省"]:
        if weight <= 3:
            totel = 4
        elif (weight >= 3) and (weight <= 30):
            totel = 6 + ceil((weight - 1)) * 3
        else:
            totel = 130 + ceil(weight - 1) * 12
        return totel

    elif adress in ["辽宁省","吉林省","黑龙江省"]:
        if weight <= 3:
            totel = 4
        elif (weight >= 3) and (weight <= 30):
            totel = 6 + ceil((weight - 1)) * 3.5
        else:
            totel = 130 + ceil(weight - 1) * 12
        return totel

    elif adress in ["贵州省"]:
        if weight <= 3:
            totel = 4
        elif (weight >= 3) and (weight <= 30):
            totel = 6 + ceil((weight - 1)) * 4
        else:
            totel = 130 + ceil(weight - 1) * 12
        return totel

    elif adress in ["山西省"]:
        if weight <= 3:
            totel = 4
        elif (weight >= 3) and (weight <= 30):
            totel = 8 + ceil((weight - 1)) * 4
        else:
            totel = 130 + ceil(weight - 1) * 12
        return totel

    elif adress in ["宁夏省","青海省"]:
        if weight <= 3:
            totel = 4
        elif (weight >= 3) and (weight <= 30):
            totel = 8 + ceil((weight - 1)) * 7
        else:
            totel = 130 + ceil(weight - 1) * 12
        return totel

    elif adress in ["云南省","内蒙古","甘肃省","海南省"]:
        if weight <= 1:
            totel = 8
        elif (weight >= 1) and (weight <= 30):
            totel = 8 + ceil((weight - 1)) * 7
        else:
            totel = 130 + ceil(weight - 1) * 12
        return totel

    elif adress in ["西藏","新疆省"]:
        if weight <= 1:
            totel = 17
        elif (weight >= 1) and (weight <= 30):
            totel = 17 + ceil(weight - 1) * 15
        else:
            totel = 130 + ceil(weight - 1) * 12
        return totel

    else:
        print("你输入的省份不合法！！！")


file_path = input("请输入文件路径：")
sheet_name = input("请输入工作簿名称：")
pf = pd.read_excel(file_path, sheet_name=sheet_name)
# 获取省份一列
pro = pf["省份"].values.tolist()
# 获取重量一列
wt = pf["重量"].values.tolist()
# 核算列
totel = []
for p, w in zip(pro, wt):
    print(p, w)
    totel.append(account(p, w))

pf["最新核算结果"] = totel
file_name = os.path.basename(file_path)
pf.to_excel(os.path.join(os.path.dirname(file_path),
                         os.path.basename(file_path).split(".")[0] + sheet_name + "最新核算结果" + ".xlsx"))
