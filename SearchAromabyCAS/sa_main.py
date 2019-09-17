#!/usr/bin/env python3

import SearchAromabyCAS
import os

print("""
该脚本可用于自动查询 flavornet 和 goodscents 上某物质的香气描述。
该脚本使用 python 3 编写，请确认安装了 openpyxl、pandas、urllib、requests、bs4、traceback、re 和 numpy 库。
该脚本只限于处理 xlsx 类型文件。
该脚本爬取基于 CAS 号，请确认 excel 表中包含 CAS 号，且列名为“CAS”（列名可在 SearchAromabyCAS.py 文件中更改）。
该脚本默认读取 excel 表中的“Sheet1”，请确认包含待查询 CAS 号的 sheet 名为“Sheet1”（sheet 名可在 SearchAromabyCAS.py 文件中更改）。
该脚本将查询结果储存在原 excel 表中名为“香气查询”的 sheet 中。
""")
pathget = input("""
获取绝对路径：（windows 系统）选中文件，按住shift键, 右键, 会出现复制为路径的选项
请输入待查询 excel 的绝对路径：
""")
pathl=pathget.split(",")
coun = 0
for path_e in pathl:
    coun = coun +1
    print("\r当前处理文件: {}/{}，文件路径：{}".format(coun,len(pathl),path_e))
    if os.path.splitext(path_e)[-1][1:] != "xlsx":
        print("文件类型不对，该脚本只限于处理 xlsx 类型文件。")
    else:
        if os.path.exists(path_e):
            file = SearchAromabyCAS.Searcharoma(path_e)
            file.write2excel()
        else:
            print("文件路径不存在，请检查文件路径。")
