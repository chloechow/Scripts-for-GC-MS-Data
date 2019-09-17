# readme
## sa_main.py
* 该脚本可用于自动查询 excel 表中某物质（**CAS号**）在 flavornet 和 goodscents 上的香气描述。

### 脚本基本注意事项
* sa_main.py 为主程序。
* 该脚本使用 **python 3** 编写，请安装 python 3 环境。
* 请确认安装了 openpyxl、pandas、urllib、requests、bs4、traceback、re 和 numpy 库。
* 该脚本只限于处理 **xlsx** 类型文件。
* 该脚本爬取基于 CAS 号，请确认 excel 表中包含 CAS 号，**且列名为“CAS”**（列名可在 SearchAromabyCAS.py 文件中更改）。
* 该脚本默认读取 excel 表中的“Sheet1”，请确认包含待查询 CAS 号的 **sheet 名为“Sheet1”**（sheet 名可在 SearchAromabyCAS.py 文件中更改）。
* 该脚本将查询结果储存在原 excel 表中名为“香气查询”的 sheet 中。
* 请将 sa_main.py 与 SearchAromabyCAS.py 置于同一文件夹下。

### 脚本使用示例
待查询 excel 表应如图所示：
![image](https://github.com/chloechow/Scripts-for-GC-MS-Data/blob/master/SearchAromabyCAS/example_excel.png)

查询后的 excel 表如图所示：
![image](https://github.com/chloechow/Scripts-for-GC-MS-Data/blob/master/SearchAromabyCAS/example_scrapy.png)

## 香气查询.ipynb
* 脚本的 ipynb 版本，注意事项同上，可直接运行。
