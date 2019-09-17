# SearchAromabyCAS.py
#!/usr/bin/env python3
# coding: utf-8
'''
整个爬取基于 CAS 号，所以提供的 excel 表格中 CAS 号一定要正确；
默认待查询表格中 CAS 号列名为“CAS”（可更改）; 
默认待查询的为“Sheet1”（可更改）；
查询结果储存在原 excel 表中名为“香气查询”的 sheet 中。
'''
from openpyxl import load_workbook,Workbook
import pandas as pd
from urllib import request, parse
import openpyxl
import requests
from bs4 import BeautifulSoup
import traceback
import re
import numpy as np

class Searcharoma():
    def __init__(self,path):
        self.path = path
        try:
            self.df = pd.read_excel(io=self.path, sheet_name = 'Sheet1') #可在此更改包含待查询 CAS 号的 sheet 名
            self.writer = pd.ExcelWriter(self.path,sheet_name = 'Sheet1', engine = 'openpyxl') #可在此更改包含待查询 CAS 号的 sheet 名
            self.book = load_workbook(self.writer.path)
            self.writer.book = self.book
        except:
            print("不存在“Sheet1”，请检查 sheet 名。")
   
    def getHTMLText(self,url,kv,cookies, code = "ascii"):
        try:
            self.r = requests.get(url, cookies = cookies, headers = kv, timeout = 30)
            self.r.raise_for_status()
            self.r.encoding = code
            return self.r.text
        except:
            return "" 

#chang0cas 函数目的是替换GC×GC的 Peak Table 中 CAS 为 0-00-0 的物质     
    def change0cas(self):
        try:
            for index, row in self.df.iterrows():
                if row['CAS'] == '0-00-0': #可在此更改 CAS 号所在列名
                    for key in dicts:
                        if row['Name'] == key: #可在此更改化合物英文名所在列名
                            self.df.at[index,'CAS'] = dicts[key] #可在此更改 CAS 号所在列名
                        else:
                            continue
                else:
                    continue
        except:
            print("CAS 号所在列名不是“CAS”，请检查列名。")
        
        return self.df

#getCasList 函数目的是获得 CAS 号的列表，默认表格中 CAS 号所在列列名为 CAS
    def getCasList(self):
        self.changeddata = self.change0cas()
        self.lst = []
        for col in self.changeddata['CAS']: #可在此更改 CAS 号所在列名
            try:
                if col == 'CAS': #可在此更改 CAS 号所在列名
                    continue
                else:
                    self.lst.append(col)
            except:    
                continue
       
        return self.lst
    
    '''
    flavornetsearch 函数的目的是爬取 flavornet 中的香气描述；
    flavornet 中 aroma 的信息储存在 p 标签的 string 中，在【Percepts: 】后，故可用正则匹配查找；
    返回的是一个列表。
    '''

    def flavornetsearch(self):
        flavorURL = 'http://www.flavornet.org/info/'
        self.flst = []
        self.clst = self.getCasList()
        count = 0
        
        for cas in self.clst:
            self.url_f = flavorURL + cas + ".html"
            coo_f = 't=85db5e7cb0133f23f29f98c7d6955615; cna=3uklFEhvXUoCAd9H6ovaVLTG; isg=BM3NGT0Oqmp6Mg4qfcGPnvDY3-pNqzF2joji8w9SGWTYBu241_taTS6UdFrF3Rk0; miid=983575671563913813; thw=cn; um=535523100CBE37C36EEFF761CFAC96BC4CD04CD48E6631C3112393F438E181DF6B34171FDA66B2C2CD43AD3E795C914C34A100CE538767508DAD6914FD9E61CE; _cc_=W5iHLLyFfA%3D%3D; tg=0; enc=oRI1V9aX5p%2BnPbULesXvnR%2BUwIh9CHIuErw0qljnmbKe0Ecu1Gxwa4C4%2FzONeGVH9StU4Isw64KTx9EHQEhI2g%3D%3D; hng=CN%7Czh-CN%7CCNY%7C156; mt=ci=0_0; hibext_instdsigdipv2=1; JSESSIONID=EC33B48CDDBA7F11577AA9FEB44F0DF3'
            cookies = {}
            for line in coo_f.split(';'):
                name,value = line.strip().split('=',1)
                cookies[name] = value
            kv = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:62.0) Gecko/20100101 Firefox/62.0'}
            self.html_f = self.getHTMLText(self.url_f, kv, cookies, code = "ascii")
        
            try:
                if self.html_f == "":
                    self.flst.append('该化合物在flavornet上查不到哦')
                    count += 1
                    print("\rflavornet查询进度: {:.2f}%".format(count*100/len(self.clst)), end = "")
                else:
                    self.soup_f = BeautifulSoup(self.html_f, 'html.parser')
                    for tag in self.soup_f.find_all('p', string=re.compile('Percepts')):
                        self.flavor_f = tag.string.split(':')[1] 
                        self.flst.append(self.flavor_f)
                        count = count + 1
                        print("\rflavornet查询进度: {:.2f}%".format(count*100/len(self.clst)), end = "")
            except:
                count += 1
                print("\rflavornet查询进度: {:.2f}%".format(count*100/len(self.clst)), end = "")
                continue
        
        return self.flst

    '''
    goodscentsearch 函数目的是爬取 goodscents 中的 aroma、flavor，返回的是索引与changeddata一致、包含两列查询结果的 dataframe 。
    爬取过程分为三步：1.对于 CAS 号仍为 0-00-0 的，爬取结果提示要检查 cas 号；
                   2.对于未能提供 odor 信息/ goodscents 上未收录该 cas 号的，爬取结果显示查不到（利用正则匹配 odor : ）；
                   3.对于可爬取的，ordor 和 flavor 信息分别储存在两个 span class="lstw11" 的 string 中；
                     部分物质均提供了 ordor 和 flavor 信息，部分只提供了 ordor 。根据长度判断情况。        
    将爬取结果先分别储存到 glst_odor 和 glst_flavor 两个列表中，再将列表写入 df_g 的 dataframe 中。
    '''

    def goodscentsearch(self):
        gsURL = 'http://www.thegoodscentscompany.com/search3.php?qName='
        self.clst = self.getCasList()
        self.df_g = pd.DataFrame(index = self.changeddata.index)
        self.glst_odor = []
        self.glst_flavor = []
        count = 0

        for cas in self.clst:
            self.url_g = gsURL + cas
            coo = 't=85db5e7cb0133f23f29f98c7d6955615; cna=3uklFEhvXUoCAd9H6ovaVLTG; isg=BM3NGT0Oqmp6Mg4qfcGPnvDY3-pNqzF2joji8w9SGWTYBu241_taTS6UdFrF3Rk0; miid=983575671563913813; thw=cn; um=535523100CBE37C36EEFF761CFAC96BC4CD04CD48E6631C3112393F438E181DF6B34171FDA66B2C2CD43AD3E795C914C34A100CE538767508DAD6914FD9E61CE; _cc_=W5iHLLyFfA%3D%3D; tg=0; enc=oRI1V9aX5p%2BnPbULesXvnR%2BUwIh9CHIuErw0qljnmbKe0Ecu1Gxwa4C4%2FzONeGVH9StU4Isw64KTx9EHQEhI2g%3D%3D; hng=CN%7Czh-CN%7CCNY%7C156; mt=ci=0_0; hibext_instdsigdipv2=1; JSESSIONID=EC33B48CDDBA7F11577AA9FEB44F0DF3'
            cookies = {}
            for line in coo.split(';'):
                name,value = line.strip().split('=', 1)
                cookies[name] = value
            kv = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:62.0) Gecko/20100101 Firefox/62.0'}
            self.html_g = self.getHTMLText(self.url_g, kv, cookies, code = "ascii")
            try:
                if cas == '0-00-0':
                    self.glst_odor.append('该化合物需要检查下cas号哦')
                    self.glst_flavor.append('该化合物需要检查下cas号哦')
                    count += 1
                    print("\rgoodscents查询进度: {:.2f}%".format(count*100/len(self.clst)), end = "")
                elif re.findall(r'Odor : ',self.html_g) == []:
                    self.glst_odor.append('该化合物在goodscents上查不到哦')
                    self.glst_flavor.append('该化合物在goodscents上查不到哦')
                    count += 1
                    print("\rgoodscents查询进度: {:.2f}%".format(count*100/len(self.clst)), end = "")
                else:
                    self.soup_g = BeautifulSoup(self.html_g, 'html.parser')
                    self.result_g = self.soup_g.find_all('span','lstw11')
                    self.odor_g = self.result_g[0].string
                    self.glst_odor.append(self.odor_g)
                    if len(self.result_g)>1:
                        self.flavor_g = self.result_g[1].string
                        self.glst_flavor.append(self.flavor_g)
                    else:
                        self.glst_flavor.append('该化合物没有提供flavor')
                    count += 1
                    print("\rgoodscents查询进度: {:.2f}%".format(count*100/len(self.clst)), end = "")
            except:
                count = count + 1
                print("\rgoodscents查询进度: {:.2f}%".format(count*100/len(self.clst)), end = "")
                continue
        self.df_g['goodscents 查询结果-odor'] = pd.Series(self.glst_odor).values
        self.df_g['goodscents 查询结果-flavor'] = pd.Series(self.glst_flavor).values
        
        return self.df_g

#使用 openpyxl 库可写入同一个 excel 表中。
    def write2excel(self):
        self.se = pd.Series(self.flavornetsearch())
        self.df_goodscent = self.goodscentsearch()
        self.changeddata['flavornet 查询结果'] = self.se.values
        self.search_result = self.changeddata.join(self.df_goodscent)
        self.search_result.to_excel(excel_writer=self.writer, sheet_name = "香气查询", index = False)
        self.writer.save()
        self.writer.close()

dicts = {"ETHYL (S)-(-)-LACTATE":"687-47-8", "4-Oxononanal":"74327-29-0",\
 "1b,5,5,6a-Tetramethyl-octahydro-1-oxa-cyclopropa[a]inden-6-one":"",\
"Valeric acid, 4-pentadecyl ester":"959021-71-7","E-11-Hexadecenoic acid, ethyl ester":"766512-32-7",\
"4-(2-Acetoxyphenyl)-1-ethyl-3-methyl-5-(4-nitrophenyl)pyrazole":"",\
"Cyclobutanecarboxylic acid, 2-propenyl ester":"959063-53-7","Indole, 3-methyl-2-(2-dimethylaminopropyl)-":"",\
"2,4-Dimethylhept-1-ene":"19549-87-2","Heptane, 3-ethyl-5-methylene-":"959078-90-1",\
"3,4-Diethyl-2-hexene":"59643-70-8","Oxalic acid, isobutyl pentyl ester":"959067-90-4","2-Ethyl-1-hexanol":"104-76-7",\
"Sulfurous acid, isobutyl pentyl ester":"959275-59-3","Terephthalic acid, butyl tridec-2-yn-1-yl ester":"",\
"Sulfurous acid, decyl 2-propyl ester":"959268-17-8","2-Ethylhexanal":"123-05-7",\
"Oxalic acid, pentyl propyl ester":"959267-78-8","Oxalic acid, allyl nonyl ester":"959078-60-5",\
"1-Hexene, 4,4-diethyl-":"959100-72-2","Sulfurous acid, butyl 2-ethylhexyl ester":"959311-34-3",\
"Oxalic acid, allyl octyl ester":"61670-32-4","1-Hexene, 2,4,4-triethyl-":"936116-63-1",\
"Oxalic acid, isobutyl octyl ester":"959275-41-3","2-Propyl-1-Pentanol, trifluoroacetate":"",\
"Bicyclo[3.2.2]non-8-en-6-ol, (1R,5-cis,6-cis)-":"683270-71-5","Cyclopent-4-ene-1,3-dione":"930-60-9",\
"1,4-Methanocycloocta[d]pyridazine, 1,4,4a,5,6,9,10,10a-octahydro-11,11-dimethyl-, (1à,4à,4aà,10aà)-":"",\
"Tricyclo[7.1.0.0[1,3]]decane-2-carbaldehyde":"898838-53-4","1-Phenyl-2-propanone":"103-79-7",\
"Oxime-, methoxy-phenyl-_":"67160-14-9","(E)-4-Oxohex-2-enal":"2492-43-5",\
"Z-1,6-Undecadiene":"","L-(+)-Threose, aldononitrile, triacetate":"","1,4:3,6-Dianhydro-à-d-glucopyranose":"4451-30-3",\
"Z-10-Tetradecen-1-ol acetate":"35153-16-3","Oxalic acid, heptyl propyl ester":"959275-46-8",\
"Oxalic acid, isobutyl nonyl ester":"959275-48-0","1-Iodo-2-methylnonane":"",\
"cis-Linaloloxide":"11063-77-7","2-Furanone, 2,5-dihydro-3,5-dimethyl":"5584-69-0",\
"(1R,5R)-4-Methylene-1-((R)-6-methylhept-5-en-2-yl)bicyclo[3.1.0]hexane, (relative configuration)":"",\
"trans-Z-à-Bisabolene epoxide":"","4,5-Oxalic acid, pentyl propyl esterdi-epi-aristolochene":"","6-epi-shyobunol":"35727-45-8",\
"Hexanoic acid, 3,5,5-trimethyl-, 2-ethylhexyl ester":"70969-70-9",\
"Sulfurous acid, butyl pentyl ester":"959311-33-2","Oxalic acid, bis(isobutyl) ester":"2050-61-5",\
"5,9-Dodecadien-2-one, 6,10-dimethyl-, (E,E))-":"13125-74-1","á-Vatirenene":"",\
"á-copaene":"3856-25-5","2-Ethyl-hexoic acid":"149-57-5","Oxalic acid, allyl heptyl ester":"959312-52-8",\
"Oxalic acid, butyl propyl ester":"26404-30-8","Z-1,9-Dodecadiene":"157887-78-0",\
"14-Oxa-1,11-diazatetracyclo[7.4.1.0(2,7).0(10,12)]tetradeca-2,4,6-triene, 11-acetyl-6,9-bis(acetyloxy)-4-formyl-8-[(aminocarbonyloxy)methyl]-":"",\
"1,3-Phenylene, bis(3-phenylpropenoate)":"22129-64-2","Oxalic acid, allyl isobutyl ester":"959312-50-6",\
"3-Butenoic acid, 2-oxo-4-phenyl-":"17451-19-3",\
"Tetracyclo[6.1.0.0(2,4).0(5,7)]nonane, 3,3,6,6,9,9-hexamethyl-, cis,cis,trans--":"51898-92-1",\
"Cyclobutane-1,1-dicarboxamide, N,N'-di-benzoyloxy-":"959242-89-8","cis-1-Ethoxy-1-butene":"1528-20-7",\
"Phthalic acid, heptyl pentyl ester":"70794-32-0","2-(2-Methoxyethoxy)ethyl acetate":"629-38-9",\
"Sulfurous acid, hexyl pentyl ester":"959059-25-7","Sulfurous acid, nonyl pentyl ester":"959274-88-5",\
"1-(3,3-Dimethyl-but-1-ynyl)-1,2-dimethyl-3-methylene-cyclopropane":"959091-65-7",\
"Hydroxymethyl 2-hydroxy-2-methylpropionate":"959259-22-4","4,5-di-epi-aristolochene":"",\
"Salicylic acid, tert.-butyl ester":"23408-05-1","Carbonic acid, heptyl vinyl ester":"",\
"4,4-Dimethylpent-2-enal":"926-37-4","(Z,Z)-à-Farnesene":"28973-99-1",\
"Bicyclo[4.1.0]-3-heptene, 2-isopropenyl-5-isopropyl-7,7-dimethyl-":"874302-33-7",\
"Oxalic acid, isobutyl propyl ester":"26404-31-9","Oxalic acid, allyl hexyl ester":"959267-74-4"}