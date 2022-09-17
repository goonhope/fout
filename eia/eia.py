# -*- coding: utf-8 -*-
'''
@FileName	:   info_eia.py
@Created  :   2021/08/19 12:51
@Updated  :   2022/06/19 12:51
@Author		:   goonhope@gmail.com
@Function	:   信息查询
@notes    ：  环评受理公告抓取 console.log(Array.from($$("div[role=rowheader]"),x=>x.innerText).filter(i=>i !=". .").join("\n"))
'''
from Project.fsel import *
import os,time
from Project.function import excel
from Project.Office.fxls import pd_xls,Analysis
import pandas as pd
import numpy as np



city_args = {"珠海": (r'http://ssthjj.zhuhai.gov.cn/zxfw/xmgsgg/slgg/index.html', '报告', 'tr:not(.firstRow) td', 1),
                  "广东": (r"http://gdee.gd.gov.cn/gsgg/index.html", "报告书受理", "#logPanel tr:not(.firstRow) td", 29),
                  "广州": (r"http://sthjj.gz.gov.cn/hjgl/jsxm/hpslgg/index.html", "报告", "font tr:not(.firstRow) td", 51),
                  "中山": [(r"http://zsepb.zs.gov.cn/xxml/ztzl/gcjslyxmxx/ssthjjhpspgs/slgs/index.html","报告","table tr:not(.firstRow) td",4),
                                ("http://zsepb.zs.gov.cn/xxml/ztzl/gcjslyxmxx/zqhbfjhpspgs/slgs/index.html","报告","table tr:not(.firstRow) td",4)],
                  }



def temp_dir(file=""):
    """临时目录"""
    return os.path.join(os.path.dirname(__file__), "config",file)


def open_txt(keys="",file="delaw.log"):
    """集成txt读取和写入"""
    file = temp_dir(file)
    if keys:
        with open(file, "w", encoding="utf-8") as f:
            if isinstance(keys,list):
                f.write(list2d(keys))
            else: print("check!!!")
    else:
        if not os.path.exists(file): return []
        else:
            with open(file, "r", encoding="utf-8") as f:
                info = f.read()
                return list2d(info)


def list2d(llist,stp=r"'\" \t"):
    """二维数列与字符串转换"""
    if isinstance(llist,list):
        return "\n".join(["\t".join([str(x).strip(stp) for x in key]) for key in llist])
    elif isinstance(llist,str):
        return [x.strip(stp).split("\t") for x in llist.strip().split("\n") if x.strip()] if llist else []
    else:
        return False


class Nd(CDriver):
    """珠海市环评"""
    def __init__(self,city="珠海",*args,**kwargs):
        super().__init__(*args,**kwargs)
        self.hold, self.ifs, self.city = dict(), False, city
        self.fname = f"eia_public_{self.city}.log"
        self.holdo = open_txt(file=self.fname)
        self.inkeys = [x[-1].split("/")[-1].replace(".pdf","") for x in self.holdo[:3]] if self.holdo else False
        self.cargs = city_args.get(self.city,city_args["珠海"])

    def get_info(self,urls,css="tr:not(.firstRow) td"):
        """获取二级页面信息"""
        for url in urls:
            try:
                self.init_page(url)
                self.choose(url)
                if self.ifs: break
                if "404" not in self.grap_tag("title")[0].text:
                    pdfs = self.tag_content("a.nfw-cms-attachment",tag="href",ink="")
                    elem = self.grap_tag(css)
                    if elem and pdfs:
                        texts = [x.text.replace(" ","").strip() for x in elem][:-1]
                        self.hold.update({"\t".join(texts):pdfs[0]})
                        print(*texts,*pdfs)
                        time.sleep(1.61)
            except Exception as e:
                print(e)

    def go(self):
        """集成"""
        func = self.get_pages3 if self.city in "中山".split() else self.get_pages2
        func(*self.cargs) if isinstance(self.cargs,tuple) else [func(*x) for x in self.cargs]
        self._quit()
        self.list_from()

    @excel(True, dir=temp_dir(),na="环评", t=False)
    def list_from(self,titles='受理日期 项目名称 建设单位 建设地点 环评单位 类型 文件地址'):
        """转二维列表输出 """
        lhold = [x.split() + [y] for x, y in self.hold.items()]
        titles = titles.strip().split()
        lhold.extend(self.holdo)
        open_txt(lhold,self.fname)
        lhold.insert(0,titles)
        # excel_raw(os.path.join(fdir(),"x.xlsx"),info=lhold)
        return lhold

    def get_pages2(self,url=r"http://ssthjj.zhuhai.gov.cn/zxfw/xmgsgg/slgg/index.html",kw="报告",css="tr:not(.firstRow) td",n=20):
        """直接抓取公示页面"""
        for i in range(1, n):
            urlx = f"_{str(i)}".join(os.path.splitext(url)) if i > 1 else url
            self.init_page(urlx)
            sub_urls =[tag.get_attribute("href") for tag in self.grap_tag("a") if tag.text and kw in tag.text]
            self.get_info(sub_urls,css)
            if self.ifs: break
            time.sleep(2)

    def get_pages3(self,url=r"http://ssthjj.zhuhai.gov.cn/zxfw/xmgsgg/slgg/index.html",kw="报告",css="tr:not(.firstRow) td",n=20):
        """中山-直接点击"""
        for i in range(1, n):
            urlx = f"_{str(i)}".join(os.path.splitext(url)) if i > 1 else url
            self.init_page(urlx)
            time.sleep(1.67)
            try: self.iter_click(self.get_info3,self.ifs,"a",kw,1.32,csss=css)
            except Exception as e: print(e)
            if self.ifs: break
            time.sleep(2)

    def get_info3(self,csss="tr:not(.firstRow) td"):
        """获取二级页面信息 - 中山"""
        if "404" not in self.grap_tag("title")[0].text:
            pdfs = self.tag_content("a.nfw-cms-attachment",tag="href",ink="")
            if pdfs and "about:blank" in pdfs[0]: pdfs[0] = self.driver.current_url
            self.choose(pdfs[0])
            elem = self.grap_tag(csss)
            if elem and pdfs:
                texts = [x.text.replace(" ","").strip() for x in elem][:-1]
                self.hold.update({"\t".join(texts):pdfs[0]})
                print(*texts,*pdfs)
                time.sleep(1.61)

    def choose(self,stri):
        """读取内容"""
        self.ifs = self.inkeys and any(x in stri for x in self.inkeys)



def eia(city="珠海"):
    """环评 """
    d = Nd(city,crm("",img=False))
    d.go()


@pd_xls
def deal_eia(file=temp_dir(f"环评.xls"),pdata=True):
    "数据分析"
    xls = Analysis(file,dt=str)
    old = xls.open()
    xls.add_clos(lambda x: x.replace("（","(").replace("）",")"), "环评单位", "环评单位")
    xls.add_clos(lambda x: 1, "all", "类型")
    xls.add_clos(lambda x: 1 if "表" in x else 0 ,"表","类型")
    xls.add_clos(lambda x: 1 if "书" in x else 0, "书", "类型")
    da = pd.pivot_table(xls.sdata,index="环评单位".split(), values="表 书 all".split(),aggfunc=np.sum)
    da = da["表 书 all".split()]
    # da.sum()
    return (old,xls.sdata,da) if pdata else da.index.tolist()


@excel(1,dir=temp_dir(),na="eia_unit",t=False)
def get_eia_unit(unit="铁汉环保集团有限公司"):
    url = r"http://114.251.10.92:8080/XYPT/unit/list"
    if isinstance(unit,str): unit = unit.strip().split()
    d = CDriver(crm(url,img=False))
    hold = ['序号\t环评单位\t统一社会信用代码\t住所\t环评工程师数量\t主要编制人员数量\t当前状态\t信用记录'.split()]
    for en,n in enumerate(unit):
        d.input_key(n,"#unitName")
        d.click_by_css("#btnSubmit")
        x = d.tag_content("#contentTable td",ink="")
        hold.append(x) if x else hold.append(["1",n])
        d.tsleep(1.2)
    d._quit()
    print(hold)
    return hold


@pd_xls
def deal_unit(file=temp_dir("eia_unit.xls")):
    """单位信息处理"""
    xls = Analysis(file,dt=str)
    old = xls.open()
    xls.add_clos(lambda x:str(x)[:2] if x else "" ,"省","住所")
    xls.add_clos(lambda x: str(x).split("-")[1][:2] if x and "-" in str(x) else "", "市", "住所")
    return old,xls.sdata


@pd_xls
def deal_all(file=temp_dir("环评_PYD.xlsx")):
    xls = Analysis(file,sheet=2)
    old = xls.open()
    da = pd.pivot_table(xls.sdata,index="省 市".split(), values="表 书 all".split(),aggfunc=np.sum)
    return old,da


def rd_excel(xls,sheet=0,skp=None,col=None,dt=None):
    """读取xlsx 转"""
    engine = 'openpyxl' if xls.lower().endswith("xlsx") else None
    data = pd.read_excel(xls,sheet,skiprows=skp,usecols=col,engine=engine,dtype=dt)
    return data


@pd_xls
def deal_all(xls1=temp_dir("环评_PYD.xlsx"),xls2=temp_dir("eia_unit_PYD.xlsx")):
    """两表数据结合"""
    pd1 = rd_excel(xls1,2)
    pd2 = rd_excel(xls2,1)
    out = pd.merge(pd1,pd2[["环评单位","省","市"]],how="inner",on="环评单位")
    da = pd.pivot_table(out, index="省 市".split(), values="表 书 all".split(), aggfunc=np.sum)
    return out,da


def eia_anly(city="广州"):
    """集成"""
    # eia(city)
    # deal_eia(temp_dir("环评.xls"), pdata=True)
    # unit = deal_eia(temp_dir("环评.xls"), pdata=False)
    # get_eia_unit(unit)
    # deal_unit(temp_dir("eia_unit.xls"))
    deal_all(temp_dir("环评_PYD.xlsx"))
 


if __name__ == "__main__":
    eia_anly()


