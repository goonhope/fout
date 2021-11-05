# -*- coding: utf-8 -*-
'''
@FileName	:   fout.py
@Created    :   2021/08/07
@Updated    :   2021/11/05
@Author		:   Teddy, goonhope@gmail.com, Zhuhai
@Function	:   function函数库规范化
@notes      ：  知乎公开函数库
'''

import os,time,shutil,zipfile
from functools import wraps
from PIL import Image
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT as align,WD_BREAK
from docx.shared import Inches,Pt
from PIL import Image, ImageDraw, ImageFont
from win32com.client import Dispatch


def shower(show=True):
    '''print装饰器'''
    def inner(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            print(f"@{func.__name__}...") if show else None
            goal = func(*args, **kwargs)
            print(goal) if len(goal[-1])  < 20 else None
            return goal
        return wrapper
    return inner


def string_check(x,enz="",ink=".",out="$$$",stz="",win=True):
    """字符串check  return bool逻辑结果
    x: string 字符串
    enz: endswith字符组,
    ink：in keyword 含有字符
    out：not in x不含字符"""
    if win:
        enz,ink,out,stz,x = enz.lower(),ink.lower(),out.lower(),stz.lower(),x.lower()
    enz = any(x.endswith(y) for y in enz.strip().split()) if enz else True
    stz = any(x.startswith(y) for y in stz.strip().split()) if stz else True
    ink = all(y in x for y in ink.strip().split()) if ink else True
    out = all(y not in x for y in out.strip().split()) if out else True
    return enz and stz and ink and out


@shower()
def get_root_sub(path,enz="img",file=True,**args):
    """获取path路径下的 文件名称或文件夹名list:
    :return root-str, files or dirs-[str]"""
    img = 'jpg jpeg tiff tif png bmp'
    enz = img if enz == "img" else enz
    files = [os.path.split(path)[-1]] if os.path.isfile(path) else [x for x in os.listdir(path) if string_check(x,enz,**args)]
    root = path if os.path.isdir(path) else os.path.split(path)[0]  # os.path.dirname(path)
    sub = files if file else [x for x in os.listdir(root) if string_check(x, enz, **args) and os.path.isdir(os.path.join(root,x))]
    return root, sub


def extracts_imgs(docx,dlt=True):
    """提取docx/pptx/xlsx中图片，docx"""
    enz = "docx pptx xlsx".split()
    inword = "word ppt xl".split()
    info = {x:y for x,y in zip(enz,inword)}
    root, docxs = get_root_sub(docx," ".join(enz),out="~$")
    for sdocx in docxs:
        docx = os.path.join(root,sdocx)
        docxb = "_Bakup".join(os.path.splitext(docx))
        shutil.copy(docx,docxb) # bakeup
        nroot = os.path.splitext(docx)[0]
        fzip = nroot + "_.zip"
        os.remove(fzip) if os.path.exists(fzip) else None # exist delete
        os.renames(docxb,fzip)
        xroot = nroot + "_old"
        os.makedirs(xroot) if not os.path.exists(xroot) else None # not exist mk
        with zipfile.ZipFile(fzip, 'r') as fi:
            for file in fi.namelist():
                fi.extract(file, xroot)  # 将压缩包里的word文件夹解压出来
        for x in enz:
            if sdocx.endswith(x):
                media = f"{info[x]}\media"
                imgf = os.path.join(xroot, media)
                shutil.rmtree(nroot) if os.path.exists(nroot) and dlt else None
                os.renames(imgf, nroot) if os.path.exists(imgf) else None # 拷贝到新目录，名称为word文件的名字
                re_name(nroot)
                emf_png(nroot)
        shutil.rmtree(xroot)
        os.remove(fzip)
    os.system(f"""start "" {root}""")


def emf_png(path,dlt=True):
    '''emf图片转png,dlt为真删除emf图片'''
    root,files = get_root_sub(path,enz=".emf")
    for file in files:
        file = os.path.join(root,file)
        Image.open(file).save(file.replace(".emf", ".png"))
        os.remove(file) if dlt else None
    root, files = get_root_sub(path)


def re_name(new_imagedir,st=1):
    '''Windows序号命名修改'''
    # 重命名文件
    l = os.listdir(new_imagedir)
    l_sorted = sorted(l, key=lambda x: int(str(x).split(".")[0].split("e")[-1]))  # 排序转换
    j_s = len(l_sorted)
    for j,pic in enumerate(l_sorted):  # 重命名
        Olddir = os.path.join(new_imagedir, pic)
        if os.path.isfile(Olddir):
            Newdir = os.path.join(new_imagedir, str(j + st).zfill(3) + "_" + pic)
            os.rename(Olddir, Newdir)
    return


def pdf_to_images(path, dpi=300,sep="_",ink=".",**kwargs):
    '''pdf转为图片，默认DPI=300  需要wand库支持
    path:支持文件名及路径'''
    from wand.image import Image
    from wand.color import Color
    path, pdf_names = get_root_sub(path, enz="PDF pdf",ink=ink,**kwargs)
    for pdf_name in pdf_names:
        fpdf = os.path.join(path, pdf_name)
        image_jpeg = Image(filename=fpdf, resolution=dpi)
        pages = len(image_jpeg.sequence)
        for i,img in enumerate(image_jpeg.sequence):
            with Image(image=img) as img_page:
                img_page.format = 'png'
                img_page.background_color = Color('white')
                img_page.alpha_channel = 'remove'
                fimg = os.path.splitext(fpdf)[0] + f"{sep + str(i).zfill(3)}.jpg"
                with open(fimg, 'wb') as img_name:
                    img_name.write(img_page.make_blob('jpg'))
        print("[Filename:\t{0},Imags:\t{1}]".format(pdf_name, pages))


def dos(cstr,show=False):
    """运行dos命令"""
    cstr = f'chcp 65001 && {cstr}'  # cmd字符代码UTF8 ，window 默认936-GBK
    result = os.popen(cstr).read().strip()
    print(result) if show else None
    return


class Word(object):
    '''word-docx文件 生成类'''
    def __init__(self,template,clear=True):
        self.template = template
        self.clear = clear
        self.out = "_插图".join(os.path.splitext(self.template))
        self.doc = Document(template)
        if self.clear: self.doc._body.clear_content()  # 清楚内容

    def add_para(self,content="",style="表图"):
        ''''添加段落：默认空行'''
        return self.doc.add_paragraph(content, style=style)

    def insert_imgs(self,root,imgs,inch=6,sp=True):
        ''''批量插入图片'''
        run = self.add_para().add_run()  # text=None, style=None
        for img in imgs:
            img = os.path.join(root,img)
            width, height = Image.open(img).size
            finch = Inches(inch) if width >= 21 / 2.54 * 72 else None
            run.add_picture(img, width=finch)
        run.add_break(WD_BREAK.PAGE) if sp else None  # 加分页符
        return run

    def replace(self,old="认定",new="技改"):
        '''替换word中字符''' # doc.sections[0].header.paragraphs[0].text
        for para in self.doc.paragraphs:
            if old in para.text:
                para.text = para.text.replace(old, new)
        for sect in self.doc.sections:  # 节
            sect.different_first_page_header_footer = 0 # 取消首页不同
            hf = sect.even_page_header.paragraphs + sect.header.paragraphs
            for hp in hf:
                if old in hp.text:
                    print(hp.text,end="\t")
                    hp.text = hp.text.replace(old, new)
                    print(hp.text)
        self._save()

    def _save(self):
        '''保存文件，如有原重名doc\docx文件先删除'''
        os.remove(self.out) if os.path.exists(self.out) else None
        self.doc.save(self.out)

    def pdf_word(self,imgdir,pdf=True):
        ''''以目录内pdf文件名为标题，批量插入（pdf转）图片到word文件中'''
        titles = [x.lower().replace(".pdf","") for x in get_root_sub(imgdir,enz=".pdf .PDF")[-1]]
        if pdf: pdf_to_images(imgdir)
        for name in titles:
            self.add_para(name , style='Heading 2')
            root, imgs = get_root_sub(imgdir, ink=name,win=False,enz=".jpg")
            print(name,len(imgs))
            self.insert_imgs(root,imgs,sp=True if not name == list(titles)[-1] else False)  # 最后不分页
        self._save()


if __name__ == "__main__":
    # main()
    # dir = r"D:\Temp\2019"
    # # pdf_to_images(dir)
    # idoc = r"D:\Temp\2019\template.docx"
    # doc = Word(r"D:\Temp\2019\template.docx")
    # doc.pdf_word(dir,False)
    dir = r"E:\Temp\Publication\out\Python\Python.docx"
    extracts_imgs(dir)
    # get_root_sub(dir,enz="txt pdf",ink="一")
