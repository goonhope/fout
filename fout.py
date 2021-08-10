# -*- coding: utf-8 -*-
'''
@FileName	:   fout.py
@Created     :   2021/08/07 12:11
@Updated    :   2021/08/07 12:11
@Author		:   Teddy, goonhope@gmail.com, Zhuhai
@Function	:   function函数库规范化
@notes        ：知乎公开函数库
'''

import os,time,shutil,zipfile
from functools import wraps
from PIL import Image


def shower(show=True):
    '''print装饰器'''
    def inner(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            print(f"@{func.__name__}...") if show else None
            goal = func(*args, **kwargs)
            print(goal)
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


if __name__ == "__main__":
    # main()
    dir = r"E:\Temp\Publication\out\Python"
    extracts_imgs(dir)
    # print(os.listdir(dir))
    # get_root_sub(dir,enz="txt pdf",ink="一")