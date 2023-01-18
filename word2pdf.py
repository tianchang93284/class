# pip install comtypes -i https://pypi.tuna.tsinghua.edu.cn/simple/

import os
import comtypes.client
from docx2pdf import convert
from win32com import client
from win32com.client import constants,gencache

def get_path():
    # 获取当前运行路径
    # path = r'./'

    # 这里需要修改为你的word的文件夹路径
    path = r'D:\trade\class\document\CBM'

    for path1 in os.listdir(path):
        pathname = os.path.join(path,path1)
        if os.path.isdir(pathname):
            # 获取所有文件名的列表
            filename_list = os.listdir(pathname)
            # 获取所有word文件名列表
            wordname_list = [filename for filename in filename_list \
                             if filename.endswith((".doc", ".docx"))]
            for wordname in wordname_list:
                # 分离word文件名称和后缀，转化为pdf名称
                pdfname = os.path.splitext(wordname)[0] + '.pdf'
                # 如果当前word文件对应的pdf文件存在，则不转化
                if pdfname in filename_list:
                    continue
                # 拼接 路径和文件名
                wordpath = os.path.join(pathname, wordname)
                pdfpath = os.path.join(pathname, pdfname)
                # 生成器
                yield wordpath, pdfpath


def doc2pdf():
    word = gencache.EnsureDispatch("kwps.Application")  # 打开word应用程序
    #word = client.dynamic.Dispatch("kwps.Application")
    for wordpath, pdfpath in get_path():
        doc = word.Documents.Open(wordpath)  # 打开word文件
        doc.ExportAsFixedFormat(pdfpath, constants.wdExportFormatPDF)  # 另存为后缀为".pdf"的文件，其中参数17表示为pdf
        doc.Close()  # 关闭原来word文件
    word.Quit()

def convert_word_to_pdf():
    # word = comtypes.client.CreateObject("Word.Application")
    # word.Visible = 0
    for wordpath, pdfpath in get_path():
        convert(wordpath,wordpath)
        # newpdf = word.Documents.Open(wordpath)
        # newpdf.SaveAs(pdfpath, FileFormat=17)
        # newpdf.Close()

if __name__ == "__main__":
    #convert_word_to_pdf()
    doc2pdf()