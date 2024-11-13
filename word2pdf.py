# pip install comtypes -i https://pypi.tuna.tsinghua.edu.cn/simple/

import os
from docx2pdf import convert
from win32com.client import constants,gencache
import shutil



def get_path(filepath):
    # 获取当前运行路径
    # path = r'./'

    # 这里需要修改为你的word的文件夹路径

    items = os.listdir(filepath)
    for item in items:
        item_path = os.path.join(filepath, item)
        if os.path.isfile(item_path) and item.endswith((".doc", ".docx")):
            pdfname = os.path.splitext(item)[0] + '.pdf'
            pdfpath = os.path.join(filepath, pdfname)
            yield item_path, pdfpath
        elif os.path.isdir(item_path):
            yield from get_path(item_path)

    # for path1 in os.listdir(path):
    #     path2 = os.path.join(path, path1)
    #     for path3 in os.listdir(path2):
    #         path4 = os.path.join(path2, path3)
    #
    #         # 获取所有文件名的列表
    #         filename_list = os.listdir(path4)
    #         # 获取所有word文件名列表
    #         wordname_list = [filename for filename in filename_list \
    #                          if filename.endswith((".doc", ".docx"))]
    #         for wordname in wordname_list:
    #             # 分离word文件名称和后缀，转化为pdf名称
    #             pdfname = os.path.splitext(wordname)[0] + '.pdf'
    #             # 如果当前word文件对应的pdf文件存在，则不转化
    #             if pdfname in filename_list:
    #                 continue
    #             # 拼接 路径和文件名
    #             wordpath = os.path.join(path4, wordname)
    #             pdfpath = os.path.join(path4, pdfname)
    #             # 生成器
    #             yield wordpath, pdfpath


def doc2pdf():
    path = r'D:\trade\class\qizhong\中期检查'#要用绝对路径
    word = gencache.EnsureDispatch("kwps.Application")  # 打开word应用程序
    #word = client.dynamic.Dispatch("kwps.Application")
    for wordpath, pdfpath in get_path(path):
        doc = word.Documents.Open(wordpath)  # 打开word文件
        doc.ExportAsFixedFormat(pdfpath, constants.wdExportFormatPDF)  # 另存为后缀为".pdf"的文件，其中参数17表示为pdf
        doc.Close()  # 关闭原来word文件
    word.Quit()

def convert_word_to_pdf():
    # word = comtypes.client.CreateObject("Word.Application")
    # word.Visible = 0
    path = r'D:\trade\class\2023'
    for wordpath, pdfpath in get_path(path):
        convert(wordpath,wordpath)
        # newpdf = word.Documents.Open(wordpath)
        # newpdf.SaveAs(pdfpath, FileFormat=17)
        # newpdf.Close()


def get_file(path):
    items = os.listdir(path)
    for item in items:
        item_path = os.path.join(path, item)
        if os.path.isfile(item_path):
            yield item_path
        elif os.path.isdir(item_path):
            yield from get_file(item_path)

def delete_file():
    pdfpath = 'D:\\trade\class\document\软件工程-提交'
    for pdf in get_file(pdfpath):
        if '_改' not in pdf:
            os.remove(pdf)

#delete_file()

if __name__ == "__main__":
    #convert_word_to_pdf()

    #删除文件
    genpath = os.path.join(os.getenv('temp'), 'gen_py')
    if os.path.exists(genpath):
        shutil.rmtree(genpath)

    doc2pdf()