''' 建议应用此工具批改作业时，先按优、良、差分别建立子目录，按优、良、差分别给随机分数和写评语'''

# pip install reportlab -i https://pypi.tuna.tsinghua.edu.cn/simple/
# pip install pdfrw -i https://pypi.tuna.tsinghua.edu.cn/simple/

import os
import random
import pandas as pd

from reportlab.pdfgen.canvas import Canvas
from pdfrw import PdfReader
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFont


#####################################################################################################
# 配置作业pdf文件路径
path = r'D:\trade\class\document\CBK'
#####################################################################################################

# 设置中文字体，避免乱码
registerFont(TTFont('Simhei', 'Simhei.ttf'))  # 黑体
registerFont(TTFont('STSong', 'STSONG.TTF'))  # 宋体
registerFont(TTFont('STKaiti', 'STKaiti.ttf')) # 楷体
registerFont(TTFont('STXINGKA', 'stxingka.ttf')) # 华文行楷

# FONT_TT = random.choice(['Simhei', 'STSong', 'STKaiti', 'STXINGKA'])
imagepath = ['gou/1.png','gou/2.png','gou/3.png','gou/4.png']
FONT_TT = 'STSong'
score_range={'A':[92,98],'B':[86,92], 'C':[81,86], 'D':[70,80], 'E':[60,70]}
conments = {
    'A':['实验比较认真，所画的图都正确', '实验结果正确，完成比较认真', '实验结果正确，内容比较充实', '报告内容详实，条理清晰', '数据分析方法严谨，数据解读恰当', '分析深入，对问题有清晰认识', '报告结论精准，结论符合实际'],
    'B':['实验比较认真，结果基本正确', '实验结果基本正确，完成认真', '实验步骤清晰，结果基本正确'],
    'C':['实验比较认真，部分结果错误', '实验结果部分有错，完成认真', '实验步骤清晰，结果部分有误'],
    'D':['实验比较认真，有些结果错误', '实验结果有些错误，完成认真', '实验步骤清晰，结果有些有误'],
    'E':['做实验不严谨, 有些结果错误', '实验结果有些错误，完成不认真', '实验步骤不清晰，结果有些有误']}

def write_score(labelname, fname, score):
    df = pd.read_excel('CBM名单-登记分数用.xlsx')
    labelnum = int(labelname[-2:])
    lab = '实验{}'.format(labelnum)
    for item in df['姓名']:
        if item in fname:
            df[lab][df['姓名']== item] = score
            df.to_excel('CBM名单-登记分数用.xlsx', index=False, header=True)
            return

    return

def get_grad(filename):
    df = pd.read_excel('CBM等级名单.xlsx')
    for item in df['姓名']:
        if item in filename:
            grad = df['等级'][df['姓名'] == item]
            return grad.values[0][0]
    #默认给一个B
    return 'B'

def create_class_test():
    test_path = r'D:\trade\class\document\CBM课堂测试word'
    paths = os.listdir(test_path)
    df = pd.read_excel('CBM名单-登记分数用.xlsx')
    for item in paths:
        grad = get_grad(item)
        txt_score = str(random.randint(score_range[grad][0], score_range[grad][1]))
        df['课堂测试'][df['姓名']==item.split('-')[1]] = txt_score
    df.to_excel('CBM名单-登记分数用.xlsx', index=False, header=True)
    return

def only_score_pdf(in_file, zuoye, scoresrange='100'):
    '''
    @in_file: 待批改的文件
    根据配置的分数做批改，并生成 in_file_批改.pdf
    '''
    #####################################################################################################
    # 需要配置批改的分数、评语、文本的坐标位置、字号
    # 数据处理与可视化作业2
    # text_conf = [['90', 400, 750, 60],
    #             ['作业比较认真', 380, 720, 18],
    #             ['郑耀东', 420, 690, 14]
    #             ]
    #
    # python 作业2
    # 文本，横坐标，纵坐标，字号
    #random.seed = 20220101
    txt_score = str(scoresrange)#str(random.randint(score_range[scoresrange][0], score_range[scoresrange][1]))
    #write_score(labelname, fname, txt_score)

    # python作业2
    #txt_comment = random.choice(conments[comment_num])
    global text_conf
    if zuoye == '作业1':#底部
        low = 40
        x = 60
        text_conf = [[txt_score, 230-x, 198-low, 60],
                    # [txt_comment, 290, 200-low, 18],
                    # ['赵稳', 185+x, 170-low, 20]
                    ]
    else:#右上角
        height = 590
        x = 100
        text_conf = [[txt_score, 230+x, 198+height, 60],
                    #[txt_comment, 290+x, 200+height, 18],
                    #['赵稳', 230+x, 180+height, 20]
                    ]
    #####################################################################################################
    output_file = f'{os.path.splitext(in_file)[0]}_改.pdf'

    template = PdfReader(in_file)
    canvas = Canvas(output_file)

    template_obj0 = pagexobj(template.pages[0])
    obj0_name = makerl(canvas, template_obj0)
    canvas.doForm(obj0_name)
    for value in text_conf:
        canvas.setFont(FONT_TT, value[3])  # 设置字号
        canvas.setFillColorRGB(255, 0, 0)
        canvas.drawString(value[1], value[2], value[0])
    #打红勾
    if zuoye != '作业1':
        imge = random.choice(imagepath)
        canvas.drawImage(imge,100,220,400,300, mask=[150,220,200,255,180,255])
    canvas.showPage()  # 关闭当前页，开始新页
    # 加入后续页面
    for i in range(1, len(template.pages)):
        template_obj1 = pagexobj(template.pages[i])
        obj1_name = makerl(canvas, template_obj1)
        canvas.doForm(obj1_name)
        imge = random.choice(imagepath)
        canvas.drawImage(imge, 100, 220, 400, 300, mask=[150, 220, 200, 255, 180, 255])
        canvas.showPage()
    canvas.save()


def score_pdf(in_file, scoresrange='100', comment_num='A', labelname = 'lab01'):
    '''
    @in_file: 待批改的文件
    根据配置的分数做批改，并生成 in_file_批改.pdf
    '''
    #####################################################################################################
    # 需要配置批改的分数、评语、文本的坐标位置、字号
    # 数据处理与可视化作业2
    # text_conf = [['90', 400, 750, 60],
    #             ['作业比较认真', 380, 720, 18],
    #             ['郑耀东', 420, 690, 14]
    #             ]
    #
    # python 作业2
    # 文本，横坐标，纵坐标，字号
    #random.seed = 20220101
    txt_score = str(scoresrange)#str(random.randint(score_range[scoresrange][0], score_range[scoresrange][1]))
    #write_score(labelname, fname, txt_score)

    # python作业2
    txt_comment = random.choice(conments[comment_num])
    global text_conf
    # if True or int(labelname[-2:])>4:
    #     low = 40
    #     x = 60
    #     text_conf = [[txt_score, 230, 198-low, 60],
    #                 [txt_comment, 290, 200-low, 18],
    #                 ['赵稳', 185+x, 170-low, 20]
    #                 ]
    # else:
    #     height = 590
    #     x = 50
    #     text_conf = [[txt_score, 230+x, 198+height, 60],
    #                 #[txt_comment, 290+x, 200+height, 18],
    #                 ['赵稳', 230+x, 180+height, 20]
    #                 ]
    #####################################################################################################
    output_file = f'{os.path.splitext(in_file)[0]}_改.pdf'

    template = PdfReader(in_file)
    canvas = Canvas(output_file)

    # template_obj0 = pagexobj(template.pages[0])
    # obj0_name = makerl(canvas, template_obj0)
    # canvas.doForm(obj0_name)
    # for value in text_conf:
    #     canvas.setFont(FONT_TT, value[3])  # 设置字号
    #     canvas.setFillColorRGB(255, 0, 0)
    #     canvas.drawString(value[1], value[2], value[0])
    #打红勾
    # if False and int(labelname[-2:]) < 5:
    #     imge = random.choice(imagepath)
    #     canvas.drawImage(imge,100,220,400,300, mask=[150,220,200,255,180,255])
    # canvas.showPage()  # 关闭当前页，开始新页
    # 加入后续页面
    for i in range(0, len(template.pages)):
        template_obj1 = pagexobj(template.pages[i])
        obj1_name = makerl(canvas, template_obj1)
        canvas.doForm(obj1_name)
        imge = random.choice(imagepath)
        canvas.drawImage(imge, 100, 220, 400, 300, mask=[150, 220, 200, 255, 180, 255])
        canvas.showPage()
    canvas.save()

def get_path(filepath):
    items = os.listdir(filepath)
    for item in items:
        item_path = os.path.join(filepath, item)
        if os.path.isfile(item_path) and item.endswith((".pdf")):
            yield item_path
        elif os.path.isdir(item_path):
            yield from get_path(item_path)


def score_pdf_excel():
    df = pd.read_excel('D:\\trade\class\\202312\BZL班级实验(1)\\bzL实验7和实验11.xlsx')

    pdfpath = 'D:\\trade\class\\202312\BZL班级实验(1)'
    for pdf in get_path(pdfpath):
        filename = os.path.splitext(pdf)[0].split('\\')[-1]
        stuno = os.path.splitext(pdf)[0].split('\\')[-2][:10]
        zuoye = os.path.splitext(pdf)[0].split('\\')[-3].split('+')[0]
        score = df[zuoye][df['学号'] == int(stuno)].values[0]
        only_score_pdf(pdf, zuoye, score)



    pass

def score_pdf_by_excel():
    pdfpath = 'D:\\trade\\class\\202312'
    # filename_list = os.listdir(pdfpath)
    # # 获取所有pdf文件名列表
    # pdf_list = [filename for filename in filename_list if filename.endswith(".pdf")]
    #df = pd.read_excel('D:\\trade\class\document\CCC班级书面作业3.xlsx')
    i = 0
    for pdf in get_path(pdfpath):
        score_pdf(pdf, scoresrange=str(random.randint(85,96)))
        # for item in df['学号']:
        #     if str(item) in pdf:
        #         score = df['作业3'][df['学号']==item]
        #         score_pdf(pdf, scoresrange=score.values[0])
        #         i+=1
        #         break
        i+=1
        # if i>=1:
        #     break
def score_pdf_all(path):
    for path1 in os.listdir(path):
        pathname = os.path.join(path, path1)
        if os.path.isdir(pathname):
            # 获取所有文件名的列表
            filename_list = os.listdir(pathname)
            # 获取所有pdf文件名列表
            pdf_list = [filename for filename in filename_list if filename.endswith(".pdf")]

            for pdf in pdf_list:
                pdf_file = (pathname + '/' + pdf)
                grad = get_grad(pdf)
                score_pdf(pdf_file, grad, grad, path1, pdf)


if __name__ == '__main__':

    # test
    n_file = r'D:\trade\class\qizhong\中期检查\BJC\实验1\2320710123-熊竟智\1041659321-01+2320710123_熊竟智_Docker容器安装与启动.pdf'
    score_pdf(in_file, scoresrange=str(random.randint(90,98)))

    #score_pdf_all(path)
    #create_class_test()
    #score_pdf_excel()