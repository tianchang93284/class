from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QMainWindow, QPushButton, QFileDialog, QVBoxLayout, QWidget
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import Qt
from mainwindow import Ui_MainWindow  # 导入生成的界面类
import sys
import fitz  # PyMuPDF
import os
from docx2pdf import convert
from win32com.client import constants,gencache
import shutil
import random
import pandas as pd
from reportlab.pdfgen.canvas import Canvas
from pdfrw import PdfReader
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from DrawingDialog import DrawingDialog
import re
# 注册自定义字体
pdfmetrics.registerFont(TTFont("MyCustomFont", "STSONG.TTF"))

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setupUi(self)  # 调用 setupUi 初始化界面
        self.setWindowTitle("作业无忧")
        self.pushButton_openpdf.clicked.connect(self.open_pdf)
        # PDF 文件路径
        self.pdf_path = None
        self.label.mousePressEvent = self.get_click_position  # 捕获鼠标点击事件
        self.signature_position = None
        self.real_name_position = None
        self.pdf_width = None
        self.pdf_height = None
        self.real_name_position_width = 199
        self.real_name_position_height = 656
        self.pushButton_word2pdf.clicked.connect(self.word2pdf)
        self.pushButton_apply.clicked.connect(self.apply_ok)
        self.comment_lists = None
        self.target_files = None
        self.FONT_TT = 'MyCustomFont'
        self.pushButton_excel.clicked.connect(self.open_excel)
        self.pushButton_draw.clicked.connect(self.open_draw)

        #self.imagepath = None#['gou/1.png','gou/2.png','gou/3.png','gou/4.png']
        gou_images = []
        semigou_images = []
        x_images = []
        gou_path = "gou"
        # 遍历文件夹下的文件
        for file_name in os.listdir(gou_path):
            if file_name.startswith("gou") and file_name.endswith(".png"):
                gou_images.append(gou_path+"/"+file_name)
            elif file_name.startswith("semigou") and file_name.endswith(".png"):
                semigou_images.append(gou_path+"/"+file_name)
            elif file_name.startswith("x") and file_name.endswith(".png"):
                x_images.append(gou_path+"/"+file_name)
        self.imagepath = (gou_images * 7) + (semigou_images * 2) + (x_images * 1)

    def open_pdf(self):
        # 打开文件对话框选择 PDF 文件
        file_path, _ = QFileDialog.getOpenFileName(self, "打开 PDF 文件", "", "PDF 文件 (*.pdf)")
        if file_path:
            self.pdf_path = file_path
            self.display_pdf_first_page(file_path)
            #self.open_location_button.setEnabled(True)

    def display_pdf_first_page(self, file_path):
        # 使用 PyMuPDF 读取 PDF 文件第一页
        doc = fitz.open(file_path)
        first_page = doc.load_page(0)  # 加载第一页

        self.pdf_width = first_page.rect.width
        self.pdf_height = first_page.rect.height

        pix = first_page.get_pixmap()  # 渲染为图像

        # 直接将图像数据转换为 QImage
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
        pixmap = QPixmap.fromImage(img)

        # 显示在 QLabel 中
        self.label.setPixmap(pixmap.scaled(600, 800, Qt.KeepAspectRatio))
        # 清理文档
        doc.close()

    def get_click_position(self, event):
        # 捕获鼠标点击的位置
        self.signature_position = event.pos()
        print(f"签名位置: {self.signature_position}")
        label_pixmap = self.label.pixmap()
        x_ratio = self.pdf_width / label_pixmap.width()
        y_ratio = self.pdf_height / label_pixmap.height()

        # 将点击位置转换为 PDF 坐标
        self.real_name_position_width = self.signature_position.x() * x_ratio
        self.real_name_position_height = self.signature_position.y() * y_ratio

    def add_annotation_to_pdf(self):
        pass

    def word2pdf(self):
        # 选择源文件夹
        source_folder = QFileDialog.getExistingDirectory(self, "选择包含文件的文件夹")
        if not source_folder:
            return

        # 获取源文件夹的上一级目录，并创建 “统计” 文件夹
        parent_folder = os.path.dirname(source_folder)
        target_folder = os.path.join(parent_folder, "统计")
        os.makedirs(target_folder, exist_ok=True)

        self.target_files = target_folder

        # 删除文件
        genpath = os.path.join(os.getenv('temp'), 'gen_py')
        if os.path.exists(genpath):
            shutil.rmtree(genpath)
        word = gencache.EnsureDispatch("kwps.Application")  # 打开word应用程序
        #将此路径下的文件夹都复制过来，并将word转为pdf
        # 遍历源文件夹中的所有文件和文件夹
        for root, _, files in os.walk(source_folder):
            for file in files:
                source_file_path = os.path.join(root, file)

                # 保留原目录结构的目标文件路径
                relative_path = os.path.relpath(root, source_folder)
                target_dir = os.path.join(target_folder, relative_path)
                os.makedirs(target_dir, exist_ok=True)

                if file.endswith(".pdf"):
                    # 如果是 PDF 文件，直接复制
                    target_file_path = os.path.join(target_dir, file)
                    shutil.copy2(source_file_path, target_file_path)
                    print(f"已复制 PDF 文件: {target_file_path}")

                elif file.endswith(".doc") or file.endswith(".docx"):
                    # 如果是 Word 文件，转换为 PDF 并保存到目标路径
                    pdf_file_name = f"{os.path.splitext(file)[0]}.pdf"
                    target_file_path = os.path.join(target_dir, pdf_file_name)

                    doc = word.Documents.Open(source_file_path)  # 打开word文件
                    doc.ExportAsFixedFormat(target_file_path, constants.wdExportFormatPDF)  # 另存为后缀为".pdf"的文件，其中参数17表示为pdf
                    doc.Close()  # 关闭原来word文件
                    print(f"已转换 Word 文件并保存为 PDF: {target_file_path}")
        word.Quit()
        print("所有 Word 文件已转换为 PDF 文件并保存在 '统计' 文件夹中。")

    def get_score(self, pdf_file):
        if self.label_exel.text() == "":
            #如果pdf_file，含有A,B,C,D代表区间，如果没有从100-90取
            grades = ["A"]
            # 拆分字符串为各个区间，并解析为数值对
            ranges_with_grades = {}
            for i, part in enumerate(self.lineEdit_score.text().split(';')):
                start, end = map(int, part.split('-'))
                ranges_with_grades[grades[i]] = (start, end)

            #for grad in grades:
             #   if grad in pdf_file:
            grad = "A"
            return random.randint(ranges_with_grades[grad][1],ranges_with_grades[grad][0])
        else:
            # 使用 pandas 打开 Excel 文件
            try:
                df = pd.read_excel(self.label_exel.text())
                # 假设 Excel 文件中包含 "名字" 和 "成绩" 列
                if "姓名" in df.columns and "成绩" in df.columns:
                    names = df["姓名"].tolist()
                    scores = df["成绩"].tolist()
                    # 打印名字和成绩
                    for name, score in zip(names, scores):
                        if name in pdf_file:
                            print(f"姓名: {name}, 成绩: {score}")
                            return score
                else:
                    print("Excel文件中未找到 '姓名' 或 '成绩' 列")

            except Exception as e:
                print(f"无法读取Excel文件: {e}")

    def get_comment(self, pdf_file):
        # 分隔数据并构建字典
        comments_dict = {}
        #grades = ["A", "B", "C", "D", "E", "F"]
        for part in self.lineEdit_pingyu.text().split(";"):
            if ":" in part:
                # 拆分等级和评语部分
                grade, comments = part.split(":")
                grade = grade.strip()  # 去除等级两侧空格
                # 去除引号和多余空格，拆分评语为列表
                comments_list = [comment.strip().strip("'") for comment in comments.split(",") if comment.strip()]
                comments_dict[grade] = comments_list

        for grad,comment_list in comments_dict.items():
            if grad in pdf_file:
                return random.choice(comment_list)

    def score_pdf(self, pdf_file):
        txt_score = str(self.get_score(pdf_file))
        txt_comment = self.get_comment(pdf_file)
        
        text_conf = [[txt_score, self.real_name_position_width, 800-self.real_name_position_height, 60],
                    [txt_comment, self.real_name_position_width+60, 800-self.real_name_position_height+2, 18],
                    [self.lineEdit_teacherName.text(), self.real_name_position_width+15, 800-self.real_name_position_height-28, 20]
                    ]
        output_file = f'{os.path.splitext(pdf_file)[0]}_改.pdf'
        template = PdfReader(pdf_file)
        canvas = Canvas(output_file)

        template_obj0 = pagexobj(template.pages[0])
        obj0_name = makerl(canvas, template_obj0)
        canvas.doForm(obj0_name)
        for value in text_conf:#第一页打分
            canvas.setFont(self.FONT_TT, value[3])  # 设置字号
            canvas.setFillColorRGB(255, 0, 0)
            canvas.drawString(value[1], value[2], value[0])

        #第一页是否打勾
        if self.checkBox_head.isChecked():
            imge = random.choice(self.imagepath)
            canvas.drawImage(imge, 100, 220, 400, 300, mask=[0, 100, 0, 100, 0, 100])
        canvas.showPage()  # 关闭当前页，开始新页

        # 加入后续页面
        for i in range(1, len(template.pages)):
            template_obj1 = pagexobj(template.pages[i])
            obj1_name = makerl(canvas, template_obj1)
            canvas.doForm(obj1_name)
            imge = random.choice(self.imagepath)
            canvas.drawImage(imge, 100, 220, 400, 300, mask=[0, 100, 0, 100, 0, 100])
            canvas.showPage()
        canvas.save()
        
        return txt_score


    def open_excel(self):
        # 打开文件对话框，设置文件过滤为 Excel 文件
        file_path, _ = QFileDialog.getOpenFileName(None, "选择Excel文件", "", "Excel文件 (*.xls *.xlsx)")
        if file_path:
            self.label_exel.setText(file_path)

    def open_draw(self):
        #打开DrawDialog
        dialog = DrawingDialog()
        dialog.exec_()

    def get_path(self, filepath):
        items = os.listdir(filepath)
        for item in items:
            item_path = os.path.join(filepath, item)
            if os.path.isfile(item_path) and item.endswith((".pdf")):
                yield item_path
            elif os.path.isdir(item_path):
                yield from self.get_path(item_path)

    def apply_ok(self):
        self.target_files = QFileDialog.getExistingDirectory(self, "选择包含文件的文件夹")
        if not self.target_files:
            return
        self.scoresExcel, _ = QFileDialog.getOpenFileName(None, "选择Excel文件", "", "Excel文件 (*.xls *.xlsx)")
        if not self.scoresExcel:
            return
        self.score_df = pd.read_excel(self.scoresExcel)
        score_cols = [col for col in self.score_df.columns if re.match(r'score_\d+$', col)]
        if score_cols:
            # 提取已有字段中的数字编号
            max_index = max([int(re.search(r'(\d+)', col).group(1)) for col in score_cols])
            new_col_name = f'score_{max_index + 1}'
        else:
            new_col_name = 'score_1'
        self.score_df[new_col_name] = 0
        realscore=0
        #遍历pdf文件
        for full_path in self.get_path(self.target_files):
            if self.checkBox_zuoye.isChecked() and "作业" in full_path:
                realscore = self.score_pdf(full_path)
            elif self.checkBox_shiyan.isChecked()and "实验" in full_path:
                realscore = self.score_pdf(full_path)
            
            self.score_df[new_col_name] = realscore
            
            # 遍历所有姓名，如果 text 包含这个姓名，则赋值
            for i, name in self.score_df["姓名"].items():
                if pd.notnull(name) and str(name) in full_path:
                    df.at[i, new_col_name] = 90  # 设置你想给的分数
                    break
            
            #删除full_path
            print(f"批改文件: {full_path}")
            os.remove(full_path)
            
        
        self.score_df.to_excel(self.scoresExcel, index=False)

# 启动应用
app = QApplication(sys.argv)
window = MainWindow()
window.show()
sys.exit(app.exec_())




