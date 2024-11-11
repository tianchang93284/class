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
        self.real_name_position_width = None
        self.real_name_position_height = None
        self.pushButton_word2pdf.clicked.connect(self.word2pdf)
        self.comment_lists = None
        self.target_files = None

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

    def apply_Ok(self):
        #先将word全部转成pdf
        pass



# 启动应用
app = QApplication(sys.argv)
window = MainWindow()
window.show()
sys.exit(app.exec_())




